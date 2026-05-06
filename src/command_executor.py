"""Execute commands sent from the Linux adapter."""
from __future__ import annotations

import logging
import re
import time
import uuid
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import urlparse

import requests

from .config import Settings
from .wechat_provider import WeChatProvider

log = logging.getLogger(__name__)


_SAFE_FILENAME = re.compile(r"[^A-Za-z0-9._\-一-鿿]")


def _sanitize_filename(name: str) -> str:
    name = name.strip().replace("/", "_").replace("\\", "_")
    name = _SAFE_FILENAME.sub("_", name)
    return name or f"file_{uuid.uuid4().hex}"


class CommandExecutor:
    def __init__(self, settings: Settings, wechat: WeChatProvider) -> None:
        self._settings = settings
        self._wechat = wechat
        self._session = requests.Session()
        self._session.headers.update({"Authorization": f"Bearer {settings.adapter_auth_token}"})

    # ---------- public ----------

    def handle(self, msg: Dict[str, Any]) -> Dict[str, Any]:
        """Run a single command message; return an ack dict (never raises).

        Linux can deliver the command in two shapes:
          1. flat:    {"type": "command", "action": "send_text", "chat_name": ..., ...}
          2. nested:  {"type": "command", "command": {"action": "send_text", ...}}
        We unwrap the nested form and use the inner dict as the working payload.
        """
        inner = msg.get("command") if isinstance(msg.get("command"), dict) else msg
        cmd_id = str(inner.get("id") or inner.get("command_id") or msg.get("id") or uuid.uuid4().hex)
        command = str(inner.get("command") or inner.get("action") or "").strip()
        try:
            result = self._dispatch(command, inner)
            return {
                "type": "command.ack",
                "id": cmd_id,
                "command": command,
                "ok": True,
                "result": result,
                "ts": time.time(),
            }
        except Exception as exc:
            log.exception("command %s failed; payload=%s", command, inner)
            return {
                "type": "command.ack",
                "id": cmd_id,
                "command": command,
                "ok": False,
                "error": str(exc),
                "ts": time.time(),
            }

    # ---------- dispatch ----------

    def _dispatch(self, command: str, msg: Dict[str, Any]) -> Dict[str, Any]:
        chat = self._resolve_chat(msg)
        if command == "send_text":
            text = str(msg.get("text") or msg.get("content") or "")
            if not text:
                raise ValueError("send_text missing 'text'")
            at = msg.get("at") or None
            self._wechat.send_text(chat, text, at=at if isinstance(at, str) else None)
            return {"chat": chat, "len": len(text)}

        if command == "send_file":
            local_path = self._fetch_attachment(msg, kind="file")
            self._wechat.send_file(chat, str(local_path))
            return {"chat": chat, "path": str(local_path)}

        if command == "send_image":
            local_path = self._fetch_attachment(msg, kind="image")
            self._wechat.send_image(chat, str(local_path))
            return {"chat": chat, "path": str(local_path)}

        raise ValueError(f"unknown command: {command!r}")

    # ---------- helpers ----------

    @staticmethod
    def _resolve_chat(msg: Dict[str, Any]) -> str:
        log.debug("_resolve_chat: msg keys=%s", list(msg.keys()))
        # Prefer chat_name (a wxauto-recognizable display name); fall back to chat_id.
        for key in ("chat_name", "to_name", "to"):
            v = msg.get(key)
            log.debug("_resolve_chat: %s=%r", key, v)
            if v:
                return str(v)
        chat_id = msg.get("chat_id") or msg.get("to_id")
        log.debug("_resolve_chat: chat_id=%r", chat_id)
        if chat_id:
            s = str(chat_id)
            for prefix in ("private:", "group:", "chat:"):
                if s.startswith(prefix):
                    s = s[len(prefix):]
                    break
            return s
        raise ValueError("command missing chat target (chat_name / chat_id)")

    def _fetch_attachment(self, msg: Dict[str, Any], kind: str) -> Path:
        spec = msg.get("file") or msg.get("attachment") or {}
        if not isinstance(spec, dict):
            spec = {}

        url = spec.get("download_url") or spec.get("url") or msg.get("download_url")
        if not url:
            raise ValueError(f"send_{kind} missing download_url")

        filename = (
            spec.get("filename")
            or spec.get("name")
            or msg.get("filename")
            or Path(urlparse(url).path).name
            or f"{kind}_{uuid.uuid4().hex}"
        )
        filename = _sanitize_filename(filename)

        target = self._settings.download_dir / f"{int(time.time())}_{filename}"
        log.info("downloading %s -> %s", url, target)

        absolute = url if url.startswith(("http://", "https://")) else f"{self._settings.adapter_base_url}{url}"
        with self._session.get(absolute, stream=True, timeout=120) as resp:
            resp.raise_for_status()
            with open(target, "wb") as f:
                for chunk in resp.iter_content(chunk_size=64 * 1024):
                    if chunk:
                        f.write(chunk)
        return target

    def upload_inbound_media(self, local_path: str, filename: Optional[str] = None) -> Optional[str]:
        """Upload media received from WeChat to the Linux adapter.

        Returns the server-assigned id/url, or None on failure.
        Endpoint convention: POST {base}/v1/files (multipart 'file').
        """
        try:
            p = Path(local_path)
            if not p.exists():
                return None
            url = f"{self._settings.adapter_base_url}/v1/files"
            with open(p, "rb") as f:
                resp = self._session.post(
                    url,
                    files={"file": (filename or p.name, f)},
                    timeout=120,
                )
            resp.raise_for_status()
            data = resp.json()
            return data.get("file_id") or data.get("id") or data.get("url")
        except Exception:
            log.exception("upload_inbound_media failed")
            return None
