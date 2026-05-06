"""Sidecar entrypoint."""
from __future__ import annotations

import asyncio
import logging
import signal
import sys
import time
from pathlib import Path
from typing import Any, Dict, Optional

from .command_executor import CommandExecutor
from .config import ROOT, Settings
from .dedup import SeenIds
from .logging_setup import configure as configure_logging
from .wechat_provider import IncomingMessage, WeChatProvider
from .ws_client import WSClient

log = logging.getLogger("sidecar")


def _msg_to_payload(settings: Settings, msg: IncomingMessage) -> Dict[str, Any]:
    chat_id = (
        f"private:{msg.chat_name}" if msg.chat_type == "private" else f"group:{msg.chat_name}"
    )
    is_group = msg.chat_type == "group"
    payload: Dict[str, Any] = {
        "type": "wechat.message",
        "id": msg.stable_id(),
        "device_id": settings.device_id,
        "ts": msg.ts,
        "chat_id": chat_id,
        "chat_name": msg.chat_name,
        "chat_type": msg.chat_type,
        "is_group": is_group,
        "at_self": msg.at_self,
        "sender_id": msg.sender,
        "sender_name": msg.sender,
        "content_type": msg.content_type,
        "text": msg.text,
        # Keep nested forms too for clients that prefer them.
        "chat": {"id": chat_id, "name": msg.chat_name, "type": msg.chat_type},
        "sender": {"id": msg.sender, "name": msg.sender},
        "content": {"type": msg.content_type, "text": msg.text},
        "raw": {"type": msg.raw_type},
    }
    if msg.file_path:
        payload["file_path"] = msg.file_path
        payload["filename"] = Path(msg.file_path).name
        payload["content"]["file_path"] = msg.file_path
        payload["content"]["filename"] = Path(msg.file_path).name
    return payload


class Sidecar:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._seen = SeenIds(settings.seen_ids_path)
        self._wechat = WeChatProvider(
            listen_chats=settings.listen_chats,
            on_message=self._handle_wechat_msg,
            inbound_media_dir=settings.inbound_media_dir,
            bot_at_name=settings.bot_at_name,
        )
        self._executor = CommandExecutor(settings, self._wechat)
        self._ws = WSClient(
            url=settings.adapter_ws_url,
            on_message=self._handle_inbound,
            heartbeat_interval=settings.heartbeat_interval,
            reconnect_min=settings.reconnect_min,
            reconnect_max=settings.reconnect_max,
        )

    # ---------- WeChat -> Linux ----------

    def _handle_wechat_msg(self, msg: IncomingMessage) -> None:
        mid = msg.stable_id()
        if not self._seen.add_if_new(mid):
            log.debug("dedup skip %s", mid)
            return

        # Group filtering is delegated to the Linux adapter
        # (WECHAT_REQUIRE_MENTION_IN_GROUPS + activation prefixes + at_self).
        # We forward all real text messages and let the adapter decide.

        payload = _msg_to_payload(self._settings, msg)
        log.info(
            "wx -> linux: chat=%s type=%s text=%r",
            payload["chat"]["name"], payload["content"]["type"], payload["content"]["text"][:80],
        )
        self._ws.send_threadsafe(payload)

    # ---------- Linux -> WeChat ----------

    async def _handle_inbound(self, msg: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        mtype = msg.get("type")
        if mtype == "hello":
            log.info("got hello: %s", {k: v for k, v in msg.items() if k != "type"})
            # Linux doesn't expect any reply — just log and move on.
            return None
        if mtype == "command":
            # Run synchronous wxauto calls on a thread to avoid blocking the loop.
            ack = await asyncio.to_thread(self._executor.handle, msg)
            return ack
        if mtype == "error":
            log.warning("server error frame: %s", {k: v for k, v in msg.items() if k != "type"})
            return None
        log.info("unhandled inbound type=%s body=%s", mtype, {k: v for k, v in msg.items() if k != "type"})
        return None

    # ---------- lifecycle ----------

    async def run(self) -> None:
        self._wechat.start()

        loop = asyncio.get_running_loop()
        if sys.platform != "win32":
            for sig in (signal.SIGINT, signal.SIGTERM):
                loop.add_signal_handler(sig, self._ws.stop)

        await self._ws.run()

    def stop(self) -> None:
        self._ws.stop()
        self._wechat.stop()


def main() -> int:
    try:
        settings = Settings.load()
    except Exception as exc:
        print(f"FATAL: {exc}", file=sys.stderr)
        return 2

    # Pre-import wxauto4 so its module-level logger setup (which calls
    # `root_logger.handlers.clear()`) runs BEFORE we install our handlers.
    # Also disable wxauto4's per-day file logger to avoid double-logging.
    try:
        import wxauto4
        wxauto4.WxParam.ENABLE_FILE_LOGGER = False
    except Exception:
        pass

    configure_logging(settings.log_level, ROOT / "logs")
    log.info("sidecar starting | device_id=%s | adapter=%s", settings.device_id, settings.adapter_base_url)
    log.info("listening chats: %s", settings.listen_chats)

    sidecar = Sidecar(settings)
    try:
        asyncio.run(sidecar.run())
    except KeyboardInterrupt:
        log.info("interrupted")
    finally:
        sidecar.stop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
