from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import List

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")


def _split_csv(value: str) -> List[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def _required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value or value == "REPLACE_ME":
        raise RuntimeError(
            f"Missing required env var {name}. Edit C:\\HermesWeChatSidecar\\.env"
        )
    return value


@dataclass(frozen=True)
class Settings:
    adapter_base_url: str
    adapter_ws_url: str
    adapter_auth_token: str
    device_id: str
    listen_chats: List[str]
    group_prefixes: List[str]
    bot_at_name: str
    heartbeat_interval: float
    reconnect_min: float
    reconnect_max: float
    download_dir: Path
    inbound_media_dir: Path
    log_level: str
    seen_ids_path: Path = field(default=ROOT / "data" / "seen_ids.json")

    @classmethod
    def load(cls) -> "Settings":
        download_dir = Path(os.getenv("DOWNLOAD_DIR", str(ROOT / "data" / "downloads")))
        inbound_dir = Path(os.getenv("INBOUND_MEDIA_DIR", str(ROOT / "data" / "inbound")))
        download_dir.mkdir(parents=True, exist_ok=True)
        inbound_dir.mkdir(parents=True, exist_ok=True)

        return cls(
            adapter_base_url=_required("ADAPTER_BASE_URL").rstrip("/"),
            adapter_ws_url=_required("ADAPTER_WS_URL"),
            adapter_auth_token=_required("ADAPTER_AUTH_TOKEN"),
            device_id=os.getenv("DEVICE_ID", "windows-main").strip() or "windows-main",
            listen_chats=_split_csv(os.getenv("LISTEN_CHATS", "文件传输助手")),
            group_prefixes=_split_csv(os.getenv("GROUP_PREFIXES", "/")),
            bot_at_name=os.getenv("BOT_AT_NAME", "").strip(),
            heartbeat_interval=float(os.getenv("HEARTBEAT_INTERVAL", "20")),
            reconnect_min=float(os.getenv("RECONNECT_MIN", "1")),
            reconnect_max=float(os.getenv("RECONNECT_MAX", "30")),
            download_dir=download_dir,
            inbound_media_dir=inbound_dir,
            log_level=os.getenv("LOG_LEVEL", "INFO").upper(),
        )
