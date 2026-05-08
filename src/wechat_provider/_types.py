"""Constants and value types shared across the wechat_provider package.

The constants describe the WeChat 4.x UI fingerprint we depend on (window
class names, UIA AutomationIds for chat sub-window controls). They are
empirical — bump them if WeChat ships a major UI rewrite.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


# Subordinate Qt window class used by Weixin 4.x for both the main window and
# popped-out chats. Filter further by Name and process.
WEIXIN_WINDOW_CLASS = "Qt51514QWindowIcon"
WEIXIN_PROCESS_NAME = "Weixin.exe"

# UIA AutomationIds inside a popped-out chat sub-window.
INPUT_AID = "chat_input_field"
MSG_LIST_AID = "chat_message_list"

# Message list-item classes — the chat-bubble row vs. system/divider row.
MSG_TEXT_CLASS = "mmui::ChatTextItemView"
MSG_DIVIDER_CLASS = "mmui::ChatItemView"   # time-of-day separators


@dataclass
class IncomingMessage:
    chat_name: str
    chat_type: str          # "private" | "group"
    sender: str
    content_type: str       # "text" | "image" | "file" | "other"
    text: str
    file_path: Optional[str]
    raw_type: str
    msg_hash: str
    ts: float
    at_self: bool = False   # group only: text contains @<bot_at_name>

    def stable_id(self) -> str:
        return f"uia_{self.msg_hash}" if self.msg_hash else f"uia_{int(self.ts*1000)}"
