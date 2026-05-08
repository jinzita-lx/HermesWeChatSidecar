"""Send a plain-text message into a popped-out chat sub-window.

We try strategies in increasing visual cost:
  1. UIA InvokePattern on the send button — pure UIA, no visual change.
  2. PostMessage VK_RETURN to the top-level hwnd — works when input has focus.
  3. Restore window without activating, real Click(), re-minimize — last
     resort, briefly visible (~150ms) but never grabs keyboard focus.
"""
from __future__ import annotations

import logging
import time
from typing import Any, Deque, Dict, Optional, Tuple

from ._types import INPUT_AID
from .uia_utils import find_by, find_by_aid

log = logging.getLogger(__name__)


def _get_root(chat_state: Dict[str, Dict[str, Any]], name: str):
    import uiautomation as auto
    return auto.ControlFromHandle(chat_state[name]["hwnd"])


def _get_input(chat_state: Dict[str, Dict[str, Any]], name: str):
    return find_by_aid(_get_root(chat_state, name), INPUT_AID)


def _get_send_btn(chat_state: Dict[str, Dict[str, Any]], name: str):
    root = _get_root(chat_state, name)
    return find_by(
        root,
        lambda c: (
            c.ControlTypeName == "ButtonControl"
            and c.Name == "发送"
            and "XOutlineButton" in (c.ClassName or "")
        ),
    )


def send_text(
    chat_state: Dict[str, Dict[str, Any]],
    recent_outbound: Deque[Tuple[str, str, float]],
    chat: str,
    text: str,
) -> None:
    if chat not in chat_state:
        raise RuntimeError(f"chat not in LISTEN_CHATS: {chat!r}")
    edit = _get_input(chat_state, chat)
    btn = _get_send_btn(chat_state, chat)
    if edit is None or btn is None:
        raise RuntimeError(f"input/send controls not found in sub-window for {chat!r}")
    edit.GetValuePattern().SetValue(text)
    time.sleep(0.05)

    hwnd = chat_state[chat]["hwnd"]
    strategy = fire_send(hwnd, edit, btn)
    if not strategy:
        raise RuntimeError(f"send_text({chat!r}): all send strategies failed")

    recent_outbound.append((chat, text, time.time()))
    log.debug("send_text(%s, %r) ok via %s", chat, text[:80], strategy)


def fire_send(hwnd: int, edit, btn) -> Optional[str]:
    """Try send strategies until the input clears. Returns the name of the
    strategy that succeeded, or None if all failed."""
    import win32con
    import win32gui

    def input_cleared() -> bool:
        try:
            return not (edit.GetValuePattern().Value or "")
        except Exception:
            return False

    # 1) InvokePattern on the send button — pure UIA, no visual change.
    try:
        ip = btn.GetInvokePattern()
        if ip is not None:
            ip.Invoke()
            time.sleep(0.15)
            if input_cleared():
                return "invoke"
    except Exception:
        log.debug("invoke strategy raised", exc_info=True)

    # 2) PostMessage Enter to the top-level HWND. Works for Qt windows that
    #    have the input as their focused widget.
    try:
        try:
            edit.SetFocus()
        except Exception:
            pass
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.2)
        if input_cleared():
            return "post_enter"
    except Exception:
        log.debug("post_enter strategy raised", exc_info=True)

    # 3) Restore window without activating, click, re-minimize.
    try:
        was_minimized = bool(win32gui.IsIconic(hwnd))
        if was_minimized:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
            time.sleep(0.12)
        btn.Click(simulateMove=False, waitTime=0)
        time.sleep(0.15)
        sent = input_cleared()
        if was_minimized:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINNOACTIVE)
        if sent:
            return "restore_click" if was_minimized else "click"
    except Exception:
        log.debug("restore_click strategy raised", exc_info=True)

    return None
