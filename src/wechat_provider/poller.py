"""Read messages out of a popped-out chat sub-window.

The chat_message_list is a virtualised Qt ListView that only exposes the
most-recent ~10 rows. We dedup by *previous-poll snapshot*, not by a
monotonic high-water mark — repeating an old text would otherwise be
mis-classified as already-seen because UIA can no longer show all the
past instances.
"""
from __future__ import annotations

import logging
import time
from collections import Counter
from typing import Any, Callable, Deque, Dict, List, Tuple

from ._types import MSG_LIST_AID, MSG_TEXT_CLASS, IncomingMessage
from .uia_utils import find_by_aid
from .win32_utils import stable_hash

log = logging.getLogger(__name__)


# Re-popout drives the WeChat UI, so a chat whose sub-window cannot be
# recovered must not be retried every poll cycle.
_RESOLVE_COOLDOWN_S = 60.0


def _hwnd_alive(hwnd: int) -> bool:
    """Cheap liveness check for a stored window handle."""
    if not hwnd:
        return False
    try:
        import win32gui
        return bool(win32gui.IsWindow(hwnd))
    except Exception:
        return True  # can't tell — let ControlFromHandle be the judge


def _resolve_subwindow(name: str) -> int:
    """Re-discover *name*'s popped-out sub-window, re-popping it out if the
    window was closed. Returns a fresh hwnd, or 0 if recovery failed."""
    try:
        from . import popout, window_finder
        pids = window_finder.find_weixin_pids()
        if not pids:
            log.warning("re-resolve(%s): no running Weixin.exe", name)
            return 0
        hwnd = window_finder.find_chat_subwindow(name, pids)
        if hwnd:
            return hwnd
        # Sub-window is gone — re-pop it out. Needs the main 微信 window
        # reachable (same precondition as startup); best-effort otherwise.
        log.info("re-resolve(%s): sub-window missing, re-popping out", name)
        popout.popout_chat(name, pids)
        for _ in range(25):  # up to ~5s for the new window to appear
            hwnd = window_finder.find_chat_subwindow(name, pids)
            if hwnd:
                return hwnd
            time.sleep(0.2)
        log.warning("re-resolve(%s): sub-window did not reappear after re-popout", name)
        return 0
    except Exception:
        log.exception("re-resolve(%s) failed", name)
        return 0


def _get_root(chat_state: Dict[str, Dict[str, Any]], name: str):
    """UIA root control of *name*'s sub-window.

    The stored hwnd goes stale if the popped-out window is closed or WeChat
    restarts. On a stale handle, re-discover the sub-window (re-popping it
    out if needed), update chat_state, re-seed the message snapshot, and
    retry. Recovery is rate-limited so an unrecoverable chat does not drive
    the WeChat UI every poll cycle. Returns None when recovery is not
    possible right now (caller treats it like a missing message list)."""
    import uiautomation as auto

    state = chat_state[name]
    hwnd = state.get("hwnd") or 0

    if _hwnd_alive(hwnd):
        try:
            root = auto.ControlFromHandle(hwnd)
            if root is not None:
                return root
        except Exception as exc:
            log.warning("get_root(%s): hwnd %s unusable: %s", name, hex(hwnd), exc)

    # Handle is stale — rebuild it, but at most once per cooldown.
    if time.time() - state.get("last_resolve_ts", 0.0) < _RESOLVE_COOLDOWN_S:
        return None
    state["last_resolve_ts"] = time.time()

    new_hwnd = _resolve_subwindow(name)
    if not new_hwnd:
        log.error("get_root(%s): could not re-resolve sub-window; retry in ~%.0fs",
                  name, _RESOLVE_COOLDOWN_S)
        return None

    state["hwnd"] = new_hwnd
    log.info("get_root(%s): re-resolved sub-window hwnd=%s", name, hex(new_hwnd))
    try:
        root = auto.ControlFromHandle(new_hwnd)
    except Exception:
        log.exception("get_root(%s): re-resolved hwnd still unusable", name)
        return None
    if root is not None:
        # Re-seed the snapshot so the rebuilt window's existing rows are not
        # all re-emitted as new messages on the next poll.
        snapshot_existing(chat_state, name)
    return root


def _get_msg_list(chat_state: Dict[str, Dict[str, Any]], name: str):
    return find_by_aid(_get_root(chat_state, name), MSG_LIST_AID)


def snapshot_existing(chat_state: Dict[str, Dict[str, Any]], name: str) -> None:
    """Seed the chat's seen_counts with whatever's currently visible in the
    message list, so we don't re-emit pre-existing history as new messages."""
    ml = _get_msg_list(chat_state, name)
    if ml is None:
        return
    items = ml.GetChildren()
    seen: Counter = Counter()
    for it in items:
        txt = it.Name or ""
        if txt and (it.ClassName or "").startswith(MSG_TEXT_CLASS):
            seen[txt] += 1
    chat_state[name]["seen_counts"] = seen
    log.info(
        "chat %r: snapshotted %d existing messages (%d unique)",
        name, sum(seen.values()), len(seen),
    )


def poll_chat(
    chat_state: Dict[str, Dict[str, Any]],
    recent_outbound: Deque[Tuple[str, str, float]],
    on_message: Callable[[IncomingMessage], None],
    bot_at_name: str,
    name: str,
) -> None:
    """One poll cycle for *name*. Pulls the message list, computes new rows
    against the previous snapshot, filters out our own outbound echoes,
    and invokes *on_message* for each genuinely-new message."""
    ml = _get_msg_list(chat_state, name)
    if ml is None:
        log.warning("poll(%s): message list not found; sub-window may be closed", name)
        return

    items = ml.GetChildren()
    if not items:
        return

    # `seen` holds the (text -> count) snapshot from the *previous* poll —
    # NOT a monotonic high-water mark. See module docstring for rationale.
    seen: Counter = chat_state[name]["seen_counts"]
    observed: Counter = Counter()
    new_msgs: List[Tuple[str, str, str]] = []   # (text, class_name, runtime_id)

    for it in items:
        txt = it.Name or ""
        cls = it.ClassName or ""
        if not txt:
            continue
        if not cls.startswith(MSG_TEXT_CLASS):
            # Time-of-day dividers, etc.
            continue
        observed[txt] += 1
        if observed[txt] > seen[txt]:
            new_msgs.append((txt, cls, ""))

    # Replace the snapshot wholesale so counts can decrease as old rows
    # scroll out of the UIA window.
    chat_state[name]["seen_counts"] = observed

    if not new_msgs:
        return

    log.debug("poll(%s): %d new", name, len(new_msgs))

    # Discard messages that were our own outbound sends within the last 30s.
    now = time.time()
    recent_self = {(c, t) for (c, t, ts) in recent_outbound if now - ts < 30}

    for txt, cls, _runtime in new_msgs:
        if (name, txt) in recent_self:
            # Consume one occurrence so a manual user-typed dup will still
            # flow next time. Remove only the matching tuple instance.
            for i, (c, t, _ts) in enumerate(recent_outbound):
                if c == name and t == txt:
                    del recent_outbound[i]
                    break
            log.debug("poll(%s): skip self-echo %r", name, txt[:60])
            continue
        _emit(chat_state, on_message, bot_at_name, name, txt, cls)


def _emit(
    chat_state: Dict[str, Dict[str, Any]],
    on_message: Callable[[IncomingMessage], None],
    bot_at_name: str,
    name: str,
    text: str,
    raw_class: str,
) -> None:
    state = chat_state[name]
    chat_type = state.get("chat_type", "private")

    at_self = False
    if chat_type == "group" and bot_at_name:
        # WeChat uses U+2005 (FOUR-PER-EM SPACE) as the @-mention separator,
        # but accept a regular space too in case the layout differs.
        for sep in ("\u2005", " "):
            if f"@{bot_at_name}{sep}" in text or text.endswith(f"@{bot_at_name}"):
                at_self = True
                break

    sender = name if chat_type == "private" else ""
    msg = IncomingMessage(
        chat_name=name,
        chat_type=chat_type,
        sender=sender,
        content_type="text",
        text=text,
        file_path=None,
        raw_type=raw_class,
        msg_hash=stable_hash(name, text, time.time()),
        ts=time.time(),
        at_self=at_self,
    )
    log.info(
        "uia: NEW msg from %s [%s]: %r%s",
        name, chat_type, text[:80], " (at_self)" if at_self else "",
    )
    try:
        on_message(msg)
    except Exception:
        log.exception("on_message callback raised")
