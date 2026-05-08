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


def _get_root(chat_state: Dict[str, Dict[str, Any]], name: str):
    import uiautomation as auto
    return auto.ControlFromHandle(chat_state[name]["hwnd"])


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
