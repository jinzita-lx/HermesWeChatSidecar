"""Dump UIA tree of the 测试虾 group sub-window to figure out sender extraction.

Run from project root with the sidecar venv:
    .venv\\Scripts\\python.exe scripts\\probe_group_uia.py
"""
from __future__ import annotations

import sys
import psutil
import win32gui
import win32process
import uiautomation as auto

CHAT_NAME = "测试虾"
WEIXIN_CLS = "Qt51514QWindowIcon"
WEIXIN_PROC = "Weixin.exe"
MAX_DEPTH = 8
MAX_BUBBLES = 6  # only dump the most recent N message-list children


def find_hwnd(name: str) -> int:
    pids = {
        p.pid for p in psutil.process_iter(["name"])
        if p.info.get("name") in (WEIXIN_PROC, "Weixin")
    }
    found = []
    def cb(h, _):
        try:
            if win32gui.GetClassName(h) != WEIXIN_CLS:
                return
            _, pid = win32process.GetWindowThreadProcessId(h)
            if pid not in pids:
                return
            if win32gui.GetWindowText(h) == name:
                found.append(h)
        except Exception:
            pass
    win32gui.EnumWindows(cb, None)
    return found[0] if found else 0


def fmt_node(c, depth: int) -> str:
    name = (c.Name or "")[:80]
    aid = c.AutomationId or ""
    cls = c.ClassName or ""
    ctl = c.ControlTypeName or ""
    extras = []
    try:
        if c.LocalizedControlType:
            extras.append(f"lct={c.LocalizedControlType!r}")
    except Exception:
        pass
    for attr in ("HelpText", "AccessKey"):
        try:
            v = getattr(c, attr, None)
            if v:
                extras.append(f"{attr.lower()}={str(v)[:40]!r}")
        except Exception:
            pass
    try:
        lp = c.GetLegacyIAccessiblePattern()
        if lp:
            d = lp.Description or ""
            v = lp.Value or ""
            if d:
                extras.append(f"acc_desc={d[:60]!r}")
            if v and v != name:
                extras.append(f"acc_val={v[:60]!r}")
    except Exception:
        pass
    extra_str = (" " + " ".join(extras)) if extras else ""
    return f"{'  '*depth}- {ctl} cls={cls!r} aid={aid!r} name={name!r}{extra_str}"


_RAW_WALKER = auto.GetRootControl().GetRawViewWalker() if hasattr(auto.GetRootControl(), 'GetRawViewWalker') else None


def raw_children(c):
    """Use UIA raw walker so we don't lose elements that ContentView filters out."""
    try:
        if _RAW_WALKER is None:
            return c.GetChildren()
        out = []
        kid = _RAW_WALKER.GetFirstChildElement(c)
        while kid:
            out.append(kid)
            kid = _RAW_WALKER.GetNextSiblingElement(kid)
        return out
    except Exception:
        return c.GetChildren()


def walk(c, depth: int):
    if depth > MAX_DEPTH:
        return
    print(fmt_node(c, depth))
    kids = raw_children(c)
    for k in kids:
        walk(k, depth + 1)


def find_by_aid(node, aid, depth=12):
    if depth <= 0 or node is None:
        return None
    if (node.AutomationId or "") == aid:
        return node
    try:
        for kid in node.GetChildren():
            r = find_by_aid(kid, aid, depth - 1)
            if r is not None:
                return r
    except Exception:
        pass
    return None


def main() -> int:
    hwnd = find_hwnd(CHAT_NAME)
    if not hwnd:
        print(f"sub-window not found for {CHAT_NAME!r}", file=sys.stderr)
        return 1
    print(f"hwnd = {hex(hwnd)}")

    root = auto.ControlFromHandle(hwnd)
    print(fmt_node(root, 0))

    msg_list = find_by_aid(root, "chat_message_list")
    if not msg_list:
        print("chat_message_list not found", file=sys.stderr)
        return 2

    print("\n=== msg list children (last %d) ===" % MAX_BUBBLES)
    items = msg_list.GetChildren()
    print(f"total children: {len(items)}")
    for it in items[-MAX_BUBBLES:]:
        print()
        walk(it, 0)

    print("\n=== input/send controls ===")
    inp = find_by_aid(root, "chat_input_field")
    if inp:
        print(fmt_node(inp, 0))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
