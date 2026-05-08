"""Send a file/image into a popped-out chat sub-window.

Strategy: click the toolbar's '发送文件' button to bring up the standard
Win32 file-open dialog, drive that dialog via UIA + PostMessage to fill
in the absolute path and submit, then fire the chat send button once
WeChat has attached the file to the composition area.

We pay a lot of attention to staying invisible:
  * The chat sub-window is un-minimized via SetWindowPlacement with
    showCmd=SW_SHOWNOACTIVATE so it can receive synthetic mouse messages
    without taking focus and without leaving its off-screen position.
  * Any new top-level Weixin window is *immediately* moved to (-30000, -30000)
    via SWP_NOACTIVATE so the user never sees the file dialog flash on screen.
  * Submission goes through PostMessage(WM_COMMAND, IDOK) directly — UIA
    only exposes the Open button's split-button DropDown halves, so a UIA
    Click() opens the dropdown menu instead of submitting.
"""
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any, Callable, Deque, Dict, List, Optional, Set, Tuple

from .uia_utils import find_by, find_by_aid, has_edit_control
from .win32_utils import post_click
from .window_finder import weixin_top_level_hwnds

log = logging.getLogger(__name__)


def _get_root(chat_state: Dict[str, Dict[str, Any]], name: str):
    import uiautomation as auto
    return auto.ControlFromHandle(chat_state[name]["hwnd"])


def _get_send_file_btn(chat_state: Dict[str, Dict[str, Any]], chat: str):
    """Locate the '发送文件' button in the chat sub-window's toolbar."""
    root = _get_root(chat_state, chat)
    toolbar = find_by_aid(root, "tool_bar_accessible")
    if toolbar is None:
        return None
    return find_by(
        toolbar,
        lambda c: c.ControlTypeName == "ButtonControl" and c.Name == "发送文件",
    )


def send_file(
    chat_state: Dict[str, Dict[str, Any]],
    chat: str,
    path: str,
) -> None:
    if chat not in chat_state:
        raise RuntimeError(f"chat not in LISTEN_CHATS: {chat!r}")
    p = Path(path)
    if not p.exists() or not p.is_file():
        raise FileNotFoundError(p)
    abspath = str(p.resolve())

    file_btn = _get_send_file_btn(chat_state, chat)
    if file_btn is None:
        raise RuntimeError(f"toolbar '发送文件' button not found in {chat!r}")

    # Snapshot existing top-level Weixin windows so we can detect the new
    # one that the file dialog opens. ClassName-based detection (#32770) is
    # too narrow: WeChat 4.x sometimes pops a Qt-skinned dialog that shares
    # the Qt51514QWindowIcon class with regular sub-windows.
    existing_hwnds = weixin_top_level_hwnds()

    # Qt drops WM_* mouse messages on minimized windows, so a stealth
    # PostMessage click no-ops while the chat window is iconic. We
    # un-minimize via SetWindowPlacement (showCmd = SW_SHOWNOACTIVATE)
    # which un-minimizes WITHOUT activating and KEEPS the original
    # off-screen rect intact — so the window stays invisible to the user
    # but is "live" enough to process synthetic input.
    import win32con
    import win32gui as _w32g
    sub_hwnd = chat_state[chat]["hwnd"]

    orig_placement = _w32g.GetWindowPlacement(sub_hwnd)
    was_minimized = bool(_w32g.IsIconic(sub_hwnd))
    placement_changed = False
    if was_minimized:
        try:
            un_min_placement = (
                orig_placement[0],
                win32con.SW_SHOWNOACTIVATE,
                orig_placement[2],
                orig_placement[3],
                orig_placement[4],
            )
            _w32g.SetWindowPlacement(sub_hwnd, un_min_placement)
            placement_changed = True
            time.sleep(0.15)
        except Exception:
            log.debug("SetWindowPlacement to un-minimize failed", exc_info=True)

    # Re-resolve the button after placement change (UIA element handles may
    # need a refresh once the window is active again).
    file_btn = _get_send_file_btn(chat_state, chat) or file_btn

    try:
        stealth_ok = post_click(sub_hwnd, file_btn)
        log.debug("send_file: stealth post-click ok=%s", stealth_ok)
        dlg = _wait_new_weixin_window(
            existing_hwnds, exclude={sub_hwnd}, timeout_s=6.0,
        )

        # Fallback only if PostMessage didn't surface a dialog: try a real
        # UIA Click. Window stays off-screen; if Click insists on screen
        # coordinates the next layer will surface a clear error.
        if dlg is None:
            log.debug("send_file: stealth produced no dialog; trying UIA Click")
            try:
                file_btn.Click(simulateMove=False, waitTime=0)
            except Exception:
                log.debug("file_btn.Click raised; trying InvokePattern", exc_info=True)
                try:
                    ip = file_btn.GetInvokePattern()
                    if ip is not None:
                        ip.Invoke()
                except Exception as exc:
                    raise RuntimeError(f"could not invoke 发送文件 button: {exc}")
            dlg = _wait_new_weixin_window(
                existing_hwnds, exclude={sub_hwnd}, timeout_s=6.0,
            )

        if dlg is None:
            snapshot = weixin_top_level_hwnds()
            added = snapshot - existing_hwnds - {sub_hwnd}
            log.debug(
                "send_file: no new window; existing=%d snapshot=%d added=%s",
                len(existing_hwnds), len(snapshot), [hex(h) for h in added],
            )
            raise RuntimeError(
                "file dialog did not appear after clicking 发送文件 (both stealth and visible)"
            )

        log.debug(
            "send_file: detected new window hwnd=%s class=%r title=%r",
            hex(getattr(dlg, "NativeWindowHandle", 0) or 0),
            _w32g.GetClassName(dlg.NativeWindowHandle) if dlg.NativeWindowHandle else "",
            _w32g.GetWindowText(dlg.NativeWindowHandle) if dlg.NativeWindowHandle else "",
        )

        try:
            _fill_file_dialog(dlg, abspath)
            _wait_dialog_closed(dlg, timeout_s=6.0)
        except Exception:
            _dismiss_file_dialog(dlg)
            raise

        # Composition area now has the attachment chip; brief settle delay
        # before firing send so the button's IsEnabled state stabilises.
        time.sleep(0.4)

        # Re-resolve input + send button after dialog closed.
        from .text_sender import _get_input as _ti, _get_send_btn as _tsb
        edit = _ti(chat_state, chat)
        btn = _tsb(chat_state, chat)
        if edit is None or btn is None:
            raise RuntimeError("send button or input vanished after dialog closed")

        def btn_enabled() -> Optional[bool]:
            try:
                return bool(btn.IsEnabled)
            except Exception:
                return None

        strategy = _fire_send_attached(sub_hwnd, edit, btn, btn_enabled)
        if not strategy:
            # Last resort: blind invoke. Dialog closed cleanly so the
            # composition area has the attachment regardless of whether
            # IsEnabled is a reliable signal on this build.
            try:
                ip = btn.GetInvokePattern()
                if ip is not None:
                    ip.Invoke()
                    strategy = "invoke_blind"
            except Exception:
                pass
        if not strategy:
            raise RuntimeError(
                f"send_file({chat!r}, {path!r}): attached via toolbar but send failed"
            )
        log.debug("send_file(%s, %r) ok via toolbar+%s", chat, abspath, strategy)
    finally:
        # Restore the original placement — back to minimized if it was.
        if placement_changed:
            try:
                _w32g.SetWindowPlacement(sub_hwnd, orig_placement)
            except Exception:
                pass


def _wait_new_weixin_window(
    existing: Set[int], exclude: Set[int], timeout_s: float = 12.0,
):
    """Poll for a newly-visible Weixin window that contains an Edit control —
    that's our signal it's the file-open dialog. As soon as ANY new top-level
    Weixin window appears we move it off-screen via SetWindowPos (no
    activate, no z-order change), so the user never sees the dialog flash.
    The dialog continues to function normally at coordinates (-30000, -30000)
    — UIA + PostMessage operations don't depend on window position."""
    import ctypes
    import uiautomation as auto
    import win32gui

    SWP_NOACTIVATE = 0x0010
    SWP_NOZORDER = 0x0004
    SWP_NOSIZE = 0x0001
    SWP_FLAGS = SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOSIZE
    SetWindowPos = ctypes.windll.user32.SetWindowPos

    examined: Set[int] = set()
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        now = weixin_top_level_hwnds()
        candidates = sorted((now - existing) - exclude, reverse=True)
        for hwnd in candidates:
            if hwnd in examined:
                continue
            examined.add(hwnd)
            # Hide ASAP — including SoPY_Status splash and any other
            # transient window so the user perceives no flicker.
            try:
                SetWindowPos(hwnd, 0, -30000, -30000, 0, 0, SWP_FLAGS)
            except Exception:
                pass
            try:
                cls = win32gui.GetClassName(hwnd)
                title = win32gui.GetWindowText(hwnd)
            except Exception:
                cls, title = "", ""
            if cls in ("SoPY_Status", "IME", "tooltips_class32",
                       "Internet Explorer_Hidden", "OleMainThreadWndClass"):
                log.debug("send_file: hid non-dialog window cls=%r", cls)
                continue
            ctrl = auto.ControlFromHandle(hwnd)
            if ctrl is None:
                continue
            if has_edit_control(ctrl):
                log.debug(
                    "send_file: matched dialog hwnd=%s cls=%r title=%r",
                    hex(hwnd), cls, title[:40],
                )
                return ctrl
            else:
                log.debug(
                    "send_file: window has no Edit, hidden cls=%r title=%r",
                    cls, title[:40],
                )
        time.sleep(0.03)
    return None


def _fill_file_dialog(dlg, abspath: str) -> None:
    """Write *abspath* into the dialog's filename field and submit via IDOK."""
    import uiautomation as auto

    # Collect every Edit + Button in the dialog for diagnosis and to pick
    # the right one(s) by name rather than position.
    edits: List[Any] = []
    buttons: List[Any] = []

    def visit(node, depth: int = 0):
        if depth > 10 or node is None:
            return
        try:
            ctl = node.ControlTypeName
            if ctl == "EditControl":
                edits.append(node)
            elif ctl == "ButtonControl":
                buttons.append(node)
            for k in node.GetChildren():
                visit(k, depth + 1)
        except Exception:
            pass

    visit(dlg)
    log.debug("send_file: dialog has %d edits, %d buttons", len(edits), len(buttons))
    for i, e in enumerate(edits):
        log.debug("  edit[%d]: name=%r aid=%r",
                  i, (e.Name or "")[:50], (e.AutomationId or "")[:30])
    for i, b in enumerate(buttons):
        log.debug("  button[%d]: name=%r aid=%r",
                  i, (b.Name or "")[:50], (b.AutomationId or "")[:30])

    if not edits:
        raise RuntimeError("file dialog has no Edit controls")

    # Prefer the edit whose name mentions "文件名" / "File name"; else last.
    fname_edit = None
    for e in edits:
        nm = e.Name or ""
        if "文件名" in nm or "File name" in nm or "file name" in nm.lower():
            fname_edit = e
            break
    if fname_edit is None:
        fname_edit = edits[-1]

    # Try ValuePattern first; verify; fall back to SendKeys.
    try:
        fname_edit.GetValuePattern().SetValue(abspath)
    except Exception as exc:
        log.debug("send_file: SetValue raised: %s", exc)
    time.sleep(0.2)
    try:
        actual = fname_edit.GetValuePattern().Value or ""
    except Exception:
        actual = ""
    log.debug("send_file: filename edit value after SetValue: %r", actual[:120])

    if actual != abspath:
        try:
            fname_edit.SetFocus()
            time.sleep(0.1)
            # Select-all + delete first so we don't append.
            auto.SendKeys("{Ctrl}a{Delete}", waitTime=0.05)
            # Send the path. {} bracket parsing in uiautomation.SendKeys
            # treats backslashes literally; colon and slash are fine.
            auto.SendKeys(abspath, waitTime=0.02)
            time.sleep(0.25)
            actual = ""
            try:
                actual = fname_edit.GetValuePattern().Value or ""
            except Exception:
                pass
            log.debug("send_file: filename edit after SendKeys: %r", actual[:120])
        except Exception as exc:
            log.debug("send_file: SendKeys fallback raised: %s", exc)

    # The Win32 file dialog's Open button is a Split-button. UIA exposes
    # only the dropdown halves (aid='DropDown'), not a clean IDOK
    # ButtonControl, so clicking what UIA calls 'Open' just opens the
    # dropdown menu — the dialog never submits. The reliable approach is
    # the Win32 protocol: PostMessage(WM_COMMAND, IDOK) to the dialog
    # hwnd; the OK handler reads the filename edit and dismisses.
    import ctypes
    WM_COMMAND = 0x0111
    IDOK = 1
    dlg_hwnd = getattr(dlg, "NativeWindowHandle", 0) or 0
    if not dlg_hwnd:
        raise RuntimeError("dialog has no NativeWindowHandle for IDOK")

    log.debug("send_file: posting WM_COMMAND IDOK to dialog hwnd=%s", hex(dlg_hwnd))
    ok = ctypes.windll.user32.PostMessageW(dlg_hwnd, WM_COMMAND, IDOK, 0)
    if not ok:
        log.debug("send_file: PostMessage WM_COMMAND failed; trying Click fallback")
        # Fallback: click any button named 打开/Open (may open dropdown).
        for b in buttons:
            nm = (b.Name or "").strip()
            if nm in ("打开(O)", "打开", "Open", "Open(O)") or "打开" in nm:
                try:
                    b.Click(simulateMove=False, waitTime=0)
                    return
                except Exception:
                    pass
        raise RuntimeError("could not submit file dialog (IDOK)")


def _wait_dialog_closed(dlg, timeout_s: float = 6.0) -> None:
    import win32gui
    hwnd = getattr(dlg, "NativeWindowHandle", 0) or 0
    if not hwnd:
        return
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if not win32gui.IsWindow(hwnd):
            return
        time.sleep(0.15)
    raise RuntimeError("file dialog did not close after Open")


def _dismiss_file_dialog(dlg) -> None:
    """Best-effort: click Cancel/取消 to close a stuck dialog so the chat
    doesn't end up with a dangling modal blocking the input."""
    cancel = find_by(
        dlg,
        lambda c: c.ControlTypeName == "ButtonControl" and (
            "取消" in (c.Name or "") or "Cancel" in (c.Name or "")
        ),
    )
    if cancel is None:
        return
    try:
        ip = cancel.GetInvokePattern()
        if ip is not None:
            ip.Invoke()
    except Exception:
        pass


def _fire_send_attached(
    hwnd: int, edit, btn, btn_enabled: Callable[[], Optional[bool]],
) -> Optional[str]:
    """Send when something is attached. Success = button flips back from
    enabled to disabled (composition area emptied)."""
    import win32con
    import win32gui

    def cleared() -> bool:
        return btn_enabled() is False

    try:
        ip = btn.GetInvokePattern()
        if ip is not None:
            ip.Invoke()
            time.sleep(0.4)
            if cleared():
                return "invoke"
    except Exception:
        log.debug("invoke strategy raised", exc_info=True)

    try:
        try:
            edit.SetFocus()
        except Exception:
            pass
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(0.4)
        if cleared():
            return "post_enter"
    except Exception:
        log.debug("post_enter strategy raised", exc_info=True)

    try:
        was_iconic = bool(win32gui.IsIconic(hwnd))
        if was_iconic:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
            time.sleep(0.12)
        btn.Click(simulateMove=False, waitTime=0)
        time.sleep(0.4)
        sent = cleared()
        if was_iconic:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINNOACTIVE)
        if sent:
            return "restore_click" if was_iconic else "click"
    except Exception:
        log.debug("restore_click strategy raised", exc_info=True)

    return None
