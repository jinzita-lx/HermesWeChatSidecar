"""Generic UIA tree walkers — no WeChat-specific knowledge."""
from __future__ import annotations

from typing import Callable, Optional


def find_by_aid(node, target_aid: str, depth: int = 12):
    """DFS search for a UIA control whose AutomationId equals *target_aid*."""
    if node is None or depth <= 0:
        return None
    if (node.AutomationId or "") == target_aid:
        return node
    try:
        for child in node.GetChildren():
            r = find_by_aid(child, target_aid, depth - 1)
            if r is not None:
                return r
    except Exception:
        pass
    return None


def find_by(node, predicate: Callable[[object], bool], depth: int = 12):
    """DFS search for the first UIA control where *predicate(ctrl)* is True."""
    if node is None or depth <= 0:
        return None
    try:
        if predicate(node):
            return node
    except Exception:
        pass
    try:
        for child in node.GetChildren():
            r = find_by(child, predicate, depth - 1)
            if r is not None:
                return r
    except Exception:
        pass
    return None


def has_edit_control(node, depth: int = 0, max_depth: int = 10) -> bool:
    """True iff *node* or any descendant within *max_depth* is an EditControl."""
    if node is None or depth > max_depth:
        return False
    try:
        if node.ControlTypeName == "EditControl":
            return True
        for k in node.GetChildren():
            if has_edit_control(k, depth + 1, max_depth):
                return True
    except Exception:
        pass
    return False
