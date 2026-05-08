"""WeChat 4.x provider via raw UIA on popped-out chat sub-windows.

Public surface preserved: external callers should import only
`WeChatProvider` and `IncomingMessage` from `src.wechat_provider`.
"""
from ._types import IncomingMessage
from .provider import WeChatProvider

__all__ = ["WeChatProvider", "IncomingMessage"]
