"""WebSocket client to the Linux adapter.

Responsibilities:
  * Maintain a long-lived WS connection with heartbeat.
  * Reconnect on failure with capped exponential backoff + jitter.
  * Dispatch inbound JSON messages to a callback.
  * Provide a thread-safe send() so the wxauto polling thread can push events.
"""
from __future__ import annotations

import asyncio
import json
import logging
import random
import time
from typing import Any, Awaitable, Callable, Dict, Optional

import websockets
from websockets.exceptions import ConnectionClosed

log = logging.getLogger(__name__)

OnMessage = Callable[[Dict[str, Any]], Awaitable[Optional[Dict[str, Any]]]]


class WSClient:
    def __init__(
        self,
        url: str,
        on_message: OnMessage,
        heartbeat_interval: float = 20.0,
        reconnect_min: float = 1.0,
        reconnect_max: float = 30.0,
    ) -> None:
        self._url = url
        self._on_message = on_message
        self._hb = heartbeat_interval
        self._rmin = reconnect_min
        self._rmax = reconnect_max

        self._loop: Optional[asyncio.AbstractEventLoop] = None
        self._ws: Optional[websockets.WebSocketClientProtocol] = None
        self._outbound: Optional[asyncio.Queue] = None
        self._stop = asyncio.Event() if False else None  # set in run()
        self._connected = asyncio.Event() if False else None

    # ---------- public ----------

    async def run(self) -> None:
        self._loop = asyncio.get_running_loop()
        self._outbound = asyncio.Queue()
        self._stop = asyncio.Event()
        self._connected = asyncio.Event()

        backoff = self._rmin
        while not self._stop.is_set():
            try:
                log.info("WS connecting -> %s", self._safe_url())
                async with websockets.connect(
                    self._url,
                    ping_interval=None,   # we send our own application-level pings
                    open_timeout=15,
                    max_size=64 * 1024 * 1024,
                ) as ws:
                    self._ws = ws
                    self._connected.set()
                    backoff = self._rmin
                    log.info("WS connected")
                    await self._session()
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                log.warning("WS session ended: %s", exc)
            finally:
                self._connected.clear()
                self._ws = None

            if self._stop.is_set():
                break

            sleep_for = backoff + random.uniform(0, backoff / 2)
            log.info("WS reconnecting in %.1fs", sleep_for)
            try:
                await asyncio.wait_for(self._stop.wait(), timeout=sleep_for)
                break  # stop signalled during sleep
            except asyncio.TimeoutError:
                pass
            backoff = min(self._rmax, backoff * 2 if backoff > 0 else self._rmin)

    def stop(self) -> None:
        if self._loop and self._stop:
            self._loop.call_soon_threadsafe(self._stop.set)

    def send_threadsafe(self, message: Dict[str, Any]) -> None:
        """Enqueue a message from any thread."""
        if not self._loop or not self._outbound:
            log.warning("WS not started; dropping message %s", message.get("type"))
            return
        self._loop.call_soon_threadsafe(self._outbound.put_nowait, message)

    async def send(self, message: Dict[str, Any]) -> None:
        if not self._outbound:
            return
        await self._outbound.put(message)

    # ---------- internals ----------

    def _safe_url(self) -> str:
        # avoid leaking token in logs
        if "token=" in self._url:
            head, _, tail = self._url.partition("token=")
            redacted = tail.split("&", 1)
            redacted[0] = "***"
            return head + "token=" + "&".join(redacted)
        return self._url

    async def _session(self) -> None:
        recv_task = asyncio.create_task(self._recv_loop(), name="ws-recv")
        send_task = asyncio.create_task(self._send_loop(), name="ws-send")
        hb_task = asyncio.create_task(self._heartbeat_loop(), name="ws-hb")

        done, pending = await asyncio.wait(
            {recv_task, send_task, hb_task},
            return_when=asyncio.FIRST_COMPLETED,
        )
        for t in pending:
            t.cancel()
        for t in pending:
            try:
                await t
            except (asyncio.CancelledError, Exception):
                pass
        # Surface the first finished task's exception (if any) so run() can log + reconnect.
        for t in done:
            exc = t.exception()
            if exc:
                raise exc

    async def _recv_loop(self) -> None:
        assert self._ws is not None
        async for raw in self._ws:
            try:
                msg = json.loads(raw)
            except Exception:
                log.warning("WS recv: non-JSON frame, ignored: %r", raw[:200])
                continue
            if not isinstance(msg, dict):
                log.warning("WS recv: non-object frame, ignored")
                continue
            mtype = msg.get("type")
            log.debug("WS <- %s", mtype)
            if mtype == "ping":
                await self.send({"type": "pong", "ts": time.time()})
                continue
            if mtype == "pong":
                continue
            try:
                reply = await self._on_message(msg)
            except Exception:
                log.exception("on_message handler raised")
                continue
            if reply:
                await self.send(reply)

    async def _send_loop(self) -> None:
        assert self._ws is not None
        assert self._outbound is not None
        while True:
            msg = await self._outbound.get()
            try:
                await self._ws.send(json.dumps(msg, ensure_ascii=False))
                log.debug("WS -> %s", msg.get("type"))
            except ConnectionClosed:
                # Re-queue the message so it is sent after reconnect.
                await self._outbound.put(msg)
                raise
            except Exception:
                log.exception("WS send failed for type=%s", msg.get("type"))

    async def _heartbeat_loop(self) -> None:
        while True:
            await asyncio.sleep(self._hb)
            await self.send({"type": "ping", "ts": time.time()})
