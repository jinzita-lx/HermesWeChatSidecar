from __future__ import annotations

import json
import logging
import threading
from collections import OrderedDict
from pathlib import Path

log = logging.getLogger(__name__)


class SeenIds:
    """LRU set of message ids, persisted to disk so dedup survives restarts."""

    def __init__(self, path: Path, capacity: int = 5000) -> None:
        self._path = path
        self._capacity = capacity
        self._lock = threading.Lock()
        self._items: "OrderedDict[str, None]" = OrderedDict()
        self._load()

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            data = json.loads(self._path.read_text(encoding="utf-8"))
        except Exception as exc:
            log.warning("seen_ids load failed: %s", exc)
            return
        if isinstance(data, list):
            for item in data[-self._capacity:]:
                self._items[str(item)] = None

    def _flush_locked(self) -> None:
        try:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            self._path.write_text(
                json.dumps(list(self._items.keys()), ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception as exc:
            log.warning("seen_ids flush failed: %s", exc)

    def add_if_new(self, key: str) -> bool:
        with self._lock:
            if key in self._items:
                self._items.move_to_end(key)
                return False
            self._items[key] = None
            while len(self._items) > self._capacity:
                self._items.popitem(last=False)
            self._flush_locked()
            return True
