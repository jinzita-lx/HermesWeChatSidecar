from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional


_BUCKET_MINUTES = 10


class BucketedFileHandler(logging.Handler):
    """Writes records into ``<root>/YYYY-MM-DD/HH/HHMM.log``, where HHMM is
    the start of the current 10-minute bucket (e.g. 0000, 0010, 0020 ...).
    The handler reopens a new file on every bucket boundary.
    """

    def __init__(self, root_dir: Path, bucket_minutes: int = _BUCKET_MINUTES,
                 encoding: str = "utf-8") -> None:
        super().__init__()
        self._root = Path(root_dir)
        self._bucket = bucket_minutes
        self._encoding = encoding
        self._stream = None
        self._current_path: Optional[Path] = None

    def _path_for(self, dt: datetime) -> Path:
        m = (dt.minute // self._bucket) * self._bucket
        hour = f"{dt.hour:02d}"
        return self._root / dt.strftime("%Y-%m-%d") / hour / f"{hour}{m:02d}.log"

    def _open(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self._stream = open(path, "a", encoding=self._encoding)
        self._current_path = path

    def emit(self, record: logging.LogRecord) -> None:
        try:
            target = self._path_for(datetime.fromtimestamp(record.created))
            if target != self._current_path:
                if self._stream is not None:
                    try:
                        self._stream.close()
                    except Exception:
                        pass
                self._open(target)
            self._stream.write(self.format(record) + "\n")
            self._stream.flush()
        except Exception:
            self.handleError(record)

    def close(self) -> None:
        if self._stream is not None:
            try:
                self._stream.close()
            except Exception:
                pass
            self._stream = None
        super().close()


def configure(level: str, log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    fmt = "%(asctime)s %(levelname)-5s [%(name)s] %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"
    formatter = logging.Formatter(fmt, datefmt)

    root = logging.getLogger()
    root.setLevel(getattr(logging, level, logging.INFO))
    for handler in list(root.handlers):
        root.removeHandler(handler)

    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    root.addHandler(stream)

    bucketed = BucketedFileHandler(log_dir)
    bucketed.setFormatter(formatter)
    root.addHandler(bucketed)

    logging.getLogger("websockets").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
