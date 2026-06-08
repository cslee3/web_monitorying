# -*- coding: utf-8 -*-
"""
logger.py — 파일 기반 비동기 로거

포맷: [INFO|HH:MM:SS.ffffff] message
파일: log/web_monitor.log

사용:
    from logger import get_logger
    log = get_logger(__name__)
    log.info("메시지")
    log.error("에러")

quiet 모드:
    setup(quiet=True)  → WARNING 이상만 기록 (serve.py 'quiet' arg로 제어)

로그 ON/OFF:
    toggle_log()       → 파일 기록 완전 중단/재개 ('l' 키)
"""

import logging
import logging.handlers
import queue
import re
import shutil
import datetime as _dt
from datetime import datetime
from pathlib import Path

_ANSI_RE = re.compile(r'\x1b\[[0-9;]*m')

try:
    _BASE = Path(__file__).parent
except NameError:
    _BASE = Path.cwd()

LOG_DIR     = _BASE / "log"
ARCHIVE_DIR = LOG_DIR / "archive"
LOG_PATH    = LOG_DIR / "web_monitor.log"
LOG_DIR.mkdir(exist_ok=True)
ARCHIVE_DIR.mkdir(exist_ok=True)


_GREEN  = "\x1b[32m"
_YELLOW = "\x1b[33m"
_RED    = "\x1b[31m"
_RESET  = "\x1b[0m"

# 로그 모드: 1=VERBOSE(전체), 2=QUIET(WARN+ERR), 3=OFF
_log_mode = 1   # 기본 VERBOSE
LOG_MODE_NAMES = {1: "VERBOSE", 2: "QUIET", 3: "OFF"}


class _Formatter(logging.Formatter):
    _TAG = {
        logging.DEBUG:    ("INFO", _GREEN),
        logging.INFO:     ("INFO", _GREEN),
        logging.WARNING:  ("WARN", _YELLOW),
        logging.ERROR:    ("ERRO", _RED),
        logging.CRITICAL: ("ERRO", _RED),
    }

    def format(self, record):
        tag, color = self._TAG.get(record.levelno, ("INFO", _GREEN))
        now = datetime.now()
        ts  = now.strftime("%H:%M:%S") + f".{now.microsecond:06d}"
        return f"{color}[{tag}|{ts}] {record.getMessage()}{_RESET}"


_log_queue = queue.Queue()
_listener  = None
_queue_handler = None   # root에 붙인 QueueHandler 참조


class _ModeFilter(logging.Filter):
    """모드별 레코드 필터 — 핸들러는 root에 유지해 lastResort 방지."""
    def filter(self, record):
        if _log_mode == 3:   # OFF
            return False
        if _log_mode == 2:   # QUIET: WARN 이상만
            return record.levelno >= logging.WARNING
        return True          # VERBOSE: 전부


_mode_filter = _ModeFilter()


def setup():
    """로거 초기화. serve.py start() 또는 standalone 실행 시 자동 호출됨."""
    global _listener, _queue_handler
    if _listener is not None:
        return
    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setFormatter(_Formatter())
    _listener = logging.handlers.QueueListener(_log_queue, fh, respect_handler_level=True)
    _listener.start()

    root = logging.getLogger()
    # conda/Flask 등 환경에서 자동 추가된 StreamHandler 포함 기존 핸들러 전부 제거
    for h in root.handlers[:]:
        root.removeHandler(h)
    root.setLevel(logging.DEBUG)   # 필터가 모드별 차단 담당

    _queue_handler = logging.handlers.QueueHandler(_log_queue)
    _queue_handler.addFilter(_mode_filter)
    root.addHandler(_queue_handler)

    # mysql.connector 내부 verbose 로그 항상 억제
    for name in ("mysql.connector", "mysql.connector.connection", "mysql.connector.network"):
        logging.getLogger(name).setLevel(logging.WARNING)


def cycle_log_mode() -> int:
    """로그 모드 순환: VERBOSE(1) → QUIET(2) → OFF(3) → VERBOSE(1) ...
    반환: 새 모드 번호.
    """
    global _log_mode
    _log_mode = (_log_mode % 3) + 1
    return _log_mode


def log_mode() -> int:
    return _log_mode


def log_mode_name() -> str:
    return LOG_MODE_NAMES[_log_mode]


def archive_log(date=None):
    """현재 로그를 log/archive/에 날짜 접미사로 복사하고 현재 파일 초기화."""
    global _listener, _queue_handler
    if date is None:
        date = _dt.date.today()
    base = f"web_monitor_{date.strftime('%Y%m%d')}"
    dest = ARCHIVE_DIR / f"{base}.log"
    if dest.exists():
        n = 1
        while (ARCHIVE_DIR / f"{base}_{n}.log").exists():
            n += 1
        dest = ARCHIVE_DIR / f"{base}_{n}.log"

    # 리스너 중지 + 핸들러 제거
    if _listener:
        _listener.stop()
        _listener = None
    root = logging.getLogger()
    for h in root.handlers[:]:
        if isinstance(h, logging.handlers.QueueHandler):
            root.removeHandler(h)
    _queue_handler = None

    # ANSI 코드 제거 후 아카이브 저장, 현재 로그 초기화
    if LOG_PATH.exists() and LOG_PATH.stat().st_size > 0:
        text = LOG_PATH.read_text(encoding="utf-8", errors="replace")
        dest.write_text(_ANSI_RE.sub("", text), encoding="utf-8")
        LOG_PATH.write_text("", encoding="utf-8")

    # 재시작 (_log_mode 상태는 필터가 유지)
    setup()


def get_logger(name: str) -> logging.Logger:
    if _listener is None:
        setup()
    return logging.getLogger(name)
