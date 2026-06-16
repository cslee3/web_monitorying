# -*- coding: utf-8 -*-
"""
serve.py — Flask 웹 서버 진입점

엔드포인트:
  GET /                  index.html 서빙
  GET /api/dashboard     SSF MM/MM2 전체 데이터 JSON (초기 로드)
  GET /api/stream        SSF MM/MM2 SSE 스트림
  GET /api/sso           SSO MM 전체 데이터 JSON (초기 로드)
  GET /api/sso/stream    SSO MM SSE 스트림
  GET /<path>            정적 파일 서빙

실행: python serve.py  (포트 8080)
"""

from flask import Flask, send_from_directory, jsonify, Response
import os, sys, threading, msvcrt
from queue import Empty
import reader_ssf
import reader_sso
import logger

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TMPL_DIR = os.path.join(BASE_DIR, "templates")


def _sse_response(monitor):
    """SSE 응답 생성 헬퍼. monitor에 클라이언트 큐를 등록하고 스트림 반환."""
    import logger as _logger_mod
    _slog = _logger_mod.get_logger("sse")
    q = monitor.add_client()
    mon_name = type(monitor).__name__

    def generate():
        _slog.info(f"{mon_name} SSE connected (total={len(monitor._clients)})")
        try:
            running = str(monitor.is_running()).lower()
            yield f"data: {{\"type\":\"connected\",\"running\":{running}}}\n\n"
            while True:
                try:
                    data = q.get(timeout=20)
                    yield f"data: {data}\n\n"
                except Empty:
                    yield ": keepalive\n\n"
        except Exception as e:
            _slog.error(f"{mon_name} SSE generator error: {e}")
        finally:
            monitor.remove_client(q)
            _slog.info(f"{mon_name} SSE disconnected (total={len(monitor._clients)})")

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/")
def index():
    resp = send_from_directory(TMPL_DIR, "index.html")
    resp.headers["Cache-Control"] = "no-store"
    return resp


# ── SSF MM / SSF MM2 (DashBoard 탭) ──────────────────────────────────────

@app.route("/api/dashboard")
def api_ssf_full():
    data = reader_ssf.monitor.get_full()
    if data is None:
        return jsonify({"error": "not ready"}), 503
    return jsonify(data)


@app.route("/api/stream")
def api_ssf_stream():
    return _sse_response(reader_ssf.monitor)


# ── SSO MM (Option_DashBoard 탭) ──────────────────────────────────────────

@app.route("/api/sso")
def api_sso_full():
    data = reader_sso.monitor.get_full()
    if data is None:
        return jsonify({"error": "not ready"}), 503
    return jsonify(data)


@app.route("/api/sso/stream")
def api_sso_stream():
    return _sse_response(reader_sso.monitor)


# ── 모니터 시작/중지 ──────────────────────────────────────────────────────────

_MONITORS = {"ssf": lambda: reader_ssf.monitor, "sso": lambda: reader_sso.monitor}

@app.route("/api/monitor/<mon>", methods=["GET"])
def api_monitor_status(mon):
    m = _MONITORS.get(mon)
    if m is None:
        return jsonify({"error": "not found"}), 404
    return jsonify({"running": m().is_running()})


@app.route("/api/monitor/<mon>", methods=["POST"])
def api_monitor_toggle(mon):
    m = _MONITORS.get(mon)
    if m is None:
        return jsonify({"error": "not found"}), 404
    monitor = m()
    if monitor.is_running():
        monitor.pause()
        logger.get_logger(__name__).warning(f"{mon} monitor PAUSED")
    else:
        monitor.resume()
        logger.get_logger(__name__).warning(f"{mon} monitor RESUMED")
    return jsonify({"running": monitor.is_running()})


# ── DB 저장 토글 ─────────────────────────────────────────────────────────────

@app.route("/api/db-save", methods=["GET"])
def api_db_save_status():
    return jsonify({"db_save": reader_ssf.monitor.db_save_enabled})


@app.route("/api/db-save", methods=["POST"])
def api_db_save_toggle():
    enabled = reader_ssf.monitor.toggle_db_save()
    logger.get_logger(__name__).warning(f"db_save {'ON' if enabled else 'OFF'}")
    return jsonify({"db_save": enabled})


# ── 정적 파일 ─────────────────────────────────────────────────────────────

ARCHIVE_DIR = os.path.join(BASE_DIR, "archive")

@app.route("/api/archive/pnl/latest")
def api_pnl_latest():
    """archive/ 폴더에서 가장 최신 PnL 파일의 날짜(YYYYMMDD) 반환."""
    import re
    pattern = re.compile(r'(?i)^pnl_report_(\d{8})\.html$')
    latest = None
    try:
        for fname in os.listdir(ARCHIVE_DIR):
            m = pattern.match(fname)
            if m:
                d = m.group(1)
                if latest is None or d > latest:
                    latest = d
    except FileNotFoundError:
        pass
    return jsonify({"date": latest})


@app.route("/archive/<path:filename>")
def archive_files(filename):
    resp = send_from_directory(ARCHIVE_DIR, filename)
    resp.headers["Cache-Control"] = "no-store"
    return resp



@app.route("/<path:filename>")
def static_files(filename):
    if filename.endswith(".html"):
        resp = send_from_directory(TMPL_DIR, filename)
        resp.headers["Cache-Control"] = "no-store"
    else:
        resp = send_from_directory(BASE_DIR, filename)
    return resp


# ── 터미널 상태 표시 ──────────────────────────────────────────────────────

def _enable_ansi() -> bool:
    """ANSI 이스케이프 코드 지원 여부 확인 및 활성화.
    IPython 환경에서는 sys.stdout이 래퍼 객체이므로 sys.__stdout__ 기준으로 판단.
    """
    raw = sys.__stdout__
    if raw is None or not raw.isatty():
        return False
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)
        mode   = ctypes.c_ulong()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass
    return True


# 터미널 출력은 IPython 래퍼를 우회해 원본 콘솔로 직접 씀
_COUT = sys.__stdout__ or sys.stdout

_ANSI  = _enable_ansi()
_GREEN = "\x1b[32m" if _ANSI else ""
_RED   = "\x1b[31m" if _ANSI else ""
_RESET = "\x1b[0m"  if _ANSI else ""

# 전체 visible 폭 = 54자 (prefix 4 + content 50)
_SEP     = "    " + "#" * 50
_CMD_SEP = "    " + "#" * 50
#           9           10    2   10    2   4    2   11       → 48
_CMD1    = "    #  DB  :  b=저장모드  p=핑  a=수동저장  f=덤프   #"
#           9           10    2   10    17                    → 48
_CMD2    = "    #  LOG :  l=로그모드  x=로그저장                 #"
#           9           12    2    8    2    6    2    6   1  → 48
_CMD3    = "    #  ETC :  t=모니터모드  d=스냅샷  c=화면  q=종료 #"
_CMD_ROW = 17  # CMD> 가 출력되는 행 번호 (1-based)

_cmd_buf      = b""
_kbd_disabled = False
_prev_status  = None


def _db_connected(ssf) -> bool:
    """네트워크 I/O 없이 모니터 스레드가 업데이트한 플래그만 읽음."""
    return getattr(ssf, '_db_ok', False)


def _live(ok):
    # visible 8자: "  LIVE  " / "  DEAD  "
    return f"{_GREEN}  LIVE  {_RESET}" if ok else f"{_RED}  DEAD  {_RESET}"

def _onoff(ok):
    # visible 8자: "   ON   " / "  OFF   "
    return f"{_GREEN}   ON   {_RESET}" if ok else f"{_RED}  OFF   {_RESET}"

def _logmode(mode):
    # visible 8자
    if mode == "VERBOSE":  return f"{_GREEN} VERB   {_RESET}"
    if mode == "QUIET":    return f"{_GREEN} QUIET  {_RESET}"
    return                        f"{_RED}  OFF   {_RESET}"


def _collect():
    ssf = reader_ssf.monitor
    sso = reader_sso.monitor
    with ssf._clients_lock:
        ssf_cli = len(ssf._clients)
    with sso._clients_lock:
        sso_cli = len(sso._clients)
    import logger as _logger_mod
    return {
        'ssf_run': ssf.is_running(),
        'ssf_xl':  getattr(ssf, '_read_ok', False),
        'sso_run': sso.is_running(),
        'sso_xl':  getattr(sso, '_read_ok', False),
        'db_save': ssf.db_save_enabled,
        'db_conn': _db_connected(ssf),
        'ssf_cli': ssf_cli,
        'sso_cli': sso_cli,
        'log_mode': _logger_mod.log_mode_name(),
    }


def _build_rows(st, t_str):
    # 각 행의 visible 폭 = 46자
    # 구조: "    #  {label:<13}#" (22) + status_8 + "#  {info:<10}  #" (16) = 46
    # ─ label 열 visible: "  " + 13chars + "#" = 16  →  prefix "    #" + 16 = 22
    # ─ status 열 visible: 8 chars (LIVE/DEAD/ON/OFF)
    # ─ info 열 visible: "#  " + 10chars + "  #" = 16
    # 합계: 22 + 8 + 16 = 46 ✓
    xl_ssf  = "Excel: OK         " if st['ssf_xl'] else "Excel:WAIT        "
    xl_sso  = "Excel: OK         " if st['sso_xl'] else "Excel:WAIT        "
    db_info = "Connected         " if st['db_conn'] else "No Conn           "
    mode = st['log_mode']   # "VERBOSE" / "QUIET" / "OFF"
    log_info = f"{mode:<18}"
    # 전체 visible 폭 = 54:
    #   label열  22chars + status열 8chars + info열 24chars = 54
    return {
        'title': f"    #  WEB MONITOR  |  {t_str}                      #",
        'ssf':   "    #  SSF Monitor   #" + _live(st['ssf_run'])  + f"#  {xl_ssf}  #",
        'sso':   "    #  SSO Monitor   #" + _live(st['sso_run'])  + f"#  {xl_sso}  #",
        'db':    "    #  DB  Save      #" + _onoff(st['db_save']) + f"#  {db_info}  #",
        'cli':   f"    #  SSE Clients   #  SSF:{st['ssf_cli']:2d}  SSO:{st['sso_cli']:2d}               #",
        'log':   "    #  Log File      #" + _logmode(mode) + f"#  {log_info}  #",
    }


def _print_status():
    global _prev_status
    import datetime
    t_str = datetime.datetime.now().strftime("%H:%M:%S")
    st  = _collect()
    row = _build_rows(st, t_str)
    out = []
    W   = out.append

    if not _ANSI:
        # ANSI 미지원 환경: cls로 화면 클리어 후 전체 재출력
        os.system('cls')
        W("\n\n")
        W(_SEP           + "\n")
        W(row['title']   + "\n")
        W(_SEP           + "\n")
        W(row['ssf']     + "\n")
        W(row['sso']     + "\n")
        W(row['db']      + "\n")
        W(row['cli']     + "\n")
        W(row['log']     + "\n")
        W(_CMD_SEP       + "\n")
        W(_CMD1          + "\n")
        W(_CMD2          + "\n")
        W(_CMD3          + "\n")
        W(_CMD_SEP       + "\n")
        W("\n")
        W(f"    CMD> {_cmd_buf.decode('ascii', 'ignore')}\n")
    elif _prev_status is None:
        W("\033[2J\033[H")
        W("\n")                    # line  1
        W("\n")                    # line  2
        W(_SEP           + "\n")   # line  3
        W(row['title']   + "\n")   # line  4
        W(_SEP           + "\n")   # line  5
        W(row['ssf']     + "\n")   # line  6
        W(row['sso']     + "\n")   # line  7
        W(row['db']      + "\n")   # line  8
        W(row['cli']     + "\n")   # line  9
        W(row['log']     + "\n")   # line 10
        W(_CMD_SEP       + "\n")   # line 11
        W(_CMD1          + "\n")   # line 12
        W(_CMD2          + "\n")   # line 13
        W(_CMD3          + "\n")   # line 14
        W(_CMD_SEP       + "\n")   # line 15
        W("\n")                    # line 16
        W(f"    CMD> {_cmd_buf.decode('ascii', 'ignore')}\n")   # line 17
    else:
        prev = _prev_status
        W(f"\033[4;1H{row['title']}")
        if st['ssf_run'] != prev['ssf_run'] or st['ssf_xl'] != prev['ssf_xl']:
            W(f"\033[6;1H{row['ssf']}")
        if st['sso_run'] != prev['sso_run'] or st['sso_xl'] != prev['sso_xl']:
            W(f"\033[7;1H{row['sso']}")
        if st['db_save'] != prev['db_save'] or st['db_conn'] != prev['db_conn']:
            W(f"\033[8;1H{row['db']}")
        if st['ssf_cli'] != prev['ssf_cli'] or st['sso_cli'] != prev['sso_cli']:
            W(f"\033[9;1H{row['cli']}")
        if st['log_mode'] != prev.get('log_mode'):
            W(f"\033[10;1H{row['log']}")

    try:
        _COUT.write("".join(out))
        _COUT.flush()
    except Exception:
        pass
    _prev_status = {**st}


def _render_cmdline():
    if not _ANSI:
        return
    buf = _cmd_buf.decode('ascii', 'ignore')
    _COUT.write(f"\033[{_CMD_ROW};1H    CMD> {buf}\033[K")
    _COUT.flush()


def _read_commands():
    global _cmd_buf, _kbd_disabled
    if _kbd_disabled:
        return []
    cmds = []
    changed = False
    try:
        while msvcrt.kbhit():
            ch = msvcrt.getch()
            if ch in (b'\x00', b'\xe0'):
                if msvcrt.kbhit():
                    msvcrt.getch()
                continue
            if ch in (b'\r', b'\n'):
                cmd = _cmd_buf.decode('ascii', 'ignore').strip().lower()
                _cmd_buf = b""
                changed = True
                if cmd:
                    cmds.append(cmd)
                continue
            if ch in (b'\x08', b'\x7f'):
                if _cmd_buf:
                    _cmd_buf = _cmd_buf[:-1]
                    changed = True
                continue
            if 0x20 <= ch[0] < 0x7f:
                _cmd_buf += ch
                changed = True
    except Exception:
        _kbd_disabled = True
    if changed:
        try:
            _render_cmdline()
        except Exception:
            pass
    return cmds


# ── 커맨드 핸들러 ─────────────────────────────────────────────────────────

_log = None   # 모듈 레벨 로거 (start() 이후 초기화)


def _toggle_monitors():
    running = not reader_ssf.monitor.is_running()
    if running:
        reader_ssf.monitor.resume()
        reader_sso.monitor.resume()
        _log.warning("모니터 시작 (SSF + SSO)")
    else:
        reader_ssf.monitor.pause()
        reader_sso.monitor.pause()
        _log.warning("모니터 중지 (SSF + SSO)")


def _toggle_db_save():
    enabled = reader_ssf.monitor.toggle_db_save()
    _log.warning(f"DB 자동저장 {'ON' if enabled else 'OFF'}")


def _manual_snapshot():
    import datetime
    today = datetime.date.today()
    def _do():
        try:
            reader_ssf.monitor._archive_html(today)
            _log.warning(f"수동 스냅샷 저장 완료 ({today.strftime('%Y%m%d')})")
        except Exception as e:
            _log.error(f"수동 스냅샷 실패: {e}")
    threading.Thread(target=_do, daemon=True).start()


def _manual_db_save():
    import time
    m = reader_ssf.monitor
    data = m._full_data
    if data is None:
        _log.warning("수동 DB 저장 실패: 데이터 없음")
        return
    def _do():
        try:
            m._save_to_db(data, time.time())
            _log.warning("수동 DB 저장 완료")
        except Exception as e:
            _log.error(f"수동 DB 저장 실패: {e}")
    threading.Thread(target=_do, daemon=True).start()


def _toggle_log():
    mode = logger.cycle_log_mode()
    name = logger.log_mode_name()
    if mode != 3:   # OFF가 아닐 때만 기록 가능
        _log.warning(f"로그 모드 변경: {name}")


def _archive_log_now():
    import datetime
    try:
        logger.archive_log(datetime.date.today())
        # archive_log 후 새 핸들러로 로그 기록
        _log.warning("로그 수동 아카이브 완료")
    except Exception as e:
        _log.error(f"로그 아카이브 실패: {e}")


def _clear_screen():
    global _prev_status
    _prev_status = None


def _manual_dump():
    import datetime
    today = datetime.date.today()
    m = reader_ssf.monitor
    def _do():
        try:
            _log.warning(f"수동 DB 덤프 시작 ({today.strftime('%Y%m%d')})")
            m._dump_and_cleanup(today)
            _log.warning("수동 DB 덤프 완료 (SQL 저장 + DB 클리어)")
        except Exception as e:
            _log.error(f"수동 DB 덤프 실패: {e}")
    threading.Thread(target=_do, daemon=True).start()


def _ping_db():
    import threading
    threading.Thread(target=reader_ssf.monitor._refresh_db_ok, daemon=True).start()
    _log.info("DB 연결 확인 중...")


def _dispatch(cmd):
    if   cmd == 't': _toggle_monitors()
    elif cmd == 'b': _toggle_db_save()
    elif cmd == 'p': _ping_db()
    elif cmd == 'd': _manual_snapshot()
    elif cmd == 'a': _manual_db_save()
    elif cmd == 'f': _manual_dump()
    elif cmd == 'l': _toggle_log()
    elif cmd == 'x': _archive_log_now()
    elif cmd == 'c':
        _clear_screen()
        if _log: _log.info("화면 새로고침")
    elif cmd == 'q':
        if _log: _log.warning("서버 종료 (q)")
        if _ANSI:
            _COUT.write("\033[?25h")   # 커서 복원
            _COUT.flush()
        os._exit(0)


def _status_loop():
    import time
    _interval = 1.0 if not _ANSI else 0.25
    while True:
        try:
            _print_status()
            time.sleep(_interval)
            for cmd in _read_commands():
                try:
                    _dispatch(cmd)
                except Exception as e:
                    if _log: _log.error(f"cmd '{cmd}' error: {e}")
            _print_status()
        except Exception as e:
            if _log: _log.error(f"status_loop error: {e}")
            time.sleep(1)
            _clear_screen()   # 다음 루프에서 전체 재출력


# ── 서버 시작 ─────────────────────────────────────────────────────────────

def start():
    """
    모니터 백그라운드 스레드 시작 후 Flask 서버 실행.
    Spyder 등 IPython 환경에서도 직접 호출 가능.
    use_reloader=False: Werkzeug 자동 리로더 비활성화 (Spyder 크래시 방지).
    """
    global _log
    import logging, flask.cli
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    flask.cli.show_server_banner = lambda *a, **kw: None
    logger.setup()
    _log = logger.get_logger(__name__)

    if _ANSI:
        _COUT.write("\033[?25l")   # 커서 숨김
        _COUT.flush()

    threading.Thread(target=reader_ssf.monitor.run, daemon=True).start()
    threading.Thread(target=reader_sso.monitor.run, daemon=True).start()
    threading.Thread(target=_status_loop, daemon=True).start()
    app.run(host="0.0.0.0", port=8080, debug=False, threaded=True, use_reloader=False)


if __name__ == "__main__":
    if "nosave" in sys.argv[1:]:
        reader_ssf.monitor.db_save_enabled = False
    start()
