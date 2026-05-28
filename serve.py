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
import os, threading
from queue import Empty
import reader_ssf
import reader_sso

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TMPL_DIR = os.path.join(BASE_DIR, "templates")


def _sse_response(monitor):
    """SSE 응답 생성 헬퍼. monitor에 클라이언트 큐를 등록하고 스트림 반환."""
    q = monitor.add_client()

    def generate():
        try:
            yield "data: {\"type\":\"connected\"}\n\n"
            while True:
                try:
                    data = q.get(timeout=20)
                    yield f"data: {data}\n\n"
                except Empty:
                    yield ": keepalive\n\n"
        finally:
            monitor.remove_client(q)

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


# ── DB 저장 토글 ─────────────────────────────────────────────────────────────

@app.route("/api/db-save", methods=["GET"])
def api_db_save_status():
    return jsonify({"db_save": reader_ssf.monitor.db_save_enabled})


@app.route("/api/db-save", methods=["POST"])
def api_db_save_toggle():
    enabled = reader_ssf.monitor.toggle_db_save()
    print(f"[serve] db_save {'ON' if enabled else 'OFF'}")
    return jsonify({"db_save": enabled})


# ── 정적 파일 ─────────────────────────────────────────────────────────────

@app.route("/<path:filename>")
def static_files(filename):
    if filename.endswith(".html"):
        resp = send_from_directory(TMPL_DIR, filename)
        resp.headers["Cache-Control"] = "no-store"
    else:
        resp = send_from_directory(BASE_DIR, filename)
    return resp


# ── 시작 시 DB 스냅샷 테스트 ────────────────────────────────────────────────

def _dump_test():
    """서버 시작 시 ssf_history 전체를 ssf_history_test.sql로 덤프. DB 상태 확인용."""
    try:
        from db_manager import DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME
        import mysql.connector, os, datetime
        conn = mysql.connector.connect(
            host=DB_HOST, port=DB_PORT,
            user=DB_USER, password=DB_PASSWORD,
            database=DB_NAME, charset="utf8mb4",
            connection_timeout=3,
        )
        cur = conn.cursor()
        cur.execute("SELECT * FROM ssf_history ORDER BY ts")
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
        cur.close()
        conn.close()

        if not rows:
            print("[serve] ssf_history 비어있음 — test dump 생략")
            return

        def _v(v):
            if v is None:          return "NULL"
            if isinstance(v, str): return "'" + v.replace("'", "''") + "'"
            return str(v)

        out = os.path.join(BASE_DIR, "ssf_history_test.sql")
        col_str = ", ".join(cols)
        with open(out, "w", encoding="utf-8") as f:
            f.write(f"-- ssf_history test dump  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}  ({len(rows):,} rows)\n")
            f.write(f"INSERT INTO ssf_history ({col_str}) VALUES\n")
            lines = ["  (" + ", ".join(_v(v) for v in row) + ")" for row in rows]
            f.write(",\n".join(lines) + ";\n")

        first_ts = rows[0][0]
        last_ts  = rows[-1][0]
        print(f"[serve] test dump → ssf_history_test.sql  {len(rows):,}행"
              f"  ({datetime.datetime.fromtimestamp(first_ts):%H:%M:%S}"
              f" ~ {datetime.datetime.fromtimestamp(last_ts):%H:%M:%S})")

    except Exception as e:
        print(f"[serve] test dump 실패 (DB 꺼짐?): {e}")


# ── 서버 시작 ─────────────────────────────────────────────────────────────

def start():
    """
    모니터 백그라운드 스레드 시작 후 Flask 서버 실행.
    Spyder 등 IPython 환경에서도 직접 호출 가능.
    use_reloader=False: Werkzeug 자동 리로더 비활성화 (Spyder 크래시 방지).
    """
    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)  # GET 로그 억제

    threading.Thread(target=reader_ssf.monitor.run, daemon=True).start()
    threading.Thread(target=reader_sso.monitor.run, daemon=True).start()
    save_state = "ON" if reader_ssf.monitor.db_save_enabled else "OFF"
    print(f"[serve] DB save mode: {save_state}")
    _dump_test()
    print("[serve] http://0.0.0.0:8080")
    app.run(host="0.0.0.0", port=8080, debug=False, threaded=True, use_reloader=False)


if __name__ == "__main__":
    import sys
    if "nosave" in sys.argv[1:]:
        reader_ssf.monitor.db_save_enabled = False
        print("[serve] db_save OFF")
    start()
