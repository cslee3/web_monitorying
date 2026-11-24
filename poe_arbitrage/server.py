"""
PoE1 화폐 3자간 차익거래 뷰어 - 로컬 프록시 서버

poe.ninja는 브라우저에서 직접 fetch 시 CORS로 막혀 있어서,
이 서버가 대신 poe.ninja API를 호출해 프론트엔드(index.html)에 넘겨준다.
표준 라이브러리만 사용 (Flask/requests 불필요) -> Python 3만 있으면 바로 실행 가능.

실행:
    python server.py
    (기본 포트 8787, http://localhost:8787 접속)
"""

import json
import time
import threading
import urllib.request
import urllib.parse
import urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

PORT = 8787
REALM = "poe1"  # poe.ninja realm segment
CACHE_TTL_SEC = 15  # 서버 측 캐시 (여러 탭/수동새로고침이 poe.ninja를 과도하게 두드리지 않도록)
REQUEST_TIMEOUT_SEC = 10

CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".svg": "image/svg+xml",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

_cache_lock = threading.Lock()
_cache = {"ts": 0.0, "data": None}


def _fetch_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT_SEC) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _pick_current_league() -> str:
    """economyLeagues 목록의 첫 항목 = poe.ninja가 보여주는 현재 시즌 리그.
    새 리그가 나오면 poe.ninja가 알아서 순서를 바꿔주므로 코드 수정 없이 자동 반영됨."""
    state = _fetch_json(f"https://poe.ninja/{REALM}/api/data/index-state")
    leagues = state.get("economyLeagues") or []
    for lg in leagues:
        if not lg.get("isPrivate"):
            return lg["name"]
    if leagues:
        return leagues[0]["name"]
    raise RuntimeError("economyLeagues 비어있음")


def _fetch_prices() -> dict:
    league = _pick_current_league()
    q = urllib.parse.urlencode({"league": league, "type": "Currency"})
    overview = _fetch_json(f"https://poe.ninja/{REALM}/api/economy/exchange/current/overview?{q}")

    prices = {}
    for line in overview.get("lines", []):
        cid = line.get("id")
        val = line.get("primaryValue")
        if cid is not None and val is not None:
            prices[cid] = val
    prices.setdefault("chaos", 1.0)

    return {
        "league": league,
        "updatedAt": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        "prices": prices,
    }


def get_prices_cached() -> dict:
    with _cache_lock:
        now = time.time()
        if _cache["data"] is not None and (now - _cache["ts"]) < CACHE_TTL_SEC:
            return _cache["data"]
    data = _fetch_prices()
    with _cache_lock:
        _cache["data"] = data
        _cache["ts"] = time.time()
    return data


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # 콘솔 조용히

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.urlsplit(self.path).path

        if path == "/api/prices":
            try:
                data = get_prices_cached()
                self._send_json(data)
            except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, RuntimeError) as e:
                self._send_json({"error": f"poe.ninja 조회 실패: {e}"}, status=502)
            return

        if path == "/" or path == "":
            file_path = BASE_DIR / "index.html"
        else:
            # 정적 파일 서빙 (디렉터리 탈출 방지)
            rel = path.lstrip("/")
            file_path = (BASE_DIR / rel).resolve()
            if BASE_DIR not in file_path.parents and file_path != BASE_DIR:
                self.send_response(403)
                self.end_headers()
                return

        if not file_path.exists() or not file_path.is_file():
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Not Found")
            return

        content_type = CONTENT_TYPES.get(file_path.suffix.lower(), "application/octet-stream")
        body = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"PoE 차익거래 서버 실행 중: http://localhost:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
