# -*- coding: utf-8 -*-
"""
db_manager.py — MySQL 프로세스 시작/종료 + 연결 헬퍼

사용법:
    from db_manager import DBManager
    db = DBManager()

    db.start()           # mysqld 기동 (이미 실행 중이면 스킵)
    db.stop()            # mysqld 종료
    db.status()          # 실행 중 여부 출력
    conn = db.connect()  # mysql.connector 연결 반환
"""

import subprocess
import time
import os
import signal

try:
    import mysql.connector
    HAS_CONNECTOR = True
except ImportError:
    HAS_CONNECTOR = False

# ── 설정 ──────────────────────────────────────────────────────────────────────
# IP 10.115.31.99 (DB 서버 PC): D:\mysql 직접 사용
# 그 외 PC: D:\OneDrive\mysql (OneDrive 공유 경로)
import socket as _socket
_LOCAL_IP = _socket.gethostbyname(_socket.gethostname())
MYSQL_BIN  = r"D:\mysql\bin" if _LOCAL_IP == "10.115.31.99" else r"D:\OneDrive\mysql\bin"
MYSQLD      = os.path.join(MYSQL_BIN, "mysqld.exe")
MYSQLADMIN  = os.path.join(MYSQL_BIN, "mysqladmin.exe")

DB_HOST     = "127.0.0.1"   # 로컬 기동 시
DB_PORT     = 3306
DB_USER     = "root"
DB_PASSWORD = ""
DB_NAME     = "market_making"

STARTUP_TIMEOUT = 15   # 초 — 기동 대기 최대 시간
# ─────────────────────────────────────────────────────────────────────────────


class DBManager:
    def __init__(self):
        self._proc = None   # Popen 객체 (이 인스턴스가 직접 띄운 경우)

    # ── 상태 확인 ─────────────────────────────────────────────────────────────
    def is_running(self) -> bool:
        """mysqld 프로세스가 실행 중인지 확인 (포트 ping 방식)."""
        if not HAS_CONNECTOR:
            return self._ping_tasklist()
        try:
            c = mysql.connector.connect(
                host=DB_HOST, port=DB_PORT,
                user=DB_USER, password=DB_PASSWORD,
                connection_timeout=2,
            )
            c.close()
            return True
        except Exception:
            return False

    def _ping_tasklist(self) -> bool:
        """mysql.connector 없을 때 tasklist로 프로세스 존재 확인."""
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq mysqld.exe", "/NH"],
            capture_output=True, text=True
        )
        return "mysqld.exe" in result.stdout

    def status(self):
        if self.is_running():
            print(f"[DB] MySQL 실행 중  ({DB_HOST}:{DB_PORT})")
        else:
            print(f"[DB] MySQL 정지 상태")

    # ── 시작 ──────────────────────────────────────────────────────────────────
    def start(self, wait: bool = True) -> bool:
        """
        mysqld 기동.
        wait=True  → 실제 연결 가능해질 때까지 대기 (최대 STARTUP_TIMEOUT초)
        반환값: 성공 여부
        """
        if self.is_running():
            print("[DB] 이미 실행 중 — 스킵")
            return True

        if not os.path.exists(MYSQLD):
            print(f"[DB] mysqld.exe 없음: {MYSQLD}")
            return False

        print(f"[DB] MySQL 기동 중... ({MYSQLD})")
        # CREATE_NO_WINDOW: 콘솔 창 안 뜸
        self._proc = subprocess.Popen(
            [MYSQLD, "--console"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

        if not wait:
            print("[DB] 기동 요청 완료 (wait=False)")
            return True

        # 연결 가능해질 때까지 폴링
        deadline = time.time() + STARTUP_TIMEOUT
        while time.time() < deadline:
            if self.is_running():
                print("[DB] MySQL 기동 완료")
                return True
            time.sleep(0.5)

        print(f"[DB] 기동 타임아웃 ({STARTUP_TIMEOUT}초)")
        return False

    # ── 종료 ──────────────────────────────────────────────────────────────────
    def stop(self, force: bool = False):
        """
        MySQL 정상 종료.
        force=True → 정상 종료 실패 시 프로세스 강제 kill
        """
        if not self.is_running():
            print("[DB] 이미 정지 상태")
            return

        # 1순위: mysqladmin shutdown (가장 안전)
        if os.path.exists(MYSQLADMIN):
            print("[DB] mysqladmin shutdown 요청...")
            result = subprocess.run(
                [MYSQLADMIN,
                 "-h", DB_HOST,
                 f"--port={DB_PORT}",
                 "-u", DB_USER,
                 f"--password={DB_PASSWORD}",
                 "shutdown"],
                capture_output=True, text=True, timeout=10,
            )
            # 종료 확인 대기
            for _ in range(20):
                if not self.is_running():
                    print("[DB] MySQL 정상 종료 완료")
                    self._proc = None
                    return
                time.sleep(0.5)

        # 2순위: 직접 띄운 프로세스면 terminate
        if self._proc and self._proc.poll() is None:
            print("[DB] 프로세스 terminate...")
            self._proc.terminate()
            try:
                self._proc.wait(timeout=8)
            except subprocess.TimeoutExpired:
                if force:
                    print("[DB] 강제 kill")
                    self._proc.kill()
            self._proc = None
            print("[DB] 종료 완료")
            return

        # 3순위: taskkill (외부에서 띄워진 mysqld)
        if force:
            print("[DB] taskkill /F 로 강제 종료...")
            subprocess.run(["taskkill", "/F", "/IM", "mysqld.exe"],
                           capture_output=True)
            print("[DB] 강제 종료 완료")
        else:
            print("[DB] 종료 실패 — force=True 로 재시도 가능")

    # ── 연결 반환 ─────────────────────────────────────────────────────────────
    def connect(self, database: str = DB_NAME) -> "mysql.connector.connection":
        """
        mysql.connector 연결 객체 반환.
        MySQL이 꺼져 있으면 자동으로 start() 먼저 실행.

        사용 예:
            conn = db.connect()
            cur  = conn.cursor(dictionary=True)
            cur.execute("SELECT ...")
            rows = cur.fetchall()
            cur.close()
            conn.close()
        """
        if not HAS_CONNECTOR:
            raise ImportError("mysql-connector-python 설치 필요")

        if not self.is_running():
            ok = self.start()
            if not ok:
                raise RuntimeError("MySQL 기동 실패")

        return mysql.connector.connect(
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASSWORD,
            database=database,
            charset="utf8mb4",
        )

    # ── context manager 지원 ─────────────────────────────────────────────────
    def __enter__(self):
        self.start()
        return self

    def __exit__(self, *_):
        self.stop()


# ── 단독 실행: 켜고 유지, 종료 시 DB도 끔 ────────────────────────────────────
if __name__ == "__main__":
    db = DBManager()

    ok = db.start()
    if not ok:
        exit(1)

    print("[DB] 실행 중... 종료하려면 Ctrl+C")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass

    db.stop()
