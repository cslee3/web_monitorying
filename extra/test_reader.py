# -*- coding: utf-8 -*-
"""
test_reader.py — dashboard_reader 모니터 동작 확인용 테스트 스크립트
monitor 스레드를 직접 띄워 초기 데이터 로드 완료 여부를 최대 60초 동안 폴링.
serve.py 없이 dashboard_reader.py 단독 동작 테스트 시 사용.
"""

import sys, time, threading

# 현재 스크립트 위치를 sys.path에 추가 (모듈 임포트 경로 보장)
sys.path.insert(0, r'D:\OneDrive\server_hana')
import dashboard_reader as dr

print("monitor thread starting...")
t = threading.Thread(target=dr.monitor.run, daemon=True)
t.start()

# 최대 60초 동안 1초 간격으로 초기 데이터 로드 완료 여부 확인
for i in range(60):
    time.sleep(1)
    data = dr.monitor.get_full()
    if data:
        print(f"OK — rows={len(data['rows'])}, updated={data['updated']}")
        break
    print(f"  waiting {i+1}s...")
else:
    print("TIMEOUT — 60초 내에 데이터 로드 실패")
