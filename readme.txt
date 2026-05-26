
================================================================================
  server_hana 구성
  - MySQL 8.0.45       : 무설치(ZIP) DB #1
  - (DB #2 추가 예정)
  - web_monitor.py     : FastAPI 엑셀 라이브 모니터링 (port 8000)

  [환경 제약 - 회사 내부망]
  - 관리자 권한 없음  ->  서비스 등록(--install), 환경변수 편집 불가
  - 인터넷 차단       ->  pip 휠 파일 수동 설치
  - 외부 프로세스     ->  회사 관리 mysqld 가 따로 떠 있을 수 있음
================================================================================


* [DB #1] MySQL 8.0.45 무설치(ZIP archive) 설치 및 운영

[설치]

1. 다운로드 (인터넷 가능한 PC에서)
   https://downloads.mysql.com/archives/community/
   - Product Version : 8.0.45
   - Operating System : Windows (x86, 64-bit)
   - ZIP Archive 선택 (mysql-8.0.45-winx64.zip)

2. 압축 해제 위치
   C:\mysql-8.0.45-winx64\

3. my.ini 생성  ->  C:\mysql-8.0.45-winx64\my.ini
   -------------------------------------------------------
   [mysqld]
   basedir   = C:/mysql-8.0.45-winx64
   datadir   = C:/mysql-8.0.45-winx64/data
   port      = 3306
   character-set-server = utf8mb4
   collation-server     = utf8mb4_unicode_ci
   default-storage-engine = InnoDB

   [client]
   port    = 3306
   default-character-set = utf8mb4
   -------------------------------------------------------

4. DB 초기화 (최초 1회, 비밀번호 없이)
   PowerShell 에서:
   > cd C:\mysql-8.0.45-winx64\bin
   > .\mysqld --initialize-insecure --console
   <- 완료 메시지 나오면 종료 (Ctrl+C 또는 창 닫기)

   ※ --initialize-insecure : root 비밀번호 없이 초기화
   ※ data 폴더가 이미 있으면 먼저 삭제 후 실행


[운영 - 서비스 등록 불가로 수동 실행]

MySQL 시작 (이 창은 떠 있는 동안 MySQL 살아있음)
   > cd C:\mysql-8.0.45-winx64\bin
   > .\mysqld --console

MySQL 접속 (새 PowerShell 창)
   > cd C:\mysql-8.0.45-winx64\bin
   > .\mysql -u root

MySQL 종료
   > 시작 창을 Ctrl+C  또는 닫기
   또는:
   > .\mysqladmin -u root shutdown


[다른 PC 복제 방법]

  조건 : 대상 PC 에도 같은 경로에 동일 버전 ZIP 압축 해제 + my.ini 생성 완료
  방법 :
   1. 원본 PC 에서 mysqld 종료
   2. C:\mysql-8.0.45-winx64\data\ 폴더 통째로 복사
   3. 대상 PC 의 같은 경로에 붙여넣기
   4. .\mysqld --console 로 시작  ->  .\mysql -u root 접속

  ※ MySQL 버전이 동일해야 함 (8.0.45)
  ※ mysqld 가 실행 중인 상태에서 data 폴더 복사하면 손상 위험


--------------------------------------------------------------------------------


* [Web Monitor] web_monitor.py  (FastAPI + xlwings)

[사전 준비 - 폐쇄망 수동 설치]

  인터넷 가능한 PC 에서 휠 파일 다운로드:
  > pip download fastapi uvicorn xlwings pandas openpyxl -d ./wheels

  대상 서버에서 설치:
  > pip install --no-index --find-links=./wheels fastapi uvicorn xlwings pandas openpyxl

[실행]

  > cd <프로젝트 폴더>
  > uvicorn web_monitor:app --host 0.0.0.0 --port 8000 --reload

[접속]

  http://localhost:8000        (로컬)
  http://<서버IP>:8000         (사내망 다른 PC)

[동작]

  - sample.xlsx 를 xlwings 로 1초마다 읽어 브라우저에 실시간 표시
  - A1 기준 왼쪽 테이블, J1 기준 오른쪽 테이블 (date 컬럼 index)

[주의]

  - Excel 파일이 열려 있어야 함  (xlwings 는 Excel COM 방식)
  - sample.xlsx 는 실행 폴더와 같은 위치에 있어야 함
  - Windows 전용


--------------------------------------------------------------------------------
