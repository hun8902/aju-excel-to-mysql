# Excel to MySQL Importer

## 목차
1. [소개](#소개)
2. [요구사항](#요구사항)
3. [설치 방법](#설치-방법)
4. [프로그램 설정](#프로그램-설정)
5. [사용 방법](#사용-방법)
6. [데이터베이스 설정](#데이터베이스-설정)
7. [텔레그램 봇 설정](#텔레그램-봇-설정)
8. [엑셀 파일 형식](#엑셀-파일-형식)
9. [결과 파일](#결과-파일)
10. [문제 해결](#문제-해결)
11. [자주 묻는 질문](#자주-묻는-질문)

## 소개
Excel to MySQL Importer는 Excel 파일의 데이터를 MySQL 데이터베이스로 쉽게 가져올 수 있는 GUI 프로그램입니다. 이 프로그램은 다음과 같은 주요 기능을 제공합니다:
일단은 엑셀에서 데이터 가져오는거에 맞춤 제작된것이지만 점차 범용으로 쓰일수 있게 개발 예정입니다.

- Excel 파일에서 데이터 읽기 (여러 시트 지원)
- MySQL 데이터베이스에 데이터 자동 입력
- 텔레그램을 통한 처리 결과 알림
- 진행 상황 실시간 모니터링
- 설정 저장 및 불러오기

## 요구사항
- Windows 7 이상
- MySQL 서버 5.7 이상
- 인터넷 연결 (텔레그램 알림 기능 사용 시)

### 필수 데이터베이스 테이블 구조
```sql
CREATE TABLE aju_facilities (
    id INT AUTO_INCREMENT PRIMARY KEY,
    fc_code VARCHAR(50),
    fc_purpose VARCHAR(50),
    fc_use VARCHAR(20),
    fc_name VARCHAR(50),
    fc_size VARCHAR(20),
    fc_model VARCHAR(50),
    fc_maker VARCHAR(30),
    fc_buy_date VARCHAR(20)
);
```

## 설치 방법

### 실행 파일 다운로드
1. 제공된 `ExcelToMySQLImporter.exe` 파일을 다운로드합니다.
2. 원하는 위치에 저장합니다.

### 소스코드에서 실행 파일 생성하기
1. Python 3.7 이상을 설치합니다.
2. 필요한 패키지들을 설치합니다:
```bash
python -m pip install pyinstaller pandas mysql-connector-python requests
```
3. 소스코드를 다운로드하고 압축을 풉니다.
4. 명령 프롬프트(CMD)를 관리자 권한으로 실행합니다.
5. 소스코드가 있는 폴더로 이동합니다.
6. 다음 명령어를 실행하여 exe 파일을 생성합니다:
```bash
python -m PyInstaller --onefile --windowed --name ExcelToMySQLImporter main.py
```
7. 생성된 exe 파일은 `dist` 폴더에서 찾을 수 있습니다.

## 프로그램 설정

### 초기 설정
1. 프로그램을 처음 실행하면 `settings.json` 파일이 자동으로 생성됩니다.
2. 데이터베이스 연결 정보와 텔레그램 설정을 입력합니다.
3. 입력한 설정은 자동으로 저장되며 다음 실행 시 불러옵니다.

### settings.json 파일 구조
```json
{
    "db_host": "localhost",
    "db_user": "root",
    "db_name": "aju_erp",
    "telegram_token": "your_telegram_bot_token",
    "telegram_chat_id": "your_telegram_chat_id"
}
```

## 사용 방법

### 기본 사용법
1. 프로그램을 실행합니다.
2. "Excel 파일 선택" 버튼을 클릭하여 처리할 Excel 파일을 선택합니다.
3. 데이터베이스 연결 정보를 입력합니다.
4. (선택사항) 텔레그램 알림을 위한 Bot Token과 Chat ID를 입력합니다.
5. "실행" 버튼을 클릭하여 데이터 처리를 시작합니다.
6. 진행 상황은 화면의 로그 창에서 확인할 수 있습니다.

### 상세 처리 과정
1. Excel 파일의 각 시트를 순차적으로 처리합니다.
2. 각 시트에서 "관리번호" 열을 찾아 데이터의 시작 위치를 확인합니다.
3. 데이터를 읽어서 MySQL 데이터베이스의 aju_facilities 테이블에 입력합니다.
4. 처리 결과는 별도의 텍스트 파일로 저장됩니다.
5. 설정된 경우 텔레그램으로 처리 완료 알림을 전송합니다.

## 데이터베이스 설정

### MySQL 서버 설정
- Host: MySQL 서버의 주소 (기본값: localhost)
- User: 데이터베이스 사용자 이름
- Password: 데이터베이스 비밀번호
- Database: 사용할 데이터베이스 이름 (기본값: aju_erp)

### 권한 설정
데이터베이스 사용자는 다음 권한이 필요합니다:
- INSERT: 데이터 입력을 위함
- SELECT: 데이터 확인을 위함

## 텔레그램 봇 설정

### 봇 생성 방법
1. Telegram에서 @BotFather를 검색합니다.
2. `/newbot` 명령어로 새 봇을 생성합니다.
3. 봇의 이름과 사용자명을 설정합니다.
4. 생성된 봇의 토큰을 프로그램에 입력합니다.

### Chat ID 확인 방법
1. 생성한 봇과 대화를 시작합니다.
2. 웹 브라우저에서 다음 주소에 접속합니다:
   `https://api.telegram.org/bot<YourBOTToken>/getUpdates`
3. 응답에서 "chat" 객체 내의 "id" 값을 확인합니다.

## 엑셀 파일 형식

### 필수 열
- 관리번호
- 사용부서
- 용도
- 설비명칭
- 규격/용량
- 모델명
- 제작사
- 구입일자

### 파일 형식 요구사항
- .xlsx 또는 .xls 파일
- 각 시트의 첫 번째 행에 열 이름이 있어야 함
- "관리번호" 열이 반드시 포함되어야 함

## 결과 파일

### 결과 파일 형식
- 파일명: result_YYYYMMDD_HHMMSS.txt
- 위치: 프로그램 실행 폴더
- 내용: 각 시트별 처리 결과 및 오류 메시지

### 결과 파일 예시
```
Sheet1 시트: 100개 데이터 삽입 완료
Sheet2 시트: '관리번호' 열을 찾을 수 없음
Sheet3 시트: 50개 데이터 삽입 완료
```

## 문제 해결

### 일반적인 문제

1. 프로그램이 실행되지 않는 경우
   - Windows defender 또는 백신 프로그램의 실행 허용 확인
   - 관리자 권한으로 실행 시도

2. 데이터베이스 연결 오류
   - MySQL 서버 실행 상태 확인
   - 사용자 이름과 비밀번호 확인
   - 방화벽 설정 확인

3. Excel 파일 처리 오류
   - 파일이 다른 프로그램에서 열려있는지 확인
   - 파일 형식이 .xlsx 또는 .xls인지 확인
   - "관리번호" 열의 존재 여부 확인

4. 텔레그램 알림 오류
   - 인터넷 연결 상태 확인
   - Bot Token과 Chat ID 정확성 확인
   - 봇과의 대화가 시작되었는지 확인

### 오류 메시지 및 해결 방법

1. "Excel 파일을 선택해주세요."
   - 파일 선택 버튼을 클릭하여 Excel 파일을 선택

2. "MySQL 서버에 연결할 수 없습니다."
   - 데이터베이스 연결 정보 확인
   - MySQL 서버 실행 상태 확인

3. "'관리번호' 열을 찾을 수 없음"
   - Excel 파일의 열 이름 확인
   - 열 이름에 공백이나 특수문자가 없는지 확인

## 자주 묻는 질문

Q: 여러 개의 Excel 파일을 한 번에 처리할 수 있나요?
A: 현재 버전에서는 한 번에 하나의 파일만 처리 가능합니다.

Q: 처리 중 오류가 발생하면 어떻게 되나요?
A: 오류가 발생한 시트는 건너뛰고 다음 시트 처리를 계속합니다. 모든 오류는 결과 파일에 기록됩니다.

Q: 이미 입력된 데이터는 어떻게 되나요?
A: 현재 버전에서는 중복 검사 없이 모든 데이터가 새로 입력됩니다.

Q: 텔레그램 알림은 필수인가요?
A: 아니요, 선택사항입니다. 텔레그램 설정을 비워두면 알림 기능은 작동하지 않습니다.

Q: 프로그램 설정은 어디에 저장되나요?
A: 프로그램 실행 폴더의 settings.json 파일에 저장됩니다.
