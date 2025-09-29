# 음성인식 구글시트 연동 프로그램

Google Cloud Speech-to-Text API를 사용한 한국어 음성인식 애플리케이션으로, 인식된 텍스트를 구글 스프레드시트에 자동으로 입력하는 프로그램입니다.

## 주요 기능

- 🎙️ 실시간 음성 녹음 (15초)
- ☁️ Google Cloud Speech-to-Text API 연동
- 📊 구글 스프레드시트 자동 저장
- 📝 사용자 지정 셀에 텍스트 입력
- 🔄 모든 시트 선택 가능
- 💾 로컬 CSV 파일 백업 저장
- ⚙️ 설정 자동 저장/복원

## 최적화된 설정

### 오디오 품질
- 샘플링 레이트: 16kHz (Google Cloud 권장값)
- 비트 깊이: 16-bit
- 채널: 모노 (1채널)
- 녹음 시간: 15초

### API 설정
- 언어: 한국어 (ko-KR) 우선
- 모델: latest_long (장문 처리 최적화)
- 향상된 모델 사용 (use_enhanced=True)
- 구두점 자동 추가
- 단어별 시간 정보 및 신뢰도 제공

### 구글 시트 연동
- 모든 스프레드시트 선택 가능
- 모든 시트 선택 가능
- 사용자 지정 셀 입력
- 자동 다음 행 이동
- 설정 자동 저장/복원

## 설치 및 설정

### 1. 저장소 클론
```bash
git clone https://github.com/tojeon37/voicetext-googlesheet.git
cd voicetext-googlesheet
```

### 2. 가상환경 생성 및 활성화 (권장)
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 3. 필요한 패키지 설치
```bash
pip install -r requirements.txt
```

### 4. Google Cloud 설정

#### 4.1 Google Cloud Console 설정
1. [Google Cloud Console](https://console.cloud.google.com/) 접속
2. 새 프로젝트 생성 또는 기존 프로젝트 선택
3. 다음 API 활성화:
   - Speech-to-Text API
   - Google Sheets API  
   - Google Drive API

#### 4.2 서비스 계정 키 생성
1. Google Cloud Console > IAM 및 관리 > 서비스 계정
2. "서비스 계정 만들기" 클릭
3. 서비스 계정 정보 입력:
   - 이름: `voicetext-service`
   - 설명: `음성인식 구글시트 연동 서비스`
4. 사용자에게 서비스 계정 액세스 권한 부여:
   - 역할: `편집자` 또는 `보안 관찰자`와 `스프레드시트 편집자`
5. 키 생성 및 JSON 다운로드 (파일명: `service_account.json`)

### 5. 구글 스프레드시트 설정
1. 원하는 이름으로 구글 스프레드시트 생성
2. 컬럼 헤더가 있는 첫 번째 행을 설정 (예: A1="타임스탬프", B1="인식된 텍스트", C1="신뢰도")
3. 생성된 서비스 계정의 이메일 주소로 스프레드시트를 공유하고 `편집자` 권한 부여

### 6. 프로젝트 설정
1. 다운로드한 서비스 계정 키 파일을 프로젝트 루트 디렉토리에 `service_account.json`으로 복사
2. 또는 환경변수 `GOOGLE_APPLICATION_CREDENTIALS`로 키 파일 경로 설정

### 7. 실행
```bash
python main.py
```

## 파일 구조

```
voicetext-googlesheet/
├── main.py                    # 메인 실행 파일
├── gui.py                     # GUI 인터페이스
├── speechtext.py              # 음성 인식 처리
├── app_settings.json          # 애플리케이션 설정
├── service_account.json       # Google Cloud API 키 (사용자가 설정)
├── requirements.txt           # 필요한 패키지 목록
├── README_GITHUB.md           # 프로젝트 설명 (GitHub용)
├── .gitignore                # Git 무시 파일 설정
└── 음성인식_데이터_GoogleCloud.csv  # 로컬 백업 데이터 (자동 생성)
```

## 사용법

1. **프로그램 실행**: `python main.py` 명령으로 실행
2. **스프레드시트 선택**: 드롭다운에서 원하는 스프레드시트 선택
3. **시트 선택**: 드롭다운에서 원하는 시트 선택  
4. **씨 주소 입력**: 입력할 셀 주소 입력 (예: A1, B5, C10)
5. **녹음 시작**: "🎙️ 녹음 시작" 버튼 클릭
6. **음성 입력**: 마이크에 명확하게 말하기 (중간에 중지 가능)
7. **녹음 중지**: "⏹️ 녹음 중지" 버튼 클릭
8. **결과 확인**: 인식된 텍스트가 지정된 셀에 자동 입력
9. **자동 이동**: 다음 행으로 자동 이동

## GUI 구성 요소

- **🎙️ 녹음 시작 버튼**: 음성 녹음 시작
- **⏹️ 녹음 중지 버튼**: 음성 녹음 중지  
- **상태 표시**: 현재 프로그램 상태 표시
- **타이머**: 남은 녹음 시간 표시
- **인식된 텍스트 영역**: 인식 결과 표시
- **스프레드시트 선택**: 드롭다운으로 스프레드시트 선택
- **시트 선택**: 드롭다운으로 시트 선택
- **새로고침 버튼**: 시트 목록 새로고침
- **셀 주소 입력**: 입력할 셀 주소 지정
- **현재 위치**: 현재 선택된 셀 주소 표시

## 설정 관리

### 자동 저장되는 설정
- 마지막 사용 스프레드시트
- 마지막 사용 시트
- 마지막 사용 셀 주소
- 프로그램 재시작 시 자동 복원

### 설정 파일 (app_settings.json)
```json
{
  "last_spreadsheet": "음성기록",
  "last_sheet": "시트1", 
  "last_cell": "A1",
  "last_row": 1,
  "last_col": 1
}
```

## 환경변수 설정 (선택사항)

```bash
# Windows
set GOOGLE_APPLICATION_CREDENTIALS=path\to\service_account.json

# macOS/Linux
export GOOGLE_APPLICATION_CREDENTIALS=path/to/service_account.json
```

## 인식률 향상 팁

- 조용한 환경에서 사용
- 마이크에 적절한 거리 유지 (10-20cm)
- 명확하고 천천히 발음
- 배경 소음 최소화
- 완전한 문장으로 말하기
- 중간에 녹음 중지하여 부분 인식도 가능

## 문제 해결

### 인식률이 낮은 경우
1. 마이크 품질 확인
2. 환경 소음 확인
3. 발음 명확도 확인
4. 마이크 거리 조정

### API 연결 오류
1. Google Cloud 프로젝트에서 Speech-to-Text API 활성화 확인
2. 서비스 계정 키 파일이 올바른 위치에 있는지 확인
3. 서비스 계정에 필요한 권한이 부여되었는지 확인
4. 인터넷 연결 상태 확인
5. API 할당량 확인

### 구글 시트 연결 오류
1. 서비스 계정이 스프레드시트에 접근 권한이 있는지 확인
2. 서비스 계정 이메일로 스프레드시트가 올바르게 공유되었는지 확인
3. Google Sheets API 활성화 확인
4. Google Drive API 활성화 확인

## 빌드 및 배포

### PyInstaller로 실행 파일 생성

#### Windows용 실행 파일 생성
```bash
# 콘솔창 있는 버전
pyinstaller --onefile --console --name VoiceRecognitionApp main.py

# 콘솔창 없는 버전  
pyinstaller --onefile --windowed --name VoiceRecognitionApp_gui main.py

# 실행 파일과 서비스 계정 키 파일 함께 패키징
pyinstaller --onefile --console --name VoiceRecognitionApp --add-data "service_account.json;." main.py
```

#### Linux/macOS용 실행 파일 생성
```bash
# 콘솔창 있는 버전
pyinstaller --onefile --console --name VoiceRecognitionApp main.py

# 콘솔창 없는 버전
pyinstaller --onefile --windowed --name VoiceRecognitionApp_gui main.py
```

## 기술 스택

- **음성 처리**: Google Cloud Speech-to-Text API, PyAudio, Wave
- **GUI**: Tkinter, ttk
- **구글 서비스**: gspread, google.oauth2.service_account
- **데이터 처리**: JSON, CSV, Threading
- **배포**: PyInstaller

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## 기여

버그 리포트나 기능 개선 제안은 언제든 환영합니다!

## 변경사항

### v2.0
- 모든 스프레드시트 선택 가능
- 모든 시트 선택 가능
- 개선된 설정 저장/복원
- 자동 다음 셀 이동 기능
- 향상된 오류 처리

### v1.0
- 기본 음성인식 기능
- 구글 스프레드시트 연동
- 기본 GUI 인터페이스
