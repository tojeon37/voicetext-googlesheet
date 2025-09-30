# Cloud Run 서버 연동 가이드

## 📋 개요
이 문서는 음성인식 구글시트 연동 프로그램이 Google Cloud Run 서버를 통해 
Google Speech API를 호출하도록 변경된 구조를 설명합니다.

## 🔄 구조 변경

### 기존 구조 (보안 위험)
```
고객 PC → Google Speech API (직접 호출)
         ↓
    API 키 노출 위험
```

### 변경된 구조 (보안 강화)
```
고객 PC → Cloud Run 서버 → Google Speech API
         ↓
    API 키 안전 보관
```

## 🛠️ 주요 변경사항

### 1. speechtext.py 수정

#### 기존 코드
```python
from google.cloud import speech

# Google Cloud Speech-to-Text 클라이언트
self.speech_client = speech.SpeechClient.from_service_account_file(
    "voicetext-472910-82f1fa0a8fbe.json"
)

# 직접 API 호출
response = self.speech_client.recognize(config=config, audio=audio_obj)
```

#### 변경된 코드
```python
import requests

# Cloud Run 서버 설정
self.api_url = "https://voicetext-api-6qtb5op6hq-du.a.run.app"

# HTTP 요청으로 API 호출
response = requests.post(
    f"{self.api_url}/transcribe",
    files={'audio': f},
    data={'language': 'ko-KR', 'sample_rate': 16000, 'encoding': 'LINEAR16'},
    timeout=60
)
```

### 2. requirements.txt 수정

#### 제거된 패키지
```
google-cloud-speech==2.21.0  # 제거
```

#### 추가된 패키지
```
requests==2.31.0  # 추가
```

## 🔧 Cloud Run 서버 설정

### 서버 URL
```
https://voicetext-api-6qtb5op6hq-du.a.run.app
```

### API 엔드포인트

#### 1. 서버 연결 테스트
```
GET /
응답: 200 OK
```

#### 2. 음성 인식 요청
```
POST /transcribe
Content-Type: multipart/form-data

파라미터:
- audio: 오디오 파일 (WAV 형식)
- language: 언어 코드 (기본값: ko-KR)
- sample_rate: 샘플링 레이트 (기본값: 16000)
- encoding: 인코딩 형식 (기본값: LINEAR16)

응답:
{
  "success": true,
  "transcript": "인식된 텍스트",
  "confidence": 0.95
}
```

## 🚀 배포 방법

### 1. 기존 배포 파일에서 제거할 항목
- `voicetext-472910-82f1fa0a8fbe.json` (Google API 키 파일)
- Google Cloud 관련 설정 파일들

### 2. 새로운 배포 파일 구성
```
배포폴더/
├── VoiceRecognitionApp_ver2.exe
├── app_settings.json
├── 방법3A_사용법.md
├── Cloud_Run_연동_가이드.md
└── README.txt
```

### 3. 고객별 설정 파일 예시

#### A고객용 app_settings.json
```json
{
  "last_sheet": "시트1",
  "last_cell": "A1",
  "last_row": 1,
  "last_col": 1,
  "last_spreadsheet": "A고객_음성기록",
  "allowed_spreadsheets": ["A고객_음성기록"],
  "cloud_run_url": "https://voicetext-api-6qtb5op6hq-du.a.run.app"
}
```

## 🔒 보안 강화 효과

### 1. API 키 보호
- 고객 PC에 Google API 키 파일이 없음
- API 키는 Cloud Run 서버에서만 관리
- 키 노출 위험 완전 제거

### 2. 접근 제어
- Cloud Run 서버에서 API 호출 제어
- 사용량 모니터링 및 제한 가능
- 비정상적인 요청 차단 가능

### 3. 중앙 관리
- 모든 API 호출이 서버를 통해 관리
- 로그 및 모니터링 중앙화
- 업데이트 및 유지보수 용이

## 📊 성능 및 안정성

### 1. 네트워크 의존성
- 인터넷 연결 필수
- 네트워크 지연 시간 고려
- 오프라인 사용 불가

### 2. 오류 처리
- 네트워크 오류 시 적절한 메시지 표시
- 타임아웃 설정 (60초)
- 재시도 로직 구현 가능

### 3. 확장성
- 서버 부하 분산 가능
- 다중 서버 구성 가능
- 자동 스케일링 지원

## 🛠️ 문제 해결

### 1. 연결 오류
```
❌ Cloud Run 서버 연결 실패
해결방법:
- 인터넷 연결 확인
- 방화벽 설정 확인
- 서버 URL 확인
```

### 2. 인식 실패
```
❌ 서버 오류: 음성 인식 실패
해결방법:
- 오디오 파일 형식 확인
- 서버 로그 확인
- 재시도
```

### 3. 타임아웃 오류
```
❌ 요청 시간 초과
해결방법:
- 네트워크 상태 확인
- 오디오 파일 크기 확인
- 서버 상태 확인
```

## 📈 모니터링 및 관리

### 1. 서버 모니터링
- Google Cloud Console에서 서버 상태 확인
- 로그 및 메트릭 모니터링
- 알림 설정

### 2. 사용량 추적
- API 호출 횟수 추적
- 응답 시간 모니터링
- 오류율 추적

### 3. 비용 관리
- Cloud Run 사용량 모니터링
- Google Speech API 비용 추적
- 최적화 방안 검토

## 🔄 마이그레이션 가이드

### 기존 고객 업데이트 방법

1. **백업 생성**
   - 기존 설정 파일 백업
   - 사용자 데이터 백업

2. **새 버전 설치**
   - 새 배포 파일 다운로드
   - 기존 파일 교체

3. **설정 확인**
   - app_settings.json 설정 확인
   - Cloud Run 서버 연결 테스트

4. **기능 테스트**
   - 음성 인식 기능 테스트
   - 스프레드시트 연동 테스트

### 롤백 방법
- 기존 버전으로 되돌리기
- 백업된 설정 파일 복원
- Google API 키 파일 복원

## 📞 지원 및 문의

### 기술 지원
- Cloud Run 서버 관련 문의
- API 연동 문제 해결
- 성능 최적화 상담

### 응답 시간
- 일반 문의: 24시간 이내
- 긴급 문의: 4시간 이내
- 서버 장애: 즉시 대응

---

이 가이드를 따라하면 Cloud Run 서버를 통한 안전하고 효율적인 
음성 인식 서비스를 구축할 수 있습니다.
