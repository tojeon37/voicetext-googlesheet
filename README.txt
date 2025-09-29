음성인식 구글시트 연동 프로그램

Google Cloud Speech-to-Text API를 사용한 한국어 음성인식 애플리케이션으로, 인식된 텍스트를 구글 스프레드시트에 자동으로 입력하는 프로그램입니다.

주요 기능

- 실시간 음성 녹음 (15초)
- Google Cloud Speech-to-Text API 연동
- 구글 스프레드시트 자동 저장
- 사용자 지정 셀에 텍스트 입력
- 모든 시트 선택 가능
- 로컬 CSV 파일 백업 저장
- 설정 자동 저장/복원

최적화된 설정

오디오 품질
- 샘플링 레이트: 16kHz (Google Cloud 권장값)
- 비트 깊이: 16-bit
- 채널: 모노 (1채널)
- 녹음 시간: 15초

API 설정
- 언어: 한국어 (ko-KR) 우선
- 모델: latest_long (장문 처리 최적화)
- 향상된 모델 사용 (use_enhanced=True)
- 구두점 자동 추가
- 단어별 시간 정보 및 신뢰도 제공

구글 시트 연동
- 스프레드시트명: "음성기록"
- 모든 시트 선택 가능
- 사용자 지정 셀 입력
- 자동 다음 행 이동
- 설정 자동 저장/복원

설치 및 설정

1. 필요한 패키지 설치
pip install -r requirements.txt

2. Google Cloud 설정
1. Google Cloud Console에서 Speech-to-Text API 활성화
2. Google Sheets API 활성화
3. Google Drive API 활성화
4. 서비스 계정 키 파일 생성 (voicetext-472910-82f1fa0a8fbe.json)
5. 서비스 계정에 구글 스프레드시트 접근 권한 부여

3. 구글 스프레드시트 설정
1. "음성기록" 이름의 스프레드시트 생성
2. 서비스 계정 이메일로 스프레드시트 공유
3. 편집 권한 부여

4. 실행
python main.py

파일 구조

├── main.py                                    # 메인 실행 파일
├── gui.py                                     # GUI 인터페이스
├── speechtext.py                              # 음성 인식 처리
├── app_settings.json                          # 애플리케이션 설정
├── voicetext-472910-82f1fa0a8fbe.json        # Google Cloud API 키
├── requirements.txt                           # 필요한 패키지 목록
├── README.txt                                 # 프로젝트 설명
└── 음성인식_데이터_GoogleCloud.csv            # 로컬 백업 데이터

사용법

1. 프로그램 실행: python main.py 명령으로 실행
2. 시트 선택: 드롭다운에서 원하는 시트 선택
3. 셀 주소 입력: 입력할 셀 주소 입력 (예: A1, B5, C10)
4. 녹음 시작: "녹음시작" 버튼 클릭
5. 음성 입력: 15초간 마이크에 명확하게 말하기
6. 녹음 중지: "녹음중지" 버튼 클릭
7. 결과 확인: 인식된 텍스트가 지정된 셀에 자동 입력
8. 자동 이동: 다음 행으로 자동 이동

GUI 구성 요소

- 녹음시작 버튼: 음성 녹음 시작
- 녹음중지 버튼: 음성 녹음 중지
- 녹음중 깜빡임: 녹음 중 상태 표시
- 인식된 텍스트 창: 인식 결과 표시
- 시트 선택: 드롭다운으로 시트 선택
- 새로고침 버튼: 시트 목록 새로고침
- 셀 주소 입력: 입력할 셀 주소 지정
- 현재 위치: 현재 선택된 셀 주소 표시

설정 관리

자동 저장되는 설정
- 마지막 사용 시트
- 마지막 사용 셀 주소
- 프로그램 재시작 시 자동 복원

설정 파일 (app_settings.json)
{
  "last_sheet": "시트1",
  "last_cell": "A1",
  "last_row": 1,
  "last_col": 1
}

인식률 향상 팁

- 조용한 환경에서 사용
- 마이크에 적절한 거리 유지 (10-20cm)
- 명확하고 천천히 발음
- 배경 소음 최소화
- 15초 내에 완전한 문장으로 말하기

문제 해결

인식률이 낮은 경우
1. 마이크 품질 확인
2. 환경 소음 확인
3. 발음 명확도 확인
4. 마이크 거리 조정

API 연결 오류
1. Google Cloud API 키 파일 확인
2. 인터넷 연결 상태 확인
3. API 할당량 확인
4. 서비스 계정 권한 확인

구글 시트 연결 오류
1. 스프레드시트 공유 설정 확인
2. 서비스 계정 이메일로 공유되었는지 확인
3. Google Sheets API 활성화 확인
4. Google Drive API 활성화 확인

기술 스택

- 음성 처리: Google Cloud Speech-to-Text API, PyAudio, Wave
- GUI: Tkinter, ttk
- 구글 서비스: gspread, google.oauth2.service_account
- 데이터 처리: JSON, CSV, Threading

라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

기여

버그 리포트나 기능 개선 제안은 언제든 환영합니다!
