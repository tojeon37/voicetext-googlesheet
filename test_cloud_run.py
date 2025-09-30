import requests

# 테스트 오디오 파일로 Cloud Run 서버 테스트
with open('test_audio.wav', 'rb') as f:
    files = {'audio': f}
    data = {
        'language': 'ko-KR',
        'sample_rate': 16000,
        'encoding': 'LINEAR16'
    }
    
    print('Cloud Run 서버 /transcribe 엔드포인트 테스트 중...')
    response = requests.post(
        'https://voicetext-api-6qtb5op6hq-du.a.run.app/transcribe',
        files=files,
        data=data,
        timeout=30
    )
    
    print(f'응답 상태 코드: {response.status_code}')
    print(f'응답 내용: {response.text}')
