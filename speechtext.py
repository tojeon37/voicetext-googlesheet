import io
import wave
import pyaudio
import threading
import os
import requests
from datetime import datetime

class SimpleVoiceProcessor:
    def __init__(self):
        """클로드간단버전 기반의 간단한 음성 처리기"""
        self.is_recording = False
        self.recording_thread = None
        self.gui = None
        
        # 클로드간단버전과 동일한 설정
        self.RATE = 16000  # Google Cloud 권장 샘플링 레이트
        self.CHUNK = 1024
        self.FORMAT = pyaudio.paInt16
        self.CHANNELS = 1
        self.RECORD_SECONDS = 15  # 15초로 연장
        
        # Cloud Run 서버 설정
        self.api_url = "https://voicetext-api-6qtb5op6hq-du.a.run.app"
        self.setup_cloud_run_api()
        
    def setup_cloud_run_api(self):
        """Cloud Run 서버 연결 설정"""
        try:
            print("🔗 Cloud Run 서버 연결 중...")
            
            # 서버 연결 테스트
            test_response = requests.get(f"{self.api_url}/", timeout=10)
            if test_response.status_code == 200:
                print("✅ Cloud Run 서버 연결 성공!")
                self.api_available = True
            else:
                print(f"⚠️ Cloud Run 서버 응답 오류: {test_response.status_code}")
                self.api_available = False
            
        except Exception as e:
            print(f"❌ Cloud Run 서버 연결 실패: {e}")
            self.api_available = False
    
    def set_gui(self, gui):
        """GUI 참조 설정"""
        self.gui = gui
        
    def start_recording(self):
        """녹음 시작"""
        if self.is_recording:
            return
            
        self.is_recording = True
        
        # 녹음 스레드 시작
        self.recording_thread = threading.Thread(target=self.record_and_recognize, daemon=True)
        self.recording_thread.start()
        
    def stop_recording(self):
        """녹음 중지"""
        self.is_recording = False
        print("녹음 중지 요청됨")
        
        # GUI 상태 업데이트
        if self.gui:
            self.gui.update_status("⏹️ 녹음 중지됨", "orange")
            self.gui.reset_buttons()
        
    def record_and_recognize(self):
        """클로드간단버전과 동일한 방식의 녹음 및 인식"""
        try:
            print(f"{self.RECORD_SECONDS}초간 녹음을 시작합니다...")
            
            # GUI 상태 업데이트
            if self.gui:
                self.gui.update_status("🎙️ 녹음 중...", "red")
            
            # 클로드간단버전과 동일한 녹음 설정
            audio = pyaudio.PyAudio()
            
            stream = audio.open(format=self.FORMAT,
                               channels=self.CHANNELS,
                               rate=self.RATE,
                               input=True,
                               frames_per_buffer=self.CHUNK)
            
            frames = []
            for i in range(0, int(self.RATE / self.CHUNK * self.RECORD_SECONDS)):
                if not self.is_recording:
                    print("사용자가 녹음을 중지했습니다.")
                    break
                    
                try:
                    data = stream.read(self.CHUNK, exception_on_overflow=False)
                    frames.append(data)
                except Exception as read_error:
                    print(f"오디오 읽기 오류: {read_error}")
                    continue
            
            stream.stop_stream()
            stream.close()
            audio.terminate()
            
            if not self.is_recording:
                # 사용자가 중지한 경우 - 수집된 데이터로 음성 인식 진행
                print("사용자가 녹음을 중지했습니다. 수집된 데이터로 음성 인식을 진행합니다.")
                if len(frames) > 0:
                    # 수집된 데이터가 있으면 음성 인식 진행
                    self.process_recorded_audio(frames, audio)
                else:
                    print("수집된 오디오 데이터가 없습니다.")
                    if self.gui:
                        self.gui.update_status("❌ 수집된 오디오 데이터가 없습니다", "red")
                        self.gui.reset_buttons()
                return
            
            if len(frames) == 0:
                print("녹음된 데이터가 없습니다.")
                if self.gui:
                    self.gui.update_status("❌ 녹음된 데이터가 없습니다", "red")
                    self.gui.reset_buttons()
                return
            
            # 클로드간단버전과 동일한 WAV 파일 저장
            filename = "simple_test.wav"
            try:
                wf = wave.open(filename, 'wb')
                wf.setnchannels(self.CHANNELS)
                wf.setsampwidth(audio.get_sample_size(self.FORMAT))
                wf.setframerate(self.RATE)
                wf.writeframes(b''.join(frames))
                wf.close()
                
                print(f"녹음 완료: {filename}")
                
                # GUI 상태 업데이트
                if self.gui:
                    self.gui.update_status("☁️ 음성 인식 중...", "blue")
                
                # 클로드간단버전과 동일한 음성 인식 수행
                self.speech_to_text_simple(filename)
                
            except Exception as wav_error:
                print(f"WAV 파일 저장 오류: {wav_error}")
                if self.gui:
                    self.gui.update_status(f"❌ WAV 파일 저장 오류: {wav_error}", "red")
                    self.gui.reset_buttons()
                return
            
            # 임시 파일 삭제
            try:
                if os.path.exists(filename):
                    os.remove(filename)
                    print("임시 파일 삭제 완료")
            except Exception as delete_error:
                print(f"임시 파일 삭제 오류: {delete_error}")
                
        except Exception as e:
            print(f"녹음 오류: {e}")
            import traceback
            traceback.print_exc()
            if self.gui:
                self.gui.update_status(f"❌ 녹음 오류: {e}", "red")
                self.gui.reset_buttons()
    
    def process_recorded_audio(self, frames, audio):
        """수집된 오디오 데이터 처리"""
        try:
            # 클로드간단버전과 동일한 WAV 파일 저장
            filename = "simple_test.wav"
            try:
                wf = wave.open(filename, 'wb')
                wf.setnchannels(self.CHANNELS)
                wf.setsampwidth(audio.get_sample_size(self.FORMAT))
                wf.setframerate(self.RATE)
                wf.writeframes(b''.join(frames))
                wf.close()
                
                print(f"녹음 완료: {filename}")
                
                # GUI 상태 업데이트
                if self.gui:
                    self.gui.update_status("☁️ 음성 인식 중...", "blue")
                
                # 클로드간단버전과 동일한 음성 인식 수행
                self.speech_to_text_simple(filename)
                
            except Exception as wav_error:
                print(f"WAV 파일 저장 오류: {wav_error}")
                if self.gui:
                    self.gui.update_status(f"❌ WAV 파일 저장 오류: {wav_error}", "red")
                    self.gui.reset_buttons()
                return
            
            # 임시 파일 삭제
            try:
                if os.path.exists(filename):
                    os.remove(filename)
                    print("임시 파일 삭제 완료")
            except Exception as delete_error:
                print(f"임시 파일 삭제 오류: {delete_error}")
                
        except Exception as e:
            print(f"오디오 처리 오류: {e}")
            import traceback
            traceback.print_exc()
            if self.gui:
                self.gui.update_status(f"❌ 오디오 처리 오류: {e}", "red")
                self.gui.reset_buttons()
    
    def speech_to_text_simple(self, audio_file):
        """Cloud Run 서버를 통한 음성 인식"""
        try:
            if not self.api_available:
                print("Cloud Run API 사용 불가능")
                text, confidence = "[오류] Cloud Run 서버 연결 실패", 0.0
                self.display_result(text, confidence)
                return
            
            # Cloud Run 서버로 HTTP 요청
            try:
                with open(audio_file, 'rb') as f:
                    files = {'audio': f}
                    data = {
                        'language': 'ko-KR',
                        'sample_rate': self.RATE,
                        'encoding': 'LINEAR16'
                    }
                    
                    print("☁️ Cloud Run 서버로 음성 인식 요청 중...")
                    response = requests.post(
                        f"{self.api_url}/transcribe",
                        files=files,
                        data=data,
                        timeout=60
                    )
                
                if response.status_code == 200:
                    result = response.json()
                    if result.get('success', False):
                        text = result.get('transcript', '')
                        confidence = result.get('confidence', 0.0)
                        
                        print(f"✅ 인식 완료: '{text}' (신뢰도: {confidence:.2f})")
                        self.display_result(text, confidence)
                    else:
                        error_msg = result.get('error', '음성 인식 실패')
                        print(f"❌ 서버 오류: {error_msg}")
                        self.display_result(f"[서버 오류] {error_msg}", 0.0)
                else:
                    print(f"❌ HTTP 오류: {response.status_code}")
                    print(f"응답 내용: {response.text}")
                    self.display_result(f"[HTTP 오류] {response.status_code}", 0.0)
                    
            except requests.exceptions.Timeout:
                print("❌ 요청 시간 초과")
                self.display_result("[오류] 서버 응답 시간 초과", 0.0)
            except requests.exceptions.RequestException as e:
                print(f"❌ 네트워크 오류: {e}")
                self.display_result(f"[네트워크 오류] {str(e)[:50]}...", 0.0)
            except Exception as e:
                print(f"❌ API 오류: {e}")
                self.display_result(f"[API 오류] {str(e)[:50]}...", 0.0)
            
        except Exception as e:
            print(f"❌ 음성 인식 오류: {e}")
            import traceback
            traceback.print_exc()
            if self.gui:
                self.gui.reset_buttons()
    
    def display_result(self, text, confidence):
        """인식 결과 표시"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        confidence_text = f" (신뢰도: {confidence:.2f})" if confidence > 0 else ""
        result = f"[{timestamp}] {text}{confidence_text}"
        print(f"인식 결과: {result}")
        
        # GUI가 있으면 GUI에도 표시
        if self.gui:
            self.gui.display_result(text, confidence)
            # 음성 인식 완료 후 GUI 상태 초기화
            self.gui.reset_buttons()

def main():
    """테스트용 메인 함수"""
    try:
        print("프로그램 시작...")
        processor = SimpleVoiceProcessor()
        print("음성 처리기 초기화 완료")
        return processor
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()