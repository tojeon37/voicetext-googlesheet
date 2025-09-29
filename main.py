import tkinter as tk
import gspread
from google.oauth2.service_account import Credentials
import os
import csv
import json
from datetime import datetime
from gui import SimpleVoiceGUI
from speechtext import SimpleVoiceProcessor

class SettingsManager:
    """설정 파일 관리 클래스"""
    def __init__(self, settings_file="app_settings.json"):
        self.settings_file = settings_file
        self.settings = self.load_settings()
    
    def load_settings(self):
        """설정 파일에서 설정 불러오기"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    print(f"✅ 설정 파일 불러오기 성공: {self.settings_file}")
                    return settings
            else:
                print("📁 설정 파일이 없습니다. 기본값을 사용합니다.")
                return self.get_default_settings()
        except Exception as e:
            print(f"❌ 설정 파일 불러오기 실패: {e}")
            return self.get_default_settings()
    
    def save_settings(self):
        """설정을 파일에 저장"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
                print(f"✅ 설정 파일 저장 성공: {self.settings_file}")
        except Exception as e:
            print(f"❌ 설정 파일 저장 실패: {e}")
    
    def get_default_settings(self):
        """기본 설정값 반환"""
        return {
            "last_spreadsheet": "음성기록",
            "last_sheet": "시트1",
            "last_cell": "A1",
            "last_row": 1,
            "last_col": 1
        }
    
    def get_setting(self, key, default=None):
        """설정값 가져오기"""
        return self.settings.get(key, default)
    
    def set_setting(self, key, value):
        """설정값 저장"""
        self.settings[key] = value
        self.save_settings()

class GoogleSheetHandler:
    def __init__(self, settings_manager=None):
        """구글 스프레드시트 핸들러"""
        self.sheet = None
        self.spreadsheet = None  # 전체 스프레드시트 객체
        self.current_row = 1  # 현재 입력할 행 번호
        self.current_col = 1  # 현재 입력할 열 번호 (A열)
        self.settings_manager = settings_manager
        self.setup_google_sheet()
    
    def setup_google_sheet(self):
        """구글 스프레드시트 설정"""
        try:
            print("🔗 구글 시트 연결 중...")
            
            # Google Sheets 연결 활성화
            print("📊 Google Sheets 연결 활성화")
            
            # 구글 인증
            print("🔑 구글 인증 중...")
            scope = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            
            # 서비스 계정 키 파일 사용
            creds = Credentials.from_service_account_file(
                "voicetext-472910-82f1fa0a8fbe.json", 
                scopes=scope
            )
            
            # 구글 시트 클라이언트 생성
            print("📊 구글 시트 클라이언트 생성 중...")
            self.gc = gspread.authorize(creds)
            
            # 서비스 계정 정보 출력 (디버깅용)
            try:
                service_account_info = creds.service_account_email
                print(f"🔑 서비스 계정: {service_account_info}")
            except Exception as e:
                print(f"⚠️ 서비스 계정 정보 확인 실패: {e}")
            
            # 스프레드시트 접근 방법 1: 직접 스프레드시트 ID 사용
            print("🔍 스프레드시트 접근 중...")
            try:
                # 방법 1: 스프레드시트 ID로 직접 접근 (가장 안정적)
                # 스프레드시트 URL에서 ID를 추출하여 사용
                # 예: https://docs.google.com/spreadsheets/d/1ABC123.../edit
                # ID 부분만 사용: 1ABC123...
                
                # 스프레드시트 이름으로 접근 시도
                try:
                    spreadsheet = self.gc.open("음성기록")
                    # 기본 시트 설정하지 않음 - GUI에서 설정하도록 함
                    self.spreadsheet = spreadsheet  # 전체 스프레드시트 객체 저장
                    print(f"✅ 스프레드시트 직접 접근 성공: 음성기록")
                    print("✅ 구글 스프레드시트 연결 성공: 음성기록")
                    return
                except Exception as direct_error:
                    print(f"직접 접근 실패: {direct_error}")
                    
                    # 방법 2: 기존 스프레드시트 목록에서 검색
                    print("🔍 기존 스프레드시트 목록에서 검색 중...")
                    spreadsheets = self.gc.openall()
                    target_sheet = None
                    
                    for spreadsheet in spreadsheets:
                        if spreadsheet.title == "음성기록":
                            target_sheet = spreadsheet
                            print(f"발견된 시트: {spreadsheet.title}")
                            break
                    
                    if target_sheet:
                        # 기본 시트 설정하지 않음 - GUI에서 설정하도록 함
                        self.spreadsheet = target_sheet
                        print(f"✅ 기존 스프레드시트 발견: {target_sheet.title}")
                        print("✅ 구글 스프레드시트 연결 성공: 음성기록")
                    else:
                        print("❌ '음성기록' 스프레드시트를 찾을 수 없습니다.")
                        print("📁 로컬 CSV 파일로 폴백")
                        self.sheet = None
                
            except Exception as e:
                print(f"스프레드시트 접근 오류: {e}")
                # 로컬 파일로 폴백
                print("📁 로컬 CSV 파일로 폴백")
                self.sheet = None
                
        except Exception as e:
            print(f"❌ 구글 시트 연결 실패: {e}")
            print("📁 로컬 CSV 파일로 폴백")
            self.sheet = None
    
    def get_current_cell_position(self):
        """현재 활성화된 셀의 위치를 가져오기"""
        try:
            # 구글 시트에서 현재 선택된 범위 정보 가져오기
            # 실제로는 사용자가 직접 셀을 선택해야 하므로, 
            # 여기서는 사용자가 지정한 위치를 사용
            return "A1"  # 기본값, 사용자가 원하는 셀로 변경 가능
        except Exception as e:
            print(f"현재 셀 위치 확인 오류: {e}")
            return "A1"
    
    def save_to_sheet(self, text, confidence, target_cell="A1"):
        """스프레드시트에 데이터 저장 (사용자 지정 셀에 텍스트만 입력)"""
        try:
            if self.sheet:
                # 구글 스프레드시트에 저장 (텍스트만 지정된 셀에 입력)
                try:
                    # 셀 위치를 행/열 번호로 변환
                    import re
                    match = re.match(r'([A-Z]+)(\d+)', target_cell)
                    if match:
                        col_letter = match.group(1)
                        row_num = int(match.group(2))
                        
                        # 열 문자를 숫자로 변환 (A=1, B=2, C=3, ...)
                        col_num = 0
                        for char in col_letter:
                            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
                        
                        # 지정된 셀에 텍스트만 입력 (타임스탬프, 신뢰도 없이)
                        self.sheet.update_cell(row_num, col_num, text)
                        print(f"✅ 구글 스프레드시트에 텍스트 입력 완료: {text[:30]}...")
                        print(f"📍 입력 위치: {target_cell} 셀")
                    else:
                        print("❌ 잘못된 셀 위치 형식입니다. A1 형식으로 입력해주세요.")
                        return
                        
                except Exception as e:
                    print(f"구글 시트 저장 오류: {e}")
                    # 로컬 파일로 폴백
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.save_to_local_file(timestamp, text, confidence)
            else:
                # 로컬 파일에 저장
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.save_to_local_file(timestamp, text, confidence)
                
        except Exception as e:
            print(f"데이터 저장 오류: {e}")
    
    def get_all_spreadsheets(self):
        """모든 스프레드시트 목록 가져오기 (gspread 사용)"""
        try:
            print("🔍 스프레드시트 목록 가져오기...")
            
            if not hasattr(self, 'gc') or not self.gc:
                # gspread 클라이언트 다시 생성
                scope = [
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
                creds = Credentials.from_service_account_file(
                    "voicetext-472910-82f1fa0a8fbe.json", 
                    scopes=scope
                )
                self.gc = gspread.authorize(creds)
            
            # 모든 스프레드시트 가져오기 (gspread 방법)
            spreadsheets = self.gc.openall()
            
            all_spreadsheets = []
            for spreadsheet in spreadsheets:
                all_spreadsheets.append({
                    'title': spreadsheet.title,
                    'id': spreadsheet.id,
                    'spreadsheet': spreadsheet
                })
                print(f"  📊 {spreadsheet.title} (ID: {spreadsheet.id})")
            
            print(f"✅ 스프레드시트 목록 가져오기 성공: {len(all_spreadsheets)}개")
            return all_spreadsheets
                
        except Exception as e:
            print(f"❌ 스프레드시트 목록 가져오기 실패: {e}")
            print("💡 해결 방법:")
            print("  1. 서비스 계정이 스프레드시트에 접근 권한이 있는지 확인")
            print("  2. 스프레드시트를 서비스 계정 이메일과 공유했는지 확인")
            print(f"  3. 서비스 계정 이메일: {getattr(self.gc.auth, 'service_account_email', 'Unknown') if hasattr(self, 'gc') and self.gc else 'Unknown'}")
            return []

    def get_all_sheets(self):
        """모든 시트 목록 가져오기"""
        try:
            if not self.spreadsheet:
                return []
            
            # 모든 워크시트 가져오기
            worksheets = self.spreadsheet.worksheets()
            
            # 모든 시트 반환
            all_sheets = []
            for sheet in worksheets:
                all_sheets.append({
                    'title': sheet.title,
                    'sheet': sheet
                })
            
            return all_sheets
            
        except Exception as e:
            print(f"모든 시트 목록 가져오기 실패: {e}")
            return []
    
    def set_target_spreadsheet(self, spreadsheet_title):
        """대상 스프레드시트 설정"""
        try:
            # 모든 스프레드시트 목록 가져오기
            all_spreadsheets = self.get_all_spreadsheets()
            
            for spreadsheet_info in all_spreadsheets:
                if spreadsheet_info['title'] == spreadsheet_title:
                    # 새로운 스프레드시트로 변경
                    self.spreadsheet = spreadsheet_info['spreadsheet']
                    self.sheet = None  # 현재 시트 초기화
                    print(f"✅ 대상 스프레드시트 변경: {spreadsheet_title}")
                    return True
            
            print(f"❌ 스프레드시트를 찾을 수 없습니다: {spreadsheet_title}")
            return False
            
        except Exception as e:
            print(f"스프레드시트 설정 오류: {e}")
            return False

    def set_target_sheet(self, sheet_title):
        """대상 시트 설정"""
        try:
            if not self.spreadsheet:
                return False
            
            worksheets = self.spreadsheet.worksheets()
            for sheet in worksheets:
                if sheet.title == sheet_title:
                    self.sheet = sheet
                    print(f"✅ 대상 시트 변경: {sheet_title}")
                    return True
            
            print(f"❌ 시트를 찾을 수 없습니다: {sheet_title}")
            return False
            
        except Exception as e:
            print(f"시트 설정 오류: {e}")
            return False
    
    def save_to_local_file(self, timestamp, text, confidence):
        """로컬 CSV 파일에 저장 (Excel 호환 UTF-8 BOM 포함)"""
        try:
            filename = "음성인식_데이터_GoogleCloud.csv"
            
            # 파일이 없으면 헤더 추가
            file_exists = os.path.exists(filename)
            
            # Excel에서 한글이 제대로 보이도록 UTF-8 BOM 포함하여 저장
            with open(filename, 'a', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)
                
                if not file_exists:
                    writer.writerow(["타임스탬프", "인식된 텍스트", "신뢰도"])
                
                writer.writerow([timestamp, text, confidence])
            
            print(f"✅ 로컬 CSV 파일에 저장 완료: {text[:30]}...")
            
        except Exception as e:
            print(f"로컬 파일 저장 오류: {e}")

def main():
    """메인 함수"""
    try:
        print("프로그램 시작...")
        
        # Tkinter 루트 윈도우 생성
        root = tk.Tk()
        print("Tkinter 창 생성 완료")
        
        # 음성 처리기 초기화
        print("음성 처리기 초기화 중...")
        voice_processor = SimpleVoiceProcessor()
        print("음성 처리기 초기화 완료")
        
        # 설정 관리자 초기화
        print("설정 관리자 초기화 중...")
        settings_manager = SettingsManager()
        print("설정 관리자 초기화 완료")
        
        # 구글 스프레드시트 핸들러 초기화
        print("구글 스프레드시트 핸들러 초기화 중...")
        sheet_handler = GoogleSheetHandler(settings_manager)
        print("구글 스프레드시트 핸들러 초기화 완료")
        
        # GUI 초기화
        print("GUI 초기화 중...")
        gui = SimpleVoiceGUI(root)
        print("GUI 초기화 성공")
        
        # GUI와 음성 처리기 연결
        print("GUI와 음성 처리기 연결 중...")
        gui.set_voice_processor(voice_processor)
        voice_processor.set_gui(gui)
        print("음성 처리기 연결 완료")
        
        # GUI와 스프레드시트 핸들러 연결
        print("GUI와 스프레드시트 핸들러 연결 중...")
        gui.set_sheet_handler(sheet_handler)
        gui.set_settings_manager(settings_manager)
        print("스프레드시트 핸들러 연결 완료")
        
        print("모든 초기화 완료 - GUI 창이 표시됩니다")
        
        # GUI 실행
        root.mainloop()
        
    except Exception as e:
        print(f"프로그램 실행 오류: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()