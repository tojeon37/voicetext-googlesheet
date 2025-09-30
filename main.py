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
            "last_col": 1,
            "allowed_spreadsheets": ["음성기록"]  # 접근 허용된 스프레드시트 목록
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
        """구글 스프레드시트 설정 (자동 감지 방식)"""
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
            
            # 자동 감지 방식으로 스프레드시트 설정
            print("🔍 스프레드시트 자동 감지 중...")
            self.auto_detect_spreadsheet()
                
        except Exception as e:
            print(f"❌ 구글 시트 연결 실패: {e}")
            print("📁 로컬 CSV 파일로 폴백")
            self.sheet = None
    
    def auto_detect_spreadsheet(self):
        """자동으로 스프레드시트 감지 및 설정"""
        try:
            # 설정에서 허용된 스프레드시트 목록 가져오기
            allowed_spreadsheets = []
            if self.settings_manager:
                allowed_spreadsheets = self.settings_manager.get_setting("allowed_spreadsheets", ["음성기록"])
            else:
                allowed_spreadsheets = ["음성기록"]  # 기본값
            
            print(f"🔒 허용된 스프레드시트: {allowed_spreadsheets}")
            
            # 모든 스프레드시트 가져오기
            all_spreadsheets = self.gc.openall()
            print(f"📊 접근 가능한 스프레드시트: {[s.title for s in all_spreadsheets]}")
            
            # 허용된 스프레드시트 중에서 우선순위에 따라 선택
            priority_names = allowed_spreadsheets.copy()
            
            # 1단계: 우선순위에 따라 스프레드시트 찾기
            for priority_name in priority_names:
                for spreadsheet in all_spreadsheets:
                    if spreadsheet.title == priority_name:
                        self.spreadsheet = spreadsheet
                        print(f"✅ 자동 감지된 스프레드시트: {priority_name}")
                        print("✅ 구글 스프레드시트 연결 성공 (자동 감지)")
                        return
            
            # 2단계: 우선순위에 없으면 허용된 스프레드시트 중 첫 번째 사용
            for spreadsheet in all_spreadsheets:
                if spreadsheet.title in allowed_spreadsheets:
                    self.spreadsheet = spreadsheet
                    print(f"✅ 허용된 스프레드시트 자동 선택: {spreadsheet.title}")
                    print("✅ 구글 스프레드시트 연결 성공 (허용된 목록에서 선택)")
                    return
            
            # 3단계: 허용된 스프레드시트가 없으면 첫 번째 스프레드시트 사용 (경고와 함께)
            if all_spreadsheets:
                self.spreadsheet = all_spreadsheets[0]
                print(f"⚠️ 허용된 스프레드시트가 없어 기본 스프레드시트 사용: {all_spreadsheets[0].title}")
                print("✅ 구글 스프레드시트 연결 성공 (기본 선택)")
            else:
                print("❌ 접근 가능한 스프레드시트가 없습니다.")
                print("📁 로컬 CSV 파일로 폴백")
                self.sheet = None
                
        except Exception as e:
            print(f"❌ 스프레드시트 자동 감지 실패: {e}")
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
        """허용된 스프레드시트 목록만 가져오기 (보안 강화)"""
        try:
            print("🔍 허용된 스프레드시트 목록 가져오기...")
            
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
            
            # 설정에서 허용된 스프레드시트 목록 가져오기
            allowed_spreadsheets = []
            if self.settings_manager:
                allowed_spreadsheets = self.settings_manager.get_setting("allowed_spreadsheets", ["음성기록"])
            else:
                allowed_spreadsheets = ["음성기록"]  # 기본값
            
            print(f"🔒 허용된 스프레드시트: {allowed_spreadsheets}")
            
            # 모든 스프레드시트 가져오기
            all_spreadsheets = self.gc.openall()
            
            # 허용된 스프레드시트만 필터링
            filtered_spreadsheets = []
            for spreadsheet in all_spreadsheets:
                if spreadsheet.title in allowed_spreadsheets:
                    filtered_spreadsheets.append({
                        'title': spreadsheet.title,
                        'id': spreadsheet.id,
                        'spreadsheet': spreadsheet
                    })
                    print(f"  📊 {spreadsheet.title} (ID: {spreadsheet.id}) - 허용됨")
                else:
                    print(f"  🚫 {spreadsheet.title} (ID: {spreadsheet.id}) - 접근 차단됨")
            
            print(f"✅ 허용된 스프레드시트 목록 가져오기 성공: {len(filtered_spreadsheets)}개")
            return filtered_spreadsheets
                
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
    
    def add_allowed_spreadsheet(self, spreadsheet_title):
        """허용된 스프레드시트 목록에 추가 (방법 3A)"""
        try:
            if not self.settings_manager:
                print("❌ settings_manager가 없습니다")
                return False
            
            # 현재 허용된 스프레드시트 목록 가져오기
            current_allowed = self.settings_manager.get_setting("allowed_spreadsheets", ["음성기록"])
            
            # 이미 목록에 있으면 추가하지 않음
            if spreadsheet_title in current_allowed:
                print(f"✅ 스프레드시트가 이미 허용 목록에 있습니다: {spreadsheet_title}")
                return True
            
            # 새 스프레드시트를 목록에 추가
            current_allowed.append(spreadsheet_title)
            self.settings_manager.set_setting("allowed_spreadsheets", current_allowed)
            
            print(f"✅ 허용된 스프레드시트 목록에 추가: {spreadsheet_title}")
            print(f"📋 현재 허용 목록: {current_allowed}")
            return True
            
        except Exception as e:
            print(f"허용된 스프레드시트 추가 오류: {e}")
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