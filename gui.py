import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import time

class SimpleVoiceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("음성 인식 프로그램")
        self.root.geometry("500x600")
        
        
        # 상태 변수
        self.is_recording = False
        self.recording_thread = None
        self.voice_processor = None
        self.sheet_handler = None
        self.settings_manager = None
        
        # 셀 위치 추적 변수
        self.current_row = 1
        self.current_col = 1
        
        # GUI 구성 요소 생성
        self.setup_gui()
        
        # 버튼 스타일 설정 (폰트 검정색)
        self.setup_styles()
        
        # 녹음 상태 표시를 위한 깜빡임 스레드
        self.blink_thread = None
        self.blink_active = False
        
        # 타이머 관련 변수
        self.timer_thread = None
        self.timer_active = False
        self.remaining_seconds = 15
        
    def setup_gui(self):
        """GUI 구성 요소 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="🎙️ 음성 인식 프로그램", 
                              font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 상태 표시 라벨 (깜빡임 효과용)
        self.status_label = ttk.Label(main_frame, text="준비됨", 
                                    font=("Arial", 12), foreground="green")
        self.status_label.grid(row=1, column=0, columnspan=2, pady=(0, 10))
        
        # 타이머 표시 라벨 (작은 글씨로 수정)
        self.timer_label = ttk.Label(main_frame, text="", 
                                   font=("Arial", 9), foreground="blue")
        self.timer_label.grid(row=2, column=0, columnspan=2, pady=(0, 5))
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(0, 20))
        
        # 녹음 시작 버튼 (크기 1.2배, 폰트 검정색)
        self.start_button = ttk.Button(button_frame, text="🎙️ 녹음 시작", 
                                      command=self.start_recording,
                                      style="Start.TButton")
        self.start_button.grid(row=0, column=0, padx=(0, 10), ipadx=20, ipady=10)
        
        # 녹음 중지 버튼 (크기 1.2배, 폰트 검정색)
        self.stop_button = ttk.Button(button_frame, text="⏹️ 녹음 중지", 
                                    command=self.stop_recording,
                                    state="disabled",
                                    style="Stop.TButton")
        self.stop_button.grid(row=0, column=1, ipadx=20, ipady=10)
        
        # 인식된 텍스트 표시 영역 (높이를 더 줄임)
        text_frame = ttk.LabelFrame(main_frame, text="인식된 텍스트", padding="5")
        text_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 스프레드시트 선택 프레임 (시트 선택 위에 추가)
        spreadsheet_frame = ttk.LabelFrame(main_frame, text="스프레드시트 선택", padding="5")
        spreadsheet_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 스프레드시트 선택 드롭다운
        ttk.Label(spreadsheet_frame, text="스프레드시트:").grid(row=0, column=0, padx=(0, 5))
        self.spreadsheet_var = tk.StringVar()
        self.spreadsheet_combo = ttk.Combobox(spreadsheet_frame, textvariable=self.spreadsheet_var, width=25, state="readonly")
        self.spreadsheet_combo.grid(row=0, column=1, padx=(0, 10))
        self.spreadsheet_combo.bind("<<ComboboxSelected>>", self.on_spreadsheet_selected)
        
        # 스프레드시트 새로고침 버튼
        refresh_spreadsheet_button = ttk.Button(spreadsheet_frame, text="새로고침", command=self.refresh_spreadsheets)
        refresh_spreadsheet_button.grid(row=0, column=2)
        
        # 시트 선택 프레임
        sheet_frame = ttk.LabelFrame(main_frame, text="시트 선택", padding="5")
        sheet_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 시트 선택 드롭다운
        ttk.Label(sheet_frame, text="시트:").grid(row=0, column=0, padx=(0, 5))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combo.grid(row=0, column=1, padx=(0, 10))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # 시트 새로고침 버튼
        refresh_button = ttk.Button(sheet_frame, text="새로고침", command=self.refresh_sheets)
        refresh_button.grid(row=0, column=2)
        
        # 스크롤 가능한 텍스트 위젯 (기존 크기로 복원)
        self.text_area = scrolledtext.ScrolledText(text_frame, height=8, width=50,
                                                  font=("Arial", 10), wrap=tk.WORD)
        self.text_area.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 셀 주소 입력 프레임 (아래쪽에 배치)
        cell_input_frame = ttk.LabelFrame(main_frame, text="셀 주소 입력", padding="5")
        cell_input_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 10))
        
        # 셀 주소 입력란
        ttk.Label(cell_input_frame, text="셀 주소:").grid(row=0, column=0, padx=(0, 5))
        self.cell_address_entry = ttk.Entry(cell_input_frame, width=15, font=("Arial", 10))
        self.cell_address_entry.grid(row=0, column=1, padx=(0, 10))
        self.cell_address_entry.insert(0, "A1")  # 기본값
        
        # 현재 입력 위치 표시
        ttk.Label(cell_input_frame, text="현재 위치:").grid(row=0, column=2, padx=(10, 5))
        self.current_cell_label = ttk.Label(cell_input_frame, text="A1", 
                                          font=("Arial", 10, "bold"), foreground="blue")
        self.current_cell_label.grid(row=0, column=3)
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)  # 텍스트 영역이 row=4
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # 버튼 스타일 설정
        self.setup_styles()
        
    def setup_styles(self):
        """버튼 스타일 설정"""
        style = ttk.Style()
        
        # 녹음 시작 버튼 스타일 (초록색)
        style.configure("Start.TButton", foreground="white", background="green")
        
        # 녹음 중지 버튼 스타일 (빨간색)
        style.configure("Stop.TButton", foreground="white", background="red")
        
    def set_voice_processor(self, voice_processor):
        """음성 처리기 설정"""
        self.voice_processor = voice_processor
        
    def set_sheet_handler(self, sheet_handler):
        """스프레드시트 핸들러 설정"""
        self.sheet_handler = sheet_handler
        # 스프레드시트 목록 초기화
        self.refresh_spreadsheets()
        # 시트 목록 초기화
        self.refresh_sheets()
    
    def set_settings_manager(self, settings_manager):
        """설정 관리자 설정"""
        self.settings_manager = settings_manager
        print(f"✅ settings_manager 설정 완료: {settings_manager is not None}")
        # 설정값 복원을 위해 refresh_sheets 호출
        if self.sheet_handler:
            self.refresh_sheets()
        
    def start_recording(self):
        """녹음 시작"""
        if self.is_recording:
            return
            
        self.is_recording = True
        
        # 버튼 상태 변경
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        # 상태 표시 업데이트
        self.update_status("🎙️ 녹음 중...", "red")
        
        # 깜빡임 효과 시작
        self.start_blinking()
        
        # 타이머 시작
        self.start_timer()
        
        # 음성 처리기로 녹음 시작
        if self.voice_processor:
            self.voice_processor.start_recording()
        else:
            print("❌ 음성 처리기가 설정되지 않았습니다!")
            self.reset_buttons()
            
    def stop_recording(self):
        """녹음 중지"""
        if not self.is_recording:
            return
            
        self.is_recording = False
        
        # 음성 처리기로 녹음 중지
        if self.voice_processor:
            self.voice_processor.stop_recording()
        else:
            print("❌ 음성 처리기가 설정되지 않았습니다!")
            self.reset_buttons()
            
    def start_blinking(self):
        """깜빡임 효과 시작"""
        self.blink_active = True
        self.blink_thread = threading.Thread(target=self.blink_effect, daemon=True)
        self.blink_thread.start()
        
    def stop_blinking(self):
        """깜빡임 효과 중지"""
        self.blink_active = False
        
    def start_timer(self):
        """타이머 시작"""
        self.timer_active = True
        self.remaining_seconds = 15
        self.timer_thread = threading.Thread(target=self.timer_countdown, daemon=True)
        self.timer_thread.start()
        
    def stop_timer(self):
        """타이머 중지"""
        self.timer_active = False
        self.timer_label.config(text="")
        
    def timer_countdown(self):
        """타이머 카운트다운"""
        while self.timer_active and self.remaining_seconds > 0:
            self.timer_label.config(text=f"⏱️ 남은 시간: {self.remaining_seconds}초")
            time.sleep(1)
            if self.timer_active:
                self.remaining_seconds -= 1
        
        if self.timer_active:
            # 15초가 완료되면 타이머 숨기기
            self.timer_label.config(text="")
        
    def blink_effect(self):
        """깜빡임 효과 구현"""
        while self.blink_active:
            self.status_label.config(foreground="red")
            time.sleep(0.5)
            if not self.blink_active:
                break
            self.status_label.config(foreground="white")
            time.sleep(0.5)
            
    def update_status(self, message, color="black"):
        """상태 메시지 업데이트"""
        self.status_label.config(text=message, foreground=color)
        
    def display_result(self, text, confidence=None):
        """인식 결과 표시"""
        timestamp = time.strftime("%H:%M:%S")
        
        if confidence:
            result_text = f"[{timestamp}] {text} (신뢰도: {confidence:.2f})\n"
        else:
            result_text = f"[{timestamp}] {text}\n"
            
        # 텍스트 영역에 결과 추가
        self.text_area.insert(tk.END, result_text)
        self.text_area.see(tk.END)  # 자동 스크롤
        
        # 스프레드시트에 저장
        if self.sheet_handler:
            # 셀 주소 입력란에서 현재 셀 주소 가져오기
            current_cell = self.cell_address_entry.get().strip().upper()
            if not current_cell:
                current_cell = "A1"  # 기본값
            
            self.sheet_handler.save_to_sheet(text, confidence or 0.0, current_cell)
            
            # 셀 주소 설정 저장
            if self.settings_manager:
                self.settings_manager.set_setting("last_cell", current_cell)
                print(f"✅ 셀 주소 설정 저장: {current_cell}")
            
            # 다음 행으로 자동 이동
            self.move_to_next_cell()
            
        # 상태 업데이트
        self.update_status("✅ 인식 완료", "green")
        
    def reset_buttons(self):
        """버튼 상태 초기화"""
        self.is_recording = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.update_status("준비됨", "green")
        self.stop_blinking()
        self.stop_timer()
    
    
    def get_current_cell_position(self):
        """현재 셀 위치를 문자열로 반환"""
        # 열 번호를 문자로 변환 (1=A, 2=B, 3=C, ...)
        col_letter = chr(ord('A') + self.current_col - 1)
        return f"{col_letter}{self.current_row}"
    
    def move_to_next_cell(self):
        """다음 행으로 이동하고 입력란 업데이트"""
        current_cell = self.cell_address_entry.get().strip().upper()
        if current_cell:
            # 셀 위치 파싱 (예: A1, B5, C10)
            import re
            match = re.match(r'([A-Z]+)(\d+)', current_cell)
            if match:
                col_letter = match.group(1)
                row_num = int(match.group(2))
                
                # 다음 행으로 이동
                next_row = row_num + 1
                next_cell = f"{col_letter}{next_row}"
                
                # 입력란과 라벨 업데이트
                self.cell_address_entry.delete(0, tk.END)
                self.cell_address_entry.insert(0, next_cell)
                self.current_cell_label.config(text=next_cell)
    
    def restore_last_settings(self):
        """이전 설정값 복원"""
        try:
            if not self.settings_manager:
                return
            
            # 마지막 사용 시트와 셀 주소 정보 출력
            last_sheet = self.settings_manager.get_setting("last_sheet", "시트1")
            last_cell = self.settings_manager.get_setting("last_cell", "A1")
            print(f"🔄 이전 설정값 확인: 시트={last_sheet}, 셀={last_cell}")
            print("✅ 실제 복원은 refresh_sheets에서 처리됩니다")
            
        except Exception as e:
            print(f"❌ 설정값 복원 실패: {e}")
    
    def refresh_spreadsheets(self):
        """스프레드시트 목록 새로고침"""
        print("🔄 refresh_spreadsheets 메서드 호출됨")
        try:
            if not self.sheet_handler:
                print("❌ sheet_handler가 없습니다")
                return
            
            # GUI 상태 업데이트
            self.update_status("🔄 스프레드시트 목록 로딩 중...", "blue")
            
            # 모든 스프레드시트 가져오기
            all_spreadsheets = self.sheet_handler.get_all_spreadsheets()
            print(f"🔄 가져온 스프레드시트 목록: {[spreadsheet['title'] for spreadsheet in all_spreadsheets]}")
            
            if all_spreadsheets:
                spreadsheet_titles = [spreadsheet['title'] for spreadsheet in all_spreadsheets]
                self.spreadsheet_combo['values'] = spreadsheet_titles
                
                # 이전 설정값이 있으면 그것을 사용, 없으면 첫 번째 스프레드시트 선택
                if self.settings_manager:
                    last_spreadsheet = self.settings_manager.get_setting("last_spreadsheet")
                    print(f"🔄 스프레드시트 복원 시도: {last_spreadsheet}")
                    if last_spreadsheet and last_spreadsheet in spreadsheet_titles:
                        self.spreadsheet_var.set(last_spreadsheet)
                        # 선택된 스프레드시트로 변경
                        self.sheet_handler.set_target_spreadsheet(last_spreadsheet)
                        print(f"✅ 스프레드시트 복원 완료: {last_spreadsheet}")
                    elif spreadsheet_titles:
                        self.spreadsheet_var.set(spreadsheet_titles[0])
                        # 첫 번째 스프레드시트로 변경
                        self.sheet_handler.set_target_spreadsheet(spreadsheet_titles[0])
                        print(f"✅ 기본 스프레드시트 선택: {spreadsheet_titles[0]}")
                elif spreadsheet_titles:
                    self.spreadsheet_var.set(spreadsheet_titles[0])
                    # 첫 번째 스프레드시트로 변경
                    self.sheet_handler.set_target_spreadsheet(spreadsheet_titles[0])
                    print(f"✅ 기본 스프레드시트 선택: {spreadsheet_titles[0]}")
                
                print(f"✅ 스프레드시트 목록 새로고침 완료: {spreadsheet_titles}")
                self.update_status(f"✅ {len(spreadsheet_titles)}개 스프레드시트 로드 완료", "green")
            else:
                print("❌ 스프레드시트 목록을 가져올 수 없습니다")
                self.spreadsheet_combo['values'] = []
                self.update_status("❌ 스프레드시트 목록을 가져올 수 없습니다", "red")
                
        except Exception as e:
            print(f"스프레드시트 목록 새로고침 오류: {e}")
            self.update_status(f"❌ 스프레드시트 목록 오류: {str(e)[:30]}...", "red")

    def refresh_sheets(self):
        """시트 목록 새로고침"""
        print("🔄 refresh_sheets 메서드 호출됨")
        try:
            if not self.sheet_handler:
                print("❌ sheet_handler가 없습니다")
                return
            
            # 모든 시트 가져오기
            all_sheets = self.sheet_handler.get_all_sheets()
            print(f"🔄 가져온 시트 목록: {[sheet['title'] for sheet in all_sheets]}")
            
            if all_sheets:
                sheet_titles = [sheet['title'] for sheet in all_sheets]
                self.sheet_combo['values'] = sheet_titles
                
                # 이전 설정값이 있으면 그것을 사용, 없으면 첫 번째 시트 선택
                print(f"🔄 settings_manager 상태: {self.settings_manager is not None}")
                if self.settings_manager:
                    last_sheet = self.settings_manager.get_setting("last_sheet")
                    print(f"🔄 시트 복원 시도: {last_sheet}")
                    print(f"🔄 사용 가능한 시트: {sheet_titles}")
                    if last_sheet and last_sheet in sheet_titles:
                        self.sheet_var.set(last_sheet)
                        self.on_sheet_selected(None)
                        print(f"✅ 시트 복원 완료: {last_sheet}")
                    elif sheet_titles:
                        self.sheet_var.set(sheet_titles[0])
                        self.on_sheet_selected(None)
                        print(f"✅ 기본 시트 선택: {sheet_titles[0]}")
                    
                    # 셀 주소 복원
                    last_cell = self.settings_manager.get_setting("last_cell", "A1")
                    print(f"🔄 셀 주소 복원 시도: {last_cell}")
                    if last_cell and hasattr(self, 'cell_address_entry'):
                        self.cell_address_entry.delete(0, tk.END)
                        self.cell_address_entry.insert(0, last_cell)
                        self.current_cell_label.config(text=last_cell)
                        print(f"✅ 셀 주소 복원 완료: {last_cell}")
                    else:
                        print(f"❌ 셀 주소 복원 실패: entry 위젯이 없음")
                else:
                    print("❌ settings_manager가 None입니다")
                    if sheet_titles:
                        self.sheet_var.set(sheet_titles[0])
                        self.on_sheet_selected(None)
                
                print(f"✅ 시트 목록 새로고침 완료: {sheet_titles}")
            else:
                print("❌ 시트 목록을 가져올 수 없습니다")
                self.sheet_combo['values'] = []
                
        except Exception as e:
            print(f"시트 목록 새로고침 오류: {e}")
    
    def on_spreadsheet_selected(self, event):
        """스프레드시트 선택 이벤트 처리"""
        try:
            selected_spreadsheet = self.spreadsheet_var.get()
            if selected_spreadsheet and self.sheet_handler:
                success = self.sheet_handler.set_target_spreadsheet(selected_spreadsheet)
                if success:
                    print(f"✅ 스프레드시트 변경 완료: {selected_spreadsheet}")
                    # 스프레드시트 변경 시 시트 목록 새로고침
                    self.refresh_sheets()
                    # 스프레드시트 변경 시 설정 저장
                    if self.settings_manager:
                        self.settings_manager.set_setting("last_spreadsheet", selected_spreadsheet)
                        print(f"✅ 스프레드시트 설정 저장: {selected_spreadsheet}")
                else:
                    print(f"❌ 스프레드시트 변경 실패: {selected_spreadsheet}")
        except Exception as e:
            print(f"스프레드시트 선택 오류: {e}")

    def on_sheet_selected(self, event):
        """시트 선택 이벤트 처리"""
        try:
            selected_sheet = self.sheet_var.get()
            if selected_sheet and self.sheet_handler:
                success = self.sheet_handler.set_target_sheet(selected_sheet)
                if success:
                    print(f"✅ 시트 변경 완료: {selected_sheet}")
                    # 시트 변경 시 설정 저장
                    if self.settings_manager:
                        self.settings_manager.set_setting("last_sheet", selected_sheet)
                        print(f"✅ 시트 설정 저장: {selected_sheet}")
                    # 셀 주소는 이전 설정값을 유지 (초기화하지 않음)
                else:
                    print(f"❌ 시트 변경 실패: {selected_sheet}")
        except Exception as e:
            print(f"시트 선택 오류: {e}")
    
    def setup_styles(self):
        """버튼 스타일 설정"""
        style = ttk.Style()
        
        # 녹음 시작 버튼 스타일 (폰트 검정색)
        style.configure("Start.TButton", 
                       foreground="black",
                       font=("Arial", 10, "bold"))
        
        # 녹음 중지 버튼 스타일 (폰트 검정색)
        style.configure("Stop.TButton", 
                       foreground="black",
                       font=("Arial", 10, "bold"))

def main():
    """GUI 테스트용 메인 함수"""
    root = tk.Tk()
    app = SimpleVoiceGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()