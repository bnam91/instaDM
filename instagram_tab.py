import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import webbrowser
from auth import get_credentials
from googleapiclient.discovery import build

# Google Sheets 시트명 기본값
DM_LIST_DEFAULT = 'dm_list'
TEMPLATE_DEFAULT = '협찬문의'

# 스프레드시트 URL 상수
DM_LIST_SPREADSHEET_ID = '1VhEWeQASyv02knIghpcccYLgWfJCe2ylUnPsQ_-KNAI'
TEMPLATE_SPREADSHEET_ID = '1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0'
DM_LIST_URL = f"https://docs.google.com/spreadsheets/d/{DM_LIST_SPREADSHEET_ID}/edit"
TEMPLATE_URL = f"https://docs.google.com/spreadsheets/d/{TEMPLATE_SPREADSHEET_ID}/edit"

class InstagramTab(ttk.Frame):
    def __init__(self, parent, user_data_parent):
        super().__init__(parent)
        self.user_data_parent = user_data_parent
        self.sheets_service = None
        self.create_widgets()
        
    def create_widgets(self):
        # 메인 컨테이너 프레임 생성
        main_container = ttk.Frame(self)
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        pad = {'padx': 10, 'pady': 5}
        # 프로필 선택
        frame1 = ttk.LabelFrame(main_container, text='사용할 인스타그램 아이디를 선택하세요')
        frame1.pack(fill='x', **pad)
        self.profile_combo = ttk.Combobox(frame1, state='readonly')
        self.profile_combo.pack(side='left', fill='x', expand=True, **pad)
        ttk.Button(frame1, text='새로고침', command=self.load_profiles).pack(side='left', **pad)
        ttk.Button(frame1, text='새 프로필 생성', command=self.create_new_profile).pack(side='left', **pad)

        # 시트명 입력
        frame2 = ttk.LabelFrame(main_container, text='스프레드시트 설정')
        frame2.pack(fill='x', **pad)
        
        # DM 목록 시트 선택
        ttk.Label(frame2, text='DM보낼 명단 시트:').pack(anchor='w', **pad)
        dm_sheet_frame = ttk.Frame(frame2)
        dm_sheet_frame.pack(fill='x', **pad)
        self.dm_sheet_combo = ttk.Combobox(dm_sheet_frame, state='readonly')
        self.dm_sheet_combo.pack(side='left', fill='x', expand=True)
        ttk.Button(dm_sheet_frame, text='시트열기', command=self.open_dm_list).pack(side='left', padx=5)
        
        # 템플릿 시트 선택
        ttk.Label(frame2, text='메시지 템플릿 시트:').pack(anchor='w', **pad)
        template_sheet_frame = ttk.Frame(frame2)
        template_sheet_frame.pack(fill='x', **pad)
        self.template_sheet_combo = ttk.Combobox(template_sheet_frame, state='readonly')
        self.template_sheet_combo.pack(side='left', fill='x', expand=True)
        ttk.Button(template_sheet_frame, text='시트열기', command=self.open_template).pack(side='left', padx=5)
        
        # 시트 목록 로드 버튼
        self.load_sheets_btn = ttk.Button(frame2, text="시트 목록 가져오기", command=self.load_sheet_lists)
        self.load_sheets_btn.pack(fill='x', **pad)

        # 버튼
        frame3 = ttk.Frame(main_container)
        frame3.pack(fill='x', pady=15)
        ttk.Button(frame3, text='선택 완료', command=self.select_all).pack(side='left', expand=True, fill='x', padx=10)
        ttk.Button(frame3, text='취소', command=self.cancel).pack(side='right', expand=True, fill='x', padx=10)
        
        # 초기 데이터 로드
        self.load_profiles()
        self.after(100, self.auto_load_sheets)

    def auto_load_sheets(self):
        """시트 목록을 자동으로 가져옴"""
        self.load_sheets_btn.config(text="시트 목록 가져오는 중...", state='disabled')
        self.after(100, self.load_sheet_lists_quietly)
        
    def load_sheet_lists_quietly(self):
        """메시지 없이 조용히 시트 목록을 가져옴"""
        if not self.get_sheets_service():
            self.load_sheets_btn.config(text="시트 목록 가져오기", state='normal')
            return
            
        try:
            # DM 목록 스프레드시트의 시트 목록 가져오기
            self.dm_sheet_combo['values'] = []
            dm_sheets = self.get_sheet_names(DM_LIST_SPREADSHEET_ID)
            self.dm_sheet_combo['values'] = dm_sheets
            
            # 기본값 dm_list 또는 첫 번째 시트 선택
            if "dm_list" in dm_sheets:
                self.dm_sheet_combo.set("dm_list")
            
            # 템플릿 스프레드시트의 시트 목록 가져오기
            self.template_sheet_combo['values'] = []
            template_sheets = self.get_sheet_names(TEMPLATE_SPREADSHEET_ID)
            self.template_sheet_combo['values'] = template_sheets
            
            # 기본값 협찬문의 또는 첫 번째 시트 선택
            if "협찬문의" in template_sheets:
                self.template_sheet_combo.set("협찬문의")
            
            # 버튼 상태 복원    
            self.load_sheets_btn.config(text="시트 목록 새로고침", state='normal')
            
        except Exception as e:
            self.load_sheets_btn.config(text="시트 목록 가져오기", state='normal')
            messagebox.showerror("오류", f"시트 목록을 가져오는 중 오류가 발생했습니다: {e}")

    def get_sheets_service(self):
        """Google Sheets API 서비스 객체를 생성"""
        if self.sheets_service is None:
            try:
                creds = get_credentials()
                self.sheets_service = build('sheets', 'v4', credentials=creds)
                return True
            except Exception as e:
                messagebox.showerror("인증 오류", f"Google API 인증에 실패했습니다: {e}")
                return False
        return True
    
    def load_sheet_lists(self):
        """두 스프레드시트의 시트 목록을 가져와 콤보박스에 표시"""
        if not self.get_sheets_service():
            return
            
        messagebox.showinfo("정보", "스프레드시트에서 시트 목록을 가져오는 중입니다...\n잠시만 기다려주세요.")
        
        try:
            # DM 목록 스프레드시트의 시트 목록 가져오기
            self.dm_sheet_combo['values'] = []
            dm_sheets = self.get_sheet_names(DM_LIST_SPREADSHEET_ID)
            self.dm_sheet_combo['values'] = dm_sheets
            
            # 기본값 dm_list 또는 첫 번째 시트 선택
            if "dm_list" in dm_sheets:
                self.dm_sheet_combo.set("dm_list")
            
            # 템플릿 스프레드시트의 시트 목록 가져오기
            self.template_sheet_combo['values'] = []
            template_sheets = self.get_sheet_names(TEMPLATE_SPREADSHEET_ID)
            self.template_sheet_combo['values'] = template_sheets
            
            # 기본값 협찬문의 또는 첫 번째 시트 선택
            if "협찬문의" in template_sheets:
                self.template_sheet_combo.set("협찬문의")
            
            messagebox.showinfo("완료", "시트 목록을 성공적으로 가져왔습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"시트 목록을 가져오는 중 오류가 발생했습니다: {e}")
    
    def get_sheet_names(self, spreadsheet_id):
        """주어진 스프레드시트 ID의 모든 시트 이름을 가져옴"""
        sheet_metadata = self.sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        return [sheet.get("properties", {}).get("title", "") for sheet in sheets]
    
    def open_dm_list(self):
        """DM 목록 스프레드시트 열기"""
        try:
            webbrowser.open(DM_LIST_URL)
        except Exception as e:
            messagebox.showwarning("링크 열기 실패", f"스프레드시트를 열지 못했습니다: {e}")
    
    def open_template(self):
        """메시지 템플릿 스프레드시트 열기"""
        try:
            webbrowser.open(TEMPLATE_URL)
        except Exception as e:
            messagebox.showwarning("링크 열기 실패", f"스프레드시트를 열지 못했습니다: {e}")

    def load_profiles(self):
        self.profile_combo['values'] = []
        if not os.path.exists(self.user_data_parent):
            os.makedirs(self.user_data_parent)
            self.profile_combo['values'] = ['프로필이 없습니다. 새 프로필을 생성하세요.']
            self.profile_combo.current(0)
            return
        profiles = []
        for item in os.listdir(self.user_data_parent):
            item_path = os.path.join(self.user_data_parent, item)
            if os.path.isdir(item_path):
                if (os.path.exists(os.path.join(item_path, 'Default')) or 
                    any(p.startswith('Profile') for p in os.listdir(item_path) if os.path.isdir(os.path.join(item_path, p)))):
                    profiles.append(item)
        if profiles:
            self.profile_combo['values'] = profiles
            self.profile_combo.current(0)
        else:
            self.profile_combo['values'] = ['프로필이 없습니다. 새 프로필을 생성하세요.']
            self.profile_combo.current(0)

    def create_new_profile(self):
        name = simpledialog.askstring('새 프로필 생성', '생성할 프로필 이름을 입력하세요:', parent=self)
        if not name:
            return
        if any(c in r'\\/:*?"<>|' for c in name):
            messagebox.showwarning('경고', '프로필 이름에 다음 문자를 사용할 수 없습니다: \\ / : * ? " < > |')
            return
        new_profile_path = os.path.join(self.user_data_parent, name)
        if os.path.exists(new_profile_path):
            messagebox.showwarning('경고', f"'{name}' 프로필이 이미 존재합니다.")
            return
        try:
            os.makedirs(new_profile_path)
            os.makedirs(os.path.join(new_profile_path, 'Default'))
            messagebox.showinfo('완료', f"'{name}' 프로필이 생성되었습니다.")
            self.load_profiles()
            idx = self.profile_combo['values'].index(name)
            self.profile_combo.current(idx)
        except Exception as e:
            messagebox.showerror('오류', f'프로필 생성 중 오류가 발생했습니다: {e}')

    def select_all(self):
        selected = self.profile_combo.get()
        if selected == '프로필이 없습니다. 새 프로필을 생성하세요.':
            messagebox.showwarning('선택 오류', '유효한 프로필이 없습니다. 새 프로필을 생성하세요.')
            return
        dm_sheet = self.dm_sheet_combo.get().strip() or DM_LIST_DEFAULT
        template_sheet = self.template_sheet_combo.get().strip() or TEMPLATE_DEFAULT
        confirm = messagebox.askyesno('설정 확인', f"다음 설정으로 DM 발송을 시작하시겠습니까?\n\n• 인스타그램 계정: {selected}\n• DM 목록 시트: {dm_sheet}\n• 메시지 템플릿 시트: {template_sheet}")
        if confirm:
            self.selected_profile_path = os.path.join(self.user_data_parent, selected)
            self.selected_dm_list_sheet = dm_sheet
            self.selected_template_sheet = template_sheet
            self.master.destroy()

    def cancel(self):
        self.selected_profile_path = None
        self.selected_dm_list_sheet = None
        self.selected_template_sheet = None
        self.master.destroy() 