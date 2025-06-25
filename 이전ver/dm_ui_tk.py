import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from instagram_tab import InstagramTab
from naver_tab import NaverTab
from youtube_tab import YoutubeTab

class ProfileSelectorTk(tk.Tk):
    def __init__(self, user_data_parent):
        super().__init__()
        self.title('인스타그램 DM 발송 프로그램')
        
        # 기본 창 크기 설정
        self.geometry('800x600')
        
        # 창 크기 조절 가능하도록 설정
        self.resizable(True, True)
        
        # 전체화면 토글 상태 변수
        self.fullscreen_state = False
        
        # 전체화면 토글 단축키 바인딩
        self.bind('<F11>', self.toggle_fullscreen)
        self.bind('<Escape>', self.end_fullscreen)
        
        self.user_data_parent = user_data_parent
        self.selected_profile_path = None
        self.selected_dm_list_sheet = None
        self.selected_template_sheet = None
        
        # 탭 컨트롤 생성
        self.tab_control = ttk.Notebook(self)
        
        # 각 탭 생성
        self.instagram_tab = InstagramTab(self.tab_control, user_data_parent)
        self.tab_control.add(self.instagram_tab, text='인스타그램')
        
        self.naver_tab = NaverTab(self.tab_control)
        self.tab_control.add(self.naver_tab, text='네이버')
        
        self.youtube_tab = YoutubeTab(self.tab_control)
        self.tab_control.add(self.youtube_tab, text='유튜브')
        
        self.tab_control.pack(expand=1, fill="both")
        
        # 창 크기 변경 이벤트 바인딩
        self.bind('<Configure>', self.on_window_resize)

    def toggle_fullscreen(self, event=None):
        """전체화면 토글"""
        self.fullscreen_state = not self.fullscreen_state
        self.attributes('-fullscreen', self.fullscreen_state)
        return "break"

    def end_fullscreen(self, event=None):
        """전체화면 종료"""
        self.fullscreen_state = False
        self.attributes('-fullscreen', False)
        return "break"

    def on_window_resize(self, event=None):
        """창 크기 변경 시 호출되는 함수"""
        pass

def select_profile_gui(user_data_parent):
    app = ProfileSelectorTk(user_data_parent)
    app.mainloop()
    if app.instagram_tab.selected_profile_path:
        return {
            'profile_path': app.instagram_tab.selected_profile_path,
            'dm_list_sheet': app.instagram_tab.selected_dm_list_sheet,
            'template_sheet': app.instagram_tab.selected_template_sheet
        }
    else:
        return None 