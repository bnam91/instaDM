import tkinter as tk
from tkinter import ttk

class YoutubeTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
        
    def create_widgets(self):
        # 메인 컨테이너 프레임 생성
        main_container = ttk.Frame(self)
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 탭 제목
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill='x', pady=(0, 20))
        
        title_label = ttk.Label(title_frame, 
                              text="유튜브 자동화", 
                              font=('맑은 고딕', 16, 'bold'))
        title_label.pack()
        
        # 구분선
        separator = ttk.Separator(main_container, orient='horizontal')
        separator.pack(fill='x', pady=10)
        
        # 준비 중 메시지
        status_frame = ttk.Frame(main_container)
        status_frame.pack(expand=True)
        
        status_label = ttk.Label(status_frame, 
                               text="유튜브 기능 준비 중...", 
                               font=('맑은 고딕', 12))
        status_label.pack(pady=20) 