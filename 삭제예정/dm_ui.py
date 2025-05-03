import os
import sys
import webbrowser
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QListWidget, QLabel, QLineEdit, 
                            QMessageBox, QDialog, QComboBox, QFrame, QSpacerItem,
                            QSizePolicy, QGroupBox)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QIcon
from auth import get_credentials
from googleapiclient.discovery import build

# PyQt5 플랫폼 플러그인 경로 설정
if hasattr(sys, 'frozen'):
    # PyInstaller로 패키징된 경우
    qt_plugin_path = os.path.join(sys._MEIPASS, 'PyQt5', 'Qt5', 'plugins')
else:
    # 일반 Python 환경
    import PyQt5
    qt_plugin_path = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = qt_plugin_path

# 스프레드시트 URL 상수
DM_LIST_SPREADSHEET_ID = '1VhEWeQASyv02knIghpcccYLgWfJCe2ylUnPsQ_-KNAI'
TEMPLATE_SPREADSHEET_ID = '1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0'
DM_LIST_URL = f"https://docs.google.com/spreadsheets/d/{DM_LIST_SPREADSHEET_ID}/edit"
TEMPLATE_URL = f"https://docs.google.com/spreadsheets/d/{TEMPLATE_SPREADSHEET_ID}/edit"

class ProfileSelector(QDialog):
    def __init__(self, user_data_parent):
        super().__init__()
        self.user_data_parent = user_data_parent
        self.selected_profile_path = None
        self.selected_dm_list_sheet = None
        self.selected_template_sheet = None
        self.sheets_service = None
        self.initUI()
        
        # UI 초기화 후 시트 목록 자동 로드 (타이머 사용하여 UI가 먼저 표시된 후 로드)
        QTimer.singleShot(100, self.auto_load_sheets)
        
    def auto_load_sheets(self):
        """시트 목록을 자동으로 가져옴"""
        # 로딩 상태 표시
        self.load_sheets_btn.setText("시트 목록 가져오는 중...")
        self.load_sheets_btn.setEnabled(False)
        
        # 비동기적으로 시트 목록 로드 (UI가 멈추지 않도록)
        QTimer.singleShot(100, self.load_sheet_lists_quietly)
        
    def load_sheet_lists_quietly(self):
        """메시지 없이 조용히 시트 목록을 가져옴"""
        # 서비스 객체 생성
        if not self.get_sheets_service():
            self.load_sheets_btn.setText("시트 목록 가져오기")
            self.load_sheets_btn.setEnabled(True)
            return
            
        try:
            # DM 목록 스프레드시트의 시트 목록 가져오기
            self.dm_sheet_combo.clear()
            dm_sheets = self.get_sheet_names(DM_LIST_SPREADSHEET_ID)
            for sheet in dm_sheets:
                self.dm_sheet_combo.addItem(sheet)
                
            # 기본값 dm_list 또는 첫 번째 시트 선택
            dm_list_index = self.dm_sheet_combo.findText("dm_list")
            if dm_list_index >= 0:
                self.dm_sheet_combo.setCurrentIndex(dm_list_index)
            
            # 템플릿 스프레드시트의 시트 목록 가져오기
            self.template_sheet_combo.clear()
            template_sheets = self.get_sheet_names(TEMPLATE_SPREADSHEET_ID)
            for sheet in template_sheets:
                self.template_sheet_combo.addItem(sheet)
                
            # 기본값 협찬문의 또는 첫 번째 시트 선택
            template_index = self.template_sheet_combo.findText("협찬문의")
            if template_index >= 0:
                self.template_sheet_combo.setCurrentIndex(template_index)
            
            # 버튼 상태 복원    
            self.load_sheets_btn.setText("시트 목록 새로고침")
            self.load_sheets_btn.setEnabled(True)
            
        except Exception as e:
            # 오류 발생 시 상태 복원
            self.load_sheets_btn.setText("시트 목록 가져오기")
            self.load_sheets_btn.setEnabled(True)
            QMessageBox.critical(self, "오류", f"시트 목록을 가져오는 중 오류가 발생했습니다: {e}")
    
    def initUI(self):
        self.setWindowTitle('인스타그램 DM 발송 프로그램')
        self.setMinimumWidth(550)
        self.setMinimumHeight(500)
        
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 제목 레이블
        title_label = QLabel('인스타그램 DM 발송 프로그램')
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)
        
        # 프로필 선택 그룹
        profile_group = QGroupBox("사용할 인스타그램 아이디를 선택하세요")
        profile_layout = QVBoxLayout()
        
        # 설명 레이블
        desc_label = QLabel('인스타 아이디:')
        profile_layout.addWidget(desc_label)
        
        # 프로필 콤보박스
        combo_layout = QHBoxLayout()
        combo_layout.setSpacing(10)
        
        self.profile_combo = QComboBox()
        self.profile_combo.setMinimumHeight(30)
        font = QFont()
        font.setPointSize(10)
        self.profile_combo.setFont(font)
        self.load_profiles()
        combo_layout.addWidget(self.profile_combo, 1)
        
        # 새로고침 버튼
        refresh_btn = QPushButton('새로고침')
        refresh_btn.setMaximumWidth(80)
        refresh_btn.clicked.connect(self.load_profiles)
        combo_layout.addWidget(refresh_btn)
        
        profile_layout.addLayout(combo_layout)
        profile_group.setLayout(profile_layout)
        layout.addWidget(profile_group)
        
        # 스프레드시트 설정 그룹
        sheet_group = QGroupBox("스프레드시트 설정")
        sheet_layout = QVBoxLayout()
        
        # DM 목록 스프레드시트 섹션
        sheet_layout.addWidget(QLabel("DM보낼 명단 시트를 선택해주세요:"))
        
        dm_sheet_layout = QHBoxLayout()
        
        # DM 목록 시트 선택 콤보박스
        self.dm_sheet_combo = QComboBox()
        self.dm_sheet_combo.setMinimumHeight(30)
        self.dm_sheet_combo.setFont(font)
        dm_sheet_layout.addWidget(self.dm_sheet_combo, 1)
        
        # DM 목록 스프레드시트 열기 버튼
        dm_open_btn = QPushButton("시트열기")
        dm_open_btn.setMaximumWidth(80)
        dm_open_btn.clicked.connect(self.open_dm_list)
        dm_sheet_layout.addWidget(dm_open_btn)
        
        sheet_layout.addLayout(dm_sheet_layout)
        
        # 템플릿 스프레드시트 섹션
        sheet_layout.addWidget(QLabel("메시지 템플릿 시트를 선택해주세요:"))
        
        template_sheet_layout = QHBoxLayout()
        
        # 템플릿 시트 선택 콤보박스
        self.template_sheet_combo = QComboBox()
        self.template_sheet_combo.setMinimumHeight(30)
        self.template_sheet_combo.setFont(font)
        template_sheet_layout.addWidget(self.template_sheet_combo, 1)
        
        # 템플릿 스프레드시트 열기 버튼
        template_open_btn = QPushButton("시트열기")
        template_open_btn.setMaximumWidth(80)
        template_open_btn.clicked.connect(self.open_template)
        template_sheet_layout.addWidget(template_open_btn)
        
        sheet_layout.addLayout(template_sheet_layout)
        
        # 시트 목록 로드 버튼
        self.load_sheets_btn = QPushButton("시트 목록 가져오기")
        self.load_sheets_btn.setMinimumHeight(35)
        self.load_sheets_btn.clicked.connect(self.load_sheet_lists)
        sheet_layout.addWidget(self.load_sheets_btn)
        
        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group)
        
        # 여백 추가
        layout.addItem(QSpacerItem(20, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # 버튼 레이아웃
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 선택 버튼
        self.selectBtn = QPushButton('선택 완료')
        self.selectBtn.setMinimumHeight(35)
        self.selectBtn.clicked.connect(self.select_all)
        button_layout.addWidget(self.selectBtn)
        
        # 새 프로필 생성 버튼
        self.newBtn = QPushButton('새 프로필 생성')
        self.newBtn.setMinimumHeight(35)
        self.newBtn.clicked.connect(self.create_new_profile)
        button_layout.addWidget(self.newBtn)
        
        # 취소 버튼
        self.cancelBtn = QPushButton('취소')
        self.cancelBtn.setMinimumHeight(35)
        self.cancelBtn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancelBtn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def get_sheets_service(self):
        """Google Sheets API 서비스 객체를 생성"""
        if self.sheets_service is None:
            try:
                creds = get_credentials()
                self.sheets_service = build('sheets', 'v4', credentials=creds)
                return True
            except Exception as e:
                QMessageBox.critical(self, "인증 오류", f"Google API 인증에 실패했습니다: {e}")
                return False
        return True
    
    def load_sheet_lists(self):
        """두 스프레드시트의 시트 목록을 가져와 콤보박스에 표시"""
        # 서비스 객체 생성
        if not self.get_sheets_service():
            return
            
        # 진행 중 메시지 표시
        QMessageBox.information(self, "정보", "스프레드시트에서 시트 목록을 가져오는 중입니다...\n잠시만 기다려주세요.")
        
        try:
            # DM 목록 스프레드시트의 시트 목록 가져오기
            self.dm_sheet_combo.clear()
            dm_sheets = self.get_sheet_names(DM_LIST_SPREADSHEET_ID)
            for sheet in dm_sheets:
                self.dm_sheet_combo.addItem(sheet)
                
            # 기본값 dm_list 또는 첫 번째 시트 선택
            dm_list_index = self.dm_sheet_combo.findText("dm_list")
            if dm_list_index >= 0:
                self.dm_sheet_combo.setCurrentIndex(dm_list_index)
            
            # 템플릿 스프레드시트의 시트 목록 가져오기
            self.template_sheet_combo.clear()
            template_sheets = self.get_sheet_names(TEMPLATE_SPREADSHEET_ID)
            for sheet in template_sheets:
                self.template_sheet_combo.addItem(sheet)
                
            # 기본값 협찬문의 또는 첫 번째 시트 선택
            template_index = self.template_sheet_combo.findText("협찬문의")
            if template_index >= 0:
                self.template_sheet_combo.setCurrentIndex(template_index)
            
            QMessageBox.information(self, "완료", "시트 목록을 성공적으로 가져왔습니다.")
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"시트 목록을 가져오는 중 오류가 발생했습니다: {e}")
    
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
            QMessageBox.warning(self, "링크 열기 실패", f"스프레드시트를 열지 못했습니다: {e}")
    
    def open_template(self):
        """메시지 템플릿 스프레드시트 열기"""
        try:
            webbrowser.open(TEMPLATE_URL)
        except Exception as e:
            QMessageBox.warning(self, "링크 열기 실패", f"스프레드시트를 열지 못했습니다: {e}")
    
    def load_profiles(self):
        """user_data 폴더 내의 프로필 디렉토리를 스캔하여 콤보박스에 추가"""
        self.profile_combo.clear()
        
        # user_data 폴더가 없으면 생성
        if not os.path.exists(self.user_data_parent):
            os.makedirs(self.user_data_parent)
            self.profile_combo.addItem("프로필이 없습니다. 새 프로필을 생성하세요.")
            return
            
        # 프로필 디렉토리 스캔
        profiles = []
        for item in os.listdir(self.user_data_parent):
            item_path = os.path.join(self.user_data_parent, item)
            # Chrome 프로필 디렉토리인지 확인 (Default 폴더 또는 Profile 폴더가 있는지)
            if os.path.isdir(item_path):
                if (os.path.exists(os.path.join(item_path, 'Default')) or 
                    any(p.startswith('Profile') for p in os.listdir(item_path) if os.path.isdir(os.path.join(item_path, p)))):
                    profiles.append(item)
        
        if profiles:
            for profile in profiles:
                self.profile_combo.addItem(profile)
        else:
            self.profile_combo.addItem("프로필이 없습니다. 새 프로필을 생성하세요.")
    
    def select_all(self):
        """프로필과 시트 선택 완료"""
        # 프로필 선택 확인
        selected_text = self.profile_combo.currentText()
        
        if selected_text == "프로필이 없습니다. 새 프로필을 생성하세요.":
            QMessageBox.warning(self, "선택 오류", "유효한 프로필이 없습니다. 새 프로필을 생성하세요.")
            return
        
        # 선택된 시트가 있는지 확인
        if self.dm_sheet_combo.count() == 0 or self.template_sheet_combo.count() == 0:
            response = QMessageBox.question(
                self, 
                "시트 미선택", 
                "시트 목록을 가져오지 않았습니다. 기본 시트 이름을 사용하시겠습니까?\n(DM 목록: dm_list, 템플릿: 협찬문의)",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if response == QMessageBox.Yes:
                self.selected_dm_list_sheet = "dm_list"
                self.selected_template_sheet = "협찬문의"
            else:
                return
        else:
            self.selected_dm_list_sheet = self.dm_sheet_combo.currentText()
            self.selected_template_sheet = self.template_sheet_combo.currentText()
        
        # 최종 선택 확인 대화상자 표시
        confirm_message = f"다음 설정으로 DM 발송을 시작하시겠습니까?\n\n" \
                        f"• 인스타그램 계정: {selected_text}\n" \
                        f"• DM 목록 시트: {self.selected_dm_list_sheet}\n" \
                        f"• 메시지 템플릿 시트: {self.selected_template_sheet}"
        
        confirm = QMessageBox.question(
            self,
            "설정 확인",
            confirm_message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No  # 기본 버튼은 '아니오'로 설정
        )
        
        if confirm == QMessageBox.Yes:
            self.selected_profile_path = os.path.join(self.user_data_parent, selected_text)
            self.accept()
        # '아니오'를 선택한 경우 아무 작업도 수행하지 않음
    
    def create_new_profile(self):
        dialog = NewProfileDialog(self.user_data_parent)
        if dialog.exec_() == QDialog.Accepted and dialog.new_profile_path:
            # 프로필만 생성하고 종료하지 않음
            self.load_profiles()
            # 새로 생성된 프로필 선택
            index = self.profile_combo.findText(os.path.basename(dialog.new_profile_path))
            if index >= 0:
                self.profile_combo.setCurrentIndex(index)


class NewProfileDialog(QDialog):
    def __init__(self, user_data_parent):
        super().__init__()
        self.user_data_parent = user_data_parent
        self.new_profile_path = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('새 Chrome 프로필 생성')
        self.setMinimumWidth(400)
        self.setMinimumHeight(180)
        
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 제목 레이블
        title_label = QLabel('새 Chrome 프로필 생성')
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 설명 레이블
        desc_label = QLabel('생성할 프로필 이름을 입력하세요:')
        layout.addWidget(desc_label)
        
        # 이름 입력 필드
        self.nameEdit = QLineEdit()
        self.nameEdit.setMinimumHeight(30)
        font = QFont()
        font.setPointSize(10)
        self.nameEdit.setFont(font)
        self.nameEdit.setPlaceholderText('프로필 이름')
        layout.addWidget(self.nameEdit)
        
        # 버튼 레이아웃
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 생성 버튼
        createBtn = QPushButton('생성')
        createBtn.setMinimumHeight(35)
        createBtn.clicked.connect(self.create_profile)
        button_layout.addWidget(createBtn)
        
        # 취소 버튼
        cancelBtn = QPushButton('취소')
        cancelBtn.setMinimumHeight(35)
        cancelBtn.clicked.connect(self.reject)
        button_layout.addWidget(cancelBtn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
    def create_profile(self):
        profile_name = self.nameEdit.text().strip()
        
        # 이름 유효성 검사
        if not profile_name:
            QMessageBox.warning(self, '경고', '프로필 이름을 입력하세요.')
            return
            
        if any(c in r'\/:*?"<>|' for c in profile_name):
            QMessageBox.warning(self, '경고', '프로필 이름에 다음 문자를 사용할 수 없습니다: \\ / : * ? " < > |')
            return
            
        # 이미 존재하는 이름인지 확인
        new_profile_path = os.path.join(self.user_data_parent, profile_name)
        if os.path.exists(new_profile_path):
            QMessageBox.warning(self, '경고', f"'{profile_name}' 프로필이 이미 존재합니다.")
            return
            
        # 새 프로필 폴더 생성
        try:
            os.makedirs(new_profile_path)
            os.makedirs(os.path.join(new_profile_path, 'Default'))
            QMessageBox.information(self, '완료', f"'{profile_name}' 프로필이 생성되었습니다.")
            self.new_profile_path = new_profile_path
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, '오류', f'프로필 생성 중 오류가 발생했습니다: {e}')


def select_profile_gui(user_data_parent):
    """GUI로 프로필과 시트를 선택하거나 생성하고 정보를 반환합니다"""
    app = QApplication.instance() or QApplication(sys.argv)
    selector = ProfileSelector(user_data_parent)
    
    if selector.exec_() == QDialog.Accepted:
        return {
            'profile_path': selector.selected_profile_path,
            'dm_list_sheet': selector.selected_dm_list_sheet,
            'template_sheet': selector.selected_template_sheet
        }
    else:
        return None 