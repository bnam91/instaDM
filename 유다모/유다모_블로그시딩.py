import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import os
from datetime import datetime, timedelta
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import gspread
from dotenv import load_dotenv
from auth import get_credentials
import webbrowser  # 웹 브라우저 열기 위한 모듈 추가

# .env 파일 로드
load_dotenv()

class NaverBlogScraper:
    def __init__(self, root):
        self.root = root
        self.root.title("Naver Blog Scraper")
        self.root.geometry("500x200")  # 높이를 조금 늘림
        self.root.resizable(False, False)
        
        self.keywords = None
        self.keywords_worksheet = None
        self.countdown_job = None
        
        # 구글 시트 설정
        self.spreadsheet_id = "1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0"
        self.second_spreadsheet_id = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos"
        self.sheet_name = "블로그_키워드"
        self.result_sheet_name = "블로그_현황판"
        self.second_result_sheet_name = "유다모"  # 두 번째 스프레드시트의 시트 이름을 '유다모'로 변경
        self.sheet_url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit#gid=0"
        
        self.create_widgets()
        # 프로그램 시작 시 자동으로 키워드 로드
        self.load_keywords_from_sheet()
    
    def convert_relative_time_to_date(self, relative_time):
        """상대적 시간 표현을 실제 날짜로 변환"""
        try:
            now = datetime.now()
            
            # 정규표현식으로 숫자와 단위 추출
            pattern = r'(\d+)\s*(시간|일|주|개월|년)\s*전'
            match = re.search(pattern, relative_time)
            
            if match:
                number = int(match.group(1))
                unit = match.group(2)
                
                if unit == '시간':
                    return (now - timedelta(hours=number)).strftime('%Y.%m.%d.')
                elif unit == '일':
                    return (now - timedelta(days=number)).strftime('%Y.%m.%d.')
                elif unit == '주':
                    return (now - timedelta(weeks=number)).strftime('%Y.%m.%d.')
                elif unit == '개월':
                    # 개월은 대략적으로 30일로 계산
                    return (now - timedelta(days=number*30)).strftime('%Y.%m.%d.')
                elif unit == '년':
                    # 년은 대략적으로 365일로 계산
                    return (now - timedelta(days=number*365)).strftime('%Y.%m.%d.')
            
            return relative_time  # 변환할 수 없는 경우 원본 반환
            
        except Exception as e:
            print(f"날짜 변환 중 오류: {str(e)}")
            return relative_time
    
    def parse_date_for_comparison(self, date_str):
        """날짜 문자열을 비교 가능한 datetime 객체로 변환"""
        try:
            # YYYY.MM.DD. 형식 파싱
            if date_str.endswith('.'):
                date_str = date_str[:-1]  # 마지막 점 제거
            
            # 점을 하이픈으로 변경하여 파싱
            date_str = date_str.replace('.', '-')
            return datetime.strptime(date_str, '%Y-%m-%d')
        except:
            # 파싱할 수 없는 경우 현재 날짜 반환
            return datetime.now()
    
    def process_and_filter_data(self, new_data, existing_data):
        """데이터를 처리하고 중복 제거 및 필터링"""
        today = datetime.now()
        
        # 기존 이메일 목록
        existing_emails = {row[5] for row in existing_data if len(row) > 5}
        
        # 새로운 데이터 필터링 (100일 이내, 중복 이메일 제외)
        filtered_new_data = []
        for row in new_data:
            date_str = row[3]
            date_obj = self.parse_date_for_comparison(date_str)
            days_diff = (today - date_obj).days
            
            email = row[5]
            
            if days_diff <= 100 and email not in existing_emails:
                filtered_new_data.append(row)
                existing_emails.add(email) # 중복 추가 방지
        
        # 기존 데이터와 새로운 데이터를 합쳐서 반환
        return filtered_new_data
    
    def is_valid_date_format(self, date_text):
        """날짜 형식인지 확인"""
        # YYYY.MM.DD. 또는 YYYY.MM.DD 형식 확인
        if len(date_text) >= 8 and (date_text.count('.') >= 2 or date_text.count('-') >= 2):
            return True
        
        # 상대적 시간 표현 확인
        relative_pattern = r'(\d+)\s*(시간|일|주|개월|년)\s*전'
        if re.search(relative_pattern, date_text):
            return True
        
        return False
    
    def create_widgets(self):
        # Google Sheets status frame
        status_frame = ttk.Frame(self.root, padding=10)
        status_frame.pack(fill=tk.X)
        
        ttk.Label(status_frame, text="구글 시트에서 키워드를 자동으로 로드합니다").pack(side=tk.LEFT, padx=5)
        
        self.status_label = ttk.Label(status_frame, text="키워드 로딩 중...")
        self.status_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Google Sheet 링크 버튼 프레임
        sheet_link_frame = ttk.Frame(self.root, padding=10)
        sheet_link_frame.pack(fill=tk.X)
        
        sheet_link_button = ttk.Button(
            sheet_link_frame, 
            text="구글 시트 열기", 
            command=self.open_google_sheet,
            style='Link.TButton'
        )
        sheet_link_button.pack(side=tk.LEFT, padx=5)
        
        # Button frame
        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.pack(fill=tk.X)
        
        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_scraping)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Cancel", command=self.root.destroy).pack(side=tk.LEFT, padx=5)

        self.countdown_label = ttk.Label(button_frame, text="")
        self.countdown_label.pack(side=tk.LEFT, padx=15)
    
    def update_countdown(self, count):
        """15초 카운트다운 후 자동 시작"""
        if count >= 0:
            self.countdown_label.config(text=f"자동 시작까지 {count}초...")
            self.countdown_job = self.root.after(1000, self.update_countdown, count - 1)
        else:
            self.countdown_label.config(text="")
            self.start_scraping()  # 카운트다운이 끝나면 바로 스크래핑 시작

    def load_keywords_from_sheet(self):
        """구글 시트에서 키워드를 로드합니다"""
        try:
            # auth.py의 인증 함수 사용
            creds = get_credentials()
            
            # gspread 클라이언트 생성
            client = gspread.authorize(creds)
            
            # 스프레드시트 열기
            spreadsheet = client.open_by_key(self.spreadsheet_id)
            self.keywords_worksheet = spreadsheet.worksheet(self.sheet_name)
            
            # 시트의 모든 데이터 읽기
            all_sheet_data = self.keywords_worksheet.get_all_values()
            
            self.keywords = []
            if len(all_sheet_data) > 1: # 헤더가 있는지 확인
                # 헤더를 제외한 데이터 (2행부터)
                for row in all_sheet_data[1:]:
                    # D열(인덱스 3)이 '시딩'이고 A열(인덱스 0)에 키워드가 있는지 확인
                    if len(row) > 3 and row[3].strip() == '시딩' and row[0].strip():
                        self.keywords.append(row[0].strip())

            if self.keywords:
                self.status_label.config(text=f"'시딩' 키워드 {len(self.keywords)}개 로드 완료")
                self.update_countdown(15) # 카운트다운 시작
            else:
                self.status_label.config(text="'시딩'으로 표시된 키워드를 찾을 수 없습니다")
                
        except Exception as e:
            self.status_label.config(text="키워드 로드 실패")
            messagebox.showerror("Error", f"구글 시트에서 키워드를 읽는 중 오류가 발생했습니다: {str(e)}")
            self.keywords = None
    
    def update_search_date(self, keyword, search_date):
        """키워드 시트의 C열에 검색 날짜를 기록합니다"""
        try:
            # A열에서 해당 키워드의 행 번호 찾기
            all_keywords = self.keywords_worksheet.col_values(1)
            row_number = None
            
            for i, kw in enumerate(all_keywords, 1):
                if kw.strip() == keyword:
                    row_number = i
                    break
            
            if row_number:
                # C열에 검색 날짜 기록 (행 번호는 1부터 시작하므로 그대로 사용)
                self.keywords_worksheet.update_cell(row_number, 3, search_date)
                
        except Exception as e:
            print(f"검색 날짜 기록 중 오류: {str(e)}")
    
    def start_scraping(self):
        # 카운트다운 중지
        if self.countdown_job:
            self.root.after_cancel(self.countdown_job)
            self.countdown_job = None

        self.start_button.config(state=tk.DISABLED)
        self.countdown_label.config(text="데이터 수집 중...")
        self.root.update_idletasks() # UI 즉시 업데이트

        if not self.keywords:
            messagebox.showerror("Error", "'시딩'으로 표시된 키워드를 찾을 수 없습니다.")
            self.start_button.config(state=tk.NORMAL)
            self.countdown_label.config(text="")
            return
            
        self.perform_scraping(40)
    
    def show_auto_closing_info(self, message):
        """5초 후 자동으로 사라지는 알림창"""
        info_window = tk.Toplevel(self.root)
        info_window.title("저장 완료")
        
        # 화면 중앙에 위치
        window_width = 300
        window_height = 100
        screen_width = info_window.winfo_screenwidth()
        screen_height = info_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        info_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 테두리 제거 및 항상 위에 표시
        info_window.overrideredirect(True)
        info_window.attributes('-topmost', True)
        
        # 메시지 표시
        label = ttk.Label(info_window, text=message, wraplength=250, justify='center')
        label.pack(expand=True, padx=20, pady=20)
        
        # 5초 후 창 닫기
        info_window.after(5000, info_window.destroy)

    def perform_scraping(self, num_posts):
        try:
            # Chrome 옵션 설정
            options = Options()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument("--start-maximized")
            
            # Chrome WebDriver 설정
            driver = webdriver.Chrome(options=options)
            
            # 구글 시트 클라이언트 생성
            creds = get_credentials()
            client = gspread.authorize(creds)
            
            # 첫 번째 스프레드시트
            spreadsheet = client.open_by_key(self.spreadsheet_id)
            
            # 두 번째 스프레드시트
            second_spreadsheet = client.open_by_key(self.second_spreadsheet_id)
            
            # 현황판 시트 가져오기 (없으면 생성)
            try:
                result_worksheet = spreadsheet.worksheet(self.result_sheet_name)
            except gspread.WorksheetNotFound:
                result_worksheet = spreadsheet.add_worksheet(title=self.result_sheet_name, rows=1000, cols=10)
            
            # 두 번째 스프레드시트의 시트 가져오기 (없으면 생성)
            try:
                second_result_worksheet = second_spreadsheet.worksheet(self.second_result_sheet_name)
            except gspread.WorksheetNotFound:
                second_result_worksheet = second_spreadsheet.add_worksheet(title=self.second_result_sheet_name, rows=1000, cols=3)
            
            # 첫 번째 시트 헤더 추가 (시트가 비어있는 경우)
            if not result_worksheet.get_all_values():
                headers = ["검색어", "작성자", "제목", "날짜", "URL", "이메일", "추가일"]
                result_worksheet.append_row(headers)
            
            # 두 번째 시트 헤더 추가 (시트가 비어있는 경우)
            if not second_result_worksheet.get_all_values():
                second_headers = ["URL", "작성자", "이메일"]
                second_result_worksheet.append_row(second_headers)
            
            # 기존 데이터 읽기 (헤더 제외)
            existing_data = result_worksheet.get_all_values()[1:] if len(result_worksheet.get_all_values()) > 1 else []
            second_existing_data = second_result_worksheet.get_all_values()[1:] if len(second_result_worksheet.get_all_values()) > 1 else []
            
            # 두 번째 스프레드시트의 기존 URL 목록
            existing_urls = {row[0] for row in second_existing_data}
            
            all_data = []
            second_all_data = []
            current_date = datetime.now().strftime('%Y-%m-%d')
            
            for search_query in self.keywords:
                # 상태 표시만 창 타이틀에 보여줌
                self.root.title(f"'{search_query}' 데이터 수집 중...")
                
                driver.get("https://www.naver.com")
                time.sleep(1)
                
                driver.find_element(By.ID, "query").send_keys(search_query)
                driver.find_element(By.ID, "search-btn").click()
                time.sleep(1)
                
                driver.find_element(By.LINK_TEXT, "블로그").click()
                time.sleep(2)
                
                html = BeautifulSoup(driver.page_source, 'html.parser')
                title_links = html.find_all('a', class_='title_link', limit=num_posts)
                name_links = html.find_all('a', class_='name', limit=num_posts)
                
                # 날짜 추출 방식 개선 - 실제 날짜 형식만 찾기
                date_links = []
                user_info_divs = html.find_all('div', class_='user_info')
                
                for user_info in user_info_divs[:num_posts]:
                    sub_spans = user_info.find_all('span', class_='sub')
                    for span in sub_spans:
                        date_text = span.text.strip()
                        # 날짜 형식 확인 (YYYY.MM.DD. 또는 YYYY.MM.DD 또는 상대적 시간)
                        if self.is_valid_date_format(date_text):
                            date_links.append(span)
                            break
                    else:
                        # 날짜를 찾지 못한 경우 빈 span 추가
                        date_links.append(None)
                
                data = []
                # 각 요소를 개별적으로 처리하여 40개씩 정확히 수집
                for i in range(min(len(title_links), len(name_links), len(date_links))):
                    title_link = title_links[i]
                    name_link = name_links[i]
                    date_link = date_links[i]
                    
                    url = title_link['href']
                    
                    # ader.naver.com으로 시작하는 URL은 제외
                    if url.startswith('https://ader.naver.com/'):
                        continue
                    
                    # 날짜 추출 개선
                    if date_link:
                        date_text = date_link.text.strip()
                        # 상대적 시간을 실제 날짜로 변환
                        date_text = self.convert_relative_time_to_date(date_text)
                    else:
                        date_text = "날짜 없음"
                    
                    email = ''
                    if "https://blog.naver.com/" in url:
                        username = url.split('/')[3]
                        email = f"{username}@naver.com"
                    elif "https://adcr.naver.com/" in url:
                        email = "파워콘텐츠"
                    elif "https://post.naver.com/" in url:
                        email = "포스트"
                    
                    # 첫 번째 스프레드시트용 데이터
                    data.append([search_query, name_link.text.strip(), title_link.text.strip(), date_text, url, email, current_date])
                    
                    # 두 번째 스프레드시트용 데이터 (URL이 중복되지 않는 경우에만)
                    if url not in existing_urls:
                        second_all_data.append([url, name_link.text.strip(), email])
                        existing_urls.add(url)  # URL 중복 체크를 위해 추가
                
                all_data.extend(data)
                
                # 키워드 시트의 C열에 검색 날짜 기록
                self.update_search_date(search_query, current_date)
            
            driver.quit()
            
            # 첫 번째 스프레드시트: 새로운 데이터에서 기존에 없는 이메일만 필터링하고 100일 이내 데이터만 남김
            newly_filtered_data = self.process_and_filter_data(all_data, existing_data)
            
            added_count = 0
            # 첫 번째 스프레드시트에 새로운 데이터 추가
            if newly_filtered_data:
                result_worksheet.append_rows(newly_filtered_data)
                added_count += len(newly_filtered_data)
            
            # 두 번째 스프레드시트에 새로운 데이터 추가 (첫 번째 시트에 추가된 데이터만)
            second_sheet_data = []
            for row in newly_filtered_data:
                # URL(4), 작성자(1), 이메일(5) 순서로 데이터 추출
                second_sheet_data.append([row[4], row[1], row[5]])
            
            if second_sheet_data:
                second_result_worksheet.append_rows(second_sheet_data)
            
            message = f"첫 번째 시트와 두 번째 시트에 각각 {len(newly_filtered_data)}개의 새로운 데이터가 추가되었습니다."
            self.show_auto_closing_info(message)
            
            # 창 제목 원래대로 복구
            self.root.title("Naver Blog Scraper")
            
            # 5초 후 프로그램 종료
            self.root.after(5000, self.root.destroy)
            
        except Exception as e:
            messagebox.showerror("Error", f"데이터 수집 중 오류가 발생했습니다: {str(e)}")
            # 에러 발생 시에도 5초 후 프로그램 종료
            self.root.after(5000, self.root.destroy)
        finally:
            # UI 상태 초기화
            self.start_button.config(state=tk.NORMAL)
            self.countdown_label.config(text="")
            self.root.title("Naver Blog Scraper")

    def open_google_sheet(self):
        """구글 시트를 기본 웹 브라우저로 엽니다"""
        webbrowser.open(self.sheet_url)

if __name__ == "__main__":
    root = tk.Tk()
    app = NaverBlogScraper(root)
    root.mainloop()