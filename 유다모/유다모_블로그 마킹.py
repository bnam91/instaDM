import tkinter as tk
from tkinter import messagebox, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import gspread
from auth import get_credentials
import webbrowser

class NaverBlogScraper:
    def __init__(self, root):
        self.root = root
        self.root.title("네이버 블로그 크롤러")
        self.root.geometry("400x150")
        self.root.resizable(False, False)
        
        self.keywords = None
        
        # 새로운 구글 시트 설정
        self.spreadsheet_id = "1Yve6JJzgJaD4KXe2a8vB53zaWOUGsubcG_1t7XeNVmQ"
        self.sheet_name = "블로그"
        self.sheet_url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit#gid=0"
        
        self.create_widgets()
        # 프로그램 시작 시 자동으로 키워드 로드
        self.load_keywords_from_sheet()
    
    def convert_relative_time_to_date(self, relative_time):
        """상대적 시간 표현을 실제 날짜로 변환"""
        try:
            from datetime import datetime, timedelta
            import re
            
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
                    return (now - timedelta(days=number*30)).strftime('%Y.%m.%d.')
                elif unit == '년':
                    return (now - timedelta(days=number*365)).strftime('%Y.%m.%d.')
            
            return relative_time  # 변환할 수 없는 경우 원본 반환
            
        except Exception as e:
            print(f"날짜 변환 중 오류: {str(e)}")
            return relative_time
    
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
            command=self.open_google_sheet
        )
        sheet_link_button.pack(side=tk.LEFT, padx=5)
        
        # Button frame
        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.pack(fill=tk.X)
        
        self.start_button = ttk.Button(button_frame, text="크롤링 시작", command=self.start_scraping)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="종료", command=self.root.destroy).pack(side=tk.LEFT, padx=5)

        self.countdown_label = ttk.Label(button_frame, text="")
        self.countdown_label.pack(side=tk.LEFT, padx=15)
    
    def load_keywords_from_sheet(self):
        """구글 시트에서 키워드를 로드합니다"""
        try:
            # auth.py의 인증 함수 사용
            creds = get_credentials()
            
            # gspread 클라이언트 생성
            client = gspread.authorize(creds)
            
            # 스프레드시트 열기
            spreadsheet = client.open_by_key(self.spreadsheet_id)
            keywords_worksheet = spreadsheet.worksheet(self.sheet_name)
            
            # 시트의 모든 데이터 읽기
            all_sheet_data = keywords_worksheet.get_all_values()
            
            self.keywords = []
            if len(all_sheet_data) > 1: # 헤더가 있는지 확인
                # 헤더를 제외한 데이터 (2행부터)
                for row in all_sheet_data[1:]:
                    # A열(인덱스 0)에 키워드가 있는지 확인
                    if len(row) > 0 and row[0].strip():
                        self.keywords.append(row[0].strip())

            if self.keywords:
                self.status_label.config(text=f"키워드 {len(self.keywords)}개 로드 완료")
            else:
                self.status_label.config(text="키워드를 찾을 수 없습니다")
                
        except Exception as e:
            self.status_label.config(text="키워드 로드 실패")
            messagebox.showerror("Error", f"구글 시트에서 키워드를 읽는 중 오류가 발생했습니다: {str(e)}")
            self.keywords = None
    
    def start_scraping(self):
        self.start_button.config(state=tk.DISABLED)
        self.status_label.config(text="H열 빈 칼럼 추가 중...")
        self.root.update_idletasks()

        if not self.keywords:
            messagebox.showerror("Error", "키워드를 찾을 수 없습니다.")
            self.start_button.config(state=tk.NORMAL)
            self.status_label.config(text="")
            return
        
        # H열 앞에 빈 칼럼 추가 후 바로 크롤링 진행
        try:
            self.add_empty_column_to_h()
            
            # H열 추가 완료 후 바로 크롤링 진행
            self.status_label.config(text="크롤링 중...")
            self.root.update_idletasks()
            self.perform_scraping()
                
        except Exception as e:
            messagebox.showerror("Error", f"H열 빈 칼럼 추가 중 오류가 발생했습니다: {str(e)}")
            self.start_button.config(state=tk.NORMAL)
            self.status_label.config(text="")
    
    def perform_scraping(self):
        try:
            # Chrome 옵션 설정
            options = Options()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument("--start-maximized")
            
            # Chrome WebDriver 설정
            driver = webdriver.Chrome(options=options)
            
            # 구글 시트 클라이언트 생성 (이미 H열이 추가된 상태)
            creds = get_credentials()
            client = gspread.authorize(creds)
            spreadsheet = client.open_by_key(self.spreadsheet_id)
            keywords_worksheet = spreadsheet.worksheet(self.sheet_name)
            
            # 실행 시각을 헤더에 추가
            from datetime import datetime
            current_time = datetime.now().strftime('%m/%d %H:%M')
            
            # 헤더에 실행 시각 추가
            try:
                keywords_worksheet.update_cell(1, 8, current_time)  # H열 헤더
                # 헤더 중앙 정렬 적용
                self.format_cell_center_align(keywords_worksheet, 1, 8)
                print(f"헤더 H열에 실행 시각 추가: {current_time}")
            except Exception as e:
                print(f"헤더 업데이트 중 오류: {str(e)}")
            
            for search_query in self.keywords:
                print(f"\n{'='*50}")
                print(f"키워드: {search_query}")
                print(f"{'='*50}")
                
                # 상태 표시
                self.root.title(f"'{search_query}' 크롤링 중...")
                
                driver.get("https://www.naver.com")
                time.sleep(1)
                
                driver.find_element(By.ID, "query").send_keys(search_query)
                driver.find_element(By.ID, "search-btn").click()
                time.sleep(1)
                
                driver.find_element(By.LINK_TEXT, "블로그").click()
                time.sleep(2)
                
                html = BeautifulSoup(driver.page_source, 'html.parser')
                title_links = html.find_all('a', class_='title_link', limit=15)
                name_links = html.find_all('a', class_='name', limit=15)
                
                # 날짜 추출
                date_links = []
                user_info_divs = html.find_all('div', class_='user_info')
                
                for user_info in user_info_divs[:15]:
                    sub_spans = user_info.find_all('span', class_='sub')
                    date_found = False
                    for span in sub_spans:
                        date_text = span.text.strip()
                        # 날짜 형식이 'YYYY.MM.DD.' 또는 'YYYY-MM-DD' 형식인지 확인
                        if (len(date_text) >= 8 and 
                            ((date_text.count('.') >= 2 and all(part.isdigit() for part in date_text.replace('.', ' ').split())) or
                             (date_text.count('-') >= 2 and all(part.isdigit() for part in date_text.replace('-', ' ').split())))):
                            date_links.append(span)
                            date_found = True
                            break
                        # 상대적 시간 표현 확인 ('시간 전', '일 전' 등)
                        elif ('전' in date_text and 
                              any(unit in date_text for unit in ['시간', '일', '주', '개월', '년']) and
                              any(char.isdigit() for char in date_text)):
                            date_links.append(span)
                            date_found = True
                            break
                    if not date_found:
                        date_links.append(None)
                
                # 유다모가 포함된 제목의 순위 찾기
                yudamo_positions = []
                for i in range(min(len(title_links), len(name_links), len(date_links))):
                    title_link = title_links[i]
                    name_link = name_links[i]
                    date_link = date_links[i]
                    
                    title = title_link.text.strip()
                    author = name_link.text.strip()
                    
                    # 유다모가 제목에 포함되어 있는지 확인
                    if '유다모' in title:
                        yudamo_positions.append(i + 1)  # 1부터 시작하는 순위
                    
                    # 날짜 정보는 콘솔 출력용으로만 사용
                    if date_link:
                        date_text = date_link.text.strip()
                        # 상대적 시간을 실제 날짜로 변환 (콘솔 출력용)
                        display_date = self.convert_relative_time_to_date(date_text)
                    else:
                        display_date = "날짜 없음"
                    
                    print(f"{i+1:2d}. 제목: {title}")
                    print(f"    작성자: {author}")
                    print(f"    작성시간: {display_date}")
                    print()
                
                # H열에 유다모 순위만 기록 (실시간 업데이트)
                self.update_yudamo_positions(keywords_worksheet, search_query, yudamo_positions)
                
                # 유다모 발견 결과 출력
                if yudamo_positions:
                    print(f"유다모 발견: {', '.join(map(str, yudamo_positions))}번째")
                else:
                    print("유다모 발견: 없음")
                print()
                
                # 잠시 대기 (구글 시트 API 제한 방지)
                time.sleep(1)
            
            driver.quit()
            
            print(f"\n{'='*50}")
            print("크롤링 완료!")
            print(f"{'='*50}")
            
            # 창 제목 원래대로 복구
            self.root.title("네이버 블로그 크롤러")
            self.status_label.config(text="크롤링 완료!")
            
            messagebox.showinfo("완료", "크롤링이 완료되었습니다. 콘솔을 확인해주세요.")
            
        except Exception as e:
            messagebox.showerror("Error", f"크롤링 중 오류가 발생했습니다: {str(e)}")
        finally:
            # UI 상태 초기화
            self.start_button.config(state=tk.NORMAL)
            self.root.title("네이버 블로그 크롤러")

    def update_yudamo_positions(self, worksheet, keyword, positions):
        """키워드 시트의 H열에 유다모 순위를 실시간으로 기록합니다"""
        try:
            # A열에서 해당 키워드의 행 번호 찾기
            all_keywords = worksheet.col_values(1)
            row_number = None
            
            for i, kw in enumerate(all_keywords, 1):
                if kw.strip() == keyword:
                    row_number = i
                    break
            
            if row_number:
                # 유다모 순위 기록
                if positions:
                    # 순위 앞에 #을 붙여서 날짜로 해석되는 것을 방지
                    position_text = ', '.join([f'#{pos}' for pos in positions])
                else:
                    position_text = '-'
                
                # H열에 새로운 데이터 입력
                worksheet.update_cell(row_number, 8, position_text)  # H열
                
                # 셀 중앙 정렬 적용
                self.format_cell_center_align(worksheet, row_number, 8)
                
                print(f"[{keyword}] H열에 기록: {position_text}")
                
                # UI 상태 업데이트
                self.status_label.config(text=f"[{keyword}] 완료 - 유다모: {position_text}")
                self.root.update_idletasks()
                
        except Exception as e:
            print(f"유다모 순위 기록 중 오류: {str(e)}")
            self.status_label.config(text=f"[{keyword}] 기록 실패")
            self.root.update_idletasks()
    
    def add_empty_column_to_h(self):
        """H열 앞에 빈 칼럼을 추가합니다"""
        try:
            print("H열 앞에 빈 칼럼 추가 중...")
            
            # auth.py의 인증 함수 사용
            creds = get_credentials()
            
            # gspread 클라이언트 생성
            client = gspread.authorize(creds)
            
            # 스프레드시트 열기
            spreadsheet = client.open_by_key(self.spreadsheet_id)
            keywords_worksheet = spreadsheet.worksheet(self.sheet_name)
            
            # H열(8번째) 앞에 열 삽입 - Google Sheets API 직접 사용
            body = {
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": keywords_worksheet.id,
                                "dimension": "COLUMNS",
                                "startIndex": 7,  # H열 앞에 삽입 (0-based이므로 7)
                                "endIndex": 8
                            },
                            "inheritFromBefore": False
                        }
                    }
                ]
            }
            spreadsheet.batch_update(body)
            print("H열 앞에 빈 칼럼 추가 완료!")
            
        except Exception as e:
            print(f"H열 추가 중 오류: {str(e)}")
            # 에러 발생 시 상세 정보 출력
            import traceback
            traceback.print_exc()
    
    def shift_all_columns_right(self, worksheet, start_col):
        """H열에만 빈 칼럼을 추가합니다"""
        try:
            print("H열에 빈 칼럼 추가 중...")
            
            # 현재 시트의 모든 데이터 가져오기
            all_values = worksheet.get_all_values()
            
            # 시트의 실제 열 수 확인 (최대 23열)
            max_cols = min(23, len(all_values[0]) if all_values else 23)
            
            # 각 행에 대해 H열부터 오른쪽으로 한 열씩 밀어내기
            for row_idx, row_data in enumerate(all_values, 1):
                # H열(8번째)부터 실제 열 수까지 역순으로 밀어내기
                for col in range(max_cols, 7, -1):  # max_cols부터 H열(8)까지 역순으로
                    if col <= len(row_data):
                        # 현재 열의 값을 오른쪽 열로 이동
                        current_value = row_data[col - 1] if col <= len(row_data) else ""
                        try:
                            worksheet.update_cell(row_idx, col + 1, current_value)
                        except:
                            # 열 제한에 도달한 경우 무시
                            break
                    else:
                        # 빈 셀인 경우에도 이동
                        try:
                            worksheet.update_cell(row_idx, col + 1, "")
                        except:
                            # 열 제한에 도달한 경우 무시
                            break
                
                # 진행 상황 표시 (10행마다)
                if row_idx % 10 == 0:
                    print(f"  {row_idx}행 처리 완료...")
                    self.status_label.config(text=f"데이터 이동 중... {row_idx}행 완료")
                    self.root.update_idletasks()
            
            print("H열에 빈 칼럼 추가 완료!")
            
            self.status_label.config(text="H열 빈 칼럼 추가 완료")
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"열 추가 중 오류: {str(e)}")
            # 에러 발생 시 상세 정보 출력
            import traceback
            traceback.print_exc()
    


    def format_cell_center_align(self, worksheet, row, col):
        """셀에 중앙 정렬 포맷을 적용합니다"""
        try:
            # 셀 범위 지정 (예: H5)
            cell_range = f"{chr(64 + col)}{row}"
            
            # 중앙 정렬 포맷 적용
            body = {
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": row - 1,
                                "endRowIndex": row,
                                "startColumnIndex": col - 1,
                                "endColumnIndex": col
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "horizontalAlignment": "CENTER",
                                    "verticalAlignment": "MIDDLE"
                                }
                            },
                            "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
                        }
                    }
                ]
            }
            
            # 포맷 적용
            worksheet.spreadsheet.batch_update(body)
            
        except Exception as e:
            print(f"셀 포맷 적용 중 오류: {str(e)}")

    def open_google_sheet(self):
        """구글 시트를 기본 웹 브라우저로 엽니다"""
        webbrowser.open(self.sheet_url)

if __name__ == "__main__":
    root = tk.Tk()
    app = NaverBlogScraper(root)
    root.mainloop() 