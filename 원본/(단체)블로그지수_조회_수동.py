"""
================================================================================
블로그지수 자동 수집 도구 (Blogdex Crawler) - 유다모용
================================================================================

목적: 
- Blogdex 사이트(https://blogdex.space)에서 네이버 블로그들의 블로그지수만 자동으로 수집
- Google Sheets에 저장된 블로그 URL 목록을 읽어서 각 블로그의 지수를 F열에 직접 저장

주요 기능:
1. Google Sheets에서 블로그 URL 목록 읽기 ('유다모' 시트 A열 기준)
2. Blogdex 사이트에 카카오톡 계정으로 자동 로그인
3. 각 블로그별로 블로그지수만 수집
4. 수집된 블로그지수를 Google Sheets의 F열에 직접 저장

안전 기능:
- 랜덤한 대기 시간으로 서버 부하 방지
- 50~80개마다 3~5분 긴 휴식 시간
- undetected_chromedriver로 봇 탐지 우회

사용법:
1. Google Sheets에서 '유다모' 시트의 A열에 블로그 URL 목록 입력
2. 자동으로 크롤링 시작 (진행 상황 실시간 출력)
3. 수집된 블로그지수가 F열에 자동 저장

입력/출력:
- 입력: Google Sheets '유다모' 시트 A열의 블로그 URL 목록
- 출력: Google Sheets '유다모' 시트 F열에 블로그지수 저장
================================================================================
"""

import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random
from time import sleep
import sys
import time
from auth import get_credentials
from googleapiclient.discovery import build

# Google Sheets 설정
SPREADSHEET_ID = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos"

def get_sheet_list():
    """Google Sheets의 시트 목록을 가져오는 함수"""
    try:
        # 시트 메타데이터 가져오기
        result = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID
        ).execute()
        
        sheets = result.get('sheets', [])
        sheet_list = []
        
        for i, sheet in enumerate(sheets, 1):
            sheet_name = sheet['properties']['title']
            sheet_list.append((i, sheet_name))
            print(f"{i}. {sheet_name}")
        
        return sheet_list
        
    except Exception as e:
        print(f"시트 목록을 가져오는 중 오류 발생: {e}")
        raise

def select_sheet():
    """사용자가 시트를 선택하는 함수"""
    print("사용 가능한 시트 목록:")
    sheet_list = get_sheet_list()
    
    while True:
        try:
            choice = input("\n사용할 시트 번호를 입력하세요: ")
            choice_num = int(choice)
            
            if 1 <= choice_num <= len(sheet_list):
                selected_sheet = sheet_list[choice_num - 1][1]
                print(f"\n선택된 시트: {selected_sheet}")
                return selected_sheet
            else:
                print(f"1부터 {len(sheet_list)} 사이의 번호를 입력해주세요.")
        except ValueError:
            print("올바른 숫자를 입력해주세요.")

def get_data_from_sheets(sheet_name):
    """Google Sheets에서 블로그 URL 목록을 한 번에 가져오는 함수"""
    try:
        # 선택된 시트에서 A열, D열, F열 데이터 가져오기
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{sheet_name}!A:F'  # A열부터 F열까지
        ).execute()
        
        values = result.get('values', [])
        
        # 처리할 URL 목록 생성
        valid_urls = []
        skipped_count = 0
        
        for row_idx, row in enumerate(values):
            if len(row) >= 1 and row[0].strip():  # A열에 URL이 있는 경우만
                url = row[0].strip()
                
                # D열 또는 F열에 데이터가 있는지 확인 (이미 처리된 행인지 체크)
                d_has_data = len(row) >= 4 and row[3].strip()  # D열에 데이터가 있는지
                f_has_data = len(row) >= 6 and row[5].strip()  # F열에 데이터가 있는지
                
                if d_has_data or f_has_data:  # D열 또는 F열에 데이터가 있으면 건너뛰기
                    skipped_count += 1
                    reason = "D열에 데이터 있음" if d_has_data else "F열에 데이터 있음"
                    print(f"이미 처리됨 (건너뜀): {url} - {reason}")
                    continue
                
                # 네이버 블로그 URL인지 확인
                if 'blog.naver.com' in url:
                    # URL 전처리: 포스트 번호 제거
                    if '/223' in url or '/224' in url or '/225' in url:  # 포스트 번호 패턴 확인
                        # https://blog.naver.com/thewonny/223918997335 -> https://blog.naver.com/thewonny
                        url_parts = url.split('/')
                        if len(url_parts) >= 4:  # 최소 4개 부분이 있어야 함
                            # 처음 4개 부분만 사용 (https:, , blog.naver.com, 블로그ID)
                            processed_url = '/'.join(url_parts[:4])
                            valid_urls.append((processed_url, row_idx))
                            print(f"URL 전처리: {url} -> {processed_url}")
                        else:
                            valid_urls.append((url, row_idx))
                    else:
                        valid_urls.append((url, row_idx))
                else:
                    print(f"네이버 블로그가 아닌 URL 제외: {url}")
                    continue
        
        if skipped_count > 0:
            print(f"\n이미 처리된 {skipped_count}개의 URL을 제외했습니다.")
        
        print(f"처리할 URL {len(valid_urls)}개를 가져왔습니다.")
        return valid_urls
        
    except Exception as e:
        print(f"Google Sheets에서 데이터를 가져오는 중 오류 발생: {e}")
        raise

def update_blog_data_to_sheets(service, sheet_name, row_index, blog_name, blog_index, max_retries=3):
    """Google Sheets의 B열에 블로그명, F열에 블로그지수를 저장하는 함수 (재시도 로직 포함)"""
    for attempt in range(max_retries):
        try:
            # B열에 블로그명 저장 (행 번호는 1부터 시작하므로 +1)
            range_name_b = f'{sheet_name}!B{row_index + 1}'
            values_b = [[blog_name]]
            
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name_b,
                valueInputOption='RAW',
                body={'values': values_b}
            ).execute()
            
            # F열에 블로그지수 저장 (행 번호는 1부터 시작하므로 +1)
            range_name_f = f'{sheet_name}!F{row_index + 1}'
            values_f = [[blog_index]]
            
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name_f,
                valueInputOption='RAW',
                body={'values': values_f}
            ).execute()
            
            # 블로그지수가 특정 값인 경우 D열에 '제외' 입력
            exclude_values = ['일반', '준최1', '준최2', '준최3', '준최4']
            if blog_index in exclude_values:
                range_name_d = f'{sheet_name}!D{row_index + 1}'
                values_d = [['제외']]
                
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=range_name_d,
                    valueInputOption='RAW',
                    body={'values': values_d}
                ).execute()
                
                print(f"B{row_index + 1}에 블로그명 '{blog_name}', F{row_index + 1}에 블로그지수 '{blog_index}', D{row_index + 1}에 '제외' 저장 완료")
            else:
                print(f"B{row_index + 1}에 블로그명 '{blog_name}', F{row_index + 1}에 블로그지수 '{blog_index}' 저장 완료")
            
            return True  # 성공 시 True 반환
            
        except Exception as e:
            error_msg = str(e).lower()
            
            # API 제한 관련 오류인지 확인
            if 'quota' in error_msg or 'rate limit' in error_msg or '429' in error_msg:
                wait_time = (attempt + 1) * 10  # API 제한 시 더 긴 대기 시간 (10초, 20초, 30초)
                print(f"API 제한 도달 (시도 {attempt + 1}/{max_retries}): {e}")
                print(f"{wait_time}초 후 재시도합니다...")
                time.sleep(wait_time)
            elif attempt < max_retries - 1:  # 일반 오류인 경우
                wait_time = (attempt + 1) * 2  # 2초, 4초, 6초로 증가
                print(f"Google Sheets 업데이트 실패 (시도 {attempt + 1}/{max_retries}): {e}")
                print(f"{wait_time}초 후 재시도합니다...")
                time.sleep(wait_time)
            else:  # 마지막 시도에서도 실패
                print(f"Google Sheets 업데이트 최종 실패 (시도 {max_retries}회): {e}")
                return False  # 실패 시 False 반환

def countdown(seconds):
    for i in range(seconds, 0, -1):
        sys.stdout.write(f"\r대기 시간: {i}초")
        sys.stdout.flush()
        sleep(1)
    sys.stdout.write("\r완료!      \n")
    sys.stdout.flush()

# Google Sheets 서비스 초기화
creds = get_credentials()
service = build('sheets', 'v4', credentials=creds)

# 사용자가 시트 선택
selected_sheet = select_sheet()

# Google Sheets에서 데이터 한 번에 가져오기
url_list = get_data_from_sheets(selected_sheet)

# Selenium 웹드라이버 설정
options = uc.ChromeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument('--disable-infobars')
options.add_argument('--lang=ko_KR')

driver = uc.Chrome(options=options)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

# Blogdex 로그인 페이지로 이동
url = "https://blogdex.space/login?from=/blog-index"
driver.get(url)
time.sleep(3.2)  # 로그인 페이지 로딩 대기

# 로그인 페이지에서 수동 로그인 대기
print("Blogdex 로그인 페이지가 열렸습니다.")
print("수동 로그인을 완료해주세요.")
print("로그인이 완료되면 Enter 키를 눌러주세요...")
input()

# 로그인 후 대기 시간을 10~60초 랜덤으로 설정
wait_time = random.randint(5, 10)
print(f"\n로그인 후 {wait_time}초 대기합니다...")
countdown(wait_time)



save_count = 0
RANDOM_BREAK_COUNT = random.randint(50, 80)  # 50~80개 사이에서 랜덤하게 휴식 시점 설정

print("스트림 방식으로 URL을 하나씩 처리합니다...")

# URL 목록을 하나씩 처리
for blog_url, row_index in url_list:
    
    try:
        blog_id = blog_url.split("/")[-1]  # URL에서 블로그 ID 추출
        url_2 = f"https://blogdex.space/blog-index/{blog_id}"
        
        print(f"\n처리 중: {blog_url}")
        
        driver.get(url_2)
        sleep_time = random.uniform(8, 10)  # 8~10초 랜덤 대기
        time.sleep(sleep_time)  # 아이디마다 크롤링할 때 랜덤 타임슬립

        # 블로그지수와 블로그명 스크래핑
        blog_index = driver.find_element(By.CSS_SELECTOR, "svg > text[font-family='Pretendard'][font-size='22px'][font-weight='700'][y='-60']").text
        blog_name_raw = driver.find_element(By.CSS_SELECTOR, "div.flex.space-x-1.pr-0.md\\:pr-24 p.text-sm.font-medium.leading-none").text
        
        # 블로그명에서 괄호와 아이디 부분 제거 (예: "블로그명 (아이디)" -> "블로그명")
        if '(' in blog_name_raw and ')' in blog_name_raw:
            blog_name = blog_name_raw.split('(')[0].strip()
        else:
            blog_name = blog_name_raw
        
        # Google Sheets의 B열에 블로그명, F열에 블로그지수 저장 (원본 행 번호 사용)
        save_success = update_blog_data_to_sheets(service, selected_sheet, row_index, blog_name, blog_index)
        
        if save_success:
            save_count += 1
            # 실시간으로 출력
            print(f"완료: {blog_id} - 블로그명: {blog_name} - 블로그지수: {blog_index}")
        else:
            print(f"저장 실패: {blog_id} - 블로그명: {blog_name} - 블로그지수: {blog_index} (나중에 다시 시도해주세요)")
            # 저장 실패 시 잠시 대기 후 다음 URL로 진행
            time.sleep(5)

        # 랜덤한 개수마다 긴 휴식 시간 가지기
        if save_count % RANDOM_BREAK_COUNT == 0:
            long_break = random.randint(180, 300)  # 3~5분 사이 랜덤 휴식
            print(f"\n{save_count}개 수집 완료. {long_break}초 동안 휴식합니다...")
            countdown(long_break)
            RANDOM_BREAK_COUNT = random.randint(50, 80)  # 다음 휴식 시점도 랜덤하게 재설정

    except Exception as e:
        print(f"블로그지수 수집 실패 (아이디: {blog_id}): {e}")
        continue

# 드라이버 종료
driver.quit()

print(f"모든 블로그지수 수집이 완료되었습니다. 시트를 확인해주세요.")
