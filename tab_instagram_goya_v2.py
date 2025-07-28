'''
## 목록 시트
https://docs.google.com/spreadsheets/d/1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0/edit?gid=1609134358#gid=1609134358

## 템플릿 시트
https://docs.google.com/spreadsheets/d/1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0/edit?gid=1655394428#gid=1655394428


## 🚩요약
인플루언서에게 자연스럽게 팔로우하고 DM을 보내는 자동화 시스템입니다.
메세지 보내는 부분은 모듈로 임포트 되어있습니다. (instagram_message_vendor.py)

## 템플릿 시트 변수
템플릿에서 {이름}, {노션리스트}를 변수로 받아서 메세지를 보냅니다.

## 🚩주요 기능
- 자연스러운 Instagram 사용자 행동 모방
- 팔로우 → 피드 살펴보기 → DM 발송 순서
- 2~7명마다 브라우저 세션 재시작으로 안정성 확보
- Instagram 메인에서 10-80회 랜덤 스크롤
- 인플루언서 피드에서 3-15개 게시물 랜덤 살펴보기
- Google Sheets와 MongoDB 연동으로 발송 상태 추적

## 🚩사용법
- 코드를 실행하면 프로필 선택합니다. (*사용 가능한 프로필 목록:)
- DM보낼 인원 목록 시트를 선택합니다. (*1. 사용 가능한 시트 목록:) 
    - 발송조건 - I열에 'DM요청' 상태인 행만 발송 대상
    
- 템플릿 시트를 선택합니다. (*2. 사용 가능한 시트 목록:)
- 최종 발송할 인플루언서 확인 후 Y를 눌러 발송합니다.

- 발송 완료 후 목록 시트에서 발송 상태를 확인할 수 있습니다. (MongoDB에도 로그 기록)

## 🚩예상 소요 시간 (30명 기준)
- 최소: 약 6시간
- 평균: 약 10시간  
- 최대: 약 17.5시간

## 🚩안전성 특징
- 랜덤한 대기 시간과 스크롤 횟수
- 브라우저 세션 격리
- 자연스러운 사용자 행동 패턴
- 에러 발생 시에도 안전하게 진행

'''

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import random
import time
import os
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pyperclip
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from auth import get_credentials
from googleapiclient.discovery import build
import logging
from datetime import datetime
from dotenv import load_dotenv
import sys
from pymongo import MongoClient
from pymongo.server_api import ServerApi
from instagram_message_vendor import InstagramMessageTemplate
from pathlib import Path

# 환경 변수 로드
load_dotenv()

# 스프레드시트 ID 설정
DM_LIST_SPREADSHEET_ID = '1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0'
TEMPLATE_SPREADSHEET_ID = '1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0'

# MongoDB 연결 설정
uri = "mongodb+srv://coq3820:JmbIOcaEOrvkpQo1@cluster0.qj1ty.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
try:
    mongo_client = MongoClient(uri, 
                             server_api=ServerApi('1'),
                             tlsAllowInvalidCertificates=True)
    mongo_client.admin.command('ping')
    print("MongoDB 연결 성공!")
    
    db = mongo_client['insta09_database']
    dm_collection = db['gogoya_DmRecords']
    mongo_connected = True
except Exception as e:
    print(f"MongoDB 연결 실패: {e}")
    mongo_connected = False

# 프로젝트 루트 디렉토리를 파이썬 경로에 추가
project_root = Path(__file__).parent.parent.absolute()
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from notion_module.notion_reader import get_database_items, extract_page_id_from_url, print_database_items

def clear_chrome_data(user_data_dir, keep_login=True):
    default_dir = os.path.join(user_data_dir, 'Default')
    if not os.path.exists(default_dir):
        print("Default 디렉토리가 존재하지 않습니다.")
        return

    dirs_to_clear = ['Cache', 'Code Cache', 'GPUCache']
    files_to_clear = ['History', 'Visited Links', 'Web Data']
    
    for dir_name in dirs_to_clear:
        dir_path = os.path.join(default_dir, dir_name)
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
            print(f"{dir_name} 디렉토리를 삭제했습니다.")

    if not keep_login:
        files_to_clear.extend(['Cookies', 'Login Data'])

    for file_name in files_to_clear:
        file_path = os.path.join(default_dir, file_name)
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"{file_name} 파일을 삭제했습니다.")

def get_sheet_names(service, spreadsheet_id):
    """스프레드시트의 모든 시트 이름을 가져옴"""
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        return [sheet.get("properties", {}).get("title", "") for sheet in sheets]
    except Exception as e:
        print(f"시트 목록을 가져오는 중 오류 발생: {e}")
        return []

def select_sheet(service, spreadsheet_id):
    """사용자에게 시트를 선택하도록 함"""
    sheet_names = get_sheet_names(service, spreadsheet_id)
    if not sheet_names:
        print("시트 목록을 가져올 수 없습니다.")
        return None
    
    # '인스타'와 '현황판'이 포함된 시트만 필터링
    filtered_sheets = [name for name in sheet_names if '인스타' in name and '현황판' in name]
    
    if not filtered_sheets:
        print("'인스타'와 '현황판'이 포함된 시트가 없습니다.")
        return None
        
    print("\n사용 가능한 시트 목록 (인스타 + 현황판):")
    for idx, name in enumerate(filtered_sheets, 1):
        print(f"{idx}. {name}")
        
    while True:
        try:
            choice = int(input("\n사용할 시트 번호를 선택하세요: "))
            if 1 <= choice <= len(filtered_sheets):
                selected_sheet = filtered_sheets[choice - 1]
                print(f"\n선택된 시트: {selected_sheet}")
                return selected_sheet
            else:
                print("유효하지 않은 번호입니다. 다시 선택해주세요.")
        except ValueError:
            print("숫자를 입력해주세요.")

def get_data_from_sheets(service, selected_sheet):
    logging.info("URL과 이름 데이터 가져오기 시작")
    try:
        sheet = service.spreadsheets()
        # A열(URL), B열(하이퍼링크), C열(이름), I열(DM요청 상태)까지 가져오도록 수정
        result = sheet.values().get(spreadsheetId=DM_LIST_SPREADSHEET_ID,
                                    range=f'{selected_sheet}!A2:I').execute()
        values = result.get('values', [])

        if not values:
            logging.warning('스프레드시트에서 데이터를 찾을 수 없습니다.')
            return []

        # URL, 이름 반환 (실제 행 번호 포함)
        url_name_pairs = []
        print(f"총 {len(values)}개의 행을 확인합니다.")
        for row_idx, row in enumerate(values, start=2):  # 실제 스프레드시트의 행 번호 사용
            print(f"행 {row_idx} 확인 중: {row}")
            if (row and 
                len(row) > 8 and row[8] == 'DM요청'):  # I열에 'DM요청'이 있는 경우만
                print(f"행 {row_idx}: DM요청 조건 만족")
                
                # B열에서 인플루언서 이름 가져오기 (URL 생성용)
                instagram_username = ""
                if len(row) > 1 and row[1]:
                    instagram_username = row[1]
                    print(f"B열에서 가져온 인스타그램 사용자명: {instagram_username}")  # 디버깅용
                
                # C열에서 인플루언서 이름 가져오기 (템플릿 변수용)
                influencer_name = ""
                if len(row) > 2 and row[2]:
                    influencer_name = row[2]
                    print(f"C열에서 가져온 인플루언서 이름: {influencer_name}")  # 디버깅용
                
                # 인플루언서 사용자명으로 인스타그램 URL 생성
                if instagram_username:
                    instagram_url = f"https://www.instagram.com/{instagram_username}/"
                    print(f"생성된 URL: {instagram_url}")  # 디버깅용
                else:
                    instagram_url = ""
                    print("인스타그램 사용자명이 없어서 URL을 생성할 수 없습니다.")
                
                # URL이 비어있거나 유효하지 않으면 건너뛰기
                if not instagram_url or not instagram_url.startswith('http'):
                    print(f"행 {row_idx}: URL이 비어있거나 유효하지 않아서 건너뜁니다. URL: '{instagram_url}'")
                    continue
                
                url_name_pairs.append((
                    instagram_url,  # B열에서 추출한 인스타그램 URL
                    influencer_name,  # B열에서 추출한 인플루언서 이름
                    "",  # 노션리스트 (사용하지 않음)
                    "",  # G열 값 (사용하지 않음)
                    "",  # 전체 리스트 URL (사용하지 않음)
                    row_idx  # 실제 스프레드시트의 행 번호
                ))
        
        return url_name_pairs

    except Exception as e:
        logging.error(f"스프레드시트에서 데이터를 가져오는 중 오류 발생: {e}")
        return []

def update_sheet_status(service, row, status, timestamp=None, sheet_name=None):
    if not sheet_name:
        return
        
    sheet_id = DM_LIST_SPREADSHEET_ID
    # I열에 상태를 기록
    range_name = f'{sheet_name}!I{row}'
    
    # 성공 시 'DM완료(오늘날짜)', 실패 시 'DM실패(오늘날짜)' 기록
    if status == 'Y':
        today = datetime.now().strftime("%m/%d")
        value = f'DM완료({today})'
    else:
        today = datetime.now().strftime("%m/%d")
        value = f'DM실패({today})'
    
    values = [[value]]
    body = {'values': values}
    
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()

def countdown(wait_time, message):
    print(f"{message}: {wait_time:.2f}초", end='\r')
    for remaining in range(int(wait_time), 0, -1):
        print(f"{message}: {remaining}초 남음    ", end='\r')
        time.sleep(1)
    print(f"{message} 완료!    ", end='\r')

def natural_scroll_on_instagram(driver, url=None):
    """Instagram 메인/릴스에서 자연스러운 스크롤 동작을 수행"""
    try:
        if url:
            driver.get(url)
            print(f"\n{url} 페이지로 이동했습니다.")
            time.sleep(random.uniform(3, 8))  # 페이지 로딩 대기

        # 초기 페이지 높이 확인
        last_height = driver.execute_script("return document.body.scrollHeight")
        print(f"초기 페이지 높이: {last_height}px")
        
        # 스크롤 횟수 랜덤 설정 (40-200회로 증가)
        scroll_attempts = random.randint(40, 200)
        print(f"스크롤 시도 횟수: {scroll_attempts}회")
        
        for attempt in range(scroll_attempts):
            print(f"\n--- {attempt + 1}번째 스크롤 시도 ---")
            
            # 페이지 스크롤 (더 큰 범위로 수정)
            scroll_multiplier = random.uniform(1.2, 2.0)  # 120% ~ 200% 사이의 랜덤 배수
            viewport_height = driver.execute_script("return window.innerHeight;")
            scroll_height = int(viewport_height * scroll_multiplier)  # 화면 높이의 120-200%만큼 스크롤

            current_position = driver.execute_script("return window.pageYOffset;")
            target_position = min(current_position + scroll_height, last_height)

            # 중간 지점들을 만들어 자연스러운 스크롤 구현
            steps = random.randint(2, 4)  # 2-4개의 중간 지점
            for step in range(steps):
                intermediate_position = current_position + (target_position - current_position) * (step + 1) / steps
                driver.execute_script(f"window.scrollTo({{top: {intermediate_position}, behavior: 'smooth'}});")
                time.sleep(random.uniform(0.5, 1))  # 각 중간 스크롤마다 짧은 대기

            # 최종 위치로 스크롤
            driver.execute_script(f"window.scrollTo({{top: {target_position}, behavior: 'smooth'}});")

            # 더 긴 대기 시간 설정 (2-4초)
            wait_time = random.uniform(2, 4)
            time.sleep(wait_time)

            # 새로운 높이 계산 전에 추가 대기
            time.sleep(random.uniform(3, 8))
            new_height = driver.execute_script("return document.body.scrollHeight")

            # 추가 스크롤 시도
            retry_count = 0
            while new_height == last_height and retry_count < 10:
                print(f"\n새로운 컨텐츠를 찾기 위해 {retry_count + 1}번째 추가 스크롤 시도...")
                
                # 현재 위치에서 조금 더 아래로 스크롤
                current_position = driver.execute_script("return window.pageYOffset;")
                scroll_amount = random.randint(300, 1000)  # 300~1000픽셀 추가 스크롤
                driver.execute_script(f"window.scrollTo({current_position}, {current_position + scroll_amount});")
                
                time.sleep(3)  # 로딩 대기
                
                # 새로운 높이 확인
                new_height = driver.execute_script("return document.body.scrollHeight")
                retry_count += 1
            
            # 새로운 컨텐츠가 로드되었으면 높이 업데이트
            if new_height > last_height:
                last_height = new_height
                print(f"새로운 컨텐츠 로드됨! 새로운 높이: {new_height}px")
            
            # 스크롤 간 랜덤 대기 (3-10초로 증가)
            if attempt < scroll_attempts - 1:  # 마지막 스크롤이 아니면
                wait_between_scrolls = random.uniform(3, 10)
                print(f"다음 스크롤까지 {wait_between_scrolls:.1f}초 대기...")
                time.sleep(wait_between_scrolls)
        
        # 스크롤 완료 후 상단으로 돌아가기 (선택적)
        if random.choice([True, False]):  # 50% 확률로 상단으로 이동
            print("상단으로 돌아가는 중...")
            driver.execute_script("window.scrollTo({top: 0, behavior: 'smooth'});")
            time.sleep(random.uniform(1, 2))
        
        print("Instagram 메인 스크롤 동작 완료!")
        
    except Exception as e:
        print(f"Instagram 스크롤 중 오류 발생: {e}")
        # 오류가 발생해도 계속 진행
        pass

def browse_feed_posts(driver, profile_url):
    """인플루언서의 피드 게시물을 랜덤으로 살펴보기"""
    try:
        print(f"\n--- {profile_url}의 피드 게시물 살펴보기 시작 ---")
        
        # 랜덤으로 살펴볼 게시물 수 결정 (10-21개로 수정)
        num_posts_to_browse = random.randint(10, 21)
        print(f"랜덤으로 {num_posts_to_browse}개의 게시물을 살펴보겠습니다.")
        
        # 첫 번째 피드 게시물이 로드될 때까지 대기
        print("첫 번째 피드 게시물 로딩 대기 중...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div._aagv"))
        )

        # 잠시 대기
        time.sleep(3)

        # 첫 번째 게시물 찾기
        first_post = driver.find_element(By.CSS_SELECTOR, "div._aagv")

        # 부모 요소로 이동하여 링크 찾기
        parent = first_post.find_element(By.XPATH, "./ancestor::a")
        post_link = parent.get_attribute("href")

        # JavaScript로 첫 번째 게시물 클릭
        print(f"\n첫 번째 게시물({post_link})을 클릭합니다...")
        driver.execute_script("arguments[0].click();", parent)
        
        # 게시물 클릭 후 로딩 대기
        time.sleep(random.uniform(2, 4))
        
        # 랜덤한 게시물들을 살펴보기
        for i in range(num_posts_to_browse - 1):  # 첫 번째는 이미 클릭했으므로 -1
            print(f"\n--- {i+2}번째 게시물 살펴보기 ---")
            
            # 게시물을 살펴보는 시간 (3-60초 랜덤으로 수정)
            browse_time = random.uniform(3, 60)
            print(f"게시물을 {browse_time:.1f}초간 살펴보는 중...")
            time.sleep(browse_time)
            
            # 다음 버튼 찾기
            next_button = None
            selector = "//span[contains(@style, 'rotate(90deg)')]/.."  # 90도 회전된 화살표(다음 버튼)의 부모 요소

            print("\n다음 버튼 찾는 중...")
            try:
                next_button = driver.find_element(By.XPATH, selector)
                if next_button.is_displayed():
                    print("다음 버튼을 찾았습니다.")
                else:
                    print("다음 버튼이 화면에 표시되지 않습니다.")
                    break
            except Exception as e:
                print(f"다음 버튼을 찾을 수 없습니다: {str(e)}")
                break

            if next_button is None:
                print(f"{i+2}번째 피드로 이동할 수 없습니다. 다음 버튼을 찾을 수 없습니다.")
                break

            print(f"\n{i+2}번째 피드로 이동합니다...")
            driver.execute_script("arguments[0].click();", next_button)
            
            # 다음 게시물 로딩 대기
            time.sleep(random.uniform(2, 4))
        
        # 피드 살펴보기 완료 후 프로필 페이지로 돌아가기
        print(f"\n피드 살펴보기 완료. 프로필 페이지({profile_url})로 돌아갑니다...")
        driver.get(profile_url)
        
        # 프로필 페이지 로딩 대기
        wait_time = random.uniform(1, 5)
        countdown(wait_time, "프로필 페이지 로딩 대기")
        
        print("피드 게시물 살펴보기 완료!")
        
    except Exception as e:
        print(f"피드 게시물 살펴보기 중 오류 발생: {e}")
        # 오류가 발생해도 프로필 페이지로 돌아가기
        try:
            driver.get(profile_url)
            time.sleep(3)
        except:
            pass

def process_url(driver, url, name, template_manager, row, service, sheet_name, user_data_dir):
    # Instagram 메인 피드 스크롤
    natural_scroll_on_instagram(driver)
    print("\n메인 피드 스크롤 완료!")
    
    # 릴스 페이지 스크롤 (주석처리)
    # print("\n릴스 페이지로 이동합니다...")
    # natural_scroll_on_instagram(driver, "https://www.instagram.com/reels/")
    # print("\n릴스 페이지 스크롤 완료!")
    
    # 인플루언서 프로필로 이동
    print(f"\n{url} 프로필로 이동합니다...")
    driver.get(url)
    print(driver.title)
    wait_time = random.uniform(5, 10)
    countdown(wait_time, "URL 접속 후 대기")

    try:
        # 먼저 팔로우 버튼 확인
        try:
            follow_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, 
                    "//button[.//div[contains(text(), '팔로우') or contains(text(), '팔로잉')]]"))
            )
            button_text = follow_button.find_element(By.XPATH, ".//div").text
            
            if button_text == "팔로우":
                follow_button.click()
                print("팔로우 완료")
                wait_time = random.uniform(4, 12)
                countdown(wait_time, "팔로우 후 대기")
        except TimeoutException:
            print("팔로우 버튼이 없거나 이미 팔로우 중입니다.")
            pass

        # 피드 게시물 랜덤 살펴보기
        browse_feed_posts(driver, url)

        # 메시지 보내기 버튼 찾기 (최대 3회 재시도) - 클릭은 주석처리
        message_button = None
        retry_count = 0
        max_retries = 3
        
        while retry_count < max_retries:
            try:
                message_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x1i10hfl') and contains(text(), '메시지 보내기')]"))
                )
                print(f"버튼 텍스트: {message_button.text}")
                print("메시지 보내기 버튼을 찾았습니다.")
                message_button.click()  # 클릭 동작 활성화
                break
            except TimeoutException:
                retry_count += 1
                if retry_count < max_retries:
                    print(f"메시지 보내기 버튼을 찾을 수 없습니다. 5초 후 재시도 ({retry_count}/{max_retries})")
                    time.sleep(5)
                else:
                    print("메시지 보내기 버튼을 찾을 수 없습니다. 최대 재시도 횟수 초과.")
                    raise TimeoutException("메시지 보내기 버튼을 찾을 수 없습니다.")
        
        wait_time = random.uniform(5, 20) #원래 60초였음
        countdown(wait_time, "DM 버튼 클릭 후 대기")

        message = template_manager.format_message(template_manager.get_message_templates()[0], name, "", "")
        print(f"생성된 메시지: {message}")
        pyperclip.copy(message)
        print("메시지를 클립보드에 복사했습니다.")
        
        # 메시지 입력 및 발송
        actions = ActionChains(driver)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@role, 'textbox')]"))
        ).click()
        
        if sys.platform == 'darwin':
            actions.key_down(Keys.COMMAND).send_keys('v').key_up(Keys.COMMAND).perform()
        else:
            actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        
        wait_time = random.uniform(5, 20)
        countdown(wait_time, "메시지 입력 후 대기")

        actions.send_keys(Keys.ENTER).perform()

        # 성공적으로 메시지를 보냈을 때
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        update_sheet_status(service, row, 'Y', timestamp, sheet_name)
        
        # MongoDB에 DM 기록 저장
        save_dm_record_to_mongodb(
            influe_name=name,
            contact_profile=os.path.basename(user_data_dir),
            status='Y',
            dm_date=timestamp,
            content='공구문의',
            message=message
        )

    except TimeoutException:
        print("'메시지 보내기' 버튼을 찾을 수 없습니다.")
        update_sheet_status(service, row, 'failed', None, sheet_name)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_dm_record_to_mongodb(
            influe_name=name,
            contact_profile=os.path.basename(user_data_dir),
            status='failed',
            dm_date=timestamp,
            content='공구문의',
            message="메시지 보내기 버튼을 찾을 수 없음"
        )
        print("메시지 보내기 버튼을 찾을 수 없어 프로그램을 종료합니다.")
        sys.exit(1)
    except NoSuchElementException:
        print("요소를 찾을 수 없습니다.")
        update_sheet_status(service, row, 'failed', None, sheet_name)
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_dm_record_to_mongodb(
            influe_name=name,
            contact_profile=os.path.basename(user_data_dir),
            status='failed',
            dm_date=timestamp,
            content='공구문의',
            message="요소를 찾을 수 없음"
        )
        print("요소를 찾을 수 없어 프로그램을 종료합니다.")
        sys.exit(1)

def save_dm_record_to_mongodb(influe_name, contact_profile, status, dm_date, content, message):
    if not mongo_connected:
        print("MongoDB에 연결되어 있지 않아 기록을 저장할 수 없습니다.")
        return
        
    try:
        record = {
            "influencer_name": influe_name,
            "contact_profile": contact_profile,
            "status": status,
            "dm_date": dm_date,
            "template_content": content,
            "message": message,
            "created_at": datetime.now()
        }
        
        dm_collection.insert_one(record)
        print(f"MongoDB에 DM 기록 저장 성공: {influe_name}")
    except Exception as e:
        print(f"MongoDB에 DM 기록 저장 실패: {e}")

def get_available_profiles(user_data_parent):
    """사용 가능한 프로필 목록을 가져옴"""
    profiles = []
    if not os.path.exists(user_data_parent):
        os.makedirs(user_data_parent)
        return profiles
        
    for item in os.listdir(user_data_parent):
        item_path = os.path.join(user_data_parent, item)
        if os.path.isdir(item_path):
            if (os.path.exists(os.path.join(item_path, 'Default')) or 
                any(p.startswith('Profile') for p in os.listdir(item_path) if os.path.isdir(os.path.join(item_path, p)))):
                profiles.append(item)
    return profiles

def select_profile(user_data_parent):
    """사용자에게 프로필을 선택하도록 함"""
    profiles = get_available_profiles(user_data_parent)
    if not profiles:
        print("\n사용 가능한 프로필이 없습니다.")
        create_new = input("새 프로필을 생성하시겠습니까? (y/n): ").lower()
        if create_new == 'y':
            while True:
                name = input("새 프로필 이름을 입력하세요: ")
                if not name:
                    print("프로필 이름을 입력해주세요.")
                    continue
                    
                if any(c in r'\\/:*?"<>|' for c in name):
                    print("프로필 이름에 다음 문자를 사용할 수 없습니다: \\ / : * ? \" < > |")
                    continue
                    
                new_profile_path = os.path.join(user_data_parent, name)
                if os.path.exists(new_profile_path):
                    print(f"'{name}' 프로필이 이미 존재합니다.")
                    continue
                    
                try:
                    os.makedirs(new_profile_path)
                    os.makedirs(os.path.join(new_profile_path, 'Default'))
                    print(f"'{name}' 프로필이 생성되었습니다.")
                    return name
                except Exception as e:
                    print(f'프로필 생성 중 오류가 발생했습니다: {e}')
                    retry = input("다시 시도하시겠습니까? (y/n): ").lower()
                    if retry != 'y':
                        return None
        return None
        
    print("\n사용 가능한 프로필 목록:")
    for idx, profile in enumerate(profiles, 1):
        print(f"{idx}. {profile}")
    print(f"{len(profiles) + 1}. 새 프로필 생성")
        
    while True:
        try:
            choice = int(input("\n사용할 프로필 번호를 선택하세요: "))
            if 1 <= choice <= len(profiles):
                selected_profile = profiles[choice - 1]
                print(f"\n선택된 프로필: {selected_profile}")
                return selected_profile
            elif choice == len(profiles) + 1:
                # 새 프로필 생성
                while True:
                    name = input("새 프로필 이름을 입력하세요: ")
                    if not name:
                        print("프로필 이름을 입력해주세요.")
                        continue
                        
                    if any(c in r'\\/:*?"<>|' for c in name):
                        print("프로필 이름에 다음 문자를 사용할 수 없습니다: \\ / : * ? \" < > |")
                        continue
                        
                    new_profile_path = os.path.join(user_data_parent, name)
                    if os.path.exists(new_profile_path):
                        print(f"'{name}' 프로필이 이미 존재합니다.")
                        continue
                        
                    try:
                        os.makedirs(new_profile_path)
                        os.makedirs(os.path.join(new_profile_path, 'Default'))
                        print(f"'{name}' 프로필이 생성되었습니다.")
                        return name
                    except Exception as e:
                        print(f'프로필 생성 중 오류가 발생했습니다: {e}')
                        retry = input("다시 시도하시겠습니까? (y/n): ").lower()
                        if retry != 'y':
                            break
            else:
                print("유효하지 않은 번호입니다. 다시 선택해주세요.")
        except ValueError:
            print("숫자를 입력해주세요.")

def main():
    # 사용자 프로필 경로 설정
    user_data_parent = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_data")
    
    # 프로필 선택
    selected_profile = select_profile(user_data_parent)
    if not selected_profile:
        print("프로필을 선택할 수 없습니다. 프로그램을 종료합니다.")
        return
        
    user_data_dir = os.path.join(user_data_parent, selected_profile)
    
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)
        os.makedirs(os.path.join(user_data_dir, 'Default'))

    # Google Sheets API 서비스 초기화
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # DM 목록 시트 선택
    selected_sheet = select_sheet(service, DM_LIST_SPREADSHEET_ID)
    if not selected_sheet:
        print("시트를 선택할 수 없습니다. 프로그램을 종료합니다.")
        return

    # URL과 이름 데이터 가져오기
    url_name_pairs = get_data_from_sheets(service, selected_sheet)
    if not url_name_pairs:
        print("데이터를 가져올 수 없습니다. 프로그램을 종료합니다.")
        return

    # 템플릿 시트 선택 (필터링 없이 모든 시트 표시)
    print("\n템플릿 시트를 선택하세요:")
    template_sheet_names = get_sheet_names(service, TEMPLATE_SPREADSHEET_ID)
    if not template_sheet_names:
        print("템플릿 시트 목록을 가져올 수 없습니다.")
        return
        
    print("\n사용 가능한 템플릿 시트 목록:")
    for idx, name in enumerate(template_sheet_names, 1):
        print(f"{idx}. {name}")
        
    while True:
        try:
            choice = int(input("\n사용할 템플릿 시트 번호를 선택하세요: "))
            if 1 <= choice <= len(template_sheet_names):
                template_sheet = template_sheet_names[choice - 1]
                print(f"\n선택된 템플릿 시트: {template_sheet}")
                break
            else:
                print("유효하지 않은 번호입니다. 다시 선택해주세요.")
        except ValueError:
            print("숫자를 입력해주세요.")
    
    if not template_sheet:
        print("템플릿 시트를 선택할 수 없습니다. 프로그램을 종료합니다.")
        return

    # 메시지 템플릿 매니저 초기화
    template_manager = InstagramMessageTemplate(TEMPLATE_SPREADSHEET_ID, template_sheet)
    
    # DM 발송 전 확인
    print(f"\n선택된 설정:")
    print(f"• 프로필: {selected_profile}")
    print(f"• DM 목록 시트: {selected_sheet}")
    print(f"• 템플릿 시트: {template_sheet}")
    print(f"• 발송할 DM 수: {len(url_name_pairs)}개")
    
    if len(url_name_pairs) == 0:
        print("\n발송할 DM이 없습니다. 프로그램을 종료합니다.")
        return
    
    # 발송할 인플루언서 목록 표시
    print("\n발송할 인플루언서 목록:")
    for idx, (url, name, notion_list, g_value, total_list, row_num) in enumerate(url_name_pairs, 1):
        print(f"{idx}. {name}")
    
    confirm = input("\nDM을 발송하시겠습니까? (Y/N): ").upper()
    if confirm != 'Y':
        print("DM 발송을 취소합니다.")
        return

    # Chrome 옵션 설정
    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)
    options.add_argument("disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument(f"user-data-dir={user_data_dir}")
    options.add_argument("--disable-application-cache")
    options.add_argument("--disable-cache")

    # 캐시와 임시 파일 정리 (로그인 정보 유지)
    clear_chrome_data(user_data_dir)

    print("\nDM 발송을 시작합니다...")
    
    # 각 URL 처리
    driver = None
    for index, (url, name, notion_list, g_value, total_list, row_num) in enumerate(url_name_pairs, 1):
        print(f"\n[{index}/{len(url_name_pairs)}] {name}")
        
        # 2~7명마다 랜덤하게 브라우저 재시작
        if driver is None or index % random.randint(2, 7) == 0:
            print(f"\n--- 브라우저 세션 재시작 (처리된 인원: {index}명) ---")
            
            # 기존 드라이버가 있으면 종료
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            
            # 캐시 정리
            clear_chrome_data(user_data_dir)
            print("캐시 정리 완료")
            
            # 새 Chrome 드라이버 시작
            driver = webdriver.Chrome(options=options)
            print("새 Chrome 드라이버 시작 완료")
            
            # Instagram 접속 (로그인 상태 확인)
            driver.get("https://www.instagram.com/")
            wait_time = random.uniform(3, 8)
            countdown(wait_time, "Instagram 접속 후 대기")
            
            # Instagram 메인에서 자연스러운 스크롤 동작 (실제 사용자처럼 둘러보기)
            print("Instagram 메인에서 자연스러운 스크롤 동작 시작...")
            natural_scroll_on_instagram(driver)
        
        process_url(driver, url, name, template_manager, row_num, service, selected_sheet, user_data_dir)
        wait_time = random.uniform(3, 30)
        countdown(wait_time, "다음 URL로 이동하기 전 대기")

    # 최종 브라우저 세션 종료
    if driver:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    main()
