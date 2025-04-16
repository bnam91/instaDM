# 인스타그램 자동 DM 발송 프로그램
# 기능:
# 1. 구글 스프레드시트에서 인스타그램 프로필 URL과 사용자 이름 목록을 가져옴
#    - 스프레드시트 ID: 1VhEWeQASyv02knIghpcccYLgWfJCe2ylUnPsQ_-KNAI
#    - 시트 이름: dm_list
#    - 데이터 구조: A열(URL), B열(이름), C열(발송상태), D열(발송시간)
# 2. 다른 스프레드시트에서 DM 메시지 템플릿을 무작위로 선택
#    - 스프레드시트 ID: 1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0
#    - 시트 이름: 협찬문의
#    - 데이터 구조: A1:A15 셀에 메시지 템플릿 목록
#    - 템플릿 내 {이름} 태그는 실제 사용자 이름으로 대체됨
# 3. 각 프로필을 방문하여 자동으로 DM 메시지 발송
#    - 실제 발송은 현재 주석 처리되어 있음 (actions.send_keys(Keys.ENTER).perform())
# 4. 메시지 발송 결과와 시간을 스프레드시트에 기록
#    - 성공 시: 'Y' + 타임스탬프
#    - 실패 시: 'failed'
# 5. 브라우저 캐시 관리 및 자동화 감지 회피 기능 포함
#    - 로그인 정보는 유지하면서 캐시만 정리
#    - 작업 간 랜덤한 시간 간격 추가
# 6. MongoDB에 DM 발송 기록 저장
#    - 데이터베이스: insta09_database
#    - 컬렉션: gogoya_DmRecords
#    - 기록 정보: 인플루언서 이름, 프로필, 상태, 발송시간, 템플릿, 메시지 내용
# 작성일: v2 버전

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
# PyQt5 프로필 선택기 임포트
from dm_ui import select_profile_gui
# 릴리즈 업데이트 임포트
from release_updater import ReleaseUpdater
# MongoDB 관련 임포트
from pymongo import MongoClient
from pymongo.server_api import ServerApi

# 환경 변수 로드
load_dotenv()

# GitHub 저장소 정보 설정 (환경 변수에서 가져오거나 기본값 사용)
owner = os.environ.get("GITHUB_OWNER", "bnam91")
repo = os.environ.get("GITHUB_REPO", "instaDM")

# 최신 버전 확인 및 업데이트
try:
    print("📦 버전 확인 중...")
    updater = ReleaseUpdater(owner=owner, repo=repo)
    update_success = updater.update_to_latest()
    
    if update_success:
        print("✅ 최신 버전으로 업데이트되었거나 이미 최신 버전입니다.")
    else:
        print("⚠️ 업데이트 실패, 이전 버전으로 계속 진행합니다...")
except Exception as e:
    print(f"❌ 버전 확인 중 오류 발생: {e}")

# MongoDB 연결 설정
uri = "mongodb+srv://coq3820:JmbIOcaEOrvkpQo1@cluster0.qj1ty.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
try:
    mongo_client = MongoClient(uri, server_api=ServerApi('1'))
    # 연결 확인
    mongo_client.admin.command('ping')
    print("MongoDB 연결 성공!")
    
    # 데이터베이스와 컬렉션 선택
    db = mongo_client['insta09_database']
    dm_collection = db['gogoya_DmRecords']
    mongo_connected = True
except Exception as e:
    print(f"MongoDB 연결 실패: {e}")
    mongo_connected = False

def select_user_profile():
    # 현재 스크립트 위치 기준으로 user_data 폴더 경로 생성
    script_dir = os.path.dirname(os.path.abspath(__file__))
    user_data_parent = os.path.join(script_dir, "user_data")
    
    # GUI를 통해 프로필 선택 또는 생성
    result = select_profile_gui(user_data_parent)
    
    if not result:
        print("프로필이 선택되지 않았습니다.")
        return None, None, None
    
    print(f"\n선택된 프로필 경로: {result['profile_path']}")
    print(f"선택된 DM 목록 시트: {result['dm_list_sheet']}")
    print(f"선택된 템플릿 시트: {result['template_sheet']}")
    
    return result['profile_path'], result['dm_list_sheet'], result['template_sheet']

# 스프레드시트 ID를 환경 변수에서 가져오기
DM_LIST_SPREADSHEET_ID = os.getenv('DM_LIST_SPREADSHEET_ID', '1VhEWeQASyv02knIghpcccYLgWfJCe2ylUnPsQ_-KNAI')
TEMPLATE_SPREADSHEET_ID = os.getenv('TEMPLATE_SPREADSHEET_ID', '1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0')

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

options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)
options.add_argument("disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-logging"])

# 사용자 프로필 선택
user_data_dir, dm_list_sheet, template_sheet = select_user_profile()
if not user_data_dir:
    print("프로필 선택 오류. 프로그램을 종료합니다.")
    sys.exit(1)

# 프로필 폴더명 추출 (경로의 마지막 부분)
profile_name = os.path.basename(user_data_dir)
print(f"사용 중인 프로필명: {profile_name}")

options.add_argument(f"user-data-dir={user_data_dir}")

# 캐시와 임시 파일 정리 (로그인 정보 유지)
clear_chrome_data(user_data_dir)

# 추가 옵션 설정
options.add_argument("--disable-application-cache")
options.add_argument("--disable-cache")

driver = webdriver.Chrome(options=options)

def get_data_from_sheets():
    logging.info("URL과 이름 데이터 가져오기 시작")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        # 선택한 시트 사용
        result = sheet.values().get(spreadsheetId=DM_LIST_SPREADSHEET_ID,
                                    range=f'{dm_list_sheet}!A2:B').execute()
        values = result.get('values', [])

        if not values:
            logging.warning('스프레드시트에서 데이터를 찾을 수 없습니다.')
            return []

        return [(row[0], row[1] if len(row) > 1 else "") for row in values if row]
    except Exception as e:
        logging.error(f"스프레드시트에서 데이터를 가져오는 중 오류 발생: {e}")
        return []

def get_message_templates():
    logging.info("메시지 템플릿 가져오기 시작")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        # 선택한 템플릿 시트 사용
        result = sheet.values().get(spreadsheetId=TEMPLATE_SPREADSHEET_ID,
                                    range=f'{template_sheet}!A1:A15').execute()
        values = result.get('values', [])

        if not values:
            logging.warning('메시지 템플릿을 찾을 수 없습니다.')
            return ["안녕하세요"]

        return [row[0] for row in values if row]
    except Exception as e:
        logging.error(f"메시지 템플릿을 가져오는 중 오류 발생: {e}")
        return ["안녕하세요"]

def update_sheet_status(service, row, status, timestamp=None):
    sheet_id = DM_LIST_SPREADSHEET_ID
    # 선택한 시트 사용
    range_name = f'{dm_list_sheet}!C{row}:D{row}'
    
    values = [[status, timestamp if timestamp else '']]
    body = {'values': values}
    
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()

def save_dm_record_to_mongodb(influe_name, contact_profile, status, dm_date, content, message):
    """MongoDB에 DM 발송 기록을 저장하는 함수"""
    if not mongo_connected:
        print("MongoDB에 연결되지 않아 기록을 저장할 수 없습니다.")
        return False
    
    try:
        dm_record = {
            "influe_name": influe_name,
            "contact_profile": contact_profile,
            "status": status,
            "DM_date": dm_date,
            "content": content,
            "message": message
        }
        
        dm_collection.insert_one(dm_record)
        print(f"MongoDB에 DM 기록이 저장되었습니다: {influe_name}")
        return True
    except Exception as e:
        print(f"MongoDB에 DM 기록 저장 중 오류 발생: {e}")
        return False

def process_url(driver, url, name, message_template, row, service):
    driver.get(url)
    print(driver.title)
    wait_time = random.uniform(5, 10)
    print(f"URL 접속 후 대기: {wait_time:.2f}초")
    time.sleep(wait_time)

    try:
        message_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x1i10hfl') and contains(text(), '메시지 보내기')]"))
        )
        print(f"버튼 텍스트: {message_button.text}")
        message_button.click()
        wait_time = random.uniform(5, 10)
        print(f"DM 버튼 클릭 후 대기: {wait_time:.2f}초")
        time.sleep(wait_time)

        message = message_template.replace("{이름}", name)
        
        # 수정된 부분: 클립보드를 사용하여 메시지 전체를 한 번에 붙여넣기
        pyperclip.copy(message)
        actions = ActionChains(driver)
        # 텍스트 입력 필드에 포커스
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@role, 'textbox')]"))
        ).click()
        # 붙여넣기 단축키 사용 (Ctrl+V 또는 Command+V)
        if sys.platform == 'darwin':  # macOS
            actions.key_down(Keys.COMMAND).send_keys('v').key_up(Keys.COMMAND).perform()
        else:  # Windows/Linux
            actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        
        wait_time = random.uniform(5, 10)
        print(f"메시지 입력 후 대기: {wait_time:.2f}초")
        time.sleep(wait_time)

        # Enter 키를 눌러 메시지 전송 (엔터)
        actions.send_keys(Keys.ENTER).perform()

        # 성공적으로 메시지를 보냈을 때
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        update_sheet_status(service, row, 'Y', timestamp)
        
        # MongoDB에 DM 기록 저장
        save_dm_record_to_mongodb(
            influe_name=name,
            contact_profile=profile_name,  # 프로필 폴더명으로 수정
            status='Y',
            dm_date=timestamp,
            content=template_sheet,
            message=message
        )

    except TimeoutException:
        print("'메시지 보내기' 버튼을 찾을 수 없습니다.")
        update_sheet_status(service, row, 'failed')
        
        # 실패 정보도 MongoDB에 저장
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_dm_record_to_mongodb(
            influe_name=name,
            contact_profile=profile_name,  # 프로필 폴더명으로 수정
            status='failed',
            dm_date=timestamp,
            content=template_sheet,
            message="메시지 보내기 버튼을 찾을 수 없음"
        )
    except NoSuchElementException:
        print("요소를 찾을 수 없습니다.")
        update_sheet_status(service, row, 'failed')
        
        # 실패 정보도 MongoDB에 저장
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_dm_record_to_mongodb(
            influe_name=name,
            contact_profile=profile_name,  # 프로필 폴더명으로 수정
            status='failed',
            dm_date=timestamp,
            content=template_sheet,
            message="요소를 찾을 수 없음"
        )

# 메인 실행 부분
message_templates = get_message_templates()
url_name_pairs = get_data_from_sheets()

creds = get_credentials()
service = build('sheets', 'v4', credentials=creds)

for index, (url, name) in enumerate(url_name_pairs, start=2):  # start=2 because row 1 is header
    message_template = random.choice(message_templates)
    process_url(driver, url, name, message_template, index, service)
    time.sleep(5)  # 다음 URL로 이동하기 전 5초 대기

# 브라우저를 닫지 않고 세션 유지
driver.quit()  # 필요한 경우 주석 해제
