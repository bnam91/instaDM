'''
## 목록 시트
https://docs.google.com/spreadsheets/d/1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0/edit?gid=1609134358#gid=1609134358

## 템플릿 시트
https://docs.google.com/spreadsheets/d/1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0/edit?gid=1655394428#gid=1655394428


## 🚩요약
셀러별 타게팅해서 dm보내는 백엔드입니다.
메세지 보내는 부분은 모듈로 임포트 되어있습니다. (instagram_message_vendor.py)

## 템플릿 시트 변수
템플릿에서 {이름}, {노션리스트}를 변수로 받아서 메세지를 보냅니다.


## 🚩사용법
- 코드를 실행하면 프로필 선택합니다. (*사용 가능한 프로필 목록:)
- DM보낼 인원 목록 시트를 선택합니다. (*1. 사용 가능한 시트 목록:) 
    - 발송조건 - G열에 리스트생성일이 있으나, H열 DM발송 시트가 빈 값인 경우 발송송조건에 해당
    
    - ### 🚩추후 업데이트 예정사항 : 리스트생성일을 보고 발송조건 수정
    - 🚩(옵션1. 오늘날짜 / 옵션2. 최근 n일 / **🚩옵션3. 리스트생성일과 DM발송시간 차이가 n일 이상인 경우** / 옵션4. 특정 날짜 이후 / 옵션5. 옵션별 번호 선택)
    - 🚩총 200명의 인원이 있습니다. 10일 기준으로 인월들의 리스트를 생성할 것입니다. 그렇다면 하루에 20명 씩 발송할 것입니다. 그렇다면 10일 소요됩니다.

- 템플릿 시트를 선택합니다. (*2. 사용 가능한 시트 목록:)
- 최종 발송할 인플루언서 확인 후 Y를 눌러 발송합니다.

- 발송 완료 후 목록 시트에서 발송 상태를 확인할 수 있습니다. (몽고db에도 로그기록록)


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
    
    # 성공 시 'DM완료(오늘날짜)', 실패 시 'DM실패' 기록
    if status == 'Y':
        today = datetime.now().strftime("%m/%d")
        value = f'DM완료({today})'
    else:
        value = 'DM실패'
    
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

def process_url(driver, url, name, template_manager, row, service, sheet_name, user_data_dir):
    driver.get(url)
    print(driver.title)
    wait_time = random.uniform(5, 10) #원래 300초였음
    countdown(wait_time, "URL 접속 후 대기")

    try:
        # 먼저 팔로우 버튼 확인 (주석처리됨)
        # try:
        #     follow_button = WebDriverWait(driver, 5).until(
        #         EC.presence_of_element_located((By.XPATH, 
        #             "//button[.//div[contains(text(), '팔로우') or contains(text(), '팔로잉')]]"))
        #     )
        #     button_text = follow_button.find_element(By.XPATH, ".//div").text
        #     
        #     if button_text == "팔로우":
        #         follow_button.click()
        #         print("팔로우 완료")
        #         wait_time = random.uniform(4, 12)
        #         countdown(wait_time, "팔로우 후 대기")
        # except TimeoutException:
        #     print("팔로우 버튼이 없거나 이미 팔로우 중입니다.")
        #     pass

        # 메시지 보내기 버튼 찾기 (최대 3회 재시도)
        message_button = None
        retry_count = 0
        max_retries = 3
        
        while retry_count < max_retries:
            try:
                message_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x1i10hfl') and contains(text(), '메시지 보내기')]"))
                )
                print(f"버튼 텍스트: {message_button.text}")
                message_button.click()
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
        pyperclip.copy(message)
        
        actions = ActionChains(driver)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@role, 'textbox')]"))
        ).click()
        
        if sys.platform == 'darwin':
            actions.key_down(Keys.COMMAND).send_keys('v').key_up(Keys.COMMAND).perform()
        else:
            actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        
        wait_time = random.uniform(10, 20)
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

    # Chrome 드라이버 초기화
    driver = webdriver.Chrome(options=options)
        
    print("\nDM 발송을 시작합니다...")
    
    # 각 URL 처리
    for index, (url, name, notion_list, g_value, total_list, row_num) in enumerate(url_name_pairs, 1):
        print(f"\n[{index}/{len(url_name_pairs)}] {name}")
        process_url(driver, url, name, template_manager, row_num, service, selected_sheet, user_data_dir)
        wait_time = random.uniform(5, 60)
        countdown(wait_time, "다음 URL로 이동하기 전 대기")

    # 브라우저 세션 유지
    driver.quit()

if __name__ == "__main__":
    main()
