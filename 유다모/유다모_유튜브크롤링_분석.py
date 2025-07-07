from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import time
import re
import random
from auth import get_credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = '1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0'

def get_channel_ids_from_sheet():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API - 채널 ID와 메일주소를 함께 가져오기
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='유튜브_현황판!K2:L').execute()
    values = result.get('values', [])

    if not values:
        print('No data found in the spreadsheet.')
        return []
    else:
        # 메일주소가 비어있거나 빈 문자열인 경우만 처리 ('-'는 제외)
        channel_ids = []
        for row in values:
            if len(row) >= 2:  # 채널 ID와 메일주소가 모두 있는 경우
                channel_id = row[0]
                email = row[1]
                # 메일주소가 비어있거나 빈 문자열인 경우만 처리 ('-'는 제외)
                if not email or email.strip() == '':
                    channel_ids.append(channel_id)
            elif len(row) == 1:  # 채널 ID만 있는 경우 (메일주소가 없는 경우)
                channel_ids.append(row[0])
        
        print(f"처리할 채널 수: {len(channel_ids)}개 (메일주소가 비어있는 채널)")
        return channel_ids

def update_sheet_with_result(channel_info, videos_info, service, is_first_run, channel_id):
    sheet = service.spreadsheets()

    # Define headers
    headers = [
        "검색어", "전체영상수", "채널명", "구독자수", "채널소개", "제목", "업로드날짜", "URL", "조회수", "영상길이", "채널아이디", "메일주소", "추가날짜"
    ]

    # Prepare the data (기존 헤더 구조에 맞춰 배치)
    row_data = [
        '',  # 검색어 (A열)
        channel_info.get('총 영상 수', ''),  # 전체영상수 (B열)
        '',  # 채널명 (C열) - 기존 데이터 유지
        convert_subscribers(channel_info.get('구독자 수', '0')),  # 구독자수 (D열)
        channel_info.get('채널 소개', ''),  # 채널소개 (E열)
        '',  # 제목 (F열) - 기존 데이터 유지
        '',  # 업로드날짜 (G열) - 기존 데이터 유지
        '',  # URL (H열) - 기존 데이터 유지
        '',  # 조회수 (I열) - 기존 데이터 유지
        '',  # 영상길이 (J열) - 기존 데이터 유지
        channel_info.get('채널ID', ''),  # 채널아이디 (K열)
        '',  # 메일주소 (L열) - 기존 데이터 유지
        '',  # 추가날짜 (M열) - 기존 데이터 유지
    ]

    if is_first_run:
        # 헤더가 이미 있는지 확인
        existing_headers = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='유튜브_현황판!A1:Z1').execute()
        if not existing_headers.get('values') or not existing_headers['values'][0]:
            # 헤더가 없으면 추가
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range='유튜브_현황판!A1',
                valueInputOption='USER_ENTERED',
                body={'values': [headers]}
            ).execute()
            print("헤더가 추가되었습니다.")
            
            # 칼럼 높이 고정 설정 (더 안정적인 방법)
            try:
                # 먼저 시트 정보를 가져와서 실제 시트 ID 확인
                spreadsheet = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
                sheet_id = None
                for sheet_info in spreadsheet['sheets']:
                    if sheet_info['properties']['title'] == '유튜브_현황판':
                        sheet_id = sheet_info['properties']['sheetId']
                        break
                
                if sheet_id is not None:
                    row_height_request = {
                        'requests': [{
                            'updateDimensionProperties': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'dimension': 'ROWS',
                                    'startIndex': 0,
                                    'endIndex': 1000  # 충분히 큰 범위
                                },
                                'properties': {
                                    'pixelSize': 25  # 행 높이를 25픽셀로 고정
                                },
                                'fields': 'pixelSize'
                            }
                        }]
                    }
                    
                    sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=row_height_request).execute()
                    print("칼럼 높이가 고정되었습니다.")
                else:
                    print("시트 ID를 찾을 수 없어 높이 고정을 건너뜁니다.")
            except Exception as e:
                print(f"칼럼 높이 고정 중 오류 발생: {e}")
        else:
            print("헤더가 이미 존재합니다. 기존 데이터를 유지합니다.")
            
            # 기존 헤더가 있어도 높이 고정 시도
            try:
                spreadsheet = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
                sheet_id = None
                for sheet_info in spreadsheet['sheets']:
                    if sheet_info['properties']['title'] == '유튜브_현황판':
                        sheet_id = sheet_info['properties']['sheetId']
                        break
                
                if sheet_id is not None:
                    row_height_request = {
                        'requests': [{
                            'updateDimensionProperties': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'dimension': 'ROWS',
                                    'startIndex': 0,
                                    'endIndex': 1000
                                },
                                'properties': {
                                    'pixelSize': 25
                                },
                                'fields': 'pixelSize'
                            }
                        }]
                    }
                    
                    sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=row_height_request).execute()
                    print("기존 시트의 높이가 고정되었습니다.")
            except Exception as e:
                print(f"기존 시트 높이 고정 중 오류 발생: {e}")

    # Find the row where the channel ID exists in column K
    k_column_data = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='유튜브_현황판!K:K').execute()
    values = k_column_data.get('values', [])
    
    target_row = None
    for i, row in enumerate(values, 1):
        if row and row[0] == channel_id:
            target_row = i
            break
    
    if target_row is None:
        print(f"채널 ID {channel_id}를 찾을 수 없습니다.")
        return

    # Get existing row data to preserve it
    existing_row_data = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f'유튜브_현황판!A{target_row}:M{target_row}').execute()
    existing_values = existing_row_data.get('values', [[]])[0] if existing_row_data.get('values') else [''] * 13
    
    # Update only specific columns while preserving existing data
    updated_values = existing_values.copy()
    updated_values[1] = channel_info.get('총 영상 수', '')  # B열: 전체영상수
    updated_values[3] = format(convert_subscribers(channel_info.get('구독자 수', '0')), ',')  # D열: 구독자수 (쉼표 포함)
    updated_values[4] = channel_info.get('채널 소개', '')  # E열: 채널소개
    updated_values[10] = channel_info.get('채널ID', '')  # K열: 채널아이디
    
    # 메일주소가 없는 경우 '-'로 표시
    email = channel_info.get('이메일 주소', '')
    updated_values[11] = email if email else '-'  # L열: 메일주소

    # Update the specific row with preserved data
    range_name = f'유튜브_현황판!A{target_row}:M{target_row}'
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body={'values': [updated_values]}
    ).execute()

    # Update channel name in column C (하이퍼링크 없이)
    channel_name = channel_info.get('채널명', '')
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'유튜브_현황판!C{target_row}',
        valueInputOption='USER_ENTERED',
        body={'values': [[channel_name]]}
    ).execute()
    print(f"채널명 '{channel_name}'이 업데이트되었습니다.")

def wait_and_find_element(driver, selector, by=By.CSS_SELECTOR, timeout=20):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
        return element
    except TimeoutException:
        print(f"요소를 찾을 수 없습니다: {selector}")
        return None

def clean_text(text, remove_words=None):
    if remove_words:
        for word in remove_words:
            text = text.replace(word, '')
    return text.strip()

def extract_email(text):
    email_regex = r'[\w\.-]+@[\w\.-]+\.\w+'
    match = re.search(email_regex, text)
    return match.group(0) if match else ''

def get_channel_info(channel_id):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--remote-debugging-port=9222')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-plugins')
    options.add_argument('--disable-images')
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

    try:
        driver = webdriver.Chrome(options=options)
    except WebDriverException as e:
        print(f"웹드라이버 초기화 오류: {e}")
        return None, None, str(e)

    base_url = "https://www.youtube.com/"
    about_url = f"{base_url}{channel_id}/about"
    videos_url = f"{base_url}{channel_id}/videos"
    
    try:
        # About 페이지 정보 수집
        driver.get(about_url)
        time.sleep(3)

        # 채널 소개 추출
        channel_description = wait_and_find_element(driver, "#description-container")
        channel_description = channel_description.text.strip() if channel_description else "정보 없음"
        
        # 이메일 주소 추출
        email = extract_email(channel_description)

        # 구독자 수, 총 영상 수, 총 조회수, 가입일 추출
        info_elements = driver.find_elements(By.CSS_SELECTOR, "td.style-scope.ytd-about-channel-renderer")
        subscribers = video_count = view_count = "정보 없음"
        # join_date = "정보 없음"  # 가입일 수집 비활성화

        for element in info_elements:
            text = element.text.strip()
            if "구독자" in text:
                subscribers = clean_text(text, ['구독자'])
            elif "동영상" in text:
                video_count = clean_text(text, ['동영상', '개'])
            elif "조회수" in text:
                view_count = clean_text(text, ['조회수', '회'])
            # elif "가입일" in text:  # 가입일 수집 비활성화
            #     join_date = clean_text(text, ['가입일:'])

        # Videos 페이지 정보 수집
        driver.get(videos_url)
        time.sleep(2)

        channel_name_selectors = [
            "h1.dynamic-text-view-model-wiz__h1 span",
            ".page-header-view-model-wiz__page-header-title span",
            "yt-formatted-string#text.ytd-channel-name"
        ]
        
        channel_name = None
        for selector in channel_name_selectors:
            channel_name_element = wait_and_find_element(driver, selector)
            if channel_name_element:
                channel_name = channel_name_element.text
                break
        
        if not channel_name:
            print("채널명을 찾을 수 없습니다.")
            return None, None, "채널명을 찾을 수 없습니다."

        video_items = driver.find_elements(By.CSS_SELECTOR, "div#dismissible.style-scope.ytd-rich-grid-media")
        
        videos_info = []
        for item in video_items[:3]:
            try:
                title_element = item.find_element(By.CSS_SELECTOR, "#video-title")
                title = title_element.text

                metadata = item.find_elements(By.CSS_SELECTOR, "#metadata-line span.inline-metadata-item")
                video_view_count = metadata[0].text if len(metadata) > 0 else "정보 없음"
                upload_time = metadata[1].text if len(metadata) > 1 else "정보 없음"

                videos_info.append({
                    "제목": title,
                    "조회수": video_view_count,
                    "업로드 시간": upload_time
                })
            except NoSuchElementException:
                continue

        channel_info = {
            "채널명": channel_name,
            "채널 소개": channel_description,
            "이메일 주소": email,
            "구독자 수": subscribers,
            "총 영상 수": video_count,
            "총 조회수": view_count,
            # "가입일": join_date,  # 가입일 수집 비활성화
            "채널ID": channel_id
        }

        return channel_info, videos_info, None

    except Exception as e:
        print(f"오류 발생: {e}")
        return None, None, str(e)
    finally:
        driver.quit()

def convert_subscribers(subs):
    if subs == '정보 없음':
        return 0  # 또는 다른 적절한 기본값
    elif '만명' in subs:
        return int(float(subs.replace('만명', '')) * 10000)
    elif '천명' in subs:
        return int(float(subs.replace('천명', '')) * 1000)
    else:
        return int(subs.replace('명', '').replace(',', ''))

def convert_views(views):
    if views == '없음' or views == '':
        return 0  # 또는 다른 적절한 기본값
    
    views = views.replace('조회수 ', '').replace('회', '').replace(',', '')
    if '만' in views:
        return int(float(views.replace('만', '')) * 10000)
    elif '천' in views:
        return int(float(views.replace('천', '')) * 1000)
    else:
        try:
            return int(views)
        except ValueError:
            print(f"경고: 조회수 '{views}'를 정수로 변환할 수 없습니다. 0을 반환합니다.")
            return 0

def check_email(email):
    if not email:
        return 'NULL'
    elif email.endswith('@naver.com') or email.endswith('@gmail.com'):
        return 'TRUE'
    else:
        return 'FALSE'


if __name__ == "__main__":
    channel_ids = get_channel_ids_from_sheet()
    if not channel_ids:
        print("스프레드시트에서 channel_id를 가져오는 데 실패했습니다.")
    else:
        error_log = []
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        is_first_run = True

        for i, channel_id in enumerate(channel_ids, 1):
            print(f"\n처리 중: {i}/{len(channel_ids)} - Channel ID: {channel_id}")
            result = get_channel_info(channel_id)
            
            if result[0] is None or result[1] is None:
                error_message = f"Channel ID {channel_id}: {result[2]}"
                print(f"채널 정보를 가져오는 데 실패했습니다. {error_message}")
                error_log.append(error_message)
            else:
                channel_info, videos_info, _ = result
                
                print("\n채널 정보:")
                for key, value in channel_info.items():
                    print(f"{key}: {value}")

                print("\n최근 업로드된 영상:")
                for j, video in enumerate(videos_info, 1):
                    print(f"\n{j}. 제목: {video['제목']}")
                    print(f"   조회수: {video['조회수']}")
                    print(f"   업로드 시간: {video['업로드 시간']}")

                # 각 채널 정보를 개별적으로 시트에 업데이트
                update_sheet_with_result(channel_info, videos_info, service, is_first_run, channel_id)
                print(f"\nChannel ID {channel_id}의 결과가 스프레드시트에 업데이트되었습니다.")

                is_first_run = False

            if i < len(channel_ids):
                # 다음 세트 전 랜덤한 시간 동안 대기 (예: 1초에서 9초 사이)
                wait_time = random.uniform(1, 4)
                print(f"\n다음 채널 처리까지 {wait_time:.2f}초 대기 중...")
                time.sleep(wait_time)

        print("\n모든 채널 처리가 완료되었습니다.")
        
        if error_log:
            print("\n발생한 오류 목록:")
            for error in error_log:
                print(error)
        else:
            print("\n오류 없이 모든 채널 정보를 처리했습니다.")