from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import time
import urllib.parse
import re
from datetime import datetime
from auth import get_credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = '1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0' #분할버전

def get_keywords_from_sheet():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='유튜브_키워드!A2:A'
    ).execute()
    values = result.get('values', [])

    if not values:
        print('스프레드시트에서 키워드를 찾을 수 없습니다.')
        return []
    else:
        return [row[0] for row in values if row]

def clear_channel_id_sheet():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # 헤더만 지우기 (A1:N1)
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range='유튜브_현황판!A1:N1'
    ).execute()

    print("'유튜브_현황판' 시트의 헤더가 지워졌습니다.")

def write_header_to_sheet():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # 먼저 헤더 행을 지우기
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range='유튜브_현황판!A1:N1'
    ).execute()

    header = [['검색어', '전체영상수', '채널명', '구독자수', '채널소개', '제목', '업로드날짜', 'URL', '조회수', '영상길이', '채널아이디', '메일주소', '추가날짜']]

    body = {
        'values': header
    }

    # 헤더 값 입력
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range='유튜브_현황판!A1',
        valueInputOption='RAW',
        body=body
    ).execute()

    # 헤더 스타일 적용 (굵은 글씨, 배경색)
    requests = [
        {
            'repeatCell': {
                'range': {
                    'sheetId': 0,  # 첫 번째 시트
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 13
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.8,
                            'green': 0.8,
                            'blue': 0.8
                        },
                        'textFormat': {
                            'bold': True,
                            'fontSize': 12
                        },
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE'
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': requests}
    ).execute()

    print("헤더가 '유튜브_현황판' 시트에 작성되었습니다.")

def get_existing_channel_ids():
    """기존 데이터에서 채널아이디 목록을 가져오기"""
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='유튜브_현황판!K:K'  # 채널아이디 컬럼
        ).execute()
        values = result.get('values', [])
        
        if not values:
            return set()
        else:
            # 헤더 제외하고 채널아이디만 추출
            return set(row[0] for row in values[1:] if row and row[0])
    except Exception as e:
        print(f"기존 채널아이디 가져오기 오류: {e}")
        return set()

def append_batch_to_sheet(data_batch):
    if not data_batch:
        return
        
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # 기존 채널아이디 목록 가져오기
    existing_channel_ids = get_existing_channel_ids()
    
    # 중복되지 않은 데이터만 필터링
    filtered_data = []
    duplicate_count = 0
    current_batch_channel_ids = set()  # 현재 배치 내 중복 체크용
    
    for item in data_batch:
        channel_id = item.get('채널아이디', '')
        if channel_id and channel_id not in existing_channel_ids and channel_id not in current_batch_channel_ids:
            filtered_data.append(item)
            current_batch_channel_ids.add(channel_id)  # 현재 배치에 추가
        else:
            duplicate_count += 1
    
    if duplicate_count > 0:
        print(f"기존 시트와 중복된 채널아이디 {duplicate_count}건 제외")
    
    if not filtered_data:
        print("추가할 새로운 데이터가 없습니다.")
        return

    values = [list(row.values()) for row in filtered_data]

    body = {
        'values': values
    }

    # 데이터는 기본 스타일로 추가
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range='유튜브_현황판!A1',
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()

    # 최종 시트 로우 수 확인
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='유튜브_현황판!A:A'
        ).execute()
        total_rows = len(result.get('values', []))
        print(f"{len(filtered_data)}개의 새로운 결과가 '유튜브_현황판' 시트에 추가되었습니다. (총 {total_rows}개 로우)")
    except Exception as e:
        print(f"{len(filtered_data)}개의 새로운 결과가 '유튜브_현황판' 시트에 추가되었습니다.")

def scroll_to_bottom(driver):
    """페이지 끝까지 스크롤하여 모든 데이터를 로드"""
    print("페이지 끝까지 스크롤 중...")
    last_height = driver.execute_script("return document.documentElement.scrollHeight")
    no_new_content_count = 0
    scroll_count = 0
    
    while True:
        # 페이지 끝까지 스크롤
        driver.execute_script("window.scrollTo(0, document.documentElement.scrollHeight);")
        time.sleep(1)  # 로딩 대기
        
        # 새로운 높이 계산
        new_height = driver.execute_script("return document.documentElement.scrollHeight")
        scroll_count += 1
        
        # 높이가 변하지 않으면 새로운 콘텐츠가 로드되지 않음
        if new_height == last_height:
            no_new_content_count += 1
            if no_new_content_count >= 3:  # 3번 연속으로 높이가 같으면 종료
                break
        else:
            no_new_content_count = 0
            last_height = new_height
    
    print(f"스크롤 완료! (총 {scroll_count}회 스크롤, 최종 높이: {new_height})")

def convert_views(views):
    if views == '없음' or views == '' or views == '정보 없음':
        return '0'
    
    views = views.replace('조회수 ', '').replace('회', '').replace(',', '')
    if '만' in views:
        number = int(float(views.replace('만', '')) * 10000)
        return f"{number:,}"
    elif '천' in views:
        number = int(float(views.replace('천', '')) * 1000)
        return f"{number:,}"
    else:
        try:
            number = int(views)
            return f"{number:,}"
        except ValueError:
            return '0'

def parse_upload_date(upload_date):
    """업로드 날짜를 파싱하여 일수로 변환"""
    try:
        if '시간' in upload_date:
            hours = int(upload_date.replace('시간 전', ''))
            return hours / 24  # 일수로 변환
        elif '일' in upload_date:
            days = int(upload_date.replace('일 전', ''))
            return days
        elif '주' in upload_date:
            weeks = int(upload_date.replace('주 전', ''))
            return weeks * 7
        elif '개월' in upload_date:
            months = int(upload_date.replace('개월 전', ''))
            return months * 30  # 대략적인 일수
        elif '년' in upload_date:
            years = int(upload_date.replace('년 전', ''))
            return years * 365  # 대략적인 일수
        else:
            return float('inf')  # 파싱할 수 없는 경우 무한대로 설정
    except:
        return float('inf')

def is_within_3_months(upload_date):
    """3개월(90일) 이내인지 확인"""
    days = parse_upload_date(upload_date)
    return days <= 90

def parse_duration(duration):
    """영상 길이를 파싱하여 초 단위로 변환"""
    if duration == "Shorts":
        return float('inf')  # Shorts는 항상 포함
    
    try:
        # "17초", "3초" 등의 형식에서 숫자만 추출
        seconds = int(duration.replace('초', ''))
        return seconds
    except:
        return float('inf')  # 파싱할 수 없는 경우 무한대로 설정

def is_valid_duration(duration):
    """15초 초과이거나 Shorts인지 확인"""
    if duration == "Shorts":
        return True
    
    seconds = parse_duration(duration)
    return seconds > 15

def process_element(element, keyword):
    try:
        # 비디오 링크 추출
        try:
            video_link = element.find_element(By.ID, "thumbnail").get_attribute("href")
        except NoSuchElementException:
            # 대체 방법으로 링크 찾기
            try:
                video_link = element.find_element(By.CSS_SELECTOR, "a#video-title").get_attribute("href")
            except NoSuchElementException:
                print(f"{keyword}: 비디오 링크를 찾을 수 없습니다.")
                return None
        
        # 제목 추출
        try:
            title = element.find_element(By.ID, "video-title").get_attribute("title")
        except NoSuchElementException:
            try:
                title = element.find_element(By.CSS_SELECTOR, "a#video-title").text
            except NoSuchElementException:
                print(f"{keyword}: 제목을 찾을 수 없습니다.")
                return None
        
        # 채널명 추출
        try:
            channel_element = element.find_element(By.CSS_SELECTOR, "#channel-info #text-container yt-formatted-string a")
            channel_name = channel_element.text
        except NoSuchElementException:
            try:
                channel_element = element.find_element(By.CSS_SELECTOR, "yt-formatted-string#text a")
                channel_name = channel_element.text
            except NoSuchElementException:
                print(f"{keyword}: 채널명을 찾을 수 없습니다.")
                return None
        
        # 조회수와 업로드 날짜 추출
        try:
            metadata_elements = element.find_elements(By.XPATH, ".//span[@class='inline-metadata-item style-scope ytd-video-meta-block']")
            if len(metadata_elements) >= 2:
                views = metadata_elements[0].text
                upload_date = metadata_elements[1].text
            else:
                print(f"{keyword}: 메타데이터를 찾을 수 없습니다.")
                return None
        except NoSuchElementException:
            print(f"{keyword}: 메타데이터를 찾을 수 없습니다.")
            return None
        
        # 3개월 이내가 아닌 경우 None 반환
        if not is_within_3_months(upload_date):
            return None
        
        # 영상 길이 추출
        try:
            video_title_element = element.find_element(By.CSS_SELECTOR, 'a#video-title')
            aria_label = video_title_element.get_attribute('aria-label')
            duration_match = re.search(r'\d+초', aria_label)
            duration = duration_match.group() if duration_match else "Shorts"
        except NoSuchElementException:
            duration = "Shorts"
        
        # 15초 이하인 경우 None 반환 (Shorts는 제외)
        if not is_valid_duration(duration):
            return None
        
        # 채널 ID 추출
        try:
            channel_id = element.find_element(By.XPATH, ".//yt-formatted-string[@id='text']/a").get_attribute("href").split("/")[-1]
        except NoSuchElementException:
            try:
                channel_id = channel_element.get_attribute("href").split("/")[-1]
            except:
                print(f"{keyword}: 채널 ID를 찾을 수 없습니다.")
                return None
        
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # 조회수를 숫자로 변환
        converted_views = convert_views(views)
        
        # 조회수가 700 이하인 경우 None 반환
        # 쉼표 제거하고 숫자로 변환하여 비교
        views_number = int(converted_views.replace(',', ''))
        if views_number <= 1000:
            return None
        
        return {
            '검색어': keyword,
            '전체영상수': '',  # 우선 빈 값으로 설정
            '채널명': channel_name,
            '구독자수': '',  # 우선 빈 값으로 설정
            '채널소개': '',  # 우선 빈 값으로 설정
            '제목': title,
            '업로드날짜': upload_date,
            'URL': video_link,
            '조회수': converted_views,
            '영상길이': duration,
            '채널아이디': channel_id,
            '메일주소': '',  # 빈 값으로 설정
            '추가날짜': current_date
        }
    except Exception as e:
        print(f"{keyword}: 데이터 추출 중 오류 발생 - {e}")
        return None

def crawl_data_stream(driver, keyword, processed_links):
    elements = driver.find_elements(By.ID, "dismissible")
    new_data = []
    total_elements = len(elements)
    
    print(f"{keyword}: {total_elements}개의 요소 발견, 처리 시작...")
    
    for i, element in enumerate(elements, 1):
        try:
            video_link = element.find_element(By.ID, "thumbnail").get_attribute("href")
            if video_link not in processed_links:
                data = process_element(element, keyword)
                if data:
                    new_data.append(data)
                    processed_links.add(video_link)
                    
                # 10개마다 진행상태 출력 (한 줄로 업데이트)
                if i % 10 == 0 or i == total_elements:
                    print(f"\r{keyword}: {i}/{total_elements} 요소 처리 완료 (현재 {len(new_data)}개 데이터 수집)", end="", flush=True)
                    
        except Exception as e:
            print(f"{keyword}: 요소 처리 중 오류 발생 - {e}")
            
    return new_data

def main():
    keywords = get_keywords_from_sheet()
    
    if not keywords:
        print("키워드를 찾을 수 없습니다. 프로그램을 종료합니다.")
        return

    crawl_input = input("키워드 당 크롤링할 데이터의 수를 입력하세요 (max 입력시 전체 크롤링): ")
    crawl_count = float('inf') if crawl_input.lower() == 'max' else int(crawl_input)

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--log-level=3')  # 오류만 표시
    chrome_options.add_argument('--silent')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=chrome_options)

    clear_channel_id_sheet()
    write_header_to_sheet()

    for keyword in keywords:
        encoded_keyword = urllib.parse.quote(keyword)
        url = f"https://www.youtube.com/results?search_query={encoded_keyword}"
        driver.get(url)

        try:
            print(f"{keyword}: Shorts 탭으로 이동 중...")
            shorts_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="chips"]/yt-chip-cloud-chip-renderer[2]'))
            )
            shorts_tab.click()
            time.sleep(2)  # 대기시간 단축
            print(f"{keyword}: Shorts 탭으로 이동 완료")
        except Exception as e:
            print(f"{keyword}: Shorts 탭 이동 실패 - {e}")
            continue

        # 먼저 페이지 끝까지 스크롤하여 모든 데이터 로드
        scroll_to_bottom(driver)
        
        # 스크롤 완료 후 모든 데이터 처리
        processed_links = set()
        
        # 한 번에 모든 데이터 크롤링
        all_data = crawl_data_stream(driver, keyword, processed_links)
        print(f"{keyword}: 크롤링 완료 - 조건 충족 데이터 {len(all_data)}개 발견")
        
        # 모든 데이터를 한 번에 처리
        if all_data:
            append_batch_to_sheet(all_data)
        
        # crawl_count 제한이 있는 경우 처리
        if crawl_count != float('inf') and len(all_data) > crawl_count:
            print(f"{keyword}: 요청된 수량({crawl_count})을 초과하여 {len(all_data)}개를 크롤링했습니다.")

    driver.quit()

if __name__ == "__main__":
    main()