# 변수

# # output_sheet_name 
#  - 'output'라고 하드코딩
#  - 1.유튜버메일주소추출 스프레드시트에서 원하는 시트명을 변경입력
#  - T열에 채널주소 양식 체크할 것

# # comment_sheet_name
#  - '협찬문의' 하드코딩 >> ### 댓글유형
#  - 2번 댓글유형 선택
 
# # commentlog_sheet_name
#  -  "시트1" 하드코딩 >> # 상품 따라서 시트 나눠서 댓글남겨둬도 됨
#  -  3번 댓글로그


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import random
from auth import get_credentials
from googleapiclient.discovery import build
import logging
from datetime import datetime

def get_data_from_sheets(spreadsheet_id, range_name):
    logging.info("스프레드시트에서 데이터 가져오기 시작")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        data = result.get('values', [])
        return data
    except Exception as e:
        logging.error(f"데이터 가져오기 실패: {str(e)}")
        return None

def append_to_sheets(spreadsheet_id, range_name, values):
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        body = {
            'values': values
        }
        result = sheet.values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        logging.info(f"{result.get('updates').get('updatedCells')} cells appended.")
    except Exception as e:
        logging.error(f"데이터 추가 실패: {str(e)}")

def update_cell_value(spreadsheet_id, range_name, value):
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        body = {
            'values': [[value]]
        }
        result = sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        logging.info(f"셀 업데이트 완료: {range_name} = {value}")
    except Exception as e:
        logging.error(f"셀 업데이트 실패: {str(e)}")

def get_sheet_list(spreadsheet_id):
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.get(spreadsheetId=spreadsheet_id).execute()
        sheets = result.get('sheets', [])
        sheet_names = [sheet['properties']['title'] for sheet in sheets]
        return sheet_names
    except Exception as e:
        logging.error(f"시트 목록 가져오기 실패: {str(e)}")
        return []

# URL 설정
url = "https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fwww.youtube.com%2Fsignin%3Faction_handle_signin%3Dtrue%26app%3Ddesktop%26hl%3Dko%26next%3Dhttps%253A%252F%252Fwww.youtube.com%252F&ec=65620&hl=ko&ifkv=ARpgrqf_A3Y62wfntR2UMjMrEHh8j1uNYKyX03-wetYe3rXSilMDOZwsIE_HZI7A0YodvJHsekBljw&passive=true&service=youtube&uilel=3&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S-519112804%3A1727418941441319&ddm=0"

# Chrome 옵션 설정
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("disable-blink-features=AutomationControlled")

# 웹드라이버 초기화
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)
long_wait = WebDriverWait(driver, 60)  # 60초 대기를 위한 새로운 WebDriverWait 객체

# 스프레드시트에서 데이터 가져오기
# 댓글 달 인원이 나타난 시트
output_sheet_id = '1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0'
output_sheet_name = '유튜브_현황판'
output_data = get_data_from_sheets(output_sheet_id, f'{output_sheet_name}!A2:T')  # 헤더를 제외하고 데이터 가져오기
# 댓글 내용 불러오기
comment_sheet_id = '1kJZDY5k-E980WVc1c5XLpImv96pulIGJ4xWY0Fp7XCo'

# 시트 목록 가져오기
print("댓글 템플릿 시트를 선택해주세요:")
sheet_list = get_sheet_list(comment_sheet_id)
for i, sheet_name in enumerate(sheet_list, 1):
    print(f"{i}. {sheet_name}")

# 사용자 입력 받기
while True:
    try:
        choice = int(input("시트 번호를 입력하세요: "))
        if 1 <= choice <= len(sheet_list):
            comment_sheet_name = sheet_list[choice - 1]
            print(f"선택된 시트: {comment_sheet_name}")
            break
        else:
            print(f"1부터 {len(sheet_list)}까지의 번호를 입력해주세요.")
    except ValueError:
        print("올바른 숫자를 입력해주세요.")

comment_data = get_data_from_sheets(comment_sheet_id, f'{comment_sheet_name}!A1:A15')

# 로그 스프레드시트 설정
commentlog_sheet_id = "1tRe1j-7cFsC9J1rj0I3kqVS4XuJlJiLrOZveORp_85s"
commentlog_sheet_name = "시트1"

try:
    # 페이지 로드 및 로그인
    driver.get(url)
    print(driver.title)

    # 이메일 입력
    email_field = wait.until(EC.presence_of_element_located((By.ID, "identifierId")))
    email_field.send_keys("bnam91@goyamkt.com")
    time.sleep(3)  # 3초 대기

    # '다음' 버튼 클릭
    next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='다음']")))
    next_button.click()
    time.sleep(3)  # 3초 대기

    # 비밀번호 입력
    password_field = wait.until(EC.presence_of_element_located((By.NAME, "Passwd")))
    password_field.send_keys("@rhdi120")
    time.sleep(3)  # 3초 대기

    # '다음' 버튼 다시 클릭
    next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='다음']")))
    next_button.click()
    time.sleep(3)  # 3초 대기

    print("로그인 완료. 채널 선택 화면을 기다리는 중...")

    # 채널 선택 화면 대기 (최대 15초)
    try:
        channel_switcher = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "ytd-channel-switcher-renderer")))
        print("채널 선택 화면이 나타났습니다.")
        
        # 15초 카운트다운
        print("15초 후 자동으로 다음 작업으로 넘어갑니다. 원하는 채널을 선택하세요!")
        for i in range(15, 0, -1):
            print(f"남은 시간: {i}초", end='\r', flush=True)
            time.sleep(1)
        print("\n시간이 종료되었습니다. 다음 작업으로 넘어갑니다.")
    except:
        print("15초 내에 채널 선택 화면이 나타나지 않았습니다. 채널이 이미 선택된 것으로 간주합니다.")
        print("다음 작업으로 넘어갑니다.")
    
    # 실제 스프레드시트 행 번호 추적 (헤더 제외하고 2행부터 시작)
    actual_row = 2
    
    for row in output_data:
        try:
            youtuber_name = row[2]  # C열의 {이름} 데이터
            channel_id = row[10]  # K열의 채널 아이디 데이터
            l_column = row[11]  # L열의 값
            n_column = row[13]  # N열의 값
            is_shorts = False  # is_shorts 변수 초기화
            current_url = ""  # current_url 변수 초기화
            
            if not channel_id:
                print(f"{youtuber_name}의 채널 아이디가 없습니다. 다음 행으로 넘어갑니다.")
                continue
                
            # 댓글 조건 체크: L열이 '-'이고 N열이 '시딩'이어야 함
            if l_column != '-' or n_column != '시딩':
                print(f"{youtuber_name}의 조건이 맞지 않습니다. (L열: {l_column}, N열: {n_column}) 다음 행으로 넘어갑니다.")
                continue

            # 채널 URL 조합
            base_url = f"https://www.youtube.com/{channel_id}"
            # 동영상 탭부터 시작
            target_url = f"{base_url}/shorts"
            driver.get(target_url)
            print(f"다음 링크로 이동했습니다: {target_url}")
            
            # 페이지 로드 대기
            wait.until(EC.presence_of_element_located((By.ID, "contents")))
            print("페이지가 로드되었습니다.")

            try:
                # 쇼츠 탭에서 최근 영상 찾기 시도
                recent_video = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "ytd-rich-item-renderer")))
                print("쇼츠 탭에서 최근 영상을 찾았습니다.")
            except:
                # 쇼츠 탭에서 찾지 못한 경우 동영상 탭으로 전환
                print("동영상 탭에서 최근 영상을 찾지 못했습니다. Shorts 탭으로 전환합니다.")
                target_url = f"{base_url}/videos"
                driver.get(target_url)
                print(f"다음 링크로 이동했습니다: {target_url}")
                
                # 동영상 탭 페이지 로드 대기
                wait.until(EC.presence_of_element_located((By.ID, "contents")))
                print("동영상 탭 페이지가 로드되었습니다.")
                
                # 동영상 탭에서 최근 영상 찾기
                recent_video = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "ytd-grid-video-renderer")))
                print("동영상 탭에서 최근 영상을 찾았습니다.")

            # 최근 영상 클릭
            recent_video.click()
            print("가장 최근 영상을 클릭했습니다.")


            # 현재 URL 출력
            current_url = driver.current_url
            print(f"현재 영상의 URL: {current_url}")

#==================


            # 댓글 버튼 찾기 및 클릭
            print("[디버그] 댓글 버튼 찾는 중...")
            comment_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#comments-button")))
            print("[디버그] 댓글 버튼 클릭")
            comment_button.click()
            time.sleep(random.uniform(2, 3))

            # 댓글 입력 필드 찾기
            print("[디버그] 댓글 입력 필드 찾는 중...")
            comment_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#simplebox-placeholder")))
            print("[디버그] 댓글 입력 필드 클릭")
            comment_input.click()
            time.sleep(random.uniform(1, 2))

            # 랜덤으로 댓글 선택
            print("[디버그] 랜덤 댓글 선택 중...")
            # 빈 값이 아닌 댓글만 필터링
            valid_comments = [comment[0] for comment in comment_data if comment and comment[0] and comment[0].strip()]
            random_comment = random.choice(valid_comments) if valid_comments else "감사히 잘봤습니다."

            # {이름} 변수 대체
            random_comment = random_comment.replace("{이름}", youtuber_name)
            print(f"[디버그] 선택된 댓글: {random_comment}")

            # 실제 입력 필드에 댓글 입력
            print("[디버그] 댓글 입력 중...")
            comment_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "ytd-comment-simplebox-renderer #contenteditable-root")))
            comment_box.send_keys(random_comment)
            time.sleep(random.uniform(2, 3))

            # 댓글 제출 버튼 클릭
            print("[디버그] 댓글 제출 버튼 찾는 중...")
            try:
                # 새로운 댓글 제출 버튼 선택자 시도
                submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "ytd-commentbox #submit-button")))
                print("[디버그] 새로운 댓글 제출 버튼 찾음")
            except:
                try:
                    # 기존 댓글 제출 버튼 선택자 시도
                    submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "ytd-comment-simplebox-renderer #submit-button")))
                    print("[디버그] 기존 댓글 제출 버튼 찾음")
                except:
                    # 가장 일반적인 댓글 버튼 선택자
                    submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#submit-button")))
                    print("[디버그] 일반 댓글 제출 버튼 찾음")
            
            print("[디버그] 댓글 제출 버튼 클릭")
            # submit_button.click()

            print(f"{youtuber_name}에게 댓글을 성공적으로 게시했습니다.")

            # N열을 '시딩(댓글완)'으로 업데이트
            update_cell_value(output_sheet_id, f'{output_sheet_name}!N{actual_row}', '시딩(댓글완)')
            print(f"{youtuber_name}의 N{actual_row}열이 '시딩(댓글완)'으로 업데이트되었습니다.")

            # 로그 데이터 준비
            comment_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            shorts_status = "O" if is_shorts else ""
            log_data = [[current_url, comment_date, youtuber_name, shorts_status]]

            # 로그 데이터를 스프레드시트에 추가
            append_to_sheets(commentlog_sheet_id, f'{commentlog_sheet_name}!A1:D', log_data)
            print("로그 데이터가 스프레드시트에 추가되었습니다.")
            

            # 2~80초 랜덤 대기
            wait_time = random.uniform(2, 80)
            print(f"{wait_time:.2f}초 대기 중...")
            time.sleep(wait_time)

        except Exception as e:
            print(f"오류가 발생했습니다: {str(e)}")
            # 실패 로그 추가
            failed_log = [[current_url, "failed", youtuber_name, "O" if is_shorts else ""]]
            append_to_sheets(commentlog_sheet_id, f'{commentlog_sheet_name}!A1:D', failed_log)
            print("실패 로그가 스프레드시트에 추가되었습니다.")
        
        print(f"{youtuber_name}에 대한 작업이 완료되었습니다. 다음 행으로 넘어갑니다.")
        actual_row += 1
        

except Exception as e:
    print("전체 프로세스에서 오류가 발생했습니다:", str(e))

finally:
    print("모든 작업이 완료되었습니다.")
    # 브라우저를 열린 상태로 유지
    input("Press Enter to close the browser...")
    driver.quit()