import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import time
import random
import pyperclip
import pyautogui
from auth import get_credentials
from googleapiclient.discovery import build
import atexit

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

SPREADSHEET_ID = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos"

def get_data_from_sheets():
    logging.info("스프레드시트에서 데이터 가져오기 시작")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        # 이메일 템플릿 및 로그인 정보 가져오기
        result = service.spreadsheets().values().batchGet(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[
                '단체메일!A6:D6', '단체메일!A10:D10', '단체메일!A14:D14',  # 제목
                '단체메일!A8:D8', '단체메일!A12:D12', '단체메일!A16:D16',  # 본문
                '단체메일!B1', '단체메일!B2', '단체메일!B3'  # 로그인 정보 및 DB 시트 이름  
            ]
        ).execute()

        values = result.get('valueRanges', [])
        email_titles = [item for sublist in values[:3] for item in sublist.get('values', [[]])[0]]
        email_contents = [item for sublist in values[3:6] for item in sublist.get('values', [[]])[0]]
        user_id = values[6].get('values', [['']])[0][0]
        user_pw = values[7].get('values', [['']])[0][0]
        db_sheet_name = values[8].get('values', [['']])[0][0]

        # 수신자 데이터 가져오기
        recipient_data = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{db_sheet_name}!B2:C'
        ).execute().get('values', [])

        logging.info("스프레드시트에서 데이터 가져오기 완료")
        return email_titles, email_contents, user_id, user_pw, recipient_data
    except Exception as e:
        logging.error(f"스프레드시트에서 데이터 가져오기 실패: {str(e)}")
        raise

def create_driver():
    logging.info("Chrome 드라이버 생성 시작")
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        driver = webdriver.Chrome(options=chrome_options)
        logging.info("Chrome 드라이버 생성 완료")
        return driver
    except Exception as e:
        logging.error(f"Chrome 드라이버 생성 실패: {str(e)}")
        raise

def prevent_browser_close(driver):
    def keep_browser_open():
        try:
            if driver.service.process:
                logging.info("브라우저 종료 방지 시도")
                driver.execute_script("Object.defineProperty(window, 'onbeforeunload', { value: function() { return true; } });")
        except Exception as e:
            logging.error(f"브라우저 종료 방지 실패: {str(e)}")
    return keep_browser_open

def send_email(email_titles, email_contents, user_id, user_pw, recipient_data):
    logging.info("이메일 전송 프로세스 시작")
    driver = create_driver()
    keep_browser_open = prevent_browser_close(driver)
    atexit.register(keep_browser_open)
    
    driver.maximize_window()
    driver.get("https://mail.naver.com/v2/new")

    try:
        # 로그인
        logging.info("로그인 시도")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#id")))
        id_element = driver.find_element(By.CSS_SELECTOR, "#id")
        pyperclip.copy(user_id)
        id_element.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        pw_element = driver.find_element(By.CSS_SELECTOR, "#pw")
        pyperclip.copy(user_pw)
        pw_element.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        driver.find_element(By.CSS_SELECTOR, ".btn_login").click()
        logging.info("로그인 완료")

        # 메일 작성 페이지 대기
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#recipient_input_element")))

        # HTML 모드 선택 (기본값으로 설정)
        html_bt = driver.find_element(By.CSS_SELECTOR, "div.editor_mode_select button[value='HTML']")
        html_bt.click()
        time.sleep(1)

        for recipient in recipient_data:
            name, email = recipient
            logging.info(f"{email}에 대한 이메일 준비 시작")

            # 랜덤으로 제목과 본문 선택
            email_title_template = random.choice(email_titles)
            email_content_template = random.choice(email_contents)

            # 수신자 입력
            address_element = driver.find_element(By.CSS_SELECTOR, "#recipient_input_element")
            address_element.clear()
            pyperclip.copy(email)
            address_element.send_keys(Keys.CONTROL, 'v')
            address_element.send_keys(Keys.ENTER)
            time.sleep(1)

            # 제목 입력
            title_element = driver.find_element(By.CSS_SELECTOR, "#subject_title")
            title_element.clear()
            email_title = email_title_template.replace("{이름}", name)
            pyperclip.copy(email_title)
            title_element.send_keys(Keys.CONTROL, 'v')
            pyautogui.press("tab")
            time.sleep(0.5)

            # 내용 입력
            email_content = email_content_template.replace("{이름}", name)
            pyperclip.copy(email_content)
            pyautogui.hotkey("ctrl", "a")  # 기존 내용 전체 선택
            pyautogui.hotkey("ctrl", "v")  # 새 내용으로 덮어쓰기
            time.sleep(2)

            #전송 버튼 클릭 (주석 처리됨)
            send_button = driver.find_element(By.CSS_SELECTOR, ".button_write_task")
            send_button.click()
            time.sleep(5)


            logging.info(f"이메일이 {email} 전송완료")

               
            wait_time = random.uniform(3, 70)
            print(f"다음 메일 전송 대기 중 {wait_time:.2f} seconds")
            time.sleep(wait_time)

            

            # 새 메일 작성 페이지로 이동
            driver.get("https://mail.naver.com/v2/new")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#recipient_input_element")))
            time.sleep(2)

            # HTML 모드 선택 (기본값으로 설정)
            html_bt = driver.find_element(By.CSS_SELECTOR, "div.editor_mode_select button[value='HTML']")
            html_bt.click()
            time.sleep(1)

    except (TimeoutException, NoSuchElementException, WebDriverException) as e:
        logging.error(f"오류 발생: {str(e)}")
    finally:
        logging.info("스크립트 실행 완료. 브라우저를 열린 상태로 유지 시도...")
        keep_browser_open()
        input("Enter 키를 눌러 브라우저를 종료하세요...")

if __name__ == "__main__":
    try:
        email_titles, email_contents, user_id, user_pw, recipient_data = get_data_from_sheets()
        send_email(email_titles, email_contents, user_id, user_pw, recipient_data)
    except Exception as e:
        logging.error(f"메인 스크립트 실행 중 오류 발생: {str(e)}")
    finally:
        logging.info("메인 스크립트 실행 완료")