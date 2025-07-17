import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from auth import get_credentials
from googleapiclient.discovery import build
from datetime import datetime

# 구글 시트 ID
SPREADSHEET_ID = '1Yve6JJzgJaD4KXe2a8vB53zaWOUGsubcG_1t7XeNVmQ'

def get_cafe_urls_from_sheet():
    """구글 시트에서 카페 URL 목록을 가져오기"""
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='카페_리스트!C2:C' # C열에서 URL 가져오기 (헤더 제외)
    ).execute()
    values = result.get('values', [])
    if not values:
        print('스프레드시트에서 카페 URL을 찾을 수 없습니다.')
        return []
    else:
        return [row[0] for row in values if row and row[0]]

def update_cafe_data(row_index, member_count, recent_articles_count, hotdeal_boards):
    """구글 시트에 카페 데이터 업데이트"""
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    # 실제 시트의 행 번호 (헤더 제외)
    sheet_row = row_index + 2 # 헤더가 1행이므로 +2
    
    # 멤버수 업데이트 (D열)
    if member_count:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'카페_리스트!D{sheet_row}',
            valueInputOption='RAW',
            body={'values': [[member_count]]}
        ).execute()
    
    # 30분 게시글 수 업데이트 (E열)
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'카페_리스트!E{sheet_row}',
        valueInputOption='RAW',
        body={'values': [[recent_articles_count]]}
    ).execute()
    
    # 핫딜 게시판 업데이트 (K열)
    if hotdeal_boards:
        hotdeal_text = ', '.join(hotdeal_boards)
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'카페_리스트!K{sheet_row}',
            valueInputOption='RAW',
            body={'values': [[hotdeal_text]]}
        ).execute()
    
    print(f"행 {sheet_row} 데이터 업데이트 완료")

def scrape_cafe_data(url, cafe_index):
    """개별 카페 데이터 스크래핑"""
    print(f"\n{'='*60}")
    print(f"카페 {cafe_index + 1} 처리 중: {url}")
    print(f"{'='*80}")

    options = Options()
    options.add_experimental_option("detach", True)
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--log-level=3')
    options.add_argument('--silent')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=options)

    try:
        # 1. URL을 통해 카페를 연다
        driver.get(url)
        time.sleep(2)  # 페이지 로딩을 위한 대기 시간
        print("1. 카페 페이지 로드 완료")

        # 현재 페이지 URL 확인
        print(f"현재 페이지 URL: {driver.current_url}")

        # 페이지 소스를 txt 파일로 저장
        try:
            page_source = driver.page_source
            with open("cafe_page_source.txt", "w", encoding="utf-8") as f:
                f.write(page_source)
            print("페이지 소스가 'cafe_page_source.txt' 파일로 저장되었습니다.")
        except Exception as e:
            print(f"페이지 소스 저장 중 오류: {e}")

        # 메인 프레임으로 돌아가기 (iframe 전환 전에)
        driver.switch_to.default_content()

        # 카페 멤버수 정보 출력 - 메인 페이지에서 찾기
        print("카페 멤버수 추출 시도...")
        member_count = None
        try:
            # 방법 0: 실제 페이지 구조에 맞는 선택자 (최우선)
            member_count_element = driver.find_element(By.CSS_SELECTOR, "li.mem-cnt-info em")
            member_count = member_count_element.text.split('비공개')[0].strip()
            print(f"카페 멤버수: {member_count}")
        except Exception as e:
            print(f"방법 0 실패: {e}")
            try:
                # 방법 1: em 태그에서 숫자만 찾기
                em_elements = driver.find_elements(By.TAG_NAME, "em")
                for em in em_elements:
                    text = em.text.strip()
                    if ',' in text and '비공개' in text:
                        member_count = text.split('비공개')[0].strip()
                        print(f"카페 멤버수: {member_count}")
                        break
            except Exception as e:
                print(f"방법 1 실패: {e}")
                try:
                    # 방법 2: XPath로 mem-cnt-info 찾기
                    member_count_element = driver.find_element(By.XPATH, "//li[@class='mem-cnt-info']//em")
                    member_count = member_count_element.text.split('비공개')[0].strip()
                    print(f"카페 멤버수: {member_count}")
                except Exception as e:
                    print(f"방법 2 실패: {e}")
                    try:
                        # 방법 3: 텍스트로 직접 찾기
                        page_source = driver.page_source
                        import re
                        match = re.search(r'<em>([0-9,]+)<span class="ico_lock2">비공개</span></em>', page_source)
                        if match:
                            member_count = match.group(1)
                            print(f"카페 멤버수: {member_count}")
                        else:
                            print("멤버수 정보를 찾을 수 없습니다.")
                    except Exception as e:
                        print(f"방법 3 실패: {e}")
                        print("멤버수 정보를 찾을 수 없습니다.")

        # 카페 메인 페이지로 직접 이동 (게시판 목록이 있는 페이지)
        print("카페 메인 페이지로 이동 중...")
        driver.get(url)
        time.sleep(3)

        # iframe으로 전환
        print("iframe으로 전환 중...")
        try:
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "cafe_main"))
            )
            print("iframe 전환 성공")
            time.sleep(2)  # iframe 내부 로딩 대기 추가

            # iframe 내부 페이지 소스도 저장
            try:
                iframe_page_source = driver.page_source
                with open("cafe_iframe_source.txt", "w", encoding="utf-8") as f:
                    f.write(iframe_page_source)
                print("iframe 내부 페이지 소스가 'cafe_iframe_source.txt' 파일로 저장되었습니다.")
            except Exception as e:
                print(f"iframe 내부 페이지 소스 저장 중 오류: {e}")

        except Exception as e:
            print(f"iframe 전환 실패: {e}")
            driver.switch_to.default_content()

        # 2. 게시판 이름들 출력
        print("2. 게시판 이름들 출력")

        # 메인 프레임으로 돌아가기
        driver.switch_to.default_content()

        try:
            # 실제 카페 메뉴 구조에 맞는 선택자 사용
            board_links = []
            
            # 방법 1: cafe-menu-list 내의 모든 링크 찾기
            try:
                links1 = driver.find_elements(By.CSS_SELECTOR, ".cafe-menu-list a.gm-tcol-c")
                board_links.extend(links1)
                print(f"방법 1로 {len(links1)}개 찾음")
            except Exception as e:
                print(f"방법 1 실패: {e}")
            
            # 방법 2: 특별 메뉴 링크 찾기
            try:
                links2 = driver.find_elements(By.CSS_SELECTOR, ".special-menu a.link_special")
                board_links.extend(links2)
                print(f"방법 2로 {len(links2)}개 찾음")
            except Exception as e:
                print(f"방법 2 실패: {e}")
            
            # 방법 3: 모든 menuLink ID를 가진 링크 찾기
            try:
                links3 = driver.find_elements(By.CSS_SELECTOR, "a[id^='menuLink']")
                board_links.extend(links3)
                print(f"방법 3으로 {len(links3)}개 찾음")
            except Exception as e:
                print(f"방법 3 실패: {e}")
            
            # 중복 제거 및 한글 포함 게시판만 필터링
            unique_links = []
            seen_hrefs = set()
            for link in board_links:
                try:
                    href = link.get_attribute('href')
                    text = link.text.strip()
                    
                    # 한글이 포함되고 '%'가 없는 게시판만 필터링
                    if href and text and href not in seen_hrefs:
                        # 한글 문자가 포함되어 있는지 확인
                        has_korean = any('\u3131' <= char <= '\u318E' or '\uAC00' <= char <= '\uD7A3' for char in text)
                        # '%' 문자가 포함되어 있는지 확인
                        has_percent = '%' in text
                        
                        if has_korean and not has_percent:
                            unique_links.append((text, href, link))
                            seen_hrefs.add(href)
                            print(f"한글 게시판 추가: {text}")
                            
                except Exception as e:
                    print(f"링크 정보 추출 실패: {e}")
                    continue
            
            print(f"총 {len(unique_links)}개의 게시판을 찾았습니다:")
            
            def is_recent_post(time_text):
                """작성 시간이 30분 이내인지 확인"""
                import re
                from datetime import datetime, timedelta
                
                # 현재 시간
                now = datetime.now()
                
                # 시간 패턴 매칭
                if '분' in time_text:
                    minutes = int(re.search(r'(\d+)분', time_text).group(1))
                    return minutes <= 30
                elif '시간' in time_text:
                    hours = int(re.search(r'(\d+)시간', time_text).group(1))
                    return hours == 0  # 0시간 = 30분 이내
                elif '일' in time_text or '월' in time_text or '년' in time_text:
                    return False  # 1일 이상은 제외
                elif ':' in time_text:  # 시:분 형식 (예: 18:24)
                    try:
                        # 시:분 형식 파싱
                        time_parts = time_text.split(':')
                        if len(time_parts) == 2:
                            hour = int(time_parts[0])
                            minute = int(time_parts[1])
                            
                            # 오늘 날짜로 시간 객체 생성
                            post_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                            
                            # 현재 시간과의 차이 계산
                            time_diff = now - post_time
                            
                            # 30분 이내인지 확인
                            return time_diff.total_seconds() <= 30 * 60
                        else:
                            return False
                    except:
                        return False
                else:
                    return False  # 알 수 없는 형식은 제외

            def scrape_recent_articles(driver, board_name):
                """현재 게시판에서 최근 30분 이내 글들을 스크래핑"""
                try:
                    # 새로운 네이버 카페 구조는 iframe이 없으므로 전환하지 않음
                    time.sleep(3)  # 페이지 로딩 대기
                    
                    total_recent_count = 0
                    current_page = 1
                    max_page = 10  # 최대 10페이지까지 검색
                    
                    print(f"\n=== {board_name} 게시판 최근 글들 ===")
                    
                    while current_page <= max_page:
                        print(f"\n--- {current_page}페이지 크롤링 중 ---")
                        
                        # 게시글 목록 찾기 (새로운 네이버 카페 구조)
                        try:
                            # 방법 1: article-board 내의 tr 찾기
                            articles = driver.find_elements(By.CSS_SELECTOR, ".article-board tr")
                            print(f"방법 1로 {len(articles)}개의 행을 찾았습니다.")
                        except:
                            try:
                                # 방법 2: article 클래스를 가진 링크가 있는 tr 찾기
                                articles = driver.find_elements(By.CSS_SELECTOR, "tr:has(.article)")
                                print(f"방법 2로 {len(articles)}개의 행을 찾았습니다.")
                            except:
                                # 방법 3: 모든 tr에서 article 링크 찾기
                                all_trs = driver.find_elements(By.TAG_NAME, "tr")
                                articles = []
                                for tr in all_trs:
                                    try:
                                        tr.find_element(By.CSS_SELECTOR, ".article")
                                        articles.append(tr)
                                    except:
                                        continue
                                print(f"방법 3으로 {len(articles)}개의 행을 찾았습니다.")
                        
                        page_recent_count = 0
                        old_articles_found = 0  # 30분 이상 된 글 개수
                        
                        for article in articles:
                            try:
                                # 공지/필독 게시글 제외
                                try:
                                    notice_element = article.find_element(By.CSS_SELECTOR, "span.inner")
                                    notice_text = notice_element.text.strip()
                                    if notice_text in ["공지", "필독"]:
                                        continue
                                except:
                                    pass  # 공지/필독 태그가 없으면 일반 게시글
                                
                                # 작성 시간 확인 (새로운 구조)
                                try:
                                    time_element = article.find_element(By.CSS_SELECTOR, "td:nth-child(4)")
                                    time_text = time_element.text.strip()
                                except:
                                    try:
                                        # 다른 시간 선택자 시도
                                        time_elements = article.find_elements(By.TAG_NAME, "td")
                                        if len(time_elements) >= 4:
                                            time_text = time_elements[3].text.strip()
                                        else:
                                            time_text = "알 수 없음"
                                    except:
                                        time_text = "알 수 없음"
                                
                                # 30분 이내 글인지 확인
                                if is_recent_post(time_text):
                                    # 게시글 정보 추출 (새로운 구조)
                                    title_element = article.find_element(By.CSS_SELECTOR, ".article")
                                    title = title_element.text.strip()
                                    url = title_element.get_attribute("href")
                                    
                                    # 댓글 수 (새로운 구조)
                                    try:
                                        comment_element = article.find_element(By.CSS_SELECTOR, ".cmt")
                                        comment_text = comment_element.text.strip()
                                        # [숫자] 형태에서 숫자만 추출
                                        import re
                                        comment_match = re.search(r'\[(\d+)\]', comment_text)
                                        comment_count = comment_match.group(1) if comment_match else "0"
                                    except:
                                        comment_count = "0"
                                    
                                    # 닉네임 (새로운 구조)
                                    try:
                                        nickname_element = article.find_element(By.CSS_SELECTOR, ".nickname")
                                        nickname = nickname_element.text.strip()
                                    except:
                                        nickname = "알 수 없음"
                                    
                                    # 조회수 (새로운 구조)
                                    try:
                                        view_element = article.find_element(By.CSS_SELECTOR, ".type_readCount")
                                        view_count = view_element.text.strip()
                                    except:
                                        try:
                                            # 다른 조회수 선택자 시도
                                            view_elements = article.find_elements(By.TAG_NAME, "td")
                                            for td in view_elements:
                                                if "type_readCount" in td.get_attribute("class") or "":
                                                    view_count = td.text.strip()
                                                    break
                                            else:
                                                view_count = "알 수 없음"
                                        except:
                                            view_count = "알 수 없음"
                                    
                                    # 결과 출력
                                    print(f"\n제목: {title}")
                                    print(f"URL: {url}")
                                    print(f"댓글 수: {comment_count}")
                                    print(f"닉네임: {nickname}")
                                    print(f"작성시각: {time_text}")
                                    print(f"조회수: {view_count}")
                                    print("-" * 50)
                                    
                                    page_recent_count += 1
                                    total_recent_count += 1
                                else:
                                    old_articles_found += 1
                            except Exception as e:
                                print(f"게시글 정보 추출 중 오류: {e}")
                                continue
                        
                        print(f"{current_page}페이지에서 {page_recent_count}개의 최근 글을 찾았습니다.")
                        
                        # 30분 이상 된 글이 많이 나오면 더 이상 검색할 필요 없음
                        if old_articles_found > 5:  # 5개 이상의 오래된 글이 나오면 중단
                            print(f"30분 이상 된 글이 많이 나와서 {current_page}페이지에서 검색을 중단합니다.")
                            break
                        
                        # 다음 페이지로 이동
                        try:
                            # 새로운 네이버 카페 구조의 페이지네이션 (button 기반)
                            page_buttons = driver.find_elements(By.CSS_SELECTOR, "button.btn.number")
                            
                            # 다음 페이지 번호 찾기
                            next_page_number = current_page + 1
                            next_page_button = None
                            
                            # 현재 표시된 페이지 범위 확인 (1-10, 11-20 등)
                            current_range_start = ((current_page - 1) // 10) * 10 + 1
                            current_range_end = current_range_start + 9
                            # 다음 페이지가 현재 범위 내에 있는지 확인
                            if next_page_number <= current_range_end:
                                # 현재 범위 내에서 다음 페이지 버튼 찾기
                                for button in page_buttons:
                                    button_text = button.text.strip()
                                    if button_text == str(next_page_number):
                                        next_page_button = button
                                        break
                            else:
                                # 다음 범위로 넘어가야 함 - "다음" 버튼 클릭
                                try:
                                    next_button = driver.find_element(By.CSS_SELECTOR, "button.btn.type_next")
                                    next_button.click()
                                    time.sleep(2)
                                    current_page += 1
                                    print(f"{current_page}페이지로 이동했습니다 (다음 범위).")
                                    continue
                                except:
                                    print(f"다음 범위로 이동할 수 없어서 {current_page}페이지에서 종료합니다.")
                                    break
                            
                            if next_page_button:
                                next_page_button.click()
                                time.sleep(2)
                                current_page += 1
                                print(f"{current_page}페이지로 이동했습니다.")
                            else:
                                print(f"{next_page_number}페이지 버튼을 찾을 수 없어서 {current_page}페이지에서 종료합니다.")
                                break
                                
                        except Exception as e:
                            print(f"다음 페이지가 없어서 {current_page}페이지에서 종료합니다. 오류: {e}")
                            break
                    
                    if total_recent_count == 0:
                        print(f"{board_name} 게시판에 최근 30분 이내 글이 없습니다.")
                    else:
                        print(f"{board_name} 게시판에서 총 {total_recent_count}개의 최근 글을 찾았습니다.")
                    
                except Exception as e:
                    print(f"{board_name} 게시판 크롤링 중 오류: {e}")
                    # 디버깅을 위해 페이지 소스 일부 출력
                    try:
                        print("페이지 소스 일부:", driver.page_source[:1000])
                    except:
                        print("페이지 소스 확인 실패")
                
                return total_recent_count
            
            # 3. '전체글보기' 게시판만 찾아서 크롤링
            print("\n3. '전체글보기' 게시판 찾기 및 크롤링 시작")
            
            # '전체글보기' 게시판 찾기
            whole_board = None
            for board_name, href, link_element in unique_links:
                if '전체글보기' in board_name:
                    whole_board = (board_name, href, link_element)
                    break
            
            if whole_board:
                board_name, href, link_element = whole_board
                print(f"'전체글보기' 게시판을 찾았습니다: {board_name}")
                
                try:
                    print(f"게시판 '{board_name}' 클릭 시도 중...")
                    
                    # 게시판으로 이동
                    if href and href.startswith('http'):
                        driver.get(href)
                        print(f"게시판 '{board_name}' 링크로 이동 성공")
                    else:
                        # 상대 경로인 경우 전체 URL로 변환
                        full_url = f"https://cafe.naver.com{href}" if href.startswith('/') else href
                        driver.get(full_url)
                        print(f"게시판 '{board_name}' 전체 URL로 이동 성공")
                    
                    time.sleep(2)  # 페이지 로딩 대기
                    
                    # 페이지 소스를 txt 파일로 저장
                    try:
                        page_source = driver.page_source
                        with open("cafe_전체글보기_page_source.txt", "w", encoding="utf-8") as f:
                            f.write(page_source)
                        print("전체글보기 페이지 소스가 'cafe_전체글보기_page_source.txt' 파일로 저장되었습니다.")
                    except Exception as e:
                        print(f"전체글보기 페이지 소스 저장 중 오류: {e}")
                    
                    # 최근 글 크롤링
                    total_recent_articles = scrape_recent_articles(driver, board_name)
                    
                    # 공구, 핫딜이 포함된 게시판 확인 (공백 제거 후 비교)
                    special_keywords = ['공구', '핫딜']
                    special_boards = []
                    for board_name, href, link_element in unique_links:
                        board_name_no_space = board_name.replace(' ', '')
                        for keyword in special_keywords:
                            if keyword in board_name_no_space:
                                special_boards.append(board_name)
                                break
                    
                    # 최종 결과 출력
                    print("\n" + "="*60)
                    print(f"🎉 전체글보기 최종 결과: 최근 30분 내 총 {total_recent_articles}개의 게시글 발견!")
                    
                    if special_boards:
                        print("📢 특별 게시판 발견:")
                        for board in special_boards:
                            print(f"   • {board}")
                        print("💡 위 게시판들에서도 최근 글을 확인해보세요!")
                    
                    print("="*60)
                    
                    return member_count, total_recent_articles, special_boards
                except Exception as e:
                    print(f"게시판 '{board_name}' 처리 중 오류: {e}")
                    try:
                        driver.switch_to.default_content()
                    except:
                        pass
                    return member_count, 0, []
            else:
                print("'전체글보기' 게시판을 찾을 수 없습니다.")
                print("사용 가능한 게시판 목록:")
                for i, (board_name, href, link_element) in enumerate(unique_links, 1):
                    print(f"{i}. {board_name}")
                return member_count, 0, []
        except Exception as e:
            print(f"게시판 찾기 중 오류 발생: {e}")
            print("현재 페이지 HTML 구조 확인 중...")
            try:
                page_source = driver.page_source
                print("페이지 소스 일부:", page_source[:1000])
            except:
                print("페이지 소스 확인 실패")
            return member_count, 0, []
    finally:
        driver.quit()

def get_processed_indices():
    """E열(30분 게시글 수)에 값이 있는 행의 인덱스(0부터 시작)를 반환"""
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='카페_리스트!E2:E'  # E열(헤더 제외)
    ).execute()
    values = result.get('values', [])
    processed_indices = []
    for i, row in enumerate(values):
        if row and row[0] and row[0].strip() != '':
            processed_indices.append(i)
    return processed_indices

def main():
    """메인 함수"""
    print("네이버 카페 크롤링 시작!")
    print("구글 시트에서 카페 URL 목록을 가져오는 중...")
    
    # 구글 시트에서 카페 URL 목록 가져오기
    cafe_urls = get_cafe_urls_from_sheet()
    processed_indices = get_processed_indices()
    
    if not cafe_urls:
        print("카페 URL을 찾을 수 없습니다. 프로그램을 종료합니다.")
        return
    
    print(f"총 {len(cafe_urls)}개의 카페 URL을 찾았습니다.")
    print(f"이미 처리된 인덱스: {processed_indices}")
    
    # 각 카페별로 처리 (이미 처리된 인덱스는 건너뜀)
    for i, url in enumerate(cafe_urls):
        if i in processed_indices:
            print(f"[SKIP] {i+1}번째 카페는 이미 처리됨. 건너뜁니다.")
            continue
        try:
            print(f"\n{'='*80}")
            print(f"카페 {i+1}/{len(cafe_urls)} 처리 중...")
            print(f"URL: {url}")
            print(f"{'='*80}")
            # 카페 데이터 스크래핑
            member_count, recent_articles_count, hotdeal_boards = scrape_cafe_data(url, i)
            
            # 구글 시트에 결과 업데이트
            update_cafe_data(i, member_count, recent_articles_count, hotdeal_boards)
            
            print(f"카페 {i+1} 처리 완료!")
            
            # 다음 카페 처리 전 잠시 대기
            if i < len(cafe_urls) - 1:
                print("다음 카페 처리 전 3초 대기...")
                time.sleep(3)
                
        except Exception as e:
            print(f"카페 {i+1} 처리 중 오류 발생: {e}")
            continue
    
    print(f"\n{'='*80}")
    print("모든 카페 처리 완료!")
    print(f"{'='*80}")

if __name__ == "__main__":
    main()
