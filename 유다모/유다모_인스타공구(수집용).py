#고고야에서 구현할것 - 에러


from pymongo import MongoClient
from pymongo.server_api import ServerApi
from datetime import datetime, timedelta
import gspread
from dotenv import load_dotenv
from auth import get_credentials
import re

# .env 파일 로드
load_dotenv()

# MongoDB 연결 설정
uri = "mongodb+srv://coq3820:JmbIOcaEOrvkpQo1@cluster0.qj1ty.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
mongo_client = MongoClient(uri, server_api=ServerApi('1'), 
                    connectTimeoutMS=60000,  # 연결 타임아웃을 60초로 증가
                    socketTimeoutMS=60000)   # 소켓 타임아웃도 60초로 증가

def find_keywords(content):
    """content에서 키워드를 찾아 반환"""
    keywords = []
    if '출산' in content:
        keywords.append('출산')
    if '산후' in content:
        keywords.append('산후')
    return ', '.join(keywords) if keywords else ''

def set_column_widths_and_wrapping(worksheet):
    """열 너비 설정 및 자동 줄바꿈 설정"""
    # 모든 셀에 자동 줄바꿈 적용
    worksheet.format(
        'A1:I1000',
        {
            "wrapStrategy": "WRAP",
            "verticalAlignment": "TOP",
            "textFormat": {"fontSize": 10}
        }
    )
    
    # 헤더를 제외한 모든 행의 높이를 60픽셀로 설정
    spreadsheet_id = worksheet.spreadsheet.id
    worksheet_id = worksheet.id
    
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet_id,
                        "dimension": "ROWS",
                        "startIndex": 1,  # 2번째 행부터 시작 (0-based index)
                        "endIndex": 1000
                    },
                    "properties": {
                        "pixelSize": 60
                    },
                    "fields": "pixelSize"
                }
            }
        ]
    }
    
    worksheet.spreadsheet.batch_update(body)

try:
    # 연결 확인
    mongo_client.admin.command('ping')
    print("MongoDB 연결 성공!")
    
    # 데이터베이스와 컬렉션 선택
    db = mongo_client['insta09_database']
    post_collection = db['01_main_newfeed_crawl_data']
    user_collection = db['02_main_influencer_data']
    
    # 현재 날짜로부터 4개월 전 날짜 계산
    four_months_ago = (datetime.now() - timedelta(days=120)).isoformat()
    
    # '출산' 또는 '산후' 키워드가 포함되고, 최근 4개월 이내인 문서 검색
    query = {
        "$and": [
            {
                "content": {
                    "$regex": "출산|산후",
                    "$options": "i"  # 대소문자 구분 없이 검색
                }
            },
            {
                "cr_at": {
                    "$gte": four_months_ago  # 4개월 이내의 데이터만
                }
            }
        ]
    }
    
    # 필요한 필드만 선택하여 검색 (limit 제거)
    results = post_collection.find(
        query,
        {"author": 1, "post_url": 1, "cr_at": 1, "content": 1, "_id": 0}
    )
    
    # 구글 시트 연동
    creds = get_credentials()
    sheets_client = gspread.authorize(creds)
    
    # 스프레드시트 열기
    spreadsheet_id = "1VbtK0Q9iUG3VvbJJAlsb0NzzuA6RFFxmIFKmIuzYEQ0"
    spreadsheet = sheets_client.open_by_key(spreadsheet_id)
    
    # '인스타_현황판' 시트 가져오기 (없으면 생성)
    try:
        worksheet = spreadsheet.worksheet('인스타_현황판')
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title='인스타_현황판', rows=1000, cols=10)
        
    # 헤더 추가 (시트가 비어있는 경우)
    existing_data = worksheet.get_all_values()
    if not existing_data:
        headers = ["키워드", "작성자", "정제된 이름", "공구유무", "프로필 링크", "릴스 조회수", "게시물 URL", "작성 시간", "내용"]
        worksheet.append_row(headers)
        existing_authors = set()
        
        # 헤더 스타일 설정
        worksheet.format('A1:I1', {
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
            "horizontalAlignment": "CENTER",
            "textFormat": {"bold": True}
        })
    else:
        # 1달 이전 데이터 삭제
        one_month_ago = (datetime.now() - timedelta(days=30)).isoformat()
        rows_to_delete = []
        
        # 헤더를 제외한 모든 행을 검사
        for i, row in enumerate(existing_data[1:], start=2):  # start=2는 실제 시트의 행 번호
            cr_at = row[7]  # 작성 시간 열 (8번째 열)
            if cr_at < one_month_ago:
                rows_to_delete.append(i)
        
        # 오래된 행 한번에 삭제
        if rows_to_delete:
            # 삭제할 범위 생성 (A2:I1000 형식)
            range_to_clear = f'A{min(rows_to_delete)}:I{max(rows_to_delete)}'
            worksheet.batch_clear([range_to_clear])
            
            # 남은 데이터를 위로 이동 (각 행을 9개 열로 제한)
            remaining_data = [row[:9] for i, row in enumerate(existing_data[1:], start=2) if i not in rows_to_delete]
            if remaining_data:
                # update 메서드 수정: values를 먼저, range를 나중에
                worksheet.update(values=remaining_data, range_name=f'A2:I{len(remaining_data)+1}')
            
            print(f"{len(rows_to_delete)}개의 오래된 데이터가 삭제되었습니다.")
        
        # 남은 데이터에서 author 목록 추출
        remaining_data = worksheet.get_all_values()
        existing_authors = set(row[1] for row in remaining_data[1:])  # 헤더 제외
    
    # 열 너비와 줄바꿈 설정
    set_column_widths_and_wrapping(worksheet)
    
    # 결과 데이터 준비
    all_data = []
    latest_posts = {}  # author별 최신 게시물을 저장할 딕셔너리
    
    for doc in results:
        author = doc.get('author')
        
        # 이미 스프레드시트에 있는 author는 스킵
        if author in existing_authors:
            continue
            
        content = doc.get('content', '')
        cr_at = doc.get('cr_at', '')
        
        # 이미 해당 author의 게시물이 있는 경우, 날짜 비교
        if author in latest_posts:
            if cr_at <= latest_posts[author]['cr_at']:  # 현재 게시물이 더 오래된 경우 스킵
                continue
        
        # username이 author와 일치하는 사용자 정보 검색
        user_info = user_collection.find_one(
            {"username": author},
            {
                "clean_name": 1,
                "09_is": 1,
                "profile_link": 1,
                "reels_views(15)": 1,
                "_id": 0
            }
        )
        
        if user_info:
            post_data = {
                'cr_at': cr_at,
                'row_data': [
                    find_keywords(content),  # A열에 키워드 표시
                    author,
                    user_info.get('clean_name', ''),
                    user_info.get('09_is', ''),
                    user_info.get('profile_link', ''),
                    user_info.get('reels_views(15)', ''),
                    doc.get('post_url', ''),
                    cr_at,
                    content
                ]
            }
            latest_posts[author] = post_data
    
    # 최신 게시물만 all_data에 추가
    all_data = [post['row_data'] for post in latest_posts.values()]
            
    # 구글 시트에 데이터 추가
    if all_data:
        # 새로운 데이터 추가
        worksheet.append_rows(all_data)
        
        print(f"{len(all_data)}개의 새로운 데이터가 구글 시트에 추가되었습니다.")
    else:
        print("추가할 새로운 데이터가 없습니다.")

except Exception as e:
    print("오류 발생:", e)

finally:
    # MongoDB 연결 종료
    mongo_client.close()
    print("\nMongoDB 연결 종료")
