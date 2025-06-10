import requests
import json
from datetime import datetime
import sys
from pathlib import Path

# 현재 디렉토리를 파이썬 경로에 추가
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.append(str(current_dir))

from config import headers

def extract_page_id_from_url(url):
    """
    노션 URL에서 페이지 ID를 추출하는 함수
    
    Args:
        url (str): 노션 페이지 URL
        
    Returns:
        str: 페이지 ID
    """
    # URL에서 마지막 부분의 ID를 추출
    page_id = url.split('-')[-1]
    return page_id

def get_database_items(page_id):
    """
    노션 페이지의 데이터베이스 항목을 가져오는 함수
    
    Args:
        page_id (str): 노션 페이지 ID
        
    Returns:
        list: 데이터베이스 항목 목록
    """
    # 페이지의 블록들을 가져옴
    url = f"https://api.notion.com/v1/blocks/{page_id}/children"
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"블록 조회 중 오류 발생: {response.status_code}")
        print(response.text)
        return None
    
    blocks = response.json()["results"]
    
    # 데이터베이스 블록 찾기
    database_id = None
    for block in blocks:
        if block.get('type') == 'child_database':
            database_id = block.get('id')
            break
    
    if not database_id:
        print("데이터베이스를 찾을 수 없습니다.")
        return None
    
    # 데이터베이스 항목 가져오기
    database_url = f"https://api.notion.com/v1/databases/{database_id}/query"
    response = requests.post(database_url, headers=headers)
    
    if response.status_code != 200:
        print(f"데이터베이스 조회 중 오류 발생: {response.status_code}")
        print(response.text)
        return None
    
    return response.json()["results"]

def print_database_items(items):
    """
    데이터베이스 항목을 출력하는 함수
    
    Args:
        items (list): 데이터베이스 항목 목록
    """
    if not items:
        print("데이터베이스 항목이 없습니다.")
        return
    
    for idx, item in enumerate(items, 1):
        properties = item.get('properties', {})
        brand = ""
        item_name = ""
        
        for prop_name, prop_value in properties.items():
            try:
                if prop_value.get('type') == 'title':
                    brand = prop_value.get('title', [{}])[0].get('plain_text', '')
                elif prop_value.get('type') == 'rich_text':
                    rich_text = prop_value.get('rich_text', [])
                    text = rich_text[0].get('plain_text', '') if rich_text else ''
                    if prop_name == '2.아이템':
                        item_name = text
            except Exception as e:
                print(f"처리 중 오류 발생 - {str(e)}")
        
        print(f"{idx}. {brand} - {item_name}")

def main():
    # 노션 URL에서 페이지 ID 추출
    notion_url = input("노션 페이지 URL을 입력하세요: ")
    page_id = extract_page_id_from_url(notion_url)
    
    # 데이터베이스 항목 가져오기
    items = get_database_items(page_id)
    
    # 결과 출력
    print_database_items(items)

if __name__ == "__main__":
    main()
