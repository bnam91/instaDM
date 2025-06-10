import random
import logging
from googleapiclient.discovery import build
from auth import get_credentials

class InstagramMessageTemplate:
    def __init__(self, template_sheet_id, template_sheet_name):
        self.template_sheet_id = template_sheet_id
        self.template_sheet_name = template_sheet_name

    def get_message_templates(self):
        """
        구글 스프레드시트에서 메시지 템플릿을 가져와 조합하는 함수
        
        Returns:
            list: 조합된 메시지 템플릿 리스트
        """
        logging.info("메시지 템플릿 가져오기 시작")
        try:
            creds = get_credentials()
            service = build('sheets', 'v4', credentials=creds)

            sheet = service.spreadsheets()
            # 각 파트별 템플릿 가져오기 (B1:D4 범위)
            result = sheet.values().get(spreadsheetId=self.template_sheet_id,
                                        range=f'{self.template_sheet_name}!B1:D4').execute()
            values = result.get('values', [])

            if not values or len(values) < 4:
                logging.warning('메시지 템플릿을 찾을 수 없습니다.')
                return ["안녕하세요"]

            # 각 파트별로 무작위 선택
            title = random.choice(values[0]) if values[0] else ""
            greeting = random.choice(values[1]) if values[1] else ""
            proposal = random.choice(values[2]) if values[2] else ""
            closing = random.choice(values[3]) if values[3] else ""

            # 선택된 파트들을 조합하여 하나의 메시지 생성
            message = f"{title}\n\n{greeting}\n\n{proposal}\n\n{closing}"
            
            return [message]  # 리스트 형태로 반환하여 기존 코드와의 호환성 유지

        except Exception as e:
            logging.error(f"메시지 템플릿을 가져오는 중 오류 발생: {e}")
            return ["안녕하세요"]

    def format_message(self, template, name="", notion_list=""):
        """
        템플릿의 변수를 실제 값으로 대체하는 함수
        
        Args:
            template (str): 메시지 템플릿
            name (str): 인플루언서 이름
            notion_list (str): 노션 리스트
            
        Returns:
            str: 변수가 대체된 메시지
        """
        return template.replace("{이름}", name).replace("{노션리스트}", notion_list) 