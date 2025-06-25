from instagram_message import InstagramMessageTemplate

def test_message_template():
    # 테스트용 스프레드시트 ID와 시트 이름을 설정해주세요
    template_sheet_id = "1mwZ37jiEGK7rQnLWp87yUQZHyM6LHb4q6mbB0A07fI0"
    template_sheet_name = "공구템플릿_1"
    
    # InstagramMessageTemplate 인스턴스 생성
    message_template = InstagramMessageTemplate(template_sheet_id, template_sheet_name)
    
    # 메시지 템플릿 가져오기
    templates = message_template.get_message_templates()
    
    # 가져온 템플릿 출력
    print("생성된 메시지 템플릿:")
    for template in templates:
        print("\n" + "="*50)
        print(template)
        print("="*50)
    
    # 변수 대체 테스트
    test_name = "홍길동"
    test_brand = "테스트 브랜드"
    test_item = "테스트 제품"
    
    formatted_message = message_template.format_message(
        templates[0],
        name=test_name,
        brand=test_brand,
        item=test_item
    )
    
    print("\n변수 대체된 메시지:")
    print("\n" + "="*50)
    print(formatted_message)
    print("="*50)

if __name__ == "__main__":
    test_message_template() 