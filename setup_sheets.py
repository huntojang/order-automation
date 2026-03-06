"""
공급처별 구글 스프레드시트 자동 생성 스크립트 (OAuth2 방식)

처음 한 번만 실행하면:
1. 공급처별 구글 스프레드시트 자동 생성
2. vendors.json에 시트 URL 자동 저장
"""
import json
import logging
from utils import Config, GoogleSheetOAuthClient, Logger

# 관리자 구글 드라이브 공유 폴더 ID
SHARED_FOLDER_ID = '1Uh3c_c7kakXIXlTX8BT89HtBcA3nCY1v'


def create_vendor_sheets():
    """공급처별 구글 스프레드시트 생성"""
    Logger.setup(log_file='logs/setup_sheets.log')

    config = Config()

    # OAuth2 클라이언트 초기화 (처음 실행 시 브라우저 로그인)
    client = GoogleSheetOAuthClient(
        oauth_credentials_file=config.get_oauth_credentials_file(),
        token_file=config.get_oauth_token_file()
    )

    # vendors.json 로드
    with open('config/vendors.json', 'r', encoding='utf-8') as f:
        vendor_data = json.load(f)

    vendors = vendor_data['vendors']
    updated = False

    for vendor in vendors:
        vendor_name = vendor['name']

        # 이미 유효한 시트 URL이 있으면 건너뛰기
        existing_url = vendor.get('google_sheet_url', '')
        if existing_url and 'YOUR_SHEET_ID_HERE' not in existing_url:
            logging.info(f'⏭️  {vendor_name}: 이미 시트가 있어요 - 건너뜀')
            continue

        # 스프레드시트 생성 (공유 폴더 안에)
        sheet_title = f'[갓샵] {vendor_name}_발주서'
        spreadsheet = client.create_spreadsheet(sheet_title, folder_id=SHARED_FOLDER_ID)

        # 헤더 설정
        worksheet = spreadsheet.sheet1
        headers = ['주문일자', '주문번호', '수취인명', '연락처', '주소',
                   '상품명', '옵션', '수량', '택배사', '송장번호']
        worksheet.update('A1', [headers])

        # 헤더 서식
        worksheet.format('A1:J1', {
            'textFormat': {'bold': True},
            'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
        })

        # 서비스 계정에도 편집 권한 부여 (일상 자동화용)
        service_email = 'balju-bot@baljoo-helper-test.iam.gserviceaccount.com'
        spreadsheet.share(service_email, perm_type='user', role='writer')

        # vendors.json에 URL 저장
        vendor['google_sheet_url'] = spreadsheet.url
        logging.info(f'   📎 {spreadsheet.url}')

        updated = True

    # vendors.json 업데이트
    if updated:
        with open('config/vendors.json', 'w', encoding='utf-8') as f:
            json.dump(vendor_data, f, ensure_ascii=False, indent=2)
        logging.info('\n✅ vendors.json 업데이트 완료!')

    # 결과 요약
    logging.info('\n' + '=' * 60)
    logging.info('📊 공급처별 스프레드시트 목록')
    logging.info('=' * 60)
    for vendor in vendors:
        logging.info(f"  {vendor['name']}: {vendor['google_sheet_url']}")


if __name__ == '__main__':
    create_vendor_sheets()
