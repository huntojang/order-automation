"""
발주 자동화 스크립트

이지어드민 엑셀 파일을 읽어서:
1. 업체별로 데이터 분류
2. 각 업체의 구글 시트에 업로드
3. 업체 사장님에게 알림톡 발송
"""
import os
import sys
import logging
from datetime import datetime

import pandas as pd
import requests

from utils import (
    Config, GoogleSheetClient, GoogleSheetOAuthClient, Logger,
    get_today_str, get_latest_file
)


def load_order_excel(file_path: str) -> pd.DataFrame:
    """
    이지어드민 발주서 엑셀 파일 로드

    Args:
        file_path: 엑셀 파일 경로

    Returns:
        DataFrame
    """
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        logging.info(f'✅ 엑셀 파일 로드 완료: {file_path}')
        logging.info(f'📊 총 주문 건수: {len(df)}')
        return df
    except Exception as e:
        logging.error(f'❌ 엑셀 파일 로드 실패: {e}')
        return None


def load_and_merge_files(directory: str) -> pd.DataFrame:
    """
    디렉토리의 모든 엑셀/CSV 파일을 읽어서 하나로 합침

    Args:
        directory: 파일이 있는 디렉토리

    Returns:
        합쳐진 DataFrame
    """
    import glob

    # 엑셀 + CSV 파일 찾기
    excel_files = glob.glob(os.path.join(directory, '*.xlsx'))
    csv_files = glob.glob(os.path.join(directory, '*.csv'))
    all_files = excel_files + csv_files

    if not all_files:
        logging.error(f'❌ {directory} 디렉토리에 데이터 파일이 없습니다.')
        return None

    logging.info(f'📂 발견된 파일: {len(all_files)}개')

    all_dataframes = []

    for file_path in sorted(all_files):
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8')
            else:
                df = pd.read_excel(file_path, engine='openpyxl')
            file_name = os.path.basename(file_path)
            logging.info(f'  ✅ {file_name}: {len(df)}건')
            all_dataframes.append(df)
        except Exception as e:
            logging.error(f'  ❌ {os.path.basename(file_path)} 로드 실패: {e}')
            continue

    if not all_dataframes:
        logging.error('❌ 로드된 파일이 없습니다.')
        return None

    # 모든 DataFrame 합치기
    merged_df = pd.concat(all_dataframes, ignore_index=True)

    logging.info(f'📊 총 주문 건수 (합계): {len(merged_df)}')

    return merged_df


def split_by_vendor(df: pd.DataFrame, vendor_column='공급처') -> dict:
    """
    업체별로 데이터 분류

    Args:
        df: 전체 주문 데이터
        vendor_column: 업체 컬럼명

    Returns:
        {업체명: DataFrame} 딕셔너리
    """
    if vendor_column not in df.columns:
        logging.error(f'❌ "{vendor_column}" 컬럼을 찾을 수 없습니다.')
        logging.info(f'사용 가능한 컬럼: {", ".join(df.columns)}')
        return {}

    vendor_data = {}
    for vendor_name, group in df.groupby(vendor_column):
        vendor_data[vendor_name] = group.reset_index(drop=True)
        logging.info(f'📦 {vendor_name}: {len(group)}건')

    return vendor_data


def prepare_sheet_data(df: pd.DataFrame) -> list:
    """
    구글 시트에 업로드할 데이터 준비

    Args:
        df: 업체별 주문 데이터

    Returns:
        2차원 리스트 (헤더 포함)
    """
    column_mapping = {
        '관리번호': ['관리번호'],
        '주문번호': ['주문번호'],
        '수취인명': ['수취인명', '수령자 이름', '수령자이름'],
        '연락처': ['연락처', '수령자 휴대폰번호', '수령자 전화', '수령자휴대폰', '수령자전화'],
        '주소': ['주소', '수령자 주소', '수령자주소'],
        '상품명': ['상품명'],
        '옵션': ['옵션', '옵션명'],
        '수량': ['수량', '상품수량'],
        '배송메모': ['배송메모'],
        '택배사': ['택배사'],
        '송장번호': ['송장번호'],
    }

    sheet_headers = list(column_mapping.keys())

    def find_value(row, candidates):
        for col in candidates:
            if col in df.columns:
                val = row[col]
                if not pd.isna(val) and str(val).strip() != '':
                    return str(val)
        return ''

    rows = []
    for _, row in df.iterrows():
        new_row = [find_value(row, candidates) for candidates in column_mapping.values()]
        rows.append(new_row)

    return [sheet_headers] + rows


def send_alimtalk(vendor_info: dict, order_count: int, config: dict) -> bool:
    """
    카카오 알림톡 발송

    Args:
        vendor_info: 업체 정보
        order_count: 주문 건수
        config: 알림톡 설정

    Returns:
        성공 여부
    """
    if not config:
        logging.warning('⚠️  알림톡 설정이 없습니다. 발송을 건너뜁니다.')
        return False

    try:
        # 메시지 템플릿 포맷팅
        message = config['message_template'].format(
            vendor_name=vendor_info['name'],
            date=datetime.now().strftime('%Y년 %m월 %d일'),
            count=order_count,
            sheet_url=vendor_info['google_sheet_url']
        )

        # API 엔드포인트 (알리고)
        if config['service'] == 'aligo':
            import json as _json
            api_url = config.get('proxy_url', 'https://kakaoapi.aligo.in/akv10/alimtalk/send/')
            button_info = _json.dumps({
                "button": [{
                    "type": "WL",
                    "name": "발주서 확인하기",
                    "url_mobile": vendor_info['google_sheet_url'],
                    "url_pc": vendor_info['google_sheet_url']
                }]
            })
            payload = {
                'apikey': config['api_key'],
                'userid': config['user_id'],
                'senderkey': config['sender_key'],
                'tpl_code': config['template_code'],
                'sender': config['sender'],
                'receiver_1': vendor_info['phone'],
                'subject_1': '발주 알림',
                'message_1': message,
                'button_1': button_info,
            }
        else:
            logging.error(f'❌ 지원하지 않는 서비스: {config["service"]}')
            return False

        # API 호출
        response = requests.post(api_url, data=payload, timeout=10)

        if response.status_code == 200:
            result = response.json()
            if result.get('code') == 0:
                logging.info(f'✅ 알림톡 발송 성공: {vendor_info["name"]} ({vendor_info["phone"]})')
                return True
            else:
                logging.error(f'❌ 알림톡 발송 실패: {result.get("message")}')
                return False
        else:
            logging.error(f'❌ API 호출 실패: {response.status_code}')
            return False

    except Exception as e:
        logging.error(f'❌ 알림톡 발송 중 오류: {e}')
        return False


def main():
    """메인 실행 함수"""
    # 로그 설정
    today = get_today_str()
    Logger.setup(log_file=f'logs/send_orders_{today}.log')

    logging.info('=' * 60)
    logging.info('📤 발주 자동화 시작')
    logging.info('=' * 60)

    # 설정 로드
    config = Config()
    vendors_info = config.load_vendors()
    alimtalk_config = config.load_alimtalk_config()

    if not vendors_info:
        logging.error('❌ 업체 정보가 없습니다. 종료합니다.')
        sys.exit(1)

    # 구글 시트 클라이언트 초기화 (서비스 계정 우선, 없으면 OAuth2)
    sheet_client = None
    try:
        google_credentials = config.get_google_credentials_file()
        if os.path.exists(google_credentials):
            sheet_client = GoogleSheetClient(google_credentials)
    except Exception:
        pass

    if not sheet_client:
        try:
            sheet_client = GoogleSheetOAuthClient(
                oauth_credentials_file=config.get_oauth_credentials_file(),
                token_file=config.get_oauth_token_file()
            )
        except Exception as e:
            logging.error(f'❌ 구글 시트 클라이언트 초기화 실패: {e}')
            logging.info('테스트 모드로 계속 진행합니다 (구글 시트 업로드 제외)')

    # 데이터 파일 로드 (엑셀 + CSV, 여러 파일 자동 병합)
    input_dir = 'input'
    df = load_and_merge_files(input_dir)

    if df is None:
        sys.exit(1)

    # 업체별 데이터 분류
    logging.info('\n📦 업체별 주문 분류 중...')
    vendor_data = split_by_vendor(df)

    if not vendor_data:
        logging.error('❌ 업체별 데이터 분류 실패')
        sys.exit(1)

    # 각 업체별 처리
    success_count = 0
    fail_count = 0

    for vendor_info in vendors_info:
        vendor_name = vendor_info['name']

        logging.info(f'\n{"=" * 60}')
        logging.info(f'처리 중: {vendor_name}')
        logging.info(f'{"=" * 60}')

        # 해당 업체의 주문이 있는지 확인
        if vendor_name not in vendor_data:
            logging.info(f'ℹ️  {vendor_name}에 대한 주문이 없습니다.')
            continue

        vendor_df = vendor_data[vendor_name]
        order_count = len(vendor_df)

        # 구글 시트 데이터 준비
        sheet_data = prepare_sheet_data(vendor_df)

        # 구글 시트 업데이트
        if sheet_client:
            sheet_url = vendor_info.get('google_sheet_url')
            if sheet_url:
                success = sheet_client.update_sheet(sheet_url, sheet_data)
                if not success:
                    logging.error(f'❌ {vendor_name} 시트 업데이트 실패')
                    fail_count += 1
                    continue
            else:
                logging.warning(f'⚠️  {vendor_name}의 구글 시트 URL이 없습니다.')
        else:
            logging.info(f'ℹ️  테스트 모드: 시트 업데이트 건너뜀')

        # 알림톡 발송
        if alimtalk_config:
            send_alimtalk(vendor_info, order_count, alimtalk_config)
        else:
            logging.info(f'ℹ️  알림톡 설정이 없어 발송을 건너뜁니다.')

        success_count += 1

    # 결과 요약
    logging.info('\n' + '=' * 60)
    logging.info('📊 처리 결과 요약')
    logging.info('=' * 60)
    logging.info(f'✅ 성공: {success_count}개 업체')
    logging.info(f'❌ 실패: {fail_count}개 업체')
    logging.info(f'📦 총 주문 건수: {len(df)}건')
    logging.info('=' * 60)

    logging.info('✨ 발주 자동화 완료!')


if __name__ == '__main__':
    main()
