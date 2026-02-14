"""
송장 수집 자동화 스크립트

각 업체의 구글 시트에서:
1. 송장번호가 입력된 데이터 수집
2. 유효성 검증
3. 이지어드민 업로드용 엑셀 파일 생성
"""
import os
import sys
import logging
from datetime import datetime

import pandas as pd

from utils import (
    Config, GoogleSheetClient, Logger,
    get_today_str, validate_tracking_number
)


def collect_invoices_from_sheet(sheet_client: GoogleSheetClient,
                                 vendor_info: dict) -> pd.DataFrame:
    """
    구글 시트에서 송장번호 수집

    Args:
        sheet_client: 구글 시트 클라이언트
        vendor_info: 업체 정보

    Returns:
        수집된 데이터 DataFrame
    """
    vendor_name = vendor_info['name']
    sheet_url = vendor_info.get('google_sheet_url')

    if not sheet_url:
        logging.warning(f'⚠️  {vendor_name}의 시트 URL이 없습니다.')
        return pd.DataFrame()

    try:
        # 시트 데이터 읽기
        data = sheet_client.read_sheet(sheet_url)

        if not data or len(data) < 2:
            logging.info(f'ℹ️  {vendor_name} 시트가 비어있습니다.')
            return pd.DataFrame()

        # DataFrame 생성
        df = pd.DataFrame(data[1:], columns=data[0])

        # 송장번호가 입력된 행만 필터링
        if '송장번호' not in df.columns:
            logging.warning(f'⚠️  {vendor_name} 시트에 "송장번호" 컬럼이 없습니다.')
            return pd.DataFrame()

        # 송장번호가 비어있지 않은 행
        df_with_tracking = df[df['송장번호'].notna() & (df['송장번호'] != '')]

        if len(df_with_tracking) == 0:
            logging.info(f'ℹ️  {vendor_name}: 입력된 송장번호가 없습니다.')
            return pd.DataFrame()

        logging.info(f'📦 {vendor_name}: {len(df_with_tracking)}건 수집')

        # 업체명 추가
        df_with_tracking['공급처'] = vendor_name

        return df_with_tracking

    except Exception as e:
        logging.error(f'❌ {vendor_name} 시트 읽기 실패: {e}')
        return pd.DataFrame()


def validate_and_clean_data(df: pd.DataFrame) -> tuple:
    """
    데이터 유효성 검증 및 정제

    Args:
        df: 수집된 데이터

    Returns:
        (정제된 DataFrame, 오류 목록)
    """
    if df.empty:
        return df, []

    errors = []
    valid_rows = []

    for idx, row in df.iterrows():
        tracking_number = str(row.get('송장번호', '')).strip()
        courier = str(row.get('택배사', '')).strip()
        order_number = str(row.get('주문번호', '')).strip()

        # 송장번호 유효성 검증
        if not validate_tracking_number(tracking_number):
            error_msg = f"주문번호 {order_number}: 유효하지 않은 송장번호 '{tracking_number}'"
            logging.warning(f'⚠️  {error_msg}')
            errors.append(error_msg)
            continue

        # 택배사 확인
        if not courier:
            error_msg = f"주문번호 {order_number}: 택배사가 입력되지 않았습니다"
            logging.warning(f'⚠️  {error_msg}')
            errors.append(error_msg)
            continue

        valid_rows.append(row)

    if valid_rows:
        clean_df = pd.DataFrame(valid_rows)
        logging.info(f'✅ 유효한 데이터: {len(clean_df)}건')
        if errors:
            logging.warning(f'⚠️  오류 발견: {len(errors)}건')
        return clean_df, errors
    else:
        return pd.DataFrame(), errors


def create_upload_excel(df: pd.DataFrame, output_file: str) -> bool:
    """
    이지어드민 업로드용 엑셀 파일 생성

    Args:
        df: 정제된 데이터
        output_file: 출력 파일 경로

    Returns:
        성공 여부
    """
    if df.empty:
        logging.error('❌ 데이터가 없어 엑셀 파일을 생성할 수 없습니다.')
        return False

    try:
        # 이지어드민 양식에 맞는 컬럼만 선택
        columns = ['주문번호', '택배사', '송장번호']

        # 필요한 컬럼만 선택
        upload_df = df[columns].copy()

        # 엑셀 파일 생성
        upload_df.to_excel(output_file, index=False, engine='openpyxl')

        logging.info(f'✅ 엑셀 파일 생성 완료: {output_file}')
        logging.info(f'📊 송장 등록 건수: {len(upload_df)}')

        return True

    except Exception as e:
        logging.error(f'❌ 엑셀 파일 생성 실패: {e}')
        return False


def save_error_log(errors: list, log_file: str):
    """
    오류 로그를 별도 파일로 저장

    Args:
        errors: 오류 목록
        log_file: 로그 파일 경로
    """
    if not errors:
        return

    try:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write('=' * 60 + '\n')
            f.write('송장 수집 오류 로그\n')
            f.write(f'생성 시간: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write('=' * 60 + '\n\n')

            for i, error in enumerate(errors, 1):
                f.write(f'{i}. {error}\n')

        logging.info(f'📝 오류 로그 저장: {log_file}')

    except Exception as e:
        logging.error(f'❌ 오류 로그 저장 실패: {e}')


def main():
    """메인 실행 함수"""
    # 로그 설정
    today = get_today_str()
    Logger.setup(log_file=f'logs/collect_invoices_{today}.log')

    logging.info('=' * 60)
    logging.info('📥 송장 수집 자동화 시작')
    logging.info('=' * 60)

    # 설정 로드
    config = Config()
    vendors_info = config.load_vendors()

    if not vendors_info:
        logging.error('❌ 업체 정보가 없습니다. 종료합니다.')
        sys.exit(1)

    # 구글 시트 클라이언트 초기화
    try:
        google_credentials = config.get_google_credentials_file()
        if not os.path.exists(google_credentials):
            logging.error(f'❌ 구글 인증 파일이 없습니다: {google_credentials}')
            logging.error('구글 API 인증 설정이 필요합니다.')
            sys.exit(1)

        sheet_client = GoogleSheetClient(google_credentials)
    except Exception as e:
        logging.error(f'❌ 구글 시트 클라이언트 초기화 실패: {e}')
        sys.exit(1)

    # 각 업체 시트에서 송장번호 수집
    logging.info('\n📦 업체별 송장번호 수집 중...\n')

    all_data = []
    all_errors = []

    for vendor_info in vendors_info:
        vendor_name = vendor_info['name']

        logging.info(f'처리 중: {vendor_name}')

        # 시트에서 데이터 수집
        df = collect_invoices_from_sheet(sheet_client, vendor_info)

        if not df.empty:
            all_data.append(df)

    # 전체 데이터 통합
    if not all_data:
        logging.warning('⚠️  수집된 송장번호가 없습니다.')
        logging.info('업체들이 아직 송장번호를 입력하지 않았거나, 시트에 접근할 수 없습니다.')
        sys.exit(0)

    combined_df = pd.concat(all_data, ignore_index=True)
    logging.info(f'\n📊 총 수집 건수: {len(combined_df)}건')

    # 데이터 유효성 검증
    logging.info('\n🔍 데이터 유효성 검증 중...\n')
    clean_df, errors = validate_and_clean_data(combined_df)

    if errors:
        all_errors.extend(errors)

    # 유효한 데이터가 없으면 종료
    if clean_df.empty:
        logging.error('❌ 유효한 데이터가 없습니다.')
        if all_errors:
            error_log_file = f'logs/invoice_errors_{today}.txt'
            save_error_log(all_errors, error_log_file)
        sys.exit(1)

    # 이지어드민 업로드용 엑셀 생성
    logging.info('\n📄 이지어드민 업로드용 엑셀 생성 중...\n')

    output_file = f'output/송장일괄등록_{today}.xlsx'

    # output 디렉토리 생성
    os.makedirs('output', exist_ok=True)

    success = create_upload_excel(clean_df, output_file)

    # 오류 로그 저장
    if all_errors:
        error_log_file = f'logs/invoice_errors_{today}.txt'
        save_error_log(all_errors, error_log_file)

    # 결과 요약
    logging.info('\n' + '=' * 60)
    logging.info('📊 처리 결과 요약')
    logging.info('=' * 60)
    logging.info(f'✅ 유효한 송장: {len(clean_df)}건')
    logging.info(f'⚠️  오류 발견: {len(all_errors)}건')
    logging.info(f'📁 출력 파일: {output_file}')

    if all_errors:
        logging.info(f'📝 오류 로그: logs/invoice_errors_{today}.txt')

    logging.info('=' * 60)

    if success:
        logging.info('✨ 송장 수집 완료!')
        logging.info(f'\n👉 다음 단계: 이지어드민에 {output_file} 파일을 업로드하세요.')
    else:
        logging.error('❌ 송장 수집 실패')
        sys.exit(1)


if __name__ == '__main__':
    main()
