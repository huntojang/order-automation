"""
공통 유틸리티 함수 모듈
"""
import json
import logging
import os
import time
from datetime import datetime
from typing import Dict, List, Any

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


class Config:
    """설정 관리 클래스 (로컬 파일 또는 Streamlit Secrets 지원)"""

    def __init__(self, config_dir='config'):
        self.config_dir = config_dir
        self._secrets = None
        try:
            import streamlit as st
            if hasattr(st, 'secrets') and len(st.secrets) > 0:
                self._secrets = st.secrets
        except Exception:
            pass

    def load_vendors(self) -> List[Dict[str, Any]]:
        """업체 정보 로드 (Secrets 우선, 없으면 파일)"""
        if self._secrets and "vendors" in self._secrets:
            vendors_section = self._secrets["vendors"]
            vendors_str = vendors_section.get("vendors", "[]") if hasattr(vendors_section, 'get') else str(vendors_section)
            if isinstance(vendors_str, str):
                parsed = json.loads(vendors_str)
                return [dict(v) for v in parsed]

        vendors_file = os.path.join(self.config_dir, 'vendors.json')
        try:
            with open(vendors_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('vendors', [])
        except (FileNotFoundError, json.JSONDecodeError):
            pass

        logging.error(f'업체 정보를 찾을 수 없습니다')
        return []

    def load_alimtalk_config(self) -> Dict[str, Any]:
        """알림톡 설정 로드"""
        if self._secrets and "alimtalk" in self._secrets:
            alimtalk_raw = self._secrets["alimtalk"]
            if isinstance(alimtalk_raw, str):
                return json.loads(alimtalk_raw)
            return dict(alimtalk_raw)

        config_file = os.path.join(self.config_dir, 'alimtalk_config.json')
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            logging.error(f'알림톡 설정 파일을 찾을 수 없습니다: {config_file}')
            return {}
        except json.JSONDecodeError as e:
            logging.error(f'알림톡 설정 파일 형식 오류: {e}')
            return {}

    def get_google_credentials_file(self) -> str:
        """구글 API 인증 파일 경로 (서비스 계정)"""
        return os.path.join(self.config_dir, 'google_credentials.json')

    def get_google_credentials_dict(self) -> dict:
        """구글 API 인증 정보 (Streamlit Secrets용)"""
        if self._secrets and "google_credentials" in self._secrets:
            cred_raw = self._secrets["google_credentials"]
            if isinstance(cred_raw, str):
                return json.loads(cred_raw)
            return dict(cred_raw)
        return None

    def get_vendor_master_url(self) -> str:
        """업체 마스터 시트 URL"""
        if self._secrets and "vendor_master" in self._secrets:
            return self._secrets["vendor_master"].get("sheet_url", "")
        return ""

    def get_shared_folder_id(self) -> str:
        """구글 드라이브 공유폴더 ID"""
        if self._secrets and "vendor_master" in self._secrets:
            return self._secrets["vendor_master"].get("shared_folder_id", "")
        return "1Uh3c_c7kakXIXlTX8BT89HtBcA3nCY1v"

    def get_oauth_credentials_file(self) -> str:
        """OAuth2 인증 파일 경로"""
        return os.path.join(self.config_dir, 'oauth_credentials.json')

    def get_oauth_token_file(self) -> str:
        """OAuth2 토큰 저장 파일 경로"""
        return os.path.join(self.config_dir, 'token.json')


class GoogleSheetClient:
    """구글 시트 클라이언트"""

    def __init__(self, credentials_file: str = None, credentials_dict: dict = None):
        """
        Args:
            credentials_file: 구글 API 인증 JSON 파일 경로
            credentials_dict: 구글 API 인증 정보 딕셔너리 (Streamlit Secrets용)
        """
        self.credentials_file = credentials_file
        self.credentials_dict = credentials_dict
        self.client = None
        self._authenticate()

    def _authenticate(self):
        """구글 API 인증"""
        try:
            scope = [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive'
            ]
            if self.credentials_dict:
                credentials = ServiceAccountCredentials.from_json_keyfile_dict(
                    self.credentials_dict, scope
                )
            else:
                credentials = ServiceAccountCredentials.from_json_keyfile_name(
                    self.credentials_file, scope
                )
            self.client = gspread.authorize(credentials)
            logging.info('✅ 구글 시트 인증 성공')
        except FileNotFoundError:
            logging.error(f'❌ 구글 인증 파일을 찾을 수 없습니다: {self.credentials_file}')
            raise
        except Exception as e:
            logging.error(f'❌ 구글 시트 인증 실패: {e}')
            raise

    _spreadsheet_cache = {}  # URL → spreadsheet 객체 캐시
    _worksheet_cache = {}    # (URL, index) → worksheet 객체 캐시

    def _retry_on_quota(self, fn, max_retries=5):
        """429 Rate Limit 에러 시 자동 재시도 (exponential backoff)"""
        for attempt in range(max_retries + 1):
            try:
                return fn()
            except gspread.exceptions.APIError as e:
                if e.response.status_code == 429 and attempt < max_retries:
                    wait = (attempt + 1) * 15
                    logging.warning(f'⏳ API 할당량 초과, {wait}초 후 재시도 ({attempt + 1}/{max_retries})')
                    time.sleep(wait)
                else:
                    raise

    def open_sheet_by_url(self, url: str, retry=True):
        """URL로 시트 열기 (항상 캐시 — spreadsheet 객체는 핸들이라 캐시해도 안전)"""
        try:
            if url in self._spreadsheet_cache:
                return self._spreadsheet_cache[url]
            if retry:
                ss = self._retry_on_quota(lambda: self.client.open_by_url(url))
            else:
                ss = self.client.open_by_url(url)
            self._spreadsheet_cache[url] = ss
            return ss
        except Exception as e:
            logging.error(f'❌ 시트 열기 실패 ({url}): {e}')
            return None

    def _get_worksheet_cached(self, sheet_url: str, spreadsheet, worksheet_index: int = 0):
        """워크시트 객체 캐시 (get_worksheet 호출 = 읽기 API 1회 절약)"""
        cache_key = (sheet_url, worksheet_index)
        if cache_key in self._worksheet_cache:
            return self._worksheet_cache[cache_key]
        worksheet = self._retry_on_quota(lambda: spreadsheet.get_worksheet(worksheet_index))
        self._worksheet_cache[cache_key] = worksheet
        return worksheet

    def update_sheet(self, sheet_url: str, data: List[List[Any]],
                     worksheet_index: int = 0):
        """
        시트 데이터 업데이트 (API 1회: 빈 행 패딩으로 덮어쓰기)
        서식은 Apps Script에서 시트 생성 시 설정되므로 여기서는 값만 업데이트.

        Returns:
            성공 시 True, 실패 시 에러 메시지 문자열
        """
        try:
            spreadsheet = self.open_sheet_by_url(sheet_url)
            if not spreadsheet:
                return "시트를 열 수 없습니다 (권한 없음 또는 URL 오류)"

            worksheet = self._get_worksheet_cached(sheet_url, spreadsheet, worksheet_index)

            # 빈 행 패딩으로 이전 데이터 덮어쓰기 (clear 불필요)
            if data:
                num_cols = len(data[0])
                MAX_ROWS = max(len(data) + 10, 100)
                padded = data + [[''] * num_cols] * (MAX_ROWS - len(data))
                self._retry_on_quota(lambda: worksheet.update('A1', padded))

            logging.info(f'✅ 시트 업데이트 완료: {spreadsheet.title}')
            return True

        except Exception as e:
            logging.error(f'❌ 시트 업데이트 실패: {e}')
            return str(e)

    def create_spreadsheet(self, title: str, folder_id: str = None):
        """새 스프레드시트 생성"""
        spreadsheet = self.client.create(title, folder_id=folder_id)
        logging.info(f'✅ 스프레드시트 생성: {title}')
        return spreadsheet

    def append_row(self, sheet_url: str, row: List[Any], worksheet_index: int = 0) -> bool:
        """시트에 행 추가"""
        try:
            spreadsheet = self.open_sheet_by_url(sheet_url)
            if not spreadsheet:
                return False
            worksheet = spreadsheet.get_worksheet(worksheet_index)
            worksheet.append_row(row, value_input_option='USER_ENTERED')
            return True
        except Exception as e:
            logging.error(f'❌ 행 추가 실패: {e}')
            return False

    def read_sheet(self, sheet_url: str, worksheet_index: int = 0) -> List[List[Any]]:
        """
        시트 데이터 읽기

        Args:
            sheet_url: 구글 시트 URL
            worksheet_index: 워크시트 인덱스

        Returns:
            시트 데이터 (2차원 리스트)
        """
        try:
            spreadsheet = self.open_sheet_by_url(sheet_url, retry=False)
            if not spreadsheet:
                return []

            worksheet = self._get_worksheet_cached(sheet_url, spreadsheet, worksheet_index)
            return worksheet.get_all_values()

        except Exception as e:
            logging.error(f'❌ 시트 읽기 실패: {e}')
            return []


class GoogleSheetOAuthClient:
    """OAuth2 기반 구글 시트 클라이언트 (시트 생성 가능)"""

    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    def __init__(self, oauth_credentials_file: str, token_file: str):
        self.oauth_credentials_file = oauth_credentials_file
        self.token_file = token_file
        self.client = None
        self._authenticate()

    def _authenticate(self):
        """OAuth2 인증 (처음 1번만 브라우저 로그인 필요)"""
        creds = None

        # 저장된 토큰이 있으면 로드
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)

        # 토큰이 없거나 만료됐으면 새로 발급
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.oauth_credentials_file, self.SCOPES
                )
                creds = flow.run_local_server(port=0)

            # 토큰 저장
            with open(self.token_file, 'w') as f:
                f.write(creds.to_json())

        self.client = gspread.authorize(creds)
        logging.info('✅ OAuth2 구글 시트 인증 성공')

    def create_spreadsheet(self, title: str, folder_id: str = None):
        """새 스프레드시트 생성"""
        spreadsheet = self.client.create(title, folder_id=folder_id)
        logging.info(f'✅ 스프레드시트 생성: {title}')
        return spreadsheet

    def open_sheet_by_url(self, url: str):
        """URL로 시트 열기"""
        try:
            return self.client.open_by_url(url)
        except Exception as e:
            logging.error(f'❌ 시트 열기 실패 ({url}): {e}')
            return None

    def update_sheet(self, sheet_url: str, data: List[List[Any]],
                     worksheet_index: int = 0) -> bool:
        """시트 데이터 업데이트"""
        try:
            spreadsheet = self.open_sheet_by_url(sheet_url)
            if not spreadsheet:
                return False

            worksheet = spreadsheet.get_worksheet(worksheet_index)
            worksheet.clear()

            if data:
                worksheet.update('A1', data)

            if len(data) > 0:
                worksheet.format('A1:J1', {
                    'textFormat': {'bold': True},
                    'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
                })

            if len(data) > 1:
                last_row = len(data)
                worksheet.format(f'J2:J{last_row}', {
                    'backgroundColor': {'red': 1.0, 'green': 1.0, 'blue': 0.8}
                })

            logging.info(f'✅ 시트 업데이트 완료: {spreadsheet.title}')
            return True

        except Exception as e:
            logging.error(f'❌ 시트 업데이트 실패: {e}')
            return False

    def read_sheet(self, sheet_url: str, worksheet_index: int = 0) -> List[List[Any]]:
        """시트 데이터 읽기"""
        try:
            spreadsheet = self.open_sheet_by_url(sheet_url)
            if not spreadsheet:
                return []
            worksheet = spreadsheet.get_worksheet(worksheet_index)
            return worksheet.get_all_values()
        except Exception as e:
            logging.error(f'❌ 시트 읽기 실패: {e}')
            return []


class Logger:
    """로그 설정"""

    @staticmethod
    def setup(log_file: str = None, level=logging.INFO):
        """
        로그 설정

        Args:
            log_file: 로그 파일 경로 (None이면 콘솔만)
            level: 로그 레벨
        """
        # 로그 포맷
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        date_format = '%Y-%m-%d %H:%M:%S'

        # 핸들러 설정
        handlers = [logging.StreamHandler()]  # 콘솔 출력

        if log_file:
            # 로그 디렉토리 생성
            log_dir = os.path.dirname(log_file)
            if log_dir and not os.path.exists(log_dir):
                os.makedirs(log_dir)

            # 파일 출력 추가
            handlers.append(logging.FileHandler(log_file, encoding='utf-8'))

        # 로깅 설정
        logging.basicConfig(
            level=level,
            format=log_format,
            datefmt=date_format,
            handlers=handlers
        )


def validate_tracking_number(tracking_number: str, courier: str = None) -> bool:
    """
    송장번호 유효성 검증

    Args:
        tracking_number: 송장번호
        courier: 택배사 (선택)

    Returns:
        유효 여부
    """
    if not tracking_number:
        return False

    # 공백 제거
    tracking_number = str(tracking_number).strip()

    # 숫자만 있는지 확인
    if not tracking_number.isdigit():
        return False

    # 길이 체크 (일반적으로 10-15자리)
    if len(tracking_number) < 10 or len(tracking_number) > 15:
        return False

    # 택배사별 추가 검증 (선택)
    # TODO: 필요시 택배사별 규칙 추가

    return True


def get_today_str(format: str = '%Y%m%d') -> str:
    """오늘 날짜 문자열 반환"""
    return datetime.now().strftime(format)


def get_latest_file(directory: str, pattern: str = '*.xlsx') -> str:
    """
    디렉토리에서 가장 최근 파일 찾기

    Args:
        directory: 검색할 디렉토리
        pattern: 파일 패턴

    Returns:
        파일 경로 (없으면 None)
    """
    import glob

    files = glob.glob(os.path.join(directory, pattern))
    if not files:
        return None

    # 수정 시간 기준 정렬
    latest_file = max(files, key=os.path.getmtime)
    return latest_file


class VendorManager:
    """업체 마스터 시트 기반 CRUD 관리"""

    HEADERS = ['업체ID', '업체명', '담당자', '전화번호', '이메일', '구글시트URL', '등록일', '상태']
    SHEET_HEADERS = ['주문일자', '주문번호', '수취인명', '연락처', '주소',
                     '상품명', '옵션', '수량', '택배사', '송장번호']

    def __init__(self, sheet_client: GoogleSheetClient, master_sheet_url: str,
                 shared_folder_id: str = '1Uh3c_c7kakXIXlTX8BT89HtBcA3nCY1v'):
        self.client = sheet_client
        self.master_url = master_sheet_url
        self.folder_id = shared_folder_id

    def _read_master(self) -> List[List[str]]:
        """마스터 시트 전체 데이터 읽기"""
        return self.client.read_sheet(self.master_url)

    def load_vendors(self) -> List[Dict[str, Any]]:
        """활성 업체 목록 반환 (발주 처리용 포맷)"""
        data = self._read_master()
        if not data or len(data) < 2:
            return []

        headers = data[0]
        vendors = []
        for row in data[1:]:
            if len(row) < len(headers):
                row += [''] * (len(headers) - len(row))
            item = dict(zip(headers, row))
            if item.get('상태', '') == '비활성':
                continue
            phone = str(item.get('전화번호', '')).strip()
            if phone and not phone.startswith('0'):
                phone = '0' + phone
            vendors.append({
                'id': item.get('업체ID', ''),
                'name': item.get('업체명', ''),
                'contact_person': item.get('담당자', ''),
                'phone': phone,
                'email': item.get('이메일', ''),
                'google_sheet_url': item.get('구글시트URL', ''),
            })
        return vendors

    def load_all_vendors(self) -> List[Dict[str, Any]]:
        """모든 업체 목록 반환 (관리용, 비활성 포함)"""
        data = self._read_master()
        if not data or len(data) < 2:
            return []

        headers = data[0]
        vendors = []
        for row in data[1:]:
            if len(row) < len(headers):
                row += [''] * (len(headers) - len(row))
            item = dict(zip(headers, row))
            phone = str(item.get('전화번호', '')).strip()
            if phone and not phone.startswith('0'):
                phone = '0' + phone
            item['전화번호'] = phone
            vendors.append(item)
        return vendors

    def add_vendor(self, name: str, contact_person: str, phone: str,
                   email: str = '', sheet_url: str = None) -> Dict[str, Any]:
        """업체 추가 (전용 시트 자동 생성 또는 URL 직접 지정)"""
        vendor_id = f"vendor_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        today = datetime.now().strftime('%Y-%m-%d')

        # 전용 발주서 시트 생성 (URL 미지정 시)
        if not sheet_url:
            sheet_url = self._create_vendor_sheet(name)

        row = [vendor_id, name, contact_person, phone, email, sheet_url, today, '활성']
        self.client.append_row(self.master_url, row)

        return {
            'id': vendor_id,
            'name': name,
            'contact_person': contact_person,
            'phone': phone,
            'email': email,
            'google_sheet_url': sheet_url,
        }

    def update_vendor(self, vendor_id: str, **fields) -> bool:
        """업체 정보 수정"""
        try:
            spreadsheet = self.client.open_sheet_by_url(self.master_url)
            if not spreadsheet:
                return False
            worksheet = spreadsheet.get_worksheet(0)
            data = worksheet.get_all_values()
            if not data:
                return False

            headers = data[0]
            id_col = headers.index('업체ID') if '업체ID' in headers else 0

            for row_idx, row in enumerate(data[1:], start=2):
                if len(row) > id_col and row[id_col] == vendor_id:
                    for field, value in fields.items():
                        if field in headers:
                            col_idx = headers.index(field)
                            worksheet.update_cell(row_idx, col_idx + 1, value)
                    return True
            return False
        except Exception as e:
            logging.error(f'❌ 업체 수정 실패: {e}')
            return False

    def delete_vendor(self, vendor_id: str) -> bool:
        """업체 비활성화 (소프트 삭제)"""
        return self.update_vendor(vendor_id, **{'상태': '비활성'})

    def restore_vendor(self, vendor_id: str) -> bool:
        """업체 재활성화"""
        return self.update_vendor(vendor_id, **{'상태': '활성'})

    def _create_vendor_sheet(self, vendor_name: str) -> str:
        """업체 전용 발주서 시트 생성 (Apps Script 우선, 실패 시 서비스 계정)"""
        import requests as _req

        # 1) Apps Script 웹앱으로 생성 시도 (사용자 계정 드라이브 사용)
        apps_script_url = self._get_apps_script_url()
        if apps_script_url:
            try:
                resp = _req.post(apps_script_url,
                    json={'vendor_name': vendor_name}, timeout=30,
                    headers={'Content-Type': 'application/json'})
                if resp.status_code == 200:
                    result = resp.json()
                    if result.get('success') and result.get('sheet_url'):
                        logging.info(f'✅ {vendor_name} 전용 시트 생성 (Apps Script): {result["sheet_url"]}')
                        return result['sheet_url']
                    else:
                        logging.warning(f'⚠️ Apps Script 응답 오류: {result}')
            except Exception as e:
                logging.warning(f'⚠️ Apps Script 호출 실패: {e}')

        # 2) 서비스 계정으로 직접 생성 시도 (fallback)
        try:
            title = f'[발주도우미] {vendor_name}_발주서'
            spreadsheet = self.client.create_spreadsheet(title, folder_id=self.folder_id)

            worksheet = spreadsheet.sheet1
            worksheet.update('A1', [self.SHEET_HEADERS])
            worksheet.format('A1:J1', {
                'textFormat': {'bold': True},
                'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
            })
            worksheet.format('I2:J100', {
                'backgroundColor': {'red': 1.0, 'green': 1.0, 'blue': 0.8}
            })

            logging.info(f'✅ {vendor_name} 전용 시트 생성: {spreadsheet.url}')
            return spreadsheet.url
        except Exception as e:
            logging.error(f'❌ 시트 생성 실패: {e}')
            return ''

    def _get_apps_script_url(self) -> str:
        """Apps Script 웹앱 URL 가져오기"""
        try:
            import streamlit as st
            if hasattr(st, 'secrets') and "vendor_master" in st.secrets:
                return st.secrets["vendor_master"].get("apps_script_url", "")
        except Exception:
            pass
        return os.environ.get('APPS_SCRIPT_URL', '')
