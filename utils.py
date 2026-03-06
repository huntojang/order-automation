"""
공통 유틸리티 함수 모듈
"""
import json
import logging
import os
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
        """업체 정보 로드"""
        if self._secrets and "vendors" in self._secrets:
            vendors_raw = self._secrets["vendors"]
            if isinstance(vendors_raw, str):
                return json.loads(vendors_raw).get('vendors', [])
            return list(vendors_raw.get('vendors', []))

        vendors_file = os.path.join(self.config_dir, 'vendors.json')
        try:
            with open(vendors_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('vendors', [])
        except FileNotFoundError:
            logging.error(f'업체 정보 파일을 찾을 수 없습니다: {vendors_file}')
            return []
        except json.JSONDecodeError as e:
            logging.error(f'업체 정보 파일 형식 오류: {e}')
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

    def open_sheet_by_url(self, url: str):
        """URL로 시트 열기"""
        try:
            return self.client.open_by_url(url)
        except Exception as e:
            logging.error(f'❌ 시트 열기 실패 ({url}): {e}')
            return None

    def update_sheet(self, sheet_url: str, data: List[List[Any]],
                     worksheet_index: int = 0) -> bool:
        """
        시트 데이터 업데이트

        Args:
            sheet_url: 구글 시트 URL
            data: 업데이트할 데이터 (2차원 리스트)
            worksheet_index: 워크시트 인덱스 (기본값: 0)

        Returns:
            성공 여부
        """
        try:
            spreadsheet = self.open_sheet_by_url(sheet_url)
            if not spreadsheet:
                return False

            worksheet = spreadsheet.get_worksheet(worksheet_index)

            # 기존 데이터 클리어
            worksheet.clear()

            # 새 데이터 삽입
            if data:
                worksheet.update('A1', data)

            # 헤더 행 서식 적용 (굵게, 배경색)
            if len(data) > 0:
                worksheet.format('A1:K1', {
                    'textFormat': {'bold': True},
                    'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
                })

            # 송장번호 칸 노란색 강조 (J열: 택배사, K열: 송장번호)
            if len(data) > 1:
                last_row = len(data)
                worksheet.format(f'J2:K{last_row}', {
                    'backgroundColor': {'red': 1.0, 'green': 1.0, 'blue': 0.8}
                })

            logging.info(f'✅ 시트 업데이트 완료: {spreadsheet.title}')
            return True

        except Exception as e:
            logging.error(f'❌ 시트 업데이트 실패: {e}')
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
            spreadsheet = self.open_sheet_by_url(sheet_url)
            if not spreadsheet:
                return []

            worksheet = spreadsheet.get_worksheet(worksheet_index)
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
