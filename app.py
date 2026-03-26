"""
발주도우미 - 데모 웹앱 (Streamlit)
"""
import os
import json
import time
import base64
import logging
from datetime import datetime, timedelta
from io import BytesIO

import requests as _requests
import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh

from utils import Config, GoogleSheetClient, GoogleSheetOAuthClient, Logger, VendorManager
from send_orders import send_alimtalk

# 페이지 설정 (파비콘을 커스텀 이미지로)
from PIL import Image
_favicon = Image.open("assets/favicon_1.png")
st.set_page_config(
    page_title="발주도우미",
    page_icon=_favicon,
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 우측 상단 running 아이콘을 커스텀 로고 애니메이션으로 교체
def _load_b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

_b64_1 = _load_b64("assets/favicon_1.png")
_b64_2 = _load_b64("assets/favicon_2.png")
_b64_3 = _load_b64("assets/favicon_3.png")

# CSS 스타일 — 우측 상단 Running 아이콘 커스텀 (base64 포함이라 별도 블록)
st.markdown(f"""
<style>
    /* Running 아이콘: 원본 숨기고 커스텀 로고로 대체 */
    [data-testid="stStatusWidget"] .e1by5rsa2,
    [data-testid="stStatusWidget"] img,
    [data-testid="stStatusWidget"] svg {{
        visibility: hidden !important;
        position: relative;
    }}
    [data-testid="stStatusWidget"] .e1by5rsa2::after,
    [data-testid="stStatusWidget"] img::after {{
        content: "";
        position: absolute;
        top: 0; left: 0;
        width: 100%; height: 100%;
        background-size: contain;
        background-repeat: no-repeat;
        background-position: center;
        visibility: visible !important;
        animation: logoSwap 1.5s steps(1) infinite;
    }}
    /* fallback: stStatusWidget 자체에도 적용 */
    [data-testid="stStatusWidget"] {{
        position: relative;
    }}
    [data-testid="stStatusWidget"]::after {{
        content: "";
        position: absolute;
        top: 50%; left: 0;
        transform: translateY(-50%);
        width: 28px; height: 28px;
        background-size: contain;
        background-repeat: no-repeat;
        background-position: center;
        animation: logoSwap 1.5s steps(1) infinite;
        pointer-events: none;
    }}
    @keyframes logoSwap {{
        0%   {{ background-image: url("data:image/png;base64,{_b64_3}"); }}
        33%  {{ background-image: url("data:image/png;base64,{_b64_2}"); }}
        66%  {{ background-image: url("data:image/png;base64,{_b64_1}"); }}
    }}
</style>
""", unsafe_allow_html=True)

# CSS 스타일 — 메인 (Green Modern)
st.markdown("""
<style>
    /* ===== 글로벌 ===== */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"], [class*="st-"] { font-family: 'Inter', sans-serif !important; }
    .block-container, [data-testid="stAppViewBlockContainer"] {
        max-width: 1100px !important; padding-top: 1.5rem !important;
    }
    iframe[title="streamlit_autorefresh.st_autorefresh"] {
        position: absolute !important; height: 0 !important; overflow: hidden !important;
    }
    /* Streamlit 기본 헤더/푸터 숨기기 */
    header[data-testid="stHeader"] { background: transparent !important; }

    /* ===== 헤더 ===== */
    .main-header {
        font-size: 1.5rem !important; font-weight: 700 !important; color: #1A1A1A !important;
        margin-bottom: 1.25rem !important; letter-spacing: -0.03em !important;
        line-height: 1.3 !important;
    }
    .sub-header { display: none !important; }
    .section-title {
        font-size: 0.78rem !important; font-weight: 600 !important; color: #999 !important;
        text-transform: uppercase !important; letter-spacing: 0.08em !important;
        margin-bottom: 0.75rem !important; margin-top: 0.5rem !important;
    }

    /* ===== 카드 ===== */
    .card {
        background: #F5F5F5 !important; border-radius: 16px !important; padding: 1.5rem !important;
        border: none !important;
    }
    .card-accent {
        background: #2E643C !important; border-radius: 16px !important; padding: 1.5rem !important;
        color: white !important; border: none !important;
    }
    .card-title { font-size: 0.85rem !important; font-weight: 500 !important; color: #888 !important; margin-bottom: 8px !important; }
    .card-value { font-size: 2rem !important; font-weight: 700 !important; color: #1A1A1A !important; }
    .card-desc { font-size: 0.8rem !important; color: #AAA !important; }
    .card-accent .card-title { color: rgba(255,255,255,0.7) !important; }
    .card-accent .card-value { color: #fff !important; }
    .card-accent .card-desc { color: rgba(255,255,255,0.5) !important; }

    /* ===== 리스트 행 ===== */
    .list-row {
        display: flex !important; align-items: center !important; justify-content: space-between !important;
        background: #F5F5F5 !important; border-radius: 14px !important; padding: 1rem 1.25rem !important;
        margin-bottom: 8px !important; border: none !important;
    }
    .list-row:hover { background: #EFEFEF !important; }
    .list-row-active { background: #2E643C !important; color: white !important; }
    .list-row-active:hover { background: #255633 !important; }
    .list-row-active .list-desc { color: rgba(255,255,255,0.6) !important; }
    .list-name { font-weight: 600 !important; font-size: 0.92rem !important; }
    .list-desc { font-size: 0.82rem !important; color: #888 !important; }
    .list-arrow {
        width: 34px !important; height: 34px !important; border-radius: 50% !important;
        background: #2E643C !important; color: white !important; display: flex !important;
        align-items: center !important; justify-content: center !important; font-size: 0.9rem !important;
        flex-shrink: 0 !important; text-decoration: none !important;
    }
    .list-row-active .list-arrow { background: rgba(255,255,255,0.2) !important; }

    /* ===== 카카오 메시지 ===== */
    .kakao-msg {
        background: #FEE500 !important; color: #3C1E1E !important; padding: 1.25rem !important;
        border-radius: 16px !important; margin: 0.5rem 0 !important; font-size: 0.88rem !important;
        max-width: 340px !important;
    }

    /* ===== 알림바 ===== */
    .notification-bar {
        background: #2E643C !important; color: white !important;
        padding: 0.8rem 1.2rem !important; border-radius: 12px !important;
        margin: 0.5rem 0 !important; font-size: 0.88rem !important;
    }

    /* ===== 사이드바 숨기기 ===== */
    [data-testid="stSidebar"],
    section[data-testid="stSidebar"],
    button[data-testid="stSidebarCollapsedControl"],
    button[kind="headerNoPadding"] { display: none !important; }

    /* ===== 상단 네비게이션 탭 ===== */
    /* 라디오 버튼을 탭처럼 보이게 */
    div[data-testid="stRadio"] > div {
        flex-direction: row !important; gap: 0 !important;
    }
    div[data-testid="stRadio"] > div > label,
    div[data-testid="stRadio"] > div[role="radiogroup"] > label {
        background: transparent !important;
        border: none !important;
        border-bottom: 2px solid transparent !important;
        border-radius: 0 !important;
        padding: 10px 24px !important;
        margin: 0 !important;
        font-weight: 500 !important;
        font-size: 0.9rem !important;
        color: #999 !important;
        cursor: pointer !important;
    }
    div[data-testid="stRadio"] > div > label:hover,
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
        color: #2E643C !important;
    }
    div[data-testid="stRadio"] > div > label[data-checked="true"],
    div[data-testid="stRadio"] > div > label:has(input:checked),
    div[data-testid="stRadio"] > div[role="radiogroup"] > label[data-checked="true"],
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
        color: #1A1A1A !important;
        font-weight: 700 !important;
        border-bottom: 2px solid #2E643C !important;
    }
    /* 라디오 동그라미 숨기기 */
    div[data-testid="stRadio"] > div > label > div:first-child,
    div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child,
    div[data-testid="stRadio"] [data-testid="stMarkdownContainer"]:empty {
        display: none !important;
    }
    /* 라디오 컨테이너 하단 보더 */
    div[data-testid="stRadio"] {
        border-bottom: 1px solid #E8E8E8 !important;
        margin-bottom: 1.5rem !important;
        padding-bottom: 0 !important;
    }
    /* 라디오 레이블 숨기기 */
    div[data-testid="stRadio"] > label { display: none !important; }

    /* ===== 버튼 ===== */
    button[kind="primary"],
    .stButton > button[kind="primary"],
    [data-testid="stButton"] > button[kind="primary"] {
        background-color: #2E643C !important; border: none !important;
        border-radius: 12px !important; padding: 0.65rem 1.5rem !important;
        font-weight: 600 !important; font-size: 0.92rem !important;
        color: white !important;
    }
    button[kind="primary"]:hover,
    .stButton > button[kind="primary"]:hover {
        background-color: #255633 !important;
        box-shadow: 0 4px 16px rgba(46,100,60,0.25) !important;
    }
    button, .stButton > button {
        border-radius: 12px !important; font-weight: 500 !important;
    }
    /* form submit 버튼도 동일 색상 */
    [data-testid="stFormSubmitButton"] > button,
    [data-testid="stFormSubmitButton"] > button[kind="secondaryFormSubmit"],
    [data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"] {
        background-color: #2E643C !important; color: white !important;
        border: none !important; border-radius: 12px !important;
        font-weight: 600 !important;
    }
    [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #255633 !important;
        box-shadow: 0 4px 16px rgba(46,100,60,0.25) !important;
    }

    /* ===== 입력 필드 ===== */
    [data-testid="stTextInput"] input {
        background-color: #F5F5F5 !important;
        border: 1px solid #E0E0E0 !important;
        border-radius: 10px !important;
    }
    [data-testid="stTextInput"] input:focus {
        border-color: #BDBDBD !important;
        box-shadow: 0 0 0 1px #BDBDBD !important;
    }
    /* 전화번호 필드 파란색 강조 (aria-label 기반) */
    input[aria-label*="전화번호"] {
        background-color: #E8F5E9 !important;
        border: 1px solid #A5D6A7 !important;
    }
    input[aria-label*="전화번호"]:focus {
        border-color: #4CAF50 !important;
        box-shadow: 0 0 0 1px #4CAF50 !important;
    }

    /* ===== 프로그레스바 ===== */
    [data-testid="stProgress"] > div > div,
    .stProgress > div > div { background-color: #E0E0E0 !important; border-radius: 8px !important; }
    [data-testid="stProgress"] > div > div > div,
    .stProgress > div > div > div { background-color: #2E643C !important; border-radius: 8px !important; }

    /* ===== metric 카드 ===== */
    [data-testid="stMetric"],
    [data-testid="stMetricValue"],
    div[data-testid="stMetric"] {
        background: #F5F5F5 !important; border: none !important;
        border-radius: 16px !important;
        display: flex !important; flex-direction: column !important;
        justify-content: center !important; align-items: center !important;
        text-align: center !important; min-height: 100px !important;
    }
    [data-testid="stMetricValue"] > div { color: #1A1A1A !important; font-weight: 700 !important; }
    [data-testid="stMetricLabel"] > div { color: #888 !important; font-weight: 500 !important; }
    [data-testid="stMetricDelta"] svg { display: none !important; }

    /* ===== 파일 업로더 ===== */
    [data-testid="stFileUploader"] section {
        border-radius: 16px !important; border: 2px dashed #C8DECE !important;
        background: #E8F3EB !important; padding: 2rem !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        font-size: 0 !important; line-height: 0 !important;
    }
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "엑셀업로드" !important; font-size: 0.875rem !important;
        line-height: normal !important;
    }
    /* 업로드된 파일명 영역 숨기기 (expander에서 표시) */
    [data-testid="stFileUploaderFile"] {
        display: none !important;
    }

    /* ===== expander ===== */
    [data-testid="stExpander"] summary {
        background: #F5F5F5 !important; border-radius: 12px !important;
        font-weight: 500 !important; padding: 0.8rem 1rem !important;
    }

    /* ===== 다운로드 버튼 ===== */
    .stDownloadButton > button,
    [data-testid="stDownloadButton"] > button {
        background-color: #2E643C !important; color: white !important;
        border: none !important; border-radius: 12px !important; font-weight: 600 !important;
    }
    .stDownloadButton > button:hover,
    [data-testid="stDownloadButton"] > button:hover {
        background-color: #255633 !important;
    }

    /* ===== 구분선 ===== */
    hr { border-color: #EBEBEB !important; }

    /* ===== 알림 메시지 스타일 ===== */
    [data-testid="stAlert"] {
        border-radius: 12px !important; border: none !important;
    }

    /* ===== Streamlit 기본 뱃지/푸터/툴바 완전 숨기기 ===== */
    .viewerBadge_container__r5tak,
    .styles_viewerBadge__CvC9N,
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    .stDeployButton,
    #MainMenu,
    footer,
    footer *,
    .reportview-container .main footer,
    a[href*="streamlit.io"],
    a[href*="github.com"][data-testid],
    img[src*="avatar"],
    div[class*="viewerBadge"],
    div[class*="deployBtn"] {
        display: none !important; visibility: hidden !important;
        height: 0 !important; min-height: 0 !important;
        max-height: 0 !important; overflow: hidden !important;
        opacity: 0 !important; pointer-events: none !important;
    }
    footer:after { display: none !important; content: "" !important; }
</style>
""", unsafe_allow_html=True)


UPLOAD_LOG_FILE = 'logs/upload_history.json'


def load_upload_history():
    """업로드 기록 로드"""
    if os.path.exists(UPLOAD_LOG_FILE):
        with open(UPLOAD_LOG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_upload_log(entry):
    """업로드 기록 저장"""
    os.makedirs('logs', exist_ok=True)
    history = load_upload_history()
    history.insert(0, entry)
    with open(UPLOAD_LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


@st.cache_resource
def get_sheet_client():
    """구글 시트 클라이언트 초기화 (앱 전체에서 1회만)"""
    config = Config()
    try:
        cred_dict = config.get_google_credentials_dict()
        if cred_dict:
            return GoogleSheetClient(credentials_dict=cred_dict)
        cred_file = config.get_google_credentials_file()
        if os.path.exists(cred_file):
            return GoogleSheetClient(credentials_file=cred_file)
    except Exception:
        pass
    return None


def load_vendors():
    """업체 정보 로드 (마스터 시트 우선, 캐시 없음)"""
    config = Config()
    master_url = config.get_vendor_master_url()
    if master_url:
        client = get_sheet_client()
        if client:
            try:
                vm = VendorManager(client, master_url, config.get_shared_folder_id())
                vendors = vm.load_vendors()
                if vendors:
                    return vendors
            except Exception:
                pass
    return config.load_vendors()


def fetch_dashboard(_client, master_url):
    """대시보드 탭에서 송장 현황 가져오기 (API 1회, 캐시 1분)
    Apps Script가 1분마다 갱신하는 대시보드 데이터를 읽음.
    """
    now = time.time()
    cache = st.session_state.get('_dashboard_cache', {})
    cache_time = st.session_state.get('_dashboard_cache_time', 0)

    if cache and (now - cache_time) < 60:
        return cache

    if not _client or not master_url:
        return {}

    try:
        spreadsheet = _client.open_sheet_by_url(master_url)
        if not spreadsheet:
            logging.warning('⚠️ 대시보드: 마스터 시트를 열 수 없음')
            return {}
        try:
            worksheet = spreadsheet.worksheet('대시보드')
        except Exception as ws_err:
            logging.warning(f'⚠️ 대시보드 탭 없음: {ws_err}')
            return {}
        data = worksheet.get_all_values()
        logging.info(f'📊 대시보드 탭 데이터: {len(data)}행, 헤더={data[0] if data else "없음"}')
        if not data or len(data) < 2:
            logging.warning(f'⚠️ 대시보드 데이터 부족: {len(data) if data else 0}행')
            return {}

        headers = data[0]
        result = {}
        for row in data[1:]:
            if len(row) < len(headers):
                row += [''] * (len(headers) - len(row))
            item = dict(zip(headers, row))
            name = item.get('업체명', '')
            if not name:
                continue
            total = int(item.get('전체주문', 0)) if str(item.get('전체주문', '')).lstrip('-').isdigit() else 0
            invoiced = int(item.get('송장완료', 0)) if str(item.get('송장완료', '')).isdigit() else 0
            if total < 0:
                total = 0  # 읽기 실패한 업체
            result[name] = {
                'total': total,
                'invoiced': invoiced,
                'pending': max(0, total - invoiced),
                'rate': int(item.get('완료율', 0)) if str(item.get('완료율', '')).isdigit() else 0,
            }

        st.session_state['_dashboard_cache'] = result
        st.session_state['_dashboard_cache_time'] = now
        return result

    except Exception as e:
        logging.error(f'❌ 대시보드 읽기 실패: {e}')
        return {}


def prepare_sheet_data(df):
    """구글 시트에 업로드할 데이터 준비"""
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


def split_by_vendor(df, vendor_column='공급처'):
    """업체별로 데이터 분류"""
    if vendor_column not in df.columns:
        return {}
    vendor_data = {}
    for vendor_name, group in df.groupby(vendor_column):
        vendor_data[vendor_name] = group.reset_index(drop=True)
    return vendor_data


# ===== SSO 인증 =====
QUESTLOOM_URL = "https://questloom.io"
SSO_SERVICE = "orderhelper"


def sso_login(email: str, password: str) -> dict:
    """questloom.io SSO 로그인"""
    try:
        resp = _requests.post(
            f"{QUESTLOOM_URL}/api/sso/login",
            json={"email": email, "password": password, "service": SSO_SERVICE},
            timeout=10
        )
        return resp.json()
    except Exception as e:
        return {"error": str(e)}


def _decode_jwt_payload(token: str) -> dict:
    """JWT 페이로드를 디코딩 (서명 검증 없이 — 만료 체크용)"""
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return {}
        payload_b64 = parts[1]
        # base64 패딩 보정
        padding = 4 - len(payload_b64) % 4
        if padding != 4:
            payload_b64 += "=" * padding
        payload_json = base64.urlsafe_b64decode(payload_b64)
        return json.loads(payload_json)
    except Exception:
        return {}


def _restore_session_from_token(token: str) -> bool:
    """JWT 토큰에서 세션 복원. 만료됐으면 False"""
    payload = _decode_jwt_payload(token)
    if not payload:
        return False

    # 만료 체크
    exp = payload.get("exp", 0)
    if time.time() > exp:
        return False

    # 구독 정보 추출
    subs = payload.get("subscriptions", [])
    has_orderhelper = any(s.get("serviceSlug") == SSO_SERVICE for s in subs)

    st.session_state["sso_token"] = token
    st.session_state["sso_authenticated"] = True
    st.session_state["sso_user"] = {
        "id": payload.get("userId", ""),
        "email": payload.get("email", ""),
        "companyId": payload.get("companyId", ""),
        "role": payload.get("role", ""),
    }
    st.session_state["sso_subscriptions"] = subs
    st.session_state["sso_has_subscription"] = has_orderhelper
    return True


def check_auth():
    """인증 상태 확인. 미인증이면 로그인 폼을 보여주고 True 반환 (= 차단)"""
    # 이미 인증됨
    if st.session_state.get("sso_authenticated"):
        return False

    # 쿼리 파라미터에 토큰이 있으면 세션 복원 시도 (새로고침 / 리다이렉트)
    token = st.query_params.get("token")
    if token:
        if _restore_session_from_token(token):
            return False  # 복원 성공
        else:
            # 토큰 만료 → 파라미터 제거
            st.query_params.clear()

    # 로그인 화면
    st.markdown("""
    <style>
    .login-title {
        text-align: center;
        font-size: 24px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .login-subtitle {
        text-align: center;
        color: #666;
        font-size: 14px;
        margin-bottom: 24px;
    }
    </style>
    """, unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        _login_logo_b64 = _load_b64("assets/logo.png")
        st.markdown(f'<img src="data:image/png;base64,{_login_logo_b64}" style="width:160px;height:auto;" />', unsafe_allow_html=True)
        st.markdown('<div class="login-title">발주도우미</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle">QuestLoom 계정으로 로그인하세요</div>', unsafe_allow_html=True)

        with st.form("login_form"):
            email = st.text_input("이메일")
            password = st.text_input("비밀번호", type="password")
            submitted = st.form_submit_button("로그인", type="primary", use_container_width=True)

            if submitted:
                if not email or not password:
                    st.error("이메일과 비밀번호를 입력해주세요.")
                else:
                    with st.spinner("로그인 중..."):
                        result = sso_login(email, password)

                    if result.get("error"):
                        error_msg = result["error"]
                        if "Invalid email" in error_msg:
                            st.error("이메일 또는 비밀번호가 올바르지 않습니다.")
                        else:
                            st.error(f"로그인 실패: {error_msg}")
                    elif result.get("token"):
                        st.session_state["sso_token"] = result["token"]
                        st.session_state["sso_user"] = result.get("user", {})
                        st.session_state["sso_subscriptions"] = result.get("subscriptions", [])
                        st.session_state["sso_authenticated"] = True
                        st.session_state["sso_has_subscription"] = result.get("hasActiveSubscription", False)
                        # 토큰을 쿼리 파라미터에 저장 (새로고침 시 유지)
                        st.query_params["token"] = result["token"]
                        st.rerun()

        st.markdown("---")
        st.caption(f"아직 계정이 없으신가요? [QuestLoom에서 가입하기]({QUESTLOOM_URL}/signup)")

    return True  # 미인증 → 차단


# 인증 체크 (개발 모드에서는 건너뛸 수 있도록)
_skip_auth = os.environ.get("SKIP_AUTH", "").lower() == "true"
if not _skip_auth and check_auth():
    st.stop()

# 구독 미활성 시 안내
if not _skip_auth and not st.session_state.get("sso_has_subscription", False):
    _user = st.session_state.get("sso_user", {})
    _user_email = _user.get("email", "")
    _user_name = _user.get("name", "")

    st.markdown(f"""
    <style>
    .paywall-wrap {{
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 70vh;
    }}
    .paywall-card {{
        max-width: 480px;
        width: 100%;
        text-align: center;
        padding: 48px 40px;
        border-radius: 16px;
        background: linear-gradient(135deg, #f8faf8 0%, #eef4ee 100%);
        border: 1px solid #d4e4d4;
        box-shadow: 0 4px 24px rgba(46,100,60,0.08);
    }}
    .paywall-icon {{
        font-size: 48px;
        margin-bottom: 16px;
    }}
    .paywall-title {{
        font-size: 22px;
        font-weight: 700;
        color: #1a1a1a;
        margin-bottom: 8px;
    }}
    .paywall-desc {{
        font-size: 15px;
        color: #666;
        line-height: 1.6;
        margin-bottom: 28px;
    }}
    .paywall-plans {{
        display: flex;
        gap: 12px;
        margin-bottom: 28px;
    }}
    .paywall-plan {{
        flex: 1;
        padding: 16px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        background: #fff;
    }}
    .paywall-plan-name {{
        font-size: 14px;
        font-weight: 600;
        color: #2E643C;
        margin-bottom: 4px;
    }}
    .paywall-plan-price {{
        font-size: 20px;
        font-weight: 700;
        color: #1a1a1a;
    }}
    .paywall-plan-unit {{
        font-size: 13px;
        color: #999;
    }}
    .paywall-cta {{
        display: inline-block;
        padding: 14px 40px;
        background: #2E643C;
        color: #fff !important;
        text-decoration: none;
        border-radius: 8px;
        font-size: 16px;
        font-weight: 600;
        transition: background 0.2s;
    }}
    .paywall-cta:hover {{
        background: #245232;
        color: #fff !important;
    }}
    .paywall-footer {{
        margin-top: 20px;
        font-size: 13px;
        color: #999;
    }}
    .paywall-footer a {{
        color: #2E643C;
    }}
    </style>
    <div class="paywall-wrap">
        <div class="paywall-card">
            <div class="paywall-icon">🔒</div>
            <div class="paywall-title">구독이 필요한 서비스입니다</div>
            <div class="paywall-desc">
                {_user_name or _user_email}님, 반갑습니다!<br>
                발주도우미를 이용하려면 QuestLoom에서 구독을 시작해주세요.
            </div>
            <div class="paywall-plans">
                <div class="paywall-plan">
                    <div class="paywall-plan-name">스타터</div>
                    <div class="paywall-plan-price">39,000<span class="paywall-plan-unit">원/월</span></div>
                </div>
                <div class="paywall-plan">
                    <div class="paywall-plan-name">프로</div>
                    <div class="paywall-plan-price">79,000<span class="paywall-plan-unit">원/월</span></div>
                </div>
            </div>
            <a href="{QUESTLOOM_URL}/console/services" target="_blank" class="paywall-cta">구독 시작하기</a>
            <div class="paywall-footer">
                {_user_email} · <a href="/?logout=1">로그아웃</a>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.query_params.get("logout"):
        for key in list(st.session_state.keys()):
            if key.startswith("sso_"):
                del st.session_state[key]
        st.query_params.clear()
        st.rerun()
    st.stop()

# 로그아웃 처리
if st.query_params.get("logout"):
    for key in list(st.session_state.keys()):
        if key.startswith("sso_"):
            del st.session_state[key]
    st.query_params.clear()
    st.rerun()


# ===== 상단 헤더 + 탭 네비게이션 =====
_col_logo, _col_spacer, _col_user = st.columns([1, 2, 1])
with _col_logo:
    _logo_b64 = _load_b64("assets/logo.png")
    st.markdown(f'<img src="data:image/png;base64,{_logo_b64}" style="width:160px;height:auto;" />', unsafe_allow_html=True)
with _col_user:
    if not _skip_auth:
        _sso_user = st.session_state.get("sso_user", {})
        _user_email = _sso_user.get("email", "")
        if _user_email:
            st.markdown(f'<div style="text-align:right;padding-top:12px;font-size:13px;color:#666;">{_user_email} · <a href="/?logout=1">로그아웃</a></div>', unsafe_allow_html=True)

st.markdown('<div class="nav-tabs">', unsafe_allow_html=True)
page = st.radio(
    "메뉴",
    ["발주 업로드", "송장 현황", "송장 다운로드", "업체 관리"],
    horizontal=True,
    label_visibility="collapsed"
)
st.markdown('</div>', unsafe_allow_html=True)


# ===== 송장 다운로드 페이지에서 백그라운드 송장 변경 감지 (5분 간격) =====
if page == "송장 다운로드":
    bg_refresh = st_autorefresh(interval=300000, limit=None, key="bg_autorefresh")

    # 대시보드에서 변경 감지 (API 1회)
    try:
        _bg_client = get_sheet_client()
        _bg_config = Config()
        _bg_master_url = _bg_config.get_vendor_master_url()
        _bg_dashboard = fetch_dashboard(_bg_client, _bg_master_url)
        if _bg_dashboard:
            _bg_total_invoices = sum(d['invoiced'] for d in _bg_dashboard.values())
            _prev = st.session_state.get('prev_invoice_count', None)
            _prev_vs = st.session_state.get('prev_vendor_status', {})

            if _prev is not None and _bg_total_invoices > _prev:
                for _vn, _vd in _bg_dashboard.items():
                    _pc = _prev_vs.get(_vn, 0)
                    if _vd['invoiced'] > _pc:
                        st.toast(f"{_vn} 송장 +{_vd['invoiced'] - _pc}건")

            st.session_state['prev_invoice_count'] = _bg_total_invoices
            st.session_state['prev_vendor_status'] = {n: d['invoiced'] for n, d in _bg_dashboard.items()}
    except Exception:
        pass


# ===== 발주 업로드 =====
if page == "발주 업로드":
    st.markdown('<div class="main-header">발주 업로드</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "이지어드민 발주서 엑셀/CSV 파일을 올려주세요",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        help="여러 파일을 한번에 올릴 수 있습니다"
    )

    if uploaded_files:
        # 현재 파일 키 (뱃지 표시 판단용)
        _current_file_key = "|".join(sorted(f.file_id for f in uploaded_files))

        all_dfs = []
        file_names = []
        for f in uploaded_files:
            if f.name.endswith('.csv'):
                df = pd.read_csv(f, encoding='utf-8')
            elif f.name.endswith('.xls'):
                df = pd.read_excel(f, engine='xlrd')
            else:
                df = pd.read_excel(f, engine='openpyxl')
            all_dfs.append(df)
            file_names.append(f.name)

        merged_df = pd.concat(all_dfs, ignore_index=True)

        # [1] 파일명 — 데이터 미리보기 (총 OO건)
        file_label = ", ".join(file_names)
        with st.expander(f"{file_label} — 데이터 미리보기 (총 {len(merged_df)}건)", expanded=False):
            st.dataframe(merged_df.head(10), use_container_width=True)

        # 프로그레스 영역 (발주 실행 시 스크롤 없이 바로 보이도록 그리드 위에 배치)
        progress_area = st.container()

        # [2] 공급처별 주문 현황
        vendor_data = split_by_vendor(merged_df)
        if vendor_data:
            # 발주 실행이 완료된 파일에 대해서만 뱃지 표시
            _processed_key = st.session_state.get('_processed_file_key', '')
            _show_badges = (_current_file_key == _processed_key)
            _upload_results = st.session_state.get('_upload_results', {}) if _show_badges else {}
            _alimtalk_results = st.session_state.get('_alimtalk_results', {}) if _show_badges else {}
            with st.expander(f"공급처별 주문 현황 ({len(vendor_data)}개 업체)", expanded=False):
                grid_html = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px;">'
                for name, vdf in vendor_data.items():
                    # 뱃지 HTML 생성
                    badges_html = ''
                    if _upload_results:
                        up_ok = _upload_results.get(name)
                        if up_ok is True:
                            badges_html += '<span style="background:#E8F5E9;color:#2E643C;padding:1px 6px;border-radius:8px;font-size:0.6rem;font-weight:600;">업로드완료</span> '
                        elif up_ok is not None:
                            badges_html += '<span style="background:#FFEBEE;color:#C62828;padding:1px 6px;border-radius:8px;font-size:0.6rem;font-weight:600;">업로드실패</span> '
                        al_ok = _alimtalk_results.get(name)
                        if al_ok is True:
                            badges_html += '<span style="background:#E8F5E9;color:#2E643C;padding:1px 6px;border-radius:8px;font-size:0.6rem;font-weight:600;">톡발송완료</span>'
                        elif al_ok is False:
                            badges_html += '<span style="background:#FFEBEE;color:#C62828;padding:1px 6px;border-radius:8px;font-size:0.6rem;font-weight:600;">톡발송실패</span>'
                    grid_html += f'''<div style="background:#F5F5F5;border-radius:10px;padding:10px 12px;text-align:center;">
                        <div style="font-size:0.78rem;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{name}">{name}</div>
                        <div style="font-size:1.2rem;font-weight:700;color:#1A1A1A;">{len(vdf)}건</div>
                        <div style="margin-top:4px;min-height:16px;">{badges_html}</div>
                    </div>'''
                grid_html += '</div>'
                st.markdown(grid_html, unsafe_allow_html=True)

        # 실패 항목 재시도 섹션
        if _upload_results:
            _failed_uploads = [n for n, ok in _upload_results.items() if ok is False or (ok is not True and ok is not None)]
            _failed_alimtalk = [n for n, ok in _alimtalk_results.items() if ok is False]
            # 전화번호 없는 업체는 톡 실패에서 제외 (재시도 불가)
            _vendors_with_phone = set()
            _retry_vendors_info = load_vendors()
            for _vi in _retry_vendors_info:
                if _vi.get('phones') or _vi.get('phone', '').strip():
                    _vendors_with_phone.add(_vi['name'])
            _failed_alimtalk = [n for n in _failed_alimtalk if n in _vendors_with_phone]

            if _failed_uploads or _failed_alimtalk:
                with st.expander(f"실패 항목 재시도 ({len(_failed_uploads) + len(_failed_alimtalk)}건)", expanded=True):
                    # 전체 재시도 버튼
                    col_retry_all_up, col_retry_all_al = st.columns(2)
                    if _failed_uploads:
                        with col_retry_all_up:
                            if st.button(f"업로드 실패 전체 재시도 ({len(_failed_uploads)}건)", key="retry_all_uploads"):
                                _r_client = get_sheet_client()
                                _r_progress = st.progress(0)
                                for _ri, _rn in enumerate(_failed_uploads):
                                    _r_progress.progress((_ri + 1) / len(_failed_uploads), text=f"재시도: {_rn}")
                                    if _rn in vendor_data:
                                        _r_vinfo = next((v for v in _retry_vendors_info if v['name'] == _rn), None)
                                        if _r_vinfo and _r_vinfo.get('google_sheet_url'):
                                            _r_data = prepare_sheet_data(vendor_data[_rn])
                                            _r_result = _r_client.update_sheet(_r_vinfo['google_sheet_url'], _r_data)
                                            if _r_result is True:
                                                _upload_results[_rn] = True
                                                st.success(f"{_rn} 업로드 성공")
                                            else:
                                                st.error(f"{_rn} 업로드 실패: {_r_result}")
                                    time.sleep(1)
                                _r_progress.empty()
                                st.session_state['_upload_results'] = _upload_results
                                st.rerun()
                    if _failed_alimtalk:
                        with col_retry_all_al:
                            if st.button(f"톡발송 실패 전체 재시도 ({len(_failed_alimtalk)}건)", key="retry_all_alimtalk"):
                                _al_config = Config().load_alimtalk_config()
                                for _rn in _failed_alimtalk:
                                    _r_vinfo = next((v for v in _retry_vendors_info if v['name'] == _rn), None)
                                    if _r_vinfo and _al_config and _rn in vendor_data:
                                        _sent = send_alimtalk(_r_vinfo, len(vendor_data[_rn]), _al_config)
                                        _alimtalk_results[_rn] = _sent
                                        if _sent:
                                            st.success(f"{_rn} 톡발송 성공")
                                        else:
                                            st.error(f"{_rn} 톡발송 실패")
                                    time.sleep(0.3)
                                st.session_state['_alimtalk_results'] = _alimtalk_results
                                st.rerun()

                    # 개별 재시도 버튼
                    st.markdown("---")
                    for _fn in sorted(set(_failed_uploads + _failed_alimtalk)):
                        _fc1, _fc2, _fc3 = st.columns([3, 1, 1])
                        with _fc1:
                            st.markdown(f"**{_fn}**")
                        with _fc2:
                            if _fn in _failed_uploads:
                                if st.button("업로드", key=f"retry_up_{_fn}"):
                                    _r_client = get_sheet_client()
                                    _r_vinfo = next((v for v in _retry_vendors_info if v['name'] == _fn), None)
                                    if _r_vinfo and _r_vinfo.get('google_sheet_url') and _fn in vendor_data:
                                        _r_data = prepare_sheet_data(vendor_data[_fn])
                                        _r_result = _r_client.update_sheet(_r_vinfo['google_sheet_url'], _r_data)
                                        if _r_result is True:
                                            _upload_results[_fn] = True
                                            st.session_state['_upload_results'] = _upload_results
                                            st.rerun()
                                        else:
                                            st.error(f"실패: {_r_result}")
                        with _fc3:
                            if _fn in _failed_alimtalk:
                                if st.button("톡발송", key=f"retry_al_{_fn}"):
                                    _al_config = Config().load_alimtalk_config()
                                    _r_vinfo = next((v for v in _retry_vendors_info if v['name'] == _fn), None)
                                    if _r_vinfo and _al_config and _fn in vendor_data:
                                        _sent = send_alimtalk(_r_vinfo, len(vendor_data[_fn]), _al_config)
                                        _alimtalk_results[_fn] = _sent
                                        st.session_state['_alimtalk_results'] = _alimtalk_results
                                        if _sent:
                                            st.rerun()
                                        else:
                                            st.error("톡발송 실패")

        # 발주 성공 결과 표시 (rerun 후)
        _last_success = st.session_state.pop('_last_success', None)
        if _last_success:
            _success_b64 = _load_b64("assets/success.png")
            st.markdown(f"""
            <div style="
                display:flex; justify-content:center; align-items:center;
                padding:2rem 0; animation: successPop 0.6s ease-out;
            ">
                <img src="data:image/png;base64,{_success_b64}"
                     style="width:260px; filter:drop-shadow(0 4px 20px rgba(64,144,195,0.3));" />
            </div>
            <style>
                @keyframes successPop {{
                    0% {{ opacity:0; transform:scale(0.5); }}
                    60% {{ opacity:1; transform:scale(1.05); }}
                    100% {{ opacity:1; transform:scale(1); }}
                }}
            </style>
            """, unsafe_allow_html=True)
            sc = _last_success['success_count']
            ac = _last_success['alimtalk_sent']
            to = _last_success['total_orders']
            if ac > 0:
                st.success(f"발주 처리 완료! {sc}개 업체 시트 업로드 + {ac}개 업체 알림톡 발송 완료")
            else:
                st.success(f"발주 처리 완료! {sc}개 업체에 총 {to}건 시트 업로드 완료 (알림톡 미설정)")

        # 발주 실행 버튼
        if st.button("발주 실행", type="primary", use_container_width=True):
            vendors_info = load_vendors()
            sheet_client = get_sheet_client()

            if not sheet_client:
                st.error("구글 시트 연결이 필요합니다")
            else:
                with progress_area:
                    progress = st.progress(0, text="발주 처리 중...")
                    status_container = st.container()

                total = len(vendors_info)
                success_count = 0
                _upload_results = {}

                for idx, vendor_info in enumerate(vendors_info):
                    vendor_name = vendor_info['name']
                    progress.progress((idx + 1) / total, text=f"처리 중: {vendor_name}")

                    if vendor_name not in vendor_data:
                        continue

                    vendor_df = vendor_data[vendor_name]
                    sheet_data = prepare_sheet_data(vendor_df)

                    # 구글 시트 업로드
                    sheet_url = vendor_info.get('google_sheet_url', '')
                    if sheet_url:
                        try:
                            result = sheet_client.update_sheet(sheet_url, sheet_data)
                            if result is True:
                                _upload_results[vendor_name] = True
                                success_count += 1
                            else:
                                _upload_results[vendor_name] = False
                                with status_container:
                                    st.error(f"{vendor_name} — 업로드 실패: {result}")
                        except Exception as e:
                            _upload_results[vendor_name] = False
                            with status_container:
                                st.error(f"{vendor_name} — 업로드 오류: {e}")
                    else:
                        _upload_results[vendor_name] = False
                        with status_container:
                            st.warning(f"{vendor_name} — 시트 URL 없음")

                    time.sleep(2)  # 429 방지: 분당 60회 Read 제한 준수

                progress.progress(1.0, text="시트 업로드 완료! 알림톡 발송 준비 중...")

                # 알림톡 발송 (실제 API 호출)
                alimtalk_config = Config().load_alimtalk_config()
                _alimtalk_logs = []
                _alimtalk_results = {}
                _alimtalk_targets = [vi for vi in vendors_info if vi['name'] in vendor_data]
                _alimtalk_total = len(_alimtalk_targets)
                for _ai, vendor_info in enumerate(_alimtalk_targets):
                    vendor_name = vendor_info['name']
                    order_count = len(vendor_data[vendor_name])
                    sheet_url = vendor_info.get('google_sheet_url', '')
                    progress.progress((_ai + 1) / _alimtalk_total, text=f"알림톡 발송 중... ({_ai + 1}/{_alimtalk_total}) {vendor_name}")
                    sent = False
                    if alimtalk_config:
                        sent = send_alimtalk(vendor_info, order_count, alimtalk_config)
                    _alimtalk_results[vendor_name] = sent
                    _alimtalk_logs.append({
                        'name': vendor_name,
                        'phone': vendor_info.get('phone', ''),
                        'phones': vendor_info.get('phones', []),
                        'count': order_count,
                        'sheet_url': sheet_url,
                        'sent': sent,
                    })
                    time.sleep(0.3)

                st.session_state['_upload_results'] = _upload_results
                st.session_state['_alimtalk_results'] = _alimtalk_results
                st.session_state['alimtalk_logs'] = _alimtalk_logs
                st.session_state['alimtalk_date'] = datetime.now().strftime('%Y년 %m월 %d일')

                # 세션 상태 업데이트
                st.session_state['total_orders'] = len(merged_df)
                st.session_state['vendor_data'] = vendor_data
                st.session_state['pending_count'] = len(merged_df)
                st.session_state['invoice_count'] = 0

                # 업로드 기록 저장
                log_entry = {
                    'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'files': [f.name for f in uploaded_files],
                    'total_orders': len(merged_df),
                    'vendors': []
                }
                for _al in _alimtalk_logs:
                    log_entry['vendors'].append({
                        'name': _al['name'],
                        'phone': _al['phone'],
                        'phones': _al.get('phones', []),
                        'orders': _al['count'],
                        'sheet_uploaded': True,
                        'alimtalk_sent': _al.get('sent', False),
                        'sheet_url': _al['sheet_url']
                    })
                save_upload_log(log_entry)

                # 발주 완료 파일 키 저장 (뱃지 표시용)
                st.session_state['_processed_file_key'] = _current_file_key

                # 성공 결과 저장 후 rerun (카드 뱃지 표시를 위해)
                alimtalk_sent_count = sum(1 for _al in _alimtalk_logs if _al.get('sent'))
                st.session_state['_last_success'] = {
                    'success_count': success_count,
                    'total_orders': len(merged_df),
                    'alimtalk_sent': alimtalk_sent_count,
                }
                st.rerun()

        # 알림톡 발송 내역 (메시지 미리보기만 유지)
        if st.session_state.get('alimtalk_logs'):
            _date = st.session_state.get('alimtalk_date', '')
            _sent_logs = [_al for _al in st.session_state['alimtalk_logs'] if _al.get('sent')]
            if _sent_logs:
                with st.expander(f"발송된 알림톡 미리보기 ({len(_sent_logs)}건)"):
                    for _al in _sent_logs:
                        st.markdown(f"""
                        <div class="kakao-msg">
                            안녕하세요, {_al['name']} 사장님.<br>
                            {_date} 신규 주문 <strong>{_al['count']}건</strong>이 등록되었습니다.<br><br>
                            아래 링크에서 주문 내역 확인 후,<br>
                            '송장번호, 택배사' 칸에 입력 부탁드립니다.<br><br>
                            ※ 발송 완료 후 별도 연락은 필요 없습니다.<br><br>
                            ▶ <a href="{_al['sheet_url']}" target="_blank">발주서 확인하기</a>
                        </div>
                        """, unsafe_allow_html=True)


    # 발주 업로드 기록
    st.markdown("---")
    st.markdown('<div class="section-title">업로드 기록</div>', unsafe_allow_html=True)

    upload_history = load_upload_history()
    if upload_history:
        # 날짜별 그룹핑
        today_str = datetime.now().strftime('%Y-%m-%d')
        yesterday_str = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

        groups = {'오늘': [], '어제': [], '이번 주': [], '이전': []}
        for log in upload_history:
            log_date = log['date'][:10]  # 'YYYY-MM-DD'
            if log_date == today_str:
                groups['오늘'].append(log)
            elif log_date == yesterday_str:
                groups['어제'].append(log)
            elif log_date >= week_ago:
                groups['이번 주'].append(log)
            else:
                groups['이전'].append(log)

        def _render_log(log):
            time_str = log['date'][11:]  # 'HH:MM'
            files = ', '.join(log['files'])
            vendors_html = ''
            for v in log.get('vendors', []):
                sheet_url = v.get('sheet_url', '')
                talk_icon = '✓' if v.get('alimtalk_sent') else '—'
                link = f'<a href="{sheet_url}" target="_blank" style="color:#2E643C;text-decoration:none;">시트</a>' if sheet_url else ''
                vendors_html += f'<span style="color:#888;font-size:0.82rem;">{v["name"]} {v["orders"]}건 · 알림톡 {talk_icon} {link}</span><br/>'
            st.markdown(f"""<div style="background:#F5F5F5;border-radius:12px;padding:0.8rem 1rem;margin-bottom:6px;">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <span style="font-weight:600;font-size:0.9rem;">{files}</span>
                    <span style="color:#999;font-size:0.8rem;">{time_str} · {log['total_orders']}건</span>
                </div>
                <div style="margin-top:4px;">{vendors_html}</div>
            </div>""", unsafe_allow_html=True)

        for label, logs in groups.items():
            if not logs:
                continue
            with st.expander(f"{label} ({len(logs)}건)", expanded=False):
                for log in logs:
                    _render_log(log)
    else:
        st.markdown('<div style="color:#999;font-size:0.88rem;padding:1rem 0;">아직 업로드 기록이 없습니다.</div>', unsafe_allow_html=True)


# ===== 송장 현황 =====
elif page == "송장 현황":
    st.markdown('<div class="main-header">송장 현황</div>', unsafe_allow_html=True)

    # 5분마다 대시보드 탭 자동 읽기 (Apps Script 5분 트리거가 백그라운드에서 갱신)
    refresh_count = st_autorefresh(interval=300000, limit=None, key="invoice_autorefresh")

    config = Config()
    master_url = config.get_vendor_master_url()
    sheet_client = get_sheet_client()
    vendors_info = load_vendors()

    # Apps Script URL (새로고침 버튼용)
    _apps_url = ""
    try:
        if "vendor_master" in st.secrets:
            _apps_url = st.secrets["vendor_master"].get("apps_script_url", "")
    except Exception:
        pass

    # 알림 로그 (최상단)
    notification_area = st.container()

    col_refresh, col_status = st.columns([1, 3])
    with col_refresh:
        if st.button("새로고침"):
            # 모든 캐시 클리어
            st.session_state['_dashboard_cache_time'] = 0
            st.session_state['_dashboard_cache'] = {}
            if sheet_client:
                if hasattr(sheet_client, '_spreadsheet_cache'):
                    sheet_client._spreadsheet_cache.clear()
                if hasattr(sheet_client, '_worksheet_cache'):
                    sheet_client._worksheet_cache.clear()
            if _apps_url:
                try:
                    with st.spinner("업체 시트 읽는 중... (약 25초)"):
                        _resp = _requests.get(_apps_url, params={"action": "refresh"}, timeout=90)
                        if _resp.status_code == 200:
                            logging.info(f'✅ Apps Script 응답: {_resp.text[:200]}')
                            st.toast("대시보드 갱신 완료!")
                        else:
                            logging.warning(f'⚠️ Apps Script 응답 코드: {_resp.status_code}')
                            st.warning(f"갱신 응답: {_resp.status_code}")
                except Exception as e:
                    logging.error(f'❌ Apps Script 호출 실패: {e}')
                    st.warning(f"Apps Script 호출 실패: {e}")
            # 캐시 한번 더 확실히 클리어 (Apps Script 완료 후)
            st.session_state['_dashboard_cache_time'] = 0
            st.session_state['_dashboard_cache'] = {}
            GoogleSheetClient._spreadsheet_cache.clear()
            GoogleSheetClient._worksheet_cache.clear()
            st.rerun()
    with col_status:
        st.caption(f"자동 새로고침 (1분)  |  {datetime.now().strftime('%H:%M:%S')}")

    if sheet_client and master_url:
        dashboard = fetch_dashboard(sheet_client, master_url)

        if not dashboard:
            st.info("대시보드 데이터가 아직 없습니다. Apps Script에서 updateDashboard를 실행해주세요.")
        else:
            # URL 매핑
            vendor_urls = {v['name']: v.get('google_sheet_url', '') for v in vendors_info} if vendors_info else {}

            # 업체 기준 집계
            _active_vendors = {n: d for n, d in dashboard.items() if d['total'] > 0}
            _total_vendors = len(_active_vendors)
            _done_vendors = sum(1 for d in _active_vendors.values() if d['invoiced'] == d['total'] and d['total'] > 0)
            _in_progress_vendors = sum(1 for d in _active_vendors.values() if 0 < d['invoiced'] < d['total'])
            _no_input_vendors = sum(1 for d in _active_vendors.values() if d['invoiced'] == 0)
            total_orders = sum(d['total'] for d in dashboard.values())
            total_invoices = sum(d['invoiced'] for d in dashboard.values())

            # 변경 감지
            prev_invoices = st.session_state.get('prev_invoice_count', None)
            prev_vendor_status = st.session_state.get('prev_vendor_status', {})

            if prev_invoices is not None and total_invoices > prev_invoices:
                for name, d in dashboard.items():
                    prev_count = prev_vendor_status.get(name, 0)
                    if d['invoiced'] > prev_count:
                        st.toast(f"{name} 송장 +{d['invoiced'] - prev_count}건")

            st.session_state['prev_invoice_count'] = total_invoices
            st.session_state['prev_vendor_status'] = {name: d['invoiced'] for name, d in dashboard.items()}

            # 전체 요약 (업체 기준 - 4단계)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("전체 업체", f"{_total_vendors}개")
            with col2:
                st.metric("입력 완료", f"{_done_vendors}개")
            with col3:
                st.metric("입력 중", f"{_in_progress_vendors}개")
            with col4:
                st.metric("미입력", f"{_no_input_vendors}개")

            if _total_vendors > 0:
                _started = _done_vendors + _in_progress_vendors
                _pct = int(_started / _total_vendors * 100)
                st.markdown(f"""
                <div style="margin:1rem 0;">
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;font-size:0.85rem;">
                        <span style="color:#64748B;font-weight:500;">진행률 (입력 시작 업체)</span>
                        <span style="color:#0F172A;font-weight:600;">{_started}/{_total_vendors} ({_pct}%)</span>
                    </div>
                    <div style="background:#E2E8F0;border-radius:8px;height:10px;overflow:hidden;">
                        <div style="background:#2E643C;width:{_pct}%;height:100%;border-radius:8px;transition:width 0.4s ease;"></div>
                    </div>
                </div>""", unsafe_allow_html=True)

            # 업체별 상세
            st.markdown("---")
            st.markdown('<div class="section-title">업체별 상세</div>', unsafe_allow_html=True)

            for name, d in dashboard.items():
                if d['total'] == 0:
                    continue
                prev_count = prev_vendor_status.get(name, 0)
                is_new = d['invoiced'] > prev_count
                is_done = d['invoiced'] == d['total'] and d['total'] > 0
                url = vendor_urls.get(name, '')

                new_badge = ' <span style="color:#2E643C;font-size:0.75rem;font-weight:600;">NEW</span>' if is_new else ""

                if is_done:
                    status_badge = '<span style="background:#E8F5E9;color:#2E643C;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">완료</span>'
                    active_cls = "list-row-active"
                elif d['invoiced'] > 0:
                    status_badge = f'<span style="background:#FFF8E1;color:#B8860B;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">{d["invoiced"]}/{d["total"]}</span>'
                    active_cls = ""
                else:
                    status_badge = '<span style="background:#F5F5F5;color:#999;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">대기</span>'
                    active_cls = ""

                st.markdown(f"""
                <div class="list-row {active_cls}">
                    <div style="flex:1;">
                        <div class="list-name">{name}{new_badge}</div>
                        <div class="list-desc">{d['total']}건</div>
                    </div>
                    <div style="display:flex;align-items:center;gap:10px;">
                        {status_badge}
                        <a href="{url}" target="_blank" style="text-decoration:none;">
                            <div class="list-arrow">→</div>
                        </a>
                    </div>
                </div>""", unsafe_allow_html=True)

            # 알림 로그 데이터 수집
            if 'notification_log' not in st.session_state:
                st.session_state['notification_log'] = []

            if prev_invoices is not None and total_invoices > prev_invoices:
                for name, d in dashboard.items():
                    prev_count = prev_vendor_status.get(name, 0)
                    if d['invoiced'] > prev_count:
                        _is_complete = d['invoiced'] == d['total']
                        # 같은 업체의 이전 로그 제거 (최신 상태만 유지)
                        st.session_state['notification_log'] = [
                            l for l in st.session_state['notification_log'] if l['vendor'] != name
                        ]
                        st.session_state['notification_log'].insert(0, {
                            'time': datetime.now().strftime('%H:%M:%S'),
                            'vendor': name,
                            'invoiced': d['invoiced'],
                            'total_orders': d['total'],
                            'complete': _is_complete,
                        })

            # 알림 로그 상단에 렌더링 (컴팩트 스크롤 방식)
            with notification_area:
                _raw_logs = st.session_state.get('notification_log', [])
                # 업체별 최신 로그만 유지 (중복 제거)
                _seen_vendors = set()
                _logs = []
                for l in _raw_logs:
                    if l['vendor'] not in _seen_vendors:
                        _seen_vendors.add(l['vendor'])
                        _logs.append(l)
                st.session_state['notification_log'] = _logs
                if _logs:
                    _done_count = sum(1 for l in _logs if l.get('complete'))
                    _prog_count = sum(1 for l in _logs if not l.get('complete'))
                    _summary_parts = []
                    if _done_count:
                        _summary_parts.append(f'<span style="background:#2E643C;color:white;padding:2px 10px;border-radius:12px;font-size:0.8rem;">입력 완료 {_done_count}건</span>')
                    if _prog_count:
                        _summary_parts.append(f'<span style="background:#B8860B;color:white;padding:2px 10px;border-radius:12px;font-size:0.8rem;">입력 중 {_prog_count}건</span>')

                    _log_items = ""
                    for log in _logs[:20]:
                        _inv = log.get('invoiced', 0)
                        _tot = log.get('total_orders', 0)
                        _is_done = log.get('complete', False)
                        _dot_color = "#2E643C" if _is_done else "#B8860B"
                        _status = "완료" if _is_done else "입력 중"
                        _log_items += f'<div style="padding:3px 0;font-size:0.78rem;color:#555;border-bottom:1px solid #f0f0f0;"><span style="color:{_dot_color};font-weight:bold;">●</span> {log["time"]} <strong>{log["vendor"]}</strong> {_status} ({_inv}/{_tot}건)</div>'

                    st.markdown(f"""
                    <div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:10px;padding:10px 14px;margin-bottom:10px;">
                        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                            <span style="font-size:0.85rem;font-weight:600;color:#333;">송장 입력 알림</span>
                            {' '.join(_summary_parts)}
                        </div>
                        <div style="max-height:120px;overflow-y:auto;scrollbar-width:thin;">
                            {_log_items}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                elif total_invoices > 0:
                    _init_items = ""
                    for name, d in dashboard.items():
                        if d['invoiced'] > 0:
                            _init_items += f'<div style="padding:3px 0;font-size:0.78rem;color:#555;border-bottom:1px solid #f0f0f0;"><span style="color:#2E643C;font-weight:bold;">●</span> <strong>{name}</strong> 송장 {d["invoiced"]}건 입력 완료</div>'
                    if _init_items:
                        st.markdown(f"""
                        <div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:10px;padding:10px 14px;margin-bottom:10px;">
                            <div style="font-size:0.85rem;font-weight:600;color:#333;margin-bottom:6px;">송장 입력 알림</div>
                            <div style="max-height:120px;overflow-y:auto;scrollbar-width:thin;">
                                {_init_items}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.info("아직 입력된 송장이 없습니다. 업체에서 입력하면 여기에 알림이 표시됩니다.")

    else:
        st.warning("구글 시트 연결이 필요합니다")


# ===== 송장 다운로드 =====
elif page == "송장 다운로드":
    st.markdown('<div class="main-header">송장 다운로드</div>', unsafe_allow_html=True)

    vendors_info = load_vendors()
    sheet_client = get_sheet_client()

    if st.button("송장 수집", type="primary", use_container_width=True):
        if not sheet_client:
            st.error("구글 시트 연결이 필요합니다")
        else:
            progress = st.progress(0, text="송장 수집 중...")

            all_invoices = []

            # 대시보드에서 송장 있는 업체만 수집 (API 최소화)
            _dl_config = Config()
            _dl_master_url = _dl_config.get_vendor_master_url()
            _dl_dashboard = fetch_dashboard(sheet_client, _dl_master_url)
            if _dl_dashboard:
                _invoiced_names = set(n for n, d in _dl_dashboard.items() if d['invoiced'] > 0)
            else:
                _invoiced_names = None
            target_vendors = [v for v in vendors_info if v['name'] in _invoiced_names] if _invoiced_names else []

            if not target_vendors:
                st.info("송장이 입력된 업체가 없습니다.")

            for idx, vendor_info in enumerate(target_vendors):
                vendor_name = vendor_info['name']
                sheet_url = vendor_info.get('google_sheet_url', '')
                progress.progress((idx + 1) / len(target_vendors), text=f"수집 중: {vendor_name}")

                if not sheet_url:
                    continue

                data = sheet_client.read_sheet(sheet_url)
                time.sleep(1)  # API rate limit 대응
                if not data or len(data) < 2:
                    continue

                df = pd.DataFrame(data[1:], columns=data[0])

                # 송장번호가 입력된 행만 필터
                df_with_invoice = df[
                    df.get('송장번호', pd.Series()).notna() &
                    (df.get('송장번호', pd.Series()) != '')
                ]

                if len(df_with_invoice) > 0:
                    df_with_invoice = df_with_invoice.copy()
                    df_with_invoice['공급처'] = vendor_name
                    all_invoices.append(df_with_invoice)
                    st.success(f"{vendor_name} — {len(df_with_invoice)}건 수집")
                else:
                    st.info(f"{vendor_name} — 입력된 송장 없음")

                if idx > 0 and idx % 10 == 0:
                    time.sleep(2)

            progress.progress(1.0, text="수집 완료!")

            if all_invoices:
                combined = pd.concat(all_invoices, ignore_index=True)
                st.session_state['_download_combined'] = combined

            if not all_invoices:
                st.warning("수집된 송장이 없습니다. 업체에서 아직 입력하지 않았어요.")

    # 수집된 데이터가 있으면 다운로드 UI 표시
    combined = st.session_state.get('_download_combined')
    if combined is not None and len(combined) > 0:
        today = datetime.now().strftime('%Y%m%d')
        upload_columns = ['주문일자', '주문번호', '수취인명', '연락처', '주소',
                          '상품명', '옵션', '수량', '택배사', '송장번호']
        available_cols = [c for c in upload_columns if c in combined.columns]

        st.markdown("---")
        st.markdown(f'<div class="section-title">수집 결과: {len(combined)}건</div>', unsafe_allow_html=True)

        # 택배사명 정규화 (띄어쓰기, 오타 통일)
        COURIER_ALIASES = {
            '로젠': '로젠택배', '로젠 택배': '로젠택배',
            'CJ': 'CJ대한통운', 'CJ 대한통운': 'CJ대한통운', 'cj대한통운': 'CJ대한통운',
            'cj 대한통운': 'CJ대한통운', 'CJ택배': 'CJ대한통운', 'CJ 택배': 'CJ대한통운',
            '대한통운': 'CJ대한통운',
            '한진': '한진택배', '한진 택배': '한진택배',
            '우체국': '우체국택배', '우체국 택배': '우체국택배', '우편': '우체국택배',
            '롯데': '롯데택배', '롯데 택배': '롯데택배',
            '경동': '경동택배', '경동 택배': '경동택배',
            '합동': '합동택배', '합동 택배': '합동택배',
            '일양': '일양로지스', '일양 로지스': '일양로지스',
            '건영': '건영택배', '건영 택배': '건영택배',
            '천일': '천일택배', '천일 택배': '천일택배',
            '호남': '호남택배', '호남 택배': '호남택배',
        }
        if '택배사' in combined.columns:
            combined['택배사'] = combined['택배사'].apply(
                lambda x: COURIER_ALIASES.get(str(x).strip(), str(x).strip()) if pd.notna(x) and str(x).strip() else x
            )

        # 택배사별 현황 카드
        courier_groups = {}
        if '택배사' in combined.columns:
            for courier, group in combined.groupby('택배사'):
                if courier and str(courier).strip():
                    courier_groups[str(courier).strip()] = group

        if courier_groups:
            courier_cols = st.columns(len(courier_groups))
            for i, (courier, group) in enumerate(courier_groups.items()):
                with courier_cols[i]:
                    st.metric(courier, f"{len(group)}건")

        # 미리보기
        with st.expander("데이터 미리보기"):
            st.dataframe(combined[available_cols] if available_cols else combined, use_container_width=True)

        st.markdown("---")
        st.markdown('<div class="section-title">다운로드</div>', unsafe_allow_html=True)

        download_mode = st.radio(
            "다운로드 방식",
            ["전체 통합 (1개 파일)", "택배사별 개별 파일"],
            horizontal=True,
            label_visibility="visible"
        )

        if download_mode == "전체 통합 (1개 파일)":
            buffer = BytesIO()
            combined[available_cols].to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            st.download_button(
                label=f"전체 다운로드 (송장일괄등록_{today}.xlsx)",
                data=buffer,
                file_name=f"송장일괄등록_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        elif download_mode == "택배사별 개별 파일":
            if not courier_groups:
                st.warning("택배사 정보가 없어요.")
            else:
                courier_summary = ", ".join([f"{c} {len(g)}건" for c, g in courier_groups.items()])
                st.caption(f"택배사별로 개별 엑셀 파일을 다운로드해요. ({courier_summary})")

                # 다운로드 이력 관리
                if '_downloaded_couriers' not in st.session_state:
                    st.session_state['_downloaded_couriers'] = set()

                # 미다운로드 항목 전체 선택 버튼
                _not_downloaded = [c for c in courier_groups if c not in st.session_state['_downloaded_couriers']]
                if _not_downloaded:
                    if st.checkbox(f"미다운로드 전체 선택 ({len(_not_downloaded)}건)", value=True, key="chk_select_all_new"):
                        st.session_state['_select_all_new'] = True
                    else:
                        st.session_state['_select_all_new'] = False
                else:
                    st.session_state['_select_all_new'] = False

                # 체크박스 + 다운로드 버튼
                _selected_couriers = []
                for courier, group in courier_groups.items():
                    _is_downloaded = courier in st.session_state['_downloaded_couriers']
                    _dl_label = f"  {courier} — {len(group)}건"
                    if _is_downloaded:
                        _dl_label += "  (다운로드 완료)"

                    col_chk, col_info, col_btn = st.columns([0.5, 5, 2])
                    with col_chk:
                        _default = (not _is_downloaded) and st.session_state.get('_select_all_new', True)
                        _checked = st.checkbox("", key=f"chk_{courier}", value=_default, label_visibility="collapsed")
                        if _checked:
                            _selected_couriers.append(courier)
                    with col_info:
                        if _is_downloaded:
                            st.markdown(f'**{courier}** — {len(group)}건 <span style="background:#E8F5E9;color:#2E643C;padding:2px 8px;border-radius:10px;font-size:0.75rem;font-weight:600;">다운로드 완료</span>', unsafe_allow_html=True)
                        else:
                            st.markdown(f'**{courier}** — {len(group)}건')
                    with col_btn:
                        buffer = BytesIO()
                        group[available_cols].to_excel(buffer, index=False, engine='openpyxl')
                        buffer.seek(0)
                        st.download_button(
                            label="다운로드",
                            data=buffer,
                            file_name=f"송장_{today}_{courier}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{courier}",
                            on_click=lambda c=courier: st.session_state['_downloaded_couriers'].add(c)
                        )

                # 선택 다운로드 (ZIP)
                if len(courier_groups) > 1:
                    st.markdown("---")
                    if _selected_couriers:
                        import zipfile
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for c in _selected_couriers:
                                g = courier_groups[c]
                                buf = BytesIO()
                                g[available_cols].to_excel(buf, index=False, engine='openpyxl')
                                zf.writestr(f"송장_{today}_{c}.xlsx", buf.getvalue())
                        zip_buffer.seek(0)

                        def _mark_selected_downloaded():
                            for c in _selected_couriers:
                                st.session_state['_downloaded_couriers'].add(c)

                        st.download_button(
                            label=f"선택한 {len(_selected_couriers)}개 파일 일괄 다운로드 (ZIP)",
                            data=zip_buffer,
                            file_name=f"송장_{today}_선택.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True,
                            on_click=_mark_selected_downloaded
                        )
                    else:
                        st.info("다운로드할 택배사를 선택해주세요.")


# ===== 업체 관리 페이지 =====
elif page == "업체 관리":
    st.markdown('<div class="main-header">업체 관리</div>', unsafe_allow_html=True)

    config = Config()
    master_url = config.get_vendor_master_url()
    sheet_client = get_sheet_client()

    if not master_url:
        st.warning("업체 마스터 시트가 설정되지 않았습니다. Streamlit Secrets에 [vendor_master] 섹션을 추가해주세요.")
    elif not sheet_client:
        st.error("구글 시트 연결에 실패했습니다.")
    else:
        vm = VendorManager(sheet_client, master_url, config.get_shared_folder_id())

        # --- 새 업체 등록 ---
        st.markdown('<div class="section-title">새 업체 등록</div>', unsafe_allow_html=True)
        with st.form("add_vendor_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                new_name = st.text_input("업체명 *")
            with col2:
                new_contact = st.text_input("담당자")
            col3, col4 = st.columns(2)
            with col3:
                new_email = st.text_input("이메일 (선택)")
            with col4:
                new_phone = st.text_input("전화번호1 * (알림톡 수신)")
            pc1, pc2 = st.columns(2)
            with pc1:
                new_phone2 = st.text_input("전화번호2 (선택)")
            with pc2:
                new_phone3 = st.text_input("전화번호3 (선택)")
            new_sheet_url = st.text_input("구글 시트 URL (비워두면 자동 생성)")
            submitted = st.form_submit_button("업체 등록", type="primary", use_container_width=True)
            if submitted:
                if not new_name or not new_phone:
                    st.error("업체명과 전화번호1은 필수입니다.")
                else:
                    with st.spinner(f"{new_name} 등록 중..."):
                        result = vm.add_vendor(
                            new_name, new_contact or '', new_phone,
                            new_email or '', sheet_url=new_sheet_url or None,
                            phone2=new_phone2 or '', phone3=new_phone3 or ''
                        )
                    if result:
                        if result.get('google_sheet_url'):
                            st.success(f"{new_name} 등록 완료!")
                        else:
                            st.warning(f"{new_name} 등록 완료! (시트 자동 생성 실패 — 수정에서 URL을 직접 입력해주세요)")
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("업체 등록에 실패했습니다.")

        st.markdown("---")

        # --- 등록된 업체 목록 ---
        st.markdown('<div class="section-title">등록된 업체</div>', unsafe_allow_html=True)
        all_vendors = vm.load_all_vendors()

        if not all_vendors:
            st.info("등록된 업체가 없습니다. 위에서 새 업체를 등록해주세요.")
        else:
            active = [v for v in all_vendors if v.get('상태') != '비활성']
            inactive = [v for v in all_vendors if v.get('상태') == '비활성']

            st.caption(f"총 {len(active)}개 업체 활성 / {len(inactive)}개 비활성")

            for v in active:
                vid = v.get('업체ID', '')
                vname = v.get('업체명', '')
                vphone = v.get('전화번호', '')
                vphone2 = v.get('전화번호2', '')
                vphone3 = v.get('전화번호3', '')
                vcontact = v.get('담당자', '')
                vemail = v.get('이메일', '')
                vurl = v.get('구글시트URL', '')
                vdate = v.get('등록일', '')
                # 표시용 전화번호 (등록된 번호 모두 표시)
                vphones_display = ', '.join([p for p in [vphone, vphone2, vphone3] if p])

                edit_key = f"editing_{vid}"
                is_editing = st.session_state.get(edit_key, False)

                if is_editing:
                    with st.form(f"edit_form_{vid}"):
                        st.markdown(f"**{vname}** 수정")
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            ed_name = st.text_input("업체명", value=vname, key=f"ed_name_{vid}")
                            ed_contact = st.text_input("담당자", value=vcontact, key=f"ed_contact_{vid}")
                        with ec2:
                            ed_phone = st.text_input("전화번호1", value=vphone, key=f"ed_phone_{vid}")
                            ed_email = st.text_input("이메일", value=vemail, key=f"ed_email_{vid}")
                        epc1, epc2 = st.columns(2)
                        with epc1:
                            ed_phone2 = st.text_input("전화번호2", value=vphone2, key=f"ed_phone2_{vid}")
                        with epc2:
                            ed_phone3 = st.text_input("전화번호3", value=vphone3, key=f"ed_phone3_{vid}")
                        bc1, bc2 = st.columns(2)
                        with bc1:
                            save_btn = st.form_submit_button("저장", type="primary", use_container_width=True)
                        with bc2:
                            cancel_btn = st.form_submit_button("취소", use_container_width=True)
                        if save_btn:
                            vm.update_vendor(vid, **{
                                '업체명': ed_name, '담당자': ed_contact,
                                '전화번호': ed_phone, '전화번호2': ed_phone2,
                                '전화번호3': ed_phone3, '이메일': ed_email
                            })
                            st.session_state[edit_key] = False
                            st.cache_data.clear()
                            st.rerun()
                        if cancel_btn:
                            st.session_state[edit_key] = False
                            st.rerun()
                else:
                    col_info, col_edit, col_del = st.columns([6, 1, 1])
                    with col_info:
                        sheet_link = f' · <a href="{vurl}" target="_blank" style="color:#2E643C;">시트 열기</a>' if vurl else ''
                        st.markdown(f"""
                        <div class="list-row">
                            <div style="flex:1;">
                                <div class="list-name">{vname}</div>
                                <div class="list-desc">{vphones_display} · {vcontact}{sheet_link}</div>
                            </div>
                        </div>""", unsafe_allow_html=True)
                    with col_edit:
                        if st.button("수정", key=f"btn_edit_{vid}", use_container_width=True):
                            st.session_state[edit_key] = True
                            st.rerun()
                    with col_del:
                        if st.button("삭제", key=f"btn_del_{vid}", use_container_width=True):
                            vm.delete_vendor(vid)
                            st.cache_data.clear()
                            st.rerun()

            # 비활성 업체
            if inactive:
                with st.expander(f"비활성 업체 ({len(inactive)}개)"):
                    for v in inactive:
                        vid = v.get('업체ID', '')
                        vname = v.get('업체명', '')
                        vphones_inactive = ', '.join([p for p in [v.get('전화번호', ''), v.get('전화번호2', ''), v.get('전화번호3', '')] if p])
                        col_info, col_restore = st.columns([6, 1])
                        with col_info:
                            st.markdown(f"""
                            <div class="list-row" style="opacity:0.5;">
                                <div style="flex:1;">
                                    <div class="list-name">{vname}</div>
                                    <div class="list-desc">{vphones_inactive} · 비활성</div>
                                </div>
                            </div>""", unsafe_allow_html=True)
                        with col_restore:
                            if st.button("복원", key=f"btn_restore_{vid}", use_container_width=True):
                                vm.restore_vendor(vid)
                                st.cache_data.clear()
                                st.rerun()

