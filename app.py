"""
발주도우미 - 데모 웹앱 (Streamlit)
"""
import os
import json
import time
import base64
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
        background-color: #F0F7F2 !important;
        border: 1px solid #D4E6D9 !important;
        border-radius: 10px !important;
    }
    [data-testid="stTextInput"] input:focus {
        border-color: #2E643C !important;
        box-shadow: 0 0 0 1px #2E643C !important;
    }

    /* ===== 프로그레스바 ===== */
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


def fetch_all_vendor_sheets(_client, vendors):
    """모든 업체 시트 데이터를 가져오기 (session_state 캐시, 20초 TTL)"""
    now = time.time()
    cache = st.session_state.get('_sheet_cache', {})
    cache_time = st.session_state.get('_sheet_cache_time', 0)

    # 20초 이내면 캐시 반환
    if cache and (now - cache_time) < 20:
        return cache

    result = {}
    if not _client or not vendors:
        return result
    for v in vendors:
        url = v.get('google_sheet_url', '')
        if not url:
            continue
        try:
            data = _client.read_sheet(url)
            result[v['name']] = data
        except Exception:
            result[v['name']] = None

    st.session_state['_sheet_cache'] = result
    st.session_state['_sheet_cache_time'] = now
    return result


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


# ===== 송장 다운로드 페이지에서 백그라운드 송장 변경 감지 (30초 간격) =====
# 발주 업로드 / 송장 현황 페이지에서는 autorefresh 끔
if page == "송장 다운로드":
    bg_refresh = st_autorefresh(interval=30000, limit=None, key="bg_autorefresh")

    # 백그라운드 송장 체크 (캐시 활용, 에러 시 무시)
    try:
        _bg_client = get_sheet_client()
        _bg_vendors = load_vendors()
        _all_sheets = fetch_all_vendor_sheets(_bg_client, _bg_vendors or [])
        if _all_sheets:
            _bg_total_invoices = 0
            _bg_vendor_status = {}
            for _v in (_bg_vendors or []):
                _data = _all_sheets.get(_v['name'])
                if _data and len(_data) >= 2:
                    _df = pd.DataFrame(_data[1:], columns=_data[0])
                    _cnt = len(_df[_df.get('송장번호', pd.Series()).notna() & (_df.get('송장번호', pd.Series()) != '')])
                    _bg_total_invoices += _cnt
                    _bg_vendor_status[_v['name']] = _cnt

            _prev = st.session_state.get('prev_invoice_count', None)
            _prev_vs = st.session_state.get('prev_vendor_status', {})

            if _prev is not None and _bg_total_invoices > _prev:
                for _vn, _vc in _bg_vendor_status.items():
                    _pc = _prev_vs.get(_vn, 0)
                    if _vc > _pc:
                        _diff = _vc - _pc
                        st.toast(f"{_vn} 송장 +{_diff}건")

            st.session_state['prev_invoice_count'] = _bg_total_invoices
            st.session_state['prev_vendor_status'] = _bg_vendor_status
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

        # [2] 공급처별 주문 현황
        vendor_data = split_by_vendor(merged_df)
        if vendor_data:
            with st.expander(f"공급처별 주문 현황 ({len(vendor_data)}개 업체)", expanded=False):
                grid_html = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px;">'
                for name, vdf in vendor_data.items():
                    grid_html += f'''<div style="background:#F5F5F5;border-radius:10px;padding:10px 12px;text-align:center;">
                        <div style="font-size:0.78rem;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{name}">{name}</div>
                        <div style="font-size:1.2rem;font-weight:700;color:#1A1A1A;">{len(vdf)}건</div>
                    </div>'''
                grid_html += '</div>'
                st.markdown(grid_html, unsafe_allow_html=True)

        # 발주 실행 버튼
        if st.button("발주 실행", type="primary", use_container_width=True):
            vendors_info = load_vendors()
            sheet_client = get_sheet_client()

            if not sheet_client:
                st.error("구글 시트 연결이 필요합니다")
            else:
                progress = st.progress(0, text="발주 처리 중...")
                status_container = st.container()

                total = len(vendors_info)
                success_count = 0

                for idx, vendor_info in enumerate(vendors_info):
                    vendor_name = vendor_info['name']
                    progress.progress((idx + 1) / total, text=f"처리 중: {vendor_name}")

                    if vendor_name not in vendor_data:
                        with status_container:
                            st.warning(f"{vendor_name} — 주문 없음")
                        continue

                    vendor_df = vendor_data[vendor_name]
                    sheet_data = prepare_sheet_data(vendor_df)

                    # 구글 시트 업로드
                    sheet_url = vendor_info.get('google_sheet_url', '')
                    if sheet_url:
                        result = sheet_client.update_sheet(sheet_url, sheet_data)
                        if result:
                            with status_container:
                                st.success(f"{vendor_name} — {len(vendor_df)}건 업로드 완료")
                            success_count += 1
                        else:
                            with status_container:
                                st.error(f"{vendor_name} — 업로드 실패")

                    time.sleep(0.3)

                progress.progress(1.0, text="시트 업로드 완료!")

                # 알림톡 발송 (실제 API 호출)
                alimtalk_config = Config().load_alimtalk_config()
                _alimtalk_logs = []
                for vendor_info in vendors_info:
                    vendor_name = vendor_info['name']
                    if vendor_name not in vendor_data:
                        continue
                    order_count = len(vendor_data[vendor_name])
                    sheet_url = vendor_info.get('google_sheet_url', '')
                    sent = False
                    if alimtalk_config:
                        sent = send_alimtalk(vendor_info, order_count, alimtalk_config)
                    _alimtalk_logs.append({
                        'name': vendor_name,
                        'phone': vendor_info.get('phone', ''),
                        'count': order_count,
                        'sheet_url': sheet_url,
                        'sent': sent,
                    })
                    time.sleep(0.3)

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
                        'orders': _al['count'],
                        'sheet_uploaded': True,
                        'alimtalk_sent': _al.get('sent', False),
                        'sheet_url': _al['sheet_url']
                    })
                save_upload_log(log_entry)

                # 성공 이펙트 (커스텀 로고 애니메이션)
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
                alimtalk_sent_count = sum(1 for _al in _alimtalk_logs if _al.get('sent'))
                if alimtalk_sent_count > 0:
                    st.success(f"발주 처리 완료! {success_count}개 업체 시트 업로드 + {alimtalk_sent_count}개 업체 알림톡 발송 완료")
                else:
                    st.success(f"발주 처리 완료! {success_count}개 업체에 총 {len(merged_df)}건 시트 업로드 완료 (알림톡 미설정)")

        # 알림톡 발송 내역
        if st.session_state.get('alimtalk_logs'):
            st.markdown("---")
            st.markdown('<div class="section-title">알림톡 발송 현황</div>', unsafe_allow_html=True)
            _date = st.session_state.get('alimtalk_date', '')
            for _al in st.session_state['alimtalk_logs']:
                _sent = _al.get('sent', False)
                _status = "발송 완료" if _sent else "미설정"
                _cls = "list-row-active" if _sent else ""
                st.markdown(f"""
                <div class="list-row {_cls}">
                    <div style="flex:1;">
                        <div class="list-name">{_al['name']}</div>
                        <div class="list-desc" style="{'color:rgba(255,255,255,0.6);' if _sent else ''}">{_al['phone']} · {_al['count']}건 · {_status}</div>
                    </div>
                    <a href="{_al['sheet_url']}" target="_blank" style="text-decoration:none;">
                        <div class="list-arrow">→</div>
                    </a>
                </div>""", unsafe_allow_html=True)
            with st.expander("발송된 메시지 미리보기"):
                for _al in st.session_state['alimtalk_logs']:
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

    # 15초마다 자동 새로고침
    refresh_count = st_autorefresh(interval=15000, limit=None, key="invoice_autorefresh")

    vendors_info = load_vendors()
    sheet_client = get_sheet_client()

    col_refresh, col_status = st.columns([1, 3])
    with col_refresh:
        if st.button("새로고침"):
            st.session_state['_sheet_cache_time'] = 0
    with col_status:
        st.caption(f"자동 새로고침 (15초)  |  {datetime.now().strftime('%H:%M:%S')}")

    if sheet_client and vendors_info:
        all_sheets = fetch_all_vendor_sheets(sheet_client, vendors_info)

        total_orders = 0
        total_invoices = 0
        vendor_statuses = []

        for vendor_info in vendors_info:
            vendor_name = vendor_info['name']
            sheet_url = vendor_info.get('google_sheet_url', '')
            if not sheet_url:
                continue

            data = all_sheets.get(vendor_name)
            if not data or len(data) < 2:
                vendor_statuses.append({
                    'name': vendor_name, 'total': 0, 'completed': 0, 'pending': 0, 'url': sheet_url
                })
                continue

            df = pd.DataFrame(data[1:], columns=data[0])
            order_count = len(df)
            invoice_count = len(df[df.get('송장번호', pd.Series()).notna() & (df.get('송장번호', pd.Series()) != '')])
            total_orders += order_count
            total_invoices += invoice_count

            vendor_statuses.append({
                'name': vendor_name,
                'total': order_count,
                'completed': invoice_count,
                'pending': order_count - invoice_count,
                'url': sheet_url
            })

        # 변경 감지
        prev_invoices = st.session_state.get('prev_invoice_count', None)
        prev_vendor_status = st.session_state.get('prev_vendor_status', {})

        if prev_invoices is not None and total_invoices > prev_invoices:
            for vs in vendor_statuses:
                prev_count = prev_vendor_status.get(vs['name'], 0)
                if vs['completed'] > prev_count:
                    diff = vs['completed'] - prev_count
                    st.toast(f"{vs['name']} 송장 +{diff}건")

        st.session_state['prev_invoice_count'] = total_invoices
        st.session_state['prev_vendor_status'] = {vs['name']: vs['completed'] for vs in vendor_statuses}
        st.session_state['invoice_count'] = total_invoices
        st.session_state['pending_count'] = total_orders - total_invoices

        # 전체 요약
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("전체 주문", f"{total_orders}건")
        with col2:
            st.metric("송장 입력 완료", f"{total_invoices}건",
                       delta=f"+{total_invoices - prev_invoices}건" if prev_invoices is not None and total_invoices > prev_invoices else None)
        with col3:
            st.metric("미입력", f"{total_orders - total_invoices}건")

        if total_orders > 0:
            _pct = int(total_invoices / total_orders * 100)
            st.markdown(f"""
            <div style="margin:1rem 0;">
                <div style="display:flex;justify-content:space-between;margin-bottom:6px;font-size:0.85rem;">
                    <span style="color:#64748B;font-weight:500;">진행률</span>
                    <span style="color:#0F172A;font-weight:600;">{total_invoices}/{total_orders} ({_pct}%)</span>
                </div>
                <div style="background:#E2E8F0;border-radius:8px;height:10px;overflow:hidden;">
                    <div style="background:#2E643C;width:{_pct}%;height:100%;border-radius:8px;transition:width 0.4s ease;"></div>
                </div>
            </div>""", unsafe_allow_html=True)

        # 전체 업체 상세 현황
        st.markdown("---")
        st.markdown('<div class="section-title">업체별 상세</div>', unsafe_allow_html=True)

        for vs in vendor_statuses:
            prev_count = prev_vendor_status.get(vs['name'], 0)
            is_new = vs['completed'] > prev_count
            is_done = vs['completed'] == vs['total'] and vs['total'] > 0

            new_badge = ' <span style="color:#2E643C;font-size:0.75rem;font-weight:600;">NEW</span>' if is_new else ""

            if is_done:
                status_badge = '<span style="background:#E8F5E9;color:#2E643C;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">완료</span>'
                active_cls = "list-row-active"
            elif vs['completed'] > 0:
                status_badge = f'<span style="background:#FFF8E1;color:#B8860B;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">{vs["completed"]}/{vs["total"]}</span>'
                active_cls = ""
            elif vs['total'] > 0:
                status_badge = '<span style="background:#F5F5F5;color:#999;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">대기</span>'
                active_cls = ""
            else:
                status_badge = '<span style="background:#F5F5F5;color:#CCC;padding:5px 14px;border-radius:20px;font-size:0.78rem;font-weight:600;">—</span>'
                active_cls = ""

            st.markdown(f"""
            <div class="list-row {active_cls}">
                <div style="flex:1;">
                    <div class="list-name">{vs['name']}{new_badge}</div>
                    <div class="list-desc">{vs['total']}건</div>
                </div>
                <div style="display:flex;align-items:center;gap:10px;">
                    {status_badge}
                    <a href="{vs['url']}" target="_blank" style="text-decoration:none;">
                        <div class="list-arrow">→</div>
                    </a>
                </div>
            </div>""", unsafe_allow_html=True)

        # 알림 로그
        st.markdown("---")
        st.markdown('<div class="section-title">알림 로그</div>', unsafe_allow_html=True)

        if 'notification_log' not in st.session_state:
            st.session_state['notification_log'] = []

        if prev_invoices is not None and total_invoices > prev_invoices:
            for vs in vendor_statuses:
                prev_count = prev_vendor_status.get(vs['name'], 0)
                if vs['completed'] > prev_count:
                    diff = vs['completed'] - prev_count
                    st.session_state['notification_log'].insert(0, {
                        'time': datetime.now().strftime('%H:%M:%S'),
                        'vendor': vs['name'],
                        'count': diff,
                        'total': vs['completed']
                    })

        if st.session_state.get('notification_log'):
            for log in st.session_state['notification_log'][:10]:
                st.markdown(f"""
                <div class="notification-bar">
                    [{log['time']}] <strong>{log['vendor']}</strong> 송장 +{log['count']}건 (총 {log['total']}건)
                </div>
                """, unsafe_allow_html=True)
        else:
            if total_invoices > 0:
                for vs in vendor_statuses:
                    if vs['completed'] > 0:
                        st.markdown(f"""
                        <div class="notification-bar">
                            {vs['name']} 송장 {vs['completed']}건 입력 완료
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

            for idx, vendor_info in enumerate(vendors_info):
                vendor_name = vendor_info['name']
                sheet_url = vendor_info.get('google_sheet_url', '')
                progress.progress((idx + 1) / len(vendors_info), text=f"수집 중: {vendor_name}")

                if not sheet_url:
                    continue

                data = sheet_client.read_sheet(sheet_url)
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
                for courier, group in courier_groups.items():
                    buffer = BytesIO()
                    group[available_cols].to_excel(buffer, index=False, engine='openpyxl')
                    buffer.seek(0)
                    st.download_button(
                        label=f"{courier} — {len(group)}건 다운로드",
                        data=buffer,
                        file_name=f"송장_{today}_{courier}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_{courier}"
                    )


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
                new_contact = st.text_input("담당자")
            with col2:
                new_phone = st.text_input("전화번호 * (알림톡 수신)")
                new_email = st.text_input("이메일 (선택)")
            new_sheet_url = st.text_input("구글 시트 URL (비워두면 자동 생성)")
            submitted = st.form_submit_button("업체 등록", type="primary", use_container_width=True)
            if submitted:
                if not new_name or not new_phone:
                    st.error("업체명과 전화번호는 필수입니다.")
                else:
                    with st.spinner(f"{new_name} 등록 중..."):
                        result = vm.add_vendor(
                            new_name, new_contact or '', new_phone,
                            new_email or '', sheet_url=new_sheet_url or None
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
                vcontact = v.get('담당자', '')
                vemail = v.get('이메일', '')
                vurl = v.get('구글시트URL', '')
                vdate = v.get('등록일', '')

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
                            ed_phone = st.text_input("전화번호", value=vphone, key=f"ed_phone_{vid}")
                            ed_email = st.text_input("이메일", value=vemail, key=f"ed_email_{vid}")
                        bc1, bc2 = st.columns(2)
                        with bc1:
                            save_btn = st.form_submit_button("저장", type="primary", use_container_width=True)
                        with bc2:
                            cancel_btn = st.form_submit_button("취소", use_container_width=True)
                        if save_btn:
                            vm.update_vendor(vid, **{
                                '업체명': ed_name, '담당자': ed_contact,
                                '전화번호': ed_phone, '이메일': ed_email
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
                                <div class="list-desc">{vphone} · {vcontact}{sheet_link}</div>
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
                        vphone = v.get('전화번호', '')
                        col_info, col_restore = st.columns([6, 1])
                        with col_info:
                            st.markdown(f"""
                            <div class="list-row" style="opacity:0.5;">
                                <div style="flex:1;">
                                    <div class="list-name">{vname}</div>
                                    <div class="list-desc">{vphone} · 비활성</div>
                                </div>
                            </div>""", unsafe_allow_html=True)
                        with col_restore:
                            if st.button("복원", key=f"btn_restore_{vid}", use_container_width=True):
                                vm.restore_vendor(vid)
                                st.cache_data.clear()
                                st.rerun()

