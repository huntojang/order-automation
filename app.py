"""
발주도우미 - 데모 웹앱 (Streamlit)
"""
import os
import json
import time
import base64
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh

from utils import Config, GoogleSheetClient, GoogleSheetOAuthClient, Logger

# 페이지 설정 (파비콘을 커스텀 이미지로)
from PIL import Image
_favicon = Image.open("assets/favicon_1.png")
st.set_page_config(
    page_title="발주도우미",
    page_icon=_favicon,
    layout="wide",
    initial_sidebar_state="expanded"
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

# CSS 스타일 — 메인
st.markdown("""
<style>
    /* ===== 글로벌 ===== */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .block-container { max-width: 1100px; padding-top: 2rem !important; }
    [data-testid="stAppViewBlockContainer"] { padding-top: 2rem !important; }
    /* autorefresh 등 hidden iframe이 높이 차지하지 않도록 */
    iframe[title="streamlit_autorefresh.st_autorefresh"] {
        position: absolute; height: 0 !important; overflow: hidden;
    }

    /* 로딩 스피너 */
    .stSpinner > div > div {
        border: 3px solid #E2E8F0 !important;
        border-top: 3px solid #4090C3 !important;
        border-radius: 50% !important;
        width: 22px !important; height: 22px !important;
        animation: spin 0.7s linear infinite !important;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    /* ===== 헤더 ===== */
    .main-header {
        font-size: 2rem; font-weight: 700; color: #0F172A;
        margin-bottom: 0.25rem; letter-spacing: -0.02em;
    }
    .sub-header { display: none; }
    .badge {
        display: inline-block; border: 1px solid #CBD5E1; color: #64748B;
        padding: 4px 14px; border-radius: 20px; font-size: 0.8rem;
        font-weight: 500; margin-bottom: 12px;
    }

    /* ===== 카드 ===== */
    .card {
        background: #F8FAFC; border-radius: 16px; padding: 1.5rem;
        border: 1px solid #E2E8F0; transition: all 0.2s ease;
    }
    .card:hover { border-color: #CBD5E1; box-shadow: 0 4px 12px rgba(0,0,0,0.04); }
    .card-blue {
        background: linear-gradient(135deg, #4090C3 0%, #5AACCF 100%);
        border-radius: 16px; padding: 1.5rem; color: white; border: none;
    }
    .card-blue .card-icon { background: rgba(255,255,255,0.2); }
    .card-icon {
        width: 42px; height: 42px; border-radius: 12px;
        background: #EBF4FA; display: flex; align-items: center;
        justify-content: center; font-size: 1.2rem; margin-bottom: 1rem;
        color: #4090C3;
    }
    .card-title { font-size: 1rem; font-weight: 600; margin-bottom: 4px; }
    .card-value { font-size: 2rem; font-weight: 700; margin-bottom: 2px; }
    .card-desc { font-size: 0.85rem; color: #94A3B8; }
    .card-blue .card-desc { color: rgba(255,255,255,0.7); }
    .card-blue .card-title { color: rgba(255,255,255,0.85); }
    .card-blue .card-value { color: #fff; }

    /* ===== 리스트 행 ===== */
    .list-row {
        display: flex; align-items: center; justify-content: space-between;
        background: #F8FAFC; border-radius: 12px; padding: 1rem 1.25rem;
        margin-bottom: 8px; border: 1px solid #E2E8F0;
        transition: all 0.2s ease;
    }
    .list-row:hover { background: #F1F5F9; }
    .list-row.list-row-static { cursor: default; }
    .list-row.list-row-static:hover { background: #F8FAFC; }
    .list-row.list-row-static.list-row-active:hover {
        background: linear-gradient(90deg, #4090C3, #5AACCF);
    }
    .list-row-active {
        background: linear-gradient(90deg, #4090C3, #5AACCF);
        color: white; border-color: transparent;
    }
    .list-row-active .list-desc { color: rgba(255,255,255,0.7); }
    .list-name { font-weight: 600; font-size: 0.95rem; }
    .list-desc { font-size: 0.85rem; color: #64748B; }
    .list-arrow {
        width: 36px; height: 36px; border-radius: 50%;
        background: #4090C3; color: white; display: flex;
        align-items: center; justify-content: center; font-size: 1rem;
        flex-shrink: 0;
    }
    .list-row-active .list-arrow { background: rgba(255,255,255,0.25); }

    /* ===== 카카오 메시지 ===== */
    .kakao-msg {
        background: #FEE500; color: #3C1E1E; padding: 1.25rem;
        border-radius: 16px; margin: 0.5rem 0; font-size: 0.9rem;
        max-width: 340px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }

    /* ===== 알림바 ===== */
    .notification-bar {
        background: linear-gradient(90deg, #4090C3, #5AACCF);
        color: white; padding: 0.8rem 1.2rem; border-radius: 10px;
        margin: 0.5rem 0; font-size: 0.9rem;
        animation: fadeIn 0.4s ease-out;
    }
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-8px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* ===== 사이드바 ===== */
    [data-testid="stSidebar"] {
        background-color: #FAFBFC; border-right: 1px solid #E2E8F0;
    }
    /* 사이드바 내부 구분선 제거 */
    [data-testid="stSidebar"] hr {
        display: none;
    }
    /* 라디오 버튼 → 네비게이션 버튼 스타일 */
    [data-testid="stSidebar"] .stRadio > div {
        gap: 4px !important;
    }
    [data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label {
        background: transparent;
        border: none;
        border-radius: 10px;
        padding: 10px 14px !important;
        margin: 0 !important;
        font-weight: 500;
        font-size: 0.92rem;
        color: #64748B;
        cursor: pointer;
        transition: all 0.15s ease;
        display: flex;
        align-items: center;
        gap: 4px;
    }
    [data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label:hover {
        background: #EBF4FA;
        color: #4090C3;
    }
    [data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label[data-checked="true"],
    [data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label:has(input:checked) {
        background: rgba(64,144,195,0.2);
        color: #4090C3 !important;
        font-weight: 600;
    }
    /* 라디오 원형 동그라미 숨기기 */
    [data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }
    /* 사이드바 알림 섹션 */
    .sidebar-section-title {
        font-size: 0.78rem;
        font-weight: 600;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        padding: 0 14px;
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
    }
    .sidebar-noti-item {
        background: #EBF4FA;
        border-left: 3px solid #4090C3;
        padding: 8px 12px;
        border-radius: 6px;
        margin: 0 8px 6px 8px;
        font-size: 0.78rem;
        color: #1E293B;
    }
    .sidebar-noti-item strong { color: #4090C3; }

    /* ===== 버튼 ===== */
    .stButton > button[kind="primary"] {
        background-color: #4090C3; border: none;
        border-radius: 10px; padding: 0.6rem 1.5rem;
        font-weight: 600; font-size: 0.95rem;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #357DA8; box-shadow: 0 4px 12px rgba(37,99,235,0.3);
    }
    .stButton > button {
        border-radius: 10px; font-weight: 500;
    }

    /* ===== 프로그레스바 ===== */
    .stProgress > div > div > div { background-color: #4090C3; border-radius: 8px; }
    .stProgress > div > div { background-color: #E2E8F0; border-radius: 8px; }

    /* ===== metric 카드 ===== */
    [data-testid="stMetric"] {
        background: #F8FAFC; border: 1px solid #E2E8F0;
        border-radius: 14px; padding: 1.2rem;
    }
    [data-testid="stMetricValue"] { color: #0F172A; font-weight: 700; }
    [data-testid="stMetricLabel"] { color: #64748B; font-weight: 500; }
    [data-testid="stMetricDelta"] svg { display: none; }

    /* ===== 파일 업로더 ===== */
    [data-testid="stFileUploader"] {
        border-radius: 14px;
    }
    [data-testid="stFileUploader"] > div > div {
        border-radius: 14px; border: 2px dashed #CBD5E1;
        background: #FAFBFC;
    }

    /* ===== expander ===== */
    .streamlit-expanderHeader {
        background: #F8FAFC; border-radius: 10px;
        font-weight: 500;
    }

    /* ===== 다운로드 버튼 ===== */
    .stDownloadButton > button {
        background-color: #4090C3; color: white; border: none;
        border-radius: 10px; font-weight: 600;
    }
    .stDownloadButton > button:hover {
        background-color: #357DA8;
    }
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


@st.cache_data(ttl=300)
def load_vendors():
    """업체 정보 로드 (5분 캐시)"""
    config = Config()
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
        '주문일자': ['주문일자'],
        '주문번호': ['주문번호'],
        '수취인명': ['수취인명', '수령자 이름'],
        '연락처': ['연락처', '수령자 휴대폰번호', '수령자 전화'],
        '주소': ['주소', '수령자 주소'],
        '상품명': ['상품명'],
        '옵션': ['옵션', '옵션명'],
        '수량': ['수량', '상품수량'],
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


# ===== 사이드바 =====
with st.sidebar:
    st.image("assets/logo.png", use_container_width=True)
    st.markdown("""
    <div style="padding:0 14px 0 14px;">
        <div style="font-size:0.78rem;color:#94A3B8;margin-top:-8px;margin-bottom:1rem;">위탁 발주 관리 시스템</div>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "메뉴",
        ["🏠 홈", "📤 발주 업로드", "📊 송장 현황", "📥 송장 다운로드"],
        label_visibility="collapsed"
    )

    st.markdown("""
    <div style="padding:0 14px;margin-top:auto;">
        <div style="font-size:0.72rem;color:#CBD5E1;margin-top:1.5rem;">v1.0 Demo · 갓샵</div>
    </div>
    """, unsafe_allow_html=True)


# ===== 홈/송장다운로드 페이지에서 송장 변경 감지 (30초 간격) =====
# 발주 업로드 페이지에서는 autorefresh 끔 (업로드 중 리셋 방지)
if page not in ["📊 송장 현황", "📤 발주 업로드"]:
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
                        st.toast(f"🔔 {_vn}에서 송장 {_diff}건 새로 입력!", icon="🔔")
                        if 'notification_log' not in st.session_state:
                            st.session_state['notification_log'] = []
                        st.session_state['notification_log'].insert(0, {
                            'time': datetime.now().strftime('%H:%M:%S'),
                            'vendor': _vn,
                            'count': _diff,
                            'total': _vc
                        })

            st.session_state['prev_invoice_count'] = _bg_total_invoices
            st.session_state['prev_vendor_status'] = _bg_vendor_status
    except Exception:
        pass


# ===== 홈 =====
if page == "🏠 홈":
    st.markdown('<span class="badge">DASHBOARD</span>', unsafe_allow_html=True)
    st.markdown('<div class="main-header">발주도우미</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">위탁 판매 발주 & 송장 관리 자동화 시스템</div>', unsafe_allow_html=True)

    _orders = st.session_state.get('total_orders', 0)
    _inv = st.session_state.get('invoice_count', 0)
    _pend = st.session_state.get('pending_count', 0)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="card-blue">
            <div class="card-icon" style="background:rgba(255,255,255,0.2);color:#fff;">📦</div>
            <div class="card-title">등록 업체</div>
            <div class="card-value">5개</div>
            <div class="card-desc">총 위탁 업체 수</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="card">
            <div class="card-icon">📤</div>
            <div class="card-title">오늘 발주</div>
            <div class="card-value" style="color:#0F172A;">{_orders}건</div>
            <div class="card-desc">발주 업로드 건수</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="card">
            <div class="card-icon">✅</div>
            <div class="card-title">송장 입력 완료</div>
            <div class="card-value" style="color:#0F172A;">{_inv}건</div>
            <div class="card-desc">업체 입력 완료</div>
        </div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="card">
            <div class="card-icon">⏳</div>
            <div class="card-title">미입력</div>
            <div class="card-value" style="color:#0F172A;">{_pend}건</div>
            <div class="card-desc">입력 대기 중</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<span class="badge">HOW IT WORKS</span>', unsafe_allow_html=True)
    st.markdown("#### 사용 방법")

    steps = [
        ("📤", "발주 업로드", "이지어드민 엑셀 파일을 업로드하면 공급처별로 자동 분류 & 알림 발송"),
        ("📊", "송장 현황", "각 업체의 송장번호 입력 현황을 실시간 모니터링"),
        ("📥", "송장 다운로드", "입력된 송장번호를 수집하여 이지어드민용 엑셀 다운로드"),
    ]
    for i, (icon, title, desc) in enumerate(steps):
        active = ""
        st.markdown(f"""
        <div class="list-row list-row-static {active}">
            <div>
                <div class="list-name">{icon} {title}</div>
                <div class="list-desc">{desc}</div>
            </div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<span class="badge">FLOW</span>', unsafe_allow_html=True)
    st.markdown("#### 시스템 흐름")

    flow_steps = [
        "이지어드민 엑셀 업로드",
        "공급처별 구글 시트 자동 분배",
        "각 업체에 카카오 알림톡 발송",
        "업체 사장님이 시트에서 송장번호 입력",
        "송장 수집 → 엑셀 다운로드 → 이지어드민 업로드",
    ]
    for i, step in enumerate(flow_steps):
        st.markdown(f"""
        <div class="list-row">
            <div>
                <div class="list-name"><span style="color:#4090C3; font-weight:700;">{i+1}.</span> {step}</div>
            </div>
        </div>""", unsafe_allow_html=True)


# ===== 발주 업로드 =====
elif page == "📤 발주 업로드":
    st.markdown('<div class="main-header">📤 발주 업로드</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">이지어드민 엑셀 파일을 업로드하면 자동으로 처리됩니다</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "이지어드민 발주서 엑셀/CSV 파일을 올려주세요",
        type=['xlsx', 'csv'],
        accept_multiple_files=True,
        help="여러 파일을 한번에 올릴 수 있습니다"
    )

    if uploaded_files:
        all_dfs = []
        for f in uploaded_files:
            if f.name.endswith('.csv'):
                df = pd.read_csv(f, encoding='utf-8')
            else:
                df = pd.read_excel(f, engine='openpyxl')
            all_dfs.append(df)
            st.success(f"✅ {f.name} — {len(df)}건 로드")

        merged_df = pd.concat(all_dfs, ignore_index=True)
        st.info(f"📊 총 주문 건수: **{len(merged_df)}건**")

        # 미리보기
        with st.expander("📋 데이터 미리보기", expanded=False):
            st.dataframe(merged_df.head(10), use_container_width=True)

        # 공급처별 분류
        vendor_data = split_by_vendor(merged_df)
        if vendor_data:
            st.markdown("### 📦 공급처별 주문 현황")
            cols = st.columns(len(vendor_data))
            for i, (name, vdf) in enumerate(vendor_data.items()):
                with cols[i % len(cols)]:
                    st.metric(name, f"{len(vdf)}건")

        st.markdown("---")

        # 발주 실행 버튼
        if st.button("🚀 발주 실행 (시트 업로드 + 알림 발송)", type="primary", use_container_width=True):
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
                            st.warning(f"⏭️ {vendor_name} — 주문 없음")
                        continue

                    vendor_df = vendor_data[vendor_name]
                    sheet_data = prepare_sheet_data(vendor_df)

                    # 구글 시트 업로드
                    sheet_url = vendor_info.get('google_sheet_url', '')
                    if sheet_url:
                        result = sheet_client.update_sheet(sheet_url, sheet_data)
                        if result:
                            with status_container:
                                st.success(f"✅ {vendor_name} — {len(vendor_df)}건 시트 업로드 완료")
                            success_count += 1
                        else:
                            with status_container:
                                st.error(f"❌ {vendor_name} — 시트 업로드 실패")

                    time.sleep(0.3)

                progress.progress(1.0, text="시트 업로드 완료!")

                # 알림톡 시뮬레이션 (session_state에 저장)
                _alimtalk_logs = []
                for vendor_info in vendors_info:
                    vendor_name = vendor_info['name']
                    if vendor_name not in vendor_data:
                        continue
                    order_count = len(vendor_data[vendor_name])
                    sheet_url = vendor_info.get('google_sheet_url', '')
                    _alimtalk_logs.append({
                        'name': vendor_name,
                        'phone': vendor_info.get('phone', ''),
                        'count': order_count,
                        'sheet_url': sheet_url,
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
                for vendor_info in vendors_info:
                    vn = vendor_info['name']
                    if vn in vendor_data:
                        log_entry['vendors'].append({
                            'name': vn,
                            'orders': len(vendor_data[vn]),
                            'sheet_uploaded': True,
                            'alimtalk_sent': True
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
                st.success(f"발주 처리 완료! {success_count}개 업체에 총 {len(merged_df)}건 시트 업로드 + 카카오 알림톡 발송 완료")

        # 알림톡 발송 내역 (접기/펼치기)
        if st.session_state.get('alimtalk_logs'):
            st.markdown("---")
            with st.expander("📱 카카오 알림톡 발송 내역", expanded=False):
                _date = st.session_state.get('alimtalk_date', '')
                for _al in st.session_state['alimtalk_logs']:
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.markdown(f"**{_al['name']}** ({_al['phone']})")
                    with col2:
                        st.markdown(f"""
                        <div class="kakao-msg">
                            <strong>[갓샵] 발주 알림</strong><br><br>
                            안녕하세요, {_al['name']} 사장님.<br>
                            {_date} 신규 주문 <strong>{_al['count']}건</strong>이 등록되었습니다.<br><br>
                            아래 링크에서 주문 내역 확인 후,<br>
                            '송장번호' 칸에 입력 부탁드립니다.<br><br>
                            ▶ <a href="{_al['sheet_url']}" target="_blank">발주서 확인하기</a>
                        </div>
                        """, unsafe_allow_html=True)
                    st.markdown(f"""
                    <div class="notification-bar">
                        ✅ {_al['name']} 사장님에게 알림톡 발송 완료!
                    </div>
                    """, unsafe_allow_html=True)


    # 발주 업로드 기록
    st.markdown("---")
    st.markdown("### 📋 발주 업로드 기록")

    upload_history = load_upload_history()
    if upload_history:
        for log in upload_history[:10]:
            with st.expander(f"📅 {log['date']} | 파일: {', '.join(log['files'])} | 총 {log['total_orders']}건"):
                for v in log.get('vendors', []):
                    sheet_icon = "✅" if v.get('sheet_uploaded') else "❌"
                    talk_icon = "✅" if v.get('alimtalk_sent') else "❌"
                    st.markdown(
                        f"&nbsp;&nbsp;&nbsp;&nbsp;**{v['name']}** — {v['orders']}건 | "
                        f"시트 업로드 {sheet_icon} | 알림톡 발송 {talk_icon}",
                        unsafe_allow_html=True
                    )
    else:
        st.info("아직 업로드 기록이 없습니다.")


# ===== 송장 현황 =====
elif page == "📊 송장 현황":
    st.markdown('<div class="main-header">📊 송장 현황</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">각 업체의 송장번호 입력 현황을 실시간으로 확인합니다</div>', unsafe_allow_html=True)

    # 15초마다 자동 새로고침
    refresh_count = st_autorefresh(interval=15000, limit=None, key="invoice_autorefresh")

    vendors_info = load_vendors()
    sheet_client = get_sheet_client()

    # 업체 선택 필터
    vendor_names = [v['name'] for v in (vendors_info or [])]
    col_vendor, col_refresh, col_status = st.columns([2, 1, 3])
    with col_vendor:
        selected_vendor = st.selectbox(
            "업체 선택", vendor_names,
            label_visibility="collapsed",
            key="invoice_vendor_filter"
        )
    with col_refresh:
        if st.button("🔄 수동 새로고침"):
            st.session_state['_sheet_cache_time'] = 0
    with col_status:
        st.caption(f"🟢 자동 새로고침 활성 (15초 간격) | 마지막 확인: {datetime.now().strftime('%H:%M:%S')}")

    if sheet_client and vendors_info:
        # 캐시된 데이터 사용 (20초 TTL)
        all_sheets = fetch_all_vendor_sheets(sheet_client, vendors_info)

        # 전체 업체 상태 수집 (알림 감지용)
        all_vendor_statuses = []
        total_orders_all = 0
        total_invoices_all = 0

        for vendor_info in vendors_info:
            vendor_name = vendor_info['name']
            sheet_url = vendor_info.get('google_sheet_url', '')
            if not sheet_url:
                continue

            data = all_sheets.get(vendor_name)
            if not data or len(data) < 2:
                all_vendor_statuses.append({
                    'name': vendor_name,
                    'total': 0, 'completed': 0, 'pending': 0, 'url': sheet_url
                })
                continue

            df = pd.DataFrame(data[1:], columns=data[0])
            order_count = len(df)
            invoice_count = len(df[df.get('송장번호', pd.Series()).notna() & (df.get('송장번호', pd.Series()) != '')])
            total_orders_all += order_count
            total_invoices_all += invoice_count

            all_vendor_statuses.append({
                'name': vendor_name,
                'total': order_count,
                'completed': invoice_count,
                'pending': order_count - invoice_count,
                'url': sheet_url
            })

        # 변경 감지 → 알림 토스트
        prev_invoices = st.session_state.get('prev_invoice_count', None)
        prev_vendor_status = st.session_state.get('prev_vendor_status', {})

        if prev_invoices is not None and total_invoices_all > prev_invoices:
            for vs in all_vendor_statuses:
                prev_count = prev_vendor_status.get(vs['name'], 0)
                if vs['completed'] > prev_count:
                    diff = vs['completed'] - prev_count
                    st.toast(f"🔔 {vs['name']}에서 송장 {diff}건 새로 입력!", icon="🔔")

        # 현재 상태 저장
        st.session_state['prev_invoice_count'] = total_invoices_all
        st.session_state['prev_vendor_status'] = {vs['name']: vs['completed'] for vs in all_vendor_statuses}
        st.session_state['invoice_count'] = total_invoices_all
        st.session_state['pending_count'] = total_orders_all - total_invoices_all

        # 선택된 업체 데이터만 필터링
        my_status = next((vs for vs in all_vendor_statuses if vs['name'] == selected_vendor), None)

        if my_status:
            my_orders = my_status['total']
            my_invoices = my_status['completed']
            my_pending = my_status['pending']

            prev_my = prev_vendor_status.get(selected_vendor, 0)
            delta_text = f"+{my_invoices - prev_my}건" if prev_invoices is not None and my_invoices > prev_my else None

            # 요약 카드
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("전체 주문", f"{my_orders}건")
            with col2:
                st.metric("송장 입력 완료", f"{my_invoices}건", delta=delta_text)
            with col3:
                st.metric("미입력", f"{my_pending}건")

            if my_orders > 0:
                _pct = int(my_invoices / my_orders * 100)
                st.markdown(f"""
                <div style="margin:1rem 0;">
                    <div style="display:flex;justify-content:space-between;margin-bottom:6px;font-size:0.85rem;">
                        <span style="color:#64748B;font-weight:500;">진행률</span>
                        <span style="color:#0F172A;font-weight:600;">{my_invoices}/{my_orders} ({_pct}%)</span>
                    </div>
                    <div style="background:#E2E8F0;border-radius:8px;height:10px;overflow:hidden;">
                        <div style="background:#4090C3;width:{_pct}%;height:100%;border-radius:8px;transition:width 0.4s ease;"></div>
                    </div>
                </div>""", unsafe_allow_html=True)

            # 내 업체 상세
            st.markdown("---")
            st.markdown(f"### {selected_vendor} 상세 현황")

            is_done = my_invoices == my_orders and my_orders > 0

            if is_done:
                status_badge = '<span style="background:#E8F5E9;color:#2E7D32;padding:5px 14px;border-radius:20px;font-size:0.8rem;font-weight:600;">✅ 입력 완료</span>'
                active_cls = "list-row-active"
            elif my_invoices > 0:
                status_badge = f'<span style="background:#FFF3E0;color:#E65100;padding:5px 14px;border-radius:20px;font-size:0.8rem;font-weight:600;">⏳ {my_invoices}/{my_orders}건 입력</span>'
                active_cls = ""
            elif my_orders > 0:
                status_badge = '<span style="background:#F1F5F9;color:#64748B;padding:5px 14px;border-radius:20px;font-size:0.8rem;font-weight:600;">⏳ 대기중</span>'
                active_cls = ""
            else:
                status_badge = '<span style="background:#F1F5F9;color:#94A3B8;padding:5px 14px;border-radius:20px;font-size:0.8rem;font-weight:600;">— 주문 없음</span>'
                active_cls = ""

            st.markdown(f"""
            <div class="list-row {active_cls}">
                <div style="flex:1;">
                    <div class="list-name">{selected_vendor}</div>
                    <div class="list-desc">{my_orders}건 주문</div>
                </div>
                <div style="display:flex;align-items:center;gap:10px;">
                    {status_badge}
                    <a href="{my_status['url']}" target="_blank" style="text-decoration:none;">
                        <div class="list-arrow">→</div>
                    </a>
                </div>
            </div>""", unsafe_allow_html=True)

        # 실시간 알림 로그 (내 업체만)
        st.markdown("---")
        st.markdown("### 🔔 알림 로그")

        if 'notification_log' not in st.session_state:
            st.session_state['notification_log'] = []

        if prev_invoices is not None and total_invoices_all > prev_invoices:
            for vs in all_vendor_statuses:
                prev_count = prev_vendor_status.get(vs['name'], 0)
                if vs['completed'] > prev_count:
                    diff = vs['completed'] - prev_count
                    log_entry = {
                        'time': datetime.now().strftime('%H:%M:%S'),
                        'vendor': vs['name'],
                        'count': diff,
                        'total': vs['completed']
                    }
                    st.session_state['notification_log'].insert(0, log_entry)

        my_logs = [log for log in st.session_state.get('notification_log', []) if log['vendor'] == selected_vendor]
        if my_logs:
            for log in my_logs[:10]:
                st.markdown(f"""
                <div class="notification-bar">
                    🔔 [{log['time']}] 송장번호 {log['count']}건 새로 입력! (총 {log['total']}건)
                </div>
                """, unsafe_allow_html=True)
        else:
            if my_status and my_status['completed'] > 0:
                st.markdown(f"""
                <div class="notification-bar">
                    🔔 송장번호 {my_status['completed']}건 입력 완료
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("아직 입력된 송장이 없습니다. 업체에서 입력하면 여기에 알림이 표시됩니다.")

    else:
        st.warning("구글 시트 연결이 필요합니다")


# ===== 송장 다운로드 =====
elif page == "📥 송장 다운로드":
    st.markdown('<div class="main-header">📥 송장 다운로드</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">입력된 송장번호를 수집하여 이지어드민 업로드용 엑셀을 생성합니다</div>', unsafe_allow_html=True)

    vendors_info = load_vendors()
    sheet_client = get_sheet_client()

    if st.button("📥 송장 수집 시작", type="primary", use_container_width=True):
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
                    st.success(f"✅ {vendor_name} — {len(df_with_invoice)}건 수집")
                else:
                    st.info(f"⏳ {vendor_name} — 입력된 송장 없음")

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
        st.markdown(f"### 📊 수집 결과: 총 {len(combined)}건")

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
                    st.metric(f"📦 {courier}", f"{len(group)}건")

        # 미리보기
        with st.expander("데이터 미리보기"):
            st.dataframe(combined[available_cols] if available_cols else combined, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📥 다운로드")

        download_mode = st.radio(
            "다운로드 방식",
            ["전체 통합 (1개 파일)", "택배사별 분리 (시트 나눔)", "택배사별 개별 파일"],
            horizontal=True,
            label_visibility="visible"
        )

        if download_mode == "전체 통합 (1개 파일)":
            buffer = BytesIO()
            combined[available_cols].to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            st.download_button(
                label=f"📥 전체 다운로드 (송장일괄등록_{today}.xlsx)",
                data=buffer,
                file_name=f"송장일괄등록_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        elif download_mode == "택배사별 분리 (시트 나눔)":
            if not courier_groups:
                st.warning("택배사 정보가 없어요.")
            else:
                st.caption(f"1개 엑셀 파일 안에 택배사별 시트로 분리돼요. ({', '.join(courier_groups.keys())})")
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    for courier, group in courier_groups.items():
                        sheet_name = courier[:31]  # 시트명 31자 제한
                        group[available_cols].to_excel(writer, sheet_name=sheet_name, index=False)
                buffer.seek(0)
                st.download_button(
                    label=f"📥 택배사별 시트 다운로드 (송장_{today}_택배사별.xlsx)",
                    data=buffer,
                    file_name=f"송장_{today}_택배사별.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

        elif download_mode == "택배사별 개별 파일":
            if not courier_groups:
                st.warning("택배사 정보가 없어요.")
            else:
                for courier, group in courier_groups.items():
                    buffer = BytesIO()
                    group[available_cols].to_excel(buffer, index=False, engine='openpyxl')
                    buffer.seek(0)
                    st.download_button(
                        label=f"📦 {courier} — {len(group)}건 다운로드",
                        data=buffer,
                        file_name=f"송장_{today}_{courier}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_{courier}"
                    )



# ===== 사이드바 실시간 알림 (맨 마지막에 렌더링) =====
with st.sidebar:
    st.markdown('<div class="sidebar-section-title">🔔 실시간 알림</div>', unsafe_allow_html=True)

    if 'notification_log' not in st.session_state:
        st.session_state['notification_log'] = []

    if st.session_state.get('notification_log'):
        for _log in st.session_state['notification_log'][:5]:
            st.markdown(
                f"<div class='sidebar-noti-item'>"
                f"<strong>{_log['time']}</strong>&nbsp;&nbsp;"
                f"{_log['vendor']} +{_log['count']}건"
                f"</div>",
                unsafe_allow_html=True
            )
    else:
        st.markdown(
            '<div style="padding:0 14px;font-size:0.78rem;color:#CBD5E1;">아직 새 알림이 없습니다</div>',
            unsafe_allow_html=True
        )
