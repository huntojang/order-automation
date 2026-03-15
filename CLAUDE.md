# 발주도우미 프로젝트

## 톤 & 스타일
- 이모지 가끔 써주기
- 문장 끝에 '~' 붙이지 말고 'ㅎㅎ' 붙이기
- 존대말 사용
- 귀여운 여자아이 같은 말투로 발랄하게 대화하기

## 프로젝트 정보
- **GitHub**: `huntojang/order-automation` (public)
- **배포 URL**: https://balju-doumi.streamlit.app
- **기술 스택**: Python, Streamlit, Google Sheets API, 알리고 (알림톡)
- **구글 드라이브 공유폴더 ID**: `1Uh3c_c7kakXIXlTX8BT89HtBcA3nCY1v`

## 현재 고객사
- 크리에이티브크루 (담당자: 010-3002-1239)
- 구글 시트: https://docs.google.com/spreadsheets/d/10EmOQgygX_OreaaQQmVe2ecTNzAcPcqlD6imutbz_gU

## 연동: questloom.io (별도 프로젝트)
- **프로젝트 경로**: `/Users/woojeonjang/Documents/dev/questloom_homepage/questloom-web`
- **기술 스택**: Next.js 16 + Vercel + Supabase + PortOne V2
- **배포 URL**: https://questloom.io
- **GitHub**: `huntojang/questloom-web`

### questloom.io 완성 상태 (2026-03-15)
- ✅ Supabase 인증 (이메일/비번, 회원가입/로그인/비번 리셋)
- ✅ PortOne V2 결제 연동 (정기결제 빌링키)
- ✅ JWT SSO 토큰 발급: `POST /api/sso/token`
- ✅ 구독 검증 API: `GET /api/subscription/verify?service=orderhelper`
- ✅ Console 대시보드 (구독/결제/설정)
- ✅ Admin 대시보드 (고객/구독/결제 관리)
- ✅ 쿠폰 시스템, 감사 로그, Rate Limiting

### 발주도우미 가격 플랜 (Supabase plans 테이블)
- 스타터: ₩39,000/월 (서비스 slug: `orderhelper`)
- 프로: ₩79,000/월

### SSO 연동 방식
- **인증**: questloom.io에서 로그인 → JWT 토큰 발급 (HS256, 1시간 만료)
- **구독 확인**: `Authorization: Bearer {JWT}` + `x-api-key` 헤더로 검증
- **플로우**: questloom.io 구독 결제 → "발주도우미 시작" → `balju-doumi.streamlit.app?token=JWT`
- **환경변수**: `SSO_JWT_SECRET`, `SSO_API_KEY` (questloom.io .env.local에 있음)

### 발주도우미 현재 완성 상태 (2026-03-15)
- ✅ 엑셀 업로드 → 업체별 구글 시트 분배
- ✅ 카카오 알림톡 발송 (알리고 API, GCP 프록시)
- ✅ 업체 관리 페이지 (구글 시트 마스터 DB + Apps Script 시트 자동 생성)
- ✅ 송장 현황/다운로드
- ⏳ questloom.io SSO 연동 (로그인/구독 체크)

## 서비스 구조
```
questloom.io (공홈 - Next.js/Vercel - 결제/인증 포털)
├── 발주도우미: balju-doumi.streamlit.app (Python/Streamlit)
└── 작업체커: TBD (Next.js, /Users/woojeonjang/Documents/dev/작업체커)
```
