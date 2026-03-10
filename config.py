import os
from dotenv import load_dotenv

load_dotenv()

# ── API 설정 ──────────────────────────────────────────────────
API_BASE_URL = (
    "https://apis.data.go.kr/1230000/ad/BidPublicInfoService"
    "/getBidPblancListInfoServcPPSSrch"
)
ROWS_PER_PAGE = 999
REQUEST_TIMEOUT_SECONDS = 30

# ── 필터 설정 ─────────────────────────────────────────────────
ALLOWED_TASK_TYPES = {"일반용역", "민간일반용역"}

KEYWORDS = ["플랫폼", "시스템", "앱", "위험성", "안전"]

# ── 엑셀 컬럼 매핑 (API 필드명 → 한국어 헤더) ─────────────────
EXCEL_COLUMNS = {
    "bidNtceNo":    "공고번호",
    "bidNtceNm":    "공고명",
    "ntceInsttNm":  "공고기관",
    "dminsttNm":    "수요기관",
    "srvceDivNm":   "업무구분",
    "bidNtceDt":    "공고일시",
    "bidClseDt":    "입찰마감일시",
    "asignBdgtAmt": "배정예산액",
    "presmptPrce":  "추정가격",
    "bidNtceUrl":   "공고URL",
}


def get_api_key() -> str:
    key = os.getenv("NARA_API_KEY")
    if not key:
        raise RuntimeError(
            "NARA_API_KEY 환경변수가 설정되지 않았습니다.\n"
            ".env.example 파일을 .env 로 복사한 뒤 API 키를 입력하세요.\n"
            "API 키 발급: https://www.data.go.kr/data/15129394/openapi.do"
        )
    return key
