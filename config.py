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

KEYWORDS = ["플랫폼", "시스템", "앱", "위험성", "안전", "서비스", "구축", "개발", "IOT", "Embedded", "임베디드"]
TABLET_KEYWORDS = ["태블릿", "타블릿", "tablet", "테블렛", "타블렛"]

# ── 엑셀 컬럼 매핑 (API 필드명 → 한국어 헤더) ─────────────────
# 순서를 유지하기 위해 리스트 사용. 동일 API 필드를 여러 컬럼에 쓸 수 있음.
# sucsfbidMthdNm: 문서상 "낙찰방법명"이며 설명은 "낙찰방법 세부기준"
#   (예: "협상에의한계약", "적격심사제-추정가격 2억원 미만인 용역")
EXCEL_COLUMNS = [
    ("bidNtceNo",       "공고번호"),
    ("bidNtceNm",       "공고명"),
    ("ntceInsttNm",     "공고기관"),
    ("dminsttNm",       "수요기관"),
    ("srvceDivNm",      "업무구분"),
    ("bidMethdNm",      "입찰방식"),
    ("sucsfbidMthdNm",  "낙찰방법"),
    ("sucsfbidMthdNm",  "낙찰방법세부기준"),
    ("bidNtceDt",       "공고일시"),
    ("bidClseDt",       "입찰마감일시"),
    ("asignBdgtAmt",    "배정예산액"),
    ("presmptPrce",     "추정가격"),
    ("sucsfbidLwltRate", "낙찰하한가율"),
    ("bidNtceUrl",      "공고URL"),
]


def get_api_key() -> str:
    key = os.getenv("NARA_API_KEY")
    if not key:
        raise RuntimeError(
            "NARA_API_KEY 환경변수가 설정되지 않았습니다.\n"
            ".env.example 파일을 .env 로 복사한 뒤 API 키를 입력하세요.\n"
            "API 키 발급: https://www.data.go.kr/data/15129394/openapi.do"
        )
    return key


# ── 기업마당 (bizinfo) API 설정 ────────────────────────────────
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
BIZINFO_ROWS_PER_PAGE = 100

BIZINFO_EXCEL_COLUMNS = {
    "pblancNm":                   "공고명",
    "jrsdInsttNm":                "주관기관",
    "excInsttNm":                 "운영기관",
    "pldirSportRealmLclasCodeNm": "분야",
    "reqstBeginEndDe":            "신청기간",
    "rceptEngnHmpgUrl":           "홈페이지URL",
    "creatPnttm":                 "등록일시",
}


def get_bizinfo_api_key() -> str:
    return os.getenv("BIZINFO_API_KEY", "AbiQPw")
