# 나라장터 입찰공고 자동 검색기

나라장터(g2b.go.kr) OpenAPI를 활용하여 매일 전날 올라온 입찰공고를 자동으로 수집하고, 조건에 맞는 공고를 필터링하여 이메일로 발송하는 자동화 프로그램입니다.

## 주요 기능

- **자동 수집**: 전날 하루 동안 올라온 입찰공고 전체 수집
- **조건 A** - 업무구분 `일반용역`, `민간일반용역` + 공고명에 `플랫폼`, `시스템`, `앱`, `위험성`, `안전`, `서비스`, `구축`, `개발`, `IOT`, `Embedded`, `임베디드` 포함
- **조건 B** - 업무구분 무관, 공고명에 `태블릿`, `타블릿`, `tablet`, `테블렛`, `타블렛` 포함 (대소문자 무시)
- 두 조건 중 하나라도 해당하면 수집 (OR)
- **제외**: 공고명에 `대행`, `감리`, `연구용역` 포함 시 제외
- **엑셀 출력**: 날짜별 파일명으로 결과 저장 (예: `입찰공고_2026-03-09.xlsx`)
- **이메일 발송**: Gmail SMTP로 결과 엑셀 첨부 자동 발송
- **클라우드 자동 실행**: GitHub Actions로 매일 오전 9시(KST) 무인 실행
- **재시도**: 공공데이터포털 연결 실패(타임아웃·5xx) 시 15/30/60초 간격으로 최대 4회 재시도

## 출력 엑셀 컬럼

| 컬럼 | 설명 |
|------|------|
| 공고번호 | 입찰공고 고유번호 |
| 공고명 | 입찰공고 제목 |
| 공고기관 | 공고를 등록한 기관 |
| 수요기관 | 실제 수요 기관 |
| 업무구분 | 일반용역 / 민간일반용역 |
| 입찰방식 | 전자입찰 등 (`bidMethdNm`) |
| 낙찰방법 | 낙찰방법명 (`sucsfbidMthdNm`) |
| 낙찰방법세부기준 | 동일 API 필드 (`sucsfbidMthdNm`, 문서상 세부기준 설명) |
| 공고일시 | 공고 등록 일시 |
| 입찰마감일시 | 입찰 마감 일시 |
| 배정예산액 | 배정 예산 (원) |
| 추정가격 | 추정 가격 (원) |
| 낙찰하한가율 | 개찰 시 낙찰하한율 % (`sucsfbidLwltRate`) |
| 공고URL | 나라장터 공고 상세 페이지 링크 |

## 사전 준비

### 1. 나라장터 OpenAPI 키 발급

1. [공공데이터포털](https://www.data.go.kr/data/15129394/openapi.do) 접속
2. 회원가입 후 **나라장터 입찰공고정보서비스** 활용 신청
3. 승인 후 **일반 인증키** 발급 (수 시간 소요될 수 있음)

### 2. Gmail 앱 비밀번호 발급

1. [Google 앱 비밀번호](https://myaccount.google.com/apppasswords) 접속
2. 2단계 인증이 활성화되어 있어야 함
3. 앱 이름 입력 후 **만들기** → 16자리 코드 저장

## 로컬 실행

```bash
# 1. 패키지 설치
pip install -r requirements.txt

# 2. 환경변수 설정
cp .env.example .env
# .env 파일 열어서 실제 값 입력

# 3. 실행
python rfp_search.py
```

### .env 파일 예시

```
NARA_API_KEY=발급받은_API_키
EMAIL_FROM=보내는계정@gmail.com
EMAIL_TO=받는계정1@example.com,받는계정2@example.com
EMAIL_PASSWORD=Gmail_앱_비밀번호_16자리
```

> `.env` 파일은 절대 git에 올리지 마세요. `.gitignore`에 등록되어 있습니다.

## GitHub Actions 자동 실행 설정

### 1. Repository Secrets 등록

GitHub 저장소 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Secret 이름 | 값 |
|-------------|-----|
| `NARA_API_KEY` | 공공데이터포털 API 키 |
| `BIZINFO_API_KEY` | 기업마당 API 키 (없으면 기본값 사용) |
| `EMAIL_FROM` | 발송 Gmail 주소 |
| `EMAIL_TO` | 수신 이메일 (쉼표로 여러 명 가능) |
| `EMAIL_PASSWORD` | Gmail 앱 비밀번호 16자리 |

### 2. 실행 일정

- **자동 실행**: 매일 오전 9시 KST (UTC 00:00)
- **수동 실행**: Actions 탭 → 워크플로우 선택 → **Run workflow**

### 3. 결과 확인

- 이메일로 엑셀 파일 첨부 발송
- Actions 탭 → 해당 실행 클릭 → **Artifacts** 에서 엑셀 파일 30일간 보관
- 조건에 맞는 공고가 없거나 API가 끝까지 실패한 날에는 엑셀이 없어 아티팩트도 없습니다

### 4. 장애 대응

공공데이터포털(`apis.data.go.kr`)이 간헐적으로 연결 타임아웃을 내면 검색이 실패하고 엑셀·메일이 빠질 수 있습니다.  
`config.py`의 `API_MAX_RETRIES`(기본 4회), `API_RETRY_BACKOFF_SECONDS`(기본 15초, 이후 2배)로 재시도합니다.

## 파일 구조

```
NaraSearch/
├── rfp_search.py          # 메인 스크립트
├── config.py              # 설정 (API, 필터 조건, 컬럼 매핑)
├── requirements.txt       # Python 패키지 목록
├── .env.example           # 환경변수 예시
├── .env                   # 실제 환경변수 (git 제외)
├── .gitignore
└── .github/
    └── workflows/
        └── rfp_search.yml # GitHub Actions 워크플로우
```

## 필터 조건 변경

[config.py](config.py) 에서 수정:

```python
# 업무구분 필터
ALLOWED_TASK_TYPES = {"일반용역", "민간일반용역"}

# 공고명 키워드 필터
KEYWORDS = ["플랫폼", "시스템", "앱", "위험성", "안전", "서비스", "구축", "개발", "IOT", "Embedded", "임베디드"]
TABLET_KEYWORDS = ["태블릿", "타블릿", "tablet", "테블렛", "타블렛"]

# 공고명 제외 키워드
EXCLUDE_KEYWORDS = ["대행", "감리", "연구용역"]
```

## 사용 API

- **서비스**: 조달청 나라장터 입찰공고정보서비스
- **엔드포인트**: `https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch`
- **출처**: [공공데이터포털](https://www.data.go.kr/data/15129394/openapi.do)
