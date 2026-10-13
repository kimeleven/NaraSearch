# Extra data.go.kr sources (나라장터 daily pipeline)

Secrets live under:

`/home/box/agent-data/connector-secrets/11fcf4be-d529-49fb-9260-cd8d9885d493/`

Never log or copy API key values outside those JSON files.

## Working endpoints

### 1. K-Startup
- Secret: `kstartup.json` (`KSTARTUP_API_KEY`, `KSTARTUP_API_BASE`)
- Base: `https://apis.data.go.kr/B552735/kisedKstartupService01`
- Operation: `getAnnouncementInformation01`
- Params: `serviceKey`, `page`, `perPage`, `returnType=json`
- Useful fields: `biz_pbanc_nm`, `pbanc_rcpt_bgng_dt`, `pbanc_rcpt_end_dt`, `rcrt_prgs_yn`, `supt_regin`, `detl_pg_url`, …
- Date filter: API `cond[...]` hints are unreliable; client-side filter on `pbanc_rcpt_bgng_dt` (YYYYMMDD).
- Sheet name: `K-Startup`

### 2. MSIT (과학기술정보통신부 사업공고)
- Secret: `msit.json` (`MSIT_API_KEY`, `MSIT_API_BASE`)
- Base: `https://apis.data.go.kr/1721000/msitannouncementinfo`
- Operation: `businessAnnouncMentList`
- Params: `serviceKey`, `pageNo`, `numOfRows`, `returnType=json`
- Note: portal often caps `numOfRows` at 10; paginate newest-first and stop when `pressDt` exits the window.
- Useful fields: `subject`, `viewUrl`, `deptName`, `pressDt`, `managerName`, `managerTel`, `files`
- Sheet name: `과기정통부`

## Blocked / skipped — 중기부 data.go.kr (1421000)

- Secret on disk: `bizinfo_data_go.json` (`BIZINFO_DATA_GO_API_KEY`, `BIZINFO_DATA_GO_API_BASE`)
- User base `https://apis.data.go.kr/1421000/bizinfo` → `NO_OPENAPI_SERVICE_ERROR` for all ops tried (invalid/unknown service path).
- Likely intended API: `https://apis.data.go.kr/1421000/mssBizService_v2/getbizList_v2`
  - Correct op name: `getbizList_v2`
  - Current key returns `403 SERVICE_KEY_IS_NOT_REGISTERED_ERROR`
  - Needs a separate 활용신청 for **중소벤처기업부_사업공고** (dataset `15113297`).
- Until a registered key exists, **skip** this source.
- Existing daily sheet **정부지원사업** continues to use **bizinfo.go.kr** (`crtfcKey` via `BIZINFO_API_KEY` / `config.get_bizinfo_api_key()`), not data.go.kr.

## Existing 나라장터 key

- Secret: `nara.json` → `NARA_API_KEY` (also award keys unused by this extra pipeline)

## Scripts

| Script | Role |
|--------|------|
| `extra_sources.py` | Fetch K-Startup + MSIT for a KST target date (retries, pagination, client date filter) |
| `apply_daily_filters.py` | 2nd filters + write `입찰공고_YYYY-MM-DD_v2.xlsx` |
| `rfp_search.py` | Original 나라장터 + 기업마당 (bizinfo.go.kr) collector |

### Example

```bash
/workspace/nara-search/.venv/bin/python /workspace/nara-search/apply_daily_filters.py --date 2026-09-08
```

Output: `/workspace/nara-search/입찰공고_2026-09-08_v2.xlsx`  
Sheets: `나라장터`, `정부지원사업`, `K-Startup`, `과기정통부`

## 키 설정 (Hermes 등 다른 환경)
`.env`에 `KSTARTUP_API_KEY`, `KSTARTUP_API_BASE`, `MSIT_API_KEY`, `MSIT_API_BASE`를 넣으면 JSON 파일 대신 환경변수를 사용. 또는 `NARA_SECRETS_DIR`로 JSON 폴더 지정. 키는 저장소에 커밋하지 말 것.

## 해외 서버에서 실행할 때 (403·타임아웃)
`apis.data.go.kr`, `bizinfo.go.kr`은 해외 IP에서 막히거나 끊길 수 있다. 메인 스크립트에서 다른 모듈보다 먼저 `import _allorigins_patch`를 하면 이 도메인 요청이 공개 중계(proxy.cors.dev, 실패 시 allorigins)를 거친다.
주의: 이때 API 키가 제3자 중계 서버를 지나간다. 한국 IP에서 돌릴 수 있으면 쓰지 말 것.
