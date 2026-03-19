import requests
import datetime

API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
API_KEY = "AbiQPw"

ROWS_PER_PAGE = 100

def fetch_page(page_no):
    params = {
        "crtfcKey": API_KEY,
        "dataType": "json",
        "pageUnit": ROWS_PER_PAGE,
        "pageIndex": page_no,
    }
    resp = requests.get(API_URL, params=params, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.json().get("jsonArray", [])

def fetch_recent(days=7):
    cutoff = (datetime.date.today() - datetime.timedelta(days=days)).strftime("%Y-%m-%d")
    print(f"기준일(cutoff): {cutoff}")
    result = []
    page = 1
    while True:
        items = fetch_page(page)
        print(f"페이지 {page}: {len(items)}건 수신")
        if not items:
            break
        found_older = False
        for item in items:
            created = item.get("creatPnttm", "")
            name = item.get("pblancNm", "")
            if created and created[:10] >= cutoff:
                result.append(item)
            elif created and created[:10] < cutoff:
                print(f"  -> 오래된 항목 발견: {created[:10]} | {name[:30]}")
                found_older = True
                break
            else:
                print(f"  -> creatPnttm 없음: {name[:30]}")
        if found_older or len(items) < ROWS_PER_PAGE:
            break
        page += 1
    return result

items = fetch_recent(days=1)
print(f"\n최종 수집: {len(items)}건")
for item in items[:5]:
    print(f"  {item.get('creatPnttm', '')} | {item.get('pblancNm', '')[:40]}")
