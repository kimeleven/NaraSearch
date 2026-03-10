import datetime
import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import config


class BidSearchClient:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.base_url = config.API_BASE_URL
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "Mozilla/5.0"})

    def _build_params(self, date_str: str, page_no: int) -> dict:
        return {
            "serviceKey": self.api_key,
            "pageNo": page_no,
            "numOfRows": config.ROWS_PER_PAGE,
            "type": "json",
            "inqryDiv": 1,
            "inqryBgnDt": date_str + "0000",
            "inqryEndDt": date_str + "2359",
        }

    def fetch_page(self, date_str: str, page_no: int) -> list:
        params = self._build_params(date_str, page_no)
        resp = self.session.get(
            self.base_url,
            params=params,
            timeout=config.REQUEST_TIMEOUT_SECONDS,
        )
        resp.raise_for_status()

        content_type = resp.headers.get("Content-Type", "")
        if "json" not in content_type:
            raise ValueError(
                f"JSON 응답이 아닙니다. API 키를 확인하세요.\n"
                f"Content-Type: {content_type}\n"
                f"응답 내용(앞 300자): {resp.text[:300]}"
            )

        data = resp.json()
        header = data.get("response", {}).get("header", {})
        result_code = header.get("resultCode", "")
        if result_code != "00":
            result_msg = header.get("resultMsg", "알 수 없는 오류")
            raise RuntimeError(f"API 오류 [{result_code}]: {result_msg}")

        body = data.get("response", {}).get("body", {})
        items = body.get("items")

        # 결과가 없으면 items가 None, 빈 문자열, 빈 리스트로 옴
        if not items:
            return []

        # 이 API는 items가 바로 list로 옴
        if isinstance(items, list):
            return items
        # 혹시 dict 구조인 경우 (item 키 하위)
        if isinstance(items, dict):
            item = items.get("item", [])
            if isinstance(item, dict):
                return [item]
            return item if item else []
        return []

    def fetch_all(self, date_str: str) -> list:
        all_items = []
        page = 1
        while True:
            print(f"  페이지 {page} 조회 중...", end="\r")
            items = self.fetch_page(date_str, page)
            if not items:
                break
            all_items.extend(items)
            if len(items) < config.ROWS_PER_PAGE:
                break
            page += 1
        print()
        return all_items


def filter_records(records: list) -> list:
    seen = set()
    result = []
    for r in records:
        name = r.get("bidNtceNm", "")
        task_type = r.get("srvceDivNm", "")
        bid_no = r.get("bidNtceNo", id(r))

        match_a = (task_type in config.ALLOWED_TASK_TYPES
                   and any(kw in name for kw in config.KEYWORDS))
        match_b = any(kw.lower() in name.lower() for kw in config.TABLET_KEYWORDS)

        if (match_a or match_b) and bid_no not in seen:
            seen.add(bid_no)
            result.append(r)
    return result


def build_dataframe(records: list) -> pd.DataFrame:
    rows = []
    for r in records:
        row = {field: r.get(field, "") for field in config.EXCEL_COLUMNS}
        rows.append(row)

    df = pd.DataFrame(rows, columns=list(config.EXCEL_COLUMNS.keys()))
    df.rename(columns=config.EXCEL_COLUMNS, inplace=True)
    df.sort_values("공고일시", ascending=True, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def export_to_excel(df: pd.DataFrame, date_str: str) -> str:
    display_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
    filename = f"입찰공고_{display_date}.xlsx"

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="입찰공고")

        ws = writer.sheets["입찰공고"]

        # 헤더 스타일
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(fill_type="solid", fgColor="2F5496")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

        # 열 너비 자동 조정
        for col_idx, col_cells in enumerate(ws.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for cell in col_cells:
                if cell.value:
                    # 한글은 2바이트 너비로 계산
                    text = str(cell.value)
                    length = sum(2 if ord(c) > 127 else 1 for c in text)
                    max_len = max(max_len, length)
            ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    return filename


def send_email(excel_path: str, date_str: str, count: int) -> None:
    sender = os.environ.get("EMAIL_FROM", "")
    recipients_raw = os.environ.get("EMAIL_TO", "")
    password = os.environ.get("EMAIL_PASSWORD", "")

    if not all([sender, recipients_raw, password]):
        print("이메일 설정(EMAIL_FROM/EMAIL_TO/EMAIL_PASSWORD)이 없어 발송 생략.")
        return

    # 쉼표로 구분된 여러 수신자 지원
    recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]

    display_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
    subject = f"[나라장터] {display_date} 입찰공고 {count}건"

    body = (
        f"{display_date} 입찰공고 검색 결과입니다.\n\n"
        f"- 업무구분: {', '.join(config.ALLOWED_TASK_TYPES)}\n"
        f"- 키워드: {', '.join(config.KEYWORDS)}\n"
        f"- 총 {count}건\n\n"
        "첨부 엑셀 파일을 확인해 주세요."
    )

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with open(excel_path, "rb") as f:
        attachment = MIMEApplication(f.read(), Name=os.path.basename(excel_path))
        attachment["Content-Disposition"] = f'attachment; filename="{os.path.basename(excel_path)}"'
        msg.attach(attachment)

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(sender, password)
        smtp.sendmail(sender, recipients, msg.as_string())

    print(f"이메일 발송 완료 → {', '.join(recipients)}")


def main():
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    date_str = yesterday.strftime("%Y%m%d")
    display_date = yesterday.strftime("%Y-%m-%d")

    try:
        api_key = config.get_api_key()
    except RuntimeError as e:
        print(f"[오류] {e}")
        sys.exit(1)

    print(f"날짜 {display_date} 의 입찰공고를 조회합니다...")
    client = BidSearchClient(api_key)

    try:
        raw_records = client.fetch_all(date_str)
    except (requests.HTTPError, requests.Timeout, requests.ConnectionError) as e:
        print(f"[오류] API 호출 실패: {e}")
        sys.exit(1)
    except (ValueError, RuntimeError) as e:
        print(f"[오류] {e}")
        sys.exit(1)

    print(f"총 {len(raw_records)}건 조회됨")

    filtered = filter_records(raw_records)
    print(
        f"필터링 결과: {len(filtered)}건 "
        f"(A: 업무구분({', '.join(config.ALLOWED_TASK_TYPES)}) + "
        f"키워드({', '.join(config.KEYWORDS)}) | "
        f"B: 태블릿({', '.join(config.TABLET_KEYWORDS)}))"
    )

    if not filtered:
        print("조건에 맞는 공고가 없습니다.")
        return

    df = build_dataframe(filtered)
    output_path = export_to_excel(df, date_str)
    print(f"저장 완료: {output_path}  ({len(df)}건)")

    send_email(output_path, date_str, len(df))


if __name__ == "__main__":
    main()
