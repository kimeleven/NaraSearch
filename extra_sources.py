"""Fetch K-Startup and MSIT (과기정통부) announcements from data.go.kr.

Secrets are loaded from connector JSON files; key values are never printed.
Bizinfo data.go.kr (1421000) is intentionally skipped — see EXTRA_SOURCES.md.
"""

from __future__ import annotations

import datetime
import json
import os
import time
from pathlib import Path
from typing import Any, Callable, Optional
from zoneinfo import ZoneInfo

import pandas as pd
import requests

KST = ZoneInfo("Asia/Seoul")

SECRETS_DIR = Path(
    os.getenv(
        "NARA_SECRETS_DIR",
        "/home/box/agent-data/connector-secrets/11fcf4be-d529-49fb-9260-cd8d9885d493",
    )
)

_ENV_KEYS = {
    "kstartup.json": ("KSTARTUP_API_KEY", "KSTARTUP_API_BASE"),
    "msit.json": ("MSIT_API_KEY", "MSIT_API_BASE"),
}

REQUEST_TIMEOUT = 45
MAX_RETRIES = 8
BACKOFF_BASE = 2.0
BACKOFF_MAX = 30.0

KSTARTUP_OP = "getAnnouncementInformation01"
MSIT_OP = "businessAnnouncMentList"

KSTARTUP_COLUMNS = [
    ("biz_pbanc_nm", "공고명"),
    ("pbanc_ntrp_nm", "공고기업명"),
    ("sprv_inst", "감독기관구분"),
    ("supt_regin", "지원지역"),
    ("supt_biz_clsfc", "지원분야"),
    ("pbanc_rcpt_bgng_dt", "접수시작일"),
    ("pbanc_rcpt_end_dt", "접수마감일"),
    ("rcrt_prgs_yn", "모집중여부"),
    ("detl_pg_url", "상세URL"),
    ("biz_aply_url", "신청URL"),
    ("biz_prch_dprt_nm", "담당부서"),
    ("prch_cnpl_no", "연락처"),
]

MSIT_COLUMNS = [
    ("subject", "제목"),
    ("deptName", "담당부서"),
    ("pressDt", "게시일"),
    ("managerName", "담당자"),
    ("managerTel", "연락처"),
    ("viewUrl", "상세URL"),
]


class TransientError(Exception):
    pass


def _load_secret(filename: str) -> dict:
    """환경변수(KSTARTUP_API_KEY 등)가 있으면 우선 사용, 없으면 SECRETS_DIR의 JSON."""
    keys = _ENV_KEYS.get(filename)
    if keys and all(os.getenv(k) for k in keys):
        return {k: os.getenv(k) for k in keys}
    path = SECRETS_DIR / filename
    return json.loads(path.read_text(encoding="utf-8"))


def yesterday_kst(today: Optional[datetime.date] = None) -> datetime.date:
    if today is None:
        today = datetime.datetime.now(KST).date()
    return today - datetime.timedelta(days=1)


def _get_json(session: requests.Session, url: str, params: dict) -> Any:
    last: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if resp.status_code in {429, 500, 502, 503, 504}:
                raise TransientError(f"HTTP {resp.status_code}")
            resp.raise_for_status()
            ctype = resp.headers.get("Content-Type", "")
            if "json" not in ctype.lower() and "javascript" not in ctype.lower():
                # some portal endpoints omit charset hints; still try JSON
                try:
                    return resp.json()
                except Exception as e:
                    raise TransientError(f"non-JSON Content-Type: {ctype}") from e
            return resp.json()
        except (requests.Timeout, requests.ConnectionError, TransientError) as e:
            last = e
            if attempt >= MAX_RETRIES:
                raise
            wait = min(BACKOFF_BASE * (2 ** (attempt - 1)), BACKOFF_MAX)
            print(f"  요청 실패: {e}. {wait:.0f}초 후 재시도 ({attempt}/{MAX_RETRIES})...", flush=True)
            time.sleep(wait)
    raise last  # pragma: no cover


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "nara-search-extra-sources/1.0"})
    return s


# ── K-Startup ───────────────────────────────────────────────────────────────


def fetch_kstartup(
    target_date: datetime.date,
    *,
    per_page: int = 100,
    open_only: bool = True,
    lookback_days: int = 0,
) -> list[dict]:
    """Fetch K-Startup announcements whose 접수시작일 is on target_date (±lookback).

    Uses client-side date filtering. Optionally tries a cond[] hint (API may ignore it).
    """
    sec = _load_secret("kstartup.json")
    base = sec["KSTARTUP_API_BASE"].rstrip("/")
    key = sec["KSTARTUP_API_KEY"]
    url = f"{base}/{KSTARTUP_OP}"

    start = target_date - datetime.timedelta(days=lookback_days)
    start_s = start.strftime("%Y%m%d")
    end_s = target_date.strftime("%Y%m%d")

    session = _session()
    # Hint only — verified inaccurate matchCount; still useful to shrink payload sometimes.
    extra = {"cond[pbanc_rcpt_bgng_dt::GTE]": start_s}
    collected: list[dict] = []
    page = 1
    while True:
        params = {
            "serviceKey": key,
            "page": page,
            "perPage": per_page,
            "returnType": "json",
            **extra,
        }
        print(f"  [K-Startup] page {page}...", end="\r", flush=True)
        data = _get_json(session, url, params)
        batch = data.get("data") or []
        if not batch:
            break
        collected.extend(batch)
        match = data.get("matchCount")
        if len(batch) < per_page:
            break
        if match is not None and len(collected) >= int(match):
            break
        # Safety: do not walk the entire 30k catalog
        if page >= 40:
            break
        page += 1
    print()

    def in_window(item: dict) -> bool:
        bgng = str(item.get("pbanc_rcpt_bgng_dt") or "").replace("-", "")[:8]
        if not bgng:
            return False
        return start_s <= bgng <= end_s

    filtered = [it for it in collected if in_window(it)]
    if open_only:
        filtered = [it for it in filtered if str(it.get("rcrt_prgs_yn", "Y")).upper() != "N"]
    print(
        f"  [K-Startup] raw_pages={len(collected)} in_window={len(filtered)} "
        f"({start_s}~{end_s})",
        flush=True,
    )
    return filtered


def kstartup_to_dataframe(records: list[dict]) -> pd.DataFrame:
    rows = [{src: r.get(src, "") for src, _ in KSTARTUP_COLUMNS} for r in records]
    df = pd.DataFrame(rows, columns=[src for src, _ in KSTARTUP_COLUMNS])
    df.rename(columns={src: dst for src, dst in KSTARTUP_COLUMNS}, inplace=True)
    if not df.empty:
        df.sort_values(["접수시작일", "공고명"], ascending=[False, True], inplace=True)
        df.reset_index(drop=True, inplace=True)
    return df


# ── MSIT (과기정통부) ────────────────────────────────────────────────────────


def _unwrap_msit_item(wrap: Any) -> dict:
    if isinstance(wrap, dict) and "item" in wrap:
        item = wrap["item"]
        if isinstance(item, dict):
            return item
        if isinstance(item, str):
            import ast

            return ast.literal_eval(item)
    if isinstance(wrap, dict):
        return wrap
    raise TypeError(f"unexpected MSIT item: {type(wrap)}")


def _msit_body(data: dict) -> dict:
    resp = data.get("response")
    if isinstance(resp, list):
        for part in resp:
            if isinstance(part, dict) and "body" in part:
                return part["body"]
        raise RuntimeError("MSIT response missing body")
    if isinstance(resp, dict):
        return resp.get("body") or {}
    raise RuntimeError(f"unexpected MSIT response shape: {type(resp)}")


def fetch_msit(
    target_date: datetime.date,
    *,
    lookback_days: int = 1,
    num_of_rows: int = 100,
) -> list[dict]:
    """Fetch MSIT business announcements with pressDt on target_date (±lookback).

    API returns newest-first; pagination stops once pressDt falls below window.
    Note: portal often caps numOfRows at 10 regardless of request.
    """
    sec = _load_secret("msit.json")
    base = sec["MSIT_API_BASE"].rstrip("/")
    key = sec["MSIT_API_KEY"]
    url = f"{base}/{MSIT_OP}"

    start = target_date - datetime.timedelta(days=lookback_days)
    start_s = start.strftime("%Y-%m-%d")
    end_s = target_date.strftime("%Y-%m-%d")

    session = _session()
    collected: list[dict] = []
    page = 1
    while page <= 50:
        params = {
            "serviceKey": key,
            "pageNo": page,
            "numOfRows": num_of_rows,
            "returnType": "json",
        }
        print(f"  [MSIT] page {page}...", end="\r", flush=True)
        data = _get_json(session, url, params)
        body = _msit_body(data)
        # header check when present
        resp = data.get("response")
        if isinstance(resp, list):
            for part in resp:
                hdr = part.get("header") if isinstance(part, dict) else None
                if hdr and hdr.get("resultCode") not in (None, "00", 0, "0"):
                    raise RuntimeError(f"MSIT API error: {hdr}")

        items = body.get("items") or []
        if isinstance(items, dict):
            items = items.get("item", [])
            if isinstance(items, dict):
                items = [items]
        if not items:
            break

        stop = False
        for wrap in items:
            it = _unwrap_msit_item(wrap)
            pdt = str(it.get("pressDt") or "")[:10]
            if not pdt:
                continue
            if start_s <= pdt <= end_s:
                collected.append(it)
            elif pdt < start_s:
                stop = True
                break
        page_size = int(body.get("numOfRows") or len(items) or 10)
        if stop or len(items) < page_size:
            break
        page += 1
    print()
    print(f"  [MSIT] in_window={len(collected)} ({start_s}~{end_s})", flush=True)
    return collected


def msit_to_dataframe(records: list[dict]) -> pd.DataFrame:
    rows = []
    for r in records:
        row = {src: r.get(src, "") for src, _ in MSIT_COLUMNS}
        # flatten files if present
        files = r.get("files") or []
        names = []
        if isinstance(files, list):
            for f in files:
                if isinstance(f, dict):
                    inner = f.get("file", f)
                    if isinstance(inner, dict) and inner.get("fileName"):
                        names.append(str(inner["fileName"]))
        row["files"] = "; ".join(names)
        rows.append(row)
    cols = [src for src, _ in MSIT_COLUMNS] + ["files"]
    df = pd.DataFrame(rows, columns=cols)
    rename = {src: dst for src, dst in MSIT_COLUMNS}
    rename["files"] = "첨부파일"
    df.rename(columns=rename, inplace=True)
    if not df.empty:
        df.sort_values(["게시일", "제목"], ascending=[False, True], inplace=True)
        df.reset_index(drop=True, inplace=True)
    return df


def fetch_all_extra(
    target_date: Optional[datetime.date] = None,
) -> dict[str, pd.DataFrame]:
    if target_date is None:
        target_date = yesterday_kst()
    print(f"[extra_sources] target_date={target_date.isoformat()} (KST)", flush=True)
    print("[extra_sources] K-Startup...", flush=True)
    ks = fetch_kstartup(target_date, lookback_days=0, open_only=True)
    print("[extra_sources] MSIT...", flush=True)
    # include day-before for timezone edge cases
    ms = fetch_msit(target_date, lookback_days=1)
    # Prefer exact yesterday for MSIT when available; keep lookback rows tagged by date
    return {
        "K-Startup": kstartup_to_dataframe(ks),
        "과기정통부": msit_to_dataframe(ms),
    }


if __name__ == "__main__":
    dfs = fetch_all_extra(datetime.date(2026, 9, 8))
    for name, df in dfs.items():
        print(f"{name}: {len(df)} rows")
        title_col = "공고명" if "공고명" in df.columns else "제목"
        if not df.empty and title_col in df.columns:
            for t in df[title_col].head(3).tolist():
                print(f"  - {t}")
