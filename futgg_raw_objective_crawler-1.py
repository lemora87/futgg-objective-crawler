# futgg_raw_objective_crawler.py
# FUT.GG objective raw-text crawler
# Purpose:
#   - Open FUT.GG objective detail pages with Playwright
#   - Save the page body text as-is so ChatGPT/user can interpret manually
#   - Do NOT over-parse goal/assist/squad conditions in Python
#
# Run examples:
#   python futgg_raw_objective_crawler.py --out outputs --headless
#   python futgg_raw_objective_crawler.py --urls https://www.fut.gg/objectives/asia-oceania/928-republic-of-korea https://www.fut.gg/objectives/asia-oceania/929-saudi-arabia
#
# Requirements:
#   pip install pandas openpyxl playwright
#   python -m playwright install chromium

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

KST = timezone(timedelta(hours=9))
BASE_URL = "https://www.fut.gg"
DEFAULT_INDEX_URL = "https://www.fut.gg/objectives/asia-oceania/"

DEFAULT_DETAIL_URLS = [
    "https://www.fut.gg/objectives/asia-oceania/928-republic-of-korea",
    "https://www.fut.gg/objectives/asia-oceania/923-uzbekistanqatar",
    "https://www.fut.gg/objectives/asia-oceania/924-new-zealand",
    "https://www.fut.gg/objectives/asia-oceania/925-australia",
    "https://www.fut.gg/objectives/asia-oceania/926-perfect-volley",
    "https://www.fut.gg/objectives/asia-oceania/927-japan",
    "https://www.fut.gg/objectives/asia-oceania/929-saudi-arabia",
    "https://www.fut.gg/objectives/asia-oceania/930-asiaoceania",
]

NOISE_LINES = {
    "home", "objectives", "players", "sbc", "evolutions", "squads", "login", "sign up",
    "ea sports fc", "fut.gg", "fifa", "clubs", "rush", "database", "tools",
}

ACTION_START_RE = re.compile(
    r"^(Win|Play|Score|Assist|Complete|Keep|Earn|Make|Get|Perform|Record|Concede|Claim|Finish|Achieve|Submit|Watch)\b",
    re.I,
)

REWARD_HINT_RE = re.compile(
    r"\b(Pack|Player Pick|Players Pick|Player|SP|Coin|Coins|Coin Boost|Evo|Evolution|Unlock|XP|Token|Tokens|Loan|Tifo|Badge|Kit|Rare Gold|FoF|Festival of Football|OVR|Rated|Rating|Item)\b",
    re.I,
)


def clean_text(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def canonical_url(url: str) -> str:
    url = (url or "").split("?", 1)[0].rstrip("/")
    return url


def slug_title(url: str) -> str:
    slug = canonical_url(url).split("/")[-1]
    slug = re.sub(r"^\d+-", "", slug)
    return " ".join(w.capitalize() for w in slug.split("-"))


def split_body_lines(text: str) -> List[str]:
    out: List[str] = []
    seen = set()
    for raw in (text or "").replace("\r", "\n").split("\n"):
        line = clean_text(raw)
        if not line:
            continue
        if line.lower() in NOISE_LINES:
            continue
        # Keep duplicates if they are objective/reward-like, otherwise reduce obvious nav noise.
        if len(line) < 30 and line.lower() in seen:
            continue
        seen.add(line.lower())
        out.append(line)
    return out


def is_probable_objective_line(line: str) -> bool:
    t = clean_text(line)
    if not t:
        return False
    if not ACTION_START_RE.search(t):
        return False
    # objective descriptions usually mention match/game/mode/etc.
    return bool(re.search(
        r"\b(matches?|games?|goals?|assists?|wins?|clean sheets?|Squad Battles|Rivals|Champions|Live Events?|Rush|Ultimate Team|UT|FUT|starting 11|using|while having|difficulty|SBC|watch)\b",
        t,
        re.I,
    ))


def is_probable_reward_line(line: str) -> bool:
    t = clean_text(line)
    if not t or is_probable_objective_line(t):
        return False
    if len(t) > 120:
        return False
    return bool(REWARD_HINT_RE.search(t))


def is_probable_title_line(line: str) -> bool:
    t = clean_text(line)
    if not t or len(t) > 80:
        return False
    if is_probable_objective_line(t) or is_probable_reward_line(t):
        return False
    if re.search(r"Expires in|Objectives not released|Challenges|Rewards", t, re.I):
        return False
    return bool(re.search(r"[A-Za-z0-9]", t))


def extract_deadline_text(body: str) -> str:
    t = clean_text(body)
    m = re.search(r"\bExpires\s+in\s+((?:[0-9]+\s*(?:days?|hours?|hrs?|minutes?|mins?)\s*){1,4})", t, re.I)
    if m:
        return "Expires in " + clean_text(m.group(1))
    return ""


def objective_type_from_url(url: str) -> str:
    m = re.search(r"/objectives/([^/]+)/", url)
    return m.group(1) if m else "objectives"


def parse_loose_objectives(lines: List[str]) -> List[Dict[str, Any]]:
    """
    Loose extractor for convenience only.
    The authoritative source remains Page_Raw.Body_Text_Raw and Page_Lines.

    Logic:
      - Find full objective-like action lines.
      - Look back up to 5 lines for a likely short title.
      - Look around nearby lines for reward-like text.
    """
    rows: List[Dict[str, Any]] = []
    seen = set()

    for i, line in enumerate(lines):
        if not is_probable_objective_line(line):
            continue
        if line in seen:
            continue
        seen.add(line)

        title = ""
        for j in range(i - 1, max(-1, i - 6), -1):
            cand = lines[j]
            if is_probable_title_line(cand):
                title = cand
                break

        rewards: List[str] = []
        # Rewards often sit before or after the objective line on FUT.GG cards.
        for j in list(range(max(0, i - 4), i)) + list(range(i + 1, min(len(lines), i + 7))):
            cand = lines[j]
            if is_probable_reward_line(cand) and cand not in rewards:
                rewards.append(cand)

        rows.append({
            "Objective_Order": len(rows) + 1,
            "Objective_Name_Guess": title,
            "Objective_Text_Raw": line,
            "Reward_Text_Guess": " / ".join(rewards[:5]),
            "Extractor_Note": "loose_guess_verify_with_Page_Raw",
        })

    return rows


def unique_links(links: List[str]) -> List[str]:
    seen = set()
    out = []
    for link in links:
        c = canonical_url(link)
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def collect_detail_links_from_index(page, index_url: str) -> List[str]:
    page.goto(index_url, wait_until="networkidle", timeout=60000)
    time.sleep(2)
    hrefs = page.eval_on_selector_all("a[href]", "els => els.map(a => a.href)")
    # Any objective detail under the same objective category.
    base_path = re.search(r"(/objectives/[^/]+/)", index_url)
    if not base_path:
        return []
    path = base_path.group(1)
    urls = []
    for h in hrefs:
        u = canonical_url(h)
        if re.search(re.escape(path) + r"[0-9]+-[^/]+$", u):
            urls.append(u)
    return unique_links(urls)


def scrape_one_page(context, url: str, order: int, wait_ms: int = 1200) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]], str]:
    page = context.new_page()
    try:
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(wait_ms)

        # Try to ensure dynamic content appears.
        for _ in range(4):
            page.mouse.wheel(0, 900)
            page.wait_for_timeout(250)
        page.mouse.wheel(0, -3000)
        page.wait_for_timeout(250)

        body = page.inner_text("body", timeout=15000)
        lines = split_body_lines(body)

        title = ""
        try:
            h1s = page.locator("h1").all_inner_texts()
            if h1s:
                title = clean_text(h1s[0])
        except Exception:
            pass
        if not title:
            title = slug_title(url)

        deadline = extract_deadline_text(body)
        status = "OK"
        if re.search(r"Objectives not released yet", body, re.I):
            status = "OBJECTIVES_NOT_RELEASED_TEXT_FOUND"
        if not any(is_probable_objective_line(x) for x in lines):
            status = "NO_OBJECTIVE_LINES_DETECTED" if status == "OK" else status

        page_row = {
            "Group_Order": order,
            "Group_Name": title,
            "URL": url,
            "Objective_Type": objective_type_from_url(url),
            "Deadline_Text_Raw": deadline,
            "Status": status,
            "Line_Count": len(lines),
            "Body_Text_Raw": body,
            "Fetched_At_KST": datetime.now(KST).isoformat(timespec="seconds"),
        }

        line_rows = []
        for idx, line in enumerate(lines, start=1):
            line_rows.append({
                "Group_Order": order,
                "Group_Name": title,
                "URL": url,
                "Line_No": idx,
                "Line_Text": line,
                "Looks_Like_Objective": "Yes" if is_probable_objective_line(line) else "No",
                "Looks_Like_Reward": "Yes" if is_probable_reward_line(line) else "No",
            })

        objective_rows = []
        for r in parse_loose_objectives(lines):
            r.update({
                "Group_Order": order,
                "Group_Name": title,
                "URL": url,
                "Deadline_Text_Raw": deadline,
            })
            objective_rows.append(r)

        return page_row, line_rows, objective_rows, ""

    except Exception as e:
        err = f"{type(e).__name__}: {e}"
        page_row = {
            "Group_Order": order,
            "Group_Name": slug_title(url),
            "URL": url,
            "Objective_Type": objective_type_from_url(url),
            "Deadline_Text_Raw": "",
            "Status": "ERROR",
            "Line_Count": 0,
            "Body_Text_Raw": "",
            "Fetched_At_KST": datetime.now(KST).isoformat(timespec="seconds"),
        }
        return page_row, [], [], err
    finally:
        page.close()


def write_outputs(out_dir: Path, page_rows: List[Dict[str, Any]], line_rows: List[Dict[str, Any]], objective_rows: List[Dict[str, Any]], error_rows: List[Dict[str, Any]], source_urls: List[str]) -> Tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(KST).strftime("%Y%m%d_%H%M%S")
    xlsx_path = out_dir / f"futgg_raw_objectives_{stamp}.xlsx"
    json_path = out_dir / f"futgg_raw_objectives_{stamp}.json"

    df_groups = pd.DataFrame(page_rows)
    if not df_groups.empty:
        df_groups_summary = df_groups.drop(columns=["Body_Text_Raw"], errors="ignore")
    else:
        df_groups_summary = pd.DataFrame()

    df_lines = pd.DataFrame(line_rows)
    df_objectives = pd.DataFrame(objective_rows)
    df_errors = pd.DataFrame(error_rows)
    df_sources = pd.DataFrame({"Source_URL": source_urls})

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_groups_summary.to_excel(writer, sheet_name="Objective_Groups", index=False)
        df_objectives.to_excel(writer, sheet_name="Objective_Raw_Guess", index=False)
        df_lines.to_excel(writer, sheet_name="Page_Lines", index=False)
        df_groups.to_excel(writer, sheet_name="Page_Raw", index=False)
        df_sources.to_excel(writer, sheet_name="Source_URLs", index=False)
        df_errors.to_excel(writer, sheet_name="Errors", index=False)

        # Basic readability. Avoid over-styling; keep full raw text accessible.
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = 12
                for cell in col[:50]:
                    try:
                        max_len = max(max_len, min(80, len(str(cell.value or ""))))
                    except Exception:
                        pass
                ws.column_dimensions[letter].width = min(max_len + 2, 80)

    data = {
        "generated_at_kst": datetime.now(KST).isoformat(timespec="seconds"),
        "source_urls": source_urls,
        "groups": page_rows,
        "objectives_guess": objective_rows,
        "errors": error_rows,
        "note": "Objective_Raw_Guess is a loose convenience extraction. Page_Raw.Body_Text_Raw and Page_Lines are the source of truth.",
    }
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    # Stable latest copies.
    latest_xlsx = out_dir / "futgg_raw_objectives_latest.xlsx"
    latest_json = out_dir / "futgg_raw_objectives_latest.json"
    latest_xlsx.write_bytes(xlsx_path.read_bytes())
    latest_json.write_text(json_path.read_text(encoding="utf-8"), encoding="utf-8")

    return xlsx_path, json_path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default="outputs_raw", help="Output directory")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--index-url", default=DEFAULT_INDEX_URL, help="Index page to discover links from")
    parser.add_argument("--no-index", action="store_true", help="Do not discover links from index; use provided/default detail URLs only")
    parser.add_argument("--urls", nargs="*", default=None, help="Detail URLs to scrape")
    args = parser.parse_args()

    out_dir = Path(args.out)
    urls = args.urls if args.urls else list(DEFAULT_DETAIL_URLS)

    page_rows: List[Dict[str, Any]] = []
    line_rows: List[Dict[str, Any]] = []
    objective_rows: List[Dict[str, Any]] = []
    error_rows: List[Dict[str, Any]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(
            viewport={"width": 1500, "height": 1200},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
        )

        if not args.no_index:
            page = context.new_page()
            try:
                discovered = collect_detail_links_from_index(page, args.index_url)
                if discovered:
                    urls = unique_links(urls + discovered)
                    print(f"Discovered {len(discovered)} detail links from index.")
            except Exception as e:
                error_rows.append({"URL": args.index_url, "Stage": "index_discovery", "Error": f"{type(e).__name__}: {e}"})
            finally:
                page.close()

        urls = unique_links(urls)
        print(f"Scraping {len(urls)} objective detail pages...")
        for i, url in enumerate(urls, start=1):
            print(f"[{i}/{len(urls)}] {url}")
            page_row, rows_lines, rows_obj, err = scrape_one_page(context, url, i)
            page_rows.append(page_row)
            line_rows.extend(rows_lines)
            objective_rows.extend(rows_obj)
            if err:
                error_rows.append({"URL": url, "Stage": "detail", "Error": err})

        context.close()
        browser.close()

    xlsx_path, json_path = write_outputs(out_dir, page_rows, line_rows, objective_rows, error_rows, urls)
    print(f"Saved xlsx: {xlsx_path}")
    print(f"Saved json: {json_path}")
    print(f"Latest xlsx: {out_dir / 'futgg_raw_objectives_latest.xlsx'}")
    print(f"Latest json: {out_dir / 'futgg_raw_objectives_latest.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
