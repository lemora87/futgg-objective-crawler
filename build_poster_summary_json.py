from pathlib import Path
from datetime import datetime, timezone, timedelta
import csv
import json
import re

ROOT = Path(__file__).resolve().parent
CSV_DIR = ROOT / "outputs" / "latest_csv"
OUT_DIR = ROOT / "outputs" / "latest_json"
OUT_DIR.mkdir(parents=True, exist_ok=True)

OUT_PATH = OUT_DIR / "mission_poster_summary.json"

KST = timezone(timedelta(hours=9))


def clean(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def read_csv(name):
    path = CSV_DIR / name
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def boolish(v):
    return clean(v).lower() in {"y", "yes", "true", "1"}


def get_any(row, keys):
    for k in keys:
        if k in row and row[k] is not None:
            return row[k]
    return ""


def collect_section(sb_rows, section_name):
    out = []
    current = ""

    for row in sb_rows:
        section = clean(get_any(row, ["구분", "Section"]))
        content = clean(get_any(row, ["내용", "Content"]))
        campaign = clean(get_any(row, ["관련 캠페인", "Campaign"]))

        if section:
            current = section

        if current == section_name and content and content != "-":
            out.append({
                "text": content,
                "campaign": campaign,
                "is_new": False
            })

    return out


def get_summary_value(sb_rows, label):
    for row in sb_rows:
        section = clean(get_any(row, ["구분", "Section"]))
        content = clean(get_any(row, ["내용", "Content"]))
        if section == label and content:
            return content
    return ""


def detect_index_status(rows, report_prefix):
    """
    New_Report_Index / Combo_Report_Index 상태 판단용
    """
    if not rows:
        return "비어 있음"

    all_values = []
    for row in rows:
        for v in row.values():
            vv = clean(v)
            if vv:
                all_values.append(vv)

    joined = " ".join(all_values)

    if report_prefix in joined:
        return "데이터 존재"

    if "없음" in joined:
        return "없음"

    return "비어 있음"


def count_new_missions(mission_rows):
    count = 0
    for row in mission_rows:
        v = get_any(row, ["Is_New", "신규여부"])
        if boolish(v):
            count += 1
    return count


def main():
    sb_rows = read_csv("SB_Report.csv")
    mission_rows = read_csv("Mission_DB.csv")
    new_idx_rows = read_csv("New_Report_Index.csv")
    combo_idx_rows = read_csv("Combo_Report_Index.csv")

    if not sb_rows:
        raise RuntimeError("SB_Report.csv를 읽지 못했습니다.")

    new_count = count_new_missions(mission_rows)
    new_status = detect_index_status(new_idx_rows, "New_Report")
    combo_status = detect_index_status(combo_idx_rows, "Combo_Report")

    minimum_matches = get_summary_value(sb_rows, "최소 필요경기")
    play_requirements = collect_section(sb_rows, "플레이 요건")

    starters = collect_section(sb_rows, "필수 선발 요건")
    bench = collect_section(sb_rows, "교체로 투입해도 가능")
    goals = collect_section(sb_rows, "To do (득점)")
    assists = collect_section(sb_rows, "To do (어시스트)")
    others = collect_section(sb_rows, "To do (기타)")

    summary = {
        "generated_at_kst": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"),
        "source": "latest_csv",
        "new_count": new_count,
        "new_report_status": new_status,
        "combo_report_status": combo_status,
        "poster_basis": "SB_Report.csv",
        "minimum_matches": minimum_matches,
        "play_requirements": [x["text"] for x in play_requirements],
        "sections": {
            "필수 선발 요건": starters,
            "교체로 투입해도 가능": bench,
            "To do (득점)": goals,
            "To do (어시스트)": assists,
            "To do (기타)": others
        }
    }

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(f"신규여부 Y 개수: {new_count}개")
    print(f"New_Report_Index status: {new_status}")
    print(f"Combo_Report_Index status: {combo_status}")
    print("포스터 기준 시트: SB_Report.csv")
    print(f"JSON saved: {OUT_PATH}")


if __name__ == "__main__":
    main()
