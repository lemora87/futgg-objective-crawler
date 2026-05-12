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


VALUE_KO = {
    "French": "프랑스",
    "France": "프랑스",
    "Ligue 1": "리그앙",
    "NWSL": "NWSL",
    "ROSHN Saudi League": "사우디 리그",
    "ROSHN Saudi Pro League": "사우디 리그",
    "Saudi Pro League": "사우디 리그",
    "Saudi Arabia": "사우디아라비아",
    "TOTS": "TOTS 카드",
    "Trendyol Süper Lig": "트렌디욜 쉬페르리그",
    "USA": "미국",
    "United States": "미국",
    "Germany": "독일",
    "German": "독일",
    "Bundesliga": "분데스리가",
    "D1 Arkema": "D1 아르케마",
    "defender": "수비수",
    "Defender": "수비수",
    "midfielder": "미드필더",
    "Midfielder": "미드필더",
    "attacker": "공격수",
    "Attacker": "공격수",
    "RW": "RW",
    "Any": "아무 선수",
}


def ko_value(value):
    v = clean(value)
    return VALUE_KO.get(v, v)

def needs_player_suffix(value):
    """
    득점/어시스트 표현에서 '~로' 앞에 '선수'가 붙어야 자연스러운 값.
    예: 미국 → 미국 선수로, 리그앙 → 리그앙 선수로
    """
    v = clean(value)

    no_suffix_values = {
        "수비수",
        "미드필더",
        "공격수",
        "GK",
        "CB",
        "LB",
        "RB",
        "CDM",
        "CM",
        "CAM",
        "LM",
        "RM",
        "LW",
        "RW",
        "ST",
        "아무 선수",
    }

    if not v:
        return False

    if v in no_suffix_values:
        return False

    if v.endswith("선수"):
        return False

    if v.endswith("카드"):
        return False

    return True


def as_actor(value):
    """
    미션 수행 주체 표현.
    예:
    미국 → 미국 선수
    리그앙 → 리그앙 선수
    TOTS 카드 → TOTS 카드
    수비수 → 수비수
    RW → RW
    아무 선수 → 아무 선수
    """
    v = ko_value(value)

    if needs_player_suffix(v):
        return f"{v} 선수"

    return v


def final_clean_text(text):
    """
    마지막 중복/조사 보정.
    """
    t = clean(text)

    fixes = {
        "TOTS 카드 카드": "TOTS 카드",
        "아무 선수 선수": "아무 선수",
        "프랑스 선수 선수": "프랑스 선수",
        "미국 선수 선수": "미국 선수",
        "리그앙 선수 선수": "리그앙 선수",
        "NWSL 선수 선수": "NWSL 선수",
        "사우디 리그 선수 선수": "사우디 리그 선수",
        "D1 아르케마 선수 선수": "D1 아르케마 선수",
        "트렌디욜 쉬페르리그 선수 선수": "트렌디욜 쉬페르리그 선수",
        "리그앙로": "리그앙 선수로",
        "NWSL로": "NWSL 선수로",
        "미국로": "미국 선수로",
        "프랑스로": "프랑스 선수로",
        "사우디 리그로": "사우디 리그 선수로",
        "D1 아르케마로": "D1 아르케마 선수로",
        "트렌디욜 쉬페르리그로": "트렌디욜 쉬페르리그 선수로",
    }

    for old, new in fixes.items():
        t = t.replace(old, new)

    return clean(t)

def strip_new_marker(text):
    return clean(text).replace("(신규미션)", "").replace("NEW", "").strip()


def normalize_campaign(campaign):
    """
    캠페인명과 마감은 그대로 유지하되, 표시용으로만 정리.
    """
    c = clean(campaign)
    c = c.replace("EA SPORTS FC 26 Objectives", "")
    c = c.replace("EA SPORTS FC 26", "")
    return clean(c)


def translate_play_requirement(text):
    t = strip_new_marker(text)

    m = re.fullmatch(r"Win\s+(\d+)", t, re.I)
    if m:
        return f"승리 {m.group(1)}회"

    m = re.fullmatch(r"Play\s+(\d+)", t, re.I)
    if m:
        return f"{m.group(1)}경기 플레이"

    m = re.fullmatch(r"Complete\s+(\d+)", t, re.I)
    if m:
        return f"{m.group(1)}회 완료"

    return translate_general_text(t)


def translate_starter(text):
    """
    French ≥1 → 프랑스 선수 1명 이상
    TOTS ≥1 → TOTS 카드 1명 이상
    """
    raw = strip_new_marker(text)

    m = re.fullmatch(r"(.+?)\s*≥\s*(\d+)", raw)
    if m:
        value = ko_value(m.group(1))
        count = m.group(2)

        if value.endswith("카드"):
            return f"{value} {count}명 이상"

        return f"{value} 선수 {count}명 이상"

    return translate_general_text(raw)


def translate_bench(text):
    """
    defender → 수비수
    D1 Arkema → D1 아르케마 선수
    RW → RW
    """
    raw = strip_new_marker(text)
    value = ko_value(raw)

    if value in {"수비수", "미드필더", "공격수"}:
        return value

    if value in {"RW", "LW", "ST", "CM", "CAM", "CDM", "LM", "RM", "LB", "RB", "CB", "GK"}:
        return value

    if value.endswith("선수"):
        return value

    return f"{value} 선수"


def translate_goal(text):
    """
    USA: Goal 3회 → 미국 선수로 3골
    TOTS: 5경기 각 Goal 1회 → TOTS 카드로 5경기 각각 1골
    Any: Goal 4회 → 아무 선수로 4골
    Ligue 1: 2경기 각 Goal 1회 → 리그앙 선수로 2경기 각각 1골
    """
    raw = strip_new_marker(text)

    m = re.fullmatch(r"(.+?):\s*(\d+)경기\s*각\s*Goal\s*1회", raw, re.I)
    if m:
        value = as_actor(m.group(1))
        count = m.group(2)
        return final_clean_text(f"{value}로 {count}경기 각각 1골")

    m = re.fullmatch(r"(.+?):\s*Goal\s*(\d+)회", raw, re.I)
    if m:
        value = as_actor(m.group(1))
        count = m.group(2)
        return final_clean_text(f"{value}로 {count}골")

    m = re.fullmatch(r"(.+?):\s*(\d+)경기\s*각\s*골\s*1회", raw, re.I)
    if m:
        value = as_actor(m.group(1))
        count = m.group(2)
        return final_clean_text(f"{value}로 {count}경기 각각 1골")

    return final_clean_text(translate_general_text(raw))


def translate_assist(text):
    """
    defender: Assist 4회 → 수비수로 4어시스트
    D1 Arkema: 2경기 각 Assist 1회 → D1 아르케마 선수로 2경기 각각 1어시스트
    French: Assist 4회 → 프랑스 선수로 4어시스트
    """
    raw = strip_new_marker(text)

    m = re.fullmatch(r"(.+?):\s*(\d+)경기\s*각\s*Assist\s*1회", raw, re.I)
    if m:
        value = as_actor(m.group(1))
        count = m.group(2)
        return final_clean_text(f"{value}로 {count}경기 각각 1어시스트")

    m = re.fullmatch(r"(.+?):\s*Assist\s*(\d+)회", raw, re.I)
    if m:
        value = as_actor(m.group(1))
        count = m.group(2)
        return final_clean_text(f"{value}로 {count}어시스트")

    return final_clean_text(translate_general_text(raw))


def translate_general_text(text):
    """
    이미 한국어화된 문장은 최대한 보존하고, 남아 있는 영어/기호만 보정.
    """
    t = strip_new_marker(text)

    # 긴 값부터 먼저 치환
    for eng in sorted(VALUE_KO.keys(), key=len, reverse=True):
        t = t.replace(eng, VALUE_KO[eng])

    t = re.sub(r"(\S+)\s*≥\s*(\d+)", r"\1 선수 \2명 이상", t)
    t = t.replace("Goal", "골")
    t = t.replace("Assist", "어시스트")
    t = t.replace("Win", "승리")
    t = t.replace("Play", "경기 플레이")
    t = t.replace("Complete", "완료")
    t = t.replace("Clean Sheet", "클린시트")

    # 어색한 반복 보정
    t = t.replace("TOTS 카드 선수", "TOTS 카드")
    t = t.replace("아무 선수 선수", "아무 선수")
    t = t.replace("리그앙 선수 선수", "리그앙 선수")
    t = t.replace("미국 선수 선수", "미국 선수")
    t = t.replace("프랑스 선수 선수", "프랑스 선수")

    return final_clean_text(t)


def translate_by_section(section_name, text):
    if section_name == "플레이 요건":
        return translate_play_requirement(text)

    if section_name == "필수 선발 요건":
        return translate_starter(text)

    if section_name == "교체로 투입해도 가능":
        return translate_bench(text)

    if section_name == "To do (득점)":
        return translate_goal(text)

    if section_name == "To do (어시스트)":
        return translate_assist(text)

    if section_name == "To do (기타)":
        return translate_general_text(text)

    return translate_general_text(text)


def collect_section(sb_rows, section_name):
    out = []
    current = ""

    for row in sb_rows:
        section = clean(get_any(row, ["구분", "Section"]))
        content = clean(get_any(row, ["내용", "Content"]))
        campaign = normalize_campaign(get_any(row, ["관련 캠페인", "Campaign"]))

        if section:
            current = section

        if current == section_name and content and content != "-":
            is_new = "신규미션" in content or "NEW" in content.upper()
            out.append({
                "text": translate_by_section(section_name, content),
                "raw_text": content,
                "campaign": campaign,
                "is_new": is_new
            })

    return out


def get_summary_value(sb_rows, label):
    for row in sb_rows:
        section = clean(get_any(row, ["구분", "Section"]))
        content = clean(get_any(row, ["내용", "Content"]))
        if section == label and content:
            return translate_by_section(label, content)
    return ""


def detect_index_status(rows, report_prefix):
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
    play_requirements_raw = collect_section(sb_rows, "플레이 요건")
    play_requirements = [x["text"] for x in play_requirements_raw]

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
        "play_requirements": play_requirements,
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
