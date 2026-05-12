from pathlib import Path
from datetime import timezone, timedelta
import csv
import os
import re
from PIL import Image, ImageDraw, ImageFont

KST = timezone(timedelta(hours=9))

ROOT = Path(__file__).resolve().parent

# crawler가 outputs/latest_csv 안에 만든 CSV를 읽음
CSV_DIR = ROOT / "outputs" / "latest_csv"

# 기존 운영 결과물과 분리하기 위한 테스트용 포스터 폴더
POSTER_DIR = ROOT / "outputs" / "latest_poster_test"
POSTER_DIR.mkdir(parents=True, exist_ok=True)

OUT_PATH = POSTER_DIR / "fc26_mission_guide_test.png"


def read_csv(name):
    path = CSV_DIR / name
    if not path.exists():
        print(f"[poster skipped] missing csv: {path}")
        return []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def clean(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def split_lines(v):
    text = str(v or "").replace("\r", "\n")
    return [clean(x) for x in text.split("\n") if clean(x) and clean(x) != "-"]


def find_font(bold=False):
    candidates = [
        "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf" if bold else "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc" if bold else "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    raise FileNotFoundError("No usable font found")


FONT_REG = find_font(False)
FONT_BOLD = find_font(True)


def font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)


def fit_text(draw, text, max_width, size, bold=False, min_size=16):
    s = size
    while s >= min_size:
        f = font(s, bold)
        if draw.textlength(text, font=f) <= max_width:
            return f
        s -= 1
    return font(min_size, bold)


def translate_symbolic(text):
    t = clean(text)

    # 기존 CSV가 영어/기호식을 일부 포함할 수 있어 최소 보정
    replacements = {
        "≥": "명 이상",
        "Goal": "골",
        "Assist": "어시스트",
        "Win": "승",
        "Play": "경기 플레이",
        "Complete": "완료",
        "Clean Sheet": "클린시트",
        "Any": "아무 선수",
        "Germany": "독일",
        "France": "프랑스",
        "USA": "미국",
        "Saudi Arabia": "사우디아라비아",
        "Ligue 1": "리그앙",
        "Bundesliga": "분데스리가",
        "Midfielder": "미드필더",
        "Defender": "수비수",
        "Attacker": "공격수",
    }

    for k, v in replacements.items():
        t = t.replace(k, v)

    # "미국명 이상2" 같은 극단적 변형 방지용은 추후 필요 시 보강
    return t


def collect_section(sb_rows, section_name):
    out = []
    current = ""

    # SB_Report 구조: 구분 / 내용 / 관련 캠페인
    for r in sb_rows:
        section = clean(r.get("구분") or r.get("Section") or "")
        content = clean(r.get("내용") or r.get("Content") or "")
        campaign = clean(r.get("관련 캠페인") or r.get("Campaign") or "")

        if section:
            current = section

        if current == section_name and content and content != "-":
            out.append({
                "text": translate_symbolic(content),
                "campaign": campaign,
                "is_new": "신규미션" in content or "NEW" in content.upper(),
            })

    return out


def get_summary_value(sb_rows, label):
    for r in sb_rows:
        section = clean(r.get("구분") or "")
        content = clean(r.get("내용") or "")
        if section == label:
            return content
    return ""


def draw_card(draw, x, y, w, h, number, text, sub="", is_new=False):
    draw.rounded_rectangle(
        [x, y, x + w, y + h],
        radius=9,
        fill=(245, 247, 250),
        outline=(170, 180, 200),
        width=1,
    )

    draw.rounded_rectangle(
        [x + 10, y + 13, x + 45, y + 48],
        radius=6,
        fill=(8, 26, 68),
    )
    draw.text((x + 21, y + 17), str(number), font=font(22, True), fill=(255, 255, 255))

    tx = x + 65
    maxw = w - 95

    if is_new:
        maxw -= 62

    main_font = fit_text(draw, text, maxw, 25, True, 17)
    draw.text((tx, y + 11), text, font=main_font, fill=(10, 15, 25))

    if is_new:
        draw.rounded_rectangle(
            [x + w - 66, y + 15, x + w - 15, y + 42],
            radius=5,
            fill=(220, 0, 20),
        )
        draw.text((x + w - 59, y + 18), "NEW", font=font(16, True), fill=(255, 255, 255))

    if sub:
        sub = sub.replace("(마감 ", "마감 ").replace(")", "")
        sub_font = fit_text(draw, sub, maxw, 14, False, 10)
        draw.text((tx, y + 47), sub, font=sub_font, fill=(35, 35, 35))


def draw_section(draw, x, y, w, title, items, icon, max_items=8):
    shown = min(len(items), max_items)
    section_h = 58 + shown * 72 + 10

    draw.rounded_rectangle(
        [x, y, x + w, y + section_h],
        radius=15,
        fill=(5, 20, 52),
        outline=(170, 220, 0),
        width=2,
    )
    draw.text((x + 28, y + 10), icon, font=font(30, True), fill=(190, 255, 0))
    draw.text((x + 78, y + 8), title, font=font(31, True), fill=(225, 255, 40))

    cy = y + 58
    for idx, item in enumerate(items[:max_items], start=1):
        draw_card(
            draw,
            x + 14,
            cy,
            w - 28,
            66,
            idx,
            item["text"],
            item.get("campaign", ""),
            is_new=item.get("is_new", False),
        )
        cy += 72

    if len(items) > max_items:
        draw.text(
            (x + 24, y + section_h - 24),
            f"외 {len(items) - max_items}개 항목 생략",
            font=font(15),
            fill=(180, 190, 210),
        )

    return section_h


def make_poster():
    sb = read_csv("SB_Report.csv")
    mission = read_csv("Mission_DB.csv")
    new_idx = read_csv("New_Report_Index.csv")
    combo_idx = read_csv("Combo_Report_Index.csv")

    if not sb:
        raise RuntimeError("SB_Report.csv를 읽지 못했습니다.")

    new_count = 0
    for r in mission:
        v = clean(r.get("Is_New") or r.get("신규여부")).lower()
        if v in {"yes", "y", "true", "1"}:
            new_count += 1

    if new_idx:
        joined = " ".join(clean(v) for row in new_idx for v in row.values())
        new_status = "신규 데이터 존재" if "New_Report" in joined and "없음" not in joined else "신규 미션 없음"
    else:
        new_status = "확인 불가"

    if combo_idx:
        joined = " ".join(clean(v) for row in combo_idx for v in row.values())
        combo_status = "신규 특수 모드 존재" if "Combo_Report" in joined and "없음" not in joined else "신규 특수 모드 없음"
    else:
        combo_status = "확인 불가"

    min_matches = get_summary_value(sb, "최소 필요경기") or "10경기"
    play_reqs = collect_section(sb, "플레이 요건")
    starters = collect_section(sb, "필수 선발 요건")
    bench = collect_section(sb, "교체로 투입해도 가능")
    goals = collect_section(sb, "To do (득점)")
    assists = collect_section(sb, "To do (어시스트)")
    others = collect_section(sb, "To do (기타)")

    W, H = 1400, 2000
    img = Image.new("RGB", (W, H), (3, 8, 25))
    draw = ImageDraw.Draw(img)

    # 배경
    draw.rectangle([0, 0, W, H], fill=(3, 8, 25))
    for i in range(8):
        draw.line(
            [(W - 340 + i * 35, 0), (W - 620 + i * 50, 220)],
            fill=(0, 120, 170),
            width=3,
        )

    # 좌상단 로고 박스
    draw.rounded_rectangle(
        [35, 30, 160, 165],
        radius=8,
        fill=(5, 20, 30),
        outline=(150, 255, 0),
        width=4,
    )
    draw.text((60, 55), "FC26", font=font(31, True), fill=(255, 255, 255))
    draw.text((69, 105), "⚽", font=font(38, True), fill=(190, 255, 0))

    # 타이틀
    draw.text((210, 28), "FC 26", font=font(88, True), fill=(255, 255, 255))
    draw.text((500, 28), "미션 가이드", font=font(88, True), fill=(205, 255, 0))

    draw.rounded_rectangle(
        [190, 130, W - 190, 178],
        radius=22,
        fill=(7, 30, 90),
        outline=(30, 70, 160),
        width=2,
    )
    draw.text(
        (360, 138),
        "Squad Battles / Rivals / Champions / Live Events 기준",
        font=font(31, True),
        fill=(255, 255, 255),
    )

    # 상단 요약
    draw.rounded_rectangle(
        [150, 200, W - 150, 260],
        radius=28,
        fill=(4, 10, 30),
        outline=(150, 160, 190),
        width=2,
    )
    draw.text((205, 214), f"▣ 최소 필요 경기: {min_matches}", font=font(30, True), fill=(255, 255, 255))

    play_summary = " / ".join([x["text"] for x in play_reqs[:2]]) if play_reqs else "승리/플레이 조건 확인"
    play_font = fit_text(draw, f"🏆 플레이 요건: {play_summary}", 610, 30, True, 21)
    draw.text((650, 214), f"🏆 플레이 요건: {play_summary}", font=play_font, fill=(255, 255, 255))

    # 본문
    left_x, right_x = 20, 710
    top_y = 285
    col_w = 670

    h1 = draw_section(draw, left_x, top_y, col_w, "누구를 선발로?", starters, "👥", max_items=8)
    h2 = draw_section(draw, left_x, top_y + h1 + 18, col_w, "교체로 투입해도 가능", bench, "🔄", max_items=4)

    h3 = draw_section(draw, right_x, top_y, col_w, "누구로 수행?", goals, "🎯", max_items=8)
    h4 = draw_section(draw, right_x, top_y + h3 + 18, col_w, "To do (어시스트)", assists, "👟", max_items=5)

    bottom_y = max(top_y + h1 + h2 + 36, top_y + h3 + h4 + 36)
    bottom_h = min(430, H - bottom_y - 140)

    draw.rounded_rectangle(
        [20, bottom_y, W - 20, bottom_y + bottom_h],
        radius=15,
        fill=(35, 10, 80),
        outline=(160, 120, 255),
        width=2,
    )
    draw.text((60, bottom_y + 10), "📋 To do (기타)", font=font(32, True), fill=(255, 255, 255))

    cy = bottom_y + 65
    card_w = (W - 70) // 2

    for idx, item in enumerate(others[:10], start=1):
        x = 35 if idx <= 5 else 35 + card_w + 15
        y = cy + ((idx - 1) % 5) * 66
        draw_card(
            draw,
            x,
            y,
            card_w,
            60,
            idx,
            item["text"],
            item.get("campaign", ""),
            is_new=item.get("is_new", False),
        )

    # 하단
    footer_y = H - 115
    draw.rounded_rectangle(
        [20, footer_y, W - 20, H - 25],
        radius=20,
        fill=(4, 12, 36),
        outline=(40, 80, 160),
        width=2,
    )
    draw.text(
        (65, footer_y + 18),
        "💡 모든 세부 목표는 지정된 조건을 만족한 상태로 진행해야 인정됩니다.",
        font=font(25, True),
        fill=(255, 255, 255),
    )
    draw.text(
        (65, footer_y + 55),
        "모드: 스쿼드 배틀 / 디비전 라이벌 / 챔피언스 / 라이브 이벤트",
        font=font(22),
        fill=(220, 230, 245),
    )

    if new_count:
        draw.rounded_rectangle([W - 310, footer_y + 26, W - 215, footer_y + 72], radius=8, fill=(230, 0, 20))
        draw.text((W - 294, footer_y + 34), "NEW", font=font(24, True), fill=(255, 255, 255))
        draw.text((W - 195, footer_y + 34), f"{new_count}개", font=font(24, True), fill=(255, 255, 255))
    else:
        draw.rounded_rectangle(
            [W - 330, footer_y + 20, W - 55, footer_y + 75],
            radius=14,
            fill=(8, 18, 45),
            outline=(120, 140, 180),
            width=1,
        )
        draw.text((W - 275, footer_y + 33), "신규 미션 없음", font=font(25, True), fill=(255, 255, 255))

    img.save(OUT_PATH)

    print(f"신규여부 Y 개수: {new_count}개")
    print(f"New_Report_Index status: {new_status}")
    print(f"Combo_Report_Index status: {combo_status}")
    print("포스터 기준 시트: SB_Report.csv")
    print(f"Poster saved: {OUT_PATH}")


if __name__ == "__main__":
    make_poster()
