from pathlib import Path
import json
import html
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent
JSON_PATH = ROOT / "outputs" / "latest_json" / "mission_poster_summary.json"
OUT_DIR = ROOT / "outputs" / "latest_poster"
OUT_DIR.mkdir(parents=True, exist_ok=True)

HTML_PATH = OUT_DIR / "fc26_mission_guide_latest.html"
PNG_PATH = OUT_DIR / "fc26_mission_guide_latest.png"


def esc(x):
    return html.escape(str(x or ""))


def build_cards(items, show_new=True):
    blocks = []
    for idx, item in enumerate(items, start=1):
        text = esc(item.get("text", ""))
        campaign = esc(item.get("campaign", ""))
        is_new = bool(item.get("is_new", False))

        new_badge = '<span class="new-badge">NEW</span>' if show_new and is_new else ""

        blocks.append(f"""
        <div class="card">
          <div class="num">{idx}</div>
          <div class="card-body">
            <div class="card-main">{text} {new_badge}</div>
            <div class="card-sub">{campaign}</div>
          </div>
        </div>
        """)
    return "\n".join(blocks)


def build_html(data):
    minimum_matches = esc(data.get("minimum_matches", ""))
    play_requirements = " / ".join(data.get("play_requirements", []))
    play_requirements = esc(play_requirements)

    new_count = int(data.get("new_count", 0))
    new_summary = "신규 미션 없음" if new_count == 0 else f"NEW {new_count}개"

    sections = data.get("sections", {})
    starters = sections.get("필수 선발 요건", [])
    bench = sections.get("교체로 투입해도 가능", [])
    goals = sections.get("To do (득점)", [])
    assists = sections.get("To do (어시스트)", [])
    others = sections.get("To do (기타)", [])

    html_doc = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<title>FC 26 미션 가이드</title>
<style>
  * {{
    box-sizing: border-box;
  }}
  body {{
    margin: 0;
    font-family: "Noto Sans KR", "Malgun Gothic", "Apple SD Gothic Neo", sans-serif;
    background: #050b22;
  }}
  .poster {{
    width: 1600px;
    min-height: 2300px;
    margin: 0 auto;
    background:
      radial-gradient(circle at 85% 0%, rgba(0,180,255,0.15), transparent 28%),
      radial-gradient(circle at 0% 100%, rgba(204,255,0,0.08), transparent 20%),
      linear-gradient(180deg, #07102d 0%, #040917 100%);
    color: #fff;
    padding: 42px 42px 36px 42px;
  }}
  .top {{
    display: flex;
    align-items: flex-start;
    gap: 28px;
  }}
  .logo {{
    width: 128px;
    height: 128px;
    border: 4px solid #d6ff00;
    border-radius: 16px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    font-weight: 800;
    background: rgba(5,20,40,0.8);
  }}
  .logo .small {{
    font-size: 34px;
    line-height: 1;
  }}
  .logo .ball {{
    font-size: 34px;
    line-height: 1.1;
    color: #d6ff00;
  }}
  .title-area {{
    flex: 1;
  }}
  .title {{
    font-size: 78px;
    font-weight: 900;
    line-height: 1;
    margin: 0;
  }}
  .title .lime {{
    color: #d6ff00;
  }}
  .subtitle {{
    margin-top: 20px;
    background: #0b2a75;
    border: 2px solid #2d59c6;
    border-radius: 28px;
    padding: 14px 24px;
    font-size: 30px;
    font-weight: 700;
    display: inline-block;
  }}
  .new-summary {{
    margin-left: auto;
    margin-top: 10px;
    border: 2px solid #74839f;
    border-radius: 20px;
    padding: 12px 24px;
    background: rgba(9,20,45,0.9);
    font-size: 28px;
    font-weight: 800;
    color: #edf3ff;
    white-space: nowrap;
  }}
  .summary-bar {{
    margin-top: 28px;
    display: flex;
    justify-content: space-between;
    gap: 24px;
    padding: 18px 28px;
    background: rgba(7,15,38,0.95);
    border: 2px solid #70809d;
    border-radius: 30px;
    font-size: 28px;
    font-weight: 800;
  }}
  .columns {{
    margin-top: 28px;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 26px;
  }}
  .panel {{
    background: #081a45;
    border: 3px solid #d6ff00;
    border-radius: 24px;
    padding: 18px;
  }}
  .panel.purple {{
    background: #271055;
    border-color: #b798ff;
  }}
  .panel-title {{
    font-size: 42px;
    font-weight: 900;
    color: #d6ff00;
    margin-bottom: 18px;
  }}
  .panel.purple .panel-title {{
    color: #ffffff;
  }}
  .subsection {{
    margin-top: 14px;
  }}
  .subsection-title {{
    font-size: 30px;
    font-weight: 800;
    margin: 10px 0 12px 4px;
    color: #f3f7ff;
  }}
  .card {{
    background: #f3f6fa;
    color: #0b1020;
    border: 1px solid #c8d1dd;
    border-radius: 14px;
    display: flex;
    gap: 14px;
    padding: 12px 14px;
    margin-bottom: 10px;
    align-items: flex-start;
  }}
  .num {{
    min-width: 36px;
    height: 36px;
    border-radius: 8px;
    background: #0b2a68;
    color: #fff;
    display: flex;
    justify-content: center;
    align-items: center;
    font-size: 21px;
    font-weight: 900;
    margin-top: 2px;
  }}
  .card-body {{
    flex: 1;
  }}
  .card-main {{
    font-size: 26px;
    font-weight: 800;
    line-height: 1.3;
  }}
  .card-sub {{
    margin-top: 5px;
    font-size: 16px;
    color: #677180;
    line-height: 1.35;
  }}
  .new-badge {{
    display: inline-block;
    margin-left: 8px;
    background: #df1020;
    color: #fff;
    border-radius: 8px;
    padding: 2px 10px;
    font-size: 15px;
    font-weight: 900;
    vertical-align: middle;
  }}
  .bottom {{
    margin-top: 26px;
  }}
  .footer {{
    margin-top: 28px;
    background: rgba(7,15,38,0.95);
    border: 2px solid #2d59c6;
    border-radius: 22px;
    padding: 18px 24px;
    font-size: 24px;
    color: #eef3ff;
    line-height: 1.5;
  }}
  .footer .small {{
    margin-top: 6px;
    font-size: 20px;
    color: #ced9eb;
  }}
</style>
</head>
<body>
<div class="poster">

  <div class="top">
    <div class="logo">
      <div class="small">FC26</div>
      <div class="ball">⚽</div>
    </div>

    <div class="title-area">
      <h1 class="title">FC 26 <span class="lime">미션 가이드</span></h1>
      <div class="subtitle">Squad Battles / Rivals / Champions / Live Events 기준</div>
    </div>

    <div class="new-summary">{new_summary}</div>
  </div>

  <div class="summary-bar">
    <div>최소 필요 경기: {minimum_matches}</div>
    <div>플레이 요건: {play_requirements}</div>
  </div>

  <div class="columns">
    <div>
      <div class="panel">
        <div class="panel-title">누구를 선발로?</div>

        <div class="subsection">
          <div class="subsection-title">필수 선발 요건</div>
          {build_cards(starters)}
        </div>

        <div class="subsection">
          <div class="subsection-title">교체로 투입해도 가능</div>
          {build_cards(bench)}
        </div>
      </div>
    </div>

    <div>
      <div class="panel">
        <div class="panel-title">누구로 수행?</div>

        <div class="subsection">
          <div class="subsection-title">To do (득점)</div>
          {build_cards(goals)}
        </div>

        <div class="subsection">
          <div class="subsection-title">To do (어시스트)</div>
          {build_cards(assists)}
        </div>
      </div>
    </div>
  </div>

  <div class="bottom">
    <div class="panel purple">
      <div class="panel-title">To do (기타)</div>
      {build_cards(others)}
    </div>
  </div>

  <div class="footer">
    모든 세부 목표는 지정된 조건을 만족한 상태로 진행해야 인정됩니다.
    <div class="small">기준: 스쿼드 배틀 / 디비전 라이벌 / 챔피언스 / 라이브 이벤트</div>
  </div>

</div>
</body>
</html>
"""
    return html_doc


def main():
    if not JSON_PATH.exists():
        raise FileNotFoundError(f"JSON not found: {JSON_PATH}")

    with open(JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    html_doc = build_html(data)
    HTML_PATH.write_text(html_doc, encoding="utf-8")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1600, "height": 2300}, device_scale_factor=1.5)
        page.goto(HTML_PATH.resolve().as_uri())
        page.screenshot(path=str(PNG_PATH), full_page=True)
        browser.close()

    print(f"Poster HTML saved: {HTML_PATH}")
    print(f"Poster PNG saved: {PNG_PATH}")


if __name__ == "__main__":
    main()
