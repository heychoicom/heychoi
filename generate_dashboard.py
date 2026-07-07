# -*- coding: utf-8 -*-
"""
서울동부지사 정비사업 뉴스 대시보드 자동 생성기 (v2 — 모바일/PWA 대응)
- 실행일 기준 최근 30일간 관할 자치구별 정비사업 관련 네이버 뉴스 수집
- GitHub Pages 배포용으로 docs/index.html 에 결과 저장
- GitHub Actions로 매일 자동 실행 (.github/workflows/daily.yml 참고)

로컬 테스트:
  set NAVER_CLIENT_ID=발급받은ID          (Windows cmd)
  set NAVER_CLIENT_SECRET=발급받은시크릿
  python generate_dashboard.py

※ GitHub에 올릴 때는 키를 코드에 절대 직접 쓰지 마세요.
   저장소 Settings → Secrets and variables → Actions 에 등록합니다.
"""

import os
import re
import html
import requests
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

# ──────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────
NAVER_CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")

DISTRICTS = ["성동구", "광진구", "동대문구", "중랑구", "도봉구", "노원구", "강북구"]

KEYWORDS = ["정비사업", "재개발", "재건축", "재정비", "모아타운", "신속통합기획", "공공주택 복합"]

DAYS_BACK = 30
MAX_PER_QUERY = 30
MAX_PER_DISTRICT = 15
OUTPUT_PATH = os.path.join("docs", "index.html")  # GitHub Pages 배포 경로

RELEVANCE_WORDS = KEYWORDS + ["조합", "관리처분", "사업시행", "안전진단", "이주", "착공",
                              "분양", "시공사", "정비구역", "조합설립", "리모델링"]

KST = timezone(timedelta(hours=9))

# ──────────────────────────────────────────────
# 수집
# ──────────────────────────────────────────────
API_URL = "https://openapi.naver.com/v1/search/news.json"
HEADERS = {
    "X-Naver-Client-Id": NAVER_CLIENT_ID,
    "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
}

TAG_RE = re.compile(r"</?b>")


def clean(text: str) -> str:
    return html.unescape(TAG_RE.sub("", text)).strip()


def fetch_news(district: str, keyword: str) -> list:
    params = {"query": f"{district} {keyword}", "display": MAX_PER_QUERY, "sort": "date"}
    try:
        r = requests.get(API_URL, headers=HEADERS, params=params, timeout=10)
        r.raise_for_status()
        return r.json().get("items", [])
    except Exception as e:
        print(f"  [경고] '{district} {keyword}' 검색 실패: {e}")
        return []


def collect(today: datetime) -> dict:
    cutoff = today - timedelta(days=DAYS_BACK)
    seen_links = set()
    result = {d: [] for d in DISTRICTS}

    for district in DISTRICTS:
        print(f"▶ {district} 수집 중...")
        articles = []
        for kw in KEYWORDS:
            for item in fetch_news(district, kw):
                link = item.get("originallink") or item.get("link", "")
                naver_link = item.get("link", "")
                if not link or link in seen_links or naver_link in seen_links:
                    continue
                try:
                    pub = parsedate_to_datetime(item["pubDate"]).astimezone(KST).replace(tzinfo=None)
                except Exception:
                    continue
                if pub < cutoff:
                    continue

                title = clean(item.get("title", ""))
                desc = clean(item.get("description", ""))
                if district not in title + desc:
                    continue
                if not any(w in title + desc for w in RELEVANCE_WORDS):
                    continue

                seen_links.add(link)
                seen_links.add(naver_link)
                articles.append({
                    "district": district, "title": title,
                    "link": naver_link or link, "summary": desc, "date": pub,
                })

        articles.sort(key=lambda a: a["date"], reverse=True)
        result[district] = articles[:MAX_PER_DISTRICT]
        print(f"  → {len(result[district])}건 채택")

    return result


# ──────────────────────────────────────────────
# HTML 생성
# ──────────────────────────────────────────────
def make_summary_bullets(summary: str) -> str:
    parts = re.split(r"(?<=[.다요음됨함])\s+", summary)
    parts = [p.strip().rstrip(".") for p in parts if len(p.strip()) > 10][:2]
    if not parts:
        parts = [summary[:120]]
    return "<br>".join(f"• {html.escape(p)}" for p in parts)


def build_card(a: dict) -> str:
    d = a["date"]
    date_str = f"{d.year}년 {d.month}월 {d.day}일"
    return f"""
            <div class="notion-card" data-district="{a['district']}">
                <div class="card-meta">
                    <span class="tag district-tag">📍 {a['district']}</span>
                    <span class="tag source-tag">📰 네이버뉴스</span>
                    <span class="tag date-tag">📅 {date_str}</span>
                </div>
                <h3 class="news-title">
                    <a href="{a['link']}" target="_blank">{html.escape(a['title'])}</a>
                </h3>
                <div class="summary-box">
                    {make_summary_bullets(a['summary'])}
                </div>
                <div class="card-footer">
                    <a href="{a['link']}" target="_blank">상세 원문 보기 →</a>
                </div>
            </div>
            """


def build_html(data: dict, today: datetime) -> str:
    total = sum(len(v) for v in data.values())

    sidebar_items = [
        f'<div class="sidebar-item active" onclick="filterDistrict(\'all\', this)">🌐 전체 ({total})</div>'
    ]
    for d in DISTRICTS:
        sidebar_items.append(
            f'<div class="sidebar-item" onclick="filterDistrict(\'{d}\', this)">📍 {d} ({len(data[d])})</div>'
        )

    all_articles = sorted(
        (a for arts in data.values() for a in arts),
        key=lambda a: a["date"], reverse=True,
    )
    cards = "".join(build_card(a) for a in all_articles)

    date_str = today.strftime("%Y-%m-%d")
    period_str = (today - timedelta(days=DAYS_BACK)).strftime("%Y-%m-%d")
    updated_str = today.strftime("%Y-%m-%d %H:%M")

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="theme-color" content="#fbfbfa">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-title" content="정비사업뉴스">
    <link rel="manifest" href="manifest.json">
    <title>서울동부지사 관할 지역 정비사업 뉴스</title>
    <style>
        body {{
            margin: 0; padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Pretendard, "Malgun Gothic", sans-serif;
            background-color: #fbfbfa; color: #37352f;
        }}
        #header {{ padding: 40px 50px 20px 50px; background-color: #fbfbfa; border-bottom: 1px solid #edf2fa; }}
        #header h1 {{ font-size: 28px; font-weight: 700; margin: 0 0 8px 0; color: #37352f; }}
        #header .subtitle {{ color: #73726e; font-size: 14px; }}
        #container {{ display: flex; padding: 0 50px 50px 50px; }}
        #sidebar {{ width: 200px; padding-right: 25px; border-right: 1px solid #eaeaea; margin-top: 30px; flex-shrink: 0; }}
        .sidebar-title {{ font-size: 11px; font-weight: 600; color: #acaba9; margin-bottom: 12px; padding-left: 10px; letter-spacing: 0.5px; }}
        .sidebar-item {{ padding: 8px 12px; font-size: 14px; border-radius: 6px; color: #37352f; margin-bottom: 4px; cursor: pointer; transition: background 0.2s; }}
        .sidebar-item:hover {{ background-color: #f1f1ef; }}
        .sidebar-item.active {{ background-color: #ececed; font-weight: 600; }}
        #main-content {{ flex: 1; padding-left: 40px; margin-top: 30px; min-width: 0; }}
        .view-bar {{ font-size: 14px; font-weight: 500; margin-bottom: 20px; border-bottom: 1px solid #eaeaea; padding-bottom: 8px; color: #37352f; }}
        .card-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 20px; }}
        .notion-card {{ background: #ffffff; border: 1px solid #e9e9e6; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); display: flex; flex-direction: column; justify-content: space-between; transition: transform 0.2s, box-shadow 0.2s; }}
        .notion-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.06); }}
        .card-meta {{ display: flex; gap: 8px; margin-bottom: 12px; align-items: center; flex-wrap: wrap; }}
        .tag {{ font-size: 12px; font-weight: 500; padding: 3px 8px; border-radius: 4px; white-space: nowrap; }}
        .district-tag {{ background-color: #edf2fa; color: #1f497d; }}
        .source-tag {{ background-color: #fbe4e4; color: #9c3a3a; }}
        .date-tag {{ background-color: #f1f1ef; color: #5a5a57; }}
        .news-title {{ font-size: 16px; font-weight: 600; margin: 0 0 14px 0; line-height: 1.4; }}
        .news-title a {{ color: #2383e2; text-decoration: none; }}
        .news-title a:hover {{ text-decoration: underline; }}
        .summary-box {{ background-color: #f7f7f5; padding: 14px; border-radius: 6px; font-size: 13.5px; line-height: 1.6; color: #37352f; margin-bottom: 14px; flex-grow: 1; }}
        .card-footer {{ font-size: 13px; margin-top: 8px; border-top: 1px dashed #eaeaea; padding-top: 10px; }}
        .card-footer a {{ color: #2383e2; text-decoration: none; font-weight: 500; display: inline-block; cursor: pointer; }}
        .card-footer a:hover {{ text-decoration: underline; color: #1a5fa8; }}

        /* ── 모바일 (폭 768px 이하) ── */
        @media (max-width: 768px) {{
            #header {{ padding: 24px 16px 14px 16px; }}
            #header h1 {{ font-size: 20px; }}
            #header .subtitle {{ font-size: 12px; }}
            #container {{ flex-direction: column; padding: 0 16px 30px 16px; }}
            /* 사이드바 → 상단 가로 스크롤 탭 */
            #sidebar {{
                width: 100%; padding-right: 0; border-right: none;
                border-bottom: 1px solid #eaeaea; margin-top: 12px; padding-bottom: 10px;
                position: sticky; top: 0; background-color: #fbfbfa; z-index: 10;
            }}
            .sidebar-title {{ display: none; }}
            #sidebar-items {{ display: flex; overflow-x: auto; gap: 6px; -webkit-overflow-scrolling: touch; scrollbar-width: none; }}
            #sidebar-items::-webkit-scrollbar {{ display: none; }}
            .sidebar-item {{ white-space: nowrap; margin-bottom: 0; background-color: #f1f1ef; font-size: 13px; padding: 7px 12px; border-radius: 16px; }}
            .sidebar-item.active {{ background-color: #37352f; color: #ffffff; }}
            #main-content {{ padding-left: 0; margin-top: 16px; }}
            .card-grid {{ grid-template-columns: 1fr; gap: 14px; }}
            .notion-card {{ padding: 16px; }}
            .notion-card:hover {{ transform: none; }}
        }}
    </style>

    <script>
        function filterDistrict(districtName, element) {{
            const items = document.querySelectorAll('.sidebar-item');
            items.forEach(item => item.classList.remove('active'));
            element.classList.add('active');

            const cards = document.querySelectorAll('.notion-card');
            cards.forEach(card => {{
                if (districtName === 'all') {{
                    card.style.display = 'flex';
                }} else {{
                    card.style.display = (card.getAttribute('data-district') === districtName) ? 'flex' : 'none';
                }}
            }});
        }}
    </script>
</head>
<body>

    <div id="header">
        <h1>서울동부지사 관할 지역 정비사업 뉴스</h1>
        <div class="subtitle">📅 {date_str} 기준 · 최근 {DAYS_BACK}일 ({period_str} ~) · 총 {total}건 · 갱신 {updated_str} KST</div>
    </div>

    <div id="container">
        <div id="sidebar">
            <div class="sidebar-title">관할 자치구</div>
            <div id="sidebar-items">
            {"".join(sidebar_items)}
            </div>
        </div>

        <div id="main-content">
            <div class="view-bar">📋 갤러리 보기 — 최신순</div>
            <div class="card-grid">
                {cards}
            </div>
        </div>
    </div>

</body>
</html>"""


# ──────────────────────────────────────────────
# 실행
# ──────────────────────────────────────────────
def main():
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        raise SystemExit("환경변수 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이 설정되지 않았습니다.")

    today = datetime.now(KST).replace(tzinfo=None)
    data = collect(today)
    html_out = build_html(data, today)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"\n✅ 생성 완료: {OUTPUT_PATH} (총 {sum(len(v) for v in data.values())}건)")


if __name__ == "__main__":
    main()
