# -*- coding: utf-8 -*-
"""
서울동부지사 정비사업 뉴스+고시공고 대시보드 자동 생성기 (v3)
- 뉴스: 네이버 검색 API (최근 30일, 7개 구)
- 고시공고: 각 구청 새올 전자민원(eminwon) 표준 시스템에서 정비사업 관련 고시 수집
- 결과를 docs/index.html 에 저장 (GitHub Pages 배포)

필요 라이브러리: requests, beautifulsoup4
"""

import os
import re
import html
import json
import requests
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

# ──────────────────────────────────────────────
# 공통 설정
# ──────────────────────────────────────────────
NAVER_CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")

DISTRICTS = ["성동구", "광진구", "동대문구", "중랑구", "도봉구", "노원구", "강북구"]
KEYWORDS = ["정비사업", "재개발", "재건축", "재정비", "모아타운", "신속통합기획", "공공주택 복합"]

DAYS_BACK = 30
MAX_PER_QUERY = 30
MAX_PER_DISTRICT = 15
OUTPUT_PATH = os.path.join("docs", "index.html")

RELEVANCE_WORDS = KEYWORDS + ["조합", "관리처분", "사업시행", "안전진단", "이주", "착공",
                              "분양", "시공사", "정비구역", "조합설립", "리모델링"]

# 고시 제목에 이 단어가 있으면 정비사업 관련 고시로 채택
NOTICE_KEYWORDS = ["정비", "재개발", "재건축", "모아타운", "신속통합", "리모델링",
                   "관리처분", "사업시행", "조합", "정비구역", "도시계획", "지구단위"]

KST = timezone(timedelta(hours=9))

# ──────────────────────────────────────────────
# 고시공고 수집 설정 (새올 전자민원 표준시스템)
# ──────────────────────────────────────────────
# 후보 호스트를 순서대로 시도해서 처음 성공하는 곳을 사용
EMINWON_HOSTS = {
    "성동구":   ["seongdong.eminwon.seoul.kr", "sd.eminwon.seoul.kr", "eminwon.sd.go.kr"],
    "광진구":   ["gwangjin.eminwon.seoul.kr", "eminwon.gwangjin.go.kr"],
    "동대문구": ["ddm.eminwon.seoul.kr", "dongdaemun.eminwon.seoul.kr", "eminwon.ddm.go.kr"],
    "중랑구":   ["eminwon.jungnang.go.kr", "jungnang.eminwon.seoul.kr"],
    "도봉구":   ["dobong.eminwon.seoul.kr", "eminwon.dobong.go.kr"],
    "노원구":   ["nowon.eminwon.seoul.kr", "eminwon.nowon.kr", "eminwon.nowon.go.kr"],
    "강북구":   ["gangbuk.eminwon.seoul.kr", "eminwon.gangbuk.go.kr"],
}

# 수집 실패 시 대신 안내할 링크 (확인된 게시판 주소 or 네이버 검색)
FALLBACK_BOARDS = {
    "성동구": "https://www.sd.go.kr/main/selectBbsNttList.do?bbsNo=184&key=1473",
    "광진구": "https://www.gwangjin.go.kr/portal/bbs/B0000003/list.do?menuNo=200192",
    "도봉구": "https://www.dobong.go.kr/wdb_dev/gosigong_go/default.asp",
}
for _d in DISTRICTS:
    FALLBACK_BOARDS.setdefault(_d, f"https://search.naver.com/search.naver?query={_d}+고시공고")

EMINWON_PATH = "/emwp/gov/mogaha/ntis/web/ofr/action/OfrAction.do"
NOTICE_PAGES = 4          # 목록 몇 페이지까지 볼지 (페이지당 30건)
MAX_NOTICE_PER_DISTRICT = 15

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

DATE_RE = re.compile(r"(20\d{2})[.\-/]\s?(\d{1,2})[.\-/]\s?(\d{1,2})")
MGT_RE = re.compile(r"(?:not_ancmt_mgt_no=|searchDetail\(\s*['\"]?)(\d{3,})")
DEPT_RE = re.compile(r"^[가-힣0-9·\s]{2,15}(과|국|소|센터|담당관|사업소|보건소)$")


def _decode(resp) -> str:
    """eminwon은 EUC-KR인 경우가 많아 인코딩을 자동 판별"""
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return resp.content.decode(enc)
        except UnicodeDecodeError:
            continue
    return resp.text


def _list_params(page: int) -> dict:
    return {
        "jndinm": "OfrNotAncmtEJB", "context": "NTIS",
        "method": "selectListOfrNotAncmt",
        "methodnm": "selectListOfrNotAncmtHomepage",
        "homepage_pbs_yn": "Y", "subCheck": "Y",
        "ofr_pageSize": "30", "pageIndex": str(page),
        "not_ancmt_se_code": "01,02,03,04,05,06",
        "title": "고시공고", "initValue": "Y", "countYn": "Y",
        "list_gubun": "A", "Key": "B_Subject",
    }


def _detail_url(base: str, mgt_no: str) -> str:
    return (f"{base}{EMINWON_PATH}?jndinm=OfrNotAncmtEJB&context=NTIS"
            f"&method=selectOfrNotAncmt&methodnm=selectOfrNotAncmtRegst"
            f"&not_ancmt_mgt_no={mgt_no}&homepage_pbs_yn=Y&subCheck=Y"
            f"&ofr_pageSize=10&not_ancmt_se_code=01,02,03,04,05,06"
            f"&title=고시공고&initValue=Y&countYn=Y&Key=B_Subject")


def _parse_notice_rows(html_text: str, base: str) -> list:
    """새올 목록 페이지의 표에서 (제목, 부서, 날짜, 상세링크) 추출"""
    soup = BeautifulSoup(html_text, "html.parser")
    items = []
    for tr in soup.find_all("tr"):
        row_html = str(tr)
        m_mgt = MGT_RE.search(row_html)
        if not m_mgt:
            continue
        a = tr.find("a")
        if not a:
            continue
        title = a.get_text(" ", strip=True)
        if len(title) < 4:
            continue

        tds = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
        m_date = DATE_RE.search(" ".join(tds))
        if not m_date:
            continue
        try:
            pub = datetime(int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3)))
        except ValueError:
            continue

        dept = ""
        for td in tds:
            if DEPT_RE.match(td):
                dept = td
                break

        items.append({
            "title": title, "dept": dept, "date": pub,
            "link": _detail_url(base, m_mgt.group(1)),
        })
    return items


def fetch_notices_for_district(district: str, cutoff: datetime) -> list | None:
    """구별 정비사업 관련 고시 수집. 완전 실패 시 None 반환"""
    if BeautifulSoup is None:
        print("  [경고] beautifulsoup4 미설치 — 고시공고 수집 생략")
        return None

    for host in EMINWON_HOSTS[district]:
        # 구청 서버는 https 미지원(http 전용)인 곳이 많아 두 방식 모두 시도
        for scheme in ("https", "http"):
            base = f"{scheme}://{host}"
            collected, ok = [], False
            for page in range(1, NOTICE_PAGES + 1):
                try:
                    r = requests.get(f"{base}{EMINWON_PATH}",
                                     params=_list_params(page), headers=UA,
                                     timeout=12, verify=False)
                    r.raise_for_status()
                except Exception as e:
                    if page == 1:
                        print(f"    [{scheme}://{host}] 접속 실패: {type(e).__name__}")
                    break  # 이 방식 포기
                rows = _parse_notice_rows(_decode(r), base)
                if not rows:
                    if page == 1:
                        print(f"    [{scheme}://{host}] 접속됐으나 목록 파싱 0행")
                    break
                ok = True
                collected.extend(rows)
                if min(x["date"] for x in rows) < cutoff:
                    break  # 이 페이지에 이미 기간 밖 고시가 있으면 중단

            if ok:
                in_period = [n for n in collected if n["date"] >= cutoff]
                seen, result = set(), []
                for n in in_period:
                    if n["link"] in seen:
                        continue
                    if not any(k in n["title"] for k in NOTICE_KEYWORDS):
                        continue
                    seen.add(n["link"])
                    n["district"] = district
                    result.append(n)
                result.sort(key=lambda x: x["date"], reverse=True)
                print(f"  → [{scheme}://{host}] 목록 {len(collected)}행 / "
                      f"기간내 {len(in_period)}건 / 정비관련 {len(result)}건 채택")
                if in_period and result:
                    return result[:MAX_NOTICE_PER_DISTRICT]
                if in_period:
                    # 목록·날짜는 정상인데 정비 관련 고시가 없는 경우 → 0건으로 확정
                    return []
                # 기간내 0건은 날짜 파싱 오류일 수 있으므로 다음 후보도 시도
                print(f"    (기간내 고시가 없어 다음 후보 주소 확인)")

    print(f"  → 고시공고 수집 실패 (모든 호스트 불가)")
    return None


def collect_notices(today: datetime) -> dict:
    cutoff = today - timedelta(days=DAYS_BACK)
    result = {}
    for district in DISTRICTS:
        print(f"▶ {district} 고시공고 수집 중...")
        notices = fetch_notices_for_district(district, cutoff)
        if notices is None:
            # 실패 → 게시판 바로가기 안내 카드
            result[district] = [{
                "district": district, "title": f"{district} 고시공고 게시판 직접 확인",
                "dept": "자동 수집 실패", "date": today,
                "link": FALLBACK_BOARDS[district], "fallback": True,
            }]
        else:
            result[district] = notices
    return result


# ──────────────────────────────────────────────
# 뉴스 수집 (네이버 검색 API) — v2와 동일
# ──────────────────────────────────────────────
API_URL = "https://openapi.naver.com/v1/search/news.json"
NAVER_HEADERS = {
    "X-Naver-Client-Id": NAVER_CLIENT_ID,
    "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
}
TAG_RE = re.compile(r"</?b>")


def clean(text: str) -> str:
    return html.unescape(TAG_RE.sub("", text)).strip()


def fetch_news(district: str, keyword: str) -> list:
    params = {"query": f"{district} {keyword}", "display": MAX_PER_QUERY, "sort": "date"}
    try:
        r = requests.get(API_URL, headers=NAVER_HEADERS, params=params, timeout=10)
        r.raise_for_status()
        return r.json().get("items", [])
    except Exception as e:
        print(f"  [경고] '{district} {keyword}' 검색 실패: {e}")
        return []


def collect_news(today: datetime) -> dict:
    cutoff = today - timedelta(days=DAYS_BACK)
    seen_links = set()
    result = {d: [] for d in DISTRICTS}

    for district in DISTRICTS:
        print(f"▶ {district} 뉴스 수집 중...")
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
                articles.append({"district": district, "title": title,
                                 "link": naver_link or link, "summary": desc, "date": pub})
        articles.sort(key=lambda a: a["date"], reverse=True)
        result[district] = articles[:MAX_PER_DISTRICT]
        print(f"  → 뉴스 {len(result[district])}건 채택")
    return result


# ──────────────────────────────────────────────
# HTML 생성 (뉴스/고시공고 탭)
# ──────────────────────────────────────────────
def make_summary_bullets(summary: str) -> str:
    parts = re.split(r"(?<=[.다요음됨함])\s+", summary)
    parts = [p.strip().rstrip(".") for p in parts if len(p.strip()) > 10][:2]
    if not parts:
        parts = [summary[:120]]
    return "<br>".join(f"• {html.escape(p)}" for p in parts)


def build_news_card(a: dict) -> str:
    d = a["date"]
    return f"""
        <div class="notion-card" data-type="news" data-district="{a['district']}">
            <div class="card-meta">
                <span class="tag district-tag">📍 {a['district']}</span>
                <span class="tag source-tag">📰 네이버뉴스</span>
                <span class="tag date-tag">📅 {d.year}년 {d.month}월 {d.day}일</span>
            </div>
            <h3 class="news-title"><a href="{a['link']}" target="_blank">{html.escape(a['title'])}</a></h3>
            <div class="summary-box">{make_summary_bullets(a['summary'])}</div>
            <div class="card-footer"><a href="{a['link']}" target="_blank">상세 원문 보기 →</a></div>
        </div>"""


def build_notice_card(n: dict) -> str:
    d = n["date"]
    dept = html.escape(n.get("dept") or "")
    dept_tag = f'<span class="tag dept-tag">🏢 {dept}</span>' if dept else ""
    fb = n.get("fallback")
    body = ("자동 수집에 실패해 구청 게시판으로 직접 연결합니다. 아래 링크에서 최신 고시를 확인하세요."
            if fb else "구청 고시공고 원문 페이지에서 상세 내용과 첨부파일(PDF/HWP)을 확인할 수 있습니다.")
    footer = "게시판 바로가기 →" if fb else "고시 원문·첨부파일 보기 →"
    return f"""
        <div class="notion-card{' fallback-card' if fb else ''}" data-type="notice" data-district="{n['district']}">
            <div class="card-meta">
                <span class="tag district-tag">📍 {n['district']}</span>
                <span class="tag notice-tag">📢 고시공고</span>
                {dept_tag}
                <span class="tag date-tag">📅 {d.year}년 {d.month}월 {d.day}일</span>
            </div>
            <h3 class="news-title"><a href="{n['link']}" target="_blank">{html.escape(n['title'])}</a></h3>
            <div class="summary-box">• {body}</div>
            <div class="card-footer"><a href="{n['link']}" target="_blank">{footer}</a></div>
        </div>"""


def build_html(news: dict, notices: dict, today: datetime) -> str:
    counts = {
        "news": {"all": sum(len(v) for v in news.values()),
                 **{d: len(news[d]) for d in DISTRICTS}},
        "notice": {"all": sum(len(v) for v in notices.values()),
                   **{d: len(notices[d]) for d in DISTRICTS}},
    }

    all_news = sorted((a for v in news.values() for a in v), key=lambda x: x["date"], reverse=True)
    all_notices = sorted((n for v in notices.values() for n in v), key=lambda x: x["date"], reverse=True)
    cards = "".join(build_news_card(a) for a in all_news) + \
            "".join(build_notice_card(n) for n in all_notices)

    sidebar = ['<div class="sidebar-item active" data-district="all">🌐 전체</div>']
    sidebar += [f'<div class="sidebar-item" data-district="{d}">📍 {d}</div>' for d in DISTRICTS]

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
    <title>서울동부지사 관할 지역 정비사업 뉴스·고시</title>
    <style>
        body {{ margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Pretendard, "Malgun Gothic", sans-serif; background-color: #fbfbfa; color: #37352f; }}
        #header {{ padding: 40px 50px 0 50px; background-color: #fbfbfa; border-bottom: 1px solid #edf2fa; }}
        #header h1 {{ font-size: 28px; font-weight: 700; margin: 0 0 8px 0; }}
        #header .subtitle {{ color: #73726e; font-size: 14px; margin-bottom: 16px; }}
        .tab-bar {{ display: flex; gap: 4px; }}
        .tab-btn {{ padding: 10px 18px; font-size: 15px; font-weight: 600; color: #73726e; cursor: pointer; border-bottom: 2px solid transparent; }}
        .tab-btn.active {{ color: #37352f; border-bottom-color: #37352f; }}
        #container {{ display: flex; padding: 0 50px 50px 50px; }}
        #sidebar {{ width: 200px; padding-right: 25px; border-right: 1px solid #eaeaea; margin-top: 30px; flex-shrink: 0; }}
        .sidebar-title {{ font-size: 11px; font-weight: 600; color: #acaba9; margin-bottom: 12px; padding-left: 10px; letter-spacing: 0.5px; }}
        .sidebar-item {{ padding: 8px 12px; font-size: 14px; border-radius: 6px; margin-bottom: 4px; cursor: pointer; transition: background 0.2s; }}
        .sidebar-item:hover {{ background-color: #f1f1ef; }}
        .sidebar-item.active {{ background-color: #ececed; font-weight: 600; }}
        #main-content {{ flex: 1; padding-left: 40px; margin-top: 30px; min-width: 0; }}
        .view-bar {{ font-size: 14px; font-weight: 500; margin-bottom: 20px; border-bottom: 1px solid #eaeaea; padding-bottom: 8px; }}
        .card-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 20px; }}
        .notion-card {{ background: #fff; border: 1px solid #e9e9e6; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); display: flex; flex-direction: column; justify-content: space-between; transition: transform 0.2s, box-shadow 0.2s; }}
        .notion-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.06); }}
        .fallback-card {{ border-style: dashed; background: #fffdf5; }}
        .card-meta {{ display: flex; gap: 8px; margin-bottom: 12px; align-items: center; flex-wrap: wrap; }}
        .tag {{ font-size: 12px; font-weight: 500; padding: 3px 8px; border-radius: 4px; white-space: nowrap; }}
        .district-tag {{ background-color: #edf2fa; color: #1f497d; }}
        .source-tag {{ background-color: #fbe4e4; color: #9c3a3a; }}
        .notice-tag {{ background-color: #fdecc8; color: #8a6116; }}
        .dept-tag {{ background-color: #e8f3ec; color: #2b6a46; }}
        .date-tag {{ background-color: #f1f1ef; color: #5a5a57; }}
        .news-title {{ font-size: 16px; font-weight: 600; margin: 0 0 14px 0; line-height: 1.4; }}
        .news-title a {{ color: #2383e2; text-decoration: none; }}
        .news-title a:hover {{ text-decoration: underline; }}
        .summary-box {{ background-color: #f7f7f5; padding: 14px; border-radius: 6px; font-size: 13.5px; line-height: 1.6; margin-bottom: 14px; flex-grow: 1; }}
        .card-footer {{ font-size: 13px; margin-top: 8px; border-top: 1px dashed #eaeaea; padding-top: 10px; }}
        .card-footer a {{ color: #2383e2; text-decoration: none; font-weight: 500; }}
        .card-footer a:hover {{ text-decoration: underline; }}

        @media (max-width: 768px) {{
            #header {{ padding: 24px 16px 0 16px; }}
            #header h1 {{ font-size: 20px; }}
            #header .subtitle {{ font-size: 12px; }}
            .tab-btn {{ flex: 1; text-align: center; padding: 10px 4px; font-size: 14px; }}
            #container {{ flex-direction: column; padding: 0 16px 30px 16px; }}
            #sidebar {{ width: 100%; padding-right: 0; border-right: none; border-bottom: 1px solid #eaeaea; margin-top: 12px; padding-bottom: 10px; position: sticky; top: 0; background-color: #fbfbfa; z-index: 10; }}
            .sidebar-title {{ display: none; }}
            #sidebar-items {{ display: flex; overflow-x: auto; gap: 6px; -webkit-overflow-scrolling: touch; scrollbar-width: none; }}
            #sidebar-items::-webkit-scrollbar {{ display: none; }}
            .sidebar-item {{ white-space: nowrap; margin-bottom: 0; background-color: #f1f1ef; font-size: 13px; padding: 7px 12px; border-radius: 16px; }}
            .sidebar-item.active {{ background-color: #37352f; color: #fff; }}
            #main-content {{ padding-left: 0; margin-top: 16px; }}
            .card-grid {{ grid-template-columns: 1fr; gap: 14px; }}
            .notion-card {{ padding: 16px; }}
            .notion-card:hover {{ transform: none; }}
        }}
    </style>
</head>
<body>

    <div id="header">
        <h1>서울동부지사 관할 지역 정비사업 뉴스·고시</h1>
        <div class="subtitle">📅 {date_str} 기준 · 최근 {DAYS_BACK}일 ({period_str} ~) · 갱신 {updated_str} KST</div>
        <div class="tab-bar">
            <div class="tab-btn active" data-tab="news">📰 뉴스</div>
            <div class="tab-btn" data-tab="notice">📢 고시공고</div>
        </div>
    </div>

    <div id="container">
        <div id="sidebar">
            <div class="sidebar-title">관할 자치구</div>
            <div id="sidebar-items">{"".join(sidebar)}</div>
        </div>
        <div id="main-content">
            <div class="view-bar" id="view-bar"></div>
            <div class="card-grid">{cards}</div>
        </div>
    </div>

    <script>
        const COUNTS = {json.dumps(counts, ensure_ascii=False)};
        let tab = 'news', district = 'all';

        function render() {{
            document.querySelectorAll('.notion-card').forEach(c => {{
                const show = c.dataset.type === tab && (district === 'all' || c.dataset.district === district);
                c.style.display = show ? 'flex' : 'none';
            }});
            document.querySelectorAll('.tab-btn').forEach(b =>
                b.classList.toggle('active', b.dataset.tab === tab));
            document.querySelectorAll('.sidebar-item').forEach(s => {{
                s.classList.toggle('active', s.dataset.district === district);
                const name = s.dataset.district === 'all' ? '🌐 전체' : '📍 ' + s.dataset.district;
                s.textContent = name + ' (' + (COUNTS[tab][s.dataset.district] || 0) + ')';
            }});
            document.getElementById('view-bar').textContent =
                tab === 'news' ? '📋 뉴스 갤러리 — 최신순' : '📋 정비사업 관련 고시공고 — 최신순';
        }}

        document.querySelectorAll('.tab-btn').forEach(b =>
            b.addEventListener('click', () => {{ tab = b.dataset.tab; render(); }}));
        document.querySelectorAll('.sidebar-item').forEach(s =>
            s.addEventListener('click', () => {{ district = s.dataset.district; render(); }}));
        render();
    </script>
</body>
</html>"""


# ──────────────────────────────────────────────
# 실행
# ──────────────────────────────────────────────
def main():
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        raise SystemExit("환경변수 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이 설정되지 않았습니다.")

    import urllib3
    urllib3.disable_warnings()  # 일부 구청 서버의 구형 SSL 인증서 경고 무시

    today = datetime.now(KST).replace(tzinfo=None)
    news = collect_news(today)
    notices = collect_notices(today)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(build_html(news, notices, today))
    total_news = sum(len(v) for v in news.values())
    total_notice = sum(len(v) for v in notices.values())
    print(f"\n✅ 생성 완료: {OUTPUT_PATH} (뉴스 {total_news}건 / 고시 {total_notice}건)")


if __name__ == "__main__":
    main()
