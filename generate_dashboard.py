# -*- coding: utf-8 -*-
"""
서울동부지사 정비사업 뉴스+고시공고 대시보드 자동 생성기 (v7)
- 뉴스: 네이버 검색 API (최근 30일, 7개 구)
- 고시공고: 구청별 공식 고시공고 게시판 바로가기 탭 제공
- 최근 실거래: 국토부 실거래가 API로 구별 아파트 매매/전세/월세 (계약일 기준 최근 7일)
- 결과를 docs/index.html 에 저장 (GitHub Pages 배포)

필요 라이브러리: requests
"""

import os
import re
import html
import json
import requests
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

# ──────────────────────────────────────────────
# 공통 설정
# ──────────────────────────────────────────────
NAVER_CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")

# 대시보드 상단 공지줄 (비우면 표시 안 됨). 내용 수정 후 커밋하면 다음 갱신에 반영
UPDATE_NOTICE = "🆕 2026-07-10 · '최근 실거래' 탭 신설 — 구별 아파트 매매/전세/월세 최근 7일 계약분 제공"

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
# 고시공고: 구청 공식 게시판 바로가기 (링크는 2026-07 확인)
# ──────────────────────────────────────────────
NOTICE_BOARDS = {
    "성동구": {
        "name": "고시공고/입법예고 게시판",
        "url": "https://www.sd.go.kr/main/selectBbsNttList.do?bbsNo=184&key=1473",
    },
    "광진구": {
        "name": "고시공고 게시판",
        "url": "https://www.gwangjin.go.kr/portal/bbs/B0000003/list.do?menuNo=200192",
    },
    "동대문구": {
        "name": "고시공고 게시판",
        "url": "https://www.ddm.go.kr/www/selectEminwonWebList.do?key=3291&searchNotAncmtSeCode=01,02,04,05,06,07",
    },
    "중랑구": {
        "name": "공고/고시 게시판",
        "url": "https://www.jungnang.go.kr/portal/bbs/list/B0000117.do?menuNo=200475",
    },
    "도봉구": {
        "name": "고시/공고 게시판",
        "url": "https://www.dobong.go.kr/WDB_DEV/gosigong_go/",
    },
    "노원구": {
        "name": "고시공고 게시판",
        "url": "https://www.nowon.kr/www/user/bbs/BD_selectBbsList.do?q_bbsCode=1003&q_clCode=0&q_estnColumn1=11&q_ntceSiteCode=11",
    },
    "강북구": {
        "name": "고시공고 게시판",
        "url": "https://www.gangbuk.go.kr/portal/bbs/B0000245/list.do?menuNo=200082",
    },
}


# ──────────────────────────────────────────────
# 최근 실거래 (국토교통부 실거래가 공개 API)
# ──────────────────────────────────────────────
MOLIT_KEY = os.environ.get("DATA_GO_KR_KEY", "")  # 공공데이터포털 인증키(Decoding)

LAWD_CD = {  # 법정동 시군구코드
    "성동구": "11200", "광진구": "11215", "동대문구": "11230", "중랑구": "11260",
    "도봉구": "11320", "노원구": "11350", "강북구": "11305",
}

DEAL_DAYS_BACK = 7          # 계약일 기준 최근 며칠
MAX_DEALS_PER_DISTRICT = 40 # 구별 표시 상한

TRADE_URL = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev"
RENT_URL = "https://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent"

# 카카오맵: REST 키는 좌표 변환(서버측), JS 키는 지도 표시(페이지측)
KAKAO_REST_KEY = os.environ.get("KAKAO_REST_KEY", "")
KAKAO_JS_KEY = os.environ.get("KAKAO_JS_KEY", "")
GEO_CACHE_PATH = os.path.join("data", "geocache.json")


def _load_geocache() -> dict:
    try:
        with open(GEO_CACHE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_geocache(cache: dict):
    os.makedirs(os.path.dirname(GEO_CACHE_PATH), exist_ok=True)
    with open(GEO_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)


def _kakao_geocode(query: str, keyword: bool = False):
    url = ("https://dapi.kakao.com/v2/local/search/keyword.json" if keyword
           else "https://dapi.kakao.com/v2/local/search/address.json")
    try:
        r = requests.get(url, params={"query": query},
                         headers={"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}, timeout=10)
        r.raise_for_status()
        docs = r.json().get("documents", [])
        if docs:
            return float(docs[0]["y"]), float(docs[0]["x"])  # (lat, lng)
    except Exception as e:
        print(f"    [지오코딩 실패] {query}: {type(e).__name__}")
    return None


def apply_geocoding(deals: dict):
    """거래 주소를 좌표로 변환해 각 항목에 lat/lng 부여 (캐시 활용)"""
    if not KAKAO_REST_KEY:
        print("▶ KAKAO_REST_KEY 미설정 — 지도 좌표 변환 생략")
        return
    cache = _load_geocache()
    new_cnt = 0
    for district, items in deals.items():
        for x in items:
            key = f"{district}|{x.get('dong','')}|{x.get('jibun','')}|{x.get('apt','')}"
            if key in cache:
                coord = cache[key]
            else:
                addr = f"서울특별시 {district} {x.get('dong','')} {x.get('jibun','')}".strip()
                coord = _kakao_geocode(addr)
                if coord is None and x.get("apt"):
                    coord = _kakao_geocode(f"{district} {x['apt']}", keyword=True)
                cache[key] = coord
                new_cnt += 1
            if coord:
                x["lat"], x["lng"] = coord
    _save_geocache(cache)
    done = sum(1 for v in deals.values() for x in v if "lat" in x)
    print(f"▶ 지도 좌표 변환: 총 {done}건 표시 가능 (신규 변환 {new_cnt}건, 캐시 재사용)")



def _txt(node, tag):
    el = node.find(tag)
    return el.text.strip() if el is not None and el.text else ""


def _num(s):
    try:
        return int(s.replace(",", "").strip() or 0)
    except ValueError:
        return 0


def _money(man: int) -> str:
    """만원 단위 → '12억 5,000' 형태"""
    eok, rest = divmod(man, 10000)
    if eok and rest:
        return f"{eok}억 {rest:,}"
    if eok:
        return f"{eok}억"
    return f"{rest:,}"


def _fetch_deal_xml(url: str, lawd: str, ymd: str) -> list:
    import xml.etree.ElementTree as ET
    params = {"serviceKey": MOLIT_KEY, "LAWD_CD": lawd, "DEAL_YMD": ymd,
              "numOfRows": "1000", "pageNo": "1"}
    try:
        r = requests.get(url, params=params, timeout=15)
        r.raise_for_status()
        root = ET.fromstring(r.content)
        code = root.findtext(".//resultCode") or ""
        if code not in ("00", "000"):
            print(f"    [API 응답 오류] {root.findtext('.//resultMsg')}")
            return []
        return root.findall(".//item")
    except Exception as e:
        print(f"    [실거래 조회 실패] {type(e).__name__}")
        return []


def collect_deals(today: datetime) -> dict:
    cutoff = today - timedelta(days=DEAL_DAYS_BACK)
    result = {d: [] for d in DISTRICTS}
    if not MOLIT_KEY:
        print("▶ DATA_GO_KR_KEY 미설정 — 실거래 수집 생략")
        return result

    # 이번 달 + (기간이 지난달에 걸치면) 지난달 조회
    months = {today.strftime("%Y%m")}
    months.add(cutoff.strftime("%Y%m"))

    for district in DISTRICTS:
        lawd = LAWD_CD[district]
        print(f"▶ {district} 실거래 수집 중...")
        deals = []
        for ymd in sorted(months):
            # 매매
            for it in _fetch_deal_xml(TRADE_URL, lawd, ymd):
                try:
                    day = datetime(int(_txt(it, "dealYear")), int(_txt(it, "dealMonth")), int(_txt(it, "dealDay")))
                except ValueError:
                    continue
                if not (cutoff <= day <= today):
                    continue
                deals.append({
                    "type": "매매", "date": day,
                    "apt": _txt(it, "aptNm") or _txt(it, "offiNm"),
                    "dong": _txt(it, "umdNm"),
                    "jibun": _txt(it, "jibun"),
                    "area": _txt(it, "excluUseAr"),
                    "floor": _txt(it, "floor"),
                    "price": _money(_num(_txt(it, "dealAmount"))),
                })
            # 전월세
            for it in _fetch_deal_xml(RENT_URL, lawd, ymd):
                try:
                    day = datetime(int(_txt(it, "dealYear")), int(_txt(it, "dealMonth")), int(_txt(it, "dealDay")))
                except ValueError:
                    continue
                if not (cutoff <= day <= today):
                    continue
                monthly = _num(_txt(it, "monthlyRent"))
                deposit = _num(_txt(it, "deposit"))
                deals.append({
                    "type": "월세" if monthly else "전세", "date": day,
                    "apt": _txt(it, "aptNm"),
                    "dong": _txt(it, "umdNm"),
                    "jibun": _txt(it, "jibun"),
                    "area": _txt(it, "excluUseAr"),
                    "floor": _txt(it, "floor"),
                    "price": (f"{_money(deposit)}/{monthly:,}" if monthly else _money(deposit)),
                })
        deals.sort(key=lambda x: x["date"], reverse=True)
        result[district] = deals[:MAX_DEALS_PER_DISTRICT]
        print(f"  → 실거래 {len(result[district])}건 (매매/전세/월세)")
    return result


def build_deal_card(district: str, deals: list, today: datetime) -> str:
    if not deals:
        rows = '<div class="deal-row"><span class="deal-empty">최근 7일 계약분 없음 (신고 지연분은 이후 반영될 수 있음)</span></div>'
    else:
        rows = ""
        for x in deals:
            d = x["date"]
            area = x["area"].split(".")[0] + "㎡" if x["area"] else ""
            rows += (f'<div class="deal-row">'
                     f'<span class="deal-type deal-{x["type"]}">{x["type"]}</span>'
                     f'<span class="deal-name">{html.escape(x["dong"])} {html.escape(x["apt"])}</span>'
                     f'<span class="deal-spec">{area} {x["floor"]}층</span>'
                     f'<span class="deal-price">{x["price"]}</span>'
                     f'<span class="deal-date">{d.month}/{d.day}</span>'
                     f'</div>')
    return f"""
        <div class="notion-card deal-card" data-type="deal" data-district="{district}">
            <div class="card-meta">
                <span class="tag district-tag">📍 {district}</span>
                <span class="tag deal-tag">🏠 아파트 실거래 {len(deals)}건</span>
            </div>
            <div class="deal-list">{rows}</div>
            <div class="card-footer">출처: 국토교통부 실거래가 공개시스템 · 계약일 기준 최근 {DEAL_DAYS_BACK}일 (금액 단위: 만원)</div>
        </div>"""


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


def build_notice_card(district: str) -> str:
    b = NOTICE_BOARDS[district]
    return f"""
        <div class="notion-card" data-type="notice" data-district="{district}">
            <div class="card-meta">
                <span class="tag district-tag">📍 {district}</span>
                <span class="tag notice-tag">📢 고시공고</span>
            </div>
            <h3 class="news-title"><a href="{b['url']}" target="_blank">{district}청 {b['name']}</a></h3>
            <div class="summary-box">• 구청 공식 홈페이지 고시공고 게시판으로 이동합니다.<br>• 정비사업 관련 고시문 원문과 첨부파일(PDF/HWP)을 게시판에서 확인하세요.</div>
            <div class="card-footer"><a href="{b['url']}" target="_blank">게시판 바로가기 →</a></div>
        </div>"""


def build_html(news: dict, deals: dict, today: datetime) -> str:
    counts = {"news": {"all": sum(len(v) for v in news.values()),
                       **{d: len(news[d]) for d in DISTRICTS}},
              "deal": {"all": sum(len(v) for v in deals.values()),
                       **{d: len(deals[d]) for d in DISTRICTS}}}

    all_news = sorted((a for v in news.values() for a in v), key=lambda x: x["date"], reverse=True)
    cards = "".join(build_news_card(a) for a in all_news) + \
            "".join(build_notice_card(d) for d in DISTRICTS) + \
            "".join(build_deal_card(d, deals[d], today) for d in DISTRICTS)

    sidebar = ['<div class="sidebar-item active" data-district="all">🌐 전체</div>']
    sidebar += [f'<div class="sidebar-item" data-district="{d}">📍 {d}</div>' for d in DISTRICTS]

    deal_points = [
        {"d": dist, "t": x["type"], "n": f"{x['dong']} {x['apt']}",
         "p": x["price"], "dt": f"{x['date'].month}/{x['date'].day}",
         "lat": x["lat"], "lng": x["lng"]}
        for dist, v in deals.items() for x in v if "lat" in x
    ]
    deals_json = json.dumps(deal_points, ensure_ascii=False)

    notice_bar = (f'<div class="update-bar">📌 <span>{html.escape(UPDATE_NOTICE)}</span></div>'
                  if UPDATE_NOTICE.strip() else "")
    date_str = today.strftime("%Y-%m-%d")
    period_str = (today - timedelta(days=DAYS_BACK)).strftime("%Y-%m-%d")
    updated_str = today.strftime("%Y-%m-%d %H:%M")

    page = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="theme-color" content="#fbfbfa">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-title" content="동부 AI toolkit">
    <link rel="manifest" href="manifest.json">
    <title>서울동부지사 AI toolkit</title>
    <style>
        body {{ margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Pretendard, "Malgun Gothic", sans-serif; background-color: #fbfbfa; color: #37352f; }}
        #header {{ padding: 40px 50px 0 50px; background-color: #fbfbfa; border-bottom: 1px solid #edf2fa; }}
        #header h1 {{ font-size: 28px; font-weight: 700; margin: 0 0 8px 0; }}
        #header .subtitle {{ color: #73726e; font-size: 14px; margin-bottom: 12px; }}
        .update-bar {{ background-color: #fdf6e3; border: 1px solid #f0e2b6; color: #6b5a1e; font-size: 13px; padding: 7px 12px; border-radius: 6px; margin-bottom: 14px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
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
        .deal-card {{ grid-column: 1 / -1; }}
        .deal-list {{ background-color: #f7f7f5; border-radius: 6px; padding: 6px 10px; margin-bottom: 12px; }}
        .deal-row {{ display: flex; align-items: center; gap: 10px; padding: 7px 4px; border-bottom: 1px solid #ececea; font-size: 13.5px; flex-wrap: wrap; }}
        .deal-row:last-child {{ border-bottom: none; }}
        .deal-type {{ font-size: 11.5px; font-weight: 700; padding: 2px 7px; border-radius: 4px; flex-shrink: 0; }}
        .deal-매매 {{ background-color: #fbe4e4; color: #9c3a3a; }}
        .deal-전세 {{ background-color: #e3ecf7; color: #1f497d; }}
        .deal-월세 {{ background-color: #e8f3ec; color: #2b6a46; }}
        .deal-name {{ font-weight: 600; min-width: 0; }}
        .deal-spec {{ color: #73726e; flex-shrink: 0; }}
        .deal-price {{ font-weight: 700; margin-left: auto; flex-shrink: 0; }}
        .deal-date {{ color: #acaba9; font-size: 12px; width: 34px; text-align: right; flex-shrink: 0; }}
        .deal-empty {{ color: #73726e; font-size: 13px; }}
        .deal-tag {{ background-color: #f3e8f7; color: #6b3f85; }}
        #deal-map-wrap {{ display: none; margin-bottom: 20px; }}
        #deal-map {{ width: 100%; height: 360px; border-radius: 8px; border: 1px solid #e9e9e6; }}
        .map-legend {{ font-size: 12.5px; color: #73726e; margin-top: 6px; }}
        .lg-매매 {{ color: #d94343; }} .lg-전세 {{ color: #2f6bd8; }} .lg-월세 {{ color: #2b8a4e; }}

        @media (max-width: 768px) {{
            #header {{ padding: 24px 16px 0 16px; }}
            #header h1 {{ font-size: 20px; }}
            #header .subtitle {{ font-size: 12px; }}
            .update-bar {{ font-size: 12px; padding: 6px 10px; }}
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
            #deal-map {{ height: 260px; }}
            .notion-card {{ padding: 16px; }}
            .notion-card:hover {{ transform: none; }}
        }}
    </style>
</head>
<body>

    <div id="header">
        <h1>서울동부지사 AI toolkit</h1>
        <div class="subtitle">📅 {date_str} 기준 · 최근 {DAYS_BACK}일 ({period_str} ~) · 갱신 {updated_str} KST · by heychoi</div>
        {notice_bar}
        <div class="tab-bar">
            <div class="tab-btn active" data-tab="news">📰 뉴스</div>
            <div class="tab-btn" data-tab="notice">📢 고시공고</div>\n            <div class="tab-btn" data-tab="deal">🏠 최근 실거래</div>
        </div>
    </div>

    <div id="container">
        <div id="sidebar">
            <div class="sidebar-title">관할 자치구</div>
            <div id="sidebar-items">{"".join(sidebar)}</div>
        </div>
        <div id="main-content">
            <div class="view-bar" id="view-bar"></div>
            <div id="deal-map-wrap">
                <div id="deal-map"></div>
                <div class="map-legend"><span class="lg lg-매매">● 매매</span> <span class="lg lg-전세">● 전세</span> <span class="lg lg-월세">● 월세</span> — 핀을 누르면 상세 표시</div>
            </div>
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
                s.textContent = tab === 'notice' ? name : name + ' (' + (COUNTS[tab][s.dataset.district] || 0) + ')';
            }});
            updateDealMap();
            document.getElementById('view-bar').textContent =
                tab === 'news' ? '📋 뉴스 갤러리 — 최신순' : tab === 'notice' ? '📋 구청별 고시공고 게시판 바로가기' : '📋 구별 아파트 실거래 — 계약일 기준 최근 7일';
        }}

        document.querySelectorAll('.tab-btn').forEach(b =>
            b.addEventListener('click', () => {{ tab = b.dataset.tab; render(); }}));
        document.querySelectorAll('.sidebar-item').forEach(s =>
            s.addEventListener('click', () => {{ district = s.dataset.district; render(); }}));
        __DEAL_MAP_JS__
        render();
    </script>
</body>
</html>"""
    page = page.replace("__DEAL_MAP_JS__", DEAL_MAP_JS if KAKAO_JS_KEY else "function updateDealMap(){}")
    page = page.replace("__DEALS__", deals_json).replace("__KAKAO_JS_KEY__", KAKAO_JS_KEY)
    return page


DEAL_MAP_JS = r"""
        const DEALS = __DEALS__;
        const PIN_COLOR = {"매매": "#d94343", "전세": "#2f6bd8", "월세": "#2b8a4e"};
        let map = null, markers = [], infoWin = null, sdkLoaded = false;

        function pinImage(t) {
            const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18">' +
                '<circle cx="9" cy="9" r="7" fill="' + PIN_COLOR[t] + '" stroke="white" stroke-width="2.5"/></svg>';
            return new kakao.maps.MarkerImage('data:image/svg+xml;utf8,' + encodeURIComponent(svg),
                new kakao.maps.Size(18, 18), {offset: new kakao.maps.Point(9, 9)});
        }

        function initMap() {
            map = new kakao.maps.Map(document.getElementById('deal-map'),
                {center: new kakao.maps.LatLng(37.60, 127.06), level: 7});
            infoWin = new kakao.maps.InfoWindow({removable: true});
            DEALS.forEach(x => {
                const m = new kakao.maps.Marker({
                    position: new kakao.maps.LatLng(x.lat, x.lng), image: pinImage(x.t)});
                kakao.maps.event.addListener(m, 'click', () => {
                    infoWin.setContent('<div style="padding:8px 12px;font-size:12.5px;line-height:1.5;">' +
                        '<b>' + x.n + '</b><br>' + x.t + ' ' + x.p + ' · ' + x.dt + ' 계약</div>');
                    infoWin.open(map, m);
                });
                m._d = x.d; markers.push(m);
            });
        }

        function updateDealMap() {
            const wrap = document.getElementById('deal-map-wrap');
            wrap.style.display = (tab === 'deal' && DEALS.length) ? 'block' : 'none';
            if (tab !== 'deal' || !DEALS.length) return;
            if (!sdkLoaded) return;  // SDK 로드 후 재호출됨
            if (!map) initMap();
            map.relayout();
            const bounds = new kakao.maps.LatLngBounds();
            let visible = 0;
            markers.forEach(m => {
                const show = district === 'all' || m._d === district;
                m.setMap(show ? map : null);
                if (show) { bounds.extend(m.getPosition()); visible++; }
            });
            if (infoWin) infoWin.close();
            if (visible) map.setBounds(bounds, 30);
        }

        (function loadKakao() {
            if (!DEALS.length) return;
            const s = document.createElement('script');
            s.src = 'https://dapi.kakao.com/v2/maps/sdk.js?appkey=__KAKAO_JS_KEY__&autoload=false';
            s.onload = () => kakao.maps.load(() => { sdkLoaded = true; updateDealMap(); });
            document.head.appendChild(s);
        })();
"""


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
    deals = collect_deals(today)
    apply_geocoding(deals)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(build_html(news, deals, today))
    total_news = sum(len(v) for v in news.values())
    total_deals = sum(len(v) for v in deals.values())
    print(f"\n✅ 생성 완료: {OUTPUT_PATH} (뉴스 {total_news}건 / 실거래 {total_deals}건 / 게시판 바로가기 {len(DISTRICTS)}개)")


if __name__ == "__main__":
    main()
