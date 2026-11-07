"""
대전 지역 공공기관 채용공고 자동 정리 프로그램

데이터 출처: 공공데이터포털 '기획재정부_공공기관 채용정보 조회서비스' (잡알리오)
  https://www.data.go.kr/data/15125273/openapi.do

정리 항목
  1. 채용기관명
  2. 채용 직무 분야 (NCS 직무분야)
  3. 서류 제출 마감일자 (접수 마감일)

실행 방법
  python daejeon_jobs.py
  -> 같은 폴더에 '대전_공공기관_채용공고.xlsx' 생성/갱신
"""

import json
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone
import html as htmllib
import re
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ───────────────────── 설정 ─────────────────────
# 공공데이터포털에서 발급받은 '일반 인증키(Decoding)'를 넣거나,
# 환경변수 JOB_API_KEY 로 지정하세요.
API_KEY = os.environ.get("JOB_API_KEY", "여기에_서비스키_입력")

BASE_URL = "http://apis.data.go.kr/1051000/recruitment/list"
REGION_CODE = "R3012"          # 근무지역 코드: 대전
ONLY_ONGOING = True            # True: 현재 접수 중인 공고만

BASE_DIR = Path(__file__).resolve().parent
# 결과 저장 폴더 (GitHub에서는 OUTPUT_DIR=docs 로 실행 → 웹페이지로 공개)
OUT_DIR = Path(os.environ.get("OUTPUT_DIR") or BASE_DIR)
if not OUT_DIR.is_absolute():
    OUT_DIR = BASE_DIR / OUT_DIR
OUT_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_NAME = "대전_공공기관_채용공고.xlsx"
EXCEL_PATH = OUT_DIR / EXCEL_NAME
SEEN_PATH = OUT_DIR / "seen_ids.json"   # 신규 공고 표시용 기록
HTML_PATH = OUT_DIR / "index.html"

KST = timezone(timedelta(hours=9))     # 서버가 해외에 있어도 한국 날짜 기준


def now_kst():
    return datetime.now(KST)


def today_kst():
    return now_kst().date()
# ────────────────────────────────────────────────


def pick(row: dict, *keys, default=""):
    """응답 필드명이 조금 달라도 동작하도록 후보 키 중 첫 값을 반환"""
    for k in keys:
        v = row.get(k)
        if v not in (None, ""):
            return str(v).strip()
    return default


def parse_ymd(s: str):
    s = (s or "").replace("-", "").replace(".", "").strip()[:8]
    try:
        return datetime.strptime(s, "%Y%m%d").date()
    except ValueError:
        return None


def fetch_all() -> list[dict]:
    """대전 근무 공고를 모든 페이지에 걸쳐 가져오기"""
    if not API_KEY or "입력" in API_KEY:
        sys.exit("[오류] API 키가 없습니다. 코드 상단 API_KEY 또는 환경변수 JOB_API_KEY 를 설정하세요.")

    rows, page = [], 1
    while True:
        params = {
            "serviceKey": API_KEY,
            "resultType": "json",
            "numOfRows": 100,
            "pageNo": page,
            "workRgnLst": REGION_CODE,
        }
        if ONLY_ONGOING:
            params["ongoingYn"] = "Y"

        for attempt in range(3):
            try:
                res = requests.get(BASE_URL, params=params, timeout=30)
                res.raise_for_status()
                data = res.json()
                break
            except (requests.RequestException, ValueError) as e:
                if attempt == 2:
                    sys.exit(f"[오류] API 호출 실패: {e}\n응답: {getattr(res, 'text', '')[:300]}")
                time.sleep(3)

        if str(data.get("resultCode")) != "200":
            sys.exit(f"[오류] API 응답 오류 ({data.get('resultCode')}): {data.get('resultMsg')}")

        batch = data.get("result") or []
        rows.extend(batch)
        total = int(data.get("totalCount") or 0)
        if not batch or len(rows) >= total:
            break
        page += 1
    return rows


def normalize(raw: list[dict]) -> list[dict]:
    today = today_kst()
    items, seen_keys = [], set()
    for r in raw:
        pid = pick(r, "recrutPblntSn", "recrutPbancSn", "sn")
        title = pick(r, "recrutPbancTtl", "title")
        key = pid or title
        if key in seen_keys:
            continue
        seen_keys.add(key)

        end = parse_ymd(pick(r, "pbancEndYmd", "rcptEndYmd", "endYmd"))
        # 안전장치: 이미 마감된 공고 제외
        if ONLY_ONGOING and end and end < today:
            continue

        items.append({
            "id": key,
            "기관명": pick(r, "instNm"),
            "직무분야": pick(r, "ncsCdNmLst", "ncsCdLst").replace(",", ", "),
            "공고제목": title,
            "고용형태": pick(r, "hireTypeNmLst", "hireTypeLst").replace(",", ", "),
            "채용구분": pick(r, "recrutSeNm", "recrutSe"),
            "공고시작일": parse_ymd(pick(r, "pbancBgngYmd", "bgngYmd")),
            "마감일": end,
            "링크": pick(r, "srcUrl", "url"),
        })

    # 마감 임박 순 정렬 (마감일 없는 공고는 맨 뒤)
    items.sort(key=lambda x: (x["마감일"] is None, x["마감일"] or date.max, x["기관명"]))
    return items


def load_seen() -> set:
    if SEEN_PATH.exists():
        return set(json.loads(SEEN_PATH.read_text(encoding="utf-8")))
    return set()


def save_seen(ids: set):
    SEEN_PATH.write_text(json.dumps(sorted(ids), ensure_ascii=False), encoding="utf-8")


def write_excel(items: list[dict], new_ids: set, gri: list[dict] = ()):
    today = today_kst()
    wb = Workbook()
    ws = wb.active
    ws.title = "채용공고"

    ws["A1"] = f"대전 지역 공공기관 채용공고 (갱신: {now_kst():%Y-%m-%d %H:%M}, 총 {len(items)}건, 신규 {len(new_ids)}건)"
    ws["A1"].font = Font(bold=True, size=13)

    headers = ["신규", "채용기관명", "채용 직무 분야", "서류 마감일", "D-day", "공고 제목", "고용형태", "채용구분", "공고 링크"]
    ws.append([])
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[3]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")

    new_fill = PatternFill("solid", fgColor="FFF2CC")
    urgent_font = Font(color="C00000", bold=True)

    for it in items:
        end = it["마감일"]
        dday = (end - today).days if end else None
        if dday is not None and dday > 3650:   # 2050-12-31 등은 상시채용
            end, dday = None, None
        dday_txt = "" if dday is None else ("D-Day" if dday == 0 else f"D-{dday}")
        is_new = it["id"] in new_ids
        ws.append([
            "NEW" if is_new else "",
            it["기관명"],
            it["직무분야"],
            end.strftime("%Y-%m-%d") if end else "상시/미정",
            dday_txt,
            it["공고제목"],
            it["고용형태"],
            it["채용구분"],
            it["링크"],
        ])
        row = ws.max_row
        if is_new:
            for c in ws[row]:
                c.fill = new_fill
        if dday is not None and dday <= 3:
            ws.cell(row, 4).font = urgent_font
            ws.cell(row, 5).font = urgent_font
        if it["링크"]:
            ws.cell(row, 9).hyperlink = it["링크"]
            ws.cell(row, 9).style = "Hyperlink"
        for c in ws[row]:
            c.alignment = Alignment(vertical="center", wrap_text=True)

    widths = [7, 28, 30, 13, 8, 50, 14, 10, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{max(ws.max_row, 3)}"

    ws2 = wb.create_sheet("정부출연연구기관")
    ws2["A1"] = f"대전 정부출연연구기관 채용공고 (NST 게시판, 최근 {NST_DAYS}일 등록분 {len(gri)}건) · 마감일은 공고문 확인"
    ws2["A1"].font = Font(bold=True, size=13)
    ws2.append([])
    ws2.append(["신규", "채용기관명", "채용 직무 분야(제목 기준)", "등록일", "서류 마감일", "공고 제목", "공고 링크"])
    for c in ws2[3]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    for it in gri:
        is_new = it["id"] in new_ids
        ws2.append(["NEW" if is_new else "", it["기관명"], it["직무분야"],
                    it["등록일"].strftime("%Y-%m-%d") if it["등록일"] else "", "공고문 확인", it["공고제목"], it["링크"]])
        row = ws2.max_row
        if is_new:
            for c in ws2[row]:
                c.fill = new_fill
        if it["링크"]:
            ws2.cell(row, 7).hyperlink = it["링크"]
            ws2.cell(row, 7).style = "Hyperlink"
        for c in ws2[row]:
            c.alignment = Alignment(vertical="center", wrap_text=True)
    for i, w in enumerate([7, 26, 28, 12, 13, 55, 45], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A4"

    try:
        wb.save(EXCEL_PATH)
    except PermissionError:
        sys.exit(f"[오류] 엑셀 파일이 열려 있어 저장할 수 없습니다. 파일을 닫고 다시 실행하세요: {EXCEL_PATH}")


def write_html(items: list[dict], new_ids: set, gri: list[dict] = (), gri_error: str = ""):
    """모바일에서도 보기 좋은 검색 가능한 웹페이지 생성"""
    today = today_kst()
    esc = lambda x: htmllib.escape(str(x or ""))
    rows = []
    for it in items:
        end = it["마감일"]
        dday = (end - today).days if end else None
        if end is None or (dday is not None and dday > 3650):
            end_txt, dday_txt, cls = "상시/미정", "", ""
        else:
            end_txt = end.strftime("%Y-%m-%d")
            dday_txt = "D-Day" if dday == 0 else f"D-{dday}"
            cls = "urgent" if dday <= 3 else ("soon" if dday <= 7 else "")
        new = '<span class="new">NEW</span>' if it["id"] in new_ids else ""
        title = esc(it["공고제목"])
        if it["링크"]:
            title = f'<a href="{esc(it["링크"])}" target="_blank" rel="noopener">{title}</a>'
        rows.append(
            f'<tr class="{cls}"><td data-l="기관명"><b>{esc(it["기관명"])}</b> {new}</td>'
            f'<td data-l="직무분야">{esc(it["직무분야"])}</td>'
            f'<td data-l="서류마감" class="end">{end_txt} <span class="dd">{dday_txt}</span></td>'
            f'<td data-l="공고">{title}<div class="sub">{" · ".join(esc(x) for x in (it["고용형태"], it["채용구분"]) if x)}</div></td></tr>'
        )
    grows = []
    for it in gri:
        new = '<span class="new">NEW</span>' if it["id"] in new_ids else ""
        grows.append(
            f'<tr><td data-l="기관명"><b>{esc(it["기관명"])}</b> {new}</td>'
            f'<td data-l="직무분야">{esc(it["직무분야"])}</td>'
            f'<td data-l="서류마감" class="end">공고문 확인<div class="sub">등록 {it["등록일"]:%Y-%m-%d}</div></td>'
            f'<td data-l="공고"><a href="{esc(it["링크"])}" target="_blank" rel="noopener">{esc(it["공고제목"])}</a></td></tr>'
        )
    if gri_error:
        gri_body = '<tr><td colspan="4">오늘은 출연연 공고를 불러오지 못했습니다. 내일 다시 갱신됩니다.</td></tr>'
    else:
        gri_body = chr(10).join(grows) if grows else '<tr><td colspan="4">최근 등록된 공고가 없습니다.</td></tr>'
    page = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>대전 공공기관 채용공고</title>
<style>
:root{{--b:#1f4e78;--line:#e3e7ec;--mut:#6b7280}}
*{{box-sizing:border-box}}
body{{margin:0;font-family:-apple-system,"Malgun Gothic","Apple SD Gothic Neo",sans-serif;background:#f6f8fa;color:#1f2328}}
header{{background:var(--b);color:#fff;padding:20px 16px}}
h1{{margin:0 0 6px;font-size:20px}}
.meta{{font-size:13px;opacity:.85}}
.meta a{{color:#fff}}
.wrap{{max-width:1100px;margin:0 auto;padding:16px}}
input{{width:100%;padding:11px 12px;font-size:15px;border:1px solid var(--line);border-radius:8px;margin-bottom:12px}}
table{{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden}}
th{{background:#eef2f6;text-align:left;font-size:13px;padding:10px;color:#374151}}
td{{padding:10px;border-top:1px solid var(--line);font-size:14px;vertical-align:top}}
td a{{color:var(--b);text-decoration:none}} td a:hover{{text-decoration:underline}}
.sub{{color:var(--mut);font-size:12px;margin-top:3px}}
.end{{white-space:nowrap}} .dd{{color:var(--mut);font-size:12px}}
tr.urgent .end,tr.urgent .dd{{color:#c00000;font-weight:700}}
tr.soon .dd{{color:#b45309;font-weight:600}}
h2{{font-size:17px;margin:28px 0 4px}} .cnt{{font-size:13px;color:var(--mut);font-weight:400}}
.new{{background:#f59e0b;color:#fff;font-size:11px;padding:1px 6px;border-radius:9px;margin-left:4px}}
@media (max-width:700px){{
 thead{{display:none}} table,tbody,tr,td{{display:block;width:100%}}
 tr{{border-top:6px solid #f6f8fa;padding:6px 0}}
 td{{border:0;padding:4px 12px}}
 td::before{{content:attr(data-l);display:inline-block;width:64px;color:var(--mut);font-size:12px}}
}}
</style></head><body>
<header><h1>대전 공공기관·공기업·출연연 채용공고</h1>
<div class="meta">갱신 {now_kst():%Y-%m-%d %H:%M} · 공공기관 {len(items)}건 · 출연연 {len(gri)}건 · 신규 {len(new_ids)}건 · 매일 오전 8시 자동 갱신 · <a href="{EXCEL_NAME}">엑셀 다운로드</a></div>
<div class="meta" id="visits">방문자 집계 중…</div></header>
<div class="wrap">
<input id="q" placeholder="기관명, 직무분야, 공고 제목으로 검색 (예: 연구, 정보통신, 원자력)">
<h2 style="margin-top:4px">공공기관·공기업 <span class="cnt">{len(items)}건</span></h2>
<table><thead><tr><th>채용기관명</th><th>채용 직무 분야</th><th>서류 마감일</th><th>공고 제목</th></tr></thead>
<tbody class="tb">
{chr(10).join(rows) if rows else '<tr><td colspan="4">현재 접수 중인 공고가 없습니다.</td></tr>'}
</tbody></table>
<p class="sub">출처: 공공데이터포털 기획재정부 공공기관 채용정보(잡알리오) · 근무지 대전 · 마감 임박순</p>
<h2>정부출연연구기관 <span class="cnt">{len(gri)}건</span></h2>
<p class="sub">원자력연·ETRI·기초지원연 등 대전 소재 과학기술 출연연 · 국가과학기술연구회(NST) 게시판 최근 {NST_DAYS}일 등록분 · 최신순<br>
마감일은 기관별 공고문(PDF)에만 있어 링크에서 확인해 주세요. 직무 분야는 공고 제목으로 분류한 참고용입니다.</p>
<table><thead><tr><th>채용기관명</th><th>채용 직무 분야</th><th>서류 마감일</th><th>공고 제목</th></tr></thead>
<tbody class="tb">
{gri_body}
</tbody></table>
</div>
<script>
const q=document.getElementById('q'),rs=[...document.querySelectorAll('.tb tr')];
q.addEventListener('input',()=>{{const v=q.value.trim().toLowerCase();rs.forEach(r=>r.style.display=!v||r.textContent.toLowerCase().includes(v)?'':'none')}});

// 방문자 수 (무료 카운터 abacus 사용, 같은 브라우저는 하루 1번만 집계)
(async()=>{{
  const API='https://abacus.jasoncameron.dev', NS='ssong1013-daejeon-jobs';
  const ymd=d=>new Intl.DateTimeFormat('en-CA',{{timeZone:'Asia/Seoul'}}).format(d).replaceAll('-','');
  const today=ymd(new Date()), yday=ymd(new Date(Date.now()-864e5));
  let first=true;
  try{{ first=!localStorage.getItem('djv-'+today); localStorage.setItem('djv-'+today,'1'); }}catch(e){{}}
  const val=async(act,key)=>{{ try{{ const r=await fetch(`${{API}}/${{act}}/${{NS}}/${{key}}`); if(!r.ok) return 0; return (await r.json()).value||0; }}catch(e){{ return null; }} }};
  const [t,y,all]=await Promise.all([val(first?'hit':'get','d'+today), val('get','d'+yday), val(first?'hit':'get','total')]);
  const el=document.getElementById('visits');
  el.textContent = t===null ? '' : `오늘 방문자 ${{t}}명 · 어제 ${{y}}명 · 누적 ${{all}}명`;
}})();
</script>
</body></html>"""
    HTML_PATH.write_text(page, encoding="utf-8")


# ───────────────── 정부출연연구기관 (NST 게시판) ─────────────────
# 과학기술 출연연은 2024년 공공기관 지정 해제로 잡알리오에 없음 →
# 국가과학기술연구회(NST) '소관기관 채용공고' 게시판에서 따로 수집
NST_LIST = "https://www.nst.re.kr/www/selectBbsNttList.do"
NST_VIEW = "https://www.nst.re.kr/www/selectBbsNttView.do"
NST_DAYS = 30   # 최근 며칠 이내 등록된 공고만 표시
DAEJEON_GRI = {  # 본원이 대전에 있는 출연연
    "한국기초과학지원연구원", "한국천문연구원", "한국생명공학연구원", "한국과학기술정보연구원",
    "한국한의학연구원", "한국표준과학연구원", "한국기계연구원", "한국에너지기술연구원",
    "한국화학연구원", "국가독성과학연구소", "안전성평가연구소", "한국원자력연구원",
    "한국핵융합에너지연구원", "한국전자통신연구원", "국가보안기술연구소", "한국항공우주연구원",
    "한국지질자원연구원", "국가수리과학연구소",
}
NOT_POSTING = re.compile(r"합격|결과|일정\s*변경|친인척|면접|취소|연기|정정")
UA = {"User-Agent": "Mozilla/5.0 (daejeon-jobs; daily job list)"}


def classify(title: str) -> str:
    """공고 제목으로 고용형태/직종을 추정 (정확한 내용은 공고문 확인)"""
    kinds = []
    rules = [
        (r"박사후|포닥|post-?doc", "박사후연구원"), (r"석사후", "석사후연구원"),
        (r"위촉연구원", "위촉연구원"), (r"학생연구원|연구연수생|인턴", "인턴·학생연구원"),
        (r"육아휴직|대체인력", "대체인력"), (r"기간제|계약직|비정규직", "계약직·기간제"),
        (r"정규직|신입직원|경력직원", "정규직"),
    ]
    for pat, name in rules:
        if re.search(pat, title, re.I) and name not in kinds:
            kinds.append(name)
    for pat, name in [(r"연구직", "연구직"), (r"기술직", "기술직"), (r"행정직", "행정직"), (r"장애인", "장애인 제한경쟁")]:
        if re.search(pat, title):
            kinds.append(name)
    return ", ".join(kinds) if kinds else "공고문 참고"


def fetch_nst() -> list[dict]:
    from bs4 import BeautifulSoup
    cutoff = today_kst() - timedelta(days=NST_DAYS)
    items = []
    for page in range(1, 15):
        r = requests.get(NST_LIST, params={"bbsNo": 19, "key": 61, "pageIndex": page}, headers=UA, timeout=30)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        rows = soup.select("table tbody tr")
        if not rows:
            break
        stop = False
        for tr in rows:
            for h in tr.select(".add-head"):
                h.decompose()
            cells = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
            a = tr.find("a", href=True)
            if len(cells) < 5 or not a:
                continue
            org, title, posted = cells[1], cells[2], parse_ymd(cells[4])
            if posted and posted < cutoff:
                stop = True
                continue
            m = re.search(r"nttNo=(\d+)", a["href"])
            if not m or org not in DAEJEON_GRI or NOT_POSTING.search(title):
                continue
            items.append({"id": f"nst-{m.group(1)}", "ntt": m.group(1), "기관명": org,
                          "공고제목": title, "직무분야": classify(title), "등록일": posted, "링크": ""})
        if stop:
            break
    # 상세 페이지에서 기관 원문 링크 가져오기 (실패하면 NST 상세 페이지로)
    for it in items:
        nst_url = f"{NST_VIEW}?key=61&bbsNo=19&nttNo={it['ntt']}"
        it["링크"] = nst_url
        try:
            r = requests.get(NST_VIEW, params={"key": 61, "bbsNo": 19, "nttNo": it["ntt"]}, headers=UA, timeout=20)
            body = BeautifulSoup(r.text, "html.parser").select_one("article") or BeautifulSoup(r.text, "html.parser")
            for u in re.findall(r"https?://[^\s\"'<>]+", htmllib.unescape(body.get_text(" "))):
                if "nst.re.kr" not in u:
                    it["링크"] = u
                    break
        except requests.RequestException:
            pass
        time.sleep(0.3)
    items.sort(key=lambda x: (x["등록일"] or date.min), reverse=True)
    return items


def main():
    print("대전 지역 공공기관 채용공고 조회 중...")
    items = normalize(fetch_all())
    try:
        gri = fetch_nst()
        gri_error = ""
    except Exception as e:   # NST 사이트 문제여도 잡알리오 결과는 그대로 저장
        gri, gri_error = [], str(e)
        print(f"[경고] 출연연(NST) 공고 수집 실패: {e}")

    seen = load_seen()
    current_ids = {it["id"] for it in items} | {it["id"] for it in gri}
    new_ids = current_ids - seen if seen else set()   # 첫 실행 땐 전부 신규로 표시하지 않음
    if seen and not any(str(x).startswith("nst-") for x in seen):
        new_ids -= {it["id"] for it in gri}   # 출연연 첫 수집 때는 NEW 표시 안 함

    write_excel(items, new_ids, gri)
    write_html(items, new_ids, gri, gri_error)
    save_seen(seen | current_ids)

    print(f"완료: 공공기관 {len(items)}건 + 출연연 {len(gri)}건 (신규 {len(new_ids)}건) → {EXCEL_PATH}")
    for it in items + gri:
        if it["id"] in new_ids:
            print(f"  [NEW] {it['기관명']} | {it['직무분야']} | 마감 {it.get('마감일', '공고문 확인')}")


if __name__ == "__main__":
    main()
