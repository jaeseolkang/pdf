# 2026-11-15 00:38 KST | 수정: 교보 보장분석 보고서(26쪽) → 1:1 요약 PDF 생성 코드, 글꼴에 없는 한자(無·多)·Ⅱ는 임베드 가능한 대체 TTF → CID 순으로 자동 대체
"""
보장분석 보고서 요약 PDF 생성기
==================================
교보생명 '보장분석 보고서'(26쪽 양식) PDF를 읽어서, 같은 쪽수(26쪽)의 요약 PDF를 만든다.

  원본 p.1 표지          → 요약 p.1
  원본 p.2 Pride of KYOBO → 요약 p.2
  원본 p.3 보험 가입 현황 → 요약 p.3
  원본 p.4 종합요약 → 요약 p.4
  원본 p.5 세대 분석      → 요약 p.5
  원본 p.6 상품별 분석 1/2 → 요약 p.6
  원본 p.7 상품별 분석 2/2 → 요약 p.7
  원본 p.8 신체부위별 보장현황 → 요약 p.8
  원본 p.9 보장밸런스(그래프) → 요약 p.9
  원본 p.10 연령별 분석 1/2 → 요약 p.10
  원본 p.11 연령별 분석 2/2 → 요약 p.11
  원본 p.12 보장항목별 분석 1/6 → 요약 p.12
  원본 p.13 보장항목별 분석 2/6 → 요약 p.13
  원본 p.14 보장항목별 분석 3/6 → 요약 p.14
  원본 p.15 보장항목별 분석 4/6 → 요약 p.15
  원본 p.16 보장항목별 분석 5/6 → 요약 p.16
  원본 p.17 보장항목별 분석 6/6 → 요약 p.17
  원본 p.18 보장 제안 → 요약 p.18
  원본 p.19 보장 제안 전/후 비교 1/6 → 요약 p.19
  원본 p.20 보장 제안 전/후 비교 2/6 → 요약 p.20
  원본 p.21 보장 제안 전/후 비교 3/6 → 요약 p.21
  원본 p.22 보장 제안 전/후 비교 4/6 → 요약 p.22
  원본 p.23 보장 제안 전/후 비교 5/6 → 요약 p.23
  원본 p.24 보장 제안 전/후 비교 6/6 → 요약 p.24
  원본 p.25 신체부위별 보장현황 전/후 비교 → 요약 p.25
  원본 p.26 보장밸런스(그래프) 전/후 비교 → 요약 p.26

원칙
  1) 숫자는 원본 문자열을 그대로 사용한다. (계산·반올림·단위 변환 없음)
  2) 표는 표로 옮기지 않고 문장/항목 텍스트로 풀어 쓴다.
  3) 계약자·주피·종피 반복, 안내 문구 3줄, 빈 행 등 중요하지 않은 내용은 뺀다.
  4) 원본 1쪽 = 요약 1쪽. 내용이 넘치면 글자 크기를 줄여서라도 한 쪽에 맞춘다.

설치
  pip install pdfplumber reportlab pypdf

사용
  python gaejang_summary_report.py 고순종.pdf
  python gaejang_summary_report.py 고순종.pdf 고순종_요약.pdf

한글 글꼴
  Windows는 '맑은 고딕'(malgun.ttf)을 자동 사용한다. 다른 환경에서는 스크립트와 같은 폴더의
  fonts/NanumGothic.ttf, fonts/NanumGothicBold.ttf 를 찾고, 없으면 내장 CID 글꼴(HYGothic)로 대체한다.
"""
import os
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape as _xml_escape

import pdfplumber
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, HRFlowable, KeepInFrame,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

# ----------------------------------------------------------------------------
# 1. 양식 상수 (교보생명 보장분석 보고서 8쪽 양식의 열 순서)
# ----------------------------------------------------------------------------
EXPECTED_PAGES = 26

# 사망·진단 표 (18열)
LABELS_A = ["일반사망", "암사망", "재해사망",
            "일반암", "고액암", "소액암", "뇌출혈", "뇌경색", "급성심근경색", "뇌혈관질환",
            "허혈성심장질환", "CI", "GI", "당뇨", "LTC", "중증치매", "중등도치매", "경도치매"]
GROUPS_A = [("사망", LABELS_A[0:3]), ("진단", LABELS_A[3:])]

# 진단·항암·수술·입원·통원·실손 표 (18열)
LABELS_B = ["재해골절진단",
            "항암약물치료", "항암방사선", "항암중입자",
            "5종수술", "암수술", "뇌졸중수술", "급성심근수술", "재해수술",
            "질병입원(4일↑)", "재해입원(4일↑)", "암입원", "뇌혈관입원", "허혈심장입원",
            "질병입원간병사용", "질병간호간병",
            "암통원",
            "질병입원실손"]
GROUPS_B = [("진단(재해골절)", LABELS_B[0:1]), ("항암", LABELS_B[1:4]), ("수술", LABELS_B[4:9]),
            ("입원", LABELS_B[9:16]), ("통원", LABELS_B[16:17]), ("의료비 실손", LABELS_B[17:18])]

# 신체부위별 보장현황(6쪽)
BODY_GROUPS = [("뇌", ["뇌출혈", "뇌경색", "뇌혈관질환", "산정특례(뇌)"]),
               ("암", ["일반암", "고액암", "소액암", "표적항암"]),
               ("당뇨", ["당뇨"]),
               ("치매", ["LTC(치매)", "중등도치매", "경도치매"]),
               ("심장", ["급성심근경색", "허혈성심장질환", "산정특례(심장)"]),
               ("말기 장기질환", ["말기폐질환", "말기간질환", "말기신부전증"])]

PAGE_TITLES = {
    2: "보험 가입 현황",
    3: "세대 분석",
    4: "상품별 분석 (1/2) 사망·진단",
    5: "상품별 분석 (2/2) 항암·수술·입원·통원·실손",
    6: "신체부위별 보장현황",
    7: "연령별 분석 (1/2) 사망·진단",
    8: "연령별 분석 (2/2) 항암·수술·입원·통원·실손",
}

# ----------------------------------------------------------------------------
# 2. 글꼴 / 색 / 스타일
# ----------------------------------------------------------------------------
FONT = "Helvetica"
FONT_B = "Helvetica-Bold"

NAVY = colors.HexColor("#17306B")
LIGHT = colors.HexColor("#EEF2F8")
INK = colors.HexColor("#1F2937")
GRAY = colors.HexColor("#6B7280")
C_OK = "#1B8A80"      # 적정/충분
C_LOW = "#E28E3A"     # 부족
C_NONE = "#C54648"    # 미보장
C_GRAY = "#6B7280"

PAGE_W, PAGE_H = A4
MARGIN_X = 16 * mm
CW = PAGE_W - 2 * MARGIN_X                 # 본문 폭
FRAME_BOTTOM = 18 * mm
FRAME_TOP_MARGIN = 14 * mm
FH = PAGE_H - FRAME_BOTTOM - FRAME_TOP_MARGIN - 2   # 본문 높이(여유 2pt)

S = {}
GLYPHS = None          # 기본 TTF 글꼴이 가진 글자 코드 집합 (None이면 대체 처리 안 함)
FALLBACKS = []         # [(글꼴이름, 글자집합)] 기본 글꼴에 없는 글자(한자·로마숫자 등)용 대체 글꼴
CID_FALLBACK = "HYGothic-Medium"

# (경로, subfontIndex) — 글자를 PDF에 임베드할 수 있는 TrueType 계열만
FALLBACK_CANDIDATES = [
    (r"C:\Windows\Fonts\malgun.ttf", 0),
    (r"C:\Windows\Fonts\batang.ttc", 0),
    (r"C:\Windows\Fonts\gulim.ttc", 0),
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0),
    ("/System/Library/Fonts/Supplemental/AppleGothic.ttf", 0),
]


def esc(text):
    """XML 이스케이프 + 현재 글꼴에 없는 글자(한자·로마숫자 등)만 대체 글꼴로 표시."""
    t = _xml_escape(text)
    if not GLYPHS:
        return t
    out = []
    for ch in t:
        if ord(ch) < 128 or ord(ch) in GLYPHS:
            out.append(ch)
            continue
        name = next((n for n, g in FALLBACKS if ord(ch) in g), CID_FALLBACK)
        out.append(f'<font name="{name}">{ch}</font>')
    return "".join(out)


def register_fonts():
    """한글 글꼴 등록. TTF를 못 찾으면 내장 CID 글꼴로 대체."""
    global FONT, FONT_B, GLYPHS
    here = Path(__file__).resolve().parent
    pdfmetrics.registerFont(UnicodeCIDFont(CID_FALLBACK))
    candidates = [
        (r"C:\Windows\Fonts\malgun.ttf", r"C:\Windows\Fonts\malgunbd.ttf"),
        (str(here / "fonts" / "NanumGothic.ttf"), str(here / "fonts" / "NanumGothicBold.ttf")),
        (str(here / "NanumGothic.ttf"), str(here / "NanumGothicBold.ttf")),
        ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
        ("/System/Library/Fonts/Supplemental/AppleGothic.ttf", "/System/Library/Fonts/Supplemental/AppleGothic.ttf"),
    ]
    for reg, bold in candidates:
        if os.path.exists(reg):
            pdfmetrics.registerFont(TTFont("KR", reg))
            pdfmetrics.registerFont(TTFont("KR-B", bold if os.path.exists(bold) else reg))
            pdfmetrics.registerFontFamily("KR", normal="KR", bold="KR-B", italic="KR", boldItalic="KR-B")
            FONT, FONT_B = "KR", "KR-B"
            GLYPHS = set(pdfmetrics.getFont("KR").face.charToGlyph.keys())
            register_fallbacks(reg)
            print(f"[글꼴] {reg}")
            return
    pdfmetrics.registerFontFamily("HYGothic-Medium", normal="HYGothic-Medium", bold="HYGothic-Medium",
                                  italic="HYGothic-Medium", boldItalic="HYGothic-Medium")
    FONT = FONT_B = "HYGothic-Medium"
    print("[글꼴] TTF를 찾지 못해 내장 CID 글꼴(HYGothic-Medium)을 사용합니다.")


def register_fallbacks(primary_path):
    """기본 글꼴에 없는 글자를 그릴 대체 TTF를 등록한다 (기본 글꼴과 같은 파일은 제외)."""
    for i, (path, idx) in enumerate(FALLBACK_CANDIDATES):
        if os.path.exists(path) and os.path.normcase(path) != os.path.normcase(primary_path):
            try:
                name = f"KR-FB{i}"
                pdfmetrics.registerFont(TTFont(name, path, subfontIndex=idx))
                FALLBACKS.append((name, set(pdfmetrics.getFont(name).face.charToGlyph.keys())))
            except Exception:
                pass


def build_styles():
    base = dict(fontName=FONT, textColor=INK)
    S["body"] = ParagraphStyle("body", fontSize=9.5, leading=14, **base)
    S["bullet"] = ParagraphStyle("bullet", parent=S["body"], leftIndent=11, bulletIndent=1, spaceAfter=3)
    S["h2"] = ParagraphStyle("h2", fontName=FONT_B, fontSize=11, leading=15, textColor=NAVY,
                             spaceBefore=9, spaceAfter=3)
    S["sub"] = ParagraphStyle("sub", fontName=FONT, fontSize=8.5, leading=12, textColor=GRAY)
    S["title"] = ParagraphStyle("title", fontName=FONT_B, fontSize=16, leading=20, textColor=colors.white)
    S["tag"] = ParagraphStyle("tag", fontName=FONT, fontSize=9, leading=12, textColor=colors.white, alignment=2)
    S["kpi_l"] = ParagraphStyle("kpi_l", fontName=FONT, fontSize=8, leading=11, textColor=GRAY, alignment=1)
    S["kpi_v"] = ParagraphStyle("kpi_v", fontName=FONT_B, fontSize=13, leading=17, textColor=NAVY, alignment=1)
    S["cover_sub"] = ParagraphStyle("cover_sub", fontName=FONT, fontSize=15, leading=22, textColor=INK)
    S["cover_title"] = ParagraphStyle("cover_title", fontName=FONT_B, fontSize=30, leading=38, textColor=NAVY,
                                      spaceAfter=6)


# ----------------------------------------------------------------------------
# 3. 원본 PDF 파싱  (숫자는 항상 '문자열 그대로' 보관)
# ----------------------------------------------------------------------------
class ParseError(Exception):
    pass


def clean(s):
    return re.sub(r"\s+", " ", (s or "").replace("\n", " ")).strip()


NUM_TOKENS = r"(?:(?:[\d,]+|-)(?:\s+|$))+"
MEMBER_RE = re.compile(r"^(\S+)\s*\((\d+)\)\s+(" + NUM_TOKENS + r")$")
AGE_RE = re.compile(r"^(\d+세|종신)\s+(" + NUM_TOKENS + r")$")


def parse_cover(page):
    t = page.extract_text() or ""

    def g(pat):
        m = re.search(pat, t)
        return m.group(1).strip() if m else ""

    return dict(name=g(r"(\S+)\s*고객님을 위한"), fp=g(r"Financial Planner\s*(\S+)"),
                branch=g(r"소속:\s*(.+)"), phone=g(r"H\.P:\s*(\S+)"), email=g(r"Email:\s*(\S+)"),
                print_date=g(r"출력\s*\|\s*([\d-]+)"))


def parse_enrollment(page):
    t = page.extract_text() or ""

    def g(pat):
        m = re.search(pat, t)
        return m.group(1) if m else ""

    info = dict(birth=g(r"\((\d{4}\.\d{2}\.\d{2})일생\)"), ins_age=g(r"보험나이\s*(\d+)세"),
                age_date=g(r"상령일\s*([\d.]+)"), n_contracts=g(r"유지계약\s*(\d+)건"),
                monthly=g(r"월보험료\s*([\d,]+)원"), paid_total=g(r"기납입보험료\s*([\d,]+)원"))
    m = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일 기준", t)
    info["basis"] = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""

    contracts = []
    for tb in page.extract_tables():
        for row in tb:
            if len(row) >= 13 and re.fullmatch(r"\d+", clean(row[0])):
                c = [clean(x) for x in row]
                contracts.append(dict(no=c[0], origin=c[1], company=c[2], name=c[3], holder=c[4],
                                      insured=c[5], coinsured=c[6], date=c[7], expiry=c[8],
                                      premium=c[9], count=c[10], last=c[11], paid=c[12]))
    if not contracts:
        raise ParseError("2쪽에서 계약 목록을 찾지 못했습니다.")
    return info, contracts


def parse_generation(page):
    t = page.extract_text() or ""
    m = re.search(r"(\d+)건\s+(\d+)건\s+(\d+)건\s+(\d+)건\s+(\d+)건", t)
    counts = dict(zip(["total", "kyobo", "life", "nonlife", "mutual"], m.groups())) if m else {}
    members = []
    for line in t.splitlines():
        mm_ = MEMBER_RE.match(line.strip())
        if mm_ and len(mm_.group(3).split()) == 18:
            members.append((mm_.group(1), mm_.group(2), mm_.group(3).split()))
    if not members or len(members) % 2:
        raise ParseError("3쪽에서 구성원별 보장 행을 찾지 못했습니다.")
    n = len(members) // 2
    result = []
    for i in range(n):
        name, age, a = members[i]
        _, _, b = members[n + i]
        result.append(dict(name=name, age=age, a=a, b=b))
    return counts, result


def parse_products(page, labels):
    """4·5쪽: 필요/준비/부족 + 상품별 준비 금액."""
    t = page.extract_text() or ""
    m = re.search(r"(\d+)\s*개\s+(\d+)\s*개\s+(\d+)\s*개\s+(\d+)\s*개", t)
    head = dict(zip(["total", "ok", "low", "none"], m.groups())) if m else {}
    n = len(labels)
    d = dict(need=None, prep=None, short=None, total=None, products=[], head=head)
    for tb in page.extract_tables():
        for row in tb:
            if len(row) < 3:
                continue
            key = clean(row[1])
            vals = [clean(c) for c in row[2:] if c is not None and clean(c) != ""]
            if key in ("필요", "준비", "부족") and len(vals) == n:
                d[{"필요": "need", "준비": "prep", "부족": "short"}[key]] = vals
            elif not key and row[0] and len(vals) == n:
                cell = row[0].strip()
                if cell.startswith("보장 합계"):
                    d["total"] = vals
                    continue
                lines = [l.strip() for l in cell.split("\n") if l.strip()]
                if len(lines) < 3:
                    continue
                tail = re.match(r"([\d,]+)원\s*\((.+)\)$", lines[-1])
                d["products"].append(dict(company=lines[0], name="".join(lines[1:-1]),
                                          premium=tail.group(1) + "원" if tail else lines[-1],
                                          expiry=tail.group(2) if tail else "", vals=vals))
    if not (d["need"] and d["prep"] and d["short"]):
        raise ParseError("상품별 분석에서 필요/준비/부족 행을 찾지 못했습니다.")
    return d


def canonical_names(products, contracts):
    """상품명이 줄바꿈으로 잘린 경우, 2쪽 계약 목록의 상품명으로 바로잡는다."""
    for p in products:
        key = re.sub(r"\s+", "", p["name"])
        for c in contracts:
            if c["company"] == p["company"] and re.sub(r"\s+", "", c["name"]).startswith(key):
                p["name"] = c["name"]
                break


def color_status(col, text):
    """6쪽 숫자 글자색 → 충분/부족/미보장 (원본 범례 색 기준)."""
    try:
        r, g, b = [float(x) for x in col[:3]]
        if r > 0.7 and g < 0.4:
            return "미보장"
        if r > 0.8 and 0.4 <= g < 0.75:
            return "부족"
        return "충분"
    except Exception:
        return "미보장" if text == "0" else "부족"


def parse_body(page):
    words = page.extract_words(extra_attrs=["non_stroking_color"])
    nums = [w for w in words if re.fullmatch(r"[\d,]+", w["text"])]
    out = {}
    for _, labs in BODY_GROUPS:
        for lab in labs:
            cand_labels = [w for w in words if w["text"] == lab]
            if not cand_labels:
                raise ParseError(f"6쪽에서 '{lab}' 항목을 찾지 못했습니다.")
            lw = cand_labels[0]
            cands = [n for n in nums if abs(n["top"] - lw["top"]) < 4 and n["x0"] > lw["x1"]]
            if not cands:
                raise ParseError(f"6쪽에서 '{lab}' 금액을 찾지 못했습니다.")
            n = min(cands, key=lambda w: w["x0"] - lw["x1"])
            out[lab] = (n["text"], color_status(n.get("non_stroking_color"), n["text"]))
    return out


def parse_ages(page, n_labels):
    rows = []
    for line in (page.extract_text() or "").splitlines():
        m = AGE_RE.match(line.strip())
        if m and len(m.group(2).split()) == n_labels:
            rows.append((m.group(1), m.group(2).split()))
    if not rows:
        raise ParseError("연령별 분석 행을 찾지 못했습니다.")
    return rows


def parse_report(pdf):
    if len(pdf.pages) != EXPECTED_PAGES:
        raise ParseError(f"원본이 {len(pdf.pages)}쪽입니다. 이 코드는 {EXPECTED_PAGES}쪽 보장분석 보고서 양식 전용입니다.")
    P_ = pdf.pages
    info, contracts = parse_enrollment(P_[1])
    data = dict(cover=parse_cover(P_[0]), info=info, contracts=contracts)
    data["gen_counts"], data["members"] = parse_generation(P_[2])
    data["prod1"] = parse_products(P_[3], LABELS_A)
    data["prod2"] = parse_products(P_[4], LABELS_B)
    for k in ("prod1", "prod2"):
        canonical_names(data[k]["products"], contracts)
    data["body"] = parse_body(P_[5])
    data["ages1"] = parse_ages(P_[6], len(LABELS_A))
    data["ages2"] = parse_ages(P_[7], len(LABELS_B))
    return data


# ----------------------------------------------------------------------------
# 4. 요약 문장 / 레이아웃 도우미
# ----------------------------------------------------------------------------
def col(text, color):
    return f'<font color="{color}">{text}</font>'


def P(text, style="body"):
    return Paragraph(text, S[style])


def bullet(text):
    return Paragraph(text, S["bullet"], bulletText="•")


def h2(text):
    return Paragraph(text, S["h2"])


def page_header(title, src, sub=""):
    t = Table([[Paragraph(title, S["title"]), Paragraph(f"원본 p.{src}", S["tag"])]],
              colWidths=[CW - 30 * mm, 30 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                           ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    out = [t]
    if sub:
        out += [Spacer(1, 4), P(sub, "sub")]
    out.append(Spacer(1, 6))
    return out


def kpi(items):
    """핵심 수치 카드. items = [(라벨, 값, 색상(옵션))]"""
    w = CW / len(items)
    labels = [Paragraph(i[0], S["kpi_l"]) for i in items]
    values = []
    for i in items:
        c = i[2] if len(i) > 2 else None
        txt = col(i[1], c) if c else i[1]
        values.append(Paragraph(txt, S["kpi_v"]))
    t = Table([labels, values], colWidths=[w] * len(items))
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("LINEAFTER", (0, 0), (-2, -1), 2, colors.white),
                           ("TOPPADDING", (0, 0), (-1, 0), 6), ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
                           ("TOPPADDING", (0, 1), (-1, 1), 0), ("BOTTOMPADDING", (0, 0), (-1, 0), 0),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    return t


def show(v):
    return "없음(-)" if v == "-" else v


def group_lines(labels, vals, groups):
    """그룹별로 '준비된 항목 금액' + '미준비(-) 항목'을 한 줄 문장으로."""
    out, idx = [], 0
    for gname, glabels in groups:
        gv = vals[idx: idx + len(glabels)]
        idx += len(glabels)
        present = [f"{esc(l)} <b>{v}</b>" for l, v in zip(glabels, gv) if v != "-"]
        absent = [esc(l) for l, v in zip(glabels, gv) if v == "-"]
        txt = f"<b>{gname}</b>  " + (" · ".join(present) if present else "전 항목 미준비(-)")
        if present and absent:
            txt += "  " + col("/ 미준비(-): " + " · ".join(absent), C_GRAY)
        out.append(bullet(txt))
    return out


def status_of(prep, short):
    if prep == "-":
        return "미보장"
    if short == "-":
        return "적정"
    return "부족"


def item_line(name, need, prep, short):
    st = status_of(prep, short)
    if st == "미보장":
        tail = col(f"미보장 · 부족 {short}", C_NONE)
        return f"<b>{esc(name)}</b> 필요 {need} · 준비 없음(-) · {tail}"
    if st == "적정":
        return f"<b>{esc(name)}</b> 필요 {need} · 준비 {prep} · {col('적정', C_OK)}"
    return f"<b>{esc(name)}</b> 필요 {need} · 준비 {prep} · {col('부족 ' + short, C_LOW)}"


# ----------------------------------------------------------------------------
# 5. 쪽별 스토리
# ----------------------------------------------------------------------------
def story_cover(d):
    cv, info = d["cover"], d["info"]
    st = [Spacer(1, 38 * mm), P(f"{esc(cv['name'])} 고객님을 위한", "cover_sub"), P("보장분석 요약 보고서", "cover_title"),
          HRFlowable(width=70 * mm, thickness=3, color=NAVY, hAlign="LEFT"), Spacer(1, 6),
          P("원본 보장분석 보고서를 쪽별로 1:1 요약했습니다. 금액은 원본 표기 그대로이며 단위는 원본과 같습니다.", "sub"),
          Spacer(1, 16 * mm),
          P(f"<b>Financial Planner</b>  {esc(cv['fp'])}"),
          P(f"소속  {esc(cv['branch'])}"),
          P(f"H.P  {esc(cv['phone'])}"),
          P(f"Email  {esc(cv['email'])}"),
          Spacer(1, 12 * mm), h2("구성 (요약 쪽 = 원본 쪽)")]
    for no, title in PAGE_TITLES.items():
        st.append(bullet(f"{no}쪽  {title}"))
    st += [Spacer(1, 10 * mm),
           P(f"분석 기준일 {info['basis'] or cv['print_date']}  ·  출력일 {cv['print_date']}", "sub")]
    return st


def story_enrollment(d):
    info, cs, cv = d["info"], d["contracts"], d["cover"]
    st = page_header("보험 가입 현황", 2, f"{esc(cv['name'])}님 ({info['birth']}생) 현재 유지 중인 계약")
    st.append(kpi([("보험나이", f"{info['ins_age']}세"), ("상령일", info["age_date"]),
                   ("유지계약", f"{info['n_contracts']}건"), ("월보험료", f"{info['monthly']}원"),
                   ("기납입보험료", f"{info['paid_total']}원")]))
    st.append(Spacer(1, 6))
    same = all(c["holder"] == c["insured"] == cv["name"] and c["coinsured"] == "-" for c in cs)
    if same:
        st.append(P(f"{len(cs)}건 모두 계약자·주피보험자가 {esc(cv['name'])}님 본인이며 종피보험자는 없습니다."))
    st.append(h2("계약 목록"))
    for c in cs:
        origin = "당사" if c["origin"] == "당사" else f"타사({esc(c['origin'])})"
        who = "" if same else f" 계약자 {esc(c['holder'])}, 주피 {esc(c['insured'])}, 종피 {esc(c['coinsured'])}."
        last = "" if c["last"] in ("", "-") else f", 최종월도 {c['last']}"
        st.append(bullet(
            f"<b>{c['no']}. {esc(c['company'])} {esc(c['name'])}</b> ({origin})<br/>"
            f"계약일 {c['date']} · 만기 {esc(c['expiry'])} · 월보험료 {c['premium']}원 · "
            f"납입횟수 {c['count']}{last} · 기납입 {c['paid']}원.{who}"))
    st.append(Spacer(1, 4))
    st.append(P("납입횟수는 ‘납입/전체(잔여)’ 형식으로 원본 표기 그대로입니다.", "sub"))
    return st


def story_generation(d):
    cnt, members = d["gen_counts"], d["members"]
    st = page_header("세대 분석", 3, "가족 구성원별 보장 준비 금액 (단위: 만원)")
    if cnt:
        st.append(P(f"유지계약 <b>{cnt['total']}건</b> — 교보생명 {cnt['kyobo']}건 · 생명보험 {cnt['life']}건 · "
                    f"손해보험 {cnt['nonlife']}건 · 공제/우체국 {cnt['mutual']}건"))
    for m in members:
        st.append(h2(f"{esc(m['name'])} ({m['age']}세)"))
        st += group_lines(LABELS_A, m["a"], GROUPS_A)
        st += group_lines(LABELS_B, m["b"], GROUPS_B)
    return st


def story_products(d, part, src, labels, groups):
    n = len(labels)
    need, prep, short = d["need"], d["prep"], d["short"]
    stat = [status_of(prep[i], short[i]) for i in range(n)]
    cnt = {k: stat.count(k) for k in ("적정", "부족", "미보장")}
    sub = f"필요자금 대비 급부별 분석 ({part}) · 단위: 만원"
    st = page_header(f"상품별 분석 ({part})", src, sub)
    if part == "1/2" and d["head"]:
        h = d["head"]
        st.append(kpi([("분석항목(전체)", f"{h['total']}개"), ("적정보장", f"{h['ok']}개", C_OK),
                       ("부족보장", f"{h['low']}개", C_LOW), ("미보장", f"{h['none']}개", C_NONE)]))
        st.append(Spacer(1, 3))
        st.append(P(f"이 쪽 {n}개 항목: 적정 {cnt['적정']}개 · 부족 {cnt['부족']}개 · 미보장 {cnt['미보장']}개", "sub"))
    else:
        st.append(kpi([("이 쪽 분석항목", f"{n}개"), ("적정보장", f"{cnt['적정']}개", C_OK),
                       ("부족보장", f"{cnt['부족']}개", C_LOW), ("미보장", f"{cnt['미보장']}개", C_NONE)]))
    st.append(h2("필요 대비 준비 현황"))
    idx = 0
    for gname, glabels in groups:
        lines = []
        for lab in glabels:
            lines.append(item_line(lab, need[idx], prep[idx], short[idx]))
            idx += 1
        st.append(bullet(f"<b>[{gname}]</b><br/>" + "<br/>".join(lines)))
    st.append(h2("상품별 준비 내용"))
    for p in d["products"]:
        items = [f"{esc(l)} <b>{v}</b>" for l, v in zip(labels, p["vals"]) if v != "-"]
        exp = p["expiry"] if p["expiry"] == "종신" else f"만기 {p['expiry']}"
        head = f"<b>{esc(p['company'])} {esc(p['name'])}</b> (월 {p['premium']} · {esc(exp)})"
        st.append(bullet(head + "<br/>" + (" · ".join(items) if items else "이 쪽 항목에 해당하는 보장 없음")))
    if d["total"] == prep:
        st.append(P("상품별 금액을 합하면 위 ‘준비’ 금액과 같습니다 (원본 ‘보장 합계’ 행).", "sub"))
    return st


def story_body(d):
    body = d["body"]
    st = page_header("신체부위별 보장현황", 6, "원본 범례 색 기준 충분/부족/미보장 · 단위: 만원")
    cnt = {"충분": 0, "부족": 0, "미보장": 0}
    for _, st_ in body.values():
        cnt[st_] += 1
    st.append(kpi([("충분", f"{cnt['충분']}개", C_OK), ("부족", f"{cnt['부족']}개", C_LOW),
                   ("미보장", f"{cnt['미보장']}개", C_NONE)]))
    palette = {"충분": C_OK, "부족": C_LOW, "미보장": C_NONE}
    for gname, labs in BODY_GROUPS:
        parts = []
        for lab in labs:
            v, s_ = body[lab]
            parts.append(f"{esc(lab)} <b>{v}</b> " + col(f"({s_})", palette[s_]))
        st.append(h2(gname))
        st.append(bullet(" · ".join(parts)))
    st.append(Spacer(1, 4))
    st.append(P(f"충분 {cnt['충분']}개 · 부족 {cnt['부족']}개 · 미보장 {cnt['미보장']}개 "
                "(금액이 0인 항목은 미보장으로 표시되어 있습니다).", "sub"))
    return st


def age_groups(rows):
    groups = []
    for age, vals in rows:
        if groups and groups[-1][1] == vals:
            groups[-1][0].append(age)
        else:
            groups.append(([age], vals))
    return groups


def story_ages(title, src, rows, labels, groups):
    st = page_header(title, src, "조회 연령에 따라 보장 금액이 어떻게 바뀌는지 · 단위: 만원")
    prev = None
    for ages, vals in age_groups(rows):
        st.append(h2(" · ".join(ages) + " 조회 시"))
        if len(ages) > 1:
            st.append(P(col(f"위 {len(ages)}개 시점의 보장 금액은 모두 같습니다.", C_GRAY), "sub"))
        if prev is None:
            st += group_lines(labels, vals, groups)
        else:
            ch = [f"<b>{esc(l)}</b> {show(a)} → {show(b)}" for l, a, b in zip(labels, prev, vals) if a != b]
            st.append(bullet("직전 구간 대비 변동: " + (", ".join(ch) if ch else "없음")))
            st.append(bullet("위에 적지 않은 항목은 직전 구간과 같습니다."))
        prev = vals
    return st


# ----------------------------------------------------------------------------
# 6. PDF 조립
# ----------------------------------------------------------------------------
def draw_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
    canvas.line(MARGIN_X, 14 * mm, PAGE_W - MARGIN_X, 14 * mm)
    canvas.setFont(FONT, 7.5)
    canvas.setFillColor(GRAY)
    canvas.drawString(MARGIN_X, 10 * mm, doc.footer_left)
    canvas.drawRightString(PAGE_W - MARGIN_X, 10 * mm, f"{canvas.getPageNumber()} / {EXPECTED_PAGES}")
    canvas.restoreState()


def build_pdf(data, out_path):
    build_styles()
    cv, info = data["cover"], data["info"]
    pages = [
        story_cover(data),
        story_enrollment(data),
        story_generation(data),
        story_products(data["prod1"], "1/2", 4, LABELS_A, GROUPS_A),
        story_products(data["prod2"], "2/2", 5, LABELS_B, GROUPS_B),
        story_body(data),
        story_ages("연령별 분석 (1/2)", 7, data["ages1"], LABELS_A, GROUPS_A),
        story_ages("연령별 분석 (2/2)", 8, data["ages2"], LABELS_B, GROUPS_B),
    ]
    doc = BaseDocTemplate(str(out_path), pagesize=A4, title=f"{cv['name']} 보장분석 요약 보고서",
                          author=cv["fp"], leftMargin=MARGIN_X, rightMargin=MARGIN_X)
    doc.footer_left = (f"{cv['name']} 고객님 보장분석 요약 · {info['basis'] or cv['print_date']} 기준 · "
                       "보험금 지급 근거가 아니며 약관·증권이 우선합니다")
    frame = Frame(MARGIN_X, FRAME_BOTTOM, CW, FH, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=draw_footer)])
    story = []
    for i, p in enumerate(pages):
        story.append(KeepInFrame(CW, FH, p, mode="shrink"))      # 넘치면 축소해서 한 쪽에 맞춤
        if i < len(pages) - 1:
            story.append(PageBreak())
    doc.build(story)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(src.stem + "_요약.pdf")
    register_fonts()
    with pdfplumber.open(str(src)) as pdf:
        data = parse_report(pdf)
    build_pdf(data, out)
    try:
        from pypdf import PdfReader
        n = len(PdfReader(str(out)).pages)
        print(f"[완료] {out}  ({n}쪽 / 원본 {EXPECTED_PAGES}쪽)")
        if n != EXPECTED_PAGES:
            print("[경고] 쪽수가 원본과 다릅니다. 내용이 한 쪽을 넘쳤는지 확인하세요.")
    except ImportError:
        print(f"[완료] {out}")


if __name__ == "__main__":
    main()
