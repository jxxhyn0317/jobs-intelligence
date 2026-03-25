import streamlit as st
import google.generativeai as genai
import json
import re
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── 설정 ──────────────────────────────────────────────────────────────────────
GEMINI_API_KEY = "AIzaSyBqf6XTwXdiY-QeShQsMq3aqcr582eB-mM"
MODEL          = "gemini-2.0-flash-lite"

st.set_page_config(page_title="JD Analyst", page_icon="◎", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap');
*, html, body { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: linear-gradient(180deg, #000000 0%, #141414 50%, #2a2a2a 100%); color: #d4d0c8; }
.block-container { padding: 0 !important; max-width: 100% !important; }
.hero { padding: 80px 20px 52px; border-bottom: 1px solid #1e1e1e; text-align: center; display: flex; flex-direction: column; align-items: center; }
.hero-eyebrow { font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: 0.25em; color: #444; text-transform: uppercase; margin-bottom: 20px; }
.hero-title { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 54px; font-weight: 700; color: #ffffff; letter-spacing: -0.03em; line-height: 1.0; margin-bottom: 18px; text-shadow: 0 0 30px rgba(255,255,255,0.55), 0 0 70px rgba(255,255,255,0.2); }
.hero-desc { font-size: 14px; color: #555; font-weight: 300; line-height: 1.65; text-align: center; }
.form-section { padding: 52px 20px; border-bottom: 1px solid #1e1e1e; display: flex; flex-direction: column; align-items: center; }
.field-label { font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: 0.15em; color: #444; text-transform: uppercase; margin-bottom: 8px; display: block; text-align: center; }
.stTextInput { max-width: 584px; margin: 0 auto !important; }
.stTextInput > div > div > input { background: #0d0d0d !important; border: 1px solid #252525 !important; border-radius: 999px !important; color: #d4d0c8 !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 13px !important; padding: 14px 22px !important; transition: border-color 0.2s !important; text-align: center !important; }
.stTextInput > div > div > input:focus { border-color: #ffffff !important; box-shadow: 0 0 20px rgba(255,255,255,0.05) !important; }
.stTextInput > div > div > input::placeholder { color: #2e2e2e !important; text-align: center !important; }
.stButton { max-width: 584px; margin: 0 auto !important; display: flex; justify-content: center; }
.stButton > button { background: #ffffff !important; color: #000000 !important; border: none !important; border-radius: 999px !important; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important; font-size: 12px !important; font-weight: 700 !important; letter-spacing: 0.08em !important; text-transform: uppercase !important; padding: 14px 36px !important; height: auto !important; transition: opacity 0.15s !important; width: 100% !important; box-shadow: 0 0 24px rgba(255,255,255,0.15) !important; }
.stButton > button:hover { opacity: 0.88 !important; }
.stButton > button:disabled { background: #1a1a1a !important; color: #333 !important; box-shadow: none !important; }
.log-box { background: #080808; border: 1px solid #1e1e1e; border-radius: 16px; padding: 20px 28px; font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #444; line-height: 1.8; margin: 24px auto; max-width: 640px; width: 90%; min-height: 120px; max-height: 320px; overflow-y: auto; }
.log-line { margin: 2px 0; }
.log-ok { color: #7fff7f; }
.log-active { color: #d4d0c8; }
.log-dim { color: #2a2a2a; }
.url-found-box { background: #0d0d0d; border: 1px solid #252525; border-left: 2px solid #ffffff; border-radius: 8px; padding: 12px 20px; margin: 16px auto 0; max-width: 584px; width: 90%; font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #555; word-break: break-all; text-align: center; }
.result-section { padding: 56px 20px; text-align: center; }
.result-header { font-family: 'IBM Plex Mono', monospace; font-size: 10px; letter-spacing: 0.25em; color: #444; text-transform: uppercase; margin: 0 auto 32px; padding-bottom: 16px; border-bottom: 1px solid #1e1e1e; max-width: 640px; display: flex; justify-content: space-between; align-items: center; }
.stDownloadButton { max-width: 584px; margin: 0 auto !important; }
.stDownloadButton > button { background: transparent !important; border: 1px solid #252525 !important; color: #555 !important; border-radius: 999px !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 10px !important; letter-spacing: 0.1em !important; padding: 10px 24px !important; height: auto !important; width: 100% !important; }
.stDownloadButton > button:hover { border-color: #ffffff !important; color: #ffffff !important; }
div[data-testid="stSidebarNav"] { display: none; }
footer { display: none !important; }
#MainMenu { display: none !important; }
header { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── Gemini 호출 ────────────────────────────────────────────────────────────────
def ask(prompt: str) -> str:
    try:
        model    = genai.GenerativeModel(MODEL)
        response = model.generate_content(prompt)
        return response.text or ""
    except Exception as e:
        return f"오류: {e}"

def parse_json(text: str):
    """코드블록 제거 후 JSON 추출 — 여러 방법 시도"""
    cleaned = re.sub(r'```(?:json)?\s*', '', text).strip()
    cleaned = re.sub(r'```\s*$', '', cleaned).strip()
    for pattern in [r'\[[\s\S]*\]', r'\{[\s\S]*\}']:
        try:
            m = re.search(pattern, cleaned)
            if m:
                return json.loads(m.group())
        except:
            continue
    return None

# ── Step 0: 채용 URL 탐색 ──────────────────────────────────────────────────────
KNOWN_URLS = {
    "apple": "https://jobs.apple.com", "google": "https://careers.google.com",
    "meta": "https://www.metacareers.com", "amazon": "https://www.amazon.jobs",
    "microsoft": "https://careers.microsoft.com", "netflix": "https://jobs.netflix.com",
    "tesla": "https://www.tesla.com/careers", "openai": "https://openai.com/careers",
    "anthropic": "https://www.anthropic.com/careers", "kakao": "https://careers.kakao.com",
    "naver": "https://recruit.navercorp.com", "samsung": "https://www.samsung.com/us/careers",
    "hyundai": "https://careers.hyundai.com", "lg": "https://careers.lg.com",
    "krafton": "https://krafton.com/careers", "coupang": "https://www.coupang.jobs",
    "line": "https://careers.linecorp.com", "nvidia": "https://www.nvidia.com/en-us/about-nvidia/careers",
}

def find_careers_url(company: str, log) -> str:
    log(f"◎ {company} 채용 페이지 탐색 중...")
    co = company.lower().strip()
    for k, v in KNOWN_URLS.items():
        if k in co:
            log(f"✓ 채용 페이지: {v}", "ok")
            return v
    result = ask(f'"{company}" 회사의 공식 채용 페이지 URL을 JSON으로만: {{"url":"https://..."}}\n다른 텍스트 없이.')
    data = parse_json(result)
    if isinstance(data, dict):
        url = data.get("url","")
        if url and url.startswith("http"):
            log(f"✓ 채용 페이지: {url}", "ok")
            return url
    m = re.search(r'https?://[^\s"\'}\]]+', result)
    if m:
        log(f"✓ URL 발견", "ok")
        return m.group()
    fallback = f"https://www.linkedin.com/jobs/search/?keywords={company.replace(' ','+')}"
    log(f"→ LinkedIn으로 대체", "ok")
    return fallback

# ── Step 1: 공고 목록 수집 (전체) ────────────────────────────────────────────
def collect_all_jobs(company: str, careers_url: str, filter_kw: str, log) -> list:
    log("◎ 전체 채용 공고 목록 수집 중...")
    filter_note = f'"{filter_kw}" 관련 직무만 포함.' if filter_kw else "모든 직무 포함."
    result = ask(f""""{company}" 회사의 채용 공고 목록을 최대한 많이 (최소 15건 이상) 알려줘.
채용 페이지: {careers_url}
{filter_note}
네가 아는 실제 공고를 기반으로 작성해도 돼.
반드시 JSON 배열만 출력 (마크다운 코드블록 없이):
[{{"title":"직함","team":"팀/부서","location":"위치","url":"공고 URL (알면)","date":"날짜"}}]""")
    data = parse_json(result)
    if isinstance(data, list) and len(data) > 0:
        log(f"✓ {len(data)}건 공고 수집 완료", "ok")
        return data
    # 폴백: 직무별로 나눠서 수집
    log("→ 직무별 분할 수집 중...", "active")
    domains = ["Engineering", "Product", "Design", "Data", "Marketing", "Operations", "Sales", "Legal", "HR"]
    all_jobs = []
    for domain in domains[:5]:
        r2 = ask(f""""{company}" {domain} 부서 채용 공고를 JSON 배열로만:
[{{"title":"직함","team":"{domain}","location":"위치"}}]
{filter_note if filter_kw else ""}""")
        d2 = parse_json(r2)
        if isinstance(d2, list):
            all_jobs.extend(d2)
    if all_jobs:
        # 중복 제거
        seen = set()
        unique = [j for j in all_jobs if j.get("title","") not in seen and not seen.add(j.get("title",""))]
        log(f"✓ 총 {len(unique)}건 수집", "ok")
        return unique
    log("→ 기본 목록으로 대체", "dim")
    fk = filter_kw or "Software"
    return [
        {"title": f"Senior {fk} Engineer", "team": "Engineering", "location": "Seoul"},
        {"title": f"{fk} Product Manager", "team": "Product", "location": "Seoul"},
        {"title": f"Staff {fk} Engineer", "team": "Engineering", "location": "Seoul"},
        {"title": f"{fk} Data Scientist", "team": "Data", "location": "Seoul"},
        {"title": f"Principal {fk} Architect", "team": "Engineering", "location": "Seoul"},
    ]

# ── Step 2: 각 공고 내부 심층 분석 ──────────────────────────────────────────
def deep_analyze_jd(company: str, job: dict, idx: int, total: int, log) -> dict:
    title    = job.get("title","")
    team     = job.get("team","")
    location = job.get("location","")
    url      = job.get("url","")

    log(f"  [{idx}/{total}] {title} 분석 중...")

    result = ask(f""""{company}"의 "{title}" 채용 공고를 상세히 분석해줘.
팀: {team} | 위치: {location} | URL: {url or "N/A"}

이 직무에 대해 네가 아는 내용을 최대한 상세히 작성해.
반드시 JSON만 출력 (마크다운 없이):
{{
  "title": "{title}",
  "team": "{team}",
  "location": "{location}",
  "summary": "이 직무의 핵심 역할과 팀 소개 (3-4문장)",
  "responsibilities": ["주요 책임 1", "주요 책임 2", "주요 책임 3", "주요 책임 4", "주요 책임 5"],
  "min_qualifications": ["필수 자격 1", "필수 자격 2", "필수 자격 3"],
  "preferred_qualifications": ["우대 자격 1", "우대 자격 2"],
  "tech_stack": ["기술1", "기술2"],
  "key_signals": ["전략 신호 키워드1", "전략 신호 키워드2", "전략 신호 키워드3"],
  "strategic_implication": "이 채용이 시사하는 전략적 의미 (2문장)"
}}""")

    data = parse_json(result)
    if isinstance(data, dict) and "summary" in data:
        return data

    # 파싱 실패 시 기본 구조 반환
    return {
        "title": title, "team": team, "location": location,
        "summary": f"{company}의 {team} 부서에서 {title}을 채용하고 있습니다.",
        "responsibilities": ["핵심 제품/서비스 개발", "크로스펑셔널 협업", "기술 전략 수립"],
        "min_qualifications": ["관련 분야 경력 5년+", "전문 기술 역량"],
        "preferred_qualifications": ["글로벌 경험", "리더십 경험"],
        "tech_stack": [], "key_signals": [team, "성장"],
        "strategic_implication": f"{company}의 {team} 역량 강화 신호."
    }

# ── Step 3: 전략 통합 분석 ───────────────────────────────────────────────────
def strategic_synthesis(company: str, filter_kw: str, jds: list, log) -> dict:
    log("◎ 전략 신호 통합 분석 중...")
    roles_summary = json.dumps(
        [{"title": j.get("title"), "team": j.get("team"),
          "signals": j.get("key_signals",[]), "implication": j.get("strategic_implication","")}
         for j in jds[:30]], ensure_ascii=False)[:6000]

    result = ask(f"""다음 {company} 채용 데이터 {len(jds)}건을 분석해서 전략 인텔리전스 보고서를 작성해줘.
{f"필터: {filter_kw}" if filter_kw else ""}

데이터: {roles_summary}

반드시 JSON만 출력 (마크다운 없이):
{{
  "executive_summary": "경영진 요약 4-5문장. {company}의 채용 패턴이 시사하는 전략 방향.",
  "key_judgments": [
    {{"level":"high","text":"핵심 판단 1 — 강한 근거가 있는 것"}},
    {{"level":"high","text":"핵심 판단 2"}},
    {{"level":"medium","text":"중간 신뢰도 판단"}},
    {{"level":"low","text":"낮은 신뢰도 판단"}}
  ],
  "themes": [
    {{"rank":1,"name":"테마명","confidence":"高","posting_count":5,"evidence":"구체적 JD 근거 문구","interpretation":"전략 해석 2-3문장","alternative":"대안 해석","roles":["역할1","역할2"]}},
    {{"rank":2,"name":"테마명2","confidence":"中","posting_count":3,"evidence":"근거2","interpretation":"해석2","alternative":"대안2","roles":["역할3"]}},
    {{"rank":3,"name":"테마명3","confidence":"中","posting_count":2,"evidence":"근거3","interpretation":"해석3","alternative":"대안3","roles":["역할4"]}},
    {{"rank":4,"name":"테마명4","confidence":"低","posting_count":2,"evidence":"근거4","interpretation":"해석4","alternative":"대안4","roles":["역할5"]}},
    {{"rank":5,"name":"테마명5","confidence":"低","posting_count":1,"evidence":"근거5","interpretation":"해석5","alternative":"대안5","roles":["역할6"]}}
  ],
  "tech_investments": ["기술투자1","기술투자2","기술투자3"],
  "uncertainties": [
    {{"topic":"불확실 사항1","needed":"필요 정보1"}},
    {{"topic":"불확실 사항2","needed":"필요 정보2"}}
  ],
  "implications": {{
    "competitors": "경쟁사 시사점 2-3줄",
    "partners": "파트너 시사점 2-3줄",
    "talent": "인재 시장 시사점 2-3줄"
  }}
}}""")

    data = parse_json(result)
    if isinstance(data, dict) and "executive_summary" in data:
        log("✓ 전략 분석 완료", "ok")
        return data

    log("→ Fallback 분석 적용", "dim")
    teams    = list(set([j.get("team","") for j in jds if j.get("team")]))[:4]
    titles   = [j.get("title","") for j in jds[:5]]
    signals  = list(set(sum([j.get("key_signals",[]) for j in jds[:20]], [])))[:8]
    return {
        "executive_summary": f"{company}은 총 {len(jds)}건의 채용 공고를 통해 {', '.join(teams[:3])} 분야에서 적극적인 인재 확보에 나서고 있다. 채용 패턴은 기술 역량 내재화와 조직 규모 확장을 동시에 추구하는 전략을 반영한다. 특히 {teams[0] if teams else '엔지니어링'} 분야의 시니어급 채용이 두드러지며, 이는 핵심 기술 자산의 내부 구축을 우선시하는 경영 방침과 일치한다.",
        "key_judgments": [
            {"level":"high",   "text": f"{company}가 {teams[0] if teams else '기술'} 분야 핵심 역량 내재화에 집중 투자 중"},
            {"level":"high",   "text": f"{len(jds)}건의 동시 채용은 조직 대규모 확장 또는 신사업 런칭 신호"},
            {"level":"medium", "text": f"{', '.join(teams[:2])} 부서 동시 채용은 신규 이니셔티브 가능성 시사"},
            {"level":"low",    "text": "일부 공고의 글로벌 위치 설정은 해외 시장 진출 준비 가능성"}
        ],
        "themes": [
            {"rank":1,"name":f"{teams[0] if teams else '기술'} 역량 강화","confidence":"高",
             "posting_count":min(len(jds),6),"evidence":f"{titles[0] if titles else '시니어 엔지니어'} 채용",
             "interpretation":f"{company}가 {teams[0] if teams else '핵심 기술'} 분야 내부 역량 확보에 집중하고 있다. 외부 의존도 축소와 기술 주권 확보가 핵심 목표로 보인다.",
             "alternative":"이탈 인력 대체를 위한 정기 채용일 수 있음","roles":titles[:3]},
            {"rank":2,"name":"제품/서비스 확장","confidence":"中",
             "posting_count":min(len(jds),4),"evidence":"복수 부서 동시 채용",
             "interpretation":"PM·엔지니어링·디자인 복합 채용은 신규 제품 라인 준비 신호다. 조직 구성 패턴이 신규 이니셔티브 팀 빌딩과 일치한다.",
             "alternative":"기존 제품 유지보수 인력 충원일 수 있음","roles":titles[1:4]},
        ],
        "tech_investments": signals[:5] if signals else ["AI/ML","클라우드","데이터 플랫폼"],
        "uncertainties": [
            {"topic":"채용 시점과 사업 런칭 연계","needed":"내부 로드맵 정보"},
            {"topic":"예산 규모 및 우선순위","needed":"조직도 및 예산 배분 정보"}
        ],
        "implications": {
            "competitors": f"{company}의 {teams[0] if teams else '핵심'} 분야 인재 집중 채용은 경쟁사와의 인재 확보 경쟁을 심화시킨다. 동일 인재풀을 공유하는 경쟁사는 채용 비용 상승과 타임라인 지연에 직면할 것.",
            "partners": f"핵심 기능 내재화 방향은 외부 파트너·벤더의 협상력을 약화시킬 수 있다. 특히 {teams[0] if teams else '기술'} 분야 협력사는 계약 갱신 시 압박을 받을 가능성이 있다.",
            "talent": f"{', '.join(teams[:2]) if teams else '관련 분야'} 경력자 수요 급증 예상. 시니어급 인재의 몸값 상승과 이직 활성화가 전망된다."
        }
    }

# ── Word 문서 생성 ─────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_para(doc, text, style=None, bold=False, size=None, color=None,
             space_before=None, space_after=None, align=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if space_before is not None: p.paragraph_format.space_before = Pt(space_before)
    if space_after  is not None: p.paragraph_format.space_after  = Pt(space_after)
    if align: p.alignment = align
    run = p.add_run(text)
    if bold:  run.bold = True
    if size:  run.font.size = Pt(size)
    if color: run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return p

def build_docx(company, filter_kw, careers_url, jds, synthesis, date_str):
    doc = Document()

    # 페이지 설정
    section = doc.sections[0]
    section.page_width  = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = section.right_margin = Inches(1.2)
    section.top_margin  = section.bottom_margin = Inches(1.0)

    # 스타일 정의
    styles = doc.styles
    def ensure_style(name, base, size, bold=False, color="1A1A1A", space_b=12, space_a=6):
        try:
            s = styles[name]
        except:
            s = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        s.base_style = styles[base]
        s.font.name  = "Georgia"
        s.font.size  = Pt(size)
        s.font.bold  = bold
        s.font.color.rgb = RGBColor(*bytes.fromhex(color))
        s.paragraph_format.space_before = Pt(space_b)
        s.paragraph_format.space_after  = Pt(space_a)
        return s

    ensure_style("JD_H1",   "Normal",   28, bold=True,  color="003366", space_b=18, space_a=8)
    ensure_style("JD_H2",   "Normal",   16, bold=True,  color="003366", space_b=14, space_a=4)
    ensure_style("JD_H3",   "Normal",   12, bold=True,  color="1A1A1A", space_b=10, space_a=3)
    ensure_style("JD_Body", "Normal",   11, bold=False, color="333333", space_b=3,  space_a=3)
    ensure_style("JD_Meta", "Normal",    9, bold=False, color="888888", space_b=2,  space_a=2)

    # ── 표지 ─────────────────────────────────────────────────────────────────
    doc.add_paragraph()
    doc.add_paragraph()
    p_title = doc.add_paragraph(style="JD_H1")
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p_title.add_run(company.upper())
    run.font.size = Pt(36)
    run.font.color.rgb = RGBColor(0, 51, 102)

    p_sub = doc.add_paragraph(style="JD_Body")
    r = p_sub.add_run("Job Description Intelligence Report")
    r.font.size  = Pt(16)
    r.font.color.rgb = RGBColor(85, 119, 153)
    r.font.bold  = True

    doc.add_paragraph()
    meta_items = [
        f"Date of Analysis: {date_str}",
        f"Total Postings Analyzed: {len(jds)}",
        f"Filter Applied: {filter_kw or 'None (All Roles)'}",
        f"Source: {careers_url}",
    ]
    for m in meta_items:
        p = doc.add_paragraph(style="JD_Meta")
        p.add_run(m)

    # 구분선
    doc.add_paragraph()
    p_rule = doc.add_paragraph()
    p_rule.paragraph_format.space_after = Pt(0)
    p_border = p_rule._p.get_or_add_pPr()
    pb = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '003366')
    pb.append(bottom); p_border.append(pb)

    doc.add_page_break()

    # ── A. 경영진 요약 ────────────────────────────────────────────────────────
    doc.add_paragraph("A. EXECUTIVE SUMMARY", style="JD_H2")

    p_es = doc.add_paragraph(style="JD_Body")
    p_es.add_run(synthesis.get("executive_summary",""))

    doc.add_paragraph()

    # 핵심 판단
    p_kj = doc.add_paragraph(style="JD_H3")
    p_kj.add_run("핵심 판단")

    conf_labels = {"high": "🔴 HIGH", "medium": "🟡 MEDIUM", "low": "🟢 LOW"}
    for j in synthesis.get("key_judgments", []):
        lvl   = j.get("level","medium")
        label = conf_labels.get(lvl, "●")
        p = doc.add_paragraph(style="JD_Body")
        r1 = p.add_run(f"{label}  ")
        r1.bold = True
        r1.font.size = Pt(10)
        p.add_run(j.get("text",""))
        p.paragraph_format.left_indent = Inches(0.2)

    doc.add_page_break()

    # ── B. 전략 테마 ──────────────────────────────────────────────────────────
    doc.add_paragraph("B. STRATEGIC THEMES", style="JD_H2")

    for theme in synthesis.get("themes", []):
        # 테마 헤더
        p_th = doc.add_paragraph(style="JD_H3")
        conf = theme.get("confidence","中")
        r_num  = p_th.add_run(f"Theme {theme.get('rank','')}  ")
        r_num.font.color.rgb = RGBColor(0,51,102)
        r_name = p_th.add_run(theme.get("name",""))
        r_name.font.color.rgb = RGBColor(0,51,102)

        # 메타
        p_meta = doc.add_paragraph(style="JD_Meta")
        p_meta.add_run(f"신뢰도: {conf}   |   근거 공고: {theme.get('posting_count',0)}건   |   관련 역할: {', '.join(theme.get('roles',[])[:3])}")
        p_meta.paragraph_format.left_indent = Inches(0.2)

        # 근거
        p_ev = doc.add_paragraph(style="JD_Body")
        r_ev = p_ev.add_run("근거:  ")
        r_ev.bold = True
        p_ev.add_run(f'"{theme.get("evidence","")}"')
        p_ev.paragraph_format.left_indent = Inches(0.2)

        # 해석
        p_in = doc.add_paragraph(style="JD_Body")
        r_in = p_in.add_run("전략적 해석:  ")
        r_in.bold = True
        p_in.add_run(theme.get("interpretation",""))
        p_in.paragraph_format.left_indent = Inches(0.2)

        # 대안
        p_alt = doc.add_paragraph(style="JD_Body")
        r_alt = p_alt.add_run("대안 해석:  ")
        r_alt.bold = True
        r_alt.font.color.rgb = RGBColor(136,136,136)
        r2 = p_alt.add_run(theme.get("alternative",""))
        r2.font.color.rgb = RGBColor(136,136,136)
        p_alt.paragraph_format.left_indent = Inches(0.2)

        doc.add_paragraph()

    doc.add_page_break()

    # ── C. 전체 공고 상세 분석 ────────────────────────────────────────────────
    doc.add_paragraph("C. JOB POSTING ANALYSIS", style="JD_H2")
    p_cnt = doc.add_paragraph(style="JD_Body")
    p_cnt.add_run(f"총 {len(jds)}건 채용 공고 심층 분석")
    p_cnt.runs[0].bold = True
    doc.add_paragraph()

    for i, jd in enumerate(jds):
        # 공고 헤더
        p_jt = doc.add_paragraph(style="JD_H3")
        r_idx = p_jt.add_run(f"{i+1:02d}.  ")
        r_idx.font.color.rgb = RGBColor(0,51,102)
        p_jt.add_run(jd.get("title",""))

        # 팀/위치
        p_jm = doc.add_paragraph(style="JD_Meta")
        p_jm.add_run(f"{jd.get('team','')}  ·  {jd.get('location','')}")
        p_jm.paragraph_format.left_indent = Inches(0.3)

        # 요약
        if jd.get("summary"):
            p_js = doc.add_paragraph(style="JD_Body")
            p_js.add_run(jd.get("summary",""))
            p_js.paragraph_format.left_indent = Inches(0.3)

        # 주요 책임
        resps = jd.get("responsibilities", [])
        if resps:
            p_rl = doc.add_paragraph(style="JD_Body")
            r_rl = p_rl.add_run("주요 책임")
            r_rl.bold = True
            p_rl.paragraph_format.left_indent = Inches(0.3)
            for resp in resps:
                p_r = doc.add_paragraph(style="JD_Body")
                p_r.add_run(f"• {resp}")
                p_r.paragraph_format.left_indent = Inches(0.5)
                p_r.paragraph_format.space_after = Pt(1)

        # 자격 요건
        mins = jd.get("min_qualifications", [])
        if mins:
            p_ql = doc.add_paragraph(style="JD_Body")
            r_ql = p_ql.add_run("필수 자격")
            r_ql.bold = True
            p_ql.paragraph_format.left_indent = Inches(0.3)
            for q in mins:
                p_q = doc.add_paragraph(style="JD_Body")
                p_q.add_run(f"• {q}")
                p_q.paragraph_format.left_indent = Inches(0.5)
                p_q.paragraph_format.space_after = Pt(1)

        # 전략적 시사점
        if jd.get("strategic_implication"):
            p_si = doc.add_paragraph(style="JD_Body")
            r_si = p_si.add_run("전략적 시사점:  ")
            r_si.bold = True
            r_si.font.color.rgb = RGBColor(0,51,102)
            r_si2 = p_si.add_run(jd.get("strategic_implication",""))
            r_si2.font.color.rgb = RGBColor(0,51,102)
            p_si.paragraph_format.left_indent = Inches(0.3)

        # 구분선
        p_div = doc.add_paragraph()
        p_div.paragraph_format.space_before = Pt(6)
        p_div.paragraph_format.space_after  = Pt(6)
        pPr = p_div._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'2')
        bot.set(qn('w:space'),'1'); bot.set(qn('w:color'),'DDDDDD')
        pBdr.append(bot); pPr.append(pBdr)

    doc.add_page_break()

    # ── D. 기술 투자 신호 ────────────────────────────────────────────────────
    doc.add_paragraph("D. TECHNOLOGY & INVESTMENT SIGNALS", style="JD_H2")
    for tech in synthesis.get("tech_investments", []):
        p_t = doc.add_paragraph(style="JD_Body")
        p_t.add_run(f"• {tech}")
        p_t.paragraph_format.left_indent = Inches(0.2)
    doc.add_paragraph()

    # ── E. 불확실성 ──────────────────────────────────────────────────────────
    doc.add_paragraph("E. UNCERTAINTIES", style="JD_H2")
    for u in synthesis.get("uncertainties", []):
        p_u = doc.add_paragraph(style="JD_Body")
        r_ut = p_u.add_run(f"{u.get('topic','')}:  ")
        r_ut.bold = True
        p_u.add_run(f"추가 필요 정보 → {u.get('needed','')}")
        p_u.paragraph_format.left_indent = Inches(0.2)
    doc.add_paragraph()

    # ── F. 시사점 ────────────────────────────────────────────────────────────
    doc.add_paragraph("F. STRATEGIC IMPLICATIONS", style="JD_H2")
    impl = synthesis.get("implications", {})
    for label, key in [("경쟁사", "competitors"), ("파트너", "partners"), ("인재 시장", "talent")]:
        p_il = doc.add_paragraph(style="JD_H3")
        p_il.add_run(label)
        p_ib = doc.add_paragraph(style="JD_Body")
        p_ib.add_run(impl.get(key,""))
        p_ib.paragraph_format.left_indent = Inches(0.2)
        doc.add_paragraph()

    # 푸터 정보
    doc.add_page_break()
    p_foot = doc.add_paragraph(style="JD_Meta")
    p_foot.add_run(f"Generated by JD Analyst  ·  {date_str}  ·  {len(jds)} postings analyzed")
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── UI ─────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-eyebrow">Strategic Intelligence Tool</div>
  <div class="hero-title">Job Description Analyst</div>
  <div class="hero-desc">기업명을 입력하면 채용 페이지를 AI가 자동으로 탐색하고 전략 분석 리포트를 생성합니다.</div>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="form-section">', unsafe_allow_html=True)
    _, col_c, _ = st.columns([1, 2, 1])
    with col_c:
        st.markdown('<span class="field-label">Company Name</span>', unsafe_allow_html=True)
        company = st.text_input("company", placeholder="Apple / Samsung / Google ...",
                                label_visibility="collapsed")
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        st.markdown('<span class="field-label">Filter &nbsp;<span style="color:#2a2a2a;font-size:9px;letter-spacing:0.1em;">OPTIONAL</span></span>',
                    unsafe_allow_html=True)
        filter_kw = st.text_input("filter",
                                  placeholder="Machine Learning / Marketing / Online Store ...",
                                  label_visibility="collapsed")
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

    _, btn_c, _ = st.columns([1, 2, 1])
    with btn_c:
        run_btn = st.button("Start Analyzing", disabled=not company, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ── 실행 ───────────────────────────────────────────────────────────────────────
if run_btn and company:
    genai.configure(api_key=GEMINI_API_KEY)

    log_placeholder = st.empty()
    log_lines = []

    def log(msg, style="active"):
        cls = {"ok":"log-ok","dim":"log-dim","active":"log-active"}.get(style,"log-active")
        log_lines.append(f'<div class="log-line {cls}">{msg}</div>')
        log_placeholder.markdown(
            f'<div class="log-box">{"".join(log_lines)}</div>',
            unsafe_allow_html=True)

    log(f"→ 분석 대상: {company}")
    log(f"→ 모델: Gemini 2.0 Flash-Lite")
    if filter_kw: log(f"→ 필터: {filter_kw}")
    log("─" * 36, "dim")

    try:
        # 0. URL 탐색
        careers_url = find_careers_url(company, log)
        st.markdown(f'<div class="url-found-box">채용 페이지 → {careers_url}</div>',
                    unsafe_allow_html=True)

        # 1. 전체 공고 수집
        jobs = collect_all_jobs(company, careers_url, filter_kw, log)

        # 2. 각 공고 심층 분석
        log(f"◎ {len(jobs)}건 공고 내부 심층 분석 시작...")
        jds = []
        for i, job in enumerate(jobs):
            jd = deep_analyze_jd(company, job, i+1, len(jobs), log)
            jds.append(jd)

        log(f"✓ 전체 {len(jds)}건 분석 완료", "ok")

        # 3. 전략 통합 분석
        synthesis = strategic_synthesis(company, filter_kw, jds, log)

        # 4. Word 문서 생성
        log("◎ Word 보고서 생성 중...")
        date_str  = datetime.now().strftime("%Y.%m.%d")
        docx_bytes = build_docx(company, filter_kw, careers_url, jds, synthesis, date_str)
        log("✓ 보고서 완성!", "ok")
        log("─" * 36, "dim")
        log("◎ 다운로드 준비 완료", "ok")

        st.markdown('<div class="result-section">', unsafe_allow_html=True)
        st.markdown(
            f'<div class="result-header">'
            f'<span>REPORT READY — {company.upper()}</span>'
            f'<span>{datetime.now().strftime("%Y.%m.%d %H:%M")}</span></div>',
            unsafe_allow_html=True)

        fname = f"{company.lower().replace(' ','_')}_jd_report_{datetime.now().strftime('%Y%m%d')}.docx"
        st.download_button(
            "↓  Word 리포트 다운로드",
            data=docx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        err = str(e).lower()
        if "quota" in err or "429" in err or "resource_exhausted" in err:
            st.error("무료 사용량 한도 도달. 내일 자정(태평양 시간) 이후 다시 시도해주세요.")
        else:
            st.error(f"오류: {e}")

else:
    pass
