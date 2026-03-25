import streamlit as st
import google.generativeai as genai
import json
import re
import io
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

st.set_page_config(
    page_title="JD Analyst",
    page_icon="◎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap');

*, html, body { font-family: 'IBM Plex Sans', sans-serif; }
.stApp {
    background: linear-gradient(180deg, #000000 0%, #141414 50%, #2a2a2a 100%);
    color: #d4d0c8;
}
.block-container { padding: 0 !important; max-width: 100% !important; }

/* 전체 중앙 정렬 래퍼 */
.hero {
    padding: 80px 20px 52px;
    border-bottom: 1px solid #1e1e1e;
    text-align: center;
    display: flex;
    flex-direction: column;
    align-items: center;
}
.hero-eyebrow {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.25em;
    color: #444;
    text-transform: uppercase;
    margin-bottom: 20px;
}
.hero-title {
    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    font-size: 54px;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -0.03em;
    line-height: 1.0;
    margin-bottom: 18px;
    text-shadow:
        0 0 30px rgba(255,255,255,0.55),
        0 0 70px rgba(255,255,255,0.2),
        0 0 120px rgba(255,255,255,0.08);
}
.hero-desc {
    font-size: 14px;
    color: #555;
    font-weight: 300;
    line-height: 1.65;
    text-align: center;
}

/* 폼 섹션: 중앙정렬 + Google 검색창 너비 */
.form-section {
    padding: 52px 20px;
    border-bottom: 1px solid #1e1e1e;
    display: flex;
    flex-direction: column;
    align-items: center;
}
.input-wrapper {
    width: 100%;
    max-width: 584px;  /* Google 검색창 기준 */
}
.field-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.15em;
    color: #444;
    text-transform: uppercase;
    margin-bottom: 8px;
    display: block;
    text-align: center;
}
.field-label-sm {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.12em;
    color: #333;
    text-transform: uppercase;
    margin-bottom: 6px;
    display: block;
    text-align: center;
}

/* 입력창: pill + Google 너비 */
.stTextInput { max-width: 584px; margin: 0 auto !important; }
.stTextInput > div > div > input {
    background: #0d0d0d !important;
    border: 1px solid #252525 !important;
    border-radius: 999px !important;
    color: #d4d0c8 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 13px !important;
    padding: 14px 22px !important;
    transition: border-color 0.2s, box-shadow 0.2s !important;
    text-align: center !important;
}
.stTextInput > div > div > input:focus {
    border-color: #ffffff !important;
    box-shadow: 0 0 0 1px rgba(255,255,255,0.1), 0 0 20px rgba(255,255,255,0.05) !important;
}
.stTextInput > div > div > input::placeholder { color: #2e2e2e !important; text-align: center !important; }

/* 버튼: pill + 중앙 */
.stButton { max-width: 584px; margin: 0 auto !important; }
.stButton > button {
    background: #ffffff !important;
    color: #000000 !important;
    border: none !important;
    border-radius: 999px !important;
    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    padding: 14px 36px !important;
    height: auto !important;
    transition: opacity 0.15s, box-shadow 0.15s !important;
    width: 100% !important;
    box-shadow: 0 0 24px rgba(255,255,255,0.15) !important;
}
.stButton > button:hover { opacity: 0.88 !important; box-shadow: 0 0 40px rgba(255,255,255,0.25) !important; }
.stButton > button:disabled { background: #1a1a1a !important; color: #333 !important; box-shadow: none !important; }

/* 로그, 결과 */
.log-box {
    background: #080808;
    border: 1px solid #1e1e1e;
    border-radius: 16px;
    padding: 20px 28px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #444;
    line-height: 1.8;
    margin: 24px auto;
    max-width: 640px;
    width: 90%;
    min-height: 120px;
    max-height: 280px;
    overflow-y: auto;
}
.log-line { margin: 2px 0; }
.log-ok { color: #7fff7f; }
.log-active { color: #d4d0c8; }
.log-dim { color: #2a2a2a; }

.url-found-box {
    background: #0d0d0d;
    border: 1px solid #252525;
    border-left: 2px solid #ffffff;
    border-radius: 8px;
    padding: 12px 20px;
    margin: 16px auto 0;
    max-width: 584px;
    width: 90%;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #555;
    word-break: break-all;
    text-align: center;
}

.result-section { padding: 56px 20px; text-align: center; }
.result-header {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.25em;
    color: #444;
    text-transform: uppercase;
    margin: 0 auto 32px;
    padding-bottom: 16px;
    border-bottom: 1px solid #1e1e1e;
    max-width: 640px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.stDownloadButton { max-width: 584px; margin: 0 auto !important; }
.stDownloadButton > button {
    background: transparent !important;
    border: 1px solid #252525 !important;
    color: #555 !important;
    border-radius: 999px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 10px !important;
    letter-spacing: 0.1em !important;
    padding: 10px 24px !important;
    height: auto !important;
    box-shadow: none !important;
    width: 100% !important;
}
.stDownloadButton > button:hover { border-color: #ffffff !important; color: #ffffff !important; }

div[data-testid="stSidebarNav"] { display: none; }
footer { display: none !important; }
#MainMenu { display: none !important; }
header { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── 에이전트 설정 ──────────────────────────────────────────────────────────────
MODEL = "gemini-2.5-flash-lite-preview-06-17"  # 무료 티어 최고 한도 (1,000회/일)


def run_agent(client, system, prompt, use_search=False, **kwargs):
    """Gemini Flash-Lite 호출. use_search는 grounding으로 처리."""
    try:
        tools = []
        if use_search:
            tools = [{"google_search": {}}]

        full_prompt = f"{system}\n\n{prompt}" if system else prompt

        if tools:
            response = client.models.generate_content(
                model=MODEL,
                contents=full_prompt,
                config={"tools": tools}
            )
        else:
            response = client.models.generate_content(
                model=MODEL,
                contents=full_prompt
            )
        return response.text or ""
    except Exception as e:
        return f"오류: {e}"


def find_careers_url(client, company, log):
    log(f"◎ {company} 채용 페이지 탐색 중...")
    prompt = f""""{company}" 회사의 공식 채용 페이지 URL을 알려줘.
잘 알려진 회사라면 알고 있는 URL을 바로 사용해.
JSON으로만 반환: {{"url": "https://...", "source": "공식/LinkedIn/기타"}}
다른 텍스트 없이 JSON만."""
    result = run_agent(client, None, prompt, use_search=False)
    try:
        match = re.search(r'\{[\s\S]*?\}', result)
        if match:
            data = json.loads(match.group())
            url = data.get("url", "")
            if url and url.startswith("http"):
                log(f"✓ 채용 페이지 발견 ({data.get('source','')})", "ok")
                return url
    except:
        pass
    url_match = re.search(r'https?://[^\s"\'}\]]+', result)
    if url_match:
        log("✓ URL 발견", "ok")
        return url_match.group()
    log("! 채용 페이지를 찾지 못했습니다.", "dim")
    return None


def collect_jobs(client, company, careers_url, filter_keyword, log):
    log("◎ 채용 공고 목록 수집 중...")
    prompt = f""""{company}" 회사의 최신 채용 공고 목록을 알려줘.
채용 페이지: {careers_url}
필터: {filter_keyword or '없음 (전체)'}

네가 알고 있는 이 회사의 실제 채용 공고를 기반으로 최대 10건을 JSON 배열로만 반환해:
[{{"title":"직함","team":"팀/부서","location":"위치","date":"날짜 또는 최근"}}]
JSON만 반환, 다른 텍스트 없이."""
    result = run_agent(client, None, prompt, use_search=False)
    try:
        match = re.search(r'\[[\s\S]*\]', result)
        if match:
            jobs = json.loads(match.group())
            log(f"✓ {len(jobs)}건 수집", "ok")
            return jobs
    except:
        pass
    log("! 공고 수집 실패", "dim")
    return []


def collect_jd_details(client, jobs, company, log):
    log(f"◎ {len(jobs)}건 JD 상세 수집 중...")
    jds = []
    for i, job in enumerate(jobs):
        title = job.get("title", "unknown")
        log(f"  [{i+1}/{len(jobs)}] {title}")
        prompt = f""""{company}"의 "{title}" 채용 공고 상세를 알려줘.
네가 알고 있는 정보 기반으로 작성해도 돼.

JSON으로만 반환:
{{
  "title": "{title}",
  "team": "{job.get('team','')}",
  "location": "{job.get('location','')}",
  "summary": "팀·역할 요약 2문장",
  "responsibilities": ["책임1","책임2","책임3"],
  "min_qualifications": ["최소자격1","최소자격2"],
  "preferred_qualifications": ["우대1"],
  "key_signals": ["전략키워드1","전략키워드2"]
}}"""
        result = run_agent(client, None, prompt, use_search=False)
        try:
            match = re.search(r'\{[\s\S]*\}', result)
            if match:
                jds.append(json.loads(match.group()))
                continue
        except:
            pass
        jds.append({**job, "summary": "", "responsibilities": [],
                    "min_qualifications": [], "preferred_qualifications": [], "key_signals": []})
    log(f"✓ {len(jds)}건 완료", "ok")
    return jds


def extract_signals(client, jds, company, log):
    log("◎ 전략 신호 분석 중...")
    batch_text = json.dumps(jds[:10], ensure_ascii=False)[:7000]
    prompt = f"""{company} JD에서 전략 신호 추출. JSON으로만 반환:
{{
  "growth":[{{"signal":"","evidence":"","roles":[],"confidence":"high"}}],
  "technology":[...], "operations":[...], "customer_experience":[...],
  "channel":[...], "geographic":[...],
  "patterns":[{{"pattern":"","count":0,"roles":[]}}],
  "team_names":[], "tech_stack":[]
}}
데이터: {batch_text}"""
    result = run_agent(client, None, prompt, use_search=False)
    try:
        match = re.search(r'\{[\s\S]*\}', result)
        if match:
            log("✓ 신호 추출 완료", "ok")
            return json.loads(match.group())
    except:
        pass
    log("✓ 분석 완료", "ok")
    return {}


def synthesize_analysis(client, company, filter_keyword, jds, signals, log):
    log("◎ 전략 분석 합성 중...")
    context = {
        "company": company, "filter": filter_keyword,
        "date": datetime.now().strftime("%Y-%m-%d"), "total": len(jds),
        "signals": signals,
        "roles": [{"title": j.get("title"), "team": j.get("team"),
                   "location": j.get("location"), "signals": j.get("key_signals", [])}
                  for j in jds[:15]]
    }
    prompt = f"""다음 채용 데이터를 분석해서 아래 JSON 구조로 반환해줘.
절대 JSON 외 텍스트 없이.

데이터: {json.dumps(context, ensure_ascii=False)[:8000]}

반환 형식:
{{
  "executive_summary": "경영진 요약 3-4문장",
  "key_judgments": [
    {{"level": "high", "text": "핵심 판단 1"}},
    {{"level": "high", "text": "핵심 판단 2"}},
    {{"level": "medium", "text": "핵심 판단 3"}},
    {{"level": "low", "text": "핵심 판단 4"}}
  ],
  "themes": [
    {{
      "rank": 1, "name": "테마명", "confidence": "高",
      "posting_count": 3,
      "evidence": "JD 근거 문구",
      "interpretation": "전략적 해석 2문장",
      "alternative": "대안 해석 1문장",
      "roles": ["역할1", "역할2"]
    }}
  ],
  "evidence_map": [
    {{"title": "직함", "team": "팀", "location": "위치", "theme": "테마", "signal": "●●●"}}
  ],
  "investments": [
    {{"area": "투자 영역", "detail": "구체적 내용", "evidence": "JD 근거"}}
  ],
  "uncertainties": [
    {{"topic": "불확실 사항", "current": "현재 신호", "needed": "필요 정보"}}
  ],
  "implications": {{
    "competitors": "경쟁사 시사점 2-3줄",
    "partners": "파트너 시사점 2-3줄",
    "talent": "인재 시장 시사점 2-3줄"
  }}
}}"""
    result = run_agent(client, None, prompt, use_search=False)
    try:
        match = re.search(r'\{[\s\S]*\}', result)
        if match:
            log("✓ 분석 완료", "ok")
            return json.loads(match.group())
    except:
        pass
    log("! JSON 파싱 실패 — 원문 반환", "dim")
    return {"raw": result}


# ── PPTX 생성 — McKinsey 타이포그래피 스타일 ──────────────────────────────────
def hex_to_rgb(hex_str):
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def tb(slide, text, x, y, w, h,
       size=11, bold=False, italic=False,
       color="1A1A1A", align=PP_ALIGN.LEFT,
       font="Georgia", wrap=True):
    """McKinsey 스타일 텍스트박스 헬퍼"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf  = box.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    r   = p.add_run()
    r.text = text
    r.font.size   = Pt(size)
    r.font.bold   = bold
    r.font.italic = italic
    r.font.color.rgb = hex_to_rgb(color)
    r.font.name   = font
    return box

def rect(slide, x, y, w, h, fill="FFFFFF", line=None, lw=0.5):
    from pptx.util import Pt as PPt
    shp = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = hex_to_rgb(fill)
    if line:
        shp.line.color.rgb = hex_to_rgb(line)
        shp.line.width = PPt(lw)
    else:
        shp.line.fill.background()
    return shp

def footer(slide, company, date_str, prs_w=13.33):
    """공통 푸터: 좌 회사명 / 우 날짜 + 페이지"""
    rect(slide, 0, 7.18, prs_w, 0.02, fill="1A1A1A")
    tb(slide, company.upper(), 0.5, 7.22, 6, 0.25,
       size=7, color="999999", font="Calibri")
    tb(slide, date_str, 6.5, 7.22, 6.33, 0.25,
       size=7, color="999999", align=PP_ALIGN.RIGHT, font="Calibri")

def build_pptx(company, filter_keyword, careers_url, jds, analysis, date_str):
    """
    McKinsey 보고서 타이포그래피 기준:
    - Serif (Georgia): 제목, 섹션 헤더
    - Sans (Calibri): 본문, 레이블, 데이터
    - 팔레트: 흰 배경 / 짙은 네이비(003366) 강조 / 미드그레이(666666) 부제
    - 레이아웃: 여백 0.5" 이상, 텍스트 좌정렬, 타이틀-본문 대비 강조
    """
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    layout = prs.slide_layouts[6]

    # 팔레트
    NAVY   = "003366"
    BLACK  = "1A1A1A"
    MID    = "555555"
    LIGHT  = "999999"
    RULE   = "CCCCCC"
    BG     = "FFFFFF"
    BGALT  = "F5F5F5"
    ACCENT = "003366"  # McKinsey 블루

    def slide_bg(s, color=BG):
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = hex_to_rgb(color)

    def section_label(s, text, y=0.38):
        """SECTION LABEL — 작은 대문자 추적 텍스트"""
        tb(s, text, 0.5, y, 12, 0.25,
           size=7, color=LIGHT, font="Calibri")
        rect(s, 0.5, y+0.28, 12.33, 0.01, fill=RULE)

    # ── 슬라이드 1: 표지 ──────────────────────────────────────────────────────
    s1 = prs.slides.add_slide(layout)
    slide_bg(s1, "003366")  # 네이비 배경

    # 좌측 흰 세로선
    rect(s1, 0.5, 0.9, 0.04, 5.5, fill="FFFFFF")

    tb(s1, "STRATEGIC INTELLIGENCE REPORT", 0.72, 0.9, 11, 0.35,
       size=8, color="7FAACC", font="Calibri")

    tb(s1, company.upper(), 0.72, 1.35, 11, 1.8,
       size=64, bold=True, color="FFFFFF", font="Georgia")

    tb(s1, "Job Description Intelligence Analysis", 0.72, 3.2, 10, 0.55,
       size=18, italic=True, color="AACCEE", font="Georgia")

    meta = f"Date of Analysis: {date_str}     Postings Reviewed: {len(jds)}"
    if filter_keyword:
        meta += f"     Filter: {filter_keyword}"
    tb(s1, meta, 0.72, 3.95, 11, 0.35,
       size=9, color="7FAACC", font="Calibri")

    rect(s1, 0.72, 4.4, 3.5, 0.01, fill="7FAACC")

    if careers_url:
        tb(s1, careers_url, 0.72, 4.52, 11, 0.3,
           size=8, color="557799", font="Calibri")

    tb(s1, "CONFIDENTIAL", 0.5, 7.1, 6, 0.28,
       size=7, color="557799", font="Calibri")
    tb(s1, "JD Analyst", 7.0, 7.1, 6.3, 0.28,
       size=7, color="557799", align=PP_ALIGN.RIGHT, font="Calibri")

    # ── 슬라이드 2: 경영진 요약 ───────────────────────────────────────────────
    s2 = prs.slides.add_slide(layout)
    slide_bg(s2)
    section_label(s2, "EXECUTIVE SUMMARY")

    tb(s2, "핵심 전략 판단", 0.5, 0.65, 11, 0.75,
       size=32, bold=True, color=BLACK, font="Georgia")

    summary = analysis.get("executive_summary", "")
    tb(s2, summary, 0.5, 1.55, 8.5, 1.0,
       size=12, color=MID, font="Calibri", italic=True)

    rect(s2, 0.5, 2.7, 12.33, 0.01, fill=RULE)
    tb(s2, "KEY JUDGMENTS", 0.5, 2.8, 5, 0.25,
       size=7, color=LIGHT, font="Calibri")

    judgments = analysis.get("key_judgments", [])
    conf_labels = {"high": "HIGH CONFIDENCE", "medium": "MEDIUM", "low": "LOW"}
    conf_colors = {"high": "003366", "medium": "666666", "low": "AAAAAA"}
    y = 3.15
    for j in judgments[:5]:
        lvl = j.get("level","medium")
        # 좌측 네이비 마커
        rect(s2, 0.5, y, 0.04, 0.36, fill=conf_colors.get(lvl, "888888"))
        tb(s2, conf_labels.get(lvl,""), 0.65, y, 2.2, 0.18,
           size=6, color=conf_colors.get(lvl,"888888"), font="Calibri", bold=True)
        tb(s2, j.get("text",""), 0.65, y+0.17, 11.8, 0.22,
           size=11, color=BLACK, font="Calibri")
        y += 0.52

    footer(s2, company, date_str)

    # ── 슬라이드 3: 전략 테마 ─────────────────────────────────────────────────
    themes = analysis.get("themes", [])[:5]
    s3 = prs.slides.add_slide(layout)
    slide_bg(s3)
    section_label(s3, "STRATEGIC THEMES")

    tb(s3, "전략 테마 Top 5", 0.5, 0.65, 11, 0.65,
       size=32, bold=True, color=BLACK, font="Georgia")

    # 5컬럼 카드
    cw = 2.42
    gaps = [0.42, 2.94, 5.46, 7.98, 10.5]
    for i, th in enumerate(themes[:5]):
        if i >= len(gaps): break
        cx = gaps[i]
        cy = 1.48
        # 카드 배경
        rect(s3, cx, cy, cw, 5.4, fill=BGALT)
        # 상단 네이비 바
        conf = th.get("confidence","中")
        bar_color = {"高": NAVY, "中": "557799", "低": LIGHT}.get(conf, NAVY)
        rect(s3, cx, cy, cw, 0.06, fill=bar_color)
        # 번호
        tb(s3, f"{i+1:02d}", cx+0.15, cy+0.12, cw-0.2, 0.3,
           size=10, bold=True, color=bar_color, font="Georgia")
        # 신뢰도
        tb(s3, f"신뢰도 {conf}  /  {th.get('posting_count',0)}건",
           cx+0.15, cy+0.4, cw-0.2, 0.25,
           size=7, color=LIGHT, font="Calibri")
        # 테마명
        tb(s3, th.get("name",""), cx+0.15, cy+0.65, cw-0.2, 0.72,
           size=13, bold=True, color=BLACK, font="Georgia")
        # 구분선
        rect(s3, cx+0.15, cy+1.42, cw-0.3, 0.01, fill=RULE)
        # 해석
        interp = th.get("interpretation","")[:100]
        tb(s3, interp, cx+0.15, cy+1.52, cw-0.2, 1.6,
           size=9, color=MID, font="Calibri")
        # 대안해석
        alt = th.get("alternative","")[:70]
        tb(s3, f"Alt: {alt}", cx+0.15, cy+3.3, cw-0.2, 0.9,
           size=7.5, italic=True, color=LIGHT, font="Calibri")

    footer(s3, company, date_str)

    # ── 슬라이드 4: 증거 맵 ───────────────────────────────────────────────────
    evidence = analysis.get("evidence_map", [])[:14]
    s4 = prs.slides.add_slide(layout)
    slide_bg(s4)
    section_label(s4, "EVIDENCE MAP")

    tb(s4, "공고별 증거", 0.5, 0.65, 11, 0.65,
       size=32, bold=True, color=BLACK, font="Georgia")

    # 헤더 행
    hdrs  = ["직함", "팀", "위치", "전략 테마", "신호 강도"]
    cxs   = [0.5, 4.3, 7.0, 9.0, 12.0]
    cws   = [3.7, 2.6, 1.9, 2.9, 1.15]
    rect(s4, 0.5, 1.48, 12.83, 0.38, fill=NAVY)
    for hi, hd in enumerate(hdrs):
        tb(s4, hd, cxs[hi]+0.08, 1.52, cws[hi], 0.3,
           size=8, bold=True, color="FFFFFF", font="Calibri")

    for ri, ev in enumerate(evidence):
        ry = 1.88 + ri * 0.36
        bg = BGALT if ri % 2 == 0 else BG
        rect(s4, 0.5, ry, 12.83, 0.35, fill=bg)
        vals = [ev.get("title","")[:40], ev.get("team","")[:22],
                ev.get("location","")[:18], ev.get("theme","")[:28], ev.get("signal","●○○")]
        for ci, v in enumerate(vals):
            col = BLACK if ci == 0 else MID
            tb(s4, v, cxs[ci]+0.08, ry+0.04, cws[ci], 0.28,
               size=9, color=col, font="Calibri", bold=(ci==0))

    footer(s4, company, date_str)

    # ── 슬라이드 5: 투자 영역 ─────────────────────────────────────────────────
    investments = analysis.get("investments", [])[:6]
    s5 = prs.slides.add_slide(layout)
    slide_bg(s5)
    section_label(s5, "INVESTMENT SIGNALS")

    tb(s5, f"{company}이 실제로 투자하는 것", 0.5, 0.65, 11, 0.65,
       size=28, bold=True, color=BLACK, font="Georgia")

    card_xs = [0.5, 4.63, 8.76]
    cw5     = 3.98
    for i, inv in enumerate(investments[:6]):
        row = i // 3
        col = i %  3
        cx  = card_xs[col]
        cy  = 1.52 + row * 2.72
        rect(s5, cx, cy, cw5, 2.55, fill=BGALT)
        rect(s5, cx, cy, 0.05, 2.55, fill=NAVY)
        tb(s5, inv.get("area",""), cx+0.2, cy+0.15, cw5-0.3, 0.45,
           size=13, bold=True, color=BLACK, font="Georgia")
        rect(s5, cx+0.2, cy+0.62, cw5-0.4, 0.01, fill=RULE)
        tb(s5, inv.get("detail","")[:95], cx+0.2, cy+0.72, cw5-0.3, 1.0,
           size=9.5, color=MID, font="Calibri")
        tb(s5, f'"{inv.get("evidence","")[:55]}"', cx+0.2, cy+1.82, cw5-0.3, 0.55,
           size=8, italic=True, color=LIGHT, font="Calibri")

    footer(s5, company, date_str)

    # ── 슬라이드 6: 불확실성 & 시사점 ────────────────────────────────────────
    s6 = prs.slides.add_slide(layout)
    slide_bg(s6)
    section_label(s6, "UNCERTAINTIES & IMPLICATIONS")

    tb(s6, "불확실성 / 전략적 시사점", 0.5, 0.65, 11, 0.65,
       size=28, bold=True, color=BLACK, font="Georgia")

    # 불확실성 (좌)
    tb(s6, "WHAT REMAINS UNCERTAIN", 0.5, 1.52, 6, 0.25,
       size=7, color=LIGHT, font="Calibri", bold=True)
    uncerts = analysis.get("uncertainties", [])[:5]
    for ui, u in enumerate(uncerts):
        uy = 1.85 + ui * 0.9
        rect(s6, 0.5, uy, 6.0, 0.82, fill=BGALT)
        rect(s6, 0.5, uy, 0.04, 0.82, fill=NAVY)
        tb(s6, u.get("topic",""), 0.65, uy+0.06, 5.7, 0.3,
           size=10, bold=True, color=BLACK, font="Georgia")
        tb(s6, u.get("needed","")[:70], 0.65, uy+0.38, 5.7, 0.36,
           size=8.5, color=MID, font="Calibri")

    # 시사점 (우)
    impl = analysis.get("implications", {})
    impl_list = [
        ("경쟁사", impl.get("competitors","")),
        ("파트너", impl.get("partners","")),
        ("인재 시장", impl.get("talent","")),
    ]
    tb(s6, "STRATEGIC IMPLICATIONS", 6.83, 1.52, 6, 0.25,
       size=7, color=LIGHT, font="Calibri", bold=True)
    for ii, (label, text) in enumerate(impl_list):
        iy = 1.85 + ii * 1.55
        rect(s6, 6.83, iy, 6.0, 1.4, fill=BGALT)
        rect(s6, 6.83, iy, 0.04, 1.4, fill=NAVY)
        tb(s6, label, 7.0, iy+0.1, 5.7, 0.35,
           size=11, bold=True, color=BLACK, font="Georgia")
        tb(s6, text[:130], 7.0, iy+0.48, 5.7, 0.82,
           size=9, color=MID, font="Calibri")

    footer(s6, company, date_str)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── UI ─────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-eyebrow">Strategic Intelligence Tool</div>
  <div class="hero-title">Job Description Analyst</div>
  <div class="hero-desc">기업명을 입력하면 채용 페이지를 자동으로 탐색하고<br>전략 인텔리전스 리포트를 PPTX로 만들어드립니다.</div>
</div>
""", unsafe_allow_html=True)

# ── 입력 폼: 중앙 컬럼으로 Google 검색창 너비 구현 ──────────────────────────────
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

        st.markdown('<span class="field-label-sm">Google AI Studio API Key</span>', unsafe_allow_html=True)
        api_key = st.text_input("api_key", placeholder="AIza...", type="password",
                                label_visibility="collapsed")
        st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

        run_btn = st.button("◎  분석 시작", disabled=not (api_key and company))

    st.markdown('</div>', unsafe_allow_html=True)


# ── 실행 ───────────────────────────────────────────────────────────────────────
if run_btn and api_key and company:
    genai.configure(api_key=api_key)

    log_placeholder = st.empty()
    log_lines = []

    def log(msg, style="active"):
        cls = {"ok":"log-ok","dim":"log-dim","active":"log-active"}.get(style,"log-active")
        log_lines.append(f'<div class="log-line {cls}">{msg}</div>')
        log_placeholder.markdown(
            f'<div class="log-box">{"".join(log_lines)}</div>',
            unsafe_allow_html=True
        )

    log(f"→ 분석 대상: {company}")
    log(f"→ 모델: Gemini 2.5 Flash-Lite (무료)")
    if filter_kw:
        log(f"→ 필터: {filter_kw}")
    log("─" * 36, "dim")

    try:
        careers_url = find_careers_url(genai, company, log)
        if not careers_url:
            st.error("채용 페이지를 찾지 못했습니다. 회사명을 확인해주세요.")
            st.stop()

        st.markdown(f'<div class="url-found-box">채용 페이지 → {careers_url}</div>',
                    unsafe_allow_html=True)

        jobs = collect_jobs(genai, company, careers_url, filter_kw, log)
        if not jobs:
            st.error("공고 목록을 수집하지 못했습니다.")
            st.stop()

        jds      = collect_jd_details(genai, jobs[:10], company, log)
        signals  = extract_signals(genai, jds, company, log)
        analysis = synthesize_analysis(genai, company, filter_kw, jds, signals, log)

        log("◎ PPTX 생성 중...", "active")
        date_str   = datetime.now().strftime("%Y.%m.%d")
        pptx_bytes = build_pptx(company, filter_kw, careers_url, jds, analysis, date_str)

        log("─" * 36, "dim")
        log("◎ 완료 — PPTX 준비됨", "ok")

        st.markdown('<div class="result-section">', unsafe_allow_html=True)
        st.markdown(
            f'<div class="result-header">'
            f'<span>REPORT READY — {company.upper()}</span>'
            f'<span>{datetime.now().strftime("%Y.%m.%d %H:%M")}</span>'
            f'</div>',
            unsafe_allow_html=True
        )

        fname = f"{company.lower().replace(' ','_')}_jd_analysis_{datetime.now().strftime('%Y%m%d')}.pptx"
        st.download_button(
            "↓  PPTX 리포트 다운로드",
            data=pptx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        if "executive_summary" in analysis:
            st.markdown(f"**경영진 요약**\n\n{analysis['executive_summary']}")
        if "themes" in analysis:
            st.markdown("**전략 테마**")
            for t in analysis["themes"][:5]:
                st.markdown(f"- **{t.get('name','')}** (신뢰도: {t.get('confidence','')}) — {t.get('interpretation','')[:80]}")

        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        err = str(e).lower()
        if "api_key" in err or "invalid" in err or "api key" in err:
            st.error("API 키가 올바르지 않습니다. aistudio.google.com에서 발급받은 키를 확인해주세요.")
        elif "quota" in err or "limit" in err or "429" in err or "resource_exhausted" in err:
            st.error("무료 사용량 한도에 도달했습니다. 내일 자정(태평양 시간) 이후 다시 시도해주세요.")
        else:
            st.error(f"오류 발생: {e}")

else:
    pass
