import streamlit as st
import anthropic
import json
import re
import asyncio
from datetime import datetime

st.set_page_config(
    page_title="Jobs Intelligence",
    page_icon="◎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap');

*, html, body { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: #0d0d0d; color: #d4d0c8; }
.block-container { padding: 0 !important; max-width: 100% !important; }

.hero {
    padding: 72px 80px 56px;
    border-bottom: 1px solid #1a1a1a;
}
.hero-eyebrow {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.25em;
    color: #404040;
    text-transform: uppercase;
    margin-bottom: 20px;
}
.hero-title {
    font-size: 52px;
    font-weight: 300;
    color: #d4d0c8;
    letter-spacing: -0.03em;
    line-height: 1.05;
    margin-bottom: 16px;
}
.hero-title em {
    font-style: normal;
    color: #b8ff57;
}
.hero-desc {
    font-size: 14px;
    color: #404040;
    font-weight: 300;
    max-width: 480px;
    line-height: 1.6;
}

.form-section {
    padding: 56px 80px;
    border-bottom: 1px solid #1a1a1a;
}
.field-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.15em;
    color: #404040;
    text-transform: uppercase;
    margin-bottom: 8px;
    display: block;
}

.stTextInput > div > div > input {
    background: #111 !important;
    border: 1px solid #222 !important;
    border-radius: 2px !important;
    color: #d4d0c8 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 13px !important;
    padding: 14px 16px !important;
    transition: border-color 0.15s !important;
}
.stTextInput > div > div > input:focus {
    border-color: #b8ff57 !important;
    box-shadow: none !important;
}
.stTextInput > div > div > input::placeholder { color: #303030 !important; }

.stButton > button {
    background: #b8ff57 !important;
    color: #0d0d0d !important;
    border: none !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
    font-weight: 500 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    padding: 14px 32px !important;
    height: auto !important;
    transition: opacity 0.15s !important;
    width: 100% !important;
}
.stButton > button:hover { opacity: 0.85 !important; }
.stButton > button:disabled {
    background: #1a1a1a !important;
    color: #303030 !important;
}

.log-box {
    background: #111;
    border: 1px solid #1a1a1a;
    border-radius: 2px;
    padding: 20px 24px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #404040;
    line-height: 1.8;
    margin: 24px 80px;
    min-height: 120px;
    max-height: 280px;
    overflow-y: auto;
}
.log-line { margin: 2px 0; }
.log-ok { color: #b8ff57; }
.log-active { color: #d4d0c8; }
.log-dim { color: #303030; }

.result-section { padding: 56px 80px; }
.result-header {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.25em;
    color: #404040;
    text-transform: uppercase;
    margin-bottom: 32px;
    padding-bottom: 16px;
    border-bottom: 1px solid #1a1a1a;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.stMarkdown h1 { font-size: 28px !important; font-weight: 400 !important; color: #d4d0c8 !important; letter-spacing: -0.02em !important; margin-top: 40px !important; }
.stMarkdown h2 { font-size: 18px !important; font-weight: 500 !important; color: #b8ff57 !important; letter-spacing: 0.05em !important; text-transform: uppercase !important; margin-top: 36px !important; font-family: 'IBM Plex Mono', monospace !important; }
.stMarkdown h3 { font-size: 15px !important; font-weight: 400 !important; color: #d4d0c8 !important; margin-top: 28px !important; }
.stMarkdown p { color: #9a9690 !important; font-size: 14px !important; line-height: 1.75 !important; }
.stMarkdown li { color: #9a9690 !important; font-size: 14px !important; line-height: 1.75 !important; }
.stMarkdown table { width: 100% !important; border-collapse: collapse !important; }
.stMarkdown th { background: #111 !important; color: #d4d0c8 !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 10px !important; letter-spacing: 0.1em !important; text-transform: uppercase !important; padding: 12px 16px !important; border: 1px solid #1a1a1a !important; font-weight: 500 !important; }
.stMarkdown td { color: #9a9690 !important; font-size: 13px !important; padding: 10px 16px !important; border: 1px solid #1a1a1a !important; }
.stMarkdown tr:hover td { background: #111 !important; }
.stMarkdown code { background: #111 !important; color: #b8ff57 !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 12px !important; padding: 2px 6px !important; border-radius: 2px !important; }
.stMarkdown blockquote { border-left: 2px solid #b8ff57 !important; padding-left: 16px !important; margin: 16px 0 !important; }
.stMarkdown blockquote p { color: #404040 !important; font-size: 12px !important; font-family: 'IBM Plex Mono', monospace !important; }

.stDownloadButton > button {
    background: transparent !important;
    border: 1px solid #222 !important;
    color: #404040 !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 10px !important;
    letter-spacing: 0.1em !important;
    padding: 10px 20px !important;
    height: auto !important;
}
.stDownloadButton > button:hover { border-color: #b8ff57 !important; color: #b8ff57 !important; }

div[data-testid="stSidebarNav"] { display: none; }
.stAlert { border-radius: 2px !important; }
footer { display: none !important; }
#MainMenu { display: none !important; }
header { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── 에이전트 로직 ──────────────────────────────────────────────────────────────
WEB_SEARCH_TOOL = {"type": "web_search_20250305", "name": "web_search"}
SUBAGENT_MODEL     = "claude-haiku-4-5-20251001"   # 저렴한 모델로 JD 수집
ORCHESTRATOR_MODEL = "claude-sonnet-4-5"            # Opus → Sonnet으로 변경


def run_agent(client, system, prompt, model=None, max_tokens=2000, use_search=True):
    model = model or SUBAGENT_MODEL
    tools = [WEB_SEARCH_TOOL] if use_search else []
    messages = [{"role": "user", "content": prompt}]
    kwargs = dict(model=model, max_tokens=max_tokens, messages=messages)
    if system:
        kwargs["system"] = system
    if tools:
        kwargs["tools"] = tools

    response = client.messages.create(**kwargs)

    iterations = 0
    while response.stop_reason == "tool_use" and iterations < 5:
        iterations += 1
        tool_results = [
            {"type": "tool_result", "tool_use_id": b.id, "content": "검색 완료"}
            for b in response.content if b.type == "tool_use"
        ]
        messages.append({"role": "assistant", "content": response.content})
        messages.append({"role": "user", "content": tool_results})
        kwargs["messages"] = messages
        response = client.messages.create(**kwargs)

    return " ".join(b.text for b in response.content if hasattr(b, "text"))


def collect_jobs(client, company, url, filter_keyword, log):
    log("◎ 공고 목록 수집 중...")
    prompt = f"""다음 채용 페이지에서 공고 목록을 수집해줘.

회사: {company}
URL: {url}
필터: {filter_keyword if filter_keyword else '없음 (전체)'}

웹 검색으로 이 URL의 채용 공고들을 찾아서 다음 JSON 배열로만 반환해.
필터가 있으면 해당 키워드와 관련된 공고만 포함해.
다른 텍스트 없이 JSON만:

[{{"title":"직함","team":"팀/부서","location":"위치","date":"날짜","role_id":"ID (있으면)"}}]

최대 10건만 수집해. 가장 최근 공고 위주로."""

    result = run_agent(client, None, prompt)
    try:
        match = re.search(r'\[[\s\S]*\]', result)
        if match:
            jobs = json.loads(match.group())
            log(f"✓ {len(jobs)}건 수집 완료", "ok")
            return jobs
    except:
        pass

    log("→ 폴백: 추가 검색으로 보완 중...")
    prompt2 = f""""{company}" 채용 공고를 검색해줘.
{f'"{filter_keyword}" 관련 직무만.' if filter_keyword else ''}
site:{url.split("/")[2] if "://" in url else url.split("/")[0]}

직함, 팀, 위치, 날짜를 JSON 배열로만 반환:
[{{"title":"...","team":"...","location":"...","date":"..."}}]"""

    result2 = run_agent(client, None, prompt2)
    try:
        match = re.search(r'\[[\s\S]*\]', result2)
        if match:
            jobs = json.loads(match.group())
            log(f"✓ 폴백 {len(jobs)}건 수집", "ok")
            return jobs
    except:
        pass
    log("! 공고 목록 수집 실패", "dim")
    return []


def collect_jd_details(client, jobs, company, log):
    log(f"◎ {len(jobs)}건 JD 상세 수집 중...")
    jds = []
    for i, job in enumerate(jobs):
        title = job.get("title", "unknown")
        log(f"  [{i+1}/{len(jobs)}] {title}")
        prompt = f""""{company}" 회사의 다음 채용 공고 상세 내용을 웹 검색으로 찾아줘.

직함: {title}
팀: {job.get('team','')}
위치: {job.get('location','')}

다음 JSON 형식으로만 반환해:
{{
  "title": "{title}",
  "team": "{job.get('team','')}",
  "location": "{job.get('location','')}",
  "summary": "팀·역할 요약 2-3문장",
  "responsibilities": ["주요 책임 1", "주요 책임 2"],
  "min_qualifications": ["최소 자격 1", "최소 자격 2"],
  "preferred_qualifications": ["우대 자격 1"],
  "key_signals": ["전략 키워드1", "전략 키워드2"]
}}"""
        result = run_agent(client, None, prompt)
        try:
            match = re.search(r'\{[\s\S]*\}', result)
            if match:
                jd = json.loads(match.group())
                jds.append(jd)
                continue
        except:
            pass
        jds.append({**job, "summary": "", "responsibilities": [],
                    "min_qualifications": [], "preferred_qualifications": [],
                    "key_signals": []})
    log(f"✓ JD {len(jds)}건 수집 완료", "ok")
    return jds


def extract_signals(client, jds, company, log):
    log("◎ 전략 신호 분석 중...")
    batch_size = 12
    all_signals = []

    for i in range(0, len(jds), batch_size):
        batch = jds[i:i+batch_size]
        batch_text = json.dumps(batch, ensure_ascii=False)[:10000]
        prompt = f"""{company} 채용 공고 데이터에서 전략 신호를 추출해줘.

데이터:
{batch_text}

다음 6개 카테고리별로 JSON으로 반환해:
{{
  "growth": [{{"signal":"신호","evidence":"JD 근거 문구","roles":["역할"],"confidence":"high/medium/low"}}],
  "customer_experience": [...],
  "operations": [...],
  "technology": [...],
  "channel": [...],
  "geographic": [...],
  "patterns": [{{"pattern":"반복 패턴","count":N,"roles":["역할들"]}}],
  "team_names": ["팀명들"],
  "tech_stack": ["기술/시스템들"]
}}"""

        result = run_agent(client, None, prompt, use_search=False)
        try:
            match = re.search(r'\{[\s\S]*\}', result)
            if match:
                all_signals.append(json.loads(match.group()))
        except:
            pass

    log("✓ 신호 추출 완료", "ok")
    return all_signals


def synthesize_briefing(client, company, url, filter_keyword, jds, signals, log):
    log("◎ 경영진 브리핑 합성 중... (1-2분 소요)")

    context = {
        "company": company,
        "source_url": url,
        "filter": filter_keyword,
        "analysis_date": datetime.now().strftime("%Y-%m-%d"),
        "total_postings": len(jds),
        "signals": signals[:3],
        "sample_roles": [
            {"title": j.get("title"), "team": j.get("team"),
             "location": j.get("location"), "key_signals": j.get("key_signals", [])}
            for j in jds[:25]
        ]
    }

    prompt = f"""다음 채용 데이터를 바탕으로 전문적인 한국어 전략 인텔리전스 브리핑을 작성해줘.

분석 데이터:
{json.dumps(context, ensure_ascii=False)[:12000]}

다음 형식의 마크다운으로 작성해:

# {company} 채용 신호 전략 분석
**기준일: {datetime.now().strftime('%Y.%m.%d')} | 분석 공고: {len(jds)}건 | 필터: {filter_keyword or '전체'}**

---

## A. 경영진 요약

[핵심 판단 4-5개. 각각 🔴고신뢰 🟡중신뢰 🟢저신뢰 이모지 사용]

---

## B. 전략 테마 Top 5

[각 테마마다:]
### 테마 N — [테마명]
**신뢰도: 高/中/低 | 근거 공고: N건**

**핵심 근거 문구:**
- "[JD 원문 인용]" — 역할명

**전략적 해석:** [2-3문장]

**대안 해석:** [1문장]

---

## C. 증거 맵

| 직함 | 팀 | 위치 | 전략 테마 | 신호 강도 |
|---|---|---|---|---|
[공고들]

---

## D. {company}이 실제로 투자하는 것

[4-5가지 투자 영역, 각각 JD 근거 포함]

---

## E. 여전히 불확실한 것

| 불확실 사항 | 현재 신호 | 추가 필요 정보 |
|---|---|---|

---

## F. 시사점

**경쟁사에게:** [2-3줄]
**파트너에게:** [2-3줄]
**인재 시장에게:** [2-3줄]

---

> **데이터 한계:** [수집 방법 및 한계 명시]

규칙:
- 팩트와 추론 반드시 구분
- 단일 공고 기반 과잉 해석 금지
- 모든 전략 추론에 대안 해석 필수
- JD 원문 구체적 인용 (15단어 미만)"""

    result = run_agent(client, None, prompt, model=ORCHESTRATOR_MODEL, max_tokens=4000, use_search=False)
    log("✓ 브리핑 완성!", "ok")
    return result


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-eyebrow">Strategic Intelligence Tool</div>
  <div class="hero-title">Jobs<br><em>Intelligence</em></div>
  <div class="hero-desc">기업의 채용 공고를 분석해 전략적 방향성을 추론합니다.<br>회사명과 채용 URL만 입력하면 됩니다.</div>
</div>
""", unsafe_allow_html=True)

# ── 입력 폼 ───────────────────────────────────────────────────────────────────
with st.container():
    st.markdown('<div class="form-section">', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown('<span class="field-label">API Key</span>', unsafe_allow_html=True)
        api_key = st.text_input("api_key", placeholder="sk-ant-...", type="password",
                                label_visibility="collapsed")

    with col2:
        st.markdown('<span class="field-label">Company Name</span>', unsafe_allow_html=True)
        company = st.text_input("company", placeholder="Apple / Samsung / Google ...",
                                label_visibility="collapsed")

    st.markdown('<span class="field-label">Jobs URL</span>', unsafe_allow_html=True)
    url = st.text_input("url",
                        placeholder="https://jobs.apple.com/en-us/search?product=apple-online-store-AWB",
                        label_visibility="collapsed")

    st.markdown('<span class="field-label">Filter (선택)</span>', unsafe_allow_html=True)
    filter_kw = st.text_input("filter",
                              placeholder="특정 직무나 부서 입력 (예: Machine Learning, Online Store, Marketing ...)",
                              label_visibility="collapsed")

    col_btn, col_empty = st.columns([1, 3])
    with col_btn:
        run_btn = st.button("◎  분석 시작", disabled=not (api_key and company and url))

    st.markdown('</div>', unsafe_allow_html=True)


# ── 실행 ──────────────────────────────────────────────────────────────────────
if run_btn and api_key and company and url:
    client = anthropic.Anthropic(api_key=api_key)

    log_placeholder = st.empty()
    log_lines = []

    def log(msg, style="active"):
        cls = {"ok": "log-ok", "dim": "log-dim", "active": "log-active"}.get(style, "log-active")
        log_lines.append(f'<div class="log-line {cls}">{msg}</div>')
        log_placeholder.markdown(
            f'<div class="log-box">{"".join(log_lines)}</div>',
            unsafe_allow_html=True
        )

    log(f"→ 분석 대상: {company}")
    log(f"→ URL: {url}")
    if filter_kw:
        log(f"→ 필터: {filter_kw}")
    log("─" * 40, "dim")

    try:
        jobs  = collect_jobs(client, company, url, filter_kw, log)
        if not jobs:
            st.error("공고 목록을 수집하지 못했습니다. URL과 회사명을 확인해주세요.")
            st.stop()

        jds     = collect_jd_details(client, jobs[:10], company, log)
        signals = extract_signals(client, jds, company, log)
        briefing = synthesize_briefing(client, company, url, filter_kw, jds, signals, log)

        log("─" * 40, "dim")
        log("◎ 완료", "ok")

        # ── 결과 출력 ──────────────────────────────────────────────────────────
        st.markdown('<div class="result-section">', unsafe_allow_html=True)

        col_title, col_dl = st.columns([4, 1])
        with col_title:
            st.markdown(
                f'<div class="result-header"><span>ANALYSIS RESULT — {company.upper()}</span><span>{datetime.now().strftime("%Y.%m.%d %H:%M")}</span></div>',
                unsafe_allow_html=True
            )
        with col_dl:
            filename = f"{company.lower().replace(' ','_')}_jobs_briefing_{datetime.now().strftime('%Y%m%d')}.md"
            st.download_button(
                "↓  MD 다운로드",
                data=briefing,
                file_name=filename,
                mime="text/markdown"
            )

        st.markdown(briefing)
        st.markdown('</div>', unsafe_allow_html=True)

    except anthropic.AuthenticationError:
        st.error("API 키가 올바르지 않습니다. 확인 후 다시 시도해주세요.")
    except anthropic.BadRequestError as e:
        if "credit" in str(e).lower():
            st.error("크레딧이 부족합니다. console.anthropic.com에서 충전해주세요.")
        else:
            st.error(f"오류: {e}")
    except Exception as e:
        st.error(f"오류 발생: {e}")

else:
    # 초기 상태 안내
    if not run_btn:
        st.markdown("""
<div style="padding: 40px 80px; color: #303030; font-family: 'IBM Plex Mono', monospace; font-size: 11px; line-height: 2;">
입력 예시<br><br>
Company Name  →  Apple<br>
Jobs URL      →  https://jobs.apple.com/en-us/search?product=apple-online-store-AWB<br>
Filter        →  Machine Learning  (선택)<br><br>
분석 결과는 화면에 바로 표시되고 .md 파일로 다운로드됩니다.
</div>
""", unsafe_allow_html=True)
