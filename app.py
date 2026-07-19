
# app.py
# Starcom Executive Platform (SEP) V4 - Executive Intelligence Edition
# Clean IA: executive dashboard with collapsible guidance and summaries.
# Focus: MCR, Revenue, Productivity, FJA, Capacity, Strategic Decision, Benchmark Framework, CEO Closing.

import os
import json
import html
import requests
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="SEP V4 Executive Intelligence",
    page_icon="📊",
    layout="wide",
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = os.path.join(APP_DIR, "data", "SEP_V4_Final_Project_Productivity_Template4.xlsx")

REQUIRED_SHEETS = [
    "PARAMETERS", "TARGETS", "FJA_MAPPING", "DIM_DEPARTMENT",
    "REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR"
]
OPTIONAL_SHEETS = [
    "DIM_EMPLOYEE", "REVENUE_BY_DEPT", "HC_MOVEMENT", "PROJECT_MARGIN",
    "CAPACITY_INDICATOR", "CALENDAR", "MONTHLY_KPI", "PROJECT_PRODUCTIVITY", "BENCHMARK_FRAMEWORK",
    "SEP_PRESENTATION_FLOW", "APPENDIX_KPI_DEFINITIONS"
]

TONE_COLOR = {
    "success": "#22c55e",
    "info": "#3b82f6",
    "warning": "#f59e0b",
    "error": "#ef4444",
    "neutral": "#64748b",
}


# =========================
# AI PRESENTER
# =========================
def _get_openai_api_key() -> str:
    """Read API key from Streamlit secrets first, then environment variable."""
    try:
        return str(st.secrets.get("OPENAI_API_KEY", "")).strip()
    except Exception:
        return os.getenv("OPENAI_API_KEY", "").strip()


def _build_ai_context(latest, period, mode, bs, bp, mq, mh, cr, currency):
    return {
        "period": period,
        "analysis_mode": mode,
        "currency": currency,
        "company_health": {"score": bs.get("score"), "status": bs.get("label")},
        "business_posture": bp,
        "margin_quality": mq,
        "mcr_health": mh,
        "capacity_risk": cr,
        "metrics": {
            "revenue": _safe_float(latest.get("Total_Revenue")),
            "revenue_yoy": _safe_float(latest.get("Revenue_YoY")),
            "headcount": _safe_float(latest.get("Total_Headcount")),
            "headcount_yoy": _safe_float(latest.get("Headcount_YoY")),
            "rpe": _safe_float(latest.get("RPE")),
            "rpe_yoy": _safe_float(latest.get("RPE_YoY")),
            "mcr": _safe_float(latest.get("MCR_Pct")),
            "gross_margin": _safe_float(latest.get("Gross_Margin_Pct")),
            "manpower_cost": _safe_float(latest.get("Total_Manpower_Cost")),
            "cost_per_hc": _safe_float(latest.get("Cost_per_HC")),
        },
    }


def _route_presenter_question(question: str) -> dict:
    q = (question or "").lower()
    routes = [
        (["capacity", "utilization", "utilisasi", "overtime", "lembur", "backlog", "sla"], "4. Capacity & Delivery Risk", "capacity", ["Capacity Score", "Avg Utilization", "Overtime", "Backlog", "SLA"]),
        (["mcr", "manpower cost", "biaya tenaga", "people cost", "payroll"], "1. Executive Summary", "mcr", ["MCR", "Manpower Cost", "Revenue", "MCR Health"]),
        (["revenue", "pendapatan", "rpe", "produktivitas"], "2. Revenue & Productivity", "revenue", ["Revenue", "RPE", "Revenue Growth", "Revenue vs Manpower Cost"]),
        (["gross margin", "margin", "cogs", "laba kotor"], "2. Revenue & Productivity", "margin", ["Gross Margin", "COGS", "Revenue", "Margin Quality"]),
        (["headcount", "jumlah karyawan", "hc", "hiring", "rekrut"], "3. Workforce Performance", "workforce", ["Headcount", "Cost per HC", "FJA Cost Mix", "Hiring Decision"]),
        (["health", "kesehatan perusahaan", "company health", "kondisi perusahaan"], "5. Strategic Decision Center", "health", ["Company Health", "Business Posture", "Management Attention", "Top Decisions"]),
        (["benchmark", "target", "metodologi", "framework"], "6. Executive Benchmark & Methodology", "benchmark", ["KPI Target Framework", "Current KPI vs Framework"]),
        (["kesimpulan", "closing", "rekomendasi", "keputusan"], "7. CEO Closing Report", "closing", ["Company Health", "Top Decisions", "90-Day Action Plan"]),
    ]
    for keys, board, route, targets in routes:
        if any(k in q for k in keys):
            return {"board": board, "route": route, "targets": targets}
    return {"board": "1. Executive Summary", "route": "summary", "targets": ["Company Health", "Revenue", "MCR", "Gross Margin", "Capacity Risk"]}


def _local_presenter_answer(question: str, ctx: dict, route: dict) -> str:
    m = ctx["metrics"]
    cr = ctx["capacity_risk"]
    period_label = ctx.get("period", {}).get("Period_Label", "periode terpilih")
    currency = ctx.get("currency", "IDR")
    focus = route.get("route")
    if focus == "mcr":
        return (
            f"Pada {period_label}, MCR tercatat {_pct(m.get('mcr'))} dengan manpower cost {_money(m.get('manpower_cost'), currency)} "
            f"dan revenue {_money(m.get('revenue'), currency)}. Status MCR adalah {ctx['mcr_health'].get('state')}. "
            "MCR harus dibaca bersama gross margin, RPE, dan capacity risk karena MCR yang rendah belum otomatis berarti kapasitas organisasi sehat."
        )
    if focus == "capacity":
        d = cr.get("details", {})
        return (
            f"Capacity Risk berada pada level {cr.get('state')} dengan skor {cr.get('score')}/100. "
            f"Utilization maksimum {_pct(d.get('max_utilization'))}, overtime {_num(d.get('total_overtime'))} jam, "
            f"backlog {_num(d.get('total_backlog'))}, dan SLA breach {_num(d.get('total_sla_breach'))}. "
            f"Rekomendasi sistem adalah {cr.get('recommendation')}."
        )
    if focus == "revenue":
        return (
            f"Revenue pada {period_label} adalah {_money(m.get('revenue'), currency)} dengan pertumbuhan {_pct(m.get('revenue_yoy'))}. "
            f"RPE tercatat {_money(m.get('rpe'), currency)} dengan perubahan {_pct(m.get('rpe_yoy'))}. "
            "Kualitas pertumbuhan dinilai baik apabila revenue dan RPE meningkat lebih cepat daripada headcount serta manpower cost."
        )
    if focus == "margin":
        return (
            f"Gross Margin tercatat {_pct(m.get('gross_margin'))} dengan status {ctx['margin_quality'].get('state')}. "
            "Indikator ini menunjukkan kualitas revenue setelah COGS langsung. Bila margin turun, prioritas pemeriksaan adalah pricing, scope creep, subcontractor, rework, dan direct project cost."
        )
    if focus == "workforce":
        return (
            f"Headcount rata-rata tercatat {_num(m.get('headcount'))} dengan perubahan {_pct(m.get('headcount_yoy'))}. "
            f"Cost per HC adalah {_money(m.get('cost_per_hc'), currency)} dan RPE {_money(m.get('rpe'), currency)}. "
            "Penambahan HC sebaiknya hanya dilakukan pada bottleneck revenue atau delivery yang didukung utilization, backlog, SLA, dan pipeline revenue."
        )
    if focus in ["health", "closing"]:
        return (
            f"Company Health berada pada skor {ctx['company_health'].get('score')}/100 dengan status {ctx['company_health'].get('status')}. "
            f"Business posture adalah {ctx['business_posture'].get('state')}, MCR {ctx['mcr_health'].get('state')}, "
            f"margin {ctx['margin_quality'].get('state')}, dan capacity risk {cr.get('state')}. "
            "Keputusan utama harus menyeimbangkan perlindungan margin, disiplin manpower cost, dan kapasitas delivery."
        )
    return (
        f"Pada {period_label}, revenue {_money(m.get('revenue'), currency)}, MCR {_pct(m.get('mcr'))}, "
        f"Gross Margin {_pct(m.get('gross_margin'))}, RPE {_money(m.get('rpe'), currency)}, dan Capacity Risk {cr.get('state')}. "
        f"Company Health berada pada {ctx['company_health'].get('score')}/100 dengan status {ctx['company_health'].get('status')}."
    )


def _openai_presenter_answer(question: str, ctx: dict, route: dict, api_key: str) -> str:
    if not api_key:
        return _local_presenter_answer(question, ctx, route)
    prompt = (
        "Anda adalah AI Executive Presenter untuk PT Starcom Solusindo. Jawab dalam Bahasa Indonesia, maksimal 120 kata, "
        "langsung, objektif, dan hanya berdasarkan JSON data. Jangan mengarang angka. Jawab pertanyaan bisnis secara langsung dan berikan satu rekomendasi. "
        f"Pertanyaan: {question}\nBoard yang relevan: {route['board']}\nData: {json.dumps(ctx, ensure_ascii=False, default=str)}"
    )
    try:
        r = requests.post(
            "https://api.openai.com/v1/responses",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": "gpt-5-mini", "input": prompt},
            timeout=60,
        )
        r.raise_for_status()
        payload = r.json()
        if payload.get("output_text"):
            return payload["output_text"].strip()
        texts = []
        for out in payload.get("output", []):
            for c in out.get("content", []):
                if c.get("type") in ("output_text", "text") and c.get("text"):
                    texts.append(c["text"])
        return " ".join(texts).strip() or _local_presenter_answer(question, ctx, route)
    except Exception:
        return _local_presenter_answer(question, ctx, route)


def _transcribe_audio(audio_bytes: bytes, api_key: str) -> tuple[str, str]:
    if not api_key:
        return "", "OPENAI_API_KEY belum tersedia. Gunakan pertanyaan teks atau tambahkan API key."
    try:
        files = {"file": ("question.wav", audio_bytes, "audio/wav")}
        data = {"model": "gpt-4o-mini-transcribe", "language": "id"}
        r = requests.post(
            "https://api.openai.com/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {api_key}"},
            files=files,
            data=data,
            timeout=90,
        )
        r.raise_for_status()
        return str(r.json().get("text", "")).strip(), ""
    except Exception as e:
        return "", f"Transkripsi gagal: {type(e).__name__}."


def _synthesize_speech(text: str, api_key: str) -> tuple[bytes | None, str]:
    if not api_key:
        return None, "OPENAI_API_KEY belum tersedia; jawaban suara belum dapat dibuat."
    try:
        r = requests.post(
            "https://api.openai.com/v1/audio/speech",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": "gpt-4o-mini-tts",
                "voice": "alloy",
                "input": text,
                "format": "mp3",
            },
            timeout=90,
        )
        r.raise_for_status()
        return r.content, ""
    except Exception as e:
        return None, f"Audio jawaban gagal dibuat: {type(e).__name__}."


def _render_presenter_visual(route: dict, ctx: dict):
    m = ctx["metrics"]
    focus = route.get("route")
    if focus == "capacity":
        cr = ctx["capacity_risk"]
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=cr.get("score", 0),
            number={"suffix": "/100"},
            title={"text": f"Capacity Risk — {cr.get('state')}"},
            gauge={"axis": {"range": [0, 100]}, "bar": {"color": TONE_COLOR.get(cr.get("tone"), "#3b82f6")}},
        ))
        fig.update_layout(height=260, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True, key="ai_capacity_visual")
    elif focus == "mcr":
        df = pd.DataFrame({"Metric": ["MCR", "Gross Margin"], "Value": [m.get("mcr"), m.get("gross_margin")]})
        fig = px.bar(df, x="Metric", y="Value", text=df["Value"].map(_pct), title="MCR vs Gross Margin")
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(height=260, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True, key="ai_mcr_visual")
    else:
        df = pd.DataFrame({
            "Metric": ["Revenue", "Manpower Cost", "RPE"],
            "Value": [m.get("revenue"), m.get("manpower_cost"), m.get("rpe")],
        })
        fig = px.bar(df, x="Metric", y="Value", text_auto=".3s", title="Executive Metric Focus")
        fig.update_layout(height=260, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True, key="ai_summary_visual")


def inject_ai_presenter_css():
    st.markdown("""
    <style>
    /* Only the keyed presenter container is floating.
       Do not use :has() on stVerticalBlock because it also captures parent containers. */
    .st-key-ai_presenter_float {
        position: fixed !important;
        top: 50px !important;
        right: 86px !important;
        z-index: 999999 !important;
        width: auto !important;
        min-width: 0 !important;
        background: transparent !important;
        pointer-events: auto !important;
    }
    .st-key-ai_presenter_float button {
        border-radius: 999px !important;
        border: 1px solid #F59E0B !important;
        box-shadow: 0 8px 24px rgba(15,23,42,.16) !important;
        font-weight: 800 !important;
        color: #172033 !important;
        background: #FFFFFF !important;
        white-space: nowrap !important;
    }
    .sep-ai-focus {
        border: 2px solid #F59E0B;
        border-radius: 14px;
        padding: 12px 14px;
        background: #FFF7ED;
        box-shadow: 0 0 0 6px rgba(245,158,11,.12);
        animation: sepPulse 1.8s ease-in-out infinite;
    }
    @keyframes sepPulse {
        0%,100% { box-shadow: 0 0 0 4px rgba(245,158,11,.10); }
        50% { box-shadow: 0 0 0 10px rgba(245,158,11,.20); }
    }
    @media (max-width: 768px) {
        .st-key-ai_presenter_float {
            top: 46px !important;
            right: 54px !important;
        }
        .st-key-ai_presenter_float button {
            font-size: 12px !important;
            padding-left: 10px !important;
            padding-right: 10px !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)


def render_ai_presenter(ctx: dict):
    inject_ai_presenter_css()
    if "ai_presenter_question" not in st.session_state:
        st.session_state.ai_presenter_question = ""
    if "ai_presenter_answer" not in st.session_state:
        st.session_state.ai_presenter_answer = ""
    if "ai_presenter_route" not in st.session_state:
        st.session_state.ai_presenter_route = _route_presenter_question("")
    if "ai_presenter_audio" not in st.session_state:
        st.session_state.ai_presenter_audio = None

    api_key = _get_openai_api_key()
    # A keyed container creates a dedicated CSS class (st-key-ai_presenter_float).
    # This prevents the entire Streamlit page from becoming position:fixed.
    host = st.container(key="ai_presenter_float")
    with host:
        with st.popover("🎙 AI Presenter", help="Tanya dashboard dengan suara atau teks", use_container_width=False):
            st.caption("Jawaban hanya memakai data aktif SEP V4.")
            audio = st.audio_input("Ajukan pertanyaan dengan suara", sample_rate=16000, key="ai_voice_input")
            typed = st.text_input(
                "Atau ketik pertanyaan",
                value=st.session_state.ai_presenter_question,
                placeholder="Contoh: Mengapa capacity risk medium?",
                key="ai_text_input",
            )
            c1, c2 = st.columns(2)
            ask = c1.button("Analisis", type="primary", use_container_width=True, key="ai_ask")
            clear = c2.button("Reset", use_container_width=True, key="ai_reset")

            if clear:
                st.session_state.ai_presenter_question = ""
                st.session_state.ai_presenter_answer = ""
                st.session_state.ai_presenter_audio = None
                st.rerun()

            if ask:
                question = typed.strip()
                if audio is not None:
                    transcript, err = _transcribe_audio(audio.getvalue(), api_key)
                    if err:
                        st.warning(err)
                    if transcript:
                        question = transcript
                if not question:
                    st.warning("Masukkan pertanyaan melalui suara atau teks.")
                else:
                    route = _route_presenter_question(question)
                    with st.spinner("Menganalisis data SEP V4..."):
                        answer = _openai_presenter_answer(question, ctx, route, api_key)
                        audio_bytes, audio_err = _synthesize_speech(answer, api_key)
                    st.session_state.ai_presenter_question = question
                    st.session_state.ai_presenter_answer = answer
                    st.session_state.ai_presenter_route = route
                    st.session_state.ai_presenter_audio = audio_bytes
                    if audio_err:
                        st.caption(audio_err)

            if st.session_state.ai_presenter_answer:
                route = st.session_state.ai_presenter_route
                st.markdown(
                    f"<div class='sep-ai-focus'><b>Fokus visual:</b> {html.escape(route['board'])}<br>"
                    f"<small>{html.escape(', '.join(route['targets']))}</small></div>",
                    unsafe_allow_html=True,
                )
                st.write(st.session_state.ai_presenter_answer)
                _render_presenter_visual(route, ctx)
                if st.session_state.ai_presenter_audio:
                    st.audio(st.session_state.ai_presenter_audio, format="audio/mp3", autoplay=True)
            elif not api_key:
                st.info("Mode lokal aktif. Pertanyaan teks tetap dapat dijawab. Tambahkan OPENAI_API_KEY untuk transkripsi dan suara.")


# =========================
# FORMATTERS
# =========================
def _safe_str(x) -> str:
    if x is None or pd.isna(x):
        return ""
    return str(x).strip()


def _safe_float(x):
    try:
        if x is None or pd.isna(x):
            return np.nan
        return float(x)
    except Exception:
        return np.nan


def _money(x, currency="IDR") -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        x = float(x)
        if abs(x) >= 1_000_000_000:
            return f"{currency} {x/1_000_000_000:,.1f}B".replace(",", ".")
        if abs(x) >= 1_000_000:
            return f"{currency} {x/1_000_000:,.1f}M".replace(",", ".")
        return f"{currency} {x:,.0f}".replace(",", ".")
    except Exception:
        return "-"


def _num(x) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):,.1f}".replace(",", ".")
    except Exception:
        return "-"


def _pct(x) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "-"


def _pp(x) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x)*100:+.1f} pp"
    except Exception:
        return "-"


def _find_logo_path() -> str:
    for p in [
        os.path.join(APP_DIR, "assets", "logo_company.png"),
        os.path.join(APP_DIR, "assets", "logo_company.jpg"),
        os.path.join(APP_DIR, "assets", "starcom_logo.png"),
        os.path.join(APP_DIR, "logo_company.png"),
        os.path.join(APP_DIR, "company_logo.png"),
    ]:
        if os.path.exists(p):
            return p
    return ""



def inject_header_feature_css():
    st.markdown("""
    <style>
    :root { color-scheme: light; }
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background: #FFFFFF !important;
        color: #172033 !important;
    }
    [data-testid="stHeader"] {
        background: rgba(255,255,255,.96) !important;
    }
    .main .block-container {
        max-width: 1520px;
        padding-top: 1.2rem;
    }
    .sep-feature-band {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 14px;
        margin: 10px 0 20px 0;
        padding: 16px 18px;
        border: 1px solid #E5EAF2;
        border-radius: 16px;
        background: #F8FAFC;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.035);
    }
    .sep-feature-item {
        display: grid;
        grid-template-columns: 38px 1fr;
        gap: 10px;
        align-items: start;
        min-height: 64px;
    }
    .sep-feature-icon {
        width: 34px;
        height: 34px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        background: #EEF6FF;
        border: 1px solid #D9EAFE;
    }
    .sep-feature-title {
        font-size: 12.5px;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.25;
        margin-bottom: 4px;
    }
    .sep-feature-text {
        font-size: 11.5px;
        color: #64748B;
        line-height: 1.45;
    }
    /* Make the logo stay in the top-right and slightly smaller without touching the dashboard content */
    div[data-testid="stHorizontalBlock"] img {
        max-height: 82px;
        object-fit: contain;
    }

    .board-question-card {
        border-left: 5px solid #2563EB;
        border-radius: 14px;
        padding: 14px 18px;
        margin: 8px 0 18px 0;
        background: linear-gradient(90deg, #EFF6FF 0%, #F8FAFC 100%);
        box-shadow: 0 6px 18px rgba(15,23,42,.04);
    }
    .board-question-label { font-size: 11px; font-weight: 900; color: #2563EB; letter-spacing: .10em; }
    .board-question-text { font-size: 17px; font-weight: 850; color: #0F172A; margin-top: 5px; line-height: 1.4; }
    .board-decision-focus { font-size: 13px; color: #475569; margin-top: 7px; }
    .board-how-card, .board-summary-card {
        border: 1px solid #E2E8F0; border-radius: 16px; padding: 18px 20px; min-height: 270px;
        background: #FFFFFF; box-shadow: 0 8px 22px rgba(15,23,42,.045);
    }
    .board-how-card { background: linear-gradient(180deg,#F2F7FF 0%,#FBFDFF 100%); border-color: #D9EAFE; }
    .board-card-title { font-size: 17px; font-weight: 900; color: #0B63CE; margin-bottom: 12px; }
    .board-how-card li { margin-bottom: 9px; line-height: 1.5; color: #172033; }
    .board-summary-text { font-size: 15px; line-height: 1.65; color: #172033; }
    .board-implication { margin-top: 14px; padding-top: 12px; border-top: 1px solid rgba(100,116,139,.22); line-height: 1.55; color: #334155; }
    .board-summary-info { border-top: 5px solid #3B82F6; }
    .board-summary-success { border-top: 5px solid #22C55E; }
    .board-summary-warning { border-top: 5px solid #F59E0B; }
    .board-summary-error { border-top: 5px solid #EF4444; }
    </style>
    """, unsafe_allow_html=True)


def render_feature_band():
    st.markdown("""
    <div class="sep-feature-band">
        <div class="sep-feature-item">
            <div class="sep-feature-icon">🎯</div>
            <div>
                <div class="sep-feature-title">Executive Decision Support</div>
                <div class="sep-feature-text">Insight yang relevan untuk keputusan direksi dan management.</div>
            </div>
        </div>
        <div class="sep-feature-item">
            <div class="sep-feature-icon">📈</div>
            <div>
                <div class="sep-feature-title">Performance & Productivity</div>
                <div class="sep-feature-text">Monitoring revenue, MCR, margin, dan produktivitas karyawan.</div>
            </div>
        </div>
        <div class="sep-feature-item">
            <div class="sep-feature-icon">🛡️</div>
            <div>
                <div class="sep-feature-title">Risk & Capacity Control</div>
                <div class="sep-feature-text">Identifikasi risiko kapasitas, delivery, dan efisiensi organisasi.</div>
            </div>
        </div>
        <div class="sep-feature-item">
            <div class="sep-feature-icon">👥</div>
            <div>
                <div class="sep-feature-title">Benchmark & Best Practice</div>
                <div class="sep-feature-text">Framework benchmark berbasis Telekomunikasi & IT Integrator.</div>
            </div>
        </div>
        <div class="sep-feature-item">
            <div class="sep-feature-icon">📄</div>
            <div>
                <div class="sep-feature-title">Board Ready Report</div>
                <div class="sep-feature-text">Ringkasan eksekutif siap untuk board meeting dan stakeholder.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_header():
    c1, c2 = st.columns([6, 1.3])
    with c1:
        st.markdown("""
        <div style="padding:4px 0 6px 0;">
          <div style="font-size:36px;font-weight:900;line-height:1.1;">Starcom Executive Platform (SEP) V4</div>
          <div style="font-size:15px;color:#64748b;margin-top:6px;">Executive Intelligence Edition — MCR, revenue leverage, workforce productivity, capacity risk, benchmark framework, and board decision.</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        logo = _find_logo_path()
        if logo:
            st.image(logo, use_container_width=True)
        else:
            st.caption("Logo: assets/logo_company.png")


def badge(label, value, tone="info", sub="") -> str:
    color = TONE_COLOR.get(tone, "#3b82f6")
    return f"""
    <div style="border:1px solid rgba(148,163,184,.35); border-radius:16px; padding:15px 16px; background:rgba(15,23,42,.04); min-height:100px;">
      <div style="font-size:12px; color:#64748b; text-transform:uppercase; letter-spacing:.08em;">{label}</div>
      <div style="font-size:22px; font-weight:800; margin-top:8px;"><span style="display:inline-block;width:11px;height:11px;border-radius:999px;background:{color};margin-right:8px;"></span>{value}</div>
      <div style="font-size:12px;color:#64748b;margin-top:6px;line-height:1.4;">{sub}</div>
    </div>
    """


# =========================
# DATA
# =========================
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: Optional[bytes], default_path: str):
    warnings = []
    try:
        if file_bytes:
            xls = pd.ExcelFile(BytesIO(file_bytes))
        else:
            if not os.path.exists(default_path):
                return {}, [f"File default tidak ditemukan: {default_path}. Upload Excel terlebih dahulu."]
            xls = pd.ExcelFile(default_path)
    except Exception as e:
        return {}, [f"Excel tidak bisa dibaca: {type(e).__name__}: {e}"]

    sheets = {}
    for sh in REQUIRED_SHEETS + OPTIONAL_SHEETS:
        if sh in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sh)
                df.columns = [str(c).strip() for c in df.columns]
                sheets[sh] = df
            except Exception as e:
                warnings.append(f"Sheet {sh} gagal dibaca: {e}")
    missing = [s for s in REQUIRED_SHEETS if s not in sheets]
    if missing:
        warnings.append("Sheet wajib belum ada: " + ", ".join(missing))
    return sheets, warnings


def load_params(df: pd.DataFrame) -> Dict[str, str]:
    out = {}
    if df is None or df.empty or not {"Parameter", "Value"}.issubset(df.columns):
        return out
    for _, r in df.iterrows():
        k = _safe_str(r.get("Parameter"))
        if k:
            out[k] = _safe_str(r.get("Value"))
    return out


def load_targets(df: pd.DataFrame) -> Dict[str, float]:
    out = {}
    if df is None or df.empty or not {"KPI", "Target_Value"}.issubset(df.columns):
        return out
    for _, r in df.iterrows():
        k = _safe_str(r.get("KPI"))
        try:
            if k:
                out[k] = float(r.get("Target_Value"))
        except Exception:
            pass
    return out


def get_target(targets: Dict[str, float], key: str, default: float) -> float:
    return float(targets.get(key, default))


def build_calendar(calendar: Optional[pd.DataFrame], years: List[int]) -> pd.DataFrame:
    if calendar is None or calendar.empty or "Year" not in calendar.columns:
        return pd.DataFrame({
            "Year": years,
            "Months_Closed": [12] * len(years),
            "Period_Type": ["FY"] * len(years),
            "Period_Label": [f"FY {y}" for y in years],
            "Data_Confidence": ["High"] * len(years),
        })
    cal = calendar.copy()
    cal["Year"] = pd.to_numeric(cal["Year"], errors="coerce").astype("Int64")
    cal["Months_Closed"] = pd.to_numeric(cal.get("Months_Closed"), errors="coerce").fillna(12).clip(1, 12)
    if "Period_Type" not in cal.columns:
        cal["Period_Type"] = np.where(cal["Months_Closed"] < 12, "YTD", "FY")
    if "Period_Label" not in cal.columns:
        cal["Period_Label"] = cal.apply(lambda r: f"{r['Period_Type']} {int(r['Year'])}", axis=1)
    if "Data_Confidence" not in cal.columns:
        cal["Data_Confidence"] = np.where(cal["Months_Closed"] >= 12, "High", "Medium")
    cal = cal.dropna(subset=["Year"]).copy()
    cal["Year"] = cal["Year"].astype(int)
    return cal[["Year", "Months_Closed", "Period_Type", "Period_Label", "Data_Confidence"]]


def mode_maps(calendar: pd.DataFrame, mode: str):
    cal = calendar.copy()
    cal["Months_Closed"] = pd.to_numeric(cal["Months_Closed"], errors="coerce").fillna(12).clip(1, 12)
    ytd = cal[cal["Months_Closed"] < 12]
    comparable_months = int(ytd["Months_Closed"].max()) if not ytd.empty else 12
    if mode == "Actual":
        amount = {int(r["Year"]): 1.0 for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = "Actual: angka sesuai input Excel."
    elif mode == "YTD Comparable":
        amount = {int(r["Year"]): float(comparable_months / r["Months_Closed"]) for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = f"YTD Comparable: semua tahun dibaca pada basis {comparable_months} bulan. Jika MONTHLY_KPI terisi, dashboard memakai realisasi aktual bulan yang sama."
    elif mode == "Annual Projection":
        amount = {int(r["Year"]): float(12 / r["Months_Closed"]) for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = "Annual Projection: YTD disetahunkan memakai faktor 12 / Months_Closed."
    elif mode == "MCR Focus":
        rev = {int(r["Year"]): float(12 / r["Months_Closed"]) for _, r in cal.iterrows()}
        cost = {int(r["Year"]): 1.0 for _, r in cal.iterrows()}
        amount = cost.copy()
        note = "MCR Focus: revenue diproyeksikan full-year, sementara COGS dan manpower cost tetap actual/YTD untuk stress-test MCR dan margin."
    else:
        amount = {int(r["Year"]): 1.0 for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = "Fallback Actual."
    return amount, rev, cost, note, comparable_months


def get_period_info(calendar: pd.DataFrame, year: int):
    if calendar is None or calendar.empty:
        return {"Period_Label": f"FY {year}", "Months_Closed": 12, "Data_Confidence": "High"}
    x = calendar[calendar["Year"] == int(year)]
    if x.empty:
        return {"Period_Label": f"FY {year}", "Months_Closed": 12, "Data_Confidence": "High"}
    r = x.iloc[0]
    return {
        "Period_Label": str(r.get("Period_Label", f"FY {year}")),
        "Months_Closed": int(r.get("Months_Closed", 12)),
        "Data_Confidence": str(r.get("Data_Confidence", "High")),
    }


def prep_data(sheets: Dict[str, pd.DataFrame], mode: str):
    dim = sheets["DIM_DEPARTMENT"].copy()
    fja_map = sheets["FJA_MAPPING"].copy()
    rev = sheets["REVENUE_YR"].copy()
    hc = sheets["HEADCOUNT_YR"].copy()
    pay = sheets["PAYROLL_YR"].copy()

    raw_years = []
    for df in [rev, hc, pay]:
        if "Year" in df.columns:
            raw_years += pd.to_numeric(df["Year"], errors="coerce").dropna().astype(int).tolist()
    calendar = build_calendar(sheets.get("CALENDAR", pd.DataFrame()), sorted(set(raw_years)))
    amount_map, rev_map, cost_map, mode_note, comparable_months = mode_maps(calendar, mode)
    annual_rev_map = {int(r["Year"]): float(12 / r["Months_Closed"]) for _, r in calendar.iterrows()}

    # FJA mapping
    for c in ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "FJA_Category_Override"]:
        if c in dim.columns:
            dim[c] = dim[c].astype(str).str.strip()
    if "FJA_Category_Override" not in dim.columns:
        dim["FJA_Category_Override"] = ""
    if not fja_map.empty and {"Function_Group", "FJA_Category"}.issubset(fja_map.columns):
        fja_lookup = dict(zip(fja_map["Function_Group"].astype(str).str.strip(), fja_map["FJA_Category"].astype(str).str.strip()))
    else:
        fja_lookup = {"Sales": "Revenue Generator", "Operations/Project": "Revenue Enabler", "Engineering": "Revenue Enabler", "Support": "Support Function", "Management": "Governance / Management"}
    dim["FJA_Category"] = np.where(
        dim["FJA_Category_Override"].astype(str).str.strip().ne(""),
        dim["FJA_Category_Override"].astype(str).str.strip(),
        dim["Function_Group"].map(fja_lookup).fillna("Unmapped"),
    )
    dim_small = dim[["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "FJA_Category"]].copy()

    # Monthly actual override for YTD comparable
    monthly_used = False
    monthly_dept_revenue = pd.DataFrame()
    if mode == "YTD Comparable" and "MONTHLY_KPI" in sheets and not sheets["MONTHLY_KPI"].empty:
        mk = sheets["MONTHLY_KPI"].copy()
        if {"Year", "Month_No"}.issubset(mk.columns):
            mk["Year"] = pd.to_numeric(mk["Year"], errors="coerce").astype("Int64")
            mk["Month_No"] = pd.to_numeric(mk["Month_No"], errors="coerce").astype("Int64")
            mk = mk[(mk["Month_No"] >= 1) & (mk["Month_No"] <= comparable_months)].copy()
            for c in ["Revenue_Recognized", "COGS_Direct", "Avg_Headcount", "New_Hires", "Exits", "Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
                if c not in mk.columns:
                    mk[c] = 0.0
                mk[c] = pd.to_numeric(mk[c], errors="coerce").fillna(0.0)
            if "Dept_ID" not in mk.columns:
                mk["Dept_ID"] = ""
            mk["Dept_ID"] = mk["Dept_ID"].astype(str).str.strip()
            if "Business_Line" not in mk.columns:
                mk["Business_Line"] = "Total Company"

            rev_m = mk[(mk["Revenue_Recognized"] != 0) | (mk["COGS_Direct"] != 0)].copy()
            if not rev_m.empty:
                rev = rev_m.groupby(["Year", "Business_Line"], as_index=False).agg(
                    Revenue_Recognized=("Revenue_Recognized", "sum"),
                    COGS_Direct=("COGS_Direct", "sum"),
                )
                rev["Gross_Profit"] = rev["Revenue_Recognized"] - rev["COGS_Direct"]
                monthly_used = True
                drr = rev_m[rev_m["Dept_ID"].ne("") & ~rev_m["Dept_ID"].str.upper().isin(["COMPANY_TOTAL", "TOTAL", "TOTAL COMPANY"])].copy()
                if not drr.empty:
                    monthly_dept_revenue = drr.groupby(["Year", "Dept_ID"], as_index=False).agg(Dept_Revenue=("Revenue_Recognized", "sum"))

            dm = mk[mk["Dept_ID"].ne("") & ~mk["Dept_ID"].str.upper().isin(["COMPANY_TOTAL", "TOTAL", "TOTAL COMPANY"])].copy()
            if not dm.empty and dm["Avg_Headcount"].sum() > 0:
                hc = dm.groupby(["Year", "Dept_ID"], as_index=False).agg(
                    Avg_Headcount=("Avg_Headcount", "mean"),
                    Avg_FTE=("Avg_Headcount", "mean"),
                    New_Hires=("New_Hires", "sum"),
                    Exits=("Exits", "sum"),
                )
            if not dm.empty and dm["Total_Manpower_Cost"].sum() > 0:
                pay = dm.groupby(["Year", "Dept_ID"], as_index=False).agg(
                    Payroll_Gross=("Payroll_Gross", "sum"),
                    Overtime=("Overtime", "sum"),
                    Bonus=("Bonus", "sum"),
                    Benefits=("Benefits", "sum"),
                    Employer_Tax=("Employer_Tax", "sum"),
                    Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
                )

    if monthly_used:
        mode_note = f"YTD Comparable memakai MONTHLY_KPI actual Jan-{comparable_months}. KPI, Board Score, dan Closing dihitung dari realisasi YTD aktual."
    elif mode == "YTD Comparable":
        mode_note = f"YTD Comparable memakai proxy FY x {comparable_months}/12 karena MONTHLY_KPI belum aktif/terisi."

    # Revenue
    rev["Year"] = pd.to_numeric(rev["Year"], errors="coerce").astype("Int64")
    for c in ["Revenue_Recognized", "COGS_Direct"]:
        if c not in rev.columns:
            rev[c] = 0.0
        rev[c] = pd.to_numeric(rev[c], errors="coerce").fillna(0.0)
    rev["Revenue_Actual"] = rev["Revenue_Recognized"]
    rev["COGS_Actual"] = rev["COGS_Direct"]
    if not monthly_used and mode in ["YTD Comparable", "Annual Projection", "MCR Focus"]:
        rev["Revenue_Recognized"] = rev["Revenue_Recognized"] * rev["Year"].astype(float).map(rev_map).fillna(1.0)
        rev["COGS_Direct"] = rev["COGS_Direct"] * rev["Year"].astype(float).map(cost_map).fillna(1.0)

    # Headcount
    hc["Year"] = pd.to_numeric(hc["Year"], errors="coerce").astype("Int64")
    hc["Dept_ID"] = hc["Dept_ID"].astype(str).str.strip()
    hc["Avg_Headcount"] = pd.to_numeric(hc.get("Avg_Headcount"), errors="coerce").fillna(0.0)
    for c in ["New_Hires", "Exits", "Avg_FTE"]:
        if c in hc.columns:
            hc[c] = pd.to_numeric(hc[c], errors="coerce").fillna(0.0)

    # Payroll
    pay["Year"] = pd.to_numeric(pay["Year"], errors="coerce").astype("Int64")
    pay["Dept_ID"] = pay["Dept_ID"].astype(str).str.strip()
    for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
        if c in pay.columns:
            pay[c] = pd.to_numeric(pay[c], errors="coerce").fillna(0.0)
    if "Total_Manpower_Cost" not in pay.columns:
        cols = [c for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax"] if c in pay.columns]
        pay["Total_Manpower_Cost"] = pay[cols].sum(axis=1) if cols else 0.0
    pay["Total_Manpower_Cost_Actual"] = pay["Total_Manpower_Cost"]
    if not monthly_used and mode in ["YTD Comparable", "Annual Projection"]:
        for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
            if c in pay.columns:
                pay[c] = pay[c] * pay["Year"].astype(float).map(cost_map).fillna(1.0)

    hc2 = hc.merge(dim_small, on="Dept_ID", how="left")
    pay2 = pay.merge(dim_small, on="Dept_ID", how="left")

    rev_tot = rev.groupby("Year", as_index=False).agg(
        Total_Revenue=("Revenue_Recognized", "sum"),
        Total_COGS=("COGS_Direct", "sum"),
        Revenue_Actual=("Revenue_Actual", "sum"),
        COGS_Actual=("COGS_Actual", "sum"),
    )
    rev_tot["Gross_Profit"] = rev_tot["Total_Revenue"] - rev_tot["Total_COGS"]
    rev_tot["Gross_Margin_Pct"] = np.where(rev_tot["Total_Revenue"] > 0, rev_tot["Gross_Profit"] / rev_tot["Total_Revenue"], np.nan)

    hc_tot = hc2.groupby("Year", as_index=False).agg(
        Total_Headcount=("Avg_Headcount", "sum"),
        New_Hires=("New_Hires", "sum") if "New_Hires" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
        Exits=("Exits", "sum") if "Exits" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
    )
    pay_tot = pay2.groupby("Year", as_index=False).agg(
        Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
        Total_Manpower_Cost_Actual=("Total_Manpower_Cost_Actual", "sum"),
    )

    yr = rev_tot.merge(hc_tot, on="Year", how="outer").merge(pay_tot, on="Year", how="outer").sort_values("Year")
    yr["RPE"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Revenue"] / yr["Total_Headcount"], np.nan)
    yr["Cost_per_HC"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Headcount"], np.nan)
    yr["MCR_Pct"] = np.where(yr["Total_Revenue"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Revenue"], np.nan)
    yr["Revenue_per_Payroll"] = np.where(yr["Total_Manpower_Cost"] > 0, yr["Total_Revenue"] / yr["Total_Manpower_Cost"], np.nan)

    for c, out in [
        ("Total_Revenue", "Revenue_YoY"),
        ("Total_Headcount", "Headcount_YoY"),
        ("Total_Manpower_Cost", "ManpowerCost_YoY"),
        ("RPE", "RPE_YoY"),
        ("Gross_Margin_Pct", "GrossMargin_Delta"),
        ("MCR_Pct", "MCR_Delta"),
    ]:
        yr[out] = yr[c].diff() if "Delta" in out else yr[c].pct_change()

    # FJA
    fja_hc = hc2.groupby(["Year", "FJA_Category"], dropna=False, as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    fja_pay = pay2.groupby(["Year", "FJA_Category"], dropna=False, as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))
    fja = fja_hc.merge(fja_pay, on=["Year", "FJA_Category"], how="outer")
    fja["Cost_Share"] = fja["Manpower_Cost"] / fja.groupby("Year")["Manpower_Cost"].transform("sum")
    fja["HC_Share"] = fja["Headcount"] / fja.groupby("Year")["Headcount"].transform("sum")

    # Department
    hcd = hc2.groupby(["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    payd = pay2.groupby(["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))
    dept = hcd.merge(payd, on=["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], how="outer")
    if not monthly_dept_revenue.empty:
        dept = dept.merge(monthly_dept_revenue, on=["Year", "Dept_ID"], how="left")
    elif "REVENUE_BY_DEPT" in sheets and not sheets["REVENUE_BY_DEPT"].empty:
        rd = sheets["REVENUE_BY_DEPT"].copy()
        rd["Year"] = pd.to_numeric(rd["Year"], errors="coerce").astype("Int64")
        rd["Dept_ID"] = rd["Dept_ID"].astype(str).str.strip()
        rd["Revenue_Recognized"] = pd.to_numeric(rd.get("Revenue_Recognized"), errors="coerce").fillna(0.0)
        if mode in ["YTD Comparable", "Annual Projection", "MCR Focus"]:
            rd["Revenue_Recognized"] = rd["Revenue_Recognized"] * rd["Year"].astype(float).map(rev_map).fillna(1.0)
        dept_rev = rd.groupby(["Year", "Dept_ID"], as_index=False).agg(Dept_Revenue=("Revenue_Recognized", "sum"))
        dept = dept.merge(dept_rev, on=["Year", "Dept_ID"], how="left")
    else:
        dept["Dept_Revenue"] = np.nan
    dept["Dept_RPE"] = np.where(dept["Headcount"] > 0, dept["Dept_Revenue"] / dept["Headcount"], np.nan)
    dept["Revenue_per_Cost"] = np.where(dept["Manpower_Cost"] > 0, dept["Dept_Revenue"] / dept["Manpower_Cost"], np.nan)
    dept["Cost_per_HC"] = np.where(dept["Headcount"] > 0, dept["Manpower_Cost"] / dept["Headcount"], np.nan)

    # Capacity
    if "CAPACITY_INDICATOR" in sheets and not sheets["CAPACITY_INDICATOR"].empty:
        cap = sheets["CAPACITY_INDICATOR"].copy()
        cap["Year"] = pd.to_numeric(cap["Year"], errors="coerce").astype("Int64")
        for c in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Incident_Count", "Turnover_Count"]:
            if c in cap.columns:
                cap[c] = pd.to_numeric(cap[c], errors="coerce").fillna(0.0)
    else:
        cap = pd.DataFrame()

    # Project
    if "PROJECT_MARGIN" in sheets and not sheets["PROJECT_MARGIN"].empty:
        pm = sheets["PROJECT_MARGIN"].copy()
        pm["Year"] = pd.to_numeric(pm["Year"], errors="coerce").astype("Int64")
        for c in ["Revenue", "COGS", "Gross_Profit"]:
            if c in pm.columns:
                pm[c] = pd.to_numeric(pm[c], errors="coerce").fillna(0.0)
        if mode in ["YTD Comparable", "Annual Projection", "MCR Focus"] and not monthly_used:
            factor = pm["Year"].astype(float).map(rev_map).fillna(1.0)
            for c in ["Revenue", "COGS", "Gross_Profit"]:
                if c in pm.columns:
                    pm[c] = pm[c] * factor
        if "Gross_Profit" not in pm.columns:
            pm["Gross_Profit"] = pm["Revenue"] - pm["COGS"]
        pm["Gross_Margin_Pct"] = np.where(pm["Revenue"] > 0, pm["Gross_Profit"] / pm["Revenue"], np.nan)
    else:
        pm = pd.DataFrame()

    # Project Productivity: revenue contribution by common telecom/IT integrator service lines.
    if "PROJECT_PRODUCTIVITY" in sheets and not sheets["PROJECT_PRODUCTIVITY"].empty:
        pp = sheets["PROJECT_PRODUCTIVITY"].copy()
        pp["Year"] = pd.to_numeric(pp.get("Year"), errors="coerce").astype("Int64")
        for c in ["Revenue", "COGS", "Manpower_Cost", "Project_Hours", "HC_Allocated", "Sites_or_Tickets", "SLA_Breach_Count"]:
            if c not in pp.columns:
                pp[c] = 0.0
            pp[c] = pd.to_numeric(pp[c], errors="coerce").fillna(0.0)
        if "Project_Category" not in pp.columns:
            pp["Project_Category"] = "Unmapped"
        if "Project_Name" not in pp.columns:
            pp["Project_Name"] = pp["Project_Category"]
        if "Business_Line" not in pp.columns:
            pp["Business_Line"] = "Telecom & IT Integrator"
        if mode in ["YTD Comparable", "Annual Projection", "MCR Focus"] and not monthly_used:
            _pp_rev_factor = pp["Year"].astype(float).map(rev_map).fillna(1.0)
            _pp_cost_factor = pp["Year"].astype(float).map(cost_map).fillna(1.0)
            pp["Revenue"] = pp["Revenue"] * _pp_rev_factor
            pp["COGS"] = pp["COGS"] * _pp_cost_factor
            pp["Manpower_Cost"] = pp["Manpower_Cost"] * _pp_cost_factor
        pp["Gross_Profit"] = pp["Revenue"] - pp["COGS"]
        pp["Gross_Margin_Pct"] = np.where(pp["Revenue"] > 0, pp["Gross_Profit"] / pp["Revenue"], np.nan)
        pp["Revenue_per_HC"] = np.where(pp["HC_Allocated"] > 0, pp["Revenue"] / pp["HC_Allocated"], np.nan)
        pp["Revenue_per_Hour"] = np.where(pp["Project_Hours"] > 0, pp["Revenue"] / pp["Project_Hours"], np.nan)
        pp["MCR_Pct"] = np.where(pp["Revenue"] > 0, pp["Manpower_Cost"] / pp["Revenue"], np.nan)
    elif pm is not None and not pm.empty:
        pp = pm.copy()
        if "Project_Category" not in pp.columns:
            pp["Project_Category"] = pp.get("Business_Line", "Project")
        if "Project_Name" not in pp.columns:
            pp["Project_Name"] = pp.get("Project_ID", pp["Project_Category"])
        if "Manpower_Cost" not in pp.columns:
            pp["Manpower_Cost"] = 0.0
        if "HC_Allocated" not in pp.columns:
            pp["HC_Allocated"] = 0.0
        if "Project_Hours" not in pp.columns:
            pp["Project_Hours"] = 0.0
        if "Sites_or_Tickets" not in pp.columns:
            pp["Sites_or_Tickets"] = 0.0
        if "SLA_Breach_Count" not in pp.columns:
            pp["SLA_Breach_Count"] = 0.0
        pp["Revenue_per_HC"] = np.where(pp["HC_Allocated"] > 0, pp["Revenue"] / pp["HC_Allocated"], np.nan)
        pp["Revenue_per_Hour"] = np.where(pp["Project_Hours"] > 0, pp["Revenue"] / pp["Project_Hours"], np.nan)
        pp["MCR_Pct"] = np.where(pp["Revenue"] > 0, pp["Manpower_Cost"] / pp["Revenue"], np.nan)
    else:
        pp = pd.DataFrame()

    return {
        "yearly": yr, "fja": fja, "dept": dept, "capacity": cap, "project_margin": pm,
        "project_productivity": pp,
        "calendar": calendar, "mode_note": mode_note, "comparable_months": comparable_months,
        "monthly_ytd_used": monthly_used,
    }


# =========================
# ENGINES
# =========================
def business_posture(row, targets):
    rev_g = row.get("Revenue_YoY")
    hc_g = row.get("Headcount_YoY")
    rpe_g = row.get("RPE_YoY")
    mcr = row.get("MCR_Pct")
    rev_min = get_target(targets, "Revenue_Growth_Min", 0.20)
    rpe_min = get_target(targets, "RPE_Growth_Min", 0.15)
    hc_max = get_target(targets, "Headcount_Growth_Max_High_Leverage", 0.10)
    mcr_watch = get_target(targets, "MCR_Watch_Max", 0.09)
    if pd.notna(rev_g) and rev_g <= 0:
        return {"state": "DEFENSIVE", "tone": "error", "hint": "Revenue melemah; proteksi margin, cash, dan fungsi kritikal."}
    if pd.notna(rev_g) and pd.notna(rpe_g) and pd.notna(hc_g) and pd.notna(mcr) and rev_g >= rev_min and rpe_g >= rpe_min and hc_g <= hc_max and mcr <= mcr_watch:
        return {"state": "HIGH_LEVERAGE", "tone": "success", "hint": "Revenue & RPE tumbuh dengan HC terkendali dan MCR sehat/watch."}
    if pd.notna(mcr) and mcr > mcr_watch:
        return {"state": "COST_PRESSURE", "tone": "warning", "hint": "People-cost mulai menekan revenue."}
    if pd.notna(rev_g) and pd.notna(hc_g) and rev_g > 0.10 and hc_g > 0.15 and (pd.isna(rpe_g) or rpe_g <= 0):
        return {"state": "CAPACITY_RISK", "tone": "warning", "hint": "Growth mulai dibeli dengan tambahan kapasitas."}
    return {"state": "BALANCED", "tone": "info", "hint": "Revenue, capacity, dan people-cost relatif seimbang."}


def margin_quality(gm, targets):
    strong = get_target(targets, "Gross_Margin_Strong_Min", 0.30)
    healthy = get_target(targets, "Gross_Margin_Healthy_Min", 0.20)
    watch = get_target(targets, "Gross_Margin_Watch_Min", 0.15)
    pressure = get_target(targets, "Gross_Margin_Pressure_Min", 0.10)
    if pd.isna(gm):
        return {"state": "UNKNOWN", "tone": "info", "hint": "Gross Margin belum tersedia."}
    if gm >= strong:
        return {"state": "STRONG", "tone": "success", "hint": "Margin sangat kuat."}
    if gm >= healthy:
        return {"state": "HEALTHY", "tone": "success", "hint": "Margin sehat untuk telco/IT integrator project-based."}
    if gm >= watch:
        return {"state": "WATCH", "tone": "warning", "hint": "Margin perlu dipantau."}
    if gm >= pressure:
        return {"state": "PRESSURE", "tone": "warning", "hint": "Margin mulai tertekan."}
    return {"state": "CRITICAL", "tone": "error", "hint": "Margin kritikal."}


def mcr_health(mcr, targets):
    ultra = get_target(targets, "MCR_Ultra_Efficiency_Max", 0.05)
    healthy = get_target(targets, "MCR_Healthy_Max", 0.07)
    watch = get_target(targets, "MCR_Watch_Max", 0.09)
    pressure = get_target(targets, "MCR_Cost_Pressure_Max", 0.14)
    if pd.isna(mcr):
        return {"state": "UNKNOWN", "tone": "info", "hint": "MCR belum tersedia."}
    if mcr < ultra:
        return {"state": "ULTRA_EFFICIENCY", "tone": "warning", "hint": "Sangat efisien; validasi capacity sustainability."}
    if mcr <= healthy:
        return {"state": "HEALTHY", "tone": "success", "hint": "MCR sehat 5–7%."}
    if mcr <= watch:
        return {"state": "WATCH", "tone": "warning", "hint": "MCR mulai perlu dipantau."}
    if mcr <= pressure:
        return {"state": "COST_PRESSURE", "tone": "warning", "hint": "People-cost menekan revenue."}
    return {"state": "CRITICAL", "tone": "error", "hint": "People-cost kritikal."}


def capacity_risk(row, cap_df):
    year = int(row.get("Year"))
    d = {"max_utilization": np.nan, "total_overtime": 0, "total_backlog": 0, "total_sla_breach": 0, "total_turnover": 0, "critical_dependency": "UNKNOWN"}
    if cap_df is None or cap_df.empty:
        return {"state": "LOW", "tone": "success", "score": 0, "recommendation": "No Hiring Required", "hint": "CAPACITY_INDICATOR belum tersedia.", "details": d}
    c = cap_df[cap_df["Year"] == year].copy()
    if c.empty:
        return {"state": "LOW", "tone": "success", "score": 0, "recommendation": "No Hiring Required", "hint": "Data capacity tahun terpilih belum tersedia.", "details": d}
    for col in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Turnover_Count"]:
        if col not in c.columns:
            c[col] = 0
        c[col] = pd.to_numeric(c[col], errors="coerce").fillna(0)
    max_util = float(c["Avg_Utilization_Pct"].max())
    overtime = float(c["Overtime_Hours"].sum())
    backlog = float(c["Backlog_Count"].sum())
    sla = float(c["SLA_Breach_Count"].sum())
    turnover = float(c["Turnover_Count"].sum())
    dep = c.get("Critical_Role_Dependency", pd.Series(dtype=str)).astype(str).str.upper().str.strip()
    dep_value = {1: "LOW", 2: "MEDIUM", 3: "HIGH", 4: "CRITICAL"}.get(int(dep.map({"LOW": 1, "MEDIUM": 2, "HIGH": 3, "CRITICAL": 4}).fillna(0).max()), "UNKNOWN") if len(dep) else "UNKNOWN"
    score, reasons = 0, []
    if max_util >= 0.90: score += 25; reasons.append("Utilization >90%")
    elif max_util >= 0.85: score += 18; reasons.append("Utilization 85-90%")
    elif max_util >= 0.75: score += 10; reasons.append("Utilization 75-85%")
    if overtime > 400: score += 20; reasons.append("Overtime >400 jam")
    elif overtime > 250: score += 15; reasons.append("Overtime 250-400 jam")
    elif overtime > 100: score += 8; reasons.append("Overtime 100-250 jam")
    if backlog >= 20: score += 20; reasons.append("Backlog >=20")
    elif backlog >= 10: score += 12; reasons.append("Backlog 10-19")
    elif backlog > 0: score += 6; reasons.append("Backlog muncul")
    if sla >= 5: score += 15; reasons.append("SLA breach >=5")
    elif sla > 0: score += 8; reasons.append("Ada SLA breach")
    if turnover >= 10: score += 10; reasons.append("Turnover >=10")
    elif turnover > 0: score += 5; reasons.append("Ada turnover")
    if dep_value in ["HIGH", "CRITICAL"]: score += 10; reasons.append("Critical dependency tinggi")
    elif dep_value == "MEDIUM": score += 5; reasons.append("Critical dependency medium")
    score = int(min(100, round(score)))
    if score > 75: state, tone, rec = "CRITICAL", "error", "Immediate Capacity Action Required"
    elif score >= 56: state, tone, rec = "HIGH", "warning", "Capacity Expansion Review"
    elif score >= 31: state, tone, rec = "MEDIUM", "warning", "Selective Hiring"
    else: state, tone, rec = "LOW", "success", "No Hiring Required"
    d = {"max_utilization": max_util, "total_overtime": overtime, "total_backlog": backlog, "total_sla_breach": sla, "total_turnover": turnover, "critical_dependency": dep_value}
    return {"state": state, "tone": tone, "score": score, "recommendation": rec, "hint": "; ".join(reasons) if reasons else "Tidak ada sinyal capacity besar.", "details": d}


def score_from_state(state):
    return {"STRONG": 100, "HEALTHY": 85, "WATCH": 60, "PRESSURE": 35, "CRITICAL": 10, "UNKNOWN": 50}.get(state, 50)


def score_mcr(state):
    return {"HEALTHY": 95, "WATCH": 75, "ULTRA_EFFICIENCY": 70, "COST_PRESSURE": 40, "CRITICAL": 10, "UNKNOWN": 50}.get(state, 50)


def score_growth(v, strong=0.20, good=0.10):
    if pd.isna(v): return 50
    if v >= strong: return 100
    if v >= good: return 80
    if v >= 0: return 60
    if v >= -0.10: return 35
    return 15


def score_capacity(state):
    return {"LOW": 100, "MEDIUM": 75, "HIGH": 40, "CRITICAL": 10}.get(state, 50)


def board_score(row, targets, mq, mh, cr):
    comps = {
        "Revenue Growth": score_growth(row.get("Revenue_YoY"), get_target(targets, "Revenue_Growth_Min", 0.20), 0.10),
        "RPE Growth": score_growth(row.get("RPE_YoY"), get_target(targets, "RPE_Growth_Min", 0.15), 0.08),
        "Margin Quality": score_from_state(mq["state"]),
        "MCR Health": score_mcr(mh["state"]),
        "Capacity Risk": score_capacity(cr["state"]),
    }
    weights = {"Revenue Growth": 0.25, "RPE Growth": 0.25, "Margin Quality": 0.20, "MCR Health": 0.15, "Capacity Risk": 0.15}
    score = int(round(sum(comps[k] * weights[k] for k in comps)))
    if score >= 85: label, tone, icon = "Growth Ready", "success", "🟢"
    elif score >= 70: label, tone, icon = "Optimize Growth", "info", "🔵"
    elif score >= 55: label, tone, icon = "Selective Hiring", "warning", "🟡"
    elif score >= 40: label, tone, icon = "Capacity Alert", "warning", "🟠"
    else: label, tone, icon = "Defensive Mode", "error", "🔴"
    return {"score": score, "label": label, "tone": tone, "icon": icon, "components": comps, "weights": weights}


def management_action(bp, mq, mh, cr):
    if bp["state"] == "HIGH_LEVERAGE":
        title = "Selective Scale-Up with Capacity Guardrail"
        actions = ["Pertahankan disiplin HC.", "Selective hiring pada bottleneck revenue/delivery.", "Monitor overtime, backlog, SLA, dan role dependency."]
    elif bp["state"] == "DEFENSIVE":
        title = "Margin & Cash Protection"
        actions = ["Freeze hiring non-critical.", "Lindungi fungsi revenue/delivery kritikal.", "Fokus cash discipline dan margin recovery."]
    elif bp["state"] == "COST_PRESSURE":
        title = "People-Cost Control"
        actions = ["Perketat approval hiring berbasis ROI.", "Audit overtime, bonus, benefit, dan manpower cost.", "Review unit economics project/service."]
    else:
        title = "Balanced Optimization"
        actions = ["Optimalkan struktur sebelum ekspansi.", "Hiring hanya pada bottleneck terbukti.", "Jaga MCR, RPE, dan support ratio."]
    if mq["state"] in ["WATCH", "PRESSURE", "CRITICAL"]:
        actions.append("Jalankan margin quality review.")
    if cr["state"] in ["HIGH", "CRITICAL"]:
        actions.append("Lakukan capacity risk review bulanan.")
    return title, actions


def build_management_insights(row, bp, mq, mh, cr):
    insights = []
    rev_g, hc_g, rpe_g = row.get("Revenue_YoY"), row.get("Headcount_YoY"), row.get("RPE_YoY")
    if pd.notna(rev_g) and pd.notna(hc_g):
        if rev_g > hc_g:
            insights.append(("Revenue grows faster than Headcount", "GOOD", "Revenue tumbuh lebih cepat dari HC.", "Growth relatif produktif."))
        else:
            insights.append(("Headcount grows faster than Revenue", "WATCH", "HC tumbuh lebih cepat dari revenue.", "Tambahan kapasitas belum sepenuhnya menjadi revenue."))
    if pd.notna(rev_g) and pd.notna(rpe_g):
        if rev_g > 0 and rpe_g < 0:
            insights.append(("Revenue rises but RPE declines", "WATCH", "Revenue naik tetapi RPE turun.", "Growth kemungkinan dibeli dengan tambahan HC."))
        elif rev_g > 0 and rpe_g > 0:
            insights.append(("Revenue and RPE improve together", "GOOD", "Revenue dan RPE sama-sama tumbuh.", "Pertumbuhan berkualitas dari sisi productivity."))
    if mh["state"] == "ULTRA_EFFICIENCY" and cr["state"] in ["HIGH", "CRITICAL"]:
        insights.append(("Low MCR but high capacity risk", "CRITICAL", "MCR rendah tetapi capacity risk tinggi.", "Efisiensi bisa menyembunyikan overload."))
    if mq["state"] in ["WATCH", "PRESSURE", "CRITICAL"]:
        insights.append(("Margin quality needs attention", "WATCH", f"Gross Margin {mq['state']}.", "Revenue perlu divalidasi kualitas marginnya."))
    if cr["state"] in ["HIGH", "CRITICAL"]:
        insights.append(("Capacity bottleneck may limit growth", "CRITICAL", f"Capacity Risk {cr['state']}.", "Growth berikutnya bisa tertahan delivery capacity."))
    return insights or [("No major conflict detected", "BALANCED", "Indikator utama relatif seimbang.", "Lanjutkan monitoring bulanan.")]


def hiring_decision(dept_df, cap_df, year):
    if dept_df is None or dept_df.empty:
        return pd.DataFrame()
    d = dept_df[dept_df["Year"] == int(year)].copy()
    if d.empty:
        return pd.DataFrame()
    if cap_df is not None and not cap_df.empty and "Dept_ID" in cap_df.columns:
        cap = cap_df[cap_df["Year"] == int(year)].copy()
        for c in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Turnover_Count"]:
            if c not in cap.columns: cap[c] = 0
            cap[c] = pd.to_numeric(cap[c], errors="coerce").fillna(0)
        sig = cap.groupby("Dept_ID", as_index=False).agg(Avg_Utilization=("Avg_Utilization_Pct", "max"), Overtime=("Overtime_Hours", "sum"), Backlog=("Backlog_Count", "sum"), SLA_Breach=("SLA_Breach_Count", "sum"), Turnover=("Turnover_Count", "sum"))
        d = d.merge(sig, on="Dept_ID", how="left")
    else:
        for c in ["Avg_Utilization", "Overtime", "Backlog", "SLA_Breach", "Turnover"]:
            d[c] = 0
    d[["Avg_Utilization", "Overtime", "Backlog", "SLA_Breach", "Turnover"]] = d[["Avg_Utilization", "Overtime", "Backlog", "SLA_Breach", "Turnover"]].fillna(0)
    recs, reasons = [], []
    for _, r in d.iterrows():
        fja = str(r.get("FJA_Category", ""))
        rpc = r.get("Revenue_per_Cost")
        util, ot, bl, sla = r.get("Avg_Utilization", 0), r.get("Overtime", 0), r.get("Backlog", 0), r.get("SLA_Breach", 0)
        high_capacity = util >= 0.85 or ot > 250 or bl >= 10 or sla > 0
        efficient = pd.notna(rpc) and rpc >= 4
        low_eff = pd.notna(rpc) and rpc < 1.5
        if high_capacity and ("Revenue Enabler" in fja or "Revenue Generator" in fja or efficient):
            rec, reason = "HIRE", "Bottleneck/capacity pressure pada fungsi revenue/delivery."
        elif high_capacity:
            rec, reason = "SELECTIVE", "Ada tekanan kapasitas, perlu validasi workload dan kontribusi."
        elif low_eff and "Support" in fja:
            rec, reason = "FREEZE", "Support cost intensity tinggi; optimasi proses dahulu."
        elif low_eff:
            rec, reason = "HOLD", "Revenue per cost belum kuat."
        else:
            rec, reason = "HOLD", "Belum ada sinyal kuat untuk tambah HC."
        recs.append(rec); reasons.append(reason)
    d["Recommendation"] = recs
    d["Reason"] = reasons
    return d


def hc_guardrails(row, targets):
    revenue = _safe_float(row.get("Total_Revenue"))
    cost = _safe_float(row.get("Total_Manpower_Cost"))
    hc = _safe_float(row.get("Total_Headcount"))
    cph = _safe_float(row.get("Cost_per_HC"))
    target = get_target(targets, "MCR_Healthy_Max", 0.07)
    max_cost = revenue * target if pd.notna(revenue) else np.nan
    remaining = max_cost - cost if pd.notna(max_cost) and pd.notna(cost) else np.nan
    add_hc = np.floor(remaining / cph) if pd.notna(remaining) and pd.notna(cph) and cph > 0 else np.nan
    max_hc = hc + add_hc if pd.notna(hc) and pd.notna(add_hc) else np.nan
    req_rev = cost / target if pd.notna(cost) and target > 0 else np.nan
    return {"target_mcr": target, "remaining_budget": remaining, "additional_hc_capacity": add_hc, "max_hc": max_hc, "required_revenue": req_rev}


def roadmap(bp, mq, mh, cr):
    rows = []
    def add(phase, priority, action, owner, metric):
        rows.append({"Phase": phase, "Priority": priority, "Action": action, "Owner": owner, "Success Metric": metric})
    if cr["state"] in ["HIGH", "CRITICAL"]:
        add("0-3 Months", "Critical", "Capacity review: overtime, backlog, SLA, dependency.", "COO / HR", "Capacity score turun 15 poin")
    if mh["state"] in ["COST_PRESSURE", "CRITICAL"]:
        add("0-3 Months", "Critical", "Freeze non-critical hiring dan review manpower cost.", "CFO / HR", "MCR turun menuju <=9%")
    if mq["state"] in ["WATCH", "PRESSURE", "CRITICAL"]:
        add("0-3 Months", "High", "Margin quality review per project/customer.", "CFO / Sales / Ops", "Gross Margin kembali sehat")
    add("3-6 Months", "High", "Selective hiring pada fungsi revenue/delivery bottleneck.", "CEO / HR / COO", "Hiring berbasis ROI dan evidence")
    add("3-6 Months", "Medium", "Productivity program: SOP, automation, redistribusi workload.", "HR / Dept Head", "RPE naik, overtime turun")
    add("6-12 Months", "Medium", "Review ratio revenue generator, enabler, support.", "CEO / HR", "Support cost share terkendali")
    return pd.DataFrame(rows)


def closing_text(row, bs, bp, mq, mh, cr, targets, currency):
    guard = hc_guardrails(row, targets)
    return (
        f"Perusahaan berada pada status **{bs['label']}** dengan Company Health Score **{bs['score']}/100**. "
        f"Revenue terbaca **{_money(row.get('Total_Revenue'), currency)}**, MCR **{_pct(row.get('MCR_Pct'))}**, "
        f"RPE **{_money(row.get('RPE'), currency)}**, Gross Margin **{_pct(row.get('Gross_Margin_Pct'))}**, "
        f"dan Capacity Risk **{cr['state']}**. Dengan target MCR sehat **{_pct(guard['target_mcr'])}**, "
        f"revenue minimum untuk menopang manpower cost saat ini adalah **{_money(guard['required_revenue'], currency)}**. "
        f"Rekomendasi: lakukan hiring selektif pada fungsi bottleneck, tahan support non-critical, dan gunakan MCR sebagai guardrail."
    )


def make_priority_df(bp, mq, mh, cr):
    def sev(x):
        s = str(x).upper()
        if s in ["CRITICAL", "DEFENSIVE", "COST_PRESSURE", "HIGH"]: return 90
        if s in ["WATCH", "PRESSURE", "MEDIUM", "CAPACITY_RISK", "ULTRA_EFFICIENCY"]: return 65
        if s in ["HEALTHY", "BALANCED"]: return 35
        if s in ["STRONG", "LOW", "HIGH_LEVERAGE"]: return 20
        return 50
    return pd.DataFrame([
        {"Area": "Business Posture", "Status": bp["state"], "Priority": sev(bp["state"])},
        {"Area": "Margin Quality", "Status": mq["state"], "Priority": sev(mq["state"])},
        {"Area": "MCR Health", "Status": mh["state"], "Priority": sev(mh["state"])},
        {"Area": "Capacity Risk", "Status": cr["state"], "Priority": sev(cr["state"])},
    ])


def framework_from_targets(targets_df, targets):
    if targets_df is not None and not targets_df.empty and "KPI" in targets_df.columns:
        base = targets_df.copy()
    else:
        base = pd.DataFrame()
    rows = [
        ["Revenue Growth", "15–25% healthy; >25% strong", "Annual reports + management target", "Measures growth momentum", "High"],
        ["RPE Growth", ">=15% strong", "Human capital productivity benchmark", "Measures revenue productivity per employee", "Medium"],
        ["Headcount Growth", "<=10% controlled", "Organization design / workforce planning practice", "Controls fixed cost expansion", "Medium"],
        ["MCR", "5–7% healthy; 7–9% watch", "Internal benchmark for project-based telco/IT integrator", "Measures people-cost leverage", "Medium"],
        ["Gross Margin", "20–30% healthy; >30% strong", "Telecom/IT integrator project-margin benchmark", "Measures project revenue quality", "Medium"],
        ["Capacity Utilization", "75–85% optimal; >90% risk", "Capacity planning and service operations practice", "Measures overload risk", "Medium"],
    ]
    return pd.DataFrame(rows, columns=["KPI", "Target Framework", "Benchmark Basis", "Why It Matters", "Confidence"])



def build_projection_scenarios(row, targets, currency):
    base_revenue = _safe_float(row.get("Total_Revenue"))
    base_cost = _safe_float(row.get("Total_Manpower_Cost"))
    base_hc = _safe_float(row.get("Total_Headcount"))
    base_cogs = _safe_float(row.get("Total_COGS"))
    cost_per_hc = _safe_float(row.get("Cost_per_HC"))
    gm_base = _safe_float(row.get("Gross_Margin_Pct"))

    if pd.isna(cost_per_hc) or cost_per_hc <= 0:
        cost_per_hc = base_cost / base_hc if pd.notna(base_cost) and pd.notna(base_hc) and base_hc else 0

    return {
        "base_revenue": base_revenue,
        "base_cost": base_cost,
        "base_hc": base_hc,
        "base_cogs": base_cogs,
        "cost_per_hc": cost_per_hc,
        "gm_base": gm_base,
        "target_mcr": get_target(targets, "MCR_Healthy_Max", 0.07),
        "target_gm": get_target(targets, "Gross_Margin_Healthy_Min", 0.20),
    }


def scenario_grid(row, targets, revenue_growth, gross_margin_target, avg_cost_new_hc, max_hc_add=50):
    base = build_projection_scenarios(row, targets, "IDR")
    base_revenue = base["base_revenue"]
    base_hc = base["base_hc"]
    base_cost = base["base_cost"]

    rows = []
    for add_hc in range(0, max_hc_add + 1):
        scenario_revenue = base_revenue * (1 + revenue_growth) if pd.notna(base_revenue) else np.nan
        scenario_cogs = scenario_revenue * (1 - gross_margin_target) if pd.notna(scenario_revenue) else np.nan
        scenario_manpower = base_cost + (add_hc * avg_cost_new_hc) if pd.notna(base_cost) else np.nan
        scenario_hc = base_hc + add_hc if pd.notna(base_hc) else np.nan
        rows.append({
            "Additional_HC": add_hc,
            "Projected_Revenue": scenario_revenue,
            "Projected_COGS": scenario_cogs,
            "Projected_Gross_Margin": gross_margin_target,
            "Projected_Manpower_Cost": scenario_manpower,
            "Projected_HC": scenario_hc,
            "Projected_RPE": scenario_revenue / scenario_hc if pd.notna(scenario_revenue) and pd.notna(scenario_hc) and scenario_hc else np.nan,
            "Projected_MCR": scenario_manpower / scenario_revenue if pd.notna(scenario_manpower) and pd.notna(scenario_revenue) and scenario_revenue else np.nan,
        })
    return pd.DataFrame(rows)


def required_revenue_for_hc(row, target_mcr, avg_cost_new_hc, add_hc):
    base_cost = _safe_float(row.get("Total_Manpower_Cost"))
    projected_cost = base_cost + (add_hc * avg_cost_new_hc) if pd.notna(base_cost) else np.nan
    return projected_cost / target_mcr if pd.notna(projected_cost) and target_mcr > 0 else np.nan





# =========================
# BOARD GUIDANCE & SUMMARIES
# =========================
def render_business_question(question: str, decision_focus: str = ""):
    focus_html = f"<div class='board-decision-focus'><b>Decision focus:</b> {html.escape(decision_focus)}</div>" if decision_focus else ""
    st.markdown(
        f"""
        <div class="board-question-card">
            <div class="board-question-label">PERTANYAAN BISNIS</div>
            <div class="board-question-text">{html.escape(question)}</div>
            {focus_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_board_guidance(title: str, how_items: list[str], summary: str, business_question: str, tone: str = "info"):
    """Render How to Read and Executive Summary as collapsed sections."""
    tone_icon = {
        "success": "🟢",
        "warning": "🟠",
        "error": "🔴",
        "info": "🔵",
    }.get(tone, "🔵")

    with st.expander(f"ⓘ How to Read — {title}", expanded=False):
        for item in how_items:
            st.markdown(f"- {item}")

    with st.expander(f"{tone_icon} Executive Summary — {title}", expanded=False):
        st.markdown(summary, unsafe_allow_html=True)
        st.markdown(f"**Pertanyaan Bisnis:** {business_question}", unsafe_allow_html=True)


def _summary_tone(bs, mq, mh, cr):
    if bs.get("score", 50) < 40 or mq.get("state") == "CRITICAL" or mh.get("state") == "CRITICAL" or cr.get("state") == "CRITICAL":
        return "error"
    if bs.get("score", 50) < 70 or mq.get("state") in ["WATCH", "PRESSURE"] or mh.get("state") in ["WATCH", "COST_PRESSURE"] or cr.get("state") in ["MEDIUM", "HIGH"]:
        return "warning"
    return "success"


def executive_summary_text(latest, bs, bp, mq, mh, cr, period, mode, currency):
    return (
        f"Pada <b>{html.escape(str(period.get('Period_Label', 'periode terpilih')))}</b> dengan mode <b>{html.escape(mode)}</b>, "
        f"Company Health berada pada <b>{bs.get('score')}/100 ({html.escape(str(bs.get('label')) )})</b>. "
        f"Revenue tercatat <b>{_money(latest.get('Total_Revenue'), currency)}</b>, MCR <b>{_pct(latest.get('MCR_Pct'))}</b>, "
        f"Gross Margin <b>{_pct(latest.get('Gross_Margin_Pct'))}</b>, RPE <b>{_money(latest.get('RPE'), currency)}</b>, "
        f"dan Capacity Risk <b>{html.escape(str(cr.get('state')))}</b>. Business posture saat ini adalah <b>{html.escape(str(bp.get('state')))}</b>."
    )


def revenue_summary_text(latest, mq, mh, currency):
    rev_g = latest.get("Revenue_YoY")
    hc_g = latest.get("Headcount_YoY")
    rpe_g = latest.get("RPE_YoY")
    quality = "produktif" if pd.notna(rev_g) and pd.notna(hc_g) and rev_g > hc_g and pd.notna(rpe_g) and rpe_g > 0 else "belum sepenuhnya produktif"
    return (
        f"Revenue tercatat <b>{_money(latest.get('Total_Revenue'), currency)}</b> dengan pertumbuhan <b>{_pct(rev_g)}</b>. "
        f"Headcount berubah <b>{_pct(hc_g)}</b> dan RPE berubah <b>{_pct(rpe_g)}</b>, sehingga pertumbuhan saat ini dinilai <b>{quality}</b>. "
        f"Gross Margin berada pada status <b>{html.escape(str(mq.get('state')))}</b> dan MCR pada status <b>{html.escape(str(mh.get('state')))}</b>."
    )


def workforce_summary_text(latest, dept, fja_latest, currency):
    top_cost = "-"
    if fja_latest is not None and not fja_latest.empty:
        r = fja_latest.sort_values("Manpower_Cost", ascending=False).iloc[0]
        top_cost = f"{r.get('FJA_Category')} ({_pct(r.get('Cost_Share'))} cost share)"
    high_cost_dept = "-"
    if dept is not None and not dept.empty:
        r = dept.sort_values("Manpower_Cost", ascending=False).iloc[0]
        high_cost_dept = f"{r.get('Dept_Name')} ({_money(r.get('Manpower_Cost'), currency)})"
    return (
        f"Rata-rata headcount adalah <b>{_num(latest.get('Total_Headcount'))}</b> dengan manpower cost <b>{_money(latest.get('Total_Manpower_Cost'), currency)}</b>. "
        f"Kelompok FJA dengan biaya terbesar adalah <b>{html.escape(str(top_cost))}</b>, sedangkan departemen dengan manpower cost terbesar adalah <b>{html.escape(str(high_cost_dept))}</b>. "
        "Struktur SDM harus dinilai dari keseimbangan antara jumlah orang, biaya, kontribusi revenue, dan perlindungan delivery."
    )


def capacity_summary_text(cr):
    d = cr.get("details", {})
    return (
        f"Capacity Risk berada pada level <b>{html.escape(str(cr.get('state')))}</b> dengan skor <b>{cr.get('score')}/100</b>. "
        f"Utilization maksimum <b>{_pct(d.get('max_utilization'))}</b>, overtime <b>{_num(d.get('total_overtime'))} jam</b>, "
        f"backlog <b>{_num(d.get('total_backlog'))}</b>, dan SLA breach <b>{_num(d.get('total_sla_breach'))}</b>. "
        f"Rekomendasi sistem: <b>{html.escape(str(cr.get('recommendation')))}</b>."
    )


def strategic_summary_text(bs, bp, mq, mh, cr, action_title):
    return (
        f"Prioritas keputusan ditentukan oleh Company Health <b>{bs.get('score')}/100</b>, business posture <b>{html.escape(str(bp.get('state')))}</b>, "
        f"margin <b>{html.escape(str(mq.get('state')))}</b>, MCR <b>{html.escape(str(mh.get('state')))}</b>, dan capacity risk <b>{html.escape(str(cr.get('state')))}</b>. "
        f"Fokus manajemen saat ini adalah <b>{html.escape(str(action_title))}</b>."
    )


def benchmark_summary_text(latest, mq, mh, cr):
    return (
        f"KPI aktual menunjukkan Revenue Growth <b>{_pct(latest.get('Revenue_YoY'))}</b>, RPE Growth <b>{_pct(latest.get('RPE_YoY'))}</b>, "
        f"MCR <b>{_pct(latest.get('MCR_Pct'))} ({html.escape(str(mh.get('state')))})</b>, Gross Margin <b>{_pct(latest.get('Gross_Margin_Pct'))} ({html.escape(str(mq.get('state')))})</b>, "
        f"dan Capacity Score <b>{cr.get('score')}/100 ({html.escape(str(cr.get('state')))})</b>. Benchmark berfungsi sebagai management guardrail, bukan standar regulasi."
    )


def appendix_summary_text(period, mode):
    return (
        f"Appendix menjelaskan formula, definisi, data source, dan governance yang dipakai SEP V4. "
        f"Periode aktif adalah <b>{html.escape(str(period.get('Period_Label')))}</b> dengan mode <b>{html.escape(mode)}</b>. "
        "Bagian ini digunakan untuk memverifikasi asal angka dan memastikan hasil dashboard dapat diaudit."
    )


def inject_presentation_readability_css():
    st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-size: 15.5px !important;
    }
    .main .block-container {
        max-width: 1540px;
        padding-top: 1.15rem;
        padding-left: 2.4rem;
        padding-right: 2.4rem;
    }
    h1 { font-size: 2.25rem !important; }
    h2 { font-size: 1.85rem !important; }
    h3 { font-size: 1.38rem !important; }
    h4 { font-size: 1.15rem !important; }
    div[data-testid="stMetricLabel"] {
        font-size: 1.02rem !important;
        font-weight: 700 !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: 800 !important;
    }
    div[data-testid="stMetricDelta"] {
        font-size: .95rem !important;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 14px !important;
        font-weight: 750 !important;
        padding-left: 8px !important;
        padding-right: 8px !important;
    }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {
        font-size: 14px !important;
    }
    .how-read-box {
        background: linear-gradient(180deg, #F2F7FF 0%, #F8FBFF 100%);
        border: 1px solid #DCEBFF;
        border-radius: 16px;
        padding: 18px 20px;
        min-height: 260px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    .how-read-title {
        font-size: 17px;
        font-weight: 850;
        color: #0B63CE;
        margin-bottom: 12px;
    }
    .how-read-box li {
        margin-bottom: 9px;
        line-height: 1.55;
        color: #172033;
        font-size: 14.5px;
    }
    .insight-side-card {
        border: 1px solid #E5EAF2;
        border-radius: 16px;
        padding: 18px;
        min-height: 170px;
        background: #FFFFFF;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04);
    }
    .insight-side-card-title {
        font-size: 16px;
        font-weight: 850;
        color: #F59E0B;
        margin-bottom: 10px;
    }
    .insight-side-card-text {
        font-size: 14px;
        line-height: 1.55;
        color: #172033;
    }
    </style>
    """, unsafe_allow_html=True)


def how_to_read_box(title, items):
    """Show chart-reading guidance only when the user opens it."""
    with st.expander(f"ⓘ {title}", expanded=False):
        for item in items:
            st.markdown(f"- {item}")


def side_insight_card(title, text):
    st.markdown(f"""
    <div class="insight-side-card">
        <div class="insight-side-card-title">💡 {title}</div>
        <div class="insight-side-card-text">{text}</div>
    </div>
    """, unsafe_allow_html=True)



def get_presentation_scale(mode: str) -> Dict[str, float]:
    """Responsive presentation scaling for phone, laptop, monitor, and boardroom display."""
    if mode == "Smartphone":
        return {
            "global_px": 14.5, "h1_rem": 1.70, "h2_rem": 1.45, "h3_rem": 1.22,
            "metric_label_rem": 0.92, "metric_value_rem": 1.65, "delta_rem": 0.86,
            "tab_px": 13, "sidebar_px": 13, "plot_font": 12, "plot_title": 15,
            "axis_title": 13, "axis_tick": 11, "legend": 11,
            "bubble_size": 50, "capacity_bubble": 42, "gauge_height": 230,
            "chart_height": 380, "how_read_px": 13,
            "container_padding": "0.65rem", "max_width": "100%", "mobile": 1,
        }
    if mode == "Compact":
        return {
            "global_px": 15.0, "h1_rem": 2.20, "h2_rem": 1.78, "h3_rem": 1.34,
            "metric_label_rem": 1.00, "metric_value_rem": 1.95, "delta_rem": 0.94,
            "tab_px": 14, "sidebar_px": 14, "plot_font": 13, "plot_title": 18,
            "axis_title": 15, "axis_tick": 13, "legend": 13,
            "bubble_size": 64, "capacity_bubble": 54, "gauge_height": 280,
            "chart_height": 470, "how_read_px": 14, "container_padding": "1.25rem", "max_width": "1540px", "mobile": 0,
        }
    if mode == "Boardroom":
        return {
            "global_px": 18.0, "h1_rem": 2.85, "h2_rem": 2.28, "h3_rem": 1.78,
            "metric_label_rem": 1.25, "metric_value_rem": 2.75, "delta_rem": 1.18,
            "tab_px": 17, "sidebar_px": 17, "plot_font": 16, "plot_title": 22,
            "axis_title": 19, "axis_tick": 16, "legend": 16,
            "bubble_size": 92, "capacity_bubble": 76, "gauge_height": 350,
            "chart_height": 560, "how_read_px": 16.5, "container_padding": "2.4rem", "max_width": "1560px", "mobile": 0,
        }
    return {
        "global_px": 16.5, "h1_rem": 2.38, "h2_rem": 1.98, "h3_rem": 1.52,
        "metric_label_rem": 1.10, "metric_value_rem": 2.18, "delta_rem": 1.03,
        "tab_px": 15, "sidebar_px": 15, "plot_font": 14, "plot_title": 20,
        "axis_title": 17, "axis_tick": 14, "legend": 14,
        "bubble_size": 78, "capacity_bubble": 64, "gauge_height": 310,
        "chart_height": 520, "how_read_px": 15.5, "container_padding": "2.4rem", "max_width": "1560px", "mobile": 0,
    }


def inject_presentation_mode_css(scale: Dict[str, float]):
    st.markdown(f"""
    <style>
    html, body, [class*="css"] {{
        font-size: {scale['global_px']}px !important;
    }}
    .main .block-container {{
        max-width: {scale.get('max_width', '1560px')};
        padding-top: 1.05rem;
        padding-left: {scale.get('container_padding', '2.4rem')};
        padding-right: {scale.get('container_padding', '2.4rem')};
    }}
    h1 {{ font-size: {scale['h1_rem']}rem !important; }}
    h2 {{ font-size: {scale['h2_rem']}rem !important; }}
    h3 {{ font-size: {scale['h3_rem']}rem !important; }}
    h4 {{ font-size: 1.25rem !important; }}
    div[data-testid="stMetricLabel"] {{
        font-size: {scale['metric_label_rem']}rem !important;
        font-weight: 750 !important;
    }}
    div[data-testid="stMetricValue"] {{
        font-size: {scale['metric_value_rem']}rem !important;
        font-weight: 850 !important;
        letter-spacing: -0.03em;
    }}
    div[data-testid="stMetricDelta"] {{
        font-size: {scale['delta_rem']}rem !important;
    }}
    .stTabs [data-baseweb="tab"] {{
        font-size: {scale['tab_px']}px !important;
        font-weight: 800 !important;
        padding-left: 9px !important;
        padding-right: 9px !important;
    }}
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span {{
        font-size: {scale['sidebar_px']}px !important;
    }}
    .how-read-box li {{
        font-size: {scale['how_read_px']}px !important;
        line-height: 1.58 !important;
    }}
    .how-read-title {{
        font-size: {scale['how_read_px'] + 2}px !important;
    }}
    .insight-side-card-title {{
        font-size: {scale['how_read_px'] + 1}px !important;
    }}
    .insight-side-card-text {{
        font-size: {scale['how_read_px']}px !important;
    }}
    .sep-feature-title {{
        font-size: {max(13, scale['tab_px'])}px !important;
    }}
    .sep-feature-text {{
        font-size: {max(12, scale['tab_px'] - 1)}px !important;
    }}
    .kpi-label {{
        font-size: {max(13, scale['tab_px'])}px !important;
    }}
    .kpi-value {{
        font-size: {scale['metric_value_rem']}rem !important;
    }}
    .kpi-foot {{
        font-size: {max(12, scale['tab_px'] - 2)}px !important;
    }}
    
    @media (max-width: 768px) {{
        .main .block-container {{
            padding-left: 0.65rem !important;
            padding-right: 0.65rem !important;
            padding-top: 0.6rem !important;
        }}
        h1, .sep-title {{
            font-size: 1.65rem !important;
            line-height: 1.12 !important;
        }}
        h2 {{ font-size: 1.42rem !important; }}
        h3 {{ font-size: 1.18rem !important; }}
        .sep-header {{ display: block !important; }}
        .sep-logo img, div[data-testid="stImage"] img {{
            max-height: 42px !important;
            width: auto !important;
        }}
        .sep-feature-band {{
            display: grid !important;
            grid-template-columns: 1fr !important;
            gap: 10px !important;
            padding: 12px !important;
            margin: 10px 0 12px 0 !important;
        }}
        .sep-feature-item {{ min-height: auto !important; }}
        .stTabs [data-baseweb="tab-list"] {{
            overflow-x: auto !important;
            flex-wrap: nowrap !important;
            white-space: nowrap !important;
        }}
        .stTabs [data-baseweb="tab"] {{
            font-size: 12.5px !important;
            min-width: max-content !important;
        }}
        div[data-testid="column"] {{
            width: 100% !important;
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }}
        div[data-testid="stMetricValue"] {{ font-size: 1.55rem !important; }}
        div[data-testid="stMetricLabel"] {{ font-size: 0.95rem !important; }}
        .how-read-box {{
            min-height: auto !important;
            padding: 14px !important;
        }}
        .how-read-box li {{ font-size: 13px !important; }}
        .insight-side-card {{
            min-height: auto !important;
            padding: 14px !important;
        }}
        .kpi-card, .status-card, .health-card {{
            min-height: auto !important;
            padding: 14px !important;
            margin-bottom: 10px !important;
        }}
        .kpi-value {{ font-size: 1.65rem !important; }}
        .memo-card {{
            display: block !important;
            padding: 14px !important;
        }}
        .memo-icon {{ margin-bottom: 8px !important; }}
        [data-testid="stSidebar"] {{ width: 86vw !important; }}
    }}
</style>
    """, unsafe_allow_html=True)


def apply_plotly_presentation(fig, scale: Dict[str, float], height_key: str = "chart_height"):
    fig.update_layout(
        height=int(scale.get(height_key, scale.get("chart_height", 520))),
        font=dict(size=int(scale.get("plot_font", 14))),
        title_font=dict(size=int(scale.get("plot_title", 20))),
        legend=dict(font=dict(size=int(scale.get("legend", 14)))),
        margin=dict(l=10, r=10, t=50, b=10),
    )
    fig.update_xaxes(
        title_font=dict(size=int(scale.get("axis_title", 17))),
        tickfont=dict(size=int(scale.get("axis_tick", 14))),
    )
    fig.update_yaxes(
        title_font=dict(size=int(scale.get("axis_title", 17))),
        tickfont=dict(size=int(scale.get("axis_tick", 14))),
    )
    return fig


# =========================
# SIDEBAR & LOAD
# =========================
with st.sidebar:
    st.header("⚙️ Data Source")
    uploaded = st.file_uploader("Upload Excel SEP V4", type=["xlsx"])
    st.caption("Jika tidak upload, app membaca file default dari folder data/.")

file_bytes = uploaded.getvalue() if uploaded is not None else None
sheets, warnings = load_workbook(file_bytes, DEFAULT_FILE)

inject_header_feature_css()
render_header()
render_feature_band()
for w in warnings:
    st.warning(w)
if any(s not in sheets for s in REQUIRED_SHEETS):
    st.error("Data belum lengkap. Sheet wajib: " + ", ".join(REQUIRED_SHEETS))
    st.stop()

# Calendar preview
_years = []
for s in ["REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR"]:
    if s in sheets and "Year" in sheets[s].columns:
        _years += pd.to_numeric(sheets[s]["Year"], errors="coerce").dropna().astype(int).tolist()
calendar_preview = build_calendar(sheets.get("CALENDAR", pd.DataFrame()), sorted(set(_years)))

with st.sidebar:
    st.header("🧮 Analysis Mode")
    default_idx = 1 if not calendar_preview.empty and (calendar_preview["Months_Closed"] < 12).any() else 0
    mode = st.radio("Mode", ["Actual", "YTD Comparable", "Annual Projection", "MCR Focus"], index=default_idx)
    st.header("🖥️ Presentation Mode")
    presentation_mode = st.radio(
        "Display Size",
        ["Smartphone", "Compact", "Executive", "Boardroom"],
        index=2,
        help="Smartphone untuk HP, Compact untuk laptop, Executive untuk monitor/meeting, Boardroom untuk TV/proyektor."
    )
    st.caption("Smartphone = HP | Compact = laptop | Executive = default meeting | Boardroom = TV/proyektor.")

data = prep_data(sheets, mode)
params = load_params(sheets.get("PARAMETERS", pd.DataFrame()))
targets = load_targets(sheets.get("TARGETS", pd.DataFrame()))
currency = params.get("Currency_Code", "IDR")

yr = data["yearly"].dropna(subset=["Year"]).copy()
available_years = sorted(yr["Year"].dropna().astype(int).unique().tolist())

with st.sidebar:
    st.header("🧭 Filters")
    year_sel = st.selectbox("Analysis Year", available_years, index=len(available_years)-1)
    year_min, year_max = st.slider("Trend Range", min_value=min(available_years), max_value=max(available_years), value=(min(available_years), max(available_years)), step=1)
    fja_all = sorted(data["fja"]["FJA_Category"].dropna().unique().tolist()) if not data["fja"].empty else []
    fja_filter = st.multiselect("FJA Category", fja_all, default=fja_all)

yr_range = yr[(yr["Year"] >= year_min) & (yr["Year"] <= year_max)].copy()
latest = yr[yr["Year"] == year_sel].iloc[0]
period = get_period_info(data["calendar"], year_sel)

bp = business_posture(latest, targets)
mq = margin_quality(latest.get("Gross_Margin_Pct"), targets)
mh = mcr_health(latest.get("MCR_Pct"), targets)
cr = capacity_risk(latest, data["capacity"])
bs = board_score(latest, targets, mq, mh, cr)
action_title, actions = management_action(bp, mq, mh, cr)

ai_presenter_context = _build_ai_context(latest, period, mode, bs, bp, mq, mh, cr, currency)
render_ai_presenter(ai_presenter_context)

plotly_display_config = {
    "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
}

presentation_scale = get_presentation_scale(presentation_mode)
inject_presentation_mode_css(presentation_scale)
if presentation_mode == "Smartphone":
    st.sidebar.info("Smartphone Mode aktif: layout single-column, tab horizontal-scroll, font dan chart disesuaikan untuk HP.")


# =========================
# CLEAN BOARD NAVIGATION
# =========================
boards = st.tabs([
    "1. Executive Summary",
    "2. Revenue & Productivity",
    "3. Workforce Performance",
    "4. Capacity & Delivery Risk",
    "5. Strategic Decision Center",
    "6. Executive Benchmark & Methodology",
    "7. CEO Closing Report",
    "8. Appendix",
])


# 1 Executive Summary
with boards[0]:
    st.markdown("## 1. Executive Summary")
    st.caption(f"Periode: {period['Period_Label']} | Months closed: {period['Months_Closed']}/12 | Mode: {mode} | Confidence: {period['Data_Confidence']}")
    st.info(data["mode_note"])
    if mode == "YTD Comparable" and data.get("monthly_ytd_used"):
        st.success("MONTHLY_KPI aktif: KPI dan board narrative memakai realisasi YTD aktual.")
    elif mode == "YTD Comparable":
        st.warning("MONTHLY_KPI belum aktif/terisi. YTD Comparable memakai proxy FY x bulan YTD.")

    a, b, c, d = st.columns(4)
    with a: st.markdown(badge("Company Health", f"{bs['icon']} {bs['score']}/100", bs["tone"], bs["label"]), unsafe_allow_html=True)
    with b: st.markdown(badge("Business Posture", bp["state"], bp["tone"], bp["hint"]), unsafe_allow_html=True)
    with c: st.markdown(badge("MCR Health", mh["state"], mh["tone"], mh["hint"]), unsafe_allow_html=True)
    with d: st.markdown(badge("Capacity Risk", cr["state"], cr["tone"], f"Score {cr['score']}/100"), unsafe_allow_html=True)

    k1, k2, k3 = st.columns(3)
    k4, k5, k6 = st.columns(3)
    with k1: st.metric("💰 Revenue", _money(latest.get("Total_Revenue"), currency), _pct(latest.get("Revenue_YoY")))
    with k2: st.metric("👥 Headcount", _num(latest.get("Total_Headcount")), _pct(latest.get("Headcount_YoY")))
    with k3: st.metric("⚡ RPE", _money(latest.get("RPE"), currency), _pct(latest.get("RPE_YoY")))
    with k4: st.metric("📊 MCR", _pct(latest.get("MCR_Pct")), _pp(latest.get("MCR_Delta")))
    with k5: st.metric("📈 Gross Margin", _pct(latest.get("Gross_Margin_Pct")), _pp(latest.get("GrossMargin_Delta")))
    with k6: st.metric("💼 Manpower Cost", _money(latest.get("Total_Manpower_Cost"), currency), _pct(latest.get("ManpowerCost_YoY")))

    st.markdown("#### Board Memo")
    st.info(
        f"Pada {period['Period_Label']} dalam mode **{mode}**, SEP membaca revenue **{_money(latest.get('Total_Revenue'), currency)}**, "
        f"RPE **{_money(latest.get('RPE'), currency)}**, MCR **{_pct(latest.get('MCR_Pct'))}**, dan Gross Margin **{_pct(latest.get('Gross_Margin_Pct'))}**. "
        f"Status utama adalah **{bs['label']}** dengan fokus manajemen: **{action_title}**."
    )
    render_board_guidance(
        "Executive Summary",
        [
            "Company Health adalah skor gabungan revenue growth, RPE growth, margin quality, MCR health, dan capacity risk.",
            "Business Posture menunjukkan posisi perusahaan: growth-ready, balanced, cost pressure, capacity risk, atau defensive.",
            "Revenue menunjukkan pendapatan yang diakui pada periode aktif; selalu baca bersama mode Actual/YTD/Projection.",
            "RPE adalah revenue per employee; kenaikan RPE berarti produktivitas per orang membaik.",
            "MCR adalah manpower cost dibagi revenue; nilai rendah baik hanya jika delivery capacity tetap aman.",
            "Gross Margin mengukur kualitas revenue setelah direct COGS; margin rendah berarti revenue belum tentu berkualitas.",
            "Capacity Risk harus dibaca bersama utilization, overtime, backlog, SLA breach, turnover, dan critical dependency."
        ],
        executive_summary_text(latest, bs, bp, mq, mh, cr, period, mode, currency),
        "Bagaimana kondisi kesehatan perusahaan saat ini, indikator apa yang paling membutuhkan perhatian Direksi, dan keputusan utama apa yang perlu diprioritaskan?",
        _summary_tone(bs, mq, mh, cr),
    )


# 2 Revenue & Productivity
with boards[1]:
    st.markdown("## 2. Revenue & Productivity")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(yr_range, x="Year", y=["Total_Revenue", "Total_Manpower_Cost"], markers=True, title="Revenue vs Manpower Cost")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.line(yr_range, x="Year", y=["RPE", "Cost_per_HC"], markers=True, title="RPE vs Cost per HC")
        st.plotly_chart(fig, use_container_width=True)
    c3, c4 = st.columns(2)
    with c3:
        growth = yr_range[["Year", "Revenue_YoY", "Headcount_YoY", "RPE_YoY", "ManpowerCost_YoY"]].melt(id_vars="Year", var_name="Metric", value_name="YoY")
        fig = px.line(growth, x="Year", y="YoY", color="Metric", markers=True, title="Growth Driver Comparison")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)
    with c4:
        fig = px.line(yr_range, x="Year", y=["Gross_Margin_Pct", "MCR_Pct"], markers=True, title="Gross Margin vs MCR")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Project / Service Line Productivity")
    st.caption("Membaca jenis project atau service line yang paling menyumbang revenue untuk perusahaan Telekomunikasi dan IT System Integrator.")
    pp = data.get("project_productivity", pd.DataFrame()).copy()
    pp_year = pp[pp["Year"] == year_sel].copy() if not pp.empty and "Year" in pp.columns else pd.DataFrame()
    if pp_year.empty:
        st.info("Sheet PROJECT_PRODUCTIVITY belum tersedia/terisi. Jika kosong, isi kategori project seperti Managed Service, Network Deployment/Installation, Maintenance, NOC, IT Infrastructure, CCTV/IoT, Cloud/Data Center, dan Professional Service.")
    else:
        total_pp_rev = pp_year["Revenue"].sum()
        pp_cat = pp_year.groupby("Project_Category", as_index=False).agg(
            Revenue=("Revenue", "sum"),
            Gross_Profit=("Gross_Profit", "sum"),
            Manpower_Cost=("Manpower_Cost", "sum"),
            Project_Hours=("Project_Hours", "sum"),
            HC_Allocated=("HC_Allocated", "sum"),
            Sites_or_Tickets=("Sites_or_Tickets", "sum"),
            SLA_Breach_Count=("SLA_Breach_Count", "sum"),
        )
        pp_cat["Revenue_Share"] = np.where(total_pp_rev > 0, pp_cat["Revenue"] / total_pp_rev, np.nan)
        pp_cat["Gross_Margin_Pct"] = np.where(pp_cat["Revenue"] > 0, pp_cat["Gross_Profit"] / pp_cat["Revenue"], np.nan)
        pp_cat["MCR_Pct"] = np.where(pp_cat["Revenue"] > 0, pp_cat["Manpower_Cost"] / pp_cat["Revenue"], np.nan)
        pp_cat["Revenue_per_HC"] = np.where(pp_cat["HC_Allocated"] > 0, pp_cat["Revenue"] / pp_cat["HC_Allocated"], np.nan)

        p1, p2 = st.columns([1.0, 1.0])
        with p1:
            fig = px.treemap(
                pp_cat,
                path=["Project_Category"],
                values="Revenue",
                color="Gross_Margin_Pct",
                color_continuous_scale="RdYlGn",
                title="Revenue Contribution by Project / Service Line",
                hover_data=["Revenue_Share", "Gross_Margin_Pct", "MCR_Pct", "Revenue_per_HC"],
            )
            fig.update_layout(height=430, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig, use_container_width=True)
        with p2:
            bubble = pp_cat.copy()
            bubble["Bubble_Size"] = bubble["Revenue"].clip(lower=0)
            if bubble["Bubble_Size"].max() <= 0:
                bubble["Bubble_Size"] = bubble["HC_Allocated"].clip(lower=1)
            fig = px.scatter(
                bubble,
                x="MCR_Pct",
                y="Gross_Margin_Pct",
                size="Bubble_Size",
                color="Project_Category",
                text="Project_Category",
                title="Project Productivity Matrix: MCR vs Gross Margin",
                hover_data=["Revenue", "Revenue_Share", "Revenue_per_HC", "Sites_or_Tickets", "SLA_Breach_Count"],
                size_max=55,
            )
            fig.update_xaxes(tickformat=".0%", title="MCR")
            fig.update_yaxes(tickformat=".0%", title="Gross Margin")
            fig.add_vline(x=get_target(targets, "MCR_Healthy_Max", 0.07), line_dash="dash", annotation_text="MCR Healthy Max")
            fig.add_hline(y=get_target(targets, "Gross_Margin_Healthy_Min", 0.20), line_dash="dash", annotation_text="GM Healthy Min")
            fig.update_traces(textposition="top center")
            fig.update_layout(height=430, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig, use_container_width=True)

        pp_show = pp_cat.sort_values("Revenue", ascending=False).copy()
        pp_show["Revenue"] = pp_show["Revenue"].map(lambda x: _money(x, currency))
        pp_show["Revenue_Share"] = pp_show["Revenue_Share"].map(_pct)
        pp_show["Gross_Margin_Pct"] = pp_show["Gross_Margin_Pct"].map(_pct)
        pp_show["MCR_Pct"] = pp_show["MCR_Pct"].map(_pct)
        pp_show["Revenue_per_HC"] = pp_show["Revenue_per_HC"].map(lambda x: _money(x, currency))
        st.dataframe(pp_show, hide_index=True, use_container_width=True)

        top_project = pp_cat.sort_values("Revenue", ascending=False).iloc[0]
        st.info(
            f"Kontributor revenue terbesar adalah **{top_project['Project_Category']}** dengan share **{_pct(top_project['Revenue_Share'])}**. "
            f"Gross Margin kategori ini **{_pct(top_project['Gross_Margin_Pct'])}** dan MCR **{_pct(top_project['MCR_Pct'])}**."
        )

    with st.expander("How to Read Project / Service Line Productivity"):
        st.markdown("""
        - **Revenue Contribution** menunjukkan service line mana yang paling besar menyumbang revenue.
        - **Project Productivity Matrix** membaca dua hal sekaligus: MCR rendah dan Gross Margin tinggi adalah kombinasi terbaik.
        - Kategori umum untuk Telekomunikasi & IT System Integrator: Managed Service, Network Deployment/Installation, Network Maintenance, NOC/Monitoring, IT Infrastructure, CCTV/IoT/Security, Cloud/Data Center, Professional Service, dan SLA Support.
        - Jika revenue besar tetapi Gross Margin rendah, cek pricing, COGS, scope creep, subcontractor, dan rework.
        - Jika Gross Margin sehat tetapi MCR tinggi, cek alokasi manpower, overtime, dan produktivitas tim project.
        """)

    with st.expander("How to Read Revenue & Productivity"):
        st.markdown("""
        - Revenue harus tumbuh lebih cepat daripada manpower cost.
        - RPE naik berarti produktivitas per karyawan membaik.
        - Gross Margin menunjukkan kualitas revenue setelah direct cost.
        - MCR menunjukkan apakah biaya manpower masih seimbang terhadap revenue.
        """)

    render_board_guidance(
        "Revenue & Productivity",
        [
            "Revenue vs Manpower Cost: garis revenue idealnya tumbuh lebih cepat daripada people cost.",
            "RPE vs Cost per HC: produktivitas membaik jika RPE naik lebih cepat daripada biaya rata-rata per karyawan.",
            "Revenue Growth menunjukkan momentum; Headcount Growth menunjukkan ekspansi kapasitas; keduanya harus dibaca bersamaan.",
            "Gross Margin adalah (Revenue - COGS) / Revenue dan menunjukkan kualitas ekonomi project/service.",
            "MCR adalah Manpower Cost / Revenue; MCR tinggi menunjukkan people cost menekan revenue.",
            "Revenue Contribution menunjukkan service line terbesar, tetapi kontribusi besar belum tentu paling menguntungkan.",
            "Project Productivity terbaik berada pada kombinasi Gross Margin tinggi, MCR rendah, SLA baik, dan revenue per HC tinggi."
        ],
        revenue_summary_text(latest, mq, mh, currency),
        "Apakah revenue tumbuh bersama produktivitas dan margin, atau pertumbuhan masih bergantung pada tambahan headcount dan manpower cost?",
        _summary_tone(bs, mq, mh, cr),
    )


# 3 Workforce
with boards[2]:
    st.markdown("## 3. Workforce Performance")
    fja = data["fja"].copy()
    fja = fja[(fja["Year"] >= year_min) & (fja["Year"] <= year_max)]
    if fja_filter:
        fja = fja[fja["FJA_Category"].isin(fja_filter)]
    fja_latest = fja[fja["Year"] == year_sel].copy()
    w1, w2 = st.columns([1.1, 1.0])
    with w1:
        if not fja_latest.empty:
            fig = px.treemap(fja_latest, path=["FJA_Category"], values="Manpower_Cost", title="FJA Cost Mix")
            st.plotly_chart(fig, use_container_width=True)
    with w2:
        if not fja_latest.empty:
            show = fja_latest[["FJA_Category", "Manpower_Cost", "Cost_Share", "Headcount", "HC_Share"]].copy()
            show["Manpower_Cost"] = show["Manpower_Cost"].map(lambda x: _money(x, currency))
            show["Cost_Share"] = show["Cost_Share"].map(_pct)
            show["HC_Share"] = show["HC_Share"].map(_pct)
            st.dataframe(show.sort_values("Cost_Share", ascending=False), hide_index=True, use_container_width=True)
    dept = data["dept"].copy()
    dept = dept[dept["Year"] == year_sel].copy()
    if fja_filter and not dept.empty:
        dept = dept[dept["FJA_Category"].isin(fja_filter)]
    if not dept.empty:
        st.markdown("### Workforce Portfolio Matrix")
        wh1, wh2, wh3 = st.columns([0.78, 1.9, 0.62])
        dept["Bubble_Size"] = pd.to_numeric(dept.get("Dept_Revenue"), errors="coerce").fillna(0)
        if dept["Bubble_Size"].max() <= 0:
            dept["Bubble_Size"] = pd.to_numeric(dept.get("Headcount"), errors="coerce").fillna(1).clip(lower=1)
            size_note = "Ukuran gelembung memakai Headcount karena revenue by department belum tersedia."
        else:
            min_pos = dept.loc[dept["Bubble_Size"] > 0, "Bubble_Size"].min()
            dept["Bubble_Size"] = dept["Bubble_Size"].replace(0, min_pos * 0.35)
            size_note = "Ukuran gelembung memakai Dept Revenue; nilai kecil tetap diperbesar agar terlihat saat presentasi."
        fja_colors = {
            "Revenue Generator": "#2563EB",
            "Revenue Enabler": "#16A34A",
            "Support Function": "#EF4444",
            "Governance / Management": "#8B5CF6",
            "Unmapped": "#64748B",
        }
        with wh1:
            how_to_read_box("How to Read", [
                "Setiap gelembung mewakili Department/FJA Category.",
                "Sumbu X = Headcount; semakin kanan berarti jumlah orang lebih besar.",
                "Sumbu Y = Manpower Cost; semakin atas berarti biaya manpower lebih besar.",
                "Ukuran gelembung = revenue atau headcount, sehingga kontribusi kecil tetap terlihat.",
                "Warna merah perlu perhatian karena biasanya support/cost center.",
                "Kanan-atas = fungsi besar dan mahal; harus punya output produktif.",
                "Kiri-bawah = fungsi kecil; evaluasi apakah tetap critical atau bisa disederhanakan."
            ])
        with wh2:
            fig = px.scatter(
                dept,
                x="Headcount",
                y="Manpower_Cost",
                size="Bubble_Size",
                color="FJA_Category",
                text="FJA_Category",
                color_discrete_map=fja_colors,
                hover_name="Dept_Name",
                hover_data=["Dept_Revenue", "Dept_RPE", "Revenue_per_Cost", "Headcount", "Manpower_Cost"],
                title="Workforce Portfolio Matrix",
                size_max=int(presentation_scale.get('bubble_size', 78)),
            )
            fig.update_traces(
                textposition="middle center",
                textfont=dict(size=13, color="white"),
                marker=dict(opacity=0.92, line=dict(width=2.5, color="white")),
            )
            fig.update_layout(
                height=520,
                font=dict(size=int(presentation_scale.get('plot_font', 14))),
                margin=dict(l=10, r=10, t=45, b=10),
                legend=dict(font=dict(size=int(presentation_scale.get('legend', 14)))),
            )
            st.plotly_chart(fig, use_container_width=True, config=plotly_display_config)
            st.caption(size_note)
        with wh3:
            side_insight_card(
                "Insight",
                "Fokus menjaga produktivitas fungsi support agar biaya tidak membebani MCR. Fungsi besar dan mahal harus dikaitkan dengan revenue, delivery, atau capacity protection."
            )

    render_board_guidance(
        "Workforce Performance",
        [
            "FJA Cost Mix menunjukkan distribusi manpower cost antara Revenue Generator, Revenue Enabler, Support, dan Governance.",
            "Cost Share adalah porsi biaya suatu fungsi terhadap total manpower cost; HC Share adalah porsi jumlah orang.",
            "Workforce Portfolio Matrix: sumbu X = headcount, sumbu Y = manpower cost, ukuran bubble = revenue atau headcount.",
            "Kanan-atas menunjukkan fungsi besar dan mahal; harus memiliki output, revenue, delivery, atau risk-protection yang jelas.",
            "Revenue per Cost mengukur berapa rupiah revenue yang didukung oleh setiap rupiah manpower cost.",
            "Dept RPE adalah revenue departemen per headcount; gunakan hanya jika revenue by department tersedia dan valid.",
            "Support Function tidak otomatis buruk, tetapi cost share-nya perlu seimbang dengan kebutuhan governance dan operasi."
        ],
        workforce_summary_text(latest, dept, fja_latest, currency),
        "Apakah struktur dan biaya tenaga kerja sudah seimbang dengan kontribusi revenue, delivery, dan kebutuhan fungsi pendukung?",
        "warning" if mh.get("state") in ["WATCH", "COST_PRESSURE", "CRITICAL"] else "info",
    )


# 4 Capacity
with boards[3]:
    st.markdown("## 4. Capacity & Delivery Risk")
    cap_year = data["capacity"][data["capacity"]["Year"] == year_sel].copy() if not data["capacity"].empty else pd.DataFrame()
    cc0, cc1, cc2 = st.columns([0.72, 1.45, 0.78])
    with cc0:
        how_to_read_box("How to Read", [
            "Setiap gelembung mewakili department/fungsi operasional.",
            "Sumbu X = Avg Utilization; semakin kanan semakin padat kapasitas.",
            "Sumbu Y = Overtime Hours; semakin atas semakin tinggi beban lembur.",
            "Ukuran gelembung = Backlog; backlog kecil tetap diperbesar agar terlihat.",
            "Merah = risiko tinggi/prioritas perhatian.",
            "Kanan-atas = risiko overload tertinggi.",
            "Jika utilization tinggi + overtime tinggi + backlog besar, perlu capacity action."
        ])
    with cc1:
        if not cap_year.empty:
            cap_plot = cap_year.copy()
            cap_plot["Backlog_Visual_Size"] = pd.to_numeric(cap_plot.get("Backlog_Count"), errors="coerce").fillna(0)
            cap_plot["Backlog_Visual_Size"] = cap_plot["Backlog_Visual_Size"].clip(lower=8)
            dep_colors = {
                "Low": "#EF4444",
                "Medium": "#2563EB",
                "High": "#16A34A",
                "Critical": "#B91C1C",
            }
            fig = px.scatter(
                cap_plot,
                x="Avg_Utilization_Pct",
                y="Overtime_Hours",
                size="Backlog_Visual_Size",
                color="Critical_Role_Dependency",
                color_discrete_map=dep_colors,
                hover_data=["Dept_ID", "Backlog_Count", "SLA_Breach_Count", "Incident_Count", "Turnover_Count"],
                title="Capacity Risk Heatmap",
                size_max=int(presentation_scale.get('capacity_bubble', 64)),
            )
            fig.update_traces(marker=dict(opacity=0.90, line=dict(width=2.5, color="white")))
            fig.update_xaxes(tickformat=".0%")
            fig.update_layout(
                height=int(presentation_scale.get('chart_height', 520)),
                font=dict(size=int(presentation_scale.get('plot_font', 14))),
                margin=dict(l=10, r=10, t=45, b=10),
                legend=dict(font=dict(size=int(presentation_scale.get('legend', 14)))),
            )
            st.plotly_chart(fig, use_container_width=True, config=plotly_display_config)
        else:
            st.info("CAPACITY_INDICATOR belum diisi.")
    with cc2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=cr["score"],
            number={"suffix": "/100", "font": {"size": 38}},
            title={"text": f"Capacity Score - {cr['state']}"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": TONE_COLOR.get(cr["tone"], "#3b82f6")},
                "steps": [
                    {"range": [0, 30], "color": "#22C55E"},
                    {"range": [31, 55], "color": "#FACC15"},
                    {"range": [56, 75], "color": "#F59E0B"},
                    {"range": [76, 100], "color": "#EF4444"},
                ],
            }
        ))
        fig.update_layout(height=int(presentation_scale.get('gauge_height', 310)), margin=dict(l=10, r=10, t=45, b=10), font=dict(size=int(presentation_scale.get('plot_font', 14))))
        st.plotly_chart(fig, use_container_width=True, config=plotly_display_config)
        d = cr["details"]
        st.markdown(f"""
        <div style="border:1px solid #E5EAF2;border-radius:14px;padding:14px;background:#FFFFFF;font-size:14px;line-height:1.9;">
            <b>Max Utilization:</b> <span style="color:#EF4444;font-weight:800;">{_pct(d.get('max_utilization'))}</span><br>
            <b>Overtime:</b> <span style="color:#EF4444;font-weight:800;">{_num(d.get('total_overtime'))} jam</span><br>
            <b>Backlog:</b> <span style="color:#EF4444;font-weight:800;">{_num(d.get('total_backlog'))}</span><br>
            <b>SLA:</b> <span style="color:#EF4444;font-weight:800;">{_num(d.get('total_sla_breach'))}</span>
        </div>
        <div style="border:1px solid #FED7AA;border-radius:14px;padding:14px;margin-top:12px;background:#FFF7ED;font-size:14px;line-height:1.7;">
            <b>Recommendation</b><br>
            <span style="color:#DC2626;font-weight:850;">{cr['recommendation']}</span>
        </div>
        """, unsafe_allow_html=True)

    render_board_guidance(
        "Capacity & Delivery Risk",
        [
            "Avg Utilization adalah persentase kapasitas kerja yang terpakai; 75-85% umumnya optimal, sedangkan >90% berisiko overload.",
            "Overtime menunjukkan beban tambahan di luar jam kerja normal; nilai tinggi yang berulang dapat menandakan kekurangan kapasitas atau proses tidak efisien.",
            "Backlog adalah pekerjaan yang belum terselesaikan; backlog besar meningkatkan risiko keterlambatan dan tekanan delivery.",
            "SLA Breach menunjukkan kegagalan memenuhi komitmen layanan; bahkan jumlah kecil perlu ditelusuri penyebabnya.",
            "Critical Role Dependency mengukur ketergantungan pada orang/keahlian tertentu; dependency tinggi meningkatkan single-point-of-failure risk.",
            "Capacity Score adalah skor gabungan, bukan sekadar utilization; skor dapat tinggi karena overtime, backlog, SLA, turnover, atau dependency.",
            "Hiring bukan satu-satunya solusi; pertimbangkan redistribusi workload, automation, SOP, vendor, dan skill development."
        ],
        capacity_summary_text(cr),
        "Apakah organisasi mulai overload, departemen mana yang paling berisiko, dan apakah solusinya hiring atau perbaikan proses serta distribusi workload?",
        "error" if cr.get("state") == "CRITICAL" else "warning" if cr.get("state") in ["MEDIUM", "HIGH"] else "success",
    )


# 5 Strategic Decision
with boards[4]:
    st.markdown("## 5. Strategic Decision Center")
    s1, s2 = st.columns([1.0, 1.0])
    with s1:
        priority = make_priority_df(bp, mq, mh, cr)
        fig = px.bar(priority.sort_values("Priority"), x="Priority", y="Area", color="Status", orientation="h", text="Status", title="Management Attention Map")
        fig.update_xaxes(range=[0, 100])
        st.plotly_chart(fig, use_container_width=True)
    with s2:
        comps = pd.DataFrame({"Factor": list(bs["components"].keys()), "Score": list(bs["components"].values())})
        fig = px.bar(comps.sort_values("Score"), x="Score", y="Factor", orientation="h", text="Score", title="Company Health Components")
        fig.update_xaxes(range=[0, 100])
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Management Insight")
    for title, status, why, meaning in build_management_insights(latest, bp, mq, mh, cr):
        if status in ["GOOD", "BALANCED"]:
            st.success(f"**{title}** — {why}\n\n{meaning}")
        elif status in ["CRITICAL"]:
            st.error(f"**{title}** — {why}\n\n{meaning}")
        else:
            st.warning(f"**{title}** — {why}\n\n{meaning}")

    hire_df = hiring_decision(data["dept"], data["capacity"], year_sel)
    if not hire_df.empty:
        st.markdown("### Hiring Decision")
        rec_summary = hire_df["Recommendation"].value_counts().reset_index()
        rec_summary.columns = ["Recommendation", "Count"]
        h1, h2 = st.columns([0.7, 1.3])
        with h1:
            fig = px.pie(rec_summary, names="Recommendation", values="Count", title="Hiring Mix")
            st.plotly_chart(fig, use_container_width=True)
        with h2:
            show = hire_df[["Dept_ID", "Dept_Name", "FJA_Category", "Headcount", "Manpower_Cost", "Dept_Revenue", "Revenue_per_Cost", "Avg_Utilization", "Recommendation", "Reason"]].copy()
            st.dataframe(show, hide_index=True, use_container_width=True)

    st.markdown("### 90-Day Action Plan")
    for i, a in enumerate(actions[:5], start=1):
        st.write(f"{i}. {a}")

    st.markdown("### Board Projection Simulator")
    st.caption("Simulasi penambahan HC dan target pertumbuhan revenue untuk melihat dampak terhadap MCR, Gross Margin, dan RPE.")

    base_projection = build_projection_scenarios(latest, targets, currency)
    sim_a, sim_b, sim_c, sim_d = st.columns(4)
    with sim_a:
        add_hc_sim = st.slider("Additional HC", 0, 100, 5, step=1, key="board_sim_add_hc")
    with sim_b:
        revenue_growth_sim = st.slider("Revenue Growth Target", -0.30, 1.50, 0.20, step=0.01, key="board_sim_rev_growth")
    with sim_c:
        gross_margin_sim = st.slider("Gross Margin Target", 0.00, 0.80, float(base_projection["gm_base"]) if pd.notna(base_projection["gm_base"]) else 0.25, step=0.01, key="board_sim_gm")
    with sim_d:
        avg_cost_new_hc = st.number_input(
            "Avg Cost / New HC",
            min_value=0.0,
            value=float(base_projection["cost_per_hc"]) if pd.notna(base_projection["cost_per_hc"]) and base_projection["cost_per_hc"] > 0 else 75_000_000.0,
            step=5_000_000.0,
            key="board_sim_cost_hc"
        )

    scenario_revenue = latest.get("Total_Revenue") * (1 + revenue_growth_sim)
    scenario_cogs = scenario_revenue * (1 - gross_margin_sim)
    scenario_manpower = latest.get("Total_Manpower_Cost") + (add_hc_sim * avg_cost_new_hc)
    scenario_hc = latest.get("Total_Headcount") + add_hc_sim
    scenario_rpe = scenario_revenue / scenario_hc if scenario_hc else np.nan
    scenario_mcr = scenario_manpower / scenario_revenue if scenario_revenue else np.nan
    required_rev_at_target_mcr = required_revenue_for_hc(latest, base_projection["target_mcr"], avg_cost_new_hc, add_hc_sim)
    revenue_gap = required_rev_at_target_mcr - scenario_revenue if pd.notna(required_rev_at_target_mcr) and pd.notna(scenario_revenue) else np.nan

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        st.metric("Projected Revenue", _money(scenario_revenue, currency), _pct(revenue_growth_sim))
    with s2:
        st.metric("Projected HC", _num(scenario_hc), f"+{add_hc_sim}")
    with s3:
        st.metric("Projected MCR", _pct(scenario_mcr), _pp(scenario_mcr - latest.get("MCR_Pct")))
    with s4:
        st.metric("Projected Gross Margin", _pct(gross_margin_sim), _pp(gross_margin_sim - latest.get("Gross_Margin_Pct")))

    s5, s6, s7 = st.columns(3)
    with s5:
        st.metric("Projected RPE", _money(scenario_rpe, currency), _pct((scenario_rpe / latest.get("RPE") - 1) if latest.get("RPE") else np.nan))
    with s6:
        st.metric("Revenue Required at Target MCR", _money(required_rev_at_target_mcr, currency))
    with s7:
        st.metric("Revenue Gap / Buffer", _money(revenue_gap, currency))

    grid = scenario_grid(latest, targets, revenue_growth_sim, gross_margin_sim, avg_cost_new_hc, max_hc_add=50)

    g1, g2 = st.columns(2)
    with g1:
        fig = px.line(
            grid,
            x="Additional_HC",
            y="Projected_MCR",
            markers=True,
            title="MCR Sensitivity by Additional HC",
            labels={"Projected_MCR": "Projected MCR", "Additional_HC": "Additional HC"},
        )
        fig.add_hline(y=base_projection["target_mcr"], line_dash="dash", annotation_text="Target MCR Healthy")
        fig.add_hline(y=get_target(targets, "MCR_Watch_Max", 0.09), line_dash="dot", annotation_text="MCR Watch")
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(height=int(presentation_scale.get('chart_height', 520)), margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        fig = px.line(
            grid,
            x="Additional_HC",
            y="Projected_RPE",
            markers=True,
            title="RPE Sensitivity by Additional HC",
            labels={"Projected_RPE": "Projected RPE", "Additional_HC": "Additional HC"},
        )
        fig.update_layout(height=int(presentation_scale.get('chart_height', 520)), margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig, use_container_width=True)

    if pd.notna(scenario_mcr):
        if scenario_mcr <= get_target(targets, "MCR_Healthy_Max", 0.07) and gross_margin_sim >= get_target(targets, "Gross_Margin_Healthy_Min", 0.20):
            st.success("Scenario layak: MCR masih dalam target sehat dan Gross Margin memenuhi benchmark.")
        elif scenario_mcr <= get_target(targets, "MCR_Watch_Max", 0.09):
            st.warning("Scenario masih bisa dipertimbangkan, tetapi masuk zona watch. Direksi perlu memastikan revenue pipeline dan margin benar-benar tercapai.")
        else:
            st.error("Scenario berisiko: tambahan HC membuat MCR melewati zona sehat/watch. Tambahan HC perlu ditahan atau revenue target dinaikkan.")

    with st.expander("How to Read Board Projection Simulator"):
        st.markdown("""
        - **Additional HC** mensimulasikan penambahan karyawan/personel.
        - **Revenue Growth Target** menunjukkan target pertumbuhan revenue yang ingin dicapai Direksi.
        - **Gross Margin Target** menunjukkan kualitas revenue yang ditargetkan.
        - **Projected MCR** adalah dampak penambahan HC terhadap biaya manpower dibanding revenue.
        - **Revenue Required at Target MCR** menjawab: berapa revenue minimum agar tambahan HC masih aman secara MCR.
        - Jika MCR naik melewati target, maka opsi manajemen adalah menaikkan revenue target, menurunkan avg cost per HC, atau menunda hiring.
        """)

    render_board_guidance(
        "Strategic Decision Center",
        [
            "Management Attention Map mengurutkan area berdasarkan tingkat urgensi, bukan besar kecilnya nominal.",
            "Company Health Components menunjukkan faktor pembentuk skor kesehatan perusahaan dan area yang menarik skor turun.",
            "Hiring Decision: HIRE untuk bottleneck terbukti; SELECTIVE untuk tekanan yang perlu validasi; HOLD/FREEZE bila business case belum kuat.",
            "Projected MCR menunjukkan dampak penambahan HC terhadap people-cost ratio.",
            "Revenue Required at Target MCR adalah revenue minimum agar manpower cost tetap berada pada guardrail sehat.",
            "Revenue Gap/Buffer positif berarti revenue masih kurang terhadap kebutuhan; negatif berarti terdapat buffer.",
            "Simulator adalah alat skenario, bukan forecast otomatis; asumsi revenue growth, GM, dan cost per HC harus disetujui manajemen."
        ],
        strategic_summary_text(bs, bp, mq, mh, cr, action_title),
        "Keputusan apa yang paling tepat untuk menjaga pertumbuhan, margin, dan kapasitas tanpa meningkatkan MCR secara tidak sehat?",
        _summary_tone(bs, mq, mh, cr),
    )


# 6 Benchmark & Methodology
with boards[5]:
    st.markdown("## 6. KPI Target Framework")
    st.markdown("### Based on Telecommunications & IT System Integrator Benchmark")

    framework = framework_from_targets(sheets.get("TARGETS", pd.DataFrame()), targets)
    st.dataframe(framework, hide_index=True, use_container_width=True)

    st.markdown("### Current KPI vs Framework")
    compare = pd.DataFrame([
        {"KPI": "Revenue Growth", "Company": _pct(latest.get("Revenue_YoY")), "Framework": "15–25% healthy; >25% strong", "Status": "Strong/Watch based on trend"},
        {"KPI": "RPE Growth", "Company": _pct(latest.get("RPE_YoY")), "Framework": ">=15% strong", "Status": "Productivity growth"},
        {"KPI": "MCR", "Company": _pct(latest.get("MCR_Pct")), "Framework": "5–7% healthy; 7–9% watch", "Status": mh["state"]},
        {"KPI": "Gross Margin", "Company": _pct(latest.get("Gross_Margin_Pct")), "Framework": "20–30% healthy; >30% strong", "Status": mq["state"]},
        {"KPI": "Capacity Score", "Company": f"{cr['score']}/100", "Framework": "0–30 low; 31–55 medium; 56–75 high; 76–100 critical", "Status": cr["state"]},
    ])
    st.dataframe(compare, hide_index=True, use_container_width=True)

    st.markdown("### Methodology")
    st.info(
        "Framework ini adalah management benchmark, bukan standar regulasi. Target disusun dari kombinasi praktik industri telekomunikasi dan IT System Integrator, analisis laporan tahunan perusahaan sejenis, praktik workforce planning, serta historical performance perusahaan. Review disarankan minimal triwulanan."
    )
    st.markdown("### Benchmark Sources to Track")
    st.write("Annual reports perusahaan telco/IT integrator, APQC, Saratoga/PwC Human Capital Benchmark, Gartner, Deloitte, McKinsey, BCG, dan benchmark internal historis perusahaan.")

    render_board_guidance(
        "Executive Benchmark & Methodology",
        [
            "Target Framework adalah guardrail manajemen yang menggabungkan historical performance dan praktik industri.",
            "Company menunjukkan nilai aktual; Framework menunjukkan rentang pembanding; Status menunjukkan posisi terhadap guardrail.",
            "Revenue Growth dan RPE Growth mengukur momentum dan produktivitas, sedangkan MCR dan Gross Margin mengukur kualitas ekonomi.",
            "Capacity Score mengukur delivery sustainability; nilai rendah berarti risiko rendah, bukan performa rendah.",
            "Confidence menunjukkan kekuatan sumber benchmark; benchmark eksternal harus disesuaikan dengan model bisnis Starcom.",
            "Benchmark bukan standar regulasi dan tidak boleh digunakan sebagai satu-satunya dasar keputusan.",
            "Threshold perlu direview minimal triwulanan atau saat model bisnis, pricing, dan struktur biaya berubah material."
        ],
        benchmark_summary_text(latest, mq, mh, cr),
        "Apakah target KPI dan threshold yang digunakan masih realistis, relevan, dan dapat dipertanggungjawabkan untuk model bisnis Starcom?",
        _summary_tone(bs, mq, mh, cr),
    )


# 7 CEO Closing
with boards[6]:
    st.markdown("## 7. CEO Closing Report")
    x1, x2 = st.columns([0.9, 1.1])
    with x1:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=bs["score"],
            number={"suffix": "/100"},
            title={"text": f"Company Health - {bs['label']}"},
            gauge={"axis": {"range": [0, 100]}, "bar": {"color": TONE_COLOR.get(bs["tone"], "#3b82f6")}},
        ))
        fig.update_layout(height=int(presentation_scale.get('chart_height', 520)))
        st.plotly_chart(fig, use_container_width=True)
    with x2:
        st.info(closing_text(latest, bs, bp, mq, mh, cr, targets, currency))
        st.markdown("#### Top Decisions")
        roadmap_df = roadmap(bp, mq, mh, cr)
        st.dataframe(roadmap_df, hide_index=True, use_container_width=True)
    st.success("Board-ready statement: SEP V4 merekomendasikan ekspansi HC secara selektif, bukan general hiring. Prioritas utama adalah menjaga MCR, meningkatkan RPE, dan menurunkan capacity risk pada fungsi bottleneck.")

    render_board_guidance(
        "CEO Closing Report",
        [
            "Gauge Company Health merangkum kondisi perusahaan dalam satu skor 0-100.",
            "Closing narrative menghubungkan revenue, MCR, RPE, gross margin, dan capacity risk dalam satu kesimpulan.",
            "Top Decisions menunjukkan tindakan menurut horizon waktu dan prioritas.",
            "Success Metric adalah indikator keberhasilan tindakan, bukan hanya daftar aktivitas.",
            "Board-ready statement harus dibaca sebagai arahan umum; keputusan rinci tetap memerlukan owner, deadline, dan evidence.",
            "Fokus utama adalah tindakan yang mengurangi risiko sekaligus menjaga revenue dan delivery.",
            "Setiap keputusan perlu dibawa kembali ke dashboard pada review berikutnya untuk melihat dampaknya."
        ],
        closing_text(latest, bs, bp, mq, mh, cr, targets, currency),
        "Apa kesimpulan akhir yang harus dibawa Direksi, risiko utama apa yang harus dikendalikan, dan tindakan apa yang harus segera dijalankan?",
        _summary_tone(bs, mq, mh, cr),
    )


# 8 Appendix
with boards[7]:
    st.markdown("## 8. Appendix")
    st.markdown("""
    ### Core Formula
    - **MCR** = Total Manpower Cost / Revenue
    - **RPE** = Revenue / Average Headcount
    - **Gross Margin** = (Revenue - COGS) / Revenue
    - **Cost per HC** = Total Manpower Cost / Average Headcount
    - **Revenue per Payroll** = Revenue / Total Manpower Cost
    - **Projected MCR** = (Current Manpower Cost + Additional HC × Avg Cost per HC) / Projected Revenue
    - **Projected Revenue Required** = Projected Manpower Cost / Target MCR

    ### Governance
    - **Owner:** Finance + HR + Operations
    - **Review cycle:** Quarterly
    - **Approval:** CEO / Board
    - **Framework:** KPI Target Framework based on telco & IT system integrator benchmark
    - **Version:** SEP V4 Executive Intelligence Edition

    ### Data Source
    - PARAMETERS, TARGETS, CALENDAR
    - REVENUE_YR / MONTHLY_KPI
    - HEADCOUNT_YR / PAYROLL_YR
    - DIM_DEPARTMENT / FJA_MAPPING
    - CAPACITY_INDICATOR
    - PROJECT_MARGIN
    - PROJECT_PRODUCTIVITY: service line/project category revenue contribution
    """)
    st.markdown("### Raw Data Preview")
    st.dataframe(yr, hide_index=True, use_container_width=True)

    render_board_guidance(
        "Appendix & Data Governance",
        [
            "Formula menjelaskan hubungan matematis setiap KPI; gunakan untuk audit dan rekonsiliasi.",
            "Data Source menunjukkan sheet Excel yang menjadi sumber indikator.",
            "Owner bertanggung jawab atas kualitas dan ketepatan waktu data, bukan hanya pengisian file.",
            "Review cycle menentukan frekuensi validasi; perubahan material perlu dicatat sebagai version update.",
            "Actual, YTD Comparable, Annual Projection, dan MCR Focus menghasilkan interpretasi berbeda; selalu sebutkan mode saat presentasi.",
            "Raw Data Preview digunakan untuk pengecekan, bukan sebagai tampilan utama Direksi.",
            "Jika data kosong atau confidence rendah, dashboard harus menyatakan keterbatasan dan tidak membuat kesimpulan berlebihan."
        ],
        appendix_summary_text(period, mode),
        "Apakah seluruh angka, formula, sumber data, periode, dan ownership sudah cukup valid untuk mendukung keputusan Direksi?",
        "info",
    )

st.divider()
st.caption("Starcom Executive Platform (SEP) V4 — clean board flow: summary, productivity, workforce, capacity, decision, benchmark, closing, appendix.")
