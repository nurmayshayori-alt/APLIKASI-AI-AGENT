import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import os
from io import BytesIO
from datetime import datetime
from groq import Groq
from dotenv import load_dotenv
from fpdf import FPDF

# =====================================================
# 1. PAGE CONFIG & CLEAN MODERN LIGHT THEME CSS
# =====================================================
st.set_page_config(
    page_title="APRIYO | AI Financial & Accounting Suite",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #0f172a;
    }

    .stApp {
        background-color: #f8fafc;
        color: #0f172a;
    }

    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }

    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 18px 22px;
        border-radius: 20px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="metric-container"] label {
        color: #64748b !important;
        font-size: 13.5px !important;
        font-weight: 600;
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #0f172a !important;
        font-size: 24px !important;
        font-weight: 700;
    }

    .hero-banner {
        background: linear-gradient(120deg, #eef2ff 0%, #f0f9ff 60%, #ffffff 100%);
        border: 1px solid #c7d2fe;
        padding: 24px 30px;
        border-radius: 22px;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(79, 70, 229, 0.05);
    }
    .hero-badge {
        display: inline-block;
        background-color: #e0e7ff;
        color: #3730a3;
        padding: 4px 14px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        margin-bottom: 8px;
        border: 1px solid #c7d2fe;
    }
    .hero-banner h1 {
        color: #0f172a !important;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    .hero-banner p {
        color: #475569 !important;
        margin: 5px 0 0 0;
        font-size: 14px;
    }

    .empty-state-box {
        background-color: #ffffff;
        border: 2px dashed #cbd5e1;
        border-radius: 22px;
        padding: 44px 30px;
        text-align: center;
        margin-top: 16px;
    }

    .audit-alert {
        background-color: #fff1f2;
        border: 1px solid #fecdd3;
        border-left: 5px solid #e11d48;
        border-radius: 14px;
        padding: 13px 20px;
        margin-bottom: 18px;
        font-size: 13.5px;
        color: #881337;
    }
    .audit-ok {
        background-color: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-left: 5px solid #10b981;
        border-radius: 14px;
        padding: 13px 20px;
        margin-bottom: 18px;
        font-size: 13.5px;
        color: #065f46;
    }

    .white-panel {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 20px 24px;
        margin-top: 10px;
        margin-bottom: 12px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.03);
    }
    .highlight-row {
        background-color: #f8fafc;
        border: 1px solid #f1f5f9;
        border-radius: 12px;
        padding: 10px 14px;
        margin-bottom: 8px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 13.5px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #ffffff;
        padding: 6px;
        border-radius: 999px;
        border: 1px solid #e2e8f0;
        width: fit-content;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 999px;
        color: #475569;
        padding: 7px 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #4f46e5, #0284c7) !important;
        color: #ffffff !important;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 999px !important;
        padding: 8px 18px !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# 2. LOAD GROQ API KEY & SHARED HELPERS
# =====================================================
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

try:
    if "GROQ_API_KEY" in st.secrets:
        GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except Exception:
    pass

if not GROQ_API_KEY:
    st.error("🚨 GROQ_API_KEY belum diatur di Streamlit Secrets!")
    st.stop()

client = Groq(api_key=GROQ_API_KEY)

def to_excel_bytes(dataframe, sheet_name="Data"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def read_any_tabular_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";")
    return pd.read_excel(uploaded_file)

def smart_column_mapper(df, schema_hints, key_prefix):
    cols = df.columns.tolist()
    mapped = {}
    need_manual = False

    for target, keywords in schema_hints.items():
        found_idx = 0
        matched = False
        for i, c in enumerate(cols):
            c_clean = str(c).lower().strip()
            if any(kw in c_clean for kw in keywords):
                found_idx = i
                matched = True
                break
        if not matched:
            need_manual = True
        mapped[target] = cols[found_idx]

    with st.expander("⚙️ Pemetaan Kolom Otomatis (Klik jika ingin ubah pasangan kolom)", expanded=need_manual):
        st.caption("Sistem mendeteksi kolom dari file Anda secara otomatis. Sesuaikan di bawah jika diperlukan:")
        m_cols = st.columns(len(schema_hints))
        final_df = pd.DataFrame()
        for idx, (target, _) in enumerate(schema_hints.items()):
            default_ix = cols.index(mapped[target]) if mapped[target] in cols else 0
            chosen = m_cols[idx].selectbox(
                f"Kolom '{target}':",
                options=cols,
                index=default_ix,
                key=f"{key_prefix}_map_{target}"
            )
            final_df[target] = df[chosen]
    return final_df

def clean_latin1(text):
    return text.encode("latin-1", "ignore").decode("latin-1")

def build_executive_pdf(company_name, title, df_table, ai_summary):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_fill_color(79, 70, 229)
    pdf.rect(0, 0, 210, 32, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 12, clean_latin1(company_name.upper()), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, clean_latin1(f"APRIYO AI Executive Report  |  Modul: {title}  |  {datetime.now().strftime('%d %b %Y %H:%M')}"), ln=True)

    pdf.ln(14)
    pdf.set_text_color(15, 23, 42)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "1. Ringkasan Data Keuangan", ln=True)
    pdf.set_font("Courier", "", 8.5)
    if df_table is not None and not df_table.empty:
        pdf.multi_cell(0, 5, clean_latin1(df_table.head(30).to_string(index=False)))
    else:
        pdf.multi_cell(0, 5, "Belum ada data tabel dari modul yang dijalankan.")

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "2. Analisis, Temuan Audit & Rekomendasi Strategis AI", ln=True)
    pdf.set_font("Helvetica", "", 10)
    clean_ai = clean_latin1(ai_summary.replace("**", "").replace("#", "")) if ai_summary else "Belum ada narasi analisis AI."
    pdf.multi_cell(0, 6, clean_ai)

    return bytes(pdf.output())

# =====================================================
# 3. SIDEBAR NAVIGATION
# =====================================================
with st.sidebar:
    st.markdown("## ✨ APRIYO AI")
    st.caption("Intelligent Accounting & FP&A Suite")
    st.markdown("<br>", unsafe_allow_html=True)

    app_mode = st.radio(
        "🧭 Pilih Modul Akuntansi & Keuangan:",
        [
            "📊 Budget vs. Actuals",
            "🚀 Scenario Planning",
            "💧 Cash Flow & Liquidity",
            "🕵️ AI Fraud & Anomaly Audit",
            "🧮 Rasio Keuangan & Pajak",
            "🤖 Agentic Copilot & PDF"
        ],
        index=0
    )

    st.markdown("---")
    selected_model = st.selectbox(
        "🧠 Model AI (Groq)",
        ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "llama-3.3-70b-versatile"],
        index=0
    )
    st.caption("Mendukung **.xlsx**, **.xls**, & **.csv** dengan nama kolom bebas.")

# =====================================================
# MODUL 1: BUDGET VS. ACTUALS (VARIANCE & AUDIT)
# =====================================================
if app_mode == "📊 Budget vs. Actuals":
    sample_var = pd.DataFrame({
        "Category": ["Sales Revenue", "COGS", "Marketing", "Salaries", "IT & Cloud", "Operations"],
        "Budget": [150000000, 60000000, 15000000, 35000000, 8000000, 5000000],
        "Actual": [145000000, 64000000, 18500000, 35000000, 7200000, 5600000]
    })

    with st.sidebar:
        uploaded_var = st.file_uploader("📂 Upload Budget vs Actual (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="upload_var")
        st.download_button(
            label="📥 Contoh Excel Variance",
            data=to_excel_bytes(sample_var, "Budget_vs_Actual"),
            file_name="APRIYO_Template_Variance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        audit_threshold = st.slider("🛡️ Batas Toleransi Audit (%)", 5, 30, 10, 1)

    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • VARIANCE & AUDIT INTELLIGENCE</span>
<h1>Budget vs. Actuals Analytics</h1>
<p>Unggah file Excel atau CSV dengan nama kolom bebas untuk deteksi anomali anggaran & skor kesehatan finansial.</p>
</div>
""", unsafe_allow_html=True)

    if uploaded_var is not None:
        raw_df = read_any_tabular_file(uploaded_var)
        df = smart_column_mapper(
            raw_df,
            {
                "Category": ["cat", "kat", "akun", "account", "pos", "nama", "item", "desc"],
                "Budget": ["bud", "angg", "plan", "target", "rencana"],
                "Actual": ["act", "real", "aktual", "spent", "terpakai"]
            },
            key_prefix="var"
        )

        df["Budget"] = pd.to_numeric(df["Budget"], errors="coerce").fillna(0)
        df["Actual"] = pd.to_numeric(df["Actual"], errors="coerce").fillna(0)
        df["Variance"] = df["Actual"] - df["Budget"]
        df["Variance %"] = (df["Variance"] / df["Budget"].replace(0, float("nan"))) * 100
        df["Audit Flag"] = df["Variance %"].apply(
            lambda x: "🚨 Anomali (>Batas)" if pd.notna(x) and abs(x) > audit_threshold else "✅ Normal"
        )

        st.session_state["last_module_name"] = "Budget vs. Actuals Variance"
        st.session_state["last_module_df"] = df

        tot_b = df["Budget"].sum()
        tot_a = df["Actual"].sum()
        tot_v = tot_a - tot_b
        v_pct = (tot_v / tot_b * 100) if tot_b != 0 else 0
        realisasi = (tot_a / tot_b * 100) if tot_b != 0 else 0
        avg_abs_dev = df["Variance %"].abs().mean() if not df.empty else 0
        health_score = max(0, min(100, int(100 - (avg_abs_dev if pd.notna(avg_abs_dev) else 0))))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 Total Budget", f"{tot_b:,.0f}", f"{len(df)} Pos Anggaran")
        c2.metric("💳 Total Actual", f"{tot_a:,.0f}", f"Rasio: {realisasi:.1f}%")
        c3.metric("📉 Total Variance", f"{tot_v:,.0f}", delta=f"{v_pct:+.1f}%", delta_color="inverse")
        c4.metric("🛡️ Budget Health Score", f"{health_score} / 100", delta="Sehat" if health_score >= 80 else "Perlu Audit")

        anomalies = df[df["Variance %"].abs() > audit_threshold]
        if not anomalies.empty:
            anom_list = ", ".join([f"{row['Category']} ({row['Variance %']:+.1f}%)" for _, row in anomalies.iterrows()])
            st.markdown(f"""
<div class="audit-alert">
<strong>🚨 Peringatan Audit Otomatis:</strong> Ditemukan <b>{len(anomalies)} kategori</b> melebihi batas toleransi ({audit_threshold}%): <b>{anom_list}</b>.
</div>
""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
<div class="audit-ok">
<strong>✅ Status Audit Aman:</strong> Seluruh kategori berada dalam batas toleransi deviasi ({audit_threshold}%).
</div>
""", unsafe_allow_html=True)

        t1, t2, t3 = st.tabs(["✨ Visualisasi Grafik", "📋 Rincian & Export", "🤖 AI Audit Analyst"])

        with t1:
            col_g1, col_g2, col_g3 = st.columns([1.2, 1.2, 0.9])
            with col_g1:
                fig_area = go.Figure()
                fig_area.add_trace(go.Scatter(x=df["Category"], y=df["Budget"], name="Budget", mode="lines+markers", line=dict(color="#4f46e5", width=3, shape="spline"), fill="tozeroy", fillcolor="rgba(79, 70, 229, 0.10)"))
                fig_area.add_trace(go.Scatter(x=df["Category"], y=df["Actual"], name="Actual", mode="lines+markers", line=dict(color="#0ea5e9", width=3, shape="spline"), fill="tozeroy", fillcolor="rgba(14, 165, 233, 0.10)"))
                fig_area.update_layout(title="Tren Budget vs. Actual", template="plotly_white", paper_bgcolor="#ffffff", plot_bgcolor="#ffffff", margin=dict(t=50, b=30, l=15, r=15), height=340)
                st.plotly_chart(fig_area, use_container_width=True)

            with col_g2:
                colors = ["#e11d48" if v > 0 else "#10b981" for v in df["Variance"]]
                fig_var = go.Figure(go.Bar(x=df["Variance"], y=df["Category"], orientation="h", marker=dict(color=colors, cornerradius=8), text=df["Variance"], texttemplate="%{text:,.0f}", textposition="auto"))
                fig_var.add_vline(x=0, line_width=1, line_dash="dot", line_color="#94a3b8")
                fig_var.update_layout(title="Deviasi (Variance) per Kategori", template="plotly_white", paper_bgcolor="#ffffff", plot_bgcolor="#ffffff", yaxis=dict(autorange="reversed"), margin=dict(t=50, b=30, l=15, r=15), height=340)
                st.plotly_chart(fig_var, use_container_width=True)

            with col_g3:
                fig_donut = go.Figure(data=[go.Pie(labels=df["Category"], values=df["Actual"].abs(), hole=0.62, textinfo="percent", marker=dict(colors=["#4f46e5", "#0ea5e9", "#e11d48", "#10b981", "#f59e0b", "#6366f1"]))])
                fig_donut.update_layout(title="Proporsi Realisasi", template="plotly_white", paper_bgcolor="#ffffff", showlegend=True, legend=dict(orientation="h", y=-0.15, font=dict(size=10)), margin=dict(t=50, b=35, l=10, r=10), height=340)
                st.plotly_chart(fig_donut, use_container_width=True)

            top_devs = df.reindex(df["Variance"].abs().sort_values(ascending=False).index).head(3)
            rows_html = ""
            for idx, (_, r) in enumerate(top_devs.iterrows(), 1):
                badge_col = "#e11d48" if r["Variance"] > 0 else "#059669"
                status_txt = "Over Budget" if r["Variance"] > 0 else "Efisiensi"
                rows_html += f'<div class="highlight-row"><span style="color: #0f172a; font-weight: 600;">{idx}. {r["Category"]}</span><span style="color: {badge_col}; font-weight: 700;">{status_txt} {r["Variance"]:+,.0f} ({r["Variance %"]:+.1f}%)</span></div>'

            st.markdown(f"""
<div class="white-panel">
<div style="color: #0f172a; font-size: 16px; font-weight: 700; margin-bottom: 12px;">🏆 Top Deviasi & Sorotan Kategori</div>
{rows_html}
</div>
""", unsafe_allow_html=True)

        with t2:
            st.download_button("💾 Export Excel Hasil Variance", data=to_excel_bytes(df, "Variance"), file_name="APRIYO_Variance.xlsx")
            st.dataframe(df.style.format({"Budget": "{:,.2f}", "Actual": "{:,.2f}", "Variance": "{:,.2f}", "Variance %": "{:.2f}%"}), use_container_width=True)

        with t3:
            if st.button("✨ Generate Analisis Audit AI", type="primary"):
                with st.spinner("AI mengaudit data..."):
                    res = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": "You are a senior financial auditor. Provide concise executive summary and recommendations in Indonesian."},
                            {"role": "user", "content": f"Analisis Budget vs Actual (Toleransi {audit_threshold}%):\n{df.to_string(index=False)}"}
                        ],
                        model=selected_model
                    )
                    st.session_state["var_ai"] = res.choices[0].message.content
                    st.session_state["last_module_ai"] = st.session_state["var_ai"]
            if "var_ai" in st.session_state:
                st.markdown(st.session_state["var_ai"])
    else:
        st.markdown("""
<div class="empty-state-box">
<h3 style="color: #0f172a; margin-bottom: 8px;">📂 Unggah File Budget vs. Actual (.xlsx atau .csv)</h3>
<p style="color: #475569; font-size: 14.5px; max-width: 560px; margin: 0 auto;">
Nama kolom bebas (contoh: <b>Akun, Anggaran, Realisasi</b> atau <b>Category, Budget, Actual</b>). Sistem otomatis mengenali kolom Anda!
</p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# MODUL 2: SCENARIO PLANNING & WHAT-IF SIMULATOR
# =====================================================
elif app_mode == "🚀 Scenario Planning":
    sample_fc = pd.DataFrame({
        "Category": ["Product Revenue", "Service Income", "Direct Labor", "Raw Materials", "Marketing Opex"],
        "Base Forecast": [250000000, 90000000, 75000000, 85000000, 30000000]
    })

    with st.sidebar:
        uploaded_fc = st.file_uploader("📂 Upload File Forecast (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="upload_fc")
        st.download_button(
            label="📥 Contoh Excel Forecast",
            data=to_excel_bytes(sample_fc, "Forecast"),
            file_name="APRIYO_Template_Forecast.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        custom_shock = st.slider("🎛️ Simulasi What-If (%)", -50, 50, 0, 5)
        scenario_prompt = st.text_area("📝 Asumsi Skenario:", placeholder="Misal: Pendapatan turun 10%, biaya naik 5%", height=80)

    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • PREDICTIVE FP&A & WHAT-IF SIMULATOR</span>
<h1>Scenario Planning & Strategic Copilot</h1>
<p>Simulasi proyeksi multi-skenario (Optimistic, Pessimistic, Worst Case, & Custom Shock) secara instan.</p>
</div>
""", unsafe_allow_html=True)

    if uploaded_fc is not None:
        raw_fc = read_any_tabular_file(uploaded_fc)
        df_fc = smart_column_mapper(
            raw_fc,
            {
                "Category": ["cat", "kat", "akun", "pos", "nama", "item"],
                "Base Forecast": ["base", "fore", "proyeksi", "nilai", "amount", "budget", "nominal"]
            },
            key_prefix="fc"
        )

        df_fc["Base Forecast"] = pd.to_numeric(df_fc["Base Forecast"], errors="coerce").fillna(0)
        np.random.seed(42)
        df_fc["Optimistic"] = df_fc["Base Forecast"] * np.random.uniform(1.10, 1.25, len(df_fc))
        df_fc["Pessimistic"] = df_fc["Base Forecast"] * np.random.uniform(0.75, 0.90, len(df_fc))
        df_fc["Worst Case"] = df_fc["Base Forecast"] * np.random.uniform(0.55, 0.70, len(df_fc))
        custom_col = f"Custom ({custom_shock:+d}%)"
        df_fc[custom_col] = df_fc["Base Forecast"] * (1 + custom_shock / 100.0)

        st.session_state["last_module_name"] = "Scenario Planning Forecast"
        st.session_state["last_module_df"] = df_fc

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📌 Base Forecast", f"{df_fc['Base Forecast'].sum():,.0f}", f"{len(df_fc)} Kategori")
        k2.metric("🌟 Optimistic", f"{df_fc['Optimistic'].sum():,.0f}", delta="+10% s/d +25%")
        k3.metric("🚨 Worst Case", f"{df_fc['Worst Case'].sum():,.0f}", delta="-30% s/d -45%", delta_color="inverse")
        k4.metric(f"🎛️ {custom_col}", f"{df_fc[custom_col].sum():,.0f}", delta=f"{custom_shock:+d}% Simulasi")

        t1, t2, t3 = st.tabs(["✨ Grafik Proyeksi", "📋 Tabel & Export", "🤖 AI Strategic Copilot"])

        with t1:
            fig_fc = go.Figure(data=[
                go.Bar(name="Base Forecast", x=df_fc["Category"], y=df_fc["Base Forecast"], marker=dict(color="#4f46e5", cornerradius=6)),
                go.Bar(name="Optimistic", x=df_fc["Category"], y=df_fc["Optimistic"], marker=dict(color="#10b981", cornerradius=6)),
                go.Bar(name="Pessimistic", x=df_fc["Category"], y=df_fc["Pessimistic"], marker=dict(color="#f59e0b", cornerradius=6)),
                go.Bar(name="Worst Case", x=df_fc["Category"], y=df_fc["Worst Case"], marker=dict(color="#e11d48", cornerradius=6)),
                go.Scatter(name=custom_col, x=df_fc["Category"], y=df_fc[custom_col], mode="lines+markers", line=dict(color="#0f172a", width=3, dash="dash"))
            ])
            fig_fc.update_layout(title="Simulasi Proyeksi Multi-Skenario", barmode="group", template="plotly_white", paper_bgcolor="#ffffff", plot_bgcolor="#ffffff", height=380)
            st.plotly_chart(fig_fc, use_container_width=True)

        with t2:
            st.download_button("💾 Export Excel Forecast", data=to_excel_bytes(df_fc, "Forecast"), file_name="APRIYO_Forecast.xlsx")
            st.dataframe(df_fc.style.format({"Base Forecast": "{:,.2f}", "Optimistic": "{:,.2f}", "Pessimistic": "{:,.2f}", "Worst Case": "{:,.2f}", custom_col: "{:,.2f}"}), use_container_width=True)

        with t3:
            if st.button("🚀 Generate Strategi AI", type="primary"):
                with st.spinner("Menyusun strategi..."):
                    res_fc = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": "You are an AI Financial Copilot specializing in scenario planning. Respond in Indonesian."},
                            {"role": "user", "content": f"Data:\n{df_fc.to_string(index=False)}\nAsumsi: {scenario_prompt}\nShock: {custom_shock:+d}%"}
                        ],
                        model=selected_model
                    )
                    st.session_state["fc_ai_summary"] = res_fc.choices[0].message.content
                    st.session_state["last_module_ai"] = st.session_state["fc_ai_summary"]
            if "fc_ai_summary" in st.session_state:
                st.markdown(st.session_state["fc_ai_summary"])
    else:
        st.markdown("""
<div class="empty-state-box">
<h3 style="color: #0f172a; margin-bottom: 8px;">📂 Unggah File Forecast (.xlsx atau .csv)</h3>
<p style="color: #475569; font-size: 14.5px; max-width: 560px; margin: 0 auto;">
Unggah tabel proyeksi dengan kolom nama akun/kategori dan nilai dasar proyeksi (nama kolom bebas).
</p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# MODUL 3: CASH FLOW & LIQUIDITY
# =====================================================
elif app_mode == "💧 Cash Flow & Liquidity":
    sample_cf = pd.DataFrame({
        "Bulan": ["Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "Mei 2026", "Jun 2026"],
        "Cash Inflow": [320000000, 290000000, 350000000, 310000000, 380000000, 410000000],
        "Cash Outflow": [280000000, 315000000, 300000000, 340000000, 330000000, 350000000]
    })

    with st.sidebar:
        uploaded_cf = st.file_uploader("📂 Upload Data Arus Kas (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="upload_cf")
        st.download_button(
            label="📥 Contoh Excel Cash Flow",
            data=to_excel_bytes(sample_cf, "CashFlow"),
            file_name="APRIYO_Template_CashFlow.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        initial_cash = st.number_input("🏦 Saldo Kas Awal (Rp)", min_value=0, value=500000000, step=10000000)

    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • TREASURY & LIQUIDITY MANAGEMENT</span>
<h1>Cash Flow & Liquidity Monitor</h1>
<p>Pantau arus kas masuk-keluar, Net Cash Flow kumulatif, rata-rata Burn Rate, dan estimasi Cash Runway perusahaan.</p>
</div>
""", unsafe_allow_html=True)

    if uploaded_cf is not None:
        raw_cf = read_any_tabular_file(uploaded_cf)
        df_cf = smart_column_mapper(
            raw_cf,
            {
                "Periode": ["bul", "month", "tgl", "date", "per", "waktu"],
                "Cash Inflow": ["in", "masuk", "terima", "pendapatan", "rev", "debet"],
                "Cash Outflow": ["out", "keluar", "bayar", "biaya", "exp", "kredit"]
            },
            key_prefix="cf"
        )

        df_cf["Cash Inflow"] = pd.to_numeric(df_cf["Cash Inflow"], errors="coerce").fillna(0)
        df_cf["Cash Outflow"] = pd.to_numeric(df_cf["Cash Outflow"], errors="coerce").fillna(0)
        df_cf["Net Cash Flow"] = df_cf["Cash Inflow"] - df_cf["Cash Outflow"]
        df_cf["Saldo Kas Akhir"] = initial_cash + df_cf["Net Cash Flow"].cumsum()

        st.session_state["last_module_name"] = "Cash Flow & Liquidity Analysis"
        st.session_state["last_module_df"] = df_cf

        tot_in = df_cf["Cash Inflow"].sum()
        tot_out = df_cf["Cash Outflow"].sum()
        ending_cash = df_cf["Saldo Kas Akhir"].iloc[-1] if not df_cf.empty else initial_cash
        avg_outflow = df_cf["Cash Outflow"].mean() if not df_cf.empty else 1
        runway_months = (ending_cash / avg_outflow) if avg_outflow > 0 else 99

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📥 Total Cash Inflow", f"{tot_in:,.0f}", f"{len(df_cf)} Periode")
        c2.metric("📤 Total Cash Outflow", f"{tot_out:,.0f}", f"Avg: {avg_outflow:,.0f}/bln")
        c3.metric("🏦 Saldo Kas Akhir", f"{ending_cash:,.0f}", delta=f"{(ending_cash - initial_cash):+,.0f}")
        c4.metric("⏳ Estimasi Cash Runway", f"{runway_months:.1f} Bulan", delta="Aman (>6 Bln)" if runway_months >= 6 else "Waspada (<6 Bln)")

        t1, t2 = st.tabs(["📈 Grafik Arus Kas & Likuiditas", "🤖 Analisis Treasury AI"])
        with t1:
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Cash Inflow", x=df_cf["Periode"], y=df_cf["Cash Inflow"], marker=dict(color="#10b981", cornerradius=6)))
            fig.add_trace(go.Bar(name="Cash Outflow", x=df_cf["Periode"], y=df_cf["Cash Outflow"], marker=dict(color="#e11d48", cornerradius=6)))
            fig.add_trace(go.Scatter(name="Saldo Kas Akhir", x=df_cf["Periode"], y=df_cf["Saldo Kas Akhir"], mode="lines+markers", line=dict(color="#4f46e5", width=3)))
            fig.update_layout(title="Arus Kas Masuk, Keluar & Tren Saldo Kas", barmode="group", template="plotly_white", paper_bgcolor="#ffffff", plot_bgcolor="#ffffff", height=380)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df_cf.style.format({"Cash Inflow": "{:,.0f}", "Cash Outflow": "{:,.0f}", "Net Cash Flow": "{:+,.0f}", "Saldo Kas Akhir": "{:,.0f}"}), use_container_width=True)

        with t2:
            if st.button("✨ Generate Rekomendasi Likuiditas AI", type="primary"):
                with st.spinner("Menganalisis likuiditas..."):
                    res = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": "You are a Corporate Treasurer & CFO Advisor. Analyze cash flow, working capital, and liquidity runway in Indonesian."},
                            {"role": "user", "content": f"Saldo Awal: {initial_cash:,.0f}\nData Arus Kas:\n{df_cf.to_string(index=False)}"}
                        ],
                        model=selected_model
                    )
                    st.session_state["cf_ai"] = res.choices[0].message.content
                    st.session_state["last_module_ai"] = st.session_state["cf_ai"]
            if "cf_ai" in st.session_state:
                st.markdown(st.session_state["cf_ai"])
    else:
        st.markdown("""
<div class="empty-state-box">
<h3 style="color: #0f172a; margin-bottom: 8px;">📂 Unggah Data Arus Kas (.xlsx atau .csv)</h3>
<p style="color: #475569; font-size: 14.5px; max-width: 560px; margin: 0 auto;">
Unggah tabel arus kas dengan kolom Periode/Bulan, Kas Masuk (Inflow), dan Kas Keluar (Outflow).
</p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# MODUL 4: AI FRAUD & ANOMALY AUDIT
# =====================================================
elif app_mode == "🕵️ AI Fraud & Anomaly Audit":
    sample_gl = pd.DataFrame({
        "Tanggal": ["2026-10-01", "2026-10-03", "2026-10-03", "2026-10-04", "2026-10-05", "2026-10-07"],
        "No_Bukti": ["JV-001", "JV-002", "JV-003", "JV-004", "JV-005", "JV-006"],
        "Keterangan": ["Pembelian ATK", "Bayar Vendor IT", "Bayar Vendor IT", "Perjalanan Dinas", "Biaya Konsultan", "Tagihan Listrik"],
        "Nominal": [1250000, 15000000, 15000000, 50000000, 3250000, 4120000]
    })

    with st.sidebar:
        uploaded_gl = st.file_uploader("📂 Upload Jurnal / Buku Besar (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="upload_gl")
        st.download_button(
            label="📥 Contoh Excel Jurnal Audit",
            data=to_excel_bytes(sample_gl, "Jurnal_Umum"),
            file_name="APRIYO_Template_Audit_Jurnal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        round_limit = st.number_input("🚨 Ambang Angka Bulat (Rp)", min_value=1000000, value=10000000, step=1000000)

    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • FORENSIC ACCOUNTING & INTERNAL AUDIT</span>
<h1>AI Fraud & Anomaly Audit</h1>
<p>Deteksi otomatis transaksi ganda (duplicate entries), jurnal akhir pekan (weekend posting), dan angka bulat mencurigakan.</p>
</div>
""", unsafe_allow_html=True)

    if uploaded_gl is not None:
        raw_gl = read_any_tabular_file(uploaded_gl)
        df_gl = smart_column_mapper(
            raw_gl,
            {
                "Tanggal": ["tgl", "date", "tanggal", "waktu"],
                "Keterangan": ["ket", "desc", "uraian", "memo", "transaksi", "nama"],
                "Nominal": ["nom", "amount", "nilai", "jumlah", "total", "rp", "debet"]
            },
            key_prefix="gl"
        )

        df_gl["Nominal"] = pd.to_numeric(df_gl["Nominal"], errors="coerce").fillna(0)
        df_gl["Tanggal_Parsed"] = pd.to_datetime(df_gl["Tanggal"], errors="coerce")

        is_dup = df_gl.duplicated(subset=["Tanggal", "Keterangan", "Nominal"], keep=False)
        is_weekend = df_gl["Tanggal_Parsed"].dt.dayofweek.isin([5, 6]).fillna(False)
        is_round = (df_gl["Nominal"] >= round_limit) & (df_gl["Nominal"] % 1000000 == 0)

        def flag_row(idx):
            flags = []
            if is_dup.iloc[idx]:
                flags.append("🔁 Duplikat")
            if is_weekend.iloc[idx]:
                flags.append("📅 Akhir Pekan")
            if is_round.iloc[idx]:
                flags.append("🎯 Angka Bulat Besar")
            return " | ".join(flags) if flags else "✅ Bersih"

        df_gl["Status Audit"] = [flag_row(i) for i in range(len(df_gl))]
        df_view = df_gl.drop(columns=["Tanggal_Parsed"])

        st.session_state["last_module_name"] = "Forensic Fraud & Anomaly Audit"
        st.session_state["last_module_df"] = df_view

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📑 Total Baris Jurnal", f"{len(df_view):,}", f"Total: Rp {df_view['Nominal'].sum():,.0f}")
        c2.metric("🔁 Potensi Duplikat", f"{is_dup.sum()} Baris", delta="Perlu Cek" if is_dup.sum() > 0 else "Aman", delta_color="inverse")
        c3.metric("📅 Jurnal Akhir Pekan", f"{is_weekend.sum()} Baris", delta="Anomali Waktu" if is_weekend.sum() > 0 else "Normal", delta_color="inverse")
        c4.metric("🎯 Angka Bulat Besar", f"{is_round.sum()} Baris", f">= Rp {round_limit:,.0f}")

        st.subheader("📋 Kertas Kerja Audit & Deteksi Red Flags")
        st.dataframe(df_view.style.format({"Nominal": "{:,.0f}"}), use_container_width=True)

        if st.button("🕵️ Minta AI Auditor Susun Temuan & Prosedur Audit", type="primary"):
            flagged = df_view[df_view["Status Audit"] != "✅ Bersih"]
            with st.spinner("AI Auditor menganalisis red flags..."):
                res = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": "You are a Certified Fraud Examiner (CFE) and Senior Auditor. Analyze flagged journal entries and suggest audit procedures in Indonesian."},
                        {"role": "user", "content": f"Berikut transaksi yang terdeteksi anomali:\n{flagged.to_string(index=False)}"}
                    ],
                    model=selected_model
                )
                st.session_state["fraud_ai"] = res.choices[0].message.content
                st.session_state["last_module_ai"] = st.session_state["fraud_ai"]

        if "fraud_ai" in st.session_state:
            st.markdown(st.session_state["fraud_ai"])
    else:
        st.markdown("""
<div class="empty-state-box">
<h3 style="color: #0f172a; margin-bottom: 8px;">📂 Unggah Data Jurnal / Mutasi Transaksi (.xlsx atau .csv)</h3>
<p style="color: #475569; font-size: 14.5px; max-width: 560px; margin: 0 auto;">
Sistem otomatis menyisir transaksi kembar, input di hari libur/akhir pekan, dan transaksi bernilai bulat besar.
</p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# MODUL 5: RASIO KEUANGAN, BEP & PAJAK CEPAT
# =====================================================
elif app_mode == "🧮 Rasio Keuangan & Pajak":
    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • FINANCIAL RATIOS, BEP & TAX CALCULATOR</span>
<h1>Rasio Keuangan, BEP & Pajak Cepat</h1>
<p>Hitung kesehatan rasio neraca/laba-rugi, titik impas (Break-Even Point), serta simulasi PPN & PPh Indonesia secara instan.</p>
</div>
""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📊 Analisis Rasio Keuangan", "⚖️ Kalkulator BEP (Titik Impas)", "🧾 Kalkulator Pajak (PPN & PPh)"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 🏛️ Data Neraca (Balance Sheet)")
            aset_lancar = st.number_input("Aset Lancar (Rp)", min_value=0, value=450000000, step=10000000)
            persediaan = st.number_input("Persediaan / Inventory (Rp)", min_value=0, value=120000000, step=5000000)
            total_aset = st.number_input("Total Aset (Rp)", min_value=1, value=950000000, step=10000000)
            liabilitas_lancar = st.number_input("Liabilitas Jangka Pendek (Rp)", min_value=1, value=200000000, step=10000000)
            total_utang = st.number_input("Total Liabilitas / Utang (Rp)", min_value=0, value=380000000, step=10000000)
            total_ekuitas = st.number_input("Total Ekuitas / Modal (Rp)", min_value=1, value=570000000, step=10000000)

        with col2:
            st.markdown("#### 📈 Data Laba Rugi (Income Statement)")
            penjualan = st.number_input("Penjualan Bersih (Rp)", min_value=1, value=1200000000, step=25000000)
            hpp = st.number_input("Harga Pokok Penjualan / COGS (Rp)", min_value=0, value=720000000, step=10000000)
            laba_bersih = st.number_input("Laba Bersih Setelah Pajak (Rp)", value=165000000, step=5000000)

        current_ratio = aset_lancar / liabilitas_lancar
        quick_ratio = (aset_lancar - persediaan) / liabilitas_lancar
        der = total_utang / total_ekuitas
        gpm = ((penjualan - hpp) / penjualan) * 100
        npm = (laba_bersih / penjualan) * 100
        roe = (laba_bersih / total_ekuitas) * 100

        st.markdown("---")
        r1, r2, r3, r4, r5, r6 = st.columns(6)
        r1.metric("Current Ratio", f"{current_ratio:.2f}x", "Ideal > 1.5x")
        r2.metric("Quick Ratio", f"{quick_ratio:.2f}x", "Ideal > 1.0x")
        r3.metric("Debt to Equity", f"{der:.2f}x", "Ideal < 1.5x")
        r4.metric("Gross Margin", f"{gpm:.1f}%")
        r5.metric("Net Margin", f"{npm:.1f}%")
        r6.metric("ROE", f"{roe:.1f}%")

        df_ratio = pd.DataFrame({
            "Indikator Rasio": ["Current Ratio", "Quick Ratio", "Debt to Equity Ratio (DER)", "Gross Profit Margin (%)", "Net Profit Margin (%)", "Return on Equity (ROE %)"],
            "Nilai": [round(current_ratio, 2), round(quick_ratio, 2), round(der, 2), round(gpm, 2), round(npm, 2), round(roe, 2)]
        })
        st.session_state["last_module_name"] = "Analisis Rasio Keuangan"
        st.session_state["last_module_df"] = df_ratio

        if st.button("✨ Interpretasi Rasio oleh AI", type="primary"):
            with st.spinner("AI menilai kesehatan rasio..."):
                res = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": "You are a Financial Analyst. Interpret liquidity, solvency, and profitability ratios clearly in Indonesian."},
                        {"role": "user", "content": f"Hasil Rasio Keuangan:\n{df_ratio.to_string(index=False)}"}
                    ],
                    model=selected_model
                )
                st.session_state["ratio_ai"] = res.choices[0].message.content
                st.session_state["last_module_ai"] = st.session_state["ratio_ai"]
        if "ratio_ai" in st.session_state:
            st.markdown(st.session_state["ratio_ai"])

    with tab2:
        b1, b2, b3 = st.columns(3)
        fixed_cost = b1.number_input("Biaya Tetap / Fixed Cost (Rp)", min_value=0, value=75000000, step=5000000)
        price_unit = b2.number_input("Harga Jual per Unit (Rp)", min_value=1, value=150000, step=5000)
        var_unit = b3.number_input("Biaya Variabel per Unit (Rp)", min_value=0, value=90000, step=5000)

        margin_kontribusi = price_unit - var_unit
        if margin_kontribusi > 0:
            bep_unit = fixed_cost / margin_kontribusi
            bep_rp = bep_unit * price_unit
            k1, k2, k3 = st.columns(3)
            k1.metric("Contribution Margin / Unit", f"Rp {margin_kontribusi:,.0f}")
            k2.metric("🎯 BEP (Dalam Unit)", f"{bep_unit:,.0f} Unit")
            k3.metric("💰 BEP (Dalam Rupiah)", f"Rp {bep_rp:,.0f}")
        else:
            st.error("⚠️ Harga jual per unit harus lebih besar dari biaya variabel per unit.")

    with tab3:
        p1, p2 = st.columns(2)
        with p1:
            st.markdown("#### 🧾 Simulasi PPN (11% / 12%)")
            dpp = st.number_input("Dasar Pengenaan Pajak / Nilai Transaksi (Rp)", min_value=0, value=50000000, step=1000000)
            tarif_ppn = st.selectbox("Tarif PPN:", [11, 12], index=0)
            nilai_ppn = dpp * (tarif_ppn / 100)
            st.info(f"**PPN ({tarif_ppn}%):** Rp {nilai_ppn:,.0f}  |  **Total Tagihan (Inc. PPN):** Rp {(dpp + nilai_ppn):,.0f}")

        with p2:
            st.markdown("#### 🏢 Simulasi PPh Badan & UMKM")
            jenis_pph = st.selectbox("Skema PPh:", ["PPh Final UMKM (0.5% Omzet)", "PPh Badan Umum (22% Laba Kena Pajak)"])
            dasar_pph = st.number_input("Omzet / Laba Kena Pajak (Rp)", min_value=0, value=200000000, step=5000000)
            tarif = 0.005 if "0.5%" in jenis_pph else 0.22
            pph_terutang = dasar_pph * tarif
            st.success(f"**Estimasi PPh Terutang:** Rp {pph_terutang:,.0f}")

# =====================================================
# MODUL 6: AGENTIC COPILOT & EXPORT PDF RESMI
# =====================================================
elif app_mode == "🤖 Agentic Copilot & PDF":
    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • AGENTIC CFO COPILOT & PDF PUBLISHER</span>
<h1>Agentic Copilot & Export Laporan PDF</h1>
<p>Diskusikan hasil analisis dari modul terakhir bersama AI CFO Advisor dan cetak langsung menjadi dokumen PDF resmi.</p>
</div>
""", unsafe_allow_html=True)

    active_mod = st.session_state.get("last_module_name", "Belum ada modul yang dijalankan")
    active_df = st.session_state.get("last_module_df", None)
    active_ai = st.session_state.get("last_module_ai", "")

    col_left, col_right = st.columns([1.3, 1])

    with col_left:
        st.subheader("💬 Konsultasi Interaktif AI CFO")
        st.caption(f"Konteks Data Aktif: **{active_mod}**")

        if "copilot_messages" not in st.session_state:
            st.session_state["copilot_messages"] = []

        for msg in st.session_state["copilot_messages"]:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        user_q = st.text_input("Tanyakan strategi, koreksi jurnal, atau kesimpulan manajemen:", placeholder="Misal: Buatkan ringkasan 3 poin untuk presentasi direksi...")
        b1, b2 = st.columns([3, 1])
        if b1.button("Kirim ke Copilot", type="primary", use_container_width=True) and user_q:
            context_str = active_df.to_string(index=False) if active_df is not None else "Tidak ada tabel terlampir."
            with st.spinner("Copilot sedang berpikir..."):
                resp = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": "You are an Agentic CFO & Senior Accounting Copilot. Give sharp, structured executive advice in Indonesian."},
                        *st.session_state["copilot_messages"],
                        {"role": "user", "content": f"Modul Aktif: {active_mod}\nData:\n{context_str}\n\nPertanyaan User: {user_q}"}
                    ],
                    model=selected_model
                ).choices[0].message.content

                st.session_state["copilot_messages"].append({"role": "user", "content": user_q})
                st.session_state["copilot_messages"].append({"role": "assistant", "content": resp})
                st.session_state["last_module_ai"] = resp
                st.rerun()

        if b2.button("Reset Chat", use_container_width=True):
            st.session_state["copilot_messages"] = []
            st.rerun()

    with col_right:
        st.markdown("""
<div class="white-panel">
<div style="color: #0f172a; font-size: 16px; font-weight: 700; margin-bottom: 6px;">🖨️ Cetak Laporan Resmi (PDF)</div>
<p style="color: #475569; font-size: 13px; margin-bottom: 0;">Gabungkan tabel hasil olahan dan narasi analisis AI terakhir menjadi laporan eksekutif siap cetak.</p>
</div>
""", unsafe_allow_html=True)

        company_name = st.text_input("Nama Perusahaan / Entitas pada Kop PDF:", value="PT APRIYO NUSANTARA SEJAHTERA")
        custom_note = st.text_area("Catatan / Narasi Laporan (Otomatis terisi dari AI terakhir):", value=active_ai, height=180)

        pdf_bytes = build_executive_pdf(company_name, active_mod, active_df, custom_note)
        st.download_button(
            label="📥 Unduh Laporan Eksekutif (.PDF)",
            data=pdf_bytes,
            file_name="APRIYO_Executive_Report.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
