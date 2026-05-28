import streamlit as st
import pandas as pd
import io

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MFolks Excel Cleaner",
    page_icon="✨",
    layout="centered",
)

# ── Global CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Google Font ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ── Root vars ── */
:root {
    --bg:        #f8fafc;
    --surface:   rgba(255, 255, 255, 0.7);
    --border:    rgba(124, 58, 237, 0.12);
    --accent1:   #7c3aed;
    --accent2:   #06b6d4;
    --accent3:   #10b981;
    --danger:    #f43f5e;
    --text:      #0f172a;
    --muted:     #64748b;
    --radius:    16px;
    --glow:      0 10px 30px rgba(124,58,237,0.06);
}

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    background-color: var(--bg) !important;
    color: var(--text) !important;
}

/* Animated mesh gradient background */
.stApp {
    background:
        radial-gradient(ellipse 80% 60% at 20% 0%,  rgba(124,58,237,0.08) 0%, transparent 60%),
        radial-gradient(ellipse 60% 50% at 80% 100%, rgba(6,182,212,0.06)  0%, transparent 60%),
        #f8fafc !important;
    min-height: 100vh;
}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.block-container {
    max-width: 760px !important;
    padding-top: 2.5rem !important;
    padding-bottom: 4rem !important;
}

/* ── Hero header ── */
.hero {
    text-align: center;
    padding: 3rem 1rem 2.5rem;
}
.hero-badge {
    display: inline-block;
    background: rgba(124,58,237,0.08);
    border: 1px solid rgba(124,58,237,0.2);
    color: #6d28d9;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 0.35rem 0.9rem;
    border-radius: 100px;
    margin-bottom: 1.2rem;
}
.hero h1 {
    font-size: 2.8rem !important;
    font-weight: 800 !important;
    line-height: 1.15 !important;
    margin: 0 0 0.8rem !important;
    background: linear-gradient(135deg, #0f172a 0%, #7c3aed 50%, #06b6d4 100%);
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
}
.hero p {
    font-size: 1.05rem;
    color: var(--muted);
    max-width: 480px;
    margin: 0 auto;
    line-height: 1.7;
}

/* ── Upload card ── */
.upload-card {
    background: var(--surface);
    border: 1.5px dashed rgba(124,58,237,0.3);
    border-radius: var(--radius);
    padding: 2.5rem 2rem;
    text-align: center;
    transition: border-color 0.3s, background 0.3s;
    margin-bottom: 1.5rem;
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.04);
}
.upload-card:hover {
    border-color: var(--accent1);
    background: rgba(124,58,237,0.04);
}
.upload-icon {
    font-size: 3rem;
    margin-bottom: 0.8rem;
    animation: float 3s ease-in-out infinite;
}
@keyframes float {
    0%,100% { transform: translateY(0); }
    50%      { transform: translateY(-8px); }
}
.upload-label {
    font-size: 1.1rem;
    font-weight: 600;
    color: var(--text);
}
.upload-sublabel {
    font-size: 0.85rem;
    color: var(--muted);
    margin-top: 0.3rem;
}

/* ── Override Streamlit file uploader ── */
[data-testid="stFileUploader"] {
    background: transparent !important;
}
[data-testid="stFileUploader"] section {
    background: rgba(255,255,255,0.6) !important;
    border: 1.5px dashed rgba(124,58,237,0.3) !important;
    border-radius: var(--radius) !important;
    padding: 2rem !important;
    transition: all 0.3s !important;
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.03) !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--accent1) !important;
    background: rgba(124,58,237,0.04) !important;
}
[data-testid="stFileUploaderDropzone"] label {
    color: var(--muted) !important;
}

/* ── Stats row ── */
.stats-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1rem;
    margin: 1.5rem 0;
}
.stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.2rem 1rem;
    text-align: center;
    backdrop-filter: blur(12px);
    transition: transform 0.2s, box-shadow 0.2s;
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.04);
}
.stat-card:hover { transform: translateY(-3px); box-shadow: 0 12px 40px rgba(31, 38, 135, 0.08); }
.stat-value {
    font-size: 2rem;
    font-weight: 800;
    background: linear-gradient(135deg, #0f172a, #475569);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.stat-value.green  { background: linear-gradient(135deg, #10b981, #047857); -webkit-background-clip: text; }
.stat-value.red    { background: linear-gradient(135deg, #f43f5e, #be123c); -webkit-background-clip: text; }
.stat-label {
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--muted);
    margin-top: 0.3rem;
}

/* ── Success banner ── */
.success-banner {
    background: linear-gradient(135deg, rgba(16,185,129,0.08), rgba(6,182,212,0.05));
    border: 1px solid rgba(16,185,129,0.20);
    border-radius: var(--radius);
    padding: 1rem 1.4rem;
    display: flex;
    align-items: center;
    gap: 0.8rem;
    font-size: 0.95rem;
    font-weight: 500;
    color: #065f46;
    margin-bottom: 1.5rem;
}

/* ── Section heading ── */
.section-title {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--muted);
    margin: 2rem 0 0.8rem;
}

/* ── Download cards ── */
.dl-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1rem;
    margin-bottom: 1.5rem;
}
.dl-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.3rem 1.2rem;
    backdrop-filter: blur(12px);
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.04);
}
.dl-card-header {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin-bottom: 0.5rem;
}
.dl-card-icon { font-size: 1.5rem; }
.dl-card-title { font-size: 0.95rem; font-weight: 700; }
.dl-card-desc  { font-size: 0.8rem; color: var(--muted); margin-bottom: 1rem; }

/* ── Streamlit button overrides ── */
.stDownloadButton > button {
    width: 100% !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    padding: 0.55rem 1rem !important;
    border: none !important;
    transition: all 0.25s !important;
    letter-spacing: 0.02em !important;
}
.stDownloadButton:first-of-type > button {
    background: linear-gradient(135deg, #10b981, #059669) !important;
    color: white !important;
    box-shadow: 0 4px 15px rgba(16,185,129,0.2) !important;
}
.stDownloadButton:first-of-type > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(16,185,129,0.3) !important;
}
.stDownloadButton:last-of-type > button {
    background: linear-gradient(135deg, #f43f5e, #e11d48) !important;
    color: white !important;
    box-shadow: 0 4px 15px rgba(244,63,94,0.2) !important;
}
.stDownloadButton:last-of-type > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(244,63,94,0.3) !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border-radius: var(--radius) !important;
    overflow: hidden !important;
    border: 1px solid var(--border) !important;
}

/* ── Spinner / alerts ── */
.stAlert {
    border-radius: var(--radius) !important;
    border: 1px solid rgba(0,0,0,0.05) !important;
    background: rgba(255,255,255,0.7) !important;
}

/* ── Footer ── */
.app-footer {
    text-align: center;
    color: var(--muted);
    font-size: 0.78rem;
    margin-top: 3rem;
    padding-top: 1.5rem;
    border-top: 1px solid rgba(0,0,0,0.06);
}
.app-footer span { color: var(--accent1); }

/* ── Expander ── */
[data-testid="stExpander"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    backdrop-filter: blur(12px) !important;
}

/* ── Tab ── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(0,0,0,0.03) !important;
    border-radius: 12px !important;
    padding: 4px !important;
    gap: 4px !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px !important;
    color: var(--muted) !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
}
.stTabs [aria-selected="true"] {
    background: rgba(124,58,237,0.12) !important;
    color: var(--accent1) !important;
}
</style>
""", unsafe_allow_html=True)

# ── Hero ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-badge">✦ Powered by Mfolks_Tech_Team</div>
    <h1>MFolks Excel Cleaner</h1>
    <p>Upload your Excel file and instantly split it into <strong>complete</strong> and <strong>incomplete</strong> rows — clean data in seconds.</p>
</div>
""", unsafe_allow_html=True)

# ── File uploader ─────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📂  Drag & drop your Excel file here, or click to browse",
    type=["xlsx"],
    help="Supports .xlsx format",
)

# ── Processing ────────────────────────────────────────────────────────────────
if uploaded_file:
    with st.spinner("🔍  Analysing your data..."):
        df = pd.read_excel(uploaded_file)

    complete_rows = df.dropna()
    missing_rows  = df[df.isnull().any(axis=1)]

    total    = len(df)
    complete = len(complete_rows)
    missing  = len(missing_rows)

    # ── Stats ─────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-value">{total}</div>
            <div class="stat-label">Total Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-value green">{complete}</div>
            <div class="stat-label">✓ Complete</div>
        </div>
        <div class="stat-card">
            <div class="stat-value red">{missing}</div>
            <div class="stat-label">⚠ Incomplete</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Success banner ────────────────────────────────────────────────────
    st.markdown("""
    <div class="success-banner">
        <span>✅</span>
        <span>File processed successfully — your download files are ready below.</span>
    </div>
    """, unsafe_allow_html=True)

    # ── Preview tabs ──────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📊 Data Preview</div>', unsafe_allow_html=True)
    tab1, tab2 = st.tabs([f"✅  Complete Rows ({complete})", f"⚠️  Incomplete Rows ({missing})"])
    with tab1:
        st.dataframe(complete_rows, use_container_width=True, height=220)
    with tab2:
        st.dataframe(missing_rows,  use_container_width=True, height=220)

    # ── Build in-memory excel files ────────────────────────────────────────
    buf_complete = io.BytesIO()
    buf_missing  = io.BytesIO()
    complete_rows.to_excel(buf_complete, index=False)
    missing_rows.to_excel(buf_missing,  index=False)
    buf_complete.seek(0)
    buf_missing.seek(0)

    # ── Download section ──────────────────────────────────────────────────
    st.markdown('<div class="section-title">⬇️ Download Files</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="dl-card">
            <div class="dl-card-header">
                <span class="dl-card-icon">✅</span>
                <span class="dl-card-title">Complete Rows</span>
            </div>
            <div class="dl-card-desc">Rows with no missing values — fully filled data ready for use.</div>
        </div>
        """, unsafe_allow_html=True)
        st.download_button(
            label="⬇️  Download Complete Rows",
            data=buf_complete,
            file_name="complete_rows.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col2:
        st.markdown("""
        <div class="dl-card">
            <div class="dl-card-header">
                <span class="dl-card-icon">⚠️</span>
                <span class="dl-card-title">Incomplete Rows</span>
            </div>
            <div class="dl-card-desc">Rows with one or more missing values — flagged for review.</div>
        </div>
        """, unsafe_allow_html=True)
        st.download_button(
            label="⬇️  Download Incomplete Rows",
            data=buf_missing,
            file_name="missing_rows.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-footer">
    Built with <span>♥</span> by <span>mfolks_Tech_Team</span>  &nbsp;·&nbsp;  MFolks Excel Cleaner
</div>
""", unsafe_allow_html=True)