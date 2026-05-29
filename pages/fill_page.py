import streamlit as st
import pandas as pd
import io

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Fill Missing Values - MFolks Excel Cleaner",
    page_icon="🛠",
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

/* ── Card Styles ── */
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
.upload-icon {
    font-size: 3rem;
    margin-bottom: 0.8rem;
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

.dl-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.3rem 1.2rem;
    backdrop-filter: blur(12px);
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.04);
    margin-bottom: 1rem;
}
.dl-card-header {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin-bottom: 0.5rem;
}
.dl-card-icon { font-size: 1.5rem; }
.dl-card-title { font-size: 0.95rem; font-weight: 700; color: var(--text); }
.dl-card-desc  { font-size: 0.8rem; color: var(--muted); margin-bottom: 0.5rem; }

/* ── Stats row ── */
.stats-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
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

/* ── Button Overrides ── */
.stButton > button, .stDownloadButton > button {
    width: 100% !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    padding: 0.55rem 1rem !important;
    border: none !important;
    transition: all 0.25s !important;
    letter-spacing: 0.02em !important;
}
.stButton > button {
    background: linear-gradient(135deg, #7c3aed, #6d28d9) !important;
    color: white !important;
    box-shadow: 0 4px 15px rgba(124,58,237,0.2) !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(124,58,237,0.3) !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #10b981, #059669) !important;
    color: white !important;
    box-shadow: 0 4px 15px rgba(16,185,129,0.2) !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(16,185,129,0.3) !important;
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
    margin-top: 1rem;
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

/* ── Spinner / alerts ── */
.stAlert {
    border-radius: var(--radius) !important;
    border: 1px solid rgba(0,0,0,0.05) !important;
    background: rgba(255,255,255,0.7) !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border-radius: var(--radius) !important;
    overflow: hidden !important;
    border: 1px solid var(--border) !important;
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

/* ── Page Link overrides ── */
[data-testid="stPageLink"] a {
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    width: 100% !important;
    background: rgba(124, 58, 237, 0.08) !important;
    color: #6d28d9 !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 0.65rem 1rem !important;
    border: 1px solid rgba(124, 58, 237, 0.2) !important;
    text-decoration: none !important;
    transition: all 0.25s !important;
}
[data-testid="stPageLink"] a:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 20px rgba(124, 58, 237, 0.08) !important;
    background: rgba(124, 58, 237, 0.12) !important;
    color: #6d28d9 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Hero ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-badge">✦ Smart Data Filling</div>
    <h1>Fill Missing Values</h1>
    <p>Review columns with missing records and fill them interactively to create a clean, complete dataset.</p>
</div>
""", unsafe_allow_html=True)

st.page_link("app.py", label="⬅️ Return to Home Page", icon="🏠")
st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

# ── Session State Check ──────────────────────────────────────────────────────
if "df" not in st.session_state:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-icon">📂</div>
        <div class="upload-label">No Active Excel Session</div>
        <div class="upload-sublabel">Please upload your Excel file on the <strong>Home</strong> page first to use this utility.</div>
    </div>
    """, unsafe_allow_html=True)
else:
    df = st.session_state.df.copy()
    
    # Find columns with missing values
    missing_columns = df.columns[df.isnull().any()].tolist()
    
    if len(missing_columns) == 0:
        st.markdown("""
        <div class="success-banner">
            <span>✅</span>
            <span>All good! Your Excel sheet has no missing values.</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        total_missing = df.isnull().sum().sum()
        cols_missing = len(missing_columns)
        
        # ── Stats ─────────────────────────────────────────────────────────────
        st.markdown(f"""
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value red">{total_missing}</div>
                <div class="stat-label">Total Empty Cells</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{cols_missing}</div>
                <div class="stat-label">Columns to Review</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="section-title">📝 Missing Values Review</div>', unsafe_allow_html=True)
        
        fill_values = {}
        
        # Loop through all columns having missing values
        for col in missing_columns:
            missing_count = df[col].isnull().sum()
            sample_values = df[col].dropna().unique()[:4]
            sample_str = ", ".join([str(x) for x in sample_values]) if len(sample_values) > 0 else "None"
            
            # Use clean card description for preview info
            st.markdown(f"""
            <div class="dl-card" style="margin-bottom: 0px; border-bottom-left-radius: 0px; border-bottom-right-radius: 0px;">
                <div class="dl-card-header">
                    <span class="dl-card-icon">📌</span>
                    <span class="dl-card-title">{col}</span>
                </div>
                <div class="dl-card-desc" style="margin-bottom: 0px;">
                    <strong>Empty records:</strong> {missing_count} rows &nbsp;·&nbsp; <strong>Sample entries:</strong> <em>{sample_str}</em>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Place the text input beautifully inside the flow
            fill_value = st.text_input(
                f"Enter replacement value for column '{col}'",
                key=col,
                placeholder="Type fallback value here..."
            )
            fill_values[col] = fill_value
            st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)
            
        # ── Quick Clean: Multiple Phones/Emails ───────────────────────────────
        import re

        phone_multi_count = 0
        email_multi_count = 0

        for col in df.columns:
            if "phone" in col.lower() or "mobile" in col.lower():
                phone_multi_count += df[col].astype(str).str.contains(
                    r'[,;/|]',
                    regex=True,
                    na=False
                ).sum()

            if "email" in col.lower():
                email_multi_count += df[col].astype(str).str.contains(
                    r'[,;/|]',
                    regex=True,
                    na=False
                ).sum()

        if "phone_clean_success" in st.session_state:
            st.success(st.session_state.phone_clean_success)
            del st.session_state.phone_clean_success

        if "email_clean_success" in st.session_state:
            st.success(st.session_state.email_clean_success)
            del st.session_state.email_clean_success

        if phone_multi_count > 0 or email_multi_count > 0:
            st.markdown('<div class="section-title">🧹 Multi-Value Clean Up</div>', unsafe_allow_html=True)

            if phone_multi_count > 0:
                st.markdown(f"""
                <div class="dl-card">
                    <div class="dl-card-title">
                    📱 Multiple Phone Numbers Found
                    </div>
                    <div class="dl-card-desc">
                    {phone_multi_count} rows contain more than one phone number.
                    </div>
                </div>
                """, unsafe_allow_html=True)

                remove_phone_btn = st.button(
                    "📱 Keep Only First Phone Number"
                )
                if remove_phone_btn:
                    for col in df.columns:
                        if (
                            "phone" in col.lower()
                            or "mobile" in col.lower()
                            or "contact" in col.lower()
                        ):
                            df[col] = df[col].apply(
                                lambda x: re.split(r'[,;/|]', str(x))[0].strip() if pd.notnull(x) and str(x) not in ['nan', 'None'] else x
                            )
                    st.session_state.df = df
                    st.session_state.phone_clean_success = f"Removed extra phone numbers from {phone_multi_count} rows"
                    st.rerun()

            if email_multi_count > 0:
                st.markdown(f"""
                <div class="dl-card">
                    <div class="dl-card-title">
                    📧 Multiple Emails Found
                    </div>
                    <div class="dl-card-desc">
                    {email_multi_count} rows contain more than one email address.
                    </div>
                </div>
                """, unsafe_allow_html=True)

                remove_email_btn = st.button(
                    "📧 Keep Only First Email Address"
                )
                if remove_email_btn:
                    for col in df.columns:
                        if "email" in col.lower():
                            df[col] = df[col].apply(
                                lambda x: re.split(r'[,;/|]', str(x))[0].strip() if pd.notnull(x) and str(x) not in ['nan', 'None'] else x
                            )
                    st.session_state.df = df
                    st.session_state.email_clean_success = f"Removed extra emails from {email_multi_count} rows"
                    st.rerun()

            st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">⚙️ Process Cleaned Dataset</div>', unsafe_allow_html=True)
        
        # Action button
        if st.button("✨ Apply & Generate Clean Excel"):
            st.session_state.cleaned = True
            
            cleaned_df = df.copy()
            for col, value in fill_values.items():
                cleaned_df[col] = cleaned_df[col].fillna(value)
            
            st.session_state.cleaned_df = cleaned_df
            
            # Create Excel file in-memory
            output = io.BytesIO()
            cleaned_df.to_excel(output, index=False)
            output.seek(0)
            st.session_state.cleaned_excel_data = output.getvalue()

        # Persist the cleaned view and download options
        if st.session_state.get("cleaned") and "cleaned_df" in st.session_state:
            cleaned_df = st.session_state.cleaned_df
            excel_data = st.session_state.cleaned_excel_data
            
            st.markdown("""
            <div class="success-banner">
                <span>🎉</span>
                <span>Dataset cleaned successfully! Preview and download options are ready below.</span>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown('<div class="section-title">📊 Cleaned Data Preview</div>', unsafe_allow_html=True)
            st.dataframe(cleaned_df, use_container_width=True, height=220)
            
            st.markdown("<div style='margin-top: 1.5rem;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="⬇️ Download Cleaned Excel File",
                data=excel_data,
                file_name="cleaned_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-footer">
    Built with <span>♥</span> by <span>mfolks_Tech_Team</span>  &nbsp;·&nbsp;  MFolks Excel Cleaner
</div>
""", unsafe_allow_html=True)
