import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
 
# ----------------------------
# SESSION STATE INITIALIZATION (MUST BE FIRST)
# ----------------------------
if 'sidebar_open' not in st.session_state:
    st.session_state.sidebar_open = False
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = None
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = {}
if 'selected_file' not in st.session_state:
    st.session_state.selected_file = None
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None
 
# ----------------------------
# PAGE CONFIG -- keep sidebar collapsed by default
# ----------------------------
st.set_page_config(
    page_title="Smart Excel Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed"
)
 
# ----------------------------
# Constants for CSS
# ----------------------------
SIDEBAR_WIDTH = 340
MAIN_LEFT_OPEN = SIDEBAR_WIDTH + 20
MAIN_LEFT_CLOSED = 20
 
sidebar_transform = "translateX(0)" if st.session_state.sidebar_open else f"translateX(-{SIDEBAR_WIDTH + 20}px)"
main_margin_left = f"{MAIN_LEFT_OPEN}px" if st.session_state.sidebar_open else f"{MAIN_LEFT_CLOSED}px"
toggle_icon = "✕" if st.session_state.sidebar_open else "☰"
 
# ----------------------------
# CUSTOM HEADER + FOOTER + STYLES
# ----------------------------
st.markdown(f"""
    <style>
        .custom-header {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            height: 70px;
            background-color: #ffffff;
            color: #222;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
            font-weight: 600;
            border-bottom: 1px solid #ddd;
            z-index: 9999;
        }}
        .custom-header img {{
            position: absolute;
            left: 20px;
            height: 45px;
        }}
 
        .custom-footer {{
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #f5f5f5;
            color: #555;
            text-align: center;
            padding: 8px 0;
            font-size: 14px;
            border-top: 1px solid #ddd;
            z-index: 9999999;
        }}
 
        .main .block-container {{
            padding-top: 100px !important;
            padding-bottom: 80px !important;
            margin-left: {main_margin_left} !important;
            transition: margin-left 0.25s ease;
        }}
 
        section[data-testid="stSidebar"] {{
            width: {SIDEBAR_WIDTH}px;
            min-width: {SIDEBAR_WIDTH}px;
            max-width: {SIDEBAR_WIDTH}px;
            position: fixed;
            top: 70px;
            left: 0;
            height: calc(100% - 70px - 40px);
            transform: {sidebar_transform};
            transition: transform 0.28s ease;
            z-index: 10000;
            box-shadow: 0 6px 18px rgba(0,0,0,0.12);
            background-color: black;
            overflow: auto;
            color: white !important;
        }}
 
        section[data-testid="stSidebar"] * {{ color: white !important; }}
        #MainMenu, footer, header {{visibility: hidden;}}
 
        .stButton > button.custom-toggle {{
            background-color: #E50914 !important;
            color: white !important;
            border-radius: 50% !important;
            width: 55px !important;
            height: 55px !important;
            font-size: 28px !important;
            font-weight: bold !important;
            position: fixed !important;
            top: 85px !important;
            left: 20px !important;
            z-index: 10001 !important;
        }}
    </style>
    <div class="custom-header">
        <img src="https://upload.wikimedia.org/wikipedia/commons/0/08/Netflix_2015_logo.svg" alt="Logo">
        Smart Excel Dashboard
    </div>
""", unsafe_allow_html=True)
 
# ----------------------------
# TOGGLE BUTTON
# ----------------------------
if st.button(toggle_icon, key="sidebar_toggle_btn"):
    st.session_state.sidebar_open = not st.session_state.sidebar_open
    st.rerun()
 
# ----------------------------
# SIDEBAR
# ----------------------------
with st.sidebar:
    st.markdown("### 📑 File & Sheet Selection")
    st.markdown("---")
   
    uploaded_files = st.file_uploader(
        "📂 Upload Excel files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )
   
    if uploaded_files:
        st.session_state.uploaded_files = uploaded_files
        for file in uploaded_files:
            try:
                xls = pd.ExcelFile(file)
                st.session_state.excel_data[file.name] = xls.sheet_names
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
   
    if st.session_state.uploaded_files and st.session_state.excel_data:
        selected_file = st.selectbox("Select File:", list(st.session_state.excel_data.keys()), key="file_selector")
        st.session_state.selected_file = selected_file
       
        if selected_file:
            selected_sheet = st.radio("Select Sheet:", st.session_state.excel_data[selected_file], key="sheet_selector")
            st.session_state.selected_sheet = selected_sheet
   
    st.markdown("---")
    st.markdown("### 💡 Tips")
    st.info("• Upload multiple Excel files\n• Select different sheets to analyze\n• Charts update automatically")
 
# ----------------------------
# MAIN PAGE CONTENT
# ----------------------------
st.title("📊 Smart Excel Dashboard Generator")
 
if not st.session_state.uploaded_files:
    st.info("👈 Click the red button on the left to open sidebar and upload Excel files.")
else:
    st.success("✅ Files uploaded - Select a file and sheet from the sidebar to view analytics.")
 
st.markdown("---")
 
if (st.session_state.uploaded_files and st.session_state.selected_file and st.session_state.selected_sheet):
    try:
        selected_file_obj = next((f for f in st.session_state.uploaded_files if f.name == st.session_state.selected_file), None)
        xls = pd.ExcelFile(selected_file_obj)
        raw_df = pd.read_excel(xls, sheet_name=st.session_state.selected_sheet, header=None).dropna(how="all")
 
        def find_header_row(df, threshold=0.5):
            for i, row in df.iterrows():
                if row.count() / len(row) >= threshold:
                    return i
            return 0
 
        header_row = find_header_row(raw_df)
        df = pd.read_excel(xls, sheet_name=st.session_state.selected_sheet, header=header_row).dropna(how="all")
 
        if any(str(c).startswith("Unnamed") for c in df.columns):
            try:
                temp = pd.read_excel(xls, sheet_name=st.session_state.selected_sheet, header=[header_row, header_row + 1])
                df.columns = [" ".join([str(a), str(b)]).replace("Unnamed", "").strip()
                              for a, b in zip(temp.columns.get_level_values(0), temp.columns.get_level_values(1))]
            except:
                pass
 
        def clean_column_names(cols):
            clean = []
            for i, col in enumerate(cols):
                col = re.sub(r'[^A-Za-z0-9 ]+', ' ', str(col)).title().strip()
                col = re.sub(r'\s+', ' ', col)
                if col == "" or col.startswith("Unnamed"):
                    col = f"Column {i+1}"
                clean.append(col)
            return clean
 
        df.columns = clean_column_names(df.columns)
 
        if df.empty:
            st.warning("⚠️ This sheet appears empty.")
        else:
            col1, col2, col3 = st.columns(3)
            with col1: st.metric("Total Rows", len(df))
            with col2: st.metric("Total Columns", len(df.columns))
            with col3: st.metric("File", st.session_state.selected_file)
           
            st.markdown("---")
           
            numeric_cols = df.select_dtypes(include="number").columns.tolist()
            categorical_cols = df.select_dtypes(exclude="number").columns.tolist()
 
            st.subheader("📈 Auto Insights Dashboard")
 
            if numeric_cols:
                st.markdown("### 🔢 Numeric Insights (Charts Only)")
                for num_col in numeric_cols:
                    st.markdown(f"#### 🔸 {num_col}")
                    fig_hist = px.histogram(df, x=num_col, nbins=20)
                    fig_line = px.line(df, y=num_col, markers=True)
                    avg_val = df[num_col].mean()
                    max_val = df[num_col].max()
                    fig_gauge = go.Figure(go.Indicator(
                        mode="gauge+number",
                        value=avg_val,
                        gauge={'axis': {'range': [0, max_val if pd.notna(max_val) else 1]}}
                    ))
                    col1, col2, col3 = st.columns(3)
                    with col1: st.plotly_chart(fig_hist, use_container_width=True)
                    with col2: st.plotly_chart(fig_line, use_container_width=True)
                    with col3: st.plotly_chart(fig_gauge, use_container_width=True)
 
            if categorical_cols:
                st.markdown("### 🗂️ Category Insights")
                for cat_col in categorical_cols:
                    vc = df[cat_col].value_counts().reset_index()
                    vc.columns = [cat_col, "Count"]
                    fig_bar = px.bar(vc, x=cat_col, y="Count")
                    fig_pie = px.pie(vc, names=cat_col, values="Count", hole=0.4)
                    col1, col2 = st.columns(2)
                    with col1: st.plotly_chart(fig_bar, use_container_width=True)
                    with col2: st.plotly_chart(fig_pie, use_container_width=True)
 
            if numeric_cols and categorical_cols:
                st.markdown("### 🔀 Numeric vs Category")
                cat_col = categorical_cols[0]
                for num_col in numeric_cols:
                    g = df.groupby(cat_col)[num_col].mean().reset_index()
                    fig = px.bar(g, x=cat_col, y=num_col)
                    st.plotly_chart(fig, use_container_width=True)
 
    except Exception as e:
        st.error(f"Error processing the selected sheet: {e}")
 
elif st.session_state.uploaded_files:
    st.info("📁 Files uploaded - Select a file and sheet from the sidebar to view analytics.")
 
st.markdown("""
    <div class="custom-footer">
        © 2025 Mayuresh | Built with ❤️ using Streamlit
    </div>
""", unsafe_allow_html=True)
