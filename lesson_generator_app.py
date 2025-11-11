"""
Somerset Lesson Generator App (Stage 7.1v2)
Auto-detecting column version with fallback for empty or variant Excel schemas.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import os

# -------------------------------------------------
# CONFIGURATION
# -------------------------------------------------
DATA_PATH = Path("YTC_CRW_Semester_1_Master_Calendar_2025_v21_STAGE4_FINAL.xlsx")
EXPORT_DIR = Path("exports")
EXPORT_DIR.mkdir(exist_ok=True)

# -------------------------------------------------
# LOAD DATABASE
# -------------------------------------------------
@st.cache_data
def load_data():
    if not DATA_PATH.exists():
        st.error("❌ Could not find the Excel file. Upload it to this directory and refresh.")
        st.stop()

    df = pd.read_excel(DATA_PATH)
    df.columns = [str(c).strip() for c in df.columns]

    # Add missing key fields if needed
    for col in ["Lesson Summary", "Terms / Vocabulary", "GIG / Bell Ringer"]:
        if col not in df.columns:
            df[col] = ""

    # Auto-detect columns
    week_col = None
    for c in df.columns:
        if str(c).strip().lower() in ["week", "week #", "week number"]:
            week_col = c
            break
    if week_col is None:
        df["Week"] = 1
        week_col = "Week"

    title_col = None
    for c in df.columns:
        if str(c).strip().lower() in ["lesson title", "title"]:
            title_col = c
            break
    if title_col is None:
        df["Lesson Title"] = "Untitled Lesson"
        title_col = "Lesson Title"

    # Clean data
    df.fillna("", inplace=True)
    df = df.dropna(how="all")
    return df, week_col, title_col

def save_data(df):
    df.to_excel(DATA_PATH, index=False)

df, week_col, title_col = load_data()

# -------------------------------------------------
# STREAMLIT INTERFACE
# -------------------------------------------------
st.set_page_config(page_title="Somerset Lesson Generator", layout="wide")
st.title("📘 Somerset Lesson Generator (v21)")
st.caption("View, edit, and generate Somerset-formatted lesson plans. Auto-saves every edit.")

# Sidebar Filters
weeks = sorted(df[week_col].dropna().unique().tolist()) or [1]
selected_week = st.sidebar.selectbox("Select Week", weeks)

week_df = df[df[week_col] == selected_week]
titles = sorted(week_df[title_col].dropna().unique().tolist()) or ["Untitled Lesson"]
selected_title = st.sidebar.selectbox("Select Lesson", titles)

# Find selected row
selected_row = week_df[week_df[title_col] == selected_title].head(1)
if selected_row.empty:
    st.warning("No lesson selected. You can add a new lesson below.")
else:
    idx = selected_row.index[0]

# -------------------------------------------------
# ADD NEW LESSON
# -------------------------------------------------
if st.sidebar.button("➕ Add New Lesson"):
    new_row = {c: "" for c in df.columns}
    new_row[week_col] = selected_week
    new_row["Date"] = datetime.now().strftime("%B %d, %Y")
    new_row[title_col] = f"New Lesson {datetime.now().strftime('%H%M%S')}"
    new_row["Lesson Status"] = "Planned"
    if "Day # (Continuous)" in df.columns:
        new_row["Day # (Continuous)"] = len(df) + 1
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(df)
    st.success("✅ New lesson added. Refresh to edit it.")

# -------------------------------------------------
# EDIT FORM
# -------------------------------------------------
if not selected_row.empty:
    st.subheader(f"✏️ Editing: {selected_title}")
    form_cols = st.columns(2)

    with form_cols[0]:
        left_fields = [
            title_col, "Date", "NV Standard Code", "NV Standard Descriptor",
            "Lesson Objective(s)", "Essential Question", "Lesson Summary",
            "Instructional Strategies and Procedures", "Accommodations and Modifications Strategies",
            "GIG / Bell Ringer", "Closure / Exit Ticket"
        ]
        for field in left_fields:
            if field in df.columns:
                new_val = st.text_area(field, str(df.at[idx, field]), key=field)
                if new_val != str(df.at[idx, field]):
                    df.at[idx, field] = new_val
                    save_data(df)
                    st.toast(f"💾 Auto-saved: {field}")

    with form_cols[1]:
        right_fields = [
            "Materials / Resources", "Terms / Vocabulary", "Learning Evidence",
            "Tiered Differentiation", "Tech Tools", "Tech Purpose",
            "Formative Check", "Summative Assessment", "Reflection / Notes", "Lesson Status"
        ]
        for field in right_fields:
            if field in df.columns:
                new_val = st.text_area(field, str(df.at[idx, field]), key=field)
                if new_val != str(df.at[idx, field]):
                    df.at[idx, field] = new_val
                    save_data(df)
                    st.toast(f"💾 Auto-saved: {field}")

# -------------------------------------------------
# PDF GENERATION
# -------------------------------------------------
def format_list(text, numbered=False):
    if not isinstance(text, str) or not text.strip():
        return ""
    items = [x.strip() for x in text.replace("\n", ";").split(";") if x.strip()]
    if not items:
        return text
    if numbered:
        return "<br/>".join([f"{i+1}. {it}" for i, it in enumerate(items)])
    else:
        return "<br/>".join([f"• {it}" for it in items])

def generate_pdf(row):
    week = str(row.get(week_col, "")).strip()
    day = str(row.get("Day # (Continuous)", "")).strip()
    title = str(row.get(title_col, "")).strip()
    date = str(row.get("Date", "")).strip()
    now = datetime.now()
    year = now.year
    month_str = now.strftime("%b")
    day_str = now.strftime("%d")
    safe_title = "_".join(title.split())
    filename = f"W{week}D{day}_LessonPlan_{month_str}{day_str}_{year}_{safe_title}.pdf"

    week_folder = EXPORT_DIR / f"Week_{week}"
    week_folder.mkdir(exist_ok=True)
    pdf_path = week_folder / filename

    styles = getSampleStyleSheet()
    normal = ParagraphStyle('Normal', parent=styles['Normal'], fontName='Times-Roman', fontSize=11, leading=14)
    header = ParagraphStyle('Header', parent=styles['Heading1'], fontSize=14, alignment=1, spaceAfter=12)
    section = ParagraphStyle('Section', parent=styles['Heading3'], fontName='Times-Bold', fontSize=12, spaceBefore=6, spaceAfter=4)

    doc = SimpleDocTemplate(str(pdf_path), pagesize=LETTER, rightMargin=60, leftMargin=60, topMargin=60, bottomMargin=40)
    story = []

    story.append(Paragraph("Somerset Academy Sky Pointe • U.S. History 8", header))
    story.append(Paragraph(f"Week {week} – Day {day} | Date: {date}", normal))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Lesson Title: {title}", section))
    story.append(Spacer(1, 8))

    def add_section(name, value, bullet=False, numbered=False):
        if not isinstance(value, str) or not value.strip():
            return
        story.append(Paragraph(name, section))
        formatted = format_list(value, numbered) if bullet or numbered else value.replace("\n", "<br/>")
        story.append(Paragraph(formatted, normal))
        story.append(Spacer(1, 8))

    add_section("NV Standards", f"{row.get('NV Standard Code', '')}<br/>{row.get('NV Standard Descriptor', '')}")
    add_section("Lesson Objective", row.get("Lesson Objective(s)", ""))
    add_section("Essential Question", row.get("Essential Question", ""))
    add_section("Lesson Summary", row.get("Lesson Summary", ""))
    add_section("Instructional Strategies and Procedures", row.get("Instructional Strategies and Procedures", ""), numbered=True)
    add_section("Accommodations and Modifications Strategies", row.get("Accommodations and Modifications Strategies", ""), bullet=True)
    add_section("GIG / Bell Ringer", row.get("GIG / Bell Ringer", ""))
    add_section("Closure / Exit Ticket", row.get("Closure / Exit Ticket", ""))
    add_section("Materials / Resources", row.get("Materials / Resources", ""), bullet=True)
    add_section("Terms / Vocabulary", row.get("Terms / Vocabulary", ""), bullet=True)
    add_section("Learning Evidence", row.get("Learning Evidence", ""))
    add_section("Reflection / Notes", row.get("Reflection / Notes", ""))

    doc.build(story)
    return pdf_path

if st.sidebar.button("📄 Generate PDF") and not selected_row.empty:
    pdf_path = generate_pdf(selected_row.iloc[0].to_dict())
    st.success(f"✅ PDF generated: {pdf_path.name}")
    st.download_button("⬇️ Download PDF", open(pdf_path, "rb"), file_name=pdf_path.name)

# -------------------------------------------------
# LIVE PREVIEW TABLE
# -------------------------------------------------
st.divider()
st.subheader(f"📊 Lessons in Week {selected_week}")
preview_cols = [c for c in df.columns if c not in ['Last Updated (PT)', 'Source Version']]
st.dataframe(week_df[preview_cols], use_container_width=True)
