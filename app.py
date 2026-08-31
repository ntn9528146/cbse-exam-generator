import streamlit as st
import os
import json
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai

# Page Config
st.set_page_config(page_title="CBSE AI Exam Paper Generator", layout="wide", page_icon="🎓")

# ----------------- GEMINI API INITIALIZATION -----------------
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]

if api_key:
    genai.configure(api_key=api_key)
else:
    with st.sidebar:
        st.warning("⚠️ API Key nahi mili! Yahan manual enter karein:")
        api_key = st.text_input("Enter Gemini API Key", type="password")
        if api_key:
            genai.configure(api_key=api_key)

def call_gemini_with_fallback(prompt):
    """Multiple model candidates check karega taaki 404 error na aaye"""
    candidate_models = [
        "gemini-2.5-flash",
        "gemini-1.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-pro"
    ]
    
    last_error = None
    for m_name in candidate_models:
        try:
            model = genai.GenerativeModel(m_name)
            response = model.generate_content(prompt)
            if response and response.text:
                return response.text
        except Exception as e:
            last_error = e
            continue
            
    raise Exception(f"AI Generation Error: Sabhi available models try kiye gaye. Last Error: {last_error}")

# ----------------- DOCX GENERATOR HELPER -----------------
def create_docx_stream(school_name, class_level, subject, total_marks, time_allowed, syllabus, generated_sections):
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # School Header
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(school_name.upper())
    r_title.bold = True
    r_title.font.size = Pt(16)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"EXAMINATION PAPER | {class_level.upper()} | SUBJECT: {subject.upper()}")
    r_sub.bold = True
    r_sub.font.size = Pt(12)

    # Info Table
    info_table = doc.add_table(rows=1, cols=2)
    info_table.autofit = False
    info_table.columns[0].width = Inches(3.5)
    info_table.columns[1].width = Inches(3.5)
    
    info_table.cell(0, 0).paragraphs[0].text = f"Time Allowed: {time_allowed}"
    p_right = info_table.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.text = f"Maximum Marks: {total_marks}"

    if syllabus.strip():
        syl_p = doc.add_paragraph()
        r_syl = syl_p.add_run(f"Syllabus / Units Covered: {syllabus.strip()}")
        r_syl.italic = True
        r_syl.font.size = Pt(9.5)

    inst_p = doc.add_paragraph()
    inst_run = inst_p.add_run("General Instructions:\n1. All questions are compulsory.\n2. Marks allocated to questions are mentioned against each.")
    inst_run.bold = True
    inst_p.paragraph_format.space_after = Pt(10)

    # Write Sections & Questions
    q_num = 1
    for sec in generated_sections:
        sec_name = sec.get("section_name", "Section")
        sec_marks = sec.get("marks_per_q", 1)
        questions = sec.get("questions", [])
        
        if not questions:
            continue

        head = doc.add_paragraph()
        h_run = head.add_run(f"\n{sec_name.upper()} ({sec_marks} Mark Each)")
        h_run.bold = True
        h_run.font.size = Pt(11)

        for q in questions:
            qp = doc.add_paragraph()
            qp.paragraph_format.left_indent = Inches(0.2)
            qp.add_run(f"Q{q_num}. {q.get('question_text', '')}   ").bold = False
            qp.add_run(f"[{sec_marks}]").bold = True

            options = q.get("options")
            if options and isinstance(options, dict):
                for k, v in options.items():
                    op = doc.add_paragraph()
                    op.paragraph_format.left_indent = Inches(0.4)
                    op.add_run(f"({k}) {v}")
            q_num += 1

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ----------------- UI SETUP -----------------
st.title("🎓 Automated CBSE Exam Paper Generator")

ALL_CLASSES = [
    "Nursery", "LKG", "UKG",
    "Class 1", "Class 2", "Class 3", "Class 4", "Class 5",
    "Class 6", "Class 7", "Class 8", "Class 9", "Class 10",
    "Class 11", "Class 12"
]

CBSE_SUBJECTS = sorted([
    "English", "Hindi", "Mathematics", "Science", "Social Science", 
    "Environmental Studies (EVS)", "Computer Applications", "Information Technology (IT - 402)",
    "Artificial Intelligence (AI - 417)", "Physics", "Chemistry", "Biology",
    "Accountancy", "Business Studies", "Economics", "Computer Science (CS - 083)",
    "Informatics Practices (IP - 065)", "History", "Political Science", "Geography",
    "Sociology", "Psychology", "Physical Education", "Sanskrit", "Applied Mathematics",
    "Painting", "Home Science", "Music", "General Knowledge (GK)"
])

ALL_Q_TYPES = [
    "Multiple Choice Questions (MCQ)",
    "Fill in the Blanks",
    "True / False",
    "Very Short Answer (VSA)",
    "Short Answer (SA)",
    "Long Answer (LA)",
    "Extract / Case Based"
]

# Section 1: Exam Header Info
with st.expander("📌 1. School & Examination Details", expanded=True):
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        school_name = st.text_input("School / Institute Name", value="DELHI PUBLIC SCHOOL")
    with col_s2:
        time_allowed = st.selectbox("Time Allowed", ["1 Hour", "1.5 Hours", "2 Hours", "2.5 Hours", "3 Hours"], index=4)

    col1, col2, col3 = st.columns(3)
    with col1:
        class_level = st.selectbox("Select Class", ALL_CLASSES, index=14)
    with col2:
        subject = st.selectbox("Select Subject", CBSE_SUBJECTS, index=CBSE_SUBJECTS.index("Mathematics"))
    with col3:
        total_target_marks = st.number_input("Target Total Marks", min_value=10, max_value=150, value=80, step=5)

    syllabus = st.text_area(
        "✍️ Enter / Paste Syllabus or Chapters", 
        placeholder="Yahan chapters aur topics type ya paste karein (e.g. Chapter 1: Real Numbers, Chapter 2: Polynomials, Triangles)...",
        height=100
    )

st.markdown("---")
st.subheader("⚙️ 2. Question Paper Blueprint")

# Session state init
for q_type in ALL_Q_TYPES:
    if f"enabled_{q_type}" not in st.session_state:
        st.session_state[f"enabled_{q_type}"] = True
    if f"marks_{q_type}" not in st.session_state:
        st.session_state[f"marks_{q_type}"] = 1 if "MCQ" in q_type or "Fill" in q_type or "True" in q_type else (2 if "VSA" in q_type else (3 if "SA" in q_type else (5 if "LA" in q_type else 4)))
    if f"count_{q_type}" not in st.session_state:
        st.session_state[f"count_{q_type}"] = 0

cols_h = st.columns([1, 4, 2, 2, 2])
cols_h[0].write("**Include**")
cols_h[1].write("**Question Type**")
cols_h[2].write("**Marks / Q**")
cols_h[3].write("**No. of Qs**")
cols_h[4].write("**Section Total**")

calculated_total = 0
active_blueprint = {}

for q_type in ALL_Q_TYPES:
    col_chk, col_label, col_m, col_c, col_tot = st.columns([1, 4, 2, 2, 2])
    
    with col_chk:
        enabled = st.checkbox("", key=f"enabled_{q_type}", label_visibility="collapsed")
    with col_label:
        st.write(f"**{q_type}**" if enabled else f"~{q_type}~")
    with col_m:
        marks = st.selectbox(f"m_{q_type}", [1, 2, 3, 4, 5, 6, 8], key=f"marks_{q_type}", disabled=not enabled, label_visibility="collapsed")
    with col_c:
        count = st.number_input(f"c_{q_type}", min_value=0, max_value=100, key=f"count_{q_type}", disabled=not enabled, label_visibility="collapsed")
        
    sec_total = (marks * count) if enabled else 0
    calculated_total += sec_total
    
    with col_tot:
        st.write(f"**{sec_total}**")
        
    if enabled and count > 0:
        active_blueprint[q_type] = {"marks": marks, "count": count}

st.markdown("---")

col_b1, col_b2 = st.columns([2, 2])
with col_b1:
    if calculated_total == total_target_marks:
        st.success(f"✅ Total Marks: **{calculated_total} / {total_target_marks}** (Matched)")
    elif calculated_total < total_target_marks:
        st.warning(f"⚠️ Current Total: **{calculated_total}** | Target: **{total_target_marks}** (Remaining: **{total_target_marks - calculated_total} Marks**)")
    else:
        st.error(f"❌ Current Total: **{calculated_total}** exceeds Target: **{total_target_marks}** by **{calculated_total - total_target_marks} Marks**")

with col_b2:
    if st.button("🪄 Auto-Calculate & Balance Remaining Questions"):
        diff = total_target_marks - calculated_total
        if diff <= 0:
            st.info("Remaining marks 0 hain ya exceed ho rahe hain.")
        else:
            open_types = [t for t in ALL_Q_TYPES if st.session_state[f"enabled_{t}"] and st.session_state[f"count_{t}"] == 0]
            if not open_types:
                open_types = [t for t in ALL_Q_TYPES if st.session_state[f"enabled_{t}"]]
            
            for t in reversed(open_types):
                m = st.session_state[f"marks_{t}"]
                if diff >= m:
                    added_count = diff // m
                    st.session_state[f"count_{t}"] += added_count
                    diff = diff % m
            st.rerun()

st.markdown("---")

# ----------------- GENERATION BUTTON -----------------
if st.button("🚀 Generate Examination Paper & Export DOCX", type="primary"):
    if not syllabus.strip():
        st.error("Kripya syllabus ya chapters ka text paste karein!")
    elif calculated_total != total_target_marks:
        st.error(f"Marks mismatch! Blueprint Total ({calculated_total}) Target Marks ({total_target_marks}) ke barabar hona chahiye.")
    elif not active_blueprint:
        st.error("Kripya kam se kam ek question type select karein jisme No. of Questions > 0 ho.")
    else:
        with st.spinner("AI Question Paper generate kar raha hai, kripya wait karein..."):
            prompt = f"""
            You are a senior CBSE Paper Setter for {class_level} - {subject}.
            Create an official exam paper based STRICTLY on the following syllabus and blueprint.

            SYLLABUS:
            {syllabus}

            BLUEPRINT:
            {json.dumps(active_blueprint, indent=2)}

            Output must be a valid JSON array of section objects with this exact structure:
            [
              {{
                "section_name": "Section A: Multiple Choice Questions",
                "marks_per_q": 1,
                "questions": [
                  {{
                    "question_text": "Question text here...",
                    "options": {{"A": "...", "B": "...", "C": "...", "D": "..."}} (only for MCQ, otherwise null)
                  }}
                ]
              }}
            ]
            Return ONLY raw JSON. No markdown ticks ```json ... ```.
            """
            
            try:
                raw_response = call_gemini_with_fallback(prompt)
                clean_json = raw_response.strip()
                if clean_json.startswith("```json"):
                    clean_json = clean_json[7:]
                if clean_json.startswith("```"):
                    clean_json = clean_json[3:]
                if clean_json.endswith("```"):
                    clean_json = clean_json[:-3]
                clean_json = clean_json.strip()

                sections_data = json.loads(clean_json)

                # Generate Word document in memory
                docx_stream = create_docx_stream(
                    school_name, class_level, subject, total_target_marks, time_allowed, syllabus, sections_data
                )

                st.success("✅ Examination Paper successfully generate ho gaya!")

                # Download Button
                st.download_button(
                    label="📥 Download Examination Paper (.docx)",
                    data=docx_stream,
                    file_name=f"{class_level}_{subject}_Exam_Paper.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                # Preview
                st.markdown("---")
                st.markdown(f"<div style='text-align:center;'><h2>{school_name.upper()}</h2><h4>{class_level.upper()} - {subject.upper()}</h4><p><b>Time Allowed:</b> {time_allowed} | <b>Max Marks:</b> {total_target_marks}</p></div>", unsafe_allow_html=True)
                st.markdown(f"**Syllabus:** *{syllabus}*")
                st.markdown("---")

                global_q_no = 1
                for sec in sections_data:
                    st.markdown(f"### {sec.get('section_name')} ({sec.get('marks_per_q')} Mark Each)")
                    for q in sec.get("questions", []):
                        st.write(f"**Q{global_q_no}.** {q.get('question_text')} `[{sec.get('marks_per_q')} Marks]`")
                        if q.get("options"):
                            opts = q.get("options")
                            cols = st.columns(len(opts))
                            for idx, (k, val) in enumerate(opts.items()):
                                with cols[idx]:
                                    st.write(f"**({k})** {val}")
                        global_q_no += 1
                    st.markdown("---")

            except Exception as ex:
                st.error(f"Error: {ex}")
