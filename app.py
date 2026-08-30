import streamlit as st
import json
import os
import re
from dotenv import load_dotenv
from google import genai
from docx_generator import generate_and_save_docx

load_dotenv()

st.set_page_config(page_title="CBSE & Junior Exam Paper Generator", page_icon="🎓", layout="wide")

st.title("🎓 Examination Paper Generator (Session 2026-27)")
st.caption("Jai Arihant International School | Automated CBSE (9-12) & Custom Junior (Upto 8th) Engine")

# ---------------- API Key & Logo Configuration ----------------
api_key = os.getenv("GEMINI_API_KEY")

with st.sidebar:
    st.header("🔑 Settings")
    sidebar_key = st.text_input("Gemini API Key:", value=api_key if api_key else "", type="password")
    if sidebar_key:
        api_key = sidebar_key.strip()
    
    st.markdown("---")
    st.subheader("🏫 School Logo")
    uploaded_logo = st.file_uploader("Upload School Logo (PNG / JPG)", type=["png", "jpg", "jpeg"])

# Auto-Logo detection logic
logo_temp_path = None
if uploaded_logo:
    with open("temp_logo.png", "wb") as f:
        f.write(uploaded_logo.getbuffer())
    logo_temp_path = "temp_logo.png"
elif os.path.exists("school_logo.png"):
    logo_temp_path = "school_logo.png"

# ---------------- Classes & Filtered Subjects ----------------
ALL_CLASSES = [
    "Nursery", "LKG", "UKG",
    "Class 1", "Class 2", "Class 3", "Class 4", "Class 5",
    "Class 6", "Class 7", "Class 8", "Class 9", "Class 10",
    "Class 11", "Class 12"
]

PRIMARY_MIDDLE_SUBJECTS = sorted([
    "English", "Hindi", "Mathematics", "Environmental Studies (EVS)", 
    "Science", "Social Science", "Computer Studies", "General Knowledge (GK)", "Sanskrit"
])

SECONDARY_SUBJECTS = sorted([
    "Information Technology (Code 402)", "Artificial Intelligence (Code 417)",
    "Mathematics (Standard - 041)", "Mathematics (Basic - 241)", "Science (086)", 
    "Social Science (087)", "English Language & Literature (184)", 
    "Hindi Course A (002)", "Hindi Course B (085)", "Sanskrit (122)"
])

SENIOR_SECONDARY_SUBJECTS = sorted([
    "Informatics Practices (Code No. 065)", "Computer Science (Code No. 083)",
    "Mathematics (041)", "Applied Mathematics (241)", "Physics (042)", 
    "Chemistry (043)", "Biology (044)", "Accountancy (055)", "Business Studies (054)", 
    "Economics (030)", "English Core (301)", "History (027)", 
    "Political Science (028)", "Geography (029)", "Physical Education (048)"
])

JUNIOR_Q_TYPES = [
    "Multiple Choice Questions (MCQ)",
    "Fill in the Blanks",
    "True / False",
    "Match the Following",
    "Very Short Answer (VSA)",
    "Short Answer (SA)",
    "Long Answer (LA)",
    "Picture / Passage Based Questions"
]

DEFAULT_JUNIOR_MARKS = {
    "Multiple Choice Questions (MCQ)": 1,
    "Fill in the Blanks": 1,
    "True / False": 1,
    "Match the Following": 2,
    "Very Short Answer (VSA)": 2,
    "Short Answer (SA)": 3,
    "Long Answer (LA)": 5,
    "Picture / Passage Based Questions": 4
}

# ---------------- Section 1: Exam Details ----------------
st.subheader("1. School & Examination Details")
col_sc1, col_sc2 = st.columns([2, 1])
with col_sc1:
    school_name = st.text_input("School / Institution Name", value="Jai Arihant International School")

col1, col2, col3 = st.columns(3)
with col1:
    class_level = st.selectbox("Select Class", ALL_CLASSES, index=10)

is_junior_class = class_level in ["Nursery", "LKG", "UKG", "Class 1", "Class 2", "Class 3", "Class 4", "Class 5", "Class 6", "Class 7", "Class 8"]

if is_junior_class:
    available_subjects = PRIMARY_MIDDLE_SUBJECTS
elif class_level in ["Class 9", "Class 10"]:
    available_subjects = SECONDARY_SUBJECTS
else:
    available_subjects = SENIOR_SECONDARY_SUBJECTS

with col2:
    subject = st.selectbox("Select Subject", available_subjects)

with col3:
    if is_junior_class:
        paper_standard = st.selectbox("Paper Standard", ["Standard School Level", "Easy / Activity Based", "Analytical / Olympiad"])
    else:
        paper_standard = st.selectbox(
            "Standard / Difficulty Level",
            ["PYQ (CBSE Board Previous Years Questions with Year Details)", "Standard CBSE (Medium)", "Basic / Easy", "Tough / Advanced"]
        )

# Marks and Time calculation
if is_junior_class:
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        total_target_marks = st.number_input("Target Total Marks", min_value=10, max_value=100, value=50, step=5)
    with col_m2:
        time_allowed = st.selectbox("Time Allowed", ["1 Hour", "1.5 Hours", "2 Hours", "2.5 Hours"], index=2)
else:
    # Official CBSE Locked Marks for Senior Classes (9 to 12)
    if "065" in subject or "083" in subject:
        total_target_marks = 70
        time_allowed = "3 Hours"
    elif "402" in subject or "417" in subject:
        total_target_marks = 50
        time_allowed = "2 Hours"
    elif "Physics" in subject or "Chemistry" in subject or "Biology" in subject:
        total_target_marks = 70
        time_allowed = "3 Hours"
    elif "Accountancy" in subject or "Business Studies" in subject or "Economics" in subject:
        total_target_marks = 80
        time_allowed = "3 Hours"
    else:
        total_target_marks = 80
        time_allowed = "3 Hours"

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.number_input("Maximum Marks (Locked as per CBSE SQP)", value=total_target_marks, disabled=True)
    with col_m2:
        st.text_input("Time Allowed (Locked as per CBSE SQP)", value=time_allowed, disabled=True)

syllabus = st.text_area(
    "✍️ Enter or Paste Syllabus / Topics / Chapters",
    placeholder="Paste syllabus topics or chapters here...",
    height=100
)

st.markdown("---")

# ---------------- Section 2: Blueprint Setup ----------------
st.subheader("2. Question Paper Blueprint")

junior_blueprint = {}

if is_junior_class:
    st.write("👉 **Select Question Types & Customize (At least 2 questions will be assigned per selected type):**")
    
    # Initialize session state for junior options
    for q_type in JUNIOR_Q_TYPES:
        if f"chk_{q_type}" not in st.session_state:
            st.session_state[f"chk_{q_type}"] = True
        if f"m_{q_type}" not in st.session_state:
            st.session_state[f"m_{q_type}"] = DEFAULT_JUNIOR_MARKS[q_type]
        if f"c_{q_type}" not in st.session_state:
            st.session_state[f"c_{q_type}"] = 0

    # Auto-Distribution Logic for Junior classes (At least 2 per active category)
    def auto_distribute_junior_marks(target_marks):
        selected_types = [t for t in JUNIOR_Q_TYPES if st.session_state.get(f"chk_{t}", False)]
        if not selected_types:
            return
        
        # Reset counts
        for t in JUNIOR_Q_TYPES:
            st.session_state[f"c_{t}"] = 0
        
        remaining = target_marks
        
        # Step 1: Guarantee at least 2 questions for each selected type
        for t in selected_types:
            m = st.session_state.get(f"m_{t}", DEFAULT_JUNIOR_MARKS[t])
            if remaining >= m * 2:
                st.session_state[f"c_{t}"] = 2
                remaining -= (m * 2)
            elif remaining >= m:
                st.session_state[f"c_{t}"] = 1
                remaining -= m

        # Step 2: Distribute leftover marks smoothly across selected types
        while remaining > 0:
            allocated = False
            for t in selected_types:
                m = st.session_state.get(f"m_{t}", DEFAULT_JUNIOR_MARKS[t])
                if remaining >= m:
                    st.session_state[f"c_{t}"] += 1
                    remaining -= m
                    allocated = True
            if not allocated:
                break

    # Blueprint Header
    cols_h = st.columns([1, 4, 2, 2, 2])
    cols_h[0].write("**Include**")
    cols_h[1].write("**Question Type**")
    cols_h[2].write("**Marks / Q**")
    cols_h[3].write("**No. of Qs**")
    cols_h[4].write("**Total Marks**")

    current_calculated = 0
    for q_type in JUNIOR_Q_TYPES:
        col_chk, col_label, col_m, col_c, col_tot = st.columns([1, 4, 2, 2, 2])
        with col_chk:
            enabled = st.checkbox("", key=f"chk_{q_type}", label_visibility="collapsed")
        with col_label:
            st.write(f"**{q_type}**" if enabled else f"~{q_type}~")
        with col_m:
            marks = st.selectbox(f"marks_sel_{q_type}", [1, 2, 3, 4, 5, 6], key=f"m_{q_type}", disabled=not enabled, label_visibility="collapsed")
        with col_c:
            count = st.number_input(f"count_in_{q_type}", min_value=0, max_value=50, key=f"c_{q_type}", disabled=not enabled, label_visibility="collapsed")

        row_total = (marks * count) if enabled else 0
        current_calculated += row_total
        with col_tot:
            st.write(f"**{row_total}**")

        if enabled and count > 0:
            junior_blueprint[q_type] = {"marks": marks, "count": count}

    col_btn1, col_btn2 = st.columns([2, 2])
    with col_btn1:
        if current_calculated == total_target_marks:
            st.success(f"✅ Marks Perfectly Matched: **{current_calculated} / {total_target_marks}**")
        else:
            st.warning(f"⚠️ Current Total: **{current_calculated}** | Target Total: **{total_target_marks}** (Diff: {total_target_marks - current_calculated})")
    with col_btn2:
        st.button("🪄 Auto-Distribute Marks (Min 2 Qs Per Type)", on_click=auto_distribute_junior_marks, args=(total_target_marks,))

else:
    # Senior CBSE Blueprint Preview (9 to 12)
    if "065" in subject or "083" in subject:
        st.info(f"📋 **Official CBSE SQP Pattern (37 Questions):** Sec A (21 Qs × 1M), Sec B (7 Qs × 2M), Sec C (4 Qs × 3M), Sec D (2 Case Studies × 4M), Sec E (3 Qs × 5M).")
    elif "402" in subject or "417" in subject:
        st.info("📋 **Official CBSE Skill Blueprint (21 Questions):** Section A (Objective - 24 Marks) & Section B (Subjective - 26 Marks).")
    else:
        st.info(f"📋 **Official CBSE SQP Structure:** Sections A to E with MCQs, Assertion-Reason, VSA, SA, LA & Case Studies.")

st.markdown("---")

# ---------------- Senior & Junior AI Prompt Engines ----------------
def generate_junior_paper(client, school_name, class_level, subject, total_marks, time_allowed, syllabus, blueprint_dict, paper_standard):
    bp_lines = []
    for q_type, cfg in blueprint_dict.items():
        bp_lines.append(f"- Section '{q_type}': Exactly {cfg['count']} questions, each worth {cfg['marks']} marks.")
    bp_text = "\n".join(bp_lines)

    prompt = f"""
    You are an expert school examination setter for {class_level}, Subject: {subject}.
    Create a complete question paper based strictly on this syllabus:
    "{syllabus}"

    Difficulty Standard: {paper_standard}
    Total Marks: {total_marks} | Time: {time_allowed}

    Required Sections & Questions:
    {bp_text}

    Strict Rules:
    - Every single question must be 100% unique and distinct.
    - If MCQ: Provide exactly 4 options labeled a, b, c, d.
    - If Fill in the blanks: Include '________'.
    - If Match the Following: Provide two matching columns (Column A and Column B).
    - If Picture/Passage: Provide a descriptive scenario/short passage followed by questions.
    - Language and difficulty must be appropriate for {class_level} students.

    Return ONLY a valid JSON object matching this schema with NO markdown backticks:
    {{
      "general_instructions": [
        "1. All questions are compulsory.",
        "2. Read each question carefully before attempting.",
        "3. Write answers neatly and legibly."
      ],
      "sections": [
        {{
          "section_header": "SECTION NAME",
          "guidelines": ["Marks details"],
          "questions": [
            {{
              "q_no": "Q1.",
              "question_text": "Question statement here...",
              "marks_text": "[1 Mark]",
              "options": {{"a": "Option 1", "b": "Option 2", "c": "Option 3", "d": "Option 4"}}
            }}
          ]
        }}
      ]
    }}
    """
    model_list = ["gemini-3.6-flash", "gemini-3-flash", "gemini-2.5-flash"]
    for m in model_list:
        try:
            resp = client.models.generate_content(model=m, contents=prompt)
            raw = resp.text.strip()
            raw = re.sub(r"^```json\s*", "", raw, flags=re.MULTILINE)
            raw = re.sub(r"^```\s*", "", raw, flags=re.MULTILINE)
            raw = raw.strip("` \n\r")
            data = json.loads(raw)
            if "sections" in data:
                return data
        except Exception:
            continue
    raise Exception("Failed to generate junior paper. Please check API Key.")

def generate_senior_cbse_paper(client, school_name, class_level, subject, total_marks, time_allowed, syllabus, paper_standard):
    is_pyq = "PYQ" in paper_standard
    pyq_mandate = "For every question and sub-question, provide the authentic CBSE Board source in 'pyq_tag' (e.g. 'CBSE 2024 Delhi Set-1', 'CBSE 2023 All India', 'CBSE 2022 Term-2')." if is_pyq else ""
    
    if "065" in subject or "083" in subject:
        prompt = f"""
        You are an official CBSE Board Examination Paper Setter for {class_level}, Subject: {subject}.
        Generate the COMPLETE 37-Question CBSE Board Paper based strictly on this syllabus:
        "{syllabus}"

        Total Marks: 70 | Time: 3 Hours | Standard: {paper_standard}
        {pyq_mandate}

        STRICT CBSE OFFICIAL SQP BLUEPRINT FOR {subject}:
        Total Questions: 37 (Q1 to Q37 across Sections A to E).
        - SECTION A (Q1 to Q21): 21 Questions. Q1-Q19 MCQs; Q20-Q21 Assertion & Reasoning. [1 Mark each].
        - SECTION B (Q22 to Q28): 7 Questions of 2 marks each (Python output/error finding). Internal choice 'OR' in 2 questions.
        - SECTION C (Q29 to Q32): 4 Questions of 3 marks each (Python logic/SQL queries). Internal choice 'OR' in 1 question.
        - SECTION D (Q33 to Q34): 2 Case Study / Application based questions of 4 marks each with sub-questions.
        - SECTION E (Q35 to Q37): 3 Long Answer questions of 5 marks each (Complete Python programs/SQL schemas).

        Return ONLY a valid JSON object matching the standard CBSE schema with NO markdown backticks.
        """
    elif "402" in subject or "417" in subject:
        prompt = f"""
        You are an official CBSE Board Examination Paper Setter for {class_level}, Subject: {subject}.
        Generate the complete 21-Question examination paper based strictly on this syllabus:
        "{syllabus}"

        Total Marks: 50 | Time: 2 Hours | Standard: {paper_standard}
        {pyq_mandate}

        STRICT IT-402/AI-417 BLUEPRINT:
        Total Questions: 21 (Q1 to Q21).
        Section A: OBJECTIVE (24 Marks) -> Q1 (Employability 4/6), Q2-Q5 (Specific Skills 5/6 each).
        Section B: SUBJECTIVE (26 Marks) -> Q6-Q10 (Employability 3/5 [2M]), Q11-Q16 (Specific Skills 4/6 [2M]), Q17-Q21 (Specific Skills 3/5 [4M]).
        
        Return ONLY a valid JSON object with NO markdown backticks.
        """
    else:
        prompt = f"""
        You are an official CBSE Board Examination Paper Setter for {class_level}, Subject: {subject}.
        Generate a complete official CBSE question paper based strictly on this syllabus:
        "{syllabus}"

        Total Marks: {total_marks}, Time Allowed: {time_allowed}, Standard: {paper_standard}
        {pyq_mandate}

        STRICT CBSE 5-SECTION BLUEPRINT (Sections A to E, Assertion-Reason, VSA, SA, LA & Case Studies).
        Return ONLY a valid JSON object with NO markdown backticks.
        """

    model_list = ["gemini-3.6-flash", "gemini-3-flash", "gemini-2.5-flash"]
    for m in model_list:
        try:
            resp = client.models.generate_content(model=m, contents=prompt)
            raw = resp.text.strip()
            raw = re.sub(r"^```json\s*", "", raw, flags=re.MULTILINE)
            raw = re.sub(r"^```\s*", "", raw, flags=re.MULTILINE)
            raw = raw.strip("` \n\r")
            data = json.loads(raw)
            if "sections" in data:
                return data
        except Exception:
            continue
    raise Exception("Failed to generate senior CBSE paper. Please check API Key.")

# ---------------- Action Button ----------------
if st.button("🚀 Generate Examination Paper & Export DOCX", type="primary", use_container_width=True):
    if not api_key:
        st.error("❌ Gemini API Key is missing! Enter it in the sidebar.")
    elif not syllabus.strip():
        st.error("❌ Please enter the syllabus or topics in the box above.")
    elif is_junior_class and not junior_blueprint:
        st.error("❌ Please select at least one question type for the junior paper.")
    elif is_junior_class and current_calculated != total_target_marks:
        st.error(f"❌ Junior Blueprint Total ({current_calculated}) must match Target Total Marks ({total_target_marks}). Use Auto-Distribute button.")
    else:
        try:
            client = genai.Client(api_key=api_key.strip())
            
            with st.spinner(f"🧠 Generating unique examination paper for {class_level} - {subject}..."):
                if is_junior_class:
                    paper_data = generate_junior_paper(
                        client=client,
                        school_name=school_name,
                        class_level=class_level,
                        subject=subject,
                        total_marks=total_target_marks,
                        time_allowed=time_allowed,
                        syllabus=syllabus,
                        blueprint_dict=junior_blueprint,
                        paper_standard=paper_standard
                    )
                else:
                    paper_data = generate_senior_cbse_paper(
                        client=client,
                        school_name=school_name,
                        class_level=class_level,
                        subject=subject,
                        total_marks=total_target_marks,
                        time_allowed=time_allowed,
                        syllabus=syllabus,
                        paper_standard=paper_standard
                    )

            gen_instructions = paper_data.get("general_instructions", [])
            sections_list = paper_data.get("sections", [])

            filepath, filename = generate_and_save_docx(
                school_name=school_name,
                class_level=class_level,
                subject=subject,
                total_marks=total_target_marks,
                time_allowed=time_allowed,
                syllabus=syllabus,
                sections_list=sections_list,
                general_instructions=gen_instructions,
                logo_path=logo_temp_path
            )

            st.success(f"🎉 Exam Paper generated successfully and saved to: **`{filepath}`**")

            with open(filepath, "rb") as f:
                st.download_button(
                    label="📥 Download Word Document (.docx)",
                    data=f,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

            # ---------------- Live Screen Preview ----------------
            st.markdown("---")
            st.markdown(
                f"""
                <div style="border: 2px solid #333; padding: 25px; border-radius: 6px; background-color: #ffffff; color: #111; font-family: 'Times New Roman', serif;">
                    <div style="text-align: center;">
                        <h2 style="margin: 0; font-size: 22px; font-weight: bold;">{school_name.upper()}</h2>
                        <h4 style="margin: 5px 0; font-size: 16px;">EXAMINATION (SESSION 2026-2027)</h4>
                        <h3 style="margin: 5px 0; color: #1a365d; font-size: 18px;">{class_level.upper()} — {subject.upper()}</h3>
                        <p style="margin: 5px 0; font-size: 14px;"><b>Maximum Marks:</b> {total_target_marks} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Time Allowed:</b> {time_allowed.upper()}</p>
                    </div>
                    <hr style="margin: 15px 0;">
                    <p style="margin: 0; font-size: 15px; font-weight: bold;">General Instructions:</p>
                    <ul style="margin-top: 5px; font-size: 14px; padding-left: 20px;">
                        {''.join(f'<li>{inst}</li>' for inst in gen_instructions)}
                    </ul>
                </div>
                """,
                unsafe_allow_html=True
            )

            for sec in sections_list:
                st.markdown(f"### {sec.get('section_header')}")
                for gl in sec.get("guidelines", []):
                    st.caption(f"*{gl}*")

                for q in sec.get("questions", []):
                    q_num = q.get("q_no", "")
                    q_h = q.get("instruction_header", "")
                    q_t = q.get("question_text", "")
                    q_m = q.get("marks_text", "")
                    q_pyq = q.get("pyq_tag", "")

                    tag_display = f" <span style='color:#c53030; font-weight:bold;'>[{q_pyq}]</span>" if q_pyq else ""

                    if q_h:
                        st.markdown(f"**{q_num} {q_h}** `{q_m}`")
                    if q_t:
                        st.markdown(f"**{q_num}** {q_t}{tag_display} `{q_m}`", unsafe_allow_html=True)

                    if q.get("options") and isinstance(q["options"], dict):
                        cols = st.columns(len(q["options"]))
                        for o_idx, (k, val) in enumerate(q["options"].items()):
                            with cols[o_idx]:
                                st.write(f"**({k})** {val}")

                    for idx, sub in enumerate(q.get("sub_items", []), 1):
                        s_tag = f" <span style='color:#c53030; font-weight:bold;'>[{sub.get('pyq_tag')}]</span>" if sub.get('pyq_tag') else ""
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**({chr(104+idx)})** {sub.get('text')}{s_tag}", unsafe_allow_html=True)
                        opts = sub.get("options")
                        if opts and isinstance(opts, dict):
                            cols = st.columns(len(opts))
                            for o_idx, (k, val) in enumerate(opts.items()):
                                with cols[o_idx]:
                                    st.write(f"({k}) {val}")
                st.markdown("---")

        except Exception as err:
            st.error(f"Error: {str(err)}")