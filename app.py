import streamlit as st
import os


st.set_page_config(page_title="CBSE Exam Paper Generator", layout="wide")

st.title("🎓 CBSE Automated Exam Paper Generator")

# Class List: Nursery to 12th
ALL_CLASSES = [
    "Nursery", "LKG", "UKG",
    "Class 1", "Class 2", "Class 3", "Class 4", "Class 5",
    "Class 6", "Class 7", "Class 8", "Class 9", "Class 10",
    "Class 11", "Class 12"
]

# Complete CBSE Subject List
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

# 1. School & Exam Information
with st.expander("📌 School & Exam Details", expanded=True):
    col_sc1, col_sc2 = st.columns([2, 1])
    with col_sc1:
        school_name = st.text_input("School / Institution Name", value="DELHI PUBLIC SCHOOL")
    with col_sc2:
        time_allowed = st.selectbox("Time Allowed", ["1 Hour", "1.5 Hours", "2 Hours", "2.5 Hours", "3 Hours"], index=4)

    col1, col2, col3 = st.columns(3)
    with col1:
        class_level = st.selectbox("Select Class", ALL_CLASSES, index=14)
    with col2:
        subject = st.selectbox("Select Subject", CBSE_SUBJECTS, index=CBSE_SUBJECTS.index("Mathematics"))
    with col3:
        total_target_marks = st.number_input("Target Total Marks", min_value=10, max_value=150, value=80, step=5)

    syllabus = st.text_area("✍️ Enter / Paste Syllabus or Chapters", 
                            placeholder="Type or paste syllabus topics here (e.g. Unit 1: Calculus, Chapter 2: Matrices, Probability)...",
                            height=100)

st.markdown("---")
st.subheader("⚙️ Blueprint Configuration")

# Session state initialization for dynamic adjustments
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
        marks = st.selectbox(
            f"m_{q_type}", 
            [1, 2, 3, 4, 5, 6, 8], 
            key=f"marks_{q_type}", 
            disabled=not enabled, 
            label_visibility="collapsed"
        )
        
    with col_c:
        count = st.number_input(
            f"c_{q_type}", 
            min_value=0, 
            max_value=100, 
            key=f"count_{q_type}", 
            disabled=not enabled, 
            label_visibility="collapsed"
        )
        
    sec_total = (marks * count) if enabled else 0
    calculated_total += sec_total
    
    with col_tot:
        st.write(f"**{sec_total}**")
        
    if enabled and count > 0:
        active_blueprint[q_type] = {"marks": marks, "count": count}

# Auto-Calculate Remaining Marks Feature
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
    # Auto Balance Button
    if st.button("🪄 Auto-Calculate & Balance Remaining Questions"):
        diff = total_target_marks - calculated_total
        if diff <= 0:
            st.info("Bache hue marks nahi hain ya marks target se zyada hain.")
        else:
            # Distribute remaining marks in enabled subjective questions where count is currently 0
            open_types = [t for t in ALL_Q_TYPES if st.session_state[f"enabled_{t}"] and st.session_state[f"count_{t}"] == 0]
            if not open_types:
                open_types = [t for t in ALL_Q_TYPES if st.session_state[f"enabled_{t}"]]
            
            # Simple heuristic distribution
            for t in reversed(open_types):
                m = st.session_state[f"marks_{t}"]
                if diff >= m:
                    added_count = diff // m
                    st.session_state[f"count_{t}"] += added_count
                    diff = diff % m
            st.rerun()

st.markdown("---")

# Generate Paper Trigger
if st.button("📄 Generate Question Paper & Preview", type="primary"):
    if not syllabus.strip():
        st.error("Kripya syllabus ya chapters enter karein!")
    elif calculated_total != total_target_marks:
        st.error(f"Total Marks mismatch! Blueprint ({calculated_total}) target marks ({total_target_marks}) ke barabar hona chahiye.")
    elif not active_blueprint:
        st.error("Kripya kam se kam ek question type select karein jisme questions ki sankhya > 0 ho.")
    else:
        # Fetch Questions from DB
        sections_data = {}
        for q_type, cfg in active_blueprint.items():
            # Standardizing name for DB query matching
            db_type_map = {
                "Multiple Choice Questions (MCQ)": "MCQ",
                "Fill in the Blanks": "Fill in the blanks",
                "True / False": "True/False",
                "Very Short Answer (VSA)": "Very Short Answer",
                "Short Answer (SA)": "Short Answer",
                "Long Answer (LA)": "Long Answer",
                "Extract / Case Based": "Extract based question"
            }
            mapped_type = db_type_map.get(q_type, q_type)
            
            db_questions = fetch_questions(subject, syllabus, mapped_type, cfg["marks"], cfg["count"])
            
            # Fallback placeholder if DB has fewer questions
            final_qs = []
            for i in range(cfg["count"]):
                if i < len(db_questions):
                    final_qs.append(db_questions[i])
                else:
                    final_qs.append({
                        "question_text": f"Sample question for {subject} ({q_type} - Question #{i+1}).",
                        "marks": cfg["marks"],
                        "options": {"A": "Option 1", "B": "Option 2", "C": "Option 3", "D": "Option 4"} if "MCQ" in q_type else None
                    })
                    
            sections_data[q_type] = {
                "marks": cfg["marks"],
                "questions": final_qs
            }

        # Save to docx in folder
        filepath, filename = generate_and_save_docx(
            school_name, class_level, subject, total_target_marks, time_allowed, syllabus, sections_data
        )

        st.success(f"✅ Paper successfully generate ho gaya aur **`{filepath}`** me auto-save ho gaya!")

        # Download Button
        with open(filepath, "rb") as docx_file:
            st.download_button(
                label="📥 Download Word Document (.docx)",
                data=docx_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # On-Screen Preview
        st.markdown("## 📋 Exam Paper Preview")
        st.markdown(f"<div style='text-align:center;'><h2>{school_name.upper()}</h2><h4>{class_level.upper()} - {subject.upper()}</h4><p><b>Time:</b> {time_allowed} | <b>Max Marks:</b> {total_target_marks}</p></div>", unsafe_allow_html=True)
        st.markdown(f"**Syllabus:** *{syllabus}*")
        st.markdown("---")

        global_q_num = 1
        for sec_name, sdata in sections_data.items():
            st.markdown(f"### Section: {sec_name} ({sdata['marks']} Mark Each)")
            for q in sdata["questions"]:
                st.write(f"**Q{global_q_num}.** {q['question_text']} `[{sdata['marks']} Marks]`")
                if q.get("options"):
                    cols = st.columns(len(q["options"]))
                    for idx, (k, val) in enumerate(q["options"].items()):
                        with cols[idx]:
                            st.write(f"**({k})** {val}")
                global_q_num += 1
            st.markdown("---")
