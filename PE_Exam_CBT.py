import streamlit as st
import openpyxl
import random
import re
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="PE Exam Simulator", page_icon="📝", layout="wide")


# --- CLASS DEFINITION ---
class Question:
    def __init__(self, text, options, correct_indices, q_id):
        self.text = text
        self.options = options
        self.correct_indices = set(correct_indices)
        self.id = q_id
        self.user_selections = set()
        self.flagged = False

    def is_correct(self):
        return self.user_selections == self.correct_indices

    def is_answered(self):
        return len(self.user_selections) > 0


# --- HELPER FUNCTIONS ---
def parse_excel(uploaded_file):
    """Parses the uploaded Excel file into Question objects."""
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    qs = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=2, values_only=False):
            # Column A: Question Text
            if not row[0].value: continue
            txt = str(row[0].value)

            # Extract ID if present in (parentheses) at end
            qid_match = re.search(r'\(([\w-]+)\)\s*$', txt)
            qid = qid_match.group(1) if qid_match else "ID"

            opts = []
            cor = []

            # Columns B-E: Options
            for i, c in enumerate(row[1:]):
                if c.value:
                    opts.append(str(c.value))
                    # Check for bold font indicating correct answer
                    if c.font and c.font.bold:
                        cor.append(i)

            if opts and cor:
                qs.append(Question(txt, opts, cor, qid))
    return qs


def generate_word_doc(questions):
    """Generates a Word document in memory."""
    doc = Document()
    doc.add_heading('PE Exam Set', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Generated ID: {random.randint(1000, 9999)}")

    # Questions Section
    for i, q in enumerate(questions):
        p = doc.add_paragraph()
        p.add_run(f"{i + 1}. {q.text}").bold = True
        lets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        for idx, o in enumerate(q.options):
            doc.add_paragraph(f"    {lets[idx] if idx < 26 else '?'}. {o}")
        doc.add_paragraph("")

    doc.add_page_break()

    # Answer Key Section
    doc.add_heading('Answer Key', 0)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Light Shading'
    tbl.rows[0].cells[0].text = 'Q#'
    tbl.rows[0].cells[1].text = 'Ans'

    for i, q in enumerate(questions):
        row = tbl.add_row().cells
        row[0].text = str(i + 1)
        # Convert indices back to letters (0 -> A, 1 -> B)
        lets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        ans_str = ", ".join([lets[x] for x in sorted(q.correct_indices)])
        row[1].text = ans_str

    # Save to memory buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- SESSION STATE INITIALIZATION ---
if 'questions_pool' not in st.session_state:
    st.session_state.questions_pool = []
if 'exam_session' not in st.session_state:
    st.session_state.exam_session = []  # The 20 active questions
if 'current_index' not in st.session_state:
    st.session_state.current_index = 0
if 'exam_stage' not in st.session_state:
    st.session_state.exam_stage = "SETUP"  # SETUP, ACTIVE, REVIEW, REPORT

# --- MAIN APP LOGIC ---

# 1. SIDEBAR: Controls & Navigation
with st.sidebar:
    st.header("Exam Controls")

    # Upload Button
    uploaded_file = st.file_uploader("1. Load Question Database", type=["xlsx"])
    if uploaded_file:
        try:
            questions = parse_excel(uploaded_file)
            st.session_state.questions_pool = questions
            st.success(f"Loaded {len(questions)} questions.")
        except Exception as e:
            st.error(f"Error parsing file: {e}")

    # Generate Word Doc Button
    if st.session_state.questions_pool:
        st.write("---")
        st.write("**Export to Word**")
        if st.button("Generate Word Doc"):
            # Create a sample of 20 (or less)
            subset = random.sample(st.session_state.questions_pool, min(20, len(st.session_state.questions_pool)))
            docx_file = generate_word_doc(subset)
            st.download_button(
                label="Download .docx",
                data=docx_file,
                file_name="PE_Exam_Set.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    # Start Exam Button
    if st.session_state.questions_pool and st.session_state.exam_stage == "SETUP":
        st.write("---")
        if st.button("2. Start CBT Exam", type="primary"):
            st.session_state.exam_session = random.sample(st.session_state.questions_pool,
                                                          min(20, len(st.session_state.questions_pool)))
            # Reset user selections for this session
            for q in st.session_state.exam_session:
                q.user_selections = set()
                q.flagged = False
            st.session_state.current_index = 0
            st.session_state.exam_stage = "ACTIVE"
            st.rerun()

    # Reset Button
    if st.session_state.exam_stage != "SETUP":
        st.write("---")
        if st.button("End Exam / Reset"):
            st.session_state.exam_stage = "SETUP"
            st.rerun()

# 2. MAIN CONTENT AREA
if st.session_state.exam_stage == "SETUP":
    st.title("PE Exam Simulator & Generator")
    st.write("""
    **Instructions:**
    1. Upload your Excel database (Questions in Col A, Options in Col B-E, **Bold** the correct answer).
    2. Click 'Start CBT Exam' to simulate a 20-question test.
    3. Or click 'Generate Word Doc' to print a paper version.
    """)

elif st.session_state.exam_stage == "ACTIVE":
    # Get current question
    q_idx = st.session_state.current_index
    q = st.session_state.exam_session[q_idx]
    total_q = len(st.session_state.exam_session)

    # Header
    col1, col2, col3 = st.columns([1, 4, 1])
    with col1:
        st.caption(f"Question {q_idx + 1} / {total_q}")
    with col3:
        # Flag Toggle
        flag_label = "🚩 Flagged" if q.flagged else "🏳️ Flag"
        if st.button(flag_label):
            q.flagged = not q.flagged
            st.rerun()

    # Progress Bar
    st.progress((q_idx + 1) / total_q)

    # Question Text
    st.subheader(q.text)

    # Options Display
    is_multi = len(q.correct_indices) > 1
    if is_multi:
        st.info("Select all that apply.")
        for i, opt in enumerate(q.options):
            # Checkbox logic
            checked = i in q.user_selections
            if st.checkbox(opt, value=checked, key=f"q{q_idx}_opt{i}"):
                q.user_selections.add(i)
            else:
                q.user_selections.discard(i)
    else:
        # Radio button logic
        # We need to map the selection back to the index
        current_selection = list(q.user_selections)[0] if q.user_selections else None

        # Helper to find index of current selection for default value
        idx_val = current_selection if current_selection is not None else 0

        selected_option = st.radio(
            "Select one:",
            q.options,
            index=idx_val if current_selection is not None else None,
            key=f"q{q_idx}_radio"
        )

        # Update state immediately based on radio selection
        if selected_option:
            sel_index = q.options.index(selected_option)
            q.user_selections = {sel_index}

    st.write("---")

    # Navigation Buttons
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        if q_idx > 0:
            if st.button("<< Previous"):
                st.session_state.current_index -= 1
                st.rerun()
    with c2:
        if st.button("Review All Questions"):
            st.session_state.exam_stage = "REVIEW"
            st.rerun()
    with c3:
        if q_idx < total_q - 1:
            if st.button("Next >>"):
                st.session_state.current_index += 1
                st.rerun()
        else:
            if st.button("Go to Review"):
                st.session_state.exam_stage = "REVIEW"
                st.rerun()

elif st.session_state.exam_stage == "REVIEW":
    st.title("Review Your Answers")

    # Grid Layout for Buttons
    cols = st.columns(5)
    for i, q in enumerate(st.session_state.exam_session):
        status = "⬛"  # Unanswered
        if q.is_answered(): status = "🟩"  # Answered
        if q.flagged: status = "🚩"  # Flagged

        label = f"Q{i + 1} {status}"
        with cols[i % 5]:
            if st.button(label, key=f"nav_{i}"):
                st.session_state.current_index = i
                st.session_state.exam_stage = "ACTIVE"
                st.rerun()

    st.write("---")
    st.caption("Legend: 🟩 Answered | 🚩 Flagged | ⬛ Unanswered")

    if st.button("SUBMIT EXAM", type="primary"):
        st.session_state.exam_stage = "REPORT"
        st.rerun()

elif st.session_state.exam_stage == "REPORT":
    st.title("Exam Results")

    score = sum(1 for q in st.session_state.exam_session if q.is_correct())
    total = len(st.session_state.exam_session)
    percent = (score / total) * 100

    st.metric("Final Score", f"{percent:.1f}%", f"{score}/{total} Correct")

    # Results Table
    results_data = []
    for i, q in enumerate(st.session_state.exam_session):
        status = "✅ CORRECT" if q.is_correct() else "❌ INCORRECT"
        results_data.append({
            "Q#": i + 1,
            "ID": q.id,
            "Result": status,
            "Flagged": "Yes" if q.flagged else ""
        })

    st.table(results_data)

    if st.button("Start New Exam"):
        st.session_state.exam_stage = "SETUP"
        st.rerun()