import streamlit as st
import openpyxl
import random
import re
import io
import json
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
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=False), start=2):
            if not row[0].value: continue
            txt = str(row[0].value)

            qid_match = re.search(r'\(([\w-]+)\)\s*$', txt)
            if qid_match:
                qid = qid_match.group(1)
            else:
                qid = f"GEN_{row_idx}_{random.randint(100, 999)}"

            opts = []
            cor = []
            for i, c in enumerate(row[1:]):
                if c.value:
                    opts.append(str(c.value))
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

    for i, q in enumerate(questions):
        p = doc.add_paragraph()
        p.add_run(f"{i + 1}. {q.text}").bold = True
        lets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        for idx, o in enumerate(q.options):
            doc.add_paragraph(f"    {lets[idx] if idx < 26 else '?'}. {o}")
        doc.add_paragraph("")

    doc.add_page_break()
    doc.add_heading('Answer Key', 0)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Light Shading'
    tbl.rows[0].cells[0].text = 'Q#'
    tbl.rows[0].cells[1].text = 'Ans'

    for i, q in enumerate(questions):
        row = tbl.add_row().cells
        row[0].text = str(i + 1)
        lets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        ans_str = ", ".join([lets[x] for x in sorted(q.correct_indices)])
        row[1].text = ans_str

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generate_war_report(questions, title="Wrong Answer Roundup"):
    """Generates a text file listing wrong questions for review."""
    output = f"--- {title} ---\n\n"
    for q in questions:
        if not q.is_correct():
            output += f"ID: {q.id}\n"
            output += f"Q: {q.text}\n"
            lets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

            # Show Correct Answer(s)
            ans_str = ", ".join([lets[x] for x in sorted(q.correct_indices)])
            output += f"Correct Answer: {ans_str}\n"
            output += "-" * 40 + "\n"
    return output


# --- SESSION STATE INITIALIZATION ---
if 'questions_pool' not in st.session_state:
    st.session_state.questions_pool = []
if 'exam_session' not in st.session_state:
    st.session_state.exam_session = []
if 'current_index' not in st.session_state:
    st.session_state.current_index = 0
if 'exam_stage' not in st.session_state:
    st.session_state.exam_stage = "SETUP"

# Persistent Tracking Variables
if 'used_ids' not in st.session_state:
    st.session_state.used_ids = set()
if 'wrong_ids' not in st.session_state:
    st.session_state.wrong_ids = set()
if 'cumulative_correct' not in st.session_state:
    st.session_state.cumulative_correct = 0
if 'cumulative_answered' not in st.session_state:
    st.session_state.cumulative_answered = 0

# --- MAIN APP LOGIC ---

# 1. SIDEBAR: Controls & Navigation
with st.sidebar:
    st.header("1. Load Data")

    # A. Load Questions
    uploaded_file = st.file_uploader("Upload Question Database (Excel)", type=["xlsx"])
    if uploaded_file:
        try:
            questions = parse_excel(uploaded_file)
            st.session_state.questions_pool = questions
            st.success(f"Loaded {len(questions)} questions.")
        except Exception as e:
            st.error(f"Error parsing file: {e}")

    # B. Load History
    history_file = st.file_uploader("Upload History File (Optional)", type=["json"])
    if history_file:
        try:
            data = json.load(history_file)

            if isinstance(data, list):
                # Legacy support
                st.session_state.used_ids.update(set(data))
                st.warning("Loaded legacy history (No score/wrong ID tracking yet).")
            elif isinstance(data, dict):
                st.session_state.used_ids.update(set(data.get("used_ids", [])))
                st.session_state.wrong_ids.update(set(data.get("wrong_ids", [])))
                st.session_state.cumulative_correct = data.get("correct", 0)
                st.session_state.cumulative_answered = data.get("answered", 0)
                st.success(
                    f"Restored: {len(st.session_state.used_ids)} Qs | {len(st.session_state.wrong_ids)} Wrong IDs")

        except Exception as e:
            st.error(f"Error loading history: {e}")

    # C. Performance Metrics
    st.write("---")
    st.header("📊 Performance")

    total_bank = len(st.session_state.questions_pool)
    used_count = len(st.session_state.used_ids)

    # Bank Progress
    st.caption("Bank Completion")
    prog_val = used_count / total_bank if total_bank > 0 else 0
    st.progress(prog_val)
    st.write(f"**{used_count} / {total_bank}** Questions Seen")

    # Cumulative Score
    st.caption("Cumulative Accuracy")
    if st.session_state.cumulative_answered > 0:
        acc = (st.session_state.cumulative_correct / st.session_state.cumulative_answered) * 100
        st.write(f"**{acc:.1f}%** ({st.session_state.cumulative_correct}/{st.session_state.cumulative_answered})")

    # DOWNLOAD ALL WRONG IDs (Anytime)
    if len(st.session_state.wrong_ids) > 0:
        st.write("---")
        wrong_ids_text = "All-Time Wrong Question IDs:\n" + "\n".join(sorted(list(st.session_state.wrong_ids)))
        st.download_button(
            label="⚠️ Download All Wrong IDs",
            data=wrong_ids_text,
            file_name="All_Wrong_IDs.txt",
            mime="text/plain",
            help="Download a list of every question ID you have ever missed."
        )

    # D. Exam Controls
    if st.session_state.questions_pool and st.session_state.exam_stage == "SETUP":
        st.write("---")
        st.header("2. Start Exam")

        available_pool = [q for q in st.session_state.questions_pool if q.id not in st.session_state.used_ids]

        if len(available_pool) == 0:
            st.warning("All questions attempted!")
            if st.button("🔄 Reset History & Scores", type="primary"):
                st.session_state.used_ids = set()
                st.session_state.wrong_ids = set()
                st.session_state.cumulative_correct = 0
                st.session_state.cumulative_answered = 0
                st.rerun()
        else:
            remaining_qs = len(available_pool)
            btn_label = "Start New Exam Block"
            if remaining_qs < 20:
                st.info(f"Only {remaining_qs} questions remaining.")
                btn_label = f"Start Final {remaining_qs} Questions"

            if st.button(btn_label, type="primary"):
                session_size = min(20, len(available_pool))
                st.session_state.exam_session = random.sample(available_pool, session_size)

                # Mark as used immediately upon start
                for q in st.session_state.exam_session:
                    st.session_state.used_ids.add(q.id)
                    q.user_selections = set()
                    q.flagged = False

                st.session_state.current_index = 0
                st.session_state.exam_stage = "ACTIVE"
                st.rerun()

    # E. Save Progress
    st.write("---")
    st.header("3. Save Progress")
    if len(st.session_state.used_ids) > 0:
        save_data = {
            "used_ids": list(st.session_state.used_ids),
            "wrong_ids": list(st.session_state.wrong_ids),
            "correct": st.session_state.cumulative_correct,
            "answered": st.session_state.cumulative_answered
        }

        history_json = json.dumps(save_data)
        st.download_button(
            label="💾 Download Progress File",
            data=history_json,
            file_name="PE_Exam_History.json",
            mime="application/json",
            help="Includes your used IDs, wrong IDs, and running score."
        )

    # Reset Button
    if st.session_state.exam_stage != "SETUP":
        st.write("---")
        if st.button("End Exam / Return to Menu"):
            st.session_state.exam_stage = "SETUP"
            st.rerun()

    # Word Doc Export
    if st.session_state.questions_pool and st.session_state.exam_stage == "SETUP":
        if st.button("Generate Word Doc"):
            subset = random.sample(st.session_state.questions_pool, min(20, len(st.session_state.questions_pool)))
            docx_file = generate_word_doc(subset)
            st.download_button(
                label="Download .docx",
                data=docx_file,
                file_name="PE_Exam_Set.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# 2. MAIN CONTENT AREA
if st.session_state.exam_stage == "SETUP":
    st.title("PE Exam Quizzes")
    st.info("""
    **Campaign Mode Active**
    1. **Upload** Excel Question Bank & History File.
    2. **Take Exams** (Questions are removed from the pool as you use them).
    3. **Download Reports:** - Get a "Session Report" immediately after an exam.
       - Get an "All-Time Wrong IDs" list from the sidebar anytime.
    4. **Save Progress** before leaving.
    """)

elif st.session_state.exam_stage == "ACTIVE":
    q_idx = st.session_state.current_index
    q = st.session_state.exam_session[q_idx]
    total_q = len(st.session_state.exam_session)

    col1, col2, col3 = st.columns([1, 4, 1])
    with col1:
        st.caption(f"Question {q_idx + 1} / {total_q}")
    with col3:
        flag_label = "🚩 Flagged" if q.flagged else "🏳️ Flag"
        if st.button(flag_label):
            q.flagged = not q.flagged
            st.rerun()

    st.progress((q_idx + 1) / total_q)
    st.subheader(q.text)

    is_multi = len(q.correct_indices) > 1
    if is_multi:
        st.info("Select all that apply.")
        for i, opt in enumerate(q.options):
            checked = i in q.user_selections
            if st.checkbox(opt, value=checked, key=f"q{q_idx}_opt{i}"):
                q.user_selections.add(i)
            else:
                q.user_selections.discard(i)
    else:
        current_selection = list(q.user_selections)[0] if q.user_selections else None
        idx_val = current_selection if current_selection is not None else 0
        selected_option = st.radio(
            "Select one:",
            q.options,
            index=idx_val if current_selection is not None else None,
            key=f"q{q_idx}_radio"
        )
        if selected_option:
            sel_index = q.options.index(selected_option)
            q.user_selections = {sel_index}

    st.write("---")

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
    cols = st.columns(5)
    for i, q in enumerate(st.session_state.exam_session):
        status = "⬛"
        if q.is_answered(): status = "🟩"
        if q.flagged: status = "🚩"

        label = f"Q{i + 1} {status}"
        with cols[i % 5]:
            if st.button(label, key=f"nav_{i}"):
                st.session_state.current_index = i
                st.session_state.exam_stage = "ACTIVE"
                st.rerun()

    st.write("---")

    if st.button("SUBMIT EXAM", type="primary"):
        # Update Scores and Wrong IDs List
        s_correct = 0
        s_total = len(st.session_state.exam_session)

        for q in st.session_state.exam_session:
            if q.is_correct():
                s_correct += 1
            else:
                # Add to cumulative wrong list
                st.session_state.wrong_ids.add(q.id)

        st.session_state.cumulative_correct += s_correct
        st.session_state.cumulative_answered += s_total

        st.session_state.exam_stage = "REPORT"
        st.rerun()

elif st.session_state.exam_stage == "REPORT":
    st.title("Exam Results")
    score = sum(1 for q in st.session_state.exam_session if q.is_correct())
    total = len(st.session_state.exam_session)
    percent = (score / total) * 100

    st.metric("Session Score", f"{percent:.1f}%", f"{score}/{total} Correct")

    # 1. WAR REPORT BUTTON (For Immediate Review)
    wrong_qs = [q for q in st.session_state.exam_session if not q.is_correct()]
    if wrong_qs:
        war_report = generate_war_report(wrong_qs, title="Session Wrong Answer Roundup")
        st.download_button(
            label="⚠️ Download Session WAR Report",
            data=war_report,
            file_name="Session_WAR_Report.txt",
            mime="text/plain",
            help="Download a text file of just the questions you missed this session."
        )
    else:
        st.success("Perfect Score! No WAR Report needed.")

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

    if st.button("Start Next Exam Block"):
        st.session_state.exam_stage = "SETUP"
        st.rerun()