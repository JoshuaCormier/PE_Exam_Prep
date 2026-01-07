import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import random
import re
import sys
import os
from typing import List

# --- WORD DOCUMENT LIBRARY ---
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    Document = None


# --- HELPER: RESOURCE PATH ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# --- HELPER: HIGH-QUALITY ICON GENERATOR ---
def ensure_icon_exists():
    icon_path = "PE_icon.ico"
    try:
        from PIL import Image, ImageDraw, ImageFont

        # 1. Create Master Image (High Res 512x512)
        size = 512
        img = Image.new('RGBA', (size, size), color="#003366")
        draw = ImageDraw.Draw(img)

        # 2. Border (Keep it thick so it remains visible at small sizes)
        draw.rectangle([0, 0, size - 1, size - 1], outline="white", width=40)

        # 3. Center Text with Padding
        try:
            # Changed font size from 300 -> 280 to give it "breathing room"
            font = ImageFont.truetype("arial.ttf", 280)
            text = "PE"

            # Smart Centering
            try:
                left, top, right, bottom = font.getbbox(text)
                w = right - left
                h = bottom - top
            except AttributeError:
                w, h = draw.textsize(text, font=font)

            # Adjusted vertical math for better optical centering
            x = (size - w) / 2
            y = (size - h) / 2 - (h * 0.1)

            draw.text((x, y), text, fill="white", font=font)
        except:
            # Redrawn Fallback (Block Letters) - cleaner coordinates
            draw.rectangle([110, 110, 160, 400], fill="white")  # P (Left)
            draw.rectangle([110, 110, 260, 160], fill="white")  # P (Top)
            draw.rectangle([210, 110, 260, 260], fill="white")  # P (Curve)
            draw.rectangle([110, 260, 260, 310], fill="white")  # P (Mid)

            draw.rectangle([300, 110, 350, 400], fill="white")  # E (Left)
            draw.rectangle([300, 110, 450, 160], fill="white")  # E (Top)
            draw.rectangle([300, 230, 420, 280], fill="white")  # E (Mid)
            draw.rectangle([300, 350, 450, 400], fill="white")  # E (Bot)

        # 4. Generate Layers (CRITICAL FOR ICON QUALITY)
        icon_sizes = []
        for s in [256, 128, 64, 48, 32, 16]:
            # Lanczos filter prevents "jagged" pixels
            resized = img.resize((s, s), resample=Image.Resampling.LANCZOS)
            icon_sizes.append(resized)

        # Save as a multi-layer ICO
        img.save(icon_path, format='ICO', sizes=[(s.size[0], s.size[1]) for s in icon_sizes], append_images=icon_sizes)

    except Exception:
        pass


class Question:
    def __init__(self, text, options, correct_indices, q_id):
        self.text = text
        self.options = options
        self.correct_indices = set(correct_indices)
        self.id = q_id
        self.user_selections = set()
        self.flagged = False

    def is_correct(self): return self.user_selections == self.correct_indices

    def is_answered(self): return len(self.user_selections) > 0


class QuizApp:
    def __init__(self, root):
        self.root = root
        ensure_icon_exists()
        try:
            icon_path = resource_path("PE_icon.ico")
            if os.path.exists(icon_path): self.root.iconbitmap(icon_path)
        except:
            pass

        self.root.title("PE Exam Simulator & Generator")
        self.root.geometry("1100x850")
        self.questions_pool = []
        self.current_session_questions = []
        self.current_index = 0
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 10))
        self.main_container = ttk.Frame(self.root, padding="20")
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.show_start_screen()

    def show_start_screen(self):
        self.clear_frame()
        ttk.Label(self.main_container, text="Work. Study. Learn.", font=("Arial", 24, "bold")).pack(pady=20)
        inst = ("NOTE: Exam database must be in Excel format.\n    "
                " - Place the question in Column A\n    "
                " - Place possible solutions in Columns B, C, D, and E\n    "
                " - Correct answer must be bold\n")
        ttk.Label(self.main_container, text=inst, font=("Arial", 12)).pack(pady=10)
        ttk.Button(self.main_container, text="1. Load Excel Database", command=self.load_file).pack(pady=10, ipadx=10,
                                                                                              ipady=5)
        self.btn_start = ttk.Button(self.main_container, text="2. Start CBT Exam", command=self.start_exam,
                                    state="disabled")
        self.btn_start.pack(pady=5, ipadx=10, ipady=5)
        self.btn_print = ttk.Button(self.main_container, text="3. Print Exam Set to Word", command=self.generate_word_doc,
                                    state="disabled")
        self.btn_print.pack(pady=5, ipadx=10, ipady=5)
        self.lbl_status = ttk.Label(self.main_container, text="Waiting...", foreground="blue")
        self.lbl_status.pack(pady=20)

    def load_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not fp: return
        try:
            self.questions_pool = self.parse_excel(fp)
            if not self.questions_pool: raise ValueError("No questions.")
            self.lbl_status.config(text=f"Loaded {len(self.questions_pool)} Questions", foreground="green")
            self.btn_start.config(state="normal");
            self.btn_print.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def parse_excel(self, fp) -> List[Question]:
        wb = openpyxl.load_workbook(fp, data_only=True)
        qs = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(min_row=2, values_only=False):
                if not row[0].value: continue
                txt = row[0].value
                qid = re.search(r'\(([\w-]+)\)\s*$', str(txt))
                qid = qid.group(1) if qid else "ID"
                opts = [];
                cor = []
                for i, c in enumerate(row[1:]):
                    if c.value:
                        opts.append(str(c.value))
                        if c.font and c.font.bold: cor.append(i)
                if opts and cor: qs.append(Question(str(txt), opts, cor, qid))
        return qs

    def generate_word_doc(self):
        if not Document: return messagebox.showerror("Error", "pip install python-docx")
        if not self.questions_pool: return
        exam = random.sample(self.questions_pool, min(20, len(self.questions_pool)))
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if not path: return
        try:
            doc = Document()
            doc.add_heading('PE Exam Set', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"ID: {random.randint(1000, 9999)}")
            for i, q in enumerate(exam):
                doc.add_paragraph().add_run(f"{i + 1}. {q.text}").bold = True
                lets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                for idx, o in enumerate(q.options): doc.add_paragraph(f"    {lets[idx] if idx < 26 else '?'}. {o}")
                doc.add_paragraph("")
            doc.add_page_break()
            doc.add_heading('Answer Key', 0)
            tbl = doc.add_table(rows=1, cols=2);
            tbl.style = 'Light Shading'
            tbl.rows[0].cells[0].text = 'Q#';
            tbl.rows[0].cells[1].text = 'Ans'
            for i, q in enumerate(exam):
                row = tbl.add_row().cells;
                row[0].text = str(i + 1)
                row[1].text = ", ".join([lets[x] for x in sorted(q.correct_indices)])
            doc.save(path)
            messagebox.showinfo("Success", "The exam set has been saved!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def start_exam(self):
        self.current_session_questions = random.sample(self.questions_pool, min(20, len(self.questions_pool)))
        for q in self.current_session_questions: q.user_selections = set(); q.flagged = False
        self.current_index = 0;
        self.show_question_screen()

    def show_question_screen(self):
        self.clear_frame()

        # --- FIX START: Initialize a list to hold variable references ---
        self.vars = []
        # ----------------------------------------------------------------

        q = self.current_session_questions[self.current_index]
        top = ttk.Frame(self.main_container)
        top.pack(fill=tk.X)
        ttk.Label(top, text=f"Q {self.current_index + 1} / {len(self.current_session_questions)}",
                  font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        t = tk.Text(self.main_container, height=6, wrap=tk.WORD, font=("Arial", 12), bg="#f0f0f0", bd=0)
        t.insert("1.0", q.text)
        t.config(state="disabled")
        t.pack(fill=tk.X, pady=10)

        opts = ttk.Frame(self.main_container)
        opts.pack(fill=tk.BOTH, expand=True)
        is_multi = len(q.correct_indices) > 1
        if is_multi:
            ttk.Label(opts, text="(Select all)", foreground="blue").pack(anchor="w")

        # Prepare the single-choice variable ONCE outside the loop
        if not is_multi:
            # Check if there is a previous selection, otherwise default to -1
            current_val = list(q.user_selections)[0] if q.user_selections else -1
            iv = tk.IntVar(value=current_val)
            self.vars.append(iv)  # Keep reference alive

        for i, o in enumerate(q.options):
            if is_multi:
                v = tk.BooleanVar(value=(i in q.user_selections))
                self.vars.append(v)  # --- FIX: Keep reference alive ---

                def c(x=i, val=v):
                    q.user_selections.add(x) if val.get() else q.user_selections.discard(x)

                tk.Checkbutton(opts, text=o, variable=v, command=c, wraplength=900, justify="left",
                               font=("Arial", 11)).pack(anchor="w")
            else:
                # Use the 'iv' created outside the loop
                def r(x=i):
                    q.user_selections = {x}

                tk.Radiobutton(opts, text=o, variable=iv, value=i, command=r, wraplength=900, justify="left",
                               font=("Arial", 11)).pack(anchor="w")

        # NAVIGATION
        nav = ttk.Frame(self.main_container)
        nav.pack(fill=tk.X, pady=20)
        if self.current_index > 0:
            ttk.Button(nav, text="<< Prev", command=lambda: self.move(-1)).pack(side=tk.LEFT)

        f_txt = "Unflag" if q.flagged else "Flag"
        ttk.Button(nav, text=f_txt, command=self.toggle_flag).pack(side=tk.LEFT, padx=20)
        ttk.Button(nav, text="Review All", command=self.show_review).pack(side=tk.RIGHT, padx=10)

        if self.current_index < len(self.current_session_questions) - 1:
            ttk.Button(nav, text="Next >>", command=lambda: self.move(1)).pack(side=tk.RIGHT)
        else:
            ttk.Button(nav, text="Finish Section", command=self.show_review).pack(side=tk.RIGHT)

    def move(self, d):
        self.current_index += d; self.show_question_screen()

    def toggle_flag(self):
        self.current_session_questions[self.current_index].flagged = not self.current_session_questions[
            self.current_index].flagged
        self.show_question_screen()

    def show_review(self):
        self.clear_frame()
        ttk.Label(self.main_container, text="Review Your Answers", font=("Arial", 18)).pack(pady=10)
        grid = ttk.Frame(self.main_container);
        grid.pack(pady=20)
        leg = ttk.Frame(self.main_container);
        leg.pack(pady=10)
        ttk.Label(leg, text="[ Answered ]", foreground="green").pack(side=tk.LEFT, padx=10)
        ttk.Label(leg, text="[ FLAGGED ]", foreground="red", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)

        for i, q in enumerate(self.current_session_questions):
            txt = f"Q{i + 1}" + (" (F)" if q.flagged else "")
            color = "red" if q.flagged else ("green" if q.is_answered() else "black")
            btn = tk.Button(grid, text=txt, width=8, fg=color, command=lambda idx=i: self.jump(idx))
            btn.grid(row=i // 5, column=i % 5, padx=5, pady=5)

        act = ttk.Frame(self.main_container);
        act.pack(fill=tk.X, pady=20)
        ttk.Button(act, text="<< Back", command=self.show_question_screen).pack(side=tk.LEFT)
        ttk.Button(act, text="SUBMIT EXAM", command=self.submit).pack(side=tk.RIGHT)

    def jump(self, idx):
        self.current_index = idx
        self.show_question_screen()

    def submit(self):
        if messagebox.askyesno("Submit", "Grade exam now?"): self.show_report()

    # --- RESTORED SCORE SHEET ---
    def show_report(self):
        self.clear_frame()
        score = sum(1 for q in self.current_session_questions if q.is_correct())
        total = len(self.current_session_questions)
        pct = (score / total) * 100 if total > 0 else 0

        ttk.Label(self.main_container, text="Exam Report", font=("Arial", 20, "bold")).pack(pady=10)
        ttk.Label(self.main_container, text=f"Score: {score}/{total} ({pct:.1f}%)", font=("Arial", 16)).pack(pady=10)

        # TABLE (TREEVIEW)
        tree_frame = ttk.Frame(self.main_container)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("#", "ID", "Result", "Flag")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for c in cols: tree.heading(c, text=c)
        tree.column("#", width=50, anchor="center")
        tree.column("ID", width=100, anchor="center")
        tree.column("Result", width=100, anchor="center")
        tree.column("Flag", width=50, anchor="center")

        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for i, q in enumerate(self.current_session_questions):
            res = "CORRECT" if q.is_correct() else "INCORRECT"
            tag = "cor" if q.is_correct() else "inc"
            tree.insert("", tk.END, values=(i + 1, q.id, res, "Yes" if q.flagged else ""), tags=(tag,))

        tree.tag_configure("cor", foreground="green")
        tree.tag_configure("inc", foreground="red")

        ttk.Button(self.main_container, text="Return to Menu", command=self.show_start_screen).pack(pady=10)

    def clear_frame(self):
        for w in self.main_container.winfo_children(): w.destroy()


if __name__ == "__main__":
    try:
        from ctypes import windll;

        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    root = tk.Tk()
    app = QuizApp(root)
    root.mainloop()