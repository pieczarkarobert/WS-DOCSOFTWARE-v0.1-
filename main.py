import json
import os
import re
import tkinter as tk
from tkinter import colorchooser, filedialog, font, messagebox, simpledialog, ttk

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas as pdf_canvas
    HAS_PDF = True
except ImportError:
    HAS_PDF = False


KEYWORDS_PY = {
    "def", "class", "return", "if", "elif", "else", "for", "while", "break",
    "continue", "import", "from", "as", "try", "except", "finally", "with",
    "pass", "True", "False", "None", "and", "or", "not", "in", "is", "lambda",
    "yield", "async", "await",
}

KEYWORDS_CPP = {
    "int", "float", "double", "char", "void", "bool", "long", "short", "unsigned",
    "const", "static", "return", "if", "else", "for", "while", "do", "switch",
    "case", "break", "continue", "class", "struct", "public", "private",
    "protected", "include", "using", "namespace", "new", "delete", "true", "false",
    "nullptr", "auto", "template", "typename",
}


class DocSoftware(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DocSoftware v0.1 — foundation for WS DocStudio")
        self.geometry("1040x700")
        self.minsize(720, 480)

        self.current_file = None
        self.code_mode = False
        self.text_color = "#1e1e1e"
        self.font_family = "Calibri" if "Calibri" in font.families() else "TkDefaultFont"
        self.font_size = 12
        self.code_font_family = "Consolas" if "Consolas" in font.families() else "TkFixedFont"
        self.code_font_size = 11

        self.style_counter = 0
        self._highlight_job = None

        self._build_menu()
        self._build_toolbar()
        self._build_editor()
        self._build_statusbar()
        self._setup_tags()
        self._bind_events()

        self.text.focus_set()

    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New", accelerator="Ctrl+N", command=self.new_file)
        file_menu.add_command(label="Open...", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="Save", accelerator="Ctrl+S", command=self.save_file)
        file_menu.add_command(label="Save As / Export...", accelerator="Ctrl+Shift+S", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=self._undo)
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=self._redo)
        edit_menu.add_separator()
        edit_menu.add_command(label="Insert Table...", command=self.insert_table_dialog)
        edit_menu.add_command(label="Insert Code Block", command=self.insert_code_block)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_checkbutton(label="Code Mode", command=self.toggle_code_mode)
        menubar.add_cascade(label="View", menu=view_menu)

        self.config(menu=menubar)

    def _build_toolbar(self):
        toolbar = ttk.Frame(self, padding=(6, 4))
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(toolbar, text="New", width=7, command=self.new_file).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="Open", width=8, command=self.open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Save", width=8, command=self.save_file).pack(side=tk.LEFT, padx=(2, 8))

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        ttk.Button(toolbar, text="B", width=3, command=lambda: self.toggle_tag("bold")).pack(side=tk.LEFT, padx=1)
        ttk.Button(toolbar, text="I", width=3, command=lambda: self.toggle_tag("italic")).pack(side=tk.LEFT, padx=1)
        ttk.Button(toolbar, text="U", width=3, command=lambda: self.toggle_tag("underline")).pack(side=tk.LEFT, padx=(1, 6))

        families = sorted(set(font.families()))
        default_f = self.font_family if self.font_family in families else families[0]
        self.font_var = tk.StringVar(value=default_f)
        font_combo = ttk.Combobox(toolbar, textvariable=self.font_var, values=families, width=18, state="readonly")
        font_combo.pack(side=tk.LEFT, padx=2)
        font_combo.bind("<<ComboboxSelected>>", self._on_font_change)

        self.size_var = tk.StringVar(value=str(self.font_size))
        size_combo = ttk.Combobox(
            toolbar,
            textvariable=self.size_var,
            values=["9", "10", "11", "12", "14", "16", "18", "20", "24", "28"],
            width=4,
            state="readonly",
        )
        size_combo.pack(side=tk.LEFT, padx=2)
        size_combo.bind("<<ComboboxSelected>>", self._on_font_change)

        ttk.Button(toolbar, text="Color", width=6, command=self.choose_color).pack(side=tk.LEFT, padx=(2, 8))

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        ttk.Button(toolbar, text="+ Table", command=self.insert_table_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="{ } Code", command=self.insert_code_block).pack(side=tk.LEFT, padx=2)

        self.code_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            toolbar,
            text="Code Mode",
            variable=self.code_mode_var,
            command=self._sync_code_mode_from_button,
        ).pack(side=tk.LEFT, padx=(8, 2))

        self.lang_var = tk.StringVar(value="Python")
        lang_combo = ttk.Combobox(
            toolbar,
            textvariable=self.lang_var,
            values=["Python", "HTML", "C++"],
            width=8,
            state="readonly",
        )
        lang_combo.pack(side=tk.LEFT, padx=2)
        lang_combo.bind("<<ComboboxSelected>>", lambda _e: self.schedule_highlight())

    def _build_editor(self):
        container = ttk.Frame(self)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        v_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.text = tk.Text(
            container,
            wrap=tk.WORD,
            undo=True,
            maxundo=-1,
            bg="#ffffff",
            fg=self.text_color,
            insertbackground="#1e1e1e",
            selectbackground="#cde4ff",
            relief=tk.FLAT,
            padx=24,
            pady=18,
            font=(self.font_family, self.font_size),
            yscrollcommand=v_scroll.set,
            tabs=("1c",),
        )
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.config(command=self.text.yview)

    def _build_statusbar(self):
        bar = ttk.Frame(self, padding=(8, 3))
        bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_left = ttk.Label(bar, text="Ready | Mode: Document")
        self.status_left.pack(side=tk.LEFT)

        self.status_right = ttk.Label(bar, text="Line 1, Col 1")
        self.status_right.pack(side=tk.RIGHT)

    def _setup_tags(self):
        base = font.Font(family=self.font_family, size=self.font_size)
        bold_f = font.Font(family=self.font_family, size=self.font_size, weight="bold")
        italic_f = font.Font(family=self.font_family, size=self.font_size, slant="italic")
        bi_f = font.Font(family=self.font_family, size=self.font_size, weight="bold", slant="italic")

        self.text.tag_configure("bold", font=bold_f)
        self.text.tag_configure("italic", font=italic_f)
        self.text.tag_configure("underline", underline=True)
        self.text.tag_configure("bold_italic", font=bi_f)
        self.text.tag_configure("_base", font=base)

        self.text.tag_configure(
            "code_block",
            background="#f4f5f7",
            foreground="#18181b",
            font=(self.code_font_family, self.code_font_size),
            lmargin1=16,
            lmargin2=16,
            rmargin=16,
        )
        # Syntax highlighting tags for code mode
        self.text.tag_configure("syn_kw", foreground="#0550ae")
        self.text.tag_configure("syn_str", foreground="#0a3069")
        self.text.tag_configure("syn_com", foreground="#6e7781")
        self.text.tag_configure("syn_tag", foreground="#116329")

    def _bind_events(self):
        self.bind("<Control-n>", lambda _e: self.new_file())
        self.bind("<Control-o>", lambda _e: self.open_file())
        self.bind("<Control-s>", lambda _e: self.save_file())
        self.bind("<Control-S>", lambda _e: self.save_file_as())
        self.bind("<Control-z>", lambda _e: self._undo())
        self.bind("<Control-y>", lambda _e: self._redo())

        self.text.bind("<KeyRelease>", self._on_key_release)
        self.text.bind("<ButtonRelease-1>", lambda _e: self._update_cursor_status())
        self.text.bind("<Tab>", self._on_tab)

    def _undo(self):
        try:
            self.text.edit_undo()
        except tk.TclError:
            pass

    def _redo(self):
        try:
            self.text.edit_redo()
        except tk.TclError:
            pass

    def _on_key_release(self, _event=None):
        self._update_cursor_status()
        if self.code_mode:
            self.schedule_highlight()

    def _update_cursor_status(self):
        idx = self.text.index(tk.INSERT)
        line, col = idx.split(".")
        mode_txt = f"Code ({self.lang_var.get()})" if self.code_mode else "Document"
        name = os.path.basename(self.current_file) if self.current_file else "Untitled"
        self.status_left.config(text=f"{name} | Mode: {mode_txt}")
        self.status_right.config(text=f"Line {line}, Col {int(col) + 1}")

    def _on_tab(self, _event):
        if self.code_mode:
            self.text.insert(tk.INSERT, "    ")
            return "break"
        return None

    def _sync_code_mode_from_button(self):
        self.toggle_code_mode(force_state=self.code_mode_var.get())

    def toggle_code_mode(self, force_state=None):
        self.code_mode = (not self.code_mode) if force_state is None else bool(force_state)
        self.code_mode_var.set(self.code_mode)

        if self.code_mode:
            self.text.configure(
                font=(self.code_font_family, self.code_font_size),
                bg="#fafafa",
                wrap=tk.NONE,
            )
            self.schedule_highlight()
        else:
            self._clear_syntax_tags()
            self.text.configure(
                font=(self.font_family, self.font_size),
                bg="#ffffff",
                wrap=tk.WORD,
            )
        self._update_cursor_status()

    def insert_code_block(self):
        snippet = "\n# Code block\ndef example():\n    return 'DocSoftware'\n\n"
        start = self.text.index(tk.INSERT)
        self.text.insert(start, snippet)
        end = self.text.index(f"{start} + {len(snippet)} chars")
        self.text.tag_add("code_block", start, end)
        if self.code_mode:
            self.schedule_highlight()

    def schedule_highlight(self):
        if self._highlight_job:
            self.after_cancel(self._highlight_job)
        self._highlight_job = self.after(120, self._apply_syntax_highlighting)

    def _clear_syntax_tags(self):
        for tag in ("syn_kw", "syn_str", "syn_com", "syn_tag"):
            self.text.tag_remove(tag, "1.0", tk.END)

    def _apply_syntax_highlighting(self):
        self._highlight_job = None
        if not self.code_mode:
            return
        self._clear_syntax_tags()
        content = self.text.get("1.0", tk.END)
        lang = self.lang_var.get()

        if lang == "Python":
            comment_pat = r"#[^\n]*"
            kw_set = KEYWORDS_PY
        elif lang == "C++":
            comment_pat = r"//[^\n]*|/\*[\s\S]*?\*/"
            kw_set = KEYWORDS_CPP
        else:
            # Basic tag highlighting for HTML
            for m in re.finditer(r"</?[a-zA-Z0-9:-]+[^>]*>", content):
                self._add_span_tag("syn_tag", m.start(), m.end())
            return

        for m in re.finditer(r"\"\"\"[\s\S]*?\"\"\"|\'\'\'[\s\S]*?\'\'\'|\"(?:\\.|[^\"\\])*\"|\'(?:\\.|[^\'\\])*\'", content):
            self._add_span_tag("syn_str", m.start(), m.end())

        for m in re.finditer(comment_pat, content):
            self._add_span_tag("syn_com", m.start(), m.end())

        for m in re.finditer(r"\b[A-Za-z_][A-Za-z0-9_]*\b", content):
            if m.group(0) in kw_set:
                self._add_span_tag("syn_kw", m.start(), m.end())

    def _add_span_tag(self, tag, start_offset, end_offset):
        s_idx = f"1.0 + {start_offset} chars"
        e_idx = f"1.0 + {end_offset} chars"
        if not self.text.tag_nextrange("syn_str", s_idx, e_idx):
            self.text.tag_add(tag, s_idx, e_idx)

    def toggle_tag(self, tag):
        try:
            start, end = self.text.index(tk.SEL_FIRST), self.text.index(tk.SEL_LAST)
        except tk.TclError:
            messagebox.showinfo("Selection", "Select the text you want to format.")
            return

        current_tags = self.text.tag_names(start)
        if tag in current_tags:
            self.text.tag_remove(tag, start, end)
        else:
            self.text.tag_add(tag, start, end)

        self._refresh_bold_italic(start, end)

    def _refresh_bold_italic(self, start, end):
        pos = start
        while self.text.compare(pos, "<", end):
            nxt = self.text.index(f"{pos} + 1 char")
            tags = set(self.text.tag_names(pos))
            if "bold" in tags and "italic" in tags:
                self.text.tag_add("bold_italic", pos, nxt)
            else:
                self.text.tag_remove("bold_italic", pos, nxt)
            pos = nxt

    def _on_font_change(self, _event=None):
        self.font_family = self.font_var.get()
        try:
            self.font_size = int(self.size_var.get())
        except ValueError:
            self.font_size = 12

        if self.code_mode:
            self.code_font_family = self.font_family
            self.code_font_size = self.font_size
            self.text.configure(font=(self.code_font_family, self.code_font_size))
            return

        self.style_counter += 1
        tag_name = f"style_{self.style_counter}"
        self.text.tag_configure(tag_name, font=(self.font_family, self.font_size))

        try:
            start, end = self.text.index(tk.SEL_FIRST), self.text.index(tk.SEL_LAST)
            self.text.tag_add(tag_name, start, end)
        except tk.TclError:
            self.text.configure(font=(self.font_family, self.font_size))

    def choose_color(self):
        picked = colorchooser.askcolor(color=self.text_color, title="Choose text color")
        if not picked or not picked[1]:
            return
        color = picked[1]
        self.style_counter += 1
        tag_name = f"color_{self.style_counter}"
        self.text.tag_configure(tag_name, foreground=color)

        try:
            start, end = self.text.index(tk.SEL_FIRST), self.text.index(tk.SEL_LAST)
            self.text.tag_add(tag_name, start, end)
        except tk.TclError:
            self.text_color = color
            self.text.configure(fg=color)

    def insert_table_dialog(self):
        rows = simpledialog.askinteger("New Table", "Number of rows:", minvalue=1, maxvalue=30, initialvalue=3)
        if not rows:
            return
        cols = simpledialog.askinteger("New Table", "Number of columns:", minvalue=1, maxvalue=15, initialvalue=3)
        if not cols:
            return

        # Fixed-width text grid so cells stay easy to edit and export across formats
        col_w = 14
        sep = "+" + "+".join(["-" * col_w for _ in range(cols)]) + "+"
        lines = ["", sep]
        for r in range(rows):
            cells = [f" {'Cell ' + str(r+1) + ',' + str(c+1):<{col_w-1}}" for c in range(cols)]
            lines.append("|" + "|".join(cells) + "|")
            lines.append(sep)
        lines.append("")

        table_text = "\n".join(lines)
        start = self.text.index(tk.INSERT)
        self.text.insert(start, table_text)
        end = self.text.index(f"{start} + {len(table_text)} chars")
        self.text.tag_add("code_block", start, end)

    def new_file(self):
        if self.text.get("1.0", "end-1c").strip():
            if not messagebox.askyesno("New Document", "Clear current content?"):
                return
        self.text.delete("1.0", tk.END)
        self.current_file = None
        self._update_cursor_status()

    def open_file(self):
        path = filedialog.askopenfilename(
            title="Open File",
            filetypes=[
                ("All Supported", "*.txt *.md *.html *.htm *.py *.cpp *.c *.h *.js *.json *.bat *.sh"),
                ("Text & Markdown", "*.txt *.md"),
                ("Source Code", "*.py *.cpp *.c *.h *.js *.json *.bat *.sh"),
                ("HTML", "*.html *.htm"),
                ("All Files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
        except OSError as exc:
            messagebox.showerror("Read Error", f"Could not open file:\n{exc}")
            return

        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", content)
        self.current_file = path

        ext = os.path.splitext(path)[1].lower()
        code_exts = {
            ".py": "Python",
            ".cpp": "C++",
            ".c": "C++",
            ".h": "C++",
            ".html": "HTML",
            ".htm": "HTML",
        }
        if ext in code_exts:
            self.lang_var.set(code_exts[ext])
            self.toggle_code_mode(force_state=True)
        elif ext in {".js", ".json", ".bat", ".sh"}:
            self.toggle_code_mode(force_state=True)
        else:
            self.toggle_code_mode(force_state=False)

        self._update_cursor_status()

    def save_file(self):
        if not self.current_file:
            self.save_file_as()
            return
        self._export_to_path(self.current_file)

    def save_file_as(self):
        filetypes = [
            ("Text Document (.txt)", "*.txt"),
            ("Markdown (.md)", "*.md"),
            ("HTML Page (.html)", "*.html"),
            ("Rich Text Format (.rtf)", "*.rtf"),
            ("Word Document (.docx)", "*.docx"),
            ("PDF Document (.pdf)", "*.pdf"),
            ("Python Script (.py)", "*.py"),
            ("C++ Source (.cpp)", "*.cpp"),
            ("JavaScript (.js)", "*.js"),
            ("JSON (.json)", "*.json"),
            ("Windows Batch Script (.bat)", "*.bat"),
            ("Shell Script (.sh)", "*.sh"),
        ]
        path = filedialog.asksaveasfilename(title="Save / Export As", filetypes=filetypes)
        if not path:
            return
        self._export_to_path(path)
        self.current_file = path
        self._update_cursor_status()

    def _export_to_path(self, path):
        ext = os.path.splitext(path)[1].lower()
        raw_text = self.text.get("1.0", "end-1c")

        try:
            if ext == ".docx":
                if not HAS_DOCX:
                    raise ImportError("Missing 'python-docx'. Install it with: pip install python-docx")
                doc = Document()
                for line in raw_text.splitlines():
                    doc.add_paragraph(line)
                doc.save(path)

            elif ext == ".pdf":
                if not HAS_PDF:
                    raise ImportError("Missing 'reportlab'. Install it with: pip install reportlab")
                c = pdf_canvas.Canvas(path, pagesize=letter)
                _, height = letter
                y = height - 50
                for line in raw_text.splitlines():
                    if y < 50:
                        c.showPage()
                        y = height - 50
                    # Fallback for non-Latin1 characters in default PDF canvas font
                    safe_line = line.encode("latin-1", "replace").decode("latin-1")
                    c.drawString(50, y, safe_line[:105])
                    y -= 15
                c.save()

            elif ext == ".html":
                escaped = (
                    raw_text.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                )
                body_tag = "pre" if self.code_mode else "div style='white-space:pre-wrap;'"
                html = (
                    "<!DOCTYPE html>\n<html lang='en'>\n<head>\n<meta charset='utf-8'>\n"
                    "<title>DocSoftware Export</title>\n</head>\n<body>\n"
                    f"<{body_tag}>{escaped}</{body_tag}>\n</body>\n</html>\n"
                )
                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)

            elif ext == ".rtf":
                # Emit RTF with signed unicode escapes for non-ASCII characters
                rtf_parts = [r"{\rtf1\ansi\deff0 "]
                for ch in raw_text:
                    if ch == "\n":
                        rtf_parts.append(r"\par " + "\n")
                    elif ch in ("\\", "{", "}"):
                        rtf_parts.append("\\" + ch)
                    elif ord(ch) < 128:
                        rtf_parts.append(ch)
                    else:
                        code = ord(ch)
                        signed = code - 65536 if code > 32767 else code
                        rtf_parts.append(f"\\u{signed}?")
                rtf_parts.append("}")
                with open(path, "w", encoding="ascii") as f:
                    f.write("".join(rtf_parts))

            elif ext == ".json":
                # Wrap plain text into a JSON object if it isn't valid JSON already
                try:
                    parsed = json.loads(raw_text)
                    out = json.dumps(parsed, indent=2, ensure_ascii=False)
                except json.JSONDecodeError:
                    out = json.dumps({"content": raw_text}, indent=2, ensure_ascii=False)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(out)

            else:
                # .txt, .md, .py, .cpp, .js, .bat, .sh and other text-based formats
                with open(path, "w", encoding="utf-8") as f:
                    f.write(raw_text)

            messagebox.showinfo("Saved", f"File saved to:\n{path}")

        except ImportError as err:
            messagebox.showwarning("Missing Dependency", str(err))
        except OSError as err:
            messagebox.showerror("Save Error", f"Could not save file:\n{err}")


if __name__ == "__main__":
    app = DocSoftware()
    app.mainloop()