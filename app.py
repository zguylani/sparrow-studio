"""
Journal Publishing System v2.0 REAL
Windows-friendly Tkinter app for:
- importing .docx journals
- splitting into manageable prompt chunks
- generating hierarchical JSON extraction prompts
- importing AI JSON results
- merging into hierarchy-aware CSV files

This app does NOT call an AI automatically. It prepares prompts and merges the JSON
responses you save into ai_json_results.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception as exc:  # pragma: no cover
    print("Tkinter could not be imported:", exc)
    raise

try:
    from docx import Document
except Exception:
    Document = None

APP_TITLE = "Journal Publishing System v2.0"
BASE_DIR = Path(__file__).resolve().parent
PROMPTS_DIR = BASE_DIR / "prompts_to_paste"
AI_JSON_DIR = BASE_DIR / "ai_json_results"
OUTPUT_DIR = BASE_DIR / "output"
LOG_PATH = BASE_DIR / "app_log.txt"

CSV_FIELDS = [
    "item_id", "parent_id", "chunk_id", "source_date", "source_heading",
    "level", "category", "rating", "themes", "keywords", "corrected_text",
    "notes", "possible_book", "journal_entry_id", "essay_id", "paragraph_id", "sequence",
]

LEVEL_TO_FILE = {
    "Journal entry": "Journal_Entries.csv",
    "Essay": "Essays.csv",
    "Paragraph": "Paragraphs.csv",
    "One-liner": "One_Liners.csv",
    "Principle/framework": "Principles.csv",
    "Story/illustration": "Stories.csv",
    "Definition": "Definitions.csv",
    "Book idea": "Book_Ideas.csv",
    "Analogy": "Analogies.csv",
    "Sermon idea": "Sermon_Ideas.csv",
    "Research idea": "Research_Ideas.csv",
    "Question": "Questions.csv",
    "Personal testimony": "Personal_Testimonies.csv",
}

DATE_RE = re.compile(r"(?=^Date:\s*)", re.MULTILINE)
DATE_LINE_RE = re.compile(r"^Date:\s*(.+)$", re.MULTILINE)
HEADING_RE = re.compile(r"^#\s*(.+)$", re.MULTILINE)


def log(msg: str) -> None:
    try:
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(msg.rstrip() + "\n")
    except Exception:
        pass


def ensure_dirs() -> None:
    for d in (PROMPTS_DIR, AI_JSON_DIR, OUTPUT_DIR):
        d.mkdir(exist_ok=True)


def safe_slug(name: str) -> str:
    name = Path(name).stem
    name = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return name or "Journal"


def read_docx_text(docx_path: Path) -> str:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Run Install_Requirements.bat first.")
    doc = Document(str(docx_path))
    parts: List[str] = []
    for p in doc.paragraphs:
        txt = p.text.rstrip()
        if txt:
            parts.append(txt)
        else:
            # Preserve paragraph breaks without creating too much noise.
            if parts and parts[-1] != "":
                parts.append("")
    text = "\n".join(parts)
    # Normalize odd spacing while preserving paragraph/line structure.
    text = text.replace("\u202f", " ").replace("\xa0", " ")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def split_entries(text: str) -> List[str]:
    """Split by Date: lines. If none found, split by paragraph blocks."""
    if "Date:" not in text:
        blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
        return blocks
    chunks = [c.strip() for c in DATE_RE.split(text) if c.strip()]
    # DATE_RE is a lookahead, so entries retain Date: at front.
    return chunks


def word_count(s: str) -> int:
    return len(re.findall(r"\S+", s))


def chunk_entries(entries: List[str], target_words: int) -> List[str]:
    """Combine journal entries into chunks near target_words without splitting entries."""
    chunks: List[str] = []
    current: List[str] = []
    current_words = 0
    target_words = max(500, int(target_words))

    for entry in entries:
        wc = word_count(entry)
        if current and current_words + wc > target_words:
            chunks.append("\n\n".join(current).strip())
            current = [entry]
            current_words = wc
        else:
            current.append(entry)
            current_words += wc
    if current:
        chunks.append("\n\n".join(current).strip())
    return chunks


def build_prompt(chunk_text: str, chunk_id: str) -> str:
    return f"""You are helping create Zuba's hierarchical journal publishing system.

CRITICAL RULES:
- Review every sentence in the supplied chunk.
- Preserve wording as close to verbatim as possible.
- Correct only spelling, punctuation, grammar, and obvious autofill/autocorrect/dictation/AI input errors.
- Do not polish, summarize, theologize, censor, or rewrite the author's voice.
- Extract complete thoughts at every natural level. A complete thought may be a whole journal entry, an essay/reflection, a paragraph, or one sentence.
- Preserve hierarchy ALWAYS, not optionally.
- Create a Journal entry item for each dated journal entry that contains usable intellectual, spiritual, narrative, or publication material.
- If a journal entry contains a sustained reflection or essay, create an Essay item under the Journal entry.
- Create Paragraph items for reusable paragraphs under the Essay or Journal entry.
- Create One-liner, Principle/framework, Story/illustration, Definition, Analogy, Question, Book idea, Sermon idea, Research idea, and Personal testimony items under their most specific parent.
- If a paragraph has several usable sentences, extract the paragraph AND the usable sentences inside it.
- Exclude mundane logistical entries unless they contain a usable phrase, story, insight, principle, image, or book material.
- Return ONLY valid JSON. No markdown. No explanation.

RETURN FORMAT:
Return a JSON array of objects. Every object must use exactly these keys:
item_id, parent_id, chunk_id, source_date, source_heading, level, category, rating, themes, keywords, corrected_text, notes, possible_book, journal_entry_id, essay_id, paragraph_id, sequence

FIELD RULES:
- item_id: Use the exact ID prefixes below with sequential numbering inside this chunk.
  Journal entry: {chunk_id}-JE001
  Essay: {chunk_id}-E001
  Paragraph: {chunk_id}-P001
  One-liner/quote: {chunk_id}-Q001
  Principle/framework: {chunk_id}-PR001
  Story/illustration: {chunk_id}-S001
  Definition: {chunk_id}-D001
  Book idea: {chunk_id}-B001
  Analogy: {chunk_id}-A001
  Sermon idea: {chunk_id}-N001
  Research idea: {chunk_id}-R001
  Question: {chunk_id}-X001
  Personal testimony: {chunk_id}-T001
- parent_id: blank if no parent. Otherwise use the item_id of the parent item.
- chunk_id: always {chunk_id}
- source_date: date of the journal entry if present.
- source_heading: heading if present.
- level: one of Journal entry, Essay, Paragraph, One-liner, Principle/framework, Story/illustration, Definition, Book idea, Analogy, Sermon idea, Research idea, Question, Personal testimony.
- category: one or more of Publishable quote, Book idea, Principle/framework, Story/illustration, Definition, Sermon idea, Analogy, Question, Personal testimony, Research idea, Essay/reflection, Paragraph.
- rating: 1-10 for usefulness/publication potential.
- themes: semicolon-separated themes.
- keywords: semicolon-separated keywords.
- corrected_text: the extracted text, lightly corrected only by the rules above.
- notes: brief note about context, parent/child relationship, uncertainty, or why included.
- possible_book: possible book/project fit, such as Belovedness, Grace, Wisdom of Sparrow, Leadership, Trauma Healing, Sustainability, Sermons, Articles, Unknown.
- journal_entry_id: item_id of the parent Journal entry if applicable.
- essay_id: item_id of the parent Essay if applicable.
- paragraph_id: item_id of the parent Paragraph if applicable.
- sequence: integer sequence number preserving the order in the chunk.

HIERARCHY REQUIREMENTS:
For a dated entry with heading and multiple paragraphs, create:
1. Journal entry item for the whole dated entry.
2. Essay item if it contains a sustained reflection.
3. Paragraph items for reusable paragraphs.
4. One-liners/principles/questions/etc. inside the paragraphs.

Do NOT choose only the best level. Extract all usable levels.

CHUNK_ID: {chunk_id}

JOURNAL CHUNK:
{chunk_text}
"""


def create_prompts(docx_path: Path, target_words: int) -> Tuple[int, Path]:
    ensure_dirs()
    # Clear old prompts only, never clear AI results/output automatically.
    for old in PROMPTS_DIR.glob("*.txt"):
        old.unlink()

    text = read_docx_text(docx_path)
    entries = split_entries(text)
    chunks = chunk_entries(entries, target_words)
    base = safe_slug(docx_path.name)

    manifest_rows = []
    for idx, chunk in enumerate(chunks, start=1):
        chunk_id = f"{base}_{idx:04d}"
        prompt = build_prompt(chunk, chunk_id)
        prompt_path = PROMPTS_DIR / f"Prompt_{idx:04d}_{chunk_id}.txt"
        prompt_path.write_text(prompt, encoding="utf-8")
        date_match = DATE_LINE_RE.search(chunk)
        heading_match = HEADING_RE.search(chunk)
        manifest_rows.append({
            "prompt_file": prompt_path.name,
            "chunk_id": chunk_id,
            "word_count": word_count(chunk),
            "first_date": date_match.group(1).strip() if date_match else "",
            "first_heading": heading_match.group(1).strip() if heading_match else "",
        })

    manifest_path = PROMPTS_DIR / "manifest.csv"
    with manifest_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["prompt_file", "chunk_id", "word_count", "first_date", "first_heading"])
        writer.writeheader()
        writer.writerows(manifest_rows)

    return len(chunks), PROMPTS_DIR


def strip_json_fences(text: str) -> str:
    s = text.strip()
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s*```$", "", s)
    return s.strip()


def load_json_file(path: Path) -> List[Dict[str, Any]]:
    raw = path.read_text(encoding="utf-8-sig")
    raw = strip_json_fences(raw)
    data = json.loads(raw)
    if isinstance(data, dict):
        # Accept {"items": [...]} just in case.
        if "items" in data and isinstance(data["items"], list):
            data = data["items"]
        else:
            data = [data]
    if not isinstance(data, list):
        raise ValueError(f"Top-level JSON must be an array or object: {path.name}")
    clean: List[Dict[str, Any]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        row = {field: item.get(field, "") for field in CSV_FIELDS}
        # Normalize field values to strings except sequence can remain number but csv handles it.
        clean.append(row)
    return clean


def sort_key(row: Dict[str, Any]) -> Tuple[str, int, str]:
    chunk = str(row.get("chunk_id", ""))
    seq_raw = row.get("sequence", 0)
    try:
        seq = int(seq_raw)
    except Exception:
        seq = 0
    return (chunk, seq, str(row.get("item_id", "")))


def write_csv(path: Path, rows: List[Dict[str, Any]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in CSV_FIELDS})


def merge_json_results() -> Tuple[int, int, Path]:
    ensure_dirs()
    json_files = sorted(AI_JSON_DIR.glob("*.json")) + sorted(AI_JSON_DIR.glob("*.txt"))
    if not json_files:
        raise RuntimeError(f"No .json or .txt result files found in {AI_JSON_DIR}")

    all_rows: List[Dict[str, Any]] = []
    errors: List[str] = []
    for path in json_files:
        try:
            rows = load_json_file(path)
            all_rows.extend(rows)
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")
            log(f"ERROR loading {path}: {traceback.format_exc()}")

    all_rows.sort(key=sort_key)
    write_csv(OUTPUT_DIR / "Master_Database.csv", all_rows)

    for level, fname in LEVEL_TO_FILE.items():
        level_rows = [r for r in all_rows if r.get("level") == level]
        write_csv(OUTPUT_DIR / fname, level_rows)

    # Book view: one CSV per possible_book value.
    book_dir = OUTPUT_DIR / "Book_View"
    book_dir.mkdir(exist_ok=True)
    by_book: Dict[str, List[Dict[str, Any]]] = {}
    for row in all_rows:
        books = str(row.get("possible_book", "Unknown") or "Unknown")
        # Allow semicolon-separated multi-book tagging.
        for book in [b.strip() for b in books.split(";") if b.strip()]:
            by_book.setdefault(book, []).append(row)
    for book, rows in by_book.items():
        fname = safe_slug(book) + ".csv"
        write_csv(book_dir / fname, rows)

    # Error report if needed.
    if errors:
        (OUTPUT_DIR / "merge_errors.txt").write_text("\n".join(errors), encoding="utf-8")

    return len(json_files), len(all_rows), OUTPUT_DIR


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x520")
        self.docx_path: Optional[Path] = None
        ensure_dirs()
        self._build_ui()
        self.set_status("Ready.")

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        tk.Label(self, text=APP_TITLE, font=("Segoe UI", 16, "bold")).pack(anchor="w", **pad)

        frame = tk.Frame(self)
        frame.pack(fill="x", **pad)
        tk.Button(frame, text="Select Journal .docx", command=self.select_docx, width=24).pack(side="left")
        self.file_label = tk.Label(frame, text="No file selected", anchor="w")
        self.file_label.pack(side="left", padx=12, fill="x", expand=True)

        chunk_frame = tk.Frame(self)
        chunk_frame.pack(fill="x", **pad)
        tk.Label(chunk_frame, text="Words per chunk:").pack(side="left")
        self.words_var = tk.StringVar(value="1800")
        tk.Entry(chunk_frame, textvariable=self.words_var, width=8).pack(side="left", padx=5)
        tk.Label(chunk_frame, text="Recommended: 1500–2000 for careful extraction.").pack(side="left", padx=10)

        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", **pad)
        tk.Button(btn_frame, text="Create Hierarchical Prompt Files", command=self.on_create_prompts, width=32).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Merge JSON Results into CSVs", command=self.on_merge, width=28).pack(side="left", padx=4)

        open_frame = tk.Frame(self)
        open_frame.pack(fill="x", **pad)
        tk.Button(open_frame, text="Open prompts_to_paste", command=lambda: self.open_folder(PROMPTS_DIR), width=22).pack(side="left", padx=4)
        tk.Button(open_frame, text="Open ai_json_results", command=lambda: self.open_folder(AI_JSON_DIR), width=22).pack(side="left", padx=4)
        tk.Button(open_frame, text="Open output", command=lambda: self.open_folder(OUTPUT_DIR), width=15).pack(side="left", padx=4)

        instructions = (
            "Workflow:\n"
            "1. Select your .docx journal.\n"
            "2. Create prompt files.\n"
            "3. Open prompts_to_paste and copy Prompt_0001 into ChatGPT.\n"
            "4. Save ChatGPT's JSON-only answer as Chunk_0001.json in ai_json_results.\n"
            "5. Repeat for each prompt.\n"
            "6. Click Merge JSON Results into CSVs.\n\n"
            "Outputs: Master_Database.csv plus separate Journal Entries, Essays, Paragraphs, One-Liners, Principles, Stories, and Book_View CSVs."
        )
        tk.Label(self, text=instructions, justify="left", anchor="w", font=("Segoe UI", 10)).pack(fill="x", **pad)

        self.status_text = tk.Text(self, height=10, wrap="word")
        self.status_text.pack(fill="both", expand=True, padx=10, pady=10)

    def set_status(self, text: str) -> None:
        self.status_text.insert("end", text.rstrip() + "\n")
        self.status_text.see("end")
        self.update_idletasks()
        log(text)

    def select_docx(self) -> None:
        filename = filedialog.askopenfilename(
            title="Select journal .docx file",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if filename:
            self.docx_path = Path(filename)
            self.file_label.config(text=str(self.docx_path))
            self.set_status(f"Selected: {self.docx_path}")

    def on_create_prompts(self) -> None:
        if not self.docx_path:
            messagebox.showwarning(APP_TITLE, "Please select a .docx journal first.")
            return
        try:
            target = int(self.words_var.get().strip())
            self.set_status("Creating prompt files...")
            count, folder = create_prompts(self.docx_path, target)
            self.set_status(f"Created {count} prompt files in: {folder}")
            messagebox.showinfo(APP_TITLE, f"Created {count} prompt files.\n\nFolder:\n{folder}")
        except Exception as exc:
            self.set_status("ERROR creating prompts: " + str(exc))
            log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, str(exc))

    def on_merge(self) -> None:
        try:
            self.set_status("Merging JSON results...")
            files, rows, folder = merge_json_results()
            self.set_status(f"Merged {files} result files into {rows} database rows.")
            self.set_status(f"Output folder: {folder}")
            messagebox.showinfo(APP_TITLE, f"Merged {files} JSON files into {rows} rows.\n\nOutput:\n{folder}")
        except Exception as exc:
            self.set_status("ERROR merging results: " + str(exc))
            log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, str(exc))

    def open_folder(self, folder: Path) -> None:
        ensure_dirs()
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open folder:\n{folder}\n\n{exc}")


if __name__ == "__main__":
    ensure_dirs()
    app = App()
    app.mainloop()
