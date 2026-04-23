import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st


WORKBOOK_PATH = Path("./Excel_Readiness_AI_Coach_Content_Pack_with_sources.xlsx")

# Local-only file (gitignored): put your sk- key on one line next to app.py — for `streamlit run` on your PC.
# Do NOT commit the real file. Public websites should use Streamlit Cloud → Settings → Secrets instead.
_LOCAL_OPENAI_KEY_FILE = "openai_key_local.txt"
_local_key_file_read: bool = False


def _load_local_openai_key_file() -> None:
    """Load OPENAI_API_KEY from openai_key_local.txt once (never commit that file)."""
    global _local_key_file_read
    if _local_key_file_read:
        return
    _local_key_file_read = True
    try:
        p = Path(__file__).resolve().parent / _LOCAL_OPENAI_KEY_FILE
        if not p.is_file():
            return
        raw = p.read_text(encoding="utf-8").strip()
        for line in raw.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("sk-") and len(line) > 20:
                os.environ.setdefault("OPENAI_API_KEY", line)
                return
    except OSError:
        pass


def normalize_col_name(name: str) -> str:
    return "".join(ch for ch in str(name).strip().lower() if ch.isalnum())


def find_col(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    normalized_map = {normalize_col_name(col): col for col in df.columns}
    for alias in aliases:
        key = normalize_col_name(alias)
        if key in normalized_map:
            return normalized_map[key]
    return None


def load_workbook_data(path: Path) -> Dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(str(path))

    xls = pd.ExcelFile(path)
    expected_sheets = [
        "Competency Map",
        "Question Bank",
        "Material Library",
        "AI Prompt Template",
    ]

    missing = [s for s in expected_sheets if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"Missing required sheet(s): {', '.join(missing)}")

    return {sheet: pd.read_excel(path, sheet_name=sheet) for sheet in expected_sheets}


def parse_questions(question_df: pd.DataFrame) -> List[Dict]:
    q_col = find_col(question_df, ["Question", "Question Text", "Prompt"])
    module_id_col = find_col(
        question_df,
        ["Module ID", "Module", "Module Name", "Competency", "Topic"],
    )
    module_name_col = find_col(question_df, ["Module Name", "Module", "Competency"])
    answer_col = find_col(question_df, ["Correct Answer", "Answer", "Correct", "Key"])
    option_cols = [
        find_col(question_df, ["Option A", "A", "Choice A", "Option 1"]),
        find_col(question_df, ["Option B", "B", "Choice B", "Option 2"]),
        find_col(question_df, ["Option C", "C", "Choice C", "Option 3"]),
        find_col(question_df, ["Option D", "D", "Choice D", "Option 4"]),
    ]

    if not q_col or not answer_col or any(c is None for c in option_cols):
        raise ValueError(
            "Question Bank is missing required columns. Needed: question, four options (A-D), and correct answer."
        )

    cleaned_questions: List[Dict] = []
    for _, row in question_df.iterrows():
        q_text = str(row.get(q_col, "")).strip()
        if not q_text or q_text.lower() == "nan":
            continue

        options = [str(row.get(c, "")).strip() for c in option_cols]
        options = [opt for opt in options if opt and opt.lower() != "nan"]
        if len(options) < 4:
            continue

        answer_raw = str(row.get(answer_col, "")).strip()
        module = "General"
        if module_id_col:
            module_raw = str(row.get(module_id_col, "")).strip()
            if module_raw and module_raw.lower() != "nan":
                m_id = re.search(r"\bM\d+\b", module_raw, re.I)
                if m_id:
                    module = m_id.group(0).upper()
                else:
                    module = module_raw
        module_name = ""
        if module_name_col:
            module_name = str(row.get(module_name_col, "")).strip()
            if module_name.lower() in ("", "nan"):
                module_name = ""

        cleaned_questions.append(
            {
                "question": q_text,
                "options": options[:4],
                "answer": answer_raw,
                "module": module,
                "module_name": module_name,
            }
        )

    if not cleaned_questions:
        raise ValueError("No valid questions found in Question Bank.")

    return cleaned_questions


QUIZ_EMPTY_PLACEHOLDER = "— Select an answer —"


def answer_matches(selected: str, answer_key: str, options: List[str]) -> bool:
    if selected is None or str(selected).strip() == "" or str(selected) == QUIZ_EMPTY_PLACEHOLDER:
        return False
    selected_norm = str(selected).strip().lower()
    answer_norm = str(answer_key).strip().lower()

    if selected_norm == answer_norm:
        return True

    letter_map = {"a": 0, "b": 1, "c": 2, "d": 3}
    if answer_norm in letter_map and letter_map[answer_norm] < len(options):
        return selected_norm == str(options[letter_map[answer_norm]]).strip().lower()

    return False


def evaluate_quiz(questions: List[Dict], responses: Dict[int, str]) -> Dict:
    total = len(questions)
    correct = 0
    module_stats: Dict[str, Dict[str, int]] = {}

    for i, q in enumerate(questions):
        module = q["module"]
        module_stats.setdefault(module, {"correct": 0, "total": 0})
        module_stats[module]["total"] += 1

        selected = responses.get(i, "")
        is_correct = answer_matches(selected, q["answer"], q["options"])
        if is_correct:
            correct += 1
            module_stats[module]["correct"] += 1

    score_pct = (correct / total * 100) if total else 0
    if score_pct >= 80:
        level = "Advanced / Pass"
    elif score_pct >= 50:
        level = "Intermediate"
    else:
        level = "Beginner"

    strong_modules = []
    weak_modules = []
    for module, stat in module_stats.items():
        module_score = (stat["correct"] / stat["total"]) * 100 if stat["total"] else 0
        if module_score >= 70:
            strong_modules.append(module)
        else:
            weak_modules.append(module)

    return {
        "total_questions": total,
        "correct_answers": correct,
        "score_pct": score_pct,
        "level": level,
        "strong_modules": strong_modules,
        "weak_modules": weak_modules,
        "module_stats": module_stats,
    }


def normalize_text(value: str) -> str:
    return str(value).strip().lower()


_EMPTY_RECS: pd.DataFrame = pd.DataFrame(
    columns=[
        "Title",
        "Module",
        "Level",
        "Type",
        "Description",
        "Author",
        "ISBN",
        "_module_id_raw",
    ]
)


def _parse_module_ids_in_cell(val: object) -> set:
    """Extract M1, M2, M2 from cells like 'M2/M4' or 'M2'."""
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return set()
    return {m.upper() for m in re.findall(r"M\d+", s, re.I)}


def _parse_weak_module_ids(weak_modules: List[str]) -> set:
    """Align diagnostic module labels (e.g. M2) with Material Library Module ID."""
    out: set = set()
    for w in weak_modules:
        s = str(w).strip()
        if not s or s.lower() == "nan":
            continue
        m = re.search(r"\bM\d+\b", s, re.I)
        if m:
            out.add(m.group(0).upper())
    return out


def _normalize_author_cell(val: object) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    return s


def _normalize_isbn_cell(val: object) -> str:
    """ISBNs are often stored as numbers in Excel; avoid 978...e+12 style strings."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if int(val) == val:
            return str(int(val))
    s = str(val).strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    return s


def _finalize_material_frame(
    raw: pd.DataFrame,
    title_col: str,
    module_name_col: Optional[str],
    level_col: Optional[str],
    type_col: Optional[str],
    desc_col: Optional[str],
    module_id_col: str,
    author_col: Optional[str],
    isbn_col: Optional[str],
) -> pd.DataFrame:
    """Map workbook columns to student-friendly card fields. Drops backend-only columns."""
    t = raw.copy()
    t["Title"] = t[title_col].astype(str).str.strip()
    t["Module"] = (
        t[module_name_col].astype(str).str.strip()
        if module_name_col
        else t[module_id_col].astype(str).str.strip()
    )
    t["Level"] = t[level_col].astype(str).str.strip() if level_col else ""
    t["Type"] = t[type_col].astype(str).str.strip() if type_col else ""
    t["Description"] = t[desc_col].astype(str).str.strip() if desc_col else ""
    t["Author"] = t[author_col].map(_normalize_author_cell) if author_col else ""
    t["ISBN"] = t[isbn_col].map(_normalize_isbn_cell) if isbn_col else ""
    t["_module_id_raw"] = t[module_id_col].astype(str).str.strip()
    return t[
        [
            "Title",
            "Module",
            "Level",
            "Type",
            "Description",
            "Author",
            "ISBN",
            "_module_id_raw",
        ]
    ]


def get_recommended_materials(
    material_df: pd.DataFrame,
    weak_modules: List[str],
    level: str,
    score_pct: float,
) -> Tuple[pd.DataFrame, str]:
    """
    Match materials by Module ID (supports compound IDs like M2/M4).
    Fallback: optional Advanced materials if no weak areas + strong score;
    else a short M1 foundational refresh. Returns (df, rec_mode) for 'why' copy.
    """
    if material_df.empty:
        return _EMPTY_RECS.copy(), "none"

    module_id_col = find_col(
        material_df, ["Module ID", "ModuleID", "Module Code"]
    )
    module_name_col = find_col(
        material_df, ["Module Name", "Skill Area", "Module", "Topic"]
    )
    level_col = find_col(material_df, ["Level", "Difficulty"])
    title_col = find_col(
        material_df, ["Material Title", "Title", "Material", "Name", "Resource Name"]
    )
    type_col = find_col(material_df, ["Type", "Format", "Resource Type"])
    desc_col = find_col(
        material_df, ["Description", "Short Description", "Summary", "Details", "Desc"]
    )
    author_col = find_col(
        material_df,
        ["Author", "Authors", "Writer", "By", "Name of Author", "Book Author"],
    )
    isbn_col = find_col(
        material_df, ["ISBN", "ISBN-13", "ISBN13", "ISBN Number", "Book ISBN", "EAN"],
    )

    if not title_col or not module_id_col:
        return _EMPTY_RECS.copy(), "none"

    weak_ids = _parse_weak_module_ids(weak_modules)
    picked: Optional[pd.DataFrame] = None
    rec_mode: str = "none"

    # 1) Primary: match weak Module ID(s) to library Module ID (including M2/M4)
    if weak_ids:
        idxs: list = []
        for i, row in material_df.iterrows():
            mcell = row.get(module_id_col)
            mat_ids = _parse_module_ids_in_cell(mcell)
            if mat_ids & weak_ids:
                idxs.append(i)
        if idxs:
            picked = material_df.loc[idxs].drop_duplicates()
            rec_mode = "weak"

    # 1b) Name-based fallback (e.g. legacy "General" or name-only matches)
    if rec_mode == "none" and weak_modules:
        idxs2: list = []
        for w in weak_modules:
            wlow = str(w).lower().strip()
            for i, row in material_df.iterrows():
                mname = (
                    str(row.get(module_name_col, "")).lower()
                    if module_name_col
                    else ""
                )
                if not mname:
                    continue
                if (len(wlow) > 2 and wlow in mname) or (len(mname) > 3 and mname in wlow):
                    idxs2.append(i)
        if idxs2:
            picked = material_df.loc[idxs2].drop_duplicates()
            rec_mode = "weak"

    # 2) No weak areas (all modules strong): optional Advanced, else short M1 refresh
    if rec_mode == "none" and not weak_modules:
        is_advanced_student = (level == "Advanced / Pass") or (score_pct >= 80.0)
        if is_advanced_student and level_col:
            mask_adv = material_df[level_col].astype(str).str.contains(
                "Advanced", case=False, na=False
            )
            m7 = (
                material_df[module_id_col]
                .astype(str)
                .str.contains("M7", case=False, na=False)
            )
            sub = material_df[mask_adv | m7]
            if not sub.empty:
                picked = sub.drop_duplicates()
                rec_mode = "advanced_optional"
        if rec_mode == "none" and level_col and module_id_col:
            m1b = material_df[module_id_col].astype(str).str.strip().str.upper().eq("M1")
            beg = material_df[level_col].astype(str).str.strip().str.lower() == "beginner"
            sub2 = material_df[m1b & beg].head(2)
            if not sub2.empty:
                picked = sub2.copy()
                rec_mode = "foundational_refresh"

    if rec_mode == "none" or picked is None or picked.empty:
        return _EMPTY_RECS.copy(), "none"

    out = _finalize_material_frame(
        picked,
        title_col,
        module_name_col,
        level_col,
        type_col,
        desc_col,
        module_id_col,
        author_col,
        isbn_col,
    )
    return out, rec_mode


def build_targeted_practice_task(weak_modules: List[str], level: str) -> str:
    """Single concrete next practice step for the study guide page."""
    if not weak_modules:
        return (
            "Spend 20 minutes in Excel: replicate one class-style mini workflow (small dataset → "
            "a few formulas → a short table or chart) and add two sentences explaining the business reason."
        )
    first = ", ".join(weak_modules[:3])
    if len(weak_modules) > 3:
        first += ", and related areas"
    return (
        f"Spend 20–30 minutes in Excel on **{first}**: build a 10-row example from scratch, "
        f"apply at least one lookup or error-checking technique, and write a 3-bullet "
        f'"client-ready" summary. Your current readiness is **{level}**—focus on accuracy and clarity, not speed.'
    )


def _display_citation_field(val: object) -> str:
    """String for showing Author/ISBN in the UI; treats Excel blanks/NaN as empty."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    return s


def _format_short_description(text: str, max_len: int = 320) -> str:
    t = str(text or "").strip()
    if not t or t.lower() == "nan":
        return "—"
    if len(t) > max_len:
        return t[: max_len - 1].rstrip() + "…"
    return t


def _why_recommended_row(
    row: pd.Series,
    student_level: str,
    rec_mode: str,
    weak_id_set: Set[str],
) -> str:
    skill = str(row.get("Module", "") or "").strip() or "this skill area"
    mid_raw = str(row.get("_module_id_raw", ""))
    mat_ids = _parse_module_ids_in_cell(mid_raw)
    overlap = mat_ids & weak_id_set if weak_id_set else set()
    if rec_mode == "advanced_optional":
        return (
            "Optional stretch: the diagnostic did not flag a weak **Module ID**, so this **Advanced** "
            "item from *your* content pack (e.g. deeper analytics) is suggested for next-level professional Excel use in accounting and finance."
        )
    if rec_mode == "foundational_refresh":
        return (
            "Foundational refresh: a short **M1** item to keep layout, structure, and formatting audit-ready for "
            "deliverables you will submit in accounting and finance courses."
        )
    if rec_mode == "weak" and overlap:
        pretty = ", ".join(sorted(overlap))
        return (
            f"Your results pointed to **{pretty}** — *{skill}* — for stronger Excel and business workflow "
            f"readiness in audit-style documentation and finance analysis."
        )
    if rec_mode == "weak":
        return (
            f"Aligned with a focus area the quiz identified for extra practice: **{skill}**, to support clear "
            "workpapers and time-saving routines common in accounting and finance roles."
        )
    return f"From your course **Material Library**, supporting **{skill}** (readiness: {student_level})."


def render_resource_cards(
    rec_df: pd.DataFrame,
    student_level: str,
    rec_mode: str,
    weak_id_set: Set[str],
) -> None:
    if rec_df is None or rec_df.empty:
        if rec_mode == "none":
            st.info(
                "No content-pack resources to show — try running the diagnostic again, or check that the "
                "workbook’s Question Bank and Material Library are loaded."
            )
        return
    for _, row in rec_df.iterrows():
        title = str(row.get("Title", "Resource") or "Resource").strip() or "Resource"
        area = str(row.get("Module", "—") or "—").strip() or "—"
        level_v = str(row.get("Level", "") or "").strip()
        mtype = str(row.get("Type", "") or "").strip()
        desc = _format_short_description(str(row.get("Description", "")))
        author = _display_citation_field(row.get("Author", ""))
        isbn = _display_citation_field(row.get("ISBN", ""))
        has_level_on_row = bool(level_v and level_v not in ("—", "nan"))
        mtype_d = mtype if mtype and mtype.lower() != "nan" else "—"
        level_d = level_v if has_level_on_row else "—"
        why = _why_recommended_row(row, student_level, rec_mode, weak_id_set)
        emdash = "—"
        author_line = author if author else emdash
        isbn_line = isbn if isbn else emdash

        with st.container(border=True):
            st.markdown(f"**{title}**")
            col_src_a, col_src_b = st.columns(2)
            with col_src_a:
                st.caption("Author")
                st.write(author_line)
            with col_src_b:
                st.caption("ISBN")
                st.write(isbn_line)
            col_a, col_b, col_t = st.columns(3)
            col_a.caption("Skill area (module name)")
            col_a.write(area)
            col_b.caption("Level")
            col_b.write(level_d)
            col_t.caption("Type")
            col_t.write(mtype_d)
            st.caption("Short description")
            st.write(desc)
            st.caption("Why this was recommended")
            st.write(why)


def extract_prompt_template(prompt_df: pd.DataFrame) -> str:
    if prompt_df.empty:
        return (
            "Create a personalized study guide for an accounting/finance student. "
            "Include weekly goals, practice tasks, and confidence tips."
        )

    text_parts = []
    for _, row in prompt_df.iterrows():
        for cell in row.tolist():
            cell_text = str(cell).strip()
            if cell_text and cell_text.lower() != "nan":
                text_parts.append(cell_text)

    if not text_parts:
        return (
            "Create a personalized study guide for an accounting/finance student. "
            "Include weekly goals, practice tasks, and confidence tips."
        )

    return "\n".join(text_parts)


def build_study_guide_prompt(
    level: str,
    weak_modules: List[str],
    materials: pd.DataFrame,
    template: str,
    score_pct: float,
) -> str:
    pack_lines: List[str] = []
    if not materials.empty:
        for _, row in materials.iterrows():
            title = str(row.get("Title", "")).strip() or "Untitled"
            skill = str(row.get("Module", "")).strip() or "—"
            mtype = str(row.get("Type", "")).strip() or "—"
            author = str(row.get("Author", "") or "").strip()
            isbn = str(row.get("ISBN", "") or "").strip()
            if author.lower() in ("", "nan"):
                author = ""
            if isbn.lower() in ("", "nan"):
                isbn = ""
            book_bits = []
            if author:
                book_bits.append(f"Author: {author}")
            if isbn:
                book_bits.append(f"ISBN: {isbn}")
            if book_bits:
                pack_lines.append(
                    f"- **{title}** — {skill} — ({mtype}) — " + " · ".join(book_bits)
                )
            else:
                pack_lines.append(f"- **{title}** — {skill} — ({mtype})")

    pack_text = "\n".join(pack_lines) if pack_lines else "*(No extra rows; give general course-coach advice only, still with no outside resources.)*"
    weak_text = (
        ", ".join(weak_modules)
        if weak_modules
        else "None — all tracked module areas met the strong band on the quiz (or use Advanced optional path)."
    )

    return f"""{template}

You are the **Excel Readiness AI Coach** for *new* accounting and finance students. Emphasize Excel habits and business **workflow** readiness: clean workpapers, traceable numbers, tie-outs, and professional documentation — the kind of skills used in **accounting, audit, and corporate finance** starter roles.

**Strict rules (must follow):**
- Suggest **only** the course materials listed under “Approved content-pack resources” below. Do **not** name or recommend **Khan Academy, Coursera, YouTube**, random blogs, or any third-party products or sites.
- Keep the full answer **short** (aim under ~500 words, tight bullets where possible).

**From our diagnostic:**
- Readiness level: {level}
- Overall score: {score_pct:.1f}%
- Weak **Module ID(s) / focus areas** from the quiz: {weak_text}

**Approved content-pack resources (Material Library only):**
{pack_text}

**Write the study guide with exactly these Markdown sections in order. Use `##` headings exactly as written:**
## Readiness summary
## Weak areas
## Why these areas matter (accounting / audit / finance)
## Recommended learning path
## Targeted practice task
## Recommended resources
In **Recommended learning path** and **Recommended resources**, map actions explicitly to the **titles** of the approved resources when available, and when **Author** and/or **ISBN** are listed for an item, include them alongside the title for proper citation.
"""


def _placeholder_study_guide(prompt: str) -> str:
    return f"""## Readiness summary
This offline sample reflects an Excel + business workflow focus for new accounting and finance students (set **OPENAI_API_KEY** for a live, tailored response).

## Weak areas
(Sample) Prioritize the module IDs the quiz flags below 70% in your class spreadsheet — not generic “Excel tips.”

## Why these areas matter (accounting / audit / finance)
Clean workbooks, consistent formatting, and traceable numbers reduce review notes and make tie-outs and footnotes easier for managers and audit teams.

## Recommended learning path
Work only through the **Material Library** items that match your weak **Module ID(s)** (M1, M2, M3, …) in the content pack, in the order that moves from setup → formulas → matching → workpaper quality.

## Targeted practice task
Build a 10-row practice sheet, apply one lookup and one check formula, and write 3 bullet points explaining an auditor would re-perform your numbers.

## Recommended resources
Re-open the list under “Approved content-pack resources” in the app — those are the only items to use by title.

---
*(Prompt length {len(prompt)} characters — truncated in logs only.)*"""


def generate_ai_study_guide(prompt: str) -> str:
    """
    Uses the OpenAI API when OPENAI_API_KEY is set; otherwise returns a local placeholder.
    """
    key = get_openai_api_key()
    if not key:
        return _placeholder_study_guide(prompt)

    coach_system = (
        "You are the Excel Readiness AI Coach for new accounting and finance students. "
        "Follow the user's required Markdown section headings and rules exactly. "
        "Never suggest Khan Academy, Coursera, YouTube, or any outside learning sites — only the course "
        "content-pack / Material Library items the user listed. When author and ISBN are provided for a resource, "
        "mention them. Be concise; avoid generic filler."
    )
    try:
        from openai import OpenAI

        client = OpenAI(api_key=key)
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": coach_system},
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            temperature=0.55,
            max_tokens=900,
        )
        text = (response.choices[0].message.content or "").strip()
        if not text:
            return _placeholder_study_guide(prompt)
        return text
    except Exception as exc:
        return (
            "Could not reach the OpenAI API. Showing the offline version instead.\n\n"
            f"**Error:** {exc}\n\n"
            + _placeholder_study_guide(prompt)
        )


def initialize_session_state() -> None:
    defaults = {
        "responses": {},
        "submitted": False,
        "results": None,
        "recommendations": pd.DataFrame(),
        "study_prompt": "",
        "study_guide": "",
        "rec_mode": "none",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if "app_step" not in st.session_state:
        st.session_state["app_step"] = (
            "post_quiz"
            if st.session_state.get("submitted")
            and st.session_state.get("results") is not None
            else "welcome"
        )


def reset_diagnostic_state() -> None:
    """Clear quiz answers, scores, and study guide so the student can retake the diagnostic."""
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and re.match(r"^q_\d+$", k):
            del st.session_state[k]
    st.session_state["responses"] = {}
    st.session_state["submitted"] = False
    st.session_state["results"] = None
    st.session_state["recommendations"] = _EMPTY_RECS.copy()
    st.session_state["study_prompt"] = ""
    st.session_state["study_guide"] = ""
    st.session_state["rec_mode"] = "none"
    st.session_state["app_step"] = "welcome"


def page_welcome() -> None:
    st.title("Excel Readiness AI Coach")
    st.subheader("Welcome")
    st.write(
        "This tool helps diagnose Excel and business workflow readiness for new accounting and finance students. "
        "Complete the diagnostic quiz, review your readiness level, and receive targeted learning materials "
        "plus an AI-generated study guide."
    )
    st.info(
        "The flow is linear: **Welcome** → **Diagnostic** → **Results** and **AI study guide** on one page after you submit. "
        "No sidebar menu—just follow the steps."
    )
    if st.button("Start diagnostic", type="primary", use_container_width=True):
        st.session_state["app_step"] = "quiz"
        st.rerun()
    if st.session_state.get("submitted") and st.session_state.get("results") is not None:
        if st.button("Continue to your results & study guide", use_container_width=True):
            st.session_state["app_step"] = "post_quiz"
            st.rerun()
        st.caption("Or start a **new** run (clears answers, score, and study guide):")
        if st.button("Start over", use_container_width=True):
            reset_diagnostic_state()
            st.rerun()


def page_diagnostic_quiz(questions: List[Dict]) -> None:
    t1, t2 = st.columns([3, 1])
    with t1:
        st.title("Diagnostic Quiz")
    with t2:
        if st.button("Back to home", use_container_width=True, help="Return to the welcome page"):
            st.session_state["app_step"] = "welcome"
            st.rerun()
    st.write("Answer each question using the multiple-choice options.")

    with st.form("quiz_form"):
        responses: Dict[int, Optional[str]] = {}
        for idx, q in enumerate(questions):
            st.markdown(f"**Q{idx + 1}. {q['question']}**")
            mod_label = str(q.get("module", "General"))
            mname = str(q.get("module_name", "") or "").strip()
            st.caption(f"Module: {mod_label}" + (f" — {mname}" if mname else ""))
            # First option is a non-answer placeholder so no real choice is pre-selected
            full_options: List[str] = [QUIZ_EMPTY_PLACEHOLDER] + list(q["options"])
            if idx in st.session_state.get("responses", {}):
                existing = st.session_state["responses"][idx]
                if existing in q["options"]:
                    default_index = 1 + q["options"].index(existing)
                else:
                    default_index = 0
            else:
                default_index = 0

            choice = st.radio(
                label=f"Select an answer for question {idx + 1}",
                options=full_options,
                key=f"q_{idx}",
                index=default_index,
                label_visibility="collapsed",
            )
            responses[idx] = None if choice == QUIZ_EMPTY_PLACEHOLDER else choice
            st.divider()

        submitted = st.form_submit_button("Submit Answers", use_container_width=True)

    if submitted:
        unanswered = [i + 1 for i in range(len(questions)) if responses.get(i) is None]
        if unanswered:
            st.error(
                "Please select an answer for every question before submitting. "
                f"Still need an answer for: {', '.join(f'Q{n}' for n in unanswered)}."
            )
        else:
            st.session_state["responses"] = {i: responses[i] for i in range(len(questions))}
            results = evaluate_quiz(questions, st.session_state["responses"])
            st.session_state["results"] = results
            st.session_state["submitted"] = True
            st.session_state["app_step"] = "post_quiz"
            st.rerun()


def page_results(
    *, hide_outer_heading: bool = False, hide_caption: bool = False
) -> bool:
    if not st.session_state["submitted"] or st.session_state["results"] is None:
        st.warning("Please complete and submit the Diagnostic Quiz first.")
        return False
    if not hide_outer_heading:
        st.subheader("Your results")
    if not hide_caption:
        st.caption("Based on this diagnostic only—not a course grade.")

    results = st.session_state["results"]
    st.metric("Total Score", f"{results['score_pct']:.1f}%")
    st.metric("Classification", results["level"])
    st.write(f"Correct Answers: **{results['correct_answers']} / {results['total_questions']}**")

    strong = results["strong_modules"] if results["strong_modules"] else ["None yet"]
    weak = results["weak_modules"] if results["weak_modules"] else ["None identified"]

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Strong Modules")
        for item in strong:
            st.write(f"- {item}")
    with col2:
        st.subheader("Weak Modules")
        for item in weak:
            st.write(f"- {item}")

    stats_rows = []
    for module, stat in results["module_stats"].items():
        pct = (stat["correct"] / stat["total"] * 100) if stat["total"] else 0
        stats_rows.append(
            {
                "Module": module,
                "Correct": stat["correct"],
                "Total": stat["total"],
                "Score %": round(pct, 1),
            }
        )
    if stats_rows:
        st.subheader("Module Breakdown")
        st.dataframe(pd.DataFrame(stats_rows), use_container_width=True, hide_index=True)
    return True


def page_ai_study_guide(
    prompt_template_df: pd.DataFrame,
    material_df: pd.DataFrame,
    *,
    embedded: bool = False,
) -> None:
    if embedded:
        st.subheader("AI study guide & resources")
    else:
        st.title("AI Study Guide")

    if not st.session_state["submitted"] or st.session_state["results"] is None:
        st.warning("Please complete the quiz first.")
        return

    results = st.session_state["results"]
    level = results["level"]
    score_pct = results["score_pct"]
    weak_modules = results["weak_modules"]
    weak_display = (
        ", ".join(weak_modules) if weak_modules else "— (no weak areas identified — great work)"
    )

    recommendations, rec_mode = get_recommended_materials(
        material_df, weak_modules, level, score_pct
    )
    st.session_state["recommendations"] = recommendations
    st.session_state["rec_mode"] = rec_mode
    weak_id_set = _parse_weak_module_ids(weak_modules)

    st.subheader("Your snapshot")
    m1, m2 = st.columns(2)
    m1.metric("1. Student level", level)
    m2.metric("2. Overall score", f"{score_pct:.1f}%")
    st.markdown("**3. Weak modules**")
    st.write(weak_display)

    template = extract_prompt_template(prompt_template_df)
    prompt = build_study_guide_prompt(
        level=level,
        weak_modules=weak_modules,
        materials=recommendations,
        template=template,
        score_pct=score_pct,
    )
    st.session_state["study_prompt"] = prompt

    st.caption(
        "**OpenAI (live text):** configure **Streamlit Cloud → Settings → Secrets** with `OPENAI_API_KEY`, "
        "or for local runs use `openai_key_local.txt` (see `openai_key_local.txt.example`)."
    )

    key_set = bool(get_openai_api_key())
    if key_set:
        st.caption("API key is available — **Generate** will call **gpt-4o-mini** (or shows an error if the key is invalid).")
    else:
        st.caption("No key yet — **Generate** will use the **offline sample** text.")
        with st.expander("How to enable live AI (Streamlit Cloud + local)", expanded=False):
            st.markdown(
                "**Streamlit Community Cloud (recommended, persistent):**  \n"
                "App **⋮** → **Settings** → **Secrets** — paste in the editor (exact format):\n\n"
                "```toml\n"
                'OPENAI_API_KEY = "sk-proj-xxxxxxxx"\n'
                "```\n\n"
                "Click **Save** → **Reboot** app. Wait 1–2 min.  \n"
                "**Troubleshooting:** key on **one line** in double quotes; valid TOML.\n\n"
                "**On your computer:** env var `OPENAI_API_KEY`, or `.streamlit/secrets.toml`, or `openai_key_local.txt`."
            )

    if st.button("Generate AI Study Guide", use_container_width=True, type="primary"):
        with st.spinner("Calling the model to build your study guide…" if get_openai_api_key() else "Building the sample study guide (no key)…"):
            st.session_state["study_guide"] = generate_ai_study_guide(prompt)
        st.success("Done. Scroll to **4. Personalized study guide** below.")

    st.subheader("4. Personalized study guide")
    if st.session_state.get("study_guide"):
        st.write(st.session_state["study_guide"])
    else:
        st.info(
            "Click **Generate AI Study Guide** above. If an OpenAI API key is set, you will get a live response; "
            "otherwise you will see the built-in sample text."
        )

    st.subheader("5. Targeted practice task")
    st.write(build_targeted_practice_task(weak_modules, level))

    st.subheader("6. Recommended resources")
    if rec_mode == "weak":
        st.caption(
            "Pulled from **your** Material Library by **Module ID** (M1–M6, M7), including combined IDs like **M2/M4**."
        )
    elif rec_mode == "advanced_optional":
        st.caption(
            "You had no weak **Module ID** on the quiz; showing **optional Advanced** items from the content pack."
        )
    elif rec_mode == "foundational_refresh":
        st.caption("Light **M1** refresh from the content pack for strong all-around results.")
    else:
        st.caption("From the Material Library (no matches this run).")
    st.caption(
        "Each card shows **Author** and **ISBN** from your **Material Library** in Excel. "
        "A dash (—) means that cell is still blank in the file."
    )
    render_resource_cards(
        recommendations, level, rec_mode, weak_id_set
    )

    with st.expander("View generated API prompt (for your future OpenAI connection)", expanded=False):
        st.code(st.session_state.get("study_prompt", prompt), language="markdown")


def page_post_quiz(
    prompt_template_df: pd.DataFrame, material_df: pd.DataFrame
) -> None:
    """After the quiz: show results, then a short bridge, then the AI study guide (one scrollable page)."""
    st.title("Your results & study plan")
    st.caption("Results first. Below that, your AI study guide and resource cards—scroll when you are ready.")
    st.subheader("1. Your results")
    if not page_results(hide_outer_heading=True, hide_caption=True):
        st.error("We could not load your scores. Return to the quiz and submit your answers again.")
        if st.button("Go to the diagnostic", type="primary"):
            st.session_state["app_step"] = "quiz"
            st.rerun()
        return

    st.divider()
    st.subheader("2. What’s next?")
    st.write(
        "Next, use **AI study guide & resources** to get a short personalized plan, a practice task, and "
        "materials matched to your weak modules. Scroll down to the next section, or use **Start over** at the bottom "
        "if you want a fresh run."
    )
    page_ai_study_guide(
        prompt_template_df,
        material_df,
        embedded=True,
    )
    st.divider()
    if st.button("Start over (new diagnostic run)", use_container_width=True, type="secondary"):
        reset_diagnostic_state()
        st.rerun()


def get_openai_api_key() -> str:
    """
    Resolve OpenAI key: env var → local file (openai_key_local.txt) → st.secrets (Streamlit Cloud).
    """
    _load_local_openai_key_file()
    k = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if k and not k.lower().startswith("sk-placeholder") and "paste" not in k.lower():
        if len(k) > 10:
            return k

    def _from_secrets() -> str:
        try:
            sec = st.secrets
        except (FileNotFoundError, OSError, RuntimeError):
            return ""
        if sec is None:
            return ""

        def try_get(n: str) -> str:
            for nm in (n, n.lower(), n.upper()):
                try:
                    v = sec[nm]  # type: ignore
                    s = str(v).strip() if v is not None else ""
                    if s and s not in ("None", ""):
                        return s
                except (KeyError, TypeError, IndexError):
                    try:
                        v = getattr(sec, nm, None)
                        if v is not None:
                            s = str(v).strip()
                            if s:
                                return s
                    except Exception:
                        pass
            return ""

        for n in (
            "OPENAI_API_KEY",
            "openai_api_key",
            "OPENAI_KEY",
            "api_key",
        ):
            val = try_get(n)
            if val:
                return val

        # Nested [openai] api_key = "..." in TOML
        try:
            inner = sec["openai"]
            if isinstance(inner, dict):
                for kk in ("api_key", "OPENAI_API_KEY", "key", "Key"):
                    if inner.get(kk):
                        s = str(inner[kk]).strip()
                        if s:
                            return s
        except Exception:
            pass

        # Any secret whose name contains OPENAI
        try:
            keys = list(sec.keys())  # type: ignore
        except Exception:
            keys = []
        for name in keys:
            if "openai" in str(name).lower() or (str(name).lower() == "api_key"):
                try:
                    val = str(sec[name]).strip()  # type: ignore
                except Exception:
                    continue
                if val and val.startswith("sk-"):
                    return val

        return ""

    skey = _from_secrets()
    if skey:
        os.environ["OPENAI_API_KEY"] = skey
        return skey

    return ""


def _apply_streamlit_secrets_to_env() -> None:
    """Copy Cloud secrets into os.environ as early as possible in the run."""
    get_openai_api_key()


def _hide_default_sidebar() -> None:
    """No sidebar nav: this app uses a linear welcome → quiz → results + study guide flow."""
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Excel Readiness AI Coach",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    _apply_streamlit_secrets_to_env()
    initialize_session_state()
    _hide_default_sidebar()

    try:
        data = load_workbook_data(WORKBOOK_PATH)
        questions = parse_questions(data["Question Bank"])
    except FileNotFoundError:
        st.error(
            "Workbook not found at `./Excel_Readiness_AI_Coach_Content_Pack_with_sources.xlsx`.\n\n"
            "Please place the file in the project root and try again."
        )
        st.stop()
    except Exception as exc:
        st.error(f"Could not load workbook data: {exc}")
        st.stop()

    app_step = st.session_state.get("app_step", "welcome")
    if app_step == "post_quiz" and (
        not st.session_state.get("submitted") or st.session_state.get("results") is None
    ):
        st.session_state["app_step"] = "quiz"
        app_step = "quiz"
    if app_step == "welcome":
        page_welcome()
    elif app_step == "quiz":
        page_diagnostic_quiz(questions)
    elif app_step == "post_quiz":
        page_post_quiz(
            data["AI Prompt Template"],
            data["Material Library"],
        )
    else:
        st.session_state["app_step"] = "welcome"
        st.rerun()


if __name__ == "__main__":
    main()
