# foundry_to_docx.py
import json
import os
import glob
import re
import subprocess
import tempfile
from datetime import datetime, timezone

# ========= Colorama (Windows-friendly console colors) =========
try:
    import colorama
    from colorama import Fore, Style
    # autoreset so we don't need to manually append reset codes
    colorama.init(autoreset=True, convert=True)
except Exception:
    # fallback no coloring (define minimal placeholders)
    class _NoColor:
        RESET_ALL = ""
    colorama = None
    Fore = type("F", (), {"BLUE": "", "GREEN": "", "RED": ""})
    Style = _NoColor()

# ========= Third-party libraries (docx, bs4) =========
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup

# ========= Configuration =========
INPUT_DIR = "./sessions"
CONFIG_DIR = "./config"
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.txt")
ACTORS_FILE = os.path.join(CONFIG_DIR, "actors.txt")
EXPORT_DIR = "./export"
OMITTED_DIR = os.path.join(EXPORT_DIR, "omitted")
PORTRAITS_DIR = "portraits"  # folder with JPG portraits
PORTRAIT_WIDTH_INCH = 0.75  # portrait image width
LEFT_CELL_WIDTH_INCH = 1.5  # portrait + gap cell width (0.75 portrait + 0.75 gap)
PAGE_MARGIN_CM = 2.5
# =================================

CONFIG = {
    "TITLE": "FoundryVTT Session Transcript",
    "DEFAULT_SPEAKER": "Handler",
    "PRINT2PDF": "YES",  # default
}

ACTORS = {}  # speaker -> username mapping
DEFAULT_FONT = "Times New Roman"

# Global storage for duplicates found (grouped by session index and title)
DELETED_DUPLICATES = []  # list of (session_index, session_title, [(speaker, message), ...])
SESSION_DATES = []  # list of parsed session dates (None if missing) in order

# Console icon
ICON = "●"

# ---- Logging helpers (colored; safe on Windows thanks to colorama) ----
def log_loading(text):
    try:
        print(f"{Fore.BLUE}{ICON} {text}{Style.RESET_ALL}")
    except Exception:
        print(f"{ICON} {text}")

def log_done(text):
    try:
        print(f"{Fore.GREEN}{ICON} {text}{Style.RESET_ALL}")
    except Exception:
        print(f"{ICON} {text}")

def log_fail(text):
    try:
        print(f"{Fore.RED}{ICON} {text}{Style.RESET_ALL}")
    except Exception:
        print(f"{ICON} {text}")

# ---------------- Config loaders ----------------
def load_config():
    if not os.path.exists(CONFIG_FILE):
        log_loading(f"No {CONFIG_FILE} found — using defaults...")
        return
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, value = line.split("=", 1)
                key = key.strip().upper()
                value = value.strip()
                CONFIG[key] = value
    log_done(f"Loaded configuration from {CONFIG_FILE}")

def load_actors():
    if not os.path.exists(ACTORS_FILE):
        log_loading(f"No {ACTORS_FILE} found — skipping Cast list...")
        return
    with open(ACTORS_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or "=" not in line:
                continue
            speaker, username = line.split("=", 1)
            ACTORS[speaker.strip()] = username.strip()
    log_done(f"Loaded {len(ACTORS)} actors from {ACTORS_FILE}")

# ---------------- Utilities ----------------
def clean_html(content):
    """Convert HTML to plain text without inserting extra newline characters."""
    soup = BeautifulSoup(content, "html.parser")
    text = soup.get_text(separator="")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def parse_iso_or_epoch(value):
    """Try to parse ISO datetime string or epoch milliseconds. Return datetime or None."""
    if value is None:
        return None
    try:
        if isinstance(value, (int, float)) or (isinstance(value, str) and value.isdigit()):
            v = int(value)
            if v > 1e12:
                return datetime.fromtimestamp(v / 1000.0, tz=timezone.utc)
            else:
                return datetime.fromtimestamp(v, tz=timezone.utc)
    except Exception:
        pass
    if isinstance(value, str):
        s = value.strip()
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        try:
            return datetime.fromisoformat(s)
        except Exception:
            pass
    return None

def get_session_date(data):
    """
    Heuristic search for a session date in the JSON data.
    Returns formatted string like 'March 2, 2025' or None.
    """
    def fmt(dt):
        if not dt:
            return None
        return f"{dt.strftime('%B')} {dt.day}, {dt.year}"

    if not isinstance(data, dict):
        return None

    candidates = []
    d_block = data.get("data", {}) if isinstance(data.get("data", {}), dict) else {}

    for key in ("created", "createdTime", "modified", "modifiedTime", "timestamp"):
        if key in d_block:
            candidates.append(d_block[key])
        if key in data:
            candidates.append(data[key])

    if isinstance(d_block.get("_stats"), dict):
        for key in ("createdTime", "modifiedTime"):
            if key in d_block["_stats"]:
                candidates.append(d_block["_stats"][key])

    for cand in candidates:
        dt = parse_iso_or_epoch(cand)
        if dt:
            return fmt(dt)

    return None

# ---------------- DOCX Helpers ----------------
def set_margins(section):
    """Set uniform margins on a section in centimeters."""
    section.top_margin = Cm(PAGE_MARGIN_CM)
    section.bottom_margin = Cm(PAGE_MARGIN_CM)
    section.left_margin = Cm(PAGE_MARGIN_CM)
    section.right_margin = Cm(PAGE_MARGIN_CM)

def paragraph_format_defaults(paragraph, space_before_pt=6, space_after_pt=6, line_spacing=1.5):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.line_spacing = line_spacing

def set_page_number_start(section, start=1):
    """Set w:pgNumType w:start on section."""
    sectPr = section._sectPr
    existing = sectPr.find(qn("w:pgNumType"))
    if existing is not None:
        sectPr.remove(existing)
    pgNumType = OxmlElement("w:pgNumType")
    pgNumType.set(qn("w:start"), str(start))
    sectPr.append(pgNumType)

def add_page_field_to_footer(section):
    """Add a centered PAGE field to the given section's footer."""
    footer = section.footer
    if not footer.paragraphs:
        p = footer.add_paragraph()
    else:
        p = footer.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.font.name = DEFAULT_FONT
    run.font.size = Pt(10)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)

def clear_footer(section):
    """Clear footer content and unlink from previous."""
    try:
        section.footer.is_linked_to_previous = False
    except Exception:
        pass
    footer = section.footer
    if not footer.paragraphs:
        return
    footer.paragraphs[0].clear()

# Table helpers
def set_table_fixed_layout(table):
    """Set table layout to fixed so column widths are honoured (works on all python-docx versions)."""
    tbl = table._tbl
    tblPr = getattr(tbl, "tblPr", None)
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.append(tblPr)
    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)
    tblLayout.set(qn("w:type"), "fixed")

def set_cell_width(cell, width_inches):
    """Set a table cell width (tcW in twips/dxa)."""
    tc = cell._tc
    tcPr = getattr(tc, "tcPr", None)
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.append(tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    twips = int(width_inches * 1440)
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(twips))
    tcW.set(qn("w:type"), "dxa")

# ---------------- Cast Section (table with portraits) ----------------
def add_cast_section(doc):
    """Add Cast: heading and actor rows as tables with portrait left, text right."""
    if not ACTORS:
        return

    h_para = doc.add_paragraph()
    h_run = h_para.add_run("Cast:")
    try:
        h_run.font.name = DEFAULT_FONT
        h_run.font.size = Pt(18)
    except Exception:
        pass
    h_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph_format_defaults(h_para)

    for speaker, username in ACTORS.items():
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        set_table_fixed_layout(table)

        cell_img = table.rows[0].cells[0]
        cell_text = table.rows[0].cells[1]

        set_cell_width(cell_img, LEFT_CELL_WIDTH_INCH)

        cell_img.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell_text.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        portrait_path = os.path.join(PORTRAITS_DIR, f"{username}.jpg")
        if os.path.exists(portrait_path):
            try:
                p_img = cell_img.paragraphs[0]
                run_img = p_img.add_run()
                run_img.add_picture(portrait_path, width=Inches(PORTRAIT_WIDTH_INCH))
            except Exception as e:
                log_fail(f"Could not insert portrait for {username}: {e}")

        p_text = cell_text.paragraphs[0]
        paragraph_format_defaults(p_text)
        p_text.alignment = WD_ALIGN_PARAGRAPH.LEFT

        r1 = p_text.add_run(f"{speaker} — ")
        r1.bold = True
        r1.font.name = DEFAULT_FONT
        r1.font.size = Pt(12)

        r2 = p_text.add_run(username)
        r2.font.name = DEFAULT_FONT
        r2.font.size = Pt(12)

    blank = doc.add_paragraph()
    paragraph_format_defaults(blank)

# ---------------- Message Formatting ----------------
def add_styled_paragraph(doc, content, style=0, speaker=None):
    """
    Add a message paragraph.
    - style==1: narration (italic)
    - style==0 or 2: dialogue/rolls (normal)
    Body text will be 12pt and fully justified.
    """
    if not speaker:
        speaker = CONFIG.get("DEFAULT_SPEAKER", "Handler")

    p = doc.add_paragraph()
    paragraph_format_defaults(p)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    run_speaker = p.add_run(f"{speaker}: ")
    run_speaker.bold = True
    run_speaker.font.name = DEFAULT_FONT
    run_speaker.font.size = Pt(12)

    keywords_pattern = re.compile(r"(Critical Success|Critical Failure|Success|Failure)", re.IGNORECASE)
    parts = keywords_pattern.split(content)

    for part in parts:
        if not part:
            continue
        if keywords_pattern.fullmatch(part):
            run = p.add_run(part)
            run.bold = True
            run.font.name = DEFAULT_FONT
            run.font.size = Pt(12)
        else:
            run = p.add_run(part)
            run.font.name = DEFAULT_FONT
            run.font.size = Pt(12)
            if style == 1:
                run.italic = True

# ---------------- Dice roll extraction ----------------
def extract_roll_info(msg):
    content = msg.get("content", "")
    flavor = msg.get("flavor", "")
    speaker_data = msg.get("speaker", {})
    speaker = speaker_data.get("alias") if speaker_data else None
    if not speaker:
        speaker = CONFIG.get("DEFAULT_SPEAKER", "Handler")

    if "dice-roll" in content:
        soup = BeautifulSoup(content, "html.parser")
        formula_el = soup.select_one(".dice-formula")
        total_el = soup.select_one(".dice-total")
        formula = formula_el.get_text(strip=True) if formula_el else "?"
        total = total_el.get_text(strip=True) if total_el else "?"
        flavor_text = re.sub(r"<[^>]+>", "", flavor).strip()
        result = f"{speaker} rolls {formula} -> {total}"
        if flavor_text:
            result += f" ({flavor_text})"
        return result
    return None

# ---------------- File processing (with consecutive duplicate removal) ----------------
def process_file(filepath, doc, session_index, is_first_session=False):
    """
    Add session header and messages.
    Duplicate removal: only remove immediate consecutive duplicates of the same (speaker_alias, cleaned_content).
    Populate global DELETED_DUPLICATES list with removed messages for this session.
    """
    if not is_first_session:
        new_sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        set_margins(new_sec)
        try:
            new_sec.header.is_linked_to_previous = True
            new_sec.footer.is_linked_to_previous = True
        except Exception:
            pass
        try:
            sectPr = new_sec._sectPr
            pg = sectPr.find(qn("w:pgNumType"))
            if pg is not None:
                sectPr.remove(pg)
        except Exception:
            pass

    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Session title (Heading 1 for bookmark generation by Word -> PDF)
    title = (data.get("data", {}) or {}).get("title") or data.get("title") or os.path.basename(filepath)
    header = doc.add_paragraph()
    header.style = doc.styles["Heading 1"]
    run = header.add_run(title)

    try:
        for run in header.runs:
            run.font.name = DEFAULT_FONT
            run.font.size = Pt(14)
            run.bold = True
            run.underline = False
            run.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format_defaults(header)

    # Session date extraction with fallback (centered, 12 pt)
    session_date = get_session_date(data)
    SESSION_DATES.append(session_date)  # may be None
    if not session_date:
        session_date = "FALLBACK DATE!"
    p_date = doc.add_paragraph()
    paragraph_format_defaults(p_date)
    run_date = p_date.add_run(session_date)
    run_date.font.name = DEFAULT_FONT
    run_date.font.size = Pt(12)
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # blank line after header/date
    blank = doc.add_paragraph()
    paragraph_format_defaults(blank)

    messages = data.get("messages", [])
    last_key = None
    removed_list = []
    removed_count = 0

    for msg in messages:
        raw_content = msg.get("content", "")
        if raw_content is None:
            last_key = None
            continue
        cleaned = clean_html(raw_content)
        speaker_data = msg.get("speaker", {}) or {}
        speaker_alias = speaker_data.get("alias") if speaker_data.get("alias") else CONFIG.get("DEFAULT_SPEAKER", "Handler")
        key = (speaker_alias.strip(), cleaned.strip())

        # Consecutive-only duplicate detection:
        if last_key is not None and key == last_key:
            removed_count += 1
            removed_list.append((speaker_alias, cleaned))
            continue

        if not cleaned:
            last_key = None
            continue

        roll_summary = extract_roll_info(msg)
        if roll_summary:
            add_styled_paragraph(doc, roll_summary, style=0, speaker=speaker_alias)
            last_key = (speaker_alias.strip(), roll_summary.strip())
            continue

        style = msg.get("style", 0)
        add_styled_paragraph(doc, cleaned, style=style, speaker=speaker_alias)
        last_key = key

    if removed_count:
        log_loading(f"Removed {removed_count} consecutive duplicate message(s) from {os.path.basename(filepath)}")
    DELETED_DUPLICATES.append((session_index, title, removed_list))

# ---------------- Deleted duplicates writer (clean title + filename) ----------------
def write_deleted_duplicates_doc(title, timestamp_str=None):
    """
    Write the deleted duplicates doc grouped by session.
    Filename and display: 'Deleted Duplicate Messages — <Title with spaces>.docx'
    (no timestamp; underscores replaced with spaces)
    """
    any_removed = any(len(lst) for (_, _, lst) in DELETED_DUPLICATES)
    if not any_removed:
        log_loading("No deleted duplicates to write.")
        return

    os.makedirs(OMITTED_DIR, exist_ok=True)
    doc = Document()
    first_section = doc.sections[0]
    set_margins(first_section)

    # title: replace underscores in provided title with spaces for display
    display_title = f"Deleted Duplicate Messages — {title.replace('_', ' ')}"
    title_para = doc.add_paragraph()
    run = title_para.add_run(display_title)
    try:
        run.font.name = DEFAULT_FONT
        run.font.size = Pt(18)
        run.bold = True
    except Exception:
        pass
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format_defaults(title_para, space_before_pt=6, space_after_pt=6, line_spacing=1.0)
    doc.add_paragraph()

    # Iterate sessions in stored order
    for (session_index, session_title, removed_list) in DELETED_DUPLICATES:
        if not removed_list:
            continue
        h = doc.add_paragraph()
        h_run = h.add_run(f"Session {session_index}: {session_title}")
        try:
            h_run.font.name = DEFAULT_FONT
            h_run.font.size = Pt(14)
            h_run.bold = True
        except Exception:
            pass
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph_format_defaults(h, space_before_pt=6, space_after_pt=6, line_spacing=1.0)

        # Session date
        date_idx = session_index - 1
        session_date = None
        if 0 <= date_idx < len(SESSION_DATES):
            session_date = SESSION_DATES[date_idx]
        if not session_date:
            session_date = "FALLBACK DATE!"
        dpara = doc.add_paragraph()
        drun = dpara.add_run(session_date)
        try:
            drun.font.name = DEFAULT_FONT
            drun.font.size = Pt(12)
            drun.bold = False
        except Exception:
            pass
        dpara.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph_format_defaults(dpara, space_before_pt=2, space_after_pt=6, line_spacing=1.0)

        # messages listed left aligned
        for speaker, message in removed_list:
            p = doc.add_paragraph()
            paragraph_format_defaults(p)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(f"{speaker}: {message}")
            r.font.name = DEFAULT_FONT
            r.font.size = Pt(12)

        doc.add_paragraph()  # blank line between sessions

    # Safe filename: keep the human-friendly display (spaces) but strip illegal filesystem chars
    safe_title = re.sub(r'[<>:"/\\\\|?*]', "", display_title).strip()
    outname = f"{safe_title}.docx"
    outpath = os.path.join(OMITTED_DIR, outname)
    doc.save(outpath)
    log_done(f"Deleted duplicates exported to: {outpath}")

# ---------------- PowerShell PDF export (external COM) ----------------
def export_docx_to_pdf_via_powershell(docx_path, pdf_path):
    """
    Use PowerShell to convert DOCX -> PDF using Word COM.
    Uses UTF-8 output encoding inside PowerShell and Resolve-Path to avoid truncation.
    """
    # Defensive: ensure docx_path and pdf_path are absolute
    docx_abs = os.path.abspath(docx_path)
    pdf_abs = os.path.abspath(pdf_path)

    # Escape single quotes for single-quoted PowerShell string by doubling them
    docx_escaped = docx_abs.replace("'", "''")
    pdf_escaped = pdf_abs.replace("'", "''")

    ps_script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$docPath = '{docx_escaped}'
$pdfPath = '{pdf_escaped}'
$docResolved = (Resolve-Path -LiteralPath $docPath).Path
$pdfResolved = (Resolve-Path -LiteralPath $pdfPath -ErrorAction SilentlyContinue)
if (-not $pdfResolved) {{
    $dir = Split-Path $pdfPath -Parent
    if (-not (Test-Path $dir)) {{
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }}
    $pdfResolved = $pdfPath
}} else {{
    $pdfResolved = $pdfResolved.Path
}}
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docResolved)
# Export with CreateBookmarks = 1 (Headings). 
# Parameters: OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor,
#             Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags
$doc.ExportAsFixedFormat($pdfResolved, 17, $false, 0, 0, 0, 0, 0, $true, $false, 1, $true)
$doc.Close($false)
$word.Quit()
"""

    tf_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ps1", mode="w", encoding="utf-8") as tf:
            tf.write(ps_script)
            tf_path = tf.name
        # Run PowerShell with bypass and no profile
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tf_path],
            capture_output=True,
            text=True,
        )
        if completed.returncode != 0:
            log_fail("PDF conversion failed. PowerShell output:")
            if completed.stdout:
                print(completed.stdout)
            if completed.stderr:
                print(completed.stderr)
            return False
        return True
    except Exception as e:
        log_fail(f"Exception while attempting PDF conversion: {e}")
        return False
    finally:
        try:
            if tf_path and os.path.exists(tf_path):
                os.remove(tf_path)
        except Exception:
            pass

# ---------------- Main ----------------
def main():
    # ensure config dir
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR, exist_ok=True)

    load_config()
    load_actors()

    print2pdf = CONFIG.get("PRINT2PDF", "YES").strip().upper()
    if print2pdf not in ("YES", "NO"):
        print2pdf = "YES"

    files = sorted(
        glob.glob(os.path.join(INPUT_DIR, "*.json")),
        key=lambda x: int(re.search(r"(\d+)", os.path.basename(x)).group(1))
        if re.search(r"(\d+)", os.path.basename(x))
        else 0,
    )

    if not files:
        log_fail(f"No JSON files found in {INPUT_DIR}")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    title_clean = re.sub(r"[^\w\s-]", "", CONFIG.get("TITLE", "FoundryVTT Session Transcript")).strip().replace(" ", "_")
    output_filename = f"{title_clean}_{timestamp}.docx"
    os.makedirs(EXPORT_DIR, exist_ok=True)
    output_path = os.path.join(EXPORT_DIR, output_filename)

    # Peek session dates to compute start/end for the title block
    peeked_dates = []
    for filepath in files:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
            sd = get_session_date(data)
            peeked_dates.append(sd)
        except Exception:
            peeked_dates.append(None)

    def fmt_date_or_fallback(d):
        return d if d else "FALLBACK DATE!"

    start_date = fmt_date_or_fallback(peeked_dates[0]) if peeked_dates else "FALLBACK DATE!"
    end_date = fmt_date_or_fallback(peeked_dates[-1]) if peeked_dates else "FALLBACK DATE!"

    # Create main document
    doc = Document()

    # Title (as Heading 1 so Word will create a PDF bookmark)
    # Use the Heading 1 style for the paragraph (ensures bookmark in PDF)
    title_para = doc.add_paragraph()
    try:
        title_para.style = doc.styles["Heading 1"]
    except Exception:
        # If style not available for some reason, just continue and set font
        pass
    title_run = title_para.add_run(CONFIG.get("TITLE", "FoundryVTT Session Transcript"))
    try:
        title_run.font.name = DEFAULT_FONT
        title_run.font.size = Pt(24)
        title_run.font.color.rgb = RGBColor(0, 0, 0)
        title_run.bold = True
    except Exception:
        pass
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format_defaults(title_para, space_before_pt=6, space_after_pt=6, line_spacing=1.0)

    # Sessions summary lines centered, 14pt not bold, 1.0 spacing
    sessions_line = doc.add_paragraph()
    sessions_run = sessions_line.add_run(f"Sessions 1 - {len(files)}")
    try:
        sessions_run.font.name = DEFAULT_FONT
        sessions_run.font.size = Pt(14)
        sessions_run.bold = False
    except Exception:
        pass
    sessions_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format_defaults(sessions_line, space_before_pt=0, space_after_pt=6, line_spacing=1.0)

    dates_line = doc.add_paragraph()
    dates_run = dates_line.add_run(f"{start_date} - {end_date}")
    try:
        dates_run.font.name = DEFAULT_FONT
        dates_run.font.size = Pt(14)
        dates_run.bold = False
    except Exception:
        pass
    dates_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph_format_defaults(dates_line, space_before_pt=0, space_after_pt=6, line_spacing=1.0)

    # one blank line before Cast
    doc.add_paragraph()

    # Cast section (first physical page)
    add_cast_section(doc)

    # Ensure first section: margins + ensure no numbering and unlink footer
    first_section = doc.sections[0]
    set_margins(first_section)
    try:
        first_section.different_first_page_header_footer = True
    except Exception:
        pass
    clear_footer(first_section)
    # Remove any pgNumType if present
    try:
        sectPr = first_section._sectPr
        existing_pg = sectPr.find(qn("w:pgNumType"))
        if existing_pg is not None:
            sectPr.remove(existing_pg)
    except Exception:
        pass

    # Numbered section (sessions begin here) - show page numbers starting at 1
    numbered_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    set_margins(numbered_section)
    try:
        numbered_section.different_first_page_header_footer = False
    except Exception:
        pass
    try:
        numbered_section.header.is_linked_to_previous = False
        numbered_section.footer.is_linked_to_previous = False
    except Exception:
        pass

    # Force numbering to start at 1 on this section and add page field
    set_page_number_start(numbered_section, 1)
    add_page_field_to_footer(numbered_section)

    # Process files: first session placed into this numbered section
    DELETED_DUPLICATES.clear()
    SESSION_DATES.clear()
    for idx, filepath in enumerate(files, start=1):
        is_first = idx == 1
        log_loading(f"Processing {os.path.basename(filepath)}...")
        process_file(filepath, doc, session_index=idx, is_first_session=is_first)

    # After processing all sessions, write the deleted duplicates doc
    # Pass title_clean (which is underscored) so write_deleted_duplicates_doc will replace underscores with spaces
    write_deleted_duplicates_doc(title_clean)

    # Save main doc
    doc.save(output_path)
    log_done(f"Export complete: {output_path}")

    # PDF conversion
    if print2pdf == "YES":
        log_loading("Please wait — converting to PDF...")
        try:
            pdf_path = os.path.splitext(output_path)[0] + ".pdf"
            ok = export_docx_to_pdf_via_powershell(output_path, pdf_path)
            if ok:
                log_done(f"PDF created: {pdf_path}")
            else:
                log_fail("PDF creation failed.")
        except Exception as e:
            log_fail(f"Exception during PDF creation: {e}")
    else:
        log_fail("PRINT2PDF=NO → Skipping PDF export")

if __name__ == "__main__":
    main()
