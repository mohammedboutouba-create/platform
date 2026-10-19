import io
import os
import re
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date
from urllib.parse import urlparse
from xml.sax.saxutils import escape

import altair as alt
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

# ==========================================
# 0. الثوابت والتسميات والكليات المعتمدة
# ==========================================
DB_PATH = "university_integrated_platform_v15.db"

FACULTIES = [
    "كلية العلوم والتكنولوجيا",
    "كلية العلوم الاقتصادية",
    "كلية الآداب واللغات والعلوم الاجتماعية",
    "كلية الحقوق",
]

ROLE_LABELS = {
    "INCUBATOR": "مدير حاضنة",
    "VICE_DEAN": "نائب عميد",
}

RESOURCE_CATEGORIES = [
    "قرارات وزارية وبيداغوجية",
    "نماذج مخطط الأعمال (BMC)",
    "أدلة براءات الاختراع",
]
ALLOWED_DOC_EXT = ["pdf", "xlsx", "xls"]
MAX_DOC_MB = 10
DOC_MIME = {
    "pdf": "application/pdf",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
}

SHEET_COLUMNS = [
    "الرقم",
    "الاسم واللقب",
    "القسم",
    "المستوى",
    "اسم المشروع",
    "ملخص عن المشروع",
]


# ==========================================
# 1. أدوات معالجة ملفات Excel (.xlsx) و PDF (.pdf)
# ==========================================
def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "المشاريع") -> bytes:
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            ws = writer.sheets[sheet_name[:31]]
            ws.sheet_view.rightToLeft = True
        return buf.getvalue()
    except Exception:
        buf = io.BytesIO()
        headers = list(df.columns)
        rows_xml = []

        cells_h = [
            f'<c r="{chr(65 + c_idx)}1" t="inlineStr"><is><t>{escape(str(h))}</t></is></c>'
            for c_idx, h in enumerate(headers)
        ]
        rows_xml.append(f'<row r="1">{"".join(cells_h)}</row>')

        for r_idx, row in enumerate(df.fillna("").to_dict("records"), start=2):
            cells = []
            for c_idx, h in enumerate(headers):
                col_letter = chr(65 + c_idx)
                val = row.get(h, "")
                if h == "الرقم" and str(val).strip().isdigit():
                    cells.append(f'<c r="{col_letter}{r_idx}"><v>{int(val)}</v></c>')
                else:
                    cells.append(
                        f'<c r="{col_letter}{r_idx}" t="inlineStr"><is><t>{escape(str(val))}</t></is></c>'
                    )
            rows_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

        sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>
            <sheetData>{"".join(rows_xml)}</sheetData>
        </worksheet>'''

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(
                "[Content_Types].xml",
                '''<?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                    <Default Extension="xml" ContentType="application/xml"/>
                    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>''',
            )
            zf.writestr(
                "_rels/.rels",
                '''<?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>''',
            )
            zf.writestr(
                "xl/workbook.xml",
                f'''<?xml version="1.0" encoding="UTF-8"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <sheets><sheet name="{escape(sheet_name[:31])}" sheetId="1" r:id="rId1"/></sheets>
                </workbook>''',
            )
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                '''<?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>''',
            )
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        return buf.getvalue()


def read_excel_to_df(uploaded_file) -> pd.DataFrame:
    raw_bytes = uploaded_file.getvalue()
    raw_df = None
    try:
        raw_df = pd.read_excel(io.BytesIO(raw_bytes))
    except Exception:
        with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
            shared_strings = []
            if "xl/sharedStrings.xml" in zf.namelist():
                ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in ss_root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                    texts = [
                        t.text or ""
                        for t in si.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                    ]
                    shared_strings.append("".join(texts))

            sheet_root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            extracted_rows = []
            for row_el in sheet_root.findall(".//m:sheetData/m:row", ns):
                row_vals = {}
                for c_el in row_el.findall("m:c", ns):
                    ref = c_el.attrib.get("r", "A1")
                    col_letters = "".join(ch for ch in ref if ch.isalpha())
                    col_idx = 0
                    for ch in col_letters:
                        col_idx = col_idx * 26 + (ord(ch.upper()) - 64)
                    col_idx -= 1

                    c_type = c_el.attrib.get("t", "")
                    v_el = c_el.find("m:v", ns)
                    is_el = c_el.find(".//m:t", ns)
                    val = ""
                    if c_type == "s" and v_el is not None and v_el.text:
                        idx = int(v_el.text)
                        val = shared_strings[idx] if idx < len(shared_strings) else ""
                    elif c_type == "inlineStr" and is_el is not None:
                        val = is_el.text or ""
                    elif v_el is not None and v_el.text is not None:
                        val = v_el.text
                    row_vals[col_idx] = val
                if row_vals:
                    max_c = max(row_vals.keys())
                    extracted_rows.append([row_vals.get(i, "") for i in range(max_c + 1)])

            if extracted_rows:
                headers = [str(x).strip() for x in extracted_rows[0]]
                data_rows = extracted_rows[1:]
                raw_df = pd.DataFrame(data_rows, columns=headers)
            else:
                raw_df = pd.DataFrame(columns=SHEET_COLUMNS)

    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=SHEET_COLUMNS)

    raw_df.columns = [str(c).strip() for c in raw_df.columns]
    result = pd.DataFrame()

    for idx, target_col in enumerate(SHEET_COLUMNS):
        if target_col in raw_df.columns:
            result[target_col] = raw_df[target_col]
        elif idx < len(raw_df.columns):
            result[target_col] = raw_df.iloc[:, idx]
        else:
            result[target_col] = ""

    result = result.dropna(how="all").fillna("")
    cleaned_nums = []
    for i, val in enumerate(result["الرقم"], start=1):
        try:
            cleaned_nums.append(int(float(str(val).strip())))
        except Exception:
            cleaned_nums.append(i)
    result["الرقم"] = cleaned_nums
    return result[SHEET_COLUMNS]


# --- تشكيل الحروف العربية لتصدير PDF ---
_AR_MAP = {
    "\u0621": ("\uFE80", "\uFE80", "\uFE80", "\uFE80"),
    "\u0622": ("\uFE81", "\uFE82", "\uFE81", "\uFE82"),
    "\u0623": ("\uFE83", "\uFE84", "\uFE83", "\uFE84"),
    "\u0624": ("\uFE85", "\uFE86", "\uFE85", "\uFE86"),
    "\u0625": ("\uFE87", "\uFE88", "\uFE87", "\uFE88"),
    "\u0626": ("\uFE89", "\uFE8A", "\uFE8B", "\uFE8C"),
    "\u0627": ("\uFE8D", "\uFE8E", "\uFE8D", "\uFE8E"),
    "\u0628": ("\uFE8F", "\uFE90", "\uFE91", "\uFE92"),
    "\u0629": ("\uFE93", "\uFE94", "\uFE93", "\uFE94"),
    "\u062A": ("\uFE95", "\uFE96", "\uFE97", "\uFE98"),
    "\u062B": ("\uFE99", "\uFE9A", "\uFE9B", "\uFE9C"),
    "\u062C": ("\uFE9D", "\uFE9E", "\uFE9F", "\uFEA0"),
    "\u062D": ("\uFEA1", "\uFEA2", "\uFEA3", "\uFEA4"),
    "\u062E": ("\uFEA5", "\uFEA6", "\uFEA7", "\uFEA8"),
    "\u062F": ("\uFEA9", "\uFEAA", "\uFEA9", "\uFEAA"),
    "\u0630": ("\uFEAB", "\uFEAC", "\uFEAB", "\uFEAC"),
    "\u0631": ("\uFEAD", "\uFEAE", "\uFEAD", "\uFEAE"),
    "\u0632": ("\uFEAF", "\uFEB0", "\uFEAF", "\uFEB0"),
    "\u0633": ("\uFEB1", "\uFEB2", "\uFEB3", "\uFEB4"),
    "\u0634": ("\uFEB5", "\uFEB6", "\uFEB7", "\uFEB8"),
    "\u0635": ("\uFEB9", "\uFEBA", "\uFEBB", "\uFEBC"),
    "\u0636": ("\uFEBD", "\uFEBE", "\uFEBF", "\uFEC0"),
    "\u0637": ("\uFEC1", "\uFEC2", "\uFEC3", "\uFEC4"),
    "\u0638": ("\uFEC5", "\uFEC6", "\uFEC7", "\uFEC8"),
    "\u0639": ("\uFEC9", "\uFECA", "\uFECB", "\uFECC"),
    "\u063A": ("\uFECD", "\uFECE", "\uFECF", "\uFED0"),
    "\u0641": ("\uFED1", "\uFED2", "\uFED3", "\uFED4"),
    "\u0642": ("\uFED5", "\uFED6", "\uFED7", "\uFED8"),
    "\u0643": ("\uFED9", "\uFEDA", "\uFEDB", "\uFEDC"),
    "\u0644": ("\uFEDD", "\uFEDE", "\uFEDF", "\uFEE0"),
    "\u0645": ("\uFEE1", "\uFEE2", "\uFEE3", "\uFEE4"),
    "\u0646": ("\uFEE5", "\uFEE6", "\uFEE7", "\uFEE8"),
    "\u0647": ("\uFEE9", "\uFEEA", "\uFEEB", "\uFEEC"),
    "\u0648": ("\uFEED", "\uFEEE", "\uFEED", "\uFEEE"),
    "\u0649": ("\uFEEF", "\uFEF0", "\uFEEF", "\uFEF0"),
    "\u064A": ("\uFEF1", "\uFEF2", "\uFEF3", "\uFEF4"),
}
_DUAL_JOIN = set("بتثجحخسشصضطظعغفقكلمنهيئ")
_RIGHT_JOIN = _DUAL_JOIN | set("اأإآدذرزوؤةى")
_LAM_ALEF = {
    "\u0622": ("\uFEF5", "\uFEF6"),
    "\u0623": ("\uFEF7", "\uFEF8"),
    "\u0625": ("\uFEF9", "\uFEFA"),
    "\u0627": ("\uFEFB", "\uFEFC"),
}


def _shape_arabic_word(text: str) -> str:
    chars = [c for c in text if not ("\u064B" <= c <= "\u065F")]
    out = []
    i = 0
    n = len(chars)
    while i < n:
        ch = chars[i]
        prev_c = chars[i - 1] if i > 0 else ""
        prev_joins = prev_c in _DUAL_JOIN

        if ch == "\u0644" and i + 1 < n and chars[i + 1] in _LAM_ALEF:
            lig = _LAM_ALEF[chars[i + 1]]
            out.append(lig[1] if prev_joins else lig[0])
            i += 2
            continue

        if ch in _AR_MAP:
            next_c = chars[i + 1] if i + 1 < n else ""
            next_joins = (ch in _DUAL_JOIN) and (next_c in _RIGHT_JOIN)
            forms = _AR_MAP[ch]
            if prev_joins and next_joins:
                out.append(forms[3])
            elif prev_joins:
                out.append(forms[1])
            elif next_joins:
                out.append(forms[2])
            else:
                out.append(forms[0])
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def _rtl_visual(text: str) -> str:
    if not text:
        return ""
    shaped = _shape_arabic_word(str(text))
    tokens = re.findall(r"[\u0600-\u06FF\uFE70-\uFEFC]+|[A-Za-z0-9_./:-]+|\s+|[^\s\w]", shaped)
    visual_tokens = []
    for tok in reversed(tokens):
        if re.match(r"^[\u0600-\u06FF\uFE70-\uFEFC]+$", tok):
            visual_tokens.append(tok[::-1])
        else:
            visual_tokens.append(tok)
    return "".join(visual_tokens)


def _load_font(size: int, bold: bool = False):
    candidates = (
        ["C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/tahomabd.ttf"]
        if bold
        else ["C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/tahoma.ttf"]
    )
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def df_to_pdf_bytes(df: pd.DataFrame, faculty_name: str) -> bytes:
    page_w, margin = 1650, 50
    col_widths = [90, 250, 240, 150, 310, 510]
    headers_rtl = list(reversed(SHEET_COLUMNS))
    widths_rtl = list(reversed(col_widths))

    font_title = _load_font(26, bold=True)
    font_sub = _load_font(17, bold=False)
    font_head = _load_font(18, bold=True)
    font_cell = _load_font(16, bold=False)

    clean_rows = df.fillna("").to_dict("records")
    row_heights = []
    wrapped_rows = []

    for r in clean_rows:
        row_cells = []
        max_lines = 1
        for col_name, w in zip(headers_rtl, widths_rtl):
            raw_val = str(r.get(col_name, "")).strip()
            max_chars = max(10, int(w / 10))
            words = raw_val.split()
            lines, cur = [], ""
            for word in words:
                if len(cur) + len(word) + 1 <= max_chars:
                    cur = f"{cur} {word}".strip()
                else:
                    if cur:
                        lines.append(cur)
                    cur = word
            if cur:
                lines.append(cur)
            if not lines:
                lines = [""]
            max_lines = max(max_lines, len(lines))
            row_cells.append(lines)
        wrapped_rows.append(row_cells)
        row_heights.append(max(44, max_lines * 26 + 16))

    total_h = max(750, 210 + sum(row_heights) + 80)
    img = Image.new("RGB", (page_w, total_h), "white")
    draw = ImageDraw.Draw(img)

    draw.rectangle([margin, 35, page_w - margin, 125], fill=(24, 62, 115))
    title_txt = _rtl_visual(f"قائمة المشاريع الابتكارية - {faculty_name}")
    sub_txt = _rtl_visual(f"تاريخ الاستخراج: {now_str()} | مدير حاضنة")
    draw.text((page_w - margin - 25, 50), title_txt, font=font_title, fill="white", anchor="ra")
    draw.text((page_w - margin - 25, 90), sub_txt, font=font_sub, fill=(220, 232, 245), anchor="ra")

    y = 145
    header_h = 48
    x = margin
    for h_name, w in zip(headers_rtl, widths_rtl):
        draw.rectangle([x, y, x + w, y + header_h], fill=(230, 238, 248), outline=(120, 140, 170), width=2)
        draw.text((x + w - 12, y + 12), _rtl_visual(h_name), font=font_head, fill=(20, 40, 80), anchor="ra")
        x += w
    y += header_h

    for r_idx, (row_cells, r_h) in enumerate(zip(wrapped_rows, row_heights)):
        bg = (255, 255, 255) if r_idx % 2 == 0 else (247, 250, 253)
        x = margin
        for lines, w in zip(row_cells, widths_rtl):
            draw.rectangle([x, y, x + w, y + r_h], fill=bg, outline=(180, 195, 210), width=1)
            for l_idx, line in enumerate(lines):
                draw.text(
                    (x + w - 12, y + 10 + l_idx * 24),
                    _rtl_visual(line),
                    font=font_cell,
                    fill=(30, 30, 30),
                    anchor="ra",
                )
            x += w
        y += r_h

    buf = io.BytesIO()
    img.save(buf, format="PDF", resolution=150.0)
    return buf.getvalue()


# ==========================================
# 2. قاعدة البيانات
# ==========================================
@st.cache_resource
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


conn = get_conn()


def q(sql, params=(), one=False):
    cur = conn.execute(sql, params)
    return cur.fetchone() if one else cur.fetchall()


def ex(sql, params=()):
    cur = conn.execute(sql, params)
    conn.commit()
    return cur.lastrowid


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def init_db():
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT NOT NULL,
            faculty TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS academic_seasons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_year TEXT NOT NULL UNIQUE,
            is_active INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS faculty_sheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty TEXT NOT NULL,
            season_year TEXT NOT NULL,
            row_num INTEGER NOT NULL,
            full_name TEXT NOT NULL,
            department TEXT NOT NULL,
            level TEXT NOT NULL,
            project_name TEXT NOT NULL,
            project_summary TEXT NOT NULL,
            updated_by TEXT,
            updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS resources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            category TEXT NOT NULL,
            link TEXT NOT NULL,
            uploaded_by INTEGER NOT NULL REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            message TEXT NOT NULL,
            created_at TEXT NOT NULL,
            is_read INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            filename TEXT NOT NULL,
            file_ext TEXT NOT NULL,
            file_data BLOB NOT NULL,
            target_faculty TEXT NOT NULL,
            uploaded_by INTEGER NOT NULL REFERENCES users(id),
            created_at TEXT NOT NULL
        );
        """
    )

    if q("SELECT COUNT(*) AS n FROM users", one=True)["n"] == 0:
        users = [
            ("مدير حاضنة", "INCUBATOR", "إدارة الحاضنة"),
            ("نائب عميد كلية العلوم والتكنولوجيا", "VICE_DEAN", "كلية العلوم والتكنولوجيا"),
            ("نائب عميد كلية العلوم الاقتصادية", "VICE_DEAN", "كلية العلوم الاقتصادية"),
            ("نائب عميد كلية الآداب واللغات والعلوم الاجتماعية", "VICE_DEAN", "كلية الآداب واللغات والعلوم الاجتماعية"),
            ("نائب عميد كلية الحقوق", "VICE_DEAN", "كلية الحقوق"),
        ]
        conn.executemany("INSERT INTO users (name, role, faculty) VALUES (?, ?, ?)", users)
        conn.commit()

    if q("SELECT COUNT(*) AS n FROM academic_seasons", one=True)["n"] == 0:
        conn.executemany(
            "INSERT INTO academic_seasons (season_year, is_active) VALUES (?, ?)",
            [("2024/2025", 0), ("2025/2026", 1)],
        )
        conn.commit()


init_db()


# ==========================================
# 3. دوال مساعدة
# ==========================================
def get_active_season() -> str:
    row = q("SELECT season_year FROM academic_seasons WHERE is_active=1 LIMIT 1", one=True)
    return row["season_year"] if row else "2025/2026"


def valid_url(url: str) -> bool:
    try:
        p = urlparse(url.strip())
        return p.scheme in ("http", "https") and bool(p.netloc)
    except Exception:
        return False


def flash(msg: str, icon: str = "✅"):
    st.session_state["_flash"] = (msg, icon)


def notify(user_ids, message: str):
    for uid in set(user_ids):
        ex(
            "INSERT INTO notifications (user_id, message, created_at) VALUES (?,?,?)",
            (uid, message, now_str()),
        )


def users_by_filter(role=None, faculty=None):
    sql = "SELECT id FROM users WHERE 1=1"
    params = []
    if role:
        sql += " AND role=?"
        params.append(role)
    if faculty:
        sql += " AND faculty=?"
        params.append(faculty)
    return [r["id"] for r in q(sql, tuple(params))]


def load_faculty_df(faculty_name: str, season: str) -> pd.DataFrame:
    rows = q(
        """SELECT row_num, full_name, department, level, project_name, project_summary
           FROM faculty_sheets WHERE faculty=? AND season_year=? ORDER BY row_num ASC, id ASC""",
        (faculty_name, season),
    )
    if not rows:
        return pd.DataFrame(columns=SHEET_COLUMNS)
    data = [
        {
            "الرقم": r["row_num"],
            "الاسم واللقب": r["full_name"],
            "القسم": r["department"],
            "المستوى": r["level"],
            "اسم المشروع": r["project_name"],
            "ملخص عن المشروع": r["project_summary"],
        }
        for r in rows
    ]
    return pd.DataFrame(data, columns=SHEET_COLUMNS)


def save_faculty_df(faculty_name: str, season: str, df: pd.DataFrame, user_name: str):
    ex("DELETE FROM faculty_sheets WHERE faculty=? AND season_year=?", (faculty_name, season))
    clean_df = df.dropna(how="all").fillna("")
    for idx, row in enumerate(clean_df.to_dict("records"), start=1):
        r_num = row.get("الرقم")
        try:
            r_num = int(float(str(r_num).strip())) if str(r_num).strip() != "" else idx
        except ValueError:
            r_num = idx

        ex(
            """INSERT INTO faculty_sheets
               (faculty, season_year, row_num, full_name, department, level, project_name, project_summary, updated_by, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (
                faculty_name,
                season,
                r_num,
                str(row.get("الاسم واللقب", "")).strip(),
                str(row.get("القسم", "")).strip(),
                str(row.get("المستوى", "")).strip(),
                str(row.get("اسم المشروع", "")).strip(),
                str(row.get("ملخص عن المشروع", "")).strip(),
                user_name,
                now_str(),
            ),
        )


def get_faculty_project_count(faculty_name: str, season: str) -> int:
    sheet_count = q(
        "SELECT COUNT(*) AS n FROM faculty_sheets WHERE faculty=? AND season_year=? AND TRIM(project_name) != ''",
        (faculty_name, season),
        one=True,
    )["n"]
    return sheet_count


# ==========================================
# 4. دالة بناء فضاء الكلية
# ==========================================
def render_faculty_hub(faculty_name: str, tab_idx: int, active_season: str):
    st.markdown(f"""
        <div class="p-4 mb-4 bg-blue-50 border-r-4 border-blue-600 rounded-lg shadow-sm">
            <h3 class="text-lg font-bold text-blue-900">🏛️ فضاء مشاريع {faculty_name} (الموسم: {active_season})</h3>
            <p class="text-sm text-gray-600">إدارة الجداول والمشاريع الابتكارية الخاصة بالكلية.</p>
        </div>
    """, unsafe_allow_html=True)

    role = current_user["role"]
    is_incubator = role == "INCUBATOR"
    can_manage_faculty = is_incubator

    if not is_incubator:
        st.info("ℹ️ أنت مسجل بصلاحية نائب العميد (قراءة واطلاع وتصدير فقط). التعديل والإضافة متاحان حصرياً لمدير الحاضنة.")

    if can_manage_faculty:
        with st.expander(f"📥 استيراد ملف Excel وتعديل جدول {faculty_name}", expanded=False):
            col_imp1, col_imp2 = st.columns([2, 1])
            with col_imp1:
                imp_file = st.file_uploader(
                    "اختر ملف إكسل (.xlsx / .xls)",
                    type=["xlsx", "xls"],
                    key=f"imp_excel_{tab_idx}",
                )
            with col_imp2:
                imp_mode = st.radio(
                    "طريقة الاستيراد:",
                    ["استبدال بيانات الجدول الحالي", "إضافة الأسطر إلى الجدول الحالي"],
                    key=f"imp_mode_{tab_idx}",
                )
                if st.button("⬆️ تعبئة الجدول من ملف Excel", key=f"btn_imp_{tab_idx}", use_container_width=True):
                    if imp_file is None:
                        st.warning("يرجى اختيار ملف أولاً.")
                    else:
                        try:
                            imported_df = read_excel_to_df(imp_file)
                            if imported_df.empty:
                                st.warning("الملف فارغ.")
                            else:
                                if imp_mode == "إضافة الأسطر إلى الجدول الحالي":
                                    existing_df = load_faculty_df(faculty_name, active_season)
                                    combined_df = pd.concat([existing_df, imported_df], ignore_index=True)
                                    combined_df["الرقم"] = list(range(1, len(combined_df) + 1))
                                    save_faculty_df(faculty_name, active_season, combined_df, current_user["name"])
                                else:
                                    save_faculty_df(faculty_name, active_season, imported_df, current_user["name"])

                                st.session_state[f"grid_ver_{tab_idx}"] = (
                                    st.session_state.get(f"grid_ver_{tab_idx}", 0) + 1
                                )
                                flash(f"تم استيراد {len(imported_df)} سطراً بنجاح!")
                                st.rerun()
                        except Exception as e:
                            st.error(f"خطأ في قراءة الملف: {e}")

    last_meta = q(
        "SELECT updated_by, updated_at FROM faculty_sheets WHERE faculty=? AND season_year=? ORDER BY id DESC LIMIT 1",
        (faculty_name, active_season),
        one=True,
    )
    if last_meta and last_meta["updated_by"]:
        st.caption(f"🕒 آخر تحديث بواسطة: **{last_meta['updated_by']}** بتاريخ: `{last_meta['updated_at']}`")

    df_current = load_faculty_df(faculty_name, active_season)
    grid_ver = st.session_state.get(f"grid_ver_{tab_idx}", 0)

    edited_df = st.data_editor(
        df_current,
        num_rows="dynamic" if can_manage_faculty else "fixed",
        disabled=not can_manage_faculty,
        use_container_width=True,
        hide_index=True,
        key=f"excel_grid_{tab_idx}_{active_season}_{grid_ver}",
        column_config={
            "الرقم": st.column_config.NumberColumn("1. الرقم", min_value=1, step=1, width="small"),
            "الاسم واللقب": st.column_config.TextColumn("2. الاسم واللقب", width="medium"),
            "القسم": st.column_config.TextColumn("3. القسم", width="medium"),
            "المستوى": st.column_config.SelectboxColumn(
                "4. المستوى",
                options=["ليسانس L3", "ماستر 1", "ماستر 2", "دكتوراه"],
                width="small",
            ),
            "اسم المشروع": st.column_config.TextColumn("5. اسم المشروع", width="medium"),
            "ملخص عن المشروع": st.column_config.TextColumn("6. ملخص عن المشروع", width="large"),
        },
    )

    # تم حذف زر مشاركة الجدول تماماً بناءً على طلبك
    col_b1, col_b2, col_b3 = st.columns([1.5, 1.5, 1.5])

    if can_manage_faculty:
        with col_b1:
            if st.button("💾 حفظ الجدول", key=f"save_sheet_{tab_idx}", use_container_width=True):
                save_faculty_df(faculty_name, active_season, edited_df, current_user["name"])
                flash("تم حفظ التعديلات.")
                st.rerun()

    with col_b2:
        xlsx_export = df_to_xlsx_bytes(edited_df, sheet_name=faculty_name)
        st.download_button(
            "📗 تحميل بصيغة Excel",
            data=xlsx_export,
            file_name=f"جدول_مشاريع_{faculty_name}.xlsx",
            mime=DOC_MIME["xlsx"],
            key=f"dl_xlsx_{tab_idx}",
            use_container_width=True,
        )

    with col_b3:
        pdf_export = df_to_pdf_bytes(edited_df, faculty_name)
        st.download_button(
            "📕 تحميل بصيغة PDF",
            data=pdf_export,
            file_name=f"جدول_مشاريع_{faculty_name}.pdf",
            mime=DOC_MIME["pdf"],
            key=f"dl_pdf_{tab_idx}",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown(f"### 📎 الوثائق المتبادلة الخاصة بـ {faculty_name}")

    if can_manage_faculty:
        with st.expander(f"⬆️ إرسال ملف إضافي (PDF / Excel)", expanded=False):
            with st.form(f"doc_form_{tab_idx}", clear_on_submit=True):
                d_file = st.file_uploader(f"اختر ملف (حد أقصى {MAX_DOC_MB} ميغابايت)", type=ALLOWED_DOC_EXT, key=f"doc_file_{tab_idx}")
                d_title = st.text_input("عنوان الوثيقة", key=f"doc_title_{tab_idx}")

                if st.form_submit_button("إرسال وحفظ الملف"):
                    if d_file is None:
                        st.warning("يرجى اختيار ملف.")
                    else:
                        ext = d_file.name.rsplit(".", 1)[-1].lower()
                        data = d_file.getvalue()
                        if ext not in ALLOWED_DOC_EXT:
                            st.warning("نوع الملف غير مسموح.")
                        elif len(data) > MAX_DOC_MB * 1024 * 1024:
                            st.warning("حجم الملف يتجاوز الحد الأقصى.")
                        else:
                            doc_title = d_title.strip() or d_file.name
                            ex(
                                """INSERT INTO documents (title, filename, file_ext, file_data,
                                                          target_faculty, uploaded_by, created_at)
                                   VALUES (?,?,?,?,?,?)""",
                                (doc_title, d_file.name, ext, data, faculty_name, current_user["id"], now_str()),
                            )
                            notify(users_by_filter(role="VICE_DEAN", faculty=faculty_name), f"وثيقة جديدة موجهة لـ {faculty_name}: {doc_title}")
                            flash("تم إرسال الملف بنجاح.")
                            st.rerun()

    docs = q(
        """SELECT d.id, d.title, d.filename, d.file_ext, d.file_data, d.target_faculty,
                  d.uploaded_by, d.created_at, u.name AS uploader
           FROM documents d
           JOIN users u ON u.id = d.uploaded_by
           WHERE d.target_faculty = ?
           ORDER BY d.id DESC""",
        (faculty_name,),
    )
    if not docs:
        st.caption("لا توجد وثائق مرسلة.")
    for d in docs:
        icon = "📕" if d["file_ext"] == "pdf" else "📗"
        with st.container(border=True):
            col_info, col_btn, col_del = st.columns([5, 2, 1])
            col_info.markdown(f"{icon} **{d['title']}** (`{d['file_ext'].upper()}`)")
            meta = f"الكلية: {d['target_faculty']} | التاريخ: {d['created_at']} | المرسل: {d['uploader']}"
            col_info.caption(meta)
            col_btn.download_button(
                f"⬇️ تحميل",
                data=bytes(d["file_data"]),
                file_name=d["filename"],
                mime=DOC_MIME.get(d["file_ext"], "application/octet-stream"),
                key=f"dl_doc_{tab_idx}_{d['id']}",
            )
            if is_incubator:
                if col_del.button("🗑️", key=f"del_doc_{tab_idx}_{d['id']}"):
                    ex("DELETE FROM documents WHERE id=?", (d["id"],))
                    flash("تم الحذف.", "🗑️")
                    st.rerun()


# ==========================================
# 5. الواجهة العامة وتسجيل الدخول
# ==========================================
st.set_page_config(page_title="منصة التكامل البيداغوجي والابتكار", layout="wide")

st.markdown(
    """
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
    .stApp, .stMarkdown, .stTextInput, .stTextArea, .stSelectbox, .stAlert,
    [data-testid="stExpander"] { direction: rtl; text-align: right; font-family: Tahoma, sans-serif; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("""
    <div class="p-6 mb-6 bg-gradient-to-r from-blue-900 to-indigo-800 text-white rounded-2xl shadow-lg">
        <h1 class="text-3xl font-bold mb-2">🎓 منصة التكامل بين مدير حاضنة ومصالح البيداغوجيا</h1>
        <p class="text-blue-200">إدارة ومتابعة المشاريع الابتكارية والوثائق بمديرية الحاضنة والكليات الجامعية.</p>
    </div>
""", unsafe_allow_html=True)

if "_flash" in st.session_state:
    _msg, _icon = st.session_state.pop("_flash")
    st.toast(_msg, icon=_icon)

seasons_rows = q("SELECT season_year, is_active FROM academic_seasons ORDER BY id DESC")
seasons_list = [r["season_year"] for r in seasons_rows]
active_season = get_active_season()

st.sidebar.header("⚙️ إدارة المواسم والبوابة")

selected_season = st.sidebar.selectbox(
    "📅 الموسم الجامعي الحالي:",
    seasons_list,
    index=seasons_list.index(active_season) if active_season in seasons_list else 0,
)

all_users = q("SELECT * FROM users ORDER BY id")
user_map = {u["id"]: u for u in all_users}

selected_id = st.sidebar.selectbox(
    "🔐 المسؤول الحالي:",
    list(user_map.keys()),
    format_func=lambda i: f"{user_map[i]['name']} ({user_map[i]['faculty']})",
)
current_user = user_map[selected_id]

# التحقق من كلمة المرور لمدير الحاضنة (123456789)
if current_user["role"] == "INCUBATOR":
    if "incubator_authenticated" not in st.session_state:
        st.session_state["incubator_authenticated"] = False

    if not st.session_state["incubator_authenticated"]:
        pwd_input = st.sidebar.text_input("أدخل كلمة مرور مدير حاضنة:", type="password")
        if st.sidebar.button("تسجيل الدخول كمدير حاضنة"):
            if pwd_input == "123456789":
                st.session_state["incubator_authenticated"] = True
                flash("تم تسجيل الدخول بنجاح.")
                st.rerun()
            else:
                st.sidebar.error("كلمة المرور غير صحيحة!")
        st.stop()
    else:
        if st.sidebar.button("🚪 تسجيل الخروج"):
            st.session_state["incubator_authenticated"] = False
            st.rerun()

    with st.sidebar.expander("➕ فتح موسم جامعي جديد"):
        new_season_input = st.text_input("أدخل سنة الموسم (مثال: 2026/2027)")
        if st.button("تفعيل واعتماد موسم جديد"):
            clean_s = new_season_input.strip()
            if clean_s:
                existing_s = q("SELECT id FROM academic_seasons WHERE season_year=?", (clean_s,), one=True)
                if not existing_s:
                    ex("UPDATE academic_seasons SET is_active=0")
                    ex("INSERT INTO academic_seasons (season_year, is_active) VALUES (?, 1)", (clean_s,))
                    flash(f"تم فتح واعتماد الموسم الجامعي {clean_s} بنجاح!")
                    st.rerun()
                else:
                    ex("UPDATE academic_seasons SET is_active=0")
                    ex("UPDATE academic_seasons SET is_active=1 WHERE season_year=?", (clean_s,))
                    flash(f"تم التبديل إلى الموسم الجامعي {clean_s} وجعله نشطاً!")
                    st.rerun()
            else:
                st.sidebar.warning("يرجى إدخال صيغة صحيحة.")

st.sidebar.markdown("---")
unread = q("SELECT * FROM notifications WHERE user_id=? AND is_read=0 ORDER BY id DESC", (current_user["id"],))
with st.sidebar.expander(f"🔔 الإشعارات ({len(unread)})", expanded=bool(unread)):
    if unread:
        for n in unread[:10]:
            st.write(f"• {n['message']}")
            st.caption(n["created_at"])
        if st.button("تعليم الكل كمقروء", key="read_all"):
            ex("UPDATE notifications SET is_read=1 WHERE user_id=?", (current_user["id"],))
            st.rerun()
    else:
        st.caption("لا توجد إشعارات جديدة.")

# ==========================================
# 6. شريط التبويبات الرئيسي
# ==========================================
tabs = st.tabs([
    "📊 مؤشر المشاريع الابتكارية",
    "كلية العلوم والتكنولوجيا",
    "كلية العلوم الاقتصادية",
    "كلية الآداب واللغات والعلوم الاجتماعية",
    "كلية الحقوق",
    "📂 مكتبة الموارد المشتركة",
])

# ------------------------------------------
# التبويب 0: مؤشر المشاريع الابتكارية
# ------------------------------------------
with tabs[0]:
    st.markdown(f"""
        <div class="p-4 mb-4 bg-indigo-50 border-r-4 border-indigo-600 rounded-lg shadow-sm">
            <h2 class="text-xl font-bold text-indigo-900">📊 مؤشر المشاريع الابتكارية الحية (الموسم النشط: {selected_season})</h2>
        </div>
    """, unsafe_allow_html=True)

    total_projects_current = sum(get_faculty_project_count(fac, selected_season) for fac in FACULTIES)
    total_all_seasons = q("SELECT COUNT(*) AS n FROM faculty_sheets WHERE TRIM(project_name) != ''", one=True)["n"]

    col_sum1, col_sum2, col_sum3 = st.columns(3)
    col_sum1.metric(f"إجمالي مشاريع موسم {selected_season}", f"{total_projects_current} مشروع")
    col_sum2.metric("إجمالي المشاريع المسجلة", f"{total_all_seasons} مشروع")
    col_sum3.metric("الموسم الجامعي المفعل حالياً", selected_season)

    st.markdown("---")
    st.markdown(f"#### 🏛️ توزيع المشاريع الحقيقية حسب الكليات (لموسم {selected_season})")

    fac_cols = st.columns(4)
    for idx, fac_name in enumerate(FACULTIES):
        fac_count = get_faculty_project_count(fac_name, selected_season)
        with fac_cols[idx]:
            with st.container(border=True):
                st.markdown(f"**🎓 {fac_name}**")
                st.metric("عدد المشاريع المسجلة", f"{fac_count} مشروع")

    st.markdown("---")
    st.markdown(f"#### 🍩 الدائرة النسبية لتوزيع المشاريع الابتكارية حسب الكليات (موسم {selected_season})")

    chart_data = []
    for fac_name in FACULTIES:
        cnt = get_faculty_project_count(fac_name, selected_season)
        chart_data.append({"الكلية": fac_name, "عدد المشاريع": cnt})
    df_chart = pd.DataFrame(chart_data)

    if df_chart["عدد المشاريع"].sum() > 0:
        base = alt.Chart(df_chart).encode(
            theta=alt.Theta(field="عدد المشاريع", type="quantitative"),
            color=alt.Color(field="الكلية", type="nominal", legend=alt.Legend(title="الكليات")),
            tooltip=["الكلية", "عدد المشاريع"]
        )
        pie = base.mark_arc(innerRadius=80, outerRadius=140)
        text = base.mark_text(radius=160, size=14).encode(text=alt.Text("عدد المشاريع:Q"))
        chart = (pie + text).properties(width=500, height=350)
        st.altair_chart(chart, use_container_width=True)
    else:
        st.info("لا توجد مشاريع مسجلة في هذا الموسم لعرضها في الدائرة النسبية.")

# ------------------------------------------
# التبويبات 1 إلى 4: فضاءات الكليات الأربع
# ------------------------------------------
with tabs[1]:
    render_faculty_hub("كلية العلوم والتكنولوجيا", tab_idx=1, active_season=selected_season)

with tabs[2]:
    render_faculty_hub("كلية العلوم الاقتصادية", tab_idx=2, active_season=selected_season)

with tabs[3]:
    render_faculty_hub("كلية الآداب واللغات والعلوم الاجتماعية", tab_idx=3, active_season=selected_season)

with tabs[4]:
    render_faculty_hub("كلية الحقوق", tab_idx=4, active_season=selected_season)

# ------------------------------------------
# التبويب 5: مكتبة الموارد المشتركة
# ------------------------------------------
with tabs[5]:
    st.markdown("""
        <div class="p-4 mb-4 bg-emerald-50 border-r-4 border-emerald-600 rounded-lg shadow-sm">
            <h2 class="text-xl font-bold text-emerald-900">📚 مكتبة الموارد والوثائق المشتركة</h2>
        </div>
    """, unsafe_allow_html=True)

    with st.form("res_form", clear_on_submit=True):
        r_title = st.text_input("عنوان المورد")
        r_cat = st.selectbox("التصنيف", RESOURCE_CATEGORIES)
        r_link = st.text_input("رابط التحميل", "https://")
        if st.form_submit_button("إضافة للمكتبة"):
            if not r_title.strip():
                st.warning("أدخل العنوان.")
            elif not valid_url(r_link):
                st.warning("رابط غير صالح.")
            else:
                ex("INSERT INTO resources (title, category, link, uploaded_by) VALUES (?,?,?,?)", (r_title.strip(), r_cat, r_link.strip(), current_user["id"]))
                flash("تمت الإضافة.")
                st.rerun()

    resources = q("SELECT r.*, u.name AS uploader FROM resources r JOIN users u ON u.id = r.uploaded_by ORDER BY r.category, r.id DESC")
    if not resources:
        st.caption("لا توجد موارد.")
    for r in resources:
        col_txt, col_del = st.columns([6, 1])
        col_txt.markdown(f"- 📄 **{r['title']}** ({r['category']}) - [تحميل]({r['link']}) | *بواسطة: {r['uploader']}*")
        if current_user["role"] == "INCUBATOR" or current_user["id"] == r["uploaded_by"]:
            if col_del.button("🗑️", key=f"del_res_{r['id']}"):
                ex("DELETE FROM resources WHERE id=?", (r["id"],))
                flash("تم الحذف.", "🗑️")
                st.rerun()
