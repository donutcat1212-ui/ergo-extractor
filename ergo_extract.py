
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ergo_extract_v3.py

Версия 3: явный учет первичных и выписных документов, приоритет выписки,
более устойчивый разбор таблиц, явный баннер в консоли.
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

VERSION = "ERGO EXTRACT V3.0"
DEFAULT_ROOT_DIR = r"\\fccps.local\dfs\ОМР ЦНС1\Эрготерапия\Эрготерапия 23-25г"
DEFAULT_OUTPUT_XLSX = r"ergotherapy_extract_result_v3.xlsx"

FIELD_DEFINITIONS = [
    ("Шкала SULCS", "", "paired"),
    ("Шкала FIM", "", "paired"),
    ("Шкала EQ-5D", "", "paired"),
    ("ОЦЕНКА COPM", "", "paired"),
    ("d445", "Использование кисти руки", "paired"),
    ("d4458", "Использование кисти и руки, другое уточненное", "paired"),
    ("d540", "Надевание одежды", "paired"),
    ("d6300", "Приготовление простых блюд", "paired"),
    ("е1151", "Вспомогательные изделия и технологии для личного повседневного пользования", "paired"),
    ("е155", "Дизайн, характер проектирования, строительства и обустройства зданий частного использования", "paired"),
    ("е150", "Дизайн, характер проектирования, строительства и обустройства зданий для общественного использования", "paired"),
    ("d4408", "Использование точных движений кисти, другое уточненное (удержание адаптационной ручки при письме правой рукой)", "paired"),
    ("d4301", "Перенос кистями рук", "paired"),
    ("d145", "Усвоение навыков письма (четкое написание букв правой рукой)", "paired"),
    ("d6201", "Приготовление сложных блюд", "paired"),
    ("d550", "Прием пищи (использование столовых приборов правой рукой)", "paired"),
    ("d2302", "Исполнение повседневного распорядка", "paired"),
    ("d5208", "Уход за частями тела, другой уточненный (паретичная верхняя конечность)", "paired"),
    ("d4108", "Изменение позы тела уточненное", "paired"),
    ("d4208", "Перемещение тела другое, уточненное", "paired"),
    ("d5202", "Уход за волосами", "paired"),
    ("d510", "Мытье", "paired"),
    ("d530", "Физиологические отправления", "paired"),
    ("d129", "Целенаправленное использование органов чувств, другое уточненное и не уточненное", "paired"),
    ("d4158", "Поддержание положения тела", "paired"),
    ("d460", "Передвижение в различных местах", "paired"),
    ("d179", "Применение знаний, другое уточненное", "paired"),
    ("норма", "", "activity"),
]

DATE_RE = re.compile(r'(?<!\d)(\d{2}\.\d{2}\.\d{4})(?!\d)')
DATETIME_PATTERNS = [
    re.compile(r'дата/время\s*[:\-]?\s*(\d{2}\.\d{2}\.\d{4}\s*/\s*\d{2}[:.]\d{2})', re.I),
    re.compile(r'дата\s*[:\-]?\s*(\d{2}\.\d{2}\.\d{4}\s*/\s*\d{2}[:.]\d{2})', re.I),
]
COPM_PAIR_RE = re.compile(r'(?<!\d)(\d+(?:[.,]\d+)?)\s*[-–—]\s*(\d+(?:[.,]\d+)?)(?!\d)')
NUMBER_RE = re.compile(r'(?<![\d/])\d+(?:[.,]\d+)?(?![:\d])')
CODE_RE = re.compile(r'^[deе]\d{3,4}$', re.I)

CODE_ALIASES = {
    "d5100": "d510",
    "d5400": "d540",
    "e1151": "е1151",
    "e155": "е155",
    "e150": "е150",
    "е1151": "е1151",
    "е155": "е155",
    "е150": "е150",
}

@dataclass
class FieldDef:
    display: str
    subtitle: str
    kind: str
    @property
    def field_id(self):
        return self.display

FIELDS: List[FieldDef] = [FieldDef(*x) for x in FIELD_DEFINITIONS]
FIELD_BY_DISPLAY = {f.display: f for f in FIELDS}
FIELD_BY_CODE = {}
for f in FIELDS:
    if f.kind == "paired" and f.display.startswith(("d", "е", "e")):
        FIELD_BY_CODE[f.display] = f
        FIELD_BY_CODE[f.display.replace("е", "e")] = f

@dataclass
class ParsedDoc:
    path: str
    year: str
    doc_kind: str
    patient: str
    surname: str
    patient_norm: str
    document_dt: Optional[dt.datetime]
    exam_date: Optional[dt.date]
    primary_eval_date: Optional[dt.date]
    repeat_eval_date: Optional[dt.date]
    activity: str
    metrics: Dict[str, Tuple[str, str]]
    warnings: List[str]

def clean_text(text: str) -> str:
    text = text or ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    return text.strip()

def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", clean_text(text)).strip()

def normalize_match(text: str) -> str:
    text = normalize_spaces(text).lower()
    return text.replace("ё", "е").replace("№", " ")

def normalize_number_text(text: str) -> str:
    return clean_text(text).replace(",", ".")

def parse_value_text(text: str) -> Optional[str]:
    text = normalize_spaces(text)
    if not text or text in {"-", "—", "–"}:
        return None
    pair = COPM_PAIR_RE.search(text)
    if pair:
        return f"{normalize_number_text(pair.group(1))}-{normalize_number_text(pair.group(2))}"
    nums = NUMBER_RE.findall(text)
    if nums:
        return normalize_number_text(nums[0])
    return None

def canonical_code(text: str) -> str:
    t = normalize_match(text).replace("e", "е")
    t = re.sub(r"\s+", "", t)
    t = CODE_ALIASES.get(t, t)
    return t

def surname_from_fio(patient: str) -> str:
    parts = [p for p in normalize_spaces(patient).split(" ") if p]
    return parts[0] if parts else patient

def normalize_patient_name(patient: str) -> str:
    patient = normalize_match(patient)
    patient = re.sub(r'[^а-яa-z0-9.\- ]+', ' ', patient)
    patient = re.sub(r'\s+', ' ', patient).strip()
    return patient

def patient_key_from_name_date(year: str, patient_norm: str, date_obj: Optional[dt.date]) -> str:
    return f"{year}|{patient_norm}|{date_obj.isoformat() if date_obj else 'unknown_date'}"

def parse_datetime(text: str) -> Optional[dt.datetime]:
    text = re.sub(r'\s+', '', text)
    m = re.match(r'^(\d{2}\.\d{2}\.\d{4})/(\d{2})[:.](\d{2})$', text)
    if not m:
        return None
    raw = f"{m.group(1)}/{m.group(2)}:{m.group(3)}"
    try:
        return dt.datetime.strptime(raw, "%d.%m.%Y/%H:%M")
    except ValueError:
        return None

def parse_date_after_label(text: str, label_pattern: str) -> Optional[dt.date]:
    m = re.search(label_pattern + r'\s*[:\-]?\s*(\d{2}\.\d{2}\.\d{4})', text, re.I)
    if not m:
        return None
    try:
        return dt.datetime.strptime(m.group(1), "%d.%m.%Y").date()
    except ValueError:
        return None

def detect_year_from_path(path: str) -> str:
    m = re.search(r'ЭТ\s*(20\d{2})', path)
    return m.group(1) if m else ""

def detect_doc_kind(path: str, full_text: str = "") -> str:
    p = normalize_match(path)
    if "2.выпис" in p or "выписн" in p or "повторн" in p or "выписк" in p:
        return "discharge"
    if "1.первич" in p or "первич" in p:
        return "primary"
    t = normalize_match(full_text)
    rdate = parse_date_after_label(t, r'повторн(?:ая|ой)\s+оценка/дата')
    return "discharge" if rdate else "primary"

def iter_docx_files(root_dir: str) -> Tuple[List[str], List[str]]:
    primaries, discharges = [], []
    for cur_root, _, files in os.walk(root_dir):
        low = normalize_match(cur_root)
        for fn in files:
            if not fn.lower().endswith(".docx") or fn.startswith("~$"):
                continue
            full = os.path.join(cur_root, fn)
            if "1.первич" in low or "первич" in low:
                primaries.append(full)
            elif "2.выпис" in low or "выпис" in low or "повторн" in low:
                discharges.append(full)
    return sorted(primaries), sorted(discharges)

def iter_block_items(doc: Document):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def extract_doc_content(path: str) -> Tuple[str, List[str], List[List[List[str]]]]:
    doc = Document(path)
    paragraphs: List[str] = []
    tables: List[List[List[str]]] = []
    full_parts: List[str] = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            txt = clean_text(block.text)
            if txt:
                paragraphs.append(txt)
                full_parts.append(txt)
        else:
            rows = []
            for row in block.rows:
                cells = [clean_text(c.text) for c in row.cells]
                rows.append(cells)
                for c in cells:
                    if c:
                        full_parts.append(c)
            tables.append(rows)
    return "\n".join(full_parts), paragraphs, tables

def detect_activity(full_text: str, paragraphs: Sequence[str]) -> Tuple[str, List[str]]:
    warnings: List[str] = []
    full_norm = normalize_match(full_text)
    tail_norm = normalize_match("\n".join(paragraphs[-12:] if paragraphs else []))

    if "групповые занятия" in full_norm or "на групповых занятиях" in full_norm:
        return "групповые", warnings
    if "индивидуальные занятия с эрготерапевтом" in full_norm or "на индивидуальных занятиях" in full_norm:
        return "индивидуальные", warnings
    if "занятия не показаны" in full_norm or "проведена консультация" in full_norm or "консультация по " in tail_norm:
        return "консультация", warnings
    if "проводились занятия" in full_norm:
        warnings.append("Тип занятий явно не указан; принято 'индивидуальные'")
        return "индивидуальные", warnings
    return "", warnings

def parse_summary_table(table_rows: List[List[str]]) -> Dict[str, Tuple[str, str]]:
    out: Dict[str, Tuple[str, str]] = {}
    if len(table_rows) < 2:
        return out
    header = " | ".join(table_rows[0])
    if "дата" not in normalize_match(header):
        return out
    for row in table_rows[1:]:
        if not row:
            continue
        label = normalize_match(row[0])
        if len(row) < 3:
            continue
        v1 = parse_value_text(row[1]) or ""
        v2 = parse_value_text(row[2]) or ""
        if "sulcs" in label:
            out["Шкала SULCS"] = (v1, v2)
        elif "fim" in label:
            out["Шкала FIM"] = (v1, v2)
        elif "eq-5d" in label:
            out["Шкала EQ-5D"] = (v1, v2)
        elif "выполнение" in label and "удовлетвор" in label:
            out["ОЦЕНКА COPM"] = (v1, v2)
    return {FIELD_BY_DISPLAY[k].field_id: v for k, v in out.items()}

def parse_problem_table(table_rows: List[List[str]]) -> Dict[str, Tuple[str, str]]:
    out = {}
    if len(table_rows) < 3:
        return out
    header = normalize_match(" | ".join(table_rows[0] + table_rows[1]))
    if "выполнение 1" not in header:
        return out
    for row in table_rows[2:]:
        row_join = normalize_match(" ".join(row))
        if not row_join or "подсчет" in row_join or "изменение" in row_join:
            continue
        def get(idx): return row[idx] if idx < len(row) else ""
        v1 = parse_value_text(get(1)) or ""
        s1 = parse_value_text(get(2)) or ""
        v2 = parse_value_text(get(3)) or ""
        s2 = parse_value_text(get(4)) or ""
        left = f"{v1}-{s1}" if v1 and s1 else ""
        right = f"{v2}-{s2}" if v2 and s2 else ""
        out[FIELD_BY_DISPLAY["ОЦЕНКА COPM"].field_id] = (left, right)
        break
    return out

def parse_narrative_scale_values(full_text: str, doc_kind: str) -> Dict[str, Tuple[str, str]]:
    out: Dict[str, Tuple[str, str]] = {}
    text = full_text

    patterns = {
        "Шкала SULCS": [
            r'емкостн\w* тест верхн\w* конечн\w*\s*sulcs\s*[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)',
            r'шкала\s*sulcs\s*[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)',
        ],
        "Шкала FIM": [
            r'шкала функциональн\w* независим\w*.*?fim.*?[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)',
            r'шкала\s*fim\s*[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)',
        ],
        "Шкала EQ-5D": [
            r'опросник\s*eq-?5d.*?[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)',
            r'шкала\s*eq-?5d\s*[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)',
        ],
    }

    for disp, plist in patterns.items():
        found = ""
        for p in plist:
            m = re.search(p, text, re.I | re.S)
            if m:
                found = normalize_number_text(m.group(1))
                break
        if found:
            out[FIELD_BY_DISPLAY[disp].field_id] = (found, "") if doc_kind == "primary" else (found, "")

    # Narrative COPM in primary docs often means no scored pair; ignore if "не рекомендовано"/"не выявил"
    return out

def parse_icf_profile_table(table_rows: List[List[str]]) -> Dict[str, Tuple[str, str]]:
    out = {}
    if len(table_rows) < 5:
        return out
    flat = normalize_match(" ".join(" ".join(r) for r in table_rows[:4]))
    if "мкф категориальный профиль" not in flat:
        return out

    date_header_row = table_rows[2]
    date_cols = [i for i, cell in enumerate(date_header_row) if DATE_RE.search(cell or "")]
    if not date_cols:
        return out

    for row in table_rows[4:]:
        code = canonical_code(row[0] if row else "")
        if not code or not CODE_RE.match(code.replace("е", "e")):
            continue
        field = FIELD_BY_CODE.get(code)
        if not field:
            continue
        vin = parse_value_text(row[date_cols[0]]) if date_cols[0] < len(row) else None
        vout = parse_value_text(row[date_cols[1]]) if len(date_cols) > 1 and date_cols[1] < len(row) else None
        out[field.field_id] = (vin or "", vout or "")
    return out

def merge_metric_dict(base: Dict[str, Tuple[str, str]], extra: Dict[str, Tuple[str, str]]) -> Dict[str, Tuple[str, str]]:
    res = dict(base)
    for k, (a, b) in extra.items():
        cur = res.get(k, ("", ""))
        if a and b:
            res[k] = (a, b)
            continue
        # avoid overwriting pair with single
        if cur[0] and cur[1]:
            res[k] = cur
            continue
        res[k] = (a or cur[0], b or cur[1])
    return res

def parse_single_doc(path: str) -> ParsedDoc:
    full_text, paragraphs, tables = extract_doc_content(path)
    warnings: List[str] = []

    patient = ""
    m = re.search(r'Ф\.И\.О\.\s*[:\-]?\s*(.+)', full_text, re.I)
    if m:
        patient = normalize_spaces(m.group(1).split("\n")[0])
    if not patient:
        patient = Path(path).stem

    doc_dt = None
    for p in DATETIME_PATTERNS:
        m = p.search(full_text)
        if m:
            doc_dt = parse_datetime(m.group(1))
            if doc_dt:
                break

    exam_date = doc_dt.date() if doc_dt else None
    p_eval = parse_date_after_label(full_text, r'первичн(?:ая|ой)\s+оценка/дата')
    r_eval = parse_date_after_label(full_text, r'повторн(?:ая|ой)\s+оценка/дата')
    activity, aw = detect_activity(full_text, paragraphs)
    warnings.extend(aw)

    doc_kind = detect_doc_kind(path, full_text)
    metrics: Dict[str, Tuple[str, str]] = {}

    # 1) pair-priority tables first
    for table in tables:
        metrics = merge_metric_dict(metrics, parse_summary_table(table))
    for table in tables:
        metrics = merge_metric_dict(metrics, parse_problem_table(table))
    for table in tables:
        metrics = merge_metric_dict(metrics, parse_icf_profile_table(table))

    # 2) fallback narrative scales only where still missing
    narrative = parse_narrative_scale_values(full_text, doc_kind)
    metrics = merge_metric_dict(metrics, narrative)

    # if primary and no repeat date -> force right side empty
    if doc_kind == "primary" and not r_eval:
        metrics = {k: (v[0], "") for k, v in metrics.items()}

    for f in FIELDS:
        if f.kind == "paired":
            metrics.setdefault(f.field_id, ("", ""))

    return ParsedDoc(
        path=path,
        year=detect_year_from_path(path),
        doc_kind=doc_kind,
        patient=patient,
        surname=surname_from_fio(patient),
        patient_norm=normalize_patient_name(patient),
        document_dt=doc_dt,
        exam_date=exam_date,
        primary_eval_date=p_eval,
        repeat_eval_date=r_eval,
        activity=activity,
        metrics=metrics,
        warnings=list(dict.fromkeys(warnings)),
    )

def choose_best_doc(docs: List[ParsedDoc]) -> ParsedDoc:
    def filled_pair_count(d: ParsedDoc) -> int:
        return sum(1 for a,b in d.metrics.values() if a or b)
    def sort_key(d: ParsedDoc):
        return (
            1 if d.doc_kind == "discharge" else 0,
            filled_pair_count(d),
            1 if d.repeat_eval_date else 0,
            1 if d.document_dt else 0,
            d.document_dt or dt.datetime.min,
        )
    return sorted(docs, key=sort_key)[-1]

def reduce_docs(all_docs: List[ParsedDoc]) -> List[ParsedDoc]:
    discharges = [d for d in all_docs if d.doc_kind == "discharge"]
    primaries = [d for d in all_docs if d.doc_kind == "primary"]

    by_key: Dict[str, List[ParsedDoc]] = {}
    for d in discharges:
        key_date = d.repeat_eval_date or d.exam_date or d.primary_eval_date
        key = patient_key_from_name_date(d.year, d.patient_norm, key_date)
        by_key.setdefault(key, []).append(d)

    selected_discharges: List[ParsedDoc] = []
    for docs in by_key.values():
        best = choose_best_doc(docs)
        if len(docs) > 1:
            best.warnings.append(f"Найдено дублей выписки: {len(docs)}. Выбран лучший документ.")
        selected_discharges.append(best)

    selected: List[ParsedDoc] = list(selected_discharges)

    for p in primaries:
        covered = False
        for d in selected_discharges:
            if p.year != d.year or p.patient_norm != d.patient_norm:
                continue
            candidates = [x for x in [p.exam_date, p.primary_eval_date] if x]
            d_candidates = [x for x in [d.primary_eval_date, d.exam_date] if x]
            for a in candidates:
                for b in d_candidates:
                    if abs((a - b).days) <= 14:
                        covered = True
                        break
                if covered:
                    break
            if covered:
                break
        if not covered:
            selected.append(p)

    final_map: Dict[str, List[ParsedDoc]] = {}
    for d in selected:
        eff_date = d.repeat_eval_date or d.exam_date or d.primary_eval_date
        key = patient_key_from_name_date(d.year, d.patient_norm, eff_date)
        final_map.setdefault(key, []).append(d)

    final_docs = []
    for docs in final_map.values():
        best = choose_best_doc(docs)
        if len(docs) > 1:
            best.warnings.append(f"После объединения осталось дублей: {len(docs)}. Выбран лучший документ.")
        final_docs.append(best)

    final_docs.sort(key=lambda d: (d.year, d.patient_norm, d.exam_date or dt.date.min))
    return final_docs

def autosize_columns(ws):
    for col in ws.columns:
        letter = col[0].column_letter
        length = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[letter].width = min(max(length + 2, 12), 60)

def style_header(ws):
    fill = PatternFill("solid", fgColor="D9EAF7")
    font = Font(bold=True)
    for c in ws[1]:
        c.fill = fill
        c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")

def export_xlsx(docs: List[ParsedDoc], out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "cases_wide"

    headers = [
        "Год", "Тип документа", "Дата осмотра", "Дата/время документа",
        "Пациент", "Фамилия", "Нормализованный пациент", "Ключ сопоставления",
        "Характер занятий", "Источник файла", "Предупреждения", "Число извлеченных показателей"
    ]
    for f in FIELDS:
        if f.kind == "paired":
            headers.append(f"{f.display} | поступление")
            headers.append(f"{f.display} | выписка")
    ws.append(headers)

    for d in docs:
        eff_date = d.repeat_eval_date or d.exam_date or d.primary_eval_date
        row = [
            d.year,
            "выписка" if d.doc_kind == "discharge" else "первичный",
            eff_date.isoformat() if eff_date else "",
            d.document_dt.strftime("%d.%m.%Y %H:%M") if d.document_dt else "",
            d.patient,
            d.surname,
            d.patient_norm,
            patient_key_from_name_date(d.year, d.patient_norm, eff_date),
            d.activity,
            d.path,
            " | ".join(d.warnings),
            sum(1 for a,b in d.metrics.values() if a or b),
        ]
        for f in FIELDS:
            if f.kind == "paired":
                a, b = d.metrics.get(f.field_id, ("", ""))
                row.extend([a, b])
        ws.append(row)
    style_header(ws)
    autosize_columns(ws)
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("files_raw")
    ws2.append([
        "Год", "Тип документа", "Пациент", "Дата/время документа", "Дата документа",
        "Первичная оценка", "Повторная оценка", "Характер занятий", "Файл", "Предупреждения"
    ])
    for d in docs:
        ws2.append([
            d.year, d.doc_kind, d.patient,
            d.document_dt.strftime("%d.%m.%Y %H:%M") if d.document_dt else "",
            d.exam_date.isoformat() if d.exam_date else "",
            d.primary_eval_date.isoformat() if d.primary_eval_date else "",
            d.repeat_eval_date.isoformat() if d.repeat_eval_date else "",
            d.activity, d.path, " | ".join(d.warnings)
        ])
    style_header(ws2)
    autosize_columns(ws2)
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("warnings")
    ws3.append(["Тип", "Год", "Пациент", "Дата осмотра", "Файл", "Сообщение"])
    for d in docs:
        for w in d.warnings:
            ws3.append([
                "document_warning", d.year, d.patient,
                (d.repeat_eval_date or d.exam_date or d.primary_eval_date).isoformat()
                if (d.repeat_eval_date or d.exam_date or d.primary_eval_date) else "",
                d.path, w
            ])
    style_header(ws3)
    autosize_columns(ws3)
    ws3.freeze_panes = "A2"

    wb.save(out_path)

def build_arg_parser():
    p = argparse.ArgumentParser()
    p.add_argument("--root", default=DEFAULT_ROOT_DIR)
    p.add_argument("--output", default=DEFAULT_OUTPUT_XLSX)
    return p

def main():
    args = build_arg_parser().parse_args()
    print("=" * 72)
    print(f"{VERSION} | режим: первичные + выписные | приоритет: выписка")
    print("=" * 72)
    print(f"Корневая папка: {args.root}")

    primaries, discharges = iter_docx_files(args.root)
    all_files = primaries + discharges

    print(f"Найдено первичных: {len(primaries)}")
    print(f"Найдено выписных: {len(discharges)}")
    print(f"Итого найдено документов: {len(all_files)}")

    if not all_files:
        print("Не найдено .docx")
        return 1

    parsed: List[ParsedDoc] = []
    for idx, path in enumerate(all_files, start=1):
        try:
            parsed.append(parse_single_doc(path))
        except Exception as e:
            year = detect_year_from_path(path)
            patient = Path(path).stem
            parsed.append(ParsedDoc(
                path=path, year=year, doc_kind=detect_doc_kind(path), patient=patient,
                surname=surname_from_fio(patient), patient_norm=normalize_patient_name(patient),
                document_dt=None, exam_date=None, primary_eval_date=None, repeat_eval_date=None,
                activity="", metrics={f.field_id: ("", "") for f in FIELDS if f.kind == "paired"},
                warnings=[f"Ошибка разбора документа: {e}"]
            ))
        if idx % 50 == 0 or idx == len(all_files):
            print(f"Обработано: {idx}/{len(all_files)}")

    selected = reduce_docs(parsed)
    primary_left = sum(1 for d in selected if d.doc_kind == "primary")
    discharge_left = sum(1 for d in selected if d.doc_kind == "discharge")
    print(f"После объединения оставлено выписок: {discharge_left}")
    print(f"После объединения оставлено первичных без выписки: {primary_left}")

    export_xlsx(selected, args.output)
    print(f"Готово: {args.output}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
