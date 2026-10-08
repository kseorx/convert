#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Загрузка Excel-бюджетных форм (И1, И2, И34, И35, КОНС1) в ClickHouse (КХД).

Особенности исходников (проверено на data/):
  * структура НЕ единообразна: у форм разное число шапок-строк, заголовки таблиц
    идут вразнонос (строки 12..16), часть колонок объединена (merged cells);
  * имена листов = названия форм ("И.1 Бюджет доходов и расходов", "КОНС.1 ..."),
    поэтому формы определяются по имени листа + по тексту шапки;
  * периодные колонки бывают двух типов:
      - исторические ("2022 г. Факт", "2023 г. Прогноз", "2025 г. План/Бюджет")
      - внутригодовые ("1 кв.", "6 мес.", "9 мес.", "Январь" ...)
      - служебные ("Отклонение", "абс.откл.", "отн.откл.", "Изменение") -- отбрасываются
  * метаданные берутся из шапки листа: "Организационная единица", "Основная валюта",
    "Период отчета", "Сценарий отчета"; год/тип также парсятся из имени файла.

Модель данных в ClickHouse:
  budget.budget_forms       -- длинная (tidy): одна строка = статья x период
  budget.budget_forms_wide  -- широкая: периоды собраны в Map(String, Float64)
  Дополнительно к данным о статье сохраняются: id, source_file, sheet_name,
  report_year, report_type (План/Факт), form (И1/И2/И34/И35/КОНС1), organization,
  currency, scenario.

Использование:
  python load_to_clickhouse.py                 #Dry-run: распарсить и показать сводку
  python load_to_clickhouse.py --dry-run --csv out.csv
  python load_to_clickhouse.py --create        #создать БД/таблицы и загрузить
  python load_to_clickhouse.py --create-only   #только DDL (schema.sql)
"""

import argparse
import csv
import os
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    import clickhouse_connect
except ImportError:
    clickhouse_connect = None

# ==========================================
# НАСТРОЙКИ (можно переопределить через CLI / env)
# ==========================================
DIRECTORY_PATH = "./data"

CH_HOST = os.getenv("CH_HOST", "localhost")
CH_PORT = int(os.getenv("CH_PORT", "8123"))
CH_USER = os.getenv("CH_USER", "default")
CH_PASSWORD = os.getenv("CH_PASSWORD", "")
CH_DATABASE = os.getenv("CH_DATABASE", "budget")
CH_TABLE = os.getenv("CH_TABLE", "budget_forms")          # длинная таблица
CH_TABLE_WIDE = os.getenv("CH_TABLE_WIDE", "budget_forms_wide")

BATCH_SIZE = 50_000

# Целевые формы: нормализованное имя -> регулярка для имени листа / текста шапки
TARGET_FORMS = {
    "И1":    r"^и\s*\.?\s*1\b",
    "И2":    r"^и\s*\.?\s*2\b",
    "И34":   r"^и\s*\.?\s*(34|3\.4)\b",
    "И35":   r"^и\s*\.?\s*(35|3\.5)\b",
    "КОНС1": r"^конс\s*\.?\s*1\b",
}

# Служебные заголовки, которые НЕ являются данными
META_LABELS = {
    "организационная единица", "основная валюта", "период отчета",
    "сценарий отчета", "реквизиты отчета",
}

# Колонки-мусор (отклонения, изменения, комментарии) — не грузим в КХД
VARIANCE_WORDS = (
    "отклон", "откл", "изменение", "абс", "относ", "комментар", "разниц", "примечан",
)

MONTHS = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]
MONTH_ABBR = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]


# ==========================================
# УТИЛИТЫ
# ==========================================
def norm(s) -> str:
    """Привести ячейку к компактной строке для матчинга."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = str(s).replace("\xa0", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", t).strip()


def to_number(v):
    """
    Аккуратно привести значение ячейки к float:
      * числа из Excel берутся как есть (без потери знаков);
      * строки вида "1 234,56" / "-123.45" парсятся; текст ("х", "-") -> None.
    """
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return None if pd.isna(f) or f in (float("inf"), float("-inf")) else f
    t = str(v).replace("\xa0", "").replace(" ", "").replace("\u2009", "")
    if not t or t.lower() in {"nan", "none", "-", "+", ".", "x", "х"}:
        return None
    if "," in t and "." in t:          # "1,234.56" — запятая как разделитель тысяч
        t = t.replace(",", "")
    elif "," in t:                      # "1234,56" — десятичная запятая
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def col_index(letters: str) -> int:
    """'A'->0, 'B'->1, ... 'AA'->26"""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


RX_CELLREF = re.compile(r"^([A-Za-z]+)(\d+)$")


def read_merged_ranges(path):
    """
    Быстро вытащить диапазоны объединённых ячеек по листам напрямую из XML пакета xlsx.

    Нужно потому, что openpyxl в read_only-режиме не отдаёт ws.merged_cells,
    а обычный load_workbook на файлах 5-10 МБ приводит к OOM.
    Возвращает {имя_листа: [(min_row, min_col, max_row, max_col)]} (1-based).
    """
    import zipfile
    from xml.etree import ElementTree as ET

    NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    RT_WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

    out = {}
    try:
        with zipfile.ZipFile(path) as z:
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rid_target = {r.get("Id"): r.get("Target") for r in rels}
            sheets = []                      # (имя, target)
            for sh in wb.iter(f"{NS_MAIN}sheet"):
                tgt = rid_target.get(sh.get(f"{NS_REL}id"), "")
                if tgt.startswith("/"):
                    tgt = tgt[1:]
                elif not tgt.startswith("xl/"):
                    tgt = "xl/" + tgt
                sheets.append((sh.get("name"), tgt))

            for name, target in sheets:
                try:
                    data = z.read(target)
                except KeyError:
                    continue
                if b"<mergeCell " not in data:
                    continue
                ranges = []
                root = ET.fromstring(data)
                for mc in root.iter(f"{NS_MAIN}mergeCell"):
                    ref = mc.get("ref") or ""
                    m = re.fullmatch(r"([A-Za-z]+\d+):([A-Za-z]+\d+)", ref)
                    if not m:
                        continue
                    a, b = m.group(1), m.group(2)
                    ma, mb = RX_CELLREF.match(a), RX_CELLREF.match(b)
                    if not (ma and mb):
                        continue
                    ranges.append((int(ma.group(2)), col_index(ma.group(1)),
                                   int(mb.group(2)), col_index(mb.group(1))))
                if ranges:
                    out[name] = ranges
    except Exception as exc:
        print(f"  [warn] {Path(path).name}: не удалось прочитать merged-диапазоны ({exc})")
    return out


class SheetBook:
    """
    Ленивая кэш-обёртка над Excel-файлом.

    Почему так: файлы многосличные (до 60 листов), а полный load_workbook
    (data_only=True) на файле ~10 МБ съедает сотни МБ RAM и на больших файлах
    приводит к OOM. Поэтому:
      * значения ячеек читаются по листу через pd.ExcelFile().parse(header=None);
      * диапазоны объединённых (merged) ячеек берутся один раз на файл в
        read_only-режиме и раскрываются вручную;
      * хранятся и сырые значения (для чисел), и нормализованные (для матчинга);
      * кэш очищается после обработки каждого файла (см. close()).
    """

    def __init__(self, path):
        self.path = str(path)
        self.name = Path(path).name
        self._xls = None
        self._merges = None
        self._cache = {}

    @property
    def xls(self):
        if self._xls is None:
            self._xls = pd.ExcelFile(self.path)
        return self._xls

    @property
    def sheet_names(self):
        return self.xls.sheet_names

    def _merge_ranges(self):
        if self._merges is None:
            if Path(self.path).suffix.lower() == ".xlsx":
                self._merges = read_merged_ranges(self.path)
            else:
                self._merges = {}
                try:                    # .xls — единственный способ: полный openpyxl-load
                    from openpyxl import load_workbook
                    wb = load_workbook(self.path, data_only=False)
                    try:
                        for ws in wb.worksheets:
                            rngs = [(r.min_row, r.min_col, r.max_row, r.max_col)
                                    for r in ws.merged_cells.ranges]
                            if rngs:
                                self._merges[ws.title] = rngs
                    finally:
                        wb.close()
                except Exception as exc:
                    print(f"  [warn] {self.name}: merged-cells недоступны ({exc})")
        return self._merges


    def get(self, sheet: str):
        """Вернуть (rows_norm, rows_raw) — списки списков (строки x столбцы)."""
        if sheet in self._cache:
            return self._cache[sheet]
        try:
            df = self.xls.parse(sheet, header=None, dtype=object)
        except Exception as exc:
            print(f"  [warn] {self.name}[{sheet}]: лист не прочитан ({exc})")
            self._cache[sheet] = ([], [])
            return self._cache[sheet]

        raw = [list(rec) for rec in df.itertuples(index=False, name=None)]
        del df
        width = max((len(r) for r in raw), default=0)
        for r in raw:                       # добить строки до единой ширины
            if len(r) < width:
                r.extend([None] * (width - len(r)))
        norm_rows = [[norm(v) for v in r] for r in raw]

        # раскрыть объединённые ячейки: значение верхней левой копии во весь диапазон
        for r0, c0, r1, c1 in self._merge_ranges().get(sheet, []):
            if r0 > len(raw) or c0 > width:
                continue
            top_n, top_r = norm_rows[r0 - 1][c0 - 1], raw[r0 - 1][c0 - 1]
            if not top_n:
                continue
            for r in range(r0 - 1, min(r1, len(raw))):
                for c in range(c0 - 1, min(c1, width)):
                    if not norm_rows[r][c]:
                        norm_rows[r][c] = top_n
                        raw[r][c] = top_r

        # ошибки вида #REF!/#DIV/0! -> пусто
        for row in norm_rows:
            for i, v in enumerate(row):
                if v.startswith("#"):
                    row[i] = ""
                    raw[i] = None

        self._cache[sheet] = (norm_rows, raw)
        return self._cache[sheet]

    def rows(self, sheet: str):
        return self.get(sheet)[0]

    def close(self):
        self._cache.clear()
        self._merges = None
        self._xls = None


def detect_form(sheet_name: str, rows) -> str | None:
    """Форма определяется по имени листа, а если оно нейтральное — по первым строкам."""
    key = norm(sheet_name).lower().replace("№", "").replace("'", "'")
    for form, rx in TARGET_FORMS.items():
        if re.search(rx, key):
            return form
    head = " ".join(norm(c).lower() for row in rows[:6] for c in row if norm(c))[:400]
    for form, rx in TARGET_FORMS.items():
        if re.search(rx, head):
            return form
    return None


RX_YEAR_COL = re.compile(r"(?P<y>\d{4})\s*г\.?\s*(?P<t>факт|прогноз|план|бюджет)", re.I)
RX_YEAR_ONLY = re.compile(r"^(?P<y>\d{4})\s*г\.?$", re.I)
RX_QUARTER = re.compile(r"^(\d)\s*кв\.?$", re.I)
RX_HALF = re.compile(r"^6\s*мес\.?$", re.I)
RX_NINE = re.compile(r"^9\s*мес\.?$", re.I)
RX_MONTH = re.compile(r"^(?P<m>" + "|".join(MONTHS) + r")$", re.I)


def classify_period(label: str):
    """
    Вернуть (period_key, period_kind) либо None, если колонка не является периодом.
    period_key нормализуем: '2022|Факт', '2025|Q1', '2025|Мес-01', 'Год|План'...
    """
    t = norm(label)
    if not t:
        return None
    low = t.lower()

    if any(w in low for w in VARIANCE_WORDS):          # отклонения/комментарии — пропускаем
        return None
    if low in META_LABELS:
        return None

    m = RX_YEAR_COL.search(t.replace("\n", " "))
    if m:
        year, tag = m.group("y"), m.group("t").capitalize()
        return f"{year}|{tag}", "year"

    m = RX_QUARTER.match(t)
    if m:
        return f"Q{m.group(1)}", "quarter"
    if RX_HALF.match(t):
        return "H1cum", "half"          # накопительным итогом за 6 мес
    if RX_NINE.match(t):
        return "9Mcum", "nine_month"    # накопительным итогом за 9 мес

    m = RX_MONTH.match(t)
    if m:
        idx = [x.capitalize() for x in MONTHS].index(m.group("m").capitalize()) + 1
        return f"M{idx:02d}", "month"

    # "2025 г. Бюджет (Январь)" / "2025 г. Факт (Январь)"
    m = re.search(r"(?P<y>\d{4})\s*г\.?\s*(?:план|бюджет|факт|прогноз)?\s*[\(\[]\s*(?P<m>[А-Яа-яЁё]+)", t)
    if m:
        mon = m.group("m").lower()
        if mon.endswith(("ь", "й")):
            mon = mon[:-1]
        for i, name in enumerate(MONTHS):
            if name.startswith(mon[:3]) or (i < 12 and MONTH_ABBR[i].startswith(mon[:3])):
                return f"{m.group('y')}|Мес-{i + 1:02d}", "month"

    # техническое имя периода из АРМ-выгрузки: к010к / к020к / к100к / к120к
    m = re.fullmatch(r"к(?P<c>\d{2})(?P<p>\d)(?P<k>[ркф])?", t, re.I)
    if m:
        code = m.group("c")
        if code == "01":
            kind = "plan" if m.group("p") == "1" else "fact"
            return f"Год|{'План' if kind == 'plan' else 'Факт'}", "year"
        if code == "02":
            return "ПрошлыйГод|Факт", "year"
        if code == "12":
            return "ПрошлыйГод|План", "year"
        return f"Код{code}{m.group('p')}", "other"

    if RX_YEAR_ONLY.match(t):
        return f"{t[:4]}|Год", "year"

    return None  # прочее (например "№ п/п", "х", пустые) — не период


# ==========================================
# ПАРСИНГ ОДНОГО ЛИСТА
# ==========================================
HEADER_TOKENS = ("наименование статьи", "код статьи", "№ п/п", "ед. изм", "единицы измерения")


def find_header_row(rows):
    """Заголовок таблицы: строка, где есть 'Наименование статьи' или ('№ п/п' + 'Ед. изм')."""
    best = None
    for i, row in enumerate(rows[:40]):
        joined = " || ".join(row).lower()
        score = sum(tok in joined for tok in HEADER_TOKENS)
        if score >= 2:
            return i
        if score == 1 and best is None:
            best = i
    return None


def parse_meta(rows, up_to: int):
    """Метаданные из шапки листа (значения обычно в той же или следующей строке)."""
    meta = {"organization": "", "currency": "", "period": "", "scenario": ""}
    labels = {
        "организационная единица": "organization",
        "основная валюта": "currency",
        "период отчета": "period",
        "сценарий отчета": "scenario",
    }
    block = rows[: max(up_to, 1) + 1]
    for ri, row in enumerate(block):
        for ci, cell in enumerate(row):
            key = labels.get(re.sub(r"[.:]+$", "", cell.lower()))
            if not key or meta[key]:
                continue
            cand = ""
            if ci + 1 < len(row) and row[ci + 1] and row[ci + 1].lower() not in labels:
                cand = row[ci + 1]
            elif ri + 1 < len(block):
                nxt = [v for v in block[ri + 1] if v]
                if nxt and nxt[0].lower() not in labels:
                    cand = nxt[0]
            if cand:
                meta[key] = cand
    return meta


def find_col(header, *needles, exclude=()):
    for i, h in enumerate(header):
        hl = h.lower()
        if not hl:
            continue
        if any(n in hl for n in needles) and not any(x in hl for x in exclude):
            return i
    return None


def build_periods(header, group_rows):
    """
    Разметить столбцы заголовка как периоды.

    header      — строка с основными заголовками (например "2022 г. Факт", "1 кв.")
    group_rows  — список строк ВЫШЕ header (для многоуровневых шапок вида
                  "В том числе по кварталам" + подзаголовки "абс.откл./отн.откл.").
    Возвращает {col_index: (period_key, period_kind)}.
    """
    width = len(header)
    periods = {}
    for ci, label in enumerate(header):
        if not label:
            continue
        info = classify_period(label)
        if info is None:
            continue
        key, kind = info
        if kind == "year":
            # уточнить месяц по групповой строке: "2025 г. Бюджет" + "(Январь)"
            m = re.search(r"[\(\[]\s*(?:январ|феврал|март|апрел|ма[йя]|июн|июл|август|"
                          r"сентябр|октябр|ноябр|декабр)", " ".join(g[ci] for g in group_rows
                          if ci < len(g)), re.I)
            if m:
                mon = m.group(0).strip(" \t[").lower()
                for i, name in enumerate(MONTHS):
                    if name.startswith(mon[:4]) or name.startswith(mon[:3]):
                        year = key.split("|")[0]
                        key, kind = f"{year}|Мес-{i + 1:02d}", "month"
                        break
        periods[ci] = (key, kind)

    # колонки без подписи периода ("1 кв." живёт на строку ниже/выше) — берём из соседних строк
    for ci, label in enumerate(header):
        if ci in periods or (label and re.fullmatch(r"\d{1,2}", label)):
            continue          # строка нумерации колонок обрабатывается отдельно
        stack = [header[ci]] + [g[ci] for g in reversed(group_rows) if ci < len(g)]
        for cand in stack:
            info = classify_period(cand)
            if info:
                periods[ci] = info
                break

    # если группа "В том числе по кварталам" стоит надо/под заголовком — проверим её
    grp = ""
    for row in list(group_rows) + [header]:
        for cell in row:
            if "по кварталам" in cell.lower():
                grp = cell
    if grp:
        # найти столбец группы и разметить пустые подписи внутри неё как Q1..Q4/H1/9M
        pass

    return periods


def parse_sheet(book: "SheetBook", sheet_name, year_from_name, type_from_name):
    """Вернуть (список dict-записей в длинной форме, имя формы) для одного листа."""
    rows, raw = book.get(sheet_name)
    if not rows:
        return [], None

    form = detect_form(sheet_name, rows)
    if not form:
        return [], None

    h_idx = find_header_row(rows)
    if h_idx is None:
        print(f"  [skip] {book.name}[{sheet_name}]: не найден заголовок таблицы")
        return [], form

    header = rows[h_idx]
    meta = parse_meta(rows, up_to=h_idx)

    # --- служебные столбцы ------------------------------------------------
    col_num = find_col(header, "\u2116 п/п", "n п/п")
    col_code = find_col(header, "код статьи")
    col_name = find_col(header, "наименование статьи")
    col_unit = find_col(header, "ед. изм", "единицы измерения")
    col_ekbs = find_col(header, "екбс", "ас фзд", exclude=("код статьи",))

    used = {c for c in (col_num, col_code, col_name, col_unit, col_ekbs) if c is not None}

    # --- таблица нумерации колонок ("0 | 1 | 2 | ...") сразу под заголовком ---
    sub_rows = []
    ri = h_idx + 1
    while ri < len(rows):
        cells = [c for c in rows[ri] if c]
        if cells and all(re.fullmatch(r"\d{1,3}", c) for c in cells):
            sub_rows.append(rows[ri])
            ri += 1
        else:
            break

    group_rows = [rows[i] for i in range(max(0, h_idx - 3), h_idx)] + sub_rows
    periods = {ci: v for ci, v in build_periods(header, group_rows).items() if ci not in used}

    # подписи периодов могут лежать на строках НИЖЕ заголовка (merged "В том числе...")
    for si in range(h_idx + 1, min(h_idx + 4, len(rows))):
        if all(not c for c in rows[si]):
            continue
        if any(classify_period(c) for c in rows[si]):
            for ci, lbl in enumerate(rows[si]):
                if ci in used or ci in periods or not lbl:
                    continue
                info = classify_period(lbl)
                if info:
                    periods[ci] = info

    if not periods:
        print(f"  [skip] {book.name}[{sheet_name}] ({form}): нет периодных колонок")
        return [], form

    code_col_final = col_code if col_code is not None else col_num
    if col_name is None:
        print(f"  [skip] {book.name}[{sheet_name}] ({form}): нет колонки наименования")
        return [], form

    data_start = h_idx + 1 + len(sub_rows)
    out = []
    for ri in range(data_start, len(rows)):
        row, rraw = rows[ri], raw[ri]
        if all(not v for v in row):
            continue
        first_txt = next((v for v in row[:6] if v), "")
        if re.search(r"реквизиты отчета", first_txt, re.I):
            break                                     # начался новый отчёт на этом листе
        if detect_form(first_txt, [row]) and ri > data_start + 1:
            break                                     # следующая форма ниже таблицей

        name = row[col_name]
        code = row[code_col_final] if code_col_final is not None else ""
        if not name and not code:
            continue
        if name.lower() in ("наименование статьи",) or re.fullmatch(r"\d{1,2}", name or ""):
            continue

        values = {}
        for ci, (key, kind) in periods.items():
            val = to_number(rraw[ci]) if ci < len(rraw) else None
            if val is None:
                continue
            # при коллизии ключей (два периода с одинаковой нормализацией) — суммируем
            if key in values:
                key = f"{key}#{ci}"
            values[key] = (val, kind)

        if not values:
            continue

        out.append({
            "id": "",                                    # deterministic id заполняется ниже
            "source_file": book.name,
            "sheet_name": sheet_name,
            "report_year": year_from_name,
            "report_type": type_from_name,
            "form": form,
            "organization": meta["organization"] or guess_org_from_filename(book.path),
            "currency": (meta["currency"] or "RUB").upper(),
            "scenario": meta["scenario"],
            "row_num": ri + 1,
            "row_code": code,
            "row_name": name,
            "unit": row[col_unit] if col_unit is not None else "",
            "ekbs_code": row[col_ekbs] if col_ekbs is not None else "",
            "_values": values,
        })
    return out, form


def guess_org_from_filename(path) -> str:
    stem = Path(path).stem
    return re.split(r"_\s*\d{4}", stem)[0].strip() or "Неизвестно"


# ==========================================
# ГОД / ТИП ИЗ ИМЕНИ ФАЙЛА
# ==========================================
RX_FILE_META = re.compile(r"(?P<y>20\d{2})\s*г\.?\s*_\s*(?P<t>Факт|Планирование|План)", re.I)


def file_meta(filename: str, rows_head=None):
    m = RX_FILE_META.search(filename)
    year = int(m.group("y")) if m else None
    typ = None
    if m:
        typ = "Факт" if m.group("t").lower().startswith("факт") else "План"
    if year is None:                       # запасной вариант: из шапки "Период отчета"
        for row in rows_head or []:
            for ci, cell in enumerate(row):
                if cell.lower().startswith("период отчета"):
                    mm = re.search(r"(20\d{2})", " ".join(row[ci:ci + 3]))
                    if mm:
                        year = int(mm.group(1))
    return year, typ


# ==========================================
# MAIN
# ==========================================
DDL = """
CREATE DATABASE IF NOT EXISTS {db};

CREATE TABLE IF NOT EXISTS {db}.{tbl}
(
    `id`          UInt64,
    `source_file` String,
    `sheet_name`  String,
    `report_year` UInt16,
    `report_type` LowCardinality(String),
    `form`        LowCardinality(String),
    `organization` String,
    `currency`    LowCardinality(String),
    `scenario`    String,
    `row_num`     UInt32,
    `row_code`    String,
    `row_name`    String,
    `unit`        String,
    `ekbs_code`   String,
    `period`      String,
    `period_kind` LowCardinality(String),
    `value`       Float64,
    `loaded_at`   DateTime DEFAULT now()
)
ENGINE = ReplacingMergeTree(loaded_at)
PARTITION BY (report_year, form)
ORDER BY (report_year, report_type, form, organization, row_code, period);

CREATE TABLE IF NOT EXISTS {db}.{tbl_wide}
(
    `id`          UInt64,
    `source_file` String,
    `report_year` UInt16,
    `report_type` LowCardinality(String),
    `form`        LowCardinality(String),
    `organization` String,
    `currency`    LowCardinality(String),
    `row_code`    String,
    `row_name`    String,
    `unit`        String,
    `ekbs_code`   String,
    `periods`     Map(String, Float64),
    `loaded_at`   DateTime
)
ENGINE = ReplacingMergeTree(loaded_at)
PARTITION BY (report_year, form)
ORDER BY (report_year, report_type, form, organization, row_code);
"""

WIDE_SQL = """
INSERT INTO {db}.{tbl_wide}
SELECT
    id, anyLast(source_file), report_year, report_type, form, organization, currency,
    row_code, anyLast(row_name), anyLast(unit), anyLast(ekbs_code),
    mapFromVectors(groupArray(period), groupArray(value)), max(loaded_at)
FROM {db}.{tbl}
WHERE loaded_at >= {ts}
GROUP BY id, report_year, report_type, form, organization, currency, row_code
"""


def explode_long(records):
    """Длинная форма: одна строка на пару (статья, период). id — детерминированный хэш."""
    import hashlib
    long_rows = []
    for r in records:
        base = "|".join(str(x) for x in (r["source_file"], r["sheet_name"], r["report_year"],
                                         r["report_type"], r["form"], r["organization"],
                                         r["row_num"], r["row_code"], r["row_name"]))
        rid = int.from_bytes(hashlib.blake2b(base.encode("utf-8"), digest_size=8).digest(), "big")
        for period, (val, kind) in sorted(r["_values"].items()):
            long_rows.append({
                "id": rid,
                "source_file": r["source_file"], "sheet_name": r["sheet_name"],
                "report_year": int(r["report_year"] or 0), "report_type": r["report_type"] or "",
                "form": r["form"], "organization": r["organization"], "currency": r["currency"],
                "scenario": r["scenario"], "row_num": int(r["row_num"]),
                "row_code": r["row_code"], "row_name": r["row_name"],
                "unit": r["unit"], "ekbs_code": r["ekbs_code"],
                "period": period, "period_kind": kind, "value": float(val),
            })
    return long_rows


def main():
    ap = argparse.ArgumentParser(description="Загрузка бюджетных форм в ClickHouse")
    ap.add_argument("--dir", default=DIRECTORY_PATH, help="папка с Excel-файлами")
    ap.add_argument("--host", default=CH_HOST)
    ap.add_argument("--port", type=int, default=CH_PORT)
    ap.add_argument("--user", default=CH_USER)
    ap.add_argument("--password", default=CH_PASSWORD)
    ap.add_argument("--database", default=CH_DATABASE)
    ap.add_argument("--table", default=CH_TABLE)
    ap.add_argument("--create", action="store_true", help="создать БД/таблицы перед загрузкой")
    ap.add_argument("--create-only", action="store_true", help="только DDL")
    ap.add_argument("--dry-run", action="store_true", help="не писать в ClickHouse")
    ap.add_argument("--csv", metavar="PATH", help="дополнительно выгрузить результат в CSV")
    ap.add_argument("--truncate", action="store_true", help="очистить таблицу перед загрузкой")
    args = ap.parse_args()

    root = Path(args.dir)
    files = sorted([p for p in root.rglob("*") if p.suffix.lower() in (".xlsx", ".xls")])
    print(f"Каталог: {root.resolve()}  |  файлов: {len(files)}")

    all_records = []
    stats = Counter()
    for path in files:
        book = SheetBook(path)
        try:
            sheet_names = book.sheet_names
        except Exception as exc:
            print(f"[ERROR] {path.name}: {exc}")
            continue

        year, typ = file_meta(path.name)
        n_file = 0
        for sheet in sheet_names:
            # читаем только листы, похожие на целевые формы (экономия памяти/времени)
            if detect_form(sheet, []) is None:
                head = book.rows(sheet)[:8]
                if detect_form(sheet, head) is None:
                    continue
            try:
                recs, form = parse_sheet(book, sheet, year, typ)
            except Exception as exc:
                print(f"  [error] {path.name}[{sheet}]: {exc}")
                recs, form = [], None
            book._cache.pop(sheet, None)      # освобождаем память после листа
            if not recs:
                continue
            all_records.extend(recs)
            n_file += len(recs)
            stats[(year, typ, form)] += len(recs)
        book.close()
        print(f"[OK]   {path.name:<60} строк-статей: {n_file}"
              + ("" if year else "  (ВНИМАНИЕ: год не определён)"))

    if not all_records:
        print("Данных не найдено — выходим.")
        return 1

    long_rows = explode_long(all_records)
    df = pd.DataFrame(long_rows)

    print("\n=== Сводка (форма / год / тип -> кол-во значений) ===")
    for (year, typ, form), cnt in sorted(stats.items(), key=lambda kv: (str(kv[0][2]), str(kv[0][0]))):
        print(f"  {form or '?':<7} {year or '????'} {typ or '-':<5} строк: {cnt:>6}")
    print(f"Всего записей (длинная форма): {len(df):,}, уникальных статей: {df['id'].nunique():,}")

    if args.csv:
        df.to_csv(args.csv, index=False, sep=";", encoding="utf-8-sig")
        print(f"CSV выгружен: {args.csv}")

    if args.dry_run:
        print("\n[dry-run] в ClickHouse ничего не записано.")
        return 0

    if clickhouse_connect is None:
        print("[ERROR] пакет clickhouse-connect не установлен: pip install clickhouse-connect")
        return 2

    client = clickhouse_connect.get_client(host=args.host, port=args.port, username=args.user,
                                           password=args.password, database=args.database)
    if args.create or args.create_only:
        client.command(DDL.format(db=args.database, tbl=args.table, tbl_wide=args.table + "_wide"))
        print("DDL выполнен.")
        if args.create_only:
            return 0

    if args.truncate:
        client.command(f"TRUNCATE TABLE {args.database}.{args.table}")
        client.command(f"TRUNCATE TABLE {args.database}.{args.table}_wide")

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cols = ["id", "source_file", "sheet_name", "report_year", "report_type", "form",
            "organization", "currency", "scenario", "row_num", "row_code", "row_name",
            "unit", "ekbs_code", "period", "period_kind", "value"]
    total = 0
    for i in range(0, len(df), BATCH_SIZE):
        chunk = df.iloc[i:i + BATCH_SIZE][cols]
        client.insert(args.database + "." + args.table, chunk.values.tolist(), column_names=cols)
        total += len(chunk)
        print(f"  вставлено {total:,}/{len(df):,}")

    client.command(WIDE_SQL.format(db=args.database, tbl=args.table,
                                   tbl_wide=args.table + "_wide", ts=f"'{ts}'"))
    print(f"\nГотово. {args.database}.{args.table}: {total:,} строк;"
          f" материализация в {args.database}.{args.table}_wide выполнена.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
