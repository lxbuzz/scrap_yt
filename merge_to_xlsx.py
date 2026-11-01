#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для об'єднання та підсумовування всіх згенерованих CSV-файлів (Capture_*.csv та Run_log_*.csv)
з папки ./output в ОДИН підсумковий Excel-файл (.xlsx) із вкладками:
1. Capture — об'єднані, очищені (без дублікатів URL і рядків <10 слів) записи з оригінальними Excel-формулами
   у стовпцях Theme ID (E), Words (P) та Row check (Q).
2. Run_log — об'єднаний лог усіх прогонів (сесій/батчів _1, _2, _3 тощо).
3. Summary — зведена таблиця виконання Stage targets за темами (Theme ID) та запитами (Query ID)
   на основі динамічних формул Excel (COUNTIF / SUMIF).
"""

import argparse
import csv
import datetime
import glob
import os
import re
import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


CAPTURE_COLUMNS = [
    "Collector",
    "Date collected",
    "Platform",
    "Query ID",
    "Theme ID",
    "Source (subreddit / channel / app + store)",
    "Parent title (thread, video, app version)",
    "Item type",
    "URL (permalink)",
    "Posted date",
    "Score (upvotes / likes / helpful)",
    "Rating (stars, stores only)",
    "Search rank (1-10 Reddit, 1-5 YouTube)",
    "Text (verbatim, full)",
    "Note (facts only)",
    "Words",
    "Row check",
]

RUNLOG_COLUMNS = [
    "Collector",
    "Date",
    "Query ID",
    "Threads / videos / reviews opened",
    "Rows added to Capture",
    "Minutes spent",
    "Notes (no relevant results, mostly ads, blocked...)",
]


def count_words(text: str) -> int:
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t:
        return 0
    return len(t.split(" "))


def natural_sort_key(path_str: str):
    """Сортування файлів за номерами батчів (_1, _2, _3...)."""
    name = os.path.basename(path_str)
    return [int(s) if s.isdigit() else s.lower() for s in re.split(r"(\d+)", name)]


def load_csv_files(file_paths: list[str]) -> list[dict]:
    rows = []
    for fp in file_paths:
        try:
            with open(fp, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    row["_source_file"] = os.path.basename(fp)
                    rows.append(row)
        except Exception as e:
            print(f"[!] Помилка читання {fp}: {e}")
    return rows


def clean_and_deduplicate_capture(raw_rows: list[dict]) -> tuple[list[dict], dict]:
    """
    Виконує фінальну перевірку згідно з правилом 'Row check column':
    - Пропускає порожні рядки або рядки без Query ID / URL / Text (MISSING)
    - Видаляє рядки < 10 слів (delete SHORT rows)
    - Видаляє дублікати URL між різними CSV-файлами (delete DUP rows)
    """
    cleaned = []
    seen_urls = set()
    stats = {
        "total_read": len(raw_rows),
        "skipped_missing": 0,
        "skipped_short": 0,
        "skipped_dup": 0,
        "kept_ok": 0,
    }

    for r in raw_rows:
        qid = (r.get("Query ID") or "").strip()
        url = (r.get("URL (permalink)") or "").strip()
        text = (r.get("Text (verbatim, full)") or "").strip()

        if not qid or not url or not text:
            stats["skipped_missing"] += 1
            continue

        words = count_words(text)
        if words < 10:
            stats["skipped_short"] += 1
            continue

        if url in seen_urls:
            stats["skipped_dup"] += 1
            continue

        seen_urls.add(url)
        cleaned.append(r)
        stats["kept_ok"] += 1

    return cleaned, stats


def clean_and_merge_runlog(raw_rows: list[dict], aggregate_by_query: bool = False) -> list[dict]:
    """
    Об'єднує рядки Run_log з усіх CSV-файлів.
    Якщо aggregate_by_query=True і той самий Query ID запускався у кількох батчах,
    підсумовує відкриті відео, додані рядки та витрачені хвилини.
    """
    valid_rows = [r for r in raw_rows if (r.get("Query ID") or "").strip()]
    if not aggregate_by_query:
        return valid_rows

    grouped = {}
    for r in valid_rows:
        qid = (r.get("Query ID") or "").strip()
        if qid not in grouped:
            grouped[qid] = {
                "Collector": r.get("Collector", ""),
                "Date": r.get("Date", ""),
                "Query ID": qid,
                "Threads / videos / reviews opened": 0,
                "Rows added to Capture": 0,
                "Minutes spent": 0.0,
                "Notes_list": [],
            }
        try:
            grouped[qid]["Threads / videos / reviews opened"] += int(
                float(r.get("Threads / videos / reviews opened") or 0)
            )
        except ValueError:
            pass
        try:
            grouped[qid]["Rows added to Capture"] += int(
                float(r.get("Rows added to Capture") or 0)
            )
        except ValueError:
            pass
        try:
            grouped[qid]["Minutes spent"] = round(
                grouped[qid]["Minutes spent"] + float(r.get("Minutes spent") or 0.0), 1
            )
        except ValueError:
            pass
        note = (r.get("Notes (no relevant results, mostly ads, blocked...)") or "").strip()
        if note and note not in grouped[qid]["Notes_list"]:
            grouped[qid]["Notes_list"].append(note)

    merged = []
    for qid, data in grouped.items():
        merged.append(
            {
                "Collector": data["Collector"],
                "Date": data["Date"],
                "Query ID": data["Query ID"],
                "Threads / videos / reviews opened": data["Threads / videos / reviews opened"],
                "Rows added to Capture": data["Rows added to Capture"],
                "Minutes spent": data["Minutes spent"],
                "Notes (no relevant results, mostly ads, blocked...)": "; ".join(data["Notes_list"]),
            }
        )
    return merged


def style_sheet_From_Scratch(wb: openpyxl.Workbook):
    """Створює вкладки Capture та Run_log з фірмовим стилем шаблону (#2F6B4F), якщо файл шаблону відсутній."""
    header_fill = PatternFill(start_color="2F6B4F", end_color="2F6B4F", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    ws_cap = wb.active
    ws_cap.title = "Capture"
    for col_idx, h in enumerate(CAPTURE_COLUMNS, start=1):
        c = ws_cap.cell(1, col_idx, h)
        c.fill = header_fill
        c.font = header_font
    ws_cap.freeze_panes = "A2"

    ws_log = wb.create_sheet("Run_log")
    for col_idx, h in enumerate(RUNLOG_COLUMNS, start=1):
        c = ws_log.cell(1, col_idx, h)
        c.fill = header_fill
        c.font = header_font
    ws_log.freeze_panes = "A2"


def build_summary_sheet(wb: openpyxl.Workbook, capture_rows: list[dict], runlog_rows: list[dict]):
    """
    Додає (або оновлює) аналітичну вкладку Summary з динамічними Excel-формулами
    для контролю лімітів за темами (30–60 рядків на тему) та статистики за кожним Query ID.
    """
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws_sum = wb.create_sheet("Summary")

    header_fill = PatternFill(start_color="2F6B4F", end_color="2F6B4F", fill_type="solid")
    sub_fill = PatternFill(start_color="E8F1EC", end_color="E8F1EC", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    reg_font = Font(name="Calibri", size=11)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Таблиця 1: Зведення за темами (Theme ID) — контроль Stage targets (30..60 rows)
    ws_sum.cell(1, 1, "1. Зведення за темами (Stage Targets: 30–60 рядків)").font = Font(
        name="Calibri", size=12, bold=True, color="2F6B4F"
    )
    theme_headers = [
        "Theme ID",
        "Captured Rows (Formula)",
        "Min Target (30)",
        "Max Cap (60)",
        "Target Status",
    ]
    for c_idx, h in enumerate(theme_headers, start=1):
        cell = ws_sum.cell(2, c_idx, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Визначаємо всі унікальні Theme ID
    themes_set = sorted(
        {
            (r.get("Theme ID") or r.get("Query ID", "").split(".")[0]).strip()
            for r in capture_rows
            if (r.get("Query ID") or "").strip()
        }
        | {"A1", "A2", "A4", "A5", "B1", "B2"}
    )
    themes_list = [t for t in themes_set if t]

    cur_row = 3
    for t_id in themes_list:
        ws_sum.cell(cur_row, 1, t_id).font = bold_font
        ws_sum.cell(cur_row, 2, f'=COUNTIF(Capture!$E$2:$E$1501, A{cur_row})').font = reg_font
        ws_sum.cell(cur_row, 3, 30).font = reg_font
        ws_sum.cell(cur_row, 4, 60).font = reg_font
        ws_sum.cell(
            cur_row,
            5,
            f'=IF(B{cur_row}>=D{cur_row},"STOP (Cap 60 reached)",IF(B{cur_row}>=C{cur_row},"TARGET MET (30+)","NEED MORE (<30)"))',
        ).font = bold_font
        for c_i in range(1, 6):
            ws_sum.cell(cur_row, c_i).border = thin_border
        cur_row += 1

    # Підсумковий рядок по темах
    ws_sum.cell(cur_row, 1, "TOTAL").font = bold_font
    ws_sum.cell(cur_row, 2, f"=SUM(B3:B{cur_row-1})").font = bold_font
    ws_sum.cell(cur_row, 5, '=COUNTIF(Capture!$Q$2:$Q$1501, "OK") & " OK rows"').font = bold_font
    for c_i in range(1, 6):
        ws_sum.cell(cur_row, c_i).fill = sub_fill
        ws_sum.cell(cur_row, c_i).border = thin_border

    # Таблиця 2: Зведення за запитами (Query ID)
    q_start = cur_row + 3
    ws_sum.cell(q_start - 1, 1, "2. Деталізація за запитами (Query ID)").font = Font(
        name="Calibri", size=12, bold=True, color="2F6B4F"
    )
    q_headers = [
        "Query ID",
        "Videos Opened (Run_log)",
        "Rows in Capture",
        "Avg Words",
        "Total Minutes",
    ]
    for c_idx, h in enumerate(q_headers, start=1):
        cell = ws_sum.cell(q_start, c_idx, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    q_set = sorted(
        {
            (r.get("Query ID") or "").strip()
            for r in (capture_rows + runlog_rows)
            if (r.get("Query ID") or "").strip()
        }
        | {"A1.7", "A1.8", "A2.7", "A4.7", "A5.6", "B1.7", "B2.7"}
    )
    r_idx = q_start + 1
    for q_id in q_set:
        ws_sum.cell(r_idx, 1, q_id).font = bold_font
        ws_sum.cell(r_idx, 2, f'=SUMIF(Run_log!$C$2:$C$301, A{r_idx}, Run_log!$D$2:$D$301)').font = reg_font
        ws_sum.cell(r_idx, 3, f'=COUNTIF(Capture!$D$2:$D$1501, A{r_idx})').font = reg_font
        ws_sum.cell(
            r_idx,
            4,
            f'=IF(C{r_idx}=0, 0, ROUND(AVERAGEIF(Capture!$D$2:$D$1501, A{r_idx}, Capture!$P$2:$P$1501), 1))',
        ).font = reg_font
        ws_sum.cell(r_idx, 5, f'=SUMIF(Run_log!$C$2:$C$301, A{r_idx}, Run_log!$F$2:$F$301)').font = reg_font
        for c_i in range(1, 6):
            ws_sum.cell(r_idx, c_i).border = thin_border
        r_idx += 1

    ws_sum.cell(r_idx, 1, "TOTAL").font = bold_font
    ws_sum.cell(r_idx, 2, f"=SUM(B{q_start+1}:B{r_idx-1})").font = bold_font
    ws_sum.cell(r_idx, 3, f"=SUM(C{q_start+1}:C{r_idx-1})").font = bold_font
    ws_sum.cell(r_idx, 5, f"=SUM(E{q_start+1}:E{r_idx-1})").font = bold_font
    for c_i in range(1, 6):
        ws_sum.cell(r_idx, c_i).fill = sub_fill
        ws_sum.cell(r_idx, c_i).border = thin_border

    for col in range(1, 6):
        ws_sum.column_dimensions[get_column_letter(col)].width = 26


def write_consolidated_workbook(
    capture_rows: list[dict],
    runlog_rows: list[dict],
    template_path: str,
    output_xlsx: str,
    add_summary_tab: bool = True,
):
    """
    Записує об'єднані дані у файл output_xlsx за зразком шаблону scrap_out.xlsx,
    гарантуючи наявність формул у стовпцях E (Theme ID), P (Words) та Q (Row check)
    для всіх рядків (навіть якщо їх більше, ніж у шаблоні).
    """
    if os.path.exists(template_path):
        shutil.copyfile(template_path, output_xlsx)
        wb = openpyxl.load_workbook(output_xlsx)
    else:
        wb = openpyxl.Workbook()
        style_sheet_From_Scratch(wb)

    ws_cap = wb["Capture"]
    max_cap_rows = max(1501, len(capture_rows) + 50)

    # Очищаємо старі значення даних у Capture (крім заголовка), але відновлюємо формули
    for idx, r in enumerate(capture_rows, start=2):
        ws_cap.cell(idx, 1, r.get("Collector", ""))
        ws_cap.cell(idx, 2, r.get("Date collected", ""))
        ws_cap.cell(idx, 3, r.get("Platform", "YouTube"))
        ws_cap.cell(idx, 4, r.get("Query ID", ""))

        # Колонка E (5): Формула Theme ID із шаблону
        ws_cap.cell(
            idx,
            5,
            f'=IF(D{idx}="","",IF(D{idx}="STORE-CORE","",IF(LEFT(D{idx},6)="STORE-",MID(D{idx},7,5),LEFT(D{idx},FIND(".",D{idx})-1))))',
        )

        ws_cap.cell(idx, 6, r.get("Source (subreddit / channel / app + store)", ""))
        ws_cap.cell(idx, 7, r.get("Parent title (thread, video, app version)", ""))
        ws_cap.cell(idx, 8, r.get("Item type", "video comment"))
        ws_cap.cell(idx, 9, r.get("URL (permalink)", ""))
        ws_cap.cell(idx, 10, r.get("Posted date", ""))

        # Числові поля конвертуємо в int, якщо можливо
        score_val = r.get("Score (upvotes / likes / helpful)", "")
        try:
            score_val = int(float(score_val)) if str(score_val).strip() != "" else 0
        except ValueError:
            pass
        ws_cap.cell(idx, 11, score_val)

        ws_cap.cell(idx, 12, r.get("Rating (stars, stores only)", ""))

        rank_val = r.get("Search rank (1-10 Reddit, 1-5 YouTube)", "")
        try:
            rank_val = int(float(rank_val)) if str(rank_val).strip() != "" else ""
        except ValueError:
            pass
        ws_cap.cell(idx, 13, rank_val)

        ws_cap.cell(idx, 14, r.get("Text (verbatim, full)", ""))
        ws_cap.cell(idx, 15, r.get("Note (facts only)", ""))

        # Колонка P (16): Формула підрахунку слів (Words) із шаблону
        ws_cap.cell(
            idx,
            16,
            f'=IF(N{idx}="","",LEN(TRIM(N{idx}))-LEN(SUBSTITUTE(TRIM(N{idx})," ",""))+1)',
        )
        # Колонка Q (17): Формула перевірки (Row check) із шаблону
        ws_cap.cell(
            idx,
            17,
            f'=IF(COUNTA(A{idx}:D{idx},F{idx}:N{idx})=0,"",IF(OR(D{idx}="",I{idx}="",N{idx}=""),"MISSING",IF(P{idx}<10,"SHORT",IF(COUNTIF($I$2:$I${max_cap_rows},I{idx})>1,"DUP","OK"))))',
        )

    ws_log = wb["Run_log"]
    for idx, r in enumerate(runlog_rows, start=2):
        ws_log.cell(idx, 1, r.get("Collector", ""))
        ws_log.cell(idx, 2, r.get("Date", ""))
        ws_log.cell(idx, 3, r.get("Query ID", ""))

        opened_val = r.get("Threads / videos / reviews opened", 0)
        try:
            opened_val = int(float(opened_val))
        except ValueError:
            pass
        ws_log.cell(idx, 4, opened_val)

        added_val = r.get("Rows added to Capture", 0)
        try:
            added_val = int(float(added_val))
        except ValueError:
            pass
        ws_log.cell(idx, 5, added_val)

        min_val = r.get("Minutes spent", 0)
        try:
            min_val = round(float(min_val), 1)
        except ValueError:
            pass
        ws_log.cell(idx, 6, min_val)

        ws_log.cell(idx, 7, r.get("Notes (no relevant results, mostly ads, blocked...)", ""))

    if add_summary_tab:
        build_summary_sheet(wb, capture_rows, runlog_rows)

    wb.save(output_xlsx)


def main():
    parser = argparse.ArgumentParser(
        description="Об'єднання всіх Capture_*.csv та Run_log_*.csv в ОДИН підсумковий Excel (.xlsx)"
    )
    parser.add_argument(
        "--input-dir",
        default="./output",
        help="Папка зі згенерованими CSV-файлами (за замовчуванням: ./output)",
    )
    parser.add_argument(
        "--template",
        default="scrap_out.xlsx",
        help="Шлях до еталонного Excel-шаблону (scrap_out.xlsx або scrap_out_2.xlsx)",
    )
    parser.add_argument(
        "--collector",
        default="AlexB",
        help="Ім'я дослідника для формування назви фінального файлу Capture_<Name>_<Date>.xlsx",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Явний шлях до вихідного .xlsx файлу (якщо не вказано: ./output/Capture_<Collector>_<Date>.xlsx)",
    )
    parser.add_argument(
        "--aggregate-runlog",
        action="store_true",
        help="Згортати однакові Query ID у вкладці Run_log в один підсумковий рядок (за замовчуванням зберігаються всі рядки батчів)",
    )
    parser.add_argument(
        "--no-summary-tab",
        action="store_true",
        help="Не додавати третю аналітичну вкладку Summary (залишити виключно Capture та Run_log як у шаблоні)",
    )
    args = parser.parse_args()

    in_dir = Path(args.input_dir)
    if not in_dir.exists():
        print(f"[!] Папка {in_dir} не знайдена.")
        return

    # Шукаємо всі файли Capture*.csv та Run_log*.csv (включно з _1, _2, -3 тощо)
    capture_files = sorted(
        glob.glob(str(in_dir / "Capture*.csv")),
        key=natural_sort_key,
    )
    runlog_files = sorted(
        glob.glob(str(in_dir / "Run_log*.csv")),
        key=natural_sort_key,
    )

    # Виключаємо можливо раніше згенеровані MERGED.csv, щоб не дублювати
    capture_files = [f for f in capture_files if "MERGED" not in os.path.basename(f).upper()]
    runlog_files = [f for f in runlog_files if "MERGED" not in os.path.basename(f).upper()]

    print(f"[*] Знайдено файлів Capture CSV ({len(capture_files)}): {[os.path.basename(f) for f in capture_files]}")
    print(f"[*] Знайдено файлів Run_log CSV ({len(runlog_files)}): {[os.path.basename(f) for f in runlog_files]}")

    raw_capture = load_csv_files(capture_files)
    raw_runlog = load_csv_files(runlog_files)

    cleaned_capture, stats = clean_and_deduplicate_capture(raw_capture)
    merged_runlog = clean_and_merge_runlog(raw_runlog, aggregate_by_query=args.aggregate_runlog)

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    if args.output:
        out_xlsx = Path(args.output)
    else:
        out_xlsx = in_dir / f"Capture_{args.collector}_{today_str}.xlsx"

    # Якщо вказаний шаблон не знайдено, перевіряємо scrap_out_2.xlsx або scrap_out.xlsx поруч
    template_path = args.template
    if not os.path.exists(template_path):
        for alt in ["scrap_out_2.xlsx", "scrap_out.xlsx"]:
            if os.path.exists(alt):
                template_path = alt
                break

    write_consolidated_workbook(
        capture_rows=cleaned_capture,
        runlog_rows=merged_runlog,
        template_path=template_path,
        output_xlsx=str(out_xlsx),
        add_summary_tab=not args.no_summary_tab,
    )

    print("\n[+] Зведення завершено успішно!")
    print(f"    - Прочитано рядків Capture: {stats['total_read']}")
    print(f"    - Відсіяно дублікатів (DUP): {stats['skipped_dup']}")
    print(f"    - Відсіяно коротких <10 слів (SHORT): {stats['skipped_short']}")
    print(f"    - Відсіяно неповних (MISSING): {stats['skipped_missing']}")
    print(f"    - Записано чистих рядків (OK) у Capture: {stats['kept_ok']}")
    print(f"    - Записано рядків у Run_log: {len(merged_runlog)}")
    print(f"[+] Фінальний Excel-файл збережено: {out_xlsx}")


if __name__ == "__main__":
    main()
