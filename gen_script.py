script_code = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пайплайн для збору коментарів YouTube через yt-dlp згідно з правилами:
- Пошук YouTube: Filter = This year (EgQIAxAB), Sort = Relevance, Skip Shorts.
- Перші 5 відео по порядку (Search rank 1..5).
- До 15 топових коментарів (Comment sort = Top), не від власника каналу.
- Eligibility gate:
    1) Текст про їжу, харчування, готування, продукти або планування їжі (food/eating/cooking/groceries/meal planning).
    2) Мінімум 10 слів.
    3) Не реклама, не пост лише з посиланням, не видалений і не дублікат URL.
    4) Написано людиною, не ботом і не автором каналу.
- Copy rules:
    - Текст verbatim (як є), якщо > 300 слів -> перші 300 слів + " [cut]".
    - Заміна @username та імен на [name].
    - Видалення телефонів, email, адрес.
    - URL = permalink коментаря (https://www.youtube.com/watch?v=VIDEO_ID&lc=COMMENT_ID).
    - Дати у форматі YYYY-MM-DD.
- Експорт у шаблон Excel (зі збереженням формул у колонках E, P, Q) або у CSV.
"""

import argparse
import csv
import datetime
import json
import os
import random
import re
import shutil
import subprocess
import time
import urllib.parse
from pathlib import Path

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# Список запитів зі скріншотів
DEFAULT_QUERIES = [
    {
        "id": "A2.7",
        "owner": "P2",
        "theme_id": "A2",
        "theme": "A2 Low appetite and under-eating",
        "type": "Pain",
        "query": "what I eat in a day on Zepbound",
        "platform": "YouTube",
        "priority": 1,
    },
    {
        "id": "A1.7",
        "owner": "P2",
        "theme_id": "A1",
        "theme": "A1 Protein and muscle preservation",
        "type": "Pain",
        "query": "high protein meal prep for GLP-1",
        "platform": "YouTube",
        "priority": 2,
    },
    {
        "id": "B1.7",
        "owner": "P2",
        "theme_id": "B1",
        "theme": "B1 Dinner decision fatigue",
        "type": "Pain",
        "query": "what's for dinner this week family of 4",
        "platform": "YouTube",
        "priority": 2,
    },
    {
        "id": "B2.7",
        "owner": "P2",
        "theme_id": "B2",
        "theme": "B2 Picky eaters and kids who refuse",
        "type": "Pain",
        "query": "picky eater family meals",
        "platform": "YouTube",
        "priority": 2,
    },
    {
        "id": "A1.8",
        "owner": "P2",
        "theme_id": "A1",
        "theme": "A1 Protein and muscle preservation",
        "type": "Pain",
        "query": "how to get 100g protein on Zepbound",
        "platform": "YouTube",
        "priority": 3,
    },
    {
        "id": "A4.7",
        "owner": "P2",
        "theme_id": "A4",
        "theme": "A4 Small portions, waste, cooking for one",
        "type": "Pain",
        "query": "GLP-1 meal prep small portions",
        "platform": "YouTube",
        "priority": 3,
    },
    {
        "id": "A5.6",
        "owner": "P2",
        "theme_id": "A5",
        "theme": "A5 Cooking for a household while on GLP-1",
        "type": "Pain",
        "query": "family dinner GLP-1 meal ideas",
        "platform": "YouTube",
        "priority": 3,
    },
]

# Ключові слова для перевірки релевантності темі їжі / кулінарії / продуктів (Eligibility gate #1)
FOOD_KEYWORDS = re.compile(
    r"\\b(food|eat|eating|ate|meal|meals|cook|cooking|cooked|recipe|recipes|"
    r"grocery|groceries|dinner|lunch|breakfast|snack|snacks|protein|calories|"
    r"appetite|hungry|hunger|portion|portions|prep|kitchen|fridge|chicken|meat|"
    r"veggies|vegetables|fruit|shake|shakes|smoothie|yogurt|eggs|oatmeal|diet|"
    r"nutrition|stomach|nausea|full|taste|flavor|flavour|drink|water|hydration|"
    r"fiber|carbs|fat|gram|grams|oz|ounces|dish|dishes|menu|kids|picky|family|"
    r"leftovers|waste|throw away|cooking for one|supper|bite|chew|swallow|digest)\\b",
    re.IGNORECASE,
)

# Ознаки реклами / спаму / ботів (Eligibility gate #3, #4)
AD_BOT_PATTERNS = re.compile(
    r"(use my code|discount code|promo code|link in bio|check out my channel|"
    r"subscribe to my|whatsapp|telegram|crypto|forex|investment|dm me|"
    r"sponsored|affiliate|http://|https://|www\\.)",
    re.IGNORECASE,
)

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+")
PHONE_REGEX = re.compile(
    r"(?:\\+?\\d{1,3}[-.\\s]?)?(?:\\(?\\d{3}\\)?[-.\\s]?)\\d{3}[-.\\s]?\\d{4}\\b"
)
HANDLE_REGEX = re.compile(r"@[a-zA-Z0-9_.-]+")
NAME_INTRO_REGEX = re.compile(
    r"\\b(my name is|i am|i'm)\\s+([A-Z][a-z]{2,15})(?=\\s|[.,!]|$)"
)


def sanitize_and_trim_text(raw_text: str) -> str:
    """
    Виконує правила Copy rules:
    - Замінює @username та явні представлення імені на [name]
    - Видаляє email та номери телефонів
    - Обрізає текст до 300 слів із додаванням [cut] у кінці, зберігаючи verbatim-структуру
    """
    text = raw_text.strip()

    # Заміна @handles на [name]
    text = HANDLE_REGEX.sub("[name]", text)
    # Заміна "My name is John" -> "My name is [name]"
    text = NAME_INTRO_REGEX.sub(r"\\1 [name]", text)

    # Видалення телефонів та email (правило: Never copy phone numbers, emails...)
    text = EMAIL_REGEX.sub("[redacted]", text)
    text = PHONE_REGEX.sub("[redacted]", text)

    # Нормалізація переносів рядків для таблиці, але без зміни орфографії (verbatim)
    text = re.sub(r"\\s+", " ", text).strip()

    words = text.split(" ")
    if len(words) > 300:
        text = " ".join(words[:300]) + " [cut]"

    return text


def count_words(text: str) -> int:
    """Підрахунок слів аналогічно до формули Excel: LEN(TRIM(N))-LEN(SUBSTITUTE(TRIM(N)," ",""))+1"""
    t = re.sub(r"\\s+", " ", text.strip())
    if not t:
        return 0
    return len(t.split(" "))


def is_eligible_comment(
    raw_text: str,
    author_is_uploader: bool,
    author_name: str,
    permalink: str,
    seen_urls: set,
    strict_food_check: bool = True,
) -> tuple[bool, str]:
    """
    Перевіряє 4 критерії Eligibility gate:
    1) Про їжу, харчування, готування, продукти чи планування меню
    2) Мінімум 10 слів
    3) Не реклама, не пост лише з лінком, не видалений, не дублікат
    4) Написано людиною, не ботом і не власником каналу
    """
    if not raw_text or raw_text.strip().lower() in ("[deleted]", "[removed]"):
        return False, "deleted_or_empty"

    if author_is_uploader:
        return False, "channel_owner"

    if author_name and "bot" in author_name.lower():
        return False, "bot_author"

    if permalink in seen_urls:
        return False, "duplicate_url"

    # Перевірка на рекламу та посилання
    if AD_BOT_PATTERNS.search(raw_text):
        return False, "ad_or_link"

    cleaned = sanitize_and_trim_text(raw_text)
    if count_words(cleaned) < 10:
        return False, "under_10_words"

    if strict_food_check and not FOOD_KEYWORDS.search(raw_text):
        return False, "not_about_food"

    return True, cleaned


def build_yt_search_url(query: str) -> str:
    """
    Генерує URL пошуку YouTube з фільтром:
    - This year + Video: sp=EgQIAxAB (або EgIIBQ%3D%3D для This year)
    - Sort = Relevance (за замовчуванням у YouTube)
    """
    encoded_q = urllib.parse.quote_plus(query)
    # EgQIAxAB = Upload date: This year + Type: Video
    return f"https://www.youtube.com/results?search_query={encoded_q}&sp=EgQIAxAB"


def search_top_5_videos(query: str, current_year: int) -> list[dict]:
    """
    Шукає відео через yt-dlp, пропускає Shorts і повертає перші 5 відео по порядку.
    """
    search_url = build_yt_search_url(query)
    cmd = [
        "yt-dlp",
        "--flat-playlist",
        "--dump-json",
        "--playlist-end",
        "25",  # Беремо з запасом, щоб відсіяти Shorts
        search_url,
    ]

    videos = []
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=90)
        for line in proc.stdout.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                continue

            vid_id = data.get("id")
            url = data.get("url") or f"https://www.youtube.com/watch?v={vid_id}"
            title = data.get("title") or ""
            duration = data.get("duration")

            # Пропускаємо Shorts (за URL, назвою або тривалістю <= 60 сек)
            if "/shorts/" in url.lower() or "#shorts" in title.lower():
                continue
            if duration is not None and duration <= 60:
                continue

            videos.append(
                {
                    "id": vid_id,
                    "title": title,
                    "url": f"https://www.youtube.com/watch?v={vid_id}",
                    "channel": data.get("channel") or data.get("uploader") or "YouTube",
                }
            )
            if len(videos) == 5:
                break
    except Exception as e:
        print(f"[!] Помилка пошуку для '{query}': {e}")

    return videos


def fetch_video_comments(video_url: str, max_top_comments: int = 40) -> dict:
    """
    Завантажує метадані відео та топові коментарі (Comment sort = Top) через yt-dlp.
    Беремо до 40 топових коментарів верхнього рівня, щоб після фільтрації (Eligibility gate)
    отримати до 15 якісних коментарів (>=10 слів, про їжу, без власника каналу).
    """
    cmd = [
        "yt-dlp",
        "--skip-download",
        "--write-comments",
        "--no-warnings",
        "--extractor-args",
        f"youtube:comment_sort=top;max_comments={max_top_comments},{max_top_comments},0,0",
        "--dump-single-json",
        video_url,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=120)
        if proc.stdout:
            return json.loads(proc.stdout)
    except Exception as e:
        print(f"[!] Помилка отримання коментарів з {video_url}: {e}")
    return {}


def format_timestamp(ts) -> str:
    """Конвертує timestamp або YYYYMMDD у формат YYYY-MM-DD."""
    if not ts:
        return datetime.date.today().strftime("%Y-%m-%d")
    if isinstance(ts, (int, float)):
        return datetime.datetime.utcfromtimestamp(ts).strftime("%Y-%m-%d")
    ts_str = str(ts)
    if len(ts_str) == 8 and ts_str.isdigit():
        return f"{ts_str[:4]}-{ts_str[4:6]}-{ts_str[6:8]}"
    return datetime.date.today().strftime("%Y-%m-%d")


def export_to_excel_and_csv(
    capture_rows: list[dict],
    run_log_rows: list[dict],
    template_path: str,
    output_xlsx: str,
    output_capture_csv: str,
    output_runlog_csv: str,
):
    """
    Записує зібрані дані:
    1) У сформований Excel-файл за шаблоном Capture_<Name>_<Date>.xlsx
       (зберігаючи оригінальні формули у стовпцях Theme ID [E], Words [P] та Row check [Q]).
    2) У два CSV-файли (Capture та Run_log) для легкого імпорту в Google Sheets / Drive.
    """
    # 1. Експорт у CSV
    capture_headers = [
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
    with open(output_capture_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=capture_headers)
        writer.writeheader()
        for r in capture_rows:
            writer.writerow(r)

    runlog_headers = [
        "Collector",
        "Date",
        "Query ID",
        "Threads / videos / reviews opened",
        "Rows added to Capture",
        "Minutes spent",
        "Notes (no relevant results, mostly ads, blocked...)",
    ]
    with open(output_runlog_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=runlog_headers)
        writer.writeheader()
        for r in run_log_rows:
            writer.writerow(r)

    # 2. Експорт у шаблон Excel (.xlsx)
    if HAS_OPENPYXL and os.path.exists(template_path):
        shutil.copyfile(template_path, output_xlsx)
        wb = openpyxl.load_workbook(output_xlsx)

        ws_cap = wb["Capture"]
        for idx, r in enumerate(capture_rows, start=2):
            ws_cap.cell(idx, 1, r["Collector"])
            ws_cap.cell(idx, 2, r["Date collected"])
            ws_cap.cell(idx, 3, r["Platform"])
            ws_cap.cell(idx, 4, r["Query ID"])
            # Колонка 5 (E) - Theme ID: залишаємо оригінальну формулу Excel!
            ws_cap.cell(idx, 6, r["Source (subreddit / channel / app + store)"])
            ws_cap.cell(idx, 7, r["Parent title (thread, video, app version)"])
            ws_cap.cell(idx, 8, r["Item type"])
            ws_cap.cell(idx, 9, r["URL (permalink)"])
            ws_cap.cell(idx, 10, r["Posted date"])
            ws_cap.cell(idx, 11, r["Score (upvotes / likes / helpful)"])
            ws_cap.cell(idx, 12, r["Rating (stars, stores only)"])
            ws_cap.cell(idx, 13, r["Search rank (1-10 Reddit, 1-5 YouTube)"])
            ws_cap.cell(idx, 14, r["Text (verbatim, full)"])
            ws_cap.cell(idx, 15, r["Note (facts only)"])
            # Колонки 16 (P - Words) та 17 (Q - Row check) вже містять формули в шаблоні!

        ws_log = wb["Run_log"]
        for idx, r in enumerate(run_log_rows, start=2):
            ws_log.cell(idx, 1, r["Collector"])
            ws_log.cell(idx, 2, r["Date"])
            ws_log.cell(idx, 3, r["Query ID"])
            ws_log.cell(idx, 4, r["Threads / videos / reviews opened"])
            ws_log.cell(idx, 5, r["Rows added to Capture"])
            ws_log.cell(idx, 6, r["Minutes spent"])
            ws_log.cell(idx, 7, r["Notes (no relevant results, mostly ads, blocked...)"])

        wb.save(output_xlsx)


def run_pipeline(args):
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    current_year = datetime.date.today().year

    capture_rows = []
    run_log_rows = []
    seen_urls = set()
    theme_counts = {}

    # Фільтруємо запити за пріоритетом (наприклад, Stage 1 = Priority 1, Stage 2 = Priority 2)
    selected_queries = [
        q for q in DEFAULT_QUERIES if args.priority is None or q["priority"] in args.priority
    ]

    print(f"[*] Запуск пайплайну для колектора: {args.collector}")
    print(f"[*] Кількість запитів у черзі: {len(selected_queries)}")

    for q_item in selected_queries:
        qid = q_item["id"]
        theme_id = q_item["theme_id"]
        query_str = q_item["query"]

        # Перевірка ліміту на тему (Stage targets: stop a theme at 60 rows)
        if theme_counts.get(theme_id, 0) >= args.max_per_theme:
            print(f"[*] Тема {theme_id} вже досягла ліміту ({args.max_per_theme} рядків). Пропуск {qid}.")
            continue

        t_start = time.time()
        print(f"\\n[>>>] Обробка {qid} (Priority {q_item['priority']}): '{query_str}'")

        videos = search_top_5_videos(query_str, current_year)
        videos_opened = 0
        rows_added_for_query = 0
        notes_list = []

        if not videos:
            notes_list.append("no relevant results")
        else:
            for rank_idx, vid in enumerate(videos, start=1):
                if theme_counts.get(theme_id, 0) >= args.max_per_theme:
                    notes_list.append(f"theme cap {args.max_per_theme} reached")
                    break

                videos_opened += 1
                print(f"  -> Відео #{rank_idx}: {vid['title'][:60]}... ({vid['url']})")

                vid_data = fetch_video_comments(vid["url"], max_top_comments=45)
                comments = vid_data.get("comments") or []
                channel_name = (
                    vid_data.get("channel")
                    or vid_data.get("uploader")
                    or vid["channel"]
                )
                video_title = vid_data.get("title") or vid["title"]

                captured_for_video = 0
                for c in comments:
                    # Беремо лише коментарі верхнього рівня (або всі, якщо parent == 'root')
                    if c.get("parent") and c.get("parent") != "root":
                        continue

                    cid = c.get("id")
                    if not cid:
                        continue

                    permalink = f"https://www.youtube.com/watch?v={vid['id']}&lc={cid}"
                    raw_text = c.get("text") or ""
                    author_is_uploader = bool(c.get("author_is_uploader"))
                    author_name = c.get("author") or ""

                    ok, cleaned_or_reason = is_eligible_comment(
                        raw_text=raw_text,
                        author_is_uploader=author_is_uploader,
                        author_name=author_name,
                        permalink=permalink,
                        seen_urls=seen_urls,
                        strict_food_check=not args.disable_food_filter,
                    )
                    if not ok:
                        continue

                    seen_urls.add(permalink)
                    posted_date = format_timestamp(c.get("timestamp"))
                    likes = c.get("like_count") or 0
                    words_cnt = count_words(cleaned_or_reason)

                    capture_rows.append(
                        {
                            "Collector": args.collector,
                            "Date collected": today_str,
                            "Platform": "YouTube",
                            "Query ID": qid,
                            "Theme ID": theme_id,
                            "Source (subreddit / channel / app + store)": f"{channel_name} (YouTube)",
                            "Parent title (thread, video, app version)": video_title,
                            "Item type": "video comment",
                            "URL (permalink)": permalink,
                            "Posted date": posted_date,
                            "Score (upvotes / likes / helpful)": likes,
                            "Rating (stars, stores only)": "",
                            "Search rank (1-10 Reddit, 1-5 YouTube)": rank_idx,
                            "Text (verbatim, full)": cleaned_or_reason,
                            "Note (facts only)": "",
                            "Words": words_cnt,
                            "Row check": "OK",
                        }
                    )

                    captured_for_video += 1
                    rows_added_for_query += 1
                    theme_counts[theme_id] = theme_counts.get(theme_id, 0) + 1

                    # До 15 топових коментарів з кожного відео
                    if captured_for_video >= 15:
                        break
                    if theme_counts[theme_id] >= args.max_per_theme:
                        break

                # Динамічна пауза між відео, щоб уникати блокувань
                pause = args.min_pause + random.uniform(0.3 * args.min_pause, 1.5 * args.min_pause)
                time.sleep(pause)

        elapsed_min = max(1, round((time.time() - t_start) / 60.0, 1))
        if rows_added_for_query == 0 and not notes_list:
            notes_list.append("no relevant results matching food/10-word gate")

        run_log_rows.append(
            {
                "Collector": args.collector,
                "Date": today_str,
                "Query ID": qid,
                "Threads / videos / reviews opened": videos_opened,
                "Rows added to Capture": rows_added_for_query,
                "Minutes spent": elapsed_min,
                "Notes (no relevant results, mostly ads, blocked...)": "; ".join(notes_list),
            }
        )

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    out_xlsx = out_dir / f"Capture_{args.collector}_{today_str}.xlsx"
    out_cap_csv = out_dir / f"Capture_{args.collector}_{today_str}.csv"
    out_log_csv = out_dir / f"Run_log_{args.collector}_{today_str}.csv"

    export_to_excel_and_csv(
        capture_rows=capture_rows,
        run_log_rows=run_log_rows,
        template_path=args.template,
        output_xlsx=str(out_xlsx),
        output_capture_csv=str(out_cap_csv),
        output_runlog_csv=str(out_log_csv),
    )

    print(f"\\n[+] Готово! Додано рядків у Capture: {len(capture_rows)}")
    print(f"[+] Excel файл: {out_xlsx}")
    print(f"[+] CSV файли: {out_cap_csv}, {out_log_csv}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="YouTube Comment Capture Pipeline (yt-dlp)")
    parser.add_argument("--collector", default="Alex", help="Ім'я дослідника (Collector)")
    parser.add_argument(
        "--priority",
        type=int,
        nargs="*",
        default=None,
        help="Фільтр за Priority (наприклад: --priority 1 для Stage 1, --priority 2 для Stage 2)",
    )
    parser.add_argument(
        "--template",
        default="scrap_out.xlsx",
        help="Шлях до шаблону Excel із вкладками Capture та Run_log",
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Папка для збереження результатів (можна вказати папку Google Drive)",
    )
    parser.add_argument(
        "--max-per-theme",
        type=int,
        default=60,
        help="Максимальна кількість рядків на одну тему (за правилами Stage targets = 60)",
    )
    parser.add_argument(
        "--min-pause",
        type=float,
        default=3.0,
        help="Мінімальна пауза між запитами (сек) + рандом 30-150%%",
    )
    parser.add_argument(
        "--disable-food-filter",
        action="store_true",
        help="Вимкнути жорсткий словниковий фільтр їжі (якщо тема вже про їжу і потрібні всі коментарі >=10 слів)",
    )
    args = parser.parse_args()
    run_pipeline(args)
'''

with open("yt_capture_pipeline.py", "w", encoding="utf-8") as f:
    f.write(script_code)

print("Saved yt_capture_pipeline.py")
