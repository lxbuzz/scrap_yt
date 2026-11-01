# YouTube Comments Capture & Aggregation Pipeline (`scrap_yt`)

Інструментарій для автоматизованого збору, фільтрації, анонімізації та зведення коментарів YouTube за заданими дослідницькими темами (GLP-1 харчування, планування сімейного меню тощо) на базі `yt-dlp` та `openpyxl`.

---

## 📁 Структура проєкту

```text
scrap_yt/
├── .venv/                      # Віртуальне середовище Python
├── output/                     # Згенеровані CSV-батчі та фінальний XLSX (ігнорується в Git)
│   ├── Capture_<Name>_<Date>_1.csv
│   ├── Run_log_<Name>_<Date>_1.csv
│   └── Capture_<Name>_<Date>.xlsx
├── yt_capture_pipeline.py      # Основний скрипт збору та первинної фільтрації коментарів
├── merge_to_xlsx.py            # Скрипт об'єднання всіх батчів в один Excel-файл зі звітом
├── queries_config.csv          # Таблиця пошукових запитів, тем (Theme ID) та пріоритетів
├── scrap_out.xlsx              # Еталонний Excel-шаблон із формулами в стовпцях E, P, Q
├── .gitignore                  # Правила ігнорування службових файлів та зібраних даних
└── README.md                   # Документація проєкту
```

---

## ⚙️ Встановлення та підготовка середовища

1. **Створіть та активуйте віртуальне середовище (Windows PowerShell):**
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. **Встановіть необхідні залежності всередину `.venv`:**
   ```powershell
   python -m pip install --upgrade pip
   python -m pip install yt-dlp openpyxl
   ```
   *(Важливо: у Windows використовуйте команду `python`, а не `py`, щоб гарантовано звертатися до пакунків усередині активованого `.venv`).*

---

## 🚀 Крок 1: Збір коментарів (`yt_capture_pipeline.py`)

Скрипт виконує пошук на YouTube із жорстким дотриманням дослідницького протоколу:
* **Параметри пошуку (Step 2):** Фільтр **This year** (`sp=EgQIAxAB`), сортування за **Relevance**, автоматичний пропуск **YouTube Shorts** (за URL `/shorts/`, тегом `#shorts` або тривалістю $\le 60$ с). Відбираються перші **5 відео** (`Search rank` 1–5).
* **Збір коментарів:** До **15 топових коментарів** (`Comment sort = Top`) верхнього рівня з кожного відео, виключаючи власника каналу (`author_is_uploader == False`).
* **Eligibility Gate (4 перевірки):**
  1. Текст стосується їжі, харчування, приготування, закупівлі продуктів або планування раціону.
  2. Довжина тексту — **щонайменше 10 слів**.
  3. Не є рекламою, промокодом, спам-посиланням, видаленим коментарем або дублікатом (`DUP`).
  4. Написаний реальною людиною (не ботом і не автором відео).
* **Copy Rules та безпека даних (GLP-1 Privacy):**
  * Текст зберігається дослівно (`verbatim`), включно з одруківками. Якщо коментар довший за 300 слів, залишаються перші 300 слів із позначкою `[cut]` у кінці.
  * Згадки `@username` та імена автоматично замінюються на `[name]`.
  * Email-адреси та номери телефонів видаляються (`[redacted]`).
  * Генерується прямий permalink на коментар (`https://www.youtube.com/watch?v=<VIDEO_ID>&lc=<COMMENT_ID>`) та дата у форматі `YYYY-MM-DD`.

### Приклади запуску збору:

* **Stage 1 (лише запити з Priority 1):**
  ```powershell
  python yt_capture_pipeline.py --collector AlexB --priority 1
  ```

* **Stage 2 (запити з Priority 2, ліміт до 60 рядків на тему):**
  ```powershell
  python yt_capture_pipeline.py --collector AlexB --priority 2 --max-per-theme 60
  ```

* **Повний прогін за всіма пріоритетами (Priority 1, 2, 3):**
  ```powershell
  python yt_capture_pipeline.py --collector AlexB
  ```

* **Збереження напряму в папку Google Drive:**
  ```powershell
  python yt_capture_pipeline.py --collector AlexB --output-dir "G:\My Drive\Capture_Output"
  ```

---

## 📊 Крок 2: Об'єднання батчів в один Excel (`merge_to_xlsx.py`)

Якщо збір проводився у кілька сесій (наприклад, у папці `output/` утворилися файли `Capture_AlexB_2026-11-01_1.csv`, `_2.csv`, `_3.csv` та відповідні `Run_log_*.csv`), скрипт `merge_to_xlsx.py` зводить їх в **єдиний фінальний файл** `Capture_<YourName>_<date>.xlsx`.

### Що робить агрегатор:
1. Зчитує всі CSV-батчі з папки `./output` у правильному хронологічному порядку.
2. Повторно перевіряє наскрізну унікальність `URL (permalink)` між різними файлами (видаляє міжбатчеві дублікати `DUP`), а також відсіває порожні (`MISSING`) або короткі (`SHORT`, $<10$ слів) записи.
3. Копіює еталонний шаблон `scrap_out.xlsx` і записує дані у вкладки **`Capture`** та **`Run_log`**, зберігаючи оригінальні формули Excel у стовпцях:
   * **E (`Theme ID`)** — автоматичне визначення коду теми з `Query ID`.
   * **P (`Words`)** — підрахунок слів формулою Excel.
   * **Q (`Row check`)** — автоматичний контроль статусу `OK` / `DUP` / `SHORT` / `MISSING`.
4. Створює додаткову аналітичну вкладку **`Summary`** з живими формулами `COUNTIF` / `SUMIF` для контролю квот за темами (від 30 до 60 рядків на тему).

### Приклади запуску агрегатора:

* **Стандартне зведення всіх CSV з папки `./output` (із вкладками `Capture`, `Run_log`, `Summary`):**
  ```powershell
  python merge_to_xlsx.py --collector AlexB
  ```

* **Суворе зведення лише у 2 вкладки (`Capture` + `Run_log` як у базовому шаблоні) зі згортанням однакових `Query ID` у логах:**
  ```powershell
  python merge_to_xlsx.py --collector AlexB --aggregate-runlog --no-summary-tab
  ```

---

## 📋 Перелік тем та запитів (`queries_config.csv`)

| Query ID | Owner | Theme ID | Тема | Запит (YouTube) | Priority |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **A2.7** | P2 | A2 | Low appetite and under-eating | `what I eat in a day on Zepbound` | **1** |
| **A1.7** | P2 | A1 | Protein and muscle preservation | `high protein meal prep for GLP-1` | **2** |
| **B1.7** | P2 | B1 | Dinner decision fatigue | `what's for dinner this week family of 4` | **2** |
| **B2.7** | P2 | B2 | Picky eaters and kids who refuse | `picky eater family meals` | **2** |
| **A1.8** | P2 | A1 | Protein and muscle preservation | `how to get 100g protein on Zepbound` | **3** |
| **A4.7** | P2 | A4 | Small portions, waste, cooking for one | `GLP-1 meal prep small portions` | **3** |
| **A5.6** | P2 | A5 | Cooking for a household while on GLP-1 | `family dinner GLP-1 meal ideas` | **3** |

---

## 🔒 Безпека та конфіденційність
* Зібрані коментарі за темами GLP-1 належать до чутливої медичної інформації: ніколи не публікуйте вміст папки `output/` у відкритих Git-репозиторіях.
* Перед відправкою фінального файлу переконайтеся, що у стовпці **Q (`Row check`)** на вкладці `Capture` всі рядки мають статус **`OK`**.