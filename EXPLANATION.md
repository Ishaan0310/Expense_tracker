# Complete Guide — Both Projects

---

## Part 1 — Which File Belongs to Which Project

```
finance_tracker/           ← Project 1: Personal Finance Variance Tracker
├── requirements.txt       – Python libraries to install
├── fetch_data.py          – downloads Kaggle dataset, builds budget_input.xlsx
├── tracker.py             – reads budget_input.xlsx, computes variance, writes report
└── README.md              – setup guide

skill_tracker/             ← Project 2: Skill Demand Intelligence Tracker
├── requirements.txt       – Python libraries to install
├── fetch_data.py          – downloads Kaggle dataset, writes job_postings.csv
├── load_to_postgres.py    – creates PostgreSQL tables, loads job_postings.csv
├── analyze.py             – runs 5 SQL analyses, saves skill_analysis.xlsx
├── queries.sql            – standalone SQL queries (run directly in psql or pgAdmin)
└── README.md              – setup guide
```

Run order — Project 1:
```
fetch_data.py  →  tracker.py
```

Run order — Project 2:
```
fetch_data.py  →  load_to_postgres.py  →  analyze.py
```

---

## Part 2 — Code Explained, File by File

### Project 1 — Finance Tracker

---

#### `requirements.txt`

```
pandas     – for reading, cleaning, and grouping the CSV data
openpyxl   – for reading and writing .xlsx files with formatting
kaggle     – for downloading the dataset from Kaggle without leaving Python
```

No version numbers — pip will always install the latest stable version.

---

#### `fetch_data.py`

This is the data pipeline. It does four things in sequence:
downloads → cleans → transforms → writes Excel.

**CONFIG block (lines 30–50)**

```python
DATASET_ID    = "bukolafatunde/personal-finance"
COL_DATE      = "Date"
COL_CATEGORY  = "Category"
COL_AMOUNT    = "Amount"
COL_TYPE      = "Type"
EXPENSE_VALUE = "Expense"
BUDGET_FACTOR = 1.10
```

These are the only lines you ever need to touch when switching datasets.
DATASET_ID is the Kaggle dataset URL slug. The COL_* variables map to
the actual column names in that dataset's CSV. If a new dataset calls
its column "Spending" instead of "Amount" — you change one line here,
everything else still works.

**`download()` function**

```python
kaggle.api.authenticate()
kaggle.api.dataset_download_files(DATASET_ID, path=DOWNLOAD_DIR, unzip=True)
```

`kaggle.api.authenticate()` reads your `~/.kaggle/kaggle.json` file
and authenticates silently. Then `dataset_download_files` downloads
the zip and extracts it into `./raw_data/`. You don't need to manually
visit the Kaggle website.

**`find_csv()` function**

```python
csvs = glob.glob(os.path.join(DOWNLOAD_DIR, "**/*.csv"), recursive=True)
return max(csvs, key=os.path.getsize)
```

Kaggle datasets sometimes contain multiple CSV files. `glob` finds
all CSVs anywhere inside `raw_data/` (including subfolders), and we
pick the largest one — which is almost always the main data file.

**`load_and_clean()` function**

```python
df = pd.read_csv(path, parse_dates=[COL_DATE])
df = df[df[COL_TYPE].str.strip().str.lower() == EXPENSE_VALUE.lower()]
df[COL_AMOUNT] = pd.to_numeric(df[COL_AMOUNT], errors="coerce").abs()
df["Month"] = df[COL_DATE].dt.strftime("%B")
```

Line by line:
- `parse_dates=[COL_DATE]` — tells pandas to read the Date column as
  actual date objects, not plain text strings.
- The filter line keeps only rows where Type == "Expense" (ignores income).
- `errors="coerce"` turns anything that can't be a number into NaN,
  then `.abs()` makes all amounts positive (some datasets store expenses
  as negatives).
- `.strftime("%B")` converts a date like 2023-03-15 into "March".

**`build_budget_vs_actual()` function**

```python
actuals = df.groupby([COL_CATEGORY, "Month"])[COL_AMOUNT].sum()
avg_per_cat = actuals.groupby("Category")["Actual"].mean()
actuals["Budgeted"] = actuals["Category"].map(avg_per_cat) * BUDGET_FACTOR
```

First `.groupby()` groups rows by Category + Month and sums spending —
this gives "total groceries spent in January", "total transport in February", etc.

Second `.groupby()` finds the average monthly spend per category across
all months. Then we multiply by BUDGET_FACTOR (1.10 = 10% above average)
to set a realistic but firm budget. The idea: if you normally spend $400
on groceries, the budget becomes $440 — ambitious but not unrealistic.

**`write_excel()` function**

This builds the Excel file using openpyxl directly — not pandas.
Pandas can write Excel but can't do formatting. openpyxl lets you
control colors, borders, cell alignment, number formats precisely.
The output file `budget_input.xlsx` has a "Budget" sheet that
`tracker.py` reads.

---

#### `tracker.py`

This reads `budget_input.xlsx`, does the variance math, and produces
the final formatted report. It doesn't touch the internet — it only
reads and writes local Excel files.

**`load_budget()` function**

```python
df = pd.read_excel(filepath, sheet_name="Budget")
needed = {"Category", "Month", "Budgeted", "Actual"}
missing = needed - set(df.columns)
```

Reads specifically the "Budget" sheet. Validates that the four required
columns exist before doing anything else — if something is missing, it
tells you exactly which column is absent rather than crashing with a
cryptic error.

**`compute_variance()` function**

```python
df["Variance"]   = df["Actual"] - df["Budgeted"]
df["Variance_%"] = (df["Actual"] - df["Budgeted"]) / df["Budgeted"] * 100
df["Flagged"]    = df["Variance_%"] > OVERSPEND_THRESHOLD
```

Simple arithmetic on pandas columns — pandas applies it to every row
at once (vectorized), which is faster than looping. Flagged becomes
True/False and drives all the color coding downstream.

**`write_summary()` / `write_anomalies()` / `write_category_pivot()` functions**

Each writes one sheet to the workbook. Key openpyxl concepts used:
- `PatternFill("solid", fgColor="2B4590")` — sets cell background color
- `Font(bold=True, color="FFFFFF")` — white bold text for headers
- `cell.number_format = "#,##0.00"` — formats numbers with commas and 2 decimals
- `Reference` + `BarChart` — defines data range and adds an embedded chart

**`main()` function**

```python
wb = openpyxl.load_workbook(INPUT_FILE)
write_summary(wb, df)
write_anomalies(wb, df)
write_category_pivot(wb, df)
wb.save(OUTPUT_FILE)
```

It loads the original `budget_input.xlsx` first (so the original Budget
sheet is preserved in the output), adds three new sheets to it, and
saves everything as `budget_report.xlsx`. The original file is untouched.

---

### Project 2 — Skill Tracker

---

#### `requirements.txt`

```
pandas          – data cleaning and CSV processing
psycopg2-binary – Python driver to talk to PostgreSQL
tabulate        – pretty-prints DataFrames as tables in the terminal
openpyxl        – writes the formatted Excel output
kaggle          – downloads dataset from Kaggle
```

`psycopg2-binary` is the "batteries included" version of psycopg2 —
it bundles the C library so you don't need to install PostgreSQL
client libraries separately.

---

#### `fetch_data.py`

**CONFIG block**

Same pattern as Project 1. DATASET_ID controls which dataset downloads.
The COL_* variables map to that dataset's column names.
One extra flag: `DROP_ROWS_NO_SALARY = False` — by default we keep
all rows even without salary, because most job postings don't list one.
Set to True if you only want analyzable salary data.

**`parse_skills()` function**

```python
items = ast.literal_eval(str(val))
return ", ".join(s.strip().title() for s in items if s)
```

The skills column in the Kaggle dataset looks like a Python list written
as a string: `['python', 'sql', 'excel']`. `ast.literal_eval` safely
converts that string into a real Python list. We then join it back into
a clean comma-separated string: `"Python, Sql, Excel"`.

Why not `json.loads()`? JSON uses double quotes. Python list literals
use single quotes. `ast.literal_eval` handles both.

**`extract_experience()` function**

```python
if any(k in t for k in ["senior", "sr.", " sr "]):
    return "Senior"
```

There's no experience_level column in the dataset. We infer it by
scanning the full job title for keywords. `any(k in t for k in [...])` 
is a compact way to check if any of the keyword patterns appear in
the lowercased title string.

**`extract_city()` function**

```python
city = loc.split(",")[0].strip()
```

Job locations come in formats like "Austin, TX" or "New York, NY".
Splitting on comma and taking the first part gives us a clean city name.
"Anywhere" gets mapped to "Remote" — that's a standard Google Jobs value
meaning the role has no location requirement.

---

#### `load_to_postgres.py`

**Why three tables instead of one?**

```sql
jobs       — one row per job posting
skills     — deduplicated list of skill names
job_skills — links jobs to skills (many-to-many)
```

A single job has multiple skills. Storing skills as a comma-separated
string in one column makes querying them very hard. The normalized design
(three tables) lets you write clean SQL like "find all jobs with Python
AND SQL" or "count how many times Python + Tableau appear together".

This is called database normalization — worth mentioning in interviews.

**`CREATE_TABLES_SQL`**

```sql
DROP TABLE IF EXISTS job_skills;
DROP TABLE IF EXISTS skills;
DROP TABLE IF EXISTS jobs;
```

The DROP order matters because of foreign keys — you can't drop `jobs`
if `job_skills` still references it. So we drop in reverse dependency order.

```sql
salary_usd NUMERIC(10, 2)    -- no NOT NULL: many listings don't show salary
```

The `--` comment explains why nullable. NUMERIC(10,2) allows salaries
up to 99,999,999.99 — wide enough for USD annual figures.

**`execute_values()` from psycopg2**

```python
execute_values(cur, "INSERT INTO jobs (...) VALUES %s RETURNING id, raw_skills", data)
```

`execute_values` does a bulk insert — it sends all rows in one SQL
statement instead of one INSERT per row. For 785K job postings, this
is the difference between finishing in 30 seconds vs 30 minutes.

`RETURNING id, raw_skills` gives back the auto-generated IDs alongside
the skills text, which we need to build the junction table next.

**`build_skill_index()` function**

Collects every unique skill name from all job rows, inserts them into
the `skills` table, then builds a Python dictionary mapping each skill
name to its database ID. This dictionary is used in `insert_job_skills()`
to create the junction rows.

---

#### `analyze.py`

**Why use pandas to run SQL queries?**

```python
def run_query(conn, sql):
    return pd.read_sql_query(sql, conn)
```

`pd.read_sql_query` runs a SQL query via the psycopg2 connection and
returns the result as a DataFrame. This is cleaner than manually
iterating cursor results — and we get the full pandas toolkit for free
(display, export, etc.).

**The self-join pattern in skill pair queries**

```sql
FROM job_skills js1
JOIN job_skills js2 ON js1.job_id = js2.job_id
                   AND js1.skill_id < js2.skill_id
```

We join the job_skills table to itself on the same job_id — this gives
us every combination of skills for each job. The condition
`js1.skill_id < js2.skill_id` ensures we get each pair only once
(without it, Python+SQL and SQL+Python would both appear).

**The window function in city rankings**

```sql
ROW_NUMBER() OVER (PARTITION BY j.city ORDER BY COUNT(*) DESC) AS rn
```

`PARTITION BY city` means "restart the row counter for each city".
`ORDER BY COUNT(*) DESC` ranks skills within that city by frequency.
The outer query then filters `WHERE rn <= 5` to keep only the top 5
per city. This is a standard SQL pattern for "top N per group".

**`PERCENTILE_CONT(0.5)`**

```sql
PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY salary_usd)
```

This calculates the median salary (the value at the 50th percentile).
Median is more meaningful than average for salaries because a few
very high earners can skew the average significantly.

---

#### `queries.sql`

This file has the exact same queries as `analyze.py` but formatted
for running directly in psql or pgAdmin4. Useful for:
- Exploring the data interactively
- Modifying a query quickly without editing Python
- Showing in an interview what your SQL looks like

---

## Part 3 — Switching Datasets: What to Change

### Project 1 (fetch_data.py)

Open `fetch_data.py` and update the CONFIG section only:

```python
DATASET_ID    = "new_author/new_dataset"   # ← Kaggle dataset slug
COL_DATE      = "TransactionDate"          # ← actual date column name in new CSV
COL_CATEGORY  = "Expense Category"         # ← actual category column name
COL_AMOUNT    = "Spent"                    # ← actual amount column name
COL_TYPE      = None                       # ← set None if all rows are already expenses
EXPENSE_VALUE = "Debit"                    # ← only matters if COL_TYPE is not None
BUDGET_FACTOR = 1.15                       # ← change the budget strictness
```

That's it. `tracker.py` does NOT need any changes — it always reads
`budget_input.xlsx` with the same four column names (Category, Month,
Budgeted, Actual), which `fetch_data.py` always writes regardless of
the source dataset.

If the new dataset's date format confuses `parse_dates`, add this to
`load_and_clean()`:
```python
df[COL_DATE] = pd.to_datetime(df[COL_DATE], dayfirst=True)  # for DD/MM/YYYY
```

---

### Project 2 (fetch_data.py)

Open `fetch_data.py` and update CONFIG:

```python
DATASET_ID       = "new_author/new_dataset"
COL_TITLE_SHORT  = "role"              # ← simplified job title column
COL_TITLE_FULL   = "full_title"        # ← title with seniority keywords
COL_COMPANY      = "employer"          # ← company name column
COL_LOCATION     = "location"          # ← city/state column
COL_SKILLS       = "required_skills"   # ← skills column
COL_SALARY       = "annual_salary"     # ← salary column (USD)
```

The four processing functions may need updates too, depending on how
the new dataset formats its data:

- `parse_skills()` — if skills are stored as JSON (double-quotes) instead
  of Python list literals, change `ast.literal_eval` to `json.loads`.
  If skills are already a plain comma-separated string, remove the
  parsing logic and just return the value directly.

- `extract_experience()` — if the new dataset has a seniority column
  directly, replace this function with: `return row["seniority_column"]`

- `extract_city()` — if location is already a clean city name (not
  "City, State" format), simplify to: `return str(location).strip()`

`load_to_postgres.py`, `analyze.py`, and `queries.sql` do NOT need
changes — they work with the standardized `job_postings.csv` columns
that `fetch_data.py` always outputs.

---

## Part 4 — Running in Your IDE

### PyCharm (recommended for Python files)

**First-time project setup:**

1. Open PyCharm → File → Open → select the project folder
   (`finance_tracker` or `skill_tracker`)

2. PyCharm will show a banner at the top: "No interpreter configured".
   Click "Add interpreter" → "Add local interpreter" → "Virtualenv environment"
   → "New" → click OK.

3. A `venv/` folder appears in your project. This is the isolated
   Python environment.

4. Open the Terminal tab (bottom of screen):
   ```bash
   pip install -r requirements.txt
   ```

5. **Set up Kaggle credentials** (one-time, for both projects):
   - Go to kaggle.com → your profile icon → Settings → API → "Create New Token"
   - This downloads a file called `kaggle.json`
   - In Terminal:
     ```bash
     mkdir -p ~/.kaggle
     cp ~/Downloads/kaggle.json ~/.kaggle/kaggle.json
     chmod 600 ~/.kaggle/kaggle.json
     ```

**Running a file:**

Right-click any `.py` file in the left panel → "Run 'filename'"

Or open the file → click the green ▶ button at the top right.

Output appears in the "Run" tab at the bottom.

**Suggested run order:**
- Project 1: right-click `fetch_data.py` → Run → then right-click `tracker.py` → Run
- Project 2: Run `fetch_data.py` → `load_to_postgres.py` → `analyze.py`

**PyCharm database tool (optional, PyCharm Professional only):**

View → Tool Windows → Database → + → Data Source → PostgreSQL
Fill in: Host=localhost, Port=5432, Database=skill_tracker, User=your_mac_username
Click "Test Connection" → Apply. You can now browse tables and run SQL
directly inside PyCharm without switching to pgAdmin.

---

### pgAdmin 4 (for SQL exploration — Project 2 only)

pgAdmin is a GUI for PostgreSQL. Use it to create the database,
inspect tables, and run queries from `queries.sql`.

**Create the database:**

1. Open pgAdmin 4
2. In the left panel: Servers → expand → right-click "Databases" → Create → Database
3. Name: `skill_tracker` → Save

**Verify data loaded (after running load_to_postgres.py):**

Expand: skill_tracker → Schemas → public → Tables
You should see: `jobs`, `skills`, `job_skills`
Right-click `jobs` → "View/Edit Data" → "First 100 rows" to preview.

**Run queries from queries.sql:**

1. Right-click `skill_tracker` → Query Tool
2. In the menu bar: File → Open → select `queries.sql`
3. Highlight any single query block (Ctrl+A to select all)
4. Click the ▶ Execute button (or press F5)
5. Results appear in the "Data Output" tab below

**Useful pgAdmin features:**
- The "Explain" button (Shift+F7) shows the query execution plan —
  good for understanding how the self-join queries work
- "Save results to file" in Data Output lets you export query results to CSV

---

### VSCode (alternative to PyCharm — works for both projects)

**Setup:**

1. Open VSCode → File → Open Folder → select the project folder

2. Install the Python extension (if not installed):
   Extensions panel (Ctrl+Shift+X) → search "Python" → install Microsoft's extension

3. Create virtual environment:
   Open terminal (Ctrl+` ) →
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

4. Select the interpreter:
   Bottom-left of VSCode shows the Python version → click it →
   "Enter interpreter path" → type `./venv/bin/python` → Enter

**Running a file:**

- Open the `.py` file → click the ▶ Run button at the top right
- Or in terminal: `python3 fetch_data.py`

**SQL files in VSCode:**

Install the "PostgreSQL" extension by Chris Kolkman to get syntax
highlighting and the ability to run SQL queries directly in VSCode
against your PostgreSQL instance. After installing:
- Open `queries.sql`
- Right-click → "Run Query" (after configuring a connection)

---

## Quick Reference

| File | Project | What it does | When to run |
|---|---|---|---|
| `finance_tracker/fetch_data.py` | 1 | Downloads Kaggle dataset, writes `budget_input.xlsx` | First, once |
| `finance_tracker/tracker.py` | 1 | Reads Excel, computes variance, writes `budget_report.xlsx` | Every time you update data |
| `finance_tracker/requirements.txt` | 1 | Library list for pip | Once during setup |
| `skill_tracker/fetch_data.py` | 2 | Downloads Kaggle dataset, writes `job_postings.csv` | First, once |
| `skill_tracker/load_to_postgres.py` | 2 | Creates DB tables, loads `job_postings.csv` into PostgreSQL | Once after fetch |
| `skill_tracker/analyze.py` | 2 | Runs 5 SQL analyses, saves `skill_analysis.xlsx` | Whenever you want results |
| `skill_tracker/queries.sql` | 2 | Same queries, formatted for psql / pgAdmin | For manual exploration |
| `skill_tracker/requirements.txt` | 2 | Library list for pip | Once during setup |
