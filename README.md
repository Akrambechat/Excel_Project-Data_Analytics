# 📊 Excel Project – Data Analytics

> An interactive Excel dashboard that uses Power Query, a Power Pivot data model and PivotCharts to explore which skills data-job postings ask for, and how pay varies by job title, country and skill.

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query-M-F2C811)
![Power Pivot](https://img.shields.io/badge/Power%20Pivot-Data%20Model-F2C811)
![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)

---

## 📌 Overview

`Data_Analytics_Project.xlsx` is a single workbook that answers three questions about data-related job postings:

1. **Which skills are requested the most?**
2. **How much do these roles pay, and how does it differ between the US and the rest of the world?**
3. **Which skills are associated with the highest median salary?**

The raw job-postings data is loaded and cleaned with **Power Query**, split into two related tables in a **Power Pivot data model**, summarised with **DAX measures** and **PivotTables**, and presented through **PivotCharts** with **slicers** (Job Title and Country) for interactive filtering.

---

## ✨ Key Features

- **ETL in Power Query:** type conversion, text cleaning, date extraction, a generated `job_id` key and a normalised skills table, all as reproducible M queries.
- **Two-table data model:** `data_jobs_salary` and `data_jobs_skills`, linked on `job_id` (one posting → many skills).
- **Custom measures in the data model**, used across the PivotTables: `median salary`, `Median salary US`, `Median salary <>US`, `Median salary skills`, `skill likelihood`, `job_per_skills`, `skills count`, `job count`.
- **5 analysis sheets**, each driven by its own PivotTable on the shared data model.
- **3 PivotCharts/charts:** horizontal bar, scatter plot with data labels, and a bar + line combo chart.
- **Interactive slicers** for **Job Title** and **Country** (111 countries, 10 job titles), connected to several PivotTables at once.

---

## 🛠 Tech Stack & Tools

| Layer | Tool |
|---|---|
| Spreadsheet / UI | Microsoft Excel (workbook `.xlsx`) |
| Data extraction & cleaning | Power Query (M language) |
| Data modelling & measures | Power Pivot (Excel Data Model, DAX measures) |
| Analysis | PivotTables with Top-N filters and measure-based columns |
| Visualisation | PivotCharts, scatter chart, combo chart, slicers |

---

## 🗂 Workbook Structure

| Sheet | PivotTable | What it shows | Visual |
|---|---|---|---|
| `Skill_Analysis` | PivotTable1 | Count of skill mentions per job title | — |
| `Skill_Job_Analysis` | PivotTable2 | Top 10 skills by number of postings | Bar chart *"What are the top skills of Data Nerds?"* + Job Title / Country slicers |
| `Salary_vs_skills` | PivotTable3 | Median salary vs. average number of skills requested, per job title | Scatter chart (median salary vs. average skills requested) |
| `Salary_analysis` | PivotTable4 | Median salary per job title: overall, US, and non-US | Country slicer |
| `Skill salary analysis` | PivotTable6 | Median salary and "skill likelihood" for the top 10 skills | Combo chart *"What's the top pay of the top 10 skills?"* + Job Title / Country slicers |

```text
.
├── Data_Analytics_Project.xlsx   # The full project (queries, data model, pivots, charts)
├── LICENSE                       # MIT
└── README.md
```

---

## 💻 Code Demonstrations & Examples

The queries below are taken from the workbook's Power Query (M) code. Long column lists are shortened with `...`, and the local file path is replaced by a placeholder.

### 1. Cleaning and enriching the job postings (`data_jobs_salary`)

```powerquery
let
    Source = Excel.Workbook(File.Contents("<path>\data_jobs_salary_all.xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",
        {{"job_title_short", type text}, ..., {"job_posted_date", type datetime},
         {"salary_year_avg", Currency.Type}, {"salary_hour_avg", Currency.Type}, ...}),
    #"Replaced ""via "" in job_via" = Table.ReplaceValue(#"Changed Type","via ","",Replacer.ReplaceText,{"job_via"}),
    #"Inserted Month" = Table.AddColumn(..., "job_posted_month", each Date.Month([job_posted_datetime]), Int64.Type),
    #"Inserted Multiplication" = Table.AddColumn(..., "salary_hour_adjusted", each [salary_hour_avg] * 2080, Currency.Type),
    #"Added Index" = Table.AddIndexColumn(..., "job_id", 10001, 1, Int64.Type)
in
    #"Reordered Id Column"
```

**What it does:** loads `Sheet1` of the source workbook, promotes headers and sets proper data types (text, logical, datetime, currency). It then:

- strips the `"via "` prefix from `job_via`;
- adds a `job_posted_month` column and a date-only `job_posted_date` column;
- converts hourly pay to an annual equivalent (`salary_hour_adjusted = salary_hour_avg × 2080`, i.e. 40 h × 52 weeks);
- generates a unique **`job_id`** (starting at 10001), which later links the two model tables.

### 2. Turning a skills list into one row per skill (`data_jobs_skills`)

```powerquery
let
    Source = data_jobs_salary,
    #"Removed Other Columns" = Table.SelectColumns(Source, {"job_id", "job_skills"}),
    #"Replaced [ Value" = Table.ReplaceValue(#"Removed Other Columns", "[", "", Replacer.ReplaceText, {"job_skills"}),
    // ... same replacement for "]" and "'"
    #"Split Skills Column by Delimiter" = Table.SplitColumn(..., "job_skills",
        Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"job_skills.1", ..., "job_skills.24"}),
    #"Unpivoted Skill Columns" = Table.UnpivotOtherColumns(#"Changed Type", {"job_id"}, "Attribute", "job_skills"),
    #"Trimmed Skill Text" = Table.TransformColumns(..., {{"job_skills", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(..., {{"job_skills", Text.Proper, type text}}),
    #"Added Conditional Column" = Table.AddColumn(#"Capitalized Each Word", "Custom",
        each if [job_skills] = "Sql" then "SQL"
        else if [job_skills] = "Aws" then "AWS"
        else if [job_skills] = "Sas" then "SAS"
        else if [job_skills] = "Nosql" then "NoSQL"
        else if [job_skills] = "Power Bi" then "Power BI"
        else if [job_skills] = "Gcp" then "GCP"
        else [job_skills])
in
    #"Renamed Columns"
```

**What it does:** builds a second table from the first. Each posting's skills string (for example `['sql', 'python']`) is stripped of brackets and quotes, split into up to 24 columns, then **unpivoted** into one row per `(job_id, skill)` pair. Skill names are trimmed, capitalised with `Text.Proper`, and acronyms that `Text.Proper` breaks (SQL, AWS, SAS, NoSQL, Power BI, GCP) are restored by a conditional column.

### 3. Data model relationship

```text
data_jobs_skills[job_id]  ──►  data_jobs_salary[job_id]
```

Every PivotTable in the workbook queries this model, so a slicer on `job_title_short` or `job_country` (both columns of `data_jobs_salary`) also filters the skill counts coming from `data_jobs_skills`.

### 4. Measures used by the PivotTables

| Measure | Used in | Purpose |
|---|---|---|
| `Count of job_skills` | `Skill_Analysis`, `Skill_Job_Analysis` | Number of skill mentions |
| `median salary` | `Salary_analysis`, `Salary_vs_skills` | Median salary per job title |
| `Median salary US` / `Median salary <>US` | `Salary_analysis` | Median salary for US vs. non-US postings |
| `job_per_skills` | `Salary_vs_skills` | Average number of skills requested per posting (scatter-chart axis) |
| `Median salary skills` | `Skill salary analysis` | Median salary for postings that list a given skill |
| `skill likelihood` | `Skill salary analysis` | Share of postings that list the skill |

> The measure definitions are stored inside the Excel Data Model. To read the DAX, open **Power Pivot → Manage → Measures** (or **Data → Manage Data Model**).

---

## 📈 Results Snapshot

The values below are the ones saved in the workbook. The slicers were saved with **Country = United States**, and the two skill sheets (`Skill_Job_Analysis`, `Skill salary analysis`) additionally with **Job Title = Data Analyst**. Change the slicers to see other views.

**Top 10 skills for US Data Analyst postings**

| Skill | Postings | Median salary (USD) | Skill likelihood |
|---|---:|---:|---:|
| SQL | 4,378 | 90,000 | 52.5 % |
| Excel | 3,457 | 84,500 | 41.5 % |
| Tableau | 2,393 | 92,500 | 28.7 % |
| Python | 2,306 | 97,087 | 27.7 % |
| SAS | 1,568 | 90,000 | 18.8 % |
| Power BI | 1,401 | 90,000 | 16.8 % |
| R | 1,337 | 90,000 | 16.0 % |
| Word | 869 | 81,682 | 10.4 % |
| PowerPoint | 805 | 85,000 | 9.7 % |
| Oracle | 566 | 96,924 | 6.8 % |

**Median salary by job title (USD, Country = United States vs. rest of world)**

| Job title | US | Non-US |
|---|---:|---:|
| Senior Data Scientist | 155,000 | 155,000 |
| Senior Data Engineer | 150,000 | 147,500 |
| Machine Learning Engineer | 150,000 | 101,029 |
| Data Scientist | 130,000 | 119,550 |
| Data Engineer | 125,000 | 123,500 |
| Software Engineer | 125,000 | 89,100 |
| Cloud Engineer | 115,000 | 89,100 |
| Senior Data Analyst | 110,000 | 111,175 |
| Data Analyst | 90,000 | 90,000 |
| Business Analyst | 90,000 | 75,000 |

---

## 🚀 Getting Started

### Prerequisites

- **Microsoft Excel for Windows** (Microsoft 365 or a version that includes Power Query and Power Pivot / the Data Model). The PivotTables, slicers and charts rely on the Excel Data Model.

### Installation

1. **Clone** the repository (or download the `.xlsx` from the GitHub page):
   ```bash
   git clone https://github.com/Akrambechat/Excel_Project-Data_Analytics.git
   cd Excel_Project-Data_Analytics
   ```
2. **Open** `Data_Analytics_Project.xlsx` in Excel. If Excel opens it in *Protected View*, click **Enable Editing**.
3. The workbook ships with its data model and cached results, so **no refresh is needed to explore it**.

### Refreshing from the source data (optional)

The Power Query queries read a source file named `data_jobs_salary_all.xlsx` (sheet `Sheet1`). That file is **not included** in this repository, and the query still points to a path on the author's machine. To refresh the data:

1. Place your copy of `data_jobs_salary_all.xlsx` somewhere on your computer.
2. Go to **Data → Queries & Connections**, right-click **data_jobs_salary → Edit**, and change the path in the `Source` step.
3. Click **Close & Load**, then **Data → Refresh All**.

---

## 🧭 Usage / How to Run

1. **Browse the tabs** in order: `Skill_Analysis` → `Skill_Job_Analysis` → `Salary_vs_skills` → `Salary_analysis` → `Skill salary analysis`.
2. **Filter with the slicers.** On `Skill_Job_Analysis` and `Skill salary analysis`, pick a **Job Title** and a **Country** (for example *Data Analyst* + *United States*) to update both the PivotTable and its chart. The Country slicer is shared by every sheet, so selecting a country updates all of them.
3. **Inspect the model:** **Data → Manage Data Model → Diagram View** shows the `job_id` relationship; the measure grid shows the measures listed above.
4. **Inspect the queries:** **Data → Queries & Connections**, then open `data_jobs_salary` or `data_jobs_skills` to step through each Power Query transformation.

---

## 📄 License

Distributed under the **MIT License**. See [`LICENSE`](LICENSE) for details.

## 👤 Author

**Akram** · GitHub: [@Akrambechat](https://github.com/Akrambechat)
