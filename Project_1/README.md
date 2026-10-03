# Tech & Data Roles Job Market Dashboard

![Dashboard Preview](Project_1/Assets/Dashboard_Final_Preview.gif)

## Introduction

This dashboard was built to give job seekers in the tech and data space a clearer picture of what to expect — average compensation, where the opportunities are concentrated, and which platforms are actively hiring for each role. Rather than scrolling through scattered job listings, users can filter by job title, country, and employment type to instantly see relevant salary and hiring trends.

The underlying dataset covers real-world job postings across 10 tech and data roles, including Data Analyst, Data Scientist, Data Engineer, Machine Learning Engineer, Cloud Engineer, and their senior-level counterparts.

### Dashboard File
The finished dashboard is available here: [Tech_Data_Roles_Dashboard.xlsx](Tech_Data_Roles_Dashboard.xlsx).

### Excel Skills Used

- **📉 Charts**
- **🧮 Formulas and Functions**
- **❎ Data Validation**

### Dataset Overview

The dataset contains job posting records with details on:

- **👨‍💼 Job Titles** — 10 roles spanning analyst, engineer, and scientist tracks
- **💰 Salaries**
- **📍 Job Locations** (country-level)
- **🧩 Hiring Platforms** (e.g., Indeed, LinkedIn, etc.)
- **⏰ Schedule Types** (Full-time, Part-time, Contractor, Temp work, Internship)

## Dashboard Build

### 📉 Charts

#### 📊 Average Salary by Role — Bar Chart

<img src="Project_1/Assets/Chart1_Salary_By_Role.png" width="850" height="550" alt="Average Salary by Role">

- 🛠️ **Excel Features:** Horizontal bar chart with `$0K` formatted axis labels for clean readability on large salary figures.
- 🎨 **Design Choice:** Sorted descending by salary, with the currently selected role highlighted in a darker shade to stand out against the rest.
- 💡 **Insight:** Senior and engineering-focused roles (Senior Data Scientist, ML Engineer, Senior Data Engineer) consistently show higher average salaries than analyst-track positions.

#### 🗺️ Average Salary by Country — Map Chart

![Country Map](Project_1/Assets/Chart2_Country_Map.gif)

- 🛠️ **Excel Features:** Map chart with color gradient reflecting average salary per country.
- 🎨 **Design Choice:** Darker shading indicates higher average salary, independent of the Country dropdown — giving a global comparison view regardless of which country is currently selected for the other KPIs.
- 💡 **Insight:** Makes it easy to spot which regions tend to offer stronger compensation for the same role, useful for anyone weighing remote or relocation options.

#### 🧩 Schedule Type Breakdown — Bar Chart

<img src="Project_1/Assets/Chart3_Schedule_Type.png" width="850" height="550" alt="Job Count by Schedule Type">

- 🛠️ **Excel Features:** Horizontal bar chart comparing job counts across schedule types (Full-time, Part-time, Contractor, Temp work, Internship).
- 🎨 **Design Choice:** The currently selected schedule type is highlighted to stay consistent with the Job Title chart's visual language.
- 💡 **Insight:** Full-time postings dominate across most roles, but the gap narrows for certain positions where contract and temp work are more common.

### 🧮 Formulas and Functions

#### 💰 Average Salary by Job Title, Country & Schedule Type

```excel
=AVERAGEIFS(
    jobs[salary_year_avg], jobs[job_country], A2,
    jobs[salary_year_avg], "<>0",
    jobs[job_title_short], title,
    jobs[job_schedule_type], type
)
```

- 🔍 **Multi-Criteria Filtering:** Matches country, job title, and schedule type simultaneously, while excluding blank/zero salary entries.
- 🎯 **Why Average, not Median:** I opted for average salary to reflect overall compensation level across postings, giving a direct sense of typical market pay for the selected combination.
- **🔢 Formula Purpose:** Powers the main "Average Salary" KPI, updating dynamically as the Job Title, Country, and Type dropdowns change.

🍽️ Background Table

![Background Table](Project_1/Assets/Screenshot1_Background_Table.png)

📉 Dashboard Implementation

<img src="Project_1/Assets/Dashboard_Job_Title_Selector.png" width="400" height="500" alt="Job Title Selector">

---

#### 🧩 Top Job Platform

This KPI is built across four steps, combined into a single text box on the dashboard.

**Step 1 — Get the unique list of platforms:**
```excel
=UNIQUE(jobs[job_via])
```

**Step 2 — Count postings per platform, filtered by the selected job title, country, and type:**
```excel
=COUNTIFS(
    jobs[job_via], A2,
    jobs[job_title_short], title,
    jobs[job_country], country,
    jobs[job_schedule_type], type
)
```

**Step 3 — Sort platforms by count, highest first:**
```excel
=SORT(A2:B594, 2, -1)
```

**Step 4 — Clean up the platform label for display:**
```excel
=SUBSTITUTE(D2, "via ", "")
```

- 🔍 **Why four steps:** The raw `job_via` values come prefixed with "via " (e.g., "via Indeed"), so the final step strips that prefix for a cleaner label. Splitting the process this way also makes each stage easy to audit and debug independently.
- **🔢 Formula Purpose:** Together, these four formulas identify and display the platform with the highest number of matching postings as a shape/text box on the dashboard.

---

#### 🔢 Total Job Postings

This KPI is also built across four steps.

**Step 1 — Get the unique list of job titles:**
```excel
=UNIQUE(jobs[job_title_short])
```

**Step 2 — Count postings matching the selected country and schedule type, for each job title:**
```excel
=COUNT(
  IF(
    (jobs[job_country]=country)*
    (jobs[job_title_short]=A2)*
    (ISNUMBER(SEARCH(type, jobs[job_schedule_type]))),
    jobs[salary_year_avg]
  )
)
```

**Step 3 — Sort by count, highest first:**
```excel
=SORT(A2:B11, 2, -1)
```

**Step 4 — Look up the count for the currently selected job title:**
```excel
=XLOOKUP(title, E2:E11, F2:F11, "No Result")
```

- 🔍 **Multi-Criteria Filtering:** Combines an exact match on country with a partial match (`SEARCH`) on schedule type, since the source data sometimes stores combined schedule types (e.g., "Full-time and Contractor") in a single cell.
- **🔢 Formula Purpose:** Displays the total number of job postings for the selected Job Title, Country, and Type combination.

🍽️ Background Table

![Count Table](Project_1/Assets/Screenshot2_Count_Table.png)

### ❎ Data Validation

#### 🔍 Dropdown Filters

- 🔒 **Controlled Input:** Applied filtered/unique lists as data validation rules for the `Job Title`, `Country`, and `Type` selectors, ensuring:
    - 🎯 Users can only select from valid, predefined options
    - 🚫 Typos or inconsistent entries are prevented from breaking the formulas
    - 👥 The dashboard stays reliable and easy to use for anyone exploring it

<img src="Project_1/Assets/Dashboard_Data_Validation.gif" width="425" height="400" alt="Data Validation Demo">

## Conclusion

This dashboard was built to turn raw job posting data into something genuinely actionable — a quick way to compare average salaries, see where demand is strongest geographically, and understand which platforms and schedule types are worth prioritizing during a job search. Beyond the end result, this project reinforced how much a well-structured Excel workbook (array formulas, `XLOOKUP`, dynamic filters, and clean validation) can achieve without needing a full BI tool.