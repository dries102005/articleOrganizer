# articleOrganizer

### RIS → Excel tool for organizing academic search results

A lightweight CLI tool to organize **RIS exports** from academic databases
(ScienceDirect, Web of Science, IEEE Xplore, SpringerLink, Wiley, ACM, PubMed, and more)
into **clean, grouped Excel files**.

This project:

* does **not** scrape websites
* does **not** require API keys
* works entirely with **manually exported RIS files**

You search on the database website, export RIS files, and this script organizes everything into structured Excel outputs.

---

## What this tool generates

For **each source folder** (for example `wos_exports/` or `ieee_exports/`), the script creates:

```
outputs/results_<folder_name>_grouped.xlsx
```

Each Excel file contains two sheets:

* **`grouped_results`**
  Human-friendly, grouped layout (for reading, reporting, and reviews)

* **`raw_results`**
  Flat table with all processed records (useful for analysis and debugging)

### Columns

* Title
* Year
* Index (source name, by UR domain or DOI)
* DOI
* Author Name (first author)
* Author Surname (first author)

---

## Setup

### Requirements

* Python 3.9+
* pandas
* openpyxl

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## How to Use

Run the script:

```bash
python ris_to_excel.py
```

When the program starts, it will ask interactively:

1. **Year filter**
   Example:

   ```
   2007-2017
   ```

   (leave empty for no year filtering)

2. **ONLY filter (keywords)**
   Example:

   ```
   MTHFR, HPV
   ```

   Keeps only articles that contain **all** keywords
   (leave empty for no keyword filtering)

3. **Duplicate handling**

   ```
   Write only ONE from duplicate articles? (y/n)
   ```

  * Uses **DOI** first
  * Falls back to `(Title + Year + Author)` if DOI is missing

4. **Source folders**

  * Type `all` to scan all subfolders
  * Or provide a comma-separated list:

    ```
    science_direct_exports, wos_exports, ieee_exports
    ```

---

## File naming & grouping behavior (IMPORTANT)

This tool supports **two grouping modes**, depending on how you name your RIS files.

---

### ✅ If you name your RIS files like this

**Example (ScienceDirect, Web of Science, IEEE Xplore, etc.):**

```text
1.(MTHFR and HPV)_0-100.ris
1.(MTHFR and HPV)_100-200.ris
2.(MTHFR and human papillomavirus)_0-100.ris
```
You may also use a simplified form if you export all results in a single file:

```text
1.(MTHFR and HPV)_all.ris
2.(Virus)_all.ris
3.(Cancer)_all.ris
```

The part after the underscore (`_`) can be **anything** (`0-100`, `all`, `results`, etc.).
It exists only to separate the **query name** from the rest of the filename.


#### Then the output will look like this:

```text
1.(MTHFR and HPV)
  Article A
  Article B
  Article C

2.(MTHFR and human papillomavirus)
  Article D
  Article E
```

This means:

* All files starting with the same prefix (e.g. `1.(...)_`) are **grouped together**
* A **header row** (`1.(...)`) is automatically inserted
* A **blank line** separates each query group

✅ This mode is ideal for:

* Systematic reviews
* Search strategy documentation
* PRISMA-style reporting
* Clearly separating results from different search queries

---

### ❗ If you do NOT name your files like that

**Example:**

```text
search_results_1.ris
export_from_ieee.ris
my_results.ris
```

#### The script will still work.

In this case, the output will be grouped like:

```text
FILE: search_results_1.ris
  Articles from that file

FILE: export_from_ieee.ris
  Articles from that file
```

This means:

* Each RIS file becomes its **own group**
* The filename itself is used as the group header
* No data is lost — only the grouping logic changes

This mode is useful when:

* You do not want to rename exported files
* You export results incrementally
* You just want structured Excel output without query-based grouping

---

### Summary

| File naming style   | Grouping behavior |
| ------------------- | ----------------- |
| `N.(query)_X-Y.ris` | Grouped by query  |
| `N.(query)_all.ris` | Grouped by query  |
| Any other filename  | Grouped by file   |

Both modes are fully supported — **no configuration changes required**.