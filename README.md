# 📊 PR. 1 Fundamental Booster — Microsoft Excel Formula & Data Analytics Project

> **A complete, multi-sheet Microsoft Excel data engineering and formula mastery project demonstrating dynamic lookups, conditional aggregations, logical classifiers, text transformations, date/time analytics, financial rounding, and dynamic array filtering across academic, sales, and HR datasets.**

## 🌟 Project Overview & Executive Summary

**PR. 1 Fundamental Booster** is a practical spreadsheet modeling and data analysis project built in **Microsoft Excel (`.xlsx`)**. The workbook processes three distinct real-world operational datasets—**Students Grade** (`20` student records), **Sales Data** (`20` regional transactions), and **Employee Data** (`20` cross-departmental HR records)—by applying core to advanced Excel functions without relying on manual or hardcoded calculations.

Every worksheet is structured with custom number masks, conditional logic, and dedicated lookup blocks so that any change in underlying raw scores, transaction amounts, or employee salaries automatically updates all dependent metrics.

| Key Project Dimension | Implementation Specification |
| :--- | :--- |
| **Author** | **PRIYAL PATEL** ([@priyal-24](https://github.com/priyal-24)) |
| **GitHub Repository** | [https://github.com/priyal-24/Fundamental-Booster](https://github.com/priyal-24/Fundamental-Booster) |
| **Deliverable Workbook** | `PR. 1 Fundamental Booster.xlsx` |
| **Total Worksheets** | `4` (`Project Instructions`, `Students Grade`, `Sales Data`, `Employee Data`) |
| **Core Formula Families** | Logical (`IF`, Nested `IF`, `AND`, `OR`), Conditional Math (`COUNTIFS`, `SUMIFS`, `AVERAGEIFS`), Lookups (`VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX` + `MATCH`), Dynamic Arrays (`FILTER`, `INDIRECT`, `OFFSET`), Text (`LEFT`, `FIND`, `UPPER`, `LOWER`), Date (`DATEDIF`, `YEAR`), Financial Rounding (`ROUND`, `CEILING`, `FLOOR`) |
| **Visual & Typography Theme** | Corporate Navy (`#1F4E78`) & Teal (`#2F5597`) headers, Soft Zebra striping (`#F2F6FC`), Sage Green result highlights (`#E2EFDA`), and frozen top headers (`A2`) |
| **Custom Number Formatting** | `"STU-"0000` (Student IDs), `"SAL-"0000` (Sales IDs), `"EMP-"0000` (Employee IDs), `$#,##0.00` (Currency), `dd-mmm-yyyy` (Standardized Dates) |

## 🗂️ Workbook Architecture

The workbook is organized into four worksheets, progressing from the project specification to three domain-specific analytical sheets:

1. **📋 `Project Instructions`:** Contains the official project brief, core topics, main tasks, and essential task guidelines formatted with clean section headers and borders.
2. **🎓 `Students Grade`:** Evaluates student performance across **Math**, **Science**, and **English**, automates letter grading (`A`–`D`), flags dual-subject high achievers, normalizes student name strings, executes multi-criteria class counts/averages, and extracts a live 10-student honor roll using `FILTER`.
3. **💼 `Sales Data`:** Tracks regional hardware sales across **North**, **South**, **East**, and **West**, computes tiered promotional discounts using absolute references, checks discount eligibility, applies financial rounding rules, and runs cross-table lookups via `VLOOKUP`, `SUMIFS`, `INDEX/MATCH`, `XLOOKUP`, `XMATCH`, and `OFFSET`.
4. **🏢 `Employee Data`:** Manages departmental payroll across **Finance**, **HR**, **IT**, and **Marketing**, calculates exact employee age (`DATEDIF`), joining year (`YEAR`), and service tenure in days, and retrieves unsorted employee compensation via `XLOOKUP` and `INDEX/MATCH`.

## 🎯 Main Tasks & Essential Tasks Traceability Matrix

Every requirement from the project specification is mapped directly to its target worksheet, exact cell range, live formula syntax, and verified output below:

| # | Task Category | Project Requirement | Target Sheet | Cell Range | Formula Applied | Verified Output |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **Main Task 1** | Apply `IF` & Nested `IF` formulas to classify student grades | `Students Grade` | `I2:I21` | `=IF(H2>=80,"A",IF(H2>=70,"B",IF(H2>=60,"C","D")))` | Grades `A`, `B`, `C`, `D` assigned dynamically |
| **2** | **Main Task 2** | Use `COUNTIFS` to count students scoring above 60 | `Students Grade` | `P2` | `=COUNTIFS(H2:H21,">60")` | **`16` Students** |
| **3** | **Main Task 3** | Use `VLOOKUP` to find a product price from sales data | `Sales Data` | `N4` | `=VLOOKUP(2005,$A$2:$E$21,5,FALSE)` | **`$43,388.00`** (`SAL-2005`) |
| **4** | **Main Task 4** | Apply `XLOOKUP` to fetch employee salaries dynamically | `Employee Data` | `L3:L4` | `=XLOOKUP("Employee 8",B2:B21,D2:D21)` | **`$91,659.00`** (`EMP-3008`) & **`$114,173.00`** (`EMP-3016`) |
| **5** | **Main Task 5** | Format and analyze date of joining in Employee data | `Employee Data` | `E2:I21` | `=DATEDIF(F2,$L$2,"Y")`, `=YEAR(E2)`, `=$L$2-E2` | Custom `dd-mmm-yyyy` dates, Age, Year & Tenure Days |
| **6** | **Main Task 6** | Use `FILTER` function to extract students with scores > 80 | `Students Grade` | `R2:V11` | `{=FILTER(A2:E21,C2:C21>80)}` | **10 Top-Performing Students** extracted across 5 columns |
| **7** | **Essential Task** | Relative (`A1`) & Absolute (`$A$1`) references + Custom Formats | All Data Sheets | `G2:G21`, `O5:P6` | `=IF(E2>=30000,E2*$N$2,IF(E2>=15000,E2*$N$3,0))` | Relative row formulas + locked `$N$2`, `$N$3`, `$L$2` parameters |
| **8** | **Essential Task** | Identify students scoring > 80 in both Math & Science (`IF` + `AND`) | `Students Grade` | `J2:J21` | `=IF(AND(C2>80,D2>80),"Yes","No")` | **`Yes`** for `STU-1012`, `STU-1013`, `STU-1020` |
| **9** | **Essential Task** | Determine product discount eligibility (`IF` + `OR`) | `Sales Data` | `H2:H21` | `=IF(OR(E2>=20000,B2="Laptop"),"Yes","No")` | **`Yes`** for high-value orders (`>= $20k`) or `Laptop` |
| **10** | **Essential Task** | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` conditional analytics | `Students Grade` & `Sales Data` | `P3:P4`, `N5` | `=COUNTIFS(C2:C21,">50")`, `=AVERAGEIFS(H2:H21,H2:H21,">60")`, `=SUMIFS(...)` | Math > 50: **`19`**, Avg > 60: **`73.63`**, East Keyboard: **`$80,349.00`** |
| **11** | **Essential Task** | Fetch student name based on ID using `VLOOKUP` | `Students Grade` | `P6` | `=VLOOKUP(P5,$A$2:$B$21,2,FALSE)` | ID `STU-1005` -> **`Student 5`** |
| **12** | **Essential Task** | `INDEX` & `MATCH` dynamic searches without column limits | `Sales Data` & `Employee Data` | `N6`, `L5` | `=INDEX(E2:E21,MATCH("Person 8",D2:D21,0))` | `Person 8`: **`$46,429.00`**, `EMP-3010` Dept: **`HR`** |
| **13** | **Essential Task** | Excel `TEXT` functions (`LEFT`, `FIND`, `UPPER`, `LOWER`) | `Students Grade` | `K2:M21` | `=LEFT(B2,FIND(" ",B2)-1)`, `=UPPER(B2)`, `=LOWER(B2)` | `"Student"`, `"STUDENT 1"`, `"student 1"` |
| **14** | **Essential Task** | `XLOOKUP` & `XMATCH` modern lookup operations | `Sales Data` | `N7:N8` | `=XLOOKUP("Person 12",D2:D21,E2:E21)`, `=XMATCH("Laptop",B2:B21)` | `Person 12`: **`$35,611.00`**, `Laptop` 1st Row: **`13`** |
| **15** | **Essential Task** | `INDIRECT` & `OFFSET` dynamic range references | `Students Grade` & `Sales Data` | `P7`, `N9` | `=SUM(INDIRECT("C2:C21"))`, `=SUM(OFFSET(E2,0,0,5,1))` | Math Total: **`1,604`**, First 5 Sales Sum: **`$96,800.00`** |
| **16** | **Essential Task** | Financial Math functions (`ROUND`, `CEILING`, `FLOOR`) | `Sales Data` | `I2:K21` | `=ROUND(E2,-2)`, `=CEILING(E2,100)`, `=FLOOR(E2,100)` | Standardized hundred-dollar financial rounding |

## 🔍 Deep-Dive: Worksheet-by-Worksheet Breakdown

### 1️⃣ Worksheet: `Students Grade` (Academic Performance & Dynamic Filtering)

This worksheet analyzes **20 student records** (`STU-1001` to `STU-1020`) enrolled between `01-Jan-2018` and `14-May-2027`.

#### A. Main Student Table (`Columns A:M`, `Rows 1:21`)

* **Primary Columns (`A:F`):**
  * `Column A` (**Student ID**): Formatted with custom mask `"STU-"0000` (`STU-1001` through `STU-1020`) while preserving numeric underlying IDs for `VLOOKUP`.
  * `Column B` (**Name**): Bold text (`Student 1` through `Student 20`).
  * `Columns C:E` (**Math, Science, English**): Raw subject scores formatted as integers (`#,##0`).
  * `Column F` (**Enrollment Date**): Standardized date format (`dd-mmm-yyyy`).
* **Calculated Columns (`G:M`):**
  * `Column G` (**Total**): `=SUM(C2:E2)` — Computes the 3-subject aggregate score using relative references.
  * `Column H` (**Average**): `=AVERAGE(C2:E2)` — Computes the mean score formatted to two decimal places (`0.00`).
  * `Column I` (**Grade (IF)**): Classifies students into 4 performance tiers using a Nested `IF` formula:
    ```excel
    =IF(H2>=80,"A",IF(H2>=70,"B",IF(H2>=60,"C","D")))
    ```
  * `Column J` (**Math & Sci > 80**): Uses `IF` combined with `AND` to identify STEM top performers:
    ```excel
    =IF(AND(C2>80,D2>80),"Yes","No")
    ```
    *Three students qualify with `"Yes"`: **Student 12** (Math `85`, Science `89`), **Student 13** (Math `98`, Science `90`), and **Student 20** (Math `83`, Science `92`).*
  * `Column K` (**First Name**): Dynamically locates the space delimiter and extracts the first word:
    ```excel
    =LEFT(B2,FIND(" ",B2)-1)
    ```
  * `Columns L & M` (**UPPER Name & LOWER Name**): Applies `=UPPER(B2)` and `=LOWER(B2)` for case standardization.

#### B. Summary & Lookup Panel (`Columns O:P`, `Rows 1:7`)

| Cell | Task / Function Label | Excel Formula | Evaluated Output |
| :--- | :--- | :--- | :--- |
| **`P2`** | Students Scoring > 60 (`COUNTIFS`) | `=COUNTIFS(H2:H21,">60")` | **`16`** |
| **`P3`** | Students > 50 in Math (`COUNTIFS`) | `=COUNTIFS(C2:C21,">50")` | **`19`** |
| **`P4`** | Average Score > 60 (`AVERAGEIFS`) | `=AVERAGEIFS(H2:H21,H2:H21,">60")` | **`73.63`** |
| **`P5`** | Lookup Student ID (Input Parameter) | `1005` (Displayed as `STU-1005`) | **`STU-1005`** |
| **`P6`** | Student Name (`VLOOKUP` with Absolute Ref) | `=VLOOKUP(P5,$A$2:$B$21,2,FALSE)` | **`Student 5`** |
| **`P7`** | Total Math (`INDIRECT`) | `=SUM(INDIRECT("C2:C21"))` | **`1,604`** |

#### C. Extracted Top-Performing Students Table (`Columns R:V`, `Rows 1:11`)

Using the array formula `{=FILTER(A2:E21,C2:C21>80)}` across `R2:V11`, the worksheet automatically extracts all **10 students** who scored above `80` in Mathematics:

| Row | Student ID (`Col R`) | Name (`Col S`) | Math (`Col T`) | Science (`Col U`) | English (`Col V`) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Row 2** | `STU-1003` | **Student 3** | `97` | `60` | `51` |
| **Row 3** | `STU-1008` | **Student 8** | `94` | `67` | `64` |
| **Row 4** | `STU-1010` | **Student 10** | `98` | `54` | `58` |
| **Row 5** | `STU-1012` | **Student 12** | `85` | `89` | `79` |
| **Row 6** | `STU-1013` | **Student 13** | `98` | `90` | `97` |
| **Row 7** | `STU-1014` | **Student 14** | `93` | `40` | `44` |
| **Row 8** | `STU-1016` | **Student 16** | `98` | `77` | `60` |
| **Row 9** | `STU-1018` | **Student 18** | `82` | `48` | `79` |
| **Row 10** | `STU-1019` | **Student 19** | `90` | `78` | `60` |
| **Row 11** | `STU-1020` | **Student 20** | `83` | `92` | `55` |

> **Tip — Why Array `FILTER` (`R2:V11`) Works Seamlessly:** The `FILTER` formula in `R2:V11` is configured as a multi-cell array range with pre-cached OpenXML values, ensuring that Excel never inserts the `@` implicit intersection operator and always displays all 10 rows and 5 columns immediately upon opening.

### 2️⃣ Worksheet: `Sales Data` (Commercial Discounts, Rounding & Advanced Lookups)

This worksheet tracks **20 hardware sales transactions** (`SAL-2001` to `SAL-2020`) across five products (**Keyboard**, **Printer**, **Mouse**, **Monitor**, **Laptop**) and four regions (**South**, **West**, **East**, **North**), totaling **$507,725.00** in gross revenue.

#### A. Main Sales Transactions Table (`Columns A:K`, `Rows 1:21`)

* **Primary Columns (`A:F`):**
  * `Column A` (**Sales ID**): Custom number format `"SAL-"0000` (`SAL-2001` through `SAL-2020`).
  * `Columns B & C` (**Product & Region**): Left-aligned categorical labels.
  * `Column D` (**Salesperson**): Italicized salesperson identifiers (`Person 1` through `Person 20`).
  * `Column E` (**Amount**): Formatted in US Currency (`$#,##0.00`).
  * `Column F` (**Date**): Formatted as `dd-mmm-yyyy` (`01-Jan-2023` through `13-Oct-2023`).
* **Calculated Columns (`G:K`):**
  * `Column G` (**Discount (IF)**): Combines a Nested `IF` with **Absolute Cell References** (`$N$2` = `10%` High Discount Rate for orders `>= $30,000`; `$N$3` = `5%` Low Discount Rate for orders `>= $15,000`; `0` for orders `< $15,000`):
    ```excel
    =IF(E2>=30000,E2*$N$2,IF(E2>=15000,E2*$N$3,0))
    ```
  * `Column H` (**Discount Eligible (OR)**): Flags transactions that either exceed `$20,000` OR involve a promotional `Laptop` unit:
    ```excel
    =IF(OR(E2>=20000,B2="Laptop"),"Yes","No")
    ```
  * `Columns I, J & K` (**ROUND, CEILING, FLOOR**): Applies financial rounding to the nearest `$100` increment:
    * `Column I` (**ROUND**): `=ROUND(E2,-2)` — Rounds to the nearest hundred dollars.
    * `Column J` (**CEILING**): `=CEILING(E2,100)` — Rounds **up** to the next multiple of `$100`.
    * `Column K` (**FLOOR**): `=FLOOR(E2,100)` — Rounds **down** to the previous multiple of `$100`.

#### B. Sales Lookup & Trend Analysis Panel (`Columns M:N`, `Rows 1:9`)

| Cell | Task / Function Label | Excel Formula | Evaluated Output |
| :--- | :--- | :--- | :--- |
| **`N2`** | High Discount Rate (Abs Ref `$N$2`) | `0.10` (`10%`) | **`10%`** |
| **`N3`** | Low Discount Rate (Abs Ref `$N$3`) | `0.05` (`5%`) | **`5%`** |
| **`N4`** | Product Price (`VLOOKUP` ID `2005`) | `=VLOOKUP(2005,$A$2:$E$21,5,FALSE)` | **`$43,388.00`** |
| **`N5`** | Total Sales East Keyboard (`SUMIFS`) | `=SUMIFS(E2:E21,C2:C21,"East",B2:B21,"Keyboard")` | **`$80,349.00`** |
| **`N6`** | Sales Value Person 8 (`INDEX/MATCH`) | `=INDEX(E2:E21,MATCH("Person 8",D2:D21,0))` | **`$46,429.00`** |
| **`N7`** | Salesperson Performance (`XLOOKUP`) | `=XLOOKUP("Person 12",D2:D21,E2:E21)` | **`$35,611.00`** |
| **`N8`** | Product Position Laptop (`XMATCH`) | `=XMATCH("Laptop",B2:B21)` | **`13`** |
| **`N9`** | Sales Trend First 5 (`OFFSET`) | `=SUM(OFFSET(E2,0,0,5,1))` | **`$96,800.00`** |

### 3️⃣ Worksheet: `Employee Data` (HR Tenure Analytics & Dynamic Salary Lookups)

This worksheet manages **20 employee records** (`EMP-3001` to `EMP-3020`) across **Finance**, **HR**, **IT**, and **Marketing**, with a combined base payroll of **$1,586,798.00**.

#### A. Main Employee Table (`Columns A:I`, `Rows 1:21`)

* **Primary Columns (`A:E`):**
  * `Column A` (**Employee ID**): Custom number format `"EMP-"0000` (`EMP-3001` through `EMP-3020`).
  * `Column B` (**Name**): Bold employee names (`Employee 1` through `Employee 20`).
  * `Column C` (**Department**): Italicized department names (`Finance`, `HR`, `IT`, `Marketing`).
  * `Column D` (**Salary**): Currency format (`$#,##0.00`), ranging from `$31,943.00` to `$116,284.00`.
  * `Column E` (**Joining Date**): Standardized date format (`dd-mmm-yyyy`), spanning `01-Jan-2015` to `03-Jan-2028`.
* **Date & Time Analytical Columns (`F:I`):**
  * `Column F` (**Date of Birth**): Employee birth dates formatted as `dd-mmm-yyyy`.
  * `Column G` (**Age**): Calculates exact completed years of age against the Analysis Reference Date (`$L$2` = `31-Dec-2028`) using `DATEDIF`:
    ```excel
    =DATEDIF(F2,$L$2,"Y")
    ```
  * `Column H` (**Joining Year**): Extracts the 4-digit calendar year from the Joining Date:
    ```excel
    =YEAR(E2)
    ```
  * `Column I` (**Days Since Joining**): Computes exact service tenure in days between the Analysis Reference Date (`$L$2`) and each employee's Joining Date:
    ```excel
    =$L$2-E2
    ```

#### B. HR Dynamic Lookup Panel (`Columns K:L`, `Rows 1:5`)

| Cell | Task / Lookup Label | Excel Formula | Evaluated Output |
| :--- | :--- | :--- | :--- |
| **`L2`** | Analysis Date (Abs Ref `$L$2`) | `2028-12-31` | **`31-Dec-2028`** |
| **`L3`** | Employee 8 Salary (`XLOOKUP` by Name) | `=XLOOKUP("Employee 8",B2:B21,D2:D21)` | **`$91,659.00`** |
| **`L4`** | Employee 3016 Salary (`XLOOKUP` by ID) | `=XLOOKUP(3016,A2:A21,D2:D21)` | **`$114,173.00`** |
| **`L5`** | Employee 3010 Dept (`INDEX/MATCH`) | `=INDEX(C2:C21,MATCH(3010,A2:A21,0))` | **`HR`** |

## 🧠 Quick Viva / Presentation Cheat Sheet

If your instructor asks you to explain how specific formulas work during project evaluation, use these concise points:

* **Why use `$` in `$A$2:$B$21` or `$N$2`?**
  > The dollar sign (`$`) creates an **Absolute Reference**. When we drag the discount formula `=IF(E2>=30000,E2*$N$2,...)` down from Row 2 to Row 21, `E2` changes relatively (`E3`, `E4`, ...), but `$N$2` stays locked on the `10%` discount cell.
* **Why is `XLOOKUP` better than `VLOOKUP`?**
  > `VLOOKUP` can only search in the leftmost column and requires a column index number that breaks if columns are inserted. `XLOOKUP(lookup_value, lookup_array, return_array)` searches any column (left or right) and works accurately on unsorted data by default.
* **How does `INDEX` + `MATCH` work together?**
  > `MATCH("Person 8", D2:D21, 0)` finds the exact row position (`8`) where `"Person 8"` appears, and `INDEX(E2:E21, 8)` returns the sales amount (`$46,429.00`) from that exact row.
* **What is the difference between `ROUND`, `CEILING`, and `FLOOR`?**
  > For `SAL-2001` (`$7,048.00`): `ROUND(7048, -2)` rounds to the nearest hundred (`$7,000`), `CEILING(7048, 100)` always rounds **up** to the next hundred (`$7,100`), and `FLOOR(7048, 100)` always rounds **down** (`$7,000`).
* **How do Custom ID Formats (`"STU-"0000`) work with `VLOOKUP`?**
  > Custom number formatting changes only how the number is **displayed** (`STU-1005`), while Excel still stores the pure integer `1005` inside the cell—allowing numeric lookups and sorting to work without type mismatch errors.

## 🚀 How to Clone, Open & Verify the Workbook

1. Clone this repository or download **`PR. 1 Fundamental Booster.xlsx`** directly from GitHub:
   ```bash
   git clone https://github.com/priyal-24/Fundamental-Booster.git
   ```
2. Open **`PR. 1 Fundamental Booster.xlsx`** in **Microsoft Excel** (Excel 365 / 2021 recommended for dynamic array formulas) or **Google Sheets**.
3. If prompted with a yellow security bar at the top, click **Enable Editing**.
4. Click on any green-highlighted result cell (`Col I` or `P2:P7` in `Students Grade`, `Col H` or `N4:N9` in `Sales Data`, `Col I` or `L3:L5` in `Employee Data`) to inspect the live formula in the Excel Formula Bar (`fx`).
5. Try changing any student's score in `C2:E21` or any transaction amount in `E2:E21` to watch the grades, discounts, `COUNTIFS`, `SUMIFS`, and `FILTER` tables recalculate automatically.

## 👩‍💻 Author & Repository Details

| Field | Details |
| :--- | :--- |
| **Created & Submitted By** | **PRIYAL PATEL** |
| **GitHub Profile** | [@priyal-24](https://github.com/priyal-24) |
| **Project Repository Link** | [https://github.com/priyal-24/Fundamental-Booster](https://github.com/priyal-24/Fundamental-Booster) |
| **Project Title** | **PR. 1 Fundamental Booster — Excel Data Analytics & Formula Mastery** |