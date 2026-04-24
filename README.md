# 📘 Professional Excel Training Course — Part 2
### Functions

> A complete hands-on Excel reference workbook covering **7 powerful Functions** for What-If Analysis, Data Validation, Smart Dropdowns, and Dynamic References — used daily in Finance, Accounting, and Data Analysis roles.

[![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?logo=microsoft-excel&logoColor=white)](https://github.com/heysubu)
[![Course](https://img.shields.io/badge/Course-Part%202%20of%203-blue)](https://github.com/heysubu)
[![Topics](https://img.shields.io/badge/Topics-7-orange)](https://github.com/heysubu)
[![Level](https://img.shields.io/badge/Level-Intermediate-green)](https://github.com/heysubu)
[![Sponsor](https://img.shields.io/badge/Sponsor-%E2%9D%A4-red)](https://github.com/sponsors/heysubu)

---

## 📚 Full Course Links

> This is **Part 2** of a 3-part Professional Excel Training Course. Jump to any part below:

| 📗 Part 1 | 📘 Part 2 — You Are Here | 📙 Part 3 |
|---|---|---|
| **Navigation & Formatting** | **Functions** | **Formulas** |
| 14 Topics | 7 Topics | 30+ Topics |
| [🔗 Open Part 1](https://github.com/heysubu/excel-course-part-1-navigation-formatting) | ✅ Current | *Coming Soon* |

---

## 📋 About This Course

**Part 2** dives into Excel's most powerful built-in functions used in real Finance and Analyst jobs. Each topic includes a real-world example, step-by-step instructions, and actual formulas you can copy and use immediately.

| 📦 Course Part | 📚 Category | 🔢 Topics |
|---|---|---|
| Part 1 | Navigation & Formatting | 14 Topics |
| **Part 2 — This File** | **Functions** | **7 Topics** |
| Part 3 | Formulas | 30+ Topics |

---

## 📥 Download the File

> 📎 **The Excel file is attached directly in this repository.**
>
> Click on `Excel_Course_Part_2-Functions.xlsx` above to download and open in Microsoft Excel.

**How to open:**
1. Click the `.xlsx` file in this repository
2. Click **Download**
3. Open in **Microsoft Excel** (2016 or later recommended)
4. Enable editing if prompted

---

## 🗺️ Quick Navigation — Click to Jump to Any Topic

> 💡 Click any topic name below and it jumps directly to that section!

| # | 🔗 Topic | ⚡ Quick Summary |
|---|---|---|
| 01 | [🎯 Goal Seek](#01--goal-seek) | Work backwards — find the input needed to hit a target |
| 02 | [✅ Data Validation](#02--data-validation) | Control what data users can enter into any cell |
| 03 | [🎭 Scenario Manager](#03--scenario-manager) | Compare multiple what-if scenarios side by side |
| 04 | [🏷️ Name Manager](#04--name-manager) | Give names to cells & ranges — use them in formulas |
| 05 | [🔍 Advanced Filter](#05--advanced-filter) | Filter large data tables using complex multi-criteria rules |
| 06 | [🔽 Dependable Drop Down Lists](#06--dependable-drop-down-lists) | Create cascading dropdowns where selection 2 depends on 1 |
| 07 | [🔗 INDIRECT Formula](#07--indirect-formula) | Pull data from different sheets using dynamic references |

---

---

## 01 · 🎯 Goal Seek

> **Work backwards — find the input value needed to reach a specific target output**

### What is Goal Seek?
Goal Seek answers the question: *"What number do I need to enter here to get THIS result?"*

Instead of guessing and testing, Goal Seek automatically calculates the exact input needed.

### When to Use It:
- Find what **conversion rate** is needed to hit a lead generation target
- Find what **cost per click** to bid on ads to be profitable
- Find what **GPA** a student needs in their next semester to maintain a target average
- Find what **sales volume** is needed to break even

### How to Run Goal Seek:
| Step | Action |
|---|---|
| 1 | Go to **Data tab** → **What-If Analysis** → **Goal Seek** |
| 2 | **Set Cell** → select the cell with the outcome/result |
| 3 | **To Value** → type the target number you want to achieve |
| 4 | **By Changing Cell** → select the input cell Excel should adjust |
| 5 | Click **OK** — Excel calculates the answer instantly |

### Real Example #1 — Marketing Channel Analysis:
ABC Inc. runs 4 marketing channels (Facebook, Affiliate, YouTube, Pinterest).
They want each channel to generate **2.85 leads per $100 spent**.

```
Channel       Cost/Click   Conversion   Leads/$100
Facebook      $1.00        3.0%         3.00
Affiliate     $0.50        ?            → Goal: 2.85
YouTube       $1.50        4.0%         2.67
Pinterest     $0.70        2.0%         2.85
```
**Goal Seek Result:** Affiliate needs a **1.425% conversion rate** to hit the 2.85 target.

### Real Example #2 — Student GPA Calculator:
A student has these semester GPAs and wants to maintain a **3.75 cumulative GPA**:

| Semester | GPA |
|---|---|
| 1st | 3.68 |
| 2nd | 3.81 |
| 3rd | 3.75 |
| 4th | 3.58 |
| **5th (Goal Seek)** | **→ Needs 3.93** |

**Goal Seek Result:** The student needs a **3.93 GPA** in the 5th semester to maintain 3.75 cumulative.

> 💡 **Pro Tip:** Goal Seek only works with **one variable**. If you need to change multiple variables at once, use Scenario Manager (Topic 03) instead.

---

---

## 02 · ✅ Data Validation

> **Control exactly what data a user can enter into any cell — no typos, no wrong values**

### What is Data Validation?
Data Validation lets you set rules for what is allowed in a cell. If someone tries to enter something outside your rules, Excel blocks it and shows an error message.

### When to Use It:
- Create a **picklist/dropdown** so users pick from a fixed list of options
- Restrict a cell to **only whole numbers** (e.g. 0 to 10)
- Restrict a cell to **only dates** within a range
- Share a workbook with others and ensure they fill it in correctly

---

### Example #1 — Creating a Picklist (Dropdown):

**Steps:**
1. Create your list of items somewhere in the workbook (e.g. account types)
2. Select the cell(s) where you want the dropdown to appear
3. Go to **Data tab** → **Data Validation**
4. Set **Allow** = `List`
5. In **Source**, click and select your list range
6. Click **OK** — a dropdown arrow now appears in the cell

**Sample Picklist Items:**
```
Entertainment
Printing & Stationery
Insurance Premium
Travelling Allowance
Miscellaneous Expenses
Repair & Maintenance
Air Conditioner
License Renewal Fee
Business Development Expense
Loan from Directors
Sales Revenue
Purchase
Advertisement
Furniture & Fixture
```

**Result — Cash Received & Payments Table (HAMMERHEAD USA):**

| SL | Date | Account | Description | Received | Payment | Balance |
|---|---|---|---|---|---|---|
| 01 | 01.01.2017 | Entertainment | Loan From Director | 1,410 | — | 1,410 |
| 02 | 10.01.2017 | Repair & maintenance | Entertainment for Meeting | — | 1,240 | 170 |
| 03 | 26.01.2017 | Entertainment | Stationery for Office | — | 170 | 0 |

---

### Example #2 — Whole Number Restriction (0 to 10):

**Use case:** Sending a questionnaire where answers must be between 0 and 10.

**Steps:**
1. Select the input cells
2. **Data** → **Data Validation** → Allow = `Whole Number`
3. Set Minimum = `0` and Maximum = `10`
4. Optional: Add an error message like *"Please enter a number between 0 and 10"*
5. Click **OK**

> ✅ If someone enters `11` or `-1` or text, Excel will block it and show your error message.

---

---

## 03 · 🎭 Scenario Manager

> **Build multiple what-if scenarios and compare them side by side in a summary table**

### What is Scenario Manager?
Scenario Manager lets you create named scenarios — each with different input values — and compare how they all affect one output. It is like taking multiple snapshots of your model under different assumptions.

### When to Use It:
- Compare **best case / base case / worst case** financial projections
- Compare **different car lease options** with different prices and rates
- Compare **budget scenarios** with different expense levels
- Present multiple options to management in a clean summary table

### How to Create Scenarios:

| Step | Action |
|---|---|
| 1 | **Data tab** → **What-If Analysis** → **Scenario Manager** |
| 2 | Click **Add** → type a name for Scenario 1 |
| 3 | In **Changing Cells**, select all input cells that vary |
| 4 | Enter the values for this scenario → Click **OK** |
| 5 | Repeat steps 2–4 for each additional scenario |
| 6 | To view: select a scenario → click **Show** |
| 7 | For summary: click **Summary** → select result cell → **OK** |

### Real Example — Car Lease Comparison:

Input cells that change per scenario: Car Model, Price, Down Payment, Yearly Rate, Loan Term

| Variable | Scenario 1: BMW 325i | Scenario 2: Lexus ES |
|---|---|---|
| Car Price | $50,000 | $40,000 |
| Down Payment | $3,000 | $4,500 |
| Yearly Rate | 3.0% | 2.5% |
| Loan Term | 2 years | 3 years |
| **Monthly Payment** | **-$2,299.49** | **Calculated** |

Excel creates a **side-by-side summary sheet** automatically showing all scenarios and results.

> 💡 **Pro Tip:** Scenario Manager works best when your input cells are clearly marked (e.g. highlighted in grey) so it is easy to identify what to put in the "changing cells" field.

---

---

## 04 · 🏷️ Name Manager

> **Give meaningful names to cells and ranges — then use those names in formulas instead of cell addresses**

### What is Name Manager?
Instead of writing `=SUM(D12:D20)` you can name that range "Expenses" and write `=SUM(Expenses)`. This makes formulas much easier to read and maintain.

### When to Use It:
- Large financial models with many formulas referencing the same ranges
- Multi-sheet workbooks where you want consistent, readable formula references
- Shared workbooks where others need to understand the formulas quickly
- Avoiding copy-paste errors (named cells are always absolute references)

### How to Create a Named Range:

| Step | Action |
|---|---|
| 1 | Go to **Formulas tab** → **Name Manager** |
| 2 | Click **New** |
| 3 | Enter a name — **no spaces allowed** (e.g. `January2019`) |
| 4 | In **Refers To**, select the cell or range |
| 5 | Click **OK** |

> ⚠️ Names cannot contain spaces. Use underscores or run words together (e.g. `TotalRevenue`, `January_Sales`)

### Real Example — Monthly Sales Summary:

**Raw Data:**

| Product | January | February | March |
|---|---|---|---|
| A0001 | 100 | 200 | 300 |
| B0002 | 400 | 500 | 600 |
| C0003 | 700 | 800 | 900 |
| D0004 | 800 | 700 | 600 |
| **Total** | **2,000** | **2,200** | **2,400** |

**After naming the totals:**
```excel
Named:   January = 2000    February = 2200    March = 2400

Formula (old way):   =SUM(D29, E29, F29)
Formula (with names): =SUM(January, February, March)     ← Much clearer! ✅
```

**Results using names:**
| Description | Result | Formula Used |
|---|---|---|
| Sales for January | 2,000 | `=January` |
| Sales for February | 2,200 | `=February` |
| Sales for March | 2,400 | `=March` |
| **Total Sales** | **6,600** | `=SUM(January, February, March)` |

---

---

## 05 · 🔍 Advanced Filter

> **Filter large data tables using complex multi-condition criteria that regular AutoFilter cannot handle**

### What is Advanced Filter?
The regular Excel filter is simple. Advanced Filter lets you define your own criteria table with multiple conditions — AND / OR logic — and copy the filtered results to a new location.

### When to Use It:
- Filter a sales table where **Total Kg > 100**
- Filter data by **multiple conditions on different columns** at once
- Extract filtered results to a **separate location** without changing the original data
- Run complex queries on large exported datasets

### Steps:

| Step | Action |
|---|---|
| 1 | Set up a **criteria table** — same headers as your data, with your filter conditions below |
| 2 | Select your entire **data table** |
| 3 | Go to **Data tab** → **Filter** → **Advanced** |
| 4 | Enter: **List Range** (your data), **Criteria Range** (your conditions), **Copy To** (output location) |
| 5 | Click **OK** |

### Real Example — Sales Report Filter (December 2019):

**Goal:** Find all products sold where **Total Kg > 100**

**Criteria Setup:**
```
Total Kg
>100
```

**Sample Results (products > 100 Kg):**

| Client | Product | Qty (PKT) | Total Kg | Total Value |
|---|---|---|---|---|
| Western Marin | Bohler Fox N EV-48, 3.15x450mm | 52 | 260 | 67,600 |
| Royal Technic JV | Bohler Fox EV 48-1, 3.2x350mm | 60 | 300 | 1,17,000 |
| NS Construction | Bohler Fox N EV-48, 3.15x450mm | 80 | 400 | 1,13,200 |
| INDWell | Bohler Fox N EV-48, 3.15x450mm | 100 | 500 | 1,45,000 |
| Royal-Pipeliners JV | Ultra CR1-3.2x450mm | 65 | 325 | 1,46,250 |

> 💡 **Pro Tip:** Advanced Filter is especially powerful when you need to find records that meet **multiple OR conditions** across different columns — something the regular filter menu cannot do in one step.

---

---

## 06 · 🔽 Dependable Drop Down Lists

> **Create cascading dropdowns where the second dropdown automatically changes based on what was selected in the first**

### What is a Dependable (Cascading) Dropdown?
A regular dropdown gives you one fixed list. A **Dependable** dropdown is smarter — when you pick from Dropdown 1, Dropdown 2 automatically shows only the options that belong to that selection.

### When to Use It:
- **Industry → Sub-Industry** (pick "Arts and Marketing" → see only "Broadcasting, Graphic Design, Photographer, Publisher")
- **Expense Category → Expense Item** (pick "Revenue" → see only revenue line items)
- **Sales Segment → Sales Rep** (pick segment → see only that segment's reps)
- Any time you have **categories and sub-categories**

### Example — Category & Sub-category Lists:

| Industry Category | Industry Subcategory |
|---|---|
| Arts and Marketing | Broadcasting, Graphic Design, Photographer, Publisher |
| Event Planning | DJ, Caterer, Entertainer, Wedding Planning, Wedding Venue |
| Business Services | Employment Agency, Recruiter, Consultants, Printing Service |

| Expense Category | Expense Subcategory |
|---|---|
| Cost of Revenue | Support, Technical Operations, Hosting |
| Sales and Marketing | Commissions, SEO, Advertising, Conferences, Payroll |
| Operating Expenses | Payroll, Legal, Rent |
| Other Expense | Interest Expense, Gain/Loss |

### Steps to Build Dependable Dropdowns:

| Step | Action |
|---|---|
| 1 | Create separate lists for each category and its items |
| 2 | Convert each list into a **Table** (Insert → Table) and give each table a meaningful name (e.g. `Revenue1`, `AdministrationCost1`) |
| 3 | Select all the **headers** together → go to **Formulas → Name Manager** → create a name for all headers (e.g. `accounttype1`) |
| 4 | Create individual names for each sub-list (e.g. name the Revenue column as `revenue1`) |
| 5 | In your main data table, select the **Category column** → Data Validation → List → Source = `=accounttype1` |
| 6 | In the **Sub-category column** → Data Validation → List → Source = `=INDIRECT($C2)` *(reference the category cell — do NOT lock the row number)* |
| 7 | Drag down the sub-category validation to all rows |

> ⚠️ **Important:** In the INDIRECT source formula, lock the **column** with `$` but NOT the row number. Example: `=INDIRECT($C142)` ✅ not `=INDIRECT($C$142)` ❌

---

---

## 07 · 🔗 INDIRECT Formula

> **Pull data from a different sheet dynamically — changing the sheet name updates all results automatically**

### What is INDIRECT?
INDIRECT builds a cell reference from text. Instead of hardcoding `=JAN!D11`, you can make the sheet name come from a dropdown cell — so changing the dropdown instantly shows that sheet's data.

### When to Use It:
- Build a **monthly P&L summary** where you select the month from a dropdown and all numbers update
- Lookup **MIN, MAX, AVERAGE** across city or branch worksheets using the city name as the reference
- Consolidate data from **consistently structured sheets** without writing separate formulas for each sheet

### Formula Syntax:
```excel
=INDIRECT("" & sheet & "!" & "reference")
```

| Part | Meaning |
|---|---|
| `sheet` | The cell containing the sheet name (comes from a dropdown) |
| `reference` | The cell address or range on that sheet to pull from |

---

### Example #1 — Dynamic Monthly P&L Summary:

**Setup:** 6 monthly sheets exist: JAN, FEB, MAR, APR, MAY, JUN

**Goal:** One summary tab with a dropdown to select the month — all P&L numbers update automatically.

```excel
Sheet dropdown cell: G38 (contains "FEB")

=VLOOKUP($D41, INDIRECT("" & $G$38 & "!$D$11:$G$36"), 4, FALSE)
↑ Looks up the account name from the selected month's sheet
```

**Result when "FEB" is selected:**

| Category | Amount |
|---|---|
| Service Fee | 3,64,331 |
| Electricity Bill Banglalion | 17,000 |
| **TOTAL RECEIPTS** | **3,82,301** |
| Staff Salary | 22,000 |
| Electricity Bill | 65,084 |
| **TOTAL PAYMENTS** | **2,90,455** |

> Changing the dropdown from "FEB" to "MAR" → all numbers update instantly to March data! ✅

---

### Example #2 — Dynamic City Temperature Lookup:

**Setup:** 4 city sheets exist: Beijing, Washington, Wellington, Canberra — each with daily temperature data.

**Goal:** Pick a city from a dropdown and see Min, Max, Average temperature.

```excel
City dropdown cell: E79 (contains "Washington")

=MIN(INDIRECT("" & $E$79 & "!" & "D12:D41"))       → 4°C
=MAX(INDIRECT("" & $E$79 & "!" & "D12:D41"))       → 21°C
=AVERAGE(INDIRECT("" & $E$79 & "!" & "D12:D41"))   → 8°C
```

> Changing dropdown from "Washington" to "Beijing" → temperatures update to Beijing data! ✅

> 💡 **Pro Tip:** INDIRECT is most powerful when all your sheets have the **same structure** — same headers, same row positions. Then you only need one formula that works for all sheets.

---

---

## 📚 Full Course Overview

| Part | Category | Topics | Link |
|---|---|---|---|
| ✅ **Part 1** | Navigation & Formatting | Navigation Shortcuts, Formatting Tips, Dates, Workday, Common Errors, Copy vs Cut, Precedents & Dependents, Group/Ungroup, Protect Ranges, Screenshot, Sorting, Remove Duplicates, Text to Columns, Conditional Formatting | [🔗 Open Part 1](https://github.com/heysubu/excel-course-part-1-navigation-formatting) |
| ✅ **Part 2** | Functions | Goal Seek, Data Validation, Scenario Manager, Name Manager, Advanced Filter, Dependable Dropdowns, INDIRECT | ✅ You are here |
| 📙 **Part 3** | Formulas | VLOOKUP, HLOOKUP, INDEX/MATCH, SUMIF/S, COUNTIF/S, IF, PMT, NPV, IRR, TRANSPOSE, TEXTJOIN, and 20+ more | *Coming Soon* |

---

## 📞 Contact

- 🐙 **GitHub:** [heysubu](https://github.com/heysubu)
- 📧 **Email:** suubhamghadge@gmail.com
- 💼 **LinkedIn:** [Subham Ghadge](https://www.linkedin.com/in/subhamghadge/)

---

## ❤️ Support This Work

If this course helped you, consider sponsoring to support more free Excel resources!

[![Sponsor](https://img.shields.io/badge/❤️%20Sponsor%20This%20Project-GitHub%20Sponsors-pink?logo=github)](https://github.com/sponsors/heysubu)

---

## 📄 License

MIT License — Free to use, share, and learn from.

---

## 🌟 Project Stats

![GitHub stars](https://img.shields.io/github/stars/heysubu/excel-course-part-2-functions?style=social)
![Excel](https://img.shields.io/badge/Excel-Functions-217346?logo=microsoft-excel)
![Sheets](https://img.shields.io/badge/Sheets-19-blue)
![Topics](https://img.shields.io/badge/Topics-7-orange)
