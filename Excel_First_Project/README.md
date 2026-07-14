# 📗 Excel Formulas Practice Project — Multi-Sheet Workbook

A hands-on Excel workbook demonstrating a wide range of formulas and functions — logical, lookup, statistical, text, date, and array formulas — applied across three real-world style datasets: Student Grades, Sales Data, and Employee Data.

---

## 📌 Project Overview

| Property | Details |
|---|---|
| **Tool** | Microsoft Excel |
| **File** | `Project-1.xlsx` |
| **Sheets** | Students Grade, Sales Data, Employees Data |
| **Key Formula Categories** | Logical, Lookup & Reference, Statistical, Text, Date & Time, Array/Dynamic Array |

---

## 📄 Sheet 1 — Students Grade

**Columns:** Student Id, Name, Maths, Science, English, Average Marks, Grade, High Score (>80), Enrollment Date

| Formula/Feature | Cell(s) | Purpose |
|---|---|---|
| `AVERAGE()` | F2:F21 | Calculates each student's average marks across Maths, Science, and English |
| Nested `IF()` | G2:G21 | Assigns a letter grade — **A** (≥80), **B** (≥70), **C** (≥60), or **Fail** — based on average marks |
| `AND()` + `IF()` | H2:H21 | Flags students who scored above 80 in **both** Maths and Science as high performers |
| `COUNTIF()` | K2 | Counts how many students scored above 50 in Maths |
| `AVERAGEIFS()` | K5 | Calculates the average score only for students whose average is above 60 |
| `VLOOKUP()` | L8 | Looks up a student's name by their Student ID |
| Dynamic Array `FILTER()` | K11 | Returns a filtered list of all top-performing students (average > 80), with a fallback message if none qualify |

---

## 📄 Sheet 2 — Sales Data

**Columns:** Sales Id, Product, Region, SalesPerson, Amount, Discount Eligible, Discount %, Final Amount, Date

| Formula/Feature | Cell(s) | Purpose |
|---|---|---|
| `OR()` + `IF()` | F2:F21 | Marks a sale as "Eligible" for discount if the amount exceeds 10,000 **or** the product is a Keyboard/Mouse |
| Nested `IF()` (tiered discount) | G2:G21 | Applies a tiered discount rate — 10% (>40,000), 5% (>20,000), 2% (>10,000), or 0% otherwise |
| Arithmetic formula | H2:H21 | Calculates the Final Amount after applying the discount percentage |
| `SUMIFS()` | L3 | Sums total sales for a specific Region + Product combination |
| `VLOOKUP()` | L6 | Retrieves a product's sale amount based on its Sales ID |
| Array formula with `INDEX()` + `MATCH()` (multi-condition) | L9 | Finds the sale value for a specific salesperson in a specific month, matching two conditions at once |
| `LEFT()` + `FIND()` | K12 | Extracts characters from a product name up to (but not including) the letter "b" |
| `UPPER()` / `LOWER()` | K13 / K14 | Converts text to uppercase / lowercase |
| `XLOOKUP()` | L16 | Looks up a salesperson's final sale amount, a modern alternative to VLOOKUP |
| `XMATCH()` inside `INDEX()`, wrapped in `IFERROR()` | L19 | Looks up a product name by Sales ID, with a graceful "Product not found" fallback |
| `INDIRECT()` + `SUM()` | L22 | Sums a range defined dynamically as text (`"E2:E6"`) rather than a hardcoded cell reference |
| `ROUND()`, `CEILING()`, `FLOOR()` | (labeled reference cells) | Practiced separately as core rounding functions |

---

## 📄 Sheet 3 — Employees Data

**Columns:** Employee ID, Name, Department, Salary, Joining Date

| Formula/Feature | Cell(s) | Purpose |
|---|---|---|
| `INDEX()` + `MATCH()` (two-way lookup) | H2 | Retrieves an employee's department by matching both Employee ID (row) and column header (e.g. "Department") — a lookup with no fixed column position |
| `UPPER()` | H6 | Converts a department name to uppercase |
| `LOWER()` | H7 | Converts an employee name to lowercase |
| `LEFT()` | H8 | Extracts the first 3 characters of an employee's name |
| `XLOOKUP()` | H10 | Retrieves an employee's salary by Employee ID |
| `TODAY()` − `XLOOKUP()` | H13 | Calculates an employee's tenure/age in days by subtracting their joining date (looked up dynamically) from today's date |
| `ABS()` on a date difference | G17 | Calculates the absolute difference between two dates, regardless of which is earlier |

---

## 🧪 Concepts Practiced

- **Logical functions:** `IF`, nested `IF`, `AND`, `OR` for grading, eligibility, and classification logic
- **Lookup & reference functions:** `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX` + `MATCH` (including two-way and multi-condition lookups)
- **Statistical/aggregation functions:** `AVERAGE`, `AVERAGEIFS`, `COUNTIF`, `SUMIFS`, `SUM`
- **Text functions:** `UPPER`, `LOWER`, `LEFT`, `FIND`
- **Date & time functions:** `TODAY`, date subtraction, `ABS` for absolute date differences
- **Rounding functions:** `ROUND`, `CEILING`, `FLOOR`
- **Dynamic arrays:** `FILTER` for returning a full list of matching records without manual filtering
- **Indirect referencing:** `INDIRECT` to build a range reference from text dynamically
- **Error handling in formulas:** `IFERROR` to show friendly fallback messages instead of `#N/A` errors

---

## 🚧 Challenges & Solutions

| Challenge | Solution |
|---|---|
| Assigning grades based on multiple score thresholds | Used nested `IF()` statements to check thresholds in descending order (A → B → C → Fail) |
| Checking two conditions at once for high performers | Combined `AND()` inside `IF()` to require both Maths and Science above 80 |
| Looking up values where the target column isn't fixed | Used `INDEX()` + `MATCH()` with a matched column header instead of a hardcoded column number |
| Summing sales for a specific Region **and** Product combination | Used `SUMIFS()` with two criteria ranges instead of a single-condition `SUMIF()` |
| Matching a salesperson's sale for a specific month | Built an array formula combining `INDEX()`, `MATCH()`, and multiplied boolean conditions to match two criteria simultaneously |
| Avoiding `#N/A` errors when a lookup value doesn't exist | Wrapped `XMATCH()`/`INDEX()` combinations in `IFERROR()` to show a clear fallback message |
| Referencing a range that needed to change dynamically | Used `INDIRECT()` to convert a text string into a usable cell range reference |

---

## 💡 What I Learned

- How to layer multiple **nested IF conditions** to classify data into more than two categories
- The difference between older lookup functions (`VLOOKUP`) and newer, more flexible ones (`XLOOKUP`, `XMATCH`) — including built-in "not found" handling
- How `INDEX` + `MATCH` can perform lookups that `VLOOKUP` can't, such as looking left of the lookup column or matching on a dynamic column header
- Building **multi-condition lookups** using array formulas when a single `MATCH()` isn't enough
- The value of `IFERROR()` for turning a broken-looking `#N/A` into a clear, human-readable message
- How `INDIRECT()` allows a formula to reference a range built from text, enabling more dynamic spreadsheet logic
- Practical use of Excel's newer **dynamic array functions** (`FILTER`) to return multiple matching rows without manual filtering or helper columns

---

## 📁 Repository Contents

```
excel-formulas-practice/
│
├── Project-1.xlsx     # Main Excel workbook (3 sheets)
└── README.md           # Project documentation
```

---

## 👤 Author

**Prerita Morashiya**
