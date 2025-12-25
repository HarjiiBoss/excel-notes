# Sorting and Filtering

This file covers how to sort and filter data in Excel to organize information, find specific records, and analyze subsets of your data.

---

## What is Sorting?

**Sorting** rearranges rows in your data based on the values in one or more columns.

### Purpose
- **Organize data** alphabetically or numerically
- **Find patterns** (highest to lowest, A to Z)
- **Prepare for analysis** (group similar items)
- **Rank items** (best to worst performance)

### Visual Example

**Before Sorting:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ Mike   │ 3500   │ West   │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │
  ├────────┼────────┼────────┤
4 │ John   │ 5000   │ North  │
  └────────┴────────┴────────┘
```

**After Sorting by Sales (High to Low):**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ Sarah  │ 9000   │ East   │
  ├────────┼────────┼────────┤
3 │ John   │ 5000   │ North  │
  ├────────┼────────┼────────┤
4 │ Mike   │ 3500   │ West   │
  └────────┴────────┴────────┘
```

---

## What is Filtering?

**Filtering** temporarily hides rows that don't meet your criteria, showing only the data you want to see.

### Purpose
- **Focus on specific records** (only West region)
- **Find items matching criteria** (sales > $5000)
- **Exclude unwanted data** (hide completed tasks)
- **Analyze subsets** (Q4 only)

### Key Difference from Sorting

| Sorting | Filtering |
|---------|-----------|
| Rearranges all rows | Hides some rows |
| All data visible | Only matching data visible |
| Changes order | Keeps original order |
| Permanent until re-sorted | Temporary, can be cleared |

---

## Basic Sorting

### Sort A to Z (Ascending)

**Steps:**
1. Click any cell in the column you want to sort
2. Data Tab → A→Z button (Sort A to Z)

**Or:**
1. Right-click column header
2. Select "Sort A to Z"

### Sort Z to A (Descending)

**Steps:**
1. Click any cell in the column
2. Data Tab → Z→A button (Sort Z to A)

### Visual: Sort Buttons

```
┌─────────────────────────────────────┐
│ Data Tab                            │
├─────────────────────────────────────┤
│  [A→Z]  [Z→A]  [Sort]  [Filter]    │
│    ↑      ↑       ↑        ↑        │
│ Ascending│    Custom   Enable       │
│      Descending  Sort    Filter     │
└─────────────────────────────────────┘
```

### Example: Sort Names Alphabetically

**Before:**
```
     A
  ┌────────┐
1 │ Name   │
  ├────────┤
2 │ Zoe    │
  ├────────┤
3 │ Alice  │
  ├────────┤
4 │ Mike   │
  └────────┘
```

**After (A→Z):**
```
     A
  ┌────────┐
1 │ Name   │
  ├────────┤
2 │ Alice  │
  ├────────┤
3 │ Mike   │
  ├────────┤
4 │ Zoe    │
  └────────┘
```

---

## Custom Sort (Multiple Columns)

Sort by multiple criteria in order of priority.

### How to Access

**Data Tab → Sort** (opens Sort dialog)

### Sort Dialog

```
┌──────────────────────────────────────┐
│ Sort                                 │
├──────────────────────────────────────┤
│ ☑ My data has headers                │
├──────────────────────────────────────┤
│ Sort by:     [Region ▼]              │
│ Sort On:     [Values  ▼]             │
│ Order:       [A to Z  ▼]             │
├──────────────────────────────────────┤
│ Then by:     [Sales   ▼]             │
│ Sort On:     [Values  ▼]             │
│ Order:       [Largest to Smallest ▼] │
├──────────────────────────────────────┤
│ [Add Level] [Delete Level] [Copy]    │
└──────────────────────────────────────┘
```

### Example: Sort by Region, Then by Sales

**Original Data:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │
  ├────────┼────────┼────────┤
4 │ Mike   │ 3500   │ West   │
  ├────────┼────────┼────────┤
5 │ Lisa   │ 7000   │ West   │
  └────────┴────────┴────────┘
```

**After Sort:**
```
Sort by: Region (A to Z)
Then by: Sales (Largest to Smallest)

     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ Sarah  │ 9000   │ East   │ ← East first
  ├────────┼────────┼────────┤
3 │ John   │ 5000   │ East   │ ← East (lower sales)
  ├────────┼────────┼────────┤
4 │ Lisa   │ 7000   │ West   │ ← West next
  ├────────┼────────┼────────┤
5 │ Mike   │ 3500   │ West   │ ← West (lower sales)
  └────────┴────────┴────────┘
```

---

## Sort Options

### Sort By Values (Default)

Sorts by the actual cell values.

### Sort By Cell Color

Sort rows by background color.

**Use case:** After conditional formatting, group colored cells together

**Steps:**
1. Data → Sort
2. Sort On: Cell Color
3. Order: Choose color to put on top

### Sort By Font Color

Sort rows by text color.

**Use case:** Group cells with red text together

### Sort By Icon

Sort by conditional formatting icons.

**Use case:** Group traffic light status together

### Sort Left to Right

Sort columns instead of rows (rare).

**Steps:**
1. Data → Sort
2. Options → Sort left to right
3. Choose row to sort by

---

## Basic Filtering

### Enable Filter

**Steps:**
1. Click any cell in your data
2. Data Tab → Filter

**Visual indicator:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │Name [▼]│Sales[▼]│Region[▼]│ ← Dropdown arrows appear
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
  └────────┴────────┴────────┘
```

### Filter Dropdown Menu

Click dropdown arrow to see:
```
┌────────────────────────────┐
│ Sort A to Z                │
│ Sort Z to A                │
│ Sort by Color              │
├────────────────────────────┤
│ Clear Filter               │
│ Filter by Color            │
├────────────────────────────┤
│ Text Filters ▶             │
├────────────────────────────┤
│ ☑ (Select All)             │
│ ☑ East                     │
│ ☑ North                    │
│ ☑ West                     │
└────────────────────────────┘
```

### Filter by Selection

**Steps:**
1. Uncheck items you don't want to see
2. Click OK

**Result:** Only checked items are visible

### Visual Example: Filter for "East" Only

**Before Filter:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │
  ├────────┼────────┼────────┤
4 │ Mike   │ 3500   │ West   │
  ├────────┼────────┼────────┤
5 │ Lisa   │ 7000   │ North  │
  └────────┴────────┴────────┘
```

**After Filter (East only):**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │Region[▼]│ ← Blue arrow = filter active
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │
  └────────┴────────┴────────┘

Rows 4-5 hidden (not deleted)
Row numbers: 1, 2, 3 (no 4, 5 shown in blue)
```

---

## Number Filters

For numeric columns, Excel provides special filter options.

### Number Filter Options

**Data → Filter → Number Filters:**
```
┌────────────────────────────┐
│ Equals...                  │
│ Does Not Equal...          │
│ Greater Than...            │
│ Greater Than or Equal To...│
│ Less Than...               │
│ Less Than or Equal To...   │
│ Between...                 │
│ Top 10...                  │
│ Above Average              │
│ Below Average              │
└────────────────────────────┘
```

### Example 1: Greater Than

**Filter for Sales > $6000:**

```
Original (5 rows):
John   │ 5000   ← Hidden
Sarah  │ 9000   ✓ Shown
Mike   │ 3500   ← Hidden
Lisa   │ 7000   ✓ Shown
Tom    │ 8500   ✓ Shown

Result: 3 rows visible (Sarah, Lisa, Tom)
```

**Steps:**
1. Click dropdown on Sales column
2. Number Filters → Greater Than
3. Enter: `6000`
4. Click OK

### Example 2: Between

**Filter for Sales between $4000 and $8000:**

**Steps:**
1. Number Filters → Between
2. Enter: `4000` AND `8000`
3. Click OK

```
Result shows:
John   │ 5000   ✓
Lisa   │ 7000   ✓
```

### Example 3: Top 10

**Show top 3 performers:**

**Steps:**
1. Number Filters → Top 10
2. Change to: Top `3` Items
3. Click OK

```
Shows highest 3 sales values
```

---

## Text Filters

For text columns, Excel provides text-specific options.

### Text Filter Options

**Data → Filter → Text Filters:**
```
┌────────────────────────────┐
│ Equals...                  │
│ Does Not Equal...          │
│ Begins With...             │
│ Ends With...               │
│ Contains...                │
│ Does Not Contain...        │
└────────────────────────────┘
```

### Example 1: Contains

**Filter names containing "ar":**

```
Sarah   ✓ (contains "ar")
Mark    ✓ (contains "ar")
John    ← Hidden
Lisa    ← Hidden
```

**Steps:**
1. Text Filters → Contains
2. Enter: `ar`
3. Click OK

### Example 2: Begins With

**Filter products starting with "Pro":**

```
Product Pro     ✓
Pro Series      ✓
Widget          ← Hidden
Gadget Pro      ← Hidden (doesn't START with "Pro")
```

**Steps:**
1. Text Filters → Begins With
2. Enter: `Pro`
3. Click OK

### Example 3: Does Not Contain

**Exclude test data:**

```
Customer A      ✓
Customer B      ✓
Test Account    ← Hidden
Customer C      ✓
Test User       ← Hidden
```

**Steps:**
1. Text Filters → Does Not Contain
2. Enter: `Test`
3. Click OK

---

## Date Filters

For date columns, Excel provides date-specific options.

### Date Filter Options

**Data → Filter → Date Filters:**
```
┌────────────────────────────┐
│ Equals...                  │
│ Before...                  │
│ After...                   │
│ Between...                 │
├────────────────────────────┤
│ Tomorrow                   │
│ Today                      │
│ Yesterday                  │
├────────────────────────────┤
│ Next Week                  │
│ This Week                  │
│ Last Week                  │
├────────────────────────────┤
│ Next Month                 │
│ This Month                 │
│ Last Month                 │
├────────────────────────────┤
│ Next Quarter               │
│ This Quarter               │
│ Last Quarter               │
├────────────────────────────┤
│ Next Year                  │
│ This Year                  │
│ Last Year                  │
├────────────────────────────┤
│ Year to Date               │
│ All Dates in the Period ▶  │
└────────────────────────────┘
```

### Example 1: This Month

**Show only current month records:**

```
1/15/2024   ← Hidden (last month)
12/5/2024   ✓ Shown (this month)
12/20/2024  ✓ Shown (this month)
```

**Steps:**
1. Date Filters → This Month
2. Click OK

### Example 2: Between Dates

**Show Q4 2024 (Oct-Dec):**

**Steps:**
1. Date Filters → Between
2. Enter: `10/1/2024` AND `12/31/2024`
3. Click OK

### Example 3: Last Week

**Show last week's orders:**

**Steps:**
1. Date Filters → Last Week
2. Click OK

**Note:** "Last Week" is dynamic - updates automatically based on today's date

---

## Filter by Color

Filter rows based on cell or font color (from conditional formatting or manual formatting).

### Filter by Cell Color

**Steps:**
1. Click filter dropdown
2. Filter by Color → Choose color
3. Only rows with that color show

### Example

```
After conditional formatting marks top performers in green:

Sarah  │ 9000  [Green]   ✓ Shown
John   │ 5000  [White]   ← Hidden
Lisa   │ 8000  [Green]   ✓ Shown
Mike   │ 3500  [White]   ← Hidden
```

**Steps:**
1. Filter by Color → Green
2. Shows only green-highlighted rows

---

## Multiple Filters

You can filter multiple columns simultaneously.

### How It Works

Each filter is combined with AND logic:
```
Region = "East" AND Sales > 5000
```

### Example: Multi-Column Filter

**Original Data:**
```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │ Status │
  ├────────┼────────┼────────┼────────┤
2 │ John   │ 8000   │ East   │ Active │
  ├────────┼────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │ Active │
  ├────────┼────────┼────────┼────────┤
4 │ Mike   │ 3500   │ West   │ Active │
  ├────────┼────────┼────────┼────────┤
5 │ Lisa   │ 7000   │ East   │ Inactive│
  └────────┴────────┴────────┴────────┘
```

**Apply Filters:**
1. Region = "East"
2. Sales > 7000
3. Status = "Active"

**Result:**
```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │ Name   │Sales[▼]│Region[▼]│Status[▼]│
  ├────────┼────────┼────────┼────────┤
2 │ John   │ 8000   │ East   │ Active │
  ├────────┼────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │ Active │
  └────────┴────────┴────────┴────────┘

Only rows meeting ALL criteria shown
```

---

## Clear Filters

### Clear Filter from One Column

**Steps:**
1. Click dropdown arrow on filtered column
2. Clear Filter from [Column Name]

### Clear All Filters

**Steps:**
1. Data Tab → Clear
2. All filters removed, all data visible again

**Or:**
- Data Tab → Filter (toggle off, then back on)

---

## Advanced Filter

For complex criteria, use Advanced Filter.

### When to Use Advanced Filter

- Multiple OR conditions
- Criteria across different columns
- Extract filtered data to new location
- Complex formulas as criteria

### How Advanced Filter Works

You create a **criteria range** on the worksheet, then Excel filters based on that range.

### Example: Advanced Filter Setup

**Data:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
3 │ Sarah  │ 9000   │ East   │
4 │ Mike   │ 3500   │ West   │
5 │ Lisa   │ 7000   │ West   │
  └────────┴────────┴────────┘
```

**Criteria Range (set up above or below data):**
```
     E         F         G
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │ ← Headers (must match data)
  ├────────┼────────┼────────┤
2 │        │ >6000  │        │ ← Criteria: Sales > 6000
  └────────┴────────┴────────┘
```

**Steps:**
1. Set up criteria range (E1:G2)
2. Click in data range
3. Data Tab → Advanced
4. List Range: A1:C5
5. Criteria Range: E1:G2
6. Click OK

**Result:** Shows Sarah (9000) and Lisa (7000)

### Advanced Filter: OR Conditions

**To use OR, put criteria on different rows:**

```
Criteria:
     E         F         G
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │        │        │ East   │ ← Region = East
  ├────────┼────────┼────────┤
3 │        │ >8000  │        │ ← OR Sales > 8000
  └────────┴────────┴────────┘

Shows: All East region OR any sales > 8000
```

### Advanced Filter: Copy to Another Location

**Steps:**
1. Data → Advanced
2. Select "Copy to another location"
3. Copy to: Select destination cell
4. Click OK

**Result:** Filtered data copied to new location, original data unchanged

---

## Sort and Filter with Tables

When you convert a range to a Table (Insert Tab → Table), sorting and filtering become more powerful.

### Table Advantages

✅ **Auto-expand:** New rows automatically included
✅ **Better filtering:** Filter buttons always visible
✅ **Easier formulas:** Use column names
✅ **Cleaner look:** Built-in styling
✅ **Total row:** Quick calculations

### Visual: Table with Filters

```
     A         B         C
  ┌────────┬────────┬────────┐
1 │Name [▼]│Sales[▼]│Region[▼]│ ← Built-in filter buttons
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ East   │
  ├────────┼────────┼────────┤
4 │ Total  │ 14000  │        │ ← Optional total row
  └────────┴────────┴────────┘
```

---

## Search in Filter

For large lists, use the search box in filter dropdown.

### How to Use

**Steps:**
1. Click filter dropdown
2. Type in search box
3. Matching items appear
4. Select items to show
5. Click OK

### Example

```
Column has 1000 customer names
Search for: "Smith"

Results:
☑ John Smith
☑ Sarah Smith
☑ Mike Smithson
☐ (All others hidden in list)
```

---

## Filter Status Indicators

Excel shows when filters are active:

### Visual Indicators

```
Active Filter:
Column[▼] → Blue funnel icon

No Filter:
Column[▼] → Gray funnel icon

Row Numbers:
1, 3, 5, 7 → Blue (indicates hidden rows between)
```

### Status Bar

When filtered, status bar shows:
```
"X of Y records found"
```

Example: "5 of 20 records found"

---

## Sorting and Filtering Best Practices

### 1. Use Headers

```
✅ Always have header row
✅ One row of headers only
✅ No blank rows between headers and data
```

### 2. Keep Data Contiguous

```
❌ Avoid blank rows/columns in middle of data
✅ Keep data in continuous range
```

### 3. Remove Merged Cells

```
❌ Merged cells break sorting/filtering
✅ Unmerge before sorting
```

### 4. Format Consistently

```
✅ Numbers as numbers, not text
✅ Dates as actual dates
✅ Consistent text format (no leading spaces)
```

### 5. Clear Filters When Done

```
✅ Clear filters before sharing
✅ Document active filters if keeping them
```

### 6. Use Tables for Better Management

```
✅ Convert range to Table for easier filtering
✅ Table filters persist better
✅ Easier to work with formulas
```

---

## Common Mistakes

### Mistake 1: Sorting Without Selecting All Columns

```
❌ Problem: Select only one column and sort
Result: Data becomes misaligned

✅ Solution: Click any cell in data, Excel selects entire range
Or: Select all columns before sorting
```

### Mistake 2: Not Including Headers

```
❌ Problem: Sort includes header row
Result: "Name" sorted with data

✅ Solution: Check "My data has headers" in Sort dialog
```

### Mistake 3: Blank Rows in Data

```
❌ Problem: Blank rows in middle of data
Result: Excel doesn't select full range

✅ Solution: Remove blank rows before sorting/filtering
```

### Mistake 4: Forgetting Active Filters

```
❌ Problem: Filters still active, data looks incomplete
Result: Formulas calculate on visible rows only

✅ Solution: Always clear filters before final analysis
Check for blue row numbers (indicates hidden rows)
```

### Mistake 5: Mixed Data Types

```
❌ Problem: Numbers stored as text
Result: Sort order incorrect (1, 10, 2, 20...)

✅ Solution: Convert text to numbers first
```

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + Shift + L` | Toggle Filter on/off |
| `Alt + ↓` | Open filter dropdown (when in header cell) |
| `Alt + D, S` | Open Sort dialog |
| `Ctrl + Home` | Go to first visible cell (when filtered) |
| `Alt + ;` | Select visible cells only |
| `Ctrl + Shift + End` | Select to last visible cell |

---

## Real-World Example: Sales Report

**Scenario:** Analyze sales data for East region, top performers only

**Original Data (100 rows):**
```
Name│Sales│Region│Quarter│Status
...many rows...
```

**Analysis Steps:**

**Step 1: Filter Region**
- Region = "East"
- Result: 30 rows visible

**Step 2: Filter Status**
- Status = "Active"
- Result: 25 rows visible

**Step 3: Sort by Sales**
- Sort Sales: Largest to Smallest
- Result: Top performers at top

**Step 4: Filter Top 10**
- Number Filters → Top 10 Items
- Result: Top 10 East region active reps

**Final View:**
```
     A         B         C         D         E
  ┌────────┬────────┬────────┬────────┬────────┐
1 │Name [▼]│Sales[▼]│Region[▼]│Quarter│Status[▼]│
  ├────────┼────────┼────────┼────────┼────────┤
3 │ Sarah  │ 12000  │ East   │ Q4     │ Active │
  ├────────┼────────┼────────┼────────┼────────┤
7 │ John   │ 11500  │ East   │ Q4     │ Active │
  ├────────┼────────┼────────┼────────┼────────┤
12│ Mike   │ 10800  │ East   │ Q3     │ Active │
  └────────┴────────┴────────┴────────┴────────┘

Shows 10 of 100 records
```

---

## Real-World Example: Task Management

**Scenario:** Find overdue high-priority tasks

**Data:**
```
Task│Priority│Status│Due Date│Owner
...
```

**Filters Applied:**
1. Status ≠ "Complete"
2. Priority = "High"
3. Due Date < TODAY()

**Sort:**
- Due Date: Oldest to Newest

**Result:** Urgent tasks requiring immediate attention, sorted by how overdue they are

---

## Troubleshooting Guide

### Problem: Can't Sort - Button Grayed Out

**Causes:**
- Cell is in Edit mode (press Esc)
- Sheet is protected
- Cells are merged

**Solution:**
- Exit edit mode
- Unprotect sheet
- Unmerge cells

### Problem: Sort Mixes Up Data

**Cause:** Didn't select entire data range

**Solution:**
- Click single cell in data
- Let Excel auto-select range
- Or manually select all columns

### Problem: Filter Doesn't Show All Options

**Cause:** Excel shows only unique values (max 10,000)

**Solution:**
- Use search box in filter
- Or use Advanced Filter

### Problem: Numbers Sort Incorrectly

**Cause:** Numbers stored as text

**Solution:**
```
1. Select column
2. Data → Text to Columns
3. Click Finish
```

### Problem: Can't See Filter Arrows

**Cause:** Filter not enabled, or outside data range

**Solution:**
- Click in data range
- Data → Filter (toggle on)

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Sorting rearranges rows
- Filtering hides rows temporarily
- Data → Filter to enable filtering
- Blue row numbers = hidden rows
- Filter dropdown has checkboxes for selection
- Multiple filters use AND logic
- Ctrl + Shift + L toggles Filter on/off

### Practice Deeply
- Sorting by single column (A→Z, Z→A)
- Custom sort with multiple levels
- Enabling and using basic filters
- Using number filters (Greater Than, Between, Top 10)
- Using text filters (Contains, Begins With)
- Using date filters (This Month, Between)
- Applying multiple filters simultaneously
- Clearing filters from columns
- Sorting by color after conditional formatting
- Understanding when rows are hidden vs deleted
- Converting ranges to Tables for better filtering
- Creating criteria for Advanced Filter
- Recognizing visual indicators (blue arrows, blue row numbers)
- Testing filters before sharing workbooks

---

## Quick Reference: Common Filter Types

| Data Type | Useful Filters | Example |
|-----------|---------------|---------|
| **Numbers** | Greater Than, Top 10, Between | Sales > 5000 |
| **Text** | Contains, Begins With, Equals | Name contains "Smith" |
| **Dates** | This Month, Between, Before | Orders this month |
| **Colors** | Filter by Color | Show only red cells |
| **Blanks** | (Blanks) checkbox | Find empty cells |

---

## Next Step

After this file, we move to:

**`13-pivot-tables.md`**
- What are Pivot Tables
- Creating your first Pivot Table
- Arranging fields (Rows, Columns, Values, Filters)
- Summarizing data (Sum, Count, Average)
- Grouping data (dates, numbers, text)
- Calculated fields
- Pivot Table styles and formatting
- Refreshing Pivot Tables
