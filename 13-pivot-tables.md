# Pivot Tables

This file covers Pivot Tables - Excel's most powerful tool for summarizing, analyzing, and exploring large datasets without writing formulas.

---

## What is a Pivot Table?

A **Pivot Table** is an interactive table that automatically summarizes data from a larger dataset.

### Purpose
- **Summarize thousands of rows** into a compact report
- **Group and aggregate data** (sum, count, average)
- **Analyze by different dimensions** (by region, by month, by product)
- **Find patterns and trends** quickly
- **Answer business questions** without formulas

### Visual Concept

**Source Data (100 rows):**
```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │ Date   │ Sales  │ Region │ Product│
  ├────────┼────────┼────────┼────────┤
2 │ 1/5/24 │ 1000   │ East   │ Widget │
3 │ 1/7/24 │ 1500   │ West   │ Gadget │
4 │ 1/8/24 │ 2000   │ East   │ Widget │
... (97 more rows)
```

**Pivot Table (Summary):**
```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │        │ East   │ West   │ Total  │
  ├────────┼────────┼────────┼────────┤
2 │ Widget │ 45000  │ 38000  │ 83000  │
  ├────────┼────────┼────────┼────────┤
3 │ Gadget │ 32000  │ 41000  │ 73000  │
  ├────────┼────────┼────────┼────────┤
4 │ Total  │ 77000  │ 79000  │ 156000 │
  └────────┴────────┴────────┴────────┘

Automatic summation by Product × Region
```

---

## When to Use Pivot Tables

### Perfect For:
✅ Summarizing sales by region, month, or product
✅ Counting occurrences (how many orders per customer)
✅ Finding averages (average order value by category)
✅ Comparing time periods (this year vs last year)
✅ Identifying top performers
✅ Creating quick reports without formulas

### Not Ideal For:
❌ Small datasets (< 20 rows) - regular formulas work fine
❌ Data entry or editing individual records
❌ Complex calculations requiring multiple steps
❌ Data that needs to be modified directly

---

## Creating Your First Pivot Table

### Prerequisites

Your source data should:
- Have **headers** in the first row
- Be in a **contiguous range** (no blank rows/columns)
- Have **consistent data types** per column
- Be **organized properly** (one record per row)

### Example Source Data

```
     A         B         C         D         E
  ┌────────┬────────┬────────┬────────┬────────┐
1 │ Date   │ Salesperson│Region│Product │ Sales  │
  ├────────┼────────┼────────┼────────┼────────┤
2 │ 1/5/24 │ John   │ East   │ Widget │ 1000   │
  ├────────┼────────┼────────┼────────┼────────┤
3 │ 1/7/24 │ Sarah  │ West   │ Gadget │ 1500   │
  ├────────┼────────┼────────┼────────┼────────┤
4 │ 1/8/24 │ John   │ East   │ Widget │ 2000   │
  ├────────┼────────┼────────┼────────┼────────┤
5 │ 1/9/24 │ Mike   │ North  │ Tool   │ 800    │
  └────────┴────────┴────────┴────────┴────────┘
```

### Steps to Create

**Method 1: Insert Tab**
1. Click any cell in your data
2. **Insert Tab → PivotTable**
3. Verify range is correct
4. Choose where to place it:
   - New Worksheet (recommended)
   - Existing Worksheet
5. Click **OK**

**Method 2: Quick Analysis**
1. Select your data range
2. Press **Ctrl + Q**
3. Select **Tables** → **PivotTable**

### What Happens Next

Excel creates:
1. A new worksheet (or uses existing)
2. An empty Pivot Table placeholder
3. A **PivotTable Fields** pane on the right

```
┌──────────────────────────────────────┐
│ PivotTable Fields                    │
├──────────────────────────────────────┤
│ Choose fields to add to report:      │
│                                      │
│ ☐ Date                               │
│ ☐ Salesperson                        │
│ ☐ Region                             │
│ ☐ Product                            │
│ ☐ Sales                              │
├──────────────────────────────────────┤
│ Drag fields between areas below:     │
│                                      │
│ ┌──────────┐  ┌──────────┐         │
│ │ Filters  │  │ Columns  │         │
│ └──────────┘  └──────────┘         │
│ ┌──────────┐  ┌──────────┐         │
│ │ Rows     │  │ Values   │         │
│ └──────────┘  └──────────┘         │
└──────────────────────────────────────┘
```

---

## Understanding the Four Areas

Every Pivot Table has four areas where you can place fields:

### 1. Filters Area
**Purpose:** Filter entire Pivot Table by a field

**Example:** Filter showing only 2024 data
```
Year: [2024 ▼]

Region │ Sales
───────┼──────
East   │ 50000
West   │ 45000
```

### 2. Rows Area
**Purpose:** Create row labels (left side)

**Example:** Regions as rows
```
Region │ Sales
───────┼──────
East   │ 50000
West   │ 45000
North  │ 38000
```

### 3. Columns Area
**Purpose:** Create column headers (top)

**Example:** Products as columns
```
       │ Widget│ Gadget│ Tool
───────┼───────┼───────┼──────
East   │ 20000 │ 18000 │ 12000
West   │ 15000 │ 20000 │ 10000
```

### 4. Values Area
**Purpose:** Numbers to calculate (sum, count, average, etc.)

**Example:** Sales amount to sum
```
Shows: Sum of Sales for each combination
```

---

## Building Your First Pivot Table: Step by Step

**Goal:** Show total sales by Region

### Step 1: Drag Region to Rows
```
PivotTable Fields:
┌────────────┐
│ Rows       │
│ Region     │
└────────────┘

Result:
Region
East
West
North
```

### Step 2: Drag Sales to Values
```
PivotTable Fields:
┌────────────┐  ┌────────────┐
│ Rows       │  │ Values     │
│ Region     │  │ Sum of Sales│
└────────────┘  └────────────┘

Result:
Region │ Sum of Sales
───────┼─────────
East   │ 50000
West   │ 45000
North  │ 38000
Total  │ 133000
```

**Congratulations!** You've created your first Pivot Table.

---

## Adding More Dimensions

### Add Product to Columns

**Drag Product to Columns area:**

```
PivotTable:
             Widget  Gadget   Tool    Total
East         20000   18000   12000   50000
West         15000   20000   10000   45000
North        13000   15000   10000   38000
Total        48000   53000   32000  133000
```

Now you see sales by **Region AND Product** simultaneously.

### Add Salesperson to Rows

**Drag Salesperson below Region in Rows:**

```
PivotTable:
Region │ Salesperson │ Sum of Sales
───────┼─────────────┼─────────
East   │ John        │ 28000
       │ Sarah       │ 22000
West   │ Mike        │ 25000
       │ Lisa        │ 20000
North  │ Tom         │ 23000
       │ Emma        │ 15000
Total                 133000
```

This creates a **hierarchical view** (Region > Salesperson).

---

## Changing Value Calculation

By default, Pivot Tables **SUM** numeric fields. You can change this.

### Available Calculations

| Function | Purpose | Use Case |
|----------|---------|----------|
| **Sum** | Total of values | Total sales |
| **Count** | Number of items | Number of orders |
| **Average** | Mean value | Average order size |
| **Max** | Largest value | Highest sale |
| **Min** | Smallest value | Lowest sale |
| **Product** | Multiply values | Rare use |
| **Count Numbers** | Count numeric entries | Data validation |
| **StdDev** | Standard deviation | Statistical analysis |
| **Var** | Variance | Statistical analysis |

### How to Change Calculation

**Steps:**
1. Click the field in Values area (or right-click value in table)
2. Select **Value Field Settings**
3. Choose calculation type
4. Click **OK**

### Example: Count Orders Instead of Sum Sales

**Before (Sum):**
```
Region │ Sum of Sales
───────┼─────────
East   │ 50000
West   │ 45000
```

**After (Count):**
```
Region │ Count of Sales
───────┼────────────
East   │ 25
West   │ 18
```

Shows: Number of transactions per region, not dollar amounts.

---

## Value Field Settings

When you click **Value Field Settings**, you see:

```
┌─────────────────────────────────────┐
│ Value Field Settings                │
├─────────────────────────────────────┤
│ Source Name: Sales                  │
│ Custom Name: [Sum of Sales     ]    │
├─────────────────────────────────────┤
│ Summarize Values By:                │
│   ○ Sum                             │
│   ○ Count                           │
│   ○ Average                         │
│   ○ Max                             │
│   ○ Min                             │
│   ... (more options)                │
├─────────────────────────────────────┤
│ Number Format: [Format...]          │
└─────────────────────────────────────┘
```

### Custom Name

Change the label that appears in your Pivot Table:
```
Default: "Sum of Sales"
Custom:  "Total Revenue"
```

### Number Format

Apply formatting to values:
- Currency: $50,000
- Percentage: 15%
- Thousands separator: 50,000
- Decimal places: 50,000.00

---

## Grouping Data

Group similar items together for better analysis.

### Grouping Dates

**Scenario:** You have daily sales data, want to see monthly totals.

**Original Pivot Table (Daily):**
```
Date     │ Sales
─────────┼──────
1/1/2024 │ 1000
1/2/2024 │ 1200
1/3/2024 │ 900
... (365 rows)
```

**Steps to Group by Month:**
1. Right-click any date in Rows area
2. Select **Group**
3. Choose **Months**
4. Click **OK**

**Result (Monthly):**
```
Date     │ Sales
─────────┼──────
Jan 2024 │ 45000
Feb 2024 │ 48000
Mar 2024 │ 52000
... (12 rows)
```

### Grouping Date Options

You can group by:
- **Seconds**
- **Minutes**
- **Hours**
- **Days**
- **Months**
- **Quarters**
- **Years**

**Multiple selections:** Group by Quarters AND Months simultaneously
```
2024
  Q1
    January   │ 45000
    February  │ 48000
    March     │ 52000
  Q2
    April     │ 50000
    ...
```

### Grouping Numbers

**Scenario:** Age data, want to group into ranges.

**Original:**
```
Age │ Count
────┼──────
18  │ 5
19  │ 8
20  │ 12
... (60 different ages)
```

**Steps:**
1. Right-click any number in Rows
2. Select **Group**
3. Starting at: `18`
4. Ending at: `65`
5. By: `10` (creates 10-year ranges)
6. Click **OK**

**Result:**
```
Age       │ Count
──────────┼──────
18-27     │ 45
28-37     │ 67
38-47     │ 52
48-57     │ 38
58-65     │ 23
```

---

## Filtering Pivot Tables

### Filter Using Row/Column Labels

Click dropdown arrow next to any row/column label:

```
Region [▼]
┌──────────────┐
│ ☑ (Select All)│
│ ☑ East       │
│ ☑ West       │
│ ☑ North      │
│ ☐ South      │
└──────────────┘
```

Uncheck items to hide them from Pivot Table.

### Filter Using Slicers

**Slicers** provide visual filter buttons.

**Steps to Add Slicer:**
1. Click anywhere in Pivot Table
2. **PivotTable Analyze Tab → Insert Slicer**
3. Select field(s) to create slicers for
4. Click **OK**

**Visual Result:**
```
┌────────────────────┐
│ Region             │
├────────────────────┤
│ [East] [West]      │
│ [North] [South]    │
└────────────────────┘

Click buttons to filter
Multiple selection: Ctrl + Click
```

### Filter Using Report Filter (Filters Area)

**Drag field to Filters area:**

```
At top of Pivot Table:
Region: [All ▼]

Click dropdown to select specific region
```

### Label Filters

Right-click row/column label → **Filter**:
- **Equals**
- **Does Not Equal**
- **Begins With**
- **Contains**
- **Greater Than** (numbers only)
- **Top 10**

### Value Filters

Filter based on values in the data:
- **Top 10** items by sales
- **Above Average** sales
- **Greater Than** $50,000

**Example:**
```
Show only regions with sales > $40,000

Result:
Region │ Sales
───────┼──────
East   │ 50000
West   │ 45000
(North hidden because 38000 < 40000)
```

---

## Show Values As (Advanced Calculations)

Display values as percentages, running totals, or comparisons.

### Common Options

| Show Values As | Purpose | Example |
|---------------|---------|---------|
| **No Calculation** | Raw values (default) | $50,000 |
| **% of Grand Total** | Percentage of total | 35% |
| **% of Column Total** | Percentage within column | 60% |
| **% of Row Total** | Percentage within row | 45% |
| **Running Total In** | Cumulative sum | Month 3: $150,000 total |
| **% Difference From** | Change from baseline | +15% vs January |
| **Rank Largest to Smallest** | Position ranking | Rank 3 of 10 |

### Example: % of Grand Total

**Steps:**
1. Right-click value in Pivot Table
2. **Show Values As → % of Grand Total**

**Before:**
```
Region │ Sales
───────┼──────
East   │ 50000
West   │ 45000
North  │ 38000
Total  │ 133000
```

**After:**
```
Region │ % of Total
───────┼──────
East   │ 37.6%
West   │ 33.8%
North  │ 28.6%
Total  │ 100.0%
```

### Example: Running Total by Month

**Steps:**
1. Value Field Settings → Show Values As
2. Select **Running Total In**
3. Base field: **Date**

**Result:**
```
Month  │ Sales │ Running Total
───────┼───────┼──────────
Jan    │ 45000 │ 45000
Feb    │ 48000 │ 93000
Mar    │ 52000 │ 145000
Apr    │ 50000 │ 195000
```

---

## Multiple Value Fields

You can add multiple calculations to Values area.

### Example: Sales and Count Side by Side

**Drag Sales to Values twice:**
1. First: Sum of Sales
2. Second: Count of Sales

**Result:**
```
Region │ Sum of Sales │ Count of Sales
───────┼──────────────┼───────────
East   │ 50000        │ 25
West   │ 45000        │ 18
North  │ 38000        │ 15
```

### Calculate Average Order Value

**With both Sum and Count visible:**
```
Average = Sum / Count
East: 50000 / 25 = $2000 per order
```

Or directly: Change one field to **Average**.

---

## Pivot Table Design and Formatting

### Apply Styles

**PivotTable Design Tab:**
- Choose from gallery of pre-designed styles
- Light, Medium, Dark color schemes
- Banded rows or columns

### Design Options

**PivotTable Design Tab → Layout:**

**Row headers:**
- In compact form (default)
- In outline form
- In tabular form

**Subtotals:**
- Show at top of group
- Show at bottom of group
- Don't show subtotals

**Grand totals:**
- Show for rows
- Show for columns
- Show for both
- Don't show

**Blank rows:**
- Insert blank line after each item (spacing)

### Example: Tabular vs Compact Form

**Compact (Default):**
```
Row Labels     │ Sales
───────────────┼──────
East           │ 50000
  Widget       │ 20000
  Gadget       │ 18000
  Tool         │ 12000
West           │ 45000
  Widget       │ 15000
  Gadget       │ 20000
```

**Tabular:**
```
Region │ Product │ Sales
───────┼─────────┼──────
East   │ Widget  │ 20000
East   │ Gadget  │ 18000
East   │ Tool    │ 12000
West   │ Widget  │ 15000
West   │ Gadget  │ 20000
```

---

## Refreshing Pivot Tables

Pivot Tables **don't update automatically** when source data changes.

### How to Refresh

**Method 1: Right-click**
1. Right-click anywhere in Pivot Table
2. Select **Refresh**

**Method 2: Ribbon**
1. PivotTable Analyze Tab → **Refresh**

**Method 3: Keyboard**
1. Press **Alt + F5**

### Refresh All Pivot Tables

If workbook has multiple Pivot Tables:
- PivotTable Analyze → **Refresh** dropdown → **Refresh All**

### Auto-Refresh on Open

**Steps:**
1. Right-click Pivot Table → **PivotTable Options**
2. Data tab
3. Check **Refresh data when opening the file**
4. Click OK

---

## Calculated Fields

Create custom calculations using existing fields.

### When to Use

- Formulas that apply to all rows (e.g., Profit = Sales - Cost)
- Calculations not in source data
- Derived metrics (e.g., Profit Margin %)

### How to Create

**Steps:**
1. Click in Pivot Table
2. **PivotTable Analyze Tab → Fields, Items, & Sets → Calculated Field**
3. Name: Enter field name (e.g., "Profit")
4. Formula: Enter formula (e.g., `=Sales - Cost`)
5. Click **Add**
6. Click **OK**

### Example: Profit Margin

**Source data has:**
- Sales
- Cost

**Create calculated field:**
```
Name: Profit Margin
Formula: =(Sales-Cost)/Sales
```

**Result in Pivot Table:**
```
Product │ Sales │ Cost  │ Profit Margin
────────┼───────┼───────┼──────────
Widget  │ 50000 │ 30000 │ 40%
Gadget  │ 45000 │ 25000 │ 44%
Tool    │ 38000 │ 20000 │ 47%
```

### Calculated Field Limitations

❌ Can't use in row/column/filter areas (only Values)
❌ Calculates on aggregated data (not row-by-row)
❌ Can't reference cells outside Pivot Table
✅ Can use +, -, *, /, ^, and basic functions

---

## Pivot Table Best Practices

### 1. Clean Source Data First
```
✅ Remove blank rows
✅ Ensure consistent headers
✅ Fix data types (text vs number)
✅ Remove merged cells
```

### 2. Use Tables for Source Data
```
✅ Insert → Table
✅ Pivot Table auto-expands when you add data
✅ No need to manually update range
```

### 3. Name Your Pivot Tables
```
✅ PivotTable Analyze → PivotTable Name
✅ Use descriptive names: "Sales_by_Region_2024"
✅ Easier to reference in formulas
```

### 4. Document Filters
```
✅ If sharing, note active filters
✅ Or clear all filters before sharing
✅ Add text box explaining what's included
```

### 5. Keep Source Data Separate
```
✅ Source data on one sheet
✅ Pivot Tables on other sheets
✅ Don't delete source data!
```

### 6. Refresh Before Important Decisions
```
✅ Always refresh before presenting
✅ Check refresh date (if shown)
✅ Verify source data is current
```

---

## Common Mistakes

### Mistake 1: Not Refreshing After Data Changes

```
❌ Problem: Source data updated, Pivot shows old values
✅ Solution: Right-click → Refresh (or Alt + F5)
```

### Mistake 2: Deleting Source Data

```
❌ Problem: Pivot Table breaks completely
✅ Solution: Keep source data sheet hidden, not deleted
```

### Mistake 3: Changing Source Range Manually

```
❌ Problem: New rows not included
✅ Solution: Use Tables, or update range via:
   Right-click → PivotTable Options → Data → Change Data Source
```

### Mistake 4: Using Pivot Tables for Small Data

```
❌ Overkill: 10 rows of data
✅ Better: Use SUMIF, COUNTIF, regular formulas
```

### Mistake 5: Too Many Fields

```
❌ Confusing: 5 row fields, 4 column fields
✅ Better: Start simple, add complexity gradually
```

### Mistake 6: Not Understanding Aggregation

```
❌ Problem: "Why doesn't my average match?"
✅ Understanding: Pivot averages the sums, not raw data
   Use raw data for true averages when needed
```

---

## Troubleshooting Guide

### Problem: Field List Disappeared

**Solution:**
- Right-click Pivot Table → **Show Field List**
- Or: PivotTable Analyze → **Field List**

### Problem: Can't Change Pivot Table

**Cause:** Sheet or workbook protected

**Solution:**
- Review Tab → **Unprotect Sheet**

### Problem: "Cannot group that selection"

**Causes:**
- Blank cells in date column
- Text mixed with dates
- Dates stored as text

**Solutions:**
- Fill blank cells
- Convert text to dates
- Filter out invalid entries first

### Problem: Count Instead of Sum

**Cause:** Numbers stored as text in source data

**Solution:**
- Fix source data (convert text to numbers)
- Refresh Pivot Table
- Change Value Field Settings to Sum if still needed

### Problem: #REF! Error

**Cause:** Source data deleted or moved

**Solution:**
- Right-click Pivot → PivotTable Options → Change Data Source
- Point to correct range

---

## Real-World Example: Sales Analysis

**Scenario:** Analyze 1000 sales transactions

**Source Data Columns:**
- Date
- Salesperson
- Region
- Product
- Quantity
- Unit Price
- Total Sales

**Analysis Goals:**
1. Total sales by region
2. Top performing salespeople
3. Monthly trends
4. Product mix by region

**Pivot Table Setup:**

**Analysis 1: Sales by Region**
```
Rows: Region
Values: Sum of Total Sales

Result:
Region │ Sales
───────┼────────
East   │ 450000
West   │ 380000
North  │ 325000
South  │ 290000
```

**Analysis 2: Top Salespeople**
```
Rows: Salesperson
Values: Sum of Total Sales
Sort: Largest to Smallest
Filter: Top 10

Result shows top 10 performers
```

**Analysis 3: Monthly Trends**
```
Rows: Date (grouped by Month)
Values: Sum of Total Sales

Add: Show Values As → % Difference From (Previous)

Result:
Month  │ Sales   │ % Change
───────┼─────────┼─────
Jan    │ 120000  │ --
Feb    │ 135000  │ +12.5%
Mar    │ 142000  │ +5.2%
```

**Analysis 4: Product Mix by Region**
```
Rows: Region
Columns: Product
Values: Sum of Total Sales

Result: Matrix showing each product's sales in each region
```

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Alt + N + V` | Create new Pivot Table |
| `Alt + F5` | Refresh Pivot Table |
| `Ctrl + -` | Remove field from Pivot Table |
| `Alt + ↓` | Open field dropdown |
| `Ctrl + Shift + *` | Select entire Pivot Table |
| `Alt + D + P` | Open PivotTable Wizard (legacy) |

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Pivot Tables summarize data automatically
- Four areas: Filters, Rows, Columns, Values
- Rows = left side labels
- Columns = top headers
- Values = numbers to calculate
- Default calculation is SUM
- Must refresh to see updated source data
- Right-click for most options

### Practice Deeply
- Creating basic Pivot Tables from scratch
- Dragging fields to different areas
- Rearranging fields (Rows vs Columns)
- Changing value calculations (Sum, Count, Average)
- Grouping dates by month, quarter, year
- Grouping numbers into ranges
- Adding and removing fields
- Filtering Pivot Tables (dropdown filters)
- Using slicers for visual filtering
- Formatting Pivot Tables (styles, layouts)
- Refreshing after source data changes
- Creating calculated fields for custom metrics
- Using "Show Values As" for percentages
- Interpreting Pivot Table results
- Starting with simple tables, adding complexity gradually

---

## Quick Reference: Common Pivot Table Layouts

### Sales by Region
```
Rows: Region
Values: Sum of Sales
```

### Sales by Month
```
Rows: Date (grouped by Month)
Values: Sum of Sales
```

### Product Performance Matrix
```
Rows: Product
Columns: Quarter
Values: Sum of Sales
```

### Customer Purchase Frequency
```
Rows: Customer Name
Values: Count of Order ID
```

### Average Order Value
```
Rows: Region or Product
Values: Average of Order Total
```

---

## Next Step

After this file, we move to:

**`14-pivot-charts.md`**
- Creating Pivot Charts from Pivot Tables
- Types of Pivot Charts (column, line, pie)
- Interactive filtering with Pivot Charts
- Syncing charts with Pivot Tables
- Formatting Pivot Charts
- Slicers and timelines for charts
- When to use Pivot Charts vs regular charts
