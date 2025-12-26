# Tables and Structured References

This file covers Excel Tables - a powerful feature that transforms ordinary ranges into dynamic, formatted, self-expanding data structures with special formula syntax called structured references.

---

## What are Excel Tables?

An **Excel Table** is a structured range of data with special properties and behaviors.

### Regular Range vs Table

**Regular Range:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Sales  │ Region │
  ├────────┼────────┼────────┤
2 │ John   │ 5000   │ East   │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 6500   │ West   │
  ├────────┼────────┼────────┤
4 │ Mike   │ 4800   │ North  │
  └────────┴────────┴────────┘

Just cells with data
```

**Excel Table:**
```
     A         B         C
  ┌════════╦════════╦════════╗
1 ║ Name ▼ ║ Sales ▼║ Region▼║ ← Headers with filters
  ╠════════╬════════╬════════╣
2 ║ John   ║ 5000   ║ East   ║
  ╠────────╬────────╬────────╣
3 ║ Sarah  ║ 6500   ║ West   ║
  ╠────────╬────────╬────────╣
4 ║ Mike   ║ 4800   ║ North  ║
  ╚════════╩════════╩════════╝

Formatted, with dropdown filters
Auto-expanding structure
Special formula syntax available
```

### Key Table Features

✅ **Auto-expansion** - Add row, table grows automatically
✅ **Built-in filtering** - Dropdown filters on every column
✅ **Structured references** - Use column names in formulas
✅ **Banded rows** - Alternating colors for readability
✅ **Total row** - Quick sum/average/count
✅ **Auto-formatting** - Consistent styling
✅ **Easy sorting** - Click header to sort
✅ **Named automatically** - Table gets unique name

---

## Creating a Table

### Method 1: Insert Tab (Recommended)

**Steps:**
1. Click anywhere in your data range
2. **Insert Tab → Table** (or **Ctrl + T**)
3. Verify range in dialog
4. Check **My table has headers**
5. Click **OK**

**Dialog:**
```
┌─────────────────────────────────────┐
│ Create Table                        │
├─────────────────────────────────────┤
│ Where is the data for your table?   │
│                                     │
│ =$A$1:$C$50                         │
│                                     │
│ ☑ My table has headers              │
│                                     │
│ [OK] [Cancel]                       │
└─────────────────────────────────────┘
```

### Method 2: Home Tab

**Steps:**
1. Select data range
2. **Home Tab → Format as Table**
3. Choose table style
4. Verify range and headers
5. Click **OK**

### Method 3: Keyboard Shortcut

**Steps:**
1. Click in data range
2. Press **Ctrl + T** (or **Ctrl + L**)
3. Verify and confirm

### What Happens After Creation

**Immediately:**
- Data formatted with table style
- Filter dropdowns appear on headers
- Table Design tab appears in ribbon
- Table gets default name (Table1, Table2, etc.)

**Visual transformation:**
```
Before:
Plain cells, no formatting

After:
┌════════╦════════╦════════╗
║ Name ▼ ║ Sales ▼║ Region▼║
╠════════╬════════╬════════╣
║ John   ║ 5000   ║ East   ║
╠────────╬────────╬────────╣
║ Sarah  ║ 6500   ║ West   ║
╚════════╩════════╩════════╝

Styled, filtered, enhanced!
```

---

## Table Components

### 1. Header Row

```
┌════════╦════════╦════════╗
║ Name ▼ ║ Sales ▼║ Region▼║ ← Header Row
╠════════╬════════╬════════╣

Features:
- Filter dropdown buttons
- Bold formatting
- Different background color
- Cannot delete (only hide)
```

### 2. Data Rows

```
╠════════╬════════╬════════╣
║ John   ║ 5000   ║ East   ║ ← Data Row
╠────────╬────────╬────────╣
║ Sarah  ║ 6500   ║ West   ║ ← Data Row
╠────────╬────────╬────────╣

Features:
- Banded rows (alternating colors)
- Auto-extend when you add data
- Structured reference formulas
```

### 3. Total Row

```
╠────────╬────────╬────────╣
║ Mike   ║ 4800   ║ North  ║
╠════════╬════════╬════════╣
║ Total  ║ 16300  ║        ║ ← Total Row
╚════════╩════════╩════════╝

Features:
- Optional (toggle on/off)
- Dropdown to choose function (Sum, Average, Count, etc.)
- Automatically calculates
```

### 4. Resize Handle

```
╠────────╬────────╬────────╣
║ Mike   ║ 4800   ║ North  ║
╚════════╩════════╩═══════╗◄ Resize handle
                           ║
                           ▼

Drag to expand table manually
```

---

## Structured References

**Structured References** use column names instead of cell addresses.

### Basic Syntax

**Regular formula:**
```excel
=B2*C2
```

**Structured reference:**
```excel
=[@Sales]*[@Price]

@ = "this row"
[Sales] = column name
```

### Syntax Components

**[@ColumnName]** - Current row, specific column
```excel
=[@Sales]*[@Quantity]

In row 2: Uses Sales from row 2
In row 3: Uses Sales from row 3
Relative reference that adjusts per row
```

**[ColumnName]** - Entire column
```excel
=SUM(SalesData[Sales])

Sums entire Sales column in SalesData table
```

**[[#This Row],[ColumnName]]** - Explicit this row
```excel
=[[#This Row],[Sales]]*[[#This Row],[Quantity]]

Same as [@Sales]*[@Quantity]
More explicit, less common
```

**[#All]** - Entire table including headers
```excel
=ROWS(SalesData[#All])

Counts rows including header
```

**[#Data]** - Data rows only (no header, no total)
```excel
=SUM(SalesData[#Data])

Sums all data, excludes total row
```

**[#Headers]** - Header row only
```excel
=SalesData[#Headers]

References header row
```

**[#Totals]** - Total row only
```excel
=SalesData[#Totals]

References total row
```

### Visual Examples

**Table: SalesData**
```
     A         B         C         D
  ┌════════╦════════╦════════╦════════╗
1 ║ Name   ║ Qty    ║ Price  ║ Total  ║ ← [#Headers]
  ╠════════╬════════╬════════╬════════╣
2 ║ Widget ║ 10     ║ 25.00  ║ 250.00 ║ ← [#Data]
  ╠────────╬────────╬────────╬────────╣
3 ║ Gadget ║ 15     ║ 30.00  ║ 450.00 ║ ← [#Data]
  ╠────────╬────────╬────────╬────────╣
4 ║ Tool   ║ 8      ║ 20.00  ║ 160.00 ║ ← [#Data]
  ╠════════╬════════╬════════╬════════╣
5 ║        ║        ║ Total  ║ 860.00 ║ ← [#Totals]
  ╚════════╩════════╩════════╩════════╝
                               ↑
                          [#All] = entire table
```

**Formulas:**

Cell D2:
```excel
=[@Qty]*[@Price]

Result: 10 * 25.00 = 250.00
```

Sum all sales:
```excel
=SUM(SalesData[Total])

Result: 250.00 + 450.00 + 160.00 = 860.00
```

Average price:
```excel
=AVERAGE(SalesData[Price])

Result: (25+30+20)/3 = 25.00
```

---

## Calculated Columns

**Calculated Column** = Formula column that auto-fills in tables.

### Creating Calculated Column

**Steps:**
1. Click first cell in empty column next to table
2. Type formula using structured references
3. Press **Enter**

**What happens:**
- Formula automatically copies down entire column
- All rows get the same formula (relative to their row)
- New rows automatically get formula

### Example

**Table before:**
```
┌════════╦════════╦════════╗
║ Qty    ║ Price  ║        ║
╠════════╬════════╬════════╣
║ 10     ║ 25.00  ║        ║
╠────────╬────────╬────────╣
║ 15     ║ 30.00  ║        ║
╠────────╬────────╬────────╣
║ 8      ║ 20.00  ║        ║
╚════════╩════════╩════════╝
```

**Type in D2:**
```excel
=[@Qty]*[@Price]
```

**Press Enter:**
```
┌════════╦════════╦════════╦════════╗
║ Qty    ║ Price  ║ Total  ║
╠════════╬════════╬════════╬════════╣
║ 10     ║ 25.00  ║ 250.00 ║ ← Auto-filled
╠────────╬────────╬────────╬────────╣
║ 15     ║ 30.00  ║ 450.00 ║ ← Auto-filled
╠────────╬────────╬────────╬────────╣
║ 8      ║ 20.00  ║ 160.00 ║ ← Auto-filled
╚════════╩════════╩════════╩════════╝

Excel automatically:
1. Added column header "Total"
2. Copied formula to all rows
3. Will auto-fill for new rows
```

### Editing Calculated Column

**Change one cell:**
- Edit formula in any cell
- Excel asks: "Update all cells in this column?"
- Click **Yes** → All cells updated
- Click **No** → Only that cell changes (breaks calculated column)

**Visual prompt:**
```
┌──────────────────────────────────────┐
│ Do you want to replace the existing  │
│ formula with this one?               │
│                                      │
│ [Yes] [No] [Cancel]                  │
└──────────────────────────────────────┘
```

### Benefits of Calculated Columns

```
✅ No need to copy formula down
✅ Automatically extends to new rows
✅ Consistent formulas (no variations)
✅ Easy to understand (structured references)
✅ Self-documenting
```

---

## Total Row

Add summary calculations at bottom of table.

### Enabling Total Row

**Steps:**
1. Click anywhere in table
2. **Table Design Tab → Total Row** (check box)

Or right-click table → **Table → Total Row**

**Result:**
```
┌════════╦════════╦════════╗
║ Qty    ║ Price  ║ Total  ║
╠════════╬════════╬════════╣
║ 10     ║ 25.00  ║ 250.00 ║
╠────────╬────────╬────────╣
║ 15     ║ 30.00  ║ 450.00 ║
╠────────╬────────╬────────╣
║ 8      ║ 20.00  ║ 160.00 ║
╠════════╬════════╬════════╣
║ Total  ║        ║ 860.00 ║ ← Total Row added
╚════════╩════════╩════════╝
```

### Choosing Calculation

**Steps:**
1. Click cell in total row
2. Click dropdown arrow
3. Select function

**Available functions:**
```
┌──────────────────┐
│ None             │
│ Average          │
│ Count            │
│ Count Numbers    │
│ Max              │
│ Min              │
│ Sum              │ ← Default for numbers
│ StdDev           │
│ Var              │
│ More Functions...│
└──────────────────┘
```

### Total Row Formula

**Behind the scenes:**
```excel
Total row uses SUBTOTAL function:

=SUBTOTAL(109,[Total])

109 = SUM (ignore hidden rows)
[Total] = column reference

Other function codes:
102 = COUNT
103 = COUNTA
104 = MAX
105 = MIN
106 = PRODUCT
107 = STDEV
109 = SUM (default)
110 = VAR
```

**Why SUBTOTAL?**
- Respects filters (only calculates visible rows)
- If you filter table, total updates automatically
- Regular SUM would include hidden rows

### Multiple Total Row Calculations

Can have different calculation per column:
```
┌════════╦════════╦════════╦════════╗
║ Name   ║ Qty    ║ Price  ║ Total  ║
╠════════╬════════╬════════╬════════╣
║ Widget ║ 10     ║ 25.00  ║ 250.00 ║
╠────────╬────────╬────────╬────────╣
║ Gadget ║ 15     ║ 30.00  ║ 450.00 ║
╠════════╬════════╬════════╬════════╣
║ Total  ║ 25     ║ Avg:   ║ 860.00 ║
╚════════╩════════╩  27.50 ╩════════╝
            ↑         ↑         ↑
          SUM    AVERAGE     SUM
```

---

## Table Design and Formatting

### Table Styles

**Accessing styles:**
1. Click in table
2. **Table Design Tab → Table Styles** gallery
3. Choose style

**Categories:**
- Light (subtle colors)
- Medium (moderate colors)
- Dark (bold colors)

**Custom styles:**
- Right-click style → **Duplicate**
- Modify colors, fonts, borders
- Save as custom style

### Table Style Options

**Table Design Tab → Table Style Options:**

```
☑ Header Row         Show/hide header
☑ Total Row          Show/hide total
☑ Banded Rows        Alternating row colors
☐ First Column       Bold/highlight first column
☐ Last Column        Bold/highlight last column
☐ Banded Columns     Alternating column colors
☐ Filter Button      Show/hide filter dropdowns
```

### Customizing Table Appearance

**Banded Rows (Recommended):**
```
☑ Banded Rows

┌════════╦════════╗
║ John   ║ 5000   ║ ← Light
╠────────╬────────╣
║ Sarah  ║ 6500   ║ ← Dark
╠────────╬────────╣
║ Mike   ║ 4800   ║ ← Light
╚════════╩════════╝

Easier to read across rows
```

**Banded Columns:**
```
☑ Banded Columns

┌════════╦════════╦════════╗
║ Name   ║ Sales  ║ Region ║
║   ↓    ║   ↓    ║   ↓    ║
║ Light  ║ Dark   ║ Light  ║
╚════════╩════════╩════════╝

Easier to read down columns
```

**First/Last Column Emphasis:**
```
☑ First Column       ☑ Last Column

┌════════╦════════╦════════╗
║►NAME   ║ Sales  ║ Total◄ ║
║►John   ║ 5000   ║ 400◄   ║
║►Sarah  ║ 6500   ║ 520◄   ║
╚════════╩════════╩════════╝

Bold font, different color
```

---

## Working with Tables

### Adding Rows

**Method 1: Tab from last cell**
1. Click last cell in table
2. Press **Tab**
3. New row appears

**Method 2: Type below table**
1. Click cell immediately below table
2. Type data
3. Press **Enter**
4. Table expands to include new row

**Method 3: Drag resize handle**
1. Click resize handle (bottom-right corner)
2. Drag down
3. New rows added

**Visual:**
```
Before:
╠────────╬────────╣
║ Mike   ║ 4800   ║
╚════════╩═══════╗◄ Grab and drag down
                 ║

After:
╠────────╬────────╣
║ Mike   ║ 4800   ║
╠────────╬────────╣
║        ║        ║ ← New row
╠────────╬────────╣
║        ║        ║ ← New row
╚════════╩════════╝
```

### Adding Columns

**Method 1: Type next to table**
1. Click cell immediately right of table
2. Type header
3. Press **Enter**
4. Table expands to include new column

**Method 2: Drag resize handle right**

**Result:**
- Calculated columns auto-fill
- Formatting applies automatically
- Structured references update

### Deleting Rows/Columns

**Delete rows:**
1. Select row(s)
2. Right-click → **Delete → Table Rows**

Or: Home Tab → Delete → Delete Table Rows

**Delete columns:**
1. Select column(s)
2. Right-click → **Delete → Table Columns**

⚠️ **Note:** Can't delete header row (only hide)

### Inserting Rows/Columns

**Insert row:**
1. Right-click row
2. **Insert → Table Rows Above**

**Insert column:**
1. Right-click column
2. **Insert → Table Columns to the Left**

### Selecting in Tables

**Select column:**
- Click column header once (selects data only)
- Click again (includes header)
- Click third time (includes total row if visible)

**Select row:**
- Click row number (if visible)
- Or select first cell, Shift+End

**Select entire table:**
- Click table selector (top-left corner)
- Or Ctrl + A (when in table)

---

## Sorting and Filtering Tables

### Sorting

**Quick sort:**
1. Click dropdown in column header
2. Choose:
   - **Sort A to Z** (ascending)
   - **Sort Z to A** (descending)

**Multi-level sort:**
1. **Data Tab → Sort**
2. Sort by: Column1
3. Then by: Column2
4. Then by: Column3
5. OK

**Example:**
```
Sort by: Region (A to Z)
Then by: Sales (Largest to Smallest)

Result:
Region  Name   Sales
East    Sarah  6500
East    John   5000
West    Mike   4800
```

### Filtering

**Filter dropdown automatically available on headers.**

**Text filters:**
```
┌──────────────────────┐
│ Text Filters         │
├──────────────────────┤
│ Equals               │
│ Does Not Equal       │
│ Begins With          │
│ Ends With            │
│ Contains             │
│ Does Not Contain     │
└──────────────────────┘
```

**Number filters:**
```
┌──────────────────────┐
│ Number Filters       │
├──────────────────────┤
│ Equals               │
│ Greater Than         │
│ Less Than            │
│ Between              │
│ Top 10               │
│ Above Average        │
│ Below Average        │
└──────────────────────┘
```

**Date filters:**
```
┌──────────────────────┐
│ Date Filters         │
├──────────────────────┤
│ Tomorrow             │
│ Today                │
│ Yesterday            │
│ This Week            │
│ Last Month           │
│ This Quarter         │
│ Last Year            │
│ Custom...            │
└──────────────────────┘
```

**Checkbox filtering:**
```
┌──────────────────────┐
│ Region        ▼      │
├──────────────────────┤
│ ☑ (Select All)       │
│ ☑ East               │
│ ☑ West               │
│ ☐ North              │
│ ☑ South              │
└──────────────────────┘

Uncheck to hide
```

### Clear Filters

**Clear from one column:**
- Click filter dropdown
- **Clear Filter from "Column"**

**Clear all filters:**
- **Data Tab → Clear**

**Visual indicator:**
```
Filtered column shows funnel icon:
║ Sales ▼≡║ ← Filter active

No filter:
║ Sales ▼ ║ ← No filter
```

---

## Table Names and References

### Renaming a Table

**Steps:**
1. Click anywhere in table
2. **Table Design Tab → Table Name** box
3. Type new name
4. Press **Enter**

**Naming rules:**
- Start with letter or underscore
- No spaces (use underscores: Sales_Data)
- No cell references (A1, XFD1, etc.)
- Must be unique in workbook

**Example:**
```
Default: Table1
Better:  SalesData or Sales_2024
```

### Using Table Names in Formulas

**Reference entire table:**
```excel
=SUM(SalesData[Sales])

Sums Sales column in SalesData table
```

**Reference from another sheet:**
```excel
=AVERAGE(SalesData[Price])

Works from any sheet
No need for sheet reference!
```

**Count rows:**
```excel
=ROWS(SalesData[#Data])

Counts data rows (excludes header/total)
```

### External References

**From different workbook:**
```excel
='[Budget.xlsx]Summary'!SalesData[Sales]

Syntax:
'[Workbook]Sheet'!TableName[Column]
```

---

## Converting Between Tables and Ranges

### Convert Table to Range

**Steps:**
1. Click in table
2. **Table Design Tab → Convert to Range**
3. Confirm

**What happens:**
- Formatting remains
- Filter dropdowns removed
- Structured references converted to cell references
- No longer auto-expands
- Total row becomes regular cells

**When to convert:**
- Need to delete specific rows (tables have restrictions)
- Exporting to non-Excel format
- Compatibility with very old Excel
- Specific formatting requirements

### Convert Range to Table

**Steps:**
1. Click in range
2. **Insert Tab → Table** (Ctrl + T)
3. Confirm range and headers

**Benefits:**
- Gain all table features
- Auto-expansion
- Structured references
- Built-in filtering
- Professional appearance

---

## Advanced Table Features

### Remove Duplicates

**Built into tables:**

**Steps:**
1. Click in table
2. **Table Design Tab → Remove Duplicates**
3. Select columns to check
4. Click **OK**

**Dialog:**
```
┌─────────────────────────────────────┐
│ Remove Duplicates                   │
├─────────────────────────────────────┤
│ Select columns:                     │
│ ☑ Name                              │
│ ☑ Email                             │
│ ☐ Phone                             │
│                                     │
│ [OK] [Cancel]                       │
└─────────────────────────────────────┘

Checks Name+Email combination
Removes duplicate rows
```

### Slicer for Tables

**Visual filtering (like Pivot Tables):**

**Steps:**
1. Click in table
2. **Table Design Tab → Insert Slicer**
3. Select fields
4. Click **OK**

**Result:**
```
┌────────────────────┐
│ Region             │
├────────────────────┤
│ [East]  [West]     │
│ [North] [South]    │
└────────────────────┘

Click buttons to filter table
```

**Benefits:**
- Visual, intuitive
- See what's selected
- Easy for non-technical users
- Can control multiple tables/Pivot Tables

### Table Relationships

**Connect related tables (Power Pivot):**

**Example:**
- Orders table (OrderID, CustomerID, Amount)
- Customers table (CustomerID, Name, Region)

**Create relationship:**
1. **Data Tab → Relationships**
2. Click **New**
3. Table: Orders
4. Column: CustomerID
5. Related Table: Customers
6. Related Column: CustomerID
7. OK

**Use in formulas:**
```excel
=RELATED(Customers[Region])

From Orders table, get related Customer Region
```

⚠️ **Note:** Requires Data Model (Power Pivot)

---

## Structured Reference Examples

### Example 1: Sales Calculation

**Table: Orders**
```
┌════════╦════════╦════════╦════════╗
║ Product║ Qty    ║ Price  ║ Total  ║
╠════════╬════════╬════════╬════════╣
║ Widget ║ 10     ║ 25.00  ║   ?    ║
╚════════╩════════╩════════╩════════╝
```

**Formula in Total column:**
```excel
=[@Qty]*[@Price]

Clear and self-documenting
```

### Example 2: Conditional Calculation

**Table: Sales**
```
┌════════╦════════╦════════╦════════╗
║ Amount ║ Target ║ Bonus  ║
╠════════╬════════╬════════╬════════╣
║ 5000   ║ 4000   ║   ?    ║
╚════════╩════════╩════════╩════════╝
```

**Formula in Bonus column:**
```excel
=IF([@Amount]>[@Target],[@Amount]*0.05,0)

If sales exceed target, 5% bonus
```

### Example 3: Lookup Within Table

**Table: Products**
```
┌════════╦════════╦════════╗
║ Code   ║ Price  ║ Disc   ║
╠════════╬════════╬════════╣
║ A100   ║ 25.00  ║   ?    ║
╚════════╩════════╩════════╝
```

**Formula in Disc column:**
```excel
=IF([@Price]>100,[@Price]*0.10,[@Price]*0.05)

10% discount if price > $100, otherwise 5%
```

### Example 4: Reference from Outside Table

**From cell outside Orders table:**
```excel
=SUM(Orders[Total])

Sums all values in Total column
```

**Average order value:**
```excel
=AVERAGE(Orders[Total])
```

**Count of orders:**
```excel
=ROWS(Orders[#Data])

Or: =COUNTA(Orders[Product])
```

### Example 5: Multiple Column Reference

**Entire table (all columns, all data):**
```excel
=ROWS(Orders[#All])

Counts all rows including header
```

**Two columns:**
```excel
=SUM(Orders[[Qty]:[Price]])

Unusual, but possible
Sums Qty column + Price column
```

---

## Table Best Practices

### When to Use Tables

```
✅ Use tables for:
- Lists that grow over time
- Data you filter/sort frequently
- Datasets for analysis
- Data entry forms
- Dashboards and reports
- Any structured data
```

### When NOT to Use Tables

```
❌ Avoid tables for:
- Single-use data
- Complex layouts with merged cells
- Data shared with Excel 2003 or earlier
- Templates where structure must not change
- When you need specific cell formatting per cell
```

### Naming Conventions

```
✅ Good table names:
Sales_2024
Customer_List
Inventory_Tracking
Monthly_Budget

❌ Poor table names:
Table1
data
tbl
List
```

### Design Guidelines

```
✅ Keep headers clear and concise
✅ One data type per column
✅ No blank rows within table
✅ No blank columns within table
✅ Use calculated columns for formulas
✅ Enable total row for summaries
✅ Use banded rows for readability
```

### Formula Guidelines

```
✅ Use structured references ([@Column])
✅ Use table names in external references
✅ Let calculated columns auto-fill
✅ Test formulas before applying to all rows
✅ Avoid absolute references ($A$1) in table formulas
```

---

## Troubleshooting Tables

### Problem: Can't Delete Rows

**Symptom:** Delete option grayed out

**Cause:** Trying to use regular delete on table row

**Solution:**
- Right-click → **Delete → Table Rows**
- Or select row, Home Tab → Delete → Delete Table Rows

### Problem: Formula Not Auto-Filling

**Cause:** Calculated column feature disabled or broken

**Solution:**
```
1. Check if formula bar shows structured reference
2. File → Options → Proofing → AutoCorrect Options
3. AutoFormat As You Type tab
4. Check "Fill formulas in tables to create calculated columns"
5. OK
```

### Problem: Structured Reference Shows Error

**Symptom:** `=[@Sales]` shows #REF!

**Causes:**
- Column deleted
- Table corrupted
- Column renamed

**Solutions:**
- Check if column exists
- Update formula with correct column name
- Recreate table if necessary

### Problem: Table Won't Expand

**Cause:** Cell below/right of table contains data

**Solution:**
- Clear cells around table
- Or manually resize table
- Or convert to range, add data, convert back

### Problem: Filter Not Working

**Symptom:** Filter dropdown missing or non-functional

**Solution:**
```
1. Table Design Tab
2. Check "Filter Button" in Table Style Options
3. Or Data Tab → Filter (toggle off/on)
```

### Problem: Total Row Calculates Wrong

**Cause:** Using SUM instead of SUBTOTAL, or filtered rows

**Solution:**
```
Total row uses SUBTOTAL automatically
If you edited it manually, restore:
1. Click cell in total row
2. Select dropdown
3. Choose appropriate function
4. SUBTOTAL respects filters automatically
```

### Problem: Structured Reference Too Long

**Symptom:** Formula like `=Sales_Data_2024_Q1[[#This Row],[Total Amount]]` is verbose

**Solution:**
```
1. Shorten table name (Sales_Data_2024_Q1 → Sales)
2. Shorten column names (Total Amount → Total)
3. Use @ syntax: [@Total] instead of [[#This Row],[Total]]
```

### Problem: Table Converted to Range Accidentally

**Solution:**
```
Undo (Ctrl + Z) immediately if possible
Or:
1. Select the data range
2. Ctrl + T to recreate table
3. Formulas will need manual fixing (structured refs lost)
```

---

## Common Patterns and Use Cases

### Pattern 1: Simple Data Entry Table

**Structure:**
```
┌════════╦════════╦════════╦════════╗
║ Date   ║ Item   ║ Amount ║ Category║
╠════════╬════════╬════════╬════════╣
║        ║        ║        ║        ║ ← Empty row for entry
╚════════╩════════╩════════╩════════╝

Total row showing sum of Amount
Filter/sort as needed
```

**Use case:** Expense tracking, log entries, simple records

### Pattern 2: Calculated Results Table

**Structure:**
```
┌════════╦════════╦════════╦════════╦════════╗
║ Product║ Qty    ║ Price  ║ Total  ║ Tax    ║
╠════════╬════════╬════════╬════════╬════════╣
║ Widget ║ 10     ║ 25.00  ║ =[@Qty]*[@Price] ║ =[@Total]*0.08 ║
╚════════╩════════╩════════╩════════╩════════╝

Calculated columns auto-compute
```

**Use case:** Order forms, invoices, price lists

### Pattern 3: Lookup Table

**Structure:**
```
┌════════╦════════╦════════╗
║ Code   ║ Name   ║ Price  ║
╠════════╬════════╬════════╣
║ A100   ║ Widget ║ 25.00  ║
║ A200   ║ Gadget ║ 30.00  ║
║ A300   ║ Tool   ║ 20.00  ║
╚════════╩════════╩════════╝

Used in VLOOKUP/XLOOKUP formulas
Filter to find specific items
```

**Use case:** Product catalogs, employee lists, reference data

### Pattern 4: Summary Dashboard Source

**Structure:**
```
┌════════╦════════╦════════╦════════╗
║ Region ║ Sales  ║ Costs  ║ Profit ║
╠════════╬════════╬════════╬════════╣
║ East   ║ 50000  ║ 30000  ║ =[@Sales]-[@Costs] ║
║ West   ║ 45000  ║ 28000  ║ =[@Sales]-[@Costs] ║
╚════════╩════════╩════════╩════════╝

Dashboard formulas reference:
=SUM(RegionData[Profit])
=AVERAGE(RegionData[Sales])
```

**Use case:** Reports, KPI tracking, dashboards

### Pattern 5: Running Balance Table

**Structure:**
```
┌════════╦════════╦════════╦════════╗
║ Date   ║ Debit  ║ Credit ║ Balance║
╠════════╬════════╬════════╬════════╣
║ 1/1    ║ 100    ║ 0      ║ =[@Debit]-[@Credit] ║
║ 1/2    ║ 0      ║ 50     ║ =...running total... ║
╚════════╩════════╩════════╩════════╝
```

**Note:** Running totals tricky in tables (need special formula)

**Use case:** Bank statements, inventory tracking

---

## Performance Considerations

### Large Tables (10,000+ Rows)

**Best practices:**
```
✅ Turn off calculated columns if not needed
✅ Use manual calculation (Formulas → Calculation Options)
✅ Avoid volatile functions (INDIRECT, OFFSET, TODAY)
✅ Consider Power Query for very large datasets
✅ Close unnecessary workbooks
✅ Save as .xlsb (binary) for faster load times
```

### Many Tables in Workbook

**Optimization:**
```
✅ Limit to 5-10 tables per sheet
✅ Use separate sheets for large tables
✅ Consider combining related tables
✅ Remove unused tables
✅ Convert to ranges if not actively using table features
```

### Formula Performance

**Faster:**
```
✅ =[@Qty]*[@Price]
✅ =SUM(Sales[Amount])
✅ Simple structured references
```

**Slower:**
```
❌ =SUMPRODUCT((Sales[Region]="East")*(Sales[Amount]))
❌ Complex array formulas in calculated columns
❌ Nested INDIRECT with structured references
```

---

## Tables vs Other Excel Features

### Tables vs Named Ranges

| Feature | Tables | Named Ranges |
|---------|--------|--------------|
| **Auto-expand** | ✅ Yes | ❌ No (unless dynamic formula) |
| **Filtering** | ✅ Built-in | ❌ Manual setup |
| **Formatting** | ✅ Automatic | ❌ Manual |
| **Formulas** | Structured references | Regular references |
| **Total row** | ✅ Built-in | ❌ Manual |
| **Flexibility** | Data only | Any cell/range/constant |
| **Use case** | Datasets | Constants, single cells |

### Tables vs Pivot Tables

| Feature | Tables | Pivot Tables |
|---------|--------|--------------|
| **Purpose** | Store data | Summarize data |
| **Editing** | ✅ Edit individual cells | ❌ Can't edit values |
| **Formulas** | ✅ Yes | ❌ Calculated fields only |
| **Size** | Unlimited rows | Summarized (fewer rows) |
| **Filtering** | Show/hide rows | Aggregate filtered data |
| **Best for** | Data entry, storage | Analysis, reporting |

**Workflow:** Table → Pivot Table
```
1. Store data in Table
2. Create Pivot Table from Table
3. Table updates → Refresh Pivot
```

### Tables vs Lists (Excel Online/SharePoint)

| Feature | Excel Tables | SharePoint Lists |
|---------|-------------|------------------|
| **Location** | Excel file | SharePoint site |
| **Collaboration** | Limited | ✅ Real-time |
| **Permissions** | File-level | Item-level |
| **Workflows** | ❌ No | ✅ Power Automate |
| **Forms** | Manual | ✅ Built-in |
| **Mobile** | Excel app | SharePoint app |

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + T` | Create table from selection |
| `Ctrl + L` | Create table (alternative) |
| `Ctrl + Shift + L` | Toggle filters on/off |
| `Alt + ↓` | Open filter dropdown (in header) |
| `Ctrl + Space` | Select table column |
| `Shift + Space` | Select table row |
| `Ctrl + A` | Select entire table |
| `Tab` | Move to next cell (creates new row at end) |
| `Shift + Tab` | Move to previous cell |
| `Ctrl + Shift + +` | Insert table row/column |

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Ctrl + T creates a table
- Tables auto-expand when you add data
- [@ColumnName] = this row's value
- [ColumnName] = entire column
- Total row uses SUBTOTAL (respects filters)
- Tab from last cell creates new row
- Calculated columns auto-fill formulas
- Tables have built-in filtering
- Structured references use column names
- Tables require headers

### Practice Deeply
- Creating tables from data ranges (Ctrl + T)
- Adding rows and columns to tables
- Using Tab to add new rows
- Creating calculated columns with structured references
- Writing formulas with [@ColumnName] syntax
- Enabling and using the total row
- Sorting and filtering tables
- Renaming tables (Table Design → Table Name)
- Using table names in formulas from other sheets
- Converting tables to ranges and back
- Applying table styles
- Working with table style options (banded rows, etc.)
- Using filter dropdowns effectively
- Creating slicers for tables
- Removing duplicates from tables
- Understanding when structured references update
- Troubleshooting #REF! errors in table formulas
- Testing formulas before applying to all rows

---

## Quick Reference: Structured Reference Syntax

### Basic Patterns

```excel
[@Sales]
Single column, current row

[Sales]
Entire Sales column

Sales[Amount]
Amount column in Sales table (from outside)

Sales[[#Headers],[Region]]
Header of Region column

Sales[#Totals]
Entire total row

Sales[#Data]
All data rows (no header, no total)

Sales[#All]
Everything (headers + data + totals)

[@[Sales Amount]]
Column name with space (brackets required)

Sales[[Amount]:[Total]]
Multiple columns (Amount through Total)
```

### Common Formula Examples

```excel
Sum entire column:
=SUM(Sales[Amount])

Average this row:
=([@Revenue]-[@Costs])/[@Revenue]

Count rows:
=ROWS(Sales[#Data])

Conditional sum:
=SUMIF(Sales[Region],"East",Sales[Amount])

Lookup in table:
=VLOOKUP([@Code],Products[[Code]:[Price]],2,FALSE)

Reference from other sheet:
=Sales[Amount]*0.08
```

---

## Checklist: Creating Effective Tables

Before creating table:
```
☐ Data has clear headers in first row
☐ No blank rows within data
☐ No blank columns within data
☐ Each column has consistent data type
☐ Headers are unique (no duplicates)
☐ Data range is contiguous
```

After creating table:
```
☐ Rename table to something descriptive
☐ Verify all data included
☐ Check calculated columns work correctly
☐ Enable total row if needed
☐ Choose appropriate table style
☐ Enable/disable banded rows as preferred
☐ Test filtering and sorting
☐ Remove filter buttons if not needed
☐ Document table purpose (comment or separate doc)
```

When using tables in formulas:
```
☐ Use structured references ([@Column])
☐ Reference table by name from other sheets
☐ Test formulas with filtered data
☐ Verify calculated columns update properly
☐ Use appropriate specifiers (#Data, #All, etc.)
☐ Keep structured references readable
☐ Document complex formulas
```

---

## Real-World Example: Sales Tracking System

### Setup

**Table: SalesData**
```
┌════════╦════════╦════════╦════════╦════════╦════════╗
║ Date   ║ Rep    ║ Product║ Qty    ║ Price  ║ Total  ║
╠════════╬════════╬════════╬════════╬════════╬════════╣
║ 1/5/24 ║ John   ║ Widget ║ 10     ║ 25.00  ║ 250.00 ║
╠────────╬────────╬────────╬────────╬────────╬────────╣
║ 1/7/24 ║ Sarah  ║ Gadget ║ 15     ║ 30.00  ║ 450.00 ║
╠────────╬────────╬────────╬────────╬────────╬────────╣
║ 1/8/24 ║ Mike   ║ Tool   ║ 8      ║ 20.00  ║ 160.00 ║
╠════════╬════════╬════════╬════════╬════════╬════════╣
║ Total  ║        ║        ║ 33     ║        ║ 860.00 ║
╚════════╩════════╩════════╩════════╩════════╩════════╝
```

**Calculated Column (Total):**
```excel
=[@Qty]*[@Price]
```

### Dashboard (Separate Sheet)

**Summary Metrics:**
```
Total Sales:      =SUM(SalesData[Total])
Average Order:    =AVERAGE(SalesData[Total])
Number of Orders: =ROWS(SalesData[#Data])
Top Salesperson:  =INDEX(SalesData[Rep],MATCH(MAX(...),...))
```

**Analysis:**
```
Sales by Rep:     =SUMIF(SalesData[Rep],A2,SalesData[Total])
Sales by Product: =SUMIF(SalesData[Product],B2,SalesData[Total])
This Month:       =SUMIFS(SalesData[Total],SalesData[Date],">="&DATE(2024,1,1))
```

### Benefits
- ✅ New sales automatically included (table auto-expands)
- ✅ Dashboard formulas never break (structured references)
- ✅ Easy to filter (by rep, product, date)
- ✅ Professional appearance
- ✅ Total row updates automatically
- ✅ Can add slicers for interactive filtering

---

## Next Step

After this file, we move to:

**`19-array-formulas-and-spill.md`**
- Understanding dynamic arrays
- Spill behavior and spill range
- Array formula basics
- SORT, FILTER, UNIQUE functions
- SEQUENCE and RANDARRAY
- Array operations
- Troubleshooting #SPILL! errors
- Legacy array formulas (Ctrl+Shift+Enter)
