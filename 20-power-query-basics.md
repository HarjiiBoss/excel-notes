# Power Query Basics

This file covers Power Query (Get & Transform) - Excel's powerful data import and transformation tool that lets you clean, reshape, and combine data from multiple sources without formulas.

---

## What is Power Query?

**Power Query** is Excel's built-in ETL (Extract, Transform, Load) tool.

### Purpose
- **Extract:** Import data from various sources
- **Transform:** Clean and reshape data
- **Load:** Bring transformed data into Excel

### Why Power Query?

**Without Power Query:**
```
1. Import messy CSV
2. Manually clean data (find/replace, split columns, remove blanks)
3. Fix data types (text to numbers, dates)
4. Repeat EVERY time you refresh data
❌ Time-consuming, error-prone
```

**With Power Query:**
```
1. Import data once
2. Build transformation steps
3. Save query
4. Click Refresh → All transformations reapply automatically
✅ Repeatable, reliable, fast
```

### Visual Concept

```
┌─────────────────────────────────────────────────────┐
│                   POWER QUERY                       │
│                                                     │
│  Sources          Transform              Load       │
│  ┌──────┐         ┌──────┐              ┌──────┐  │
│  │ CSV  │────────>│Filter│────────────> │Sheet │  │
│  │ Web  │         │Clean │              │Table │  │
│  │ DB   │         │Split │              │Model │  │
│  └──────┘         └──────┘              └──────┘  │
│                       ↓                             │
│                  Each step saved                    │
│                  Replay on refresh                  │
└─────────────────────────────────────────────────────┘
```

---

## Accessing Power Query

### Where to Find It

**Data Tab → Get & Transform Data section:**

```
┌──────────────────────────────────────┐
│ Data Tab                             │
├──────────────────────────────────────┤
│ Get Data ▼                           │
│   From File ▶                        │
│   From Database ▶                    │
│   From Azure ▶                       │
│   From Online Services ▶             │
│   From Other Sources ▶               │
│                                      │
│ Queries & Connections                │
│ Refresh All                          │
│ Edit Query                           │
└──────────────────────────────────────┘
```

### Opening Power Query Editor

**Method 1: Data Tab**
- **Data → Get Data → Launch Power Query Editor**

**Method 2: From Query**
- **Data → Queries & Connections**
- Right-click query → **Edit**

**Method 3: When importing**
- Get Data → From File → CSV
- Click **Transform Data** (not Load)

---

## Power Query Editor Interface

### Main Components

```
┌───────────────────────────────────────────────────────┐
│ Power Query Editor                                    │
├─────────────┬─────────────────────────────────────────┤
│             │ Home | Transform | Add Column | View    │← Ribbon
├─────────────┼─────────────────────────────────────────┤
│ Queries     │ Data Preview                            │
│             │                                         │
│ ▼ Queries   │     Column1    Column2    Column3       │
│   Query1    │   ┌──────────┬──────────┬──────────┐   │
│   Query2    │   │ Value1   │ Value2   │ Value3   │   │
│             │   ├──────────┼──────────┼──────────┤   │
│             │   │ Value4   │ Value5   │ Value6   │   │
│             │   └──────────┴──────────┴──────────┘   │
├─────────────┼─────────────────────────────────────────┤
│ Properties  │ Applied Steps:                          │
│             │ ► Source                                │
│ Name:       │ ► Promoted Headers                      │
│ [Query1]    │ ► Changed Type                          │
│             │ ► Filtered Rows                         │
│             │ ► Added Custom Column                   │
│             │                  ↑                      │
│             │          Each transformation recorded   │
└─────────────┴─────────────────────────────────────────┘
```

### Key Areas

**1. Queries Pane (Left)**
- Lists all queries in workbook
- Organize into groups
- Shows query dependencies

**2. Data Preview (Center)**
- Shows sample of data (top 1000 rows)
- Can scroll and inspect
- Apply transformations here

**3. Query Settings (Right)**
- Query properties (name, description)
- **Applied Steps** - transformation history
- Each step can be edited or deleted

**4. Ribbon (Top)**
- **Home:** Common transformations
- **Transform:** Data manipulation
- **Add Column:** Create new columns
- **View:** Display options

**5. Formula Bar**
- Shows M code for selected step
- Can edit directly (advanced)

---

## Basic Workflow

### Step-by-Step Process

**1. Get Data**
```
Data Tab → Get Data → From File → From Text/CSV
Select file → Click Open
```

**2. Preview Data**
```
┌─────────────────────────────────────┐
│ Import Data Preview                 │
├─────────────────────────────────────┤
│ File: sales.csv                     │
│                                     │
│ Name      Sales    Region          │
│ John      5000     East            │
│ Sarah     6500     West            │
│                                     │
│ [Load ▼] [Transform Data] [Cancel] │
└─────────────────────────────────────┘
```

**3. Choose Action**
- **Load:** Import as-is into Excel
- **Transform Data:** Opens Power Query Editor for cleaning

**4. Transform (if chosen)**
```
Power Query Editor opens
Apply transformations:
- Remove columns
- Filter rows
- Change data types
- Split columns
- Etc.
```

**5. Load to Excel**
```
Home Tab → Close & Load
Or: Close & Load To... (choose destination)

Data appears in Excel worksheet or Table
Query saved for future refresh
```

---

## Common Transformations

### 1. Remove Columns

**Scenario:** CSV has 20 columns, you need only 5

**Steps:**
1. Select columns to **keep** (Ctrl + click)
2. Right-click → **Remove Other Columns**

Or:

1. Select columns to **remove**
2. Right-click → **Remove Columns**

**Visual:**
```
Before:
Col1 | Col2 | Col3 | Col4 | Col5 | Col6 | Col7

After (remove Col3, Col5, Col6, Col7):
Col1 | Col2 | Col4
```

### 2. Filter Rows

**Scenario:** Show only sales > $1000

**Steps:**
1. Click dropdown in column header
2. **Number Filters → Greater Than**
3. Enter: 1000
4. Click **OK**

**Filter types:**
```
Text Filters:
- Contains, Begins With, Ends With
- Equals, Does Not Equal

Number Filters:
- Greater Than, Less Than
- Between, Top N
- Above Average, Below Average

Date Filters:
- Before, After, Between
- Last 7 Days, This Month, Last Year
```

**Visual:**
```
Before:
Sales
500
1500
800
2000

After (Sales > 1000):
Sales
1500
2000
```

### 3. Change Data Type

**Scenario:** Numbers imported as text

**Steps:**
1. Select column
2. **Transform Tab → Data Type**
3. Choose: Whole Number, Decimal, Date, Text, etc.

**Visual indicator:**
```
Column header icons:
123  = Whole Number
1.2  = Decimal Number
ABC  = Text
📅   = Date
🕐   = Time
T/F  = True/False
```

**Why important:**
- Text "100" won't calculate
- Date as text won't sort correctly
- Decimals as text won't average properly

### 4. Remove Duplicates

**Scenario:** Customer list with duplicate entries

**Steps:**
1. Select columns to check for duplicates
2. **Home Tab → Remove Rows → Remove Duplicates**

**Example:**
```
Before:
Name     Email
John     john@ex.com
Sarah    sarah@ex.com
John     john@ex.com  ← Duplicate

After:
Name     Email
John     john@ex.com
Sarah    sarah@ex.com
```

### 5. Split Column

**Scenario:** Full Name column needs to split into First and Last

**Steps:**
1. Select column
2. **Transform Tab → Split Column → By Delimiter**
3. Choose: Space, Comma, Custom, etc.
4. Split at: Each occurrence of delimiter
5. Click **OK**

**Example:**
```
Before:
Full Name
John Smith
Sarah Jones

After:
Full Name.1  | Full Name.2
John         | Smith
Sarah        | Jones

Rename columns: First Name, Last Name
```

### 6. Merge Columns

**Scenario:** Combine First and Last Name

**Steps:**
1. Select columns (Ctrl + click)
2. **Transform Tab → Merge Columns**
3. Choose separator: Space, Comma, Custom, None
4. New column name: Full Name
5. Click **OK**

**Example:**
```
Before:
First    | Last
John     | Smith
Sarah    | Jones

After:
Full Name
John Smith
Sarah Jones
```

### 7. Replace Values

**Scenario:** Replace "N/A" with blank or 0

**Steps:**
1. Select column
2. **Transform Tab → Replace Values**
3. Value to find: N/A
4. Replace with: (leave blank or enter 0)
5. Click **OK**

**Example:**
```
Before:
Sales
1000
N/A
1500

After:
Sales
1000
0
1500
```

### 8. Fill Down

**Scenario:** Category only in first row, needs to fill down

**Steps:**
1. Select column
2. Right-click → **Fill → Down**

**Example:**
```
Before:
Category  | Item
Fruit     | Apple
          | Banana
          | Cherry
Vegetable | Carrot
          | Lettuce

After:
Category  | Item
Fruit     | Apple
Fruit     | Banana
Fruit     | Cherry
Vegetable | Carrot
Vegetable | Lettuce
```

### 9. Remove Blank Rows

**Scenario:** CSV has empty rows scattered throughout

**Steps:**
1. **Home Tab → Remove Rows → Remove Blank Rows**

Or filter specific column:
1. Select column
2. **Home Tab → Remove Rows → Remove Empty**

### 10. Trim and Clean

**Scenario:** Data has extra spaces

**Steps:**
1. Select text column
2. **Transform Tab → Format → Trim** (removes leading/trailing spaces)
3. **Transform Tab → Format → Clean** (removes non-printable characters)

**Example:**
```
Before:
"  John  "
" Sarah"
"Mike   "

After Trim:
"John"
"Sarah"
"Mike"
```

---

## Applied Steps

### Understanding Steps

**Each transformation creates a step:**

```
Applied Steps:
► Source              ← Get data from file
► Promoted Headers    ← First row becomes headers
► Changed Type        ← Set data types
► Filtered Rows       ← Remove rows < 1000
► Removed Columns     ← Delete unwanted columns
► Added Custom        ← Calculate new column
```

### Managing Steps

**View step:**
- Click step name
- Preview shows data at that point
- See what transformation did

**Edit step:**
- Click gear icon ⚙ next to step
- Modify parameters
- Click OK

**Delete step:**
- Click X next to step
- Confirmation: "Delete Step?"
- Dependent steps may break

**Rename step:**
- Right-click step → **Rename**
- Use descriptive names

**Reorder steps:**
- Click and drag (limited - dependencies matter)

**Insert step:**
- Click step where you want to insert
- Apply new transformation
- New step inserts after selected

### Step Dependencies

```
Applied Steps:
1. Source
2. Changed Type     ← Depends on Source
3. Filtered Rows    ← Depends on Changed Type
4. Removed Columns  ← Depends on Filtered Rows

If you delete step 2, steps 3-4 might break!
```

---

## Creating Custom Columns

### Add Custom Column

**Scenario:** Calculate Total = Quantity × Price

**Steps:**
1. **Add Column Tab → Custom Column**
2. New column name: Total
3. Formula: `[Quantity] * [Price]`
4. Click **OK**

**Dialog:**
```
┌──────────────────────────────────────┐
│ Custom Column                        │
├──────────────────────────────────────┤
│ New column name:                     │
│ [Total____________]                  │
│                                      │
│ Custom column formula:               │
│ = [Quantity] * [Price]               │
│                                      │
│ Available columns:                   │
│ - Quantity                           │
│ - Price                              │
│ - Product                            │
│                                      │
│ [OK] [Cancel]                        │
└──────────────────────────────────────┘
```

### Custom Column Formula Syntax

**Reference columns:**
```
[Column Name]

Examples:
[Sales]
[First Name]
[Order Date]
```

**Arithmetic:**
```
[Price] * [Quantity]
[Revenue] - [Costs]
[Total] / [Count]
```

**Text concatenation:**
```
[First Name] & " " & [Last Name]
"Order #" & Text.From([Order ID])
```

**Conditional logic:**
```
if [Sales] > 1000 then "High" else "Low"

if [Region] = "East" then [Sales] * 1.1
else if [Region] = "West" then [Sales] * 1.05
else [Sales]
```

**Date calculations:**
```
Date.Year([Order Date])
Date.Month([Invoice Date])
Duration.Days([End Date] - [Start Date])
```

### Common Custom Column Examples

**Example 1: Full Name**
```
[First Name] & " " & [Last Name]
```

**Example 2: Profit**
```
[Revenue] - [Cost]
```

**Example 3: Profit Margin**
```
([Revenue] - [Cost]) / [Revenue]
```

**Example 4: Sales Category**
```
if [Sales] >= 10000 then "Platinum"
else if [Sales] >= 5000 then "Gold"
else if [Sales] >= 1000 then "Silver"
else "Bronze"
```

**Example 5: Days Since Order**
```
Duration.Days(DateTime.LocalNow() - [Order Date])
```

---

## Combining Queries

### Append Queries (Union)

**Scenario:** Combine data from multiple similar files

**When to use:** Same columns, different rows

**Example:**
```
Query1 (Jan Sales):
Name  | Sales | Region
John  | 5000  | East

Query2 (Feb Sales):
Name  | Sales | Region
Sarah | 6500  | West

Result (Appended):
Name  | Sales | Region
John  | 5000  | East
Sarah | 6500  | West
```

**Steps:**
1. **Home Tab → Append Queries → Append Queries as New**
2. Select: Query1
3. Add: Query2 (and others)
4. Click **OK**
5. New combined query created

**Visual:**
```
┌──────┐       ┌──────┐
│Query1│       │Query2│
└──┬───┘       └──┬───┘
   │              │
   └──────┬───────┘
          ↓
    ┌───────────┐
    │ Combined  │
    └───────────┘
    All rows together
```

### Merge Queries (Join)

**Scenario:** Add customer details to orders

**When to use:** Match rows based on common column (like VLOOKUP)

**Example:**
```
Orders Query:
OrderID | CustomerID | Amount
1001    | C01        | 500

Customers Query:
CustomerID | Name  | Region
C01        | John  | East

Result (Merged):
OrderID | CustomerID | Amount | Name | Region
1001    | C01        | 500    | John | East
```

**Steps:**
1. Select Orders query
2. **Home Tab → Merge Queries → Merge Queries as New**
3. Select Customers query
4. Select matching columns (CustomerID in both)
5. Join Kind: Left Outer (most common)
6. Click **OK**
7. Expand new column to show Name, Region

**Join Types:**
```
┌──────────────────┬────────────────────────┐
│ Join Kind        │ Result                 │
├──────────────────┼────────────────────────┤
│ Left Outer       │ All from left + matches│
│ Right Outer      │ All from right + matches│
│ Full Outer       │ All from both          │
│ Inner            │ Only matches           │
│ Left Anti        │ Left without matches   │
│ Right Anti       │ Right without matches  │
└──────────────────┴────────────────────────┘
```

**Expanding merged column:**
```
After merge:
OrderID | CustomerID | Amount | Customers
1001    | C01        | 500    | Table    ← Click expand icon

Choose columns:
☑ Name
☑ Region
☐ Use original column name as prefix

Result:
OrderID | CustomerID | Amount | Name | Region
1001    | C01        | 500    | John | East
```

---

## Loading Data to Excel

### Load Options

**Home Tab → Close & Load dropdown:**

**Option 1: Close & Load**
- Loads to new worksheet
- Creates Table
- Query saved

**Option 2: Close & Load To...**

**Dialog:**
```
┌──────────────────────────────────────┐
│ Import Data                          │
├──────────────────────────────────────┤
│ How do you want to view this data?   │
│   ○ Table                            │
│   ○ PivotTable Report                │
│   ○ PivotChart                       │
│   ○ Only Create Connection           │
│                                      │
│ Where do you want to put the data?   │
│   ○ Existing worksheet: [$A$1]       │
│   ○ New worksheet                    │
│                                      │
│ ☐ Add this data to the Data Model    │
│                                      │
│ [OK] [Cancel]                        │
└──────────────────────────────────────┘
```

**Option 3: Only Create Connection**
- Doesn't load data to sheet
- Query available for:
  - Merging with other queries
  - Loading to Data Model
  - Using later

### Load to Data Model

**For large datasets (1M+ rows):**

**Steps:**
1. Close & Load To
2. Check **Add this data to the Data Model**
3. Data stored in Power Pivot (not worksheet)
4. Use PivotTable to analyze

**Benefits:**
- Handle millions of rows
- Relationships between tables
- DAX calculations
- Faster than worksheet

---

## Refreshing Queries

### Manual Refresh

**Refresh single query:**
1. **Data Tab → Queries & Connections**
2. Right-click query
3. **Refresh**

**Refresh all queries:**
- **Data Tab → Refresh All**

**Visual:**
```
┌──────────────────────────────┐
│ Queries & Connections        │
├──────────────────────────────┤
│ Queries (2)                  │
│                              │
│ 📊 Sales Data                │
│    100 rows loaded           │
│    Right-click → Refresh     │
│                              │
│ 📊 Customer List             │
│    50 rows loaded            │
└──────────────────────────────┘
```

### Automatic Refresh

**On workbook open:**
1. Data Tab → Queries & Connections
2. Right-click query → **Properties**
3. Check **Refresh data when opening the file**
4. OK

**Scheduled refresh (Excel Online):**
- Save to OneDrive/SharePoint
- Open in Excel Online
- Configure scheduled refresh (requires proper permissions)

### What Happens on Refresh

```
1. Re-executes all Applied Steps
2. Gets latest data from source
3. Applies same transformations
4. Updates results in Excel

Source data changes → Refresh → Excel updates!
```

---

## Query Properties

### Accessing Properties

**Right-click query → Properties:**

```
┌──────────────────────────────────────┐
│ Query Properties                     │
├──────────────────────────────────────┤
│ General:                             │
│   Name: [Sales_Data]                 │
│   Description: [Monthly sales...]    │
│                                      │
│ ☑ Enable load to worksheet           │
│ ☐ Enable load to Data Model          │
│                                      │
│ Refresh Control:                     │
│ ☑ Refresh data when opening file     │
│ ☐ Include in Refresh All             │
│                                      │
│ [OK] [Cancel]                        │
└──────────────────────────────────────┘
```

### Important Settings

**Enable load to worksheet:**
- Checked: Data appears in Excel
- Unchecked: Query available but not loaded (connection only)

**Include in Refresh All:**
- Checked: Refreshes when you click Refresh All
- Unchecked: Must refresh manually

**Refresh on open:**
- Automatic update when workbook opens
- Good for dashboards

---

## Power Query M Language

### What is M?

**M** is the formula language Power Query uses behind the scenes.

**Example - Filter step:**
```m
= Table.SelectRows(#"Changed Type", each [Sales] > 1000)

Function: Table.SelectRows
Source: #"Changed Type" (previous step)
Condition: each [Sales] > 1000
```

### Viewing M Code

**Method 1: Formula bar**
- Click any Applied Step
- Formula bar shows M code for that step

**Method 2: Advanced Editor**
- **View Tab → Advanced Editor**
- Shows complete M script

**Example full script:**
```m
let
    Source = Csv.Document(File.Contents("C:\data.csv")),
    Promoted = Table.PromoteHeaders(Source),
    Changed = Table.TransformColumnTypes(Promoted, {{"Sales", Int64.Type}}),
    Filtered = Table.SelectRows(Changed, each [Sales] > 1000)
in
    Filtered
```

### Common M Functions

**Table functions:**
```m
Table.SelectRows        ← Filter
Table.RemoveColumns     ← Remove columns
Table.RenameColumns     ← Rename
Table.AddColumn         ← Add column
Table.Sort              ← Sort
Table.Group             ← Group by
```

**Text functions:**
```m
Text.Upper              ← Uppercase
Text.Trim               ← Remove spaces
Text.Split              ← Split by delimiter
Text.Replace            ← Replace values
```

**Date functions:**
```m
Date.Year               ← Extract year
Date.Month              ← Extract month
Date.AddDays            ← Add days
Duration.Days           ← Days between dates
```

### When to Use M

✅ **Use GUI (recommended):**
- Most transformations
- Learning Power Query
- Standard operations

✅ **Use M code:**
- Complex custom logic
- Reusable functions
- Advanced transformations
- Parameters

❌ **Avoid if possible:**
- You're just starting
- GUI can do it
- Maintenance by others

---

## Best Practices

### Naming Conventions

```
✅ Good query names:
Sales_2024_Raw
Customer_Master_Cleaned
Orders_Filtered
Product_Catalog

❌ Poor names:
Query1
data
temp
Sheet1
```

### Query Organization

**Use groups:**
```
Queries:
├─ 📁 Source Data
│   ├─ Sales_Raw
│   └─ Customers_Raw
├─ 📁 Transformed
│   ├─ Sales_Cleaned
│   └─ Customers_Cleaned
└─ 📁 Final
    └─ Sales_Dashboard
```

### Step Naming

**Rename steps descriptively:**
```
❌ Default:
- Filtered Rows
- Changed Type1
- Changed Type2

✅ Descriptive:
- Filter Sales Over 1000
- Set Date Types
- Set Number Types
```

### Connection Management

```
✅ Use "Only Create Connection" for intermediate queries
✅ Load only final results to worksheet
✅ Disable refresh for static reference data
✅ Document complex queries (Description field)
```

### Performance Tips

```
✅ Filter early (reduce rows ASAP)
✅ Remove unnecessary columns early
✅ Use specific data types
✅ Avoid volatile functions in custom columns
✅ Fold queries when possible (pushes to source)
```

---

## Troubleshooting

### Problem: Query Fails to Refresh

**Error:** "DataFormat.Error: Invalid URI"

**Causes:**
- Source file moved or deleted
- Network path unavailable
- Permissions changed

**Solutions:**
```
1. Right-click query → Edit
2. Click Source step
3. Click gear icon ⚙
4. Update file path
5. OK → Close & Load
```

### Problem: Column Not Found Error

**Error:** "Expression.Error: The column 'Sales' was not found"

**Cause:** Source data structure changed (column renamed/removed)

**Solutions:**
```
1. Check source data
2. Update column references in Applied Steps
3. Or delete broken steps and recreate
```

### Problem: Type Conversion Errors

**Error:** "DataFormat.Error: We couldn't convert to Number"

**Cause:** Non-numeric data in number column ("N/A", spaces, etc.)

**Solutions:**
```
1. Go to step before type change
2. Clean data first (replace N/A, trim spaces)
3. Then convert type
```

### Problem: Slow Performance

**Symptoms:** Query takes minutes to refresh

**Solutions:**
```
✅ Filter rows early
✅ Remove columns early
✅ Check if "query folding" works (see step context)
✅ Use native data source queries (SQL) when possible
✅ Reduce custom column complexity
```

### Problem: #REF! Error After Refresh

**Cause:** Query deleted or renamed, formulas reference old query

**Solution:**
```
1. Update formulas to reference correct query/table
2. Or recreate query with original name
```

---

## Common Workflows

### Workflow 1: Clean CSV Import

```
1. Get Data → From CSV
2. Transform Data
3. Promote first row to headers
4. Change data types
5. Remove blank rows
6. Remove unnecessary columns
7. Filter to relevant data
8. Trim text columns
9. Replace null values
10. Close & Load
```

### Workflow 2: Combine Multiple Files

```
1. Get Data → From Folder
2. Select folder with multiple CSVs
3. Transform Data
4. Filter to keep only CSV files
5. Click "Combine Files" button
6. Power Query combines all files automatically
7. Apply additional transformations
8. Close & Load
```

### Workflow 3: Lookup/Join

```
1. Load main data query (Orders)
2. Load reference data query (Customers)
3. Select Orders query
4. Merge Queries → Select Customers
5. Match on CustomerID
6. Expand customer details
7. Close & Load
```

---

## Quick Reference: Common Tasks

| Task | Steps |
|------|-------|
| **Remove column** | Select → Right-click → Remove |
| **Filter rows** | Click dropdown → Filter options |
| **Change type** | Select column → Transform → Data Type |
| **Split column** | Transform → Split Column → By Delimiter |
| **Merge columns** | Select both → Transform → Merge Columns |
| **Replace values** | Transform → Replace Values |
| **Remove duplicates** | Home → Remove Rows → Remove Duplicates |
| **Add custom column** | Add Column → Custom Column |
| **Append queries** | Home → Append Queries |
| **Merge queries** | Home → Merge Queries |

---

## Keyboard Shortcuts (Power Query Editor)

| Shortcut | Action |
|----------|--------|
| `Ctrl + Enter` | Apply step |
| `Delete` | Delete step |
| `F2` | Rename step |
| `Ctrl + Q` | Close editor |
| `Alt + F5` | Refresh preview |
| `Ctrl + A` | Select all rows/columns |

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Power Query = Get & Transform Data
- Each transformation creates a step
- Steps are repeatable on refresh
- Data Tab → Get Data to import
- Transform Data to open Power Query Editor
- Close & Load to bring data to Excel
- Refresh reruns all steps automatically
- Can't edit data in Power Query (read-only preview)
- M language powers Power Query behind scenes

### Practice Deeply
- Importing CSV files via Power Query
- Opening Power Query Editor
- Navigating the interface (queries pane, preview, steps)
- Removing unnecessary columns
- Filtering rows with various criteria
- Changing data types (text to number, date, etc.)
- Removing duplicate rows
- Splitting columns by delimiter
- Merging columns together
- Replacing values in columns
- Trimming and cleaning text
- Adding custom columns with simple formulas
- Understanding Applied Steps
- Editing and deleting steps
- Closing and loading data to Excel
- Refreshing queries manually
- Appending queries (combining similar data)
- Merging queries (joining like VLOOKUP)
- Setting query properties (refresh on open)
- Troubleshooting basic errors
- Organizing queries with groups and names

---

## Next Step

After this file, we move to:

**`21-macros-and-vba-introduction.md`**
- What are macros
- Recording macros
- Running macros
- Assigning macros to buttons
- Macro security settings
- Introduction to VBA Editor
- Basic VBA concepts
- When to use macros vs formulas/Power Query
