# Data Import and Export

This file covers importing data from external sources (CSV, TXT, databases, web) and exporting Excel data to various formats - essential skills for working with data from different systems.

---

## Why Import/Export Data?

### Common Scenarios

**Importing:**
- Download sales data from e-commerce platform (CSV)
- Get financial data from accounting software (Excel/CSV)
- Pull customer list from CRM database
- Extract data from company website
- Combine data from multiple sources

**Exporting:**
- Send report to colleague without Excel (PDF)
- Upload data to web application (CSV)
- Share with older Excel versions (compatibility)
- Import into database or other software
- Archive data in universal format

### Benefits of Proper Import/Export

✅ Avoid manual data entry (errors + time)
✅ Maintain data integrity
✅ Enable automation
✅ Connect to live data sources
✅ Share across platforms

---

## Understanding File Formats

### Common Data Formats

| Format | Extension | Description | Use Case |
|--------|-----------|-------------|----------|
| **Excel Workbook** | .xlsx | Modern Excel format | Standard Excel files |
| **Excel Binary** | .xlsb | Compressed binary | Large files, faster |
| **Excel Macro** | .xlsm | With VBA macros | Automated workbooks |
| **CSV** | .csv | Comma-separated | Universal, simple |
| **Text** | .txt | Plain text | Generic data |
| **Tab-delimited** | .txt | Tab-separated | Alternative to CSV |
| **PDF** | .pdf | Portable document | Read-only sharing |
| **HTML** | .html | Web page | Web publishing |
| **XML** | .xml | Structured data | Data exchange |
| **JSON** | .json | JavaScript object | Web APIs |

### Format Characteristics

**CSV (Comma-Separated Values):**
```
Visual representation:
Name,Age,City
John,25,New York
Sarah,30,Boston
Mike,28,Chicago

Characteristics:
✅ Universal (any software can read)
✅ Small file size
✅ No formatting preserved
❌ One sheet only
❌ No formulas
❌ No colors, fonts, etc.
```

**Excel (.xlsx):**
```
Characteristics:
✅ Multiple sheets
✅ Formulas preserved
✅ Formatting included
✅ Charts, images, objects
✅ Data validation, protection
❌ Larger file size
❌ Requires Excel or compatible software
```

**PDF:**
```
Characteristics:
✅ Looks same everywhere
✅ Can't accidentally edit
✅ Professional appearance
❌ Not editable (by design)
❌ Hard to extract data from
```

---

## Importing CSV Files

CSV is the most common format for data exchange.

### Method 1: Open Directly

**Steps:**
1. **File → Open**
2. Browse to CSV file
3. Select file
4. Click **Open**

**What Happens:**
- Excel auto-detects delimiters (commas)
- Creates new workbook
- One sheet with data

**Example:**
```
CSV file content:
Product,Price,Stock
Widget,25.99,150
Gadget,45.50,200

Opens as:
     A       B      C
  ┌────────┬──────┬──────┐
1 │Product │Price │Stock │
  ├────────┼──────┼──────┤
2 │Widget  │25.99 │150   │
  ├────────┼──────┼──────┤
3 │Gadget  │45.50 │200   │
  └────────┴──────┴──────┘
```

### Method 2: Import into Existing Workbook

**Steps:**
1. **Data Tab → Get Data → From File → From Text/CSV**
2. Select CSV file
3. Excel shows preview
4. Click **Load** or **Transform Data**

**Load vs Transform:**
- **Load:** Import as-is
- **Transform Data:** Opens Power Query editor (clean/modify first)

### Method 3: Drag and Drop

**Steps:**
1. Open Excel workbook
2. Find CSV file in File Explorer
3. Drag file onto Excel worksheet
4. Data appears

⚠️ **Warning:** Drag and drop may not always parse correctly

### Common CSV Import Issues

**Problem 1: Wrong Delimiter**
```
CSV uses semicolons (;) instead of commas:
Product;Price;Stock
Widget;25.99;150

Solution:
- Use Text Import Wizard
- Specify semicolon as delimiter
```

**Problem 2: Numbers Imported as Text**
```
Symptom: "25.99" stored as text, can't calculate

Solution:
- Select column
- Data Tab → Text to Columns
- Choose "General" format
Or use Power Query to set data types
```

**Problem 3: Leading Zeros Removed**
```
CSV has: 00123
Excel shows: 123

Solution:
- Format column as Text BEFORE importing
Or use Power Query with Text type
```

**Problem 4: Date Format Issues**
```
CSV has: 01/02/2024
Excel interprets as: Feb 1 (MM/DD) or Jan 2 (DD/MM)?

Solution:
- Use ISO format in CSV: 2024-01-02
- Or specify date format in Power Query
```

---

## Text Import Wizard (Legacy)

For precise control over text file imports.

### When to Use

✅ Non-standard delimiters (semicolon, pipe, tab)
✅ Fixed-width text files
✅ Need to skip rows
✅ Need to specify data types per column

### Steps

**1. Start Wizard:**
- **Data Tab → From Text/CSV** (in older Excel)
- Or **File → Open** → Select text file

**2. Step 1 of 3: File Type**
```
┌─────────────────────────────────────┐
│ Text Import Wizard - Step 1 of 3    │
├─────────────────────────────────────┤
│ Original data type:                 │
│   ○ Delimited (comma, tab, etc.)    │
│   ○ Fixed width (spaces)            │
│                                     │
│ Start import at row: [1]            │
│                                     │
│ Preview:                            │
│ Name,Age,City                       │
│ John,25,New York                    │
└─────────────────────────────────────┘

Choose Delimited or Fixed Width
```

**3. Step 2 of 3: Delimiters**
```
┌─────────────────────────────────────┐
│ Text Import Wizard - Step 2 of 3    │
├─────────────────────────────────────┤
│ Delimiters:                         │
│   ☑ Comma                           │
│   ☐ Tab                             │
│   ☐ Semicolon                       │
│   ☐ Space                           │
│   ☐ Other: [    ]                   │
│                                     │
│ Text qualifier: ["]                 │
│                                     │
│ Preview:                            │
│ Name    │Age    │City               │
│ John    │25     │New York           │
└─────────────────────────────────────┘

Select delimiter(s) that separate columns
```

**4. Step 3 of 3: Column Data Format**
```
┌─────────────────────────────────────┐
│ Text Import Wizard - Step 3 of 3    │
├─────────────────────────────────────┤
│ Column data format:                 │
│   ○ General                         │
│   ○ Text                            │
│   ○ Date: MDY ▼                     │
│   ○ Do not import (skip)            │
│                                     │
│ Click column to select format:      │
│ ↓                                   │
│ Name    │Age    │City               │
│ General │General│General            │
└─────────────────────────────────────┘

Set data type for each column
```

**5. Finish:**
- Click **Finish**
- Choose where to put data
- Data imports

### Fixed-Width Files

**Example file:**
```
Name     Age City
John     25  New York
Sarah    30  Boston
Mike     28  Chicago

No delimiters - spaces separate columns
```

**Steps:**
1. Choose **Fixed Width** in Step 1
2. Step 2: Set column breaks by clicking ruler
3. Step 3: Format each column
4. Finish

**Visual of setting breaks:**
```
┌─────────────────────────────────────┐
│ Click ruler to create column breaks │
│                                     │
│ 0    5    10   15   20   25        │
│ ├────┼────┼────┼────┼────┤         │
│ Name     Age City                   │
│      ↑   ↑   ↑                      │
│    Set breaks here                  │
└─────────────────────────────────────┘
```

---

## Importing from Excel Files

### Opening Another Workbook

**Simplest method:**
1. **File → Open**
2. Select .xlsx/.xls file
3. Opens in new window

### Copying Between Workbooks

**Steps:**
1. Open both workbooks
2. Select data in source
3. **Copy** (Ctrl + C)
4. Switch to destination
5. **Paste** (Ctrl + V)

**Paste Options:**
- Values only (no formulas)
- Formulas
- Formatting
- All

### Linking to External Workbook

**Create live connection:**

**Formula in destination workbook:**
```excel
='[SourceFile.xlsx]Sheet1'!A1

Syntax:
='[Filename.xlsx]SheetName'!CellReference
```

**Example:**
```excel
Get value from A1 in Budget.xlsx:
='[Budget.xlsx]Summary'!A1

When Budget.xlsx changes, this updates
```

⚠️ **Warning:** Links break if source file moves/renames

**Managing Links:**
1. **Data Tab → Edit Links**
2. View all external links
3. Update, Change Source, or Break Links

---

## Importing from Databases

Connect Excel to SQL databases, Access, or other data sources.

### Types of Connections

| Type | Description | Use Case |
|------|-------------|----------|
| **Microsoft Access** | .accdb/.mdb files | Small business databases |
| **SQL Server** | Enterprise database | Large corporate data |
| **MySQL/PostgreSQL** | Open-source databases | Web applications |
| **ODBC** | Generic connection | Various database systems |
| **OLEDB** | Windows connection | Legacy systems |

### Connecting to Access Database

**Steps:**
1. **Data Tab → Get Data → From Database → From Microsoft Access Database**
2. Browse to .accdb file
3. Select file
4. Choose table or query
5. Click **Load** or **Transform Data**

**Navigator shows:**
```
┌─────────────────────────────────────┐
│ Navigator                           │
├─────────────────────────────────────┤
│ Database: Sales.accdb               │
│                                     │
│ Tables:                             │
│   ☐ Customers                       │
│   ☐ Orders                          │
│   ☐ Products                        │
│                                     │
│ Queries:                            │
│   ☐ SalesByRegion                   │
│   ☐ TopCustomers                    │
│                                     │
│ [Load] [Transform Data] [Cancel]    │
└─────────────────────────────────────┘

Select what to import
```

### Connecting to SQL Server

**Steps:**
1. **Data Tab → Get Data → From Database → From SQL Server Database**
2. Enter server name
3. Enter database name (optional)
4. Choose authentication:
   - Windows (use your login)
   - Database (username/password)
5. Click **OK**
6. Select tables
7. Load or Transform

**Connection Dialog:**
```
┌─────────────────────────────────────┐
│ SQL Server Database                 │
├─────────────────────────────────────┤
│ Server: [server.company.com]        │
│                                     │
│ Database (optional): [SalesDB]      │
│                                     │
│ Authentication:                     │
│   ○ Windows                         │
│   ○ Database                        │
│     Username: [____]                │
│     Password: [____]                │
│                                     │
│ [OK] [Cancel]                       │
└─────────────────────────────────────┘
```

### Writing SQL Query

**Advanced users can write custom queries:**

**Steps:**
1. Connect to database
2. In Navigator: **Advanced Options**
3. Write SQL query:

```sql
SELECT 
    CustomerName,
    SUM(OrderTotal) AS TotalSpent
FROM Orders
WHERE OrderDate >= '2024-01-01'
GROUP BY CustomerName
ORDER BY TotalSpent DESC
```

4. Click **OK**
5. Data loads with query results

### Refreshing Database Connections

**Data doesn't update automatically:**

**To refresh:**
1. **Data Tab → Refresh All**
2. Or right-click table → **Refresh**

**Schedule auto-refresh:**
1. Right-click table → **Table Design**
2. **Properties**
3. Check **Refresh data when opening the file**
4. Set refresh interval if desired

---

## Importing from Web

Get data directly from web pages.

### Method 1: From Web (Simple)

**Steps:**
1. **Data Tab → Get Data → From Other Sources → From Web**
2. Enter URL
3. Click **OK**
4. Excel detects tables on page
5. Select table(s)
6. Click **Load**

**Example URL:**
```
https://en.wikipedia.org/wiki/List_of_countries_by_GDP

Excel finds and lists all tables on page
```

**Navigator shows:**
```
┌─────────────────────────────────────┐
│ Navigator                           │
├─────────────────────────────────────┤
│ Tables detected:                    │
│                                     │
│   ☐ Table 0 (GDP by Country)        │
│   ☐ Table 1 (Historical Data)       │
│   ☐ Table 2 (Regional)              │
│                                     │
│ Preview: (shows selected table)     │
│                                     │
│ [Load] [Transform Data] [Cancel]    │
└─────────────────────────────────────┘
```

### Method 2: Web Query (Advanced)

**For dynamic pages or authentication:**

**Steps:**
1. **Data Tab → Get Data → From Other Sources → From Web**
2. Click **Advanced**
3. URL parts: Enter URL
4. HTTP request header parameters (if needed)
5. Authentication (if required)

**Authentication Options:**
- Anonymous (no login)
- Windows (current user)
- Basic (username/password)
- Web API Key
- Organizational account

### Refreshing Web Data

**Manual:**
- Right-click table → **Refresh**

**Automatic:**
- Table Design → Properties
- Check "Refresh every X minutes"

⚠️ **Warning:** Some websites block automated scraping

---

## Power Query Basics

**Power Query** = Excel's modern data import/transformation tool

### What is Power Query?

**Purpose:**
- Import data from any source
- Clean and transform data
- Combine multiple sources
- Create repeatable processes
- Refresh automatically

**Where to find:**
- **Data Tab → Get Data**
- **Data Tab → Get & Transform Data** section

### Power Query Editor

**Opens when you choose "Transform Data":**

```
┌───────────────────────────────────────────────────────┐
│ Power Query Editor                                    │
├─────────────┬─────────────────────────────────────────┤
│ Queries     │ Preview                                 │
│             │                                         │
│ ▼ Queries   │     A       B       C                   │
│   Query1    │  ┌──────┬──────┬──────┐               │
│             │  │Name  │Age   │City  │               │
│             │  ├──────┼──────┼──────┤               │
│             │  │John  │25    │NY    │               │
│             │  │Sarah │30    │LA    │               │
│             │  └──────┴──────┴──────┘               │
├─────────────┼─────────────────────────────────────────┤
│             │ Applied Steps:                          │
│             │ ► Source                                │
│             │ ► Changed Type                          │
│             │ ► Filtered Rows                         │
└─────────────┴─────────────────────────────────────────┘
```

**Components:**
- **Queries pane:** List of all queries
- **Preview:** Shows data
- **Applied Steps:** Each transformation recorded
- **Ribbon:** Transformation tools

### Common Power Query Transformations

**Remove Columns:**
1. Select column(s)
2. Right-click → **Remove Columns**

**Filter Rows:**
1. Click dropdown in column header
2. Uncheck values to exclude
3. Or use filters (Text, Number, Date filters)

**Change Data Type:**
1. Select column
2. **Transform Tab → Data Type**
3. Choose: Text, Number, Date, etc.

**Replace Values:**
1. Select column
2. **Transform Tab → Replace Values**
3. Value to find: `Old`
4. Replace with: `New`

**Split Column:**
1. Select column
2. **Transform Tab → Split Column → By Delimiter**
3. Choose delimiter (comma, space, etc.)

**Example:**
```
Before:
FullName
John Smith
Sarah Jones

After splitting by space:
FirstName | LastName
John      | Smith
Sarah     | Jones
```

**Merge Columns:**
1. Select columns (Ctrl + click)
2. **Transform Tab → Merge Columns**
3. Choose separator
4. Name new column

**Add Custom Column:**
1. **Add Column Tab → Custom Column**
2. Enter formula:
   ```
   = [Price] * [Quantity]
   ```
3. Name: "Total"

**Group By (Aggregate):**
1. **Transform Tab → Group By**
2. Group by: Region
3. New column: Total Sales
4. Operation: Sum
5. Column: Sales

### Applying Changes

**When done transforming:**
1. **Home Tab → Close & Load**
2. Data loads into Excel worksheet
3. Query saved for future refreshes

**Visual:**
```
Before Power Query:
Messy data with errors, blanks, wrong types

Power Query transformations:
► Remove blank rows
► Fix data types
► Split name column
► Filter to 2024 only
► Calculate totals

After Power Query:
Clean, structured, ready-to-analyze data!
```

---

## Exporting Data

### Save As Excel Formats

**File → Save As:**

**Excel Workbook (.xlsx):**
- Default modern format
- Multiple sheets
- Formulas, formatting, charts
- Compatible with Excel 2007+

**Excel Binary (.xlsb):**
- Compressed format
- Faster to open/save
- Smaller file size
- Same features as .xlsx
- Use for very large files (100MB+)

**Excel Macro-Enabled (.xlsm):**
- Contains VBA macros
- If workbook has macros, must save as .xlsm
- Security warning when opening

**Excel 97-2003 (.xls):**
- Old format for compatibility
- Limited to 65,536 rows
- Missing modern features
- Only if recipient has old Excel

### Export to CSV

**Method 1: Save As**

**Steps:**
1. **File → Save As**
2. File type: **CSV (Comma delimited) (*.csv)**
3. Save

⚠️ **Warnings shown:**
```
┌──────────────────────────────────────┐
│ Warning!                             │
├──────────────────────────────────────┤
│ File may contain features not        │
│ compatible with CSV:                 │
│ - Multiple sheets (only active saved)│
│ - Formulas (converted to values)     │
│ - Formatting (lost)                  │
│ - Charts (lost)                      │
│                                      │
│ Continue?     [Yes] [No]             │
└──────────────────────────────────────┘
```

**What Gets Saved:**
```
✅ Values from active sheet only
✅ Text and numbers
❌ Formulas (replaced with calculated values)
❌ Formatting (colors, fonts, etc.)
❌ Multiple sheets
❌ Charts, images
❌ Data validation
```

**CSV File Result:**
```
Name,Age,City
John,25,New York
Sarah,30,Boston
Mike,28,Chicago

Simple text file, comma-separated
```

**Method 2: Export to CSV UTF-8**

**For international characters:**

**Steps:**
1. **File → Save As**
2. File type: **CSV UTF-8 (Comma delimited) (*.csv)**
3. Save

**Difference:**
- UTF-8 preserves international characters (é, ñ, 中, 日)
- Standard CSV may show garbled text
- Use UTF-8 for multilingual data

### Export to Text (Tab-Delimited)

**Alternative to CSV:**

**Steps:**
1. **File → Save As**
2. File type: **Text (Tab delimited) (*.txt)**
3. Save

**Result:**
```
Name    Age    City
John    25     New York
Sarah   30     Boston

Tabs separate columns instead of commas
```

**When to use:**
- Data contains commas in values
- Some systems prefer tab-delimited
- Alternative if CSV doesn't work

### Export to PDF

**Create non-editable document:**

**Method 1: Save As PDF**

**Steps:**
1. **File → Save As**
2. File type: **PDF (*.pdf)**
3. Options:
   - **Standard:** Full quality
   - **Minimum size:** Compressed
4. Choose what to export:
   - Entire workbook
   - Active sheet(s)
   - Selection
5. Save

**Method 2: Export to PDF**

**Steps:**
1. **File → Export → Create PDF/XPS**
2. Choose location and name
3. Options (button):
   - Page range
   - What to publish
4. Publish

**PDF Options:**
```
┌──────────────────────────────────────┐
│ Options                              │
├──────────────────────────────────────┤
│ Publish what:                        │
│   ○ Active sheet(s)                  │
│   ○ Entire workbook                  │
│   ○ Selection                        │
│   ○ Table                            │
│                                      │
│ ☑ Include document properties        │
│ ☐ ISO 19005-1 compliant (PDF/A)      │
│                                      │
│ [OK] [Cancel]                        │
└──────────────────────────────────────┘
```

**Before exporting PDF:**
```
✅ Set print area (Page Layout → Print Area)
✅ Check page breaks (View → Page Break Preview)
✅ Hide unnecessary rows/columns
✅ Adjust scaling (fit to page)
✅ Set headers/footers
```

### Export to HTML

**Publish to web:**

**Steps:**
1. **File → Save As**
2. File type: **Web Page (*.htm, *.html)**
3. Options:
   - **Web Page:** Interactive, keeps some Excel features
   - **Web Page, Filtered:** Removes Excel-specific code, smaller
4. Choose:
   - Entire Workbook
   - Selection
5. Save

**What happens:**
- Creates HTML file
- Creates folder with images/supporting files
- Can open in web browser
- Basic interactivity preserved (sorting)

### Export Specific Range

**Export only part of worksheet:**

**Method 1: Copy to new sheet, save as**
1. Select range
2. Copy
3. New sheet → Paste Values
4. Save that sheet as CSV/PDF/etc.

**Method 2: Define Print Area, save as PDF**
1. Select range
2. **Page Layout Tab → Print Area → Set Print Area**
3. Save as PDF

**Method 3: Use Power Query**
1. Load data to Power Query
2. Filter to desired rows
3. Close & Load to new sheet
4. Export that sheet

---

## Connecting to Online Sources

### SharePoint/OneDrive

**Collaborate with cloud storage:**

**Import from SharePoint:**
1. **Data Tab → Get Data → From File → From SharePoint Folder**
2. Enter SharePoint site URL
3. Authenticate
4. Select file(s)
5. Load or Transform

**Benefits:**
- Data stays in cloud
- Multiple users can access
- Auto-updates when source changes
- Version control

### Microsoft Dataverse

**Connect to Power Platform:**

1. **Data Tab → Get Data → From Power Platform → From Dataverse**
2. Enter environment URL
3. Select tables
4. Load

**Use case:**
- Business applications built on Power Platform
- Integrate with Dynamics 365
- Custom apps

### OData Feed

**Connect to web services:**

1. **Data Tab → Get Data → From Other Sources → From OData Feed**
2. Enter URL of OData endpoint
3. Authenticate if required
4. Select entities
5. Load

**Example sources:**
- Public data APIs
- Internal company services
- Third-party data providers

### JSON Files

**Import from APIs or JSON files:**

**Method 1: From File**
1. **Data Tab → Get Data → From File → From JSON**
2. Select JSON file
3. Power Query parses structure
4. Transform as needed

**Method 2: From Web**
1. **Data Tab → Get Data → From Other Sources → From Web**
2. Enter API URL that returns JSON
3. Add authentication if needed
4. Power Query converts to table

**Example JSON:**
```json
{
  "employees": [
    {"name": "John", "age": 25, "dept": "Sales"},
    {"name": "Sarah", "age": 30, "dept": "IT"}
  ]
}

Power Query converts to table:
Name  | Age | Dept
John  | 25  | Sales
Sarah | 30  | IT
```

### XML Files

**Import structured XML:**

1. **Data Tab → Get Data → From File → From XML**
2. Select XML file
3. Navigate XML structure in Navigator
4. Select table(s)
5. Load or Transform

---

## Data Refresh Strategies

### Manual Refresh

**Simplest approach:**

**Refresh single query/table:**
- Right-click table
- **Refresh**

**Refresh all:**
- **Data Tab → Refresh All**
- Or **Alt + F5**

### Automatic Refresh on Open

**Refresh when workbook opens:**

**Steps:**
1. Click in table/query
2. **Data Tab → Queries & Connections**
3. Right-click query
4. **Properties**
5. Check **Refresh data when opening the file**
6. OK

### Scheduled Refresh (Excel Online)

**For files in OneDrive/SharePoint:**

1. Save workbook to OneDrive/SharePoint
2. Open in Excel Online
3. **Data Tab → Refresh All** (button menu)
4. **Connection Settings**
5. Enable scheduled refresh
6. Set frequency (daily, weekly, etc.)

⚠️ **Note:** Requires Power BI or Microsoft 365 E5

### Background Refresh

**Don't wait for refresh to complete:**

**Steps:**
1. Query Properties
2. Check **Enable background refresh**
3. Continue working while data refreshes

**Visual indicator:**
```
Status bar shows:
"Refreshing query 'Sales Data'..." (🔄)

When done:
"Refresh completed" (✓)
```

---

## Troubleshooting Import/Export

### Problem: CSV Imports Wrong

**Symptom:** All data in one column

**Cause:** Wrong delimiter

**Solution:**
```
Use Text Import Wizard:
1. Data Tab → From Text/CSV
2. Delimiter: Change from comma to correct one
3. Preview updates
4. Load
```

**Symptom:** Numbers treated as text

**Cause:** Quotes around numbers in CSV

**Solution:**
```
Power Query:
1. Transform Data
2. Select column
3. Transform Tab → Data Type → Whole Number/Decimal
4. Close & Load
```

### Problem: Import Fails with Error

**Common errors:**

**"File not found"**
- File moved or deleted
- Check file path
- Update connection

**"Permission denied"**
- File open in another program
- Close file and retry
- Check file permissions

**"Data source error"**
- Database offline
- Network issue
- Credentials expired
- Re-enter login information

### Problem: Export Loses Formatting

**Cause:** CSV doesn't support formatting

**Solutions:**
```
Option 1: Keep in Excel format (.xlsx)
Option 2: Export to PDF (formatting preserved)
Option 3: Save CSV + separate Excel template
```

### Problem: PDF Export Cuts Off Data

**Cause:** Page scaling issues

**Solutions:**
```
1. Page Layout View
2. Adjust scaling: Fit to 1 page wide
3. Or adjust column widths
4. Check Print Preview before exporting
5. Consider landscape orientation
```

### Problem: Data Refresh Breaks

**Cause:** Source file moved/renamed

**Solution:**
```
1. Data Tab → Queries & Connections
2. Right-click query → Properties
3. Definition Tab
4. Edit connection string
5. Update file path/name
6. OK
```

### Problem: Imported Dates Wrong

**Cause:** Regional settings mismatch

**Solution:**
```
Power Query:
1. Transform Data
2. Select date column
3. Transform Tab → Data Type → Date
4. Using Locale → Choose correct region
5. Close & Load
```

---

## Best Practices

### Importing Data

```
✅ Use Power Query for repeatable imports
✅ Store source files in consistent location
✅ Document connection information
✅ Test with sample data first
✅ Verify data types after import
✅ Check for blank rows/columns
✅ Keep source data separate from analysis
```

### Exporting Data

```
✅ Choose format based on recipient's needs
✅ CSV for universal compatibility
✅ PDF for read-only distribution
✅ Excel for further analysis
✅ Test exported file before sending
✅ Include metadata (date, source, notes)
✅ Remove sensitive data before sharing
```

### Data Connections

```
✅ Use Tables for source data (auto-expand)
✅ Name queries descriptively
✅ Document transformation steps
✅ Test refresh process
✅ Set appropriate refresh schedule
✅ Monitor for errors
✅ Have backup plan if source unavailable
```

### File Management

```
✅ Organized folder structure
✅ Consistent naming convention (Date_Description.xlsx)
✅ Version control for important files
✅ Regular backups
✅ Archive old versions
✅ Document data sources and refresh procedures
```

---

## Real-World Scenarios

### Scenario 1: Daily Sales Report

**Source:** E-commerce platform exports CSV daily

**Process:**
1. Save CSV to consistent folder: `C:\Reports\Sales\`
2. Excel workbook with Power Query connection
3. Query points to folder, gets latest file
4. Transformations: Filter, clean, calculate
5. Pivot Table summarizes data
6. Chart visualizes trends
7. Save/refresh each morning
8. Export summary as PDF, email to team

### Scenario 2: Database Dashboard

**Source:** SQL Server database with sales data

**Process:**
1. Create Excel workbook
2. Data Tab → From SQL Server
3. Write query to get last 90 days
4. Load to Excel
5. Create Pivot Tables and Charts
6. Save to SharePoint
7. Schedule refresh: Every 4 hours
8. Team accesses live dashboard online

### Scenario 3: Web Data Analysis

**Source:** Company website with product prices

**Process:**
1. Data Tab → From Web
2. Enter competitor website URL
3. Select pricing table
4. Transform: Clean product names, fix prices
5. Load to Excel
6. Compare with our prices (using formulas)
7. Highlight price differences (conditional formatting)
8. Refresh weekly to monitor changes
9. Export report as PDF for management

### Scenario 4: Multi-File Consolidation

**Source:** 50 CSV files from different stores

**Process:**
1. Place all CSV files in one folder
2. Data Tab → Get Data → From Folder
3. Power Query: Combine all files
4. Add column for store name (from filename)
5. Transform: Standardize formats
6. Load to Excel
7. Create Pivot Table showing all stores
8. Monthly: Add new files, refresh query
9. Automatic consolidation!

---

## Quick Reference: Format Selection

| Need | Best Format | Why |
|------|-------------|-----|
| **Share with anyone** | CSV | Universal compatibility |
| **Preserve formulas** | .xlsx | Full Excel features |
| **Read-only distribution** | PDF | Can't be edited |
| **Import to database** | CSV | Standard import format |
| **Web publishing** | HTML | Browser-viewable |
| **Large file (100MB+)** | .xlsb | Compressed, faster |
| **Contains macros** | .xlsm | VBA preserved |
| **Old Excel (2003)** | .xls | Backward compatibility |
| **Data exchange APIs** | JSON/XML | Structured format |
| **International data** | CSV UTF-8 | Preserves special characters |

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Alt + A + FT` | Get Data → From Text/CSV |
| `Alt + A + FW` | Get Data → From Web |
| `Alt + A + FA` | Refresh All |
| `Ctrl + S` | Save |
| `F12` | Save As |
| `Ctrl + P` | Print/Export to PDF |
| `Alt + A + P + A` | Queries & Connections |

---

## Common File Extension Reference

### Excel Formats
```
.xlsx  → Excel Workbook (default)
.xlsm  → Excel Macro-Enabled Workbook
.xlsb  → Excel Binary Workbook
.xls   → Excel 97-2003 Workbook (legacy)
.xlt   → Excel Template
.xltx  → Excel Template (new)
.xltm  → Excel Macro-Enabled Template
```

### Text Formats
```
.csv   → Comma-Separated Values
.txt   → Plain Text
.prn   → Formatted Text (Space-delimited)
.dif   → Data Interchange Format
```

### Data Formats
```
.xml   → Extensible Markup Language
.json  → JavaScript Object Notation
.mdb   → Microsoft Access Database (old)
.accdb → Microsoft Access Database (new)
```

### Output Formats
```
.pdf   → Portable Document Format
.xps   → XML Paper Specification
.htm   → Web Page
.html  → Web Page
```

---

## Data Source Connection Strings

### Common Connection Examples

**CSV File:**
```
C:\Data\sales.csv
or
\\server\share\data\sales.csv
```

**SQL Server:**
```
Server=servername;Database=dbname;Trusted_Connection=True;
or
Server=servername;Database=dbname;User ID=username;Password=password;
```

**Access Database:**
```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\database.accdb;
```

**Web URL:**
```
https://example.com/data/prices.html
or
https://api.example.com/v1/data?format=json&key=APIKEY
```

**SharePoint:**
```
https://company.sharepoint.com/sites/sitename/
```

---

## Power Query M Language Basics

Power Query uses **M language** for transformations.

### Common M Functions

**Filter Rows:**
```m
= Table.SelectRows(Source, each [Sales] > 1000)

Shows only rows where Sales > 1000
```

**Add Custom Column:**
```m
= Table.AddColumn(Source, "Total", each [Quantity] * [Price])

Creates Total column
```

**Replace Values:**
```m
= Table.ReplaceValue(Source, "old", "new", Replacer.ReplaceText, {"Column1"})

Replaces "old" with "new" in Column1
```

**Remove Columns:**
```m
= Table.RemoveColumns(Source, {"Column1", "Column2"})

Removes specified columns
```

**Change Type:**
```m
= Table.TransformColumnTypes(Source, {{"Date", type date}, {"Amount", type number}})

Sets data types
```

**Group By:**
```m
= Table.Group(Source, {"Region"}, {{"Total Sales", each List.Sum([Sales]), type number}})

Groups by Region, sums Sales
```

### Viewing M Code

**Steps:**
1. Power Query Editor
2. Click any Applied Step
3. Formula bar shows M code
4. **View Tab → Advanced Editor** (see all code)

**Example:**
```m
let
    Source = Csv.Document(File.Contents("C:\data.csv")),
    Promoted = Table.PromoteHeaders(Source),
    Changed = Table.TransformColumnTypes(Promoted, {{"Sales", type number}}),
    Filtered = Table.SelectRows(Changed, each [Sales] > 1000)
in
    Filtered
```

---

## Security Considerations

### Importing Data

⚠️ **Risks:**
- Malicious macros in Excel files
- SQL injection in database queries
- Untrusted web sources
- Credential exposure

✅ **Best Practices:**
```
✅ Only import from trusted sources
✅ Scan files for malware
✅ Use read-only database accounts
✅ Never hardcode passwords in queries
✅ Validate data after import
✅ Enable Protected View for external files
```

### Exporting Data

⚠️ **Risks:**
- Sensitive data exposure
- Accidental sharing of confidential info
- Data leakage through metadata

✅ **Best Practices:**
```
✅ Review data before exporting
✅ Remove sensitive columns
✅ Password-protect files if needed
✅ Use secure file transfer methods
✅ Check document properties (File → Info)
✅ Consider PDF for sensitive data (harder to copy)
```

### Connection Security

```
✅ Use Windows Authentication when possible
✅ Encrypt database connections (SSL/TLS)
✅ Don't save passwords in connections
✅ Use OAuth for cloud services
✅ Regularly rotate credentials
✅ Monitor access logs
```

---

## Performance Optimization

### Large File Imports

**Problem:** Importing 1 million+ rows is slow

**Solutions:**

**1. Filter in Source**
```
SQL: Use WHERE clause to limit rows
Power Query: Filter early in process
Don't import unnecessary columns
```

**2. Use Binary Format**
```
Save as .xlsb instead of .xlsx
50% smaller file size
Opens/saves faster
```

**3. Disable Auto-Calculate**
```
Formulas Tab → Calculation Options → Manual
Import data
Re-enable Automatic
```

**4. Load to Data Model**
```
Instead of worksheet, load to Power Pivot
Handles millions of rows
Query in Data Model, not worksheet
```

**5. Split Large Files**
```
Import in chunks
Use Power Query to append
Or aggregate before loading
```

### Connection Performance

**Slow database queries:**
```
✅ Add indexes to database tables
✅ Optimize SQL query (avoid SELECT *)
✅ Limit date ranges
✅ Use views/stored procedures
✅ Increase connection timeout if needed
```

**Slow web imports:**
```
✅ Check internet connection
✅ Try off-peak hours
✅ Increase timeout settings
✅ Cache data locally, refresh less often
✅ Contact website if consistently slow
```

---

## Migration Strategies

### Moving from Manual Entry to Automated Import

**Phase 1: Assessment**
```
1. Document current manual process
2. Identify data sources
3. Determine refresh frequency needed
4. Test with sample data
```

**Phase 2: Setup**
```
1. Create Power Query connections
2. Build transformations
3. Create analysis templates (Pivots, Charts)
4. Test thoroughly
```

**Phase 3: Parallel Run**
```
1. Run both manual and automated
2. Compare results
3. Fix discrepancies
4. Build confidence
```

**Phase 4: Cutover**
```
1. Switch to automated only
2. Document refresh procedures
3. Train team members
4. Monitor for issues
```

### Upgrading Legacy Imports

**From Excel 2003 .xls files:**
```
1. Open in modern Excel
2. Save as .xlsx
3. Replace external links
4. Test all functionality
5. Update dependent files
```

**From manual copy/paste:**
```
1. Identify source systems
2. Check if API/export available
3. Create Power Query connection
4. Set up refresh schedule
5. Retire manual process
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- CSV = universal, no formatting
- Power Query = modern import/transform tool
- Refresh doesn't happen automatically
- File → Save As for format changes
- Data Tab → Get Data for imports
- Connection breaks if source moves
- Choose format based on use case

### Practice Deeply
- Opening CSV files and checking data
- Importing CSV with correct delimiters
- Using Text Import Wizard for custom imports
- Connecting to databases (if applicable)
- Using Power Query to clean data
- Creating queries that can be refreshed
- Transforming data in Power Query (filter, clean, calculate)
- Saving files in different formats (.xlsx, .csv, .pdf)
- Exporting data as CSV
- Creating and saving PDFs from Excel
- Setting up data refresh (manual and automatic)
- Importing data from web pages
- Troubleshooting import errors
- Managing external data connections
- Combining multiple files with Power Query
- Using Get Data → From Folder for multiple files
- Verifying data types after import
- Testing connections before production use

---

## Troubleshooting Checklist

Before asking for help, check:

```
☐ Is the source file accessible?
☐ Is the file path correct?
☐ Are credentials valid?
☐ Is the file open in another program?
☐ Has the file structure changed?
☐ Are you using the correct delimiter?
☐ Is the data type correct for each column?
☐ Have you refreshed the connection?
☐ Are there any error messages? (read them!)
☐ Does it work with a smaller sample file?
☐ Have you tried closing and reopening Excel?
☐ Is your Excel up to date?
```

---

## Quick Import/Export Decision Tree

```
Need to import data?
│
├─ From CSV/TXT file?
│  └─ Data Tab → From Text/CSV
│
├─ From Excel file?
│  └─ File → Open (or copy/paste)
│
├─ From database?
│  └─ Data Tab → Get Data → From Database
│
├─ From website?
│  └─ Data Tab → Get Data → From Web
│
└─ Multiple files?
   └─ Data Tab → Get Data → From Folder

Need to export data?
│
├─ For anyone to use?
│  └─ Save As → CSV
│
├─ Preserve all features?
│  └─ Save As → Excel Workbook (.xlsx)
│
├─ Read-only sharing?
│  └─ Export → PDF
│
├─ Upload to system?
│  └─ Check system requirements (usually CSV)
│
└─ Web publishing?
   └─ Save As → HTML
```

---

## Common Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| **"File not found"** | Source missing/moved | Update file path |
| **"Permission denied"** | No access rights | Check file permissions |
| **"Data source error"** | Connection failed | Verify source is available |
| **"Cannot connect"** | Network/auth issue | Check credentials/network |
| **"Timeout expired"** | Query too slow | Optimize query or increase timeout |
| **"Invalid connection"** | Connection string wrong | Re-enter connection info |
| **"Cannot refresh"** | Source changed/offline | Verify source and update connection |
| **"Type mismatch"** | Data type conflict | Check data types in Power Query |

---

## Resources for Learning More

### Power Query
- **Data Tab → Get Data** (explore all sources)
- **Power Query Editor** (experiment with transformations)
- Microsoft Learn: Power Query documentation
- Practice with sample datasets

### M Language
- **Power Query Editor → View → Advanced Editor**
- See generated M code
- Modify and test
- Learn by doing

### Connection Types
- Identify data sources in your organization
- Request access to databases
- Test connections in sandbox environment
- Document connection procedures

---

## Next Step

After this file, we move to:

**`17-named-ranges.md`**
- Creating and managing named ranges
- Using names in formulas
- Dynamic named ranges
- Name Manager
- Benefits of using names
- Named range best practices
- Scope (workbook vs worksheet)
- Named constants
