# Workbook Basics

This file covers how to create, manage, and organize workbooks and worksheets,
along with essential features like freezing panes and printing.

---

## Creating a New Workbook

### In Excel Online

**Method 1: From Excel Home**
1. Go to [excel.cloud.microsoft.com](https://excel.cloud.microsoft.com)
2. Click **"New blank workbook"**
3. Your workbook opens immediately

**Method 2: From File Menu**
1. Open any workbook
2. Click **File** → **New**
3. Choose **Blank workbook**

### In Excel Desktop

**Keyboard Shortcut:**
- `Ctrl + N` (Windows)
- `Cmd + N` (Mac)

**From File Menu:**
- File → New → Blank Workbook

---

## Understanding Workbooks vs Worksheets

### Visual Hierarchy
```
Workbook: SalesReport.xlsx
┌──────────────────────────────────────┐
│  File  Home  Insert  Formulas  Data  │
├──────────────────────────────────────┤
│                                      │
│        Worksheet Content             │
│                                      │
├──────────────────────────────────────┤
│ Sheet1  Sheet2  Sheet3  [+]         │ ← Sheet Tabs
└──────────────────────────────────────┘
         ↑
    Active Worksheet
```

### Key Concepts

**Workbook:**
- The entire Excel file
- Can contain multiple worksheets
- Saved as `.xlsx` file
- Think of it as a **book**

**Worksheet:**
- Individual sheet/page within the workbook
- Contains the grid of cells
- Think of it as a **page in the book**

### Default Setup
When you create a new workbook:
- **Excel Online:** 1 worksheet (Sheet1)
- **Excel Desktop:** 1-3 worksheets (configurable)

---

## Working with Worksheets

### Sheet Tab Location
```
┌──────────────────────────────────────┐
│                                      │
│        Main Worksheet Grid           │
│                                      │
├──────────────────────────────────────┤
│ ◄ ► │ Sheet1 │ Sheet2 │ Sheet3 │ + │ ← Sheet Tabs
└──────────────────────────────────────┘
   ↑      ↑        ↑         ↑      ↑
 Scroll  Active  Inactive  Inactive Add
 Arrows  Sheet   Sheets    Sheets   New
```

---

## Adding New Worksheets

### Method 1: Click the Plus Button
Click the **[+]** button next to the last sheet tab.

### Method 2: Keyboard Shortcut
- `Shift + F11` (Excel Desktop)
- Or right-click any sheet tab → **Insert**

### Method 3: Right-Click Menu
1. Right-click any sheet tab
2. Select **Insert**
3. Choose **Worksheet**

**Result:** A new blank worksheet is added.

---

## Renaming Worksheets

Default names like "Sheet1" aren't descriptive.

### How to Rename

**Method 1: Double-Click**
1. Double-click the sheet tab
2. Type new name
3. Press `Enter`

**Method 2: Right-Click**
1. Right-click the sheet tab
2. Select **Rename**
3. Type new name
4. Press `Enter`

### Best Practices for Sheet Names

**✅ Good Names:**
- `January_Sales`
- `Customer_Data`
- `Q1_Summary`
- `Budget_2024`

**❌ Avoid:**
- `Sheet1`, `Sheet2` (not descriptive)
- Names over 31 characters (too long)
- Special characters: `\ / ? * [ ]`
- Names starting or ending with apostrophe `'`

### Visual Example
```
Before:
│ Sheet1 │ Sheet2 │ Sheet3 │

After:
│ Jan_Sales │ Feb_Sales │ Summary │
```

---

## Moving and Copying Worksheets

### Moving Sheets (Same Workbook)

**Method 1: Drag and Drop**
1. Click and **hold** the sheet tab
2. Drag left or right
3. Release when positioned correctly

**Visual:**
```
Original:
│ Summary │ January │ February │

Dragging January left:
│ Summary │ ← January │ February │

Result:
│ January │ Summary │ February │
```

**Method 2: Right-Click Menu**
1. Right-click the sheet tab
2. Select **Move or Copy**
3. Choose position
4. Click **OK**

---

### Copying Sheets

**Method 1: Drag with Ctrl**
1. Hold `Ctrl` key (Windows) or `Option` key (Mac)
2. Drag the sheet tab
3. Release mouse, then release key

**Visual Indicator:**  
A small **[+]** icon appears next to cursor while dragging.

**Method 2: Right-Click Menu**
1. Right-click the sheet tab
2. Select **Move or Copy**
3. ✓ Check **"Create a copy"**
4. Choose position
5. Click **OK**

**Result:** You get an exact duplicate (e.g., `January (2)`)

---

### Moving/Copying to Another Workbook

1. Right-click the sheet tab
2. Select **Move or Copy**
3. In **"To book:"** dropdown, select destination workbook
4. Check **"Create a copy"** if you want to keep original
5. Click **OK**

⚠️ **Note:** Both workbooks must be open in Excel Desktop. This is limited in Excel Online.

---

## Deleting Worksheets

### How to Delete

**Method 1: Right-Click**
1. Right-click the sheet tab
2. Select **Delete**
3. Confirm if prompted

**Method 2: Alt + E + L (Desktop Only)**
1. Select the sheet
2. Press `Alt`, then `E`, then `L`

### ⚠️ CRITICAL WARNING

**Deleting a worksheet is PERMANENT.**

You **cannot** undo worksheet deletion.

**Best Practice:**
- Create a backup copy before deleting
- Or rename it to "OLD_[SheetName]" instead of deleting
- Double-check you're deleting the correct sheet

---

## Hiding and Unhiding Worksheets

Sometimes you want to hide sheets without deleting them.

### Hide a Worksheet

**Method 1: Right-Click**
1. Right-click the sheet tab
2. Select **Hide**

**Result:** Sheet disappears from tabs but data remains in the workbook.

**Use Cases:**
- Hide calculation sheets
- Hide template sheets
- Hide reference data
- Clean up tab clutter

### Unhide a Worksheet

**Method 1: Right-Click**
1. Right-click **any** visible sheet tab
2. Select **Unhide**
3. Choose which sheet to unhide
4. Click **OK**

### Visual Flow
```
Before Hide:
│ Data │ Calculations │ Summary │
          ↓ (Hide)
After Hide:
│ Data │ Summary │
          ↓ (Unhide)
Back:
│ Data │ Calculations │ Summary │
```

---

## Sheet Tab Colors

Color-coding sheets helps organize workbooks.

### How to Color Sheet Tabs

1. Right-click the sheet tab
2. Select **Tab Color**
3. Choose a color

### Practical Color Coding Example
```
│ Jan_Sales │ Feb_Sales │ Mar_Sales │ Q1_Summary │ Archive │
   (Blue)      (Blue)       (Blue)      (Green)     (Gray)
    ↑           ↑            ↑             ↑           ↑
  Monthly     Monthly      Monthly      Summary    Old Data
```

**Best Practice:**
- **Blue:** Monthly data sheets
- **Green:** Summary/dashboard sheets
- **Orange:** Calculation/working sheets
- **Red:** Important/urgent sheets
- **Gray:** Archive/reference sheets

---

## Navigating Between Worksheets

### Using Mouse
Click the sheet tab you want to view.

### Using Keyboard

| Shortcut | Action |
|----------|--------|
| `Ctrl + Page Down` | Move to next sheet (right) |
| `Ctrl + Page Up` | Move to previous sheet (left) |

### Scroll Through Sheet Tabs

If you have many sheets, use the scroll arrows:

```
┌──────────────────────────────────┐
│ ◄ ► │ Sheet5 │ Sheet6 │ Sheet7 │
└──────────────────────────────────┘
   ↑
Click to scroll through tabs
```

---

## Selecting Multiple Worksheets

Sometimes you want to work on multiple sheets at once.

### Group Sheets

**Method 1: Ctrl + Click**
1. Click first sheet tab
2. Hold `Ctrl`
3. Click additional sheet tabs

**Method 2: Shift + Click (Consecutive Sheets)**
1. Click first sheet tab
2. Hold `Shift`
3. Click last sheet tab in range

**Result:** All sheets between first and last are selected.

### Visual Indicator
```
Selected sheets appear highlighted:
│ [Jan] │ [Feb] │ [Mar] │ Summary │
   ↑       ↑       ↑
 All three selected (grouped)
```

### When Sheets are Grouped

**What happens:**
- Changes apply to **all** grouped sheets simultaneously
- Formatting applied to all
- Data entered in all
- Useful for identical structure across sheets

**Warning Indicator:**  
Title bar shows `[Group]` when sheets are grouped.

### Ungroup Sheets

**Method 1:** Click any non-selected sheet tab

**Method 2:** Right-click any grouped sheet → **Ungroup Sheets**

---

## Freeze Panes

**Freeze Panes** keeps specific rows and/or columns visible while scrolling.

This is **extremely useful** for large datasets.

### Why Use Freeze Panes?

**Problem:**
```
Scrolling down, headers disappear:

Initially:
│ Name    │ Sales  │ Region │
│ John    │ 1500   │ East   │
│ Sarah   │ 2000   │ West   │
       ↓ (Scroll down)
After scrolling (headers gone!):
│ Mike    │ 1800   │ North  │
│ Lisa    │ 2200   │ South  │
     ↑ What column is this?
```

**Solution: Freeze top row!**

---

### Freeze Top Row

**Steps:**
1. Click **View** tab
2. Click **Freeze Panes** dropdown
3. Select **Freeze Top Row**

**Result:**
```
Row 1 stays visible while scrolling:

│ Name    │ Sales  │ Region │ ← Always visible
├─────────┴────────┴────────┤
│ Mike    │ 1800   │ North  │ ← Scrolled content
│ Lisa    │ 2200   │ South  │
```

---

### Freeze First Column

**Steps:**
1. Click **View** tab
2. Click **Freeze Panes** dropdown
3. Select **Freeze First Column**

**Result:** Column A stays visible when scrolling right.

---

### Freeze Panes (Custom)

Freeze both rows AND columns at a specific position.

**Steps:**
1. Click the cell **below and to the right** of where you want to freeze
2. Click **View** tab
3. Click **Freeze Panes** → **Freeze Panes**

**Example:**
```
To freeze rows 1-2 AND columns A-B:
Click cell C3, then freeze.

Result:
┌─────────┬─────────┬─────────────────┐
│ A (frozen) │ B (frozen) │ C  D  E ... (scroll) │
├─────────┼─────────┼─────────────────┤
│ Row 1 (frozen)      │                 │
│ Row 2 (frozen)      │                 │
├─────────┴─────────┼─────────────────┤
│ Row 3              │ Scrollable area │
│ Row 4              │                 │
```

### Visual Rules
```
Click cell C3 to freeze:

     A    B  │  C    D    E
  ┌─────────┼──────────────
1 │ Frozen  │  Scrolls horizontally
2 │ Frozen  │  Scrolls horizontally
──┼─────────┼──────────────
3 │ Scrolls │  Scrolls both ways
4 │ vertically│
```

**Everything to the left and above the selected cell is frozen.**

---

### Unfreeze Panes

**Steps:**
1. Click **View** tab
2. Click **Freeze Panes** dropdown
3. Select **Unfreeze Panes**

---

## Split Panes (Alternative to Freeze)

**Split Panes** divides the worksheet into separate scrollable areas.

### Split vs Freeze

| Feature | Freeze Panes | Split Panes |
|---------|--------------|-------------|
| **Purpose** | Keep headers visible | View different parts of same sheet |
| **Scrolling** | Frozen area doesn't move | All areas scroll independently |
| **Use case** | Large datasets with headers | Compare distant cells |

### How to Split

1. Select the cell where you want to split
2. Click **View** tab
3. Click **Split**

**Result:** Gray split bars appear allowing independent scrolling.

### Remove Split

Click **View** → **Split** again (it toggles).

---

## Zoom and View Options

### Zoom Level

**Method 1: Zoom Slider (Bottom-Right)**
```
┌──────────────────────────────┐
│                              │
│        Worksheet             │
│                              │
├──────────────────────────────┤
│             [-][====][+] 100%│ ← Zoom Slider
└──────────────────────────────┘
```

**Method 2: View Tab**
1. Click **View** tab
2. Click **Zoom**
3. Select percentage (50%, 75%, 100%, 150%, 200%)

**Keyboard Shortcut:**
- `Ctrl + Mouse Wheel` - Zoom in/out

### Common Zoom Levels

| Zoom | Use Case |
|------|----------|
| 50-75% | See more data at once |
| 100% | Default, comfortable viewing |
| 150-200% | Easier reading, presentations |

---

## View Modes

Excel has three view modes:

### 1. Normal View (Default)
Standard grid view for working with data.

### 2. Page Layout View
Shows how the worksheet will look when printed.
- Displays margins
- Shows headers/footers
- Indicates page breaks

### 3. Page Break Preview
Shows where pages will break when printing.
- Blue lines = automatic page breaks
- Drag lines to adjust where pages break

### Switching Views

**Location:** Bottom-right of Excel window

```
┌──────────────────────────────┐
│                              │
│        Worksheet             │
│                              │
├──────────────────────────────┤
│ [Normal] [Page] [Break]  Zoom│
└──────────────────────────────┘
     ↑       ↑      ↑
  Normal  Layout  Preview
```

**Or:** Click **View** tab → Choose view mode

---

## Page Setup and Printing

### Access Page Setup

**Method 1: Page Layout Tab**
1. Click **Page Layout** tab
2. Use options in **Page Setup** group

**Method 2: Page Setup Dialog**
1. Click **Page Layout** tab
2. Click small arrow in **Page Setup** group corner

---

### Page Orientation

**Portrait:**
```
┌───────┐
│       │
│       │
│       │
│       │
└───────┘
Taller than wide
```

**Landscape:**
```
┌─────────────┐
│             │
│             │
└─────────────┘
Wider than tall
```

**Change Orientation:**
1. **Page Layout** tab
2. Click **Orientation**
3. Choose **Portrait** or **Landscape**

---

### Page Size

**Standard Sizes:**
- Letter (8.5" × 11") - US standard
- A4 (210mm × 297mm) - International standard
- Legal (8.5" × 14")

**Change Size:**
1. **Page Layout** tab
2. Click **Size**
3. Choose paper size

---

### Margins

**Preset Margins:**
- Normal: 0.75" all sides
- Wide: 1" all sides
- Narrow: 0.25" all sides

**Custom Margins:**
1. **Page Layout** tab
2. Click **Margins** → **Custom Margins**
3. Enter values
4. Click **OK**

---

### Print Area

**Set Print Area:**
1. Select the range you want to print
2. **Page Layout** tab
3. Click **Print Area** → **Set Print Area**

**Result:** Only the selected area will print.

**Clear Print Area:**
**Page Layout** tab → **Print Area** → **Clear Print Area**

---

### Headers and Footers

Headers appear at the top of each printed page.  
Footers appear at the bottom.

**Add Header/Footer:**
1. Click **Insert** tab
2. Click **Header & Footer**
3. Click in header/footer area
4. Type text or insert elements:
   - Page number
   - Date/Time
   - File name
   - Sheet name

**Quick Elements:**
While in header/footer, click **Header & Footer Tools** tab for quick inserts.

---

### Print Preview

**Before printing, always preview!**

**Open Print Preview:**
1. Click **File** → **Print**
2. Or press `Ctrl + P`

**What You See:**
- How pages will look
- How many pages
- Page breaks
- Margins

**Adjust from Preview:**
- Change orientation
- Adjust margins
- Fit to page options

---

### Fit to Page Options

**Problem:** Data spans multiple pages

**Solution:** Scale to fit

**Options:**

**1. Fit Sheet on One Page:**
```
Page Layout tab → Scale to Fit group
Width: 1 page
Height: 1 page
```

**2. Fit All Columns on One Page:**
```
Width: 1 page
Height: Automatic
```

**3. Fit All Rows on One Page:**
```
Width: Automatic
Height: 1 page
```

⚠️ **Warning:** Scaling too much makes text tiny and unreadable. Use judiciously.

---

### Printing Gridlines and Headings

By default, gridlines and row/column headers don't print.

**Print Gridlines:**
1. **Page Layout** tab
2. In **Sheet Options** group
3. Check **Print** under **Gridlines**

**Print Headings (Row/Column Labels):**
1. **Page Layout** tab
2. In **Sheet Options** group
3. Check **Print** under **Headings**

**Result:**
```
Prints with A, B, C... and 1, 2, 3... visible
```

---

## Saving Workbooks

### Save vs Save As

**Save (Ctrl + S):**
- Updates the existing file
- Use for ongoing work

**Save As (Ctrl + Shift + S or F12):**
- Creates a new file
- Use for creating copies
- Use to change file name or location
- Use to change file format

---

### Saving in Excel Online

**Auto-Save:**
Excel Online **automatically saves** your work to OneDrive.

Look for the **"Saved"** indicator near the file name.

**Rename File:**
Click the file name at top and type new name.

**Download a Copy:**
1. Click **File** → **Save As**
2. Click **Download a Copy**
3. Choose format (Excel Workbook, PDF, etc.)

---

### Saving in Excel Desktop

**Initial Save:**
1. Press `Ctrl + S`
2. Choose location (OneDrive, This PC, etc.)
3. Enter file name
4. Choose file type (if not `.xlsx`)
5. Click **Save**

**Subsequent Saves:**
- `Ctrl + S` quickly updates the file

---

### File Formats When Saving

| Format | Extension | When to Use |
|--------|-----------|-------------|
| Excel Workbook | `.xlsx` | Standard format (default) |
| Excel Macro-Enabled | `.xlsm` | Contains macros/VBA code |
| Excel 97-2003 | `.xls` | Compatibility with old Excel |
| CSV | `.csv` | Plain text, import to other apps |
| PDF | `.pdf` | Share non-editable version |

**Recommendation:** Use `.xlsx` for most work.

---

## Opening Existing Workbooks

### Excel Online

**From Home Screen:**
1. Go to [excel.cloud.microsoft.com](https://excel.cloud.microsoft.com)
2. Click on recent file
3. Or click **Open** → browse OneDrive

**Upload File:**
1. Click **Upload**
2. Choose file from computer
3. File opens in Excel Online

---

### Excel Desktop

**Method 1: File Menu**
1. Click **File** → **Open**
2. Browse to file location
3. Select file
4. Click **Open**

**Method 2: Keyboard**
Press `Ctrl + O`

**Method 3: Recent Files**
1. Click **File**
2. Recent files appear on the left
3. Click to open

---

## Protecting Workbooks and Worksheets

### Protect Sheet (Prevent Editing)

**Why:** Prevent accidental changes to formulas or structure.

**Steps:**
1. Click **Review** tab
2. Click **Protect Sheet**
3. Optional: Set password
4. Check options (what users can still do)
5. Click **OK**

**What Gets Protected:**
- Cell content
- Formulas
- Formatting (if selected)

**What Users Can Still Do:**
- Select cells (if allowed)
- View data
- Copy data (if allowed)

**Unprotect:**
1. **Review** tab → **Unprotect Sheet**
2. Enter password (if set)

---

### Protect Workbook (Structure)

**Why:** Prevent adding/deleting/renaming sheets.

**Steps:**
1. Click **Review** tab
2. Click **Protect Workbook**
3. Optional: Set password
4. Click **OK**

**What Gets Protected:**
- Can't add/delete sheets
- Can't rename sheets
- Can't hide/unhide sheets
- Can't move sheets

**Unprotect:**
1. **Review** tab → **Unprotect Workbook**

---

## Best Practices

### 1. Use Descriptive Sheet Names
```
✅ Good: Jan_Sales, Customer_List, Summary_Report
❌ Bad: Sheet1, Sheet2, Sheet3
```

### 2. Organize Sheets Logically
```
Left to right: Input → Calculations → Summary
│ Raw_Data │ Calculations │ Dashboard │
```

### 3. Freeze Panes for Large Datasets
Always freeze headers when working with data.

### 4. Color-Code Sheet Tabs
Helps navigate complex workbooks quickly.

### 5. Use Print Preview Before Printing
Saves paper and catches formatting issues.

### 6. Save Frequently (Desktop)
`Ctrl + S` every few minutes (or enable AutoSave to OneDrive).

### 7. Hide, Don't Delete
Hide sheets you might need later instead of deleting.

### 8. Protect Important Sheets
Prevent accidental changes to formulas and structure.

---

## Common Mistakes

### Mistake 1: Not Freezing Panes
Working with large datasets without freezing headers = constant scrolling up to remember what column is what.

### Mistake 2: Deleting Sheets Without Backup
Sheet deletion is permanent. Always make a copy first or rename to "OLD_".

### Mistake 3: Too Many Worksheets
Having 50+ sheets makes navigation difficult. Consider splitting into multiple workbooks.

### Mistake 4: Not Using Print Preview
Printing without preview wastes paper and time.

### Mistake 5: Forgetting to Ungroup Sheets
Making changes while sheets are grouped affects ALL grouped sheets (often unexpected).

---

## What to PRACTICE vs MEMORIZE

### Memorize
- `Ctrl + N` - New workbook
- `Ctrl + S` - Save
- `Ctrl + Page Up/Down` - Navigate sheets
- Freeze Panes location: **View** tab
- Sheet deletion is permanent

### Practice Deeply
- Creating and renaming worksheets
- Moving and copying sheets
- Using Freeze Panes effectively
- Setting up print areas and page setup
- Organizing workbooks with multiple sheets
- Protecting sheets and workbooks
- Using zoom and view modes
- Grouping and ungrouping sheets

---

## Quick Reference: Workbook Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + N` | New workbook |
| `Ctrl + O` | Open workbook |
| `Ctrl + S` | Save |
| `Ctrl + P` | Print/Print Preview |
| `Ctrl + Page Down` | Next sheet |
| `Ctrl + Page Up` | Previous sheet |
| `Shift + F11` | Insert new worksheet |
| `Alt + Page Down` | One screen right |
| `Alt + Page Up` | One screen left |
| `Ctrl + F1` | Show/Hide Ribbon |

---

## Next Step

After this file, we move to:

**`02-cell-references-and-ranges.md`**
- Understanding cell references (A1, B2, etc.)
- Absolute vs Relative references ($A$1)
- Mixed references ($A1, A$1)
- Named ranges
- Selecting and working with ranges
- The Name Box
