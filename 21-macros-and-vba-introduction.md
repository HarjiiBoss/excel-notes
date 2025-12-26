# Macros and VBA Introduction

This file covers Excel macros and VBA (Visual Basic for Applications) - tools for automating repetitive tasks and extending Excel's capabilities beyond formulas and built-in features.

---

## What is a Macro?

**Macro** = A recorded or written sequence of actions that Excel can replay automatically.

### Purpose
- **Automate repetitive tasks:** Do in 1 click what takes 20 clicks
- **Standardize processes:** Everyone does the task the same way
- **Save time:** Minutes become seconds
- **Reduce errors:** Computer doesn't forget steps

### Real-World Examples

**Without Macro:**
```
Monthly Report Process:
1. Format headers (bold, color, size)
2. Apply borders to data
3. Freeze top row
4. Auto-fit columns
5. Add company logo
6. Set print area
7. Adjust margins
8. Add footer with date

❌ Takes 5 minutes
❌ Easy to forget steps
❌ Inconsistent formatting
```

**With Macro:**
```
Monthly Report Process:
1. Click "Format Report" button

✅ Takes 2 seconds
✅ Never misses steps
✅ Always looks identical
```

### Visual Concept

```
┌─────────────────────────────────────────────┐
│              MACRO WORKFLOW                 │
│                                             │
│  You Do Once:              Computer Does:   │
│  ┌──────────────┐         ┌──────────────┐ │
│  │ 1. Format    │         │ Repeats all  │ │
│  │ 2. Color     │  ─────> │ steps in     │ │
│  │ 3. Borders   │ Record  │ 1 second     │ │
│  │ 4. Adjust    │         │ every time   │ │
│  └──────────────┘         └──────────────┘ │
│                                             │
│  Click button → All steps execute           │
└─────────────────────────────────────────────┘
```

---

## Macro vs Formula vs Power Query

### When to Use Each

**Formulas:**
```
✅ Calculate values
✅ Dynamic results that update
✅ Reference other cells
✅ Built-in functions available

Example: =SUM(A1:A10)
```

**Power Query:**
```
✅ Import and transform data
✅ Clean messy data
✅ Combine multiple sources
✅ Repeatable ETL process

Example: Import CSV, clean, filter, load
```

**Macros:**
```
✅ Automate formatting
✅ Perform multi-step processes
✅ Interact with workbook structure
✅ Create custom commands

Example: Format report, add chart, save as PDF
```

### Comparison Table

| Task | Best Tool | Why |
|------|-----------|-----|
| Calculate sales totals | Formula | Dynamic, updates automatically |
| Import and clean CSV | Power Query | Repeatable, handles changes |
| Format 50 sheets identically | Macro | Multi-step automation |
| Look up customer info | Formula (XLOOKUP) | Real-time lookup |
| Combine 100 CSV files | Power Query | Handles file structures |
| Add logo to every sheet | Macro | Workbook interaction |
| Sum by category | Formula (SUMIF) | Dynamic calculation |
| Remove blank rows | Power Query | Data transformation |
| Create monthly report layout | Macro | Complex formatting |

---

## Macro Basics: Recording Your First Macro

### Accessing the Developer Tab

⚠️ **Important:** Macros require the Developer tab, which is hidden by default.

**Enable Developer Tab:**

```
1. File → Options
2. Click "Customize Ribbon"
3. Right side: Check ☑ Developer
4. Click OK

Developer tab appears in ribbon
```

**Visual:**
```
Excel Ribbon:
┌────────────────────────────────────────────┐
│ Home | Insert | Draw | Page Layout |      │
│ Formulas | Data | Review | View |         │
│ DEVELOPER ← Now visible                    │
└────────────────────────────────────────────┘
```

### Recording a Macro

**Developer Tab → Code section:**

```
┌──────────────────────────────────┐
│ Developer Tab                    │
├──────────────────────────────────┤
│ Visual Basic                     │
│ Macros                           │
│ Record Macro  ← Start here       │
│ Use Relative References          │
│ Macro Security                   │
└──────────────────────────────────┘
```

### Step-by-Step: Record Simple Macro

**Example: Format Headers**

**Step 1: Start Recording**
```
1. Developer Tab → Record Macro
2. Macro name: FormatHeaders (no spaces!)
3. Shortcut key: Ctrl+Shift+H (optional)
4. Store in: This Workbook
5. Description: "Formats column headers"
6. Click OK

🔴 Recording... (indicator appears in status bar)
```

**Record Macro Dialog:**
```
┌──────────────────────────────────────┐
│ Record Macro                         │
├──────────────────────────────────────┤
│ Macro name:                          │
│ [FormatHeaders________________]      │
│                                      │
│ Shortcut key:                        │
│ Ctrl+Shift+[H]                       │
│                                      │
│ Store macro in:                      │
│ [This Workbook          ▼]           │
│                                      │
│ Description:                         │
│ [Formats column headers...]          │
│                                      │
│ [OK] [Cancel]                        │
└──────────────────────────────────────┘
```

**Step 2: Perform Actions**
```
While recording, do these actions:
1. Select row 1
2. Bold (Ctrl+B)
3. Fill color: Blue
4. Font color: White
5. Font size: 12
6. Center align

Every action is recorded!
```

**Step 3: Stop Recording**
```
Developer Tab → Stop Recording
Or: Click 🔴 in status bar

Macro saved!
```

### Running Your Macro

**Method 1: Keyboard Shortcut**
```
Press: Ctrl+Shift+H
Macro runs instantly!
```

**Method 2: Macros Dialog**
```
1. Developer Tab → Macros
2. Select: FormatHeaders
3. Click: Run

Or: Alt+F8 (shortcut to open dialog)
```

**Macros Dialog:**
```
┌──────────────────────────────────────┐
│ Macro                                │
├──────────────────────────────────────┤
│ Macro name:                          │
│ ┌──────────────────────────────┐    │
│ │ FormatHeaders                │    │
│ │ CreateMonthlyReport          │    │
│ │ ExportToPDF                  │    │
│ └──────────────────────────────┘    │
│                                      │
│ [Run] [Step Into] [Edit]             │
│ [Create] [Delete] [Options...]       │
│                                      │
│ Macros in: [All Open Workbooks ▼]   │
│ Description: Formats column headers  │
└──────────────────────────────────────┘
```

**Method 3: Button (covered later)**

---

## Macro Naming Rules

### Valid Names

```
✅ FormatHeaders
✅ MonthlyReport
✅ Export_To_PDF
✅ Calculate_Totals
✅ MyMacro123

Must start with letter
Can contain letters, numbers, underscores
No spaces allowed
Max 255 characters
Not case-sensitive (Excel treats same)
```

### Invalid Names

```
❌ Format Headers    (space)
❌ 123Macro          (starts with number)
❌ Export-To-PDF     (hyphen not allowed)
❌ My Macro!         (special character)
❌ Print             (Excel keyword/function)
```

---

## Relative vs Absolute Recording

### Absolute References (Default)

**Records exact cell addresses:**

```
You record:
1. Click cell B2
2. Type "Total"
3. Bold

Macro does:
1. Go to cell B2
2. Type "Total"
3. Bold

Every time → always cell B2
```

**When to use:**
- Always format specific cells
- Same structure every time
- Fixed report layouts

**Example - Always format A1:**
```
✅ Absolute recording
Click A1, make bold
→ Macro always formats A1
```

### Relative References

**Records movements from current position:**

```
You record (from A1):
1. Move right 1 cell (→ B1)
2. Type "Total"
3. Bold

Macro does:
1. Move right 1 cell from current position
2. Type "Total"
3. Bold

Start at C5 → formats D5
Start at A10 → formats B10
```

**When to use:**
- Apply to different locations
- Flexible positioning
- Process multiple ranges

**Enable relative recording:**
```
Developer Tab → Use Relative References
(Toggle on before recording)

Icon appears pressed/highlighted when active
```

### Comparison Example

**Task:** Add "Total" label next to selected cell

**Absolute Recording:**
```
Select B5, click Record, type "Total" in C5
Run from anywhere → always puts "Total" in C5

❌ Not flexible
```

**Relative Recording:**
```
Enable Relative References
Select B5, click Record, move right, type "Total"
Run from B5 → puts in C5
Run from B10 → puts in C10
Run from D3 → puts in E3

✅ Flexible, works anywhere
```

**Visual:**
```
Absolute:
Start: [B5] → Run → [C5] gets "Total"
Start: [D10] → Run → [C5] gets "Total" (same!)

Relative:
Start: [B5] → Run → [C5] gets "Total"
Start: [D10] → Run → [E10] gets "Total" (relative!)
```

---

## Assigning Macros to Buttons

### Why Use Buttons?

```
✅ Visual reminder of available macros
✅ No need to remember shortcuts
✅ Easy for others to use
✅ Professional dashboard appearance
✅ Can be labeled clearly
```

### Creating a Button

**Method 1: Shape Button**

**Step 1: Insert Shape**
```
1. Insert Tab → Shapes
2. Choose: Rectangle (or any shape)
3. Draw shape on worksheet
4. Right-click shape → Edit Text
5. Type: "Format Report" (or macro name)
```

**Step 2: Assign Macro**
```
1. Right-click shape
2. Choose: Assign Macro
3. Select macro from list
4. Click OK

Button ready to use!
Click button → macro runs
```

**Method 2: Form Control Button**

```
1. Developer Tab → Insert → Form Controls
2. Click Button icon (first option)
3. Draw button on sheet
4. "Assign Macro" dialog appears automatically
5. Select macro → OK
6. Right-click button → Edit Text to rename
```

**Visual:**
```
Worksheet with buttons:
┌──────────────────────────────────────┐
│ A     B     C     D     E     F      │
├──────────────────────────────────────┤
│                                      │
│  ┌───────────────┐  ┌──────────────┐│
│  │ Format Report │  │ Export PDF   ││
│  └───────────────┘  └──────────────┘│
│                                      │
│  ┌───────────────┐  ┌──────────────┐│
│  │ Clear Filters │  │ Refresh Data ││
│  └───────────────┘  └──────────────┘│
│                                      │
│  Data appears below buttons          │
└──────────────────────────────────────┘
```

### Formatting Buttons

**Right-click button → Format Shape:**
```
Fill: Choose color
Border: Style and color
Effects: Shadow, glow, 3D
Text: Font, size, color, alignment
```

**Tips:**
```
✅ Use consistent colors
✅ Label clearly (verb + object)
✅ Group related buttons
✅ Use appropriate size
✅ Add icons (via Insert → Icons, group with shape)
```

---

## Macro Security Settings

### Why Security Matters

```
⚠️ Macros can contain malicious code
⚠️ Can delete files, steal data, corrupt workbooks
⚠️ Only run macros from trusted sources
```

### Security Levels

**Access settings:**
```
Developer Tab → Macro Security

Or:
File → Options → Trust Center → Trust Center Settings → Macro Settings
```

**Security Options:**
```
┌──────────────────────────────────────────────┐
│ Trust Center - Macro Settings               │
├──────────────────────────────────────────────┤
│ ○ Disable all macros without notification   │
│   ↳ All macros disabled (most secure)       │
│                                              │
│ ● Disable all macros with notification      │
│   ↳ Default - shows warning, you choose     │
│                                              │
│ ○ Disable all except digitally signed       │
│   ↳ Only macros from trusted publishers     │
│                                              │
│ ○ Enable all macros (not recommended)       │
│   ↳ No protection - risky!                  │
│                                              │
│ ☑ Trust access to VBA project object model  │
└──────────────────────────────────────────────┘
```

### Recommended Setting

```
✅ Disable all macros with notification (default)

When you open workbook with macros:
┌──────────────────────────────────────┐
│ ⚠️ SECURITY WARNING                  │
│ Macros have been disabled.           │
│ [Enable Content]                     │
└──────────────────────────────────────┘

Only click Enable if you trust the source!
```

### Trusted Locations

**Add folder as trusted:**
```
1. File → Options → Trust Center
2. Trust Center Settings
3. Trusted Locations
4. Add new location
5. Browse to folder (e.g., C:\MyMacros\)
6. ☑ Subfolders of this location are also trusted
7. OK

Files in this folder always enabled - no warning
```

**Use for:**
```
✅ Your personal macro files
✅ Company-approved macro folder
✅ Development/testing folder

❌ Downloads folder
❌ Shared network drives (unless controlled)
```

---

## Introduction to VBA Editor

### What is VBA?

**VBA = Visual Basic for Applications**

- Programming language built into Excel
- More powerful than recording
- Can do things macros can't
- Write custom functions
- Complex logic and decisions

**Recorded macro = VBA code automatically generated**

### Opening VBA Editor

**Methods:**
```
1. Developer Tab → Visual Basic
2. Alt+F11 (keyboard shortcut)
3. In Macros dialog → Edit button
```

### VBA Editor Interface

```
┌─────────────────────────────────────────────────────┐
│ Microsoft Visual Basic for Applications - Excel    │
├────────────┬────────────────────────────────────────┤
│ Project    │ Code Window                            │
│ Explorer   │                                        │
│            │ Sub FormatHeaders()                    │
│ ▼ VBAProj  │     Range("A1").Select                 │
│   ├─ Micr │     Selection.Font.Bold = True         │
│   ├─ Sheet│     Selection.Interior.Color = Blue    │
│   └─ Modul│ End Sub                                │
│     └─ Mod│                                        │
│            │                                        │
├────────────┼────────────────────────────────────────┤
│ Properties │ Immediate Window                       │
│            │ (for testing/debugging)                │
│            │                                        │
└────────────┴────────────────────────────────────────┘
```

### Key Components

**1. Project Explorer (Left)**
```
Shows workbook structure:
├─ Microsoft Excel Objects
│   ├─ ThisWorkbook (workbook-level code)
│   ├─ Sheet1 (sheet-level code)
│   ├─ Sheet2
│   └─ Sheet3
└─ Modules
    └─ Module1 (recorded macros stored here)
```

**2. Code Window (Center)**
```
Where you write/edit VBA code
Each module, sheet, or workbook has own code area
Syntax highlighting (keywords in blue)
```

**3. Properties Window (Bottom Left)**
```
Shows properties of selected object
Change names, settings
```

**4. Immediate Window (Bottom)**
```
Test code snippets
Debug and inspect values
Press Ctrl+G to show if hidden
```

### Finding Your Recorded Macro

```
1. Open VBA Editor (Alt+F11)
2. Project Explorer → VBAProject → Modules
3. Double-click Module1
4. Your macro code appears in Code Window

Example:
Sub FormatHeaders()
    Range("A1:Z1").Select
    Selection.Font.Bold = True
    Selection.Interior.Color = RGB(0, 0, 255)
    Selection.Font.Color = RGB(255, 255, 255)
End Sub
```

---

## Understanding Basic VBA Code

### Structure of a Macro

```vba
Sub MacroName()
    ' Code goes here
    ' Comments start with apostrophe
    Range("A1").Value = "Hello"
End Sub
```

**Parts:**
```
Sub          = Subroutine (start of macro)
MacroName()  = Name you gave it
End Sub      = End of macro

Everything between = the actions
```

### Common VBA Commands

**Select cells:**
```vba
Range("A1").Select                  ' Single cell
Range("A1:B10").Select              ' Range
Range("A:A").Select                 ' Entire column
Rows("1:1").Select                  ' Entire row
Cells(2, 3).Select                  ' Row 2, Column 3 (C2)
```

**Modify cells:**
```vba
Range("A1").Value = "Sales"         ' Enter text
Range("B2").Value = 1000            ' Enter number
Range("C3").Formula = "=A1+B1"      ' Enter formula
```

**Format cells:**
```vba
Range("A1").Font.Bold = True        ' Bold
Range("A1").Font.Size = 14          ' Font size
Range("A1").Font.Color = RGB(255, 0, 0)  ' Red text
Range("A1").Interior.Color = RGB(255, 255, 0)  ' Yellow fill
```

**Other actions:**
```vba
Rows("1:1").Delete                  ' Delete row
Columns("A:A").Insert               ' Insert column
ActiveSheet.Name = "Sales Data"     ' Rename sheet
Range("A1:B10").Copy                ' Copy range
Range("D1").PasteSpecial            ' Paste
```

### Reading Recorded Code

**Example recorded macro:**
```vba
Sub FormatReport()
    Range("A1:E1").Select
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    With Selection.Interior
        .Pattern = xlSolid
        .Color = RGB(68, 114, 196)
    End With
    Range("A2").Select
End Sub
```

**Breaking it down:**
```
Line 1: Select range A1:E1
Line 2: Make selected range bold
Line 3: Set font size to 12
Lines 4-7: Fill with blue color (compact way)
Line 8: Select cell A2 (where recording stopped)
```

### Variables (Basic Concept)

**Variables store information:**

```vba
Sub UseVariables()
    Dim salesAmount As Double       ' Declare variable
    salesAmount = 1000              ' Assign value
    Range("A1").Value = salesAmount ' Use variable
End Sub
```

**Variable types:**
```
String   = Text ("Hello")
Integer  = Whole numbers (-32768 to 32767)
Long     = Larger whole numbers
Double   = Decimal numbers (1.5, 3.14159)
Boolean  = True/False
Date     = Dates and times
```

---

## Editing Recorded Macros

### Why Edit?

```
✅ Remove unnecessary steps (Select commands)
✅ Make code more efficient
✅ Add flexibility (variables, logic)
✅ Fix issues
✅ Combine multiple macros
```

### Common Edits

**Remove unnecessary Select:**

**Recorded (inefficient):**
```vba
Sub FormatCell()
    Range("A1").Select
    Selection.Font.Bold = True
    Selection.Value = "Total"
End Sub
```

**Edited (efficient):**
```vba
Sub FormatCell()
    Range("A1").Font.Bold = True
    Range("A1").Value = "Total"
End Sub
```

**Or even better:**
```vba
Sub FormatCell()
    With Range("A1")
        .Font.Bold = True
        .Value = "Total"
    End With
End Sub
```

### Testing Code Changes

**Step 1: Edit in VBA Editor**
```
1. Open VBA Editor (Alt+F11)
2. Find macro code
3. Make changes
4. Save (Ctrl+S)
```

**Step 2: Test**
```
1. Return to Excel (Alt+F11 again)
2. Run macro (Alt+F8 or button)
3. Check results
```

**Step 3: Debug if needed**
```
If error occurs:
- VBA Editor opens automatically
- Error line highlighted in yellow
- Read error message
- Fix code
- Test again
```

---

## Basic VBA Concepts

### Comments

```vba
' This is a comment - explains what code does
' Ignored by Excel when running
' Use apostrophe to start comment

Sub MyMacro()
    ' Format the headers
    Range("A1:E1").Font.Bold = True
    
    ' Enter current date
    Range("F1").Value = Date  ' Today's date
End Sub
```

### Message Boxes

**Display information to user:**

```vba
Sub ShowMessage()
    MsgBox "Report formatting complete!"
End Sub
```

**With variable:**
```vba
Sub ShowTotal()
    Dim total As Double
    total = 1000
    MsgBox "Total Sales: $" & total
End Sub
```

**Message box options:**
```vba
' Question with Yes/No buttons
result = MsgBox("Continue?", vbYesNo + vbQuestion)
If result = vbYes Then
    ' User clicked Yes
End If
```

### Input Boxes

**Ask user for input:**

```vba
Sub GetUserName()
    Dim userName As String
    userName = InputBox("Enter your name:")
    Range("A1").Value = userName
End Sub
```

**With default value:**
```vba
userName = InputBox("Enter name:", "Name Entry", "John Doe")
```

### If Statements

**Make decisions:**

```vba
Sub CheckValue()
    If Range("A1").Value > 1000 Then
        Range("B1").Value = "High"
    Else
        Range("B1").Value = "Low"
    End If
End Sub
```

**Multiple conditions:**
```vba
Sub CategorizeScore()
    Dim score As Integer
    score = Range("A1").Value
    
    If score >= 90 Then
        Range("B1").Value = "A"
    ElseIf score >= 80 Then
        Range("B1").Value = "B"
    ElseIf score >= 70 Then
        Range("B1").Value = "C"
    Else
        Range("B1").Value = "F"
    End If
End Sub
```

### Loops (Preview)

**Repeat actions:**

```vba
' Loop through rows 1 to 10
Sub LoopExample()
    Dim i As Integer
    For i = 1 To 10
        Cells(i, 1).Value = i  ' Put row number in column A
    Next i
End Sub
```

**Loop until condition met:**
```vba
Sub LoopUntilBlank()
    Dim row As Integer
    row = 1
    Do While Cells(row, 1).Value <> ""
        Cells(row, 2).Value = "Processed"
        row = row + 1
    Loop
End Sub
```

---

## When to Use Macros vs Alternatives

### Use Macros When:

```
✅ Formatting multiple sheets/ranges
✅ Generating standardized reports
✅ Creating custom buttons/tools
✅ Automating multi-step workflows
✅ Interacting with other Office apps
✅ Custom data entry forms
✅ Tasks that need human-like interaction
```

**Examples:**
```
- Apply company formatting template
- Generate monthly report layout
- Export data to multiple formats
- Create summary sheet from multiple sources
- Custom data validation workflows
```

### Don't Use Macros When:

```
❌ Formula can do it (formulas update automatically)
❌ Power Query better suited (data transformation)
❌ Built-in Excel feature available
❌ Calculation needs to be dynamic
❌ Simple one-time task
```

**Examples:**
```
- Calculate totals → Use SUM
- Look up values → Use XLOOKUP
- Import/clean CSV → Use Power Query
- Conditional formatting → Use built-in feature
- Filter data → Use AutoFilter
```

### Comparison Scenarios

**Scenario 1: Monthly Sales Report**

```
Task: Sum sales by region, format results

❌ Macro approach:
- Calculate sums with VBA
- Format results
- Must re-run when data changes

✅ Formula approach:
- Use SUMIF for sums
- Format once
- Updates automatically
```

**Scenario 2: Import Weekly Data Files**

```
Task: Import 50 CSV files, clean, combine

❌ Macro approach:
- Loop through files
- Import each
- Clean with code
- Complex to maintain

✅ Power Query approach:
- Get Data → From Folder
- Automatic file combination
- Repeatable transformations
- Easy to refresh
```

**Scenario 3: Format Quarterly Presentation**

```
Task: Apply formatting, add charts, logo, footers

✅ Macro approach:
- Multi-step formatting
- Insert objects
- Page setup
- Save as PDF
- Perfect for automation!

❌ Formula: Can't format or insert objects
❌ Power Query: Not for formatting
```

---

## Macro Best Practices

### Naming

```
✅ Use descriptive names:
FormatMonthlyReport
ExportToPDF
CalculateCommissions
ClearAllFilters

❌ Avoid vague names:
Macro1
Test
Update
DoStuff
```

### Documentation

**Add comments:**
```vba
'
' Macro: FormatSalesReport
' Purpose: Applies standard formatting to monthly sales report
' Author: John Smith
' Date: 2024-12-26
' Last Modified: 2024-12-26
'
Sub FormatSalesReport()
    ' Bold and color header row
    Range("A1:F1").Font.Bold = True
    Range("A1:F1").Interior.Color = RGB(68, 114, 196)
    
    ' Auto-fit all columns
    Columns("A:F").AutoFit
    
    ' Freeze header row
    Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub
```

### Error Handling

**Add basic error handling:**
```vba
Sub SafeMacro()
    On Error GoTo ErrorHandler
    
    ' Your code here
    Range("A1").Value = "Test"
    
    Exit Sub  ' Exit before error handler
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
```

### Testing

```
✅ Test on copy of workbook first
✅ Test with different data scenarios
✅ Include undo capability if possible
✅ Add confirmation messages
✅ Test error conditions
```

### Performance

```
✅ Turn off screen updating for speed:
Application.ScreenUpdating = False
' Your code
Application.ScreenUpdating = True

✅ Avoid selecting when possible:
❌ Range("A1").Select
   Selection.Value = "Test"

✅ Range("A1").Value = "Test"

✅ Use With statements:
With Range("A1:E1")
    .Font.Bold = True
    .Font.Size = 12
    .Interior.Color = Blue
End With
```

---

## Saving Workbooks with Macros

### File Format Requirements

**Macro-enabled format:**
```
.xlsm = Excel Macro-Enabled Workbook
Must use this format to save macros!

.xlsx = Cannot contain macros
If you save as .xlsx, macros deleted!
```

### Saving Process

**First time save:**
```
1. File → Save As
2. Choose location
3. Save as type: Excel Macro-Enabled Workbook (*.xlsm)
4. Enter filename
5. Save
```

**Warning if saving as .xlsx:**
```
┌──────────────────────────────────────────┐
│ ⚠️ Warning                               │
├──────────────────────────────────────────┤
│ The following features cannot be saved   │
│ in macro-free workbooks:                 │
│ • VB project                             │
│                                          │
│ To save a file with these features,     │
│ click No, and choose a macro-enabled    │
│ file type in the Save As Type list.     │
│                                          │
│ [Yes] [No] [Help]                        │
└──────────────────────────────────────────┘

Click No → Change to .xlsm
```

### Distribution Considerations

**Sharing macros with others:**
```
1. Save as .xlsm
2. Document what macros do
3. Include instructions
4. Warn about enabling content
5. Consider digital signature (advanced)
```

---

## Troubleshooting Common Issues

### Problem: Macro Not Appearing in List

**Possible causes:**
```
1. Macro saved in different workbook
2. Macro is Private (not Public)
3. Workbook not open
4. Name doesn't meet requirements
```

**Solutions:**
```
1. Click "Debug" to see which line failed
2. Check if sheet/range exists
3. Unprotect sheet if protected
4. Add error handling (On Error)
5. Verify references match actual workbook
```

### Problem: Macro Runs Too Slowly

**Causes:**
```
- Too many Select statements
- Screen updating on
- Calculation mode automatic
- Inefficient loops
```

**Solutions:**
```vba
Sub FastMacro()
    ' Turn off updates
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Your code here
    
    ' Turn back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### Problem: Button Disappeared

**Cause:** Button accidentally deleted or moved off screen

**Solution:**
```
1. Press Ctrl+G (Go To)
2. Special → Objects
3. Select All → OK
4. Arrow keys to find button
Or:
1. Developer → Insert → Form Controls
2. Insert new button
3. Assign same macro
```

---

## Quick Reference: Macro Tasks

| Task | How To |
|------|--------|
| **Enable Developer tab** | File → Options → Customize Ribbon → Check Developer |
| **Record macro** | Developer → Record Macro |
| **Stop recording** | Developer → Stop Recording |
| **Run macro** | Alt+F8 → Select → Run |
| **Edit macro** | Alt+F11 → Find in Modules |
| **Delete macro** | Alt+F8 → Select → Delete |
| **Assign to button** | Right-click shape → Assign Macro |
| **Change security** | Developer → Macro Security |
| **Save with macros** | Save as .xlsm (not .xlsx) |
| **View VBA Editor** | Alt+F11 |

---

## Keyboard Shortcuts: Macros & VBA

| Shortcut | Action |
|----------|--------|
| `Alt+F8` | Open Macros dialog |
| `Alt+F11` | Open/close VBA Editor |
| `F5` | Run macro (in VBA Editor) |
| `F8` | Step through code (debug) |
| `Ctrl+Break` | Stop running macro |
| `Ctrl+G` | Show Immediate Window (VBA) |
| `Ctrl+R` | Show Project Explorer (VBA) |
| `F2` | Object Browser (VBA) |

---

## Practical Examples

### Example 1: Simple Formatting Macro

**Task:** Format table headers consistently

```vba
Sub FormatTableHeaders()
    ' Select header row
    Range("A1:F1").Select
    
    ' Apply formatting
    With Selection
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)  ' White
        .Interior.Color = RGB(68, 114, 196)  ' Blue
        .HorizontalAlignment = xlCenter
    End With
    
    ' Auto-fit columns
    Columns("A:F").AutoFit
    
    ' Confirmation
    MsgBox "Headers formatted successfully!"
End Sub
```

### Example 2: Export to PDF

**Task:** Save active sheet as PDF

```vba
Sub ExportToPDF()
    Dim filePath As String
    Dim fileName As String
    
    ' Get current workbook path
    filePath = ThisWorkbook.Path & "\"
    
    ' Create filename with date
    fileName = "Report_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    
    ' Export active sheet
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath & fileName, _
        Quality:=xlQualityStandard, _
        OpenAfterPublish:=True
    
    MsgBox "PDF saved as: " & fileName
End Sub
```

### Example 3: Clear All Filters

**Task:** Remove all filters from active sheet

```vba
Sub ClearAllFilters()
    On Error Resume Next  ' In case no filters exist
    
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilter.ShowAllData
    End If
    
    MsgBox "All filters cleared!"
End Sub
```

### Example 4: Insert Timestamp

**Task:** Insert current date and time in selected cell

```vba
Sub InsertTimestamp()
    ActiveCell.Value = Now
    ActiveCell.NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
End Sub
```

### Example 5: Protect All Sheets

**Task:** Protect all sheets in workbook

```vba
Sub ProtectAllSheets()
    Dim ws As Worksheet
    Dim password As String
    
    ' Ask for password
    password = InputBox("Enter password to protect sheets:")
    
    If password = "" Then
        MsgBox "Protection cancelled."
        Exit Sub
    End If
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=password
    Next ws
    
    MsgBox "All sheets protected!"
End Sub
```

---


```
1. Check "Macros in" dropdown → Select correct workbook
2. In VBA: Change "Private Sub" to "Sub"
3. Open the workbook containing macro
4. Check macro name follows rules
```

### Problem: Security Warning Won't Go Away

**Cause:** Macros blocked by security settings

**Solution:**
```
1. Close workbook
2. Move to Trusted Location
   OR
3. Developer Tab → Macro Security
4. Change to "Disable all macros with notification"
5. Reopen workbook
6. Click "Enable Content"
```

### Problem: Compile Error

**Message:** "Compile error: Syntax error"

**Cause:** VBA code has typo or incorrect syntax

**Solution:**
```
1. VBA Editor opens, highlights error
2. Check spelling of commands
3. Check for missing quotes, parentheses
4. Compare to working examples
5. Fix and test
```

## Excel Online Limitations

⚠️ **Important:** Excel Online (excel.cloud.microsoft.com) has significant macro limitations:

```
❌ Cannot record macros
❌ Cannot create new macros
❌ Cannot edit VBA code
❌ Limited macro execution

✅ CAN run certain existing macros from desktop files
✅ Office Scripts available (alternative - JavaScript-based)
```

**If you need macros in Excel Online:**
```
1. Create macros in Excel Desktop
2. Save as .xlsm
3. Upload to OneDrive/SharePoint
4. Open in Excel Online
5. Some macros may run (simple ones)

OR:

Use Office Scripts instead:
- Automate Tab → Record Actions
- JavaScript-based (not VBA)
- Cloud-native, shareable
- Modern alternative to macros
```

**For this course:**
```
✅ Focus on Desktop Excel for macro learning
✅ Concepts transfer to Office Scripts
✅ VBA knowledge valuable for Desktop Excel
```

---

## Next Steps After Learning Basics

### Continue Learning VBA

**Topics to explore:**
```
1. Loops (For, Do While, For Each)
2. Arrays (store multiple values)
3. User forms (custom dialog boxes)
4. Error handling (robust error management)
5. Working with multiple workbooks
6. Interacting with other Office apps
7. Advanced functions and methods
8. Events (worksheet change, open, etc.)
9. Class modules (object-oriented)
10. Add-ins (distribute tools)
```

### Resources

**Built-in help:**
```
VBA Editor → Help → Microsoft Visual Basic Help
F1 key on any VBA command
Object Browser (F2) - explore objects
```

**Online resources:**
```
- Microsoft VBA documentation
- Excel VBA forums
- YouTube VBA tutorials
- Stack Overflow (programming Q&A)
- VBA blogs and websites
```

**Practice projects:**
```
1. Custom data entry form
2. Report generator
3. Data validation tool
4. Email automation (with Outlook)
5. Dashboard with refresh button
6. File processing automation
7. Custom ribbon tools
8. Workbook backup utility
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Macro = automated sequence of actions
- Recording captures actions as VBA code
- Developer tab required for macros
- .xlsm format required to save macros
- Macros can be security risk - only enable if trusted
- VBA = Visual Basic for Applications
- Alt+F11 opens VBA Editor
- Alt+F8 opens Macros dialog
- Relative vs Absolute recording affects flexibility
- Excel Online has limited macro support

### Practice Deeply
- Enabling Developer tab
- Recording simple macros (formatting, data entry)
- Starting and stopping recording
- Running macros via dialog (Alt+F8)
- Creating keyboard shortcuts for macros
- Using relative references when appropriate
- Creating shape buttons
- Assigning macros to buttons
- Testing buttons work correctly
- Understanding macro security warning
- Enabling content when safe
- Opening VBA Editor (Alt+F11)
- Navigating VBA Editor (Project Explorer, Code Window)
- Finding recorded macro code
- Reading basic VBA code structure
- Understanding Sub...End Sub
- Recognizing common VBA commands (Range, Select, Value)
- Adding simple comments to code
- Making minor edits to recorded macros
- Removing unnecessary Select statements
- Testing edited macros
- Saving as .xlsm format
- Deciding when to use macros vs formulas vs Power Query
- Creating practical macros (format headers, clear filters)
- Troubleshooting basic macro errors
- Using message boxes for confirmation
- Managing macro security settings

---

## Next Step

After this file, we move to:

**`22-data-cleaning-techniques.md`**
- Identifying common data issues
- Removing duplicates
- Finding and fixing errors
- Text cleaning (trim, proper case, etc.)
- Handling missing values
- Data validation for clean entry
- Split and merge techniques
- Standardizing formats
- Power Query for cleaning
- Flash Fill for pattern recognition:**


