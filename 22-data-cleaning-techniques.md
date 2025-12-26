# Data Cleaning Techniques

This file covers techniques for identifying and fixing common data quality issues - essential skills for preparing data for analysis, reporting, and decision-making.

---

## What is Data Cleaning?

**Data Cleaning** = The process of detecting and correcting (or removing) corrupt, inaccurate, or inconsistent data.

### Why Clean Data?

**Dirty data causes problems:**
```
❌ Incorrect calculations (text instead of numbers)
❌ Failed lookups (extra spaces, spelling variations)
❌ Duplicate records (inflated counts, wrong totals)
❌ Misleading analysis (garbage in = garbage out)
❌ Time wasted troubleshooting formulas
❌ Lost credibility in reports
```

**Clean data enables:**
```
✅ Accurate calculations
✅ Reliable analysis
✅ Successful automation
✅ Confident decision-making
✅ Efficient workflows
✅ Professional results
```

### Visual Concept

```
┌─────────────────────────────────────────────────────┐
│                DATA CLEANING FLOW                   │
│                                                     │
│  Dirty Data          Clean              Analysis    │
│  ┌──────────┐      ┌──────────┐      ┌──────────┐ │
│  │ Spaces   │      │ Trimmed  │      │ Accurate │ │
│  │ Dupes    │─────>│ Unique   │─────>│ Insights │ │
│  │ Errors   │      │ Fixed    │      │ Results  │ │
│  │ Mixed    │      │ Standard │      │          │ │
│  └──────────┘      └──────────┘      └──────────┘ │
│                                                     │
│  Time spent cleaning = Time saved analyzing        │
└─────────────────────────────────────────────────────┘
```

---

## Common Data Issues

### Issue 1: Extra Spaces

**Problem:**
```
Cell appears correct but formulas don't work
```

**Example:**
```
Column A (Customer Name):
"John Smith"     ← Looks fine
" John Smith"    ← Leading space (invisible)
"John Smith "    ← Trailing space (invisible)
"John  Smith"    ← Double space in middle

VLOOKUP fails to find "John Smith" because:
"John Smith" ≠ " John Smith" (space makes them different!)
```

**Visual indicators:**
```
Select cell, look at formula bar:
┌──────────────────────────────┐
│ fx  " John Smith"            │← Space before quote
└──────────────────────────────┘

Or cell appears left-aligned when should be right-aligned
```

### Issue 2: Inconsistent Text

**Problem:**
```
Same item entered different ways
```

**Example:**
```
Product names:
iPhone 13
Iphone 13
IPHONE 13
iphone13
IPhone 13 Pro  ← Different product!

Each counts as different item in Pivot Tables!
```

**Impact:**
```
PivotTable shows:
iPhone 13    : 10 units
Iphone 13    : 5 units
IPHONE 13    : 8 units
Total shown  : 23 units (across 3 rows)
Actual total : 23 units (but should be 1 row!)
```

### Issue 3: Mixed Data Types

**Problem:**
```
Numbers stored as text
```

**Example:**
```
Column A (Sales):
1000      ← Number (right-aligned)
"1500"    ← Text (left-aligned, green triangle)
2000      ← Number
"N/A"     ← Text
=SUM(A1:A4) = 3000 (ignores text!)
```

**Indicators:**
```
✅ Number: Right-aligned, no green triangle
❌ Text: Left-aligned, green triangle in corner
❌ Mixed: Some left, some right in same column
```

**Green triangle warning:**
```
┌──────────┐
│ △ 1500   │← Triangle = "Number stored as text"
└──────────┘
```

### Issue 4: Duplicate Records

**Problem:**
```
Same record appears multiple times
```

**Example:**
```
Customer list:
CustomerID | Name        | Email
C001       | John Smith  | john@email.com
C002       | Sarah Jones | sarah@email.com
C001       | John Smith  | john@email.com  ← Duplicate!
C003       | Mike Brown  | mike@email.com
```

**Impact:**
```
Customer count = 4 (should be 3)
Email campaign sends twice to John
Reports show inflated numbers
```

### Issue 5: Blank Cells and Rows

**Problem:**
```
Empty cells scattered in data
```

**Example:**
```
Row 1:  Data
Row 2:  Data
Row 3:  [BLANK]  ← Empty row
Row 4:  Data
Row 5:  [BLANK]  ← Empty row
Row 6:  Data
```

**Impact:**
```
COUNTA includes blanks in range
Sorting leaves blanks at top
AutoFill stops at blank
Formulas get confused
```

### Issue 6: Date Format Issues

**Problem:**
```
Dates stored as text or wrong format
```

**Example:**
```
Column A (Date):
1/15/2024     ← Actual date
"01/15/2024"  ← Text (can't calculate)
"Jan 15 2024" ← Text
15-Jan-2024   ← Actual date (different format)
01152024      ← Number (not recognized as date)
```

**Impact:**
```
Can't sort chronologically
Can't calculate date differences
Formulas like MONTH() fail
Filtering doesn't work properly
```

### Issue 7: Special Characters

**Problem:**
```
Non-printable characters from external sources
```

**Example:**
```
Text from web or database:
"Product A [non-breaking space] [line break] Extra text"

Looks weird, breaks formulas, hard to match
```

**Common culprits:**
```
- Line breaks (Char 10)
- Carriage returns (Char 13)
- Non-breaking spaces (Char 160)
- Tab characters (Char 9)
```

### Issue 8: Leading Zeros Lost

**Problem:**
```
ZIP codes, product codes lose leading zeros
```

**Example:**
```
ZIP code entered: 02134
Excel shows: 2134  ← Lost the zero!

Product ID: 00456
Excel shows: 456   ← Lost zeros!
```

**Why:**
```
Excel treats as number, removes leading zeros
```

---

## Identifying Data Issues

### Visual Inspection

**Look for:**
```
✅ Green triangles in corners (text numbers)
✅ Alignment inconsistencies (mixed types)
✅ Unexpected blanks in data
✅ Visible spaces or odd characters
✅ Different capitalization patterns
```

### Use Data Tab Tools

**Data Tab → Data Tools:**

```
┌────────────────────────────────┐
│ Text to Columns                │← Split data
│ Remove Duplicates              │← Find dupes
│ Data Validation                │← Check rules
└────────────────────────────────┘
```

### Check with Formulas

**Count blanks:**
```
=COUNTBLANK(A1:A100)
Should be 0 in data columns
```

**Find duplicates:**
```
=COUNTIF($A$1:$A$100,A1)>1
TRUE = duplicate exists
```

**Check data type:**
```
=ISTEXT(A1)   ← TRUE if text
=ISNUMBER(A1) ← TRUE if number
```

**Find extra spaces:**
```
=LEN(A1)<>LEN(TRIM(A1))
TRUE = has extra spaces
```

---

## Cleaning Technique 1: Remove Extra Spaces

### Method 1: TRIM Function

**Removes all extra spaces except single spaces between words**

**Syntax:**
```
=TRIM(text)
```

**Example:**
```
Column A (Original)    | Column B (Formula)      | Result
" John Smith"          | =TRIM(A1)               | "John Smith"
"John  Smith "         | =TRIM(A2)               | "John Smith"
"  John   Smith  "     | =TRIM(A3)               | "John Smith"
```

**Steps:**
```
1. Insert helper column next to data
2. Enter formula: =TRIM(A1)
3. Copy down for all rows
4. Copy results → Paste Values over original
5. Delete helper column
```

**Visual:**
```
Before:
┌──────────────────┐
│ A                │
│ " John Smith"    │← Extra spaces
│ "Sarah Jones "   │
└──────────────────┘

Helper column:
┌──────────────────┬──────────────────┐
│ A                │ B                │
│ " John Smith"    │ =TRIM(A1)        │
│ "Sarah Jones "   │ =TRIM(A2)        │
└──────────────────┴──────────────────┘

After cleanup:
┌──────────────────┐
│ A                │
│ "John Smith"     │← Cleaned
│ "Sarah Jones"    │
└──────────────────┘
```

### Method 2: Find and Replace

**For specific cases (leading/trailing spaces):**

```
1. Select data range
2. Ctrl+H (Find and Replace)
3. Find what: [space]  ← Type a space
4. Replace with: [leave empty]
5. Click "Replace All"

⚠️ Warning: Removes ALL spaces (even between words!)
Better for data that shouldn't have ANY spaces
```

**Example use case:**
```
Product codes with accidental spaces:
"PROD 123" → "PROD123"  ✅ Good

But:
"John Smith" → "JohnSmith"  ❌ Bad
```

### Method 3: Power Query

**Best for large datasets:**

```
1. Select data → Data Tab → From Table/Range
2. Power Query Editor opens
3. Select column
4. Transform Tab → Format → Trim
5. Close & Load

All extra spaces removed!
Refreshable if source data changes
```

---

## Cleaning Technique 2: Standardize Text Case

### Method 1: Text Functions

**UPPER - Convert to UPPERCASE:**
```
=UPPER(A1)

"john smith" → "JOHN SMITH"
"John Smith" → "JOHN SMITH"
```

**LOWER - Convert to lowercase:**
```
=LOWER(A1)

"JOHN SMITH" → "john smith"
"John Smith" → "john smith"
```

**PROPER - Convert to Proper Case:**
```
=PROPER(A1)

"JOHN SMITH" → "John Smith"
"john smith" → "John Smith"
"jOHN sMITH" → "John Smith"
```

**Example workflow:**
```
Original:
CustomerName
JOHN SMITH
sarah jones
Mike BROWN

Formula in B1: =PROPER(A1)
Copy down

Result:
CustomerName
John Smith
Sarah Jones
Mike Brown

Copy B → Paste Values to A → Delete B
```

### Method 2: Flash Fill (Quick Pattern)

**For simple patterns:**

```
1. Type desired format in first cell manually
2. Start typing second cell
3. Excel suggests pattern (Ctrl+E)
4. Press Enter to accept

Example:
A1: "JOHN SMITH"
B1: "John Smith"  ← You type
B2: [Start typing "Sa..."]
Excel suggests: "Sarah Jones"  ← Auto-detected pattern
Press Ctrl+E to fill down
```

⚠️ **Note:** Flash Fill not always reliable, verify results

### Method 3: Power Query

```
1. Load data to Power Query
2. Select column
3. Transform Tab → Format → 
   - Uppercase
   - Lowercase
   - Capitalize Each Word
4. Close & Load
```

---

## Cleaning Technique 3: Remove Duplicates

### Method 1: Built-in Remove Duplicates

**Data Tab → Remove Duplicates:**

```
1. Select data range (including headers)
2. Data Tab → Remove Duplicates
3. Choose columns to check
4. Click OK
```

**Dialog:**
```
┌────────────────────────────────────┐
│ Remove Duplicates                  │
├────────────────────────────────────┤
│ Select columns to check:           │
│ ☑ Name                             │
│ ☑ Email                            │
│ ☐ Phone                            │
│ ☐ Address                          │
│                                    │
│ ☑ My data has headers              │
│                                    │
│ [OK] [Cancel]                      │
└────────────────────────────────────┘
```

**Result:**
```
Before (4 rows):
Name       | Email
John Smith | john@email.com
Sarah Jones| sarah@email.com
John Smith | john@email.com  ← Duplicate
Mike Brown | mike@email.com

After (3 rows):
Name       | Email
John Smith | john@email.com  ← Kept first occurrence
Sarah Jones| sarah@email.com
Mike Brown | mike@email.com
```

**Message shown:**
```
"2 duplicate values found and removed;
2 unique values remain."
```

⚠️ **Important:** This permanently deletes rows! Save backup first.

### Method 2: Advanced Filter

**Non-destructive method (keeps original data):**

```
1. Select data range
2. Data Tab → Advanced Filter
3. Choose: Copy to another location
4. Check: ☑ Unique records only
5. Copy to: (select new location)
6. OK
```

**Result:**
```
Original data unchanged
Unique records copied to new location
```

### Method 3: Conditional Formatting to Highlight

**Find duplicates without deleting:**

```
1. Select data range
2. Home Tab → Conditional Formatting
3. Highlight Cells Rules → Duplicate Values
4. Choose formatting (e.g., red fill)
5. OK

Duplicates highlighted visually
Manually review and delete as needed
```

### Method 4: Formula to Identify

**Helper column approach:**

```
In column B (next to data in A):
=COUNTIF($A$1:$A$100,A1)>1

Drag down

Result:
TRUE  = Duplicate
FALSE = Unique

Filter for TRUE, review, delete if needed
```

---

## Cleaning Technique 4: Convert Text to Numbers

### Problem: Numbers Stored as Text

**Indicators:**
```
✅ Green triangle in cell corner
✅ Left-aligned (numbers usually right-aligned)
✅ Error checking icon appears
✅ SUM ignores these cells
```

### Method 1: Error Checking Smart Tag

**Quickest for small ranges:**

```
1. Click cell with green triangle
2. Yellow diamond icon appears
3. Click dropdown arrow
4. Select "Convert to Number"
5. Repeat for other cells
```

**Visual:**
```
┌──────────────────────────────┐
│ △ 1500                       │
│    ⚠ Error checking options  │
│       Convert to Number      │← Click this
│       Help on this error     │
│       Ignore Error           │
│       Edit in Formula Bar    │
└──────────────────────────────┘
```

### Method 2: Multiply by 1

**Works for any size range:**

```
1. Enter 1 in empty cell
2. Copy that cell
3. Select text-number range
4. Paste Special → Multiply
5. OK

All text numbers converted to real numbers!
```

**Steps detail:**
```
┌────────────────────────────────┐
│ Paste Special                  │
├────────────────────────────────┤
│ Paste:                         │
│ ○ All                          │
│ ○ Formulas                     │
│                                │
│ Operation:                     │
│ ○ None                         │
│ ● Multiply  ← Select this      │
│ ○ Divide                       │
│                                │
│ [OK] [Cancel]                  │
└────────────────────────────────┘
```

### Method 3: VALUE Function

**Formula approach:**

```
=VALUE(A1)

"1500" → 1500
" 2000 " → 2000 (also trims spaces)
```

**Steps:**
```
1. Helper column: =VALUE(A1)
2. Copy down
3. Copy results → Paste Values over original
4. Delete helper column
```

### Method 4: Text to Columns

**Bulk conversion:**

```
1. Select range with text numbers
2. Data Tab → Text to Columns
3. Click "Next" → "Next" → "Finish"
   (Don't change any settings)
4. Text numbers converted automatically!
```

### Method 5: Power Query

```
1. Load to Power Query
2. Select column
3. Transform Tab → Data Type → Whole Number
   (or Decimal Number)
4. Close & Load
```

---

## Cleaning Technique 5: Fix Date Formats

### Problem: Dates Stored as Text

**Example:**
```
"01/15/2024"  ← Text (can't calculate with)
01/15/2024    ← Date (can calculate)
```

### Method 1: DATEVALUE Function

**Convert text to date:**

```
=DATEVALUE(A1)

"01/15/2024" → 45307 (date serial number)
"Jan 15 2024" → 45307
"1/15/24" → 45307

Then format as date: Ctrl+1 → Date format
```

### Method 2: Text to Columns

```
1. Select date column (text dates)
2. Data Tab → Text to Columns
3. Next → Next
4. Column data format: Date (MDY or DMY)
5. Finish

Text dates converted to real dates
```

### Method 3: Combined Text Functions

**For specific text patterns:**

```
If date is: "2024-01-15" (text)
=DATE(LEFT(A1,4), MID(A1,6,2), RIGHT(A1,2))

Breaks down:
LEFT(A1,4) = "2024" (year)
MID(A1,6,2) = "01" (month)
RIGHT(A1,2) = "15" (day)
Combines into date
```

### Method 4: Flash Fill

```
1. Type correct date format in B1
2. Start typing B2
3. Ctrl+E (Flash Fill)
4. Excel detects pattern, fills down
```

---

## Cleaning Technique 6: Remove Special Characters

### Common Special Characters

```
Line Break (Alt+Enter inside cell)
Carriage Return (from external data)
Non-breaking Space (from web)
Tab characters
Smart quotes vs straight quotes
```

### Method 1: CLEAN Function

**Removes non-printable characters:**

```
=CLEAN(A1)

Removes characters with ASCII values 0-31
(line breaks, tabs, etc.)
```

### Method 2: SUBSTITUTE Function

**Remove specific characters:**

```
Remove line breaks:
=SUBSTITUTE(A1, CHAR(10), " ")

CHAR(10) = Line Feed
Replace with space or nothing
```

**Remove multiple characters:**
```
=SUBSTITUTE(SUBSTITUTE(A1, CHAR(10), ""), CHAR(13), "")

First SUBSTITUTE removes Char 10
Second SUBSTITUTE removes Char 13 from result
```

### Method 3: Find and Replace

**For visible characters:**

```
1. Ctrl+H
2. Find what: [type or paste character]
3. Replace with: [leave blank or desired replacement]
4. Replace All
```

**For line breaks:**
```
1. Ctrl+H
2. Find what: Press Ctrl+J (enters line break)
3. Replace with: [space or nothing]
4. Replace All
```

### Method 4: Combine TRIM and CLEAN

**Best practice for external data:**

```
=TRIM(CLEAN(A1))

CLEAN removes non-printable chars
TRIM removes extra spaces
Gets most common issues!
```

---

## Cleaning Technique 7: Handle Missing Values

### Identifying Missing Data

```
Truly blank: ISBLANK(A1) = TRUE
Contains space: A1=" ", looks blank but isn't
Contains zero: A1=0
Contains "N/A": Text that means "not available"
Error values: #N/A, #VALUE!, etc.
```

### Strategy 1: Replace with Zero

**When zeros make sense (sales, counts):**

```
Formula approach:
=IF(ISBLANK(A1),0,A1)

Or Find and Replace:
1. Select range
2. Ctrl+H
3. Find what: [leave blank]
4. Replace with: 0
5. Options → Match entire cell contents
6. Replace All
```

### Strategy 2: Replace with Dash

**For text fields where blank = no data:**

```
=IF(A1="", "-", A1)

Empty cells show: -
Clarifies "no data" vs "forgot to enter"
```

### Strategy 3: Keep Blank but Handle in Formulas

**Use functions that ignore blanks:**

```
AVERAGEIF, COUNTIF, SUMIF automatically ignore blanks

=AVERAGEIF(A1:A100,">0")
Only averages non-zero values
```

### Strategy 4: Remove Rows with Blanks

**If incomplete rows not needed:**

```
1. Select data
2. Home Tab → Find & Select → Go To Special
3. Select: Blanks
4. OK
5. Home Tab → Delete → Delete Sheet Rows

All rows with any blank deleted
```

⚠️ **Warning:** Destructive! Save backup first.

### Strategy 5: Fill Down/Up

**When blank should inherit value from above:**

```
Example:
Category | Item
Fruit    | Apple
         | Banana  ← Should be Fruit
         | Cherry  ← Should be Fruit
Vegetable| Carrot

Solution:
1. Select Category column
2. Home Tab → Fill → Down
Or: Ctrl+D

Result:
Category | Item
Fruit    | Apple
Fruit    | Banana  ← Filled
Fruit    | Cherry  ← Filled
Vegetable| Carrot
```

---

## Cleaning Technique 8: Split Data

### Problem: Multiple Values in One Cell

**Example:**
```
Full Name: "John Smith"
Need: First Name and Last Name in separate columns
```

### Method 1: Text to Columns

**Best for delimited data:**

```
1. Select column to split
2. Data Tab → Text to Columns
3. Choose: Delimited
4. Next
5. Select delimiter: Space, Comma, Tab, Other
6. Next
7. Preview split, adjust if needed
8. Finish
```

**Example with Space delimiter:**
```
Before:
┌──────────────┐
│ Full Name    │
│ John Smith   │
│ Sarah Jones  │
└──────────────┘

After:
┌────────────┬────────────┐
│ First Name │ Last Name  │
│ John       │ Smith      │
│ Sarah      │ Jones      │
└────────────┴────────────┘
```

**Example with Comma delimiter:**
```
Before:
┌────────────────────┐
│ City, State        │
│ Boston, MA         │
│ New York, NY       │
└────────────────────┘

After:
┌──────────┬───────┐
│ City     │ State │
│ Boston   │ MA    │
│ New York │ NY    │
└──────────┴───────┘
```

### Method 2: Flash Fill

**For pattern-based splitting:**

```
Original in A:         Type in B:      Type in C:
John Smith            John            Smith
Sarah Jones           Sarah           [Ctrl+E fills "Jones"]
Mike Brown            [Ctrl+E fills "Mike"]  [Ctrl+E fills "Brown"]
```

### Method 3: Formulas

**LEFT, RIGHT, MID, FIND:**

**Split "First Last" format:**
```
First Name (B1):
=LEFT(A1, FIND(" ", A1)-1)

Last Name (C1):
=RIGHT(A1, LEN(A1)-FIND(" ", A1))
```

**Split "Last, First" format:**
```
Last Name (B1):
=LEFT(A1, FIND(",", A1)-1)

First Name (C1):
=TRIM(RIGHT(A1, LEN(A1)-FIND(",", A1)))
```

**Split email (username and domain):**
```
Email: "john@company.com"

Username (B1):
=LEFT(A1, FIND("@", A1)-1)
Result: "john"

Domain (C1):
=RIGHT(A1, LEN(A1)-FIND("@", A1))
Result: "company.com"
```

### Method 4: Power Query

```
1. Load to Power Query
2. Select column
3. Transform Tab → Split Column → By Delimiter
4. Choose delimiter (comma, space, etc.)
5. Close & Load
```

---

## Cleaning Technique 9: Merge Data

### Problem: Data Split Across Columns

**Example:**
```
First Name | Last Name
John       | Smith
Need: "John Smith" in one column
```

### Method 1: Concatenation with &

**Simple join:**
```
=A1&" "&B1

John + [space] + Smith = "John Smith"
```

**Multiple columns:**
```
=A1&" "&B1&" "&C1

First + [space] + Middle + [space] + Last
```

### Method 2: CONCAT Function

**Join range of cells:**
```
=CONCAT(A1:C1)

Joins all cells in range (no separator)
"JohnDSmith"

Not ideal for names, better for codes
```

### Method 3: TEXTJOIN Function

**Join with custom separator:**
```
=TEXTJOIN(" ", TRUE, A1:C1)

Arguments:
" " = Separator (space)
TRUE = Ignore empty cells
A1:C1 = Range to join

Result: "John D Smith"
If middle initial empty: "John Smith"
```

**Examples:**
```
Create full address:
=TEXTJOIN(", ", TRUE, A1:D1)
Street, City, State, ZIP
"123 Main St, Boston, MA, 02134"

Create product code:
=TEXTJOIN("-", TRUE, A1:C1)
Category, Subcategory, ID
"ELEC-PHONE-12345"
```

### Method 4: Flash Fill

```
1. Type desired format in first cell manually
2. Excel detects pattern with Ctrl+E
```

---

## Cleaning Technique 10: Preserve Leading Zeros

### Problem

**Excel removes leading zeros:**
```
Enter: 02134
Shows: 2134  ← Zero gone!
```

### Solution 1: Format as Text BEFORE Entering

```
1. Select cells
2. Ctrl+1 → Number tab
3. Category: Text
4. OK
5. Now enter: 02134
Displays: 02134  ← Preserved!
```

**Visual indicator:**
```
┌──────────┐
│ '02134   │← Apostrophe visible in formula bar
└──────────┘
```

### Solution 2: Apostrophe Prefix

```
Type: '02134
(Single quote before number)

Quote not visible in cell
Formula bar shows it
Excel treats as text
```

### Solution 3: Custom Number Format

```
1. Select cells
2. Ctrl+1 → Custom category
3. Type format code: 00000
   (Five zeros = always 5 digits)
4. OK

Enter: 2134
Displays: 02134  ← Leading zero added!

Enter: 123
Displays: 00123  ← Leading zeros added!
```

### Solution 4: TEXT Function

**Convert number to text with leading zeros:**

```
=TEXT(A1, "00000")

A1 = 2134
Result = "02134" (text)

Format code:
"00000" = 5 digits minimum
"000" = 3 digits minimum
"0000000000" = 10 digits (for phone numbers)
```

---

## Data Cleaning with Power Query

### Why Use Power Query for Cleaning?

```
✅ Non-destructive (original data unchanged)
✅ Repeatable (refresh applies all steps)
✅ Multiple transformations in one place
✅ Handle large datasets efficiently
✅ Visual interface (no formulas needed)
✅ Combines well with other Power Query features
```

### Common Cleaning Steps in Power Query

**1. Remove blank rows:**
```
Home Tab → Remove Rows → Remove Blank Rows
```

**2. Trim text:**
```
Select column → Transform Tab → Format → Trim
```

**3. Change case:**
```
Transform Tab → Format → 
- Uppercase
- Lowercase  
- Capitalize Each Word
```

**4. Remove duplicates:**
```
Home Tab → Remove Rows → Remove Duplicates
```

**5. Replace values:**
```
Transform Tab → Replace Values
Find: "N/A"
Replace: 0
```

**6. Split column:**
```
Transform Tab → Split Column → By Delimiter
```

**7. Change data types:**
```
Transform Tab → Data Type → Text/Number/Date
```

**8. Fill down:**
```
Transform Tab → Fill → Down
```

**9. Remove special characters:**
```
Transform Tab → Format → Clean
(Removes non-printable)
```

### Example Workflow

**Clean messy sales data:**
```
1. Get Data → From File → From CSV
2. Transform Data (opens Power Query)
3. Promote first row to headers
4. Change data types (numbers, dates)
5. Trim all text columns
6. Replace "N/A" with 0
7. Remove blank rows
8. Remove duplicate orders
9. Filter to current year only
10. Close & Load

All steps saved - click Refresh to reapply!
```

---

## Flash Fill for Pattern Recognition

### What is Flash Fill?

**Excel's AI-powered data cleaning tool:**
- Detects patterns you type
- Automatically fills remaining cells
- Works for splitting, combining, formatting

**Shortcut:** Ctrl+E

### When Flash Fill Works Well

```
✅ Consistent patterns
✅ Simple transformations
✅ Small to medium datasets
✅ One-time cleanup

❌ Complex logic
❌ Variable patterns
❌ Large datasets (use formulas/Power Query)
❌ Need to repeat (formulas better)
```

### Example 1: Extract First Name

```
Column A (Full Name) | Column B (First Name)
John Smith           | John  ← You type
Sarah Jones          | [Start typing "Sa..."]
                     | Flash Fill suggests "Sarah"
Press Ctrl+E         | All names extracted
```

### Example 2: Format Phone Numbers

```
Column A (Raw)       | Column B (Formatted)
2025551234           | (202) 555-1234  ← You type
6175559876           | [Ctrl+E]
                     | (617) 555-9876  ← Auto-filled
```

### Example 3: Combine and Format

```
First | Last  | Email → Full Name & Email
John  | Smith | john@ex.com → John Smith (john@ex.com)
Type first result, Ctrl+E fills rest
```

### Flash Fill Limitations

```
❌ Can't always detect complex patterns
❌ No formula to edit later
❌ Doesn't update if source changes
❌ Hard to troubleshoot errors

For important/repeated tasks: Use formulas instead
```

---

## Data Validation for Prevention

### Prevent Dirty Data at Entry

**Better to prevent than clean later:**

```
Data Validation = Rules that control what can be entered
```

### Set Up Validation

**Data Tab → Data Validation:**

```
1. Select cells where data will be entered
2. Data Tab → Data Validation
3. Choose validation criteria
4. Set input message (optional)
5. Set error alert (optional)
6. OK
```

**Validation Dialog:**
```
┌────────────────────────────────────┐
│ Data Validation                    │
├────────────────────────────────────┤
│ Settings | Input Message | Error   │
│                                    │
│ Allow: [Whole number    ▼]        │
│ Data:  [between         ▼]        │
│ Minimum: [1]                       │
│ Maximum: [100]                     │
│                                    │
│ ☑ Ignore blank                     │
│                                    │
│ [OK] [Cancel]                      │
└────────────────────────────────────┘
```

### Common Validation Rules

**1. Dropdown list (most common):**
```
Allow: List
Source: "Yes,No,Maybe"
or
Source: =$A$1:$A$10 (reference list)

User sees dropdown arrow
Can only select from list
```

**2. Whole numbers in range:**
```
Allow: Whole number
Data: between
Minimum: 0
Maximum: 100

Only accepts 0-100
```

**3. Date restrictions:**
```
Allow: Date
Data: between
Start date: 1/1/2024
End date: 12/31/2024

Only dates in 2024
```

**4. Text length:**
```
Allow: Text length
Data: less than or equal to
Maximum: 50

Limits character count
```

**5. Custom formula:**
```
Allow: Custom
Formula: =LEN(TRIM(A1))>0

Requires non-blank after trimming
```

### Error Messages

**Input Message (helpful hint):**
```
Title: "Enter Age"
Message: "Please enter age between 0 and 120"

Shows when cell selected
```

**Error Alert (validation failed):**
```
Style: Stop (❌), Warning (⚠️), Information (ℹ️)
Title: "Invalid Entry"
Message: "Age must be between 0 and 120"

Shows if invalid data entered
```

---

## Complete Cleaning Workflow Example

### Scenario: Clean Customer Database

**Raw data issues:**
```
- Extra spaces in names
- Inconsistent capitalization
- Phone numbers in different formats
- Duplicate records
- Missing email addresses
- Dates stored as text
```

### Step-by-Step Cleaning

**Step 1: Create backup copy**
```
1. Save original file
2. Work on copy
```

**Step 2: Remove duplicates**
```
1. Select all data
2. Data Tab → Remove Duplicates
3. Check: Email (unique identifier)
4. OK
Result: 245 duplicates removed
```

**Step 3: Clean names**
```
1. Insert helper column after Name
2. Formula: =PROPER(TRIM(A2))
3. Copy down all rows
4. Copy results → Paste Values over original
5. Delete helper column
Result: Names standardized, spaces removed
```

**Step 4: Fix phone numbers**
```
Option A - Flash Fill:
1. Type first formatted number: (617) 555-1234
2. Ctrl+E to fill pattern

Option B - Formula:
="("&LEFT(A2,3)&") "&MID(A2,4,3)&"-"&RIGHT(A2,4)
```

**Step 5: Convert dates**
```
1. Select date column
2. Data Tab → Text to Columns
3. Next → Next → Date format: MDY → Finish
Result: Text dates converted to real dates
```

**Step 6: Handle missing emails**
```
1. Filter Email column for blanks
2. Review: Delete rows or fill with "N/A"
3. Clear filter
```

**Step 7: Verify with checks**
```
Count blanks: =COUNTBLANK(range) → Should be 0
Check duplicates: =COUNTIF($A:$A,A2)>1 → All FALSE
Verify types: All consistent in each column
```

**Step 8: Set validation for future entries**
```
1. Name column: Text length ≤ 50 chars
2. Phone column: Text length = 10 or 12
3. Email column: Custom formula for @ symbol
4. Date column: Date between valid range
```

**Before vs After:**
```
BEFORE:
Name             Phone        Email           Date
" JOHN  SMITH"   6175551234   john@email.com  "01/15/2024"
"sarah Jones "   617-555-9876                 1/20/24
" JOHN  SMITH"   6175551234   john@email.com  "01/15/2024"

AFTER:
Name         Phone          Email           Date
John Smith   (617) 555-1234 john@email.com  1/15/2024
Sarah Jones  (617) 555-9876 N/A             1/20/2024
(Duplicate removed)
```

---

## Quick Reference: Cleaning Functions

| Problem | Formula | Example |
|---------|---------|---------|
| Extra spaces | `=TRIM(A1)` | " John " → "John" |
| Uppercase | `=UPPER(A1)` | "john" → "JOHN" |
| Lowercase | `=LOWER(A1)` | "JOHN" → "john" |
| Proper case | `=PROPER(A1)` | "john smith" → "John Smith" |
| Text to number | `=VALUE(A1)` | "1500" → 1500 |
| Remove chars | `=SUBSTITUTE(A1,"x","")` | "axa" → "aa" |
| Clean special | `=CLEAN(A1)` | Removes non-printable |
| Trim + Clean | `=TRIM(CLEAN(A1))` | Combined cleaning |
| Text to date | `=DATEVALUE(A1)` | "1/15/2024" → date |
| Leading zeros | `=TEXT(A1,"00000")` | 2134 → "02134" |
| Split first word | `=LEFT(A1,FIND(" ",A1)-1)` | "John Smith" → "John" |
| Join with space | `=A1&" "&B1` | Join columns |
| Join smart | `=TEXTJOIN(" ",TRUE,A1:C1)` | Ignores blanks |

---

## Quick Reference: Cleaning Methods

| Method | Best For | Speed | Skill Level |
|--------|----------|-------|-------------|
| **Find & Replace** | Simple substitutions | ⚡⚡⚡ | Beginner |
| **Formulas** | Flexible, repeatable | ⚡⚡ | Intermediate |
| **Flash Fill** | Pattern recognition | ⚡⚡⚡ | Beginner |
| **Text to Columns** | Split delimited data | ⚡⚡ | Beginner |
| **Remove Duplicates** | Delete duplicate rows | ⚡⚡⚡ | Beginner |
| **Power Query** | Large, complex datasets | ⚡ | Intermediate |
| **Data Validation** | Prevention at entry | ⚡⚡ | Beginner |
| **VBA Macro** | Custom automation | ⚡⚡ | Advanced |

---

## Data Cleaning Checklist

### Before Starting

```
☐ Save backup copy of original data
☐ Document issues found
☐ Plan cleaning steps
☐ Test on sample rows first
☐ Verify cleaning doesn't break references
```

### During Cleaning

```
☐ Work in helper columns (don't overwrite immediately)
☐ Check results after each step
☐ Use COUNTBLANK to verify no unintended blanks
☐ Use COUNTIF to check for duplicates
☐ Verify data types (numbers right-aligned, etc.)
☐ Check min/max values make sense
☐ Look for outliers or anomalies
```

### After Cleaning

```
☐ Final visual inspection
☐ Test formulas that depend on data
☐ Remove helper columns
☐ Document what was cleaned
☐ Set up data validation for future entries
☐ Save cleaned version separately
☐ Create refresh process if data updated regularly
```

---

## Common Mistakes to Avoid

### Mistake 1: Not Saving Backup

```
❌ Clean original file directly
✅ Always work on copy
✅ Keep original for reference
```

### Mistake 2: Overwriting Formulas

```
❌ Paste cleaned data over formulas
✅ Check for formulas first
✅ Paste in correct location
```

### Mistake 3: Ignoring Data Types

```
❌ Leave numbers as text
✅ Convert to proper type
✅ Verify calculations work
```

### Mistake 4: Deleting Without Checking

```
❌ Remove duplicates without review
✅ Check if "duplicates" are actually different
✅ Verify which record to keep
```

### Mistake 5: Not Documenting Changes

```
❌ Clean without notes
✅ Document what changed
✅ Record steps for repeatability
✅ Note any manual decisions made
```

### Mistake 6: Using Flash Fill for Critical Data

```
❌ Rely on Flash Fill for important patterns
✅ Use formulas for repeatability
✅ Verify Flash Fill results carefully
```

### Mistake 7: Cleaning Too Much

```
❌ Delete rows with any blank
✅ Understand if blanks are valid
✅ Check with data owner before deleting
```

---

## Performance Tips for Large Datasets

### Speed Up Cleaning

**Turn off calculations:**
```
Formulas Tab → Calculation Options → Manual

Do all cleaning, then:
Calculate Now (F9)

Turn back to Automatic when done
```

**Disable screen updating (VBA):**
```vba
Sub FastClean()
    Application.ScreenUpdating = False
    ' Your cleaning code
    Application.ScreenUpdating = True
End Sub
```

**Use Power Query:**
```
Better than formulas for 100K+ rows
Loads only preview in editor
Final load faster than formula calculations
```

**Work with filtered data:**
```
Filter to smaller subset
Clean that section
Repeat for other sections
Faster than processing all at once
```

**Use Tables:**
```
Convert range to Table (Ctrl+T)
Formulas auto-copy when you add rows
Structured references easier to manage
```

---

## When to Use Each Cleaning Tool

### Use Find & Replace When:
```
✅ Simple text substitution
✅ Known values to replace
✅ Quick one-time fix
✅ Removing specific characters
```

### Use Formulas When:
```
✅ Need to repeat process
✅ Complex logic required
✅ Multiple conditions
✅ Results need to update automatically
```

### Use Flash Fill When:
```
✅ Pattern is obvious to Excel
✅ Small dataset (< 1000 rows)
✅ One-time cleanup
✅ Quick demo or exploration
```

### Use Power Query When:
```
✅ Large datasets (10K+ rows)
✅ Multiple cleaning steps needed
✅ Data refreshes regularly
✅ Combining cleaning with import
✅ Working with multiple sources
```

### Use Text to Columns When:
```
✅ Data delimited by consistent character
✅ Need to split into fixed columns
✅ Simple split operation
```

### Use VBA When:
```
✅ Very complex custom logic
✅ Interacting with other applications
✅ Need button-triggered automation
✅ Multiple workbooks/sheets involved
```

---

## Real-World Cleaning Scenarios

### Scenario 1: Web Scraping Data

**Issues:**
```
- HTML tags in text
- Non-breaking spaces
- Irregular spacing
- Mixed encodings
```

**Solution:**
```
1. Power Query → Clean function
2. Find & Replace HTML tags
3. TRIM for spacing
4. Manual review of odd characters
```

### Scenario 2: Legacy System Export

**Issues:**
```
- Fixed-width columns (no delimiters)
- Leading zeros lost
- Dates as text in custom format
- Special codes instead of names
```

**Solution:**
```
1. Text to Columns (Fixed Width)
2. TEXT function for leading zeros
3. Custom date parsing formula
4. VLOOKUP to replace codes with names
```

### Scenario 3: User-Entered Forms

**Issues:**
```
- Inconsistent capitalization
- Extra spaces everywhere
- Optional fields blank
- Misspellings
```

**Solution:**
```
1. PROPER + TRIM on text fields
2. Replace blanks with "N/A" or appropriate default
3. Spell check or fuzzy matching for names
4. Data validation for future entries
```

### Scenario 4: Multiple Source Files

**Issues:**
```
- Different column orders
- Different column names
- Some columns missing in some files
- Different date formats
```

**Solution:**
```
Power Query:
1. Get Data → From Folder
2. Combine Files
3. Standardize column names
4. Handle missing columns (add as null)
5. Unified date format transformation
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Clean data = reliable analysis
- TRIM removes extra spaces
- PROPER/UPPER/LOWER change text case
- VALUE converts text to number
- Remove Duplicates deletes duplicate rows
- Text to Columns splits delimited data
- Flash Fill detects patterns (Ctrl+E)
- Green triangle = number stored as text
- Leading zeros need text format or apostrophe
- Data Validation prevents dirty data at entry
- Always save backup before cleaning
- Power Query best for large/repeated cleaning
- CLEAN removes non-printable characters
- SUBSTITUTE replaces specific text

### Practice Deeply
- Identifying data quality issues visually
- Using TRIM to remove extra spaces
- Using PROPER/UPPER/LOWER for text standardization
- Converting text numbers to real numbers (multiply by 1)
- Using VALUE function for text-to-number conversion
- Removing duplicate records safely
- Using Text to Columns to split data
- Using Find & Replace for simple substitutions
- Creating formulas with LEFT, RIGHT, MID, FIND
- Using TEXTJOIN to combine columns
- Using Flash Fill for pattern-based cleaning (Ctrl+E)
- Preserving leading zeros with TEXT function
- Using SUBSTITUTE to remove specific characters
- Combining TRIM and CLEAN for external data
- Handling blank cells appropriately
- Converting text dates with DATEVALUE
- Using Remove Blanks and Fill Down
- Setting up Data Validation rules
- Building complete cleaning workflows
- Testing cleaned data before finalizing
- Using helper columns for formula-based cleaning
- Copy → Paste Values to replace formulas with results
- Cleaning in Power Query (trim, case, types)
- Documenting cleaning steps taken
- Creating validation to prevent future issues

---

## Next Step

After this file, we move to:

**`23-what-if-analysis.md`**
- Goal Seek (find input for desired output)
- Data Tables (one-variable and two-variable)
- Scenario Manager (save and compare scenarios)
- Solver (optimization with constraints)
- Sensitivity analysis
- Break-even analysis
- Best/worst case modeling
- Financial modeling applications
