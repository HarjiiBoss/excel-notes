# Data Validation

This file covers Excel's data validation features that control what users can enter
into cells. Data validation prevents errors, ensures consistency, and guides users
through dropdown lists and custom rules.

---

## What is Data Validation?

**Data validation** restricts what can be entered into a cell.

Think of it as **creating rules** that:
- Show dropdown lists for selection
- Limit numbers to specific ranges
- Ensure dates fall within periods
- Require text of certain length
- Display helpful messages
- Prevent invalid data entry

### Why Use Data Validation?

**Without validation:**
```
     A
  ┌──────────┐
1 │ Region   │
2 │ East     │
3 │ east     │ ← Inconsistent
4 │ Eastern  │ ← Typo
5 │ E        │ ← Abbreviation
  └──────────┘

Problems: Inconsistent, hard to filter/analyze
```

**With validation:**
```
     A
  ┌──────────┐
1 │ Region   │ ← Click shows: East, West, North, South
2 │ East     │ ← Selected from list
3 │ West     │ ← Selected from list
4 │ East     │ ← Consistent!
  └──────────┘

Benefits: Consistent, accurate, easy to use
```

---

## Accessing Data Validation

**Location:** Data tab → Data Validation button

**Steps:**
1. Select cell(s) where validation should apply
2. Click **Data** tab
3. Click **Data Validation** button
4. Configure settings

### Visual Interface
```
┌─────────────────────────────────────┐
│ Data Validation Dialog              │
├─────────────────────────────────────┤
│ Settings | Input Message | Error Alert│
├─────────────────────────────────────┤
│ Allow:  [List ▼]                    │
│ Data:   [between ▼]                 │
│ Source: [=A1:A5]                    │
│                                     │
│ ☑ Ignore blank                      │
│ ☑ Apply to range                    │
│                                     │
│         [OK]  [Cancel]              │
└─────────────────────────────────────┘
```

---

## Dropdown Lists (List Validation)

**Most common validation type**

### Method 1: Type List Directly

**Steps:**
1. Select cell(s)
2. Data → Data Validation
3. Allow: **List**
4. Source: Type items separated by commas
5. Click OK

**Example:**
```
Source: East,West,North,South

Result in cell:
     A
  ┌──────────┐
1 │ East  ▼  │ ← Dropdown arrow appears
  └──────────┘

Click arrow shows:
  ┌──────────┐
  │ East     │
  │ West     │
  │ North    │
  │ South    │
  └──────────┘
```

### Method 2: Reference a Range

**Setup:**
```
List in column F:
     F
  ┌──────────┐
1 │ East     │
2 │ West     │
3 │ North    │
4 │ South    │
  └──────────┘

Validation:
- Select cells where dropdown should appear
- Allow: List
- Source: =F1:F4 or =$F$1:$F$4
```

### Method 3: Named Range (Best Practice)

**Setup:**
```
Step 1: Create list
     F
  ┌──────────┐
1 │ East     │
2 │ West     │
3 │ North    │
4 │ South    │
  └──────────┘

Step 2: Name the range
- Select F1:F4
- Click in Name Box (left of formula bar)
- Type: Regions
- Press Enter

Step 3: Use in validation
- Source: =Regions
```

**Benefits of Named Ranges:**
- Easier to read: `=Regions` vs `=$F$1:$F$4`
- Updates automatically if list grows
- Can reference from other sheets
- Self-documenting

### Real-World Example 1: Status Dropdown
```
     A              B
  ┌────────────┬──────────┐
1 │ Task       │ Status   │
2 │ Design     │ [List ▼] │
3 │ Code       │ [List ▼] │
4 │ Test       │ [List ▼] │
  └────────────┴──────────┘

Source: Not Started,In Progress,Complete,On Hold

Users select from predefined statuses
```

### Real-World Example 2: Department List
```
     A              B
  ┌────────────┬──────────┐
1 │ Employee   │ Dept     │
2 │ Alice      │ [List ▼] │
3 │ Bob        │ [List ▼] │
  └────────────┴──────────┘

List in F1:F5:
HR
IT
Finance
Sales
Marketing

Source: =$F$1:$F$5

Ensures consistent department names
```

---

## Number Validation

**Purpose:** Restrict numeric entries to specific ranges

### Validation Types

| Type | Description | Example Use |
|------|-------------|-------------|
| **Whole number** | Integers only | Quantity, Age |
| **Decimal** | Any number | Price, Weight |
| **Between** | Range of values | 1 to 100 |
| **Greater than** | Minimum value | > 0 |
| **Less than** | Maximum value | < 1000 |

### Example 1: Age Validation
```
Allow: Whole number
Data: between
Minimum: 18
Maximum: 65

     A
  ┌──────┐
1 │ Age  │
2 │ 25   │ ✓ Accepted
3 │ 17   │ ✗ Error: Must be 18-65
4 │ 32.5 │ ✗ Error: Must be whole number
  └──────┘
```

### Example 2: Price Validation
```
Allow: Decimal
Data: greater than
Minimum: 0

     A
  ┌──────────┐
1 │ Price    │
2 │ 25.50    │ ✓ Accepted
3 │ -10.00   │ ✗ Error: Must be > 0
4 │ 0        │ ✗ Error: Must be > 0
  └──────────┘
```

### Example 3: Percentage Validation
```
Allow: Decimal
Data: between
Minimum: 0
Maximum: 1

     A
  ┌──────────┐
1 │ Discount │
2 │ 0.15     │ ✓ Accepted (15%)
3 │ 1.5      │ ✗ Error: Must be 0-1
4 │ -0.05    │ ✗ Error: Must be 0-1
  └──────────┘

Format as percentage to display 15%
```

### Using Cell References
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Min     │ Max     │ Quantity│
2 │ 1       │ 100     │ [Input] │
  └─────────┴─────────┴─────────┘

Validation for C2:
Allow: Whole number
Data: between
Minimum: =A2
Maximum: =B2

Dynamic! Changes if A2 or B2 change
```

---

## Date Validation

**Purpose:** Restrict dates to specific ranges or conditions

### Example 1: Future Dates Only
```
Allow: Date
Data: greater than
Start date: =TODAY()

     A
  ┌────────────┐
1 │ Deadline   │
2 │ 12/31/2025 │ ✓ Accepted (future)
3 │ 1/1/2020   │ ✗ Error: Must be after today
  └────────────┘
```

### Example 2: Date Range
```
Allow: Date
Data: between
Start date: 1/1/2024
End date: 12/31/2024

     A
  ┌────────────┐
1 │ Event Date │
2 │ 6/15/2024  │ ✓ Accepted
3 │ 1/15/2025  │ ✗ Error: Must be in 2024
  └────────────┘
```

### Example 3: Business Days Only (Weekday)
```
Allow: Custom
Formula: =WEEKDAY(A2,2)<=5

     A
  ┌────────────┐
1 │ Meeting    │
2 │ 12/23/2025 │ ✓ Accepted (Tuesday)
3 │ 12/27/2025 │ ✗ Error: Weekend not allowed
  └────────────┘

WEEKDAY(date,2) returns 1-5 for Mon-Fri, 6-7 for Sat-Sun
```

### Example 4: Within 30 Days
```
Allow: Date
Data: between
Start date: =TODAY()
End date: =TODAY()+30

     A
  ┌────────────┐
1 │ Due Date   │
2 │ [Input]    │ Must be within next 30 days
  └────────────┘
```

---

## Text Length Validation

**Purpose:** Control the length of text entries

### Example 1: Exact Length (ZIP Code)
```
Allow: Text length
Data: equal to
Length: 5

     A
  ┌──────────┐
1 │ ZIP Code │
2 │ 12345    │ ✓ Accepted
3 │ 1234     │ ✗ Error: Must be 5 characters
4 │ 123456   │ ✗ Error: Must be 5 characters
  └──────────┘
```

### Example 2: Maximum Length
```
Allow: Text length
Data: less than or equal to
Maximum: 50

     A
  ┌────────────────┐
1 │ Description    │
2 │ [Input]        │ Max 50 characters
  └────────────────┘
```

### Example 3: Minimum Length (Password)
```
Allow: Text length
Data: greater than or equal to
Minimum: 8

     A
  ┌──────────┐
1 │ Password │
2 │ pass     │ ✗ Error: Too short
3 │ password123 │ ✓ Accepted
  └──────────┘
```

---

## Custom Validation (Formula-Based)

**Purpose:** Create complex validation rules using formulas

**Key:** Formula must return TRUE (valid) or FALSE (invalid)

### Example 1: No Duplicates
```
Formula: =COUNTIF($A$2:$A$100,A2)=1

     A
  ┌──────────┐
1 │ ID       │
2 │ 001      │ ✓ Accepted (first occurrence)
3 │ 002      │ ✓ Accepted
4 │ 001      │ ✗ Error: Duplicate not allowed
  └──────────┘

COUNTIF counts occurrences
If count = 1, it's unique (TRUE)
If count > 1, it's duplicate (FALSE)
```

### Example 2: Email Format
```
Formula: =AND(ISNUMBER(FIND("@",A2)),ISNUMBER(FIND(".",A2)))

     A
  ┌────────────────────┐
1 │ Email              │
2 │ user@company.com   │ ✓ Accepted (has @ and .)
3 │ usercompany.com    │ ✗ Error: Missing @
4 │ user@company       │ ✗ Error: Missing .
  └────────────────────┘

Checks for both @ and .
```

### Example 3: Greater Than Another Cell
```
Formula: =B2>A2

     A          B
  ┌─────────┬─────────┐
1 │ Start   │ End     │
2 │ 100     │ 150     │ ✓ Accepted (150 > 100)
3 │ 200     │ 180     │ ✗ Error (180 not > 200)
  └─────────┴─────────┘
```

### Example 4: Uppercase Only
```
Formula: =EXACT(A2,UPPER(A2))

     A
  ┌──────────┐
1 │ Code     │
2 │ ABC123   │ ✓ Accepted (all uppercase)
3 │ Abc123   │ ✗ Error: Must be uppercase
  └──────────┘

EXACT is case-sensitive
Compares cell to its uppercase version
```

### Example 5: Sum Must Equal Target
```
Formula: =SUM($B$2:$B$10)=100

     A          B
  ┌─────────┬─────────┐
1 │ Item    │ %       │
2 │ A       │ 40      │
3 │ B       │ 35      │
4 │ C       │ 25      │ ✓ Total = 100
5 │ D       │ 5       │ ✗ Total = 105
  └─────────┴─────────┘

Ensures percentages sum to 100
```

### Example 6: Date Not in Past
```
Formula: =A2>=TODAY()

     A
  ┌────────────┐
1 │ Event Date │
2 │ 12/31/2025 │ ✓ Accepted (future)
3 │ 1/1/2020   │ ✗ Error: Date in past
  └────────────┘
```

---

## Input Messages

**Purpose:** Show helpful hints when cell is selected

### Setting Input Message

**Steps:**
1. Data Validation dialog
2. Click **Input Message** tab
3. Check "Show input message when cell is selected"
4. Enter Title and Message
5. Click OK

**Example:**
```
Title: Enter Quantity
Message: Please enter a number between 1 and 100

Visual result when cell selected:
     A
  ┌──────────┐
1 │ Qty  ▼   │
  └──────────┘
  ┌────────────────────────┐
  │ Enter Quantity         │
  ├────────────────────────┤
  │ Please enter a number  │
  │ between 1 and 100      │
  └────────────────────────┘
```

### Real-World Examples

**Example 1: Date Entry**
```
Title: Deadline
Message: Enter a date within the next 30 days
```

**Example 2: Code Format**
```
Title: Product Code
Message: Format: ABC-123 (3 letters, dash, 3 numbers)
```

**Example 3: Selection Guide**
```
Title: Select Region
Message: Choose from: East, West, North, South
```

---

## Error Alerts

**Purpose:** Show message when invalid data is entered

### Error Alert Styles

| Style | Icon | Behavior |
|-------|------|----------|
| **Stop** | ⛔ | Prevents invalid entry (default) |
| **Warning** | ⚠️ | Allows override with Yes/No |
| **Information** | ℹ️ | Allows override with OK/Cancel |

### Setting Error Alert

**Steps:**
1. Data Validation dialog
2. Click **Error Alert** tab
3. Check "Show error alert after invalid data is entered"
4. Choose Style
5. Enter Title and Error message
6. Click OK

### Example: Stop Error
```
Style: Stop
Title: Invalid Entry
Error message: Quantity must be between 1 and 100

When user enters 150:
  ┌────────────────────────┐
  │ ⛔ Invalid Entry       │
  ├────────────────────────┤
  │ Quantity must be       │
  │ between 1 and 100      │
  │                        │
  │        [Retry] [Cancel]│
  └────────────────────────┘

Cannot proceed until valid entry
```

### Example: Warning
```
Style: Warning
Title: Unusual Value
Error message: This value is outside normal range. Continue?

When user enters unusual value:
  ┌────────────────────────┐
  │ ⚠️ Unusual Value       │
  ├────────────────────────┤
  │ This value is outside  │
  │ normal range. Continue?│
  │                        │
  │    [Yes] [No] [Cancel] │
  └────────────────────────┘

User can override if needed
```

---

## Dependent Dropdowns (Cascading Lists)

**Purpose:** Second dropdown changes based on first selection

### Example: Country → City

**Setup:**
```
Country list in F1:F3:
     F
  ┌──────────┐
1 │ USA      │
2 │ Canada   │
3 │ Mexico   │
  └──────────┘

City lists:
     G          H          I
  ┌─────────┬─────────┬─────────┐
1 │ USA     │ Canada  │ Mexico  │
2 │ New York│ Toronto │ Tijuana │
3 │ LA      │ Montreal│ Cancun  │
4 │ Chicago │ Vancouver│ Leon    │
  └─────────┴─────────┴─────────┘

Name each column:
G2:G4 = USA
H2:H4 = Canada
I2:I4 = Mexico
```

**Validation:**
```
Cell A2 (Country):
- Allow: List
- Source: =F1:F3

Cell B2 (City):
- Allow: List
- Source: =INDIRECT(A2)

INDIRECT uses A2's value as range name
If A2="USA", source becomes =USA
```

**Result:**
```
     A          B
  ┌─────────┬─────────┐
1 │ Country │ City    │
2 │ USA  ▼  │ [List ▼]│
  └─────────┴─────────┘

Select "USA" in A2 → B2 shows: New York, LA, Chicago
Select "Canada" in A2 → B2 shows: Toronto, Montreal, Vancouver
```

### Real-World Example: Category → Product

**Setup:**
```
Categories: Electronics, Clothing, Food

Named ranges:
Electronics: Laptop, Phone, Tablet
Clothing: Shirt, Pants, Shoes
Food: Apple, Bread, Milk
```

**Implementation:**
```
     A              B
  ┌────────────┬────────────┐
1 │ Category   │ Product    │
2 │ [List ▼]   │ [List ▼]   │
  └────────────┴────────────┘

A2 validation: =Categories
B2 validation: =INDIRECT(A2)

User selects "Electronics" → Product shows tech items
User selects "Clothing" → Product shows clothing items
```

---

## Dynamic Lists with Tables

**Purpose:** Lists that automatically expand

### Method: Excel Tables

**Setup:**
```
Step 1: Create list and convert to Table
     F
  ┌──────────┐
1 │ Region   │
2 │ East     │
3 │ West     │
4 │ North    │
  └──────────┘

Select F1:F4 → Insert → Table → OK

Step 2: Name the table column
- Click in table
- Table Design → Properties → Name: RegionList

Step 3: Use in validation
- Source: =RegionList[Region]
```

**Benefits:**
```
Add new region to table:
     F
  ┌──────────┐
1 │ Region   │
2 │ East     │
3 │ West     │
4 │ North    │
5 │ South    │ ← New entry
  └──────────┘

Dropdown automatically includes "South"!
No need to update validation formula
```

---

## Finding Cells with Validation

**Steps:**
1. Press **F5** (or Ctrl+G) for Go To
2. Click **Special**
3. Select **Data validation**
4. Choose **All** or **Same**
5. Click OK

All cells with validation are selected

---

## Copying Validation

### Method 1: Copy/Paste
```
1. Select cell with validation
2. Copy (Ctrl+C)
3. Select destination cells
4. Paste (Ctrl+V)

Validation copies along with content
```

### Method 2: Format Painter
```
1. Select cell with validation
2. Click Format Painter (Home tab)
3. Select destination cells

Copies validation (and formatting)
```

### Method 3: Paste Special → Validation
```
1. Copy cell with validation
2. Select destination
3. Paste Special (Ctrl+Alt+V)
4. Choose "Validation"
5. OK

Copies ONLY validation (not content/format)
```

---

## Removing Validation

### Remove from Selected Cells
```
1. Select cells
2. Data → Data Validation
3. Click "Clear All"
4. OK
```

### Find and Remove All
```
1. Find cells (F5 → Special → Data Validation)
2. Data → Data Validation
3. Clear All
4. OK
```

---

## Circle Invalid Data

**Purpose:** Highlight existing data that violates validation rules

**Steps:**
1. Data tab → Data Validation dropdown
2. Click **Circle Invalid Data**

**Visual:**
```
     A
  ┌──────────┐
1 │ Quantity │
2 │ 50       │ ← Valid
3 │( 150 )   │ ← Invalid (circled in red)
4 │ 75       │ ← Valid
  └──────────┘

Validation: Must be ≤ 100
Cell 3 has 150 (violates rule)
```

**Clear Circles:**
Data Validation dropdown → Clear Validation Circles

---

## Common Validation Patterns

### Pattern 1: Required Field
```
Allow: Custom
Formula: =LEN(A2)>0

Prevents blank entries
```

### Pattern 2: Unique Values
```
Allow: Custom
Formula: =COUNTIF($A$2:$A$100,A2)=1

No duplicates allowed
```

### Pattern 3: Phone Number Format
```
Allow: Custom
Formula: =AND(LEN(A2)=10,ISNUMBER(VALUE(A2)))

Must be exactly 10 digits
```

### Pattern 4: Email Domain
```
Allow: Custom
Formula: =RIGHT(A2,12)="@company.com"

Must end with @company.com
```

### Pattern 5: Conditional Based on Another Cell
```
Allow: Custom
Formula: =IF(A2="Yes",B2<>"",TRUE)

     A          B
  ┌─────────┬─────────┐
1 │ Member? │ ID      │
2 │ Yes     │ [Required if Yes]
3 │ No      │ [Optional]
  └─────────┴─────────┘

If A2="Yes", B2 cannot be blank
If A2="No", B2 can be anything
```

---

## Real-World Application: Order Form

Let's build a validated order form.

### Form Setup
```
     A          B          C          D          E
  ┌─────────┬─────────┬─────────┬─────────┬─────────┐
1 │ Item    │ Qty     │ Price   │ Total   │ Status  │
2 │ [List ▼]│ [1-100] │ [>0]    │ =B2*C2  │ [List ▼]│
  └─────────┴─────────┴─────────┴─────────┴─────────┘
```

### Validation Rules

**Column A (Item) - Dropdown:**
```
Allow: List
Source: =ProductList

Input Message:
  Title: Select Product
  Message: Choose from available products

Error Alert:
  Style: Stop
  Title: Invalid Product
  Message: Please select a product from the list
```

**Column B (Quantity) - Number Range:**
```
Allow: Whole number
Data: between
Minimum: 1
Maximum: 100

Input Message:
  Title: Enter Quantity
  Message: Enter quantity between 1 and 100

Error Alert:
  Style: Stop
  Title: Invalid Quantity
  Message: Quantity must be between 1 and 100
```

**Column C (Price) - Positive Number:**
```
Allow: Decimal
Data: greater than
Minimum: 0

Input Message:
  Title: Enter Price
  Message: Enter price greater than $0.00

Error Alert:
  Style: Stop
  Title: Invalid Price
  Message: Price must be greater than zero
```

**Column E (Status) - Dropdown:**
```
Allow: List
Source: Pending,Approved,Shipped,Delivered

Input Message:
  Title: Order Status
  Message: Select current order status

Error Alert:
  Style: Warning
  Title: Status Change
  Message: Are you sure you want to change status?
```

---

## Common Mistakes and Best Practices

### Mistake 1: Forgetting Absolute References
```
❌ Wrong: Source: =F1:F5
Copy down: References move to F2:F6, F3:F7...

✅ Right: Source: =$F$1:$F$5
Copy down: References stay fixed
```

### Mistake 2: List Too Long
```
❌ Problem: Typing 50 items in Source box
Hard to maintain, prone to errors

✅ Solution: Use range reference
Source: =$F$1:$F$50
or
Source: =MyList (named range)
```

### Mistake 3: Spaces in List Items
```
❌ Problem:
Source: East, West, North, South
         ↑    ↑     ↑      ↑
      Spaces added after commas

User types "East" → Not valid (needs " East" with space)

✅ Solution: No spaces
Source: East,West,North,South
```

### Mistake 4: Not Testing Validation
```
Always test:
- Valid entries (should work)
- Invalid entries (should reject)
- Edge cases (boundaries)
- Blank entries (if allowed)
```

### Mistake 5: Unclear Error Messages
```
❌ Bad: "Invalid entry"
User doesn't know what's wrong

✅ Good: "Age must be between 18 and 65"
Clear guidance on what's expected
```

### Mistake 6: Validating Wrong Range
```
❌ Problem: Applied validation to entire column
Includes header row

✅ Solution: Select data range only
Validate A2:A100 (not A:A)
```

---

## Best Practices

### 1. Use Named Ranges
```
❌ Hard to read: =Sheet2!$F$1:$F$20
✅ Clear: =StatusList

Easier to maintain and understand
```

### 2. Provide Input Messages
```
Always tell users what's expected:
- Format required
- Valid range
- Selection options
```

### 3. Use Appropriate Error Styles
```
Stop: For critical data (IDs, required fields)
Warning: For unusual but possible values
Information: For helpful reminders
```

### 4. Keep Lists Updated
```
Use Excel Tables for dynamic lists
Lists automatically expand when rows added
```

### 5. Document Validation Logic
```
Add comments to cells with complex validation:
Right-click → Insert Comment
Explain the business rule
```

### 6. Test Edge Cases
```
Test:
- Minimum and maximum values
- Exactly at boundaries
- Just outside boundaries
- Blank entries
- Special characters
```

### 7. Use Conditional Formatting with Validation
```
Visually highlight:
- Valid entries (green)
- Invalid entries (red)
- Required fields (yellow background)

Combines visual cues with validation rules
```

---

## Validation with Conditional Formatting

**Enhance validation with visual feedback**

### Example: Highlight Invalid Entries
```
Select data range
Conditional Formatting → New Rule
Use formula: =ISNA(MATCH(A2,ValidationList,0))
Format: Red fill

Cells not in validation list turn red
```

### Example: Color-Code Status
```
     A          B
  ┌─────────┬─────────┐
1 │ Task    │ Status  │
2 │ Design  │ Complete│ ← Green
3 │ Code    │ Pending │ ← Yellow
4 │ Test    │ Delayed │ ← Red
  └─────────┴─────────┘

Conditional Formatting rules:
- Status = "Complete" → Green
- Status = "Pending" → Yellow  
- Status = "Delayed" → Red
```

---

## Troubleshooting Validation Issues

### Issue: Dropdown Not Showing
**Causes:**
- Validation not applied to cell
- "In-cell dropdown" unchecked
- Cell protection enabled

**Solution:**
```
Check validation settings:
Data → Data Validation
Ensure "In-cell dropdown" is checked
```

### Issue: Validation Not Working
**Causes:**
- Data entered before validation applied
- Copy/paste overwrites validation
- Formula returns FALSE

**Solution:**
```
Use Circle Invalid Data to find problems
Re-apply validation if needed
Check formula logic
```

### Issue: List Not Updating
**Causes:**
- Static range reference
- Named range not expanded
- Source data changed location

**Solution:**
```
Use Excel Tables for dynamic lists
Or use OFFSET with COUNTA:
=OFFSET($F$1,0,0,COUNTA($F:$F),1)
```

### Issue: Formula Validation Errors
**Causes:**
- Wrong cell references
- Formula syntax errors
- Returns text instead of TRUE/FALSE

**Solution:**
```
Test formula separately in cell
Ensure returns TRUE or FALSE
Check cell references are correct
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Data Validation is in Data tab
- Three tabs: Settings, Input Message, Error Alert
- List validation uses comma-separated values or ranges
- Custom validation uses formulas that return TRUE/FALSE
- INDIRECT enables dependent dropdowns
- Always use absolute references ($) in validation sources
- Three error styles: Stop, Warning, Information

### Practice Deeply
- Creating dropdown lists from ranges
- Setting up number and date validation
- Writing custom validation formulas
- Creating helpful input messages
- Configuring appropriate error alerts
- Building dependent dropdowns with INDIRECT
- Using named ranges for validation
- Creating dynamic lists with Tables
- Testing validation rules thoroughly
- Troubleshooting validation issues
- Combining validation with conditional formatting

### Don't Memorize
- Every possible validation formula pattern
- All formula validation examples
- Exact dialog box layouts
- Every error message variation

---

## Quick Reference: Validation Types

### List Validation
```
Source: Item1,Item2,Item3
or
Source: =$F$1:$F$10
or  
Source: =ListName
```

### Number Validation
```
Whole number: between 1 and 100
Decimal: greater than 0
Decimal: less than or equal to 1
```

### Date Validation
```
Date: greater than =TODAY()
Date: between 1/1/2024 and 12/31/2024
Date: less than =TODAY()+30
```

### Text Length
```
Text length: equal to 5
Text length: less than or equal to 50
Text length: greater than or equal to 8
```

### Custom Formula
```
No duplicates: =COUNTIF($A$2:$A$100,A2)=1
Email format: =AND(ISNUMBER(FIND("@",A2)),ISNUMBER(FIND(".",A2)))
Uppercase only: =EXACT(A2,UPPER(A2))
Not blank: =LEN(A2)>0
```

---

## Advanced Validation Example: Multi-Field Form

### Complete Employee Form

**Setup:**
```
     A              B              C              D
  ┌────────────┬────────────┬────────────┬────────────┐
1 │ Name       │ Dept       │ Start Date │ Salary     │
2 │ [Required] │ [List]     │ [Future]   │ [Range]    │
  └────────────┴────────────┴────────────┴────────────┘
```

**Validation Rules:**

**A2 (Name) - Required:**
```
Custom: =LEN(A2)>0
Input: "Enter employee full name"
Error: "Name is required"
```

**B2 (Department) - List:**
```
List: =Departments
Input: "Select department from list"
Error: "Please select a valid department"
```

**C2 (Start Date) - Future, Weekday:**
```
Custom: =AND(C2>TODAY(),WEEKDAY(C2,2)<=5)
Input: "Enter future start date (weekday only)"
Error: "Start date must be future weekday"
```

**D2 (Salary) - Range by Department:**
```
Custom: =AND(D2>=VLOOKUP(B2,SalaryRanges,2,0),D2<=VLOOKUP(B2,SalaryRanges,3,0))
Input: "Enter salary within department range"
Error: "Salary outside allowed range for department"
```

---

## Next Step

After mastering data validation, you're ready to explore:

**`11-conditional-formatting.md`**
- Highlighting cells based on values
- Color scales and data bars
- Icon sets for visual indicators
- Formula-based formatting rules
- Managing formatting rules
- Conditional formatting best practices
- Creating visual dashboards
