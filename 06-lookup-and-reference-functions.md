# Lookup and Reference Functions

This file covers Excel's lookup functions that find and retrieve data from tables
and ranges. These functions are essential for creating dynamic spreadsheets that
automatically pull information based on search criteria.

---

## What are Lookup Functions?

**Lookup functions** search for a value in a table and return related information.

Think of it like **looking up a word in a dictionary**:
- You know the word (lookup value)
- You search the dictionary (table)
- You find the definition (return value)

### Real-World Analogy
```
Employee ID → Lookup → Employee Database → Return Name

    1001    →  Search  →  ┌──────┬─────────┐
                           │ ID   │ Name    │
                           ├──────┼─────────┤
                           │ 1001 │ Alice   │
                           │ 1002 │ Bob     │
                           └──────┴─────────┘
                                      ↓
                                   "Alice"
```

### Why Use Lookup Functions?

**Instead of manually copying data:**
```
❌ Manual (breaks when source changes):
Cell B2: Type "Alice" manually
```

**Use lookups (stays synchronized):**
```
✅ Dynamic:
Cell B2: =VLOOKUP(A2, EmployeeTable, 2, FALSE)
Updates automatically when data changes!
```

---

## VLOOKUP Function

**Purpose:** Searches **vertically** down the first column of a table and returns a value from the same row

**Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

### The Four Arguments

| Argument | What It Is | Example |
|----------|------------|---------|
| **lookup_value** | What to search for | `A2` or `"John"` or `1001` |
| **table_array** | The data table to search | `B2:E10` |
| **col_index_num** | Which column to return (1=first column) | `2` or `3` |
| **range_lookup** | Exact match (FALSE) or approximate (TRUE) | `FALSE` |

### Visual Structure
```
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
         ↓             ↓              ↓             ↓
         What to       Where to       Which         Exact
         find          search         column        match
```

### How VLOOKUP Works
```
Lookup: Find "1002" in first column, return from column 3

     B          C          D
  ┌──────┬──────────┬──────────┐
1 │ ID   │ Name     │ Dept     │  ← Headers (not included in range)
  ├──────┼──────────┼──────────┤
2 │ 1001 │ Alice    │ Sales    │
3 │ 1002 │ Bob      │ IT       │  ← Found! Return column 3
4 │ 1003 │ Carol    │ HR       │
  └──────┴──────────┴──────────┘
   ↑      Col 1      Col 2      Col 3
   Searches here                Returns from here

=VLOOKUP(1002, B2:D4, 3, FALSE)  →  "IT"
```

### Basic Example
```
Data Table (B2:D5):
     B          C          D
  ┌──────┬──────────┬──────────┐
1 │ ID   │ Name     │ Salary   │
2 │ 101  │ Alice    │ 50000    │
3 │ 102  │ Bob      │ 60000    │
4 │ 103  │ Carol    │ 55000    │
5 │ 104  │ David    │ 65000    │
  └──────┴──────────┴──────────┘

Formula:
=VLOOKUP(102, B2:D5, 2, FALSE)  →  "Bob"
=VLOOKUP(102, B2:D5, 3, FALSE)  →  60000
```

### Example 1: Product Lookup
```
Product Table (E2:G6):
     E          F          G
  ┌──────────┬─────────┬─────────┐
1 │ Product  │ Price   │ Stock   │
2 │ Widget   │ 25.00   │ 150     │
3 │ Gadget   │ 40.00   │ 200     │
4 │ Tool     │ 15.00   │ 75      │
5 │ Device   │ 60.00   │ 120     │
  └──────────┴─────────┴─────────┘

Order Form:
     A          B          C
  ┌──────────┬─────────┬─────────────────────────┐
1 │ Item     │ Price   │ Stock                   │
2 │ Gadget   │ =VLOOKUP(A2,$E$2:$G$6,2,FALSE) │ =VLOOKUP(A2,$E$2:$G$6,3,FALSE)
3 │ Widget   │ =VLOOKUP(A3,$E$2:$G$6,2,FALSE) │ =VLOOKUP(A3,$E$2:$G$6,3,FALSE)
  └──────────┴─────────┴─────────────────────────┘
              ↓         ↓
            40.00      200
            25.00      150
```

### Example 2: Grade Lookup
```
Grade Table (E2:F6):
     E          F
  ┌─────────┬──────────┐
1 │ Score   │ Grade    │
2 │ 90      │ A        │
3 │ 80      │ B        │
4 │ 70      │ C        │
5 │ 60      │ D        │
  └─────────┴──────────┘

Student Scores:
     A          B          C
  ┌─────────┬─────────┬─────────────────────────┐
1 │ Student │ Score   │ Grade                   │
2 │ Alice   │ 92      │ =VLOOKUP(B2,$E$2:$F$6,2,TRUE)
3 │ Bob     │ 75      │ =VLOOKUP(B3,$E$2:$F$6,2,TRUE)
4 │ Carol   │ 88      │ =VLOOKUP(B4,$E$2:$F$6,2,TRUE)
  └─────────┴─────────┴─────────────────────────┘
                          ↓      ↓      ↓
                         "A"   "C"   "B"

Note: Using TRUE for approximate match (explained below)
```

### FALSE vs TRUE (Exact vs Approximate Match)

**FALSE (Exact Match):**
- Searches for **exact** value
- Returns error if not found
- Table can be in any order
- **Most common usage**

**TRUE (Approximate Match):**
- Finds the **largest value less than or equal to** lookup value
- Table MUST be sorted in ascending order (smallest to largest)
- Used for ranges (like grade scales, tax brackets)

### Approximate Match Example
```
Tax Bracket Table (MUST be sorted):
     A          B
  ┌─────────┬──────────┐
1 │ Income  │ Rate     │
2 │ 0       │ 10%      │
3 │ 10000   │ 12%      │
4 │ 40000   │ 22%      │
5 │ 85000   │ 24%      │
  └─────────┴──────────┘

=VLOOKUP(35000, A2:B5, 2, TRUE)  →  "12%"

How it works:
- 35000 is not in the table
- Finds largest value ≤ 35000
- That's 10000
- Returns corresponding rate: 12%
```

### ⚠️ Common VLOOKUP Errors

**Error: #N/A**
```
Cause: Lookup value not found

     B          C
  ┌──────┬──────────┐
2 │ 101  │ Alice    │
3 │ 102  │ Bob      │
  └──────┴──────────┘

=VLOOKUP(999, B2:C3, 2, FALSE)  →  #N/A

Fix: Use IFERROR or check spelling/data
=IFERROR(VLOOKUP(999,B2:C3,2,FALSE),"Not Found")
```

**Error: #REF!**
```
Cause: col_index_num is larger than columns in table

=VLOOKUP(101, B2:C3, 5, FALSE)  →  #REF!
                      ↑
                   Only 2 columns, asking for 5th

Fix: Use valid column number (1 or 2 in this case)
```

**Error: Wrong result with TRUE**
```
Cause: Table not sorted when using TRUE

     A          B
  ┌─────────┬──────────┐
2 │ 85000   │ 24%      │  ← Unsorted!
3 │ 10000   │ 12%      │
4 │ 0       │ 10%      │
  └─────────┴──────────┘

=VLOOKUP(35000, A2:B4, 2, TRUE)  →  Wrong result!

Fix: Sort A2:A4 in ascending order OR use FALSE
```

---

## VLOOKUP Best Practices

### 1. Always Use Absolute References for Tables
```
❌ Wrong: =VLOOKUP(A2, B2:D10, 2, FALSE)
          When copied down, B2:D10 becomes B3:D11, B4:D12...

✅ Right: =VLOOKUP(A2, $B$2:$D$10, 2, FALSE)
          Table reference stays fixed
```

### 2. Use FALSE for Most Lookups
```
✅ Most common: =VLOOKUP(A2, Table, 2, FALSE)
⚠️ Only use TRUE when doing range lookups with sorted data
```

### 3. Combine with IFERROR
```
=IFERROR(VLOOKUP(A2, Table, 2, FALSE), "Not Found")

Handles #N/A gracefully
```

### 4. Name Your Ranges
```
❌ Hard to read:
=VLOOKUP(A2, $E$2:$G$100, 3, FALSE)

✅ Clear:
=VLOOKUP(A2, ProductTable, 3, FALSE)
```

---

## HLOOKUP Function

**Purpose:** Searches **horizontally** across the first row of a table and returns a value from the same column

**Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

### When to Use HLOOKUP
Use when your data is organized **horizontally** (across rows instead of down columns).

### Visual Structure
```
Horizontal Table:
     A          B          C          D
  ┌──────────┬──────────┬──────────┬──────────┐
1 │ Product  │ Widget   │ Gadget   │ Tool     │ ← Search this row
2 │ Price    │ 25.00    │ 40.00    │ 15.00    │ ← Row 2
3 │ Stock    │ 150      │ 200      │ 75       │ ← Row 3
  └──────────┴──────────┴──────────┴──────────┘
                ↑                               
            Search for "Gadget"

=HLOOKUP("Gadget", A1:D3, 2, FALSE)  →  40.00
=HLOOKUP("Gadget", A1:D3, 3, FALSE)  →  200
```

### Example: Monthly Sales
```
Data:
     A          B          C          D          E
  ┌──────────┬──────────┬──────────┬──────────┬──────────┐
1 │ Month    │ Jan      │ Feb      │ Mar      │ Apr      │
2 │ Sales    │ 50000    │ 55000    │ 48000    │ 62000    │
3 │ Expenses │ 35000    │ 38000    │ 36000    │ 40000    │
  └──────────┴──────────┴──────────┴──────────┴──────────┘

Lookup:
=HLOOKUP("Mar", A1:E3, 2, FALSE)  →  48000 (Sales)
=HLOOKUP("Feb", A1:E3, 3, FALSE)  →  38000 (Expenses)
```

### HLOOKUP vs VLOOKUP

| Aspect | VLOOKUP | HLOOKUP |
|--------|---------|---------|
| **Search direction** | Vertical (down) | Horizontal (across) |
| **Table layout** | Columns | Rows |
| **Return parameter** | col_index_num | row_index_num |
| **Common usage** | Very common (90%+) | Rare (10%-) |

**Note:** Most data in Excel is organized in columns, so VLOOKUP is used far more frequently.

---

## XLOOKUP Function

**Purpose:** Modern replacement for VLOOKUP/HLOOKUP with more flexibility

**Syntax:** `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`

**Available in:** Excel 2021, Microsoft 365, Excel Online

### Why XLOOKUP is Better

| Feature | VLOOKUP | XLOOKUP |
|---------|---------|---------|
| **Direction** | Left to right only | Any direction |
| **Return value** | Must be to right of lookup column | Can be anywhere |
| **Column numbers** | Must count columns | References actual range |
| **Not found** | Returns #N/A | Can specify custom message |
| **Default match** | Approximate (risky) | Exact (safer) |
| **Simpler syntax** | No | Yes |

### Basic XLOOKUP Structure
```
=XLOOKUP(what_to_find, where_to_find_it, what_to_return)
```

### Example 1: Basic XLOOKUP
```
Data:
     A          B          C
  ┌──────┬──────────┬──────────┐
1 │ ID   │ Name     │ Dept     │
2 │ 101  │ Alice    │ Sales    │
3 │ 102  │ Bob      │ IT       │
4 │ 103  │ Carol    │ HR       │
  └──────┴──────────┴──────────┘

VLOOKUP way:
=VLOOKUP(102, A2:C4, 2, FALSE)  →  "Bob"

XLOOKUP way:
=XLOOKUP(102, A2:A4, B2:B4)  →  "Bob"
        ↑     ↑       ↑
        Find  Here    Return this
```

### Example 2: Lookup to the Left (VLOOKUP can't do this!)
```
Data:
     A          B          C
  ┌──────────┬──────┬──────────┐
1 │ Name     │ ID   │ Dept     │
2 │ Alice    │ 101  │ Sales    │
3 │ Bob      │ 102  │ IT       │
4 │ Carol    │ 103  │ HR       │
  └──────────┴──────┴──────────┘

Find name based on ID:
=XLOOKUP(102, B2:B4, A2:A4)  →  "Bob"
             ↑       ↑
          Search   Return (to the LEFT!)

VLOOKUP can't do this! Would need helper column or INDEX/MATCH.
```

### Example 3: Custom "Not Found" Message
```
=XLOOKUP(999, A2:A10, B2:B10, "ID Not Found")
                              ↑
                     If not found, return this

Instead of #N/A error, shows "ID Not Found"
```

### Example 4: Return Entire Row
```
Data:
     A          B          C          D
  ┌──────┬──────────┬──────────┬──────────┐
1 │ ID   │ Name     │ Dept     │ Salary   │
2 │ 101  │ Alice    │ Sales    │ 50000    │
3 │ 102  │ Bob      │ IT       │ 60000    │
4 │ 103  │ Carol    │ HR       │ 55000    │
  └──────┴──────────┴──────────┴──────────┘

Return all info for ID 102:
=XLOOKUP(102, A2:A4, B2:D4)
             ↑       ↑
          Search   Return entire range

Returns: Bob    IT    60000 (spills across cells)
```

### Real-World Example: Product Catalog
```
Catalog:
     A          B          C          D
  ┌──────────┬─────────┬─────────┬──────────┐
1 │ SKU      │ Product │ Price   │ Category │
2 │ WDG-001  │ Widget  │ 25.00   │ Tools    │
3 │ GAD-002  │ Gadget  │ 40.00   │ Tech     │
4 │ TOL-003  │ Tool    │ 15.00   │ Tools    │
  └──────────┴─────────┴─────────┴──────────┘

Order Form:
     E          F          G
  ┌──────────┬──────────┬─────────┐
1 │ SKU      │ Product  │ Price   │
2 │ GAD-002  │ =XLOOKUP(E2,$A$2:$A$4,$B$2:$B$4) │ =XLOOKUP(E2,$A$2:$A$4,$C$2:$C$4)
  └──────────┴──────────┴─────────┘
              ↓          ↓
           "Gadget"    40.00
```

### XLOOKUP Match Modes

| Mode | Value | Description |
|------|-------|-------------|
| **Exact** | 0 (default) | Exact match, #N/A if not found |
| **Approximate** | -1 | Exact or next smallest |
| **Approximate** | 1 | Exact or next largest |
| **Wildcard** | 2 | Wildcard match (*, ?) |

### Example: Wildcard Match
```
=XLOOKUP("*son", A2:A10, B2:B10, , 2)
         ↑                         ↑
      Wildcard                  Match mode

Finds: "Johnson", "Anderson", "Wilson", etc.
```

---

## INDEX and MATCH Combination

**Purpose:** Flexible lookup that works in any direction (the "old way" before XLOOKUP)

**Why use INDEX/MATCH?**
- Works in all Excel versions
- Can lookup left
- More flexible than VLOOKUP
- Building block for advanced formulas

### How INDEX Works

**Syntax:** `=INDEX(array, row_num, [col_num])`

**Purpose:** Returns a value from a specific position in a range

```
Range:
     A          B          C
  ┌──────┬──────────┬──────────┐
1 │ 101  │ Alice    │ Sales    │ ← Row 1
2 │ 102  │ Bob      │ IT       │ ← Row 2
3 │ 103  │ Carol    │ HR       │ ← Row 3
  └──────┴──────────┴──────────┘
   Col1    Col2       Col3

=INDEX(A1:C3, 2, 3)  →  "IT"
              ↑   ↑
            Row 2, Column 3
```

### How MATCH Works

**Syntax:** `=MATCH(lookup_value, lookup_array, [match_type])`

**Purpose:** Returns the **position** (row number) of a value in a range

```
Range:
     A
  ┌──────┐
1 │ 101  │ ← Position 1
2 │ 102  │ ← Position 2
3 │ 103  │ ← Position 3
  └──────┘

=MATCH(102, A1:A3, 0)  →  2
       ↑    ↑        ↑
      Find  Here   Exact match

Returns position: 2 (102 is in the 2nd row)
```

### Combining INDEX and MATCH
```
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
       ↑             ↑
       Get value     At this position
       from here     (found by MATCH)
```

### Visual Flow
```
Step 1: MATCH finds the position
=MATCH(102, A2:A4, 0)  →  2

Step 2: INDEX returns value at that position
=INDEX(B2:B4, 2)  →  "Bob"

Combined:
=INDEX(B2:B4, MATCH(102, A2:A4, 0))  →  "Bob"
```

### Example 1: Basic INDEX/MATCH
```
Data:
     A          B          C
  ┌──────┬──────────┬──────────┐
1 │ ID   │ Name     │ Dept     │
2 │ 101  │ Alice    │ Sales    │
3 │ 102  │ Bob      │ IT       │
4 │ 103  │ Carol    │ HR       │
  └──────┴──────────┴──────────┘

Find name for ID 102:
=INDEX(B2:B4, MATCH(102, A2:A4, 0))  →  "Bob"
       ↑      ↑
     Return   Find 102's position (returns 2)
     from     Then get 2nd value from B2:B4
     here
```

### Example 2: Lookup to the Left
```
Data:
     A          B          C
  ┌──────────┬──────┬──────────┐
1 │ Name     │ ID   │ Dept     │
2 │ Alice    │ 101  │ Sales    │
3 │ Bob      │ 102  │ IT       │
4 │ Carol    │ 103  │ HR       │
  └──────────┴──────┴──────────┘

Find name (column A) based on ID (column B):
=INDEX(A2:A4, MATCH(102, B2:B4, 0))  →  "Bob"
       ↑      ↑
    Return    Find position of 102 in column B
    from A    Then get that position from column A

VLOOKUP can't do this! (can only look right)
```

### Example 3: Two-Way Lookup
```
Table:
     A          B          C          D
  ┌──────────┬──────────┬──────────┬──────────┐
1 │ Product  │ Q1       │ Q2       │ Q3       │
2 │ Widget   │ 1000     │ 1200     │ 1100     │
3 │ Gadget   │ 1500     │ 1600     │ 1800     │
4 │ Tool     │ 800      │ 900      │ 850      │
  └──────────┴──────────┴──────────┴──────────┘

Find Q2 sales for Gadget:
=INDEX(B2:D4, MATCH("Gadget",A2:A4,0), MATCH("Q2",B1:D1,0))
       ↑      ↑                         ↑
      Data    Find row (returns 2)     Find column (returns 2)
              Then return B2:D4[row 2, column 2] = 1600
```

### INDEX/MATCH vs VLOOKUP

| Feature | VLOOKUP | INDEX/MATCH |
|---------|---------|-------------|
| **Lookup direction** | Right only | Any direction |
| **Insert columns** | Breaks formula | Still works |
| **Speed** | Fast | Slightly slower |
| **Complexity** | Simple | More complex |
| **Flexibility** | Limited | Very flexible |

**When to use each:**
- **VLOOKUP:** Simple right lookups, newer users
- **INDEX/MATCH:** Need to lookup left, professional templates, complex scenarios
- **XLOOKUP:** If available, best of both worlds!

---

## Lookup Function Comparison

### Summary Table

| Function | Direction | Excel Version | Complexity | Flexibility |
|----------|-----------|---------------|------------|-------------|
| **VLOOKUP** | Vertical, right only | All | Low | Low |
| **HLOOKUP** | Horizontal only | All | Low | Low |
| **XLOOKUP** | Any | 2021+, 365 | Low | High |
| **INDEX/MATCH** | Any | All | Medium | High |

### Which Should You Learn?

**Priority Order:**
1. **VLOOKUP** - Most common, learn first
2. **XLOOKUP** - If you have Excel 365/2021
3. **INDEX/MATCH** - For professional work, maximum flexibility
4. **HLOOKUP** - Only if your data is horizontal (rare)

---

## Advanced Lookup Techniques

### Technique 1: Multiple Criteria Lookup

**Problem:** Lookup based on TWO conditions

```
Data:
     A          B          C
  ┌──────────┬──────────┬──────────┐
1 │ Product  │ Region   │ Sales    │
2 │ Widget   │ East     │ 1000     │
3 │ Widget   │ West     │ 1200     │
4 │ Gadget   │ East     │ 1500     │
5 │ Gadget   │ West     │ 1800     │
  └──────────┴──────────┴──────────┘

Find: Sales for "Widget" in "West"
```

**Solution 1: Helper Column**
```
Add column D (Concat):
D2: =A2&"-"&B2  →  "Widget-East"
D3: =A3&"-"&B3  →  "Widget-West"

Then use VLOOKUP:
=VLOOKUP("Widget-West", D2:E5, 2, FALSE)  →  1200
```

**Solution 2: XLOOKUP (if available)**
```
=XLOOKUP(1, (A2:A5="Widget")*(B2:B5="West"), C2:C5)
            ↑
         Returns 1 where both conditions are TRUE
```

**Solution 3: INDEX/MATCH with multiple criteria**
```
=INDEX(C2:C5, MATCH(1, (A2:A5="Widget")*(B2:B5="West"), 0))

Note: Enter as array formula (Ctrl+Shift+Enter in older Excel)
```

### Technique 2: Partial Match Lookup

**Problem:** Find value containing specific text

```
Data:
     A                    B
  ┌──────────────────┬──────────┐
1 │ Description      │ Code     │
2 │ Widget Pro 2000  │ WDG      │
3 │ Gadget Ultra     │ GAD      │
4 │ Tool Master      │ TOL      │
  └──────────────────┴──────────┘

Find code for any item containing "Widget"
```

**Solution: XLOOKUP with wildcards**
```
=XLOOKUP("*Widget*", A2:A4, B2:B4, , 2)
         ↑                           ↑
      Wildcard                   Match mode 2
                                 (wildcard)
```

**Solution: INDEX/MATCH with wildcards**
```
=INDEX(B2:B4, MATCH("*Widget*", A2:A4, 0))
```

### Technique 3: Closest Match Lookup

**Problem:** Find nearest value (not exact)

```
Price Tier Table (sorted):
     A          B
  ┌─────────┬──────────┐
1 │ Qty     │ Price    │
2 │ 1       │ 19.99    │
3 │ 10      │ 17.99    │
4 │ 50      │ 14.99    │
5 │ 100     │ 12.99    │
  └─────────┴──────────┘

Customer orders 35 units. What's the price?
```

**Solution: VLOOKUP with TRUE**
```
=VLOOKUP(35, A2:B5, 2, TRUE)  →  17.99

Finds: Largest value ≤ 35, which is 10
Returns: 17.99
```

---

## Real-World Application: Order Processing System

Let's build a complete order system using lookup functions.

### Setup Tables

**Product Table (E2:H6):**
```
     E          F          G          H
  ┌──────────┬─────────┬─────────┬──────────┐
1 │ SKU      │ Product │ Price   │ Stock    │
2 │ WDG-001  │ Widget  │ 25.00   │ 150      │
3 │ GAD-002  │ Gadget  │ 40.00   │ 200      │
4 │ TOL-003  │ Tool    │ 15.00   │ 75       │
5 │ DEV-004  │ Device  │ 60.00   │ 120      │
  └──────────┴─────────┴─────────┴──────────┘
```

**Order Form:**
```
     A          B          C          D          E
  ┌──────────┬──────────┬─────────┬─────────┬──────────┐
1 │ SKU      │ Product  │ Price   │ Qty     │ Total    │
2 │ GAD-002  │          │         │ 3       │          │
3 │ WDG-001  │          │         │ 5       │          │
4 │ DEV-004  │          │         │ 2       │          │
  └──────────┴──────────┴─────────┴─────────┴──────────┘
```

### Formulas

**Column B (Product Name):**
```
B2: =IFERROR(VLOOKUP(A2,$E$2:$H$6,2,FALSE),"")
```

**Column C (Price):**
```
C2: =IFERROR(VLOOKUP(A2,$E$2:$H$6,3,FALSE),"")
```

**Column E (Total):**
```
E2: =IF(C2="","",C2*D2)
```

**Complete Order Form with Results:**
```
     A          B          C          D          E
  ┌──────────┬──────────┬─────────┬─────────┬──────────┐
1 │ SKU      │ Product  │ Price   │ Qty     │ Total    │
2 │ GAD-002  │ Gadget   │ 40.00   │ 3       │ 120.00   │
3 │ WDG-001  │ Widget   │ 25.00   │ 5       │ 125.00   │
4 │ DEV-004  │ Device   │ 60.00   │ 2       │ 120.00   │
5 │          │          │         │         │          │
6 │ Grand Total:        │         │         │ =SUM(E2:E4) → 365.00
  └──────────┴──────────┴─────────┴─────────┴──────────┘
```

**Add Stock Validation:**
```
Column F (In Stock?):
F2: =IF(A2="","",IF(VLOOKUP(A2,$E$2:$H$6,4,FALSE)>=D2,"✓","Out of Stock"))
```

**Enhanced Version with XLOOKUP (if available):**
```
B2: =IFERROR(XLOOKUP(A2,$E$2:$E$6,$F$2:$F$6),"")
C2: =IFERROR(XLOOKUP(A2,$E$2:$E$6,$G$2:$G$6),"")
F2: =IF(A2="","",IF(XLOOKUP(A2,$E$2:$E$6,$H$2:$H$6)>=D2,"✓","Out of Stock"))
```

---

## Reference Functions

Beyond lookups, Excel has functions that work with cell references.

### ROW Function

**Purpose:** Returns the row number of a reference

**Syntax:** `=ROW([reference])`

```
     A          B
  ┌─────────┬──────────┐
1 │ Data    │ Row #    │
2 │ Apple   │ =ROW()      → 2
3 │ Banana  │ =ROW()      → 3
4 │ Cherry  │ =ROW()      → 4
  └─────────┴──────────┘

=ROW(A5)  → 5
=ROW(C10) → 10
```

**Use Case: Auto-numbering**
```
     A          B
  ┌─────────┬──────────┐
1 │ #       │ Item     │
2 │ =ROW()-1│ Widget   │  → 1
3 │ =ROW()-1│ Gadget   │  → 2
4 │ =ROW()-1│ Tool     │  → 3
  └─────────┴──────────┘

Subtract 1 because data starts in row 2
```

### COLUMN Function

**Purpose:** Returns the column number of a reference

**Syntax:** `=COLUMN([reference])`

```
=COLUMN(A1)  → 1  (A is column 1)
=COLUMN(B1)  → 2  (B is column 2)
=COLUMN(Z1)  → 26 (Z is column 26)
=COLUMN(AA1) → 27 (AA is column 27)
```

**Use Case: Dynamic column selection**
```
Header row with dynamic lookup:
     A          B          C          D
  ┌──────────┬──────────┬──────────┬──────────┐
1 │ ID       │ Name     │ Dept     │ Salary   │
2 │ 101      │ Alice    │ Sales    │ 50000    │
  └──────────┴──────────┴──────────┴──────────┘

Formula to get column number:
=MATCH("Dept", A1:D1, 0)  → 3
```

### ROWS Function

**Purpose:** Returns the number of rows in a reference

**Syntax:** `=ROWS(array)`

```
=ROWS(A1:A10)   → 10
=ROWS(A1:C5)    → 5
=ROWS(B2:B100)  → 99
```

**Use Case: Count entries**
```
     A          B
  ┌─────────┬──────────────────┐
1 │ Data    │ Count            │
2 │ Apple   │                  │
3 │ Banana  │                  │
4 │ Cherry  │                  │
5 │         │ =ROWS(A2:A4) → 3 │
  └─────────┴──────────────────┘
```

### COLUMNS Function

**Purpose:** Returns the number of columns in a reference

**Syntax:** `=COLUMNS(array)`

```
=COLUMNS(A1:D1)   → 4
=COLUMNS(A1:Z10)  → 26
=COLUMNS(B:E)     → 4
```

### INDIRECT Function

**Purpose:** Returns a reference specified by a text string

**Syntax:** `=INDIRECT(ref_text)`

**Power:** Build dynamic cell references from text

```
     A          B          C
  ┌─────────┬─────────┬──────────────┐
1 │ B2      │ 100     │              │
2 │ C3      │         │ 200          │
3 │         │         │              │
4 │ Result: │ =INDIRECT(A1)  → 100   │
  └─────────┴─────────┴──────────────┘

A1 contains "B2"
INDIRECT converts "B2" text to actual cell reference B2
Returns value: 100
```

**Example: Dynamic Sheet Reference**
```
     A              B
  ┌────────────┬──────────────────────┐
1 │ Sheet1     │                      │
2 │            │ =INDIRECT(A1&"!A1")  │
  └────────────┴──────────────────────┘

Builds reference: "Sheet1!A1"
Returns value from Sheet1, cell A1
```

**Use Case: Summary from multiple sheets**
```
Sheets: Jan, Feb, Mar, Apr

     A          B
  ┌─────────┬────────────────────────┐
1 │ Jan     │ =INDIRECT(A1&"!B10")   │ → Gets Jan!B10
2 │ Feb     │ =INDIRECT(A2&"!B10")   │ → Gets Feb!B10
3 │ Mar     │ =INDIRECT(A3&"!B10")   │ → Gets Mar!B10
  └─────────┴────────────────────────┘

Pulls B10 from each month's sheet dynamically
```

### OFFSET Function

**Purpose:** Returns a reference offset from a starting cell

**Syntax:** `=OFFSET(reference, rows, cols, [height], [width])`

**Visual Concept:**
```
Starting point: A1

     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────┐
1 │ Start   │         │         │         │
2 │         │         │         │         │
3 │         │         │ Target  │         │
  └─────────┴─────────┴─────────┴─────────┘

=OFFSET(A1, 2, 2)
         ↑  ↑  ↑
      Start|  └── Move 2 columns right
           └───── Move 2 rows down
Result: Points to C3
```

**Example: Dynamic range**
```
     A          B
  ┌─────────┬─────────┐
1 │ Item    │ Value   │
2 │ Apple   │ 10      │
3 │ Banana  │ 20      │
4 │ Cherry  │ 30      │
5 │ Date    │ 40      │
  └─────────┴─────────┘

Get value 3 rows down from B1:
=OFFSET(B1, 3, 0)  → 40
        ↑   ↑  ↑
       B1   3  0 columns
            rows
```

**Use Case: Last N entries**
```
Dynamic SUM of last 5 entries:
=SUM(OFFSET(A1, COUNTA(A:A)-5, 0, 5, 1))
             ↑   ↑              ↑  ↑
            Start from A1       |  1 column wide
                Count-5         5 rows tall
                Starting row
```

---

## Common Lookup Mistakes

### Mistake 1: Not Using Absolute References
```
❌ Wrong: =VLOOKUP(A2, E2:G10, 2, FALSE)
Copy down: =VLOOKUP(A3, E3:G11, 2, FALSE)  ← Table moves!

✅ Right: =VLOOKUP(A2, $E$2:$G$10, 2, FALSE)
Copy down: =VLOOKUP(A3, $E$2:$G$10, 2, FALSE)  ← Table fixed!
```

### Mistake 2: Wrong Column Index
```
Table:
     E          F          G
  ┌──────┬──────────┬──────────┐
2 │ ID   │ Name     │ Dept     │
  └──────┴──────────┴──────────┘
   Col1    Col2       Col3

❌ Wrong: =VLOOKUP(A2, E2:G10, 3, FALSE)
Looking for column 3 to get Name → Returns Dept instead!

✅ Right: =VLOOKUP(A2, E2:G10, 2, FALSE)
Column 2 gets Name correctly
```

### Mistake 3: Case Sensitivity (Not an Issue!)
```
Excel lookups are NOT case-sensitive:

"widget" = "Widget" = "WIDGET"

All these will match!
```

### Mistake 4: Extra Spaces
```
Lookup value: "Widget"
Table value:  "Widget " (trailing space)

=VLOOKUP("Widget", Table, 2, FALSE)  → #N/A

Fix: Use TRIM to remove spaces
=VLOOKUP(TRIM(A2), Table, 2, FALSE)
```

### Mistake 5: Numbers Stored as Text
```
Lookup: 123 (number)
Table:  "123" (text)

Won't match!

Fix: Convert text to number or number to text
=VLOOKUP(TEXT(A2,"0"), Table, 2, FALSE)
or
=VLOOKUP(A2, VALUE(Table_Column), 2, FALSE)
```

### Mistake 6: Approximate Match on Unsorted Data
```
❌ Wrong:
=VLOOKUP(A2, Table, 2, TRUE)  when table is not sorted

Returns wrong results!

✅ Fix: Either sort table or use FALSE
=VLOOKUP(A2, Table, 2, FALSE)
```

---

## Performance Tips

### 1. Use Exact Match When Possible
```
✅ Faster: =VLOOKUP(A2, Table, 2, FALSE)
⚠️ Slower: =VLOOKUP(A2, Table, 2, TRUE) then checks sorting
```

### 2. Limit Table Range
```
❌ Slow: =VLOOKUP(A2, E:G, 2, FALSE)  (checks 1M+ rows)
✅ Fast: =VLOOKUP(A2, E2:G1000, 2, FALSE)  (checks 999 rows)
```

### 3. Use Named Ranges
```
✅ Better performance and readability:
=VLOOKUP(A2, ProductTable, 2, FALSE)
```

### 4. XLOOKUP is Generally Faster
```
VLOOKUP: Must count columns, check each row
XLOOKUP: Direct reference, optimized search

For large datasets, XLOOKUP can be significantly faster
```

### 5. Consider Helper Columns for Complex Lookups
```
Instead of:
=INDEX(C:C, MATCH(1, (A:A=E2)*(B:B=F2), 0))  ← Array formula, slow

Use helper column:
D2: =A2&"-"&B2
Then: =VLOOKUP(E2&"-"&F2, D:E, 2, FALSE)  ← Much faster
```

---

## Troubleshooting Lookup Errors

### Error: #N/A (Value Not Available)
**Causes:**
- Lookup value doesn't exist in table
- Extra spaces or formatting differences
- Numbers vs text mismatch
- Spelling differences

**Solutions:**
```
1. Wrap in IFERROR:
=IFERROR(VLOOKUP(A2, Table, 2, FALSE), "Not Found")

2. Check for spaces:
=VLOOKUP(TRIM(A2), Table, 2, FALSE)

3. Check data types:
Use TEXT() or VALUE() to convert

4. Use wildcard match (if appropriate):
=XLOOKUP("*"&A2&"*", Table_Col, Return_Col, , 2)
```

### Error: #REF! (Invalid Reference)
**Causes:**
- Column index exceeds table width
- Deleted rows/columns that formulas reference

**Solution:**
```
Check column count:
Table E2:G10 has 3 columns (E, F, G)
Max col_index_num = 3

=VLOOKUP(A2, E2:G10, 4, FALSE)  ← #REF! (no 4th column)
=VLOOKUP(A2, E2:G10, 3, FALSE)  ← Works
```

### Error: #VALUE! (Wrong Type)
**Causes:**
- Array formulas not entered correctly
- Wrong argument type

**Solution:**
```
For array formulas in older Excel:
Press Ctrl+Shift+Enter (not just Enter)

Or use newer functions that don't require array entry
```

### Wrong Results (No Error)
**Causes:**
- Using TRUE when table isn't sorted
- Searching in wrong column
- Wrong col_index_num

**Solution:**
```
1. Verify table is sorted (if using TRUE)
2. Check formula refers to correct range
3. Count columns carefully (1, 2, 3...)
4. Test with simple examples first
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- VLOOKUP syntax: `=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)`
- FALSE = exact match (most common)
- TRUE = approximate match (requires sorted data)
- Column index starts at 1 (first column of table)
- XLOOKUP syntax: `=XLOOKUP(lookup_value, lookup_array, return_array)`
- INDEX returns value at position
- MATCH returns position of value
- Always use $ for table references

### Practice Deeply
- Writing VLOOKUP formulas for different scenarios
- Using FALSE vs TRUE appropriately
- Combining VLOOKUP with IFERROR
- Using XLOOKUP (if available in your Excel)
- Building INDEX/MATCH formulas
- Creating two-way lookups
- Troubleshooting #N/A errors
- Using named ranges with lookups
- Building order forms with lookups
- Creating dynamic reference systems
- Handling missing data gracefully

### Don't Memorize
- Every possible lookup scenario
- Exact error messages
- Which Excel version has which function
- Complex nested lookup formulas (build step by step)
- All reference functions (learn as needed)

---

## Quick Reference: Lookup Functions

### Most Common Lookups

```
VLOOKUP (Vertical Lookup):
=VLOOKUP(A2, $E$2:$H$10, 3, FALSE)
         ↑    ↑          ↑   ↑
       Find  Table      Col  Exact

XLOOKUP (Modern, Flexible):
=XLOOKUP(A2, $E$2:$E$10, $F$2:$F$10)
         ↑    ↑           ↑
       Find  Search      Return

INDEX/MATCH (Any Direction):
=INDEX($F$2:$F$10, MATCH(A2, $E$2:$E$10, 0))
       ↑           ↑
     Return      Find position
```

### With Error Handling

```
=IFERROR(VLOOKUP(A2, Table, 2, FALSE), "Not Found")
=IFERROR(XLOOKUP(A2, Search, Return), "Not Found")
=IFERROR(INDEX(Return, MATCH(A2, Search, 0)), "Not Found")
```

### Common Patterns

```
Lookup with default value:
=VLOOKUP(A2, Table, 2, FALSE)
Alternative: =XLOOKUP(A2, Search, Return, "Default")

Lookup returning multiple columns (XLOOKUP):
=XLOOKUP(A2, Search, Return_Range)  ← Spills across cells

Two-way lookup (INDEX/MATCH):
=INDEX(Data, MATCH(Row_Value, Row_Range, 0), MATCH(Col_Value, Col_Range, 0))
```

---

## Real-World Scenarios

### Scenario 1: Employee Directory
```
Master List (E2:H100):
SKU | Name | Department | Extension

Lookup Form:
A2: [Enter Employee ID]
B2: =IFERROR(VLOOKUP(A2,$E$2:$H$100,2,FALSE),"")
C2: =IFERROR(VLOOKUP(A2,$E$2:$H$100,3,FALSE),"")
D2: =IFERROR(VLOOKUP(A2,$E$2:$H$100,4,FALSE),"")
```

### Scenario 2: Invoice Generator
```
Product Database: SKU | Product | Price | Tax Rate

Invoice:
Column A: SKU
Column B: =VLOOKUP(A2, ProductDB, 2, FALSE)  ← Product
Column C: =VLOOKUP(A2, ProductDB, 3, FALSE)  ← Price
Column D: Quantity (manual entry)
Column E: =C2*D2  ← Subtotal
Column F: =E2*VLOOKUP(A2, ProductDB, 4, FALSE)  ← Tax
Column G: =E2+F2  ← Total
```

### Scenario 3: Grade Calculator
```
Grading Scale: Score | Letter

Student Scores:
Column A: Student Name
Column B: Numeric Score
Column C: =VLOOKUP(B2, GradeScale, 2, TRUE)  ← Letter Grade

Note: Uses TRUE for range lookup
```

### Scenario 4: Multi-Sheet Consolidation
```
Sheets: North, South, East, West (same structure)

Summary Sheet:
Column A: Region Names (North, South, East, West)
Column B: =INDIRECT(A2&"!B20")  ← Gets B20 from each regional sheet
Column C: =INDIRECT(A3&"!C20")  ← Dynamic reference

Pulls data from multiple sheets automatically
```

---

## Advanced: Combining Functions

### Lookup + Logical Functions
```
Priority lookup with fallback:
=IF(ISNA(VLOOKUP(A2, Table1, 2, FALSE)), 
    VLOOKUP(A2, Table2, 2, FALSE), 
    VLOOKUP(A2, Table1, 2, FALSE))

Try Table1 first, if not found, try Table2
```

### Lookup + Math Functions
```
Lookup and calculate:
=VLOOKUP(A2, Table, 2, FALSE) * 1.15

Apply 15% markup to looked-up price
```

### Lookup + Text Functions
```
Lookup with partial match:
=VLOOKUP("*"&A2&"*", Table, 2, FALSE)

Or using XLOOKUP:
=XLOOKUP("*"&A2&"*", Search, Return, , 2)
```

### Nested Lookups
```
Lookup within a lookup:
=VLOOKUP(
    VLOOKUP(A2, Table1, 2, FALSE),  ← Inner lookup
    Table2, 
    3, 
    FALSE
)

First lookup finds a value, second lookup uses that value
```

---

## Next Step

After mastering lookup functions, you're ready to explore:

**`07-text-functions.md`**
- Combining text with CONCATENATE and TEXTJOIN
- Extracting text with LEFT, RIGHT, MID
- Finding text with FIND and SEARCH
- Cleaning text with TRIM, CLEAN, SUBSTITUTE
- Converting case with UPPER, LOWER, PROPER
- Text-to-columns functionality
- Advanced text manipulation techniques
