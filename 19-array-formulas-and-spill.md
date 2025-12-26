# Array Formulas and Spill

This file covers dynamic arrays and spill behavior - Excel's modern approach to array formulas that automatically fills multiple cells with results, making complex calculations simpler and more powerful.

---

## What are Dynamic Arrays?

**Dynamic Arrays** are formulas that return multiple values automatically, "spilling" results into adjacent cells.

### The Revolution

**Before (Excel 2019 and earlier):**
```
Formula in A1: =UNIQUE(B1:B10)
Press Ctrl+Shift+Enter
Results appear only in A1
Must manually copy down
```

**After (Excel 365/2021):**
```
Formula in A1: =UNIQUE(B1:B10)
Press Enter (normal)
Results automatically spill into A1:A5
Update source → Spill updates automatically
```

### Visual Concept

**Traditional formula:**
```
     A         B
  ┌────────┬────────┐
1 │ =B1*2  │ 5      │
  ├────────┼────────┤
2 │ =B2*2  │ 10     │ ← Must copy formula down
  ├────────┼────────┤
3 │ =B3*2  │ 15     │ ← Each cell needs formula
  └────────┴────────┘

3 separate formulas
```

**Dynamic array (spill):**
```
     A         B
  ┌────────┬────────┐
1 │ =B1:B3*2│ 5     │ ← One formula
  ├────────┼────────┤
2 │ 10     │ 10     │ ← Spilled result
  ├────────┼────────┤
3 │ 20     │ 15     │ ← Spilled result
  │ 30     │        │ ← Spilled result
  └────────┴────────┘

ONE formula, multiple results!
Gray border = spill range
```

---

## Understanding Spill Behavior

### What is "Spilling"?

**Spilling** = Formula automatically fills adjacent cells with array results.

**Characteristics:**
- Only **first cell** contains the formula
- Other cells show **spilled results** (no formula)
- Spilled cells have **gray border**
- Entire spill range updates when formula changes
- Can't edit individual spilled cells

### Spill Range

**Visual indicators:**
```
     A         B
  ┌════════┬────────┐
1 │ =SORT(...) │    │ ← Formula cell (blue border)
  ╠════════╣────────┤
2 ║ Apple  ║        │
  ╠────────╣        │ ← Spill range (gray border)
3 ║ Banana ║        │
  ╠────────╣        │
4 ║ Cherry ║        │
  ╚════════╝────────┘

Click any spilled cell → Selects entire range
```

### Spill Range Reference

**Use # (spill operator) to reference entire spilled range:**

```excel
Formula in A1: =SORT(B1:B10)
Results spill to A1:A5

Reference spilled range:
=SUM(A1#)

# means "all spilled results from A1"
Instead of trying to guess A1:A5
```

**Example:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ =UNIQUE(B:B) │ Apple│ =SUM(A1#) │
  ├────────┼────────┼────────┤
2 │ Banana │ Apple  │ 155    │
  ├────────┼────────┤        │
3 │ Cherry │ Banana │        │
  ├────────┼────────┤        │
4 │ Date   │ Apple  │        │
  └────────┴────────┴────────┘

A1# = A1:A4 (all unique values)
SUM adds them all
```

---

## Dynamic Array Functions

Excel 365/2021 introduced new functions that return arrays.

### SORT Function

**Sort data automatically**

**Syntax:**
```excel
=SORT(array, [sort_index], [sort_order], [by_col])

array: Range to sort
sort_index: Column number to sort by (default 1)
sort_order: 1=ascending, -1=descending (default 1)
by_col: FALSE=sort rows, TRUE=sort columns (default FALSE)
```

**Example 1: Simple sort**
```
     A         B
  ┌────────┬────────┐
1 │ =SORT(B1:B5) │ Cherry │
  ├────────┼────────┤
2 │ Apple  │ Apple  │
  ├────────┼────────┤
3 │ Banana │ Date   │
  ├────────┼────────┤
4 │ Cherry │ Banana │
  ├────────┼────────┤
5 │ Date   │        │
  └────────┴────────┘

Alphabetically sorted, spills A1:A4
```

**Example 2: Sort by second column, descending**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Score  │ =SORT(A1:B4,2,-1) │
  ├────────┼────────┼────────┤
2 │ John   │ 85     │ Sarah  │ 95 │
  ├────────┼────────┼────────┼────┤
3 │ Sarah  │ 95     │ Mike   │ 88 │
  ├────────┼────────┼────────┼────┤
4 │ Mike   │ 88     │ John   │ 85 │
  └────────┴────────┴────────┴────┘

Sorted by score (column 2), highest first
```

### FILTER Function

**Extract rows that meet criteria**

**Syntax:**
```excel
=FILTER(array, include, [if_empty])

array: Range to filter
include: Condition (TRUE/FALSE for each row)
if_empty: Value if no results (default: #CALC! error)
```

**Example 1: Filter by value**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Region │ =FILTER(A1:B5,B1:B5="East") │
  ├────────┼────────┼────────┤
2 │ John   │ East   │ John   │ East │
  ├────────┼────────┼────────┼──────┤
3 │ Sarah  │ West   │ Mike   │ East │
  ├────────┼────────┤        │      │
4 │ Mike   │ East   │        │      │
  ├────────┼────────┤        │      │
5 │ Lisa   │ South  │        │      │
  └────────┴────────┴────────┴──────┘

Shows only East region rows
```

**Example 2: Multiple criteria (AND)**
```excel
=FILTER(A1:C10,(B1:B10="East")*(C1:C10>1000))

Both conditions must be TRUE
* works as AND
```

**Example 3: Multiple criteria (OR)**
```excel
=FILTER(A1:C10,(B1:B10="East")+(B1:B10="West"))

Either condition can be TRUE
+ works as OR
```

**Example 4: With default if empty**
```excel
=FILTER(A1:B10,B1:B10="North","No results found")

If no matches, shows "No results found"
Instead of #CALC! error
```

### UNIQUE Function

**Extract unique values, remove duplicates**

**Syntax:**
```excel
=UNIQUE(array, [by_col], [exactly_once])

array: Range to analyze
by_col: FALSE=compare rows (default), TRUE=compare columns
exactly_once: FALSE=all unique, TRUE=only items that appear once
```

**Example 1: Simple unique list**
```
     A         B
  ┌────────┬────────┐
1 │ =UNIQUE(B1:B7) │ Apple  │
  ├────────┼────────┤
2 │ Apple  │ Banana │
  ├────────┼────────┤
3 │ Banana │ Apple  │
  ├────────┼────────┤
4 │ Cherry │ Cherry │
  └────────┼────────┤
           │ Banana │
           ├────────┤
           │ Apple  │
           ├────────┤
           │ Cherry │
           └────────┘

Returns: Apple, Banana, Cherry
```

**Example 2: Items that appear exactly once**
```excel
=UNIQUE(B1:B7,FALSE,TRUE)

Returns only values that appear once in list
If Apple appears 3 times, it's excluded
```

**Example 3: Unique combinations**
```excel
=UNIQUE(A1:B10)

Returns unique combinations of both columns
```

### SORTBY Function

**Sort by another column/range**

**Syntax:**
```excel
=SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], ...)

array: Range to sort
by_array1: Range to sort by
sort_order1: 1=ascending, -1=descending
```

**Example 1: Sort names by scores**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Score  │ =SORTBY(A1:A4,B1:B4,-1) │
  ├────────┼────────┼────────┤
2 │ John   │ 85     │ Sarah  │
  ├────────┼────────┼────────┤
3 │ Sarah  │ 95     │ Mike   │
  ├────────┼────────┼────────┤
4 │ Mike   │ 88     │ John   │
  └────────┴────────┴────────┘

Names sorted by their scores (descending)
```

**Example 2: Multiple sort keys**
```excel
=SORTBY(A1:B10,B1:B10,1,C1:C10,-1)

Sort by column B ascending, then column C descending
```

### SEQUENCE Function

**Generate sequence of numbers**

**Syntax:**
```excel
=SEQUENCE(rows, [columns], [start], [step])

rows: Number of rows
columns: Number of columns (default 1)
start: Starting number (default 1)
step: Increment (default 1)
```

**Example 1: Simple sequence**
```excel
=SEQUENCE(5)

Results:
1
2
3
4
5
```

**Example 2: Start at 10, step by 5**
```excel
=SEQUENCE(4,1,10,5)

Results:
10
15
20
25
```

**Example 3: 2D sequence**
```excel
=SEQUENCE(3,4)

Results:
1  2  3  4
5  6  7  8
9  10 11 12
```

**Example 4: Dates sequence**
```excel
=SEQUENCE(7,1,TODAY(),1)

Next 7 days starting today
```

### RANDARRAY Function

**Generate array of random numbers**

**Syntax:**
```excel
=RANDARRAY([rows], [columns], [min], [max], [integer])

rows: Number of rows (default 1)
columns: Number of columns (default 1)
min: Minimum value (default 0)
max: Maximum value (default 1)
integer: TRUE=whole numbers, FALSE=decimals (default FALSE)
```

**Example 1: 5 random decimals**
```excel
=RANDARRAY(5)

Results (between 0 and 1):
0.523
0.891
0.234
0.678
0.145
```

**Example 2: Random integers 1-100**
```excel
=RANDARRAY(10,1,1,100,TRUE)

10 random whole numbers between 1 and 100
```

**Example 3: Random 3x3 grid**
```excel
=RANDARRAY(3,3,1,10,TRUE)

Results:
7  3  9
2  8  1
5  4  6
```

---

## Array Operations

Perform calculations on entire arrays at once.

### Basic Arithmetic

**Multiply array by scalar:**
```
     A         B
  ┌────────┬────────┐
1 │ =B1:B3*2│ 5     │
  ├────────┼────────┤
2 │ 10     │ 10     │ ← Results spill
  ├────────┼────────┤
3 │ 20     │ 15     │
  │ 30     │        │
  └────────┴────────┘

Each value doubled
```

**Add two arrays:**
```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ =B1:B3+C1:C3│ 5  │ 2      │
  ├────────┼────────┼────────┤
2 │ 15     │ 10     │ 5      │
  ├────────┼────────┼────────┤
3 │ 25     │ 15     │ 10     │
  │ 7      │        │        │
  ├────────┤        │        │
4 │ 17     │        │        │
  ├────────┤        │        │
5 │ 35     │        │        │
  └────────┴────────┴────────┘

Element-wise addition
```

### Combining Functions

**Sort unique values:**
```excel
=SORT(UNIQUE(A1:A100))

1. UNIQUE removes duplicates
2. SORT alphabetizes result
```

**Filter and sort:**
```excel
=SORT(FILTER(A1:C100,B1:B100="East"))

1. FILTER gets East region rows
2. SORT arranges them
```

**Count unique values:**
```excel
=ROWS(UNIQUE(A1:A100))

1. UNIQUE gets distinct values
2. ROWS counts them
```

### Text Operations

**Convert to uppercase:**
```excel
=UPPER(A1:A10)

All text in range converted to uppercase
Results spill
```

**Concatenate:**
```excel
=A1:A5&" "&B1:B5

Combines first and last names with space
```

**Extract first word:**
```excel
=LEFT(A1:A10,FIND(" ",A1:A10)-1)

Gets text before first space
Works on entire array
```

---

## Advanced Array Formulas

### Nested Dynamic Functions

**Example 1: Top 5 unique values**
```excel
=SORT(UNIQUE(A1:A100),-1)

Then manually look at first 5
Or use TAKE:
=TAKE(SORT(UNIQUE(A1:A100),-1),5)
```

**Example 2: Filtered unique list**
```excel
=UNIQUE(FILTER(A1:B100,B1:B100>1000))

1. Filter rows where column B > 1000
2. Get unique combinations
```

**Example 3: Sorted filtered list**
```excel
=SORT(FILTER(A1:C100,(B1:B100="East")*(C1:C100>5000)),3,-1)

1. Filter: East region AND sales > 5000
2. Sort by column 3, descending
```

### XLOOKUP with Arrays

**Return multiple columns:**
```excel
=XLOOKUP(E1,A1:A100,B1:D100)

Looks up E1 in column A
Returns entire row from columns B:D
```

**Array of lookups:**
```excel
=XLOOKUP(E1:E5,A1:A100,B1:B100)

Looks up E1, E2, E3, E4, E5
Returns 5 results
```

### SUMIFS with Arrays

**Multiple criteria sums:**
```excel
=SUMIFS(C1:C100,A1:A100,E1:E5,B1:B100,"East")

Sums for each value in E1:E5
Where region = East
Returns 5 sums
```

---

## #SPILL! Error

The most common array formula error.

### What Causes #SPILL!

**Blocked spill range:**
```
     A         B
  ┌────────┬────────┐
1 │ =SORT(B1:B5)    │
  ├────────┼────────┤
2 │ #SPILL!│ Cherry │
  ├────────┼────────┤
3 │ DATA   │ Apple  │ ← Cell not empty!
  ├────────┼────────┤
4 │        │ Banana │
  └────────┴────────┘

A3 contains "DATA"
SORT wants to spill to A1:A4
A3 is blocking it → #SPILL!
```

### Identifying the Blockage

**Click #SPILL! cell:**
```
┌──────────────────────────────────┐
│ Spill range isn't blank          │
├──────────────────────────────────┤
│ Clear cells: A3                  │ ← Shows blocking cells
│ Or move the formula              │
└──────────────────────────────────┘
```

**Visual indicator:**
- Dashed blue box shows where formula wants to spill
- Blocking cells highlighted

### Fixing #SPILL! Errors

**Solution 1: Clear blocking cells**
```
Delete or move content from A3
Formula spills correctly
```

**Solution 2: Move formula**
```
Move formula to column with more space
E.g., from A1 to F1
```

**Solution 3: Move blocking data**
```
Move "DATA" from A3 to another column
```

### Other Spill Errors

**#CALC! Error:**
```
=FILTER(A1:A100,B1:B100="North","No results")

If no matches found and no if_empty argument:
Results in #CALC!

Fix: Add third argument for empty result
```

**Array size mismatch:**
```excel
=A1:A10+B1:B5

Can't add arrays of different sizes
Fix: Ensure ranges match
```

---

## Legacy Array Formulas (CSE)

Before dynamic arrays, array formulas required **Ctrl+Shift+Enter**.

### What are CSE Formulas?

**CSE** = Ctrl+Shift+Enter

**Old method:**
```
Type formula: =SUM(A1:A10*B1:B10)
Press: Ctrl+Shift+Enter
Result: {=SUM(A1:A10*B1:B10)}
         ↑                  ↑
      Curly braces appear
```

### Modern vs Legacy

**Legacy (CSE required):**
```excel
{=SUM(IF(A1:A10="East",B1:B10,0))}

Must press Ctrl+Shift+Enter
Curly braces added by Excel
Formula in single cell
```

**Modern (Dynamic array):**
```excel
=SUMIF(A1:A10,"East",B1:B10)

Just press Enter
No curly braces
Simpler syntax
```

### When You'll See CSE Formulas

```
✓ Old Excel files (pre-2019)
✓ Legacy workbooks
✓ Files from colleagues using older Excel
✓ Online tutorials written before 2019
```

### Converting CSE to Modern

**Example: Count unique values**

**Legacy CSE:**
```excel
{=SUM(1/COUNTIF(A1:A100,A1:A100))}
Ctrl+Shift+Enter required
```

**Modern:**
```excel
=ROWS(UNIQUE(A1:A100))
Just Enter
```

**Example: Array multiplication**

**Legacy CSE:**
```excel
{=SUM(A1:A10*B1:B10)}
```

**Modern:**
```excel
=SUMPRODUCT(A1:A10,B1:B10)
or
=SUM(A1:A10*B1:B10)  ← Works without CSE in Excel 365!
```

---

## Practical Examples

### Example 1: Dynamic Dropdown List

**Problem:** Dropdown list of unique regions that updates automatically

**Solution:**
```
Cell A1: =SORT(UNIQUE(DataTable[Region]))

Data Validation in another cell:
Source: =A1#

Dropdown shows all unique regions
Updates when new regions added
```

### Example 2: Top 10 Customers

**Data:**
```
Column A: Customer names
Column B: Total sales
```

**Formula:**
```excel
=SORT(A1:B100,2,-1)

Shows all customers sorted by sales (descending)

To get only top 10:
=TAKE(SORT(A1:B100,2,-1),10)
```

### Example 3: Filtered Report

**Problem:** Show only completed orders over $1000

**Formula:**
```excel
=FILTER(Orders[#All],(Orders[Status]="Complete")*(Orders[Amount]>1000),"No matching orders")

Returns complete table
Only rows meeting both criteria
Shows message if none found
```

### Example 4: Ranking with SEQUENCE

**Create rank column:**
```excel
=SEQUENCE(ROWS(A1:A10))

Results:
1, 2, 3, 4, 5, 6, 7, 8, 9, 10

Or custom ranking after sort:
=SORT(A1:B10,2,-1)
Then add column:
=SEQUENCE(ROWS(A1#))
```

### Example 5: Random Sample

**Get 10 random rows from dataset:**
```excel
=SORTBY(A1:C100,RANDARRAY(ROWS(A1:C100)))

Sorts rows by random numbers
Result: Random order

Then take first 10:
=TAKE(SORTBY(A1:C100,RANDARRAY(ROWS(A1:C100))),10)
```

---

## Combining with Tables

### Dynamic Arrays with Structured References

**Works seamlessly:**
```excel
=SORT(Sales[Amount])
=FILTER(Sales[#All],Sales[Region]="East")
=UNIQUE(Sales[Product])
```

**Benefits:**
- Table auto-expands
- Formula always includes new data
- Structured references + dynamic arrays = powerful!

### Example: Filtered Table Display

**Table: SalesData**
```
Columns: Date, Product, Region, Amount
```

**Filtered view:**
```excel
Cell G1:
=FILTER(SalesData[#All],SalesData[Region]="East","No Eastern sales")

Shows all East region rows
Updates automatically
Includes headers
```

---

## Best Practices

### When to Use Dynamic Arrays

```
✅ Perfect for:
- Extracting unique values
- Filtering datasets
- Sorting results
- Generating sequences
- Creating dynamic reports
- Removing duplicates
- Top N analysis
```

### When to Avoid

```
❌ Not ideal for:
- Very large datasets (10,000+ rows might be slow)
- When you need to edit individual results
- Files shared with Excel 2016 or earlier users
- Complex calculations that can be done simpler ways
```

### Performance Tips

```
✅ Use specific ranges (A1:A1000) not entire columns (A:A)
✅ Filter before sorting (reduces data to sort)
✅ Limit nested functions (2-3 max)
✅ Use tables as source (clearer references)
✅ Test with small dataset first
```

### Organization Tips

```
✅ Place dynamic arrays in dedicated area
✅ Leave space below/right for spill
✅ Document formulas (add comment)
✅ Use named ranges for source data
✅ Color-code spill ranges differently
```

---

## Troubleshooting Guide

### Problem: Formula Returns Single Value

**Expected:** Array of results
**Got:** One value

**Cause:** Function doesn't return array, or Excel version doesn't support dynamic arrays

**Solution:**
```
Check Excel version (must be 365 or 2021)
Verify function is array function (SORT, FILTER, UNIQUE, etc.)
Test with simple example: =SEQUENCE(5)
```

### Problem: #SPILL! Error

**Solution steps:**
```
1. Click #SPILL! cell
2. Read error message
3. Check dashed blue box (shows intended spill range)
4. Clear or move blocking cells
5. Or move formula to clearer area
```

### Problem: #CALC! Error in FILTER

**Cause:** No matches found, no default specified

**Solution:**
```excel
=FILTER(A1:A100,B1:B100="North","No results")
                                 ↑
                         Add this parameter
```

### Problem: Results Not Updating

**Cause:** Source data changed but spill didn't update

**Solution:**
```
Press F9 (recalculate)
Check calculation mode (Formulas Tab)
Ensure "Automatic" not "Manual"
Edit formula (F2) and press Enter to force recalc
```

### Problem: Can't Edit Spilled Cell

**Symptom:** Click spilled result, can't type

**Explanation:** Spilled cells are read-only

**Solution:**
```
Edit the formula cell (first cell with formula)
Or click spilled cell → Press Delete (clears entire spill)
```

### Problem: Slow Performance

**Cause:** Large array operation

**Solutions:**
```
Use specific ranges not entire columns
Filter data before other operations
Reduce nested functions
Consider Power Query for very large data
Test calculation time (Formulas Tab → Calculation Options → Manual)
```

---

## Quick Reference: Dynamic Array Functions

| Function | Purpose | Example |
|----------|---------|---------|
| **SORT** | Sort array | `=SORT(A1:A100)` |
| **SORTBY** | Sort by another range | `=SORTBY(A1:A10,B1:B10,-1)` |
| **FILTER** | Extract matching rows | `=FILTER(A1:B100,B1:B100>1000)` |
| **UNIQUE** | Remove duplicates | `=UNIQUE(A1:A100)` |
| **SEQUENCE** | Generate numbers | `=SEQUENCE(10)` |
| **RANDARRAY** | Random numbers | `=RANDARRAY(5,1,1,100,TRUE)` |
| **XLOOKUP** | Lookup (array mode) | `=XLOOKUP(E1:E5,A:A,B:B)` |
| **XMATCH** | Find positions | `=XMATCH(E1:E5,A:A)` |

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl + Shift + Enter` | Legacy array formula (older Excel) |
| `F9` | Recalculate formulas |
| `Ctrl + ` | Show formulas (see spill sources) |
| `Esc` | Cancel spilled range selection |
| `Delete` | Clear entire spill range |

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Dynamic arrays automatically spill into adjacent cells
- Only first cell contains formula, others show results
- #SPILL! means cells are blocking the spill range
- A1# references entire spilled range from A1
- Excel 365/2021+ required for dynamic arrays
- SORT, FILTER, UNIQUE are main dynamic functions
- Can't edit individual spilled cells (edit formula cell)
- Gray border indicates spill range
- Spill range updates automatically when source changes

### Practice Deeply
- Using SORT to organize data
- Using FILTER with single criteria
- Using FILTER with multiple criteria (AND/OR)
- Using UNIQUE to remove duplicates
- Combining SORT and UNIQUE together
- Creating sequences with SEQUENCE
- Referencing spilled ranges with # operator
- Troubleshooting #SPILL! errors
- Clearing blocking cells
- Using SORTBY to sort by another column
- Using array operations (multiply, add arrays)
- Creating dynamic dropdown lists with UNIQUE
- Using FILTER with tables and structured references
- Understanding when formulas will spill
- Testing dynamic arrays with small datasets first
- Converting legacy CSE formulas to modern syntax
- Combining multiple dynamic functions (nested)
- Using IF with FILTER for complex conditions

---

## Common Patterns

### Pattern 1: Unique Sorted List
```excel
=SORT(UNIQUE(A1:A1000))

Clean, alphabetized list of distinct values
```

### Pattern 2: Filtered and Sorted
```excel
=SORT(FILTER(A1:C100,B1:B100="East"),3,-1)

East region only, sorted by column 3 descending
```

### Pattern 3: Top N
```excel
=TAKE(SORT(A1:B100,2,-1),10)

Top 10 by column 2 values
```

### Pattern 4: Dynamic Validation List
```excel
Data Validation Source: =SORT(UNIQUE(Products!A:A))

Dropdown always current, no duplicates, alphabetized
```

### Pattern 5: Conditional Unique
```excel
=UNIQUE(FILTER(A1:A100,B1:B100>1000))

Unique values where column B > 1000
```

---

## Next Step

After this file, we move to:

**`20-power-query-basics.md`**
- Introduction to Power Query
- Get & Transform interface
- Loading and transforming data
- Common transformations (filter, sort, remove columns)
- Merging and appending queries
- Creating reusable queries
- Refreshing data
- Power Query M language basics
