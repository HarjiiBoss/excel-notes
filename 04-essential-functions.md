# Essential Functions

This file introduces the most commonly used Excel functions that form the foundation
of spreadsheet analysis. These functions handle basic calculations and operations
you'll use in nearly every workbook.

---

## What is a Function?

A **function** is a pre-built formula that performs a specific calculation or operation.

Think of functions as **shortcuts** that save you from writing complex formulas manually.

### Function Structure
```
=FUNCTION_NAME(argument1, argument2, ...)
  ↑            ↑
  |            └── Inputs (data the function needs)
  └── Always starts with =
```

### Visual Example
```
Cell A1: 10
Cell A2: 20
Cell A3: 30

Cell A4: =SUM(A1:A3)
         ↓
      Result: 60
```

---

## Anatomy of a Function

Every function has three parts:

### 1. Equals Sign (=)
All functions **must** start with `=`

### 2. Function Name
The name that identifies what the function does.

**Examples:** `SUM`, `AVERAGE`, `COUNT`, `MAX`

**Note:** Function names are **not case-sensitive**
- `=SUM(A1:A10)` works
- `=sum(a1:a10)` also works
- `=Sum(A1:A10)` also works

### 3. Arguments
The inputs enclosed in parentheses `( )`

**Types of arguments:**
- **Cell references:** `A1`, `B2:B10`
- **Ranges:** `A1:A10`
- **Numbers:** `5`, `100`, `3.14`
- **Text:** `"Hello"` (must be in quotes)
- **Other functions:** `SUM(A1:A10)`

### Visual Breakdown
```
=SUM(A1:A10)
│ │  │ │  │
│ │  │ │  └── End of range
│ │  │ └───── Range separator (:)
│ │  └──────── Start of range
│ └─────────── Opening parenthesis
└───────────── Function name
```

---

## Function Categories

Excel has hundreds of functions organized into categories:

| Category | Purpose | Example Functions |
|----------|---------|-------------------|
| **Math & Trig** | Calculations, rounding | SUM, ROUND, ABS |
| **Statistical** | Analysis, averages | AVERAGE, COUNT, MAX, MIN |
| **Logical** | Decision-making | IF, AND, OR |
| **Lookup & Reference** | Find values | VLOOKUP, XLOOKUP, INDEX |
| **Text** | Manipulate text | LEFT, RIGHT, CONCATENATE |
| **Date & Time** | Work with dates | TODAY, NOW, DATE |
| **Financial** | Money calculations | PMT, FV, RATE |
| **Database** | Filter and count | DSUM, DCOUNT, DAVERAGE |

**This file covers:** The most essential statistical and math functions you'll use daily.

---

## SUM Function

**Purpose:** Adds numbers together

**Syntax:** `=SUM(number1, [number2], ...)`

### Basic Usage

**Example 1: Sum a range**
```
     A
  ┌──────┐
1 │  10  │
2 │  20  │
3 │  30  │
4 │  40  │
5 │      │
6 │ =SUM(A1:A4)
  └──────┘
     ↓
   Result: 100
```

**Example 2: Sum multiple ranges**
```
=SUM(A1:A4, C1:C4)
```

**Example 3: Sum individual cells**
```
=SUM(A1, A3, A5)
```

**Example 4: Mix ranges and numbers**
```
=SUM(A1:A10, 100)
```

### Real-World Example
```
     A          B
  ┌─────────┬────────┐
1 │ Month   │ Sales  │
2 │ Jan     │ 5000   │
3 │ Feb     │ 6200   │
4 │ Mar     │ 5800   │
5 │ Apr     │ 7100   │
6 │         │        │
7 │ Total   │ =SUM(B2:B5)
  └─────────┴────────┘
              ↓
           Result: 24100
```

### Common Patterns

✅ **Good:**
```
=SUM(A1:A100)        ← One range
=SUM(A:A)            ← Entire column
=SUM(1:1)            ← Entire row
=SUM(A1:A10, B1:B10) ← Multiple ranges
```

❌ **Avoid:**
```
=A1+A2+A3+A4+A5+A6+A7+A8+A9+A10  ← Too long, use SUM instead
```

### ⚠️ Important Notes
- SUM **ignores text** and blank cells
- SUM counts `0` as a number
- Use `Ctrl + Shift + T` (Windows) or `Cmd + Shift + T` (Mac) to quickly insert SUM

---

## AVERAGE Function

**Purpose:** Calculates the arithmetic mean of numbers

**Syntax:** `=AVERAGE(number1, [number2], ...)`

### Basic Usage

**Example: Average of a range**
```
     A
  ┌──────┐
1 │  80  │
2 │  90  │
3 │  70  │
4 │  85  │
5 │      │
6 │ =AVERAGE(A1:A4)
  └──────┘
     ↓
   Result: 81.25
```

**Calculation:** (80 + 90 + 70 + 85) ÷ 4 = 81.25

### Real-World Example: Student Grades
```
     A          B       C
  ┌─────────┬────────┬──────────┐
1 │ Student │ Score  │ Class Avg│
2 │ Alice   │ 92     │          │
3 │ Bob     │ 78     │ =AVERAGE(B2:B6)
4 │ Carol   │ 85     │          │
5 │ David   │ 88     │          │
6 │ Emma    │ 95     │          │
  └─────────┴────────┴──────────┘
                        ↓
                    Result: 87.6
```

### Comparison: AVERAGE vs AVERAGEA

| Function | Behavior |
|----------|----------|
| **AVERAGE** | Ignores text and empty cells |
| **AVERAGEA** | Treats text as 0, counts empty cells |

**Example:**
```
     A
  ┌─────────┐
1 │  10     │
2 │  20     │
3 │  "N/A"  │
4 │         │
  └─────────┘

=AVERAGE(A1:A4)  → 15    (only counts 10 and 20)
=AVERAGEA(A1:A4) → 7.5   (counts text as 0: (10+20+0+0)÷4)
```

### ⚠️ Important Notes
- AVERAGE **ignores** text, blank cells, and logical values
- To include zeros, use `AVERAGEA`
- To average with conditions, use `AVERAGEIF` (covered later)

---

## COUNT Functions Family

Excel has three COUNT functions for different scenarios:

### COUNT Function

**Purpose:** Counts cells containing **numbers**

**Syntax:** `=COUNT(value1, [value2], ...)`

**Example:**
```
     A
  ┌─────────┐
1 │  10     │ ← Number
2 │  "Bob"  │ ← Text (ignored)
3 │  20     │ ← Number
4 │         │ ← Blank (ignored)
5 │  30     │ ← Number
6 │         │
7 │ =COUNT(A1:A5)
  └─────────┘
     ↓
   Result: 3
```

### COUNTA Function

**Purpose:** Counts cells containing **any value** (numbers, text, dates, etc.)

**Syntax:** `=COUNTA(value1, [value2], ...)`

**Example:**
```
     A
  ┌─────────┐
1 │  10     │ ← Counted
2 │  "Bob"  │ ← Counted
3 │  20     │ ← Counted
4 │         │ ← NOT counted
5 │  30     │ ← Counted
6 │         │
7 │ =COUNTA(A1:A5)
  └─────────┘
     ↓
   Result: 4
```

### COUNTBLANK Function

**Purpose:** Counts **empty cells**

**Syntax:** `=COUNTBLANK(range)`

**Example:**
```
     A
  ┌─────────┐
1 │  10     │
2 │         │ ← Blank
3 │  20     │
4 │         │ ← Blank
5 │  30     │
6 │         │
7 │ =COUNTBLANK(A1:A5)
  └─────────┘
     ↓
   Result: 2
```

### COUNT Functions Comparison

```
     A          COUNT  COUNTA  COUNTBLANK
  ┌─────────┐
1 │  10     │    ✓      ✓
2 │  "Bob"  │           ✓
3 │  20     │    ✓      ✓
4 │         │                    ✓
5 │  TRUE   │           ✓
  └─────────┘
   Result:     2       3         1
```

### Real-World Example: Survey Responses
```
     A           B
  ┌──────────┬─────────┐
1 │ Name     │ Rating  │
2 │ John     │ 5       │
3 │ Sarah    │         │ ← Not responded
4 │ Mike     │ 4       │
5 │ Lisa     │ 3       │
6 │          │         │
7 │ Total Responses:  │ =COUNTA(B2:B5)  → 3
8 │ Numeric Ratings:  │ =COUNT(B2:B5)   → 3
9 │ No Response:      │ =COUNTBLANK(B2:B5) → 1
  └──────────┴─────────┘
```

---

## MAX and MIN Functions

### MAX Function

**Purpose:** Returns the **largest** value in a range

**Syntax:** `=MAX(number1, [number2], ...)`

**Example:**
```
     A
  ┌──────┐
1 │  45  │
2 │  92  │ ← Highest
3 │  67  │
4 │  81  │
5 │      │
6 │ =MAX(A1:A4)
  └──────┘
     ↓
   Result: 92
```

### MIN Function

**Purpose:** Returns the **smallest** value in a range

**Syntax:** `=MIN(number1, [number2], ...)`

**Example:**
```
     A
  ┌──────┐
1 │  45  │ ← Lowest
2 │  92  │
3 │  67  │
4 │  81  │
5 │      │
6 │ =MIN(A1:A4)
  └──────┘
     ↓
   Result: 45
```

### Real-World Example: Temperature Tracking
```
     A          B
  ┌─────────┬────────┐
1 │ Day     │ Temp°F │
2 │ Monday  │ 72     │
3 │ Tuesday │ 68     │ ← Coldest
4 │ Wed     │ 75     │
5 │ Thu     │ 80     │ ← Hottest
6 │ Friday  │ 77     │
7 │         │        │
8 │ High:   │ =MAX(B2:B6) → 80
9 │ Low:    │ =MIN(B2:B6) → 68
10│ Range:  │ =MAX(B2:B6)-MIN(B2:B6) → 12
  └─────────┴────────┘
```

### ⚠️ Important Notes
- MAX and MIN **ignore text** and empty cells
- Returns 0 if range contains no numbers
- For dates, MAX returns latest date, MIN returns earliest date

---

## ROUND Function

**Purpose:** Rounds a number to a specified number of digits

**Syntax:** `=ROUND(number, num_digits)`

**Arguments:**
- `number` - The value to round
- `num_digits` - Number of decimal places
  - Positive: rounds to decimals
  - Zero: rounds to nearest integer
  - Negative: rounds to left of decimal

### Examples

**Example 1: Round to 2 decimal places**
```
=ROUND(3.14159, 2)  →  3.14
=ROUND(3.14159, 3)  →  3.142
=ROUND(3.14159, 0)  →  3
```

**Example 2: Round to tens, hundreds**
```
=ROUND(1234.56, -1)  →  1230   (nearest 10)
=ROUND(1234.56, -2)  →  1200   (nearest 100)
=ROUND(1234.56, -3)  →  1000   (nearest 1000)
```

### Visual Guide
```
Number: 1234.56789
         │ │ │││││
         │ │ │││└└┴─ num_digits = 4
         │ │ ││└──── num_digits = 3
         │ │ │└───── num_digits = 2
         │ │ └────── num_digits = 1
         │ └──────── num_digits = 0
         └────────── num_digits = -1, -2, -3...

=ROUND(1234.56789, 2)   →  1234.57
=ROUND(1234.56789, 0)   →  1235
=ROUND(1234.56789, -1)  →  1230
```

### Related Rounding Functions

| Function | Purpose | Example |
|----------|---------|---------|
| **ROUND** | Rounds to specified digits | `=ROUND(2.5, 0)` → 3 |
| **ROUNDUP** | Always rounds up | `=ROUNDUP(2.1, 0)` → 3 |
| **ROUNDDOWN** | Always rounds down | `=ROUNDDOWN(2.9, 0)` → 2 |
| **MROUND** | Rounds to nearest multiple | `=MROUND(13, 5)` → 15 |
| **INT** | Rounds down to integer | `=INT(2.9)` → 2 |
| **TRUNC** | Truncates decimals | `=TRUNC(2.9)` → 2 |

### Real-World Example: Invoice Calculation
```
     A            B          C
  ┌──────────┬──────────┬──────────────┐
1 │ Item     │ Price    │ Rounded      │
2 │ Widget   │ 12.3456  │ =ROUND(B2,2) │ → $12.35
3 │ Gadget   │ 8.7891   │ =ROUND(B3,2) │ → $8.79
4 │ Tool     │ 15.9234  │ =ROUND(B4,2) │ → $15.92
5 │          │          │              │
6 │ Total    │ =SUM(B2:B4) │ =SUM(C2:C4) │
7 │          │ 37.0581  │ 37.06        │
  └──────────┴──────────┴──────────────┘
```

### ⚠️ Important Notes
- ROUND uses "round half up" rule (2.5 → 3, not 2)
- ROUND changes the **displayed value**, not just formatting
- For currency, always use `ROUND(..., 2)`

---

## ABS Function

**Purpose:** Returns the **absolute value** (distance from zero, always positive)

**Syntax:** `=ABS(number)`

### Examples
```
=ABS(5)    →  5
=ABS(-5)   →  5
=ABS(0)    →  0
=ABS(-3.7) →  3.7
```

### Visual Concept
```
Number Line:
    -5   -4   -3   -2   -1    0    1    2    3    4    5
    ●─────────────────────────┼─────────────────────────●
    ↑                         ↑                         ↑
  ABS(-5) = 5            ABS(0) = 0              ABS(5) = 5
  
  Distance from 0 is always positive
```

### Real-World Example: Variance Analysis
```
     A          B          C            D
  ┌─────────┬─────────┬──────────┬───────────────┐
1 │ Budget  │ Actual  │ Variance │ Abs Variance  │
2 │ 1000    │ 1100    │ =B2-A2   │ =ABS(C2)      │
3 │ 500     │ 450     │ =B3-A3   │ =ABS(C3)      │
4 │ 750     │ 800     │ =B4-A4   │ =ABS(C4)      │
  └─────────┴─────────┴──────────┴───────────────┘
                          ↓            ↓
                       100          100
                       -50           50
                        50           50
```

**Use case:** When you want the magnitude of difference, not direction.

---

## MOD Function

**Purpose:** Returns the **remainder** after division

**Syntax:** `=MOD(number, divisor)`

### Examples
```
=MOD(10, 3)  →  1    (10 ÷ 3 = 3 remainder 1)
=MOD(15, 4)  →  3    (15 ÷ 4 = 3 remainder 3)
=MOD(20, 5)  →  0    (20 ÷ 5 = 4 remainder 0)
=MOD(7, 2)   →  1    (7 ÷ 2 = 3 remainder 1)
```

### Visual Concept
```
MOD(10, 3) = ?

10 ÷ 3 = 3 with remainder 1
┌───┬───┬───┐  ┌─┐
│ 3 │ 3 │ 3 │  │1│ ← This is the MOD
└───┴───┴───┘  └─┘
    9          + 1 = 10
```

### Real-World Example 1: Even/Odd Detection
```
     A        B
  ┌──────┬─────────────┐
1 │ Num  │ Even or Odd │
2 │  5   │ =IF(MOD(A2,2)=0,"Even","Odd")  → Odd
3 │  8   │ =IF(MOD(A3,2)=0,"Even","Odd")  → Even
4 │  13  │ =IF(MOD(A4,2)=0,"Even","Odd")  → Odd
  └──────┴─────────────┘
```

### Real-World Example 2: Alternating Row Colors
```
     A         B
  ┌───────┬──────────────────┐
1 │ Item  │ Row Color Formula│
2 │ A     │ =MOD(ROW(),2)=0  │ → FALSE (odd row)
3 │ B     │ =MOD(ROW(),2)=0  │ → TRUE (even row)
4 │ C     │ =MOD(ROW(),2)=0  │ → FALSE (odd row)
5 │ D     │ =MOD(ROW(),2)=0  │ → TRUE (even row)
  └───────┴──────────────────┘

Use in Conditional Formatting to alternate colors
```

### Real-World Example 3: Every Nth Item
```
     A          B
  ┌─────────┬────────────────────┐
1 │ Item #  │ Every 3rd Item?    │
2 │   1     │ =MOD(A2,3)=0       │ → FALSE
3 │   2     │ =MOD(A3,3)=0       │ → FALSE
4 │   3     │ =MOD(A4,3)=0       │ → TRUE ✓
5 │   4     │ =MOD(A5,3)=0       │ → FALSE
6 │   5     │ =MOD(A6,3)=0       │ → FALSE
7 │   6     │ =MOD(A7,3)=0       │ → TRUE ✓
  └─────────┴────────────────────┘
```

---

## SUMIF Function

**Purpose:** Sums values based on a **condition**

**Syntax:** `=SUMIF(range, criteria, [sum_range])`

**Arguments:**
- `range` - The range to check against criteria
- `criteria` - The condition to meet
- `sum_range` - (Optional) The actual values to sum

### Basic Pattern
```
=SUMIF(where_to_check, what_to_look_for, what_to_sum)
```

### Example 1: Sum if equal to value
```
     A          B
  ┌─────────┬────────┐
1 │ Region  │ Sales  │
2 │ East    │ 1000   │
3 │ West    │ 1500   │
4 │ East    │ 1200   │
5 │ West    │ 900    │
6 │ East    │ 800    │
7 │         │        │
8 │ East Total: │ =SUMIF(A2:A6,"East",B2:B6)
  └─────────┴────────┘
              ↓
           Result: 3000 (1000+1200+800)
```

**How it works:**
1. Checks each cell in A2:A6
2. If it equals "East"
3. Sum the corresponding value from B2:B6

### Example 2: Sum if greater than value
```
     A          B
  ┌─────────┬────────┐
1 │ Product │ Sales  │
2 │ Widget  │ 500    │
3 │ Gadget  │ 1500   │
4 │ Tool    │ 800    │
5 │ Item    │ 2000   │
6 │         │        │
7 │ Sales over 1000: │ =SUMIF(B2:B5,">1000")
  └─────────┴────────┘
                 ↓
              Result: 3500 (1500+2000)
```

### Criteria Examples

| Criteria | Meaning | Example |
|----------|---------|---------|
| `"East"` | Equals "East" | `=SUMIF(A:A,"East",B:B)` |
| `">100"` | Greater than 100 | `=SUMIF(A:A,">100")` |
| `">=50"` | Greater than or equal to 50 | `=SUMIF(A:A,">=50",B:B)` |
| `"<100"` | Less than 100 | `=SUMIF(A:A,"<100")` |
| `"<>0"` | Not equal to 0 | `=SUMIF(A:A,"<>0",B:B)` |
| `A1` | Equals value in A1 | `=SUMIF(B:B,A1,C:C)` |

### Real-World Example: Sales Report
```
     A          B          C
  ┌─────────┬──────────┬────────┐
1 │ Item    │ Category │ Amount │
2 │ Laptop  │ Tech     │ 1200   │
3 │ Desk    │ Furniture│ 300    │
4 │ Mouse   │ Tech     │ 25     │
5 │ Chair   │ Furniture│ 250    │
6 │ Monitor │ Tech     │ 400    │
7 │ Table   │ Furniture│ 500    │
8 │         │          │        │
9 │ Tech Total:     │ =SUMIF(B2:B7,"Tech",C2:C7)
10│ Furniture Total:│ =SUMIF(B2:B7,"Furniture",C2:C7)
  └─────────┴──────────┴────────┘
                         ↓           ↓
                      1625         1050
```

### ⚠️ Important Notes
- Text criteria must be in **quotes**: `"East"`, `">100"`
- Cell references don't need quotes: `A1`
- Wildcard characters: `*` (multiple) and `?` (single)
  - `"*son"` matches "Johnson", "Anderson"
  - `"A?"` matches "AB", "A1", but not "ABC"

---

## COUNTIF Function

**Purpose:** Counts cells based on a **condition**

**Syntax:** `=COUNTIF(range, criteria)`

### Example 1: Count specific values
```
     A
  ┌─────────┐
1 │ Status  │
2 │ Pass    │
3 │ Fail    │
4 │ Pass    │
5 │ Pass    │
6 │ Fail    │
7 │         │
8 │ Passes: │ =COUNTIF(A2:A6,"Pass")  → 3
9 │ Fails:  │ =COUNTIF(A2:A6,"Fail")  → 2
  └─────────┘
```

### Example 2: Count with comparison
```
     A
  ┌──────┐
1 │ Age  │
2 │ 25   │
3 │ 32   │
4 │ 18   │
5 │ 45   │
6 │ 29   │
7 │      │
8 │ Over 30: │ =COUNTIF(A2:A6,">30")  → 2
9 │ Under 21:│ =COUNTIF(A2:A6,"<21")  → 1
  └──────┘
```

### Real-World Example: Survey Analysis
```
     A             B
  ┌──────────┬──────────┐
1 │ Response │ Count    │
2 │ Yes      │          │
3 │ No       │          │
4 │ Yes      │          │
5 │ Yes      │          │
6 │ Maybe    │          │
7 │ No       │          │
8 │ Yes      │          │
9 │          │          │
10│ Yes:   │ =COUNTIF(A2:A8,"Yes")   → 4
11│ No:    │ =COUNTIF(A2:A8,"No")    → 2
12│ Maybe: │ =COUNTIF(A2:A8,"Maybe") → 1
  └──────────┴──────────┘
```

---

## AVERAGEIF Function

**Purpose:** Calculates average based on a **condition**

**Syntax:** `=AVERAGEIF(range, criteria, [average_range])`

### Example: Average by category
```
     A          B
  ┌─────────┬────────┐
1 │ Product │ Price  │
2 │ Laptop  │ 1200   │
3 │ Mouse   │ 25     │
4 │ Laptop  │ 1500   │
5 │ Mouse   │ 30     │
6 │ Laptop  │ 1100   │
7 │         │        │
8 │ Avg Laptop: │ =AVERAGEIF(A2:A6,"Laptop",B2:B6)
9 │ Avg Mouse:  │ =AVERAGEIF(A2:A6,"Mouse",B2:B6)
  └─────────┴────────┘
                 ↓              ↓
              1266.67          27.5
```

---

## Combining Functions (Nesting)

You can use functions **inside** other functions.

### Example 1: Average of highest values
```
=AVERAGE(MAX(A1:A10), MAX(B1:B10), MAX(C1:C10))
```

### Example 2: Rounded average
```
=ROUND(AVERAGE(A1:A10), 2)
  │    └──────────────┘  │
  │           │          │
  │      Calculate avg   │
  └──── Then round to 2 decimals
```

### Example 3: Sum of absolute values
```
=SUM(ABS(A1), ABS(A2), ABS(A3))
```

### Real-World Example: Grade Calculation
```
     A         B        C        D
  ┌────────┬────────┬────────┬──────────┐
1 │ Test 1 │ Test 2 │ Test 3 │ Final    │
2 │ 85     │ 92     │ 78     │ =ROUND(AVERAGE(A2:C2),1)
  └────────┴────────┴────────┴──────────┘
                                  ↓
                              Result: 85.0
```

**Step-by-step:**
1. `AVERAGE(A2:C2)` calculates 85
2. `ROUND(..., 1)` rounds to 1 decimal place

### ⚠️ Important: Function Limits
- Excel supports up to **64 levels of nesting**
- More than 3-4 levels becomes hard to read
- Break complex formulas into helper columns when possible

---

## Common Mistakes and Best Practices

### Mistake 1: Forgetting the Equals Sign
```
❌ Wrong: SUM(A1:A10)
✅ Right: =SUM(A1:A10)
```

### Mistake 2: Using Commas vs. Semicolons
Depends on your Excel region settings:
- **US/UK:** Use commas: `=SUM(A1,A2,A3)`
- **Europe:** Use semicolons: `=SUM(A1;A2;A3)`

**Check:** Type `=SUM(` and see what Excel suggests.

### Mistake 3: Not Anchoring References
```
❌ Problem: =SUM(A1:A10)  (copies as A2:A11, A3:A12...)
✅ Solution: =SUM($A$1:$A$10)  (always references A1:A10)
```
(More on this in File 02: Cell References and Ranges)

### Mistake 4: Text vs. Numbers
```
❌ Problem:
     A
  ┌──────┐
1 │ "5"  │ ← Text (left-aligned)
2 │ "10" │
3 │ =SUM(A1:A2)  → Result: 0

✅ Solution: Enter as numbers, not text
     A
  ┌──────┐
1 │  5   │ ← Number (right-aligned)
2 │  10  │
3 │ =SUM(A1:A2)  → Result: 15
```

### Mistake 5: Mixing Data Types in Ranges
```
❌ Avoid:
     A
  ┌──────────┐
1 │ Sales    │ ← Header in data range
2 │ 100      │
3 │ 200      │
4 │ =SUM(A1:A3)  → Result: 300 (ignores text)

✅ Better:
     A
  ┌──────────┐
1 │ Sales    │ ← Header
2 │ 100      │
3 │ 200      │
4 │ =SUM(A2:A3)  → Result: 300 (clear intent)
```

### Mistake 6: Using Entire Columns with Large Datasets
```
❌ Slow: =SUM(A:A)  (checks all 1,048,576 rows)
✅ Fast: =SUM(A2:A1000)  (specific range)
```

**Exception:** For small datasets or growing lists, entire columns are fine.

---

## Best Practices

### 1. Use Specific Ranges
```
✅ Good: =SUM(A2:A50)
❌ Avoid: =SUM(A2:A99999) if you only have 50 rows
```

### 2. Keep Formulas Simple
If a formula is too complex, break it into steps:

**Complex (hard to debug):**
```
=ROUND(AVERAGE(IF(A2:A10>0,A2:A10)),2)
```

**Better (multiple cells):**
```
B2: =IF(A2>0,A2,"")  (filter positive)
C2: =AVERAGE(B2:B10)  (calculate average)
D2: =ROUND(C2,2)  (round result)
```

### 3. Use Meaningful Names
Instead of `=SUM(C2:C50)`, consider using named ranges:
```
=SUM(Sales_Amount)
```
(Covered in File 17: Named Ranges)

### 4. Document Complex Formulas
Add comments to cells explaining what formulas do:
- Right-click → Insert Comment
- Explain the purpose and logic

### 5. Check Your Data Types
Before using functions, verify:
- Numbers are truly numbers (right-aligned)
- Dates are properly formatted
- No hidden spaces in text

---

## Function Categories Quick Reference

### Math & Statistical Functions

| Function | Purpose | Example |
|----------|---------|---------|
| **SUM** | Add numbers | `=SUM(A1:A10)` |
| **AVERAGE** | Calculate mean | `=AVERAGE(A1:A10)` |
| **COUNT** | Count numbers | `=COUNT(A1:A10)` |
| **COUNTA** | Count non-empty cells | `=COUNTA(A1:A10)` |
| **COUNTBLANK** | Count empty cells | `=COUNTBLANK(A1:A10)` |
| **MAX** | Find largest | `=MAX(A1:A10)` |
| **MIN** | Find smallest | `=MIN(A1:A10)` |
| **MEDIAN** | Find middle value | `=MEDIAN(A1:A10)` |
| **MODE** | Find most common | `=MODE(A1:A10)` |
| **ROUND** | Round numbers | `=ROUND(A1,2)` |
| **ROUNDUP** | Round up | `=ROUNDUP(A1,2)` |
| **ROUNDDOWN** | Round down | `=ROUNDDOWN(A1,2)` |
| **ABS** | Absolute value | `=ABS(A1)` |
| **MOD** | Remainder | `=MOD(A1,B1)` |
| **SQRT** | Square root | `=SQRT(A1)` |
| **POWER** | Raise to power | `=POWER(A1,2)` |

### Conditional Functions

| Function | Purpose | Example |
|----------|---------|---------|
| **SUMIF** | Sum with condition | `=SUMIF(A:A,"East",B:B)` |
| **COUNTIF** | Count with condition | `=COUNTIF(A:A,">100")` |
| **AVERAGEIF** | Average with condition | `=AVERAGEIF(A:A,"Pass",B:B)` |

---

## Practical Exercise: Sales Dashboard

Let's combine multiple functions to create a simple sales dashboard.

### Setup
```
     A          B        C         D
  ┌─────────┬────────┬────────┬──────────┐
1 │ Region  │ Sales  │ Target │ Variance │
2 │ East    │ 5000   │ 4500   │          │
3 │ West    │ 4200   │ 5000   │          │
4 │ North   │ 6100   │ 6000   │          │
5 │ South   │ 3900   │ 4000   │          │
6 │ East    │ 5200   │ 4500   │          │
  └─────────┴────────┴────────┴──────────┘
```

### Formulas to Add

**1. Variance (Column D):**
```
D2: =B2-C2
```

**2. Total Sales (Below table):**
```
B7: =SUM(B2:B6)
```

**3. Average Sales:**
```
B8: =AVERAGE(B2:B6)
```

**4. Highest Sales:**
```
B9: =MAX(B2:B6)
```

**5. Lowest Sales:**
```
B10: =MIN(B2:B6)
```

**6. East Region Total:**
```
B11: =SUMIF(A2:A6,"East",B2:B6)
```

**7. Regions Over Target:**
```
B12: =COUNTIF(D2:D6,">0")
```

**8. Average Variance:**
```
B13: =ROUND(AVERAGE(D2:D6),0)
```

### Complete Dashboard
```
     A          B        C         D
  ┌─────────┬────────┬────────┬──────────┐
1 │ Region  │ Sales  │ Target │ Variance │
2 │ East    │ 5000   │ 4500   │ 500      │
3 │ West    │ 4200   │ 5000   │ -800     │
4 │ North   │ 6100   │ 6000   │ 100      │
5 │ South   │ 3900   │ 4000   │ -100     │
6 │ East    │ 5200   │ 4500   │ 700      │
7 │         │        │        │          │
8 │ Total Sales:      │ 24400  │          │
9 │ Average Sales:    │ 4880   │          │
10│ Highest Sales:    │ 6100   │          │
11│ Lowest Sales:     │ 3900   │          │
12│ East Total:       │ 10200  │          │
13│ Over Target:      │ 3      │          │
14│ Avg Variance:     │ 80     │          │
  └─────────┴────────┴────────┴──────────┘
```

---

## Quick Function Finder

### "I want to..."

**Add numbers together**
→ `=SUM(range)`

**Find the average**
→ `=AVERAGE(range)`

**Count how many numbers**
→ `=COUNT(range)`

**Count how many cells have any value**
→ `=COUNTA(range)`

**Count empty cells**
→ `=COUNTBLANK(range)`

**Find the largest value**
→ `=MAX(range)`

**Find the smallest value**
→ `=MIN(range)`

**Round to 2 decimals**
→ `=ROUND(number, 2)`

**Remove negative sign**
→ `=ABS(number)`

**Find remainder after division**
→ `=MOD(number, divisor)`

**Add only if condition is met**
→ `=SUMIF(range, criteria, sum_range)`

**Count only if condition is met**
→ `=COUNTIF(range, criteria)`

**Average only if condition is met**
→ `=AVERAGEIF(range, criteria, average_range)`

---

## Troubleshooting Common Errors

### Error: #DIV/0!
**Cause:** Dividing by zero or empty cell

**Example:**
```
=AVERAGE(A1:A10)  where all cells are empty
```

**Fix:**
- Ensure cells contain numbers
- Use `IFERROR` to handle (covered in File 05)

### Error: #VALUE!
**Cause:** Wrong type of argument

**Example:**
```
=SUM(A1:A10)  where cells contain text like "N/A"
```

**Fix:**
- Clean data to remove text from number columns
- SUM automatically ignores text, but nested functions may not

### Error: #NAME?
**Cause:** Excel doesn't recognize function name

**Example:**
```
=SUMM(A1:A10)  ← Typo (extra M)
```

**Fix:**
- Check spelling
- Ensure function exists in your Excel version

### Error: #REF!
**Cause:** Invalid cell reference

**Example:**
- Deleting rows/columns that formulas reference

**Fix:**
- Update formulas with valid references
- Use named ranges to avoid this issue

---

## Tips for Faster Formula Writing

### 1. AutoComplete Function Names
Start typing `=SU` and Excel suggests:
- SUM
- SUMIF
- SUMIFS
- And more...

Press **Tab** to accept suggestion.

### 2. Function Tooltips
After typing `=SUM(`, Excel shows:
```
SUM(number1, [number2], ...)
```

This reminds you what arguments are needed.

### 3. Select Range with Mouse
1. Type `=SUM(`
2. **Click and drag** to select range
3. Press Enter

Excel fills in the range automatically.

### 4. Use F4 to Toggle Reference Types
1. Type `=SUM(A1:A10)`
2. Click on `A1:A10` in formula
3. Press **F4** to cycle:
   - `A1:A10` (relative)
   - `$A$1:$A$10` (absolute)
   - `A$1:A$10` (mixed)
   - `$A1:$A10` (mixed)

### 5. Double-Click Fill Handle
After entering formula in first cell:
1. **Hover** over bottom-right corner (fill handle)
2. **Double-click** to auto-fill down to last row with data

```
     A         B
  ┌──────┬──────────┐
1 │ Num  │ Doubled  │
2 │ 5    │ =A2*2    │ ← Enter formula
3 │ 10   │          │
4 │ 15   │          │ ← Double-click fill handle
5 │ 20   │          │    to auto-fill
  └──────┴──────────┘
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- All functions start with `=`
- Basic syntax: `=FUNCTION_NAME(arguments)`
- SUM, AVERAGE, COUNT are the most used functions
- SUMIF, COUNTIF, AVERAGEIF use pattern: `(range, criteria, [sum_range])`
- Criteria in quotes: `"East"`, `">100"`
- Functions ignore text and blank cells (generally)

### Practice Deeply
- Writing SUM, AVERAGE, COUNT, MAX, MIN formulas
- Using SUMIF, COUNTIF, AVERAGEIF with different criteria
- Rounding numbers with ROUND function
- Using ABS for variance analysis
- Using MOD for even/odd and patterns
- Nesting functions (function inside function)
- Selecting ranges with mouse while typing formulas
- Using AutoComplete for function names
- Reading and fixing function errors
- Building a complete dashboard with multiple functions
- Combining functions to solve real problems

### Don't Memorize
- Every single function Excel has (hundreds!)
- Exact argument order (use tooltips)
- All error codes (look them up when needed)

---

## Function Cheat Sheet

### Most Used Functions (Learn These First)
```
=SUM(A1:A10)           → Add numbers
=AVERAGE(A1:A10)       → Calculate mean
=COUNT(A1:A10)         → Count numbers
=MAX(A1:A10)           → Find highest
=MIN(A1:A10)           → Find lowest
=ROUND(A1,2)           → Round to 2 decimals
```

### Conditional Functions (Very Powerful)
```
=SUMIF(A:A,"East",B:B)      → Sum if condition met
=COUNTIF(A:A,">100")        → Count if condition met
=AVERAGEIF(A:A,"Pass",B:B)  → Average if condition met
```

### Counting Variations
```
=COUNT(A:A)        → Count numbers only
=COUNTA(A:A)       → Count any value
=COUNTBLANK(A:A)   → Count empty cells
```

### Helpful Math Functions
```
=ABS(A1)           → Absolute value
=MOD(A1,2)         → Remainder (even/odd)
=ROUNDUP(A1,2)     → Always round up
=ROUNDDOWN(A1,2)   → Always round down
```

---

## Real-World Scenarios

### Scenario 1: Student Grade Book
**Goal:** Calculate final grades

```
     A        B      C      D      E        F
  ┌───────┬──────┬──────┬──────┬──────┬─────────┐
1 │ Name  │ T1   │ T2   │ T3   │ Final│ Pass?   │
2 │ Alice │ 85   │ 90   │ 88   │ =AVERAGE(B2:D2) │
3 │ Bob   │ 72   │ 68   │ 75   │ =AVERAGE(B3:D3) │
4 │       │      │      │      │      │         │
5 │ Class Average: │ =AVERAGE(E2:E3)  │         │
6 │ Highest Score: │ =MAX(E2:E3)      │         │
7 │ Passed (≥70):  │ =COUNTIF(E2:E3,">=70") │   │
  └───────┴──────┴──────┴──────┴──────┴─────────┘
```

### Scenario 2: Expense Tracking
**Goal:** Summarize expenses by category

```
     A           B
  ┌──────────┬────────┐
1 │ Category │ Amount │
2 │ Food     │ 150    │
3 │ Transport│ 75     │
4 │ Food     │ 200    │
5 │ Bills    │ 300    │
6 │ Transport│ 50     │
7 │          │        │
8 │ Food Total:      │ =SUMIF(A2:A6,"Food",B2:B6)
9 │ Transport Total: │ =SUMIF(A2:A6,"Transport",B2:B6)
10│ Bills Total:     │ =SUMIF(A2:A6,"Bills",B2:B6)
11│ Grand Total:     │ =SUM(B2:B6)
  └──────────┴────────┘
```

### Scenario 3: Inventory Check
**Goal:** Identify low stock items

```
     A          B          C
  ┌─────────┬──────────┬─────────┐
1 │ Item    │ Stock    │ Status  │
2 │ Widget  │ 45       │ =IF(B2<50,"Low","OK")
3 │ Gadget  │ 120      │ =IF(B3<50,"Low","OK")
4 │ Tool    │ 30       │ =IF(B4<50,"Low","OK")
5 │         │          │         │
6 │ Low Stock Count: │ =COUNTIF(C2:C4,"Low")
7 │ Total Items:     │ =SUM(B2:B4)
8 │ Avg Stock:       │ =AVERAGE(B2:B4)
  └─────────┴──────────┴─────────┘
```

---

## Next Step

After mastering these essential functions, you're ready to move to:

**`05-logical-functions.md`**
- IF function for decision-making
- AND, OR, NOT for complex conditions
- Nested IF statements
- IFS function (multiple conditions)
- IFERROR for error handling
- Combining logical functions with calculations
