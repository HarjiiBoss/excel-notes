# Mathematical and Statistical Functions

This file covers Excel's advanced mathematical and statistical functions for
data analysis, including multi-criteria functions, ranking, percentiles, and
statistical measures. These functions are essential for business analytics.

---

## Beyond Basic Math

We covered basic functions (SUM, AVERAGE, COUNT, etc.) in **File 04**.

This file focuses on:
- **Multi-criteria functions** (SUMIFS, COUNTIFS, AVERAGEIFS)
- **Rounding variations** (CEILING, FLOOR, MROUND)
- **Statistical analysis** (MEDIAN, MODE, STDEV, VAR)
- **Ranking and percentiles** (RANK, PERCENTILE, QUARTILE)
- **Random numbers** (RAND, RANDBETWEEN)
- **Advanced calculations**

---

## SUMIFS - Sum with Multiple Criteria

**Syntax:** `=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Purpose:** Sum values that meet **multiple** conditions (AND logic)

### Structure Breakdown
```
=SUMIFS(what_to_sum, where_to_check1, condition1, where_to_check2, condition2, ...)
        ↑            ↑                ↑              ↑                ↑
        Sum this     Check here       Must equal     Check here       Must equal
```

### Basic Example
```
Data:
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Region  │ Product │ Sales   │
2 │ East    │ Widget  │ 1000    │
3 │ East    │ Gadget  │ 1500    │
4 │ West    │ Widget  │ 1200    │
5 │ West    │ Gadget  │ 900     │
6 │ East    │ Widget  │ 800     │
  └─────────┴─────────┴─────────┘

Formula:
=SUMIFS(C2:C6, A2:A6,"East", B2:B6,"Widget")
        ↑      ↑      ↑       ↑      ↑
        Sum    Check  Must be Check  Must be
        Sales  Region "East"  Product "Widget"

Result: 1800 (1000 + 800)
```

### Visual Logic
```
Row 2: East + Widget?  ✓ → Include 1000
Row 3: East + Gadget?  ✗ (wrong product)
Row 4: West + Widget?  ✗ (wrong region)
Row 5: West + Gadget?  ✗ (both wrong)
Row 6: East + Widget?  ✓ → Include 800

Total: 1800
```

### Example 1: Sales by Region and Month
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────┐
1 │ Region  │ Month   │ Sales   │ Product │
2 │ East    │ Jan     │ 5000    │ Widget  │
3 │ East    │ Jan     │ 3000    │ Gadget  │
4 │ West    │ Jan     │ 4000    │ Widget  │
5 │ East    │ Feb     │ 5500    │ Widget  │
  └─────────┴─────────┴─────────┴─────────┘

East + Widget sales:
=SUMIFS(C2:C5, A2:A5,"East", D2:D5,"Widget")
→ 10500 (5000 + 5500)

East + January sales:
=SUMIFS(C2:C5, A2:A5,"East", B2:B5,"Jan")
→ 8000 (5000 + 3000)
```

### Example 2: Using Cell References
```
     A          B          C          E          F
  ┌─────────┬─────────┬─────────┬─────────┬─────────┐
1 │ Region  │ Product │ Sales   │ Region: │ East    │
2 │ East    │ Widget  │ 1000    │ Product:│ Widget  │
3 │ East    │ Gadget  │ 1500    │         │         │
4 │         │         │         │ Total:  │ =SUMIFS($C$2:$C$6,$A$2:$A$6,F1,$B$2:$B$6,F2)
  └─────────┴─────────┴─────────┴─────────┴─────────┘

Dynamic: Change F1 or F2 to get different totals
```

### Example 3: Numeric Criteria
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Product │ Qty     │ Price   │
2 │ Widget  │ 150     │ 25.00   │
3 │ Gadget  │ 50      │ 40.00   │
4 │ Tool    │ 200     │ 15.00   │
5 │ Device  │ 75      │ 60.00   │
  └─────────┴─────────┴─────────┘

Sum sales where quantity > 100:
=SUMIFS(C2:C5, B2:B5,">100")
→ 40.00 (Widget: 25 + Tool: 15)

Sum sales where qty > 100 AND price > 20:
=SUMIFS(C2:C5, B2:B5,">100", C2:C5,">20")
→ 25.00 (Widget only)
```

### Comparison Operators in SUMIFS

| Operator | Example | Meaning |
|----------|---------|---------|
| `"="` or just value | `"East"` or `100` | Equals |
| `">"` | `">100"` | Greater than |
| `"<"` | `"<100"` | Less than |
| `">="` | `">=100"` | Greater or equal |
| `"<="` | `"<=100"` | Less or equal |
| `"<>"` | `"<>0"` | Not equal |

### Real-World Example: Commission Calculator
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────┐
1 │ Rep     │ Region  │ Sales   │ Tier    │
2 │ Alice   │ East    │ 15000   │ Gold    │
3 │ Bob     │ West    │ 8000    │ Silver  │
4 │ Alice   │ East    │ 12000   │ Gold    │
5 │ Carol   │ East    │ 20000   │ Gold    │
6 │ Bob     │ West    │ 9000    │ Silver  │
  └─────────┴─────────┴─────────┴─────────┘

Alice's Gold tier East sales:
=SUMIFS(C2:C6, A2:A6,"Alice", B2:B6,"East", D2:D6,"Gold")
→ 27000 (15000 + 12000)

East region Gold tier sales over 10000:
=SUMIFS(C2:C6, B2:B6,"East", D2:D6,"Gold", C2:C6,">10000")
→ 32000 (12000 + 20000)
```

---

## COUNTIFS - Count with Multiple Criteria

**Syntax:** `=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Purpose:** Count cells that meet **multiple** conditions

### Basic Example
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Region  │ Product │ Sales   │
2 │ East    │ Widget  │ 1000    │
3 │ East    │ Gadget  │ 1500    │
4 │ West    │ Widget  │ 1200    │
5 │ East    │ Widget  │ 800     │
  └─────────┴─────────┴─────────┘

Count East + Widget:
=COUNTIFS(A2:A5,"East", B2:B5,"Widget")
→ 2 (rows 2 and 5)
```

### Example 1: Sales Performance
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Rep     │ Sales   │ Quarter │
2 │ Alice   │ 15000   │ Q1      │
3 │ Bob     │ 8000    │ Q1      │
4 │ Alice   │ 18000   │ Q2      │
5 │ Carol   │ 12000   │ Q1      │
  └─────────┴─────────┴─────────┘

Count Q1 sales over 10000:
=COUNTIFS(C2:C5,"Q1", B2:B5,">10000")
→ 2 (Alice and Carol)

Count Alice's total entries:
=COUNTIFS(A2:A5,"Alice")
→ 2
```

### Example 2: Date Ranges
```
     A              B
  ┌────────────┬─────────┐
1 │ Date       │ Amount  │
2 │ 1/15/2024  │ 100     │
3 │ 2/20/2024  │ 150     │
4 │ 3/10/2024  │ 200     │
5 │ 1/25/2024  │ 120     │
  └────────────┴─────────┘

Count entries in January 2024:
=COUNTIFS(A2:A5,">=1/1/2024", A2:A5,"<2/1/2024")
→ 2

Or using DATE function:
=COUNTIFS(A2:A5,">="&DATE(2024,1,1), A2:A5,"<"&DATE(2024,2,1))
→ 2
```

### Real-World Example: Inventory Management
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────┐
1 │ Product │ Stock   │ Category│ Status  │
2 │ Widget  │ 5       │ Tools   │ Active  │
3 │ Gadget  │ 50      │ Tech    │ Active  │
4 │ Tool    │ 15      │ Tools   │ Active  │
5 │ Device  │ 3       │ Tech    │ Low     │
  └─────────┴─────────┴─────────┴─────────┘

Count active items with low stock (< 10):
=COUNTIFS(D2:D5,"Active", B2:B5,"<10")
→ 1 (Widget)

Count Tech category items:
=COUNTIFS(C2:C5,"Tech")
→ 2
```

---

## AVERAGEIFS - Average with Multiple Criteria

**Syntax:** `=AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Purpose:** Calculate average of values that meet **multiple** conditions

### Basic Example
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Region  │ Product │ Sales   │
2 │ East    │ Widget  │ 1000    │
3 │ East    │ Widget  │ 800     │
4 │ West    │ Widget  │ 1200    │
5 │ East    │ Gadget  │ 1500    │
  └─────────┴─────────┴─────────┘

Average East Widget sales:
=AVERAGEIFS(C2:C5, A2:A6,"East", B2:B5,"Widget")
→ 900 (average of 1000 and 800)
```

### Example 1: Student Performance
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Student │ Subject │ Score   │
2 │ Alice   │ Math    │ 92      │
3 │ Alice   │ English │ 88      │
4 │ Bob     │ Math    │ 78      │
5 │ Alice   │ Math    │ 95      │
  └─────────┴─────────┴─────────┘

Alice's average Math score:
=AVERAGEIFS(C2:C5, A2:A5,"Alice", B2:B5,"Math")
→ 93.5 (average of 92 and 95)

Math scores over 80:
=AVERAGEIFS(C2:C5, B2:B5,"Math", C2:C5,">80")
→ 93.5 (average of 92 and 95)
```

### Example 2: Sales Analysis
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────┐
1 │ Rep     │ Quarter │ Sales   │ Region  │
2 │ Alice   │ Q1      │ 15000   │ East    │
3 │ Bob     │ Q1      │ 8000    │ West    │
4 │ Alice   │ Q2      │ 18000   │ East    │
5 │ Bob     │ Q2      │ 9000    │ West    │
  └─────────┴─────────┴─────────┴─────────┘

Average Q1 sales:
=AVERAGEIFS(C2:C5, B2:B5,"Q1")
→ 11500

Average East region sales:
=AVERAGEIFS(C2:C5, D2:D5,"East")
→ 16500

Average Alice Q2 sales:
=AVERAGEIFS(C2:C5, A2:A5,"Alice", B2:B5,"Q2")
→ 18000
```

---

## Summary: IFS Functions

### Function Comparison

| Function | Purpose | Example |
|----------|---------|---------|
| **SUMIFS** | Sum with conditions | `=SUMIFS(C:C, A:A,"East", B:B,"Widget")` |
| **COUNTIFS** | Count with conditions | `=COUNTIFS(A:A,"East", B:B,"Widget")` |
| **AVERAGEIFS** | Average with conditions | `=AVERAGEIFS(C:C, A:A,"East", B:B,"Widget")` |

### Key Pattern
```
All three follow the same pattern:

=FUNCTION(what_to_calculate, where1, criteria1, where2, criteria2, ...)
          ↑
    SUMIFS: range to sum
    COUNTIFS: N/A (counts the criteria ranges)
    AVERAGEIFS: range to average
```

### Real-World Dashboard Example
```
Data:
     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────┐
1 │ Rep     │ Region  │ Sales   │ Month   │
2 │ Alice   │ East    │ 15000   │ Jan     │
3 │ Bob     │ West    │ 8000    │ Jan     │
4 │ Alice   │ East    │ 18000   │ Feb     │
5 │ Carol   │ East    │ 12000   │ Jan     │
  └─────────┴─────────┴─────────┴─────────┘

Dashboard:
     F                  G
  ┌────────────────┬──────────────────────┐
1 │ East Jan Total:│ =SUMIFS($C$2:$C$5,$B$2:$B$5,"East",$D$2:$D$5,"Jan")
2 │ East Jan Count:│ =COUNTIFS($B$2:$B$5,"East",$D$2:$D$5,"Jan")
3 │ East Jan Avg:  │ =AVERAGEIFS($C$2:$C$5,$B$2:$B$5,"East",$D$2:$D$5,"Jan")
  └────────────────┴──────────────────────┘
                      ↓          ↓          ↓
                    27000        2       13500
```

---

## Advanced Rounding Functions

### CEILING Function

**Syntax:** `=CEILING(number, significance)`

**Purpose:** Rounds **up** to nearest multiple of significance

```
=CEILING(12.3, 1)    → 13   (round up to nearest 1)
=CEILING(12.3, 5)    → 15   (round up to nearest 5)
=CEILING(12.3, 10)   → 20   (round up to nearest 10)
=CEILING(123, 100)   → 200  (round up to nearest 100)
```

### Visual Example
```
Number: 12.3

Significance: 1
  |----|----|----|----|
  11   12   13   14
          ↑    ↑
         12.3  Result: 13

Significance: 5
  |---------|---------|
  10        15        20
          ↑     ↑
         12.3   Result: 15
```

### CEILING Use Cases

**Example 1: Packaging**
```
     A          B          C
  ┌─────────┬─────────┬──────────────────┐
1 │ Items   │ Per Box │ Boxes Needed     │
2 │ 23      │ 10      │ =CEILING(A2/B2,1)│
  └─────────┴─────────┴──────────────────┘
                          ↓
                          3

23 items / 10 per box = 2.3 boxes
Round up: 3 boxes needed
```

**Example 2: Pricing**
```
     A          B
  ┌─────────┬──────────────────┐
1 │ Cost    │ Price (round to $.99)│
2 │ 12.34   │ =CEILING(A2,1)-0.01  │
  └─────────┴──────────────────┘
              ↓
           $12.99

Round up to next dollar, subtract 1 cent
```

### FLOOR Function

**Syntax:** `=FLOOR(number, significance)`

**Purpose:** Rounds **down** to nearest multiple of significance

```
=FLOOR(12.8, 1)    → 12   (round down to nearest 1)
=FLOOR(12.8, 5)    → 10   (round down to nearest 5)
=FLOOR(123, 100)   → 100  (round down to nearest 100)
```

### FLOOR Use Cases

**Example 1: Discounts**
```
     A          B
  ┌─────────┬──────────────────────┐
1 │ Price   │ Discount Price       │
2 │ 19.99   │ =FLOOR(A2*0.9,0.05)  │
  └─────────┴──────────────────────┘
              ↓
           $17.95

90% of 19.99 = 17.991
Round down to nearest nickel: $17.95
```

**Example 2: Time Rounding**
```
     A              B
  ┌────────────┬────────────────────────┐
1 │ Time       │ Round to 15 min        │
2 │ 2:37 PM    │ =FLOOR(A2,"0:15")      │
  └────────────┴────────────────────────┘
                  ↓
               2:30 PM

Rounds down to nearest 15 minutes
```

### MROUND Function

**Syntax:** `=MROUND(number, multiple)`

**Purpose:** Rounds to nearest multiple (up or down)

```
=MROUND(12.3, 5)   → 10   (closest multiple of 5)
=MROUND(13.8, 5)   → 15   (closest multiple of 5)
=MROUND(123, 25)   → 125  (closest multiple of 25)
```

### MROUND Use Cases

**Example: Time Tracking**
```
     A          B
  ┌─────────┬────────────────────────┐
1 │ Hours   │ Billable (round to 0.25)│
2 │ 3.4     │ =MROUND(A2,0.25)       │
3 │ 2.6     │ =MROUND(A3,0.25)       │
  └─────────┴────────────────────────┘
              ↓      ↓
            3.50   2.50

Rounds to nearest quarter hour
```

---

## Statistical Functions

### MEDIAN Function

**Syntax:** `=MEDIAN(number1, [number2], ...)`

**Purpose:** Returns the middle value in a dataset

```
     A
  ┌──────┐
1 │ 10   │
2 │ 20   │
3 │ 30   │ ← Middle value
4 │ 40   │
5 │ 50   │
6 │      │
7 │ =MEDIAN(A1:A5)  → 30
  └──────┘

For even count:
1, 2, 3, 4
Median = (2+3)/2 = 2.5
```

### MEDIAN vs AVERAGE

```
Data: 10, 20, 30, 40, 1000 (one outlier)

AVERAGE: (10+20+30+40+1000)/5 = 220
MEDIAN:  30 (middle value)

Median is better when outliers exist!
```

### MODE Function

**Syntax:** `=MODE.SNGL(number1, [number2], ...)` (Excel 2010+)
**Old syntax:** `=MODE(number1, [number2], ...)`

**Purpose:** Returns the most frequently occurring value

```
     A
  ┌──────┐
1 │ 10   │
2 │ 20   │
3 │ 20   │ ← Most common
4 │ 30   │
5 │ 20   │
6 │      │
7 │ =MODE.SNGL(A1:A5)  → 20
  └──────┘
```

### MODE.MULT Function

**Purpose:** Returns all modes (if multiple values tie for most frequent)

```
Data: 10, 20, 20, 30, 30

MODE.SNGL: Returns 20 (first mode found)
MODE.MULT: Returns both 20 and 30
```

---

## Standard Deviation and Variance

### STDEV.S and STDEV.P Functions

**STDEV.S:** Sample standard deviation (most common)
**STDEV.P:** Population standard deviation

```
     A
  ┌──────┐
1 │ 10   │
2 │ 20   │
3 │ 30   │
4 │ 40   │
5 │ 50   │
6 │      │
7 │ =STDEV.S(A1:A5)  → 15.81
  └──────┘

Measures spread/variability of data
```

### When to Use Which

| Function | Use When |
|----------|----------|
| **STDEV.S** | Analyzing a sample (most common) |
| **STDEV.P** | Analyzing entire population |

### Variance Functions

**VAR.S:** Sample variance
**VAR.P:** Population variance

```
Variance = (Standard Deviation)²

If STDEV = 15.81
Then VAR = 250
```

---

## RANK Functions

### RANK.EQ Function

**Syntax:** `=RANK.EQ(number, ref, [order])`

**Purpose:** Returns the rank of a number in a list

```
     A          B
  ┌─────────┬──────────────────┐
1 │ Score   │ Rank             │
2 │ 95      │ =RANK.EQ(A2,$A$2:$A$5,0) → 1
3 │ 88      │ =RANK.EQ(A3,$A$2:$A$5,0) → 3
4 │ 92      │ =RANK.EQ(A4,$A$2:$A$5,0) → 2
5 │ 88      │ =RANK.EQ(A5,$A$2:$A$5,0) → 3
  └─────────┴──────────────────┘

Order: 0 = descending (highest gets rank 1)
       1 = ascending (lowest gets rank 1)

Note: Ties get same rank, next rank is skipped
```

### RANK.AVG Function

**Purpose:** Ties get average rank

```
     A          B              C
  ┌─────────┬──────────────┬──────────────┐
1 │ Score   │ RANK.EQ      │ RANK.AVG     │
2 │ 95      │ 1            │ 1            │
3 │ 88      │ 3            │ 3.5          │
4 │ 92      │ 2            │ 2            │
5 │ 88      │ 3            │ 3.5          │
  └─────────┴──────────────┴──────────────┘

Two values tied for rank 3-4
Average: (3+4)/2 = 3.5
```

---

## PERCENTILE and QUARTILE Functions

### PERCENTILE.INC Function

**Syntax:** `=PERCENTILE.INC(array, k)`

**Purpose:** Returns the kth percentile (0 to 1)

```
     A
  ┌──────┐
1 │ 10   │
2 │ 20   │
3 │ 30   │
4 │ 40   │
5 │ 50   │
6 │      │
7 │ 50th percentile: =PERCENTILE.INC(A1:A5, 0.5)  → 30
8 │ 75th percentile: =PERCENTILE.INC(A1:A5, 0.75) → 40
9 │ 90th percentile: =PERCENTILE.INC(A1:A5, 0.9)  → 46
  └──────┘
```

### QUARTILE.INC Function

**Syntax:** `=QUARTILE.INC(array, quart)`

**Purpose:** Returns quartiles (0=min, 1=Q1, 2=median, 3=Q3, 4=max)

```
     A
  ┌──────┐
1 │ 10   │
2 │ 20   │
3 │ 30   │
4 │ 40   │
5 │ 50   │
6 │      │
7 │ Q1 (25th): =QUARTILE.INC(A1:A5, 1)  → 20
8 │ Q2 (50th): =QUARTILE.INC(A1:A5, 2)  → 30
9 │ Q3 (75th): =QUARTILE.INC(A1:A5, 3)  → 40
  └──────┘
```

---

## Random Number Functions

### RAND Function

**Syntax:** `=RAND()`

**Purpose:** Returns random decimal between 0 and 1

```
=RAND()  → 0.742...  (random, changes on recalc)
=RAND()  → 0.123...
=RAND()  → 0.918...

Each recalculation generates new number
```

### RANDBETWEEN Function

**Syntax:** `=RANDBETWEEN(bottom, top)`

**Purpose:** Returns random integer between bottom and top

```
=RANDBETWEEN(1, 100)   → Random number 1-100
=RANDBETWEEN(1, 6)     → Random dice roll
=RANDBETWEEN(1, 52)    → Random card number
```

### Use Cases

**Example 1: Sample Selection**
```
     A          B
  ┌─────────┬──────────────────┐
1 │ Name    │ Random           │
2 │ Alice   │ =RAND()          │
3 │ Bob     │ =RAND()          │
4 │ Carol   │ =RAND()          │
  └─────────┴──────────────────┘

Sort by column B to randomize order
```

**Example 2: Test Data**
```
     A          B
  ┌─────────┬──────────────────────────┐
1 │ Row     │ Random Sales             │
2 │ 1       │ =RANDBETWEEN(1000,10000) │
3 │ 2       │ =RANDBETWEEN(1000,10000) │
  └─────────┴──────────────────────────┘

Generate random sales figures for testing
```

### Make Random Numbers Static

**Problem:** Random numbers change every time sheet recalculates

**Solution:**
1. Select cells with formulas
2. Copy (Ctrl+C)
3. Paste Special → Values (Ctrl+Alt+V, V)

This converts formulas to static values

---

## LARGE and SMALL Functions

### LARGE Function

**Syntax:** `=LARGE(array, k)`

**Purpose:** Returns the kth largest value

```
     A
  ┌──────┐
1 │ 95   │
2 │ 88   │
3 │ 92   │
4 │ 78   │
5 │ 85   │
6 │      │
7 │ 1st largest: =LARGE(A1:A5, 1)  → 95
8 │ 2nd largest: =LARGE(A1:A5, 2)  → 92
9 │ 3rd largest: =LARGE(A1:A5, 3)  → 88
  └──────┘
```

### SMALL Function

**Syntax:** `=SMALL(array, k)`

**Purpose:** Returns the kth smallest value

```
=SMALL(A1:A5, 1)  → 78  (smallest)
=SMALL(A1:A5, 2)  → 85  (2nd smallest)
=SMALL(A1:A5, 3)  → 88  (3rd smallest)
```

### Use Case: Top 3 Performers
```
     A          B          C
  ┌─────────┬─────────┬──────────────────┐
1 │ Rep     │ Sales   │ Top 3?           │
2 │ Alice   │ 95000   │ =IF(B2>=LARGE($B$2:$B$6,3),"✓","")
3 │ Bob     │ 88000   │ =IF(B3>=LARGE($B$2:$B$6,3),"✓","")
4 │ Carol   │ 92000   │ =IF(B4>=LARGE($B$2:$B$6,3),"✓","")
5 │ David   │ 78000   │ =IF(B5>=LARGE($B$2:$B$6,3),"✓","")
6 │ Emma    │ 85000   │ =IF(B6>=LARGE($B$2:$B$6,3),"✓","")
  └─────────┴─────────┴──────────────────┘
                          ↓   ↓   ↓
                         ✓   ✓   ✓ (only top 3)

LARGE($B$2:$B$6,3) finds 3rd highest value (88000)
Marks anyone >= 88000
```

---

## AGGREGATE Function

**Syntax:** `=AGGREGATE(function_num, options, array, [k])`

**Purpose:** Performs various calculations with ability to ignore errors and hidden rows

### Function Numbers

| Number | Function | Number | Function |
|--------|----------|--------|----------|
| 1 | AVERAGE | 8 | STDEV.S |
| 2 | COUNT | 9 | SUM |
| 3 | COUNTA | 10 | VAR.S |
| 4 | MAX | 11 | VAR.P |
| 5 | MIN | 12 | MEDIAN |
| 6 | PRODUCT | 13 | MODE.SNGL |
| 7 | STDEV.P | 14-19 | LARGE, SMALL, etc. |

### Options

| Option | Ignores |
|--------|---------|
| 0 or omitted | Nothing |
| 1 | Hidden rows |
| 2 | Error values |
| 3 | Hidden rows and error values |
| 4 | Nothing |
| 5 | Hidden rows |
| 6 | Error values |
| 7 | Hidden rows and error values |

### Example: Ignoring Errors
```
     A
  ┌──────┐
1 │ 10   │
2 │ 20   │
3 │ #DIV/0! ← Error
4 │ 30   │
5 │      │
6 │ Regular SUM: =SUM(A1:A4)       → #DIV/0!
7 │ Ignore errors: =AGGREGATE(9,6,A1:A4) → 60
  └──────┘

Function 9 = SUM
Option 6 = Ignore error values
```

### Real-World Example: Filtered Data
```
     A          B
  ┌─────────┬─────────┐
1 │ Item    │ Value   │
2 │ A       │ 100     │
3 │ B       │ 200     │ ← Hidden row
4 │ C       │ 300     │
5 │         │         │
6 │ With hidden: =SUM(B2:B4)           → 600
7 │ Without hidden: =AGGREGATE(9,5,B2:B4) → 400
  └─────────┴─────────┘

Function 9 = SUM
Option 5 = Ignore hidden rows
```

---

## MAXIFS and MINIFS Functions

**Available in:** Excel 2019, Microsoft 365, Excel Online

### MAXIFS Function

**Syntax:** `=MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Purpose:** Returns maximum value that meets criteria

```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Region  │ Product │ Sales   │
2 │ East    │ Widget  │ 1000    │
3 │ East    │ Widget  │ 1500    │
4 │ West    │ Widget  │ 1200    │
5 │ East    │ Gadget  │ 900     │
  └─────────┴─────────┴─────────┘

Max East Widget sales:
=MAXIFS(C2:C5, A2:A5,"East", B2:B5,"Widget")
→ 1500
```

### MINIFS Function

**Syntax:** `=MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Purpose:** Returns minimum value that meets criteria

```
Min East Widget sales:
=MINIFS(C2:C5, A2:A5,"East", B2:B5,"Widget")
→ 1000
```

### Real-World Example: Price Analysis
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Store   │ Product │ Price   │
2 │ Store A │ Widget  │ 25.00   │
3 │ Store B │ Widget  │ 27.00   │
4 │ Store A │ Widget  │ 24.00   │
5 │ Store C │ Widget  │ 26.00   │
  └─────────┴─────────┴─────────┘

Lowest Widget price:
=MINIFS(C2:C5, B2:B5,"Widget")
→ 24.00

Highest Widget price at Store A:
=MAXIFS(C2:C5, A2:A5,"Store A", B2:B5,"Widget")
→ 25.00
```

---

## POWER and SQRT Functions

### POWER Function

**Syntax:** `=POWER(number, power)`

**Purpose:** Raises number to a power

```
=POWER(2, 3)    → 8   (2³ = 2×2×2)
=POWER(5, 2)    → 25  (5² = 5×5)
=POWER(10, 3)   → 1000 (10³)
=POWER(4, 0.5)  → 2   (square root of 4)

Alternative: =2^3 (same as POWER(2,3))
```

### SQRT Function

**Syntax:** `=SQRT(number)`

**Purpose:** Returns square root

```
=SQRT(16)   → 4
=SQRT(25)   → 5
=SQRT(2)    → 1.414...

Same as: =POWER(number, 0.5)
```

### Use Case: Distance Calculation
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────────┐
1 │ X       │ Y       │ Distance from Origin     │
2 │ 3       │ 4       │ =SQRT(A2^2 + B2^2)       │
  └─────────┴─────────┴──────────────────────────┘
                          ↓
                          5

Pythagorean theorem: distance = √(x² + y²)
```

---

## PRODUCT and QUOTIENT Functions

### PRODUCT Function

**Syntax:** `=PRODUCT(number1, [number2], ...)`

**Purpose:** Multiplies all numbers together

```
     A
  ┌──────┐
1 │ 2    │
2 │ 3    │
3 │ 4    │
4 │      │
5 │ =PRODUCT(A1:A3)  → 24
  └──────┘

2 × 3 × 4 = 24

Same as: =A1*A2*A3
```

### QUOTIENT Function

**Syntax:** `=QUOTIENT(numerator, denominator)`

**Purpose:** Returns integer portion of division (no remainder)

```
=QUOTIENT(10, 3)   → 3   (10÷3 = 3 remainder 1)
=QUOTIENT(17, 5)   → 3   (17÷5 = 3 remainder 2)
=QUOTIENT(20, 4)   → 5   (20÷4 = 5 remainder 0)

Compare to MOD (returns remainder):
=MOD(10, 3)        → 1
```

---

## GCD and LCM Functions

### GCD Function

**Syntax:** `=GCD(number1, [number2], ...)`

**Purpose:** Returns Greatest Common Divisor

```
=GCD(12, 18)   → 6
=GCD(15, 25)   → 5
=GCD(8, 12, 16) → 4

Largest number that divides evenly into all inputs
```

### LCM Function

**Syntax:** `=LCM(number1, [number2], ...)`

**Purpose:** Returns Least Common Multiple

```
=LCM(4, 6)     → 12
=LCM(3, 5)     → 15
=LCM(2, 3, 4)  → 12

Smallest number that all inputs divide into evenly
```

---

## SIGN and ABS Functions

### SIGN Function

**Syntax:** `=SIGN(number)`

**Purpose:** Returns 1 (positive), -1 (negative), or 0 (zero)

```
=SIGN(10)    → 1
=SIGN(-10)   → -1
=SIGN(0)     → 0
```

### Use Case: Profit/Loss Indicator
```
     A          B
  ┌─────────┬──────────────────────────────┐
1 │ Amount  │ Status                       │
2 │ 1000    │ =CHOOSE(SIGN(A2)+2,"Loss","Break Even","Profit")
3 │ -500    │ =CHOOSE(SIGN(A3)+2,"Loss","Break Even","Profit")
4 │ 0       │ =CHOOSE(SIGN(A4)+2,"Loss","Break Even","Profit")
  └─────────┴──────────────────────────────┘
              ↓         ↓            ↓
           "Profit"  "Loss"   "Break Even"

SIGN returns: 1, -1, or 0
Add 2: 3, 1, or 2
CHOOSE selects based on position
```

---

## Common Patterns and Use Cases

### Pattern 1: Conditional Statistics
```
Average of positive values only:
=AVERAGEIF(A:A, ">0")

Count non-zero values:
=COUNTIF(A:A, "<>0")

Sum top 10 values:
=SUMPRODUCT(LARGE(A:A,ROW(1:10)))
```

### Pattern 2: Weighted Average
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Score   │ Weight  │         │
2 │ 90      │ 40%     │         │
3 │ 85      │ 30%     │         │
4 │ 95      │ 30%     │         │
5 │         │         │         │
6 │ Weighted Avg: =SUMPRODUCT(A2:A4,B2:B4)/SUM(B2:B4)
  └─────────┴─────────┴─────────┘
                ↓
              89.5

(90×0.4) + (85×0.3) + (95×0.3) = 89.5
```

### Pattern 3: Running Total
```
     A          B
  ┌─────────┬──────────────────┐
1 │ Value   │ Running Total    │
2 │ 10      │ =SUM($A$2:A2)    │
3 │ 20      │ =SUM($A$2:A3)    │
4 │ 30      │ =SUM($A$2:A4)    │
  └─────────┴──────────────────┘
              ↓   ↓    ↓
             10  30   60

Fixed start ($A$2), expanding end (A2, A3, A4)
```

### Pattern 4: Percent of Total
```
     A          B
  ┌─────────┬──────────────────┐
1 │ Value   │ % of Total       │
2 │ 10      │ =A2/SUM($A$2:$A$4)
3 │ 20      │ =A3/SUM($A$2:$A$4)
4 │ 30      │ =A4/SUM($A$2:$A$4)
  └─────────┴──────────────────┘
              ↓      ↓      ↓
            16.7%  33.3%  50%

Format as percentage
```

### Pattern 5: Growth Rate
```
     A          B          C
  ┌─────────┬─────────┬──────────────────┐
1 │ Period  │ Sales   │ Growth %         │
2 │ Q1      │ 10000   │ -                │
3 │ Q2      │ 12000   │ =(B3-B2)/B2      │
4 │ Q3      │ 13500   │ =(B4-B3)/B3      │
  └─────────┴─────────┴──────────────────┘
                          ↓        ↓
                         20%     12.5%

(New - Old) / Old
```

---

## Real-World Application: Sales Dashboard

Let's build a comprehensive sales analysis dashboard.

### Data Setup
```
     A          B          C          D          E
  ┌─────────┬─────────┬─────────┬─────────┬─────────┐
1 │ Rep     │ Region  │ Product │ Sales   │ Quarter │
2 │ Alice   │ East    │ Widget  │ 15000   │ Q1      │
3 │ Bob     │ West    │ Gadget  │ 8000    │ Q1      │
4 │ Alice   │ East    │ Widget  │ 18000   │ Q2      │
5 │ Carol   │ East    │ Gadget  │ 12000   │ Q1      │
6 │ Bob     │ West    │ Widget  │ 9000    │ Q2      │
7 │ Alice   │ East    │ Widget  │ 16000   │ Q1      │
  └─────────┴─────────┴─────────┴─────────┴─────────┘
```

### Dashboard Formulas

**Total Sales:**
```
=SUM(D2:D7)
```

**Average Sale:**
```
=AVERAGE(D2:D7)
```

**East Region Total:**
```
=SUMIFS(D2:D7, B2:B7, "East")
```

**Alice's Widget Sales in Q1:**
```
=SUMIFS(D2:D7, A2:A7,"Alice", C2:C7,"Widget", E2:E7,"Q1")
```

**Top Performer:**
```
=INDEX(A2:A7, MATCH(MAX(D2:D7), D2:D7, 0))
```

**Number of Sales Over 10000:**
```
=COUNTIF(D2:D7, ">10000")
```

**Median Sale Amount:**
```
=MEDIAN(D2:D7)
```

**Sales Standard Deviation:**
```
=STDEV.S(D2:D7)
```

**Alice's Rank:**
```
=RANK.EQ(SUMIF(A2:A7,"Alice",D2:D7), 
         {SUMIF(A2:A7,"Alice",D2:D7),
          SUMIF(A2:A7,"Bob",D2:D7),
          SUMIF(A2:A7,"Carol",D2:D7)}, 0)
```

### Complete Dashboard
```
     G                      H
  ┌────────────────────┬──────────┐
1 │ Metric             │ Value    │
2 │ Total Sales        │ 78,000   │
3 │ Average Sale       │ 13,000   │
4 │ Median Sale        │ 13,500   │
5 │ East Region Total  │ 61,000   │
6 │ West Region Total  │ 17,000   │
7 │ Sales > 10K        │ 5        │
8 │ Top Performer      │ Alice    │
9 │ Std Deviation      │ 3,674    │
10│ Alice Total        │ 49,000   │
11│ Bob Total          │ 17,000   │
12│ Carol Total        │ 12,000   │
  └────────────────────┴──────────┘
```

---

## Common Mistakes and Best Practices

### Mistake 1: Forgetting Absolute References in SUMIFS
```
❌ Wrong: =SUMIFS(C2:C10, A2:A10, "East")
Copy down: =SUMIFS(C3:C11, A3:A11, "East")  ← Range moves!

✅ Right: =SUMIFS($C$2:$C$10, $A$2:$A$10, "East")
Copy down: =SUMIFS($C$2:$C$10, $A$2:$A$10, "East")  ← Fixed!
```

### Mistake 2: Using AVERAGE with Blanks
```
Data: 10, 20, (blank), 30

=SUM(A1:A3)/3        → 20  (incorrect, includes blank)
=AVERAGE(A1:A3)      → 20  (correct, ignores blank)

AVERAGE automatically ignores blank cells
```

### Mistake 3: RANK Not Updating
```
Problem: RANK references not absolute

❌ =RANK.EQ(A2, A2:A10, 0)
Copy down: =RANK.EQ(A3, A3:A11, 0)  ← Wrong range!

✅ =RANK.EQ(A2, $A$2:$A$10, 0)
Copy down: =RANK.EQ(A3, $A$2:$A$10, 0)  ← Correct!
```

### Mistake 4: Wrong Rounding Function
```
Want to round 12.8 to nearest 5:

❌ =ROUND(12.8, 5)     → Wrong (rounds to 5 decimals)
✅ =MROUND(12.8, 5)    → Correct (rounds to multiple of 5)
```

### Mistake 5: Criteria as Numbers vs Text
```
In SUMIFS/COUNTIFS:

✅ =SUMIFS(C:C, A:A, 100)       Works (number)
✅ =SUMIFS(C:C, A:A, ">100")    Works (comparison)
❌ =SUMIFS(C:C, A:A, >100)      Error (no quotes)
```

---

## Best Practices

### 1. Use Absolute References for Criteria Ranges
```
Always use $ for ranges in IFS functions:
=SUMIFS($C$2:$C$100, $A$2:$A$100, "East")
```

### 2. Name Important Ranges
```
Instead of: =SUMIFS(C:C, A:A, "East")
Better:     =SUMIFS(SalesAmount, Region, "East")

Easier to read and maintain
```

### 3. Break Complex Formulas into Steps
```
❌ Complex:
=SUMIFS(D:D,A:A,"East",B:B,"Widget")/COUNTIFS(A:A,"East",B:B,"Widget")

✅ Better:
E1: =SUMIFS(D:D,A:A,"East",B:B,"Widget")    [Total]
F1: =COUNTIFS(A:A,"East",B:B,"Widget")       [Count]
G1: =E1/F1                                    [Average]
```

### 4. Document Criteria
```
Create a criteria table:
     F          G
  ┌─────────┬─────────┐
1 │ Region: │ East    │
2 │ Product:│ Widget  │
  └─────────┴─────────┘

Then reference:
=SUMIFS(Sales, Region, G1, Product, G2)

Easier to change criteria
```

### 5. Use Helper Columns for Multiple Criteria
```
If checking same criteria multiple times:

Add column: =A2&"-"&B2  (concatenate criteria)
Then use: =SUMIF(helper_column, "East-Widget", values)

Faster than multiple SUMIFS
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- SUMIFS, COUNTIFS, AVERAGEIFS syntax
- All check criteria before summing/counting/averaging
- Criteria use quotes for text and comparisons
- CEILING rounds up, FLOOR rounds down, MROUND rounds to nearest
- MEDIAN is middle value (better than AVERAGE for outliers)
- RANK requires absolute references for the list
- RAND() and RANDBETWEEN() are volatile
- LARGE(array, k) = kth largest, SMALL(array, k) = kth smallest

### Practice Deeply
- Using SUMIFS/COUNTIFS/AVERAGEIFS with multiple criteria
- Building dashboards with IFS functions
- Using CEILING/FLOOR/MROUND for business rules
- Calculating MEDIAN and STDEV for analysis
- Ranking data with RANK functions
- Finding top/bottom N values with LARGE/SMALL
- Using PERCENTILE and QUARTILE for distributions
- Creating weighted averages
- Building running totals
- Calculating growth rates
- Combining statistical functions with logical functions
- Using AGGREGATE to ignore errors

### Don't Memorize
- Every AGGREGATE function number (look up when needed)
- All statistical formulas (Excel does the math)
- Exact variance/standard deviation calculations
- Every possible criteria combination (build as needed)

---

## Quick Reference: Key Functions

### Multi-Criteria
```
=SUMIFS(sum_range, criteria_range1, criteria1, ...)
=COUNTIFS(criteria_range1, criteria1, ...)
=AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)
=MAXIFS(max_range, criteria_range1, criteria1, ...)
=MINIFS(min_range, criteria_range1, criteria1, ...)
```

### Rounding
```
=CEILING(number, significance)    Round up
=FLOOR(number, significance)      Round down
=MROUND(number, multiple)         Round to nearest
```

### Statistical
```
=MEDIAN(range)           Middle value
=MODE.SNGL(range)        Most common
=STDEV.S(range)          Standard deviation
=VAR.S(range)            Variance
```

### Ranking
```
=RANK.EQ(number, ref, order)
=PERCENTILE.INC(array, k)
=QUARTILE.INC(array, quart)
=LARGE(array, k)         kth largest
=SMALL(array, k)         kth smallest
```

### Random
```
=RAND()                  0 to 1
=RANDBETWEEN(low, high)  Integer between
```

---

## Next Step

After mastering mathematical and statistical functions, you're ready to explore:

**`10-data-validation.md`**
- Creating dropdown lists
- Setting numeric and date constraints
- Custom validation rules
- Input messages and error alerts
- Dynamic validation lists
- Dependent dropdowns
- Data validation best practices
