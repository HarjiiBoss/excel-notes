# Conditional Formatting

This file covers conditional formatting in Excel - how to automatically format cells based on their values or formulas to highlight important data, identify trends, and create visual dashboards.

---

## What is Conditional Formatting?

**Conditional formatting** automatically applies formatting (colors, icons, data bars) to cells based on rules you define.

### Purpose
- **Highlight important values** (high sales, low inventory)
- **Identify trends** (increases, decreases)
- **Spot outliers** (values above/below threshold)
- **Create visual dashboards** (heat maps, progress bars)
- **Make data easier to scan** (alternating colors, duplicates)

### Where to Find It
**Home Tab → Conditional Formatting**

---

## Visual Example: Before and After

**Before (Plain Data):**
```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │ Name   │ Jan    │ Feb    │ Mar    │
  ├────────┼────────┼────────┼────────┤
2 │ John   │ 5000   │ 8000   │ 4500   │
  ├────────┼────────┼────────┼────────┤
3 │ Sarah  │ 9000   │ 12000  │ 11000  │
  ├────────┼────────┼────────┼────────┤
4 │ Mike   │ 3000   │ 3500   │ 2800   │
  └────────┴────────┴────────┴────────┘
```

**After (With Conditional Formatting):**
```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │ Name   │ Jan    │ Feb    │ Mar    │
  ├────────┼────────┼────────┼────────┤
2 │ John   │ 5000   │ [🟢8000]│ 4500   │ ← Highest value in green
  ├────────┼────────┼────────┼────────┤
3 │ Sarah  │ 9000   │[🟢12000]│[🟢11000]│ ← Top values highlighted
  ├────────┼────────┼────────┼────────┤
4 │ Mike   │[🔴3000]│ 3500   │[🔴2800]│ ← Lowest values in red
  └────────┴────────┴────────┴────────┘
```

---

## Types of Conditional Formatting

Excel offers several built-in conditional formatting options:

### Quick Overview

| Type | Purpose | Example Use |
|------|---------|-------------|
| **Highlight Cells Rules** | Mark cells meeting criteria | Values > 100 |
| **Top/Bottom Rules** | Highlight extremes | Top 10 items |
| **Data Bars** | Show value as bar | Sales progress |
| **Color Scales** | Gradient by value | Heat maps |
| **Icon Sets** | Show with symbols | Trend arrows |
| **Custom Formula** | Advanced conditions | Complex logic |

---

## Highlight Cells Rules

Mark cells that meet specific criteria.

### Available Rules

**Home → Conditional Formatting → Highlight Cells Rules:**
- Greater Than
- Less Than
- Between
- Equal To
- Text that Contains
- A Date Occurring
- Duplicate Values

### Example 1: Highlight Values Greater Than

**Scenario:** Highlight sales above $7,000

```
     A         B
  ┌────────┬────────┐
1 │ Name   │ Sales  │
  ├────────┼────────┤
2 │ John   │ 5000   │
  ├────────┼────────┤
3 │ Sarah  │ 9000   │ ← Highlighted (>7000)
  ├────────┼────────┤
4 │ Mike   │ 3500   │
  ├────────┼────────┤
5 │ Lisa   │ 12000  │ ← Highlighted (>7000)
  └────────┴────────┘
```

**Steps:**
1. Select range B2:B5
2. Home → Conditional Formatting → Highlight Cells Rules → Greater Than
3. Enter: `7000`
4. Choose format: Light Red Fill
5. Click OK

### Example 2: Highlight Text Containing

**Scenario:** Highlight products containing "Pro"

```
     A              B
  ┌────────────┬─────────┐
1 │ Product    │ Price   │
  ├────────────┼─────────┤
2 │ Widget     │ 25      │
  ├────────────┼─────────┤
3 │ Widget Pro │ 50      │ ← Highlighted (contains "Pro")
  ├────────────┼─────────┤
4 │ Gadget     │ 40      │
  ├────────────┼─────────┤
5 │ Pro Series │ 75      │ ← Highlighted (contains "Pro")
  └────────────┴─────────┘
```

**Steps:**
1. Select range A2:A5
2. Conditional Formatting → Highlight Cells Rules → Text that Contains
3. Enter: `Pro`
4. Choose format
5. Click OK

### Example 3: Highlight Duplicates

**Scenario:** Find duplicate entries

```
     A
  ┌─────────┐
1 │ Email   │
  ├─────────┤
2 │ a@x.com │
  ├─────────┤
3 │ b@x.com │
  ├─────────┤
4 │ a@x.com │ ← Highlighted (duplicate)
  ├─────────┤
5 │ c@x.com │
  └─────────┘
```

**Steps:**
1. Select range A2:A5
2. Conditional Formatting → Highlight Cells Rules → Duplicate Values
3. Choose format
4. Click OK

**Options:**
- Highlight Duplicate values
- Highlight Unique values

---

## Top/Bottom Rules

Highlight the highest or lowest values in a range.

### Available Rules

**Home → Conditional Formatting → Top/Bottom Rules:**
- Top 10 Items
- Top 10%
- Bottom 10 Items
- Bottom 10%
- Above Average
- Below Average

### Example 1: Top 3 Sales

```
     A         B
  ┌────────┬────────┐
1 │ Name   │ Sales  │
  ├────────┼────────┤
2 │ John   │ 5000   │
  ├────────┼────────┤
3 │ Sarah  │ 12000  │ ← Highlighted (Top 3)
  ├────────┼────────┤
4 │ Mike   │ 3500   │
  ├────────┼────────┤
5 │ Lisa   │ 9000   │ ← Highlighted (Top 3)
  ├────────┼────────┤
6 │ Tom    │ 8000   │ ← Highlighted (Top 3)
  ├────────┼────────┤
7 │ Emma   │ 4500   │
  └────────┴────────┘
```

**Steps:**
1. Select range B2:B7
2. Conditional Formatting → Top/Bottom Rules → Top 10 Items
3. Change `10` to `3`
4. Choose format
5. Click OK

### Example 2: Below Average

```
     A          B
  ┌─────────┬─────────┐
1 │ Student │ Score   │
  ├─────────┼─────────┤
2 │ Ann     │ 85      │
  ├─────────┼─────────┤
3 │ Bob     │ 72      │ ← Below average (78)
  ├─────────┼─────────┤
4 │ Carl    │ 90      │
  ├─────────┼─────────┤
5 │ Dana    │ 68      │ ← Below average (78)
  └─────────┴─────────┘

Average: (85+72+90+68)/4 = 78.75
```

**Steps:**
1. Select range B2:B5
2. Conditional Formatting → Top/Bottom Rules → Below Average
3. Choose format
4. Click OK

---

## Data Bars

Show values as horizontal bars within cells.

### Visual Example

```
     A              B
  ┌────────────┬──────────────────────┐
1 │ Product    │ Sales                │
  ├────────────┼──────────────────────┤
2 │ Widget     │ 5000  [████░░░░░░]   │
  ├────────────┼──────────────────────┤
3 │ Gadget     │ 8000  [███████░░░]   │
  ├────────────┼──────────────────────┤
4 │ Tool       │ 3000  [██░░░░░░░░]   │
  ├────────────┼──────────────────────┤
5 │ Device     │ 10000 [██████████]   │
  └────────────┴──────────────────────┘

Longest bar = highest value
```

### How to Apply

**Steps:**
1. Select range B2:B5
2. Conditional Formatting → Data Bars
3. Choose color (Gradient or Solid)

### Data Bar Options

**Right-click cell → Manage Rules → Edit Rule:**

**Bar Direction:**
- Left to Right (default)
- Right to Left

**Bar Appearance:**
- Gradient Fill
- Solid Fill
- Border (with or without)

**Value Display:**
- Show bar only (hide number)
- Show bar and number (default)

### Example: Show Bar Only

```
     A              B
  ┌────────────┬──────────────────┐
1 │ Month      │ Revenue          │
  ├────────────┼──────────────────┤
2 │ January    │ [████░░░░░░░░]   │ ← Number hidden
  ├────────────┼──────────────────┤
3 │ February   │ [████████░░░░]   │
  ├────────────┼──────────────────┤
4 │ March      │ [██████████░░]   │
  └────────────┴──────────────────┘
```

**To hide numbers:**
1. Manage Rules → Edit Rule
2. Check "Show Bar Only"
3. Click OK

---

## Color Scales

Apply gradient colors based on cell values.

### Types of Color Scales

**2-Color Scale:**
- Low values → One color
- High values → Another color
- Gradient in between

**3-Color Scale:**
- Low values → Color 1
- Middle values → Color 2
- High values → Color 3

### Visual Example: 3-Color Scale

```
     A         B         C         D
  ┌────────┬────────┬────────┬────────┐
1 │ Region │ Q1     │ Q2     │ Q3     │
  ├────────┼────────┼────────┼────────┤
2 │ East   │ [🔴50] │ [🟡75] │ [🟢95] │
  ├────────┼────────┼────────┼────────┤
3 │ West   │ [🟡70] │ [🟢90] │ [🟢98] │
  ├────────┼────────┼────────┼────────┤
4 │ North  │ [🔴45] │ [🔴55] │ [🟡72] │
  └────────┴────────┴────────┴────────┘

Red = Low, Yellow = Medium, Green = High
```

### How to Apply

**Steps:**
1. Select range B2:D4
2. Conditional Formatting → Color Scales
3. Choose preset (e.g., Red-Yellow-Green)

### Common Color Scale Schemes

| Scheme | Best For |
|--------|----------|
| **Red-Yellow-Green** | Performance (bad to good) |
| **Red-White-Blue** | Temperature data |
| **Green-Yellow-Red** | Financial (profit/loss) |
| **Blue-White-Red** | Deviation from center |

### Customizing Color Scales

**Manage Rules → Edit Rule:**

**Set minimum, midpoint, maximum:**
- Lowest Value
- Highest Value
- Number
- Percent
- Percentile
- Formula

**Example: Set specific thresholds**
```
Minimum: 0 (Red)
Midpoint: 50 (Yellow)
Maximum: 100 (Green)
```

---

## Icon Sets

Display icons based on value ranges.

### Available Icon Sets

**Directional:**
- Arrows (3, 4, or 5 arrows)
- Triangles

**Shapes:**
- Traffic lights
- Flags
- Symbols

**Indicators:**
- Stars
- Ratings
- Boxes

### Visual Example: 3 Arrows

```
     A              B                C
  ┌────────────┬──────────────┬──────────┐
1 │ Product    │ Growth %     │ Icon     │
  ├────────────┼──────────────┼──────────┤
2 │ Widget     │ 15%          │ ↑        │ ← Up arrow
  ├────────────┼──────────────┼──────────┤
3 │ Gadget     │ -5%          │ ↓        │ ← Down arrow
  ├────────────┼──────────────┼──────────┤
4 │ Tool       │ 2%           │ →        │ ← Right arrow
  └────────────┴──────────────┴──────────┘
```

### How to Apply

**Steps:**
1. Select range B2:B4
2. Conditional Formatting → Icon Sets
3. Choose icon type

### Icon Set Rules

**By default, icons split into thirds:**
- Top 33% → Green/Up
- Middle 33% → Yellow/Right
- Bottom 33% → Red/Down

### Customizing Icon Rules

**Manage Rules → Edit Rule:**

**Example: Custom thresholds**
```
↑ Green:  When value >= 10
→ Yellow: When value >= 0 and < 10
↓ Red:    When value < 0
```

### Show Icons Only (Hide Numbers)

**Edit Rule → Show Icon Only:**
```
     A              B
  ┌────────────┬──────────┐
1 │ Status     │ Icon     │
  ├────────────┼──────────┤
2 │ Project A  │ ●        │ ← Green circle only
  ├────────────┼──────────┤
3 │ Project B  │ ●        │ ← Yellow circle only
  ├────────────┼──────────┤
4 │ Project C  │ ●        │ ← Red circle only
  └────────────┴──────────┘
```

---

## Custom Formula Rules

Create advanced conditional formatting using formulas.

### When to Use Custom Formulas

- Highlight entire rows based on one column
- Compare cells across columns
- Complex logic with AND/OR
- Reference other sheets
- Use functions (IF, COUNTIF, etc.)

### Important Rule

**Formula must return TRUE or FALSE:**
```
✅ =A2>100          Returns TRUE or FALSE
✅ =B2="Pending"    Returns TRUE or FALSE
✅ =AND(A2>0,B2>0)  Returns TRUE or FALSE

❌ =A2+B2           Returns number, not TRUE/FALSE
```

### Example 1: Highlight Entire Row

**Scenario:** Highlight entire row where Status = "Overdue"

```
     A         B            C         D
  ┌────────┬──────────┬─────────┬─────────┐
1 │ Task   │ Status   │ Owner   │ Days    │
  ├────────┼──────────┼─────────┼─────────┤
2 │ Task 1 │ Complete │ John    │ 5       │
  ├────────┼──────────┼─────────┼─────────┤
3 │[Task 2]│[Overdue] │[Sarah]  │[12]     │ ← Entire row highlighted
  ├────────┼──────────┼─────────┼─────────┤
4 │ Task 3 │ Pending  │ Mike    │ 3       │
  └────────┴──────────┴─────────┴─────────┘
```

**Steps:**
1. Select range A2:D4 (entire data range)
2. Conditional Formatting → New Rule
3. Select "Use a formula to determine which cells to format"
4. Enter formula: `=$B2="Overdue"`
5. Click Format → Choose red fill
6. Click OK

**Key Point:** Use `$B2` (mixed reference)
- `$B` locks column B
- `2` changes with each row (2, 3, 4...)

### Example 2: Alternate Row Colors

**Scenario:** Create striped rows for readability

```
     A         B         C
  ┌────────┬────────┬────────┐
1 │ Name   │ Age    │ City   │
  ├────────┼────────┼────────┤
2 │ [John] │ [25]   │ [NYC]  │ ← Light gray
  ├────────┼────────┼────────┤
3 │ Sarah  │ 30     │ LA     │ ← White
  ├────────┼────────┼────────┤
4 │ [Mike] │ [28]   │ [SF]   │ ← Light gray
  ├────────┼────────┼────────┤
5 │ Lisa   │ 32     │ Boston │ ← White
  └────────┴────────┴────────┘
```

**Formula:** `=MOD(ROW(),2)=0`

**Explanation:**
- `ROW()` returns row number
- `MOD(ROW(),2)` returns 0 for even rows, 1 for odd
- `=0` checks if even
- Even rows get formatted

### Example 3: Highlight Duplicates in Column

**Scenario:** Highlight duplicate emails

```
     A
  ┌─────────────┐
1 │ Email       │
  ├─────────────┤
2 │ a@test.com  │
  ├─────────────┤
3 │ [b@test.com]│ ← Highlighted (appears twice)
  ├─────────────┤
4 │ c@test.com  │
  ├─────────────┤
5 │ [b@test.com]│ ← Highlighted (duplicate)
  └─────────────┘
```

**Formula:** `=COUNTIF($A$2:$A$5,A2)>1`

**Explanation:**
- `COUNTIF($A$2:$A$5,A2)` counts how many times A2 appears
- `>1` means it appears more than once
- Absolute range `$A$2:$A$5` doesn't change
- Relative `A2` changes for each row

### Example 4: Highlight Row Based on Multiple Conditions

**Scenario:** Highlight orders over $500 AND status is "Pending"

```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Order   │ Amount  │ Status  │
  ├─────────┼─────────┼─────────┤
2 │ 1001    │ 250     │ Pending │
  ├─────────┼─────────┼─────────┤
3 │ [1002]  │ [750]   │[Pending]│ ← Highlighted (>500 AND Pending)
  ├─────────┼─────────┼─────────┤
4 │ 1003    │ 600     │ Complete│
  └─────────┴─────────┴─────────┘
```

**Formula:** `=AND($B2>500,$C2="Pending")`

### Example 5: Highlight Weekends

**Scenario:** Highlight weekend dates

```
     A            B
  ┌──────────┬─────────┐
1 │ Date     │ Sales   │
  ├──────────┼─────────┤
2 │ 1/1/2024 │ 1000    │
  ├──────────┼─────────┤
3 │ [1/6/24] │ [800]   │ ← Saturday (highlighted)
  ├──────────┼─────────┤
4 │ [1/7/24] │ [750]   │ ← Sunday (highlighted)
  └──────────┴─────────┘
```

**Formula:** `=OR(WEEKDAY($A2)=1,WEEKDAY($A2)=7)`

**Explanation:**
- `WEEKDAY()` returns 1 for Sunday, 7 for Saturday
- `OR()` returns TRUE if either condition is met

### Example 6: Compare Columns

**Scenario:** Highlight where Actual < Budget

```
     A          B         C
  ┌─────────┬─────────┬─────────┐
1 │ Item    │ Budget  │ Actual  │
  ├─────────┼─────────┼─────────┤
2 │ Rent    │ 1500    │ 1500    │
  ├─────────┼─────────┼─────────┤
3 │ Food    │ 500     │ [550]   │ ← Highlighted (over budget)
  ├─────────┼─────────┼─────────┤
4 │ Gas     │ 200     │ 180     │
  └─────────┴─────────┴─────────┘
```

**Formula (apply to column C):** `=$C2>$B2`

---

## Managing Conditional Formatting Rules

### View All Rules

**Home → Conditional Formatting → Manage Rules**

Shows:
- All rules applied to selection
- Rule order (top to bottom priority)
- What each rule does
- Where it applies

### Rule Priority

**Rules are evaluated top to bottom:**
```
Rule 1: If value > 100, green
Rule 2: If value > 50, yellow
Rule 3: If value > 0, red
```

**Cell with value 150:**
- Rule 1 applies (green)
- Rules 2 & 3 ignored (unless "Stop If True" is unchecked)

**To change order:**
- Select rule
- Click arrows to move up/down

### Edit Existing Rule

**Steps:**
1. Manage Rules
2. Select rule
3. Click "Edit Rule"
4. Make changes
5. Click OK

### Delete Rule

**Steps:**
1. Manage Rules
2. Select rule
3. Click "Delete Rule"
4. Click OK

### Clear All Formatting

**Clear from cells:**
- Select cells
- Conditional Formatting → Clear Rules → Clear Rules from Selected Cells

**Clear from entire sheet:**
- Conditional Formatting → Clear Rules → Clear Rules from Entire Sheet

---

## Stop If True Option

Controls whether Excel continues evaluating rules after one applies.

### How It Works

**With "Stop If True" checked:**
```
Rule 1: Value > 100 → Green [Stop If True: ✓]
Rule 2: Value > 50  → Yellow
Rule 3: Value > 0   → Red

Value = 150:
✓ Rule 1 applies (green)
✗ Rule 2 not evaluated
✗ Rule 3 not evaluated
```

**With "Stop If True" unchecked:**
```
Rule 1: Value > 100 → Green [Stop If True: ☐]
Rule 2: Value > 50  → Yellow
Rule 3: Value > 0   → Red

Value = 150:
✓ Rule 1 applies (green)
✓ Rule 2 also applies (yellow)
✓ Rule 3 also applies (red)

Result: Multiple formats may apply
```

### When to Use

**Check "Stop If True" when:**
- Rules are mutually exclusive
- First match is final answer
- Performance optimization needed

**Uncheck "Stop If True" when:**
- Want multiple formats to combine
- Layering effects (e.g., bold + color)

---

## Real-World Example: Sales Dashboard

**Create a visual sales dashboard:**

```
     A         B         C            D
  ┌────────┬────────┬──────────┬──────────────┐
1 │ Rep    │ Sales  │ Quota    │ % of Quota   │
  ├────────┼────────┼──────────┼──────────────┤
2 │ John   │ 45000  │ 50000    │ 90% [███████░]│
  ├────────┼────────┼──────────┼──────────────┤
3 │ Sarah  │ 62000  │ 50000    │ 124% [████████]│ ← Green (>100%)
  ├────────┼────────┼──────────┼──────────────┤
4 │ Mike   │ 38000  │ 50000    │ 76% [█████░░░]│ ← Red (<80%)
  ├────────┼────────┼──────────┼──────────────┤
5 │ Lisa   │ 51000  │ 50000    │ 102% [████████]│ ← Green (>100%)
  └────────┴────────┴──────────┴──────────────┘
```

**Applied Formatting:**

**Column B (Sales):**
- Data bars to show relative sales

**Column D (% of Quota):**
- Green fill if >= 100%
- Yellow fill if >= 80% and < 100%
- Red fill if < 80%

**Formula in D2:** `=B2/C2`

**Conditional Formatting on D2:D5:**
```
Rule 1: =$D2>=1    → Green fill
Rule 2: =$D2>=0.8  → Yellow fill
Rule 3: =$D2<0.8   → Red fill
```

---

## Real-World Example: Project Status Tracker

**Track project milestones with visual indicators:**

```
     A            B          C          D
  ┌──────────┬──────────┬──────────┬──────────┐
1 │ Task     │ Status   │ Due Date │ Days Left│
  ├──────────┼──────────┼──────────┼──────────┤
2 │ Design   │ Complete │ 12/1/24  │ --       │ ● Green
  ├──────────┼──────────┼──────────┼──────────┤
3 │ Build    │ Active   │ 12/20/24 │ 15       │ ● Yellow
  ├──────────┼──────────┼──────────┼──────────┤
4 │ Test     │ Overdue  │ 12/15/24 │ -5       │ ● Red
  ├──────────┼──────────┼──────────┼──────────┤
5 │ Deploy   │ Pending  │ 1/5/25   │ 30       │ ● Gray
  └──────────┴──────────┴──────────┴──────────┘
```

**Applied Formatting:**

**Column B (Status):**
```
Formula: =$B2="Complete"   → Green
Formula: =$B2="Active"     → Yellow
Formula: =$B2="Overdue"    → Red
Formula: =$B2="Pending"    → Gray
```

**Column D (Days Left):**
```
Icon Set: 3 Traffic Lights
> 14 days  → Green
7-14 days  → Yellow
< 7 days   → Red
```

**Entire Row formatting:**
```
Formula: =$B2="Overdue"    → Light red fill for entire row
```

---

## Real-World Example: Expense Report

**Highlight over-budget items:**

```
     A            B         C         D
  ┌──────────┬─────────┬─────────┬─────────┐
1 │ Category │ Budget  │ Actual  │ Variance│
  ├──────────┼─────────┼─────────┼─────────┤
2 │ Travel   │ 5000    │ 4800    │ 200     │
  ├──────────┼─────────┼─────────┼─────────┤
3 │ [Food]   │ [1000]  │ [1200]  │ [-200]  │ ← Red (over budget)
  ├──────────┼─────────┼─────────┼─────────┤
4 │ Supplies │ 800     │ 650     │ 150     │
  ├──────────┼─────────┼─────────┼─────────┤
5 │ [Other]  │ [500]   │ [750]   │ [-250]  │ ← Red (over budget)
  └──────────┴─────────┴─────────┴─────────┘
```

**Formulas:**
- D2: `=B2-C2` (positive = under budget)

**Conditional Formatting:**
```
Highlight row if: =$D2<0 (Actual > Budget)
Format: Light red fill
```

**Data Bars in Column C:**
- Shows relative spending across categories

---

## Common Mistakes

### Mistake 1: Wrong Cell Reference in Formula

```
❌ Wrong: Selected A2:D10, formula uses A1
Result: First row doesn't format correctly

✅ Right: Formula references first row of selection
If selection starts at row 2, use $B2 not $B1
```

### Mistake 2: Absolute vs Relative References

```
❌ Wrong: =$B$2>100
Result: Only checks row 2 for all rows

✅ Right: =$B2>100
Column locked ($B), row relative (2)
```

### Mistake 3: Formula Doesn't Return TRUE/FALSE

```
❌ Wrong: =A2+B2
Result: Error or unexpected behavior

✅ Right: =A2+B2>100
Returns TRUE or FALSE
```

### Mistake 4: Formatting Not Showing

**Possible causes:**
- Cell already has manual formatting (overrides conditional)
- Rule applies to wrong range
- Rule order priority issue
- "Stop If True" blocking subsequent rules

**Solution:** Clear manual formatting first

### Mistake 5: Overlapping Rules Conflict

```
❌ Problem:
Rule 1: Value > 100 → Green
Rule 2: Value < 50 → Red
Value = 75 → No formatting (doesn't meet either)

✅ Solution: Add rule for middle range
Rule 3: Value >= 50 AND <= 100 → Yellow
```

---

## Performance Considerations

### Tips for Large Datasets

**1. Limit conditional formatting range:**
```
❌ Slow: Apply to entire column (A:A)
✅ Fast: Apply to specific range (A2:A1000)
```

**2. Minimize number of rules:**
```
❌ Slow: 10 separate rules
✅ Fast: Combine into 2-3 rules using OR/AND
```

**3. Avoid volatile functions in formulas:**
```
❌ Slow: =TODAY(), =NOW(), =INDIRECT()
✅ Fast: Static references and simple comparisons
```

**4. Use built-in rules when possible:**
```
✅ Fast: Built-in "Greater Than"
❌ Slower: Custom formula =A2>100
```

**5. Remove unused rules:**
- Regularly audit and delete old rules
- Conditional Formatting → Manage Rules → Delete

---

## Copying Conditional Formatting

### Method 1: Format Painter

**Steps:**
1. Select cell with conditional formatting
2. Home → Format Painter (paintbrush icon)
3. Click target cell/range

⚠️ **Note:** Copies ALL formatting, not just conditional

### Method 2: Paste Special

**Steps:**
1. Copy source cell (Ctrl + C)
2. Select target range
3. Right-click → Paste Special → Formats
4. Click OK

### Method 3: Manage Rules

**Steps:**
1. Conditional Formatting → Manage Rules
2. Select rule
3. Click "Edit Rule"
4. Change "Applies to" range
5. Click OK

---

## Conditional Formatting with Tables

When you apply conditional formatting to an Excel Table:

### Benefits
- Automatically extends to new rows
- Easier to manage with structured references
- Cleaner formula syntax

### Example with Table

**Regular range formula:**
```
=$B2>1000
```

**Table formula:**
```
=[@Sales]>1000

Clearer: "This row's Sales column > 1000"
```

**Table advantages:**
- Formulas are more readable
- Auto-expands with table
- No absolute references needed

---

## Best Practices

### 1. Start Simple
```
✅ Begin with built-in rules
✅ Test on small range first
✅ Add complexity gradually
```

### 2. Use Consistent Color Schemes
```
✅ Red = Bad/Low/Overdue
✅ Yellow = Warning/Medium/Due Soon
✅ Green = Good/High/Complete
```

### 3. Don't Over-Format
```
❌ Too many colors → Confusing
✅ 2-3 colors maximum per data set
```

### 4. Document Complex Rules
```
✅ Add cell note explaining formula logic
✅ Use descriptive rule names in Manage Rules
```

### 5. Test with Edge Cases
```
✅ Test with blank cells
✅ Test with zeros
✅ Test with text vs numbers
✅ Test with dates
```

### 6. Combine with Data Validation
```
Conditional Formatting: Visual feedback
Data Validation: Prevents invalid entry
Together: Powerful data quality control
```

---

## Conditional Formatting Gallery

### Popular Patterns

**1. Traffic Light System:**
```
Green:  >= 90%
Yellow: >= 70% and < 90%
Red:    < 70%
```

**2. Heat Map:**
```
Use 3-Color Scale:
Red → Yellow → Green (or Blue → White → Red)
```

**3. Progress Bars:**
```
Data Bars showing % complete
0% = empty, 100% = full
```

**4. Expiration Warnings:**
```
Expires in 30+ days: Green
Expires in 7-30 days: Yellow
Expires in < 7 days: Red
Expired: Dark red
```

**5. Variance Analysis:**
```
Positive variance: Green
Near zero: Yellow
Negative variance: Red
```

**6. Priority Flags:**
```
High priority: Red flag icon
Medium priority: Yellow flag
Low priority: Green flag
```

---

## Troubleshooting Guide

### Problem: Formatting Not Visible

**Check:**
- [ ] Is cell manually formatted? (Clear formats first)
- [ ] Is rule applied to correct range?
- [ ] Is formula returning TRUE?
- [ ] Are there conflicting rules?
- [ ] Is "Stop If True" blocking the rule?

### Problem: Formula Not Working

**Check:**
- [ ] Does formula return TRUE/FALSE?
- [ ] Are cell references correct?
- [ ] Is $ placement correct (absolute vs relative)?
- [ ] Are you referencing the first row of your range?
- [ ] Are there typos in cell references?

### Problem: Rule Applies to Wrong Cells

**Check:**
- [ ] "Applies to" range in Manage Rules
- [ ] Absolute vs relative references
- [ ] Did you select the entire range before creating rule?

### Problem: Slow Performance

**Solutions:**
- [ ] Reduce range size
- [ ] Simplify formulas
- [ ] Delete unused rules
- [ ] Use built-in rules instead of custom formulas
- [ ] Avoid volatile functions (TODAY, NOW, INDIRECT)

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Alt + H, L` | Open Conditional Formatting menu |
| `Alt + H, L, C` | Clear rules from selected cells |
| `Alt + H, L, R` | Manage rules |
| `Ctrl + 1` | Format Cells dialog |
| `F4` | Repeat last action (including formatting) |

---

## Quick Reference: Rule Types

### When to Use Each Type

| Rule Type | Best For | Example |
|-----------|----------|---------|
| **Highlight Cells** | Simple thresholds | Sales > $1000 |
| **Top/Bottom** | Rankings | Top 10 performers |
| **Data Bars** | Comparing values | Sales by month |
| **Color Scales** | Heat maps | Performance matrix |
| **Icon Sets** | Trends/status | Up/down/flat arrows |
| **Custom Formula** | Complex logic | Entire row formatting |

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Conditional formatting is in Home tab
- Rules can use formulas that return TRUE/FALSE
- Use `$B2` (not `$B$2`) for column-locked row-flexible formulas
- Format Painter copies conditional formatting
- Manage Rules shows all active rules
- Rules evaluate top to bottom

### Practice Deeply
- Applying built-in rules (Greater Than, Top 10, etc.)
- Creating data bars for visual comparisons
- Using color scales for heat maps
- Setting up icon sets with custom thresholds
- Writing custom formulas for row highlighting
- Using mixed references correctly ($B2 vs $B$2 vs B$2)
- Highlighting duplicates and unique values
- Creating alternate row colors
- Building dashboards with multiple formatting rules
- Comparing columns with conditional formatting
- Managing and editing existing rules
- Testing formulas to ensure they return TRUE/FALSE
- Combining multiple conditions with AND/OR
- Formatting based on dates and date ranges

---

## Advanced Tips

### Tip 1: Use Named Ranges in Formulas
```
Instead of: =$B2>$G$1
Better: =$B2>Threshold

More readable and easier to maintain
```

### Tip 2: Combine with Data Validation
```
Data Validation: Dropdown list
Conditional Formatting: Color based on selection

Example: Priority dropdown → Color codes task rows
```

### Tip 3: Create Dynamic Thresholds
```
Formula: =$B2>AVERAGE($B$2:$B$10)

Highlights values above average
Threshold updates automatically as data changes
```

### Tip 4: Use INDIRECT for Cross-Sheet Formatting
```
Formula: =INDIRECT("Sheet2!A"&ROW())>100

References same row from another sheet
Enables complex multi-sheet formatting
```

### Tip 5: Conditional Formatting in Charts
```
While not direct conditional formatting:
- Create helper column with IF formulas
- Chart the helper column
- Produces color-coded charts based on conditions
```

---

## Common Use Cases Checklist

Use conditional formatting to:

- [ ] Highlight top performers
- [ ] Flag overdue items
- [ ] Show budget variances
- [ ] Visualize progress toward goals
- [ ] Identify duplicates
- [ ] Mark weekend dates
- [ ] Color-code priorities
- [ ] Create heat maps
- [ ] Show trends with arrows
- [ ] Highlight blank or error cells
- [ ] Compare actual vs. target
- [ ] Visualize survey results
- [ ] Track project status
- [ ] Monitor inventory levels
- [ ] Flag outliers in data

---

## Next Step

After this file, we move to:

**`12-sorting-and-filtering.md`**
- Creating dropdown lists
- Setting input restrictions (numbers, dates, text length)
- Custom validation rules
- Input messages and error alerts
- Circle invalid data
- Dependent dropdowns
- Protecting data integrity
