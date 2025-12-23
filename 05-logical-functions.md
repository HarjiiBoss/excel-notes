# Logical Functions

This file covers Excel's logical functions that enable decision-making in formulas.
These functions let your spreadsheets respond dynamically to different conditions
and handle complex business logic.

---

## What are Logical Functions?

**Logical functions** make decisions based on conditions being TRUE or FALSE.

Think of them as **if-then statements** in everyday logic:
- "**IF** it's raining, **THEN** bring an umbrella"
- "**IF** score is above 90, **THEN** grade is A"
- "**IF** inventory is low **AND** price is good, **THEN** reorder"

### The Power of Logic

Logical functions transform static spreadsheets into **dynamic tools** that:
- Make automatic decisions
- Flag problems
- Calculate different results based on conditions
- Handle errors gracefully
- Validate data

---

## TRUE and FALSE Values

Excel has two logical values:

### TRUE
Represents a condition that is met or correct.

**Examples that return TRUE:**
```
=5>3        → TRUE
=10=10      → TRUE
="Yes"="Yes" → TRUE
=A1>B1      → TRUE (if A1 is greater than B1)
```

### FALSE
Represents a condition that is not met or incorrect.

**Examples that return FALSE:**
```
=5>10       → FALSE
=10=20      → FALSE
="Yes"="No"  → FALSE
=A1>B1      → FALSE (if A1 is not greater than B1)
```

### Visual Concept
```
Comparison Operators:

=5>3     Is 5 greater than 3?        → TRUE
=5<3     Is 5 less than 3?           → FALSE
=5=5     Is 5 equal to 5?            → TRUE
=5<>3    Is 5 not equal to 3?        → TRUE
=5>=5    Is 5 greater or equal to 5? → TRUE
=5<=3    Is 5 less or equal to 3?    → FALSE
```

---

## Comparison Operators

These operators compare values and return TRUE or FALSE:

| Operator | Meaning | Example | Result |
|----------|---------|---------|--------|
| `=` | Equal to | `=A1=B1` | TRUE if equal |
| `>` | Greater than | `=A1>B1` | TRUE if A1 > B1 |
| `<` | Less than | `=A1<B1` | TRUE if A1 < B1 |
| `>=` | Greater than or equal | `=A1>=B1` | TRUE if A1 ≥ B1 |
| `<=` | Less than or equal | `=A1<=B1` | TRUE if A1 ≤ B1 |
| `<>` | Not equal to | `=A1<>B1` | TRUE if not equal |

### Examples
```
     A      B       C
  ┌──────┬──────┬──────────────┐
1 │ 100  │ 50   │ =A1>B1       │ → TRUE
2 │ 30   │ 30   │ =A2=B2       │ → TRUE
3 │ 75   │ 80   │ =A3<B3       │ → TRUE
4 │ 45   │ 45   │ =A4<>B4      │ → FALSE
5 │ 60   │ 55   │ =A5<=B5      │ → FALSE
  └──────┴──────┴──────────────┘
```

---

## IF Function

**Purpose:** Returns one value if condition is TRUE, another if FALSE

**Syntax:** `=IF(logical_test, value_if_true, value_if_false)`

### Structure Breakdown
```
=IF(condition, do_this_if_true, do_this_if_false)
    ↑          ↑                 ↑
    |          |                 └── FALSE result
    |          └──────────────────── TRUE result
    └─────────────────────────────── Test to evaluate
```

### Basic Example
```
     A          B
  ┌─────────┬─────────┐
1 │ Score   │ Result  │
2 │ 85      │ =IF(A2>=70, "Pass", "Fail")
  └─────────┴─────────┘
              ↓
           "Pass"

How it works:
1. Test: Is A2 >= 70?
2. A2 is 85, so YES (TRUE)
3. Return "Pass"
```

### Visual Logic Flow
```
                    IF Function
                        │
                        ▼
              ┌─────────────────┐
              │   A2 >= 70?     │
              └─────────────────┘
                /            \
              YES            NO
              │               │
              ▼               ▼
          "Pass"          "Fail"
```

### Example 1: Pass/Fail Grading
```
     A          B
  ┌─────────┬─────────┐
1 │ Score   │ Grade   │
2 │ 85      │ =IF(A2>=60,"Pass","Fail")  → Pass
3 │ 45      │ =IF(A3>=60,"Pass","Fail")  → Fail
4 │ 72      │ =IF(A4>=60,"Pass","Fail")  → Pass
5 │ 58      │ =IF(A5>=60,"Pass","Fail")  → Fail
  └─────────┴─────────┘
```

### Example 2: Sales Bonus Calculation
```
     A          B          C
  ┌─────────┬─────────┬─────────────┐
1 │ Sales   │ Target  │ Bonus       │
2 │ 15000   │ 10000   │ =IF(A2>B2, 500, 0)  → 500
3 │ 8000    │ 10000   │ =IF(A3>B3, 500, 0)  → 0
4 │ 12000   │ 10000   │ =IF(A4>B4, 500, 0)  → 500
  └─────────┴─────────┴─────────────┘
```

### Example 3: Stock Alert
```
     A          B          C
  ┌─────────┬─────────┬──────────────────┐
1 │ Item    │ Stock   │ Status           │
2 │ Widget  │ 15      │ =IF(B2<20,"Low Stock","OK")
3 │ Gadget  │ 45      │ =IF(B3<20,"Low Stock","OK")
4 │ Tool    │ 8       │ =IF(B4<20,"Low Stock","OK")
  └─────────┴─────────┴──────────────────┘
                          ↓      ↓      ↓
                     "Low Stock" "OK" "Low Stock"
```

### IF with Calculations
You can perform calculations in the TRUE/FALSE parts:

```
     A          B
  ┌─────────┬──────────────────────┐
1 │ Sales   │ Commission           │
2 │ 5000    │ =IF(A2>=10000, A2*0.15, A2*0.10)
3 │ 12000   │ =IF(A3>=10000, A3*0.15, A3*0.10)
  └─────────┴──────────────────────┘
              ↓           ↓
            500         1800
           (10%)       (15%)
```

### ⚠️ Important Notes
- Text values must be in **quotes**: `"Pass"`, `"Fail"`
- Numbers don't need quotes: `100`, `0`, `-50`
- Cell references don't need quotes: `A1`, `B2*2`
- Empty result: use `""` (empty quotes)

---

## Nested IF Statements

**Purpose:** Test multiple conditions in sequence

**Pattern:** IF inside another IF

### Basic Structure
```
=IF(test1, result1, IF(test2, result2, result3))
    │              │
    │              └── If test1 is FALSE, check test2
    └──────────────── If test1 is TRUE, return result1
```

### Example 1: Letter Grades
```
     A          B
  ┌─────────┬─────────────────────────────────────┐
1 │ Score   │ Grade                               │
2 │ 92      │ =IF(A2>=90,"A",IF(A2>=80,"B","C"))  │
  └─────────┴─────────────────────────────────────┘
              ↓
           "A"

Logic Flow:
1. Is score >= 90? YES → "A" ✓
2. (Doesn't check further)
```

### Example 2: Expanded Grading Scale
```
=IF(A2>=90,"A",IF(A2>=80,"B",IF(A2>=70,"C",IF(A2>=60,"D","F"))))
```

**Visual breakdown:**
```
     Score      Grade
     ┌────┐
     │ 92 │
     └────┘
       │
       ▼
    >= 90? ──YES──> "A" ✓
       │
       NO
       ▼
    >= 80? ──YES──> "B"
       │
       NO
       ▼
    >= 70? ──YES──> "C"
       │
       NO
       ▼
    >= 60? ──YES──> "D"
       │
       NO
       ▼
      "F"
```

### Example 3: Sales Commission Tiers
```
     A          B
  ┌─────────┬──────────────────────────────────────────────┐
1 │ Sales   │ Rate                                         │
2 │ 25000   │ =IF(A2>=20000,"15%",IF(A2>=10000,"10%","5%"))│
  └─────────┴──────────────────────────────────────────────┘
              ↓
           "15%"

Tiers:
- $20,000+   → 15%
- $10,000+   → 10%
- Under $10k → 5%
```

### Real-World Example: Shipping Costs
```
     A          B          C
  ┌─────────┬─────────┬────────────────────────────────────┐
1 │ Order   │ Weight  │ Shipping Cost                      │
2 │ #1001   │ 2.5     │ =IF(B2>10,25,IF(B2>5,15,IF(B2>1,8,5)))│
3 │ #1002   │ 0.5     │ =IF(B3>10,25,IF(B3>5,15,IF(B3>1,8,5)))│
4 │ #1003   │ 7.3     │ =IF(B4>10,25,IF(B4>5,15,IF(B4>1,8,5)))│
5 │ #1004   │ 12.0    │ =IF(B5>10,25,IF(B5>5,15,IF(B5>1,8,5)))│
  └─────────┴─────────┴────────────────────────────────────┘
                          ↓    ↓    ↓    ↓
                          8    5    15   25

Rules:
- Over 10 lbs  → $25
- Over 5 lbs   → $15
- Over 1 lb    → $8
- 1 lb or less → $5
```

### ⚠️ Nested IF Limitations
- Excel allows up to **64 nested IF statements**
- Beyond 3-4 levels becomes hard to read
- Consider using **IFS** function instead (covered later)
- Or use lookup functions (VLOOKUP, XLOOKUP)

---

## AND Function

**Purpose:** Returns TRUE only if **ALL** conditions are TRUE

**Syntax:** `=AND(logical1, [logical2], ...)`

### Truth Table
```
Condition 1    Condition 2    AND Result
─────────────────────────────────────────
   TRUE           TRUE          TRUE ✓
   TRUE           FALSE         FALSE
   FALSE          TRUE          FALSE
   FALSE          FALSE         FALSE
```

### Visual Concept
```
All conditions must be TRUE:

        ┌─────┐
        │Test1│────TRUE───┐
        └─────┘           │
                          ├──> AND ──> TRUE
        ┌─────┐           │
        │Test2│────TRUE───┘
        └─────┘

If ANY condition is FALSE:

        ┌─────┐
        │Test1│────TRUE────┐
        └─────┘            │
                           ├──> AND ──> FALSE
        ┌─────┐            │
        │Test2│────FALSE───┘
        └─────┘
```

### Example 1: Basic AND
```
     A      B       C
  ┌──────┬──────┬────────────────┐
1 │ Num1 │ Num2 │ Result         │
2 │ 10   │ 20   │ =AND(A2>5, B2>15)  → TRUE
3 │ 10   │ 12   │ =AND(A3>5, B3>15)  → FALSE
4 │ 3    │ 20   │ =AND(A4>5, B4>15)  → FALSE
  └──────┴──────┴────────────────┘

Row 2: 10>5 is TRUE AND 20>15 is TRUE  → TRUE
Row 3: 10>5 is TRUE BUT 12>15 is FALSE → FALSE
Row 4: 3>5 is FALSE                    → FALSE
```

### Example 2: AND with IF
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────────────┐
1 │ Score   │ Attend  │ Status                       │
2 │ 85      │ 95%     │ =IF(AND(A2>=70,B2>=80),"Eligible","Not Eligible")
3 │ 65      │ 90%     │ =IF(AND(A3>=70,B3>=80),"Eligible","Not Eligible")
4 │ 75      │ 75%     │ =IF(AND(A4>=70,B4>=80),"Eligible","Not Eligible")
  └─────────┴─────────┴──────────────────────────────┘
                          ↓           ↓           ↓
                     "Eligible"  "Not Eligible"  "Not Eligible"

Requirements: Score >= 70 AND Attendance >= 80%
```

### Example 3: Multiple Conditions
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────┐
1 │ Age     │ License │ Insured │ Can Rent           │
2 │ 25      │ Yes     │ Yes     │ =IF(AND(A2>=21,B2="Yes",C2="Yes"),"Approved","Denied")
3 │ 19      │ Yes     │ Yes     │ =IF(AND(A3>=21,B3="Yes",C3="Yes"),"Approved","Denied")
4 │ 30      │ Yes     │ No      │ =IF(AND(A4>=21,B4="Yes",C4="Yes"),"Approved","Denied")
  └─────────┴─────────┴─────────┴────────────────────┘
                                      ↓         ↓         ↓
                                 "Approved" "Denied" "Denied"

All THREE conditions must be TRUE
```

### Real-World Example: Loan Approval
```
     A           B          C         D
  ┌──────────┬─────────┬────────┬────────────────────────────┐
1 │ Name     │ Income  │ Credit │ Decision                   │
2 │ John     │ 60000   │ 720    │ =IF(AND(B2>=50000,C2>=680),"Approved","Denied")
3 │ Sarah    │ 45000   │ 750    │ =IF(AND(B3>=50000,C3>=680),"Approved","Denied")
4 │ Mike     │ 70000   │ 650    │ =IF(AND(B4>=50000,C4>=680),"Approved","Denied")
  └──────────┴─────────┴────────┴────────────────────────────┘
                                      ↓         ↓         ↓
                                 "Approved" "Denied" "Denied"

Criteria:
- Income >= $50,000 AND
- Credit Score >= 680
```

---

## OR Function

**Purpose:** Returns TRUE if **ANY** condition is TRUE

**Syntax:** `=OR(logical1, [logical2], ...)`

### Truth Table
```
Condition 1    Condition 2    OR Result
────────────────────────────────────────
   TRUE           TRUE          TRUE ✓
   TRUE           FALSE         TRUE ✓
   FALSE          TRUE          TRUE ✓
   FALSE          FALSE         FALSE
```

### Visual Concept
```
ANY condition can be TRUE:

        ┌─────┐
        │Test1│────TRUE───┐
        └─────┘           │
                          ├──> OR ──> TRUE
        ┌─────┐           │
        │Test2│────FALSE──┘
        └─────┘

Only FALSE if ALL are FALSE:

        ┌─────┐
        │Test1│────FALSE──┐
        └─────┘           │
                          ├──> OR ──> FALSE
        ┌─────┐           │
        │Test2│────FALSE──┘
        └─────┘
```

### Example 1: Basic OR
```
     A      B       C
  ┌──────┬──────┬────────────────┐
1 │ Num1 │ Num2 │ Result         │
2 │ 10   │ 3    │ =OR(A2>5, B2>5)   → TRUE
3 │ 3    │ 8    │ =OR(A3>5, B3>5)   → TRUE
4 │ 2    │ 3    │ =OR(A4>5, B4>5)   → FALSE
  └──────┴──────┴────────────────┘

Row 2: 10>5 is TRUE → TRUE (doesn't matter that 3>5 is FALSE)
Row 3: 3>5 is FALSE but 8>5 is TRUE → TRUE
Row 4: Both FALSE → FALSE
```

### Example 2: OR with IF
```
     A          B          C
  ┌─────────┬─────────┬────────────────────────────────┐
1 │ Method  │ Days    │ Status                         │
2 │ Express │ 2       │ =IF(OR(A2="Express",B2<=3),"Rush","Standard")
3 │ Regular │ 5       │ =IF(OR(A3="Express",B3<=3),"Rush","Standard")
4 │ Regular │ 2       │ =IF(OR(A4="Express",B4<=3),"Rush","Standard")
  └─────────┴─────────┴────────────────────────────────┘
                          ↓         ↓         ↓
                       "Rush"  "Standard"  "Rush"

Rush if: Express shipping OR 3 days or less
```

### Example 3: Discount Eligibility
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────────────┐
1 │ Customer│ Loyalty │ Order   │ Discount                   │
2 │ John    │ Gold    │ 150     │ =IF(OR(B2="Gold",C2>=200),"15%","No Discount")
3 │ Sarah   │ Silver  │ 250     │ =IF(OR(B3="Gold",C3>=200),"15%","No Discount")
4 │ Mike    │ Bronze  │ 100     │ =IF(OR(B4="Gold",C4>=200),"15%","No Discount")
  └─────────┴─────────┴─────────┴────────────────────────────┘
                                      ↓            ↓            ↓
                                   "15%"        "15%"    "No Discount"

Discount if: Gold Member OR Order >= $200
```

### Real-World Example: Alert System
```
     A          B          C         D
  ┌─────────┬─────────┬────────┬────────────────────────────────┐
1 │ Item    │ Stock   │ Days   │ Alert                          │
2 │ Widget  │ 5       │ 25     │ =IF(OR(B2<10,C2>30),"Action Required","OK")
3 │ Gadget  │ 50      │ 15     │ =IF(OR(B3<10,C3>30),"Action Required","OK")
4 │ Tool    │ 15      │ 35     │ =IF(OR(B4<10,C4>30),"Action Required","OK")
  └─────────┴─────────┴────────┴────────────────────────────────┘
                                      ↓              ↓              ↓
                              "Action Required"   "OK"    "Action Required"

Alert if: Stock < 10 OR Days since order > 30
```

---

## NOT Function

**Purpose:** Reverses TRUE to FALSE and FALSE to TRUE

**Syntax:** `=NOT(logical)`

### Truth Table
```
Input     NOT Result
─────────────────────
TRUE      FALSE
FALSE     TRUE
```

### Visual Concept
```
NOT flips the result:

    Input: TRUE  ──> NOT ──> FALSE
    Input: FALSE ──> NOT ──> TRUE
```

### Example 1: Basic NOT
```
     A          B
  ┌─────────┬────────────┐
1 │ Value   │ Result     │
2 │ 10      │ =NOT(A2>5)    → FALSE
3 │ 3       │ =NOT(A3>5)    → TRUE
  └─────────┴────────────┘

Row 2: A2>5 is TRUE, NOT flips it to FALSE
Row 3: A3>5 is FALSE, NOT flips it to TRUE
```

### Example 2: NOT with IF
```
     A          B
  ┌─────────┬────────────────────────────┐
1 │ Status  │ Action                     │
2 │ Active  │ =IF(NOT(A2="Active"),"Contact","OK")
3 │ Paused  │ =IF(NOT(A3="Active"),"Contact","OK")
  └─────────┴────────────────────────────┘
              ↓           ↓
           "OK"      "Contact"

Contact if status is NOT "Active"
```

### Example 3: Exclude Conditions
```
     A          B          C
  ┌─────────┬─────────┬────────────────────────────────┐
1 │ Item    │ Stock   │ Status                         │
2 │ Widget  │ 50      │ =IF(NOT(B2<20),"Sufficient","Low Stock")
3 │ Gadget  │ 15      │ =IF(NOT(B3<20),"Sufficient","Low Stock")
  └─────────┴─────────┴────────────────────────────────┘
                          ↓               ↓
                    "Sufficient"    "Low Stock"

"Sufficient" if stock is NOT less than 20
(In other words: if stock >= 20)
```

### When to Use NOT
- Reversing a complex condition
- Making logic more readable in some cases
- Combining with AND/OR for inverse logic

**Note:** Often you can avoid NOT by changing the comparison:
```
❌ =NOT(A1>5)
✅ =A1<=5     (simpler, same result)

❌ =NOT(A1="Yes")
✅ =A1<>"Yes"  (simpler, same result)
```

---

## Combining AND, OR, NOT

You can nest logical functions for complex conditions.

### Example 1: AND with OR
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────────────┐
1 │ Age     │ Student │ Senior  │ Discount                   │
2 │ 25      │ Yes     │ No      │ =IF(OR(B2="Yes",C2="Yes",A2<18),"10% Off","No Discount")
3 │ 70      │ No      │ Yes     │ =IF(OR(B3="Yes",C3="Yes",A3<18),"10% Off","No Discount")
4 │ 35      │ No      │ No      │ =IF(OR(B4="Yes",C4="Yes",A4<18),"10% Off","No Discount")
  └─────────┴─────────┴─────────┴────────────────────────────┘
                                      ↓            ↓            ↓
                                  "10% Off"   "10% Off"  "No Discount"

Discount if: Student OR Senior OR Under 18
```

### Example 2: Complex Eligibility
```
=IF(AND(A2>=21, OR(B2="Yes",C2="Yes")), "Eligible", "Not Eligible")
                │         │      │
                │         └──────┴─── At least ONE must be TRUE
                └──────────────────── AND this must be TRUE

Requirements:
- Age >= 21 AND
- (Has License OR Has Insurance)
```

### Example 3: Inventory Reorder Logic
```
     A          B          C         D          E
  ┌─────────┬─────────┬────────┬─────────┬────────────────────────────┐
1 │ Item    │ Stock   │ Demand │ Lead    │ Reorder                    │
2 │ Widget  │ 25      │ High   │ 7       │ =IF(AND(B2<50,OR(C2="High",D2>5)),"Yes","No")
3 │ Gadget  │ 60      │ Low    │ 3       │ =IF(AND(B3<50,OR(C3="High",D3>5)),"Yes","No")
4 │ Tool    │ 30      │ Medium │ 10      │ =IF(AND(B4<50,OR(C4="High",D4>5)),"Yes","No")
  └─────────┴─────────┴────────┴─────────┴────────────────────────────┘
                                              ↓       ↓       ↓
                                           "Yes"   "No"   "Yes"

Reorder if:
- Stock < 50 AND
- (Demand is High OR Lead time > 5 days)
```

### Visual Logic Tree
```
                    AND
                     │
          ┌──────────┴──────────┐
          │                     │
      Stock < 50               OR
                                │
                    ┌───────────┴───────────┐
                    │                       │
              Demand = "High"        Lead Time > 5

All branches must be satisfied for "Yes"
```

---

## IFS Function

**Purpose:** Test multiple conditions without nesting IF statements

**Syntax:** `=IFS(test1, value1, test2, value2, ...)`

**Available in:** Excel 2019, Microsoft 365, Excel Online

### Why Use IFS?

**Old Way (Nested IF):**
```
=IF(A2>=90,"A",IF(A2>=80,"B",IF(A2>=70,"C",IF(A2>=60,"D","F"))))
                 └──────────────────────────────────────────────┘
                          Hard to read and maintain
```

**New Way (IFS):**
```
=IFS(A2>=90,"A", A2>=80,"B", A2>=70,"C", A2>=60,"D", TRUE,"F")
      │     │      │     │      │     │      │     │    │   │
      Test1 Result Test2 Result Test3 Result Test4 Result Default
```

### Example 1: Letter Grades
```
     A          B
  ┌─────────┬────────────────────────────────────────────────┐
1 │ Score   │ Grade                                          │
2 │ 92      │ =IFS(A2>=90,"A", A2>=80,"B", A2>=70,"C", A2>=60,"D", TRUE,"F")
3 │ 75      │ =IFS(A3>=90,"A", A3>=80,"B", A3>=70,"C", A3>=60,"D", TRUE,"F")
4 │ 58      │ =IFS(A4>=90,"A", A4>=80,"B", A4>=70,"C", A4>=60,"D", TRUE,"F")
  └─────────┴────────────────────────────────────────────────┘
              ↓      ↓      ↓
           "A"    "C"    "F"
```

**How it works:**
1. Tests conditions in order
2. Returns the value for the **first TRUE** condition
3. Stops testing once a condition is met
4. `TRUE` as the last test acts as a "catch-all" default

### Example 2: Shipping Speed
```
     A          B
  ┌─────────┬────────────────────────────────────────────┐
1 │ Days    │ Method                                     │
2 │ 1       │ =IFS(A2=1,"Overnight", A2<=3,"Express", A2<=7,"Standard", TRUE,"Slow")
3 │ 3       │ =IFS(A3=1,"Overnight", A3<=3,"Express", A3<=7,"Standard", TRUE,"Slow")
4 │ 10      │ =IFS(A4=1,"Overnight", A4<=3,"Express", A4<=7,"Standard", TRUE,"Slow")
  └─────────┴────────────────────────────────────────────┘
              ↓            ↓          ↓
        "Overnight"  "Express"   "Slow"
```

### Example 3: Commission Tiers
```
     A          B
  ┌─────────┬──────────────────────────────────────────────┐
1 │ Sales   │ Rate                                         │
2 │ 25000   │ =IFS(A2>=50000,0.20, A2>=20000,0.15, A2>=10000,0.10, TRUE,0.05)
3 │ 8000    │ =IFS(A3>=50000,0.20, A3>=20000,0.15, A3>=10000,0.10, TRUE,0.05)
4 │ 55000   │ =IFS(A4>=50000,0.20, A4>=20000,0.15, A4>=10000,0.10, TRUE,0.05)
  └─────────┴──────────────────────────────────────────────┘
              ↓        ↓        ↓
            0.15     0.05     0.20
            (15%)    (5%)     (20%)

Tiers:
- $50,000+  → 20%
- $20,000+  → 15%
- $10,000+  → 10%
- Under $10k → 5%
```

### ⚠️ Important Notes
- Tests are evaluated in order (first to last)
- Order matters! Put most restrictive conditions first
- Always include a default with `TRUE` at the end
- If no condition is TRUE and no default, returns `#N/A` error

### IFS vs Nested IF Comparison

| Aspect | Nested IF | IFS |
|--------|-----------|-----|
| **Readability** | Hard to read with many conditions | Much clearer |
| **Maintenance** | Difficult to update | Easy to update |
| **Parentheses** | Must balance many `)` | No nested parentheses |
| **Availability** | All Excel versions | Excel 2019+ only |

---

## IFERROR Function

**Purpose:** Returns a custom value if a formula produces an error

**Syntax:** `=IFERROR(value, value_if_error)`

### Why Use IFERROR?

Errors like `#DIV/0!`, `#N/A`, `#VALUE!` look unprofessional and confuse users.

**Without IFERROR:**
```
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ Sales   │ Reps    │ Per Rep │
2 │ 10000   │ 5       │ 2000    │
3 │ 5000    │ 0       │ #DIV/0! │ ← Ugly error
4 │ 8000    │ 4       │ 2000    │
  └─────────┴─────────┴─────────┘
```

**With IFERROR:**
```
     A          B          C
  ┌─────────┬─────────┬─────────────────────┐
1 │ Sales   │ Reps    │ Per Rep             │
2 │ 10000   │ 5       │ =IFERROR(A2/B2,"N/A")  → 2000
3 │ 5000    │ 0       │ =IFERROR(A3/B3,"N/A")  → N/A
4 │ 8000    │ 4       │ =IFERROR(A4/B4,"N/A")  → 2000
  └─────────┴─────────┴─────────────────────┘
```

### Example 1: Division Errors
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────┐
1 │ Total   │ Count   │ Average              │
2 │ 100     │ 10      │ =IFERROR(A2/B2,0)       → 10
3 │ 200     │ 0       │ =IFERROR(A3/B3,0)       → 0
4 │ 150     │ 5       │ =IFERROR(A4/B4,0)       → 30
  └─────────┴─────────┴──────────────────────┘

Instead of #DIV/0!, shows 0
```

### Example 2: Lookup Errors
```
     A          B                    C
  ┌─────────┬────────────────────┬──────────┐
1 │ ID      │ Formula            │ Result   │
2 │ 101     │ =IFERROR(VLOOKUP(A2,Table,2,FALSE),"Not Found")
3 │ 999     │ =IFERROR(VLOOKUP(A3,Table,2,FALSE),"Not Found")
  └─────────┴────────────────────┴──────────┘
              ↓                    ↓
         "John Smith"         "Not Found"

If VLOOKUP fails, shows "Not Found" instead of #N/A
```

### Example 3: Clean Reporting
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬──────────────────────┐
1 │ Product │ Sales   │ Target  │ % of Target          │
2 │ Widget  │ 1000    │ 800     │ =IFERROR(B2/C2*100,"--")  → 125%
3 │ Gadget  │ 500     │         │ =IFERROR(B3/C3*100,"--")  → --
4 │ Tool    │ 750     │ 1000    │ =IFERROR(B4/C4*100,"--")  → 75%
  └─────────┴─────────┴─────────┴──────────────────────┘

Shows "--" when target is empty instead of error
```

### Common Error Types Caught by IFERROR

| Error | Cause | IFERROR Handles It? |
|-------|-------|---------------------|
| `#DIV/0!` | Division by zero | ✅ Yes |
| `#N/A` | Value not available (VLOOKUP) | ✅ Yes |
| `#VALUE!` | Wrong type of argument | ✅ Yes |
| `#REF!` | Invalid cell reference | ✅ Yes |
| `#NAME?` | Unrecognized function name | ✅ Yes |
| `#NUM!` | Invalid numeric value | ✅ Yes |
| `#NULL!` | Incorrect range operator | ✅ Yes |

### ⚠️ Important Notes
- Use specific error messages that help users understand the issue
- Common replacements: `"N/A"`, `"--"`, `0`, `""`
- Don't hide errors during development—fix the root cause first
- IFERROR catches **all** errors—use carefully

---

## IFNA Function

**Purpose:** Returns a custom value only for `#N/A` errors (more specific than IFERROR)

**Syntax:** `=IFNA(value, value_if_na)`

### When to Use IFNA

Use IFNA when you only want to handle `#N/A` errors (common with lookups) but want other errors to show normally.

**Example: VLOOKUP with IFNA**
```
     A          B
  ┌─────────┬────────────────────────────────────┐
1 │ ID      │ Name                               │
2 │ 101     │ =IFNA(VLOOKUP(A2,Table,2,FALSE),"Not Found")
3 │ 999     │ =IFNA(VLOOKUP(A3,Table,2,FALSE),"Not Found")
  └─────────┴────────────────────────────────────┘
              ↓                   ↓
         "John Smith"        "Not Found"

Shows "Not Found" for missing IDs
But would still show #REF! or #VALUE! if those occurred
```

### IFERROR vs IFNA

| Function | Handles | When to Use |
|----------|---------|-------------|
| **IFERROR** | All errors | General error handling |
| **IFNA** | Only #N/A | Lookup functions specifically |

**Example showing the difference:**
```
=IFERROR(VLOOKUP(A2,Table,2,FALSE),"Error")  
→ Hides ALL errors with "Error"

=IFNA(VLOOKUP(A2,Table,2,FALSE),"Not Found")
→ Only handles #N/A, other errors still show
```

---

## SWITCH Function

**Purpose:** Match one value against multiple possible values (like a switch statement in programming)

**Syntax:** `=SWITCH(expression, value1, result1, [value2, result2], ..., [default])`

**Available in:** Excel 2019, Microsoft 365, Excel Online

### How SWITCH Works
```
=SWITCH(value_to_match, 
        match1, return_this1,
        match2, return_this2,
        match3, return_this3,
        default_result)
```

### Example 1: Day of Week
```
     A          B
  ┌─────────┬────────────────────────────────────────┐
1 │ Day #   │ Day Name                               │
2 │ 1       │ =SWITCH(A2, 1,"Mon", 2,"Tue", 3,"Wed", 4,"Thu", 5,"Fri", 6,"Sat", 7,"Sun", "Invalid")
3 │ 5       │ =SWITCH(A3, 1,"Mon", 2,"Tue", 3,"Wed", 4,"Thu", 5,"Fri", 6,"Sat", 7,"Sun", "Invalid")
4 │ 9       │ =SWITCH(A4, 1,"Mon", 2,"Tue", 3,"Wed", 4,"Thu", 5,"Fri", 6,"Sat", 7,"Sun", "Invalid")
  └─────────┴────────────────────────────────────────┘
              ↓         ↓         ↓
           "Mon"     "Fri"   "Invalid"
```

### Example 2: Grade to GPA
```
     A          B
  ┌─────────┬────────────────────────────────────┐
1 │ Grade   │ GPA                                │
2 │ A       │ =SWITCH(A2, "A",4.0, "B",3.0, "C",2.0, "D",1.0, "F",0.0)
3 │ B       │ =SWITCH(A3, "A",4.0, "B",3.0, "C",2.0, "D",1.0, "F",0.0)
4 │ F       │ =SWITCH(A4, "A",4.0, "B",3.0, "C",2.0, "D",1.0, "F",0.0)
  └─────────┴────────────────────────────────────┘
              ↓      ↓      ↓
            4.0    3.0    0.0
```

### Example 3: Department Codes
```
     A          B
  ┌─────────┬────────────────────────────────────────────┐
1 │ Code    │ Department                                 │
2 │ HR      │ =SWITCH(A2, "HR","Human Resources", "IT","Information Tech", "FIN","Finance", "MKT","Marketing", "Unknown")
3 │ IT      │ =SWITCH(A3, "HR","Human Resources", "IT","Information Tech", "FIN","Finance", "MKT","Marketing", "Unknown")
4 │ XYZ     │ =SWITCH(A4, "HR","Human Resources", "IT","Information Tech", "FIN","Finance", "MKT","Marketing", "Unknown")
  └─────────┴────────────────────────────────────────────┘
              ↓                   ↓              ↓
      "Human Resources"  "Information Tech"  "Unknown"
```

### SWITCH vs Nested IF

**Using Nested IF:**
```
=IF(A2="HR","Human Resources",IF(A2="IT","Information Tech",IF(A2="FIN","Finance","Unknown")))
```

**Using SWITCH:**
```
=SWITCH(A2, "HR","Human Resources", "IT","Information Tech", "FIN","Finance", "Unknown")
```

**SWITCH is:**
- ✅ Cleaner and more readable
- ✅ Easier to maintain
- ✅ Better for exact matches
- ❌ Only available in Excel 2019+
- ❌ Doesn't work with conditions (use IFS for that)

---

## XOR Function

**Purpose:** Returns TRUE if an **odd number** of arguments are TRUE (exclusive OR)

**Syntax:** `=XOR(logical1, [logical2], ...)`

### Truth Table (2 arguments)
```
Condition 1    Condition 2    XOR Result
─────────────────────────────────────────
   TRUE           TRUE          FALSE
   TRUE           FALSE         TRUE ✓
   FALSE          TRUE          TRUE ✓
   FALSE          FALSE         FALSE
```

### Visual Concept
```
XOR returns TRUE if:
- Exactly ONE argument is TRUE (not both, not neither)

For 2 arguments:
    TRUE + FALSE = TRUE  ✓
    FALSE + TRUE = TRUE  ✓
    TRUE + TRUE = FALSE
    FALSE + FALSE = FALSE
```

### Example 1: Basic XOR
```
     A      B       C
  ┌──────┬──────┬────────────────┐
1 │ Val1 │ Val2 │ Result         │
2 │ TRUE │ FALSE│ =XOR(A2,B2)       → TRUE
3 │ TRUE │ TRUE │ =XOR(A3,B3)       → FALSE
4 │ FALSE│ FALSE│ =XOR(A4,B4)       → FALSE
  └──────┴──────┴────────────────┘
```

### Example 2: Validation Check
```
     A          B          C
  ┌─────────┬─────────┬────────────────────────────┐
1 │ HasCar  │ NeedRide│ Status                     │
2 │ Yes     │ No      │ =IF(XOR(A2="Yes",B2="Yes"),"OK","Problem")
3 │ No      │ No      │ =IF(XOR(A3="Yes",B3="Yes"),"OK","Problem")
4 │ Yes     │ Yes     │ =IF(XOR(A4="Yes",B4="Yes"),"OK","Problem")
  └─────────┴─────────┴────────────────────────────┘
              ↓         ↓         ↓
           "OK"    "Problem"  "Problem"

Logic: Either has car OR needs ride (but not both or neither)
```

### When to Use XOR
- Validation: ensure exactly one option is selected
- Logical puzzles or complex conditions
- Rare in typical business spreadsheets (OR and AND are more common)

---

## Practical Examples and Patterns

### Pattern 1: Tiered Pricing
```
     A          B
  ┌─────────┬────────────────────────────────────────┐
1 │ Qty     │ Price                                  │
2 │ 150     │ =IFS(A2>=100,9.99, A2>=50,12.99, A2>=10,14.99, TRUE,19.99)
  └─────────┴────────────────────────────────────────┘
              ↓
            9.99

Pricing:
- 100+ units: $9.99
- 50-99 units: $12.99
- 10-49 units: $14.99
- 1-9 units: $19.99
```

### Pattern 2: Status Flags
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────────────┐
1 │ Days    │ Amount  │ Paid    │ Status                     │
2 │ 45      │ 1000    │ No      │ =IF(AND(A2>30,C2="No"),"Overdue",IF(C2="Yes","Paid","Current"))
3 │ 15      │ 500     │ No      │ =IF(AND(A3>30,C3="No"),"Overdue",IF(C3="Yes","Paid","Current"))
4 │ 60      │ 2000    │ Yes     │ =IF(AND(A4>30,C4="No"),"Overdue",IF(C4="Yes","Paid","Current"))
  └─────────┴─────────┴─────────┴────────────────────────────┘
                                      ↓          ↓         ↓
                                 "Overdue"  "Current"   "Paid"

Logic:
- If over 30 days AND not paid → "Overdue"
- If paid → "Paid"
- Otherwise → "Current"
```

### Pattern 3: Complex Approval
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────────────────────┐
1 │ Amount  │ Manager │ Budget  │ Decision                           │
2 │ 5000    │ Yes     │ Yes     │ =IF(A2<=1000,"Auto-Approve",IF(AND(A2<=10000,B2="Yes",C2="Yes"),"Approved","Needs Review"))
3 │ 500     │ No      │ No      │ =IF(A3<=1000,"Auto-Approve",IF(AND(A3<=10000,B3="Yes",C3="Yes"),"Approved","Needs Review"))
4 │ 15000   │ Yes     │ Yes     │ =IF(A4<=1000,"Auto-Approve",IF(AND(A4<=10000,B4="Yes",C4="Yes"),"Approved","Needs Review"))
  └─────────┴─────────┴─────────┴────────────────────────────────────┘
                                      ↓               ↓                ↓
                                 "Approved"   "Auto-Approve"   "Needs Review"

Rules:
- $1,000 or less → Auto-approve
- $1,001-$10,000 with manager & budget approval → Approved
- Otherwise → Needs Review
```

### Pattern 4: Data Quality Check
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────────────────┐
1 │ Name    │ Email   │ Phone   │ Status                         │
2 │ John    │ j@co.com│ 555-1234│ =IF(AND(A2<>"",B2<>"",C2<>""),"Complete","Incomplete")
3 │ Sarah   │         │ 555-5678│ =IF(AND(A3<>"",B3<>"",C3<>""),"Complete","Incomplete")
4 │ Mike    │ m@co.com│ 555-9999│ =IF(AND(A4<>"",B4<>"",C4<>""),"Complete","Incomplete")
  └─────────┴─────────┴─────────┴────────────────────────────────┘
                                      ↓            ↓            ↓
                                 "Complete"  "Incomplete"  "Complete"

Check if all required fields are filled
```

---

## Common Mistakes and Best Practices

### Mistake 1: Wrong Quote Usage
```
❌ Wrong: =IF(A1>100,Yes,No)
✅ Right: =IF(A1>100,"Yes","No")

Text must be in quotes!
```

### Mistake 2: Comparing Text Case-Sensitive
```
Excel comparisons are NOT case-sensitive:
"yes" = "Yes" = "YES"  → All TRUE

To force case-sensitive:
=EXACT(A1,"Yes")
```

### Mistake 3: Using = Instead of ==
```
In Excel, comparison uses single =

❌ Wrong: =IF(A1==100,"Yes","No")
✅ Right: =IF(A1=100,"Yes","No")
```

### Mistake 4: Forgetting Default in IFS
```
❌ Problem: =IFS(A1>90,"A", A1>80,"B")
If A1=75, returns #N/A error

✅ Solution: =IFS(A1>90,"A", A1>80,"B", TRUE,"Other")
Always include a catch-all
```

### Mistake 5: Testing Blank Cells Wrong
```
❌ Wrong: =IF(A1="","Empty","Has Value")
          (This treats 0 as empty!)

✅ Better: =IF(A1="","Empty","Has Value")
✅ Best: =IF(ISBLANK(A1),"Empty","Has Value")
```

### Mistake 6: Order of Tests in IFS
```
❌ Wrong Order:
=IFS(A1>=60,"Pass", A1>=90,"A")
Problem: 90 is >=60, so always returns "Pass"

✅ Right Order:
=IFS(A1>=90,"A", A1>=60,"Pass")
Most restrictive conditions FIRST
```

---

## Best Practices

### 1. Keep Logic Simple
```
❌ Too Complex:
=IF(AND(OR(A1>100,B1<50),NOT(C1="X"),D1="Y"),"Yes","No")

✅ Better: Break into steps
E1: =OR(A1>100,B1<50)
F1: =NOT(C1="X")
G1: =IF(AND(E1,F1,D1="Y"),"Yes","No")
```

### 2. Use Helper Columns
```
Instead of:
=IF(AND(A2>100,B2>50),A2*0.15,IF(OR(A2>50,B2>25),A2*0.10,A2*0.05))

Better:
D2: =AND(A2>100,B2>50)  (High Tier)
E2: =OR(A2>50,B2>25)    (Mid Tier)
F2: =IF(D2,A2*0.15,IF(E2,A2*0.10,A2*0.05))  (Calculate)
```

### 3. Document Complex Logic
```
Add comments to cells explaining:
- What conditions are being tested
- Why certain thresholds exist
- Expected outcomes
```

### 4. Test Edge Cases
Always test:
- Minimum values
- Maximum values
- Boundary conditions (exactly at threshold)
- Empty/null values
- Unexpected text

### 5. Use Named Ranges
```
❌ Hard to read:
=IF(A2>=1000,"Gold","Silver")

✅ Clear:
=IF(Sales>=Gold_Threshold,"Gold","Silver")
```

---

## Quick Reference: Logical Functions

| Function | Purpose | Example |
|----------|---------|---------|
| **IF** | Basic condition | `=IF(A1>100,"High","Low")` |
| **IFS** | Multiple conditions | `=IFS(A1>100,"High", A1>50,"Mid", TRUE,"Low")` |
| **AND** | All must be TRUE | `=AND(A1>10, B1<20)` |
| **OR** | Any can be TRUE | `=OR(A1>10, B1<20)` |
| **NOT** | Reverse TRUE/FALSE | `=NOT(A1>10)` |
| **XOR** | Exactly one TRUE | `=XOR(A1>10, B1<20)` |
| **IFERROR** | Handle all errors | `=IFERROR(A1/B1,"Error")` |
| **IFNA** | Handle #N/A only | `=IFNA(VLOOKUP(...),"Not Found")` |
| **SWITCH** | Match exact values | `=SWITCH(A1, 1,"One", 2,"Two", "Other")` |

---

## Troubleshooting Common Errors

### Error: #N/A in IFS
**Cause:** No condition was TRUE and no default provided

**Fix:**
```
❌ =IFS(A1>90,"A", A1>80,"B")
✅ =IFS(A1>90,"A", A1>80,"B", TRUE,"Other")
```

### Error: #VALUE! in IF
**Cause:** Usually comparison of incompatible types

**Example:**
```
=IF(A1>100,"Yes","No")  where A1 contains "text"
```

**Fix:** Ensure data types match your comparison

### Error: Unexpected FALSE Results
**Cause:** Text has extra spaces

**Example:**
```
A1 contains: "Yes "  (with trailing space)
=IF(A1="Yes","Match","No Match")  → "No Match"
```

**Fix:**
```
=IF(TRIM(A1)="Yes","Match","No Match")
```

### Error: Wrong Results with Nested Conditions
**Cause:** Conditions overlap or wrong order

**Fix:** Test conditions from most specific to least specific

---

## What to PRACTICE vs MEMORIZE

### Memorize
- IF syntax: `=IF(test, true_result, false_result)`
- Comparison operators: `=`, `>`, `<`, `>=`, `<=`, `<>`
- AND requires ALL conditions TRUE
- OR requires ANY condition TRUE
- NOT reverses TRUE/FALSE
- Text in formulas must use quotes
- IFS tests conditions in order (first match wins)
- IFERROR handles all errors, IFNA handles only #N/A

### Practice Deeply
- Writing IF statements for different scenarios
- Using AND/OR with IF for complex conditions
- Building nested IF statements (2-3 levels)
- Converting nested IF to IFS
- Using IFERROR to clean up error displays
- Combining logical functions (AND with OR, etc.)
- Testing boundary conditions in your logic
- Building multi-criteria decision formulas
- Using logical functions for data validation
- Creating status flags and alerts

### Don't Memorize
- Every possible combination of logical functions
- Exact error messages
- Which Excel version introduced which function
- Complex nested patterns (build them step by step instead)

---

## Real-World Application: Employee Bonus Calculator

Let's build a complete bonus calculator using logical functions.

### Requirements
- Base salary in column A
- Years of service in column B
- Performance rating in column C (1-5)
- Bonus calculation rules:
  - 5+ years AND rating 4-5: 15% bonus
  - 3-5 years AND rating 4-5: 10% bonus
  - Any years with rating 5: minimum 5% bonus
  - Rating 1-2: no bonus

### Setup
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────────────────┐
1 │ Salary  │ Years   │ Rating  │ Bonus %                    │
2 │ 50000   │ 6       │ 5       │                            │
3 │ 45000   │ 3       │ 4       │                            │
4 │ 60000   │ 2       │ 5       │                            │
5 │ 40000   │ 7       │ 3       │                            │
6 │ 55000   │ 4       │ 2       │                            │
  └─────────┴─────────┴─────────┴────────────────────────────┘
```

### Formula (Column D)
```
=IFS(
  AND(B2>=5,C2>=4), 0.15,
  AND(B2>=3,C2>=4), 0.10,
  C2=5, 0.05,
  C2<=2, 0,
  TRUE, 0
)
```

### Results
```
     A          B          C          D          E
  ┌─────────┬─────────┬─────────┬─────────┬────────────┐
1 │ Salary  │ Years   │ Rating  │ Bonus % │ Bonus $    │
2 │ 50000   │ 6       │ 5       │ 0.15    │ 7500       │
3 │ 45000   │ 3       │ 4       │ 0.10    │ 4500       │
4 │ 60000   │ 2       │ 5       │ 0.05    │ 3000       │
5 │ 40000   │ 7       │ 3       │ 0       │ 0          │
6 │ 55000   │ 4       │ 2       │ 0       │ 0          │
  └─────────┴─────────┴─────────┴─────────┴────────────┘

Column E formula: =A2*D2
```

### Add Error Handling
```
Column E: =IFERROR(A2*D2,"Error in calculation")
```

---

## Next Step

After mastering logical functions, you're ready to explore:

**`06-lookup-and-reference-functions.md`**
- VLOOKUP for finding data in tables
- XLOOKUP (modern, more flexible)
- INDEX and MATCH combination
- HLOOKUP for horizontal lookups
- Finding and referencing data across sheets
- Advanced lookup techniques
