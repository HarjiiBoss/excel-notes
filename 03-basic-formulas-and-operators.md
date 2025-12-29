# Basic Formulas and Operators

This file covers the fundamental operators in Excel, how to build formulas, order of operations, and how to combine calculations.

---

## What is a Formula?

A **formula** is an equation that performs calculations on values in your worksheet.

### Key Rules
- Every formula **must** start with an equals sign `=`
- Formulas can contain:
  - Cell references (`A1`, `B5`)
  - Numbers (`100`, `3.14`)
  - Operators (`+`, `-`, `*`, `/`)
  - Functions (`SUM`, `AVERAGE`)
  - Text in quotes (`"Hello"`)

### Basic Structure
```
=    Value    Operator    Value
│      │          │          │
│      │          │          └─ Cell reference or number
│      │          └─ Mathematical operation
│      └─ Cell reference or number
└─ Required starting character
```

---

## Arithmetic Operators

Excel uses standard mathematical operators to perform calculations.

### The Six Basic Operators

| Operator | Operation | Example | Result |
|----------|-----------|---------|--------|
| `+` | Addition | `=10+5` | 15 |
| `-` | Subtraction | `=10-5` | 5 |
| `*` | Multiplication | `=10*5` | 50 |
| `/` | Division | `=10/5` | 2 |
| `^` | Exponentiation | `=10^2` | 100 |
| `%` | Percent | `=10%` | 0.1 |

### Visual Examples

**Addition:**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ Price  │ Tax    │ Total    │
  ├────────┼────────┼──────────┤
2 │ 100    │ 15     │ =A2+B2   │
  └────────┴────────┴──────────┘

Result: 115
```

**Subtraction:**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ Budget │ Spent  │ Remaining│
  ├────────┼────────┼──────────┤
2 │ 5000   │ 3200   │ =A2-B2   │
  └────────┴────────┴──────────┘

Result: 1800
```

**Multiplication:**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ Qty    │ Price  │ Total    │
  ├────────┼────────┼──────────┤
2 │ 50     │ 25.50  │ =A2*B2   │
  └────────┴────────┴──────────┘

Result: 1275
```

**Division:**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ Total  │ Count  │ Average  │
  ├────────┼────────┼──────────┤
2 │ 1000   │ 25     │ =A2/B2   │
  └────────┴────────┴──────────┘

Result: 40
```

**Exponentiation (Power):**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ Base   │ Exp    │ Result   │
  ├────────┼────────┼──────────┤
2 │ 2      │ 10     │ =A2^B2   │
  └────────┴────────┴──────────┘

Result: 1024
```

**Percent:**
```
     A              B
  ┌────────────┬──────────┐
1 │ Discount   │ Decimal  │
  ├────────────┼──────────┤
2 │ 15%        │ =A2      │
  └────────────┴──────────┘

Result: 0.15

Note: 15% is stored as 0.15 internally
```

---

## Combining Multiple Operators

You can combine multiple operations in a single formula.

### Simple Combined Formula
```
=A1+B1-C1
```

### Real-World Example: Profit Calculation
```
     A         B         C         D            E
  ┌────────┬────────┬────────┬─────────┬──────────────┐
1 │ Revenue│ COGS   │ OpEx   │ Taxes   │ Net Profit   │
  ├────────┼────────┼────────┼─────────┼──────────────┤
2 │ 10000  │ 4000   │ 2000   │ 800     │ =A2-B2-C2-D2 │
  └────────┴────────┴────────┴─────────┴──────────────┘

Result: 3200

Formula reads: Revenue minus COGS minus OpEx minus Taxes
```

### Percentage Calculations
```
     A         B              C
  ┌────────┬────────────┬──────────────┐
1 │ Price  │ Discount % │ Final Price  │
  ├────────┼────────────┼──────────────┤
2 │ 200    │ 15%        │ =A2*(1-B2)   │
  └────────┴────────────┴──────────────┘

Result: 170

Explanation:
1 - B2      → 1 - 0.15 = 0.85 (85%)
A2 * 0.85   → 200 * 0.85 = 170
```

---

## Order of Operations (PEMDAS)

Excel follows standard mathematical order of operations.

### The Order

**PEMDAS (or BEDMAS):**
1. **P**arentheses (Brackets)
2. **E**xponents (Powers)
3. **M**ultiplication and **D**ivision (left to right)
4. **A**ddition and **S**ubtraction (left to right)

### Excel's Order
```
Priority 1:  ( )      Parentheses
Priority 2:  ^        Exponentiation
Priority 3:  * /      Multiplication and Division
Priority 4:  + -      Addition and Subtraction
```

### Example 1: Without Parentheses
```
=10+5*2

Step 1: Multiply first    → 5*2 = 10
Step 2: Then add          → 10+10 = 20

Result: 20
```

### Example 2: With Parentheses
```
=(10+5)*2

Step 1: Parentheses first → (10+5) = 15
Step 2: Then multiply     → 15*2 = 30

Result: 30
```

### Visual Comparison
```
     A                      B                C
  ┌──────────────────┬──────────────────┬─────────┐
1 │ Without Parens   │ With Parens      │ Result  │
  ├──────────────────┼──────────────────┼─────────┤
2 │ =10+5*2          │ =(10+5)*2        │ 20 vs 30│
  ├──────────────────┼──────────────────┼─────────┤
3 │ =100-50/2        │ =(100-50)/2      │ 75 vs 25│
  ├──────────────────┼──────────────────┼─────────┤
4 │ =5+3*2^2         │ =((5+3)*2)^2     │ 17 vs 256│
  └──────────────────┴──────────────────┴─────────┘
```

### Real-World Example: Sales Commission

**Scenario:** Base salary + commission on sales above quota

```
     A         B         C         D              E
  ┌────────┬────────┬────────┬──────────┬─────────────────────┐
1 │ Base   │ Sales  │ Quota  │ Rate     │ Total Pay           │
  ├────────┼────────┼────────┼──────────┼─────────────────────┤
2 │ 3000   │ 50000  │ 30000  │ 10%      │ =A2+(B2-C2)*D2      │
  └────────┴────────┴────────┴──────────┴─────────────────────┘

Calculation:
(B2-C2)*D2   → (50000-30000)*0.10 = 2000
A2 + 2000    → 3000 + 2000 = 5000

Result: 5000
```

**Without parentheses (wrong):**
```
❌ =A2+B2-C2*D2

Calculation:
C2*D2     → 30000*0.10 = 3000
A2+B2     → 3000+50000 = 53000
53000-3000 → 50000 (WRONG!)
```

---

## Comparison Operators

Comparison operators test relationships between values and return `TRUE` or `FALSE`.

### The Six Comparison Operators

| Operator | Meaning | Example | Result |
|----------|---------|---------|--------|
| `=` | Equal to | `=5=5` | TRUE |
| `>` | Greater than | `=10>5` | TRUE |
| `<` | Less than | `=10<5` | FALSE |
| `>=` | Greater than or equal | `=5>=5` | TRUE |
| `<=` | Less than or equal | `=5<=4` | FALSE |
| `<>` | Not equal to | `=5<>3` | TRUE |

### Visual Examples

```
     A         B         C              D
  ┌────────┬────────┬──────────────┬─────────┐
1 │ Value1 │ Value2 │ Formula      │ Result  │
  ├────────┼────────┼──────────────┼─────────┤
2 │ 100    │ 100    │ =A2=B2       │ TRUE    │
  ├────────┼────────┼──────────────┼─────────┤
3 │ 100    │ 50     │ =A3>B3       │ TRUE    │
  ├────────┼────────┼──────────────┼─────────┤
4 │ 75     │ 100    │ =A4<B4       │ TRUE    │
  ├────────┼────────┼──────────────┼─────────┤
5 │ 50     │ 50     │ =A5>=B5      │ TRUE    │
  ├────────┼────────┼──────────────┼─────────┤
6 │ 25     │ 30     │ =A6<=B6      │ TRUE    │
  ├────────┼────────┼──────────────┼─────────┤
7 │ 100    │ 50     │ =A7<>B7      │ TRUE    │
  └────────┴────────┴──────────────┴─────────┘
```

### Common Use Case: Conditional Logic

Comparison operators are typically used with IF functions:

```
     A         B              C
  ┌────────┬──────────┬─────────────────────┐
1 │ Score  │ Grade    │ Formula             │
  ├────────┼──────────┼─────────────────────┤
2 │ 85     │ Pass     │ =IF(A2>=60,"Pass","Fail") │
  ├────────┼──────────┼─────────────────────┤
3 │ 45     │ Fail     │ =IF(A3>=60,"Pass","Fail") │
  └────────┴──────────┴─────────────────────┘
```

(IF function covered in detail in File 09)

---

## Text Concatenation Operator

The ampersand `&` joins (concatenates) text strings together.

### Basic Syntax
```
=Text1 & Text2
```

### Examples

**Combining text:**
```
="Hello" & " " & "World"

Result: Hello World
```

**Combining cells:**
```
     A         B              C
  ┌────────┬──────────┬─────────────────┐
1 │ First  │ Last     │ Full Name       │
  ├────────┼──────────┼─────────────────┤
2 │ John   │ Smith    │ =A2&" "&B2      │
  └────────┴──────────┴─────────────────┘

Result: John Smith
```

**Combining text and numbers:**
```
     A              B
  ┌───────────┬─────────────────────┐
1 │ Sales     │ Message             │
  ├───────────┼─────────────────────┤
2 │ 50000     │ ="Total: $"&A2      │
  └───────────┴─────────────────────┘

Result: Total: $50000
```

**Creating email addresses:**
```
     A         B              C
  ┌────────┬──────────┬─────────────────────────┐
1 │ User   │ Domain   │ Email                   │
  ├────────┼──────────┼─────────────────────────┤
2 │ jsmith │ company  │ =A2&"@"&B2&".com"       │
  └────────┴──────────┴─────────────────────────┘

Result: jsmith@company.com
```

### Concatenation with Line Breaks

Use `CHAR(10)` for line breaks within cells:

```
=A2&CHAR(10)&B2

To see line breaks:
1. Enable "Wrap Text" in Home tab
2. Or double-click cell border to auto-fit
```

**Example:**
```
     A              B                C
  ┌──────────┬──────────────┬─────────────────────────┐
1 │ Name     │ Address      │ Full Address            │
  ├──────────┼──────────────┼─────────────────────────┤
2 │ John     │ 123 Main St  │ =A2&CHAR(10)&B2         │
  └──────────┴──────────────┴─────────────────────────┘

Result (with Wrap Text enabled):
John
123 Main St
```

---

## Parentheses for Clarity and Control

Use parentheses to:
1. **Override** order of operations
2. **Clarify** complex formulas
3. **Group** related calculations

### Example 1: Forcing Addition First
```
❌ Without parentheses:
=A1+B1*C1
(Multiplication happens first)

✅ With parentheses:
=(A1+B1)*C1
(Addition happens first)
```

### Example 2: Complex Percentage
```
     A         B         C              D
  ┌────────┬────────┬──────────┬───────────────────┐
1 │ Price  │ Tax    │ Discount │ Final             │
  ├────────┼────────┼──────────┼───────────────────┤
2 │ 100    │ 8%     │ 10%      │ =A2*(1+B2)*(1-C2) │
  └────────┴────────┴──────────┴───────────────────┘

Step by step:
(1+B2)         → 1+0.08 = 1.08
(1-C2)         → 1-0.10 = 0.90
A2*1.08*0.90   → 100*1.08*0.90 = 97.20

Result: 97.20
```

### Example 3: Average of Ratios
```
     A         B         C                      D
  ┌────────┬────────┬──────────────────┬────────────────┐
1 │ Sales1 │ Cost1  │ Sales2 | Cost2   │ Avg Margin     │
  ├────────┼────────┼──────────────────┼────────────────┤
2 │ 1000   │ 600    │ 2000   │ 1100    │ =((A2-B2)/A2+(C2-D2)/C2)/2 │
  └────────┴────────┴──────────────────┴────────────────┘

Better formatted:
=(  (A2-B2)/A2  +  (C2-D2)/C2  ) / 2

Calculates:
Margin1: (1000-600)/1000 = 0.40 (40%)
Margin2: (2000-1100)/2000 = 0.45 (45%)
Average: (0.40+0.45)/2 = 0.425 (42.5%)
```

---

## Working with Negative Numbers

### Subtraction vs Negative
```
=10-5     → Subtraction (Result: 5)
=-5       → Negative number (Result: -5)
=10+-5    → Add negative (Result: 5)
```

### Negating a Cell Value
```
     A         B
  ┌────────┬──────────┐
1 │ Value  │ Opposite │
  ├────────┼──────────┤
2 │ 100    │ =-A2     │
  └────────┴──────────┘

Result: -100
```

### Absolute Value
Use `ABS()` function to remove negative sign:
```
     A         B
  ┌────────┬──────────┐
1 │ Value  │ Absolute │
  ├────────┼──────────┤
2 │ -100   │ =ABS(A2) │
  └────────┴──────────┘

Result: 100
```

---

## Common Formula Patterns

### Pattern 1: Calculate Total
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ Item1  │ Item2  │ Total    │
  ├────────┼────────┼──────────┤
2 │ 100    │ 200    │ =A2+B2   │
  └────────┴────────┴──────────┘
```

### Pattern 2: Calculate Percentage
```
     A         B              C
  ┌────────┬──────────┬─────────────┐
1 │ Part   │ Whole    │ Percentage  │
  ├────────┼──────────┼─────────────┤
2 │ 25     │ 200      │ =A2/B2      │
  └────────┴──────────┴─────────────┘

Result: 0.125 (Format as percentage: 12.5%)
```

### Pattern 3: Calculate Change
```
     A         B              C
  ┌────────┬──────────┬─────────────────┐
1 │ Old    │ New      │ Change          │
  ├────────┼──────────┼─────────────────┤
2 │ 100    │ 150      │ =B2-A2          │
  └────────┴──────────┴─────────────────┘

Result: 50
```

### Pattern 4: Calculate Percent Change
```
     A         B              C
  ┌────────┬──────────┬─────────────────────┐
1 │ Old    │ New      │ % Change            │
  ├────────┼──────────┼─────────────────────┤
2 │ 100    │ 150      │ =(B2-A2)/A2         │
  └────────┴──────────┴─────────────────────┘

Result: 0.50 (Format as percentage: 50%)
```

### Pattern 5: Calculate Weighted Average
```
     A         B         C
  ┌────────┬────────┬──────────────────────┐
1 │ Value  │ Weight │ Formula              │
  ├────────┼────────┼──────────────────────┤
2 │ 90     │ 0.3    │                      │
  ├────────┼────────┼──────────────────────┤
3 │ 85     │ 0.7    │                      │
  ├────────┼────────┼──────────────────────┤
4 │ Result │        │ =A2*B2+A3*B3         │
  └────────┴────────┴──────────────────────┘

Calculation:
90*0.3 = 27
85*0.7 = 59.5
Total: 86.5
```

---

## Real-World Example: Invoice Calculator

**Scenario:** Calculate invoice total with tax and discount.

```
     A                   B
  ┌──────────────────┬─────────┐
1 │ Description      │ Amount  │
  ├──────────────────┼─────────┤
2 │ Subtotal         │ 1000    │
  ├──────────────────┼─────────┤
3 │ Discount %       │ 10%     │
  ├──────────────────┼─────────┤
4 │ Discount Amount  │ =B2*B3  │
  ├──────────────────┼─────────┤
5 │ After Discount   │ =B2-B4  │
  ├──────────────────┼─────────┤
6 │ Tax %            │ 8.5%    │
  ├──────────────────┼─────────┤
7 │ Tax Amount       │ =B5*B6  │
  ├──────────────────┼─────────┤
8 │ Total            │ =B5+B7  │
  └──────────────────┴─────────┘

Results:
B4: 100 (discount)
B5: 900 (after discount)
B7: 76.50 (tax)
B8: 976.50 (total)
```

**Alternative (Single Formula):**
```
Total = =B2*(1-B3)*(1+B6)

Breakdown:
(1-B3)    → 1-0.10 = 0.90 (apply discount)
(1+B6)    → 1+0.085 = 1.085 (add tax)
B2*0.90*1.085 → 1000*0.90*1.085 = 976.50
```

---

## Real-World Example: Loan Payment Calculator

**Scenario:** Calculate monthly payment components.

```
     A                      B
  ┌─────────────────────┬──────────┐
1 │ Loan Amount         │ 200000   │
  ├─────────────────────┼──────────┤
2 │ Annual Rate %       │ 4.5%     │
  ├─────────────────────┼──────────┤
3 │ Monthly Rate        │ =B2/12   │
  ├─────────────────────┼──────────┤
4 │ Loan Term (months)  │ 360      │
  ├─────────────────────┼──────────┤
5 │ Monthly Payment     │ (use PMT function) │
  ├─────────────────────┼──────────┤
6 │ Total Paid          │ =B5*B4   │
  ├─────────────────────┼──────────┤
7 │ Total Interest      │ =B6-B1   │
  └─────────────────────┴──────────┘

Results (approximation):
B3: 0.00375 (monthly rate)
B5: ~1013 (monthly payment using PMT)
B6: 364,680 (total paid)
B7: 164,680 (interest paid)
```

---

## Common Mistakes with Formulas

### Mistake 1: Forgetting the Equals Sign
```
❌ Wrong: SUM(A1:A10)
✅ Right: =SUM(A1:A10)
```

### Mistake 2: Spaces in Formulas
```
❌ Wrong: = A1 + B1
✅ Right: =A1+B1

Spaces are usually ignored but can cause issues
```

### Mistake 3: Wrong Order of Operations
```
❌ Wrong: =10+5*2 expecting 30
✅ Right: =(10+5)*2 to get 30

Remember PEMDAS!
```

### Mistake 4: Division by Zero
```
❌ Error: =100/0
Result: #DIV/0!

✅ Prevention: =IF(B2=0,0,A2/B2)
(Covered more in File 09)
```

### Mistake 5: Mixing Data Types
```
❌ Confusing: ="100"+50
Result: 150 (Excel converts text to number)

⚠️ Unpredictable: ="abc"+50
Result: #VALUE! error
```

### Mistake 6: Text Not in Quotes
```
❌ Wrong: =A1&Hello
✅ Right: =A1&"Hello"

Text must be in quotes unless it's a cell reference
```

### Mistake 7: Incorrect Percent Calculation
```
❌ Wrong: =A2*15% thinking it adds 15%
Result: Calculates 15% of A2

✅ Right: =A2*1.15 (to add 15%)
✅ Or: =A2*(1+15%)
```

---

## Formula Auditing Tips

### Show Formulas Instead of Results

**Toggle formula display:**
- Press `Ctrl + `` (backtick, usually under Esc key)
- Or: Formulas Tab → Show Formulas

**Before:**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ 100    │ 50     │ 150      │
  └────────┴────────┴──────────┘
```

**After (showing formulas):**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ 100    │ 50     │ =A1+B1   │
  └────────┴────────┴──────────┘
```

### Trace Precedents and Dependents

**Formulas Tab → Trace Precedents:**
Shows which cells feed into the formula

**Formulas Tab → Trace Dependents:**
Shows which cells depend on the current cell

**Visual:**
```
     A         B         C
  ┌────────┬────────┬──────────┐
1 │ 100 ───┐        │          │
  ├────────┤ └─────>│ =A1+B1   │
2 │ 50  ───┘        │          │
  └────────┴────────┴──────────┘

Arrows show dependencies
```

### Evaluate Formula Step-by-Step

**Formulas Tab → Evaluate Formula (Desktop only)**

Shows each step of calculation:
```
Formula: =(10+5)*2

Step 1: =(15)*2
Step 2: =30
```

---

## Best Practices for Writing Formulas

### 1. Keep Formulas Simple
```
❌ Avoid: =(A1+B1+C1)*(D1-E1)/(F1+G1)-H1*I1
✅ Better: Break into multiple cells

Cell J1: =A1+B1+C1
Cell K1: =D1-E1
Cell L1: =F1+G1
Cell M1: =J1*K1/L1-H1*I1
```

### 2. Use Cell References, Not Values
```
❌ Avoid: =1000*0.15
✅ Better: =A2*B2
(Put 1000 in A2, 0.15 in B2)
```

### 3. Use Meaningful Cell Locations
```
✅ Good: Put constants at top of sheet
     Tax rate in B1
     Discount in B2
     All formulas reference these cells
```

### 4. Add Comments for Complex Formulas
```
Right-click cell → New Note (or Insert Comment)
"Calculates net profit: Revenue - COGS - OpEx"
```

### 5. Use Parentheses for Clarity
```
❌ Hard to read: =A1+B1*C1-D1/E1
✅ Clearer: =(A1+(B1*C1))-(D1/E1)

Even if not mathematically necessary
```

### 6. Test with Known Values
```
✅ Use simple test data first:
     If 10+5 should equal 15, test with those values
     Then replace with cell references
```

---

## Keyboard Shortcuts for Formulas

| Shortcut | Action |
|----------|--------|
| `=` | Start a formula |
| `Ctrl + `` | Toggle show formulas |
| `F2` | Edit active cell |
| `Esc` | Cancel formula entry |
| `Ctrl + Shift + Enter` | Array formula (advanced) |
| `Alt + =` | AutoSum |
| `Ctrl + '` | Copy formula from cell above |
| `Ctrl + Shift + "` | Copy value from cell above |
| `F9` | Calculate all sheets (Desktop) |
| `Shift + F9` | Calculate active sheet (Desktop) |

---

## Quick Reference: Operator Precedence

**From highest to lowest priority:**

```
1. ( )              Parentheses
2. -                Negation (negative number)
3. %                Percent
4. ^                Exponentiation
5. * /              Multiplication and Division
6. + -              Addition and Subtraction
7. &                Concatenation
8. = < > <= >= <>   Comparison
```

**Memory aid:** **P**lease **E**xcuse **M**y **D**ear **A**unt **S**ally

---

## What to PRACTICE vs MEMORIZE

### Memorize
- All formulas start with `=`
- Basic operators: `+ - * / ^`
- Order of operations: PEMDAS
- Use `&` for text concatenation
- Use `( )` to control calculation order
- Comparison operators return TRUE/FALSE: `= > < >= <= <>`

### Practice Deeply
- Writing simple arithmetic formulas
- Using parentheses to override order of operations
- Combining multiple operators in one formula
- Building formulas with cell references (not hardcoded values)
- Creating percentage calculations
- Concatenating text and numbers
- Testing formulas with known values
- Reading and understanding complex formulas
- Breaking complex calculations into steps
- Identifying and fixing formula errors

---

## Practice Exercises

### Exercise 1: Sales Tax Calculator
Create formulas to calculate:
- Subtotal: Sum of items
- Tax: Subtotal × Tax rate
- Total: Subtotal + Tax

### Exercise 2: Grade Calculator
Given test scores and weights:
- Calculate weighted scores
- Sum for final grade
- Calculate percentage

### Exercise 3: Profit Margin
Calculate:
- Profit = Revenue - Costs
- Profit Margin % = Profit / Revenue
- Test with different values

### Exercise 4: Temperature Converter
Create formulas to convert:
- Celsius to Fahrenheit: `=(C*9/5)+32`
- Fahrenheit to Celsius: `=(F-32)*5/9`

### Exercise 5: Compound Interest
Calculate final amount:
- Formula: `=Principal*(1+Rate)^Years`
- Try different rates and time periods

---

## Troubleshooting Formula Errors

### Error: Formula Not Calculating

**Problem:** Formula shows as text (you see `=A1+B1` not the result)

**Causes:**
1. Cell formatted as Text
2. Space or apostrophe before `=`
3. Leading apostrophe: `'=A1+B1`

**Solutions:**
```
✅ Change cell format to General or Number
✅ Remove any characters before =
✅ Retype the formula
✅ Click cell, press F2, then Enter
```

### Error: #DIV/0!

**Problem:** Division by zero

```
❌ =A1/B1 where B1 is 0 or empty

✅ Prevention:
=IF(B1=0,"N/A",A1/B1)
or
=IFERROR(A1/B1,0)
```

### Error: #VALUE!

**Problem:** Wrong data type in calculation

```
❌ =100+A1 where A1 contains "ABC"

✅ Check: Is the cell actually a number?
✅ Convert text to number: =VALUE(A1)
```

### Error: #NAME?

**Problem:** Excel doesn't recognize text in formula

**Common causes:**
1. Misspelled function name: `=SOM(A1:A10)` instead of `=SUM(A1:A10)`
2. Missing quotes around text: `=A1&Hello` instead of `=A1&"Hello"`
3. Reference to undefined name

**Solutions:**
```
✅ Check spelling of functions
✅ Put text in quotes: "text"
✅ Verify named ranges exist
```

### Error: #REF!

**Problem:** Invalid cell reference

**Causes:**
1. Referenced cell was deleted
2. Copy/paste broke references
3. Invalid cross-sheet reference

**Solution:**
```
✅ Check formula for broken references
✅ Update to valid cell addresses
✅ Undo deletion if possible
```

### Error: Circular Reference Warning

**Problem:** Formula refers to itself

```
❌ Cell A1: =A1+10
❌ Cell A1: =SUM(A1:A10) (includes itself)

✅ Fix: Remove self-reference
✅ Or restructure calculation logic
```

---

## Advanced Formula Techniques (Preview)

### Nested Formulas

Formulas inside formulas:
```
=IF(A1>100,A1*0.15,A1*0.10)

Breakdown:
- IF checks condition
- Returns different calculations based on result
```

### Array Formulas (Dynamic Arrays)

Perform calculations on multiple values:
```
=A1:A10*2

Returns array of results (Excel 365)
```

### Named Formulas

Assign names to formulas for reuse:
```
Define: TaxRate = 0.085
Use: =Amount*TaxRate
```

(These topics covered in later files)

---

## Real-World Example: Budget Tracker

**Complete monthly budget with formulas:**

```
     A                B            C            D
  ┌──────────────┬───────────┬───────────┬─────────────┐
1 │ Category     │ Budget    │ Actual    │ Difference  │
  ├──────────────┼───────────┼───────────┼─────────────┤
2 │ Rent         │ 1500      │ 1500      │ =B2-C2      │
  ├──────────────┼───────────┼───────────┼─────────────┤
3 │ Food         │ 500       │ 550       │ =B3-C3      │
  ├──────────────┼───────────┼───────────┼─────────────┤
4 │ Transport    │ 200       │ 180       │ =B4-C4      │
  ├──────────────┼───────────┼───────────┼─────────────┤
5 │ Entertainment│ 300       │ 350       │ =B5-C5      │
  ├──────────────┼───────────┼───────────┼─────────────┤
6 │ Utilities    │ 150       │ 145       │ =B6-C6      │
  ├──────────────┼───────────┼───────────┼─────────────┤
7 │ TOTAL        │ =SUM(B2:B6)│=SUM(C2:C6)│=B7-C7      │
  ├──────────────┼───────────┼───────────┼─────────────┤
8 │ % of Budget  │           │           │=C7/B7       │
  └──────────────┴───────────┴───────────┴─────────────┘

Results:
D2: 0 (on budget)
D3: -50 (over budget)
D4: 20 (under budget)
D5: -50 (over budget)
D6: 5 (under budget)
D7: -75 (total over budget)
D8: 1.0288 (102.88% of budget)
```

**Enhanced with formatting:**
```
Column D: Conditional formatting
  - Green if positive (under budget)
  - Red if negative (over budget)

Column D8: Format as percentage
  - Shows: 102.88%
```

---

## Real-World Example: ROI Calculator

**Calculate Return on Investment:**

```
     A                      B
  ┌─────────────────────┬──────────┐
1 │ Initial Investment  │ 10000    │
  ├─────────────────────┼──────────┤
2 │ Final Value         │ 13500    │
  ├─────────────────────┼──────────┤
3 │ Gain/Loss           │ =B2-B1   │
  ├─────────────────────┼──────────┤
4 │ ROI %               │ =B3/B1   │
  ├─────────────────────┼──────────┤
5 │ Years Held          │ 3        │
  ├─────────────────────┼──────────┤
6 │ Annual Return %     │ =(B2/B1)^(1/B5)-1 │
  └─────────────────────┴──────────┘

Results:
B3: 3500 (profit)
B4: 0.35 or 35% (total ROI)
B6: 0.1055 or 10.55% (annual return)

Formula B6 explained:
(B2/B1)      → 13500/10000 = 1.35
^(1/B5)      → 1.35^(1/3) = 1.1055
-1           → 1.1055-1 = 0.1055
```

---

## Tips for Learning Formulas

### 1. Start Simple
```
Begin with: =A1+B1
Then build to: =(A1+B1)*C1
Then add: =(A1+B1)*C1*(1-D1)
```

### 2. Use the Formula Bar
- Click cell to see formula in formula bar
- Edit formulas in formula bar for better visibility
- Easier to see long formulas

### 3. Color-Coded References
Excel color-codes cell references while editing:
```
=A1+B1
  │  │
  │  └─ Blue
  └─ Red

Corresponding cells highlighted in same colors
```

### 4. Learn from Examples
- Download sample workbooks
- Study formulas in templates
- Reverse-engineer existing spreadsheets

### 5. Test Incrementally
```
Build formula step by step:
Step 1: =A1
Step 2: =A1*B1
Step 3: =A1*B1+C1
Step 4: =(A1*B1+C1)*D1

Test after each step
```

### 6. Use Comments and Notes
Document your formulas:
```
Right-click → Insert Note
"Calculates monthly payment with 15% down payment factored in"
```

---

## Formula Building Workflow

### Step 1: Plan the Calculation
Write out the logic in plain English:
```
"I need to calculate net profit by taking revenue,
subtracting cost of goods sold, subtracting operating
expenses, and subtracting taxes"
```

### Step 2: Identify the Cells
Map values to cell references:
```
Revenue: A2
COGS: B2
OpEx: C2
Taxes: D2
Result goes in: E2
```

### Step 3: Write the Formula
```
=A2-B2-C2-D2
```

### Step 4: Test with Known Values
```
Use simple numbers:
A2: 100
B2: 30
C2: 20
D2: 10

Expected result: 40
If formula shows 40, it works!
```

### Step 5: Apply to Real Data
```
Replace test values with real data
Copy formula down as needed
```

### Step 6: Document (if complex)
```
Add comment explaining the formula logic
Add labels to make spreadsheet self-documenting
```

---

## Common Formula Use Cases

### Financial
- `=Revenue-Costs` (Profit)
- `=(New-Old)/Old` (Percent change)
- `=Payment*Periods` (Total paid)
- `=Principal*(1+Rate)^Years` (Compound interest)

### Academic
- `=Total/Count` (Average)
- `=Points/MaxPoints` (Percentage score)
- `=Score1*Weight1+Score2*Weight2` (Weighted average)

### Business
- `=Units*Price` (Revenue)
- `=(Price-Cost)/Price` (Margin)
- `=Sales*Commission%` (Commission)
- `=Hours*Rate` (Wages)

### Personal
- `=Income-Expenses` (Savings)
- `=Expense/Income` (Expense ratio)
- `=Amount/Months` (Monthly budget)
- `=Miles/Gallons` (MPG)

---

## Formula Audit Checklist

Before finalizing formulas, check:

- [ ] Does the formula start with `=`?
- [ ] Are all cell references correct?
- [ ] Are parentheses balanced?
- [ ] Is order of operations correct?
- [ ] Are text values in quotes?
- [ ] Do percentage calculations work as expected?
- [ ] Does the formula copy correctly to other cells?
- [ ] Are constant values referenced with `# Basic Formulas and Operators

---

## Next Step

After this file, we move to:

**`04-essential-functions.md`**
- SUM, AVERAGE, COUNT, COUNTA
- MIN, MAX, MEDIAN, MODE
- ROUND, ROUNDUP, ROUNDDOWN
- INT, ABS, MOD
- Function syntax and arguments
- Nesting functions
- Common function errors
