# What-If Analysis

This file covers Excel's What-If Analysis tools - powerful features for exploring scenarios, finding optimal solutions, and answering "what if" questions about your data and models.

---

## What is What-If Analysis?

**What-If Analysis** = Tools that let you test different values in formulas to see how changes affect results.

### The Core Question

```
"What if I change this input?
 How will it affect the outcome?"
```

### Purpose

```
✅ Explore different scenarios
✅ Find required inputs for desired outputs
✅ Compare multiple alternatives
✅ Optimize decisions with constraints
✅ Perform sensitivity analysis
✅ Model uncertainty and risk
```

### Visual Concept

```
┌─────────────────────────────────────────────────┐
│           WHAT-IF ANALYSIS FLOW                 │
│                                                 │
│  Input Values    →    Formula    →    Result   │
│  ┌──────────┐        ┌────────┐      ┌──────┐ │
│  │ Price    │        │ Profit │      │ $$$  │ │
│  │ Quantity │   ───> │  =     │ ───> │      │ │
│  │ Cost     │        │Formula │      │      │ │
│  └──────────┘        └────────┘      └──────┘ │
│       ↕                                   ↕     │
│   Change these              See impact here    │
└─────────────────────────────────────────────────┘
```

---

## Four What-If Analysis Tools

### Overview

**Data Tab → Forecast section → What-If Analysis dropdown:**

```
┌────────────────────────────────┐
│ What-If Analysis          ▼    │
├────────────────────────────────┤
│ Scenario Manager               │
│ Goal Seek                      │
│ Data Table                     │
└────────────────────────────────┘
```

### Tool Comparison

| Tool | Purpose | When to Use |
|------|---------|-------------|
| **Goal Seek** | Find input for target output | "What price gives $10K profit?" |
| **Data Table** | Show multiple results | "Show profit at prices $10-$50" |
| **Scenario Manager** | Save & compare scenarios | "Best case vs Worst case" |
| **Solver** | Optimize with constraints | "Max profit with resource limits" |

### Quick Decision Guide

```
Question Type → Tool to Use

"What input gives me X result?" → Goal Seek
"Show results for input range?" → Data Table
"Compare named scenarios?" → Scenario Manager
"What's the optimal solution?" → Solver
```

---

## Goal Seek

### What is Goal Seek?

**Goal Seek** = Finds the input value needed to achieve a specific goal.

**You provide:**
- Target result (what you want)
- Formula cell (where result is)
- Input cell (what to change)

**Excel calculates:**
- Input value that produces target result

### Visual Concept

```
┌─────────────────────────────────────────┐
│         GOAL SEEK LOGIC                 │
│                                         │
│  Want: Profit = $10,000                 │
│                                         │
│  Formula: =Revenue - Costs              │
│           =(Price × Qty) - Costs        │
│                                         │
│  Question: What Price achieves this?    │
│                                         │
│  Excel tries different Prices until:    │
│  Profit = $10,000                       │
│                                         │
│  Answer: Price = $45.50                 │
└─────────────────────────────────────────┘
```

### Example 1: Break-Even Analysis

**Scenario:** Find break-even price

**Setup:**
```
Cell A1: Price              $0  ← Input to find
Cell A2: Quantity           1000
Cell A3: Fixed Costs        $5,000
Cell A4: Variable Cost/Unit $10

Cell A5: Revenue     =A1*A2      (Price × Quantity)
Cell A6: Total Costs =A3+(A4*A2) (Fixed + Variable)
Cell A7: Profit      =A5-A6      ← Target cell
```

**Goal:** Find price where Profit = $0 (break-even)

**Steps:**
```
1. Data Tab → What-If Analysis → Goal Seek
2. Set cell: A7 (Profit)
3. To value: 0
4. By changing cell: A1 (Price)
5. Click OK
```

**Goal Seek Dialog:**
```
┌────────────────────────────────┐
│ Goal Seek                      │
├────────────────────────────────┤
│ Set cell:    [$A$7]            │
│              (Profit)          │
│                                │
│ To value:    [0]               │
│              (Break-even)      │
│                                │
│ By changing cell: [$A$1]       │
│                   (Price)      │
│                                │
│ [OK] [Cancel]                  │
└────────────────────────────────┘
```

**Result:**
```
┌─────────────────────────────────┐
│ Goal Seek Status                │
├─────────────────────────────────┤
│ Goal Seeking with Cell A7       │
│ found a solution.               │
│                                 │
│ Target value: 0                 │
│ Current value: 0                │
│                                 │
│ [OK] [Cancel]                   │
└─────────────────────────────────┘

Cell A1 now shows: $15 (break-even price)
```

### Example 2: Loan Payment

**Scenario:** Find loan amount for $500 monthly payment

**Setup:**
```
Cell B1: Loan Amount    $0      ← Input to find
Cell B2: Annual Rate    5%
Cell B3: Years          30
Cell B4: Periods        360     (30 × 12)
Cell B5: Monthly Rate   0.417%  (5% / 12)

Cell B6: Payment =PMT(B5,B4,-B1) ← Target cell
```

**Goal:** Find loan amount where Payment = $500

**Steps:**
```
1. Goal Seek
2. Set cell: B6 (Payment)
3. To value: 500
4. By changing: B1 (Loan Amount)
5. OK

Result: Loan Amount = $93,279
```

### Example 3: Sales Target

**Scenario:** Units needed to reach revenue goal

**Setup:**
```
Cell C1: Units Sold     0    ← Input to find
Cell C2: Price per Unit $25
Cell C3: Revenue        =C1*C2  ← Target cell
```

**Goal:** Find units for $50,000 revenue

**Steps:**
```
1. Goal Seek
2. Set cell: C3
3. To value: 50000
4. By changing: C1
5. OK

Result: Units Sold = 2,000
```

### Goal Seek Limitations

```
❌ Only one input cell (can't change multiple)
❌ Only one target value at a time
❌ Doesn't save scenarios
❌ May not find solution if none exists
❌ Finds first solution (may not be only one)

✅ Quick and simple
✅ Great for single-variable problems
✅ Instant results
```

---

## Data Tables

### What is a Data Table?

**Data Table** = Shows how changing one or two inputs affects formula results.

**Creates table of results:**
- Rows = different input values
- Columns = resulting outputs

**Two types:**
1. **One-Variable:** Change one input
2. **Two-Variable:** Change two inputs

### One-Variable Data Table

**Shows results for multiple values of ONE input**

#### Example: Profit at Different Prices

**Setup:**
```
Model (anywhere on sheet):
A1: Price        $20
A2: Quantity     1000
A3: Cost/Unit    $12
A4: Profit       =(A1-A3)*A2

Data Table (different area):
    D         E
1   Price     Profit
2   $10       =A4  ← Reference to formula
3   $15
4   $20
5   $25
6   $30
7   $35
```

**Note:** Column E2 contains =A4 (references your profit formula)

**Steps:**
```
1. Select entire table (D1:E7)
2. Data Tab → What-If Analysis → Data Table
3. Column input cell: A1 (Price)
   (Because prices are in a column)
4. Leave Row input cell: BLANK
5. OK
```

**Data Table Dialog:**
```
┌────────────────────────────────┐
│ Data Table                     │
├────────────────────────────────┤
│ Row input cell:                │
│ [_______]                      │
│                                │
│ Column input cell:             │
│ [$A$1]  ← Price cell           │
│                                │
│ [OK] [Cancel]                  │
└────────────────────────────────┘
```

**Result:**
```
    D         E
1   Price     Profit
2   $10       -$2,000  ← Calculated
3   $15       $3,000   ← Calculated
4   $20       $8,000   ← Calculated
5   $25       $13,000  ← Calculated
6   $30       $18,000  ← Calculated
7   $35       $23,000  ← Calculated

Each profit calculated with that price!
```

**How it works:**
```
For each price in column D:
1. Excel puts price in cell A1
2. Recalculates profit in A4
3. Displays result in table
4. Repeats for all prices
```

#### Vertical Layout (Column Input)

```
Use when: Input values in a column
Column input cell: The cell Excel should vary
Row input cell: Leave blank

Example above uses this layout
```

#### Horizontal Layout (Row Input)

```
Use when: Input values in a row

Setup:
    D        E      F      G      H      I
1            $10    $15    $20    $25    $30
2   Profit   =A4

Steps:
1. Select D1:I2
2. Data Table
3. Row input cell: A1
4. Column input cell: [blank]
5. OK

Result:
    D        E        F        G        H        I
1            $10      $15      $20      $25      $30
2   Profit   -$2,000  $3,000   $8,000   $13,000  $18,000
```

### Two-Variable Data Table

**Shows results for combinations of TWO inputs**

#### Example: Profit by Price AND Quantity

**Setup:**
```
Model:
A1: Price        $20
A2: Quantity     1000
A3: Cost/Unit    $12
A4: Profit       =(A1-A3)*A2

Data Table:
    D       E       F       G       H
1           500     750     1000    1250  ← Quantities (row)
2   $15     =A4
3   $20               ← Results go here
4   $25
5   $30
↑ Prices (column)
```

**Note:** 
- Cell E2 contains =A4 (formula reference)
- Top row = Quantity values
- First column = Price values

**Steps:**
```
1. Select entire table (D1:H5)
2. Data Tab → What-If Analysis → Data Table
3. Row input cell: A2 (Quantity - values in row)
4. Column input cell: A1 (Price - values in column)
5. OK
```

**Result:**
```
    D       E       F       G       H
1           500     750     1000    1250
2   $15     $1,500  $2,250  $3,000  $3,750
3   $20     $4,000  $6,000  $8,000  $10,000
4   $25     $6,500  $9,750  $13,000 $16,250
5   $30     $9,000  $13,500 $18,000 $22,500

Each cell = Profit at that Price × Quantity combo
```

**How to read:**
```
Cell F3 ($6,000) = Profit when:
- Price = $20 (row 3)
- Quantity = 750 (column F)
```

### Data Table Formatting

**Add clarity:**
```
1. Bold headers
2. Number formatting ($#,##0)
3. Conditional formatting for results
4. Freeze panes for large tables
5. Add title above table
```

**Example with formatting:**
```
┌─────────────────────────────────────────────┐
│ Profit Analysis by Price and Quantity       │
├─────────┬─────────┬─────────┬───────────────┤
│         │   500   │   750   │   1000  │ 1250│
├─────────┼─────────┼─────────┼───────────────┤
│ $15     │ $1,500  │ $2,250  │ $3,000  │...  │
│ $20     │ $4,000  │ $6,000  │ $8,000  │...  │
│ $25     │ $6,500  │ $9,750  │ $13,000 │...  │
│ $30     │ $9,000  │ $13,500 │ $18,000 │...  │
└─────────┴─────────┴─────────┴───────────────┘
```

### Data Table Limitations

```
❌ Can't edit results (recalculates automatically)
❌ Slows down large workbooks (many calculations)
❌ Max 2 variables (one-variable or two-variable)
❌ Single formula output per table

✅ Great for sensitivity analysis
✅ Visual comparison of scenarios
✅ Updates automatically
✅ Easy to create charts from results
```

### Data Table Best Practices

```
✅ Keep model separate from table
✅ Use meaningful input ranges
✅ Format results for readability
✅ Add conditional formatting for insights
✅ Document assumptions
✅ Consider manual calculation mode for large tables
```

---

## Scenario Manager

### What is Scenario Manager?

**Scenario Manager** = Saves different sets of input values as named scenarios.

**Use for:**
- Best case / Worst case / Most likely
- Compare multiple alternatives
- Present different options to stakeholders
- Document assumptions

### Visual Concept

```
┌─────────────────────────────────────────────┐
│         SCENARIO MANAGER CONCEPT            │
│                                             │
│  Same Model, Different Inputs:              │
│                                             │
│  Best Case:     Price=$30, Qty=2000         │
│  Most Likely:   Price=$25, Qty=1500         │
│  Worst Case:    Price=$20, Qty=1000         │
│                                             │
│  Save each scenario with name               │
│  Switch between them instantly              │
│  Create summary report comparing all        │
└─────────────────────────────────────────────┘
```

### Example: Sales Forecast Scenarios

**Setup:**
```
Model:
B1: Units Sold       1000  ← Input 1
B2: Price per Unit   $25   ← Input 2
B3: Growth Rate      5%    ← Input 3
B4: Year 1 Revenue   =B1*B2
B5: Year 2 Revenue   =B4*(1+B3)
B6: Total Revenue    =B4+B5
```

**Create scenarios for different assumptions:**

**Scenario 1: Optimistic**
```
Units: 2000
Price: $30
Growth: 10%
```

**Scenario 2: Realistic**
```
Units: 1500
Price: $25
Growth: 5%
```

**Scenario 3: Pessimistic**
```
Units: 1000
Price: $20
Growth: 2%
```

### Creating Scenarios

**Steps:**
```
1. Data Tab → What-If Analysis → Scenario Manager
2. Click "Add"
3. Scenario name: "Optimistic"
4. Changing cells: B1:B3 (select input cells)
5. Click OK
6. Enter values for this scenario:
   B1: 2000
   B2: 30
   B3: 0.10
7. Click OK
8. Repeat for other scenarios (Add → enter values)
```

**Scenario Manager Dialog:**
```
┌────────────────────────────────────┐
│ Scenario Manager                   │
├────────────────────────────────────┤
│ Scenarios:                         │
│ ┌────────────────────────────────┐│
│ │ Optimistic                     ││
│ │ Realistic                      ││
│ │ Pessimistic                    ││
│ └────────────────────────────────┘│
│                                    │
│ [Show] [Close] [Add...]            │
│ [Delete] [Edit...] [Merge...]      │
│ [Summary...]                       │
│                                    │
│ Changing cells: $B$1:$B$3          │
│ Comment: Created by User           │
└────────────────────────────────────┘
```

**Add Scenario Dialog:**
```
┌────────────────────────────────────┐
│ Add Scenario                       │
├────────────────────────────────────┤
│ Scenario name:                     │
│ [Optimistic___________________]    │
│                                    │
│ Changing cells:                    │
│ [$B$1:$B$3]                        │
│                                    │
│ Comment:                           │
│ [Created by John on 12/26/2024]    │
│                                    │
│ Protection:                        │
│ ☐ Prevent changes                  │
│ ☐ Hide                             │
│                                    │
│ [OK] [Cancel]                      │
└────────────────────────────────────┘
```

**Scenario Values Dialog:**
```
┌────────────────────────────────────┐
│ Scenario Values                    │
├────────────────────────────────────┤
│ Enter values for scenario:         │
│ Optimistic                         │
│                                    │
│ $B$1:  [2000]  Units Sold          │
│ $B$2:  [30]    Price per Unit      │
│ $B$3:  [0.10]  Growth Rate         │
│                                    │
│ [Add] [OK] [Cancel]                │
└────────────────────────────────────┘
```

### Switching Between Scenarios

**Show a scenario:**
```
1. Data Tab → Scenario Manager
2. Select scenario: "Optimistic"
3. Click "Show"
4. Excel updates input cells (B1:B3)
5. Results recalculate automatically
6. Close Scenario Manager
```

**Result:**
```
Optimistic scenario displayed:
B1: 2000 (changed)
B2: $30 (changed)
B3: 10% (changed)
B4: $60,000 (recalculated)
B5: $66,000 (recalculated)
B6: $126,000 (recalculated)
```

### Creating Scenario Summary

**Compare all scenarios at once:**

```
1. Data Tab → Scenario Manager
2. Click "Summary"
3. Choose report type:
   ○ Scenario summary
   ○ Scenario PivotTable report
4. Result cells: B4:B6 (cells to show in report)
5. OK
```

**Scenario Summary Dialog:**
```
┌────────────────────────────────────┐
│ Scenario Summary                   │
├────────────────────────────────────┤
│ Report type:                       │
│ ● Scenario summary                 │
│ ○ Scenario PivotTable report       │
│                                    │
│ Result cells:                      │
│ [$B$4:$B$6]                        │
│                                    │
│ [OK] [Cancel]                      │
└────────────────────────────────────┘
```

**Generated Summary Report:**
```
                    Current   Optimistic  Realistic  Pessimistic
                    Values    
Changing Cells:
$B$1 Units Sold     1000      2000        1500       1000
$B$2 Price          $25       $30         $25        $20
$B$3 Growth         5%        10%         5%         2%

Result Cells:
$B$4 Year 1 Rev     $25,000   $60,000     $37,500    $20,000
$B$5 Year 2 Rev     $26,250   $66,000     $39,375    $20,400
$B$6 Total Rev      $51,250   $126,000    $76,875    $40,400

Notes:
- Current Values column shows values before showing scenarios
- Each scenario column shows results for those inputs
- Created on: 12/26/2024
```

### Scenario Manager Tips

```
✅ Use descriptive names (Optimistic vs Scenario1)
✅ Add comments explaining assumptions
✅ Protect scenarios if sharing workbook
✅ Keep changing cells to reasonable number (< 10)
✅ Include result cells that matter to decision
✅ Update scenarios as assumptions change
✅ Save scenario summary as separate report
```

---

## Solver Add-In

### What is Solver?

**Solver** = Advanced optimization tool that finds best solution given constraints.

**Capabilities:**
- Maximize or minimize a target
- Multiple changing cells (not just 1-2)
- Constraints (limits on inputs/outputs)
- Complex relationships

⚠️ **Note:** Solver is an add-in and must be enabled first.

### Enabling Solver

**Steps:**
```
1. File → Options
2. Add-ins
3. Manage: Excel Add-ins → Go
4. Check ☑ Solver Add-in
5. OK

Solver appears in Data Tab → Analyze section
```

**If Solver not in list:**
- May need to install from Office installation
- Excel Online has limited Solver support

### Solver vs Other Tools

| Feature | Goal Seek | Data Table | Scenario Mgr | Solver |
|---------|-----------|------------|--------------|--------|
| Multiple inputs | ❌ | 2 max | Many | Many |
| Constraints | ❌ | ❌ | ❌ | ✅ |
| Optimization | ❌ | ❌ | ❌ | ✅ |
| Save scenarios | ❌ | ❌ | ✅ | ✅ |
| Complexity | Low | Low | Medium | High |

### Example 1: Maximize Profit

**Scenario:** Maximize profit with limited resources

**Setup:**
```
Products:     A      B      C
Price:        $50    $75    $100
Cost:         $30    $45    $60
Profit/Unit:  $20    $30    $40
Hours/Unit:   2      3      4

Resources available: 100 hours

Decision variables (units to produce):
B1: Product A    0  ← Solver changes
B2: Product B    0  ← Solver changes
B3: Product C    0  ← Solver changes

Calculations:
B5: Total Hours  =B1*2+B2*3+B3*4  ← Constraint
B6: Total Profit =B1*20+B2*30+B3*40  ← Maximize this

Constraint: B5 <= 100 (don't exceed 100 hours)
```

**Solver Settings:**
```
1. Data Tab → Solver
2. Set Objective: B6 (Total Profit)
3. To: Max
4. By Changing Variable Cells: B1:B3
5. Subject to Constraints:
   Add: B5 <= 100 (hours constraint)
   Add: B1:B3 >= 0 (can't produce negative)
   Add: B1:B3 = integer (whole units only)
6. Click Solve
```

**Solver Parameters Dialog:**
```
┌────────────────────────────────────────┐
│ Solver Parameters                      │
├────────────────────────────────────────┤
│ Set Objective: [$B$6]                  │
│ To: ● Max  ○ Min  ○ Value of: [____]  │
│                                        │
│ By Changing Variable Cells:            │
│ [$B$1:$B$3]                            │
│                                        │
│ Subject to the Constraints:            │
│ ┌────────────────────────────────────┐│
│ │ $B$5 <= 100                        ││
│ │ $B$1:$B$3 >= 0                     ││
│ │ $B$1:$B$3 = integer                ││
│ └────────────────────────────────────┘│
│ [Add] [Change] [Delete] [Reset All]    │
│                                        │
│ ☑ Make Unconstrained Variables         │
│   Non-Negative                         │
│                                        │
│ Select a Solving Method:               │
│ [GRG Nonlinear            ▼]           │
│                                        │
│ [Solve] [Close]                        │
└────────────────────────────────────────┘
```

**Result:**
```
Solver found optimal solution:
B1: Product A = 0 units
B2: Product B = 0 units  
B3: Product C = 25 units

Total Hours: 100 (used all available)
Total Profit: $1,000 (maximized!)

Make 25 units of Product C for highest profit
```

### Example 2: Minimize Cost

**Scenario:** Meet nutritional requirements at minimum cost

**Setup:**
```
Foods:        Chicken  Rice  Vegetables
Cost/serving: $3       $1    $2
Calories:     300      200   50
Protein(g):   30       5     3

Requirements (daily minimum):
Calories: >= 2000
Protein: >= 60g

Decision variables (servings):
B1: Chicken servings     0  ← Solver changes
B2: Rice servings        0  ← Solver changes
B3: Vegetable servings   0  ← Solver changes

Calculations:
B5: Total Cost      =B1*3+B2*1+B3*2     ← Minimize
B6: Total Calories  =B1*300+B2*200+B3*50
B7: Total Protein   =B1*30+B2*5+B3*3

Constraints:
B6 >= 2000 (minimum calories)
B7 >= 60 (minimum protein)
B1:B3 >= 0 (can't have negative servings)
```

**Solver Settings:**
```
Set Objective: B5 (Total Cost)
To: Min
By Changing: B1:B3
Constraints:
- B6 >= 2000
- B7 >= 60
- B1:B3 >= 0
Solve
```

**Result:**
```
B1: Chicken = 2 servings
B2: Rice = 7 servings
B3: Vegetables = 0 servings

Total Cost: $13 (minimized!)
Calories: 2000 (met requirement)
Protein: 95g (exceeds requirement)
```

### Solver Solving Methods

**Three algorithms available:**

**1. GRG Nonlinear**
```
Use for: Smooth nonlinear problems
Example: Optimization with multiplication/division
Default choice for most problems
```

**2. Simplex LP**
```
Use for: Linear problems only
Example: All formulas are sums/products by constants
Fastest for linear optimization
```

**3. Evolutionary**
```
Use for: Complex, non-smooth problems
Example: IF statements, discontinuous formulas
Slowest but handles difficult problems
```

### Solver Constraints

**Adding constraints:**
```
In Solver dialog, click "Add":

┌────────────────────────────────┐
│ Add Constraint                 │
├────────────────────────────────┤
│ Cell Reference: [$B$5]         │
│                                │
│ Relationship: [<=  ▼]          │
│   Options: <=, =, >=, int, bin │
│                                │
│ Constraint: [100]              │
│                                │
│ [OK] [Add] [Cancel]            │
└────────────────────────────────┘
```

**Common constraint types:**
```
Cell <= Value     : Don't exceed limit
Cell >= Value     : Meet minimum requirement
Cell = Value      : Must equal exactly
Cell = integer    : Whole numbers only
Cell = binary     : 0 or 1 only (yes/no decisions)
```

### Solver Reports

**After solving, choose reports:**
```
┌────────────────────────────────┐
│ Solver Results                 │
├────────────────────────────────┤
│ ● Keep Solver Solution         │
│ ○ Restore Original Values      │
│                                │
│ Reports:                       │
│ ☑ Answer                       │
│ ☑ Sensitivity                  │
│ ☑ Limits                       │
│                                │
│ [OK] [Cancel] [Save Scenario...]│
└────────────────────────────────┘
```

**Report types:**
- **Answer:** Shows original vs optimal values
- **Sensitivity:** How sensitive to constraint changes
- **Limits:** Upper/lower bounds on variables

### Solver Limitations

```
❌ Can be slow for complex problems
❌ May find local optimum (not global best)
❌ Requires some trial and error
❌ Limited in Excel Online

✅ Very powerful for optimization
✅ Handles many variables and constraints
✅ Saves as scenarios
✅ Multiple solving methods
```

---

## Practical Applications

### Business Applications

**Break-Even Analysis:**
```
Tool: Goal Seek
Question: "What sales volume breaks even?"
Input: Units sold
Target: Profit = $0
```

**Pricing Strategy:**
```
Tool: Data Table
Question: "How do price changes affect revenue?"
Inputs: Price (rows), Market share (columns)
Output: Total revenue
```

**Budget Scenarios:**
```
Tool: Scenario Manager
Scenarios: Conservative, Moderate, Aggressive spending
Inputs: Department budgets
Outputs: Total spend, ROI, savings
```

**Resource Allocation:**
```
Tool: Solver
Objective: Maximize profit
Variables: Production quantities
Constraints: Limited labor, materials, machine time
```

### Financial Applications

**Loan Analysis:**
```
Tool: Data Table
Show: Monthly payment at different rates/terms
Visual: Easy comparison matrix
```

**Investment Portfolio:**
```
Tool: Solver
Objective: Maximize return
Variables: Investment allocations
Constraints: Risk limits, diversification rules
```

**Retirement Planning:**
```
Tool: Scenario Manager
Scenarios: Conservative, Moderate, Aggressive returns
Inputs: Savings rate, return%, inflation
Output: Retirement fund value
```

### Project Management

**Schedule Optimization:**
```
Tool: Solver
Objective: Minimize project duration
Variables: Task start dates
Constraints: Dependencies, resource availability
```

**Cost-Benefit Analysis:**
```
Tool: Data Table
Show: NPV at different discount rates and time periods
Helps: Make go/no-go decisions
```

---

## Best Practices

### Model Setup

```
✅ Keep inputs separate from formulas
✅ Use named ranges for clarity
✅ Color-code inputs (blue) vs outputs (black)
✅ Document assumptions clearly
✅ Test with simple values first
✅ Save backup before running What-If tools
```

**Example layout:**
```
┌──────────────────────────────────┐
│ INPUTS (Blue cells)              │
│ Price:           $25             │
│ Quantity:        1000            │
│ Cost per unit:   $15             │
├──────────────────────────────────┤
│ CALCULATIONS                     │
│ Revenue:   =Price*Quantity       │
│ Costs:     =Cost*Quantity        │
│ Profit:    =Revenue-Costs        │
└──────────────────────────────────┘
```

### Documentation

```
✅ Add text boxes explaining model
✅ Label all scenarios descriptively
✅ Note data sources and assumptions
✅ Include date created/modified
✅ List constraints and limits
```

### Verification

```
✅ Test with known values
✅ Check formulas reference correct cells
✅ Verify constraints make sense
✅ Review Solver reports for issues
✅ Compare results to manual calculations
```

### Performance

```
✅ Use Manual calculation for large Data Tables
✅ Limit Data Table size when possible
✅ Break complex models into sections
✅ Avoid volatile functions in What-If models
✅ Save Solver models separately
```

---

## Sensitivity Analysis

### What is Sensitivity Analysis?

**Sensitivity Analysis** = Testing how sensitive results are to input changes.

**Questions answered:**
```
"Which inputs matter most?"
"How much can inputs vary before decision changes?"
"Where should we focus data collection efforts?"
```

### Using Data Tables for Sensitivity

**Example: Loan Payment Sensitivity**

**Model:**
```
A1: Loan Amount    $200,000
A2: Interest Rate  5%
A3: Years          30
A4: Payment        =PMT(A2/12, A3*12, -A1)
```

**Two-Variable Data Table:**
```
Test payment at different rates and loan amounts

         $150K   $175K   $200K   $225K   $250K
3.0%     $633    $738    $843    $949    $1,054
3.5%     $674    $786    $898    $1,011  $1,123
4.0%     $716    $836    $955    $1,074  $1,194
4.5%     $760    $887    $1,013  $1,140  $1,267
5.0%     $805    $939    $1,074  $1,208  $1,342
5.5%     $851    $993    $1,136  $1,278  $1,420
```

**Insights:**
```
- Each 0.5% rate increase = ~$60/month (for $200K)
- Each $25K loan increase = ~$130/month (at 5%)
- Rate changes have MORE impact than originally thought
- Should lock in rate if expecting increases
```

### Tornado Diagram (Visual Sensitivity)

**Create chart showing impact of each variable:**

**Steps:**
```
1. Create Data Table for each input
2. Calculate range of outputs for each
3. Sort by impact (largest to smallest)
4. Create horizontal bar chart

Visual shows which inputs matter most
```

**Example:**
```
Impact on Profit:
Price       ████████████████████  ±$5,000
Volume      ███████████████       ±$3,500
Cost        ██████████            ±$2,000
Overhead    ████                  ±$800

Focus on Price and Volume!
```

---

## Scenario Planning Best Practices

### Three-Scenario Framework

**Standard approach:**

**Optimistic (Best Case):**
```
Assumptions:
- Best reasonable outcomes
- Favorable conditions
- High end of estimates
- Not wildly unrealistic

Example:
- Sales growth: 15%
- Costs decrease: 5%
- Market share: +3%
```

**Most Likely (Base Case):**
```
Assumptions:
- Expected outcomes
- Normal conditions
- Central estimates
- Your best guess

Example:
- Sales growth: 8%
- Costs stable
- Market share: +1%
```

**Pessimistic (Worst Case):**
```
Assumptions:
- Challenging outcomes
- Unfavorable conditions
- Low end of estimates
- Still plausible

Example:
- Sales growth: 2%
- Costs increase: 3%
- Market share: -1%
```

### Probability-Weighted Scenarios

**Assign probabilities to scenarios:**

```
Scenario      Probability   Result
Optimistic    20%           $150,000
Most Likely   60%           $100,000
Pessimistic   20%           $50,000

Expected Value = (0.20 × 150,000) + 
                 (0.60 × 100,000) + 
                 (0.20 × 50,000)
               = $100,000
```

### Monte Carlo Simulation (Advanced)

**Beyond basic scenarios:**
```
Instead of 3 scenarios, test thousands
Use random values within ranges
Excel add-ins available for this
Produces probability distributions

Example: "95% confidence profit will be $80K-$120K"
```

---

## Troubleshooting What-If Analysis

### Problem: Goal Seek Can't Find Solution

**Error:** "Goal Seek could not find a solution"

**Possible causes:**
```
❌ No mathematical solution exists
❌ Target value outside possible range
❌ Formula doesn't depend on changing cell
❌ Circular reference in formulas
```

**Solutions:**
```
1. Verify formula references changing cell
2. Check if target is achievable
3. Try different starting value in changing cell
4. Simplify formula if too complex
5. Check for circular references (Formulas Tab)
```

### Problem: Data Table Shows #N/A

**Cause:** Formula error or table not set up correctly

**Solutions:**
```
1. Check formula cell contains valid formula
2. Verify row/column input cells correct
3. Ensure changing cells in formula
4. Rebuild table from scratch
5. Check for errors in source formulas
```

### Problem: Data Table Slows Down Workbook

**Cause:** Large tables with many calculations

**Solutions:**
```
1. Switch to Manual calculation:
   Formulas Tab → Calculation Options → Manual
   
2. Press F9 to calculate when needed

3. Reduce table size (fewer rows/columns)

4. Split into multiple smaller tables

5. Consider using Scenario Manager instead
```

### Problem: Solver Won't Run

**Error:** "Solver could not find a feasible solution"

**Possible causes:**
```
❌ Constraints conflict (impossible to satisfy all)
❌ Starting values too far from solution
❌ Wrong solving method selected
```

**Solutions:**
```
1. Review constraints for conflicts
2. Relax constraints slightly
3. Try different solving method
4. Use better starting values
5. Simplify model
6. Check "Options" → increase iterations/time
```

### Problem: Solver Stops Before Optimal

**Message:** "Solver converged to current solution"

**Cause:** Found local optimum, not best overall

**Solutions:**
```
1. Use different starting values
2. Try Evolutionary solving method
3. Add more iterations (Options)
4. Use multistart (Options → GRG Nonlinear)
5. Run Solver multiple times
```

### Problem: Scenarios Don't Update Results

**Cause:** Calculation mode or formula issues

**Solutions:**
```
1. Press F9 to force recalculation
2. Check calculation mode (should be Automatic)
3. Verify formulas reference changing cells
4. Close and reopen Scenario Manager
5. Re-show scenario
```

---

## Quick Reference: What-If Analysis Tools

| Tool | Access Path | Key Dialog Option |
|------|-------------|-------------------|
| **Goal Seek** | Data → What-If Analysis → Goal Seek | Set cell, To value, By changing |
| **Data Table** | Data → What-If Analysis → Data Table | Row/Column input cell |
| **Scenario Manager** | Data → What-If Analysis → Scenario Manager | Add, Show, Summary |
| **Solver** | Data → Solver (if enabled) | Set Objective, By Changing, Constraints |

---

## Quick Reference: When to Use Each Tool

| Question | Tool | Why |
|----------|------|-----|
| "What input gives X result?" | Goal Seek | Single target, single input |
| "Show results across range?" | Data Table | Multiple scenarios at once |
| "Compare 3-5 named scenarios?" | Scenario Manager | Save and compare |
| "What's optimal with limits?" | Solver | Multiple variables + constraints |
| "Which input matters most?" | Data Table | Sensitivity analysis |
| "Test hundreds of scenarios?" | Data Table or VBA | Systematic testing |

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `F9` | Recalculate all |
| `Shift+F9` | Recalculate active sheet |
| `Ctrl+F3` | Name Manager (create named ranges) |
| `Alt+A+W+G` | Goal Seek (quick access) |
| `Alt+A+W+T` | Data Table (quick access) |
| `Alt+A+W+S` | Scenario Manager (quick access) |

---

## Real-World Example: Complete Analysis

### Scenario: Product Launch Decision

**Question:** Should we launch new product?

**Model Setup:**

```
INPUTS:
Units Year 1:        10,000
Price per Unit:      $50
Unit Cost:           $30
Marketing Budget:    $100,000
Growth Rate Y2-5:    20%

CALCULATIONS:
Year 1 Revenue:      =Units*Price
Year 1 Costs:        =(Units*UnitCost)+Marketing
Year 1 Profit:       =Revenue-Costs

5-Year NPV:          [Complex formula with growth]
```

**Analysis 1: Break-Even Price (Goal Seek)**
```
Question: What price breaks even Year 1?
Goal Seek:
- Set cell: Year 1 Profit
- To value: 0
- By changing: Price

Result: $40 = break-even price
Decision: $50 price has $100K cushion
```

**Analysis 2: Sensitivity to Units and Price (Data Table)**
```
Two-variable table:
Rows: Units (5K, 7.5K, 10K, 12.5K, 15K)
Columns: Price ($40, $45, $50, $55, $60)
Result: Year 1 Profit

Insights:
- All combinations profitable except lowest
- Volume more important than price
- $45 at 10K units = $50K profit (acceptable)
```

**Analysis 3: Three Scenarios (Scenario Manager)**
```
Optimistic:
- Units: 15,000
- Price: $55
- Growth: 25%
- 5-Year NPV: $2.5M

Realistic:
- Units: 10,000
- Price: $50
- Growth: 20%
- 5-Year NPV: $1.2M

Pessimistic:
- Units: 7,000
- Price: $45
- Growth: 10%
- 5-Year NPV: $400K

Decision: Even worst case is acceptable
```

**Analysis 4: Optimize Marketing Budget (Solver)**
```
Objective: Maximize 5-Year NPV
By changing: Marketing Budget
Constraints:
- Marketing >= $50K (minimum for awareness)
- Marketing <= $200K (budget limit)
- Must be profitable Year 1

Result: Optimal marketing = $125,000
Increases NPV by $150,000
```

**Final Recommendation:**
```
✅ Launch product
✅ Price at $50
✅ Marketing budget $125,000
✅ Expect realistic scenario most likely
✅ All scenarios show positive return
✅ Break-even price $10 below planned price
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- What-If Analysis = testing different input scenarios
- Goal Seek finds input for specific output
- Data Table shows multiple results in grid
- Scenario Manager saves named scenarios
- Solver optimizes with constraints
- Data Tab → What-If Analysis dropdown
- One-variable table changes one input
- Two-variable table changes two inputs
- Column input cell for values in column
- Row input cell for values in row
- Solver requires add-in (File → Options → Add-ins)
- Sensitivity analysis tests which inputs matter most
- Best/Most Likely/Worst case standard framework
- Can't edit Data Table results directly
- Data Tables recalculate automatically

### Practice Deeply
- Setting up What-If Analysis models
- Using Goal Seek for break-even analysis
- Using Goal Seek for target values
- Creating one-variable Data Tables
- Creating two-variable Data Tables
- Interpreting Data Table results
- Formatting Data Tables for clarity
- Creating scenarios in Scenario Manager
- Naming scenarios descriptively
- Switching between scenarios
- Generating Scenario Summary reports
- Reading Scenario Summary reports
- Enabling Solver add-in
- Setting up Solver problems
- Defining Solver objective (max/min/target)
- Adding Solver constraints
- Running Solver
- Interpreting Solver results
- Choosing appropriate Solver method
- Performing sensitivity analysis with Data Tables
- Building three-scenario models (Optimistic/Realistic/Pessimistic)
- Combining multiple What-If tools in one analysis
- Documenting assumptions clearly
- Verifying What-If results make sense
- Troubleshooting Goal Seek issues
- Troubleshooting Data Table errors
- Troubleshooting Solver convergence problems
- Creating comparison charts from scenarios
- Using What-If Analysis for business decisions

---

## Next Step

After this file, we move to:

**`24-protection-and-security.md`**
- Protecting worksheets
- Protecting workbooks
- Password protection
- Allowing specific users/ranges
- Locking and unlocking cells
- Read-only recommendations
- Digital signatures
- Sharing and collaboration settings
- Protecting formulas while allowing data entry
- Removing protection
