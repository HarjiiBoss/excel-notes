# Date and Time Functions

This file covers Excel's date and time functions that let you work with dates,
calculate durations, extract date parts, and perform calendar calculations.
Understanding how Excel handles dates is crucial for business analytics.

---

## How Excel Stores Dates and Times

**Key Concept:** Excel stores dates as **numbers**, not actual dates.

### Date Serial Numbers

Excel counts days starting from **January 1, 1900** (serial number 1).

```
Visual Timeline:
January 1, 1900  →  Serial number: 1
January 2, 1900  →  Serial number: 2
January 3, 1900  →  Serial number: 3
...
January 1, 2024  →  Serial number: 45292
December 24, 2025 → Serial number: 45645
```

### See the Serial Number
```
     A              B
  ┌────────────┬──────────────┐
1 │ Date       │ Serial #     │
2 │ 1/1/2024   │ =A2+0        │
  └────────────┴──────────────┘
              Format as General
                  ↓
               45292

Adding 0 converts date to number
```

### Why This Matters

**Because dates are numbers, you can do math:**
```
Later Date - Earlier Date = Number of Days

1/15/2024 - 1/1/2024 = 14 days
```

### Time as Decimals

**Times** are stored as **decimal fractions** of a day:

```
12:00 AM (midnight) = 0.0
6:00 AM             = 0.25  (1/4 of day)
12:00 PM (noon)     = 0.5   (1/2 of day)
6:00 PM             = 0.75  (3/4 of day)
11:59 PM            = 0.999...
```

### Date + Time Combined
```
Date:         1/15/2024      = 45306.0
Time:         3:30 PM        = 0.645833...
Combined:     1/15/2024 3:30 PM = 45306.645833
              ↑                    ↑
              Date                 Decimal = time
```

### Visual Representation
```
Serial Number: 45306.75

     45306        .75
       ↓           ↓
   Date part   Time part
  (1/15/2024)   (6:00 PM)
```

---

## TODAY and NOW Functions

### TODAY Function

**Syntax:** `=TODAY()`

**Purpose:** Returns the current date (updates when file opens)

```
     A
  ┌────────────┐
1 │ =TODAY()   │
  └────────────┘
     ↓
  12/24/2025

Changes to current date when file is opened
```

**No arguments needed** - just `=TODAY()`

### NOW Function

**Syntax:** `=NOW()`

**Purpose:** Returns current date AND time (updates when file recalculates)

```
     A
  ┌────────────────────┐
1 │ =NOW()             │
  └────────────────────┘
     ↓
  12/24/2025 10:30 AM

Updates every time Excel recalculates
```

### Difference Between TODAY and NOW

| Function | Returns | Updates When |
|----------|---------|--------------|
| **TODAY** | Date only | File opens |
| **NOW** | Date and time | File recalculates (any change) |

### Example 1: Current Date Header
```
     A
  ┌────────────────────────┐
1 │ ="Report Date: "&TEXT(TODAY(),"mmmm d, yyyy")
  └────────────────────────┘
     ↓
  "Report Date: December 24, 2025"
```

### Example 2: Age Calculation
```
     A              B              C
  ┌────────────┬────────────┬────────────────┐
1 │ Birthdate  │ Today      │ Age            │
2 │ 1/15/1990  │ =TODAY()   │ =(B2-A2)/365.25│
  └────────────┴────────────┴────────────────┘
                               ↓
                             35.9... years

Note: Better methods shown later with DATEDIF
```

### Example 3: Days Until Deadline
```
     A              B              C
  ┌────────────┬────────────┬────────────────┐
1 │ Deadline   │ Today      │ Days Left      │
2 │ 12/31/2025 │ =TODAY()   │ =A2-B2         │
  └────────────┴────────────┴────────────────┘
                               ↓
                              7 days
```

### ⚠️ Important Notes
- TODAY() has no serial number - it's always "now"
- Both functions are **volatile** (recalculate frequently)
- In large workbooks, many NOW/TODAY calls can slow performance
- Use sparingly in large spreadsheets

---

## DATE Function

**Syntax:** `=DATE(year, month, day)`

**Purpose:** Creates a date from individual year, month, day numbers

```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────┐
1 │ Year    │ Month   │ Day     │ Date           │
2 │ 2024    │ 1       │ 15      │ =DATE(A2,B2,C2)│
  └─────────┴─────────┴─────────┴────────────────┘
                                    ↓
                                 1/15/2024
```

### Why Use DATE?

**Instead of typing dates directly:**
```
❌ Problem: ="Payment due: "&A2
If A2 is 1/15/2024, result is: "Payment due: 45306"

✅ Solution: ="Payment due: "&TEXT(A2,"m/d/yyyy")
Result: "Payment due: 1/15/2024"
```

**Or building dates from parts:**
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────┐
1 │ Year    │ Month   │ Day     │ Hire Date      │
2 │ 2024    │ 3       │ 15      │ =DATE(A2,B2,C2)│
  └─────────┴─────────┴─────────┴────────────────┘
                                    ↓
                                 3/15/2024
```

### DATE with Calculations

**Add months to a date:**
```
=DATE(2024, 1+3, 15)  →  4/15/2024
       ↑    ↑
     Year  Jan + 3 months = April
```

**Add years:**
```
=DATE(2024+1, 1, 15)  →  1/15/2025
```

### Example 1: End of Month
```
     A              B
  ┌────────────┬────────────────────────┐
1 │ Date       │ End of Month           │
2 │ 1/15/2024  │ =DATE(YEAR(A2),MONTH(A2)+1,0)
  └────────────┴────────────────────────┘
                  ↓
               1/31/2024

Day 0 = last day of previous month
So Month+1, Day 0 = last day of current month
```

### Example 2: First Day of Year
```
     A              B
  ┌────────────┬────────────────────────┐
1 │ Date       │ First Day of Year      │
2 │ 5/15/2024  │ =DATE(YEAR(A2),1,1)    │
  └────────────┴────────────────────────┘
                  ↓
               1/1/2024
```

### DATE Handles Overflow

Excel automatically adjusts invalid dates:

```
=DATE(2024, 13, 1)   →  1/1/2025  (13th month = Jan next year)
=DATE(2024, 1, 32)   →  2/1/2024  (32nd day of Jan = 1st of Feb)
=DATE(2024, 0, 1)    →  12/1/2023 (Month 0 = Dec previous year)
```

---

## TIME Function

**Syntax:** `=TIME(hour, minute, second)`

**Purpose:** Creates a time from individual hour, minute, second numbers

```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬────────────────┐
1 │ Hour    │ Minute  │ Second  │ Time           │
2 │ 14      │ 30      │ 0       │ =TIME(A2,B2,C2)│
  └─────────┴─────────┴─────────┴────────────────┘
                                    ↓
                                 2:30 PM

Note: 14 = 2:00 PM (24-hour format converts)
```

### Examples
```
=TIME(8, 30, 0)    →  8:30 AM
=TIME(14, 45, 30)  →  2:45:30 PM
=TIME(0, 0, 0)     →  12:00 AM
=TIME(23, 59, 59)  →  11:59:59 PM
```

### Combining DATE and TIME
```
     A          B
  ┌─────────┬────────────────────────────────┐
1 │ Date    │ DateTime                       │
2 │ 1/15/24 │ =A2+TIME(14,30,0)              │
  └─────────┴────────────────────────────────┘
              ↓
           1/15/2024 2:30 PM

Date + Time = DateTime
```

---

## YEAR, MONTH, DAY Functions

**Purpose:** Extract parts of a date

### YEAR Function

**Syntax:** `=YEAR(date)`

```
     A              B
  ┌────────────┬──────────────┐
1 │ Date       │ Year         │
2 │ 1/15/2024  │ =YEAR(A2)    │
  └────────────┴──────────────┘
                  ↓
                2024
```

### MONTH Function

**Syntax:** `=MONTH(date)`

```
     A              B
  ┌────────────┬──────────────┐
1 │ Date       │ Month        │
2 │ 1/15/2024  │ =MONTH(A2)   │
  └────────────┴──────────────┘
                  ↓
                  1

Returns number (1-12), not name
```

### DAY Function

**Syntax:** `=DAY(date)`

```
     A              B
  ┌────────────┬──────────────┐
1 │ Date       │ Day          │
2 │ 1/15/2024  │ =DAY(A2)     │
  └────────────┴──────────────┘
                  ↓
                 15
```

### Example: Breaking Down Dates
```
     A              B            C            D
  ┌────────────┬──────────┬──────────┬──────────┐
1 │ Date       │ Year     │ Month    │ Day      │
2 │ 12/24/2025 │ =YEAR(A2)│ =MONTH(A2)│ =DAY(A2)│
  └────────────┴──────────┴──────────┴──────────┘
                  ↓          ↓          ↓
                2025        12         24
```

### Real-World Example: Fiscal Year
```
     A              B
  ┌────────────┬────────────────────────────────┐
1 │ Date       │ Fiscal Year                    │
2 │ 10/15/2024 │ =IF(MONTH(A2)>=7,YEAR(A2)+1,YEAR(A2))
  └────────────┴────────────────────────────────┘
                  ↓
                2025

If fiscal year starts July 1:
- Oct 2024 → FY 2025
- May 2024 → FY 2024
```

---

## HOUR, MINUTE, SECOND Functions

**Purpose:** Extract parts of a time

### HOUR Function

**Syntax:** `=HOUR(time)`

```
     A              B
  ┌────────────┬──────────────┐
1 │ Time       │ Hour         │
2 │ 2:30 PM    │ =HOUR(A2)    │
  └────────────┴──────────────┘
                  ↓
                 14

Returns 24-hour format (0-23)
```

### MINUTE Function

**Syntax:** `=MINUTE(time)`

```
     A              B
  ┌────────────┬──────────────┐
1 │ Time       │ Minute       │
2 │ 2:30 PM    │ =MINUTE(A2)  │
  └────────────┴──────────────┘
                  ↓
                 30
```

### SECOND Function

**Syntax:** `=SECOND(time)`

```
     A              B
  ┌────────────┬──────────────┐
1 │ Time       │ Second       │
2 │ 2:30:45 PM │ =SECOND(A2)  │
  └────────────┴──────────────┘
                  ↓
                 45
```

### Example: Time Breakdown
```
     A              B          C          D
  ┌────────────┬──────────┬──────────┬──────────┐
1 │ DateTime   │ Hour     │ Minute   │ Second   │
2 │ 1/15/24 2:30:45 PM │ =HOUR(A2)│ =MINUTE(A2)│ =SECOND(A2)
  └────────────┴──────────┴──────────┴──────────┘
                  ↓          ↓          ↓
                 14         30         45
```

---

## WEEKDAY Function

**Syntax:** `=WEEKDAY(date, [return_type])`

**Purpose:** Returns the day of the week as a number

### Return Types

| Return Type | Sunday | Monday | Tuesday | ... | Saturday |
|-------------|--------|--------|---------|-----|----------|
| **1 (default)** | 1 | 2 | 3 | ... | 7 |
| **2** | 7 | 1 | 2 | ... | 6 |
| **3** | 6 | 0 | 1 | ... | 5 |

**Most common:** Type 1 (Sunday = 1) or Type 2 (Monday = 1)

### Examples
```
     A              B                  C
  ┌────────────┬──────────────────┬──────────────┐
1 │ Date       │ Weekday (Type 1) │ Weekday (Type 2)│
2 │ 12/24/2025 │ =WEEKDAY(A2,1)   │ =WEEKDAY(A2,2)  │
  └────────────┴──────────────────┴──────────────┘
   (Tuesday)         ↓                  ↓
                     3                  2

Type 1: Sunday=1, so Tuesday=3
Type 2: Monday=1, so Tuesday=2
```

### Example 1: Is it a Weekend?
```
     A              B
  ┌────────────┬────────────────────────────────┐
1 │ Date       │ Weekend?                       │
2 │ 12/24/2025 │ =IF(OR(WEEKDAY(A2)=1,WEEKDAY(A2)=7),"Yes","No")
  └────────────┴────────────────────────────────┘
                  ↓
                "No"

Weekday 1=Sunday, 7=Saturday
```

### Example 2: Day Name
```
     A              B
  ┌────────────┬────────────────────────────────┐
1 │ Date       │ Day Name                       │
2 │ 12/24/2025 │ =TEXT(A2,"dddd")               │
  └────────────┴────────────────────────────────┘
                  ↓
              "Tuesday"

Or use: =CHOOSE(WEEKDAY(A2),"Sun","Mon","Tue","Wed","Thu","Fri","Sat")
```

### Example 3: Weekday vs Weekend
```
     A              B
  ┌────────────┬────────────────────────────────┐
1 │ Date       │ Type                           │
2 │ 12/24/2025 │ =IF(WEEKDAY(A2,2)<=5,"Weekday","Weekend")
  └────────────┴────────────────────────────────┘
                  ↓
              "Weekday"

Type 2: Mon-Fri = 1-5, Sat-Sun = 6-7
```

---

## EOMONTH Function

**Syntax:** `=EOMONTH(start_date, months)`

**Purpose:** Returns the **End Of MONTH** for a date plus/minus specified months

```
     A              B
  ┌────────────┬──────────────────┐
1 │ Date       │ End of Month     │
2 │ 1/15/2024  │ =EOMONTH(A2,0)   │
  └────────────┴──────────────────┘
                  ↓
               1/31/2024

0 months = same month
```

### Examples
```
Date: 1/15/2024

=EOMONTH(A2, 0)   →  1/31/2024  (end of same month)
=EOMONTH(A2, 1)   →  2/29/2024  (end of next month)
=EOMONTH(A2, -1)  →  12/31/2023 (end of previous month)
=EOMONTH(A2, 3)   →  4/30/2024  (end of 3 months ahead)
```

### Example 1: Contract End Date
```
     A              B          C
  ┌────────────┬─────────┬──────────────────┐
1 │ Start Date │ Months  │ End Date         │
2 │ 1/15/2024  │ 12      │ =EOMONTH(A2,B2)  │
  └────────────┴─────────┴──────────────────┘
                            ↓
                         1/31/2025

12-month contract ending on last day of month
```

### Example 2: Beginning of Next Month
```
     A              B
  ┌────────────┬──────────────────────┐
1 │ Date       │ Start of Next Month  │
2 │ 1/15/2024  │ =EOMONTH(A2,0)+1     │
  └────────────┴──────────────────────┘
                  ↓
               2/1/2024

End of this month + 1 day = first day of next month
```

---

## EDATE Function

**Syntax:** `=EDATE(start_date, months)`

**Purpose:** Returns a date that is N months before/after start date (same day of month)

```
     A              B
  ┌────────────┬──────────────────┐
1 │ Date       │ 3 Months Later   │
2 │ 1/15/2024  │ =EDATE(A2,3)     │
  └────────────┴──────────────────┘
                  ↓
               4/15/2024

Same day (15th), 3 months later
```

### EDATE vs EOMONTH

| Function | Returns |
|----------|---------|
| **EDATE** | Same day of month, N months away |
| **EOMONTH** | Last day of month, N months away |

```
Date: 1/15/2024

=EDATE(A2, 3)    →  4/15/2024  (15th of April)
=EOMONTH(A2, 3)  →  4/30/2024  (last day of April)
```

### Example: Subscription Renewal
```
     A              B              C
  ┌────────────┬────────────┬──────────────────┐
1 │ Start      │ Term(mo)   │ Renewal Date     │
2 │ 1/15/2024  │ 12         │ =EDATE(A2,B2)    │
  └────────────┴────────────┴──────────────────┘
                               ↓
                            1/15/2025
```

---

## WORKDAY and NETWORKDAYS Functions

**Purpose:** Calculate business days (excluding weekends and holidays)

### WORKDAY Function

**Syntax:** `=WORKDAY(start_date, days, [holidays])`

**Purpose:** Returns a date N **workdays** from start date

```
     A              B          C
  ┌────────────┬─────────┬──────────────────┐
1 │ Start      │ Days    │ Due Date         │
2 │ 1/15/2024  │ 10      │ =WORKDAY(A2,B2)  │
  └────────────┴─────────┴──────────────────┘
   (Monday)                  ↓
                          1/29/2024 (Monday)

Skips weekends: adds 10 business days
```

### WORKDAY with Holidays
```
     A              B          C              D
  ┌────────────┬─────────┬────────────┬──────────────────┐
1 │ Start      │ Days    │ Holidays   │ Due Date         │
2 │ 1/15/2024  │ 10      │ 1/20/2024  │ =WORKDAY(A2,B2,C2)
  └────────────┴─────────┴────────────┴──────────────────┘
                                          ↓
                                       1/30/2024

Skips weekend AND holiday
```

### NETWORKDAYS Function

**Syntax:** `=NETWORKDAYS(start_date, end_date, [holidays])`

**Purpose:** Counts **workdays** between two dates

```
     A              B              C
  ┌────────────┬────────────┬──────────────────────┐
1 │ Start      │ End        │ Business Days        │
2 │ 1/1/2024   │ 1/31/2024  │ =NETWORKDAYS(A2,B2)  │
  └────────────┴────────────┴──────────────────────┘
                               ↓
                              23

January 2024 has 23 workdays
```

### Real-World Example: Project Timeline
```
     A              B          C              D
  ┌────────────┬─────────┬────────────┬──────────────────┐
1 │ Start      │ Duration│ Holidays   │ Completion       │
2 │ 1/2/2024   │ 15      │ 1/15/2024  │ =WORKDAY(A2,B2,C2)
3 │            │         │            │                  │
4 │ Actual Days: =NETWORKDAYS(A2,D2,C2)                 │
  └────────────┴─────────┴────────────┴──────────────────┘

Calculates realistic project timeline excluding weekends/holidays
```

---

## DATEDIF Function

**Syntax:** `=DATEDIF(start_date, end_date, unit)`

**Purpose:** Calculates difference between dates in various units

**⚠️ Important:** This is an **undocumented** function (doesn't appear in autocomplete) but works!

### Units

| Unit | Returns |
|------|---------|
| **"Y"** | Complete years |
| **"M"** | Complete months |
| **"D"** | Complete days |
| **"YM"** | Months, ignoring years |
| **"YD"** | Days, ignoring years |
| **"MD"** | Days, ignoring months and years |

### Example 1: Age Calculation
```
     A              B              C
  ┌────────────┬────────────┬──────────────────────┐
1 │ Birthdate  │ Today      │ Age                  │
2 │ 1/15/1990  │ =TODAY()   │ =DATEDIF(A2,B2,"Y")  │
  └────────────┴────────────┴──────────────────────┘
                               ↓
                              35 years
```

### Example 2: Employment Duration
```
     A              B              C
  ┌────────────┬────────────┬──────────────────────────────────┐
1 │ Hire Date  │ Today      │ Years/Months                     │
2 │ 3/15/2020  │ =TODAY()   │ =DATEDIF(A2,B2,"Y")&" years, "&DATEDIF(A2,B2,"YM")&" months"
  └────────────┴────────────┴──────────────────────────────────┘
                               ↓
                          "5 years, 9 months"
```

### Example 3: Days Until Event
```
     A              B              C
  ┌────────────┬────────────┬──────────────────────┐
1 │ Today      │ Event      │ Days Until           │
2 │ =TODAY()   │ 12/31/2025 │ =DATEDIF(A2,B2,"D")  │
  └────────────┴────────────┴──────────────────────┘
                               ↓
                              7 days
```

### Complete Age Display
```
     A              B              C
  ┌────────────┬────────────┬──────────────────────────────────┐
1 │ Birthdate  │ Today      │ Age                              │
2 │ 1/15/1990  │ =TODAY()   │ =DATEDIF(A2,B2,"Y")&" yrs, "&DATEDIF(A2,B2,"YM")&" mos, "&DATEDIF(A2,B2,"MD")&" days"
  └────────────┴────────────┴──────────────────────────────────┘
                               ↓
                         "35 yrs, 11 mos, 9 days"
```

---

## Date Arithmetic

### Basic Operations

**Adding Days:**
```
=A2 + 7   → Date 7 days later
=TODAY() + 30  → 30 days from today
```

**Subtracting Days:**
```
=A2 - 7   → Date 7 days earlier
=B2 - A2  → Days between dates
```

**Adding Months (use EDATE):**
```
=EDATE(A2, 3)  → 3 months later
```

**Adding Years:**
```
=DATE(YEAR(A2)+1, MONTH(A2), DAY(A2))  → 1 year later
```

### Example: Age Ranges
```
     A              B              C
  ┌────────────┬────────────┬──────────────────────┐
1 │ Birthdate  │ Today      │ Age Range            │
2 │ 1/15/2000  │ =TODAY()   │ =IF(DATEDIF(A2,B2,"Y")<18,"Child",IF(DATEDIF(A2,B2,"Y")<65,"Adult","Senior"))
  └────────────┴────────────┴──────────────────────┘
                               ↓
                           "Adult"

< 18 = Child
18-64 = Adult
65+ = Senior
```

### Example: Quarter Calculation
```
     A              B
  ┌────────────┬────────────────────────────────┐
1 │ Date       │ Quarter                        │
2 │ 3/15/2024  │ ="Q"&ROUNDUP(MONTH(A2)/3,0)    │
  └────────────┴────────────────────────────────┘
                  ↓
                "Q1"

Month 3 / 3 = 1, roundup = 1, result: Q1
```

---

## Common Date Patterns

### Pattern 1: First Day of Month
```
=DATE(YEAR(A2), MONTH(A2), 1)
```

### Pattern 2: Last Day of Month
```
=EOMONTH(A2, 0)
```

### Pattern 3: First Day of Year
```
=DATE(YEAR(A2), 1, 1)
```

### Pattern 4: Last Day of Year
```
=DATE(YEAR(A2), 12, 31)
```

### Pattern 5: First Day of Quarter
```
=DATE(YEAR(A2), (ROUNDUP(MONTH(A2)/3,0)-1)*3+1, 1)
```

### Pattern 6: Number of Days in Month
```
=DAY(EOMONTH(A2,0))
```

### Pattern 7: Is Leap Year?
```
=IF(DAY(DATE(YEAR(A2),2,29))=29,"Yes","No")

If Feb 29 exists, it's a leap year
```

---

## Date Formatting

### TEXT Function with Dates

```
Date: 12/24/2025

=TEXT(A2,"mm/dd/yyyy")     → "12/24/2025"
=TEXT(A2,"mmm d, yyyy")    → "Dec 24, 2025"
=TEXT(A2,"mmmm d, yyyy")   → "December 24, 2025"
=TEXT(A2,"dddd")           → "Tuesday"
=TEXT(A2,"ddd, mmm d")     → "Tue, Dec 24"
=TEXT(A2,"m/d/yy")         → "12/24/25"
```

### Format Codes

| Code | Description | Example |
|------|-------------|---------|
| **d** | Day (1-31) | 24 |
| **dd** | Day with leading zero | 24 |
| **ddd** | Day name abbreviated | Tue |
| **dddd** | Day name full | Tuesday |
| **m** | Month (1-12) | 12 |
| **mm** | Month with leading zero | 12 |
| **mmm** | Month name abbreviated | Dec |
| **mmmm** | Month name full | December |
| **yy** | Year 2-digit | 25 |
| **yyyy** | Year 4-digit | 2025 |

### Time Format Codes

| Code | Description | Example |
|------|-------------|---------|
| **h** | Hour (0-23) | 14 |
| **hh** | Hour with leading zero | 14 |
| **m** | Minute (0-59) | 30 |
| **mm** | Minute with leading zero | 30 |
| **s** | Second (0-59) | 45 |
| **ss** | Second with leading zero | 45 |
| **AM/PM** | AM or PM | PM |

**⚠️ Note:** `m` means month UNLESS immediately after `h` or `hh`, then it means minute!

### Time Formatting Examples
```
Time: 2:30:45 PM

=TEXT(A2,"h:mm AM/PM")     → "2:30 PM"
=TEXT(A2,"hh:mm:ss")       → "14:30:45"
=TEXT(A2,"h:mm")           → "14:30"
```

---

## Working with Time

### Time Calculations

**Time is a fraction of a day:**
```
1 hour = 1/24 of a day = 0.041666...
1 minute = 1/1440 of a day = 0.000694...
1 second = 1/86400 of a day = 0.000011...
```

### Adding Time
```
     A          B          C
  ┌─────────┬─────────┬──────────────┐
1 │ Start   │ Duration│ End          │
2 │ 9:00 AM │ 2:30    │ =A2+B2       │
  └─────────┴─────────┴──────────────┘
                          ↓
                      11:30 AM
```

### Calculating Duration
```
     A          B          C
  ┌─────────┬─────────┬──────────────┐
1 │ Start   │ End     │ Duration     │
2 │ 9:00 AM │ 5:00 PM │ =B2-A2       │
  └─────────┴─────────┴──────────────┘
                          ↓
                        8:00
                     (8 hours)
```

### Convert Hours to Decimal
```
     A          B
  ┌─────────┬──────────────┐
1 │ Time    │ Hours        │
2 │ 2:30    │ =A2*24       │
  └─────────┴──────────────┘
              ↓
             2.5

2:30 hours = 2.5 decimal hours
```

### Convert Decimal to Time
```
     A          B
  ┌─────────┬──────────────┐
1 │ Hours   │ Time         │
2 │ 2.5     │ =A2/24       │
  └─────────┴──────────────┘
              ↓
            2:30
         (format as time)
```

### Real-World: Timesheet Calculation
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬──────────────┐
1 │ Clock In│ Clock Out│ Hours  │ Pay          │
2 │ 9:00 AM │ 5:30 PM  │ =(B2-A2)*24 │ =C2*15   │
  └─────────┴─────────┴─────────┴──────────────┘
                          ↓         ↓
                         8.5      $127.50

Multiply by 24 to get hours as number
Then multiply by hourly rate
```

### Handling Overnight Shifts
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────┐
1 │ Start   │ End     │ Hours                │
2 │ 11:00 PM│ 7:00 AM │ =IF(B2<A2,1+B2-A2,B2-A2)*24
  └─────────┴─────────┴──────────────────────┘
                          ↓
                          8

If end < start, it crossed midnight, add 1 day
```

---

## Common Date and Time Mistakes

### Mistake 1: Text vs Date
```
❌ Problem:
Cell contains: "1/15/2024" (as text, left-aligned)
=A2+7  → Error or wrong result

✅ Solution:
Enter dates properly or use DATEVALUE:
=DATEVALUE(A2)+7
```

### Mistake 2: Wrong Date Format
```
In some regions: dd/mm/yyyy
In others: mm/dd/yyyy

"3/4/2024" could be:
- March 4 (US format)
- April 3 (European format)

Always verify date format in regional settings
```

### Mistake 3: Time Without Date
```
Problem: Time alone is stored as decimal < 1
Solution: If showing as weird number, format as Time
```

### Mistake 4: Date Math with Text
```
❌ Wrong:
="Report: "&TODAY()
Result: "Report: 45645" (serial number)

✅ Right:
="Report: "&TEXT(TODAY(),"mm/dd/yyyy")
Result: "Report: 12/24/2025"
```

### Mistake 5: DATEDIF Start > End
```
❌ Wrong:
=DATEDIF(B2, A2, "Y")  where B2 > A2
Result: #NUM! error

✅ Right:
Start date must be earlier than end date
=DATEDIF(A2, B2, "Y")
```

### Mistake 6: February 29 in Non-Leap Years
```
=DATE(2023, 2, 29)  → 3/1/2023 (Excel adjusts)

Excel handles invalid dates but may give unexpected results
```

---

## Best Practices

### 1. Use Functions for Dates, Not Text
```
❌ Avoid: ="1/15/2024"
✅ Better: =DATE(2024,1,15)

Functions are unambiguous and region-independent
```

### 2. Store Dates as Dates, Not Text
```
If imported data has text dates:
- Use DATEVALUE() to convert
- Or use Text to Columns feature
```

### 3. Use TODAY() Sparingly
```
TODAY() is volatile - recalculates frequently
In large workbooks, this slows performance

If date doesn't need to update, enter manually
```

### 4. Consider Time Zones
```
Excel doesn't have built-in time zone support
Track time zones in separate column if needed

     A              B          C
  ┌────────────┬─────────┬─────────┐
1 │ Time       │ Zone    │ Notes   │
2 │ 2:00 PM    │ EST     │         │
  └────────────┴─────────┴─────────┘
```

### 5. Format Dates Consistently
```
Choose one format and stick with it:
- mm/dd/yyyy (US)
- dd/mm/yyyy (Europe)
- yyyy-mm-dd (ISO standard, unambiguous)
```

### 6. Document Date Assumptions
```
Add comments explaining:
- Fiscal year start dates
- Holiday lists
- Business day definitions
- Time zone assumptions
```

---

## Real-World Application: Project Timeline

Let's build a project timeline calculator.

### Project Setup
```
     A                  B              C
  ┌────────────────┬────────────┬────────────┐
1 │ Task           │ Start Date │ Duration   │
2 │ Planning       │ 1/2/2024   │ 10         │
3 │ Development    │            │ 30         │
4 │ Testing        │            │ 15         │
5 │ Deployment     │            │ 5          │
  └────────────────┴────────────┴────────────┘
```

### Formulas

**Column B (Start Date):**
```
B2: 1/2/2024  (manual entry)
B3: =WORKDAY(D2,1)  (day after previous task ends)
B4: =WORKDAY(D3,1)
B5: =WORKDAY(D4,1)
```

**Column D (End Date):**
```
D2: =WORKDAY(B2,C2-1)
D3: =WORKDAY(B3,C3-1)
D4: =WORKDAY(B4,C4-1)
D5: =WORKDAY(B5,C5-1)
```

**Column E (Status):**
```
E2: =IF(TODAY()>D2,"Complete",IF(TODAY()>=B2,"In Progress","Not Started"))
```

### Complete Timeline
```
     A              B          C          D          E
  ┌────────────┬──────────┬─────────┬──────────┬────────────┐
1 │ Task       │ Start    │ Days    │ End      │ Status     │
2 │ Planning   │ 1/2/24   │ 10      │ 1/15/24  │ Complete   │
3 │ Development│ 1/16/24  │ 30      │ 2/26/24  │ Complete   │
4 │ Testing    │ 2/27/24  │ 15      │ 3/18/24  │ Complete   │
5 │ Deployment │ 3/19/24  │ 5       │ 3/25/24  │ Complete   │
6 │            │          │         │          │            │
7 │ Total Days:│ =NETWORKDAYS(B2,D5)          │            │
8 │ Project End:│ =D5                         │            │
  └────────────┴──────────┴─────────┴──────────┴────────────┘
```

### Add Holidays
```
Create a holiday list:
     G          H
  ┌──────────┬────────────┐
1 │ Holidays │            │
2 │ 1/15/24  │ MLK Day    │
3 │ 2/19/24  │ Presidents │
  └──────────┴────────────┘

Update WORKDAY formulas:
D2: =WORKDAY(B2,C2-1,$G$2:$G$10)
```

---

## Advanced Date Techniques

### Technique 1: Age on Specific Date
```
     A              B              C
  ┌────────────┬────────────┬──────────────────────┐
1 │ Birthdate  │ As Of Date │ Age Then             │
2 │ 1/15/1990  │ 12/31/2024 │ =DATEDIF(A2,B2,"Y")  │
  └────────────┴────────────┴──────────────────────┘
                               ↓
                              34

Age as of specific date (not today)
```

### Technique 2: Working Days Between Dates (Excluding Weekends Only)
```
=NETWORKDAYS.INTL(start, end, weekend, holidays)

Weekend codes:
1 = Sat-Sun (default)
2 = Sun-Mon
3 = Mon-Tue
...
11 = Sun only
```

### Technique 3: Date Validation
```
     A              B
  ┌────────────┬────────────────────────────────┐
1 │ Date       │ Valid?                         │
2 │ 1/15/2024  │ =IF(ISNUMBER(A2),"Valid Date","Invalid")
  └────────────┴────────────────────────────────┘

Dates are numbers, text is not
```

### Technique 4: Conditional Formatting by Date
```
Highlight dates in the past:
Rule: =A2<TODAY()

Highlight dates within 7 days:
Rule: =AND(A2>=TODAY(),A2<=TODAY()+7)
```

### Technique 5: Dynamic Date Ranges
```
Current Month:
Start: =DATE(YEAR(TODAY()),MONTH(TODAY()),1)
End:   =EOMONTH(TODAY(),0)

Last Month:
Start: =DATE(YEAR(TODAY()),MONTH(TODAY())-1,1)
End:   =EOMONTH(TODAY(),-1)
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- Excel stores dates as serial numbers (days since 1/1/1900)
- Times are decimal fractions of a day
- TODAY() returns current date
- NOW() returns current date and time
- DATE(year, month, day) creates a date
- YEAR, MONTH, DAY extract date parts
- Date math: Later Date - Earlier Date = Days
- DATEDIF(start, end, unit) calculates differences
- WORKDAY counts business days
- Format dates with TEXT function

### Practice Deeply
- Using TODAY() and NOW() in formulas
- Creating dates with DATE function
- Extracting date parts (YEAR, MONTH, DAY)
- Calculating age with DATEDIF
- Using EOMONTH for end-of-month dates
- Using EDATE to add months
- Working with WORKDAY and NETWORKDAYS
- Date arithmetic (adding/subtracting days)
- Using WEEKDAY to identify day of week
- Formatting dates with TEXT
- Calculating time durations
- Building project timelines
- Handling overnight time calculations
- Troubleshooting date display issues

### Don't Memorize
- Every date format code (look up when needed)
- Exact serial numbers for dates
- All WEEKDAY return types (use most common)
- All DATEDIF units (reference as needed)
- Complex nested date formulas (build step by step)

---

## Quick Reference: Date Functions

### Getting Current Date/Time
```
=TODAY()              Current date
=NOW()                Current date and time
```

### Creating Dates
```
=DATE(2024,1,15)      Create date from parts
=TIME(14,30,0)        Create time from parts
```

### Extracting Date Parts
```
=YEAR(A2)             Extract year
=MONTH(A2)            Extract month (1-12)
=DAY(A2)              Extract day (1-31)
=WEEKDAY(A2)          Day of week (1-7)
=HOUR(A2)             Extract hour
=MINUTE(A2)           Extract minute
=SECOND(A2)           Extract second
```

### Date Calculations
```
=A2+7                 Add 7 days
=B2-A2                Days between dates
=EDATE(A2,3)          Add 3 months
=EOMONTH(A2,0)        End of month
=DATEDIF(A2,B2,"Y")   Years between dates
```

### Business Days
```
=WORKDAY(A2,10)       10 workdays from date
=NETWORKDAYS(A2,B2)   Workdays between dates
```

### Formatting
```
=TEXT(A2,"mm/dd/yyyy")      Format as date
=TEXT(A2,"dddd")            Day name
=TEXT(A2,"mmmm")            Month name
```

---

## Troubleshooting Date Issues

### Issue: Date Shows as Number
**Cause:** Cell formatted as General or Number

**Solution:**
1. Right-click cell → Format Cells
2. Select "Date" category
3. Choose desired format

### Issue: Date Shows as #####
**Cause:** Column too narrow

**Solution:**
- Double-click column border to auto-fit
- Or drag column wider

### Issue: Wrong Century (1/15/24 shows as 1924)
**Cause:** Excel's 2-digit year cutoff

**Solution:**
- Use 4-digit years: 2024 instead of 24
- Or check Excel's date cutoff settings

### Issue: Date Arithmetic Returns Decimal
**Cause:** Dividing by days returns fractional result

**Solution:**
```
Convert to integer:
=INT(date_formula)
```

### Issue: Dates From Import Show as Text
**Cause:** Dates imported as text strings

**Solution:**
```
Method 1: =DATEVALUE(A2)
Method 2: Data → Text to Columns → Date format
Method 3: =VALUE(A2) then format as date
```

---

## Regional Date Considerations

### Date Order Varies by Region

**US Format:** mm/dd/yyyy (12/24/2025)
**European Format:** dd/mm/yyyy (24/12/2025)
**ISO Format:** yyyy-mm-dd (2025-12-24) ← Recommended for clarity

### Recommendations
1. Use DATE function: `=DATE(2025,12,24)` (unambiguous)
2. Use ISO format for international work
3. Document your format choice
4. Test with dates like 3/4 vs 4/3 to verify

---

## Next Step

After mastering date and time functions, you're ready to explore:

**`09-mathematical-and-statistical-functions.md`**
- Advanced math functions (ROUND, CEILING, FLOOR)
- Statistical analysis (MEDIAN, MODE, STDEV)
- SUMIFS, COUNTIFS, AVERAGEIFS (multiple criteria)
- RANK and PERCENTILE functions
- Random number generation
- Array formulas for complex calculations
- Statistical analysis techniques
