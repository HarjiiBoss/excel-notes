# Text Functions

This file covers Excel's text manipulation functions that let you combine, extract,
clean, and transform text data. These functions are essential for working with
names, addresses, codes, and any text-based information.

---

## What are Text Functions?

**Text functions** manipulate and transform text (strings) in cells.

Think of them as **text editing tools** that:
- Combine multiple pieces of text
- Extract specific parts of text
- Change text case (upper/lower)
- Remove unwanted characters
- Find text within text
- Convert between text and numbers

### Why Text Functions Matter

Real-world data is messy:
- Names in wrong format: "SMITH, JOHN" → "John Smith"
- Extra spaces: "  Widget  " → "Widget"
- Need to split: "john.doe@company.com" → "john.doe" and "company.com"
- Need to combine: "John" + "Smith" → "John Smith"

---

## The Ampersand (&) - Text Concatenation

**Purpose:** Joins text together (simplest way to combine text)

**Syntax:** `=text1 & text2 & text3...`

### Basic Examples
```
     A          B          C
  ┌─────────┬─────────┬──────────────┐
1 │ First   │ Last    │ Full Name    │
2 │ John    │ Smith   │ =A2&" "&B2   │
  └─────────┴─────────┴──────────────┘
                          ↓
                     "John Smith"

Breakdown:
A2      →  "John"
&       →  join
" "     →  space
&       →  join
B2      →  "Smith"
Result: "John Smith"
```

### Example 1: Creating Full Names
```
     A          B          C
  ┌─────────┬─────────┬──────────────┐
1 │ First   │ Last    │ Full Name    │
2 │ Alice   │ Johnson │ =A2&" "&B2      → "Alice Johnson"
3 │ Bob     │ Lee     │ =A3&" "&B3      → "Bob Lee"
4 │ Carol   │ Davis   │ =A4&" "&B4      → "Carol Davis"
  └─────────┴─────────┴──────────────┘
```

### Example 2: Creating Email Addresses
```
     A          B                    C
  ┌─────────┬────────────────┬──────────────────────┐
1 │ Username│ Domain         │ Email                │
2 │ jsmith  │ company.com    │ =A2&"@"&B2           │
  └─────────┴────────────────┴──────────────────────┘
                                ↓
                         "jsmith@company.com"
```

### Example 3: Building Product Codes
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬──────────────┐
1 │ Category│ Type    │ Number  │ SKU          │
2 │ WDG     │ PRO     │ 001     │ =A2&"-"&B2&"-"&C2  → "WDG-PRO-001"
3 │ GAD     │ STD     │ 002     │ =A3&"-"&B3&"-"&C3  → "GAD-STD-002"
  └─────────┴─────────┴─────────┴──────────────┘
```

### Example 4: Combining with Text and Numbers
```
     A          B
  ┌─────────┬────────────────────────────┐
1 │ Sales   │ Report                     │
2 │ 15000   │ ="Total Sales: $"&A2       │
  └─────────┴────────────────────────────┘
              ↓
         "Total Sales: $15000"

Note: Number is automatically converted to text
```

### ⚠️ Important Notes
- Text must be in quotes: `"Hello"`
- Spaces must be added explicitly: `" "`
- Numbers are automatically converted to text
- Result is always text (not a number)

---

## CONCAT and CONCATENATE Functions

**Purpose:** Join multiple text strings into one

**CONCATENATE Syntax:** `=CONCATENATE(text1, [text2], ...)`
**CONCAT Syntax:** `=CONCAT(text1, [text2], ...)`

### Difference Between Them

| Function | Excel Version | Can Join Ranges? |
|----------|---------------|------------------|
| **CONCATENATE** | All versions | No (individual cells only) |
| **CONCAT** | Excel 2019+, 365 | Yes (can reference ranges) |

### Basic Examples
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────────┐
1 │ First   │ Last    │ Full Name                │
2 │ John    │ Smith   │ =CONCATENATE(A2," ",B2)  │
  └─────────┴─────────┴──────────────────────────┘
                          ↓
                     "John Smith"

Same as: =A2&" "&B2
```

### CONCAT with Ranges (Excel 2019+)
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬──────────────────┐
1 │ W       │ i       │ d       │ Result           │
2 │ g       │ e       │ t       │ =CONCAT(A1:C2)   │
  └─────────┴─────────┴─────────┴──────────────────┘
                                    ↓
                                "Widget"

Joins all cells in range A1:C2
```

### When to Use Each

**Use & (Ampersand):**
- Simple combinations
- Just a few pieces of text
- Most readable for beginners

**Use CONCATENATE:**
- Working with older Excel versions
- Need function-based approach

**Use CONCAT:**
- Excel 2019+ or 365
- Need to join ranges
- More flexible

---

## TEXTJOIN Function

**Purpose:** Join text with a delimiter, with option to ignore empties

**Syntax:** `=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

**Available in:** Excel 2019, Microsoft 365, Excel Online

### Arguments

| Argument | What It Is | Example |
|----------|------------|---------|
| **delimiter** | Text to put between each value | `", "` or `"-"` or `" "` |
| **ignore_empty** | TRUE = skip blanks, FALSE = include | `TRUE` |
| **text1, text2...** | Values or ranges to join | `A2:A10` |

### Basic Example
```
     A          B
  ┌─────────┬────────────────────────────┐
1 │ Items   │ List                       │
2 │ Apple   │ =TEXTJOIN(", ",TRUE,A2:A5) │
3 │ Banana  │                            │
4 │         │  ← Empty cell              │
5 │ Cherry  │                            │
  └─────────┴────────────────────────────┘
              ↓
         "Apple, Banana, Cherry"

Empty cell in A4 is ignored because ignore_empty=TRUE
```

### Example 1: Creating Address from Parts
```
     A          B          C          D          E
  ┌─────────┬─────────┬─────────┬─────────┬─────────────────────┐
1 │ Street  │ City    │ State   │ ZIP     │ Full Address        │
2 │ 123 Main│ Boston  │ MA      │ 02101   │ =TEXTJOIN(", ",TRUE,A2:D2)
  └─────────┴─────────┴─────────┴─────────┴─────────────────────┘
                                              ↓
                                "123 Main, Boston, MA, 02101"
```

### Example 2: Combining Names with Custom Separator
```
     A          B          C          D
  ┌─────────┬─────────┬─────────┬─────────────────────┐
1 │ Alice   │ Bob     │ Carol   │ Team                │
2 │         │         │         │ =TEXTJOIN(" | ",TRUE,A1:C1)
  └─────────┴─────────┴─────────┴─────────────────────┘
                                    ↓
                              "Alice | Bob | Carol"
```

### Example 3: Ignoring Blanks
```
     A          B
  ┌─────────┬────────────────────────────┐
1 │ Name    │ List                       │
2 │ John    │ =TEXTJOIN(",",TRUE,A2:A6)  │
3 │         │  ← Blank                   │
4 │ Sarah   │                            │
5 │         │  ← Blank                   │
6 │ Mike    │                            │
  └─────────┴────────────────────────────┘
              ↓
         "John,Sarah,Mike"

With ignore_empty=FALSE, result would be: "John,,Sarah,,Mike"
```

### Real-World: Creating Tag Lists
```
     A          B          C          D          E
  ┌─────────┬─────────┬─────────┬─────────┬─────────────────────┐
1 │ Tag1    │ Tag2    │ Tag3    │ Tag4    │ Tags                │
2 │ urgent  │         │ review  │ client  │ =TEXTJOIN(", ",TRUE,A2:D2)
  └─────────┴─────────┴─────────┴─────────┴─────────────────────┘
                                              ↓
                                    "urgent, review, client"
```

---

## LEFT, RIGHT, and MID Functions

**Purpose:** Extract specific characters from text

### LEFT Function

**Syntax:** `=LEFT(text, [num_chars])`

**Purpose:** Extract characters from the **beginning** of text

```
     A                  B
  ┌────────────────┬──────────────┐
1 │ Text           │ First 3      │
2 │ Widget         │ =LEFT(A2,3)  │
  └────────────────┴──────────────┘
                      ↓
                    "Wid"

Takes 3 characters from left: W-i-d
```

**Examples:**
```
=LEFT("Hello", 2)      → "He"
=LEFT("Excel", 4)      → "Exce"
=LEFT("12345", 3)      → "123"
=LEFT(A2, 1)           → First character
```

### RIGHT Function

**Syntax:** `=RIGHT(text, [num_chars])`

**Purpose:** Extract characters from the **end** of text

```
     A                  B
  ┌────────────────┬──────────────┐
1 │ Text           │ Last 3       │
2 │ Widget         │ =RIGHT(A2,3) │
  └────────────────┴──────────────┘
                      ↓
                    "get"

Takes 3 characters from right: g-e-t
```

**Examples:**
```
=RIGHT("Hello", 2)     → "lo"
=RIGHT("Excel", 3)     → "cel"
=RIGHT("12345", 2)     → "45"
=RIGHT(A2, 4)          → Last 4 characters
```

### MID Function

**Syntax:** `=MID(text, start_num, num_chars)`

**Purpose:** Extract characters from the **middle** of text

```
     A                  B
  ┌────────────────┬──────────────┐
1 │ Text           │ Middle       │
2 │ Widget         │ =MID(A2,3,2) │
  └────────────────┴──────────────┘
                      ↓
                    "dg"

Start at position 3, take 2 characters: d-g
```

**Visual Position Map:**
```
Text: "Widget"
Pos:   123456

=MID("Widget", 1, 3)  → "Wid"  (positions 1-3)
=MID("Widget", 3, 2)  → "dg"   (positions 3-4)
=MID("Widget", 4, 3)  → "get"  (positions 4-6)
```

**Examples:**
```
=MID("Hello World", 7, 5)   → "World"
=MID("ABC-123-XYZ", 5, 3)   → "123"
=MID(A2, 2, 4)              → 4 chars starting at position 2
```

### Real-World Example 1: Extracting from Product Codes
```
     A              B          C          D
  ┌────────────┬──────────┬──────────┬──────────┐
1 │ SKU        │ Category │ Type     │ Number   │
2 │ WDG-PRO-001│ =LEFT(A2,3) │ =MID(A2,5,3) │ =RIGHT(A2,3)
  └────────────┴──────────┴──────────┴──────────┘
                  ↓          ↓          ↓
               "WDG"      "PRO"      "001"
```

### Real-World Example 2: Phone Number Parts
```
     A              B          C          D
  ┌────────────┬──────────┬──────────┬──────────┐
1 │ Phone      │ Area     │ Exchange │ Number   │
2 │ 5551234567 │ =LEFT(A2,3) │ =MID(A2,4,3) │ =RIGHT(A2,4)
  └────────────┴──────────┴──────────┴──────────┘
                  ↓          ↓          ↓
               "555"      "123"      "4567"
```

### Real-World Example 3: Email Username
```
     A                      B
  ┌────────────────────┬──────────────────────┐
1 │ Email              │ Username             │
2 │ john@company.com   │ =LEFT(A2,FIND("@",A2)-1)
  └────────────────────┴──────────────────────┘
                          ↓
                       "john"

Finds @ position, then extracts everything before it
```

---

## LEN Function

**Purpose:** Returns the length (number of characters) of text

**Syntax:** `=LEN(text)`

### Basic Examples
```
     A              B
  ┌────────────┬──────────────┐
1 │ Text       │ Length       │
2 │ Widget     │ =LEN(A2)     │
  └────────────┴──────────────┘
                  ↓
                  6

"Widget" has 6 characters: W-i-d-g-e-t
```

**More Examples:**
```
=LEN("Hello")          → 5
=LEN("Hello World")    → 11 (space counts!)
=LEN("123")            → 3
=LEN("")               → 0 (empty string)
=LEN("  Hello  ")      → 9 (spaces count)
```

### Use Case 1: Validation
```
     A              B                C
  ┌────────────┬──────────┬────────────────────┐
1 │ Password   │ Length   │ Valid?             │
2 │ abc123     │ =LEN(A2) │ =IF(B2>=8,"Yes","Too Short")
  └────────────┴──────────┴────────────────────┘
                  ↓          ↓
                  6       "Too Short"
```

### Use Case 2: Character Counter
```
     A              B
  ┌────────────┬────────────────────────────┐
1 │ Tweet      │ Characters Remaining       │
2 │ Hello!     │ =280-LEN(A2)               │
  └────────────┴────────────────────────────┘
                  ↓
                274

Twitter limit: 280 characters
```

### Use Case 3: Extract Last Word
```
Extract last N characters:
=RIGHT(A2, LEN(A2)-FIND(" ", A2))

For "Hello World":
- LEN gives 11
- FIND finds space at position 6
- 11-6 = 5
- RIGHT takes 5 characters: "World"
```

---

## FIND and SEARCH Functions

**Purpose:** Find the position of text within text

### FIND Function

**Syntax:** `=FIND(find_text, within_text, [start_num])`

**Features:**
- **Case-sensitive**
- Returns position number
- Returns #VALUE! if not found

```
     A                  B
  ┌────────────────┬──────────────────┐
1 │ Text           │ Position         │
2 │ Hello World    │ =FIND("o",A2)    │
  └────────────────┴──────────────────┘
                      ↓
                      5

"o" first appears at position 5 (Hell-o)
```

### SEARCH Function

**Syntax:** `=SEARCH(find_text, within_text, [start_num])`

**Features:**
- **NOT case-sensitive**
- Supports wildcards (* and ?)
- Returns position number
- Returns #VALUE! if not found

```
     A                  B
  ┌────────────────┬──────────────────┐
1 │ Text           │ Position         │
2 │ Hello World    │ =SEARCH("WORLD",A2)
  └────────────────┴──────────────────┘
                      ↓
                      7

Finds "World" even though we searched for "WORLD"
```

### FIND vs SEARCH

| Feature | FIND | SEARCH |
|---------|------|--------|
| **Case-sensitive** | Yes | No |
| **Wildcards** | No | Yes |
| **Speed** | Faster | Slightly slower |
| **Use when** | Exact match needed | Flexible matching |

### Examples with Position
```
Text: "Hello World"
Positions: 123456789...

=FIND("H", A2)      → 1
=FIND("o", A2)      → 5 (first "o")
=FIND("W", A2)      → 7
=FIND("world", A2)  → #VALUE! (case-sensitive, not found)
=SEARCH("world", A2) → 7 (not case-sensitive, found)
```

### Example 1: Extract Email Domain
```
     A                      B
  ┌────────────────────┬──────────────────────┐
1 │ Email              │ Domain               │
2 │ john@company.com   │ =MID(A2,FIND("@",A2)+1,LEN(A2))
  └────────────────────┴──────────────────────┘
                          ↓
                    "company.com"

Steps:
1. FIND("@"...) finds @ at position 5
2. +1 = start at position 6
3. Extract rest of string
```

### Example 2: Extract First Name
```
     A                  B
  ┌────────────────┬──────────────────────────┐
1 │ Full Name      │ First Name               │
2 │ John Smith     │ =LEFT(A2,FIND(" ",A2)-1) │
  └────────────────┴──────────────────────────┘
                      ↓
                   "John"

Steps:
1. FIND finds space at position 5
2. -1 = take 4 characters
3. LEFT extracts "John"
```

### Example 3: Check if Text Contains Word
```
     A                  B
  ┌────────────────┬────────────────────────────────┐
1 │ Description    │ Contains "urgent"?             │
2 │ urgent review  │ =IF(ISNUMBER(SEARCH("urgent",A2)),"Yes","No")
  └────────────────┴────────────────────────────────┘
                      ↓
                    "Yes"

SEARCH returns position (number) if found, error if not
ISNUMBER checks if result is a number
```

### Using Wildcards with SEARCH
```
* = any number of characters
? = exactly one character

=SEARCH("*son", A2)    → Finds "Johnson", "Anderson", etc.
=SEARCH("a?e", A2)     → Finds "age", "are", "ate", etc.
=SEARCH("*@*.*", A2)   → Finds any email format
```

---

## UPPER, LOWER, and PROPER Functions

**Purpose:** Change the case of text

### UPPER Function

**Syntax:** `=UPPER(text)`

**Purpose:** Converts all text to UPPERCASE

```
     A              B
  ┌────────────┬──────────────┐
1 │ Text       │ Uppercase    │
2 │ hello      │ =UPPER(A2)   │
3 │ Hello      │ =UPPER(A3)   │
4 │ HELLO      │ =UPPER(A4)   │
  └────────────┴──────────────┘
                  ↓   ↓   ↓
               "HELLO" "HELLO" "HELLO"
```

### LOWER Function

**Syntax:** `=LOWER(text)`

**Purpose:** Converts all text to lowercase

```
     A              B
  ┌────────────┬──────────────┐
1 │ Text       │ Lowercase    │
2 │ HELLO      │ =LOWER(A2)   │
3 │ Hello      │ =LOWER(A3)   │
4 │ hello      │ =LOWER(A4)   │
  └────────────┴──────────────┘
                  ↓   ↓   ↓
               "hello" "hello" "hello"
```

### PROPER Function

**Syntax:** `=PROPER(text)`

**Purpose:** Converts text to Title Case (first letter of each word capitalized)

```
     A                      B
  ┌────────────────────┬──────────────┐
1 │ Text               │ Title Case   │
2 │ hello world        │ =PROPER(A2)  │
3 │ HELLO WORLD        │ =PROPER(A3)  │
4 │ hElLo WoRlD        │ =PROPER(A4)  │
  └────────────────────┴──────────────┘
                          ↓   ↓   ↓
                   All result in "Hello World"
```

### Real-World Example 1: Standardizing Names
```
     A                  B
  ┌────────────────┬──────────────┐
1 │ Name (raw)     │ Name (clean) │
2 │ JOHN SMITH     │ =PROPER(A2)  │
3 │ alice johnson  │ =PROPER(A3)  │
4 │ bOb LeE        │ =PROPER(A4)  │
  └────────────────┴──────────────┘
                      ↓
                "John Smith"
                "Alice Johnson"
                "Bob Lee"
```

### Real-World Example 2: Email Addresses
```
     A                      B
  ┌────────────────────┬──────────────┐
1 │ Email (input)      │ Email (clean)│
2 │ John@COMPANY.COM   │ =LOWER(A2)   │
  └────────────────────┴──────────────┘
                          ↓
                   "john@company.com"

Email addresses are case-insensitive, but lowercase is standard
```

### Real-World Example 3: Product Codes
```
     A              B
  ┌────────────┬──────────────┐
1 │ Code       │ Standardized │
2 │ wdg-001    │ =UPPER(A2)   │
3 │ Gad-002    │ =UPPER(A3)   │
  └────────────┴──────────────┘
                  ↓      ↓
              "WDG-001" "GAD-002"
```

---

## TRIM, CLEAN, and SUBSTITUTE Functions

### TRIM Function

**Syntax:** `=TRIM(text)`

**Purpose:** Removes **extra spaces** from text
- Removes leading spaces (before text)
- Removes trailing spaces (after text)
- Reduces multiple spaces to single space

```
     A                      B
  ┌────────────────────┬──────────────┐
1 │ Text with spaces   │ Cleaned      │
2 │ "  Hello  World  " │ =TRIM(A2)    │
  └────────────────────┴──────────────┘
                          ↓
                    "Hello World"

All extra spaces removed!
```

**Examples:**
```
=TRIM("  Hello")           → "Hello"
=TRIM("Hello  ")           → "Hello"
=TRIM("Hello    World")    → "Hello World"
=TRIM("  Hello  World  ")  → "Hello World"
```

### CLEAN Function

**Syntax:** `=CLEAN(text)`

**Purpose:** Removes non-printable characters (like line breaks)

```
Useful for data imported from other systems that
contains hidden characters you can't see
```

**Common use:**
```
=CLEAN(A2)  → Removes line breaks, tabs, etc.

Often combined with TRIM:
=TRIM(CLEAN(A2))  → Removes both unprintable chars AND extra spaces
```

### SUBSTITUTE Function

**Syntax:** `=SUBSTITUTE(text, old_text, new_text, [instance_num])`

**Purpose:** Replaces specific text with new text

```
     A              B
  ┌────────────┬────────────────────────────┐
1 │ Text       │ Result                     │
2 │ Hello World│ =SUBSTITUTE(A2,"World","Excel")
  └────────────┴────────────────────────────┘
                  ↓
             "Hello Excel"
```

### SUBSTITUTE Examples

**Example 1: Replace all occurrences**
```
=SUBSTITUTE("Hello Hello", "Hello", "Hi")
→ "Hi Hi"

Replaces ALL instances
```

**Example 2: Replace specific occurrence**
```
=SUBSTITUTE("Hello Hello Hello", "Hello", "Hi", 2)
                                                  ↑
                                       Replace only 2nd occurrence
→ "Hello Hi Hello"
```

**Example 3: Remove characters**
```
=SUBSTITUTE(A2, "-", "")
             Replace with nothing (empty string)

"555-123-4567" → "5551234567"
```

### Real-World Example 1: Clean Phone Numbers
```
     A              B
  ┌────────────────┬──────────────────────────────┐
1 │ Phone (raw)    │ Phone (clean)                │
2 │ (555) 123-4567 │ =SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2,"(",""),")",""),"-","")
  └────────────────┴──────────────────────────────┘
                      ↓
                 "5551234567"

Removes (, ), and -
```

**Better approach - nested:**
```
Step 1: Remove (  → =SUBSTITUTE(A2,"(","")
Step 2: Remove )  → =SUBSTITUTE(Step1,")","")
Step 3: Remove -  → =SUBSTITUTE(Step2,"-","")
Step 4: Remove space → =SUBSTITUTE(Step3," ","")
```

### Real-World Example 2: Replace Line Breaks
```
     A                  B
  ┌────────────────┬────────────────────────────┐
1 │ Address        │ Single Line                │
2 │ 123 Main       │ =SUBSTITUTE(A2,CHAR(10)," ")
  │ Boston MA      │                            │
  └────────────────┴────────────────────────────┘
                      ↓
              "123 Main Boston MA"

CHAR(10) is line break character
```

### Real-World Example 3: Standardize Data
```
     A                  B
  ┌────────────────┬────────────────────────────┐
1 │ Status         │ Standardized               │
2 │ Yes            │ =SUBSTITUTE(SUBSTITUTE(A2,"Yes","Y"),"No","N")
3 │ No             │                            │
  └────────────────┴────────────────────────────┘
                      ↓     ↓
                     "Y"   "N"
```

---

## TEXT Function

**Syntax:** `=TEXT(value, format_text)`

**Purpose:** Converts numbers to text with specific formatting

### Common Format Codes

| Format Code | Result | Example |
|-------------|--------|---------|
| `"0"` | Number with no decimals | `=TEXT(1234.5,"0")` → "1235" |
| `"0.00"` | Two decimal places | `=TEXT(1234.5,"0.00")` → "1234.50" |
| `"#,##0"` | Thousands separator | `=TEXT(1234,"#,##0")` → "1,234" |
| `"$#,##0.00"` | Currency format | `=TEXT(1234.5,"$#,##0.00")` → "$1,234.50" |
| `"0%"` | Percentage | `=TEXT(0.85,"0%")` → "85%" |
| `"mmm d, yyyy"` | Date format | `=TEXT(TODAY(),"mmm d, yyyy")` → "Dec 24, 2025" |
| `"dddd"` | Day name | `=TEXT(TODAY(),"dddd")` → "Wednesday" |

### Example 1: Format Numbers as Text
```
     A          B
  ┌─────────┬────────────────────────┐
1 │ Number  │ Formatted              │
2 │ 1234.5  │ =TEXT(A2,"$#,##0.00")  │
  └─────────┴────────────────────────┘
              ↓
          "$1,234.50"
```

### Example 2: Date to Text
```
     A          B
  ┌─────────┬────────────────────────┐
1 │ Date    │ Text                   │
2 │ 1/15/24 │ =TEXT(A2,"mmmm d, yyyy")
  └─────────┴────────────────────────┘
              ↓
          "January 15, 2024"
```

### Example 3: Combine Number with Text
```
     A          B
  ┌─────────┬──────────────────────────────────┐
1 │ Sales   │ Report                           │
2 │ 15000   │ ="Sales: "&TEXT(A2,"$#,##0")     │
  └─────────┴──────────────────────────────────┘
              ↓
          "Sales: $15,000"

Without TEXT: "Sales: 15000" (no formatting)
```

### Example 4: Leading Zeros
```
     A          B
  ┌─────────┬────────────────────────┐
1 │ ID      │ Formatted ID           │
2 │ 5       │ =TEXT(A2,"00000")      │
3 │ 123     │ =TEXT(A3,"00000")      │
  └─────────┴────────────────────────┘
              ↓       ↓
          "00005"  "00123"

Pads with leading zeros
```

### Real-World: Invoice Number
```
     A          B          C
  ┌─────────┬─────────┬────────────────────────┐
1 │ Year    │ Number  │ Invoice #              │
2 │ 2024    │ 5       │ ="INV-"&A2&"-"&TEXT(B2,"0000")
  └─────────┴─────────┴────────────────────────┘
                          ↓
                    "INV-2024-0005"
```

---

## VALUE Function

**Syntax:** `=VALUE(text)`

**Purpose:** Converts text that looks like a number into an actual number

```
     A          B
  ┌─────────┬──────────────┐
1 │ Text    │ Number       │
2 │ "123"   │ =VALUE(A2)   │
  └─────────┴──────────────┘
              ↓
            123 (as a number, not text)

Now can be used in calculations
```

### When to Use VALUE

**Problem: Numbers stored as text**
```
     A          B          C
  ┌─────────┬─────────┬──────────┐
1 │ "100"   │ "200"   │ =A1+B1   │  → "100200" (concatenates!)
  └─────────┴─────────┴──────────┘

Text "100" + Text "200" = "100200" (joined, not added)
```

**Solution: Convert to numbers first**
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────┐
1 │ "100"   │ "200"   │ =VALUE(A1)+VALUE(B1) │  → 300
  └─────────┴─────────┴──────────────────────┘

Number 100 + Number 200 = 300 (added correctly)
```

### Real-World Example: Cleaning Imported Data
```
Imported data often has numbers as text:
     A          B
  ┌─────────┬──────────────┐
1 │ Sales   │ Numeric      │
2 │ "1000"  │ =VALUE(A2)   │
3 │ "2000"  │ =VALUE(A3)   │
4 │         │              │
5 │ Total:  │ =SUM(B2:B3)  │  → 3000
  └─────────┴──────────────┘
```

---

## EXACT Function

**Syntax:** `=EXACT(text1, text2)`

**Purpose:** Compares two text strings and returns TRUE if they are exactly the same (case-sensitive)

```
     A          B          C
  ┌─────────┬─────────┬─────────────────┐
1 │ Text1   │ Text2   │ Exact Match?    │
2 │ Hello   │ Hello   │ =EXACT(A2,B2)   │  → TRUE
3 │ Hello   │ hello   │ =EXACT(A3,B3)   │  → FALSE
4 │ Hello   │ Hello   │ =EXACT(A4,B4)   │  → FALSE (trailing space)
  └─────────┴─────────┴─────────────────┘
```

### Regular Comparison vs EXACT

```
Regular = operator (NOT case-sensitive):
="Hello"="hello"  → TRUE

EXACT function (case-sensitive):
=EXACT("Hello","hello")  → FALSE
```

### Use Case: Password Validation
```
     A              B              C
  ┌────────────┬────────────┬─────────────────────┐
1 │ Password   │ Confirm    │ Match?              │
2 │ SecureP@ss │ SecureP@ss │ =EXACT(A2,B2)       │
  └────────────┴────────────┴─────────────────────┘
                                ↓
                              TRUE

Ensures exact match including case
```

---

## REPT Function

**Syntax:** `=REPT(text, number_times)`

**Purpose:** Repeats text a specified number of times

```
     A          B
  ┌─────────┬──────────────────┐
1 │ Char    │ Result           │
2 │ *       │ =REPT(A2,5)      │
  └─────────┴──────────────────┘
              ↓
           "*****"
```

### Use Case 1: Progress Bars
```
     A          B          C
  ┌─────────┬─────────┬────────────────────────┐
1 │ Progress│ Total   │ Bar                    │
2 │ 7       │ 10      │ =REPT("█",A2)&REPT("░",B2-A2)
  └─────────┴─────────┴────────────────────────┘
                          ↓
                    "███████░░░"

7 filled blocks, 3 empty blocks
```

### Use Case 2: Star Ratings
```
     A          B
  ┌─────────┬────────────────────────┐
1 │ Rating  │ Stars                  │
2 │ 4       │ =REPT("★",A2)&REPT("☆",5-A2)
  └─────────┴────────────────────────┘
              ↓
           "★★★★☆"
```

### Use Case 3: Indentation
```
     A          B          C
  ┌─────────┬─────────┬──────────────────────┐
1 │ Level   │ Item    │ Indented             │
2 │ 0       │ Parent  │ =REPT("  ",A2)&B2    │
3 │ 1       │ Child   │ =REPT("  ",A3)&B3    │
4 │ 2       │ SubChild│ =REPT("  ",A4)&B4    │
  └─────────┴─────────┴──────────────────────┘
                          ↓
                      "Parent"
                      "  Child"
                      "    SubChild"
```

---

## TEXTBEFORE and TEXTAFTER Functions

**Available in:** Excel 365 (newer versions)

### TEXTBEFORE Function

**Syntax:** `=TEXTBEFORE(text, delimiter, [instance_num])`

**Purpose:** Extracts text before a delimiter

```
     A                      B
  ┌────────────────────┬──────────────────────┐
1 │ Email              │ Username             │
2 │ john@company.com   │ =TEXTBEFORE(A2,"@")  │
  └────────────────────┴──────────────────────┘
                          ↓
                       "john"
```

### TEXTAFTER Function

**Syntax:** `=TEXTAFTER(text, delimiter, [instance_num])`

**Purpose:** Extracts text after a delimiter

```
     A                      B
  ┌────────────────────┬──────────────────────┐
1 │ Email              │ Domain               │
2 │ john@company.com   │ =TEXTAFTER(A2,"@")   │
  └────────────────────┴──────────────────────┘
                          ↓
                    "company.com"
```

### Real-World Example: Splitting Full Names
```
     A                  B                      C
  ┌────────────────┬──────────────────────┬──────────────────────┐
1 │ Full Name      │ First Name           │ Last Name            │
2 │ John Smith     │ =TEXTBEFORE(A2," ")  │ =TEXTAFTER(A2," ")   │
  └────────────────┴──────────────────────┴──────────────────────┘
                      ↓                      ↓
                   "John"                 "Smith"
```

**Note:** If these functions aren't available in your Excel version, use LEFT/RIGHT/MID/FIND instead.

---

## Common Text Manipulation Patterns

### Pattern 1: Split Full Name
```
     A              B                          C
  ┌────────────┬──────────────────────────┬──────────────────────────┐
1 │ Full Name  │ First Name               │ Last Name                │
2 │ John Smith │ =LEFT(A2,FIND(" ",A2)-1) │ =RIGHT(A2,LEN(A2)-FIND(" ",A2))
  └────────────┴──────────────────────────┴──────────────────────────┘
```

### Pattern 2: Extract Email Parts
```
Email: john.doe@company.com

Username: =LEFT(A2,FIND("@",A2)-1)
         → "john.doe"

Domain: =MID(A2,FIND("@",A2)+1,LEN(A2))
       → "company.com"

Domain Name: =LEFT(MID(A2,FIND("@",A2)+1,LEN(A2)),FIND(".",MID(A2,FIND("@",A2)+1,LEN(A2)))-1)
            → "company"
```

### Pattern 3: Reverse Name (Last, First → First Last)
```
     A                  B
  ┌────────────────┬──────────────────────────────────┐
1 │ Name           │ Reversed                         │
2 │ Smith, John    │ =TRIM(RIGHT(A2,LEN(A2)-FIND(",",A2)))&" "&LEFT(A2,FIND(",",A2)-1)
  └────────────────┴──────────────────────────────────┘
                      ↓
                  "John Smith"
```

### Pattern 4: Remove Non-Numeric Characters
```
     A                  B
  ┌────────────────┬──────────────────────────────────┐
1 │ Phone          │ Numbers Only                     │
2 │ (555) 123-4567 │ =SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2,"(",""),")",""),"-","")," ",""),".","")
  └────────────────┴──────────────────────────────────┘
                      ↓
                 "5551234567"

Or use Text to Columns feature for easier approach
```

### Pattern 5: Proper Case with Exceptions
```
PROPER function capitalizes every word, but sometimes you want exceptions:

"the big apple" → should be "The Big Apple" (not "The big Apple")

Use nested SUBSTITUTE:
=SUBSTITUTE(SUBSTITUTE(PROPER(A2)," The "," the ")," Of "," of ")
```

---

## Text to Columns Feature

**Location:** Data tab → Text to Columns

**Purpose:** Split one column into multiple columns

Not a function, but essential for text manipulation!

### Use Case 1: Split by Delimiter
```
Before:
     A
  ┌────────────────┐
1 │ John,Smith,30  │
2 │ Alice,Jones,25 │
  └────────────────┘

After using Text to Columns (comma delimiter):
     A          B          C
  ┌─────────┬─────────┬─────────┐
1 │ John    │ Smith   │ 30      │
2 │ Alice   │ Jones   │ 25      │
  └─────────┴─────────┴─────────┘
```

### Use Case 2: Split Fixed Width
```
Before:
     A
  ┌────────────┐
1 │ John  Smith│
2 │ Alice Jones│
  └────────────┘

After using Text to Columns (fixed width at position 6):
     A          B
  ┌─────────┬─────────┐
1 │ John    │ Smith   │
2 │ Alice   │ Jones   │
  └─────────┴─────────┴─────────┘
```

### Steps to Use Text to Columns:
1. Select the column with data
2. Go to **Data** tab → **Text to Columns**
3. Choose **Delimited** or **Fixed Width**
4. Follow the wizard
5. Click **Finish**

---

## Common Text Function Mistakes

### Mistake 1: Forgetting Quotes Around Text
```
❌ Wrong: =SUBSTITUTE(A2,Hello,Hi)
✅ Right: =SUBSTITUTE(A2,"Hello","Hi")

Text must be in quotes!
```

### Mistake 2: Case Sensitivity Confusion
```
FIND is case-sensitive:
=FIND("hello","Hello World")  → #VALUE! (not found)

SEARCH is NOT case-sensitive:
=SEARCH("hello","Hello World")  → 1 (found)
```

### Mistake 3: Not Handling Errors
```
❌ Problem: =FIND("@",A2)
If A2 doesn't contain @, returns #VALUE!

✅ Solution: =IFERROR(FIND("@",A2),"No @ found")
```

### Mistake 4: Extra Spaces
```
" Hello " ≠ "Hello"

Always use TRIM:
=TRIM(A2)

Or in comparisons:
=TRIM(A2)="Hello"
```

### Mistake 5: Mixing Text and Numbers
```
❌ Wrong: ="Total: " & 1234.5
Result: "Total: 1234.5" (no formatting)

✅ Right: ="Total: " & TEXT(1234.5,"$#,##0.00")
Result: "Total: $1,234.50"
```

### Mistake 6: Position Count Starting at 0
```
❌ Wrong thinking: String positions start at 0
✅ Right: In Excel, positions start at 1

"Hello"
 12345  ← Positions are 1-based
```

---

## Best Practices

### 1. Use Helper Columns
```
❌ Complex (hard to debug):
=PROPER(TRIM(SUBSTITUTE(A2,","," ")))

✅ Better (step by step):
B2: =SUBSTITUTE(A2,","," ")  (Replace comma)
C2: =TRIM(B2)                (Remove spaces)
D2: =PROPER(C2)              (Title case)

Then hide columns B and C if needed
```

### 2. Combine TEXT Functions with IFERROR
```
=IFERROR(MID(A2,FIND("@",A2)+1,LEN(A2)),"No email")

Handles cases where @ doesn't exist
```

### 3. Use Named Ranges for Readability
```
❌ Hard to read:
=LEFT(A2,FIND(" ",A2)-1)

✅ Clearer:
=LEFT(FullName,FIND(" ",FullName)-1)
```

### 4. Document Complex Text Formulas
```
Add cell comments explaining:
- What the formula does
- Why certain positions/lengths are used
- Expected input format
```

### 5. Test with Edge Cases
```
Always test with:
- Empty cells
- Single words (no spaces)
- Extra spaces
- Special characters
- Very long text
- Very short text
```

---

## Real-World Application: Data Cleaning System

Let's build a complete data cleaning workflow.

### Messy Data Import
```
     A                      B              C
  ┌────────────────────┬──────────────┬──────────────┐
1 │ Name (raw)         │ Email (raw)  │ Phone (raw)  │
2 │ "  JOHN  SMITH  "  │ John@CO.com  │ (555)123-4567│
3 │ "alice johnson"    │ alice@co.COM │ 555.123.4567 │
4 │ "  Bob Lee  "      │ BOB@CO.COM   │ 5551234567   │
  └────────────────────┴──────────────┴──────────────┘
```

### Cleaning Formulas

**Column D (Clean Name):**
```
D2: =PROPER(TRIM(A2))
```

**Column E (Clean Email):**
```
E2: =LOWER(TRIM(B2))
```

**Column F (Clean Phone):**
```
F2: =SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(C2),"(",""),")",""),"-",""),".","")
```

**Column G (Formatted Phone):**
```
G2: ="("&LEFT(F2,3)&") "&MID(F2,4,3)&"-"&RIGHT(F2,4)
```

### Results
```
     D              E              F            G
  ┌────────────┬──────────────┬────────────┬────────────────┐
1 │ Name       │ Email        │ Phone #    │ Phone Format   │
2 │ John Smith │ john@co.com  │ 5551234567 │ (555) 123-4567 │
3 │ Alice Johnson│alice@co.com│ 5551234567 │ (555) 123-4567 │
4 │ Bob Lee    │ bob@co.com   │ 5551234567 │ (555) 123-4567 │
  └────────────┴──────────────┴────────────┴────────────────┘
```

### Add Validation

**Column H (Validation):**
```
H2: =IF(AND(LEN(D2)>0,ISNUMBER(SEARCH("@",E2)),LEN(F2)=10),"✓","Check Data")
```

---

## What to PRACTICE vs MEMORIZE

### Memorize
- & (ampersand) joins text
- LEFT, RIGHT, MID extract parts of text
- LEN returns text length
- TRIM removes extra spaces
- UPPER, LOWER, PROPER change case
- FIND is case-sensitive, SEARCH is not
- SUBSTITUTE replaces text
- TEXT converts numbers to formatted text
- Text must always be in quotes in formulas
- Positions start at 1, not 0

### Practice Deeply
- Combining text with & and CONCATENATE
- Extracting parts of text with LEFT/RIGHT/MID
- Using FIND/SEARCH to locate text positions
- Cleaning data with TRIM and SUBSTITUTE
- Converting case with UPPER/LOWER/PROPER
- Using TEXT to format numbers
- Building formulas that split names
- Extracting email components
- Cleaning phone numbers
- Creating complex text transformations
- Combining multiple text functions
- Handling errors in text formulas

### Don't Memorize
- Every possible text format code
- Complex nested formulas (build step by step)
- All edge cases (test as you encounter them)
- Exact character codes (look up CHAR codes when needed)

---

## Quick Reference: Text Functions

### Combining Text
```
& operator:        =A2&" "&B2
CONCATENATE:       =CONCATENATE(A2," ",B2)
CONCAT:           =CONCAT(A2:C2)
TEXTJOIN:         =TEXTJOIN(", ",TRUE,A2:C2)
```

### Extracting Text
```
LEFT:             =LEFT(A2,5)
RIGHT:            =RIGHT(A2,5)
MID:              =MID(A2,3,5)
LEN:              =LEN(A2)
```

### Finding Text
```
FIND:             =FIND("@",A2)        (case-sensitive)
SEARCH:           =SEARCH("@",A2)     (not case-sensitive)
```

### Changing Case
```
UPPER:            =UPPER(A2)
LOWER:            =LOWER(A2)
PROPER:           =PROPER(A2)
```

### Cleaning Text
```
TRIM:             =TRIM(A2)
CLEAN:            =CLEAN(A2)
SUBSTITUTE:       =SUBSTITUTE(A2,"old","new")
```

### Converting
```
TEXT:             =TEXT(A2,"$#,##0.00")
VALUE:            =VALUE(A2)
```

### Comparing
```
EXACT:            =EXACT(A2,B2)       (case-sensitive)
```

### Repeating
```
REPT:             =REPT("*",5)
```

---

## Troubleshooting Text Formulas

### Error: #VALUE!
**Causes:**
- Using FIND and text not found
- Using MID with position beyond text length
- Wrong argument types

**Solution:**
```
Wrap in IFERROR:
=IFERROR(FIND("@",A2),"Not found")
```

### Error: Wrong Results
**Causes:**
- Case sensitivity (FIND vs SEARCH)
- Extra spaces
- Position counting

**Solution:**
```
Use TRIM to remove spaces:
=FIND("@",TRIM(A2))

Use SEARCH for case-insensitive:
=SEARCH("text",A2)
```

### Text Looks Like Number But Won't Calculate
**Cause:** Number stored as text

**Solutions:**
```
1. Use VALUE: =VALUE(A2)
2. Multiply by 1: =A2*1
3. Use Text to Columns feature
```

---

## Next Step

After mastering text functions, you're ready to explore:

**`08-date-and-time-functions.md`**
- Understanding how Excel stores dates
- TODAY, NOW functions
- DATE, TIME functions
- DATEVALUE, TIMEVALUE
- Date arithmetic and calculations
- YEAR, MONTH, DAY extraction
- WEEKDAY, WORKDAY functions
- Date formatting and display
- Working with time zones and durations
