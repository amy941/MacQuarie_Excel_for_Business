# WEEK 1: 
# 🔗Link: [Week 1_folder](https://github.com/amy941/MacQuarie_Excel_for_Business/tree/main/3_Excel_Advanced/Week%201)
### - Spreadsheet Design Principles
- **Accurate, Flexible, Responsive, Easy to Maintain, User Friendly.**
- **Keep raw data separate** to ensure your original input stays intact and don't end up damaging it.
- Consider **leaving at least 2-3 rows at the Top** to later on for advanced filters, summary calc., or headings.
- Important to **keep related data in a continuous table**, if fail to do, it's impossible to use features like Pivot Tables, and subtotalling.
- **Don't put text data in columns** that you later on want to perform calculations.
- **Organize similar info stays together**
- **Don't put any blank rows or columns inside a single data set.**
- **Avoid merge cells.**
- Good practice with raw data is to **put it into a table.** Name a table **"tbl_..."**
- Use **Name Ranges**

### - Calculations
- **Calculations must be auditable and transparent**
- **VLOOKUP limitations**: going to break if columns move around/ not auditable and quite inefficient
- **INDEX MATCH, MONTH, YEAR, DATE**
- **EOMONTH**: End-Of-Month
  * =EOMONTH(B3, **-12**) **+1**--> **-12**: get us the last data in the month 12 months ago, **+1:** get the first date of the following month.
  
- **EDATE**
  
### - Formatting
- **Pivot Tables limitations:** unless they are refreshed, they won't reflect updated data/ serious performance issues occur if the data is too large/ hard to share
- When working with **grouped sheets**, ensure to **ungroup them** as soon as you've finished your task at hand.
  
### - Documentation
- Use standard **naming conventions** for naming everything
- Use **sensible headings** where possible
- Take advantage of **Excel's tools** (data validation, comments..)  

### - Interface & Navigation


💥 **- Week 1_Practice Challenge:** None 🚫

💥💥 **- Week 1_Assessment:** [assessment_Week 1](https://github.com/amy941/MacQuarie_Excel_for_Business/blob/main/3_Excel_Advanced/Week%201/C4-W1-Final-Assessment.xlsx)

---

# WEEK 2
# 🔗Link: [Week 2_folder](https://github.com/amy941/MacQuarie_Excel_for_Business/tree/main/3_Excel_Advanced/Week%202)
### - Tables & Structured Referencing
- =ROWS(Sales)
- **Structured References**:
  * looks like Named Ranges. Inside square brackets [...], we can list parts of the table we want to work with.
  * **"&" symbol** indicates current row --> =SUM(Sales[@[Australia]:[China]])
  * **Use Headers**--> =Sales[[#Headers], [Australia]]
  * are **Mixed Cell References**

- LARGE(array, k) =LARGE(Sales[Australia])

### - Using Functions to Sort Data
- **Link URL into Excel file:** ```Data``` tab--> Get Data--> From Web--> paste URL--> click Go. Pick table--> Import--> Ok
- **Sort data**:
  * use COUNTIFS (=COUNTIFS(rateCodes, "<="&'Current Rates'!D4))
  * INDEX(MATCH(ROW_))
    
  
### - Introduction to Array Formulas
- **CSE formula:** Ctrl+Shift+Enter
- **Multi-cell array formula:** use a single calculation to return multiple results.
  * Start by selecting all the cells we want the answers to go
  * Type "=", then select the first set of values (**the whole array**) we want to multiply
  * Multiply "*"...
  * Ctrl+Shift+Enter

- **Array formula:**
  * Always shows in **curly braces** {_}
    {=C4:C36*H4}
  * cannot delete part of array formula. It can update, but has to be all or nothing.

- **Single Cell formula:** =SUM(H7:H9*I7:I9), then Ctrl+Shift+Enter
- **array constant:** an array of values that you specify they are constant rather than referring to cell references.
  * done by putting the values in curly braces, and separating them with semi-colons
    =LARGE(D4:D36, {1; 2; 3}), then Ctrl+Shift+Enter


### - Working w Array Function (TRANSPOSE)
- **TRANSPOSE function:**
  * Start by selecting all the cells we want the values to go,
  * =TRANSPOSE(select all the values we want to swing thro 90 degree)
  * Ctrl+Shift+Enter

  
### - Solving Problems w Array Formulas
- **If returning ERRORs:** =SUM(**IFNA**(H7:H17*I7:I17, 0))
  
💥 **- Week 2_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_for_Business/blob/main/3_Excel_Advanced/Week%202/C4-W2-Practice-Challenge.xlsx)

💥💥 **- Week 2_Assessment:** [assessment_Week 2](https://github.com/amy941/MacQuarie_Excel_for_Business/blob/main/3_Excel_Advanced/Week%202/C4-W2-Final-Assessment.xlsx)

---

# WEEK 3
# 🔗Link: [Week 3_folder](https://github.com/amy941/MacQuarie_Excel_for_Business/tree/main/3_Excel_Advanced/Week%203)
### - Replace blanks with repeating values
- **Method 1: Go-to special**
  * ```Find & Select``` tab--> Go To Special--> Blanks (find the blanks in that area and select them)--> Ctrl+V
  * Problem: we've lost our original data

- **Method 2: Use Calculation**
  * IF
  * ISBLANK
    =IF(ISBLANK(Sheet1!A2), A1, Sheet1!A2)
  
  
### - Fix Dates (DATE, MONTH, YEAR, DAY, TEXT)
-  **TEXt** function:
  * =TEXT(value, format_text) = TEXT(J1, "mmmm")  ⚠️dd=17, ddd=Tue, dddd=Tuesday

- **DATE** function:
  * =DATE(year, month, day) = DATE(LEFT(Sheet1!F2,4), MID(Sheet1!F2,6,2), RIGHT(Sheet1!F2,2))
 
  
### - Remove Unwanted Spaces (TRIM, CLEAN)
- **TRIM** function: removes all spaces from text except for single spaces between words. Ex:= =TRIM(CLEAN(Data!B2))
- **CLEAN** function: to remove non-printable characters from a text string. Ex: =CLEAN(TRIM(Data!B2))

### - Diagnostic Tools (ISNUMBER, LEN, CODE)
- **ISNUMBER:** check whether a value is a number, returns TRUE or FALSE
- **CODE:** if we apply CODE function to a string of more than 1 character in length, we would get the code result of the first character in the string.


### - Remove Unwanted Characters (SUBSTITUTE, CHAR, VALUE)
- **SUBSTITUTE** will remove spaces between words which we may want to keep, whereas the **TRIM** function will maintain single spaces between words.
- **SUBSTITUTE** is typically **"all or nothing"** when used in simple formulas.

- **=VALUE(A1)** vs. **=VALUE(VALUE(A1))** will produce the same results

- **CHAR** 

💥 **- Week 3_Practice Challenge:**  [challenge]()

💥💥 **- Week 3_Assessment:** [assessment_Week 3]()

---

# WEEK 4
# 🔗Link: [Week 4_folder]()
### - Working with Dates (EOMONTH, EDATE, WORKDAY.INTL)
- **EOMONTH(start_date, months)**: Returns the **last day of the month** that is the specified number of months before or after the start date.
- **EDATE(start_date, months)**: Returns the **same day of the month**, but a specified number of months before or after the start date.
- **WORKDAY.INTL(_)**:  Returns a date after adding/subtracting **workdays**, excluding weekends and optionally holidays.

### - Financial Functions (FV, PV, PMT)
- **Future Value:** **FV**(rate, nper, pmt, [pv], [type])
- **Present Value:** **PV**(rate, nper, pmt, [fv], [type])
- **Payment**: **PMT**(rate, nper, pv, [fv], [type])

 
### - Loan Schedule (PMT, EDATE)
- the optional **[FV]** argument in the **PMT** function allows us to specify the balance to be attained after the last payment is made.
- **PMT** function is used when we want to calc. the fixed monthly repayments of a car loan.

  
### - Net Present Value & Internal Rate of Return (NPV, IRR)
- **Net Present Value:** **NPV**(rate, value1, [value2], …). _Purpose:_ Calculates the present value of a series of **future cash flows**, discounted **at a given rate.**
- **Internal Rate of Return:** **IRR**(values, [guess]). _Purpose:_ Calculates the **discount rate** that makes the NPV of cash flows equal to zero.


### - Depreciation Functions (SLN, SYD, DDB)



💥 **- Week 4_Practice Challenge:** None 🚫

💥💥 **- Week 4_Assessment:** [assessment_Week 4]()

---

# WEEK 5
# 🔗Link: [Week 5_folder]()
### - INDIRECT
- **INDIRECT** funct uses **A1-style** referencing
- **INDIRECT** funct contains:
  * a reference to a cell as a text string
  * **A1-style** referencing
  * **R1C1-style** reference
  * a named defined as a ref
  * a formula

### - ADDRESS
- **ADDRESS:** returns the **cell address** (as a text string) for a cell at a specified **row and column number**.

### - Intro to OFFSET
- **OFFSET**(reference, rows, cols, [height], [width])
- Returns a **cell or range** that is a specific number of **rows and columns** away from a starting reference.
- It's super handy for **dynamic ranges** and **lookups**. 

### - Solving Problems w OFFSET



💥 **- Week 5_Practice Challenge:** None 🚫

💥💥 **- Week 5_Assessment:** [assessment_Week 5]()

---

# WEEK 6
# 🔗Link: [Week 6_folder]()
### - Dashboard Design
- **Segmentation** in dash board is important because it allows us to display info in groupings.
- **Shapes** and **Background fill** are really flexible and useful in dashboard design.

### - Prepare Data


### - Construct Dashboard


### - Creative Charting


### - Interactive Dashboard



  
💥 **- Week 6_Practice Challenge:** None 🚫

💥💥 **- Week 6_Assessment:** [assessment_Week 6]()

---


# CERTIFICATE

![final_cert](https://github.com/user-attachments/assets/0eee8749-d9b1-4974-b8b1-39ecf9f0cb79)
