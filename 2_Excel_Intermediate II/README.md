# WEEK 1: 
# 🔗Link: [Week 1_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%201)
### - Data Validation
- ```Data``` tab--> Data Validation (Settings/Input Message/ Error Alert)
   - *Settings*: define the validation criteria. The default allows ```any value```, meaning **no validation** is taking place. Drop-down shows: Any value, Whole number, Decimal, List, Date, Time, ...
     ```Ignore blank```: Excel won't consider a blank cell to be invalid.
     
   - *Input Message*:
 
   - *Error Alert*: Stop 🚫, Warning ⚠️, Information ℹ️

- If data is **copy-pasted**, or **imported**, it actually **doesn't enforce** data validation rules. **Only works for data that's been entered manually.**
- **Text length** refers to any characters, or combination of text and numeric characters.

### - Create Drop-down Lists
- ```Data``` --> ```Data Validation``` --> Settings --> **Allow: List** | Source: better type in alphabetically
- Converting lookup list into a named range and table so we don't need to update the validation criteria as the look-up list changes.
- **Drop-down list**, items should be seperated by **comma** or **comma and Space**

### - Using Formulas in Data Validation
- **Duplicate code:** ```Data``` --> ```Data Validation``` --> Settings --> **Allow: Custom** | **Formula: =countifs(Product_Code,A4) <= 1**
- **Allow** in **Data Validation** use a formula: **Custom**, **List**
    
### - Working w Data Validation
- ```Data Validation``` drop-down: Circle Invalid Data ⭕
- ```Find & Select``` tab --> Go to Special... --> Data Validation: All or Same
- **Copy data Validation** from one sheet to another: **Paste Special**

### - Advanced Conditional Formatting
- ```Conditional Formatting``` --> New Rule...--> "Use a formula to determine which cells to format" --> **Format values where this formula is true:** = H4 < J4 (w/o $ signs) --> Preview: Format (Font:Bold, Fill:Color)
- ```Conditional Formatting``` --> New Rule...--> "Use a formula to determine which cells to format" --> **Format values where this formula is true:** = **$E4** = $O$4 **(⚠️ Row to go Relative while Column remain Abs)** --> Preview: Format (Fill:Color)
  
💥 **- Week 1_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%201/C3-W1-Practice-Challenge.xlsx)

💥💥 **- Week 1_Assessment:** [assessment_Week 1](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%201/C3-W1-Assessment.xlsx)

---

# WEEK 2
# 🔗Link: [Week 2_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%202)
### - Logical Functions I: IF

**=IF(logical test, [value_if_true], [value_if_false])**
- First argument is **a logical test**, compares 2 values using a **logical operator**
  ![logical operator](https://github.com/user-attachments/assets/f9586ddc-cc68-4062-bd26-f23d58c0051b)
  
  
- Second argument in brackets is the **"value_if_true"**, could be a value we just type into the cell /or a calculated value.
  * if the logical test equates to True, then whatever we've got between two commas will occur.
  * if the logical test equates to False, then it's going to do the third and last argument **"value_if_false"**

- If working w text, put double quotes **" "** /or quotation marks **' '** /or **""** (leave Blank) 
- When comparing text, the equals is **not case sensitive**
- =IF(F4="Y",D4*5%,0)

### - Logical Functions II: AND, OR

**=AND(logical1, [logical2], ...)
  =OR(logical1, [logical2], ...)
  Up to 255 logical tests❗,
  Only returns TRUE/FALSE**

- **=AND(logical1, [logical2], ...)**
  * =AND(B4>0,C4<>"Y")
  * evaluate multiple logical tests
  * If x & y & z & ... are **ALL** True, then it returns True

- **=OR(logical1, [logical2], ...)**
  * =OR(I4>=16, J4)
  * If **any** of these are True: x,y,z,..., then returns True


### - Combining Logical Functions I: IF, AND, OR

**=IF(AND(logical1, logical2, ...), [value_if_true], [value_if_false])
  =IF(OR(logical1, logical2, ...), [value_if_true], [value_if_false])**

- =IF(AND(B4>0,C4<>"Y"),B4*10%,0)
- =IF(OR(I4>=16,J4),250,0)


### - Combining Logical Functions II: Nested IFs
![nested IF](https://github.com/user-attachments/assets/f16df025-17bb-4866-a423-f19d984edc3b)

**=IF(Balance= 0, "A", IF(Balance > 0, "B", "C"))**

### - Handling Errors: IFERROR, IFNA
- =IFERROR(AVERAGE('Invoice Data'!$O$4:$O$654),"")
- =IFNA(VLOOKUP('Invoice Data'!$A4,BPay!$B$4:$D$10,3,0),0)

💥 **- Week 2_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%202/C3-W2-Practice-Challenge.xlsx)

💥💥 **- Week 2_Assessment:** [assessment_Week 2](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%202/C3-W2-Final-Assessment.xlsx) 

---

# WEEK 3
# 🔗Link: [Week 3_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%203)
### - Introduction to lookups: CHOOSE
- **CHOOSE**: retrieving a value from a list based on a given numeric value.
  
  =CHOOSE(**index_num**, value1, [value2], ...)
  
  =CHOOSE(**[@[Loc Code]]**, $K$6,$K$7,$K$8,$K$9,$K$10)
  
- ⚠️ have to individually list each list item

- **CHOOSE** function can handle up to **254 values** for the list specified.


### - Approximate Matches: Range VLOOKUP
- **VLOOKUP**: given a value, it will go and try and match it into a master dataset. When it finds a match, it will then return a corresponding value from the same row.
- **V means VERTICAL**, can only be used on lookup data that is organized vertically.
- VLOOKUP has 2 slight variations --- **a range lookup & an exact match**
- ⚠️ a range lookup: your data must be organized **smallest to largest**

- =VLOOKUP(**lookup_value**, table_array, col_index_num, [range_lookup])
  =VLOOKUP(**D4**,$G$4:$H$7,2)
  
  * table_array: data block, not just a column or row || DON'T include the headings |} make it ABSOLUTE Reference 
  * col_index_num: number of column that contains the value we want Excel **to return** from the lookup table

### - Exact Matches: Exact Match VLOOKUP
- =VLOOKUP(**lookup_value**, table_array, col_index_num, [range_lookup])
  =VLOOKUP(**[@Item]**,'International Price List'!$A$4:$E$1254,3,0)
  
  * **0** or **FALSE** means **exact match**

- Less cumbersome version --> to put your lookup data either in a **named range** or a **table**. Benefit of a table: table is auto grow if new row is added to the bottom
  
  =VLOOKUP(**lookup_value**, table_array, col_index_num, [range_lookup])
  
  =VLOOKUP(**[@Item]**,parts,3,0)
  

### - Finding a Position: MATCH
- **MATCH**: given a particular value, it will go and look it up in either a row or a column. It's not worried about horizontal, vertical. When it finds a match, instead of returning a corresponding value, however, it will **return the position** within that column or row.
  
 - =MATCH(**lookup_value**, lookup_array, [match_type])
   =MATCH(**Table2[[#Headers]**,[Short Description]],parts[#Headers],0)
   
   ⬇️ ⬇️ ⬇️
   
   =VLOOKUP([@Item],parts,**E$4**,FALSE) -- **E$4** is a **mixed reference**
   
   ⚠️ with structured references, when **dragging** VLOOKUP formulas across, it actually treats it as a **relative reference**
      To solve it, **Copy + Paste(formulas,fx)** (it's absolute!!!)


### - Dynamic Lookups: INDEX, MATCH
- =INDEX(**array**, row_num, [column_num])
  * array: can be a single column /or a single row /or an entire matrix.
  * row_num:
  * colum_num: 0 (exact match)
 
- =INDEX(Discounts,**MATCH(C11,Discount_Categories,0)**, **MATCH(D11,Customer_Categories,0)**)

  ⬇️ ⬇️ ⬇️
  
- =INDEX(Discounts,MATCH([@Category],Discount_Categories,0), $I$4)

- **Other benefit is unlike VLOOKUP** where your lookup column must **sit to the left** of the value you want to return.

  The **INDEX MATCH** has no such restriction --> more versatile. It also has the advantage that you can split out the lookup from the return while VLOOKUP cannot.
 
💥 **- Week 3_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%203/C3-W3-Practice-Challenge.xlsx)

💥💥 **- Week 3_Assessment:** [assessment_Week 3](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%203/C3-W3-Final-Assessment.xlsx)

---

# WEEK 4
# 🔗Link: [Week 4_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%204)
### - Error Checking
- Error: **#N/A**, **#REF**, **#VALUE!**, **#DIV/0!**, **#NAME?**
  * Errors occur when typing mistakes, incorrect syntax, or invalid arguments.
  * To locate errors: Click ```Home``` tab --> ```Find & Select``` --> Go to Special... --> Formulas (✅ Errors) --> highlight errors
    
- ```Formulas``` tab --> Error Checking --> Show Calculation Steps... || Edit in Formular bar || Next    
  * 🛑 **VALUE!** occurs when you make a **typo** /or one of the **inputs is invalid.**
  * 🛑 **#N/A** means Excel tried to do a lookup but it **hasn't found the look up value.**
  * 🛑 **#REF** occurs either when you **copy paste** relative references to cells where they cannot refer to the correct values, /or they happen quite often with lookup errors **when you refer to a range that doesn't actually exist.**
  * **Potential errors**: Excel has flagged as looking like it might be problematic, eventhough it hasn't yet produced an error message.

- ```Formulas``` tab --> Show Formulas (shows all formulas in the Workbook)
   * 🛑 **#DIV/0!** happens when one of the **input cells is Blank** /or **contains a zero.**
   * 🛑 **#NAME?** occurs either **typed the function name in wrong** /or **forgotten double quotes** when working with text.

- **Trace errors**: ```Formulas``` tab --> Error Checking (click drop-down) --> Trace Error 

![trace error](https://github.com/user-attachments/assets/47772e8a-8e8d-466e-8117-77daa91f4794)

### - Formula Calc Options
- 🔁 **Circular references**: is when the calc. cell includes itself as part of that calc., and as a result, gets into an **infinite loop.**
  * They can also occur when a cell **indirectly references itself**, so it refers to another cell which refers to it.
  * ```Formulas``` tab --> Error Checking (drop-down) --> Circular References 🔁
    
- 🟢 **Green flag** error: not necessarily an error, but might be incorrect in some way. The most common reason is **an inconsistent formula.** -- the one that looks a bit different than the rest.
  * Fix **Inconsistent Formula**: ⚠️ Warning sign --> drop-down --> Copy Formula from Above

- Change **Error checking options**: Workbook Calculation --> Automatic / Manual
  * **Automatic**: everytime you make a change in your workbook, all the calc. will re-calculate.
  * **Manual**: works better when you want to make a small change and don't want to wait for long for Excel the re-calculate the whole Workbook.
    * set to **Manual** --> ```Formulas``` tab --> Calculate Now (force Excel to immediately recalculate all the values in Workbook)
    * set back to **Automatic** --> ```Formulas``` tab --> Calculation Options --> Automatic

### - Tracing Precedents & Dependents
- **Trace Precedents** is a cell that is referred to in a formula.
  
![trace_dependent](https://github.com/user-attachments/assets/a551ed27-82ea-43f2-ab90-87e74a4b6e43)
  
- **Trace Dependents** is a formula that refers to your cell.
  
![trace_dependent1](https://github.com/user-attachments/assets/ad5922d4-3b65-4ec1-a990-afee9191af18)

⬇️ ⬇️ ⬇️ CLICK TWICE

![trace_dependent2](https://github.com/user-attachments/assets/2c91457e-2095-4bc6-89cc-3336268475b2)

- **dash black arrow**: indicates we are referring to a value in a different worksheet.
  
![dash_black_arrow](https://github.com/user-attachments/assets/23ebed23-b18e-481f-931b-4f17bf54fa05)


### - Evaluate Formula, Watch Window
- ```Formulas``` tab --> Evaluate Formula
  * Evaluate: work through formula step-by-step
  * Step In: displays a formula that is referred to in your active formula in a separate box.

![evaluate_formula](https://github.com/user-attachments/assets/c6dabac2-157d-4bc3-a073-47eb382ff732)


- ```Formulas``` tab --> Watch Window (watch the cells even when we are not on the same worksheet)

![watch_windown](https://github.com/user-attachments/assets/79af3fc7-a039-4126-8a66-af33e2dab3cc)

### - Protecting Workbooks & Worksheets
- to prevent unauthorized access or accidental damage
- Protection can be added at 3 different levels
  * **at Workbook** itself.
  * **at the structural level**, so you can prevent people from adding, moving, or unhiding worksheets.
  * **at worksheet level** itself, where you can lock all the cells or just selected cells.
    
- **Method 1:** ```File``` tab --> Info --> 🔏 Protect Workbook --> Encrypt with Password (only allow 1 level of access)
  
- **Method 2:** ```File``` tab --> Save as --> Browse --> Next to **Save** button, press **Tools** --> General Options (allows to add two levels of password protection)
  * Password to open: add a password w/o which the user will not be able to open the workbook at all.
  * Password to modify: users can open the workbook, but only to view the contents.
  * users cannot be able to change the contents, meaning the file is **open read-only**
    
- **Method 3**: ```Review``` tab --> Protect Sheet || Protect Workbook
  * **Protect Workbook: protect the structure of the workbook**

    ![protect_workbook](https://github.com/user-attachments/assets/a0fe6a58-ec90-4cf2-bc83-b99d2cb869bd)

    * Impact: Structure of workbook is locked, means: cannot add New sheet ➕, cannot move the Sheet, right-click on any of the tabs nearly all the options have been grayed out.
    * Benefits: great for 3D cell references as we don't want the sheets to move around too much /or want sheets to remain hidden.
      
  * **Protect Sheet: lock down the contents of the sheet**
    
    ![protect_sheet](https://github.com/user-attachments/assets/5ee78cd9-bed7-4f34-94df-f01ce834841b)
    
    * Impact: locked the sheet but still can click & view the content of the cells.

   * **Protect Sheet** (unlock certain cells)

     ![protect_sheet_certain_cell](https://github.com/user-attachments/assets/0d66548b-647d-4c51-8f33-5539bede5d3f)
 
     ⬇️ ⬇️ ⬇️
 
     ![protect_sheet_certain_cel1](https://github.com/user-attachments/assets/d1769709-0632-4642-a6be-1139fdc5b7fa)

     * Impact: users cannot view the content/formulas behind the cells, cannot even click on them. 
     * Benefits: allow the users to change certain cells but locked down the ones that contain the formulas so those remain protected.
    
   * **Protect Sheet**: ```Review``` tab --> Allow Users to Edit Ranges (can specify certain ranges to have their own password) --> New

  
💥 **- Week 4_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%204/C3-W4-Practice-Challenge.xlsx)

💥💥 **- Week 4_Assessment:** [assessment_Week 4](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%204/C3-W4-Final-Assessment.xlsx)

---

# WEEK 5
# 🔗Link: [Week 5_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%205)
### - Modelling Functions: SUMPRODUCT
- **SUMPRODUCT()** finds the product of multiple arrays and then sums up the products.
  * **=SUMPRODUCT(**array1**, [array2], [array3], ...)**
      =SUMPRODUCT(B6:D6,$B$4:$D$4)/SUM($B$4:$D$4)
    
  * All the arrays must have the same number of rows & columns.❗❗❗
  * No need for a particular order ❗❗❗
    
- Narrow down data & add it up where a certain criterion is met
  * =SUMPRODUCT(1*(E6:E11>=100)) --> each result is going to return a True or False 
    

### - Data Tables
- **Data Tables** see a range of different outcomes using different inputs using just one formula
- ```Data``` tab --> What-if Analysis (drop-down) --> Data Table
  * a single variable
  * a dual variable

- **One-input Data Tables** (a single variable)

  ![single_variable](https://github.com/user-attachments/assets/42c818a0-4a59-4a21-8315-6c54a6f868c1) 

  * { =TABLE(,E17) } --> means we're dealing with a data table
  * ❗Cannot delete a single cell. If want to delete, have to select all of those results.

- **Two-input Data Tables** (a dual variable)

  ![dual_variable](https://github.com/user-attachments/assets/08f56c8a-a91e-477d-bcec-283f7517d254)

  * { =TABLE(E21,E23) }


### - Goal Seek
- **Goal Seek**: given a cell that has a calculation in it, it will adjust that cell to a specified value by changing one of the inputs that you provide.
  
- ```Data``` tab --> What-if Analysis --> Goal Seek
  * **Set cell**: is the cell that **contains your calculation** that you want to change the value it's returning.
  * **To value**: want this to equal to ...
  * **By changing cell**: which input we want to change to get to that result. Must be in a **typed-in value** rather than a formula ❗

![goal_seek_2](https://github.com/user-attachments/assets/8799bb47-0d9e-49a1-b5fa-9450287594f3)

⬇️ ⬇️ ⬇️

![goal_seek_2_result](https://github.com/user-attachments/assets/87d0295f-3ce0-471e-bd28-00d2dfe57fe0)
  

### - Scenario Manager
- **Scenario Manager**: allows you to keep different data inputs in a single worksheet.
- ‼️ Better to name your ranges prior to producing scenarios.
- ```Data``` tab--> What-If Analysis --> Scenario Manager


### - Solver
- **Solver:** allows you to model different situations with a variety of inputs and constraints and even integrate that with Scenario Manager so you can store different solutions.
- **Add Solver:** ```File``` tab--> Option--> Add-ins--> at the bottom: Manage: Excel Add-ins , Go--> Solver Add-ins--> OK
  * **Set Objective:** cell contains **formula** that you want to return a different result.
  * **To:** ✅Max , ✅Min , ✅Value Of: ....
  * **By Changing Variable Cells:** which our inputs can change



💥 **- Week 5_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%205/C3-W5-Practice-Challenge.xlsx)

💥💥 **- Week 5_Assessment:** [assessment_Week 5](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%205/C3-W5-Final-Assessment.xlsx)

---

# WEEK 6
# 🔗Link: [Week 6_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%206)
### - Record a Macro
- ➕ **Add macro:** ```File```--> Options--> Customize Ribbon--> Developer. You should now see the **Developer tab on the Ribbon**
- 📽️ **Record macro:** hit a Record button, give a name, specify where to store it, go thro all the steps we want Excel to remember us doing, then press **Stop.** ‼️
- Language: Visual Basic for Applications (VBA)
- 💾 **Save macro**: Save as... Excel Macro-Enabled Workbook **(.xlsm)**, Excel Binary Workbook (**.xlsb)**, Excel Macro-Enabled Template **(.xltm)**

### - Run a Macro
- 🏃🏻 **Run macro: **```Developer```--> Macros--> Run
- 🔘 **Create a Button to run macro:** ```Insert```--> Shape--> Right-click--> Assign Macro


### - Edit a Marco
- 🛠️ **Edit macro:** ```Developer```--> Macro--> Edit--> open up VBA Editor
  * Left-hand side: a Project Explorer - allows you to look at different macros in different open workbooks
  * Right-hand side: a code window
    
- All recorded macro **begins with a Sub** (subroutine), and **ends with an End Sub**
  * **All of your code sits between the Sub and the End Sub**
  * **Each line of code represents one instruction or one step that you performed**
  * some of the codes are GREEN and prefixed with an apostrophe are called **comments**, totally ignored by the compiler
    
![VBA editor](https://github.com/user-attachments/assets/0417f9ca-577c-4d28-83c7-b82bba9dd51c)

- **Add new lines to VBA Editor:**
  * Range("B2").Value = inputbox("Please enter week commencing date", "New Timesheet Date") --> Save 💾


### - Work w Marcos
- VBA code to widen the column: record yourself widening the column, then copy and paste that code into the original macro.

### - Relative Reference Macros
- ```Developer```--> Use Relative References--> Record Macro
- Ctrl + Home, Ctrl + down arrow, and down arrow
- ```Data```--> Get Data--> From File--> From Text

---

# CERTIFICATE

![Intermediate II_cert](https://github.com/user-attachments/assets/51f13bb1-61b9-45de-bee5-4beaa98e589e)
