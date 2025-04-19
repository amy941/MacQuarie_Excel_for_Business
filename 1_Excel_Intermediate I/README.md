# WEEK 1: 
# 🔗Link: [Week 1_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250218_Week%201)
### - Multiple Worksheets
- Copy Excel sheet: right click--> move or copy /or hold Ctrl key down, drag, release.
- Cannot undo the sheet delete
- **color coding sheets**: right click--> Tab color
- **group sheet**: selecting multiple sheets simultaneously.
  
### - 3D Formulas
- Shift: =SUM(Sean:Carlos!C7)
- Limit moving Sheets around too much when working with 3D formulas ⚠️
- The structure of the workbooks must be identical ⚠️
  
### - Linking Workbooks
- ```View```--> Arrange All--> Tiled
- Once those workbooks start moving or being renamed, the links can get damaged. Check links: Data--> Edit Links
  
### - Consolidating by Positions
- ```Data```--> Consolidate
- This is, however, a **snapshot**, no formula, simply the value at the time that the consolidation was run.
- If there's an update, need to rerun Consolidate.
- Cannot undo a consolidation of links
  

### - Consolidating by Reference (category)
- ```Data```--> Consolidate. Choose labels (Top row/Left column)
  
💥 **- Week 1_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250218_Week%201/W1_PracticeChallenge_HeadOffice.xlsx)

💥 **- Week 1_Advanced Practice Challenge:** [adv challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250218_Week%201/W1_AdvPracticeChallenge.xlsx)

💥💥 **- Week 1_Assessment:** [assessment_Week 1](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250218_Week%201/assessment)

---

# WEEK 2
# 🔗Link: [Week 2_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250225_Week%202)
### - Combining text (CONCAT, &):
  
  *FirstName.LastName@pushpin.com* = CONCAT(C4,".",B4,"@pushpin.com") /or = C4&"."B4&"@pushpin.com"

### - Text case (UPPER, LOWER, PROPER):

  =PROPER(CONCAT(C4," ",B4))

  =LOWER(C4&"."&B4&"@pushpin.com")
  
### - Extracting text (LEFT, MID, RIGHT):

  =LEFT(**text**,[num_chars]) = LEFT(K4,2)
  
  =RIGHT(**text**,[num_chars]) = RIGHT(K4,4)

  =MID(**text**,start_num,[num_chars]) = MID(K4,4,4)

  
### - Finding text (FIND)

  = FIND(**find_text**, within_text, [start_num]) = FIND(" ", K4)-4
  
  =CONCAT(RIGHT(Inventory!F4,3),MID(Inventory!F4,FIND(",",Inventory!F4)+2,4)) --- nesting function
  
### - Date calculation (DATE, NOW, TODAY, YEARFRAC)

  = YEARFRAC(start_date, end_date) = YEARFRAC(F4,TODAY())
  
💥 **- Week 2_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250225_Week%202/C2-W2-Practice-Challenge.xlsx)

💥💥 **- Week 2_Assessment:** [assessment_Week 2](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250225_Week%202/C2-W2-Assessment-Workbook.xlsx)

---

# WEEK 3
# 🔗Link: [Week 3_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250228_Week%203)
### - Names Ranges:
  
  =N4*Pension_Rate
  
### - Create and manage ranges
  
  =AVERAGE(Annual_Salary)
  
  =MIN(Next_Review)
  
  =MAX(Date_of_Hire)
  
### - Apply ranges to formulas
  
💥 **- Week 3_Practice Challenge:** None🚫

💥💥 **- Week 3_Assessment:** [assessment_Week 3](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250228_Week%203/C2-W3-Assessment-Workbook.xlsx)

---

# WEEK 4
# 🔗Link: [Week 4_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250301_Week%204)
### - COUNT funct.

- **COUNT():** only counts the number of occurrences of cells that contain **NUMERIC** values. However, COUNT() still works on **dates**.

- **COUNTA()** only counts the number of cells that contain **alpha/numeric/alphanumeric** values.

- **COUNTBLANK():** only counts the number of occurrences of **blank** cells.

### - Counting w Criteria (COUNTIFs)

- COUNTIFS(**criteria_range1,** criteria1,...) = COUNTIFS(**State,** A5)
  
- Others:

  =COUNTIFS(**State,** A5)
  
  =COUNTIFS(**Order_Year,** 2013)
  
  =COUNTIFS(**Order_Priority,** "High") ⚠️ place a string in " "
  
  =COUNTIFS(**Order_Quantity** ">40") ⚠️ place a condition in " "
 
### - Adding w Criteria (SUMIFs)

- SUMIFS(sum_range, criteria_range1, **criteria1**, [criteria_range2, criteria2], [criteria_range3, criteria3]...)
  
  = SUMIFS(Total, Account_Manager, A21)
  
  = SUMIFS(Total, Account_Manager, **$A21**, Order_Year, **C$20**)
  
### - Sparklines

- Both **Row data** and **Column Data** can be used to create a single Sparkline
  
-  Features can be highlighted on a sparkline: high point, low point, first point, last point, markers, negative points


### - Advanced Charting

- Switching row & column
- Selecting data,
- Changing chart type
- Adding a secondary axis

  
### - Trendlines
  
💥 **- Week 4_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250301_Week%204/C2-W4-Practice-Challenge.xlsx)

💥💥 **- Week 4_Assessment:** [assessment_Week 4](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250301_Week%204/C2-W4-Assessment-Workbook.xlsx)

---

# WEEK 5
# 🔗Link: [Week 5_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250302_Week%205)
### - Create and format tables
- Ctrl + T
- ```Table Design``` tab on ribbon
- In tables, named ranges will automatically extend when you add an extra row/column.

### - Working w Tables
- Selecting cells: **Ctrl + Shift + -->** doesn't work when you encounter **a blank cell**. Same for both table and range.
  
  To select just your data in a row, drag /or use **black arrow** /or **Ctrl+A** /or move your mouse to the Top Left corner until it turns into a little black diagonal arrow (clicks 2x to include the header)
- In range, **Ctrl+A** doesn't work if you have broken that block of contiguous data. But, works on Tables because Excel remembers that table represents one entity so it will select all of your table.
- Once you convert data into a table, it effectively converts the rows in your table into database records. All the information in that record belongs together and Excel treats the cells accordingly.
- ```Conditional Formatting```: higlight duplicate values --> ```Table Design``` tab--> Remove Duplicates
- No need to worry about *freezing panes* when working w tables as the headers are still visible when we scroll down.
  
### - Sort and filter in tables
- **Total Rows**: a row is added at the bottom of table, providing a range of calculations (sum, avg, count, min, max, etc.). Also, able to update according to the filtered data in your table.
- It's a good practice to always clear the Filter when working in a shared environment.
  
### - Automation
- To insert a row: Ctrl +
- To enter today's date: Ctrl + ;
- Table updated the named ranges when a new row is added to the table. But, doesn't work for Range.
- Press Tab once in the last cell of the last table row.
- **Structured References**: =[@[Annual_Salary]] +[@Pension]. Create auto in a table, work similarly to name ranges but they're not absolute.

### - Subtotalling
- ```Data``` tab --> Subtotal. Convert Table to Range first.
- Before converting to Range, Sort first. Then, remove Banded Rows and turn off the Total Row. Finally, click ```Convert to Range```
- **Subtotal** feature is not supported by Tables because we have summary data mixed with data entries.

💥 **- Week 5_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250302_Week%205/C2-W5-Practice-Challenge.xlsx)

💥💥 **- Week 5_Assessment:** [assessment_Week 5](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250302_Week%205/C2-W5-Assessment-Workbook.xlsx)

---

# WEEK 6
# 🔗Link: [Week 6_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/tree/main/20250313_Week%206)
### - Create & Modify a Pivot table
- Convert data into **a Table** before performing Pivot Table.
- Click ```Refresh``` to update new values.
- Drag Fields between areas below: Filters, Columns, Rows, Values

### - Value Field Settings
- Right Click--> Summarize Values by --> Sum, Count, Avg, etc.
- Right Click--> Show Values As--> No Calculation, %...
- Right Click--> Number Format

- ```Design``` tab: Subtotals, Grand Totals, Report Layout,...

### - Sort & Filter a Pivot table
- ```PivotTable Analyze``` tab: Ungroup, Group Field, Expand Field, Collapse Field, ...
- Can Sort data in Pivot as normal.

### - Report Filter Pages
- Drag Fields between areas below: Filters
- ```PivotTable Analyze``` tab: Options --> Show Reports Filter Pages...

### - Pivoting Charts
- ```PivotTable Analyze``` tab--> Pivot Chart
- ```Format```, ```Design```, ```PivotTable Analyze``` tabs
- Pivot charst always represents the pivot table. If pivot table changes, pivot chart will change and vice versa.

### - Pivoting Sliders
- ```PivotTable Analyze``` tab: Insert Slicer
- Slicer can connect slicers to multiple pivots. ```Options``` tab --> Report Connections

  
💥 **- Week 6_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250313_Week%206/C2-W6-Practice-Challenge.xlsx)

💥💥 **- Week 6_Assessment:** [assessment_Week 6](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250313_Week%206/C2-W6-Assessment-Workbook.xlsx)

---

# Final Course Assessment: [Final Course Assessment](https://github.com/amy941/MacQuarie_Excel_Intermediate-I/blob/main/20250313_Week%206/C2-Final-Assessment.xlsx)

---
# CERTIFICATE

![Image](https://github.com/user-attachments/assets/96377273-fd75-48b2-886a-321b185e74b9)





