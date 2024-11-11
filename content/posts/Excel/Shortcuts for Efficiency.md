Excel is a powerful tool for data management and analysis, but it's even more effective when you know how to use shortcuts and formulas to speed up your workflow. Below are some of my favorite Excel shortcuts, formulas, and tips that help me work efficiently and solve common problems in Excel.

---
#### **Navigation and Selection Shortcuts:**

|**Action**|**Shortcut**|
|---|---|
|Select the entire row|`Shift + Space`|
|Select the entire column|`Ctrl + Space`|
|Select all data in a worksheet|`Ctrl + A`|
|Select a range of cells|`Shift + Arrow Keys`|
|Extend selection to the last non-empty cell|`Ctrl + Shift + Arrow`|
|Move to the next cell|`Tab`|
|Move to the previous cell|`Shift + Tab`|
|Move to the last cell in a row|`Ctrl + Right Arrow`|
|Move to the last cell in a column|`Ctrl + Down Arrow`|
|Go to the beginning of a worksheet|`Ctrl + Home`|
|Go to a specific cell|`Ctrl + G` or `F5`|

#### **Editing and Formatting Shortcuts:**

|**Action**|**Shortcut**|
|---|---|
|Copy selected cells|`Ctrl + C`|
|Paste copied cells|`Ctrl + V`|
|Cut selected cells|`Ctrl + X`|
|Undo last action|`Ctrl + Z`|
|Redo last undone action|`Ctrl + Y`|
|Bold text|`Ctrl + B`|
|Italicize text|`Ctrl + I`|
|Underline text|`Ctrl + U`|
|Open Format Cells dialog|`Ctrl + 1`|
|Toggle showing formulas in the worksheet|`Ctrl + ~`|
|Insert the SUM function|`Alt + =`|
|Add or remove filter|`Ctrl + Shift + L`|

#### Data and Row Management Shortcuts:

|**Action**|**Shortcut**|
|---|---|
|Delete empty rows|`Ctrl + G` → "Special" → "Blanks" → `Ctrl + -`|
|Insert a new worksheet|`Shift + F11`|
|Open the Find and Replace dialog|`Ctrl + H`|
|Open the Go To Special dialog|`Ctrl + G` (then click "Special")|
|Insert current date|`Ctrl + ;`|
|Insert current time|`Ctrl + Shift + ;`|

---

### **Favorite Formulas:**

**XLOOKUP with Two Criteria:**
   - This formula looks up a value based on two criteria:

``` excel
  =XLOOKUP(1, (J2:J8=O12)*(P2:P8=P12), Q2:Q8, "not found")
  ```
**Remove Duplicates:**
  - Use the `UNIQUE` function to find unique values in a range:
   
``` excel
 =UNIQUE(SELECT ARRAY)
```

 **Find and Replace Non-breaking Spaces:**

- Use `CTRL + H` to find and replace non-breaking spaces. In the "Find" field, enter `Alt + 0160` to find non-breaking spaces and replace them with regular spaces.

 **Delete Empty Rows:**
- Select the data, press `CTRL + G`, choose "Special", select "Blanks", and then press `CTRL + -` to delete the empty rows.

**Conditional Formatting to Highlight Active Row:**

- Use this formula in Conditional Formatting to highlight the active row:

``` excel
=CELL("row")=ROW()
```

**Look for a Specific Word in a Cell:**
   - To check if a cell contains a specific word and return a value based on that:
   

``` excel
  =IF(ISNUMBER(SEARCH("disp", A2)), "display", "other")
  ```

 **Extract Text After a Specific Character (e.g., after @ in an email):**
   - Use the `RIGHT` and `SEARCH` functions to extract text after the "@" symbol:
   

``` excel
   =RIGHT(N2, LEN(N2) - SEARCH("@", N2))
   ```

 **Define a Dynamic Named Range:**
   - Use the `OFFSET` function to create a dynamic named range:
   
   ``` excel
   = OFFSET('Intro to Dynamic Array'!$B$3, 0, 0, 
   COUNTA('Intro to Dynamic Arrays'!$B:$B) - 1, 1)
  ```

**Separate Data into Rows:**
   - Use the `WRAPROW` function to split data into rows:
   
   ``` excel
  =WRAPROW(SELECT DATA, HOW MANY ROWS)
  ```
 
 **Remove Everything After a Space:**
   - Use `CTRL + H` to find and replace everything after a space:
   - **Find**: Enter `*` after a space.
   - **Replace**: Leave blank or add your desired replacement.
 **LOOKUP Function:**
   - Use the `LOOKUP` function to find a value within a range:
   
``` excel
  =LOOKUP(B2, $E$2:$F$2)
 ```


---
### **Additional Tips & Tricks:**

##### **Activate a Row Based on Conditions:**
- Use Conditional Formatting and a formula to highlight rows based on conditions.
- Right-click the sheet tab, choose "View Code," and use this VBA code to calculate the active cell:

vba
``` vba
   ActiveCell.Calculate
   ```
##### **`Cells.EntireColumn.AutoFit`**  
- The `AutoFit` method automatically resizes the width of the columns to fit the longest entry in each column. It adjusts the column width based on the content of the cells, ensuring that all data is fully visible without having to manually adjust each column.
vba

``` vba
Cells.EntireColumn.AutoFit
```

- `Cells`: Refers to all the cells in the worksheet.
- `EntireColumn`: Refers to all the columns in the worksheet.
- `AutoFit`: Automatically adjusts the width of the columns based on the content.


 **Formula to Move Text to the End of a Cell:**
   - As mentioned earlier, you can use `CTRL + 1` to open the Format Cells dialog and customize the format to move text to the end of the cell. For example, use `@* :` to add a colon at the end of the cell content.

---

 ^ Help you get the job done faster and more accurately.


---
Excel has many more advanced functions and tools that can further enhance productivity.
