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

This formula looks up a value from the range `Q2:Q8` based on **two conditions**:
1. The value in `J2:J8` must match the value in `O12`.
2. The value in `P2:P8` must match the value in `P12`.
If both conditions are met in the same row, it returns the corresponding value from `Q2:Q8`. If no match is found, it returns `"not found"`.
### **Explanation:**

1. **`(J2:J8=O12)*(P2:P8=P12)`**:
- This creates an array of `1`s and `0`s (True = 1, False = 0) by checking both conditions.
- The multiplication (`*`) ensures that only rows where **both conditions** are `TRUE` will return `1`.
2. **`XLOOKUP(1, ...)`**:
- `XLOOKUP` looks for the first `1` (where both conditions are `TRUE`) in the array created by the multiplication.
3. **`Q2:Q8`**:
- Once a match is found, it returns the corresponding value from the range `Q2:Q8`.
4. **`"not found"`**:
 - If no match is found, the formula returns `"not found"`.

### **Example:**
If `O12 = "C"` and `P12 = 3`, and your data is:

|**J**|**P**|**Q**|
|---|---|---|
|A|1|100|
|B|2|200|
|C|3|300|
|D|4|400|

The formula will return **`300`**, since both conditions are met in the third row.

---

**Remove Duplicates:**
  - Use the `UNIQUE` function to find unique values in a range:
``` excel
 =UNIQUE(SELECT ARRAY)
```

 **Find and Replace Non-breaking Spaces:**
- Use `CTRL + H` to find and replace non-breaking spaces. In the "Find" field, enter `Alt + 0160` to find non-breaking spaces and replace them with regular spaces.

 **Delete Empty Rows:**
- Select the data, press `CTRL + G`, choose "Special", select "Blanks", and then press `CTRL + -` to delete the empty rows.

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

1. **Open the VBA Editor**: Press `Alt + F11` in Excel.
2. **Insert a Module**: In the VBA editor, go to **Insert > Module**.
3. **Paste the Code**: Copy and paste the desired VBA code into the module.
4. **Run the Macro**: Press `F5` to run the macro, or go back to Excel and run it from the **Macros** menu (`Alt + F8`).
##### **Cells.EntireColumn.AutoFit**
- The `AutoFit` method automatically resizes the width of the columns to fit the longest entry in each column. It adjusts the column width based on the content of the cells, ensuring that all data is fully visible without having to manually adjust each column.
vba
``` vba
Cells.EntireColumn.AutoFit
```

- `Cells`: Refers to all the cells in the worksheet.
- `EntireColumn`: Refers to all the columns in the worksheet.
- `AutoFit`: Automatically adjusts the width of the columns based on the content.

##### **Loop Through All Worksheets**
If there is a need to perform the same operation on every worksheet in the workbook (e.g., formatting, data cleanup, etc.), the following loop can be used:

vba
```
Sub LoopThroughWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Perform your actions here
        ws.Cells.EntireColumn.AutoFit ' Example: Auto-fit columns on each sheet
    Next ws
End Sub
```

##### **Delete Blank Rows**
This macro will delete all blank rows in the active worksheet.

vba
```
Sub DeleteBlankRows()
    Dim rng As Range
    On Error Resume Next
    Set rng = ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If Not rng Is Nothing Then rng.EntireRow.Delete
End Sub
```

##### **Highlight Duplicates in a Range**
This code will highlight duplicate values in a selected range.

vba
```
Sub HighlightDuplicates()
    Dim rng As Range
    Dim cell As Range
    Set rng = Selection ' Set the range to the current selection
    
    For Each cell In rng
        If WorksheetFunction.CountIf(rng, cell.Value) > 1 Then
            cell.Interior.Color = RGB(255, 0, 0) ' Highlight in red
        End If
    Next cell
End Sub
```
 **Formula to Move Text to the End of a Cell:**
   - As mentioned earlier, you can use `CTRL + 1` to open the Format Cells dialog and customize the format to move text to the end of the cell. For example, use `@* :` to add a colon at the end of the cell content.

---

 > Help yourself get the job done faster and more accurately.


---
Remember excel has many more advanced functions and tools that can further enhance productivity.
