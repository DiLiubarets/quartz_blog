These formulas are used to populate various fields such as "Ship to ID," "Ship to Name," "Invoice Date," "COGS," etc. Let's break down and clean up the formulas for clarity and ensure that they are ready for use in your report.

#### Below is organized version of the formulas I used:

|**Field**|**Formula**|
|---|---|
|**OF**|`=IF(R2<>"", "OF", "")`|
|**91 (Ship to ID)**|`=IF(R2<>"", "91", "")`|
|**Ship to Name**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!K:K, "No DATA")`|
|**Ship Address**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!L:L, "No DATA")`|
|**Ship City**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!AC:AC, "No DATA")`|
|**Ship Province**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!AF:AF, "No DATA")`|
|**Ship Postal Code**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!AE:AE, "No DATA")`|
|**Ship Country**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!AG:AG, "No DATA")`|
|**Country Code (CA)**|`=IFERROR(IF(R2<>"", "CA", ""), "NO DATA")`|
|**Item ID**|`='Data '!E2`|
|**Product Code**|`=IFERROR(RIGHT('Data '!E2, LEN('Data '!E2)-6), "NO DATA")`|
|**Product Description**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!F:F, "No DATA")`|
|**Invoice Date**|`=IFERROR(TEXT(DATE(RIGHT('Data '!J2, 4), LEFT('Data '!J2, FIND("/", 'Data '!J2) - 1), MID('Data '!J2, FIND("/", 'Data '!J2) + 1, FIND("/", 'Data '!J2, FIND("/", 'Data '!J2) + 1) - FIND("/", 'Data '!J2) - 1)), "YYYYMMDD"), "NO DATA")`|
|**Quantity Shipped**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!U:U, "No DATA")`|
|**Unit of Measure**|`=XLOOKUP(Q2, 'Data '!E:E, 'Data '!I:I, "No DATA")`|
|**COGS (Cost of Goods Sold)**|`=IFERROR(ROUNDUP(AB2/Y2, 2), "NO DATA")`|
|**COGS Formula**|`=IFERROR(ROUNDUP(XLOOKUP(@Q:Q, 'Data '!E:E, 'Data '!AQ:AQ), 2), "NO DATA")`|

---

### Explanation of Key Formulas:

1. **OF (Order Flag)**:
        - `=IF(R2<>"", "OF", "")`  
        This formula checks if the value in `R2` is not empty. If it's not empty, it returns "OF" (likely indicating an order flag); otherwise, it returns an empty string.
2. **Ship to ID (91)**:
        - `=IF(R2<>"", "91", "")`  
        Similar to the previous formula, this checks if `R2` is not empty and returns "91" (which seems to be the vendor or store ID).        
3. **Ship to Name, Address, City, Province, Postal Code, Country**:
        - These formulas use `XLOOKUP` to pull data from the "Data" sheet based on the product ID in column `Q`. For example:  
        `=XLOOKUP(Q2, 'Data '!E:E, 'Data '!K:K, "No DATA")`  
        This searches for the value in `Q2` in column `E` of the "Data" sheet and returns the corresponding value from column `K`. If no match is found, it returns "No DATA."
4. **Country Code**:
        - `=IFERROR(IF(R2<>"", "CA", ""), "NO DATA")`  
        This formula checks if `R2` is not empty. If true, it returns "CA" (likely for Canada); otherwise, it returns an empty string. If an error occurs, it returns "NO DATA."        
5. **Product Code**:
        - `=IFERROR(RIGHT('Data '!E2, LEN('Data '!E2)-6), "NO DATA")`  
        This formula extracts the product code by removing the first 6 characters from the value in `'Data '!E2`. If an error occurs, it returns "NO DATA."
6. **Invoice Date**:
        - `=IFERROR(TEXT(DATE(RIGHT('Data '!J2, 4), LEFT('Data '!J2, FIND("/", 'Data '!J2) - 1), MID('Data '!J2, FIND("/", 'Data '!J2) + 1, FIND("/", 'Data '!J2, FIND("/", 'Data '!J2) + 1) - FIND("/", 'Data '!J2) - 1)), "YYYYMMDD"), "NO DATA")`  
        This formula converts a date stored as text in a non-standard format into the `YYYYMMDD` format. If it fails, it returns "NO DATA."
7. **Quantity Shipped**:
        - `=XLOOKUP(Q2, 'Data '!E:E, 'Data '!U:U, "No DATA")`  
        This pulls the quantity shipped from column `U` of the "Data" sheet based on the product ID in `Q2`.
8. **COGS (Cost of Goods Sold)**:
        - `=IFERROR(ROUNDUP(AB2/Y2, 2), "NO DATA")`  
        This formula calculates the COGS by dividing the value in `AB2` by the value in `Y2` and rounding it up to 2 decimal places. If an error occurs, it returns "NO DATA."
9. **COGS Formula**:
        - `=IFERROR(ROUNDUP(XLOOKUP(@Q:Q, 'Data '!E:E, 'Data '!AQ:AQ), 2), "NO DATA")` This formula uses `XLOOKUP` to retrieve the COGS value from column `AQ` in the "Data" sheet, based on the product ID in column `Q`. It rounds the result to 2 decimal places. If an error occurs, it returns "NO DATA."

---

### Next Steps:

1. **Validation**: Ensure that data in the "Data" sheet is properly structured and that the columns referenced in the formulas (e.g., `E:E`, `K:K`, `L:L`, etc.) contain the correct information.
2. **Testing**: Once the formulas are applied, validate the output by checking a few rows of data to ensure that everything is pulling correctly from the "Data" sheet.

### Formulas for Condition formatting **Data sheet:**

The vendor requires the product codes in column AC to start with a number. This conditional formatting formula checks if the first character in AC1 is **not** a number and highlights the cell if the condition is true.

```
=NOT(ISNUMBER(VALUE(LEFT(AC1,1))))
```

- **Explanation**:
    - `LEFT(AC1,1)` extracts the first character of the value in AC1.
    - `VALUE(LEFT(AC1,1))` converts the first character into a number (if possible).
    - `ISNUMBER(VALUE(...))` checks if the first character is a number.
    - `NOT(...)` ensures that the cell is highlighted **only if the first character is not a number**.

##### Ensuring Correct Formatting in Column AG (7 Characters with a Space in the 4th Position)

The vendor requires the values in column AG to be **exactly 7 characters long**, with a **space in the 4th position**. This formula checks if the value in AG1 fails to meet these criteria and highlights the cell if the format is incorrect.

```
=NOT(AND(LEN(AG1)=7, MID(AG1,4,1)=" "))
```

- **Explanation**:
    - `LEN(AG1)=7` checks if the length of the value in AG1 is exactly 7 characters.
    - `MID(AG1,4,1)=" "` checks if the 4th character in AG1 is a space.
    - `AND(...)` ensures both conditions are met.
    - `NOT(...)` highlights the cell if either condition is **not** met (i.e., the length is not 7 characters or the 4th character is not a space).

##### Highlighting Missing Data in Columns AE and AF

The vendor's template requires that columns AE and AF should not be left blank. These formulas check if the cells in AE1 or AF1 are either empty or blank and highlight the cells if data is missing.

```
=OR(AE1="", ISBLANK(AE1))
```

##### Formula for Column AF:

```
=OR(AF1="", ISBLANK(AF1))
```

- **Explanation**:
    - `AE1=""` or `AF1=""` checks if the cell is empty.
    - `ISBLANK(AE1)` or `ISBLANK(AF1)` checks if the cell is blank (sometimes a cell may appear empty but isn't technically blank).
    - `OR(...)` highlights the cell if **either** condition is true (i.e., the cell is empty or blank).

---

##### **How to Apply These Conditional Formatting Rules in Excel:**

1. **Select the range** of cells where you want to apply the conditional formatting (e.g., select the entire column AC, AG, AE, or AF).
2. Go to the **Home** tab in Excel.
3. Click on **Conditional Formatting** > **New Rule**.
4. Choose **Use a formula to determine which cells to format**.
5. Enter the appropriate formula from above (depending on the column you're applying it to).
6. Click on **Format...** to choose the formatting style (e.g., background color) for the highlighted cells.
7. Click **OK** to apply the rule.

---

##### Summary 

- **Column AC**: Highlights cells where the first character is not a number.
- **Column AG**: Highlights cells where the value is not exactly 7 characters long or does not have a space in the 4th position.
- **Columns AE and AF**: Highlights cells that are empty or blank.

These conditional formatting rules helping me ensure that report is formatted correctly and that any inconsistencies or missing data are quickly identified for correction before submission.
