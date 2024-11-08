In my role as a **Data Analyst/Purchase Manager** at a tool store, I am responsible for generating a biweekly report that must be uploaded to a vendor’s website. The report follows a specific template provided by the vendor and must be validated before submission to ensure that all data is correct and properly formatted.

Initially, this process was manual and time-consuming, but by leveraging Excel’s powerful formulas and Conditional Formatting, I was able to automate the report generation and validation process. In this article, I’ll explain how I streamlined this task.

---
#### **Situation:**

The vendor requires a biweekly report from our tool store, which must be uploaded to their website in a specific format. This report includes key data such as product orders, stock levels, and other relevant metrics. The vendor’s template is strict, and any deviation from the required format or missing data could result in the report being rejected.

As the **Data Analyst/Purchase Manager**, I was responsible for ensuring that the report was accurate, complete, and ready for validation before submission. Initially, the process involved manually entering data, checking for errors, and formatting the report to meet the vendor’s template. This was time-consuming and prone to human error, especially when handling large datasets.

---
#### **My goal**  was to automate the report generation process to:

1. **Ensure data accuracy** by automatically pulling data from our internal systems and applying the necessary transformations.
2. **Validate the data** to meet the vendor’s strict formatting requirements.
3. **Highlight any errors or missing data** using Conditional Formatting, so they could be addressed before submitting the report.
4. **Prepare the report in the vendor’s template** for easy upload to their website.

The report needed to be generated every two weeks, and I wanted to ensure that the process was as efficient and error-free as possible.

---
#### Here’s how I approached automating the biweekly vendor report using Excel:

###### 1. **Data Extraction and Transformation:**

- I used **XLOOKUP** to pull relevant data from our internal systems (stored in the "Data" sheet) into the report template. This allowed me to automatically retrieve product information, stock levels, and order statuses based on product IDs.
- For example, to retrieve product details from the "Data" sheet, I used the following formula:
    
    ```excel
=XLOOKUP(
    Q2, 
    'Data '!E:E, 
    'Data '!K:K, 
    "No DATA"
)
```
    
    This formula looks up the value in `Q2` (the product ID) in column `E` of the "Data" sheet and returns the corresponding data from column `K`. If no match is found, it returns "No DATA" to indicate missing information.

###### 2. **Conditional Formatting for Data Validation:**

- To ensure that the report met the vendor’s strict formatting requirements, I used **Conditional Formatting** with custom formulas to highlight any data inconsistencies or missing values.
    
    - **Highlighting Invalid Data in Column AC**:  
        The vendor required certain fields, like product codes, to start with a number. To ensure this, I applied the following Conditional Formatting formula: 
        
        ``` excel
        =NOT(
	        ISNUMBER(
		        VALUE(
			        LEFT(AC1,1)
			        )
			    )
			)
        ```
        
        This formula checks if the first character of the value in `AC1` is not a number. If the condition is true, the cell is highlighted, indicating that the data needs to be corrected before submission.
        
    - **Ensuring Correct Formatting in Column AG**:  
        The vendor required values in column `AG` to be exactly 7 characters long, with a space in the 4th position. To check this, I used: 
        ```excel
        =NOT(
	        AND(
		        LEN(AG1)=7, 
		        MID(AG1,4,1)=" ")
		    )
        ```
        
        This formula ensures that the length of the value is exactly 7 characters and that the 4th character is a space. If the condition is not met, the cell is highlighted, signaling that the data needs to be reformatted.
        
    - **Highlighting Missing Data in Columns AE and AF**:  
        The vendor’s template required certain fields to be filled out. To ensure that no data was missing in columns `AE` and `AF`, I used the following formulas:
        ```excel
        =OR(
		    AE1 = "", 
		    ISBLANK(AE1)
			)
        =OR(
	        AF1="", 
	        ISBLANK(AF1)
	        )
        ```
        
        These formulas check if the cells in `AE` or `AF` are either empty or blank. If so, the cells are highlighted, indicating that the missing data needs to be filled in.
        
###### 3. **Text Manipulation and Date Formatting:**

- Some fields required specific formatting, such as dates and text fields. I used the **TEXT** and **DATE** functions to ensure that dates were formatted correctly according to the vendor’s template. For example:    
    ``` excel
    =IFERROR(
    TEXT(
        DATE(
            RIGHT('Data '!J188, 4), 
            LEFT('Data '!J188, FIND("/", 'Data '!J188) - 1), 
            MID(
                'Data '!J188, 
                FIND("/", 'Data '!J188) + 1, 
                FIND("/", 'Data '!J188, FIND("/", 'Data '!J188) + 1) - FIND("/", 'Data '!J188) - 1
            )
        ), 
        "YYYYMMDD"
    ),
    "NO DATA"
)
    ```
    
    This formula converts a date stored as text in a non-standard format into the vendor’s required `YYYYMMDD` format. If the operation fails, the formula returns "NO DATA."

###### 4. **Final Report Preparation:**

- After applying data validation and formatting, I ensured that the report was structured according to the vendor’s template. This included ensuring that all columns were in the correct order, and that any missing or incorrect data was highlighted for review before submission.

---
##### By automating the biweekly report using Excel, I was able to achieve several key benefits:

1. **Time Savings:** The automation significantly reduced the time spent on generating the report. What used to take hours of manual data entry and formatting now updates automatically whenever new data is added to the "Data" sheet.
2. **Improved Data Accuracy:** The use of **IFERROR** and **XLOOKUP** ensured that missing or incorrect data was easily identifiable, reducing the risk of errors in the final report.
3. **Validation Compliance:** By using **Conditional Formatting**, I was able to ensure that the report met the vendor’s strict formatting requirements. Any data that didn’t conform to the template was automatically highlighted, allowing for quick corrections before submission.
4. **Error Reduction:** Automating the report reduced the risk of human error, ensuring that the data was accurate and consistent every time.
5. **Efficient Upload:** Since the report was already formatted according to the vendor’s template, it was easy to upload to the vendor’s website without additional adjustments.

---
### **Conclusion:**
Automating the biweekly vendor report using Excel has significantly improved both efficiency and accuracy. By ensuring that the report meets the vendor’s strict template and validation requirements, I’ve been able to streamline the process and reduce the risk of errors. This automation not only saves time but also ensures that the report is ready for submission with minimal manual intervention.

---

If you’re responsible for generating regular reports that need to meet strict formatting or validation requirements, consider using Excel’s advanced formulas and Conditional Formatting to automate the process. It can save you time and ensure that your reports are always accurate and ready for submission.

[[Formulas are used]]

