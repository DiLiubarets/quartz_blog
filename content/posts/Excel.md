### **Excel Automation Playground**



In my role, I regularly prepare a detailed report that tracks inventory and sales data. This report includes critical metrics like product costs, shipping details, and cost of goods sold (COGS) across multiple customers and products. Initially, this was a manual process that required extensive data extraction, filtering, calculations, and formatting. Each report update was time-consuming and left room for human error. Recognizing these challenges, I aimed to develop an automated solution to streamline the process.

  

Task

  

The goal was to create an automated report that could:

1. Pull data dynamically from a raw Data sheet.

2. Calculate essential metrics like COGS automatically.

3. Filter data by date and other criteria without manual input.

4. Highlight important information, such as discrepancies in quantities.

5. Provide a high-level view through summaries or aggregations.

  

By achieving these, I could significantly reduce preparation time, improve accuracy, and make the report more insightful for stakeholders.

  

Action

  

To automate the report, I focused on optimizing each step in the report sheet, where all calculations and transformations take place:

1. Data Extraction and Lookup:

• I used INDEX/MATCH and XLOOKUP/VLOOKUP functions to pull key data points, like product codes and quantities, from the Data sheet. This setup allows data in the report sheet to update dynamically based on any changes in the Data sheet.

2. COGS Calculation:

• I implemented automated COGS calculations by creating formulas that reference shipment quantities and unit costs, calculating COGS per quantity shipped. By embedding these formulas directly into the report sheet, each COGS value adjusts whenever the underlying data changes.

3. Automated Filtering by Date:

• To filter data by date without manual effort, I used functions like IF, AND, and DATEVALUE. These functions ensured that only relevant records—based on invoice or shipment dates—appeared in the report, focusing on current data without clutter.

4. Conditional Formatting for Key Metrics:

• I applied conditional formatting to automatically highlight values of interest, such as quantities above a threshold or discrepancies between ordered and shipped amounts. This step provided an instant visual check, making any issues immediately visible.

5. Aggregation for Summary Insights:

• To create a high-level overview, I set up pivot tables to aggregate data by product and location. This allowed me to provide quick summaries on inventory levels, costs, and other key metrics, reducing the need for manual data aggregation.

  

Result

  

The automation transformed the reporting process with immediate benefits:

• Time Savings: I reduced preparation time drastically, freeing up time for more analytical tasks.

• Improved Accuracy: Automated lookups and formulas ensured consistent and error-free calculations.

• Real-Time Data Updates: The dynamic setup ensured that the report was always current and reflected the latest raw data.

• Enhanced Insights: With conditional formatting and aggregated summaries, the report now offers deeper insights with minimal effort.

  

Conclusion

  

Automating this report enabled me to deliver a more efficient, reliable, and insightful reporting tool. This approach to automation demonstrates how powerful Excel can be when leveraged effectively, saving time and increasing accuracy for data-heavy reporting tasks.