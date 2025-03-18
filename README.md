

## 1. Overview

**A. Part A of the exercise:**
- Create a Google Script function that the user can run on this spreadsheet to update the "Averages" sheet with new data from "New 24 Hour Data" sheet.

**B. Part B of the exercise:**
- Create two line charts on the "Charts" sheet, one for the weekly average prices of each product and another for the 4-week average prices. The unit on the X-axis is days.
- Modify your Part A function so that each time it is run, it creates a new set of data points for each chart.


## 2. Function Descriptions


1. **`updateAverages()`**
   - **Purpose:** This is the main function that coordinates the entire process of updating the price data and charts.
   - **Steps:**
     - Retrieves the 24-hour price data that was on the 'New 24 Hour Data' Sheet.
     - Appends the data to the "Historical Data" sheet.
     - Calculates weekly and four-week price averages.
     - Updates the "Averages" sheet with the calculated averages.
     - Generates charts for the weekly and four-week average prices and appends them to respective sheets.

2. **`getOrCreateSheet(ss, sheetName)`**
   - **Purpose:** This is a helper function, itetrieves a sheet by name or creates it if it does not exist.
   - **Steps:**
     - Checks if the specified sheet exists in the spreadsheet.
     - If not, a new sheet with the specified name is created.


3. **`getNew24HourData(ss)`**
   - **Purpose:** Fetches the 24-hour price data for products from the "New 24 Hour Data" sheet.
   - **Steps:**
     - Retrieves the product names and prices from the sheet.
     - Combines them into an array of objects representing the product and its corresponding price.

4. **`appendHistoricalData(ss, combinedData)`**
   - **Purpose:** Appends new 24-hour price data to the "Historical Data" sheet.
   - **Steps:**
     - Retrieves the last row of data in the "Historical Data" sheet.
     - Checks if data for the current day already exists.
     - If not, appends the new data with the current date and time.

5. **`getLast7DaysPrices(ss)`**
   - **Purpose:** Retrieves price data for the last 7 days for each product from the "Historical Data" sheet.
   - **Steps:**
     - Filters the historical data for entries within the last 7 days.
     - Organizes the data by product and stores the prices in an object.
     - Logs a warning if there’s insufficient data for a product.

6. **`getLast4WeekPrices(ss)`**
   - **Purpose:** Retrieves price data for the last 4 weeks for each product from the "Historical Data" sheet.
   - **Steps:**
     - Filters the historical data for entries within the last 4 weeks (28 days).
     - Organizes the data by product and stores the prices in an object.
     - Logs a warning if there’s insufficient data for a product.

7. **`calculateAverage(lastData)`**
   - **Purpose:** Calculates the average price for each product based on the provided price data.
   - **Steps:**
     - Loops through each product's price data.
     - Computes the average of the prices and stores the result.


8. **`updateAveragesSheet(ss, weeklyAverages, fourWeekAverages)`**
   - **Purpose:** Updates the "Averages" sheet with the weekly and four-week average prices.
   - **Steps:**
     - Clears any previous average data in the sheet.
     - Adds the calculated averages for each product to the sheet.

3. **`createCharts(ss, weeklyAverages, fourWeeksAverage)`**
   - **Purpose:** Generates and updates charts based on the weekly and four-week price averages.
   - **Steps:**
     - Creates charts on the "Weekly Chart" and "FourWeeks Chart" sheets.
     - Sets chart options such as titles, axis labels, and series data.
     - Removes any existing charts before adding new ones to ensure only up-to-date charts are displayed.



## 3. Additional Notes

### **A. Data Sheets and Their Purpose**

1. **"Historical Data":**  
   This sheet keeps a record of all appended 24-hour prices for each product. It serves as the foundation for calculating weekly and four-week averages. Each time the `updateAverages()` function runs, new data from the "New 24 Hour Data" sheet is appended to this sheet.

2. **"Averages":**  
   This sheet is updated with the calculated averages for both the weekly and four-week periods. The data is presented alongside the product names, ensuring clear visibility of the price averages over the two periods.

3. **"Weekly Chart":**  
   This sheet stores the weekly average prices for each product. The chart generated here pulls data from the stored weekly averages, ensuring that the charts reflect up-to-date information.

4. **"Four Weeks Chart":**  
   Similar to the "Weekly Chart" sheet, this one stores the four-week average prices. It visualizes the trend of price averages over the past month, pulling data from the stored 4-week averages.

---

### **B. Missing Data Handling**

1. **Missing Data for Averages:**  
   If there is insufficient data for a product (e.g., fewer than 7 data points in the last 7 days or fewer than 28 data points in the last 4 weeks), the script **logs a warning** using `Logger.log()` but continues executing without interrupting the process. This ensures that incomplete data does not prevent the rest of the operations from being performed. The missing averages for such products are represented as empty cells on the "Averages" sheet.

2. **Strict Validation of Data:**  
   The script performs strict validation by requiring a minimum number of data points (7 for the last 7 days and 28 for the last 4 weeks) before proceeding with the average calculations. If any product does not meet this minimum data requirement, the product’s average is left blank, and a log is generated to notify the user of the insufficiency.
   - **For Weekly Average Calculation**: Only products with at least 7 valid data points in the past week are included in the weekly average calculation.
   - **For Four-Week Average Calculation**: Only products with at least 28 valid data points in the last four weeks are included in the four-week average calculation.

3. **Data Appending:**  
   New data is appended to the "Historical Data" sheet only if it doesn't already exist for the current day. This prevents duplication and ensures that the same date is not logged multiple times for the same product. If data for the current day already exists in the sheet, the script simply skips appending and proceeds with the rest of the execution.

---

### **C. Chart Creation and Updates**

1. **Weekly and Four-Week Charts:**  
   The script generates two separate line charts — one for the weekly average prices and one for the four-week average prices. Each time the script is run, it:
   - Updates the charts with the most recent data.
   - Clears any previous charts to ensure that only the current, up-to-date charts are displayed.

2. **Chart Customization:**  
   Each chart is configured with the following options:
   - **Title:** The title of the chart is set to reflect whether it represents weekly or four-week averages.
   - **X-Axis (Date):** The dates are shown on the X-axis, and they are formatted in the `yyyy/MM/dd` format.
   - **Y-Axis (Price):** The prices are displayed on the Y-axis with labels indicating the average price for that period.
   - **Legend:** A legend is included to display the products being tracked.

3. **Chart Sheet Creation:**  
   If the "Weekly Chart" or "Four Weeks Chart" sheets do not exist, they are created automatically. The script checks for the existence of the sheets and creates them if they are missing. This ensures that chart data is always available in the spreadsheet, even if the user is running the script for the first time or after deleting the charts.

---
### **E. Considerations for Future Enhancements**

1. **Graceful Handling of Missing Data:**  
   The current approach strictly checks for enough data before calculating averages. Future enhancements could consider more flexible handling, such as interpolating or using the available data to calculate approximate averages when full datasets are missing.

2. **Dynamic Chart Updates:**  
   The chart creation and update logic is currently fixed for weekly and four-week averages. Depending on the needs, it might be beneficial to allow the user to specify different time frames (e.g., monthly, quarterly) for which charts and averages are generated dynamically.

---





