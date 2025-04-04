# 📊 VBA Challenge – Stock Market Data Analysis

## 🗂️ Overview

This project utilizes **VBA (Visual Basic for Applications)** within **Microsoft Excel** to analyze stock market data efficiently. The VBA scripts automate the processing of large datasets, enabling quick computation of key financial metrics across multiple stock tickers.

---

## 🛠️ Tools & Technologies

- **Microsoft Excel**
- **VBA (Visual Basic for Applications)**

---

## 📁 Project Structure

```
VBA-challenge/
├── Resources/
│   └── stock_data.xlsx      # Excel workbook containing stock market data
├── Scripts/
│   └── stock_analysis.bas   # VBA script module for analysis
└── README.md                # Project documentation
```

---

## 🔑 Key Features

- **Ticker-wise Analysis:** Calculates yearly change, percentage change, and total stock volume for each ticker symbol.
- **Conditional Formatting:** Highlights positive and negative yearly changes with distinct colors for quick visualization.
- **Performance Metrics:** Identifies and displays the stocks with the greatest percentage increase, greatest percentage decrease, and highest total volume.

---

## 🖥️ How to Use

1. **Load Data:**
   - Open `stock_data.xlsx` in Microsoft Excel.
   - Ensure the dataset is structured with columns for ticker symbols, dates, opening prices, closing prices, and volumes.

2. **Import VBA Script:**
   - Press `ALT + F11` to open the VBA editor.
   - In the editor, go to `File` > `Import File...` and select `stock_analysis.bas` from the `Scripts` folder.
   - The script will be added to your VBA project.

3. **Run the Script:**
   - Close the VBA editor to return to Excel.
   - Press `ALT + F8`, select `StockAnalysis`, and click `Run`.
   - The script will process the data and output the analysis directly into the workbook.

---

## 📈 Expected Output

After running the script, the Excel workbook will display:

- **Yearly Change:** Difference between the opening price on the first trading day and the closing price on the last trading day for each ticker.
- **Percentage Change:** Yearly change expressed as a percentage.
- **Total Stock Volume:** Sum of the trading volumes for the year for each ticker.
- **Highlighting:** Positive changes in green and negative changes in red for easy interpretation.
- **Summary Table:** A consolidated table showing the stocks with the greatest increases, decreases, and highest volumes.

---

## 🚀 Future Enhancements

- **User Input:** Allow users to specify the year or range of dates for analysis.
- **Dynamic Data Handling:** Adapt the script to handle datasets of varying structures or additional financial metrics.
- **Visualization:** Integrate chart generation for a graphical representation of stock performance.

---

## 📝 Notes

- Ensure macros are enabled in Excel to run the VBA script.
- Always save your work before running scripts to prevent unintended data loss.
- The script assumes a specific data structure; modifications may be necessary for differently structured datasets.

---

**Author:** [Geo222222](https://github.com/Geo222222)  
**Focus:** Financial Data Analysis • VBA Automation • Excel Scripting

