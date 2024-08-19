**STOCK PRICES CHALLENGE USING VISUAL BASIC FOR APPLICATIONS [VBA]**

**(1) Project Overview and Purpose:**

For this exercise, stock market data will be used to an generate analysis of a variety of stocks via Visual Basic for Applications [VBA] scripting. The dataset spans a period between 2018 and 2020, which will allow for identification of specific metrics as described below.


**(2) Dataset Description:**

The dataset is stored in a Microsoft Excel workbook with three tabs, 2018; 2019; and 2020. It contains information on stock opening price; high price; low price; close price; and volume traded for each day of each year for numerous stocks (denoted by their ticker symbol). There are over 2.26 million unique rows within the the three tabs of this workbook.


**(3) Data Cleaning and Preprocessing:**

No cleaning or preprocessing was necessary for this MS Excel workbook. Rows and columns did not contain missing or incomplete data.


**(4) Data Visualization Techniques:**

VBA script loops through the stock data workbook, then reads and stores important stock information like the ticker symbol; the stock volume; the open price; and the close price.


Additional columns are added to hold values for the ticker symbol and newly generated metrics, like total stock volume; quarterly change; and percent change. Conditional formatting was added to two columns, where red indicates a negative change and green indicates a positive change.


Further new metrics were added to provide the greatest percent increase; greatest percent decrease; and greatest total volume, and the ticker symbol and associated value for each.



**(5) Results and Analysis:**

The results of this exercise is that change metrics and total stock volume were generated per ticker symbol for each year from 2018 to 2020. The positive and negative changes were correctly formatted as green and red, respectively, which allows for quicker perusal of the data. 


For each year, stocks with the greatest percent increase; greatest percent decrease; and greatest total volume are correctly identified and presented to the viewer, which could help with investment decisions.


The visualizations and metrics provide the means to select specific information on stock prices by day; year; and stock as well as summary information intended to promote further understand of stock activity including overarching performance metrics beneficial to traders and/or owners.


**(6) Ethical Considerations:**

While it would be difficult to compile the dataset in its current form (e.g., daily and over years from numerous stocks) without significant effort, there don't seem to be any privacy issues associated with this dataset, as stock prices are available via online sources or intermediaries (such as brokers) over the days, months, and years covered in this dataset.


**(7) Instructions for Interacting with the Project:**

Output from this exercise are stored in the main folder and contain the following-

(a) .vbs file for VBA code to run the stock analysis on a single worksheet, namely 2018, in the Multiple_year_stock_data file

(b) .vbs file for VBA code to run the stock analysis on all worksheets in the Multiple_year_stock_data file (2018, 2019, 2020)

(c) .png screenshot of the 2018 results

(d) .png screenshot of the 2019 results

(e) .png screenshot of the 2020 results

(f) .txt file for VBA code to run the stock analysis on a single worksheet, namely 2018, in the Multiple_year_stock_data file

(g) .txt file for VBA code to run the stock analysis on all worksheets in the Multiple_year_stock_data file (2018, 2019, 2020)


**(8) Citations:**

The following are sources of assistance I received to complete this exercise-

(a) Microsoft. (2021, September 12). Font.Bold property (Excel). Learn.Microsoft.com. https://learn.microsoft.com/en-	us/office/vba/api/excel.font.bold

(b) Peer Support via Zoom, 09-Apr-2024, outside class

(c) Peer Support within class, 10-Apr-2024
