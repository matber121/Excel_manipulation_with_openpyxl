Excel Manipulation and Analysis with Python (openpyxl)
This project demonstrates how to automate and analyze Excel data using the openpyxl library. It works with a dataset of video game sales and performs a variety of tasks, including data extraction, modification, formula generation, and chart creation.

🔧 Requirements
Python 3.6+

openpyxl library

Install the required package with:

bash
Copy
Edit
pip install openpyxl
📁 Dataset Used
The script works with an Excel file named videogamesales.xlsx, which contains sales data of various video games across different regions and genres.

Sheet vgsales: Contains the raw data.

Sheet Total Sales by Genre: Contains aggregated data by genre.

🚀 Features
✅ Basic Operations
Load Workbook & Worksheet

Read Data:

Value of specific cells.

Values from a range of cells and rows.

Write Data: Modify existing cells or insert new data.

🧮 Data Processing
Add a New Column (sum of sales):

Computes the total sales by summing regional sales (NA, EU, JP, Other).

Append New Rows: Adds new video game sales entries.

🧹 Cleanup Tasks
Delete Last Row: Demonstrates how to remove a specific row from the worksheet.

📐 Excel Formulas
The script generates the following formulas dynamically:

Column	Description
P1:P2	Calculates average sales across all rows
Q1:Q2	Counts populated cells in a column
R1:R2	Counts rows where the genre is 'Sports'
S1:S2	Total sales for 'Sports' genre
T1:T2	Rounds up total sports sales to the nearest 25

📈 Chart Creation
Generates Bar Charts to visually represent total sales by genre.

Adds labels, titles, and custom layout using ManualLayout.

