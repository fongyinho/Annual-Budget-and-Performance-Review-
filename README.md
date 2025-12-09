## Annual Budget and Performance Review
A Dash/Plotly web application for analysing annual budgets, project spending, category-level cost distribution, and department performance.
This dashboard is designed to support R&D and corporate finance teams by presenting a clear, interactive view of budget vs. actual spending.


## Features

- Annual Budget Summary with overall usage %, total spent, and remaining balance

- Budget vs Actual by Division with usage rates and detailed comparison

- Actual Cost Distribution (Labour, R&D, Assets, Admin) visualized via pie chart

- Budget vs Actual by Category with usage percentages

- Top 10 Cost Drivers across Total Cost, Labour, R&D+Admin, and Asset categories

- Department-Level Dashboards, each showing:

   - Budget usage donut
   
   - Project-level Budget vs Actual comparison
   
   - Usage percentage for each project

- Clean UI built with Dash and Plotly



## How to Run the Dashboard:
1. Clone the repository
   
   git clone https://github.com/fongyinho/Annual-Budget-and-Performance-Review-.git

   cd Annual-Budget-and-Performance-Review-

3. Install dependencies
   
   pip install -r requirements.txt

5. Run the Dash app
   
   python app2.py

   The dashboard will start on:
   http://127.0.0.1:8090/


## Project Structure

Annual-Budget-and-Performance-Review-/

│

├── app2.py               # Main Dash application

├── EXCEL_BI_ALLDATA.xlsx # Sample depersonalized dataset


├── requirements.txt      # Python dependencies
└── README.md             # Project documentation
