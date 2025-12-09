A Dash/Plotly web application for analysing annual budgets, project spending, category-level cost distribution, and department performance.
This dashboard is designed to support R&D and corporate finance teams by presenting a clear, interactive view of budget vs. actual spending.


✨ Features
Budget vs Actual visualization across all cost categories

Top 10 spending projects (overall, R&D, labour, assets)

Clean and interactive UI using Dash + Plotly

Annual Budget Usage by Division, Team, Project



✨ How to Run the Dashboard:
1. Clone the repository
git clone https://github.com/fongyinho/Annual-Budget-and-Performance-Review-.git
cd Annual-Budget-and-Performance-Review-

2. Install dependencies
pip install -r requirements.txt

3. Run the Dash app
python app2.py

The dashboard will start on:
http://127.0.0.1:8090/


✨ Project Structure
Annual-Budget-and-Performance-Review-/
│
├── app2.py               # Main Dash application
├── EXCEL_BI_ALLDATA.xlsx # Sample depersonalized dataset
├── requirements.txt      # Python dependencies
└── README.md             # Project documentation
