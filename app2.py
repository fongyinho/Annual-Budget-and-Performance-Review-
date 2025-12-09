import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

file_path = "EXCEL_BI_ALLDATA.xlsx"
df_master = pd.read_excel(file_path, sheet_name="Master")
df_budget = pd.read_excel(file_path, sheet_name="项目预算数据（测试版本）")
df_actual = pd.read_excel(file_path, sheet_name="项目实际数据（测试版本）")

for df in [df_master, df_budget, df_actual]:
    df.columns = df.columns.str.strip().str.replace("\n", "").str.replace(" ", "")

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY],
    suppress_callback_exceptions=True  
)
app.title = "Annual Budget and Performance Review"

year_dropdown = dcc.Dropdown(
    id="year-select",
    options=[{"label": y, "value": y} for y in [2025, 2026, 2027, 2028]],
    value=2025,
    clearable=False,
    style={"width": "200px", "fontFamily": "Microsoft YaHei"},
)

def overview_layout():
    return dbc.Container(
        [
            dbc.Row(
                [
                    dbc.Col(html.Div(id="chart1", style=box_style), width=6),
                    dbc.Col(html.Div(id="chart2", style=box_style), width=6),
                ],
                className="mb-3 print-page",
            ),
            dbc.Row(
                [
                    dbc.Col(html.Div(id="chart3", style=box_style), width=6),
                    dbc.Col(html.Div(id="chart4", style=box_style), width=6),
                ],
                className="mb-3 print-page",
            ),
            dbc.Row(
                [
                    dbc.Col(html.Div(id="chart5", style=box_style2), width=6),
                    dbc.Col(html.Div(id="chart6", style=box_style2), width=6),
                ],
                className="mb-3 print-page",
            ),
        ],
        fluid=True,
    )


r1_depts = [
    "Team7",
    "Team3",
]

r2_depts = [
    "Team9",
    "Team10",
    "Team19"
]

lab2030_depts = [
    "Team1",
    "Team8",
    "Team17",
    "Team12"
]

office_depts = [
    "Team16",
    "Team11"
]

CARD_BG = "#FBFBFB"
PLOT_BG = "#FBFBFB"
box_style = {
    "padding": "10px",
    "border": "1px solid #E0E0E0",
    "borderRadius": "6px",
    "fontFamily": "Microsoft YaHei",
    "backgroundColor": CARD_BG,
    "height": "450px",
    "overflow": "hidden",
}
box_style2 = {
    "padding": "10px",
    "border": "1px solid #E0E0E0",
    "borderRadius": "6px",
    "fontFamily": "Microsoft YaHei",
    "backgroundColor": CARD_BG,
    "height": "800px",
    "overflow": "hidden",
}
box_style3 = {
    "padding": "10px",
    "border": "1px solid #E0E0E0",
    "borderRadius": "6px",
    "fontFamily": "Microsoft YaHei",
    "backgroundColor": CARD_BG,
    "height": "500px",
    "overflow": "hidden",
}
style_not_select = {
    "fontFamily": "Microsoft YaHei",
    "padding": "6px 12px",
    "fontSize": "13px",
    "height": "35px",  
}
style_select = {
    "fontFamily": "Microsoft YaHei",
    "padding": "6px 12px",
    "fontSize": "13px",
    "height": "40px",   
    "fontWeight": "bold",
    "borderBottom": "0px",
}
def apply_fig_theme(fig):

    fig.update_layout(
        paper_bgcolor=CARD_BG,   
        plot_bgcolor=PLOT_BG,   
    )
    return fig
def apply_axis_style1(fig):
    fig.update_yaxes(
        showgrid=True,         
        showline=True,         
        gridcolor="rgba(0,0,0,0.05)",  
        gridwidth=0.5,
    )
    fig.update_xaxes(
        showgrid=False,
        showline=True, 
        gridcolor="rgba(0,0,0,0.05)", 
        gridwidth=0.5,
    )
    return fig
def apply_axis_style2(fig):
    fig.update_yaxes(
        showgrid=False,         
        showline=True,          
        gridcolor="rgba(0,0,0,0.05)", 
        gridwidth=0.5,
    )
    fig.update_xaxes(
        showgrid=True,
        showline=True, 
        gridcolor="rgba(0,0,0,0.05)",
        gridwidth=0.5,
    )
    return fig
def generate_multi_dept_layout(tab_id, dept_list, rows):
    layout_rows = []
    graph_ids = []
    total = len(dept_list)
    per_row = int((total + rows - 1) / rows)
    idx = 0
    for r in range(rows):
        row_cols = []
        for c in range(per_row):
            if idx < total:
                graph_id = f"{tab_id}_{idx}"
                graph_ids.append(graph_id)
                row_cols.append(
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id=f"{graph_id}_donut",
                                          config={"displayModeBar": False},
                                          style={"height": "260px"}),
                                html.Div(id=f"{graph_id}_overlay",
                                         className="fade-in")
                            ],
                            style=box_style3
                        ),
                        width=int(12 / per_row)
                    )
                )
            idx += 1
        layout_rows.append(dbc.Row(row_cols, className="mb-3"))
    return layout_rows, graph_ids

def r1_layout():
    return dbc.Container(
        [
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="r1_deptA_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="r1_deptA_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="r1_deptA_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="r1_deptB_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="r1_deptB_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="r1_deptB_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
        ],
        fluid=True,
    )
def r2_layout():
    return dbc.Container(
        [
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="r2_dept1_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="r2_dept1_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="r2_dept1_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="r2_dept2_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="r2_dept2_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="r2_dept2_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="r2_dept3_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="r2_dept3_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="r2_dept3_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
        ],
        fluid=True,
    )
def lab2030_layout():
    return dbc.Container(
        [
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="lab2030_dept1_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="lab2030_dept1_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="lab2030_dept1_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="lab2030_dept2_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="lab2030_dept2_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="lab2030_dept2_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="lab2030_dept3_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="lab2030_dept3_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="lab2030_dept3_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="lab2030_dept4_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="lab2030_dept4_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="lab2030_dept4_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
        ],
        fluid=True,
    )
def office_layout():
    return dbc.Container(
        [
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="office_dept1_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="office_dept1_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="office_dept1_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
            dbc.Row(
                [
                    dbc.Col(
                        html.Div(
                            [
                                dcc.Graph(id="office_dept2_donut", config={"displayModeBar": False}, style={"height": "390px"}),
                                html.Div(id="office_dept2_donut_table")
                            ],
                            style=box_style3
                        ),
                        width=3
                    ),
                    dbc.Col(
                        html.Div(
                            dcc.Graph(id="office_dept2_overlay", config={"displayModeBar": False}),
                            style=box_style3
                        ),
                        width=9
                    ),
                ],
                className="mb-3 print-page"
            ),
        ],
        fluid=True,
    )

dept_order = ["Dep1", "Dep2", "Dep3", "Dep4"]
my_colors = ["#A7C4BA", "#9B9AC1", "#FAC6CA", "#B198BB", "#DFC9A9"]
donut_budget_color = my_colors[0]     # dark blue
donut_actual_color = my_colors[1]     # light blue

app.layout = dbc.Container(
    [
        html.H2(id="page-title", style={"fontFamily": "Microsoft YaHei", "marginTop": "10px", "fontSize": "24px","fontWeight": "bold"}),
        html.Div(
            [
                html.Label("Select Year：", style={"fontFamily": "Microsoft YaHei", "marginRight": "10px"}),
                year_dropdown,
                html.Div(id="print-dummy", style={"display": "none"}),
            ],
            style={"display": "flex", "alignItems": "center", "gap": "10px"},
        ),
        html.Hr(),
        dcc.Tabs(
            id="tabs",
            value="tab-overview",
            children=[
                dcc.Tab(label="Summary", value="tab-overview", style=style_not_select, selected_style=style_select),
                dcc.Tab(label="Department 1", value="tab-r1", style=style_not_select, selected_style=style_select),                
                dcc.Tab(label="Department 2", value="tab-r2", style=style_not_select, selected_style=style_select),
                dcc.Tab(label="Department 3", value="tab-lab2030", style=style_not_select, selected_style=style_select),
                dcc.Tab(label="Department 4", value="tab-office", style=style_not_select, selected_style=style_select),
            ],
            className="no-print",
        ),
        html.Div(id="tab-content", style={"marginTop": "10px"}),
    ],
    fluid=True,
)

@app.callback(
    Output("tab-content", "children"),
    Input("tabs", "value")
)
def render_tab(tab):
    if tab == "tab-overview":
        return overview_layout()
    elif tab == "tab-r1":
        return r1_layout()
    elif tab == "tab-r2":
        return r2_layout()
    elif tab == "tab-lab2030":
        return lab2030_layout()
    elif tab == "tab-office":
        return office_layout()
    else:
        return html.Div("该页面内容稍后制作。", style={"fontFamily": "Microsoft YaHei"})

@app.callback(
    Output("page-title", "children"),
    Input("year-select", "value"),
)
def update_title(year):
    main_text = f"{year} Annual Budget and Performance Review"
    return [
        html.Span(
            main_text, 
            style={"marginRight":"8px"},
        ),
        html.Span(
            "(Show only ongoing projects)",
            style={
                "fontSize":"14px",
                "fontWeight":"normal",
                "color":"#555555",
                "verticalAlign":"baseline",
            },
        ),
    ]
@app.callback(
    [Output(f"chart{i}", "children") for i in range(1, 7)],
    Input("year-select", "value")
)
def update_overview(year):

    df_bud = df_budget[
        (df_budget["申请年份"] == year) &
        (df_budget["申请类型"] == "项目") &
        (df_budget["执行状态"] == "Ongoing")
    ]
    total_budget = df_bud["金额(元)"].sum() / 1000  

    df_act = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing")
    ]
    total_actual = df_act["SIPM125.BXJE"].sum() / 1000 

    usage_rate = 0 if total_budget == 0 else total_actual / total_budget

    fig1 = px.pie(
        names=["Used", "Balance"],
        values=[total_actual, max(total_budget - total_actual, 0)],
        hole=0.5,
        color=["Used", "Balance"],
        color_discrete_map={"Used": "#505D8B", "Balance": "#B9BED1"}
    )
    fig1.update_traces(
        textinfo="none",
        hovertemplate="%{label}: %{value:.2f} kCNY"
    )

    fig1.add_annotation(
        text=f"{usage_rate:.0%}",
        x=0.5, y=0.5,
        font=dict(size=28, family="Microsoft YaHei", color="#505D8B"),
        showarrow=False
    )
    fig1.update_layout(
        title=dict(
            text="Annual Budget Usage",
            font=dict(size=16, family="Microsoft YaHei")
        ),
        margin=dict(l=20, r=20, t=40, b=20),
        legend_title_text="",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.10,
            xanchor="center",
            x=0.5,
            font=dict(family="Microsoft YaHei", size=12)
        ),
    )
    apply_fig_theme(fig1)

    chart1 = dcc.Graph(figure=fig1, style={"height": "320px"}, config={"displayModeBar": False}, className="fade-in")

    header_row = html.Tr(
        [html.Th(f"Total Budget = {total_budget:,.0f} kCNY")] 
    )
    table_html = html.Table(
        [
            html.Thead(header_row),
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "18px",
            "textAlign": "center",
            "borderCollapse": "collapse"
        }
    )
    chart1 = html.Div([chart1, table_html])

    df_bud_dept = df_budget[
        (df_budget["申请年份"] == year) &
        (df_budget["申请类型"] == "项目") &
        (df_budget["执行状态"] == "Ongoing")
    ]
    bud_group = (
        df_bud_dept.groupby("一级部门")["金额(元)"].sum().reset_index())
    bud_group["预算(千元)"] = bud_group["金额(元)"] / 1000

    df_act_dept = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing")
    ]
    act_group = (
        df_act_dept.groupby("PROJ.DEPARTMENT1")["SIPM125.BXJE"].sum().reset_index())
    act_group["实际(千元)"] = act_group["SIPM125.BXJE"] / 1000
    act_group.rename(columns={"PROJ.DEPARTMENT1": "一级部门"}, inplace=True)

    df_dept = pd.merge(bud_group[["一级部门", "预算(千元)"]],
                       act_group[["一级部门", "实际(千元)"]],
                       on="一级部门", how="outer").fillna(0)

    df_dept["使用率"] = df_dept.apply(
        lambda r: 0 if r["预算(千元)"] == 0 else r["实际(千元)"] / r["预算(千元)"],
        axis=1
    )

    df_dept["排序"] = df_dept["一级部门"].apply(lambda x: dept_order.index(x) if x in dept_order else 999)

    df_dept = df_dept.sort_values("实际(千元)", ascending=False)   

    df_plot = df_dept.rename(columns={
        "预算(千元)": "Budget",
        "实际(千元)": "Actual"
    })

    fig2 = px.bar(
        df_plot,
        x="一级部门",
        y=["Budget", "Actual"],
        barmode="group",
        title="Annual Budget Usage by Division",
        labels={"value": "金额（千元）"},
        color_discrete_map={
            "Budget": "#B9BED1",
            "Actual": "#5B8DDE"
        },
        category_orders={"一级部门": df_plot["一级部门"].tolist()}  
    )
    fig2.update_layout(
        xaxis_title="",
        yaxis_title="kCNY",
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        margin=dict(l=20, r=20, t=20, b=10),
        legend_title_text="",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.10,
            xanchor="left",
            x=0.5,
            font=dict(family="Microsoft YaHei", size=12)
        ),
        bargap=0.35,
        bargroupgap=0, 
    )
    fig2.update_traces(
        hovertemplate=(
            "%{x}<br>"                     
            "%{y:.0f} kCNY"                
        ),
        hoverlabel=dict(
            bgcolor="#1C6DD0",
            font_size=11,
            font_family="Microsoft YaHei",
            font_color="white"
        ),
    )
    apply_fig_theme(fig2)
    
    fig2 = apply_axis_style1(fig2)

    chart2_graph = dcc.Graph(
        figure=fig2,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "350px"}
    )

    header_row = html.Tr(
        [html.Th("")] +
        [html.Th(dept) for dept in df_plot["一级部门"]]
    )

    budget_row = html.Tr(
        [html.Th("Budget")] +
        [html.Td(f"{v:.0f}") for v in df_plot["Budget"]]
    )

    actual_row = html.Tr(
        [html.Th("Actual")] +
        [html.Td(f"{v:.0f}") for v in df_plot["Actual"]]
    )

    usage_row = html.Tr(
        [html.Th("Usage")] +
        [html.Td(f"{v:.0%}") for v in df_plot["使用率"]]
    )
    table_html = html.Table(
        [
            html.Thead(header_row),
            html.Tbody([budget_row, actual_row, usage_row])
        ],
        style={
            "width": "100%",
            "marginTop": "5px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "12px",
            "textAlign": "left",
            "borderCollapse": "collapse"
        }
    )
    chart2 = html.Div([chart2_graph, table_html])

    df_act_pie = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing")
    ]

    df_pie = (df_act_pie.groupby("SIPM127.FYDTYPE")["SIPM125.BXJE"].sum().reset_index())
    df_pie["金额(千元)"] = df_pie["SIPM125.BXJE"] / 1000

    df_pie = df_pie.sort_values("金额(千元)", ascending=False)

    fig3 = px.pie(
        df_pie,
        values="金额(千元)",
        names="SIPM127.FYDTYPE",
        title="Actual Cost Distribution",
        color_discrete_sequence=my_colors
    )
    fig3.update_traces(
        texttemplate="%{label}<br>%{percent:.1%}",
        hovertemplate=(
            "%{label}<br>"
            "Amount：%{value:.0f} kCNY<br>"
            "Usage：%{percent:.1%}"
        ),
    )
    fig3.update_layout(
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        legend=dict(
            font=dict(family="Microsoft YaHei", size=12),
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.05,
        ),
        margin=dict(l=10, r=10, t=40, b=10)
    )
    apply_fig_theme(fig3)

    chart3 = dcc.Graph(
        figure=fig3,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "350px"}
    )

    df_bud_cat = df_budget[
        (df_budget["申请年份"] == year) &
        (df_budget["申请类型"] == "项目") &
        (df_budget["执行状态"] == "Ongoing")
    ]
    df_bud_cat_group = (df_bud_cat.groupby("费用大类")["金额(元)"].sum().reset_index())
    df_bud_cat_group["预算(千元)"] = df_bud_cat_group["金额(元)"] / 1000

    df_act_cat = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing")
    ]
    df_act_cat_group = (df_act_cat.groupby("SIPM127.FYDTYPE")["SIPM125.BXJE"].sum().reset_index())
    df_act_cat_group.rename(columns={"SIPM127.FYDTYPE": "费用大类"}, inplace=True)
    df_act_cat_group["实际(千元)"] = df_act_cat_group["SIPM125.BXJE"] / 1000

    df_cat = pd.merge(
        df_bud_cat_group[["费用大类", "预算(千元)"]],
        df_act_cat_group[["费用大类", "实际(千元)"]],
        on="费用大类",
        how="outer"
    ).fillna(0)

    df_cat["使用率"] = df_cat.apply(
        lambda r: 0 if r["预算(千元)"] == 0 else r["实际(千元)"] / r["预算(千元)"],
        axis=1
    )

    rename_map = {
        "Labour Cost": "Labour",
        "R&D Expense": "R&D",
        "Fixed Assets / Intangible Assets": "Assets",
        "Administrative Expense": "Admin",
    }
    df_cat["费用大类"] = df_cat["费用大类"].replace(rename_map)

    df_cat = df_cat.sort_values("实际(千元)", ascending=False)
    df_cat1 = df_cat.rename(columns={
        "预算(千元)": "Budget",
        "实际(千元)": "Actual"
    })

    fig4 = px.bar(
        df_cat1,
        x="费用大类",
        y=["Budget", "Actual"],
        barmode="group",
        title="Budget vs Actual by Category",
        labels={"value": "金额（千元）"},
        color_discrete_map={
            "Budget": "#B9BED1",
            "Actual": "#5B8DDE"
        }
    )

    fig4.update_traces(
        hovertemplate=(
            "%{x}<br>"
            "Amount：%{y:.0f} kCNY"
        )
    )
    fig4.update_layout(
        xaxis_title="",
        yaxis_title="kCNY",
        legend_title_text="",
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.12,
            xanchor="left",
            x=0.5,
            font=dict(family="Microsoft YaHei", size=12),
        ),
        margin=dict(l=10, r=20, t=40, b=10),
        bargap=0.35,
        bargroupgap=0, 
    )
    apply_fig_theme(fig4)
    fig4 = apply_axis_style1(fig4)
    chart4_graph = dcc.Graph(
        figure=fig4,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "350px"}
    )

    header_row4 = html.Tr(
        [html.Th("")] +
        [html.Th(cat) for cat in df_cat1["费用大类"]]
    )
    budget_row4 = html.Tr(
        [html.Th("Budget")] +
        [html.Td(f"{v:.0f}") for v in df_cat1["Budget"]]
    )
    actual_row4 = html.Tr(
        [html.Th("Actual")] +
        [html.Td(f"{v:.0f}") for v in df_cat1["Actual"]]
    )
    usage_row4 = html.Tr(
        [html.Th("Usage")] +
        [html.Td(f"{v:.0%}") for v in df_cat1["使用率"]]
    )
    table4 = html.Table(
        [
            html.Thead(header_row4),
            html.Tbody([budget_row4, actual_row4, usage_row4])
        ],
        style={
            "width": "100%",
            "marginTop": "5px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "12px",
            "textAlign": "left",
            "borderCollapse": "separate",
            "borderSpacing": "30px 0px"
        }
    )
    chart4 = html.Div([chart4_graph, table4])

    df_act_base = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing")
    ]

    df_top10_total = (df_act_base.groupby("SIPM125.NAME")["SIPM125.BXJE"].sum().reset_index())
    df_top10_total["金额(千元)"] = df_top10_total["SIPM125.BXJE"] / 1000

    df_top10_total = df_top10_total.sort_values("金额(千元)", ascending=False).head(10)

    fig5_top = px.bar(
        df_top10_total,
        x="金额(千元)",
        y="SIPM125.NAME",
        orientation="h",
        title="TOP10 Total Cost",
        labels={"金额(千元)": "千元"},
        color_discrete_sequence=[my_colors[4]],  
        category_orders={"SIPM125.NAME": df_top10_total["SIPM125.NAME"].tolist()[::-1]},
        text=df_top10_total["金额(千元)"].round(0).astype(int).astype(str)
    )
    max_val = df_top10_total["金额(千元)"].max()
    fig5_top.update_xaxes(range=[0, max_val * 1.15])      
    fig5_top.update_yaxes(autorange="reversed")
    fig5_top.update_layout(
        xaxis_title="kCNY",
        yaxis_title="",
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        margin=dict(l=0, r=10, t=30, b=10)
    )
    fig5_top.update_traces(
        hovertemplate=(
            "Proj：%{y}<br>"
            "Amount：%{x:.0f}"
        ),
        width=0.55,
        textposition="outside",
    )
    apply_fig_theme(fig5_top)
    fig5_top = apply_axis_style2(fig5_top)
    chart5_top_graph = dcc.Graph(
        figure=fig5_top,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "380px"}
    )

    df_top10_rd = df_act_base[
        (df_act_base["SIPM127.FYDTYPE"] == "R&D Expense") |
        (df_act_base["SIPM127.FYDTYPE"] == "Administrative Expense")
    ]
    df_top10_rd = (df_top10_rd.groupby("SIPM125.NAME")["SIPM125.BXJE"].sum().reset_index())
    df_top10_rd["金额(千元)"] = df_top10_rd["SIPM125.BXJE"] / 1000
    df_top10_rd = df_top10_rd.sort_values("金额(千元)", ascending=False).head(10)
    fig5_bottom = px.bar(
        df_top10_rd,
        x="金额(千元)",
        y="SIPM125.NAME",
        orientation="h",
        title="TOP10 R&D + Admin Cost",
        labels={"金额(千元)": "千元"},
        color_discrete_sequence=[my_colors[1]],  
        category_orders={"SIPM125.NAME": df_top10_rd["SIPM125.NAME"].tolist()[::-1]},
        text=df_top10_rd["金额(千元)"].round(0).astype(int).astype(str)
    )
    max_val = df_top10_rd["金额(千元)"].max()
    fig5_bottom.update_xaxes(range=[0, max_val * 1.15])      
    fig5_bottom.update_yaxes(autorange="reversed")
    fig5_bottom.update_layout(
        xaxis_title="kCNY",
        yaxis_title="",
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        margin=dict(l=10, r=10, t=30, b=10)
    )
    fig5_bottom.update_traces(
        hovertemplate=(
            "Proj：%{y}<br>"
            "Amount：%{x:.0f}"
        ),
        width=0.55,
        textposition="outside",
    )
    apply_fig_theme(fig5_bottom)
    fig5_bottom = apply_axis_style2(fig5_bottom)
    chart5_bottom_graph = dcc.Graph(
        figure=fig5_bottom,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "380px"}
    )

    chart5 = html.Div(
        [
            chart5_top_graph,
            html.Br(),
            chart5_bottom_graph,
        ]
    )

    df_top10_mp = df_act_base[
        df_act_base["SIPM127.FYDTYPE"] == "Labour Cost"
    ]
    df_top10_mp = (df_top10_mp.groupby("SIPM125.NAME")["SIPM125.BXJE"].sum().reset_index())
    df_top10_mp["金额(千元)"] = df_top10_mp["SIPM125.BXJE"] / 1000
    df_top10_mp = df_top10_mp.sort_values("金额(千元)", ascending=False).head(10)
    fig6_top = px.bar(
        df_top10_mp,
        x="金额(千元)",
        y="SIPM125.NAME",
        orientation="h",
        title="TOP10 Labour Cost",
        labels={"金额(千元)": "千元"},
        color_discrete_sequence=[my_colors[0]],
        category_orders={"SIPM125.NAME": df_top10_mp["SIPM125.NAME"].tolist()[::-1]},
        text=df_top10_mp["金额(千元)"].round(0).astype(int).astype(str)
    )
    max_val = df_top10_mp["金额(千元)"].max()
    fig6_top.update_xaxes(range=[0, max_val * 1.15])      
    fig6_top.update_yaxes(autorange="reversed")
    fig6_top.update_layout(
        xaxis_title="kCNY",
        yaxis_title="",
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        margin=dict(l=10, r=10, t=30, b=10)
    )
    fig6_top.update_traces(
        hovertemplate=(
            "Proj：%{y}<br>"
            "Amount：%{x:.0f}"
        ),
        width=0.55,
        textposition="outside",
    )
    apply_fig_theme(fig6_top)
    fig6_top = apply_axis_style2(fig6_top)
    chart6_top_graph = dcc.Graph(
        figure=fig6_top,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "380px"}
    )

    df_top10_asset = df_act_base[
        df_act_base["SIPM127.FYDTYPE"] == "Fixed Assets / Intangible Assets"
    ]
    df_top10_asset = (
        df_top10_asset.groupby("SIPM125.NAME")["SIPM125.BXJE"].sum().reset_index())
    df_top10_asset["金额(千元)"] = df_top10_asset["SIPM125.BXJE"] / 1000
    df_top10_asset = df_top10_asset.sort_values("金额(千元)", ascending=False).head(10)
    fig6_bottom = px.bar(
        df_top10_asset,
        x="金额(千元)",
        y="SIPM125.NAME",
        orientation="h",
        title="TOP10 Asset Cost",
        labels={"金额(千元)": "千元"},
        color_discrete_sequence=[my_colors[2]],
        category_orders={"SIPM125.NAME": df_top10_asset["SIPM125.NAME"].tolist()[::-1]},
        text=df_top10_asset["金额(千元)"].round(0).astype(int).astype(str)
    )
    max_val = df_top10_asset["金额(千元)"].max()
    fig6_bottom.update_xaxes(range=[0, max_val * 1.15])      
    fig6_bottom.update_yaxes(autorange="reversed")
    fig6_bottom.update_layout(
        xaxis_title="kCNY",
        yaxis_title="",
        title=dict(font=dict(size=16, family="Microsoft YaHei")),
        margin=dict(l=10, r=10, t=30, b=10)
    )
    fig6_bottom.update_traces(
        hovertemplate=(
            "Proj：%{y}<br>"
            "Amount：%{x:.0f}"
        ),
        width=0.55,
        textposition="outside",
    )
    apply_fig_theme(fig6_bottom)
    fig6_bottom = apply_axis_style2(fig6_bottom)
    chart6_bottom_graph = dcc.Graph(
        figure=fig6_bottom,
        config={"displayModeBar": False},
        className="fade-in",
        style={"height": "380px"}
    )

    chart6 = html.Div(
        [
            chart6_top_graph,
            html.Br(),
            chart6_bottom_graph,
        ]
    )
    return [chart1, chart2, chart3, chart4, chart5, chart6]

def get_dept_total(year, dept):

    df_bud = df_budget[
        (df_budget["申请年份"] == year) &
        (df_budget["申请类型"] == "项目") & 
        (df_budget["执行状态"] == "Ongoing")]
    total_bud = df_bud[df_bud["二级部门"] == dept]["金额(元)"].sum() / 1000

    df_act = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing")
    ]
    total_act = df_act[df_act["PROJ.DEPARTMENT2"] == dept]["SIPM125.BXJE"].sum() / 1000
    return total_bud, total_act

def get_max_projects_in_tab(year, dept_list):
    max_count = 0
    for dept in dept_list:

        df_bud = df_budget[
            (df_budget["申请年份"] == year) &
            (df_budget["申请类型"] == "项目") &
            (df_budget["执行状态"] == "Ongoing") &
            (df_budget["二级部门"] == dept)
        ]
        bud_ids = set(df_bud["项目编号"].dropna().astype(str))

        df_act = df_actual[
            (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
            (df_actual["SIPM124.SQLX"] == "项目") &
            (df_actual["PROJ.PSTATE"] == "Ongoing") &
            (df_actual["PROJ.DEPARTMENT2"] == dept)
        ]
        act_ids = set(df_act["SIPM125.NO"].dropna().astype(str))

        all_ids = bud_ids | act_ids
        max_count = max(max_count, len(all_ids))

    if max_count == 0:
        max_count = 1
    return max_count

def make_donut(total_budget, total_actual, title_text):
    used = total_actual
    remaining = max(total_budget - total_actual, 0)
    df = pd.DataFrame({
        "类型": ["Used", "Balance"],
        "金额": [used, remaining]
    })
    fig = px.pie(
        df,
        names="类型",
        values="金额",
        hole=0.45,
        title=title_text,
        color="类型",
        color_discrete_map={
            "Used": "#3366CC",
            "Balance": "#DDDDDD",
        }
    )

    fig.update_traces(
        hovertemplate="%{label}: %{value:.2f} kCNY",
        textinfo="none",       
    )

    usage_rate = 0 if total_budget == 0 else used / total_budget
    fig.add_annotation(
        text=f"{usage_rate:.0%}",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=28, color="#3366CC", family="Microsoft YaHei")
    )
    fig.update_layout(
        title=dict(
            text=title_text,
            font=dict(size=16, family="Microsoft YaHei"),
            x=0.5,
            y=0.97,        
        ),
        margin=dict(l=60, r=60, t=70, b=30),   
        showlegend=True,
        legend=dict(
            orientation="h",
            x=0.5, xanchor="center",
            y=1.10, yanchor="top",            
            font=dict(size=11, family="Microsoft YaHei"),
        )
    )
    apply_fig_theme(fig)
    return fig
def make_overlay_chart(dept, year, max_projects):
    df_master_local = df_master[["项目编号", "项目名称"]].copy()

    df_bud = df_budget[
        (df_budget["申请年份"] == year) &
        (df_budget["申请类型"] == "项目") &
        (df_budget["执行状态"] == "Ongoing") &
        (df_budget["二级部门"] == dept)
    ]
    df_bud_group = (
        df_bud.groupby("项目编号")["金额(元)"]
        .sum()
        .reset_index()
    )
    df_bud_group["预算(千元)"] = df_bud_group["金额(元)"] / 1000

    df_act = df_actual[
        (pd.to_datetime(df_actual["SIPM125.SQSJ"]).dt.year == year) &
        (df_actual["SIPM124.SQLX"] == "项目") &
        (df_actual["PROJ.PSTATE"] == "Ongoing") &
        (df_actual["PROJ.DEPARTMENT2"] == dept)
    ]
    df_act_group = (
        df_act.groupby("SIPM125.NO")["SIPM125.BXJE"].sum().reset_index().rename(columns={"SIPM125.NO": "项目编号"}))
    df_act_group["实际(千元)"] = df_act_group["SIPM125.BXJE"] / 1000

    df_merge = pd.merge(
        df_bud_group[["项目编号", "预算(千元)"]],
        df_act_group[["项目编号", "实际(千元)"]],
        on="项目编号",
        how="outer"
    ).fillna(0)

    df_merge = pd.merge(df_merge, df_master_local, on="项目编号", how="left")
    df_merge["项目名称"] = df_merge["项目名称"].fillna(df_merge["项目编号"].astype(str))

    df_merge["使用率"] = df_merge.apply(
        lambda r: 0 if r["预算(千元)"] == 0 else r["实际(千元)"] / r["预算(千元)"],
        axis=1
    )

    df_merge = df_merge.sort_values("实际(千元)", ascending=False)

    def wrap(text, width=6):
        return "<br>".join([text[i:i + width] for i in range(0, len(text), width)])
    df_merge["项目名称换行"] = df_merge["项目名称"].apply(lambda x: wrap(x, 7))

    n_projects = len(df_merge)
    x_positions = list(range(n_projects))  
    budget_width = 0.6  
    actual_width = 0.4
    fig = go.Figure()

    fig.add_bar(
        x=x_positions,
        y=df_merge["预算(千元)"],
        name="预算",
        marker_color="#D3D3D3",
        width=budget_width,
        hovertemplate="Budget：%{y:.2f}<extra> kCNY</extra>",
    )

    fig.add_bar(
        x=x_positions,
        y=df_merge["实际(千元)"],
        name="实际",
        marker_color="#5B8DDE",
        width=actual_width,
        hovertemplate="Actual：%{y:.2f}<extra> kCNY</extra>",
    )
 
    safe_max = max(max_projects, 1)
    fig.update_xaxes(
        range=[-0.5, safe_max - 0.5],
        tickmode="array",
        tickvals=x_positions,
        ticktext=df_merge["项目名称换行"],
        tickangle=0, 
    )

    if len(df_merge) > 0:
        max_budget = df_merge["预算(千元)"].max()
        max_actual = df_merge["实际(千元)"].max()
        ymax_base = max(max_budget, max_actual)
        if ymax_base <= 0:
            ymax_base = 1
    else:
        ymax_base = 1
    fig.update_yaxes(range=[0, ymax_base * 1.1])

    y_budget = -0.21
    y_actual = y_budget - 0.05
    y_rate   = y_actual - 0.05
    LABEL_X = -0.06 
    fig.add_annotation(
        xref="paper", yref="paper",
        x=LABEL_X, y=y_budget,
        text="Budget",
        showarrow=False,
        font=dict(size=12),
    )
    fig.add_annotation(
        xref="paper", yref="paper",
        x=LABEL_X, y=y_actual,
        text="Actual",
        showarrow=False,
        font=dict(size=12),        
    )
    fig.add_annotation(
        xref="paper", yref="paper",
        x=LABEL_X, y=y_rate,
        text="Usage",
        showarrow=False,
        font=dict(size=12),        
    )

    for idx, row in df_merge.iterrows():
        x = x_positions[list(df_merge.index).index(idx)]
        fig.add_annotation(
            x=x, yref="paper", y=y_budget,
            text=f"{row['预算(千元)']:.2f}",
            showarrow=False, font=dict(size=12),
            align="center"
        )
        fig.add_annotation(
            x=x, yref="paper", y=y_actual,
            text=f"{row['实际(千元)']:.2f}",
            showarrow=False, font=dict(size=12),
            align="center"
        )
        fig.add_annotation(
            x=x, yref="paper", y=y_rate,
            text=f"{row['使用率']:.0%}",
            showarrow=False, font=dict(size=12),
            align="center"
        )
    fig.update_layout(
        title=dept,
        barmode="overlay",
        margin=dict(l=60, r=60, t=40, b=110),
        xaxis_title="",
        yaxis_title="kCNY",
        showlegend=False, 
    )
    apply_fig_theme(fig)
    apply_axis_style1(fig)
    return fig
@app.callback(
    [
        Output("r1_deptA_donut", "figure"),
        Output("r1_deptA_donut_table", "children"),  
        Output("r1_deptA_overlay", "figure"),
        Output("r1_deptB_donut", "figure"),
        Output("r1_deptB_donut_table", "children"),  
        Output("r1_deptB_overlay", "figure"),
    ],
    Input("year-select", "value")
)
def update_r1_tab(year):
    all_depts = r1_depts + r2_depts + lab2030_depts + office_depts
    max_projects_all = get_max_projects_in_tab(year, all_depts)
    deptA = "Team7"
    deptB = "Team3"

    budA, actA = get_dept_total(year, deptA)
    donutA = make_donut(budA, actA, f"{deptA} Budget Usage")
    overlayA = make_overlay_chart(deptA, year, max_projects_all)
    tableA = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {budA:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )
    
    budB, actB = get_dept_total(year, deptB)
    donutB = make_donut(budB, actB, f"{deptB} Budget Usage")
    overlayB = make_overlay_chart(deptB, year, max_projects_all)
    tableB = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {budB:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )
    return donutA, tableA, overlayA, donutB, tableB, overlayB

@app.callback(
    [
        Output("r2_dept1_donut", "figure"),
        Output("r2_dept1_donut_table", "children"),  
        Output("r2_dept1_overlay", "figure"),
        Output("r2_dept2_donut", "figure"),
        Output("r2_dept2_donut_table", "children"),  
        Output("r2_dept2_overlay", "figure"),
        Output("r2_dept3_donut", "figure"),
        Output("r2_dept3_donut_table", "children"),  
        Output("r2_dept3_overlay", "figure"),
    ],
    Input("year-select", "value")
)
def update_r2_tab(year):

    all_depts = r1_depts + r2_depts + lab2030_depts + office_depts
    max_projects_all = get_max_projects_in_tab(year, all_depts)

    dept1 = "Team9"
    dept2 = "Team10"
    dept3 = "Team19"

    bud1, act1 = get_dept_total(year, dept1)
    donut1 = make_donut(bud1, act1, f"{dept1} Budget Usage")
    overlay1 = make_overlay_chart(dept1, year, max_projects_all)
    table1 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud1:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )

    bud2, act2 = get_dept_total(year, dept2)
    donut2 = make_donut(bud2, act2, f"{dept2} Budget Usage")
    overlay2 = make_overlay_chart(dept2, year, max_projects_all)
    table2 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud2:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )

    bud3, act3 = get_dept_total(year, dept3)
    donut3 = make_donut(bud3, act3, f"{dept3} Budget Usage")
    overlay3 = make_overlay_chart(dept3, year, max_projects_all)
    table3 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud3:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )
    return donut1, table1, overlay1, donut2, table2, overlay2, donut3, table3, overlay3

@app.callback(
    [
        Output("lab2030_dept1_donut", "figure"),
        Output("lab2030_dept1_donut_table", "children"),  
        Output("lab2030_dept1_overlay", "figure"),
        Output("lab2030_dept2_donut", "figure"),
        Output("lab2030_dept2_donut_table", "children"),  
        Output("lab2030_dept2_overlay", "figure"),
        Output("lab2030_dept3_donut", "figure"),
        Output("lab2030_dept3_donut_table", "children"),  
        Output("lab2030_dept3_overlay", "figure"),
        Output("lab2030_dept4_donut", "figure"),
        Output("lab2030_dept4_donut_table", "children"),  
        Output("lab2030_dept4_overlay", "figure"),
    ],
    Input("year-select", "value")
)
def update_lab2030_tab(year):
    all_depts = r1_depts + r2_depts + lab2030_depts + office_depts
    max_projects_all = get_max_projects_in_tab(year, all_depts)
    dept1 = "Team1"
    dept2 = "Team8"
    dept3 = "Team17"
    dept4 = "Team12"

    bud1, act1 = get_dept_total(year, dept1)
    donut1 = make_donut(bud1, act1, f"{dept1} Budget Usage")
    overlay1 = make_overlay_chart(dept1, year, max_projects_all)
    table1 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud1:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )

    bud2, act2 = get_dept_total(year, dept2)
    donut2 = make_donut(bud2, act2, f"{dept2} Budget Usage")
    overlay2 = make_overlay_chart(dept2, year, max_projects_all)
    table2 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud2:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )

    bud3, act3 = get_dept_total(year, dept3)
    donut3 = make_donut(bud3, act3, f"{dept3} Budget Usage")
    overlay3 = make_overlay_chart(dept3, year, max_projects_all)
    table3 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud3:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )

    bud4, act4 = get_dept_total(year, dept4)
    donut4 = make_donut(bud4, act4, f"{dept4} Budget Usage")
    overlay4 = make_overlay_chart(dept4, year, max_projects_all)
    table4 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud4:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )
    return donut1, table1, overlay1, donut2, table2, overlay2, donut3, table3, overlay3, donut4, table4, overlay4

@app.callback(
    [
        Output("office_dept1_donut", "figure"),
        Output("office_dept1_donut_table", "children"),  
        Output("office_dept1_overlay", "figure"),
        Output("office_dept2_donut", "figure"),
        Output("office_dept2_donut_table", "children"),  
        Output("office_dept2_overlay", "figure"),
    ],
    Input("year-select", "value")
)
def update_office_tab(year):
    all_depts = r1_depts + r2_depts + lab2030_depts + office_depts
    max_projects_all = get_max_projects_in_tab(year, all_depts)
    dept1 = "Team16"
    dept2 = "Team11"

    bud1, act1 = get_dept_total(year, dept1)
    donut1 = make_donut(bud1, act1, f"{dept1} Budget Usage")
    overlay1 = make_overlay_chart(dept1, year, max_projects_all)
    table1 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud1:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )

    bud2, act2 = get_dept_total(year, dept2)
    donut2 = make_donut(bud2, act2, f"{dept2} Budget Usage")
    overlay2 = make_overlay_chart(dept2, year, max_projects_all)
    table2 = html.Table(
        [
            html.Thead(
                html.Tr([
                    html.Th(f"Total Budget = {bud2:,.0f} kCNY")
                ])
            )
        ],
        style={
            "width": "100%",
            "marginTop": "10px",
            "fontFamily": "Microsoft YaHei",
            "fontSize": "16px",
            "textAlign": "center",
            "borderCollapse": "collapse",
        },
    )
    return donut1, table1, overlay1, donut2, table2, overlay2

app.clientside_callback(
    """
    function (n_clicks) {
        // 第一次还没点时，什么都不做
        if (!n_clicks) {
            return window.dash_clientside.no_update;
        }
        // 调起浏览器打印对话框
        window.print();
        // 返回一个空字符串给隐藏 div
        return "";
    }
    """,
    Output("print-dummy", "children"),
    Input("btn-print-pdf", "n_clicks"),
)

if __name__ == "__main__":
    app.run(debug=True, port=8090)