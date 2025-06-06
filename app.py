import openpyxl
import pandas as pd
import dash
import plotly.express as px
from dash import dcc, html, dash_table
from dash.dependencies import Output, Input

data = pd.read_excel(r'C:\Users\user\Downloads\Melano Acnes Sales Report.xlsx', sheet_name='Data')
data["Channel"] = data["Channel"].astype(str).str.strip()
data["Brand"] = data["Brand"].astype(str).str.strip()
data['Brand'] = data["Brand"].apply(lambda x: x.encode('ascii', 'ignore').decode())
data["Channel"] = data["Channel"].apply(lambda x: x.encode('ascii', 'ignore').decode())
data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
data["Month_"] = data["Date"].dt.to_period("M")
monthly_sales = data.groupby("Month_")["Value"].sum().reset_index()



app = dash.Dash(__name__)

app.layout = html.Div([
            html.H1("Activation Insights Report"),
            dcc.RadioItems(
                id='brand-radio',
                options=[
                    {'label': 'Melano', 'value':'Melano'},
                    {'label':'Acnes', 'value': 'Acnes'}
                ],
                value='Melano',
                labelStyle={'display': 'block'}
            ),
            dcc.Dropdown(
                id='Channel-dropdown',
                options=[{'label': c, 'value': c} for c in sorted(data['Channel'].unique())],
                placeholder="Select a channel",
                style ={'width': '50%', 'margin':'auto', 'border_radius': '8px'}
            ),
            html.Div([
                dcc.Graph(id="Monthly-trend", style={"width": "48%", 'height':'350px'}),
                dcc.Graph(id="MoM-Trend", style={"width":'48%','height':'350px'}) 
            ],style={'display':'flex'}),
            dcc.Graph(id="SKU-Trend", style={'height':'380px'}),
            html.Br(),
            html.Div(id="output-content",
                     style={
                         "margin-top": "15px",
                         "font-size": "20px",
                         "font-weight": "bold",
                         "text-align": "center",
                         "padding": "15px",
                         "border-radius": "8px",
                         "border" : "2px solid #ccc",
                         "background-color" : "#f5f5f5",
                         "color" : "#333"
                     }),
            html.Div(id="Summary-content",
                    style = {'margin-top': '15px', 'font-size':'18px', 'font-weight':'bold',
                          'text-align':'left', 'padding':'15px', 'border-radius':'8px',
                           'border':'2px solid #ccc', 'background-color':'#f9f9f9',
                           'color':'#333'}
            )
])

@app.callback(
    Output("output-content", "children"),
    Input("Channel-dropdown", "value")
)
def output_data(select_channel):
    if select_channel:
        return f"Selected Channel: {select_channel}"
    return "Select Channel to view the insights"

@app.callback(
    Output("Monthly-trend", "figure"),
    [Input('brand-radio', 'value'), Input("Channel-dropdown", "value")]
)
def update_trend(selected_brand, selected_channel):
    title_text = f"Monthly Sales Trend"
    filtered_df = data[data['Brand'].str.strip() == selected_brand]
    
    if selected_channel:
        filtered_df = filtered_df[filtered_df["Channel"].str.strip() == selected_channel]
        title_text += f" ({selected_brand} in {selected_channel})"
        
    else:
        filtered_df = filtered_df
    monthly_sales_filtered = filtered_df.groupby("Month_")['Value'].sum().reset_index()
    monthly_sales_filtered['Month_'] = monthly_sales_filtered["Month_"].astype(str)
    fig = px.line(
            monthly_sales_filtered,
            x="Month_",
            y="Value",
            title=title_text,
            markers=True
        )
    fig.update_layout(title_font_size=15, xaxis_title_font_size=12, yaxis_title_font_size=12, template='plotly_white')
    return fig

@app.callback(
        Output('MoM-Trend', 'figure'),
        [Input('brand-radio', 'value'), Input('Channel-dropdown', 'value')]

)
def update_mom_trend(selected_brand, selected_channel):
    filtered_df = data[data['Brand'].str.strip() == selected_brand]
    if selected_channel:
        filtered_df = filtered_df[filtered_df["Channel"].str.strip() == selected_channel]
    monthly_sales_filtered = filtered_df.groupby("Month_")['Value'].sum().reset_index()
    monthly_sales_filtered['MoM Growth (%)'] = round(monthly_sales_filtered['Value'].pct_change() * 100,2)
    monthly_sales_filtered['Month_'] = monthly_sales_filtered['Month_'].astype(str)

    fig_mom = px.bar(
        monthly_sales_filtered, x = 'Month_', y = 'MoM Growth (%)',
        text="MoM Growth (%)",
        color= "MoM Growth (%)",
        color_continuous_scale=['red', 'green']
    )
    fig_mom.update_layout(title_font_size=15, xaxis_title_font_size=12, yaxis_title_font_size=12, template='plotly_white')

    return fig_mom

@app.callback(
    Output('SKU-Trend', 'figure'),
    [Input('brand-radio', 'value'), Input('Channel-dropdown', 'value')]
)

def updated_sku_trend(selected_brand, selected_channel):
    title_text = f"SKU Breakdown ({selected_brand}) "
    filtered_df = data[data['Brand'].str.strip() == selected_brand]
    if selected_channel:
        title_text += f"in {selected_channel}"
        filtered_df = filtered_df[filtered_df["Channel"].str.strip() == selected_channel]
    else:
        title_text += "in all channels"

    sku_sales = filtered_df.groupby(["Attribute"])["Value"].sum().reset_index()

    fig_sku =px.bar(
        sku_sales, x = 'Attribute', y = 'Value',
        title=title_text,
        text= 'Value',
        color='Value',
        color_continuous_scale=['blue', 'purple']
    )
    fig_sku.update_layout(title_font_size=15, xaxis_title_font_size=12, yaxis_title_font_size=12, template='plotly_white')

    return fig_sku

@app.callback(
    Output('Summary-content', 'children'),
    [Input('brand-radio', 'value'), Input('Channel-dropdown', 'value')]
)

def generate_summary(selected_brand, selected_channel):
    filtered_df = data[data['Brand'].str.strip() == selected_brand]

    if selected_channel:
        filtered_df = filtered_df[filtered_df["Channel"].str.strip() == selected_channel]

    monthly_sales_filtered = filtered_df.groupby("Month_")['Value'].sum().reset_index()
    monthly_sales_filtered['MoM Growth (%)'] = round(monthly_sales_filtered['Value'].pct_change() * 100,2)

    last_month = monthly_sales_filtered.iloc[-1]["Month_"]
    latest_value = monthly_sales_filtered.iloc[-1]['Value']
    latest_growth = monthly_sales_filtered.iloc[-1]['MoM Growth (%)']

    summary_text = f"In {last_month}, {selected_brand} recorded sales of {latest_value} "
    summary_text += f"with a {'growth' if latest_growth > 0 else 'decline'} of {abs(latest_growth)}% MoM"

    return summary_text


    
        
    

if __name__ =='__main__':
    app.run(debug=True)