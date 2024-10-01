import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import pandas as pd

# Load the Excel file
file_path = 'DUT Research.xlsx'  # Update this path to your actual file path

# Read data from each sheet into DataFrames
df_sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')
df_sheet3 = pd.read_excel(file_path, sheet_name='Sheet3')
df_sheet4 = pd.read_excel(file_path, sheet_name='Sheet4')

# Prepare the Dash app
app = dash.Dash(__name__)

# Layout of the app
app.layout = html.Div([
    html.Div([
        html.H1("DUT Research Dashboard", style={'display': 'inline-block', 'verticalAlign': 'middle'}),
        # Increase the logo size to 300%
        html.Img(src='/assets/my_image.png', style={'display': 'inline-block', 'transform': 'scale(3)', 'float': 'right'})
    ], style={'width': '50%', 'display': 'flex', 'alignItems': 'center', 'justifyContent': 'space-between'}),

    dcc.Dropdown(
        id='graph-selector',
        options=[
            {'label': 'Postgraduate Enrolment (2020-2023)', 'value': 'graph1'},
            {'label': 'FAS Postgraduate Enrolment (2020-2023)', 'value': 'graph2'},
            {'label': '2023 Student Enrolment by Level', 'value': 'graph3'},
            {'label': 'Postgraduate Graduation Rate (2015-2023)', 'value': 'graph4'},
            {'label': 'Postgraduate Enrolment 2024 (Image)', 'value': 'image1'},
            {'label': 'Current Postdoctoral Fellows (Image)', 'value': 'image2'},
            {'label': 'Emeritus/Honorary/Adjunct Professors (Image)', 'value': 'image3'},
            {'label': 'Departmental Research Outputs 2023 (Image)', 'value': 'image4'}
        ],
        value='graph1'
    ),

    dcc.Graph(id='graph-output', style={'display': 'block'}),
    # Decrease the width of other images
    html.Img(id='image-output', style={'display': 'none', 'width': '50%', 'height': 'auto'}),  # Reduced width to 50%
    html.Div(id='image-title', style={'text-align': 'center', 'font-size': '20px', 'margin-top': '10px'})
])

# Callback function to update the graph or image based on the dropdown selection
@app.callback(
    [Output('graph-output', 'figure'), Output('graph-output', 'style'),
     Output('image-output', 'src'), Output('image-output', 'style'),
     Output('image-title', 'children')],
    Input('graph-selector', 'value')
)
def update_output(selected_graph):
    if selected_graph == 'graph1':
        # Filter data for Sheet1
        df_filtered = df_sheet1[['Postgraduate Enrolment', '2020', '2021', '2022', '2023']].copy()
        df_filtered.set_index('Postgraduate Enrolment', inplace=True)

        # Create the bar chart for graph1
        fig = go.Figure()
        for col in df_filtered.columns:
            fig.add_trace(go.Bar(
                x=df_filtered.index,
                y=df_filtered[col],
                name=col,
                text=df_filtered[col],  # Display actual values
                textposition='auto'
            ))
        fig.update_layout(
            title='Postgraduate Enrolment - Actual Student Numbers (2020-2023)',
            xaxis_title='Subjects',
            yaxis_title='Number of Students',
            barmode='group'
        )
        return fig, {'display': 'block'}, None, {'display': 'none'}, ''

    elif selected_graph == 'graph2':
        # Filter data for Sheet2
        df_filtered = df_sheet2[['FAS Postgraduate Enrolment', '2020', '2021', '2022', '2023']].copy()
        df_filtered.set_index('FAS Postgraduate Enrolment', inplace=True)
        df_filtered *= 100  # Convert values to percentage

        # Create the bar chart for graph2
        fig = go.Figure()
        for col in df_filtered.columns:
            fig.add_trace(go.Bar(
                x=df_filtered.index,
                y=df_filtered[col],
                name=col,
                text=[f'{val:.0f}%' for val in df_filtered[col]],  # Display percentages
                textposition='auto'
            ))
        fig.update_layout(
            title='FAS Postgraduate Enrolment (2020-2023)',
            xaxis_title='Category',
            yaxis_title='Enrolment (%)',
            barmode='group'
        )
        return fig, {'display': 'block'}, None, {'display': 'none'}, ''

    elif selected_graph == 'graph3':
        # Filter data for Sheet3
        df_filtered = df_sheet3[['2023 Student Enrolment by Level', 'UG (NQF 5-7)', 'PG upto Masters (NQF8)', 'PG (NQF9-10)']].copy()
        df_filtered.set_index('2023 Student Enrolment by Level', inplace=True)

        # Create the bar chart for graph3
        fig = go.Figure()
        for col in df_filtered.columns:
            fig.add_trace(go.Bar(
                x=df_filtered.index,
                y=df_filtered[col],
                name=col,
                text=df_filtered[col],  # Display actual values
                textposition='auto'
            ))
        fig.update_layout(
            title='2023 Student Enrolment by Level',
            xaxis_title='Programs',
            yaxis_title='Number of Students',
            barmode='group'
        )
        return fig, {'display': 'block'}, None, {'display': 'none'}, ''

    elif selected_graph == 'graph4':
        # Strip any leading/trailing spaces from column names and ensure integer type for graduation rate
        df_sheet4.columns = df_sheet4.columns.str.strip()
        df_sheet4['Postgraduate Graduation Rate'] = df_sheet4['Postgraduate Graduation Rate'].astype(int)

        # Create the line graph for graph4
        fig = go.Figure(data=go.Scatter(
            x=df_sheet4['Postgraduate Graduation Rate'],
            y=df_sheet4['Faculty'],
            mode='lines+markers',
            text=[f'{val}%' for val in df_sheet4['Faculty']],  # Display percentage values
            textposition='bottom center',
            line=dict(color='blue')
        ))
        fig.update_layout(
            title='Postgraduate Graduation Rate (2015-2023)',
            xaxis_title='Year',
            yaxis_title='Graduation Rate (%)',
            xaxis=dict(tickmode='linear')  # Ensure all years are displayed
        )
        return fig, {'display': 'block'}, None, {'display': 'none'}, ''

    elif selected_graph == 'image1':
        # Display first image (Postgraduate Enrolment 2024)
        return {}, {'display': 'none'}, '/assets/1.png', {'display': 'block', 'width': '70%'}, 'Postgraduate Enrolment 2024'

    elif selected_graph == 'image2':
        # Display second image (Current Postdoctoral Fellows)
        return {}, {'display': 'none'}, '/assets/2.png', {'display': 'block', 'width': '70%'}, 'Current Postdoctoral Fellows'

    elif selected_graph == 'image3':
        # Display third image (Emeritus/Honorary/Adjunct Professors)
        return {}, {'display': 'none'}, '/assets/3.png', {'display': 'block', 'width': '70%'}, 'Emeritus/Honorary/Adjunct Professors'

    elif selected_graph == 'image4':
        # Display fourth image (Departmental Research Outputs 2023)
        return {}, {'display': 'none'}, '/assets/4.png', {'display': 'block', 'width': '70%'}, 'Departmental Research Outputs 2023'

# Run the Dash app
server = app.server  # Expose the Flask server for Gunicorn

if __name__ == '__main__':
    app.run_server(debug=True)
