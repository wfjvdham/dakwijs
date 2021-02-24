# -*- coding: utf-8 -*-

# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.

import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div([
    html.Label('Indak / Opdak'),
    dcc.Dropdown(
        id='Daksysteem',
        options=[
            {'label': 'Indak', 'value': 'Indak'},
            {'label': 'Opdak', 'value': 'Opdak'},
        ],
        value='Indak'
    ),

    html.Label('Landscape / Portrait'),
    dcc.Dropdown(
        id='Indeling',
        options=[
            {'label': 'Landscape', 'value': 'LND'},
            {'label': 'Portrait', 'value': 'POR'},
        ],
        value='Landscape'
    ),

    html.Label('Paneellengte [mm]'),
    dcc.Input(id='Paneellengte',value='0',type='value'),

    html.Label('Paneelbreedte [mm]'),
    dcc.Input(id='Paneelbreedte',value='0',type='value'),

    html.Label('Paneeldikte [mm]')
    dcc.Input(id='Paneeldikte',value='0',type='value')

    html.Label('Aantal rijen'),
    dcc.Input(id='Rijen',value='0',type='value'),

    html.Label("Aantal kolommen"),
    dcc.Input(id='Kolommen',value='0',type='value')

    html.Label("Kleur frame"),
    dcc.Dropdown(
        id='KleurFrame',
        options=[
            {'label': 'ALU', 'value': 'ALU'},
            {'label': 'ALU Zwart', 'value': 'ALZ'}
        ],
        value='ALU'
    ),

    html.Label("Plaat / Paneel")
    dcc.Dropdown(
        id='Toepassing'
        options=[
            {'label': 'Plaat', 'value': 'PLA'},
            {'label': 'Paneel', 'value': 'PAN'}
        ]
        value='Paneel'
    )

    html.Label('Slider'),
    dcc.Slider(
        min=5,
        max=90,
        value=5,
    ),
    
    html.Br(),
    html.Div(id='lengte_rail'),
], style={'columnCount': 3})

@app.callback(
    Output(component_id='lengte_rail', component_property='children'),
    Input(component_id='Indeling', component_property='value'),
    Input(component_id='Paneelbreedte', component_property='value'),
    Input(component_id='Rijen', component_property= 'value')
)
def update_output_div(indeling, paneelbreedte, rijen):
    return (input_value_1 + ' en ' + ', '.join(input_value_2)

if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8080, debug=True)