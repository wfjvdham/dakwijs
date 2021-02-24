# -*- coding: utf-8 -*-

# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.

import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
from docx import Document
from docx.shared import Inches

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

    html.Label('Paneeldikte [mm]'),
    dcc.Input(id='Paneeldikte',value='0',type='value'),

    html.Label('Aantal rijen'),
    dcc.Input(id='Rijen',value='0',type='value'),

    html.Label("Aantal kolommen"),
    dcc.Input(id='Kolommen',value='0',type='value'),

    html.Label("Kleur frame"),
    dcc.Dropdown(
        id='KleurFrame',
        options=[
            {'label': 'ALU', 'value': 'ALU'},
            {'label': 'ALU Zwart', 'value': 'ALZ'}
        ],
        value='ALU'
    ),

    html.Label("Plaat / Paneel"),
    dcc.Dropdown(
        id='Toepassing',
        options=[
            {'label': 'Plaat', 'value': 'PLA'},
            {'label': 'Paneel', 'value': 'PAN'}
        ],
        value='Paneel'
    ),

    html.Label('Slider'),
    dcc.Slider(
        min=5,
        max=90,
        value=5,
    ),
    
    html.Br(),
    html.Button('Download document', id='button'),
    html.Br(),
    html.P(id='placeholder'),
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

@app.callback(
    Output("placeholder", "children"),
    Input('button', 'n_clicks')
)
def download(n_clicks):
    document = Document()

    document.add_heading('Document Title', 0)

    p = document.add_paragraph('A plain paragraph having some ')
    p.add_run('bold').bold = True
    p.add_run(' and some ')
    p.add_run('italic.').italic = True

    document.add_heading('Heading, level 1', level=1)
    document.add_paragraph('Intense quote', style='Intense Quote')

    document.add_paragraph(
        'first item in unordered list', style='List Bullet'
    )
    document.add_paragraph(
        'first item in ordered list', style='List Number'
    )

    #document.add_picture('monty-truth.png', width=Inches(1.25))

    records = (
        (3, '101', 'Spam'),
        (7, '422', 'Eggs'),
        (4, '631', 'Spam, spam, eggs, and spam')
    )

    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Qty'
    hdr_cells[1].text = 'Id'
    hdr_cells[2].text = 'Desc'
    for qty, id, desc in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(qty)
        row_cells[1].text = id
        row_cells[2].text = desc

    document.add_page_break()
    print('save document')
    document.save('demo.docx')

if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8080, debug=True)