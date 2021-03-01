# -*- coding: utf-8 -*-

# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.

import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import dash_table
import math

import pandas as pd

from docx import Document

from mailmerge import MailMerge
from datetime import date

template = "template.docx"
document = MailMerge(template)
print(document.get_merge_fields())

df = pd.read_excel("./Solor 2021.xlsm", sheet_name=1, names=['id', 'desc', 'price'], usecols=[0, 1, 2])
df['count'] = 0

app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            dbc.FormGroup([
                dbc.Label('Indak / Opdak', html_for='Daksysteem'),
                dcc.Dropdown(
                    id='Daksysteem',
                    options=[
                        {'label': 'Indak', 'value': 'Indak'},
                        {'label': 'Opdak', 'value': 'Opdak'},
                    ],
                    value='Indak'
                ),
            ]),
            dbc.FormGroup([
                dbc.Label('Landscape / Portrait', html_for='Indeling'),
                dcc.Dropdown(
                    id='Indeling',
                    options=[
                        {'label': 'Landscape', 'value': 'LND'},
                        {'label': 'Portrait', 'value': 'POR'},
                    ],
                    value='Landscape'
                ),
            ]),
            dbc.FormGroup([
                dbc.Label('Paneellengte [mm]', html_for='Paneellengte'),
                dbc.Input(id='Paneellengte', value=0, min=0, type='number'),
            ]),
            dbc.FormGroup([
                dbc.Label('Paneelbreedte [mm]', html_for='Paneelbreedte'),
                dbc.Input(id='Paneelbreedte', value=0, type='number'),
            ]),
            dbc.FormGroup([
                dbc.Label('Paneeldikte [mm]', html_for='Paneeldikte'),
                dbc.Input(id='Paneeldikte', value=0, type='number'),
            ]),
            dbc.FormGroup([
                dbc.Label('Aantal rijen', html_for='Rijen'),
                dbc.Input(id='Rijen', value=0, type='number'),
            ]),
            dbc.FormGroup([
                dbc.Label("Aantal kolommen", html_for='Kolommen'),
                dbc.Input(id='Kolommen', value=0, type='number'),
            ]),
            dbc.FormGroup([
                dbc.Label("Kleur frame", html_for='KleurFrame'),
                dcc.Dropdown(
                    id='KleurFrame',
                    options=[
                        {'label': 'ALU', 'value': 'ALU'},
                        {'label': 'ALU Zwart', 'value': 'ALZ'}
                    ],
                    value='ALU'
                ),
            ]),
            dbc.FormGroup([
                dbc.Label("Plaat / Paneel", html_for='Toepassing'),
                dcc.Dropdown(
                    id='Toepassing',
                    options=[
                        {'label': 'Plaat', 'value': 'PLA'},
                        {'label': 'Paneel', 'value': 'PAN'}
                    ],
                    value='Paneel'
                ),
            ])
        ], width={"size": 3, "offset": 1}),
        dbc.Col([
            dbc.Button('Download document', id='button'),
            html.P(id='placeholder'),
        ], width=4),
        dbc.Col(
            html.Div([
                dbc.Label('Lengte Rail', html_for='lengte_rail'),
                html.Div(id='lengte_rail'),
            ]),
            width=4
        ),
    ], form=True),
    dbc.Row([
        dbc.Col([
            html.Div(id='table')
        ], width={"size": 10, "offset": 1})
    ])
], fluid=True)

@app.callback(
    Output(component_id='lengte_rail', component_property='children'),
    Input(component_id='Indeling', component_property='value'),
    Input(component_id='Paneelbreedte', component_property='value'),
    Input(component_id='Rijen', component_property='value')
)
def update_output_div(indeling, paneelbreedte, rijen):
    result = paneelbreedte * rijen
    if indeling == 'LND':
        result += 10
    return result

@app.callback(
    Output('table', 'children'),
    Input('lengte_rail', 'children')
)
def update_datatable(rijen):
    df.loc[df['id'] == 770003, ['count']] = round(1.123345667, 1)
    df.loc[df['id'] == 770212, ['count']] = math.ceil(rijen) * 2

    df_result = df.loc[df['count'] > 0]
    data = df_result.to_dict('rows')
    columns = [{"name": i, "id": i, } for i in df.columns]
    return dash_table.DataTable(data=data, columns=columns)

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