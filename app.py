# -*- coding: utf-8 -*-

# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.

import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
import dash_table
import math
import pandas as pd
from mailmerge import MailMerge
import os
import flask
import base64
import json

df = pd.read_excel("./Solor 2021.xlsm", sheet_name=1, names=['id', 'desc', 'price'], usecols=[0, 1, 2], dtype={'id': str, 'desc': str, 'price': str})
df['count'] = 0

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.JOURNAL])

input_tab = dbc.Card(
    dbc.CardBody([
        dbc.FormGroup([
            dbc.Label('Indak / Opdak', html_for='daksysteem'),
            dcc.Dropdown(
                id='daksysteem',
                options=[
                    {'label': 'Indak', 'value': 'Indak'},
                    {'label': 'Opdak', 'value': 'Opdak'},
                ],
                value='Indak', clearable=False
            ),
        ]),
        dbc.FormGroup([
            dbc.Label('Landscape / Portrait', html_for='indeling'),
            dcc.Dropdown(
                id='indeling',
                options=[
                    {'label': 'Landscape', 'value': 'LND'},
                    {'label': 'Portrait', 'value': 'POR'},
                ],
                value='LND', clearable=False
            ),
        ]),
        dbc.FormGroup([
            dbc.Label('Paneellengte [mm]', html_for='paneellengte'),
            dbc.Input(id='paneellengte', value=0, min=0, type='number'),
        ]),
        dbc.FormGroup([
            dbc.Label('Paneelbreedte [mm]', html_for='paneelbreedte'),
            dbc.Input(id='paneelbreedte', value=0, min=0, type='number'),
        ]),
        dbc.FormGroup([
            dbc.Label('Paneeldikte [mm]', html_for='paneeldikte'),
            dbc.Input(id='paneeldikte', value=0, min=0, type='number'),
        ]),
        dbc.FormGroup([
            dbc.Label('Aantal rijen', html_for='rijen'),
            dbc.Input(id='rijen', value=0, min=0, type='number'),
        ]),
        dbc.FormGroup([
            dbc.Label("Aantal kolommen", html_for='kolommen'),
            dbc.Input(id='kolommen', value=0, min=0, type='number'),
        ]),
        dbc.FormGroup([
            dbc.Label("Kleur frame", html_for='kleurFrame'),
            dcc.Dropdown(
                id='kleurFrame',
                options=[
                    {'label': 'ALU', 'value': 'ALU'},
                    {'label': 'ALU Zwart', 'value': 'ALZ'}
                ],
                value='ALU', clearable=False
            ),
        ]),
        dbc.FormGroup([
            dbc.Label("Plaat / Paneel", html_for='toepassing'),
            dcc.Dropdown(
                id='toepassing',
                options=[
                    {'label': 'Plaat', 'value': 'PLA'},
                    {'label': 'Paneel', 'value': 'PAN'}
                ],
                value='PAN', clearable=False
            ),
        ]),
        dbc.FormGroup([
            dbc.Label('Dakhelling'),
            dcc.Slider(
                id='dakhelling',
                min=5,
                max=90,
                value=5,
            )
        ])
    ])
)

table_header = [
    html.Thead(html.Tr([html.Th("Naam"), html.Th("Waarde")]))
]
table_body = [html.Tbody([
    html.Tr([html.Td("Raillengte"), html.Td(3000, id='raillengte')]),
    html.Tr([html.Td("Rol"), html.Td("1140 x 10000")]),
    html.Tr([html.Td("Eindklem"), html.Td(40, id="eindklem")]),
    html.Tr([html.Td("Tussenklem"), html.Td(22, id="tussenklem")]),
    html.Tr([html.Td("Anker plaatsen om de"), html.Td(800, id="anker_plaatsen_om_de")]),
    html.Tr([html.Td("Benodigde overlap"), html.Td(50, id="benodigde_overlap")])
])]

constants_tab = dbc.Card(
    dbc.CardBody([
        dbc.Table(table_header + table_body, bordered=True, hover=True, responsive=True,
                  striped=True)
    ])
)

table_header = [
    html.Thead(html.Tr([html.Th("Naam"), html.Th("Waarde")]))
]
table_body = [html.Tbody([
    html.Tr([html.Td("Lengthe Rail"), html.Td(id='lengte_rail')]),
    html.Tr([html.Td("Aantal rijen rails"), html.Td(id="aantal_rijen_rails")]),
    html.Tr([html.Td("Totale lengte rails"), html.Td(id="totale_lengte_rails")]),
    html.Tr([html.Td("Aantal rails van 3 meter per rij"), html.Td(id="aantal_rails_van_3_meter_per_rij")]),
    html.Tr([html.Td("Lengte 1 rol"), html.Td(id="lengte_1_rol")]),
    html.Tr([html.Td("Breedte pv"), html.Td(id="breedte_pv")]),
    html.Tr([html.Td("Aantal rijen rollen"), html.Td(id="aantal_rijen_rollen")]),
    html.Tr([html.Td("Dakgoten"), html.Td(id="dakgoten")]),
    html.Tr([html.Td("Schuimstrook driehoek profiel"), html.Td(id="schuimstrook_driehoek_profiel")]),
    html.Tr([html.Td("Railverbinder"), html.Td(id="railverbinder")]),
    html.Tr([html.Td("Aantal ankers op 1 rail"), html.Td(id="aantal_ankers_op_1_rail")]),
    html.Tr([html.Td("Ankers"), html.Td(id="ankers")]),
    html.Tr([html.Td("Schroeven voor beugels"), html.Td(id="schroeven_voor_beugels")]),
    html.Tr([html.Td("Eindklemmen"), html.Td(id="eindklemmen")]),
    html.Tr([html.Td("Middenklemmen"), html.Td(id="middenklemmen")]),
    html.Tr([html.Td("Haak"), html.Td(id="haak")]),
    html.Tr([html.Td("Schroeven voor hoek"), html.Td(id="schroeven_voor_hoek")]),
    html.Tr([html.Td("Schroeven voor ankers"), html.Td(id="schroeven_voor_ankers")]),
    html.Tr([html.Td("Totaal aantal rails van 3m"), html.Td(id="totaal_aantal_rails_van_3m")]),
])]

results_tab = dbc.Card(
    dbc.CardBody([
        dbc.Table(table_header + table_body, bordered=True, hover=True, responsive=True,
                  striped=True),
        html.Div(id='data', style={'display': 'none'})
    ])
)

download_tab = dbc.Card(
    dbc.CardBody([
        dbc.FormGroup([
            dbc.Label("Referentie nr.", html_for='referentie_nr'),
            dbc.Input(id='referentie_nr', type='text'),
        ]),
        dbc.FormGroup([
            dbc.Label("Relatie", html_for='relatie'),
            dbc.Input(id='relatie', type='text'),
        ]),
        dbc.FormGroup([
            dbc.Label("Contactpersoon", html_for='contactpersoon'),
            dbc.Input(id='contactpersoon', type='text'),
        ]),
        dbc.FormGroup([
            dbc.Label("Project", html_for='project'),
            dbc.Input(id='project', type='text'),
        ]),
        dbc.FormGroup([
            dbc.Label("Partijen", html_for='partijen'),
            dbc.Input(id='partijen', type='text'),
        ]),
        dbc.FormGroup([
            dbc.Label("Adviseur", html_for='adviseur'),
            dbc.Input(id='adviseur', type='text'),
        ]),
        html.A(
            id='download-link', children='Download Advies',
            className='btn btn-primary'
        )
    ])
)

app.layout = dbc.Tabs(
    [
        dbc.Tab(input_tab, label="Invoer"),
        dbc.Tab(constants_tab, label="Constanten"),
        dbc.Tab(results_tab, label="Resultaten"),
        dbc.Tab(
            dbc.Col(
                html.Div(id='table', className="pt-3"),
                width={'size': 10, 'offset': 1}
            ), label="Leverlijst"
        ),
        dbc.Tab(html.Div(id='square'), label="Visual"),
        dbc.Tab(download_tab, label="Download Advies")
    ]
)


@app.callback(
    Output('download-link', 'href'),
    Input('referentie_nr', 'value'),
    Input('relatie', 'value'),
    Input('data', 'children')
)
def update_href(referentie_nr, relatie, json_data):
    template = "Solar template 2021.docx"
    document = MailMerge(template)
    #print(document.get_merge_fields())
    document.merge(
        Relatie=relatie, Referentienummer=referentie_nr
    )
    if json_data is not None:
        data = json.loads(json_data)
        document.merge_rows('Aantal', data)
    filename = f"downloads/advies.docx"
    document.write(filename)
    return '/{}'.format(filename)


@app.server.route('/downloads/<path:path>')
def serve_static(path):
    root_dir = os.getcwd()
    return flask.send_from_directory(
        os.path.join(root_dir, 'downloads'), path
    )


@app.callback(
    Output(component_id='lengte_rail', component_property='children'),
    Input(component_id='indeling', component_property='value'),
    Input(component_id='paneelbreedte', component_property='value'),
    Input(component_id='rijen', component_property='value'),
    Input(component_id='kolommen', component_property='value'),
    Input(component_id='tussenklem', component_property='children'),
    Input(component_id='eindklem', component_property='children'),
)
def update_output_div(indeling, paneelbreedte, rijen, kolommen, tussenklem, eindklem):
    if indeling == 'LND':
        result = rijen * paneelbreedte + ((rijen-1) * tussenklem) + 2 * eindklem + 50
    else:
        result = kolommen * paneelbreedte + ((kolommen - 1) * tussenklem) + 2 * eindklem + 50
    return result


@app.callback(
    Output(component_id='aantal_rijen_rails', component_property='children'),
    Input(component_id='indeling', component_property='value'),
    Input(component_id='rijen', component_property='value'),
    Input(component_id='kolommen', component_property='value'),
)
def update_output_div(indeling, rijen, kolommen):
    if indeling == 'LND':
        result = 2 * kolommen
    else:
        result = 2 * rijen
    return result


@app.callback(
    Output(component_id='totale_lengte_rails', component_property='children'),
    Input(component_id='aantal_rijen_rails', component_property='children'),
    Input(component_id='lengte_rail', component_property='children'),
)
def update_output_div(aantal_rijen_rails, lengte_rail):
    result = aantal_rijen_rails * lengte_rail
    return result


@app.callback(
    Output(component_id='aantal_rails_van_3_meter_per_rij', component_property='children'),
    Input(component_id='lengte_rail', component_property='children'),
    Input(component_id='raillengte', component_property='children'),
)
def update_output_div(lengte_rail, raillengte):
    result = math.ceil(lengte_rail / raillengte)
    return result


@app.callback(
    Output(component_id='lengte_1_rol', component_property='children'),
    Input(component_id='kolommen', component_property='value'),
    Input(component_id='paneellengte', component_property='value'),
    Input(component_id='tussenklem', component_property='children'),
    Input(component_id='eindklem', component_property='children'),
    Input(component_id='lengte_rail', component_property='children'),
    Input(component_id='indeling', component_property='value')
)
def update_output_div(kolommen, paneellengte, tussenklem, eindklem, lengte_rail, indeling):
    if indeling == 'LND':
        result = ((kolommen * paneellengte + ((kolommen - 1) * tussenklem) + 2 * eindklem) - 100)
    else:
        result = lengte_rail - 100
    return result


@app.callback(
    Output(component_id='breedte_pv', component_property='children'),
    Input(component_id='rijen', component_property='value'),
    Input(component_id='paneelbreedte', component_property='value'),
    Input(component_id='paneellengte', component_property='value'),
    Input(component_id='tussenklem', component_property='children'),
    Input(component_id='eindklem', component_property='children'),
    Input(component_id='indeling', component_property='value')
)
def update_output_div(rijen, paneelbreedte, paneellengte, tussenklem, eindklem, indeling):
    if indeling == 'LND':
        result = (rijen * paneelbreedte + ((rijen - 1) * eindklem) + 2 * eindklem)
    else:
        result = (rijen * paneellengte + ((rijen - 1) * tussenklem) + 2 * eindklem)
    return result


@app.callback(
    Output(component_id='aantal_rijen_rollen', component_property='children'),
    Input(component_id='breedte_pv', component_property='children'),
    Input(component_id='tussenklem', component_property='children'),
)
def update_output_div(breedte_pv, tussenklem):
    result = math.ceil(1 + (breedte_pv + (tussenklem * 10) - 1140) / 940)
    return result


@app.callback(
    Output(component_id='dakgoten', component_property='children'),
    Input(component_id='aantal_rijen_rollen', component_property='children'),
)
def update_output_div(aantal_rijen_rollen):
    result = aantal_rijen_rollen * 2
    return result


@app.callback(
    Output(component_id='schuimstrook_driehoek_profiel', component_property='children'),
    Input(component_id='tussenklem', component_property='children'),
    Input(component_id='lengte_1_rol', component_property='children'),
    Input(component_id='eindklem', component_property='children'),
)
def update_output_div(tussenklem, lengte_1_rol, eindklem):
    result = math.ceil((((lengte_1_rol + (tussenklem * 10)) * 2) + lengte_1_rol + (eindklem * 10)) / 1280)
    return result


@app.callback(
    Output(component_id='railverbinder', component_property='children'),
    Input(component_id='aantal_rails_van_3_meter_per_rij', component_property='children'),
    Input(component_id='aantal_rijen_rails', component_property='children'),
)
def update_output_div(aantal_rails_van_3_meter_per_rij, aantal_rijen_rails):
    if aantal_rails_van_3_meter_per_rij == 1:
        result = ((aantal_rails_van_3_meter_per_rij - 1) * aantal_rijen_rails)
    else:
        result = ((aantal_rails_van_3_meter_per_rij - 1) * aantal_rijen_rails) + 2
    return result


@app.callback(
    Output(component_id='aantal_ankers_op_1_rail', component_property='children'),
    Input(component_id='lengte_rail', component_property='children'),
    Input(component_id='tussenklem', component_property='children'),
    Input(component_id='anker_plaatsen_om_de', component_property='children')
)
def update_output_div(lengte_rail, tussenklem, anker_plaatsen_om_de):
    result = math.ceil(1 + ((lengte_rail - (20 * tussenklem)) / anker_plaatsen_om_de))
    return result


@app.callback(
    Output(component_id='ankers', component_property='children'),
    Input(component_id='aantal_ankers_op_1_rail', component_property='children'),
    Input(component_id='aantal_rijen_rails', component_property='children'),
)
def update_output_div(aantal_ankers_op_1_rail, aantal_rijen_rails):
    result = math.ceil(aantal_ankers_op_1_rail * aantal_rijen_rails)
    return result


@app.callback(
    Output(component_id='schroeven_voor_ankers', component_property='children'),
    Input(component_id='ankers', component_property='children'),
)
def update_output_div(ankers):
    result = math.ceil(ankers)
    return result


@app.callback(
    Output(component_id='schroeven_voor_beugels', component_property='children'),
    Input(component_id='aantal_ankers_op_1_rail', component_property='children'),
    Input(component_id='aantal_rijen_rails', component_property='children'),
)
def update_output_div(aantal_ankers_op_1_rail, aantal_rijen_rails):
    result = math.ceil(aantal_ankers_op_1_rail * aantal_rijen_rails)
    return result


@app.callback(
    Output(component_id='eindklemmen', component_property='children'),
    Input(component_id='aantal_rijen_rails', component_property='children'),
)
def update_output_div(aantal_rijen_rails):
    result = aantal_rijen_rails * 2
    return result


@app.callback(
    Output(component_id='middenklemmen', component_property='children'),
    Input(component_id='aantal_rijen_rails', component_property='children'),
    Input(component_id='rijen', component_property='value'),
    Input(component_id='kolommen', component_property='value'),
    Input(component_id='indeling', component_property='value')
)
def update_output_div(aantal_rijen_rails, rijen, kolommen, indeling):
    if indeling == 'LND':
        result = aantal_rijen_rails * (rijen - 1)
    else:
        result = aantal_rijen_rails * (kolommen - 1)
    return result


@app.callback(
    Output(component_id='haak', component_property='children'),
    Input(component_id='indeling', component_property='value'),
    Input(component_id='kolommen', component_property='value'),
    Input(component_id='rijen', component_property='value'),
)
def update_output_div(indeling, kolommen, rijen):
    if indeling == 'LND':
        result = (kolommen * 2 * (rijen + 1))
    else:
        result = (rijen * 2 * (kolommen + 1))
    return result


@app.callback(
    Output(component_id='schroeven_voor_hoek', component_property='children'),
    Input(component_id='haak', component_property='children')
)
def update_output_div(haak):
    result = math.ceil(haak * 2)
    return result


@app.callback(
    Output(component_id='totaal_aantal_rails_van_3m', component_property='children'),
    Input(component_id='totale_lengte_rails', component_property='children'),
    Input(component_id='raillengte', component_property='children')
)
def update_output_div(totale_lengte_rails, raillengte):
    result = math.ceil(totale_lengte_rails / raillengte)
    return result


@app.callback(
    Output('table', 'children'),
    Output('data', 'children'),
    Input('ankers', 'children'),
    Input('totaal_aantal_rails_van_3m', 'children'),
    Input('dakgoten', 'children'),
    Input('schuimstrook_driehoek_profiel', 'children'),
    Input('railverbinder', 'children'),
    Input('schroeven_voor_beugels', 'children'),
    Input('schroeven_voor_ankers', 'children')
)
def update_datatable(ankers, totaal_aantal_rails_van_3m, dakgoten,
                     schuimstrook_driehoek_profiel, railverbinder, schroeven_voor_beugels,
                     schroeven_voor_ankers):

    df.loc[df['id'] == "0770003", ['count']] = ankers
    df.loc[df['id'] == "0770212", ['count']] = totaal_aantal_rails_van_3m
    df.loc[df['id'] == "0770037", ['count']] = dakgoten
    df.loc[df['id'] == "0340139", ['count']] = schuimstrook_driehoek_profiel
    df.loc[df['id'] == "0703967", ['count']] = railverbinder
    df.loc[df['id'] == "0770500", ['count']] = math.ceil(schroeven_voor_beugels / 4)
    df.loc[df['id'] == "0770501", ['count']] = math.ceil(schroeven_voor_ankers / 30)

    df_result = df.loc[df['count'] > 0]
    df_result = df_result.astype({'count': 'str'})
    df_result.columns = ['Artikelnummer', 'Omschrijving', 'Bruto', 'Aantal']
    data = df_result.to_dict('rows')
    columns = [{"name": i, "id": i, } for i in df_result.columns]
    return dash_table.DataTable(data=data, columns=columns), json.dumps(data)


@app.callback(
    Output('square', 'children'),
    Input('rijen', 'value'),
    Input('kolommen', 'value')
)
def update_square(rijen, kolommen):
    image_filename = 'paneel.png'  # replace with your own image
    encoded_image = base64.b64encode(open(image_filename, 'rb').read())

    imgList = []
    for r in range(rijen):
        rows = []
        for k in range(kolommen):
            rows.append(html.Img(src='data:image/png;base64,{}'.format(encoded_image.decode())))
        imgList.append(html.Div(rows, className="row"))
    return imgList
    #return {'width': '40vw', 'height': '10vw'}


if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8080, debug=True)
