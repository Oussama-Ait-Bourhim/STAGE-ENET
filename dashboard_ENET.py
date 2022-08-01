#################importation des biblio necessaire########################
import dash
from dash import dcc, html
from dash.dependencies import Output, Input
import plotly.express as px
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import plotly.graph_objects as go

############imporation des données ####################v
df=pd.read_excel("data\carnet_adulte_ENET2012.xlsx",engine='openpyxl')


#################changement des valeur des colones d'aprés la nomonclature et nos techniques et changement noms des colones####################
df['Milieu']= df['Milieu'].replace([1,2],['urbain','rural'])
df['Sexe']= df['Sexe'].replace([1,2],['hommes','femmes'])
df['Ramadan']= df['Ramadan'].replace([0,1],['hors ramadan','ramadan'])
df['Groupe_âge']= df['Groupe_âge'].replace([1,2,3,4,5],['15 à 24','25 à 34','35 à 45','46 à 59','60 ans au plus'])
df['Etat_matrimonial']= df['Etat_matrimonial'].replace([1,2,3,4],['celibataire','marié','divorcé','veuf'])
df['Niveau_scolaire']= df['Niveau_scolaire'].replace([0,1,2,3,4,5],['sans niveau','primaire','secondaire collegial','secondaire qualifiant','superieur','autre niveau'])
df['Type_activite']= df['Type_activite'].replace([1,2,3,4,5],['actif occupé','chomeur','femme foyer','etudiant','autres'])
df['Catégorie_socioprofessionnelle']= df['Catégorie_socioprofessionnelle'].replace([1,2,3,4,5,6,9],['directeur','cadre moyen','commerçants','exploitant','artisants','monoeuvres','non declaré'])
df['Statut_professionnel']= df['Statut_professionnel'].replace([1,2,3,9],['salarié','autoemployé','non remineré','non declaré'])
df = df.rename(columns = {"dureeG0":"someil"})
df = df.rename(columns = {"dureeG1":"repas"})
df = df.rename(columns = {"dureeG2":"soins personnel"})
df = df.rename(columns = {"dureeG3":"travail profesionel"})
df = df.rename(columns = {"dureeG4":"formation et education"})
df = df.rename(columns = {"dureeG5":"travaux menagers"})
df = df.rename(columns = {"dureeG6":"soins menages"})
df = df.rename(columns = {"dureeG7":"loisirs"})
df = df.rename(columns = {"dureeG8":"sociabilité"})
df = df.rename(columns = {"dureeG9":"pratique reigieuses"})


#################les fonction creation graphe########################
def pourc():
    o=html.Div([
        html.H4('Partition durée des activités en min/jour selon les variables qualitatives pour les adultes (agés 15 ans et plus)',
                className='text-white'),
        html.P("Variables qualitatives:"),
        dcc.Dropdown(id='names',
                     options=['Milieu', 'Sexe', 'Etat_matrimonial', 'Groupe_âge','Niveau_scolaire', 'Type_activite','Ramadan'],
                     value='Milieu', clearable=False,
                     #style={'width': '50%', 'height': '3vh'},
                     ),
        html.P("les activités:"),
        dcc.Dropdown(id='values',
                     options=['someil', 'repas', 'soins personnel', 'travail profesionel', 'formation et education',
                              'travaux menagers', 'soins menages', 'loisirs', 'sociabilité', 'pratique reigieuses'],
                     value='someil', clearable=False,
                     #style={'width': '50%', 'height': '3vh'}
                     ),
        dcc.Graph(id="g" )#,style = {"width": "auto", "height": "270px"}),
    ])
    return o

def boxplot():
    o=html.Div([
        html.H4("Boxplot de durré activités en min par jour selon les variables qualitatives pour les adultes (agés 15 ans et plus)",
                className='text-white'),
        html.P("x-axis:"),
        dcc.Checklist(
            id='x-axis',
            options=['Milieu', 'Sexe', 'Etat_matrimonial', 'Niveau_scolaire', 'Type_activite','Groupe_âge','Ramadan'],
            value=['Sexe'],
            className='text-white',
            #labelStyle={'display': 'inline-block'}#style={'width': '50%', 'height': '3vh'}
        ),
        html.P("y-axis:"),
        dcc.RadioItems(
            id='y-axis',
            options=['someil', 'repas', 'soins personnel', 'travail profesionel', 'formation et education',
                     'travaux menagers', 'soins menages', 'loisirs', 'sociabilité', 'pratique reigieuses'],
            value='someil',
            className='text-white',
            #labelStyle={'display': 'inline-block'}#style={'width': '50%', 'height': '3vh'}
        ),
        dcc.Graph(id="f"),
        #style={'width': '50%', 'height': '40vh'}),

    ])
    return o

def barchart1():
    dff = pd.read_excel("data/b.xlsx")
    fig = px.bar(dff, x=dff['Activité'], y=dff['Ensemble']).update_layout({
            "plot_bgcolor": "rgba(16, 15, 15, 0.32)",
            "paper_bgcolor": "rgba(16, 15, 15, 0.32)",
        },        font=dict(family="Lato, monospace", size=12, color="#fff"),
        xaxis={'showgrid': False},
        yaxis={'showgrid': True}
    )
    return html.Div(
        [
            html.Div(
                [
                        html.H4('Partitions par section des activités selon temps moyenne en minute par jour pour les adultes (agés 15 ans et plus)',
                                className='text-white'),
                    dcc.Graph(
                        id="example-graph",
                        figure=fig,
                        responsive=True,
                        #style={"width": "auto", "height": "270px"},
                    ),
                ],
                #style={"textAlign": "center"},
                className="mycard",
            )
        ]
    )


def barchart2():
    dff = pd.read_excel("data\g.xlsx")
    fig = px.bar(dff, x=dff['Activité'], y=dff['Ensemble'],color_discrete_sequence =['red']*13).update_layout({
            "plot_bgcolor": "rgba(16, 15, 15, 0.32)",
            "paper_bgcolor": "rgba(16, 15, 15, 0.32)",
        },        font=dict(family="Lato, monospace", size=12, color="#fff"),
        xaxis={'showgrid': False},
        yaxis={'showgrid': True}
    )
    return html.Div(
        [
            html.Div(
                [
                        html.H4('Partitions par section des activités selon temps moyenne en minute par jour pour les enfants ( age -15 ans)',
                                className='text-white'),
                    dcc.Graph(
                        id="example-graphg",
                        figure=fig,
                        responsive=True,
                        #style={"width": "auto", "height": "270px"},
                    ),
                ],
                #style={"textAlign": "center"},
                className="mycard",
            )
        ]
    )

def matrixcorr1():
    colums = ['someil','repas','soins personnel','travail profesionel','formation et education','travaux menagers','soins menages','loisirs','sociabilité','pratique reigieuses']
    dff=df.filter(colums)
    df_corr = dff.corr()
    fig = go.Figure().add_trace(
        go.Heatmap(
            x=df_corr.columns,
            y=df_corr.index,
            z=np.array(df_corr),
            text=df_corr.values,
            texttemplate='%{text:.2f}'
        )
    ).update_layout({
            "plot_bgcolor": "rgba(16, 15, 15, 0.32)",
            "paper_bgcolor": "rgba(16, 15, 15, 0.32)",
        },        font=dict(family="Lato, monospace", size=12, color="#fff"),
        xaxis={'showgrid': False},
        yaxis={'showgrid': True}
    )
    return html.Div(
        [
            html.Div(
                [
                    html.H4('Matrice correlation entre les differents activités pour les adultes (agés 15 ans et plus)',
                            className='text-white'),
                    dcc.Graph(
                        id="example-graphv",
                        figure=fig,
                        responsive=True,
                        #style={"width": "auto", "height": "270px"},
                    ),
                ],
                #style={"textAlign": "center"},
                className="mycard",
            )
        ]
    )

def matrixcorr2():
    d=pd.read_excel("data\carnet_enfant_ENET2012.xlsx")
    d = d.rename(columns={"dureeG0": "someil"})
    d = d.rename(columns={"dureeG1": "repas"})
    d = d.rename(columns={"dureeG2": "soins personnel"})
    d = d.rename(columns={"dureeG3": "travail profesionel"})
    d = d.rename(columns={"dureeG4": "formation et education"})
    d = d.rename(columns={"dureeG5": "travaux menagers"})
    d = d.rename(columns={"dureeG6": "soins menages"})
    d = d.rename(columns={"dureeG7": "loisirs"})
    d = d.rename(columns={"dureeG8": "sociabilité"})
    d = d.rename(columns={"dureeG9": "pratique reigieuses"})
    colums = ['someil','repas','soins personnel','travail profesionel','formation et education','travaux menagers','soins menages','loisirs','sociabilité','pratique reigieuses']
    dff=d.filter(colums)
    df_corr = dff.corr()
    fig = go.Figure().add_trace(
        go.Heatmap(
            x=df_corr.columns,
            y=df_corr.index,
            z=np.array(df_corr),
            text=df_corr.values,
            texttemplate='%{text:.2f}'
        )
    ).update_layout({
            "plot_bgcolor": "rgba(16, 15, 15, 0.32)",
            "paper_bgcolor": "rgba(16, 15, 15, 0.32)",
        },        font=dict(family="Lato, monospace", size=12, color="#fff"),
        xaxis={'showgrid': False},
        yaxis={'showgrid': True}
    )
    return html.Div(
        [
            html.Div(
                [
                    html.H4('Matrice correlation entre les differents activités pour les enfants (age -15 ans)',
                            className='text-white'),
                    dcc.Graph(
                        id="example-graphvn",
                        figure=fig,
                        responsive=True,
                        #style={"width": "auto", "height": "270px"},
                    ),
                ],
                #style={"textAlign": "center"},
                className="mycard",
            )
        ]
    )







# https://stooq.com/



# df.to_csv("mystocks.csv", index=False)
# df = pd.read_csv("mystocks.csv")
# print(df[:15])


# https://www.bootstrapcdn.com/bootswatch/
#app = dash.Dash(__name__, external_stylesheets=[dbc.themes.QUARTZ],
                #meta_tags=[{'name': 'viewport',
                       #'content': 'width=device-width, initial-scale=1.0'}]
                #)


# Layout section: Bootstrap (https://hackerthemes.com/bootstrap-cheatsheet/)
##################### creation  website dashboard ########################
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.SLATE])
app.layout = dbc.Container([

    dbc.Row(
        dbc.Col(html.H1("Dashboard Enquête Nationale sur l'Emploi du Temps - ENET - 2012 ",
                        className='bg-danger text-white text-sm-center'),
                        #className='text-center text-primary mb-4'),
                width=12)
    ),

    dbc.Row([

        dbc.Col([
            pourc()

        ],# width={'size':5, 'offset':1, 'order':1},
           xs=12, sm=12, md=6, lg=6, xl=6
        ),

        dbc.Col([
            boxplot()



        ], #width={'size':5, 'offset':0, 'order':2},
           xs=12, sm=12, md=6, lg=6, xl=6
        ),

    ], justify='start'),  # Horizontal:start,center,end,between,around

    dbc.Row([
        dbc.Col([
        barchart1()
        ], #width={'size':5, 'offset':1},
           xs=12, sm=12, md=6, lg=6, xl=6
        ),

        dbc.Col([
             barchart2()

        ], #width={'size':5, 'offset':1},
           xs=12, sm=12, md=6, lg=6, xl=6
        )
    ], ),#align="center")  # Vertical: start, center, end

    dbc.Row([

        dbc.Col([
            matrixcorr1()
        ],  # width={'size':5, 'offset':1, 'order':1},
            xs=12, sm=12, md=6, lg=6, xl=6
        ),

        dbc.Col([
            matrixcorr2()

        ],  # width={'size':5, 'offset':0, 'order':2},
            xs=12, sm=12, md=6, lg=6, xl=6
        ),

    ], justify='start')

], fluid=True)


######################## Callback section: connecting the components########################
# ************************************************************************
# Line chart - Single
# Line chart - multiple
@app.callback(
    Output("g", "figure"),
    Input("names", "value"),
    Input("values", "value"))


def generate_chart1(names, values):
    fig1 = px.pie(df, values=values, names=names, hole=.3).update_layout({
            "paper_bgcolor": "rgba(16, 15, 15, 0.32)",
        },
        font=dict(family="Lato, monospace", size=12, color="#fff"),
        xaxis={'showgrid': False},
        yaxis={'showgrid': True}
    )
    return fig1

@app.callback(
    Output("f", "figure"),
    Input("x-axis", "value"),
    Input("y-axis", "value"))
def o(a,b):
    fig2 = px.box(df, x=a, y=b).update_layout({
            "plot_bgcolor": "rgba(16, 15, 15, 0.32)",
            "paper_bgcolor": "rgba(16, 15, 15, 0.32)",
        },
        font=dict(family="Lato, monospace", size=12, color="#fff"),
        xaxis={'showgrid': False},
        yaxis={'showgrid': True}
    )
    return fig2

app.run_server()
