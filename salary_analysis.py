from dash import Dash, html, dash_table, dcc, callback, Output, Input
import dash_bootstrap_components as dbc
from threading import Timer
import webbrowser
import pymupdf
import pandas as pd
import plotly.express as px
import os
import datetime

# Initializing the app
app = Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])


def load_pdf(pdf_name):
    
    doc = pymupdf.open(pdf_name) 
    text_divided = []
    for page in doc:
        content = page.get_text("dict")             # Stacked dictionary object
        for block in content["blocks"]:             # Blocks are the top hierarchy
            if block["type"] == 0 :                 # If text in block and no image
                record_line = []
                for line in block["lines"]:         # A text block consists of lines (single words)
                    if line["dir"] == (1, 0):       # Keep only horizontal text
                        for span in line["spans"]:  # span = text with identical properties
                            record_line.append(span["text"])
                text_divided.append(record_line)
        break                                       # We only want page 1 for now


    # Get the periode of time
    for i in text_divided:
        if 'Période du' in i :
            date_full = i[2].split('/')
            #date_light = date_full[1] + ' ' + date_full[2]
            date_light = datetime.datetime(int(date_full[2]), int(date_full[1]), 1)
            break
    
    # Transform text to float
    def num_str_to_float(number_string):
        return float(number_string.replace('.', '').replace(',', '.'))
    
    # Keep only lines with amount at the end
    lines_amount = []
    for line in text_divided:
        try :
            line[-1] = num_str_to_float(line[-1])
            lines_amount.append(line)
        except:
            continue
    
    # arrays to form the data frame
    array_description = []
    array_amount = []
    for line_amount in lines_amount:

        if 'montant brut' in line_amount :
            array_description.append(['0', 'Total brut', 'Brut_total'])
            array_amount.append(line_amount[-1]) 
        elif 'imposable' in line_amount :
            array_description.append(['1', 'Total imposable', 'Imposable_total'])
            array_amount.append(line_amount[-1]) 
        elif 'salaire net' in line_amount :
            array_description.append(['2', 'Total net', 'Net_total'])
            array_amount.append(line_amount[-1]) 
        elif len(line_amount[0]) == 4 :
            array_description.append([line_amount[0], line_amount[1], 'N.A.'])
            array_amount.append(line_amount[-1])

    # Create empty array of correct length
    array_description_code = [0] * len(array_description)
    for i in range(len(array_description)):
        array_description_code[i] = array_description[i][0]
    data_frame = pd.DataFrame(data={date_light:array_amount}, index=array_description_code)
    data_frame.index.name = 'code'
    return(array_description, data_frame)

def load_excel(description, data):

    try :
        data_excel = pd.read_excel('perdiems.xlsx', usecols='A,H,I', dtype={'Date':str,'Perdiem total':float,'Spent':float})
    except:
        print("No perdiems Excel file found.")
        return(description, data)
    
    # Modifiy excel date presentation to fit pandas dataframe
    list_date = [datetime.datetime(int(date.split(' ')[0].split('-')[0]), int(date.split(' ')[0].split('-')[1]), 1) for date in data_excel['Date'].tolist()]
    list_perdiem = data_excel['Perdiem total'].tolist()
    list_spent = data_excel['Spent'].tolist()    

    data_excel_clean = pd.DataFrame(data={list_date[0]:[list_perdiem[0], list_spent[0]]}, index=['3', '4'])
    for i in range(1, len(list_date)) :
        df2 = pd.DataFrame(data={list_date[i]:[list_perdiem[i], list_spent[i]]}, index=['3', '4'])
        data_excel_clean = pd.concat([data_excel_clean, df2], axis=1)
    data_excel_clean.index.name = 'code'

    # Sum duplicates (multiple missions with the same date)
    data_excel_clean = data_excel_clean.T.groupby(data_excel_clean.columns).sum().T
    data = pd.concat([data, data_excel_clean]).fillna(0)

    # Add final net salary with perdiems and missions expenses
    data.loc['5',:] = data.loc[['3', '4']].sum(axis=0)
    data.loc['6',:] = data.loc[['2', '3', '4']].sum(axis=0)

    description.append(['3', 'Mission perdiem', 'Net'])
    description.append(['4', 'Mission spent', 'Net'])
    description.append(['5', 'Net mission', 'Net'])
    description.append(['6', 'Net with mission', 'Net'])

    return(description, data)

def process_folder(dir_name):

    # Move working directory down
    os.chdir(dir_name)
    pdf_names = os.listdir()

    data_full = []
    description_full = []
    for i in pdf_names :
        try :
            description, data = load_pdf(i)
            description_full = description_full + description
            data_full.append(data)
            print(i)
        except: 
            print("Error loading file : " + i + " -> skipping file")

    # Move working directory back up
    os.chdir('..')

    # Create description array
    description_clean = []
    for description in description_full:
        if description not in description_clean:
            description_clean.append(description)

    # Mix dataframes of the same month and rename duplicates
    for i in range(len(data_full)) : 
        for j in range(i+1, len(data_full)) :
            if data_full[i].columns[-1] == data_full[j].columns[-1] and data_full[i].columns[-1] != '0':
                data_full[i] = pd.concat([data_full[i], data_full[j]]).groupby('code').sum()
                data_full[j] = data_full[j].rename(columns={data_full[j].columns[-1]: '0'})

    # Creat final dataframe
    data_clean = data_full[0]
    for i in range(1, len(data_full)) :
        if data_full[i].columns[-1] != '0':
            data_clean = pd.merge(data_clean, data_full[i], on='code', how='outer')       
    data_clean = data_clean.fillna(0)

    return(description_clean, data_clean)

def extract_data():
    # Extracting salary data
    print()
    print("Searching for sdworks folder...")

    folders_sdworks = []
    folders_names = os.listdir()
    for i in folders_names:
        if "sdworks" in i :
            folders_sdworks.append(i)

    if len(folders_sdworks) == 1:
        print("Folder found : " + folders_sdworks[0])
        folder_to_process = folders_sdworks[0]
    else :
        print("Multiple sdworks folders found. Please select one.")
        for i in range(len(folders_sdworks)):
            print(" " + str(i) + " - " + folders_sdworks[i])
        print()
        try :
            i = int(input())
            folder_to_process = folders_sdworks[i]
        except : 
            print("Wrong value...")
            exit()
        

    
    print()
    print('Extracting salary data...')
    print()
    description, data = process_folder(folder_to_process)

    # Extracting perdiems and mission spent data
    print()
    print("Loading perdiem data...")
    print()
    description, data = load_excel(description, data)

    # Create legend from description array
    codes = []
    descr = []
    for i in range(len(description)):
        codes.append(description[i][0])
        descr.append(description[i][1])

    print()
    print('---- Full data ----')
    print(data)

    print()
    print('---- Description ----')
    for i in description:
        print(i)

    return codes, descr, data

def open_browser():
    webbrowser.open_new("http://127.0.0.1:8050/")

def main():
    
    codes, descr, data = extract_data()

    # App layout
    app.layout = [
        html.H1("Salary data"),
        html.Hr(),
        dbc.Row(
            [
                dbc.Col([html.H2("Select code to display"),
                         dcc.Dropdown(
                            options=[{'label': k, 'value': v} for k, v in zip(descr, codes)],        
                            value=['6'],
                            multi=True,
                            searchable=True,
                            maxHeight=400,
                            id='code-selection'),
                        html.H2("Average"),
                        dcc.RangeSlider(0, len(data.columns)-1, 1, value=[0, len(data.columns)-1], 
                                        marks=None,
                                        tooltip={"placement": "bottom",
                                                 "always_visible": True}, 
                                        allowCross=False,
                                        id='slider-average')
                        ], style = {'margin-left':'10px', 'margin-top':'7px', 'margin-right':'10px'}),
                dbc.Col(dcc.Graph(figure={}, id='controls-and-graph'), width = 9, style = {'margin-left':'5px', 'margin-top':'7px', 'margin-right':'5px'})
        ])
    ]

    # callback for sdworx code selection
    @callback(
        Output(component_id='controls-and-graph', component_property='figure'),
        Input(component_id='code-selection', component_property='value'),
        Input(component_id='slider-average', component_property='value'),
    )
    def update_graph_code(codes_selected, dates_average):

        data_graph = data.loc[codes_selected].transpose()

        # manipulation to preset index datatype as datetime
        data_graph['dates'] = pd.to_datetime(data_graph.index)
        data_graph = data_graph.set_index('dates')
        data_mean = data_graph.iloc[dates_average[0]:dates_average[1]+1].mean(axis=0)
        
        fig = px.bar(data_graph)

        x0_average = data_graph.index[dates_average[0]]
        x1_average = data_graph.index[dates_average[1]]
        fig.add_vrect(x0=x0_average, x1=x1_average, 
                      annotation_text="average", annotation_position="top left",
                      fillcolor="grey", opacity=0.2, line_width=0)
        
        for code in codes_selected:
            y_value = data_mean[code]
            text_mean = str(code) + " : " + str("{:.2f}".format(y_value) + "€")
            fig.add_hline(y=y_value, line_dash="dot", annotation_text=text_mean, annotation_position="top left")            
    
        fig.update_layout(yaxis_title="Amount [€]")
        return fig
    

if __name__ == "__main__":
    main()
    Timer(1, open_browser).start()
    app.run(debug=True, use_reloader=True)
    