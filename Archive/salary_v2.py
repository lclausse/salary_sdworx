import matplotlib.pyplot as plt
from pypdf import PdfReader 
import pandas as pd
import os



def load_pdf(pdf_name):
    
    reader = PdfReader(pdf_name) 
    text_raw = ""
    for i in range(len(reader.pages)) :
        text_raw += reader.pages[i].extract_text()
    text_divided = text_raw.split('\n')

    def num_str_to_float(number_string):
        return float(number_string.replace('.', '').replace(',', '.'))


    # get the periode of time
    for i in text_divided:   
        if 'Périodedu' in i :
            date_full = i.split('Périodedu')[1].split('au')[0].split('/')
            date_light = date_full[1] + ' ' + date_full[2]

    # arrays to form the data frame
    array_description = []
    array_amount = []

    # boolean activators
    record_brut =      0
    record_imposable = 0
    record_net =       0
    for i in text_divided:

        if 'Salairemensueldebase:€' in i :
            #array_Type.append('Base')
            #array_code.append(0)
            #array_description.append('Salaire de base')
            #array_amount.append(num_str_to_float(i.split('Salairemensueldebase:€')[1])) 
            salaire = 0 #useless
        elif 'CodeDescription JoursHeures Montants' in i :
            record_brut = 1
        elif 'montantbrut ' in i :
            array_description.append(['0', 'Total brut', 'Brut_total'])
            array_amount.append(num_str_to_float(i.split('montantbrut ')[1])) 
            record_brut = 0
            record_imposable = 1
        elif 'imposable ' in i :
            array_description.append(['1', 'Total imposable', 'Imposable_total'])
            array_amount.append(num_str_to_float(i.split('imposable ')[1])) 
            record_imposable = 0
            record_net = 1
        elif 'salairenet ' in i :
            array_description.append(['2', 'Total net', 'Net_total'])
            array_amount.append(num_str_to_float(i.split('salairenet ')[1]))  
            record_net = 0

        elif record_brut == 1:
            array_description.append([i.split(' ')[0][0:4], i.split(' ')[0][4:], 'Brut'])
            array_amount.append(num_str_to_float(i.split(' ')[-1]))
        elif record_imposable == 1:
            array_description.append([i.split(' ')[0][0:4], i.split(' ')[0][4:], 'Imposable'])
            array_amount.append(num_str_to_float(i.split(' ')[-1]))
        elif record_net == 1:
            array_description.append([i.split(' ')[0][0:4], i.split(' ')[0][4:], 'Net'])
            array_amount.append(num_str_to_float(i.split(' ')[-1]))
    
    # Create empty array of correct length
    array_description_code = [0] * len(array_description)
    for i in range(len(array_description)):
        array_description_code[i] = array_description[i][0]
    data_frame = pd.DataFrame(data={date_light:array_amount}, index=array_description_code)
    data_frame.index.name = 'code'
    return(array_description, data_frame)

def load_excel(description, data):

    data_excel = pd.read_excel('perdiems.xlsx', usecols='A,H,I', dtype={'Date':str,'Perdiem total':float,'Spent':float})

    # Modifiy excel date presentation to fit pandas dataframe
    list_date = data_excel['Date'].tolist()
    list_date = [date.split(' ')[0].split('-')[1] + ' ' + date.split(' ')[0].split('-')[0] for date in list_date]

    list_perdiem = data_excel['Perdiem total'].tolist()
    list_spent = data_excel['Spent'].tolist()
    #('3', 'Mission perdiem', 'Net'), ('4', 'Mission spent', 'Net')
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

def sum_between(df, date_1, date_2, code):
    index_1 = df.columns.get_loc(date_1)
    index_2 = df.columns.get_loc(date_2)
    if index_2 < index_1:
        print('Pas top...')
    return(df.loc[[code]].iloc[:, index_1:index_2+1].sum(axis=1))

def mean_between(df, date_1, date_2, code):
    index_1 = df.columns.get_loc(date_1)
    index_2 = df.columns.get_loc(date_2)
    if index_2 < index_1:
        print('Pas top...')
    return(df.loc[[code]].iloc[:, index_1:index_2+1].mean(axis=1))

def plot_side(description, df, date_1, date_2, codes):
    # Check if dates are coherent
    index_1 = df.columns.get_loc(date_1)
    index_2 = df.columns.get_loc(date_2)
    if index_2 < index_1:
        print('Pas top...')

    data = df.loc[codes].iloc[:, index_1:index_2+1].T
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4), gridspec_kw={'width_ratios': [4, 1]})
    data.plot.bar(ax=ax1)

    if data.max().max() < 500 :
        ax1.yaxis.set_major_locator(plt.MultipleLocator(50))
    else :
        ax1.yaxis.set_major_locator(plt.MultipleLocator(500))
    ax1.grid(axis='y')

    # Create legend from description array
    legend = []
    for i in codes:
        for j in range(len(description)-1):
            if description[j][0] == i:
                legend.append(description[j][1])
    ax1.legend(legend)

    data.plot.box(ax=ax2)
    ax2.grid(axis='y')

def plot_stacked(description, df, date_1, date_2, codes):
    # Check if dates are coherent
    index_1 = df.columns.get_loc(date_1)
    index_2 = df.columns.get_loc(date_2)
    if index_2 < index_1:
        print('Pas top...')

    data = df.loc[codes].iloc[:, index_1:index_2+1].T
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4), gridspec_kw={'width_ratios': [4, 1]})
    data.plot.bar(ax=ax1, stacked=True)
    
    if data.max().max() < 500 :
        ax1.yaxis.set_major_locator(plt.MultipleLocator(50))
    else :
        ax1.yaxis.set_major_locator(plt.MultipleLocator(500))
    ax1.grid(axis='y')

    # Create legend from description array
    legend = []
    for i in codes:
        for j in range(len(description)-1):
            if description[j][0] == i:
                legend.append(description[j][1])
    ax1.legend(legend)

    data.plot.box(ax=ax2)
    ax2.grid(axis='y')
    
    
    

def _main():
    # Extracting salary data
    print()
    print('Extracting salary data...')
    print()
    description, data = process_folder('sdworks_LCL')

    # Extracting perdiems and mission spent data
    print()
    print("Loading perdiem data...")
    print()
    description, data = load_excel(description, data)

    date1 = '01 2023'
    date2 = '10 2024'
    code = '2'
    """
    2 -> Net salary
    3 -> Per diems
    4 -> Spent on site
    5 -> Net mission (per diems - spent)
    6 -> Net salary + perdiems
    """

    print()
    print('---- Full data ----')
    print(data)

    print()
    print('---- Description ----')
    for i in description:
        print(i)

    print()
    print('---- Sum ----')
    print(sum_between(data, date1, date2, code))

    print()
    print('---- Mean ----')
    print(mean_between(data, date1, date2, code))

    print()
    print('---- Plot of stacked data ----')
    print(plot_stacked(description, data, date1, date2, ['2', '5']))

    print()
    print('---- Plot of side by side data ----')
    print(plot_side(description, data, date1, date2, ['1', '2']))

    print()
    print('---- Plot of side by side data ----')
    print(plot_side(description, data, date1, date2, ['1701', '1702']))

    plt.show()

_main()