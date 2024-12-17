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
    array_Description = []
    array_Amount = []

    # boolean activators
    record_brut =      0
    record_imposable = 0
    record_net =       0
    for i in text_divided:

        if 'Salairemensueldebase:€' in i :
            #array_Type.append('Base')
            #array_code.append(0)
            #array_Description.append('Salaire de base')
            #array_Amount.append(num_str_to_float(i.split('Salairemensueldebase:€')[1])) 
            salaire = 0 #useless
        elif 'CodeDescription JoursHeures Montants' in i :
            record_brut = 1
        elif 'montantbrut ' in i :
            array_Description.append(('0', 'Total brut', 'Brut_total'))
            array_Amount.append(num_str_to_float(i.split('montantbrut ')[1])) 
            record_brut = 0
            record_imposable = 1
        elif 'imposable ' in i :
            array_Description.append(('1', 'Total imposable', 'Imposable_total'))
            array_Amount.append(num_str_to_float(i.split('imposable ')[1])) 
            record_imposable = 0
            record_net = 1
        elif 'salairenet ' in i :
            array_Description.append(('2', 'Total net', 'Net_total'))
            array_Amount.append(num_str_to_float(i.split('salairenet ')[1]))  
            record_net = 0

        elif record_brut == 1:
            array_Description.append((i.split(' ')[0][0:4], i.split(' ')[0][4:], 'Brut'))
            array_Amount.append(num_str_to_float(i.split(' ')[-1]))
        elif record_imposable == 1:
            array_Description.append((i.split(' ')[0][0:4], i.split(' ')[0][4:], 'Imposable'))
            array_Amount.append(num_str_to_float(i.split(' ')[-1]))
        elif record_net == 1:
            array_Description.append((i.split(' ')[0][0:4], i.split(' ')[0][4:], 'Net'))
            array_Amount.append(num_str_to_float(i.split(' ')[-1]))
    
    return(pd.DataFrame(data={date_light:array_Amount}, index=pd.MultiIndex.from_tuples(array_Description, names=('Code', 'Description', 'Type'))))

def load_excel(data):

    data_excel = pd.read_excel('perdiems.xlsx', usecols='A,H,I', dtype={'Date':str,'Perdiem total':float,'Spent':float})

    # Modifiy excel date presentation to fit pandas dataframe
    list_date = data_excel['Date'].tolist()
    list_date = [date.split(' ')[0].split('-')[1] + ' ' + date.split(' ')[0].split('-')[0] for date in list_date]

    list_perdiem = data_excel['Perdiem total'].tolist()
    list_spent = data_excel['Spent'].tolist()

    data_excel_clean = pd.DataFrame(data={list_date[0]:[list_perdiem[0], list_spent[0]]}, index=pd.MultiIndex.from_tuples([('3', 'Mission perdiem', 'Net'), ('4', 'Mission spent', 'Net')], names=('Code', 'Description', 'Type')))
    for i in range(1, len(list_date)) :
        df2 = pd.DataFrame(data={list_date[i]:[list_perdiem[i], list_spent[i]]}, index=pd.MultiIndex.from_tuples([('3', 'Mission perdiem', 'Net'), ('4', 'Mission spent', 'Net')], names=('Code', 'Description', 'Type')))
        data_excel_clean = pd.concat([data_excel_clean, df2], axis=1)

    # Sum duplicates (multiple missions with the same date)
    data_excel_clean = data_excel_clean.T.groupby(data_excel_clean.columns).sum().T
    data = pd.concat([data, data_excel_clean]).fillna(0)

    # Add final net salary with perdiems and missions expenses
    data.loc[('5', 'Net with mission', 'Net'),:] = data.loc[['2', '3', '4']].sum(axis=0)

    return(data)

def process_folder(dir_name):

    # Move working directory down
    os.chdir(dir_name)
    pdf_names = os.listdir()

    data_full = []
    for i in pdf_names :
        try :
            data_full.append(load_pdf(i))
            print(i)
        except: 
            print("Error loading file : " + i + " -> skipping file")
    
    # Move working directory back up
    os.chdir('..')

    # Mix dataframes of the same month and rename duplicates
    for i in range(len(data_full)) : 
        for j in range(i+1, len(data_full)) :
            if data_full[i].columns[-1] == data_full[j].columns[-1] and data_full[i].columns[-1] != '0':
                data_full[i] = pd.concat([data_full[i], data_full[j]]).groupby(['Code', 'Description', 'Type'])[data_full[i].columns[-1]].sum().reset_index()
                data_full[j] = data_full[j].rename(columns={data_full[j].columns[-1]: '0'})

    # Creat final dataframe
    data_clean = data_full[0]
    for i in range(1, len(data_full)) :
        if data_full[i].columns[-1] != '0':
            data_clean = pd.merge(data_clean, data_full[i], on=(['Code', 'Description', 'Type']), how='outer')
    data_clean = data_clean.set_index(['Code', 'Description', 'Type']).fillna(0)

    return(data_clean)

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

def plot_between(df, date_1, date_2, code):
    index_1 = df.columns.get_loc(date_1)
    index_2 = df.columns.get_loc(date_2)
    if index_2 < index_1:
        print('Pas top...')
    data = df.loc[[code]].iloc[:, index_1:index_2+1].T
    #-------------------------
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4), gridspec_kw={'width_ratios': [4, 1]})
    data.plot.bar(ax=ax1)
    data.plot.box(ax=ax2)
    ax1.yaxis.set_major_locator(plt.MultipleLocator(500))
    ax1.grid(axis='y')
    ax2.grid(axis='y')
    plt.show()

def plot_between_stacked(df, date_1, date_2):
    index_1 = df.columns.get_loc(date_1)
    index_2 = df.columns.get_loc(date_2)
    if index_2 < index_1:
        print('Pas top...')

    
    data = df.loc[['2','3','4']].iloc[:, index_1:index_2+1]
    data['sum'] = df.loc[['3','4']].iloc[:, index_1:index_2+1].sum(axis=0)
    #data_missions = df.loc[['3','4']].iloc[:, index_1:index_2+1].sum(axis=0)
    print("---- data ----")
    print(data)
    print("---- data missions ----")
    #print(data_missions)

    #print(data_salary+data_missions)
    """
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4), gridspec_kw={'width_ratios': [4, 1]})
    data_salary.plot.bar(ax=ax1, color='r')
    #data_missions.plot.bar(ax=ax1, bottom=data_salary, color='b')
    ax1.yaxis.set_major_locator(plt.MultipleLocator(500))
    ax1.grid(axis='y')
    #ax1.legend(["Salary", "Perdiems"])

    data_salary.plot.box(ax=ax2)
    
    ax2.grid(axis='y')
    plt.show()
    """
    

def _main():
    # Extracting salary data
    print()
    print('Extracting salary data...')
    print()
    data = process_folder('sdworks_LCL')

    # Extracting perdiems and mission spent data
    print()
    print("Loading perdiem data...")
    print()
    data = load_excel(data)

    date1 = '01 2023'
    date2 = '09 2024'
    code = '2'
    """
    2 -> Net salary
    5 -> Net salary + perdiems
    """

    print()
    print('---- Full data ----')
    print(data)

    print()
    print('---- Sum ----')
    print(sum_between(data, date1, date2, code))

    print()
    print('---- Mean ----')
    print(mean_between(data, date1, date2, code))

    print()
    print('---- Plot of data ----')
    #print(plot_between(data, date1, date2, code))

    print(plot_between_stacked(data, date1, date2))



_main()