import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font
import os




def data_collec():

    #path:Enter the data file path (you can drag the file directly into the terminal)
    path = input('Drag original file here:').strip()
    year = input('Report Period:')
    prefer = input('\nPlease select a position viewing method:\n1.Market value of holdings\n2.Number of shares\nPlease enter serial number:')
    inv_mkt = input('\nPlease select a market:\n1.SSE 2.SZSE 3.SEHK\nPlease enter the serial number(split by spaces):').split()
    inv_ty = input(
                '\nPlease select investment type:\n1.REITs\n2.Passive Funds\n3.Stocks Long/Short\
                \n4.Debt Hybrid Fund(Secondary Market)\n5.Debt Hybrid Fund(Primary Market)\n6.Dynamic Asset Allocation Funds\
                \n7.Equity-oriented Hybrid Funds\n8.Debt-oriented Hybrid Funds\n9.Balanced Fund\
                \n10.Equity Fund\n11.Enhanced Index Funds\n12.Enhanced Index Debt Funds\
                \n13.Debt Funds(Medium- and long-term)\nPlease enter the serial number(split by spaces):'
                ).split()

    #Create dictionaries and collect input data
    markets = {'1':'SSE', '2':'SZSE', '3':'SEHK'}
    invest_types = {
                    '1':'REITs', '2':'Passive Funds', '3':'Stocks Long/Short',\
                    '4':'Debt Hybrid Fund(Secondary Market)','5':'Debt Hybrid Fund(Primary Market)',\
                    '6':'Dynamic Asset Allocation Funds','7':'Equity-oriented Hybrid Funds','8':'Debt-oriented Hybrid Funds',\
                    '9':'Balanced Fund','10':'Equity Fund','11':'Enhanced Index Funds',\
                    '12':'Enhanced Index Debt Funds','13':'Debt Funds(Medium- and long-term)'
                    }

    m = []
    for mkt in inv_mkt:
        if mkt in markets.keys():
            m.append(markets[mkt])

    i = []
    for ty in inv_ty:
        if ty in invest_types.keys():
            i.append(invest_types[ty])

    #Create a string to make it easier to generate a file name
    str_m = "+".join(m)
    str_i = "+".join(i)
    str = str_m + '+' + str_i


    #Judging by the conditions of the type of investment
    if prefer == '1':
        companies = pd.read_excel(path, header=0)

        #Get the new excel path
        (root, filename) = os.path.split(path)

        #Get the new file path
        str1 = str + '+' + 'Market_value_of_holdings.xlsx'
        path1 = os.path.join(root,str1)

        #Filter the data：

        companies = companies.loc[companies['market'].isin(m)]
        companies = companies.loc[companies['investment type'].isin(i)]

        shares_num = companies['Market value of holdings'].sum()

        companies = companies.loc[companies['industry'].isin(['bank'])]
        banks_num = companies['Market value of holdings'].sum()

        companies_merge = companies.groupby(companies['Stock abbr.'],\
        as_index=False).agg({'Market value of holdings':'sum'})


        a = pd.pivot_table(companies_merge,columns=[u'Stock abbr.'],values=[u'Market value of holdings'])
        wb = openpyxl.load_workbook(path)
        a.to_excel(path1)


        #Write data in the new file
        wb = openpyxl.load_workbook(path1)
        ws = wb.active
        ws.delete_cols(1)
        ws.insert_cols(1,3)

        ws.cell(1,2,'Market value of bank stocks')
        ws.cell(1,3,'Market value of heavy positions')
        ws.cell(2,1, year)
        ws.cell(2,2,banks_num)
        ws.cell(2,3,shares_num)
        ws.column_dimensions['B'].width=17.5
        ws.column_dimensions['C'].width=17.5
        #Format the font
        #font = Font(bold=True)
        #cell2 = ws["B1"]
        #cell2.font = font
        #cell3 = ws["C1"]
        #cell3.font = font
        wb.save(path1)

        print ('File created.')


    elif prefer == '2':
        companies = pd.read_excel(path, header=0)

        (root, filename) = os.path.split(path)

        str2 = str + '+' + 'Number of shares.xlsx'
        path1 = os.path.join(root,str2)

        companies = companies.loc[companies['market'].isin(m)]
        companies = companies.loc[companies['investment type'].isin(i)]

        shares_num = companies['Number of shares'].sum()

        companies = companies.loc[companies['industry'].isin(['bank'])]
        banks_num = companies['Number of shares'].sum()

        companies_merge = companies.groupby(companies['Stock abbr.'],\
        as_index=False).agg({'Number of shares':'sum'})

        a = pd.pivot_table(companies_merge,columns=[u'Stock abbr.'],values=[u'Number of shares'])
        a.to_excel(path1)

        wb = openpyxl.load_workbook(path1)
        ws = wb.active
        ws.delete_cols(1)
        ws.insert_cols(1,3)

        ws.cell(1,2,'Number of bank stocks')
        ws.cell(1,3,'Number of heavy position stocks')
        ws.cell(2,1, year)
        ws.cell(2,2,banks_num)
        ws.cell(2,3,shares_num)
        ws.column_dimensions['B'].width=17.5
        ws.column_dimensions['C'].width=17.5
        wb.save(path1)

        print ('File created')

    else:
        print ('\nPlease enter valid data.')
