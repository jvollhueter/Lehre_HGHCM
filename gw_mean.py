#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 21:22:21 2021

@author: jonas
"""
import pandas as pd

file = r'Nitrat.xlsx'

try:
    datas = pd.read_excel(file, skiprows=6)

except:
    print('An error occured. Input file not found')
    exit()

try:
    parameter = datas['Parameter'][0]

    datas = datas.sort_values('MKZ')

    mean_MKZ = datas.groupby(['MKZ']).mean().drop(columns=['Probenbezug'])

    mean_year = datas.groupby(['Jahr']).mean()

    mean_year = mean_year[['Ergebnis']]

except:
    print('An error occured. The structure of your input file is not correct')
    exit()

try:
    mean_MKZ['Einheit'] = datas['Einheit'][0]
    mean_year['Einheit'] = datas['Einheit'][0]

except:
    print('Please note: The structure of your input file is not correct')
    exit()

mean_year['n_Messwerte'] = datas.groupby(['Jahr']).size()
mean_MKZ['n_Messwerte'] = datas.groupby(['MKZ']).size()

try:
    with pd.ExcelWriter('output_' + parameter + '.xlsx') as writer:
        mean_MKZ.to_excel(writer, sheet_name='Mittelwerte_Messstellen')
        mean_year.to_excel(writer, sheet_name='Jahresmittelwerte')
except:
    print('An error occured: Output file could not be written')
    exit()
