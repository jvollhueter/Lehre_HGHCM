#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Importieren der Bibliothek 'pandas'
"""
import pandas as pd

"""
Definieren der Variable 'file' (string)

Var. 'file': Dateipfad
"""
file = r'Nitrat.xlsx'

"""
Variable 'datas' (pandas dataframe):

Pandas (pd)-Funktion read_excel: Einlesen von Excel-Dateien:
file: Dateipfad zu Datei
skiprow=6: Ersten 6 Zeilen werden beim Einlesen uebersprungen
"""
datas = pd.read_excel(file, skiprows=6)

"""
Variable 'Parameter' (string):

Inhalt der Variable wird am Ende in Name der Ergebnisdatei geschrieben
"""
parameter = datas['Parameter'][0]

"""
Variable mean_MKZ (pandas dataframe):

Gruppieren der Variable 'datas' nach Spalte 'MKZ' (.groupby) und
Bildung von Mittelwerten innerhalb jeder Gruppe (.mean) sowie
löschen der Spalte 'Probenbezug' (.drop)
"""
mean_MKZ = datas.groupby(['MKZ']).mean().drop(columns=['Probenbezug'])

"""
Variable mean_year (pandas dataframe):

Gruppieren der Variable 'datas' nach Spalte 'Jahr' (.groupby) und
Bildung von Mittelwerten innerhalb jeder Gruppe (.mean)
"""
mean_year = datas.groupby(['Jahr']).mean()

"""
Löschen aller Spalten außer 'Ergebnis'
Genauer: Überschreiben der 'Variable mean_year' mit einem dataframe bestehend
aus der Spalte 'Ergebnis' des dataframes 'mean_years'
"""
mean_year = mean_year[['Ergebnis']]

"""
Hinzufuegen einer neuen Spalte 'Einheit' in den dataframes 'mean_MKZ' und
'mean_year' mit dem Wert aus dem dataframe datas in der Spalte 'Einheit' und
der Zeile 0 (Indexing in Python beginnt immer bei 0, in z.B. Excel wäre das
also Zeile 1)
"""
mean_MKZ['Einheit'] = datas['Einheit'][0]
mean_year['Einheit'] = datas['Einheit'][0]

"""
Hinzufuegen einer neuen Spalte 'n_Messwerte' in den dataframes 'mean_MKZ' und
'mean_year' mit der Anzahl der Häufigkeit des Auftretens jeden Wertes in der
Spalte 'Jahr' im dataframe datas (.size)
"""
mean_year['n_Messwerte'] = datas.groupby(['Jahr']).size()
mean_MKZ['n_Messwerte'] = datas.groupby(['MKZ']).size()

"""
Schreiben der Ergebnisdatei mit Hilfe von pandas (pd)-Funktion ExcelWriter()
Dateiname setzt sich zusammen aus 'output_<InhaltVariableParameter>.xlsx'
"""
with pd.ExcelWriter('output_' + parameter + '.xlsx') as writer:
    mean_MKZ.to_excel(writer, sheet_name='Mittelwerte_Messstellen')
    mean_year.to_excel(writer, sheet_name='Jahresmittelwerte')
