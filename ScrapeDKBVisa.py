# -*- coding: utf-8 -*-
"""
Created on Thu Jun 25 12:24:21 2020

@author: brock
"""

for name in dir():
    if not name.startswith('_'):
        del globals()[name]
        
import camelot
import re
import pandas as pd

import PyPDF2
from decimal import *
import datetime
from datetime import date

from os import listdir
from os.path import isfile, join
from pandas import ExcelWriter

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()  # Hide the main window

mypath = filedialog.askdirectory(title="Wähle Ordner mit Kreditkartenabrechnungen")

onlyfiles = [f for f in listdir(mypath) if isfile(join('Kreditkartenabrechnung', f))]
getcontext().prec = 10 # Beträge bis 1 Milliarde-1 können verarbeitet werden


df = pd.DataFrame(columns=['Name', 'Kartennr', 'DatumStart', 'DatumAusstellung', 'Startwert', 'Endwert', 'Monat', 'Jahr'])

for x in range(len(onlyfiles)):
  if len(onlyfiles[x])==58 or len(onlyfiles[x])==60:
      df.loc[x] = [onlyfiles[x]] + [onlyfiles[x].split("_")[1]] + [date(int(onlyfiles[x].split("_")[3]),int(onlyfiles[x].split("_")[4]),int(onlyfiles[x].split("_")[5].split(".")[0]))-datetime.timedelta(days=30)] + [date(int(onlyfiles[x].split("_")[3]),int(onlyfiles[x].split("_")[4]),int(onlyfiles[x].split("_")[5].split(".")[0]))] + [0] +[0] +[''] + ['']
  else:
      print("ERROR-Wrong number of Charakters in Filename")
      
del x, onlyfiles
# Doppelte Dateien entfernen
df=df.drop_duplicates(subset=['Kartennr', 'DatumAusstellung'], keep='first')
# Sortiere alle Belege
df=df.sort_values(['DatumAusstellung'])
# Index reparieren
df=df.reset_index(drop=True)

# Durchsuche alle verfügbaren PDFs
for i in range(len(df.index)): #range(52,55): #
    pdfFileObj = open(mypath + "\\" + df.loc[i][0], 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    
    Pages=[]
    Daten=[] # Startdatum & Enddatum & Monat & Jahr
    Werte=['',''] # Kontostand vor/nach Abrechnung
    
    # Durchsuche alle Seiten des PDF nach Buchungstabellen (PDF einlesen im Textmodus)
    for x in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(x)
        Seitentext=pageObj.extractText()
        
        # Überprüfe, ob die Kontoauszugsnummer mit dem Dateinamen übereinstimmt
        if Daten==[] and Seitentext.find("Ihre Abrechnung vom ")>-1:
            bla=re.findall("Ihre Abrechnung vom\s\d{2}.\d{2}.\d{4}\sbis\s\d{2}.\d{2}.\d{4}", Seitentext) #https://regex101.com/
            posErgebnis=[Seitentext.split('\n').index(f) for f in Seitentext.split('\n') if "Abrechnung:" in f][0]
            Ergebnis=Seitentext.split('\n')[posErgebnis].split(" ")
            Ergebnis[0]=Ergebnis[-2]
            Ergebnis[1]=Ergebnis[-1]
            Daten=[date(int(re.findall("\d{2}.\d{2}.\d{4}", bla[0])[0].split(".")[2]),int(re.findall("\d{2}.\d{2}.\d{4}", bla[0])[0].split(".")[1]),int(re.findall("\d{2}.\d{2}.\d{4}", bla[0])[0].split(".")[0])), date(int(re.findall("\d{2}.\d{2}.\d{4}", bla[0])[1].split(".")[2]),int(re.findall("\d{2}.\d{2}.\d{4}", bla[0])[1].split(".")[1]),int(re.findall("\d{2}.\d{2}.\d{4}", bla[0])[1].split(".")[0])),Ergebnis[0],Ergebnis[1]]
            if not (df.loc[i][3]==Daten[1]):
                print("ERROR!!! Dateiname passt nicht zu Auszugsnummer/Jahr")
            else:
                df.iat[i, 2]=Daten[0]
                df.iat[i, 6]=Daten[2]
                df.iat[i, 7]=Daten[3]
            del bla, posErgebnis, Ergebnis
        
        # Suche alle Seiten, auf denen Buchungen vermerkt sind; Speichern in Pages
        if Seitentext.find("Angabe des Unternehmens /\nVerwendungszweck") > -1 and Seitentext.find("Betrag in\nEUR") > -1:
            Pages.extend([x])
        
        # Suche die Seite, wo der alte und neue Kontostand aufgelistet ist; Speichern in Werte; Haben und Soll in Vorzeichen umwandeln
        # Übertrag von voriger Abrechnung
        Number_idx = [f for f, item in enumerate(Seitentext.split('\n')) if re.search('\d*.*\d{1,},\d{2}', item)]
        watchdog=0
        #Alter Kontostand
        if Werte[0]=='' and len(Number_idx)>0:
            for f in [Seitentext.split('\n')[i] for i in range(Number_idx[0]-2,Number_idx[1])]:
                if len(re.findall(r"\+\s*\d*.*\d{1,},\d{2}$",f))>0:
                    Werte[0]=re.search(r"\+\s*\d*.*\d{1,},\d{2}$",f)
                    Werte[0]=Werte[0][0]
                    Werte[0]=Werte[0].replace("+","").strip()
                    Werte[0]=Decimal(Werte[0].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                    watchdog=1
                    break
                elif len(re.findall(r"\-\s*\d*.*\d{1,},\d{2}$",f))>0:
                    Werte[0]=re.search(r"\+\s*\d*.*\d{1,},\d{2}$",f)
                    Werte[0]=Werte[0][0]
                    Werte[0]=Werte[0].replace("+","").strip()
                    Werte[0]=Decimal(Werte[0].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*-1
                    watchdog=1
                    break
                elif f =='+':
                    Werte[0]=Seitentext.split('\n')[Number_idx[0]]
                    Werte[0]=Decimal(Werte[0].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                    watchdog=1
                    break
                elif f=='-':
                    Werte[0]=Seitentext.split('\n')[Number_idx[0]]
                    Werte[0]=Decimal(Werte[0].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*-1
                    watchdog=1
                    break
                elif len(re.findall(r"\d{0,3}\.*\d{0,3},\d{2}",f))>0:
                    Werte[0]=re.search(r"\d{0,3}\.*\d{0,3},\d{2}",f)
                    Werte[0]=Werte[0][0]
                    Werte[0]=Werte[0].replace("+","").strip()
                    if '-' in f:
                        Werte[0]=Decimal(Werte[0].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*-1
                    else: 
                        Werte[0]=Decimal(Werte[0].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                    watchdog=1
                    break
            if watchdog==0:
                print("Das Zeichen nach dem ersten Betrag ist weder + oder -.")
        del watchdog
        # Saldo (Neuer Kontostand)
        if Seitentext.find("Neuer Saldo") > -1:
            if Werte[1]=='' and len(Number_idx)>0:
                posErgebnis=[Seitentext.split('\n').index(f) for f in Seitentext.split('\n') if "Neuer Saldo" in f][0]
                # Schauen welche Zahl ist am nächsten am String "Neuer Saldo" dran
                watchdog=0;
                for f in Seitentext.split('\n')[(posErgebnis+1):(sorted(Number_idx, key=lambda x: abs(x - posErgebnis))[0]-1):-1]:
                    if '0,00' in f:
                        Werte[1]=re.findall(r"^0,00",f)
                        Werte[1]=Werte[1][0].strip()
                        Werte[1]=Decimal(Werte[1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                        watchdog=5
                        break
                    elif watchdog==0 and f=='+':
                        watchdog=1
                    elif watchdog==1 and f=='-':
                        watchdog=2
                    elif re.match(r"^\d*.*\d{1,},\d{2}\s*\+",f):
                        Werte[1]=re.findall(r"^\d*.*\d{1,},\d{2}\s*\+",f)
                        Werte[1]=Werte[1][0].replace("+","").strip()
                        Werte[1]=Decimal(Werte[1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                        watchdog=3
                        break
                    elif re.match(r"^\d*.*\d{1,},\d{2}\s*\-",f):
                        Werte[1]=re.findall(r"^\d*.*\d{1,},\d{2}\s*\-",f)
                        Werte[1]=Werte[1][0].replace("+","").strip()
                        Werte[1]=Decimal(Werte[1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*-1
                        watchdog=4
                        break
                    elif watchdog==1 and re.match(r"^\s*\d*.*\d{1,},\d{2}\s*$",f):
                        Werte[1]=f.strip()
                        Werte[1]=Decimal(Werte[1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                    elif watchdog==2 and re.match(r"^\s*\d*.*\d{1,},\d{2}\s*$",f):
                        Werte[1]=f.strip()
                        Werte[1]=Decimal(Werte[1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*-1
                if watchdog==0:
                    print("Es wurde kein Saldo gefunden.")
                del watchdog
            else:
                print("Es keine Beträge auf der Seite oder es gibt mehrere Salden.")
        if x==pdfReader.numPages-1 and (Werte[0]=='' or Werte[1]==''):
            print("Es gibt nicht Start- und Endwert")
        
        # Setze die Seiten mit einer Tabelle in einen String
        Seiten=','.join(str(e+1) for e in Pages)
        # Ermittle die Tabellenmaße auf den entsprechenden Seiten
        df.iat[i,4]=Werte[0]
        df.iat[i,5]=Werte[1]
        
    del x, Seitentext, pageObj, pdfFileObj, pdfReader
    
    if i>0 and df.loc[i][1]==df.loc[i-1][1] and df.loc[i][2]==df.loc[i-1][2] and df.loc[i][3]==df.loc[i-1][3] and df.loc[i][4]==df.loc[i-1][4] and df.loc[i][5]==df.loc[i-1][5]:
        print("Eine Datei wurde übersprungen, da sie doppelt vorkam: " + df.loc[i][0])
    else:
        tables3 = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=Seiten, flavor='stream', table_areas=['30,460,590,0'], columns=['77,116,284,371,449,513'])
        if 'Buchungen' not in locals():    
            Buchungen = pd.DataFrame(columns=['Datum1', 'Datum2', 'Text', 'Währung', 'Betrag', 'Kurs', 'BetragEUR', 'Gesamtbetrag'])
        for e in Pages:
            Kontoauszug=tables3[e].df
            if Kontoauszug.loc[0][6]=='Betrag in' and Kontoauszug.loc[1][6]=='EUR' and (Kontoauszug.loc[0][0]=='Beleg-' or Kontoauszug.loc[0][0]=='Datum') and (Kontoauszug.loc[1][0]=='datum' or Kontoauszug.loc[1][0]=='Beleg'):
                # Falls in der Datei kein Startsaldo steht, wird vom Neuen Saldo zurück gerechnet
                if i>0 and (df.loc[i][2]-df.loc[i-1][3]).days>5 and e==0 and Kontoauszug.loc[2][2]!='Saldo letzte Abrechnung':
                    Startsaldo=[]
                    for f in range(len(Kontoauszug)-1,0,-1):
                        if Kontoauszug[2][f]=='Neuer Saldo':
                            Startsaldo=Decimal(Kontoauszug[6][f].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug[6][f].strip()[-1]+'1')
                        elif re.match(r"^\d*.*\d{1,}.\d{2}\s*",str(Startsaldo)) and re.match(r"^\d*.*\d{1,},\d{2}\s*",Kontoauszug[6][f]):
                            Startsaldo=Startsaldo+Decimal(Kontoauszug[6][f].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug[6][f].strip()[-1]+'1')*-1
                    if df.loc[i][4]!=Startsaldo:
                        df.loc[i][4]=Startsaldo
                        print('Der Startsaldo wurde neu berechnet, weil er in der Datei ' + str(df.loc[i][0]) + ' fehlte.')
                for f in range(0,len(Kontoauszug)):
                    if re.match(r"^\d*.*\d{1,},\d{2}\s*",Kontoauszug.loc[f][6]):
                        if Kontoauszug.loc[f][2]=='Saldo letzte Abrechnung':
                            if Decimal(Kontoauszug.loc[f][6].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug.loc[f][6].strip()[-1]+'1')!=df.loc[i][4]:
                                print("Der Startsaldo stimmt so nicht!")
                        elif re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][0]) or re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][1]) or re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][1]) or re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][0]):
                            Buchungen.loc[len(Buchungen)]=['']+['']+['']+['']+['']+['']+['']+['']
                            Buchungen.loc[len(Buchungen)-1][2]=Kontoauszug.loc[f][2]
                            Buchungen.loc[len(Buchungen)-1][3]=Kontoauszug.loc[f][3]
                            Buchungen.loc[len(Buchungen)-1][4]=Kontoauszug.loc[f][4]
                            Buchungen.loc[len(Buchungen)-1][5]=Kontoauszug.loc[f][5]
                            Buchungen.loc[len(Buchungen)-1][6]=Decimal(Kontoauszug.loc[f][6].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug.loc[f][6].strip()[-1]+'1')
                            if 'Startsaldo' in locals() and re.match(r"^\d*.*\d{1,}.\d{2}\s*",str(Startsaldo)):
                                Buchungen.loc[len(Buchungen)-1][7]=Startsaldo+Buchungen.loc[len(Buchungen)-1][6]
                                del Startsaldo
                            elif e==Pages[0] and f==3:
                                Buchungen.loc[len(Buchungen)-1][7]=df.loc[i][4]+Buchungen.loc[len(Buchungen)-1][6]
                            else:
                                Buchungen.loc[len(Buchungen)-1][7]=Buchungen.loc[len(Buchungen)-2][7]+Buchungen.loc[len(Buchungen)-1][6]
                            if re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][0]):
                                Buchungen.loc[len(Buchungen)-1][0]=Kontoauszug.loc[f][0]
                            if re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][1]):
                                Buchungen.loc[len(Buchungen)-1][1]=Kontoauszug.loc[f][1]
                            elif re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*Habenzins auf",Kontoauszug.loc[f][1]):
                                Buchungen.loc[len(Buchungen)-1][1]=Kontoauszug.loc[f][1][0:Kontoauszug.loc[f][1].find("Habenzins")].strip()
                                Buchungen.loc[len(Buchungen)-1][2]=Kontoauszug.loc[f][1][Kontoauszug.loc[f][1].find("Habenzins"):]
                                Buchungen.loc[len(Buchungen)-1][6]=Decimal(Kontoauszug.loc[f][6].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug.loc[f][6].strip()[-1]+'1')
                            elif re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][1]):
                                Buchungen.loc[len(Buchungen)-1][0]=Kontoauszug.loc[f][1].split()[0].strip()
                                Buchungen.loc[len(Buchungen)-1][1]=Kontoauszug.loc[f][1].split()[1].strip()
                            elif re.match(r"^\s*\d{2}.\d{2}.\d{2}\s*\d{2}.\d{2}.\d{2}\s*$",Kontoauszug.loc[f][0]):
                                Buchungen.loc[len(Buchungen)-1][0]=Kontoauszug.loc[f][0].split()[0].strip()
                                Buchungen.loc[len(Buchungen)-1][1]=Kontoauszug.loc[f][0].split()[1].strip()
                        elif Kontoauszug.loc[f][0]=='' and Kontoauszug.loc[f][1]=='' and (Kontoauszug.loc[f][2].find('für Auslandseinsatz')>-1 or Kontoauszug.loc[f][2].find('GS Auslandseinsatz')>-1):
                            Buchungen.loc[len(Buchungen)]=['']+['']+['']+['']+['']+['']+['']+['']
                            Buchungen.loc[len(Buchungen)-1][0]=Buchungen.loc[len(Buchungen)-2][0]
                            Buchungen.loc[len(Buchungen)-1][1]=Buchungen.loc[len(Buchungen)-2][0]
                            Buchungen.loc[len(Buchungen)-1][2]=Kontoauszug.loc[f][2]
                            Buchungen.loc[len(Buchungen)-1][3]=Kontoauszug.loc[f][3]
                            Buchungen.loc[len(Buchungen)-1][4]=Kontoauszug.loc[f][4]
                            Buchungen.loc[len(Buchungen)-1][5]=Kontoauszug.loc[f][5]
                            Buchungen.loc[len(Buchungen)-1][6]=Decimal(Kontoauszug.loc[f][6].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug.loc[f][6].strip()[-1]+'1')
                            Buchungen.loc[len(Buchungen)-1][7]=Buchungen.loc[len(Buchungen)-2][7]+Buchungen.loc[len(Buchungen)-1][6]
                    elif Kontoauszug.loc[f][6]=='' and Kontoauszug.loc[f][0]=='' and Kontoauszug.loc[f][1]=='':
                        Buchungen.loc[len(Buchungen)-1][2]=Buchungen.loc[len(Buchungen)-1][2] + Kontoauszug.loc[f][2]
            elif Kontoauszug.loc[2][2].find('Neuer Saldo')>-1:
                if Decimal(Kontoauszug.loc[2][6].strip()[0:-1].replace(".", "").replace(",", ".")).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)*int(Kontoauszug.loc[2][6].strip()[-1]+'1')!=Buchungen.loc[len(Buchungen)-1][7]:
                    print('Beim Summieren der Buchungswerte gibt es einen Fehler!')
            else:
                print("Anfang der Tabelle stimmt nicht!")
            
            if e>0:
                Kontoauszug=pd.concat([tables3[0].df, Kontoauszug], ignore_index=True)
            else:
                Kontoauszug=tables3[0].df

    if Buchungen.loc[len(Buchungen)-1][7]!=df.loc[i][5]:
        print("Beim Aufsummieren der Buchungen im Kontoauszug: " + str(df.loc[i][0]) + " gibt es einen Fehler!")
    del e, tables3, f, posErgebnis, Pages, Number_idx, Werte, Daten
    print(str(i) + " von " + str(len(df)-1) + ": " + str(df.loc[i][6]) + ' ' + str(df.loc[i][7]))
    
df.to_pickle(mypath + "\\VISA-Übersicht.pkl")  # where to save it, usually as a .pkl
Buchungen.to_pickle(mypath + "\\VISA-Buchungen.pkl")  # where to save it, usually as a .pkl
# df = pd.read_pickle(mypath + "\\Kontoauszuege.pkl")
# Buchungen = pd.read_pickle(mypath + "\\Buchungen.pkl")

writer = ExcelWriter(mypath + "\\ÜbersichtPDF-Visa.xlsx")
df.to_excel(writer,'Übersicht')
writer.save()
writer = ExcelWriter(mypath + "\\Buchungen-Visa.xlsx")
Buchungen.to_excel(writer,'Buchungen')
writer.save()