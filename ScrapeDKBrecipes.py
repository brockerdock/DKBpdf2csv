# -*- coding: utf-8 -*-
"""
Created on Tue May 12 14:53:46 2020

@author: mlaptop
"""

for name in dir():
    if not name.startswith('_'):
        del globals()[name]
        
import camelot
import re
import pandas as pd

import PyPDF2
from decimal import *
from datetime import date
import os
from os import listdir
from os.path import isfile, join

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()  # Hide the main window

mypath = filedialog.askdirectory(title="Wähle Ordner mit Kontoauszügen")

onlyfiles = [f for f in os.listdir('.') if re.match(r'[0-9]+.*\.jpg', f)]
#[f for f in listdir(mypath) if isfile(join('Kontoauszug', f))]
getcontext().prec = 10 # Beträge bis 1 Milliarde-1 können verarbeitet werden

# Checken, ob ein neues Jahr begonnen hat
def PayYear(Start,End,Date):
   month=int(Date.strip().split(".")[1])
   day=int(Date.strip().split(".")[0])
   if (month+11)==Start.month or (month+10)==Start.month:
       Datum=date(Start.year+1,month,day)
   elif (month-11)==Start.month:
       Datum=date(Start.year-1,month,day)
   else:
       Datum=date(Start.year,month,day)
   if (Datum-Start).days<-7 or (Datum-End).days>7:
       print("Das erzeugte Datum stimmt womöglich nicht!")
   return Datum


df = pd.DataFrame(columns=['Name', 'Jahr', 'Nummer', 'DatumAusstellung', 'Startwert', 'Endwert'])
for x in range(len(onlyfiles)):
  if len(onlyfiles[x])==53 or len(onlyfiles[x])==55:
      df.loc[x] = [onlyfiles[x]] + [int(onlyfiles[x].split("_")[3])] + [int(onlyfiles[x].split("_")[4])] + [date(int(onlyfiles[x].split("_")[6]),int(onlyfiles[x].split("_")[7]),int(re.sub("\D", "", onlyfiles[x].split("_")[8])))] + [0] + [0]
  else:
      print("ERROR-Wrong number of Charakters in Filename")

del x, onlyfiles
#Sortiere alle Belege
df=df.sort_values(['Jahr','Nummer'])

# Durchsuche alle verfügbaren PDFs
for i in range(len(df.index)): #range(12,14):
    pdfFileObj = open(mypath + "\\" + df.loc[i][0], 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    
    Pages=[]
    Daten=[] # Kontoauszugsnummer & Datum
    Werte=[] # Kontostand vor/nach Abrechnung
    
    # Durchsuche alle Seiten des PDF nach Buchungstabellen
    for x in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(x)
        Seitentext=pageObj.extractText()
        
        # Überprüfe, ob die Kontoauszugsnummer mit dem Dateinamen übereinstimmt
        if Daten==[] and Seitentext.find("Kontoauszug Nummer "):
            bla=re.findall("Kontoauszug Nummer \d+\s\/\s\d{4}\svom\s\d{2}.\d{2}.\d{4}\sbis\s\d{2}.\d{2}.\d{4}", Seitentext) #https://regex101.com/
            posDaten=bla[0].find('/') #24
            Daten=[int(bla[0][19:posDaten-1]), int(bla[0][posDaten+2:posDaten+6]), int(bla[0][posDaten+11:posDaten+13]), int(bla[0][posDaten+14:posDaten+16]), int(bla[0][posDaten+17:posDaten+21]), int(bla[0][posDaten+26:posDaten+28]), int(bla[0][posDaten+29:posDaten+31]), int(bla[0][posDaten+32:posDaten+36])]
            if not (df.loc[i][1]==Daten[1] and df.loc[i][2]==Daten[0] and df.loc[i][3]==date(Daten[7],Daten[6],Daten[5])):
                print("ERROR!!! Dateiname passt nicht zu Auszugsnummer/Jahr")
            del bla, posDaten
        
        # Suche alle Seiten, auf denen Buchungen vermerkt sind; Speichern in Pages
        if Seitentext.find("Bu.TagGutschrift in EURBelastung in EURWertWir haben für Sie gebucht") > -1:
            Pages.extend([x])
        
        # Suche die Seite, wo der alte und neue Kontostand aufgelistet ist; Speichern in Werte; Haben und Soll in Vorzeichen umwandeln
        if Seitentext.find("ALTER KONTOSTANDNEUER KONTOSTAND") > -1:
            if Werte==[]:
                bla=re.findall("(KONTOSTAND(\d*\.)?(\d*\.)?\d+,\d\d\s(H|S)(\d*\.)?(\d*\.)?\d+,\d\d\s(S|H)EUREUR)", Seitentext)[0][0] #https://regex101.com/
                bla=bla.replace(".", "")
                bla=bla.replace(",", ".")
                bla=bla.replace(" ", "")
                bla=bla[10:]
                Ha=bla.find("H")
                Hb=bla.find("H",bla.find("H")+1)
                Sa=bla.find("S")
                Sb=bla.find("S",bla.find("S")+1)
                Werte=re.split("[HS]", bla)
                Werte.pop(2)
                Werte[0]=Decimal(Werte[0]).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                Werte[1]=Decimal(Werte[1]).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                if (Sa>-1 and Ha==-1) or (Sa>-1 and Sa<Ha):
                    Werte[0]=Werte[0]*-1
                elif (Sb>-1) or (Ha>-1 and Sa>-1 and Sa>Ha):
                    Werte[1]=Werte[1]*-1
                df.loc[i][4]=Werte[0]
                df.loc[i][5]=Werte[1]
                del bla, Sa, Sb, Ha, Hb
            else:
                print("Es gibt mehrere Einträge zu alter/neuer Kontostand")
        
        # Setze die Seiten mit einer Tabelle in einen String
        Seiten=','.join(str(e+1) for e in Pages)
        # Ermittle die Tabellenmaße auf den entsprechenden Seiten
    
    tables = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=Seiten, flavor='lattice')
    if len(tables)<len(Pages):
        tables = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=Seiten, flavor='lattice', line_scale=40)
        if len(tables)<len(Pages):
            tables = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=Seiten, flavor='lattice', line_scale=65)
            if len(tables)!=len(Pages):
                print("Die Anzahl der gefunden Tabellen stimmt nicht mit den entsprechenden Titeln überein.")
        elif len(tables)>len(Pages):
            print("Es wurden zu viele Tabellen entdeckt2")
    elif len(tables)>len(Pages):
        print("Es wurden zu viele Tabellen entdeckt")
    
    del x, Seitentext, Seiten, pageObj, pdfFileObj, pdfReader
    
    if i>0 and df.loc[i][1]==df.loc[i-1][1] and df.loc[i][2]==df.loc[i-1][2] and df.loc[i][3]==df.loc[i-1][3] and df.loc[i][4]==df.loc[i-1][4] and df.loc[i][5]==df.loc[i-1][5]:
        print("Eine Datei wurde übersprungen, da sie doppelt vorkam: " + df.loc[i][0])
    else:
        # Lese alle Buchungen von einer Seite aus einem markierten Tabllenbereich
        for e in Pages:
            #bla=tables[e].df # get a pandas DataFrame!
            #camelot.plot(tables[e], kind='grid')
            #camelot.plot(tables[e], kind='joint')
            #camelot.plot(tables[e], kind='line')
            #camelot.plot(tables[e], kind='contour')
            
            #tables1 = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=str(e+1), flavor='stream', table_areas=[str(tables[e]._bbox[0]-1) + ',' + str(tables[e]._bbox[3]+1) + ',' + str(tables[e]._bbox[2]+1) + ',' + str(tables[e]._bbox[1]-1)])
            #tables2 = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=str(e+1), flavor='stream', table_areas=[str(tables[e]._bbox[0]) + ',' + str(tables[e]._bbox[3]) + ',' + str(tables[e]._bbox[2]) + ',' + str(tables[e]._bbox[1])])
            tables3 = camelot.read_pdf(mypath + "\\" + df.loc[i][0], pages=str(e+1), flavor='stream', table_areas=[str(tables[e]._bbox[0]) + ',' + str(tables[e]._bbox[3]) + ',' + str(tables[e]._bbox[2]) + ',' + str(tables[e]._bbox[1])], columns=['96.1,140.1,395.2,483.5'])
            #bla1=tables1[0].df
            #bla2=tables2[0].df
            #bla3=tables3[0].df
            if e>0:
                Kontoauszug=pd.concat([Kontoauszug, tables3[0].df], ignore_index=True)
            else:
                Kontoauszug=tables3[0].df
    
        del e, tables, tables3
        
        # Start- und Enddatum des Überweisungsträgers
        Begindate=date(Daten[4],Daten[3],Daten[2])
        Enddate=date(Daten[7],Daten[6],Daten[5])
        
        # Checken der Daten auf Plausibilität
        if Daten[4]==Daten[7] and Daten[1]==Daten[4] and (Enddate-Begindate).days<=34:
            YearS=0
        elif Daten[4]<Daten[7] and (Enddate-Begindate).days<=34:
            YearS=1
        else:
            print("Die Jahresangabe spinnt.")
    
        if 'Buchungen' not in locals():    
            Buchungen = pd.DataFrame(columns=['Datum1', 'Datum2', 'Text', 'Betrag', 'Gesamtbetrag'])
    
        # ausgelesene Buchungen formatieren; Leerzellen entfernen; Buchungsdatum überprüfen
        for e in range(len(Kontoauszug)):
            # Titelzeile auf jeder Seite übersrpringen
            if Kontoauszug[0][e]!='Bu.Tag':
                # Weitere Zeilen Überweisungstext; Kein Datum = kein Seitenumbruch
                if Kontoauszug[3][e]=='' and Kontoauszug[4][e]=='' and Kontoauszug[0][e]=='' and Kontoauszug[1][e]=='':
                    Buchungen.loc[len(Buchungen)-1]=[Buchungen.loc[len(Buchungen)-1][0]]+[Buchungen.loc[len(Buchungen)-1][1]]+[Buchungen.loc[len(Buchungen)-1][2]+' '+Kontoauszug[2][e]]+[Buchungen.loc[len(Buchungen)-1][3]]+[Buchungen.loc[len(Buchungen)-1][4]]
                # Weitere Zeilen Überweisungstext; Datum = Seitenumbruch
                elif Kontoauszug[3][e]=='' and Kontoauszug[4][e]=='' and Kontoauszug[0][e]!='' and Kontoauszug[1][e]!='':
                    if Kontoauszug[0][e].strip()==Buchungen.loc[len(Buchungen)-1][0].strftime("%d.%m.") and Kontoauszug[1][e].strip()==Buchungen.loc[len(Buchungen)-1][1].strftime("%d.%m."):
                        Buchungen.loc[len(Buchungen)-1]=[Buchungen.loc[len(Buchungen)-1][0]]+[Buchungen.loc[len(Buchungen)-1][1]]+[Buchungen.loc[len(Buchungen)-1][2]+' '+Kontoauszug[2][e]]+[Buchungen.loc[len(Buchungen)-1][3]]+[Buchungen.loc[len(Buchungen)-1][4]]
                    else:
                        print("Der vorige Eintrag zu der Überweisung hat ein anderes Datum.")
                # Auszahlung
                elif Kontoauszug[3][e]!='' and Kontoauszug[4][e]=='':
                    if len(Buchungen)==0:
                        Datum1=PayYear(Begindate,Enddate,Kontoauszug[0][e])
                        Datum2=PayYear(Begindate,Enddate,Kontoauszug[1][e])
                        Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Werte[0]-Decimal(Kontoauszug[3][e].strip().replace('.', '.').replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                    else:
                        Datum1=PayYear(Begindate,Enddate,Kontoauszug[0][e])
                        Datum2=PayYear(Begindate,Enddate,Kontoauszug[1][e])
                        
                        if e==1 and Buchungen.loc[len(Buchungen)-1][4]==Werte[0]:
                            Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Buchungen.loc[len(Buchungen)-1][4]-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                        elif e==1 and Buchungen.loc[len(Buchungen)-1][4]!=Werte[0] and (Datum1-Buchungen.loc[len(Buchungen)-1][0]).days>30:
                            Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Werte[0]-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                        elif e==1 and Buchungen.loc[len(Buchungen)-1][4]!=Werte[0] and (Datum1-Buchungen.loc[len(Buchungen)-1][0]).days<=30:
                            print("Achtung! Irgendetwas stimmt mit dem Addieren des Gesamtwertes nicht! #Auszahlung")
                        else:
                            Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Buchungen.loc[len(Buchungen)-1][4]-Decimal(Kontoauszug[3][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                # Einzahlung
                elif Kontoauszug[3][e]=='' and Kontoauszug[4][e]!='':
                    if len(Buchungen)==0:
                        Datum1=PayYear(Begindate,Enddate,Kontoauszug[0][e])
                        Datum2=PayYear(Begindate,Enddate,Kontoauszug[1][e])
                        Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Werte[0]+Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                    else:
                        Datum1=PayYear(Begindate,Enddate,Kontoauszug[0][e])
                        Datum2=PayYear(Begindate,Enddate,Kontoauszug[1][e])
                        # 
                        if e==1 and Buchungen.loc[len(Buchungen)-1][4] ==Werte[0]:
                            Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Buchungen.loc[len(Buchungen)-1][4]+Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                        elif e==1 and Buchungen.loc[len(Buchungen)-1][4] !=Werte[0] and (Datum1-Buchungen.loc[len(Buchungen)-1][0]).days>30:
                            Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Werte[0]+Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
                        elif e==1 and Buchungen.loc[len(Buchungen)-1][4] !=Werte[0] and (Datum1-Buchungen.loc[len(Buchungen)-1][0]).days<=30:
                            print("Achtung! Irgendetwas stimmt mit dem Addieren des Gesamtwertes nicht! #Einzahlung")
                        else:
                            Buchungen.loc[len(Buchungen)]=[Datum1] + [Datum2] + [Kontoauszug[2][e]] + [Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)] + [Buchungen.loc[len(Buchungen)-1][4]+Decimal(Kontoauszug[4][e].strip().replace('.', '').replace(',', '.')).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)]
        if Buchungen.loc[len(Buchungen)-1][4]!=Werte[1]:
            print("Achtung! Die Berechnung stimmt nicht!!!!!!!!!!!!!!!!!")
        del e, Pages, Datum1, Datum2, Daten, Begindate, Enddate, Kontoauszug, Werte, YearS
    print(str(i) + " von " + str(len(df)) + ": " + str(df.loc[i][1]) + "-" + str(df.loc[i][2]))

df.to_pickle(mypath + "\\Kontoauszuege.pkl")  # where to save it, usually as a .pkl
Buchungen.to_pickle(mypath + "\\Buchungen.pkl")  # where to save it, usually as a .pkl
# df = pd.read_pickle(mypath + "\\Kontoauszuege.pkl")
# Buchungen = pd.read_pickle(mypath + "\\Buchungen.pkl")