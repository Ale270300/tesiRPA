import pandas as pd
import openpyxl

file_path = r'C:\Users\aless\OneDrive\Desktop\Tesi\PE13 - Data for KPI.xlsx'
# Carico il file Excel usando openpyxl per mantenere la formattazione
file = openpyxl.load_workbook(file_path)
foglio = file["Publishing"]
#carico su dataframe pandas
dati = pd.read_excel(file_path, sheet_name="Publishing")

n_l =['Elena Agliari', 'Roberto Basili', 'Emanuele Caglioti', 'Chiara Cammarota', 
    'Giuseppe De Giacomo', 'Alessandro De Luca', 'Piergiorgio Donatelli', 
    'Stefano Faralli', 'Fabio Galasso', 'Stefano Giagu', 'Giorgio Grisetti', 
    'Luca Iocchi', 'Domenico Lembo', 'Maurizio Lenzerini', 'Stefano Leonardi', 
    'Andrea Marrella', 'Iacopo Masi', 'Andrea Messina', 'Daniele Nardi', 
    'Roberto Navigli', 'Fabio Patrizi', 'Giuseppe Perelli', 'Veronica Piccialli', 
    'Antonella Poggi', 'Emanuele Rodolà', 'Riccardo Rosati', 'Fabrizio Silvestri', 
    'Aurelio Uncini', 'Barbara Vantaggi', 'Paola Velardi', 'Simone Agostinelli', 
    'Gianluca Cima', 'Nicola Scianca', 'Edoardo Barba', 'Indro Spinelli', 
    'Alberto Fachechi', 'Silvia Marconi', 'Anna Livia Croella', 'Matteo Negri', 
    'Federico Scafoglieri', 'Elena Umili', 'Leandro de Souza Rosa', 
    'Simone Conia', 'Florin Cuconasu', 'Antonio D\'Orazio', 'Donatella Genovese', 
    'Lorenzo De Rebotti', 'Matteo Benati', 'Matteo Mancanelli', 
    'Lorenzo Colantonio', 'Lorenzo Colantonio', 'Roberto Maria Delfino', 
    'Maria Sofia Bucarelli', 'Pere-Lluis Huguet Cabot', 'E. Agliari', 
    'R. Basili', 'E. Caglioti', 'C. Cammarota', 'G. De Giacomo', 
    'A. De Luca', 'P. Donatelli', 'S. Faralli', 'F. Galasso', 'S. Giagu', 
    'G. Grisetti', 'L. Iocchi', 'D. Lembo', 'M. Lenzerini', 'S. Leonardi', 
    'A. Marrella', 'I. Masi', 'A. Messina', 'D. Nardi', 'R. Navigli', 
    'F. Patrizi', 'G. Perelli', 'V. Piccialli', 'A. Poggi', 'E. Rodolà', 
    'R. Rosati', 'F. Silvestri', 'A. Uncini', 'B. Vantaggi', 'P. Velardi', 
    'S. Agostinelli', 'G. Cima', 'N. Scianca', 'E. Barba', 'I. Spinelli', 
    'A. Fachechi', 'S. Marconi', 'A. L. Croella', 'M. Negri', 
    'F. Scafoglieri', 'E. Umili', 'L. de Souza Rosa', 'S. Conia', 
    'F. Cuconasu', 'A. D\'Orazio', 'D. Genovese', 'L. De Rebotti', 
    'M. Benati', 'M. Mancanelli', 'L. Colantonio', 'R. M. Delfino', 
    'M. S. Bucarelli', 'P. Huguet Cabot']


nome_colonna = 'Authors (critical Mass and new recruitment in bold)'

authors_values = dati.loc[418:428, nome_colonna]

#modifica del DataFrame(le print sono state utili per effettuaare controlli intermedi)
for index, value in authors_values.items():
    s = str(value)
    #print(s)
    if " and" in s:
        val=s.replace(" and",",")
    else:
        val=s
    #print(val)
    
    mia_lista = [autore.strip() for autore in val.split(",")]
    #print (mia_lista)
    n_s=[]
    for autore in mia_lista:
        if autore in n_l:
            n_s.append(autore)
    #print(n_s)
    stringa_concatenata = ', '.join(n_s)
    #print(stringa_concatenata)
    dati.at[index, 'Critical Mass or Recruited or Support Group (CM, R, SG)'] = stringa_concatenata

#aggiornamento del file excel iterando attraverso le righe mantenendo la formattazione originale
for index, value in dati.loc[418:462, 'Critical Mass or Recruited or Support Group (CM, R, SG)'].items():
    cell = foglio.cell(row=index+2, column=4)
    cell.value = value

file.save(file_path)
