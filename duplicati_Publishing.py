import pandas as pd

file_path = r"C:\Users\aless\OneDrive\Desktop\Tesi\PE13 - Data for KPI.xlsx"

dati = pd.read_excel(file_path, sheet_name="Publishing")

titoli= dati.iloc[:418, 4].tolist()

d={}
for titolo in titoli:
    if titolo not in d:
        d[titolo]=1
    else:
        d[titolo]+=1

chiavi = list(d.keys())

s = ', '.join(f'"{elemento}"' for elemento in chiavi)

print(s)

