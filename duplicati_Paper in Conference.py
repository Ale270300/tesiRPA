import pandas as pd

file_path = r"C:\Users\aless\OneDrive\Desktop\Tesi\PE13 - Data for KPI.xlsx"

dati = pd.read_excel(file_path, sheet_name="Paper in conference")

titoli= dati.iloc[:641, 2].tolist()

d={}
for titolo in titoli:
    if titolo not in d:
        d[titolo]=1
    else:
        d[titolo]+=1

chiavi = list(d.keys())


s = ', '.join(f'"{elemento}"' for elemento in chiavi)


file_w = r"C:\Users\aless\OneDrive\Desktop\Tesi\output_paper.txt"

# Scrittura della stringa su un file di testo per facilitarne la lettura
with open(file_w, "w", encoding="utf-8") as file:
    file.write(s)

print(s)

