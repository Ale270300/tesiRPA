import pandas as pd
import pprint

file_path = r"C:\Users\aless\OneDrive\Desktop\Tesi\PE13 - Data for KPI.xlsx"


dati = pd.read_excel(file_path, sheet_name="Paper in conference")

titoli= dati.iloc[:641, 2].tolist()
new_titoli= dati.iloc[641:665, 2].tolist()



d={}
for titolo in titoli:
    if titolo not in d:
        d[titolo]=1
    else:
        d[titolo]+=1

#pp = pprint.PrettyPrinter(indent=2)
#pp.pprint(d)

chiavi = list(d.keys())
#print(len(titoli))
#print(len(chiavi))

#chiavi_m = ['"' + s.replace('"', "'") + '"' for s in chiavi]

#s = ','.join(chiavi_m)

s = ', '.join(f'"{elemento}"' for elemento in chiavi)


file_w = r"C:\Users\aless\OneDrive\Desktop\Tesi\output_paper.txt"

# Scrivi la stringa su un file di testo
with open(file_w, "w", encoding="utf-8") as file:
    file.write(s)

print(s)
set_titoli = set(chiavi)
set_new_titoli = set(new_titoli)

common_elements = set_titoli.intersection(set_new_titoli)

if common_elements:
    print("Gli elementi comuni sono:", common_elements)
else:
    print("Non ci sono elementi comuni.")

