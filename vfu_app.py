
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = 26
program = st.selectbox("Program", ["LAFOV","LAGRV"])

# ===== GEO =====
geo = {
    "Kalmar": (56.66,16.36),
    "Oskarshamn": (57.26,16.45),
    "Karlskrona": (56.16,15.59)
}

def norm(x):
    t=str(x).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

def distance(a,b):
    if a not in geo or b not in geo:
        return 999
    return ((geo[a][0]-geo[b][0])**2+(geo[a][1]-geo[b][1])**2)**0.5*111

# ===== DATA =====
skolor = pd.read_excel(system_file)
skolor.columns = skolor.columns.str.strip()

skolor = skolor[
    (skolor["Kull"]==kull) &
    (skolor["Inriktning"].str.upper()==program)
]

kap = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

students = pd.read_excel(form_file, sheet_name="Data")
students.columns = students.columns.str.strip()

fn=[c for c in students.columns if "förnamn" in c.lower()][0]
ln=[c for c in students.columns if "efternamn" in c.lower()][0]
bost=[c for c in students.columns if "bostadsort" in c.lower()][0]

students["Namn"]=students[fn]+" "+students[ln]
students["Ort"]=students[bost].apply(norm)

# ===== PLACERING =====
cap_used = {}
placements = []

for _, s in students.iterrows():

    namn = s["Namn"]
    ort = s["Ort"]

    region = ort

    möjliga = skolor[
        skolor["Partnerområde"].str.contains(region, case=False, na=False)
    ]["Skolenhet"].tolist()

    if len(möjliga) < 2:
        möjliga = list(skolor["Skolenhet"])

    bästa=None
    bästa_cost=99999

    if region in ["Oskarshamn","Karlskrona"]:

        # A B A B
        for i in range(len(möjliga)-1):

            A = möjliga[i]
            B = möjliga[i+1]

            ok = (
                cap_used.get((A,1),0)<kap.get(A,999) and
                cap_used.get((B,2),0)<kap.get(B,999) and
                cap_used.get((A,3),0)<kap.get(A,999) and
                cap_used.get((B,4),0)<kap.get(B,999)
            )

            if ok:
                cost = max(distance(ort,region), distance(ort,region))

                if cost < bästa_cost:
                    bästa_cost = cost
                    bästa = (A,B)

        if bästa:
            A,B = bästa
            placements.append((namn,A,B,None,region))

            cap_used[(A,1)] = cap_used.get((A,1),0)+1
            cap_used[(B,2)] = cap_used.get((B,2),0)+1
            cap_used[(A,3)] = cap_used.get((A,3),0)+1
            cap_used[(B,4)] = cap_used.get((B,4),0)+1

    else:
        # Kalmar A B B C
        skol_list = möjliga

        for i in range(len(skol_list)-2):

            A = skol_list[i]
            B = skol_list[i+1]
            C = skol_list[i+2]

            ok = (
                cap_used.get((A,1),0)<kap.get(A,999) and
                cap_used.get((B,2),0)<kap.get(B,999) and
                cap_used.get((B,3),0)<kap.get(B,999) and
                cap_used.get((C,4),0)<kap.get(C,999)
            )

            if ok:
                cost = max(distance(ort,region), distance(ort,region))

                if cost < bästa_cost:
                    bästa_cost = cost
                    bästa = (A,B,C)

        if bästa:
            A,B,C = bästa
            placements.append((namn,A,B,C,region))

            cap_used[(A,1)] = cap_used.get((A,1),0)+1
            cap_used[(B,2)] = cap_used.get((B,2),0)+1
            cap_used[(B,3)] = cap_used.get((B,3),0)+1
            cap_used[(C,4)] = cap_used.get((C,4),0)+1

# ===== EXCEL =====
wb = Workbook()
ws = wb.active

ws.column_dimensions["A"].width=40
for col in ["B","C","D","E"]:
    ws.column_dimensions[col].width=25

fill = PatternFill(start_color="DDDDDD", fill_type="solid")

ws.append(["Skola","År1","År2","År3","År4"])

skol_data={}

def add(skola, student, col):
    skol_data.setdefault(skola,{})
    skol_data[skola].setdefault(student,{"År1":"","År2":"","År3":"","År4":""})
    skol_data[skola][student][col]=student

for namn,A,B,C,region in placements:

    if region in ["Oskarshamn","Karlskrona"]:
        add(A,namn,"År1")
        add(B,namn,"År2")
        add(A,namn,"År3")
        add(B,namn,"År4")
    else:
        add(A,namn,"År1")
        add(B,namn,"År2")
        add(B,namn,"År3")
        add(C,namn,"År4")

for skola in skol_data:

    ws.append([skola])
    r=ws.max_row

    for c in range(1,6):
        ws.cell(r,c).fill=fill
        ws.cell(r,c).font=Font(bold=True)

    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

    for student,years in skol_data[skola].items():
        ws.append(["",years["År1"],years["År2"],years["År3"],years["År4"]])

    ws.append([])

wb.save("kull_resultat.xlsx")

with open("kull_resultat.xlsx","rb") as f:
    st.download_button("⬇️ Ladda ner", f)
