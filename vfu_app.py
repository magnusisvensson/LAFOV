
import streamlit as st
import pandas as pd
import hashlib
import random
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV","LAGRV","LGFRI"])

# ===== GEO =====
geo = {
    "Kalmar": (56.66,16.36),
    "Oskarshamn": (57.26,16.45),
    "Karlskrona": (56.16,15.59),
    "Ronneby": (56.21,15.28)
}

def norm(text):
    t=str(text).lower()
    if "påskallavik" in t: return "Oskarshamn"
    if "kallinge" in t: return "Ronneby"
    if any(x in t for x in ["kalmar","lindsdal","nybro","emmaboda","mönsterås","färjestaden"]):
        return "Kalmar"
    return text

def distance_km(a,b):
    if a not in geo or b not in geo:
        return 999
    d=((geo[a][0]-geo[b][0])**2+(geo[a][1]-geo[b][1])**2)**0.5
    return round(d*111,1)

# ================= MAIN =================
if system_file and form_file:

    skolor_all = pd.read_excel(system_file)
    skolor_all.columns = skolor_all.columns.str.strip()

    skolor = skolor_all[
        (skolor_all["Kull"]==kull) &
        (skolor_all["Inriktning"].str.upper()==program)
    ].copy()

    kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

    def get_region(ort):
        ort = ort.lower()
        if "oskarshamn" in ort: return "Oskarshamn"
        if "karlskrona" in ort or "ronneby" in ort: return "Karlskrona"
        return "Kalmar"

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn=[c for c in students.columns if "förnamn" in c.lower()][0]
    ln=[c for c in students.columns if "efternamn" in c.lower()][0]
    bost=[c for c in students.columns if "bostadsort" in c.lower()][0]
    alt=[c for c in students.columns if "alternativ" in c.lower()][0]
    val=[c for c in students.columns if "helst" in c.lower()][0]

    def choose(row):
        if "alternativ" in str(row.get(val,"")).lower():
            if pd.notna(row.get(alt)):
                return row.get(alt)
        return row.get(bost)

    students["Ort"]=students.apply(choose,axis=1)
    students["Ort"]=students["Ort"].apply(norm)
    students["Namn"]=students[fn]+" "+students[ln]

    best_result=None
    best_log=None
    best_count=0

    for _ in range(25):  # optimering

        cap={}
        placements=[]
        logg=[]

        shuffled=students.sample(frac=1)

        for _,student in shuffled.iterrows():

            namn=student["Namn"]
            ort=student["Ort"]
            region=get_region(ort)

            möjliga = []

            # filtrera region
            for _,row in skolor.iterrows():
                partner = str(row["Partnerområde"])
                if region=="Kalmar" and any(x in partner for x in ["Kalmar","Nybro","Mönsterås"]):
                    möjliga.append(row["Skolenhet"])
                if region=="Oskarshamn" and "Oskarshamn" in partner:
                    möjliga.append(row["Skolenhet"])
                if region=="Karlskrona" and "Karlskrona" in partner:
                    möjliga.append(row["Skolenhet"])

            if len(möjliga) < 2:
                möjliga = list(skolor["Skolenhet"])

            skol_sorted = sorted(möjliga, key=lambda s: distance_km(ort, region))

            placed=False
            best_choice=None
            best_cost=9999

            # ===== SPECIAL REGIONER =====
            if region in ["Oskarshamn","Karlskrona"] and program in ["LAFOV","LAGRV"]:

                for i in range(len(skol_sorted)-1):

                    A=skol_sorted[i]
                    B=skol_sorted[i+1]

                    if (
                        cap.get((A,1),0)<kap_map.get(A,999) and
                        cap.get((B,2),0)<kap_map.get(B,999) and
                        cap.get((A,3),0)<kap_map.get(A,999) and
                        cap.get((B,4),0)<kap_map.get(B,999)
                    ):

                        dA=distance_km(ort,region)
                        dB=distance_km(ort,region)
                        cost=max(dA,dB)

                        if cost<best_cost:
                            best_cost=cost
                            best_choice=(A,B)

                if best_choice:
                    A,B=best_choice
                    placement={
                        "Student":namn,
                        "A":A,
                        "B":B,
                        "C":None,
                        "Ort":ort,
                        "Region":region
                    }
                    placed=True

            # ===== KALMAR =====
            else:
                for i in range(len(skol_sorted)-2):

                    A=skol_sorted[i]
                    B=skol_sorted[i+1]
                    C=skol_sorted[i+2]

                    if (
                        cap.get((A,1),0)<kap_map.get(A,999) and
                        cap.get((B,2),0)<kap_map.get(B,999) and
                        cap.get((B,3),0)<kap_map.get(B,999) and
                        cap.get((C,4),0)<kap_map.get(C,999)
                    ):

                        cost=max(
                            distance_km(ort,region),
                            distance_km(ort,region),
                            distance_km(ort,region)
                        )

                        if cost<best_cost:
                            best_cost=cost
                            best_choice=(A,B,C)

                if best_choice:
                    A,B,C=best_choice
                    placement={
                        "Student":namn,
                        "A":A,
                        "B":B,
                        "C":C,
                        "Ort":ort,
                        "Region":region
                    }
                    placed=True

            if not placed:
                logg.append({"Student":namn,"Status":"Får ej plats","Avstånd":"-"})
                continue

            # uppdatera kapacitet
            if region in ["Oskarshamn","Karlskrona"] and program in ["LAFOV","LAGRV"]:
                cap[(A,1)]=cap.get((A,1),0)+1
                cap[(B,2)]=cap.get((B,2),0)+1
                cap[(A,3)]=cap.get((A,3),0)+1
                cap[(B,4)]=cap.get((B,4),0)+1
            else:
                cap[(A,1)]=cap.get((A,1),0)+1
                cap[(B,2)]=cap.get((B,2),0)+1
                cap[(B,3)]=cap.get((B,3),0)+1
                cap[(C,4)]=cap.get((C,4),0)+1

            placements.append(placement)

            logg.append({"Student":namn,"Status":"OK","Avstånd":f"Maxdist ~{best_cost}"})

        if len(placements)>best_count:
            best_count=len(placements)
            best_result=placements
            best_log=logg

    # ===== EXCEL =====
    wb=Workbook()
    ws=wb.active

    ws.column_dimensions["A"].width=40
    for col in ["B","C","D","E"]:
        ws.column_dimensions[col].width=25

    fill_header=PatternFill(start_color="DDDDDD",fill_type="solid")

    ws.append(["Skola","År 1","År 2","År 3","År 4"])

    skol_data={}

    def add(skola, student, col):
        skol_data.setdefault(skola,{})
        skol_data[skola].setdefault(student,{
            "År1":"","År2":"","År3":"","År4":""
        })
        skol_data[skola][student][col]=student

    for p in best_result:

        if p["Region"] in ["Oskarshamn","Karlskrona"] and program in ["LAFOV","LAGRV"]:
            add(p["A"],p["Student"],"År1")
            add(p["B"],p["Student"],"År2")
            add(p["A"],p["Student"],"År3")
            add(p["B"],p["Student"],"År4")
        else:
            add(p["A"],p["Student"],"År1")
            add(p["B"],p["Student"],"År2")
            add(p["B"],p["Student"],"År3")
            add(p["C"],p["Student"],"År4")

    for skola in sorted(skol_data):

        ws.append([f"{skola} (max {int(kap_map.get(skola,0))})"])
        r=ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill=fill_header
            ws.cell(r,c).font=Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        for student,years in sorted(skol_data[skola].items()):
            ws.append(["",years["År1"],years["År2"],years["År3"],years["År4"]])

        ws.append([])

    ws2=wb.create_sheet("Rapport")
    ws2.append(["Student","Status","Avstånd"])

    for r in best_log:
        ws2.append([r["Student"],r["Status"],r["Avstånd"]])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp filer")
