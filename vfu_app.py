
import streamlit as st
import pandas as pd
import hashlib
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV","LAGRV","LGFRI"])

# ================= GEO =================
geo = {
    "Kalmar": (56.66,16.36),
    "Oskarshamn": (57.26,16.45),
    "Karlskrona": (56.16,15.59),
    "Ronneby": (56.21,15.28)
}

def normalize_location(text):
    t=str(text).lower()
    if "påskallavik" in t: return "Oskarshamn"
    if "kallinge" in t: return "Ronneby"
    if any(x in t for x in ["kalmar","lindsdal","nybro","emmaboda","mönsterås","färjestaden"]):
        return "Kalmar"
    return text

def distance_km_raw(a,b):
    if a not in geo or b not in geo:
        return None
    d=((geo[a][0]-geo[b][0])**2+(geo[a][1]-geo[b][1])**2)**0.5
    return round(d*111,1)

# ================= MAIN =================
if system_file and form_file:

    skolor_all=pd.read_excel(system_file)
    skolor_all.columns=skolor_all.columns.str.strip()

    skolor=skolor_all[
        (skolor_all["Kull"]==kull) &
        (skolor_all["Inriktning"].str.upper()==program)
    ].copy()

    kap_map=dict(zip(skolor["Skolenhet"],skolor["Antal platser"]))

    # ===== AUTO GEO =====
    base_geo={
        "Kalmar":(56.66,16.36),
        "Oskarshamn":(57.26,16.45),
        "Karlskrona":(56.16,15.59)
    }

    school_geo={}

    for _,row in skolor.iterrows():
        skola=row["Skolenhet"]
        partner=str(row.get("Partnerområde",""))

        if any(x in partner for x in ["Kalmar","Nybro","Mönsterås"]):
            base=base_geo["Kalmar"]
        elif "Oskarshamn" in partner:
            base=base_geo["Oskarshamn"]
        elif "Karlskrona" in partner:
            base=base_geo["Karlskrona"]
        else:
            continue

        h=int(hashlib.md5(skola.encode()).hexdigest(),16)
        lat=((h%1000)/1000-0.5)*0.12
        lon=(((h//1000)%1000)/1000-0.5)*0.12

        school_geo[skola]=(base[0]+lat, base[1]+lon)

    def distance_km(student_ort,skola):
        student_ort=normalize_location(student_ort)

        if student_ort not in geo:
            return None

        if skola in school_geo:
            lat1,lon1=geo[student_ort]
            lat2,lon2=school_geo[skola]
            d=((lat1-lat2)**2+(lon1-lon2)**2)**0.5
            return round(d*111,1)

        # fallback
        partner=skolor.loc[skolor["Skolenhet"]==skola,"Partnerområde"].values[0]

        if "Kalmar" in partner:
            return distance_km_raw(student_ort,"Kalmar")
        if "Oskarshamn" in partner:
            return distance_km_raw(student_ort,"Oskarshamn")
        if "Karlskrona" in partner:
            return distance_km_raw(student_ort,"Karlskrona")

        return None

    # ===== STUDENTER =====
    students=pd.read_excel(form_file, sheet_name="Data")
    students.columns=students.columns.str.strip()

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
    students["Ort"]=students["Ort"].apply(normalize_location)
    students["Namn"]=students[fn]+" "+students[ln]

    cap={}
    placements=[]
    logg=[]

    skol_lista=list(skolor["Skolenhet"])

    for i,(_,student) in enumerate(students.iterrows()):

        namn=student["Namn"]
        ort=student["Ort"]

        skol_sorted=sorted(
            skol_lista,
            key=lambda s: distance_km(ort,s) or 999
        )

        placed=False

        for shift in range(len(skol_sorted)-2):
            A=skol_sorted[shift]
            B=skol_sorted[shift+1]
            C=skol_sorted[shift+2]

            if (
                cap.get((A,1),0)<kap_map.get(A,999) and
                cap.get((B,2),0)<kap_map.get(B,999) and
                cap.get((B,3),0)<kap_map.get(B,999) and
                cap.get((C,4),0)<kap_map.get(C,999)
            ):
                placed=True
                break

        if not placed:
            logg.append({
                "Student":namn,
                "Status":"Får ej plats",
                "Kommentar":"",
                "Tips":"Kapacitet slut",
                "Avstånd":"-"
            })
            continue

        cap[(A,1)]=cap.get((A,1),0)+1
        cap[(B,2)]=cap.get((B,2),0)+1
        cap[(B,3)]=cap.get((B,3),0)+1
        cap[(C,4)]=cap.get((C,4),0)+1

        placements.append({
            "Student":namn,
            "A":A,
            "B":B,
            "C":C,
            "Ort":ort
        })

        # avstånd
        dists=[]
        for s in [A,B,C]:
            d=distance_km(ort,s)
            if d:
                dists.append((s,d))

        if dists:
            school,dist=max(dists, key=lambda x:x[1])
            avst=f"Längsta pendling: {school}, {dist} km"
        else:
            avst="Okänd"

        logg.append({
            "Student":namn,
            "Status":"OK",
            "Kommentar":"",
            "Tips":"",
            "Avstånd":avst
        })

    # ===== EXCEL =====
    wb=Workbook()
    ws=wb.active

    fill_header=PatternFill(start_color="DDDDDD",fill_type="solid")
    fill_green=PatternFill(start_color="CCFFCC",fill_type="solid")
    fill_dark=PatternFill(start_color="99CC66",fill_type="solid")

    ws.append(["Skola","År 1","År 2","År 3","År 4"])

    skol_data={}
    for p in placements:

        A,B,C=p["A"],p["B"],p["C"]
        namn=p["Student"]

        skol_data.setdefault(A,[]).append({"År1":namn,"År2":"","År3":"","År4":""})
        skol_data.setdefault(B,[]).append({"År1":"","År2":namn,"År3":namn,"År4":""})
        skol_data.setdefault(C,[]).append({"År1":"","År2":"","År3":"","År4":namn})

    for skola in sorted(skol_data):

        ws.append([f"{skola} (max {int(kap_map.get(skola,0))})"])
        start=ws.max_row

        for c in range(1,6):
            ws.cell(start,c).fill=fill_header
            ws.cell(start,c).font=Font(bold=True)

        rows=skol_data[skola]

        for r in rows:
            ws.append([
                "",
                r["År1"],
                r["År2"],
                r["År3"],
                r["År4"]
            ])
            rr=ws.max_row
            ws.cell(rr,3).fill=fill_green
            ws.cell(rr,4).fill=fill_green
            ws.cell(rr,5).fill=fill_dark

        ws.append([])
        ws.append([])

    ws2=wb.create_sheet("Rapport")
    ws2.append(["Student","Status","Kommentar","Tips","Avstånd"])

    for r in logg:
        ws2.append([
            r["Student"],
            r["Status"],
            r["Kommentar"],
            r["Tips"],
            r["Avstånd"]
        ])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp filer")
