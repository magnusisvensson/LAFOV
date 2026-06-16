
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

# ========= GEO =========
geo = {
    "Kalmar": (56.66, 16.36),
    "Oskarshamn": (57.26, 16.45),
    "Karlskrona": (56.16, 15.59),
    "Ronneby": (56.21, 15.28),
    "Nybro": (56.74, 15.91)
}

def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar","Nybro","Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

def normalize_location(text):
    t = str(text).lower()
    mapping = {
        "kallinge": "Ronneby",
        "lindsdal": "Kalmar",
        "rinkabyholm": "Kalmar",
        "färjestaden": "Kalmar",
        "nybro": "Kalmar",
        "emmaboda": "Kalmar",
        "mönsterås": "Kalmar"
    }
    for k in mapping:
        if k in t:
            return mapping[k]

    if "påskallavik" in t:
        return "Oskarshamn"

    return text

def distance(a,b):
    a = normalize_location(a)
    b = normalize_location(b)
    if a not in geo or b not in geo:
        return 999
    return ((geo[a][0]-geo[b][0])**2 + (geo[a][1]-geo[b][1])**2)**0.5

def distance_km(a,b):
    return round(distance(a,b)*111,1)

def clean_text(t):
    return str(t).lower().replace(" ","").replace("-","")

def match_school(a,s):
    return clean_text(a) in clean_text(s) or clean_text(s) in clean_text(a)

def find_column(cols,keywords):
    for c in cols:
        if any(k.lower() in c.lower() for k in keywords):
            return c
    return None

if system_file and form_file:

    try:
        skolor_all = pd.read_excel(system_file)
        skolor_all.columns = skolor_all.columns.str.strip()

        # ===== VAKANTA =====
        vakanta = skolor_all[
            (skolor_all["Kull"].astype(str).str.contains("VAKANT", case=False, na=False)) &
            (skolor_all["Inriktning"].str.upper() == program)
        ]

        vakant_info = [
            f"{r['Skolenhet']} ({int(r['Antal platser'])} platser)"
            for _, r in vakanta.iterrows()
            if pd.notna(r["Antal platser"]) and r["Antal platser"] > 0
        ]

        # ===== ANDRA KULLAR =====
        andra_kullar = skolor_all[
            (skolor_all["Inriktning"].str.upper() == program) &
            (skolor_all["Kull"] != kull) &
            (~skolor_all["Kull"].astype(str).str.contains("VAKANT", case=False, na=False))
        ]

        andra_info = [
            f"{r['Skolenhet']} ({int(r['Antal platser'])} platser)"
            for _, r in andra_kullar.iterrows()
            if pd.notna(r["Antal platser"]) and r["Antal platser"] > 0
        ]

        skolor = skolor_all[
            (skolor_all["Kull"] == kull) &
            (skolor_all["Inriktning"].str.upper() == program)
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_region)

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # ===== STUDENTER =====
        students = pd.read_excel(form_file, sheet_name="Data")
        students.columns = students.columns.str.strip()

        fn = find_column(students.columns,["förnamn"])
        ln = find_column(students.columns,["efternamn"])
        bost = find_column(students.columns,["bostadsort"])
        ank = find_column(students.columns,["anknytning"])
        alt_bost = find_column(students.columns,["alternativ"])
        val_ort = find_column(students.columns,["helst"])

        def välj_bostadsort(row):
            if "alternativ" in str(row.get(val_ort,"")).lower():
                alt = row.get(alt_bost)
                if pd.notna(alt):
                    return alt
            return row.get(bost)

        students["Aktiv bostadsort"] = students.apply(välj_bostadsort, axis=1)
        students["Aktiv bostadsort"] = students["Aktiv bostadsort"].apply(normalize_location)
        students["Region"] = students["Aktiv bostadsort"].apply(get_region)
        students["Namn"] = students[fn] + " " + students[ln]

        best_result, best_log = None, None
        best_unplaced = 999

        for _ in range(30):

            result, logg, ej_placerade = [], [], []
            cap = {}

            for i, (_, student) in enumerate(students.iterrows()):

                namn = student["Namn"]
                ort = student["Aktiv bostadsort"]
                skol_lista = list(skolor["Skolenhet"])

                # 🔥 sortera efter avstånd
                skol_lista = sorted(skol_lista, key=lambda s: distance(ort, s))

                A = skol_lista[0]
                dist = distance_km(ort, A)

                if cap.get((A,1),0) >= kap_map.get(A,999):
                    ej_placerade.append(namn)

                    if vakant_info:
                        tips = "Vakant: " + ", ".join(vakant_info)
                    elif andra_info:
                        tips = "Andra kullar: " + ", ".join(andra_info)
                    else:
                        tips = "Övertaligt"

                    logg.append({
                        "Student":namn,
                        "Status":"Får ej plats",
                        "Kommentar":"",
                        "Tips":tips,
                        "Avstånd":"-"
                    })
                    continue

                cap[(A,1)] = cap.get((A,1),0)+1

                result.append({"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""})

                logg.append({
                    "Student":namn,
                    "Status":"OK",
                    "Kommentar":"",
                    "Tips":"",
                    "Avstånd":f"{dist} km"
                })

        df = pd.DataFrame(result)

        # ===== LAYOUT =====
        skol_data = {}
        for _, r in df.iterrows():
            s = r["Skola"]
            skol_data.setdefault(s, {"År 1":[],"År 2":[],"År 3":[],"År 4":[]})
            if r["År 1"]:
                skol_data[s]["År 1"].append(r["År 1"])

        wb = Workbook()
        ws = wb.active

        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")
        fill_red = PatternFill(start_color="FF9999", fill_type="solid")

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        for skola in sorted(skol_data):

            kap = kap_map.get(skola)

            if kap is None or kap == 0:
                rubrik = f"{skola} (MAX SAKNAS)"
                color = fill_red
            else:
                rubrik = f"{skola} (max {int(kap)})"
                color = fill_header

            row = ws.max_row+1
            ws.append([rubrik])

            for c in range(1,6):
                ws.cell(row,c).fill = color
                ws.cell(row,c).font = Font(bold=True)

            ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)

            for namn in skol_data[skola]["År 1"]:
                ws.append(["",namn,"","",""])

            ws.append([])
            ws.append([])

        ws2 = wb.create_sheet("Rapport")
        ws2.append(["Student","Status","Kommentar","Tips","Avstånd"])

        for r in logg:
            ws2.append([
                r["Student"],
                r["Status"],
                r["Kommentar"],
                r.get("Tips",""),
                r.get("Avstånd","")
            ])

        file = "kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

        st.success(f"✅ Klar – {len(ej_placerade)} ej placerade")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
