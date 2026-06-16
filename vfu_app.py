
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

def normalize_location(text):
    t = str(text).lower()
    if "påskallavik" in t: return "Oskarshamn"
    if "kallinge" in t: return "Ronneby"
    if any(x in t for x in ["lindsdal","rinkabyholm","färjestaden","nybro","emmaboda"]):
        return "Kalmar"
    return text

def distance(a, b):
    a = normalize_location(a)
    b = normalize_location(b)

    if a not in geo or b not in geo:
        return None

    return ((geo[a][0]-geo[b][0])**2 + (geo[a][1]-geo[b][1])**2)**0.5

def distance_km(a, b):
    d = distance(a, b)
    if d is None:
        return None
    return round(d * 111, 1)

def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar","Nybro","Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

def find_column(cols,keywords):
    for c in cols:
        if any(k.lower() in c.lower() for k in keywords):
            return c
    return None

# ===============================
if system_file and form_file:

    try:
        skolor_all = pd.read_excel(system_file)
        skolor_all.columns = skolor_all.columns.str.strip()

        # ===== Tipsdata =====
        vakanta = skolor_all[
            (skolor_all["Kull"].astype(str).str.contains("VAKANT", case=False, na=False)) &
            (skolor_all["Inriktning"].str.upper() == program)
        ]

        vakant_info = [
            f"{r['Skolenhet']} ({int(r['Antal platser'])} platser)"
            for _, r in vakanta.iterrows()
            if pd.notna(r["Antal platser"]) and r["Antal platser"] > 0
        ]

        andra_kullar = skolor_all[
            (skolor_all["Inriktning"].str.upper() == program) &
            (skolor_all["Kull"] != kull) &
            (~skolor_all["Kull"].astype(str).str.contains("VAKANT", case=False, na=False))
        ]

        andra_info = [
            f"{r['Skolenhet']} ({int(r['Antal platser'])} platser)"
            for _, r in andra_kullar.iterrows()
            if pd.notna(r["Antal platser"])
        ]

        # ===== filtrera kull =====
        skolor = skolor_all[
            (skolor_all["Kull"] == kull) &
            (skolor_all["Inriktning"].str.upper() == program)
        ].copy()

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

        # ===== STUDENTER =====
        students = pd.read_excel(form_file, sheet_name="Data")
        students.columns = students.columns.str.strip()

        fn = find_column(students.columns, ["förnamn"])
        ln = find_column(students.columns, ["efternamn"])
        bost = find_column(students.columns, ["bostadsort"])
        alt = find_column(students.columns, ["alternativ"])
        val = find_column(students.columns, ["helst"])

        def choose_location(row):
            if "alternativ" in str(row.get(val,"")).lower():
                if pd.notna(row.get(alt)):
                    return row.get(alt)
            return row.get(bost)

        students["Ort"] = students.apply(choose_location, axis=1)
        students["Ort"] = students["Ort"].apply(normalize_location)
        students["Namn"] = students[fn] + " " + students[ln]

        cap = {}
        result = []
        logg = []

        for i, (_, student) in enumerate(students.iterrows()):

            namn = student["Namn"]
            ort = student["Ort"]

            skol_lista = list(skolor["Skolenhet"])
            skol_lista = sorted(skol_lista, key=lambda s: distance(ort, s) or 999)

            placed = False

            for shift in range(len(skol_lista)):
                A = skol_lista[(i+shift)%len(skol_lista)]
                B = skol_lista[(i+1+shift)%len(skol_lista)]
                C = skol_lista[(i+2+shift)%len(skol_lista)]

                if (
                    cap.get((A,1),0) < kap_map.get(A,999) and
                    cap.get((B,2),0) < kap_map.get(B,999) and
                    cap.get((B,3),0) < kap_map.get(B,999) and
                    cap.get((C,4),0) < kap_map.get(C,999)
                ):
                    placed = True
                    break

            if not placed:
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

            # ✅ uppdatera kapacitet
            cap[(A,1)] = cap.get((A,1),0)+1
            cap[(B,2)] = cap.get((B,2),0)+1
            cap[(B,3)] = cap.get((B,3),0)+1
            cap[(C,4)] = cap.get((C,4),0)+1

            # ✅ ROTATION
            result += [
                {"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""},
                {"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""},
                {"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn}
            ]

            # ✅ AVSTÅND (max av A,B,C)
            dists = []

            for skola in [A,B,C]:
                d = distance_km(ort, skola)
                if d is not None:
                    dists.append((skola, d))

            if dists:
                longest_school, longest_dist = max(dists, key=lambda x: x[1])
                avst_text = f"Längsta pendling: {longest_school}, {longest_dist} km"
            else:
                avst_text = "Okänt avstånd"

            logg.append({
                "Student":namn,
                "Status":"OK",
                "Kommentar":"",
                "Tips":"",
                "Avstånd":avst_text
            })

        df = pd.DataFrame(result)

        # ===== EXCEL =====
        wb = Workbook()
        ws = wb.active

        fill_green = PatternFill(start_color="CCFFCC", fill_type="solid")
        fill_dark = PatternFill(start_color="99CC66", fill_type="solid")
        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")
        fill_red = PatternFill(start_color="FF9999", fill_type="solid")

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        for skola in sorted(df["Skola"].unique()):

            kap = kap_map.get(skola)

            if kap is None or kap == 0:
                rubrik = f"{skola} (MAX SAKNAS)"
                color = fill_red
            else:
                rubrik = f"{skola} (max {int(kap)})"
                color = fill_header

            start = ws.max_row+1
            ws.append([rubrik])

            for c in range(1,6):
                ws.cell(start,c).fill = color
                ws.cell(start,c).font = Font(bold=True)

            ws.merge_cells(start_row=start,start_column=1,end_row=start,end_column=5)

            subset = df[df["Skola"] == skola]

            for _, r in subset.iterrows():
                ws.append([
                    "",
                    r.get("År 1",""),
                    r.get("År 2",""),
                    r.get("År 3",""),
                    r.get("År 4","")
                ])

                rr = ws.max_row
                ws.cell(rr,3).fill = fill_green
                ws.cell(rr,4).fill = fill_green
                ws.cell(rr,5).fill = fill_dark

            ws.append([])
            ws.append([])

        # ===== RAPPORT =====
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

        st.success("✅ Klar")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
``
