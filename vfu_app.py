
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

# ===== FUNKTIONER =====
def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar","Nybro","Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

def clean_text(text):
    return str(text).lower().replace(" ","").replace("-","")

def match_school(a,s):
    return clean_text(a) in clean_text(s) or clean_text(s) in clean_text(a)

def find_column(cols,keywords):
    for c in cols:
        if any(k.lower() in c.lower() for k in keywords):
            return c
    return None

if system_file and form_file:
    try:
        # ===== SKOLOR =====
        skolor_all = pd.read_excel(system_file)
        skolor_all.columns = skolor_all.columns.str.strip()

        # ✅ vakanta skolor (rätt program)
        vakanta = skolor_all[
            (skolor_all["Kull"].astype(str).str.contains("VAKANT", case=False, na=False)) &
            (skolor_all["Inriktning"].str.upper() == program)
        ]

        vakant_info = [
            f"{r['Skolenhet']} ({int(r['Antal platser'])} platser)"
            for _, r in vakanta.iterrows()
            if pd.notna(r["Antal platser"]) and r["Antal platser"] > 0
        ]

        # ✅ andra kullar
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

        # ===== filtrera rätt kull =====
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

        students["Namn"] = students[fn] + " " + students[ln]
        students["Region"] = students[bost].apply(get_region)

        best_result, best_log = None, None
        best_unplaced = 999

        for _ in range(30):

            result, logg, ej_placerade = [], [], []
            cap = {}

            s_run = students.sample(frac=1)

            for region in s_run["Region"].unique():

                stud = s_run[s_run["Region"] == region]
                skol_lista_full = list(skolor[skolor["Region"]==region]["Skolenhet"])

                for i,(_, student) in enumerate(stud.iterrows()):

                    namn = student["Namn"]
                    ank_raw = str(student.get(ank,"")).strip()

                    if ank_raw.lower() in ["","ingen","-","nej"]:
                        skol_lista = skol_lista_full
                    else:
                        skol_lista = [s for s in skol_lista_full if not match_school(ank_raw,s)]
                        if not skol_lista:
                            skol_lista = skol_lista_full

                    placed = False
                    status = "OK"
                    kommentar = ""

                    for shift in range(len(skol_lista)):
                        A = skol_lista[(i+shift)%len(skol_lista)]
                        B = skol_lista[(i+1+shift)%len(skol_lista)]
                        C = skol_lista[(i+2+shift)%len(skol_lista)]

                        if (
                            cap.get((A,1),0)<kap_map.get(A,999) and
                            cap.get((B,2),0)<kap_map.get(B,999) and
                            cap.get((B,3),0)<kap_map.get(B,999) and
                            cap.get((C,4),0)<kap_map.get(C,999)
                        ):
                            placed=True
                            break

                    if not placed:
                        ej_placerade.append(namn)

                        if vakant_info:
                            tips = "Placera på vakant skola: " + ", ".join(vakant_info)
                        elif andra_info:
                            tips = "Lediga platser på andra kullar: " + ", ".join(andra_info)
                        else:
                            tips = "Övertaligt – fler platser behövs"

                        logg.append({
                            "Student":namn,
                            "Status":"Får ej plats",
                            "Kommentar":"",
                            "Tips":tips
                        })
                        continue

                    cap[(A,1)] = cap.get((A,1),0)+1
                    cap[(B,2)] = cap.get((B,2),0)+1
                    cap[(B,3)] = cap.get((B,3),0)+1
                    cap[(C,4)] = cap.get((C,4),0)+1

                    result += [
                        {"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""},
                        {"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""},
                        {"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn}
                    ]

                    logg.append({
                        "Student":namn,
                        "Status":status,
                        "Kommentar":kommentar,
                        "Tips":""
                    })

            if len(ej_placerade)<best_unplaced:
                best_unplaced=len(ej_placerade)
                best_result=result
                best_log=logg

        df = pd.DataFrame(best_result)

        # ===== LAYOUT =====
        skol_data={}
        for _,row in df.iterrows():
            s=row["Skola"]
            if s not in skol_data:
                skol_data[s]={"År 1":[], "År 2":[], "År 3":[], "År 4":[]}

            for c in skol_data[s]:
                if c in row and row[c]!="":
                    skol_data[s][c].append(row[c])

        wb = Workbook()
        ws = wb.active

        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")
        fill_red = PatternFill(start_color="FF9999", fill_type="solid")
        fill_green = PatternFill(start_color="CCFFCC", fill_type="solid")
        fill_dark = PatternFill(start_color="99CC66", fill_type="solid")

        thin = Side(style="thin")
        thick = Side(style="medium")

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        for skola in sorted(skol_data.keys()):

            kap = kap_map.get(skola)

            if kap is None or kap == 0:
                rubrik = f"{skola} (MAX SAKNAS)"
                color = fill_red
            else:
                rubrik = f"{skola} (max {int(kap)})"
                color = fill_header

            start = ws.max_row+1
            ws.append([rubrik])

            for col in range(1,6):
                ws.cell(start,col).fill=color
                ws.cell(start,col).font=Font(bold=True)

            ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=5)

            data = skol_data[skola]
            max_len = max(len(v) for v in data.values())

            for i in range(max_len):
                ws.append([
                    "",
                    data["År 1"][i] if i<len(data["År 1"]) else "",
                    data["År 2"][i] if i<len(data["År 2"]) else "",
                    data["År 3"][i] if i<len(data["År 3"]) else "",
                    data["År 4"][i] if i<len(data["År 4"]) else ""
                ])

                r = ws.max_row
                ws.cell(r,3).fill = fill_green
                ws.cell(r,4).fill = fill_green
                ws.cell(r,5).fill = fill_dark

            ws.append([])
            ws.append([])

        # ===== RAPPORT =====
        ws2 = wb.create_sheet("Rapport")
        ws2.append(["Student","Status","Kommentar","Tips"])

        for r in best_log:
            ws2.append([
                r["Student"],
                r["Status"],
                r["Kommentar"],
                r.get("Tips","")
            ])

        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

        st.success(f"✅ Klar – {best_unplaced} ej placerade")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
``
