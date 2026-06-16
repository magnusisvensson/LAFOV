
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

# =========================
# REGION
# =========================
def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar","Nybro","Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

# =========================
# TEXTMATCHNING
# =========================
def clean_text(text):
    return str(text).lower().replace(" ","").replace("-","")

def match_school(a,s):
    return clean_text(a) in clean_text(s) or clean_text(s) in clean_text(a)

# =========================
# AUTO-KOLUMNER
# =========================
def find_column(cols,keywords):
    for c in cols:
        if any(k.lower() in c.lower() for k in keywords):
            return c
    return None

if system_file and form_file:

    try:
        # ===== SKOLOR =====
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == program)
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

        best_result = None
        best_log = None
        best_unplaced = 999

        # =========================
        # OPTIMERING
        # =========================
        for _ in range(30):

            result = []
            logg = []
            ej_placerade = []
            cap = {}

            s_run = students.sample(frac=1)

            for region in s_run["Region"].unique():

                stud = s_run[s_run["Region"] == region]
                skol_lista_full = list(skolor[skolor["Region"]==region]["Skolenhet"])

                if not skol_lista_full:
                    continue

                for i,(_, student) in enumerate(stud.iterrows()):

                    namn = student["Namn"]
                    ank_raw = str(student.get(ank,"")).strip()

                    if ank_raw.lower() in ["","ingen","-","nej"]:
                        skol_lista = skol_lista_full
                        exkl = []
                    else:
                        skol_lista, exkl = [], []
                        for s in skol_lista_full:
                            if match_school(ank_raw, s):
                                exkl.append(s)
                            else:
                                skol_lista.append(s)
                        if not skol_lista:
                            skol_lista = skol_lista_full

                    placed = False
                    status = "OK"
                    kommentar = ""

                    # ===== LGFRI =====
                    if program == "LGFRI":

                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]

                            if (
                                cap.get((A,1),0)<kap_map[A] and
                                cap.get((A,2),0)<kap_map[A] and
                                cap.get((B,3),0)<kap_map[B]
                            ):
                                placed=True
                                break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({"Student":namn,"Status":"Får ej plats","Kommentar":""})
                            continue

                        cap[(A,1)] = cap.get((A,1),0)+1
                        cap[(A,2)] = cap.get((A,2),0)+1
                        cap[(B,3)] = cap.get((B,3),0)+1

                        result += [
                            {"Skola":A,"År 1":namn,"År 2":namn,"År 3":""},
                            {"Skola":B,"År 1":"","År 2":"","År 3":namn}
                        ]

                    else:

                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]
                            C = skol_lista[(i+2+shift)%len(skol_lista)]

                            if (
                                cap.get((A,1),0)<kap_map[A] and
                                cap.get((B,2),0)<kap_map[B] and
                                cap.get((B,3),0)<kap_map[B] and
                                cap.get((C,4),0)<kap_map[C]
                            ):
                                placed=True
                                break

                        if not placed:
                            for A in skol_lista:
                                for B in skol_lista:
                                    if A!=B:
                                        if (
                                            cap.get((A,1),0)<kap_map[A] and
                                            cap.get((B,2),0)<kap_map[B] and
                                            cap.get((B,3),0)<kap_map[B] and
                                            cap.get((B,4),0)<kap_map[B]
                                        ):
                                            C=B
                                            placed=True
                                            status="Avvikelse"
                                            kommentar="Fallback använd"
                                            break
                                if placed:
                                    break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({"Student":namn,"Status":"Får ej plats","Kommentar":""})
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

                    if exkl:
                        status="Avvikelse"
                        kommentar+=" Anknytning"

                    logg.append({"Student":namn,"Status":status,"Kommentar":kommentar})

            if len(ej_placerade)<best_unplaced:
                best_unplaced=len(ej_placerade)
                best_result=result
                best_log=logg

        df = pd.DataFrame(best_result)

        # =========================
        # ✅ PIXEL LAYOUT
        # =========================
        skol_data={}
        for _,row in df.iterrows():
            s=row["Skola"]

            if s not in skol_data:
                skol_data[s]={"År 1":[],"År 2":[],"År 3":[],"År 4":[]}

            for c in ["År 1","År 2","År 3","År 4"]:
                if c in row and row[c]!="":
                    skol_data[s][c].append(row[c])

        def region_order(s):
            return {"Kalmarregion":1,"Oskarshamn":2,"Karlskrona":3}.get(region_map.get(s,""),0)

        sorted_skolor=sorted(skol_data.keys(), key=lambda x:(region_order(x),x))

        wb=Workbook()
        ws=wb.active

        ws.column_dimensions["A"].width=40
        for c in ["B","C","D","E"]:
            ws.column_dimensions[c].width=30

        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")
        fill_green = PatternFill(start_color="CCFFCC", fill_type="solid")
        fill_dark = PatternFill(start_color="99CC66", fill_type="solid")

        thin = Side(style="thin")
        thick = Side(style="medium")

        align = Alignment(vertical="center", wrap_text=True)

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        current_region=None

        for skola in sorted_skolor:

            region = region_map.get(skola,"")

            if region != current_region:
                ws.append([region.upper()])
                current_region = region

            kap = kap_map.get(skola,0)

            start_row = ws.max_row + 1

            ws.append([f"{skola} (max {int(kap)})"])

            for col in range(1,6):
                ws.cell(start_row,col).fill = fill_header
                ws.cell(start_row,col).font = Font(bold=True)

            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)

            data = skol_data[skola]
            max_len = max(len(v) for v in data.values())

            for i in range(max_len):

                ws.append([
                    "",
                    data["År 1"][i] if i < len(data["År 1"]) else "",
                    data["År 2"][i] if i < len(data["År 2"]) else "",
                    data["År 3"][i] if i < len(data["År 3"]) else "",
                    data["År 4"][i] if i < len(data["År 4"]) else ""
                ])

                r = ws.max_row

                for col in range(2,6):
                    ws.cell(r,col).alignment = align

                ws.cell(r,3).fill = fill_green
                ws.cell(r,4).fill = fill_green
                ws.cell(r,5).fill = fill_dark

                for col in range(1,6):
                    ws.cell(r,col).border = Border(left=thin, right=thin, top=thin, bottom=thin)

            end_row = ws.max_row

            for rr in range(start_row, end_row+1):
                for cc in range(1,6):
                    ws.cell(rr,cc).border = Border(
                        left=thick if cc==1 else thin,
                        right=thick if cc==5 else thin,
                        top=thick if rr==start_row else thin,
                        bottom=thick if rr==end_row else thin
                    )

            ws.append([])
            ws.append([])

        # ===== RAPPORT =====
        ws2 = wb.create_sheet("Rapport")
        ws2.append(["Student","Status","Kommentar"])

        for r in best_log:
            ws2.append([r["Student"],r["Status"],r["Kommentar"]])

        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

        st.success(f"✅ Klar – {best_unplaced} ej placerade")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
``
