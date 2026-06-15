
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Optimerad placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# =========================
# REGION (inkl Övrigt → Kalmar)
# =========================
def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

if system_file and form_file:

    try:
        # =========================
        # SKOLOR
        # =========================
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == "LAFOV")
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_region)

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # =========================
        # STUDENTER
        # =========================
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()
        students["Region"] = students["Bostadsort"].apply(get_region)

        best_result = None
        best_unplaced = 999

        # =========================
        # 🔥 OPTIMERING (flertalet körningar)
        # =========================
        for attempt in range(30):

            result = []
            ej_placerade = []
            capacity_counter = {}

            students_run = students.sample(frac=1).reset_index(drop=True)

            for region in students_run["Region"].unique():

                stud_grp = students_run[students_run["Region"] == region]
                skol_lista = list(skolor[skolor["Region"] == region]["Skolenhet"])

                for i, (_, student) in enumerate(stud_grp.iterrows()):

                    namn = f"{student['Förnamn']} {student['Efternamn']}"
                    placed = False

                    # ✅ 1. Full rotation
                    for shift in range(len(skol_lista)):
                        A = skol_lista[(i+shift)%len(skol_lista)]
                        B = skol_lista[(i+1+shift)%len(skol_lista)]
                        C = skol_lista[(i+2+shift)%len(skol_lista)]

                        if (
                            capacity_counter.get((A,1),0) < kap_map[A] and
                            capacity_counter.get((B,2),0) < kap_map[B] and
                            capacity_counter.get((B,3),0) < kap_map[B] and
                            capacity_counter.get((C,4),0) < kap_map[C]
                        ):
                            placed = True
                            break

                    # ✅ 2. fallback (minst ett byte)
                    if not placed:
                        for A in skol_lista:
                            for B in skol_lista:
                                if A == B:
                                    continue

                                if (
                                    capacity_counter.get((A,1),0) < kap_map[A] and
                                    capacity_counter.get((B,2),0) < kap_map[B] and
                                    capacity_counter.get((B,3),0) < kap_map[B] and
                                    capacity_counter.get((B,4),0) < kap_map[B]
                                ):
                                    C = B
                                    placed = True
                                    break
                            if placed:
                                break

                    if not placed:
                        ej_placerade.append(namn)
                        continue

                    # ✅ uppdatera kapacitet
                    capacity_counter[(A,1)] = capacity_counter.get((A,1),0)+1
                    capacity_counter[(B,2)] = capacity_counter.get((B,2),0)+1
                    capacity_counter[(B,3)] = capacity_counter.get((B,3),0)+1
                    capacity_counter[(C,4)] = capacity_counter.get((C,4),0)+1

                    # ✅ resultat
                    result.append({"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""})
                    result.append({"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""})
                    result.append({"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn})

            if len(ej_placerade) < best_unplaced:
                best_unplaced = len(ej_placerade)
                best_result = (result, ej_placerade)

        result, ej_placerade = best_result
        df = pd.DataFrame(result)

        # =========================
        # KOMPAKT DATA
        # =========================
        skol_data = {}

        for _, row in df.iterrows():
            sk = row["Skola"]
            if sk not in skol_data:
                skol_data[sk] = {"År 1":[],"År 2":[],"År 3":[],"År 4":[]}

            for c in ["År 1","År 2","År 3","År 4"]:
                if row[c] != "":
                    skol_data[sk][c].append(row[c])

        # =========================
        # SORTERING
        # =========================
        def region_order(s):
            r = region_map.get(s,"")
            if r == "Kalmarregion": return 1
            if r == "Oskarshamn": return 2
            if r == "Karlskrona": return 3
            return 0

        sorted_skolor = sorted(skol_data.keys(), key=lambda x: (region_order(x), x))

        # =========================
        # EXCEL
        # =========================
        wb = Workbook()
        ws = wb.active

        ws.column_dimensions["A"].width = 40
        for c in ["B","C","D","E"]:
            ws.column_dimensions[c].width = 30

        fill_green = PatternFill(start_color="CCFFCC", fill_type="solid")
        fill_dark = PatternFill(start_color="99CC66", fill_type="solid")
        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")

        thin = Side(style="thin")
        thick = Side(style="medium")

        align = Alignment(vertical="center", horizontal="left", wrap_text=True)

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        current_region = None

        for skola in sorted_skolor:

            region = region_map.get(skola,"")

            if region != current_region:
                ws.append([region.upper(),"","","",""])
                current_region = region

            data = skol_data[skola]
            antal = len(set(data["År 1"]+data["År 2"]+data["År 3"]+data["År 4"]))
            kap = kap_map.get(skola,"-")

            start_row = ws.max_row + 1

            ws.append([f"{skola} ({antal}/{kap})","","","",""])

            for col in range(1,6):
                cell = ws.cell(row=start_row, column=col)
                cell.font = Font(bold=True)
                cell.fill = fill_header

            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)

            max_len = max(len(data["År 1"]),len(data["År 2"]),len(data["År 3"]),len(data["År 4"]))

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
                    ws.cell(row=r, column=col).alignment = align

                ws.cell(row=r, column=3).fill = fill_green
                ws.cell(row=r, column=4).fill = fill_green
                ws.cell(row=r, column=5).fill = fill_dark

                for col in range(1,6):
                    ws.cell(row=r, column=col).border = Border(
                        left=thin,right=thin,top=thin,bottom=thin
                    )

            end_row = ws.max_row

            for r in range(start_row, end_row+1):
                for c in range(1,6):
                    ws.cell(row=r,column=c).border = Border(
                        left=thick if c==1 else thin,
                        right=thick if c==5 else thin,
                        top=thick if r==start_row else thin,
                        bottom=thick if r==end_row else thin
                    )

            ws.append(["","","","",""])
            ws.append(["","","","",""])

        # EJ PLACERADE
        if ej_placerade:
            ws.append(["EJ PLACERADE","","","",""])
            for s in ej_placerade:
                ws.append(["",s,"","",""])

        file = "kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

        st.success(f"✅ Bästa lösning hittad – endast {len(ej_placerade)} ej placerade")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
