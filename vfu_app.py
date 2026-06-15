
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering (full version)")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# =========================
# GRUPPER
# =========================
def get_student_group(bostadsort):
    bostadsort = str(bostadsort)
    if any(x in bostadsort for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Karlskrona" in bostadsort:
        return "Karlskrona"
    if "Oskarshamn" in bostadsort:
        return "Oskarshamn"
    return "Övrigt"

def get_school_group(partnerområde):
    område = str(partnerområde)
    if any(x in område for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Karlskrona" in område:
        return "Karlskrona"
    if "Oskarshamn" in område:
        return "Oskarshamn"
    return "Övrigt"

# =========================
# MAIN
# =========================
if system_file and form_file:

    try:
        # === SKOLOR ===
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == "LAFOV")
        ].copy()

        skolor["Grupp"] = skolor["Partnerområde"].apply(get_school_group)

        kapacitet_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

        st.write("✅ Antal skolor:", len(skolor))

        # === STUDENTER ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()
        students["Grupp"] = students["Bostadsort"].apply(get_student_group)

        result = []
        ej_placerade = []

        capacity_counter = {}

        # =========================
        # ROTATION MED FULL KAPACITETSKOLL
        # =========================
        for grupp in students["Grupp"].unique():

            stud_grp = students[students["Grupp"] == grupp]
            skol_grp = skolor[skolor["Grupp"] == grupp]

            skol_lista = list(skol_grp["Skolenhet"])

            if not skol_lista:
                continue

            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student['Förnamn']} {student['Efternamn']}"
                placed = False

                for shift in range(len(skol_lista)):

                    A = skol_lista[(i + shift) % len(skol_lista)]
                    B = skol_lista[(i + 1 + shift) % len(skol_lista)]
                    C = skol_lista[(i + 2 + shift) % len(skol_lista)]

                    if (
                        capacity_counter.get((A,1),0) < kapacitet_map.get(A,999) and
                        capacity_counter.get((B,2),0) < kapacitet_map.get(B,999) and
                        capacity_counter.get((B,3),0) < kapacitet_map.get(B,999) and
                        capacity_counter.get((C,4),0) < kapacitet_map.get(C,999)
                    ):
                        placed = True
                        break

                # ✅ OM INGEN PLATS FINNS
                if not placed:
                    ej_placerade.append(namn)
                    continue

                # ✅ RÄKNA
                capacity_counter[(A,1)] = capacity_counter.get((A,1),0)+1
                capacity_counter[(B,2)] = capacity_counter.get((B,2),0)+1
                capacity_counter[(B,3)] = capacity_counter.get((B,3),0)+1
                capacity_counter[(C,4)] = capacity_counter.get((C,4),0)+1

                # ✅ ROTATION OUTPUT
                result.append({"Skola": A, "År 1": namn, "År 2": "", "År 3": "", "År 4": ""})
                result.append({"Skola": B, "År 1": "", "År 2": namn, "År 3": namn, "År 4": ""})
                result.append({"Skola": C, "År 1": "", "År 2": "", "År 3": "", "År 4": namn})

        df = pd.DataFrame(result)

        # =========================
        # ANALYS (VIKTIG!)
        # =========================
        total_studenter = len(students)
        placerade = total_studenter - len(ej_placerade)

        st.subheader("✅ Sammanfattning")
        st.write("Totalt studenter:", total_studenter)
        st.write("Placerade:", placerade)
        st.write("Ej placerade:", len(ej_placerade))

        if ej_placerade:
            st.warning("⚠️ Studenter utan plats:")
            st.write(ej_placerade)

        # =========================
        # KOMPAKT DATA
        # =========================
        skol_data = {}

        for _, row in df.iterrows():
            skola = row["Skola"]

            if skola not in skol_data:
                skol_data[skola] = {"År 1": [], "År 2": [], "År 3": [], "År 4": []}

            for col in ["År 1","År 2","År 3","År 4"]:
                if row[col]:
                    skol_data[skola][col].append(row[col])

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

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        for skola, data in skol_data.items():

            start_row = ws.max_row + 1

            unika = set(data["År 1"] + data["År 2"] + data["År 3"] + data["År 4"])
            antal = len(unika)
            kap = kapacitet_map.get(skola,"-")

            titel = f"{skola} ({antal}/{kap})"

            ws.append([titel,"","","",""])

            for col in range(1,6):
                ws.cell(row=start_row, column=col).fill = fill_header
                ws.cell(row=start_row, column=col).font = Font(bold=True)

            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)

            max_len = max(len(data["År 1"]), len(data["År 2"]), len(data["År 3"]), len(data["År 4"]))

            for i in range(max_len):
                vals = [
                    "",
                    data["År 1"][i] if i < len(data["År 1"]) else "",
                    data["År 2"][i] if i < len(data["År 2"]) else "",
                    data["År 3"][i] if i < len(data["År 3"]) else "",
                    data["År 4"][i] if i < len(data["År 4"]) else ""
                ]

                ws.append(vals)
                r = ws.max_row

                ws.cell(row=r,column=3).fill = fill_green
                ws.cell(row=r,column=4).fill = fill_green
                ws.cell(row=r,column=5).fill = fill_dark

                for c in range(1,6):
                    ws.cell(row=r,column=c).border = Border(
                        left=thin,right=thin,top=thin,bottom=thin
                    )

            end_row = ws.max_row

            for row_i in range(start_row, end_row+1):
                for col_i in range(1,6):
                    ws.cell(row=row_i,column=col_i).border = Border(
                        left=thick if col_i==1 else thin,
                        right=thick if col_i==5 else thin,
                        top=thick if row_i==start_row else thin,
                        bottom=thick if row_i==end_row else thin
                    )

            ws.append(["","","","",""])
            ws.append(["","","","",""])

        # ✅ EXTRA FLÖDE: EJ PLACERADE
        if ej_placerade:
            ws.append(["EJ PLACERADE","","","",""])
            for namn in ej_placerade:
                ws.append(["", namn, "", "", ""])

        filename = "kull_resultat.xlsx"
        wb.save(filename)

        with open(filename,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=filename)

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
