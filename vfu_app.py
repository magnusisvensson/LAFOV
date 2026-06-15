
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# =========================
# GRUPPLOGIK
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
        skolor = skolor[skolor["Kull"] == kull].copy()
        skolor["Grupp"] = skolor["Partnerområde"].apply(get_school_group)

        # === STUDENTER ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()
        students["Grupp"] = students["Bostadsort"].apply(get_student_group)

        result = []
        capacity_counter = {}

        # =========================
        # ROTATION (RÄTT MODELL)
        # =========================
        for grupp in students["Grupp"].unique():

            stud_grp = students[students["Grupp"] == grupp]
            skol_grp = skolor[skolor["Grupp"] == grupp]

            skol_lista = list(skol_grp["Skolenhet"])
            kap_map = dict(zip(skol_grp["Skolenhet"], skol_grp["Antal platser"]))

            if not skol_lista:
                continue

            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student['Förnamn']} {student['Efternamn']}"

                for shift in range(len(skol_lista)):
                    A = skol_lista[(i + shift) % len(skol_lista)]
                    B = skol_lista[(i + 1 + shift) % len(skol_lista)]
                    C = skol_lista[(i + 2 + shift) % len(skol_lista)]

                    if capacity_counter.get((B, 2), 0) < kap_map.get(B, 999):
                        break

                capacity_counter[(B, 2)] = capacity_counter.get((B, 2), 0) + 1
                capacity_counter[(B, 3)] = capacity_counter.get((B, 3), 0) + 1

                # ✅ VIKTIGT: skapa RÄTT outputstruktur
                result.append({
                    "Skola": A,
                    "År 1": namn,
                    "År 2": "",
                    "År 3": "",
                    "År 4": ""
                })

                result.append({
                    "Skola": B,
                    "År 1": "",
                    "År 2": namn,
                    "År 3": namn,
                    "År 4": ""
                })

                result.append({
                    "Skola": C,
                    "År 1": "",
                    "År 2": "",
                    "År 3": "",
                    "År 4": namn
                })

        df = pd.DataFrame(result).sort_values("Skola")

        # =========================
        # EXCEL GENERERING
        # =========================
        wb = Workbook()
        ws = wb.active

        # kolumnbredd
        ws.column_dimensions["A"].width = 35
        for col in ["B","C","D","E"]:
            ws.column_dimensions[col].width = 30

        # färger
        fill_green = PatternFill(start_color="CCFFCC", fill_type="solid")
        fill_dark = PatternFill(start_color="99CC66", fill_type="solid")
        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")

        # borders
        thin = Side(style="thin")
        thick = Side(style="medium")

        align = Alignment(vertical="center", horizontal="left", wrap_text=True)

        # header
        ws.append(["Skola", "År 1", "År 2", "År 3", "År 4"])

        # =========================
        # BLOCK PER SKOLA
        # =========================
        for skola, grp in df.groupby("Skola"):

            start_row = ws.max_row + 1

            # rubrik
            ws.append([skola, "", "", "", ""])

            for col in range(1,6):
                cell = ws.cell(row=start_row, column=col)
                cell.font = Font(bold=True)
                cell.fill = fill_header

            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)

            # studentrader
            for _, r in grp.iterrows():

                ws.append([
                    "",
                    r["År 1"],
                    r["År 2"],
                    r["År 3"],
                    r["År 4"]
                ])

                row_i = ws.max_row

                for col in range(2,6):
                    ws.cell(row=row_i, column=col).alignment = align

                # färgkodning
                ws.cell(row=row_i, column=3).fill = fill_green
                ws.cell(row=row_i, column=4).fill = fill_green
                ws.cell(row=row_i, column=5).fill = fill_dark

                # grid
                for col in range(1,6):
                    ws.cell(row=row_i, column=col).border = Border(
                        left=thin, right=thin, top=thin, bottom=thin
                    )

                ws.row_dimensions[row_i].height = 22

            end_row = ws.max_row

            # stor ram
            for row in range(start_row, end_row + 1):
                for col in range(1,6):
                    ws.cell(row=row, column=col).border = Border(
                        left=thick if col == 1 else thin,
                        right=thick if col == 5 else thin,
                        top=thick if row == start_row else thin,
                        bottom=thick if row == end_row else thin
                    )

            ws.append(["","","","",""])
            ws.append(["","","","",""])

        # spara
        file_name = "kull_resultat.xlsx"
        wb.save(file_name)

        with open(file_name, "rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=file_name)

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
``
