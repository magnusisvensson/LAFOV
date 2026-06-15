
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# === STUDENTGRUPP ===
def get_student_group(bostadsort):
    bostadsort = str(bostadsort)

    if any(x in bostadsort for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"

    if "Karlskrona" in bostadsort:
        return "Karlskrona"

    if "Oskarshamn" in bostadsort:
        return "Oskarshamn"

    return "Övrigt"

# === SKOLGRUPP ===
def get_school_group(partnerområde):
    område = str(partnerområde)

    if any(x in område for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"

    if "Karlskrona" in område:
        return "Karlskrona"

    if "Oskarshamn" in område:
        return "Oskarshamn"

    return "Övrigt"

if system_file and form_file:

    try:
        # === SKOLOR ===
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolkol = "Skolenhet"
        gruppkol = "Partnerområde"
        kull_kol = "Kull"
        kap_kol = "Antal platser"

        skolor["Grupp"] = skolor[gruppkol].apply(get_school_group)
        skolor = skolor[skolor[kull_kol] == kull]

        # === STUDENTER ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()

        fnamn_kol = "Förnamn"
        enamn_kol = "Efternamn"
        ort_kol = "Bostadsort"

        students["Grupp"] = students[ort_kol].apply(get_student_group)

        result = []
        capacity_counter = {}

        st.write("Kör placering...")

        for grupp in students["Grupp"].unique():

            stud_grp = students[students["Grupp"] == grupp]
            skol_grp = skolor[skolor["Grupp"] == grupp]

            skol_lista = list(skol_grp[skolkol])
            kapacitet_map = dict(zip(skol_grp[skolkol], skol_grp[kap_kol]))

            if len(skol_lista) == 0:
                st.warning(f"Inga skolor för grupp: {grupp}")
                continue

            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student[fnamn_kol]} {student[enamn_kol]}"

                for shift in range(len(skol_lista)):
                    A = skol_lista[(i + shift) % len(skol_lista)]
                    B = skol_lista[(i + 1 + shift) % len(skol_lista)]
                    C = skol_lista[(i + 2 + shift) % len(skol_lista)]

                    count = capacity_counter.get((B, 2), 0)

                    if count < kapacitet_map.get(B, 999):
                        break

                capacity_counter[(B, 2)] = capacity_counter.get((B, 2), 0) + 1
                capacity_counter[(B, 3)] = capacity_counter.get((B, 3), 0) + 1

                # ✅ korrekt studentrad (med skolbyten)
                result.append({
                    "Startskola": A,
                    "År 1": namn,
                    "År 2": namn,
                    "År 3": namn,
                    "År 4": namn,
                    "Skola_År2": B,
                    "Skola_År4": C
                })

        df = pd.DataFrame(result)

        # sortera
        df = df.sort_values(by="Startskola")

        # === VISA BLOCK I APP ===
        st.subheader("✅ Placering per skola")

        for skola, grupp_df in df.groupby("Startskola"):

            st.markdown(f"### {skola}")

            vis = grupp_df[["År 1", "År 2", "År 3", "År 4"]].reset_index(drop=True)
            st.dataframe(vis)

        # === EXCEL EXPORT ===
        wb = Workbook()
        ws = wb.active

        # färger
        color_y1 = PatternFill(start_color="FFFFFF", fill_type="solid")
        color_y2 = PatternFill(start_color="CCFFCC", fill_type="solid")  # B-skolan
        color_y3 = PatternFill(start_color="CCFFCC", fill_type="solid")  # B-skolan
        color_y4 = PatternFill(start_color="99CC66", fill_type="solid")  # C-skolan

        # rubrik
        ws.append(["Skola", "År 1", "År 2", "År 3", "År 4"])

        for skola, grupp_df in df.groupby("Startskola"):

            ws.append([skola, "", "", "", ""])

            for _, row in grupp_df.iterrows():

                excel_row = [
                    "",
                    row["År 1"],
                    row["År 2"],
                    row["År 3"],
                    row["År 4"]
                ]

                ws.append(excel_row)

                last_row = ws.max_row

                # färga bas
                ws.cell(row=last_row, column=2).fill = color_y1

                # markera B-skolan (år 2–3)
                ws.cell(row=last_row, column=3).fill = color_y2
                ws.cell(row=last_row, column=4).fill = color_y3

                # markera skolbyte (År 4)
                ws.cell(row=last_row, column=5).fill = color_y4

            ws.append(["", "", "", "", ""])

        output_file = "kull_resultat.xlsx"
        wb.save(output_file)

        with open(output_file, "rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=output_file)

    except Exception as e:
        st.error("Fel i appen:")
        st.write(e)

else:
    st.info("Ladda upp båda filer.")
