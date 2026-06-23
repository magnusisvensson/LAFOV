
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(layout="wide")
st.title("VFU-placering (full modell)")

# ==============================
# FILUPPLADDNING
# ==============================
stud_file = st.file_uploader("Studentfil", type=["xlsx"])
skol_file = st.file_uploader("Översiktsfil", type=["xlsx"])

if stud_file and skol_file:

    stud_df = pd.read_excel(stud_file, sheet_name="Studenter", engine="openpyxl")
    skol_df = pd.read_excel(skol_file, sheet_name="SKOLOR", engine="openpyxl")

    stud_df.columns = stud_df.columns.str.strip()
    skol_df.columns = skol_df.columns.str.strip()

    # ==============================
    # VAL
    # ==============================
    inriktningar = sorted(skol_df["Inriktning"].dropna().unique())
    vald_inriktning = st.selectbox("Inriktning", inriktningar)

    kullar = sorted(skol_df["Kull"].dropna().unique())
    vald_kull = st.selectbox("Kull", kullar)

    filtrerad_df = skol_df[
        (skol_df["Kull"] == vald_kull) &
        (skol_df["Inriktning"] == vald_inriktning)
    ]

    if filtrerad_df.empty:
        st.warning("Inga skolor")
        st.stop()

    # ==============================
    # KAPACITET
    # ==============================
    kapacitet = {}
    for _, row in filtrerad_df.iterrows():
        if pd.notna(row["Skolenhet"]) and pd.notna(row["Antal platser"]):
            kapacitet[row["Skolenhet"]] = int(row["Antal platser"])

    skolor = list(kapacitet.keys())
    studenter = stud_df["Student"].tolist()

    # ==============================
    # REGIONVAL
    # ==============================
    region_typ = st.selectbox(
        "Regionmodell (gäller LAFOV/LAGRV)",
        ["Kalmar (ABBC)", "Karlskrona/Oskarshamn (ABAB)"]
    )

    # ==============================
    # TILLDELNING AV SKOLOR
    # ==============================
    def tilldela(studenter, skolor, kapacitet, antal):

        skolor_lista = list(skolor)
        idx = 0
        result = {}

        kvar = kapacitet.copy()

        for student in studenter:

            val = []

            while len(val) < antal:
                skola = skolor_lista[idx % len(skolor_lista)]

                if kvar[skola] > 0:
                    val.append(skola)
                    kvar[skola] -= 1

                idx += 1

            result[student] = val

        return result

    # ==============================
    # SCHEMA
    # ==============================
    if vald_inriktning == "LGFRI":

        student_skolor = tilldela(studenter, skolor, kapacitet, 2)

        schema = []
        for student, (A, B) in student_skolor.items():
            schema.append([student, A, A, B])

        columns = ["Student", "År1", "År2", "År3"]

    else:

        if "Kalmar" in region_typ:
            student_skolor = tilldela(studenter, skolor, kapacitet, 3)

            schema = []
            for student, (A, B, C) in student_skolor.items():
                schema.append([student, A, B, B, C])

        else:
            student_skolor = tilldela(studenter, skolor, kapacitet, 2)

            schema = []
            for student, (A, B) in student_skolor.items():
                schema.append([student, A, B, A, B])

        columns = ["Student", "År1", "År2", "År3", "År4"]

    schema_df = pd.DataFrame(schema, columns=columns)

    # ==============================
    # VISA
    # ==============================
    st.subheader("Schema")
    st.dataframe(schema_df)

    # ==============================
    # PENDLINGSKONTROLL
    # ==============================
    st.subheader("Pendlingskontroll")

    ort_val = {}
    ok_status = {}

    orter = sorted(stud_df["Ort"].dropna().unique())

    for i, row in stud_df.iterrows():

        col1, col2, col3 = st.columns([3,3,1])

        student = row["Student"]

        with col1:
            st.write(f"{student} ({row['Ort']})")

        with col2:
            ort_val[student] = st.selectbox("Vald ort", orter, key=f"ort_{i}")

        with col3:
            ok_status[student] = st.checkbox("OK", key=f"ok_{i}")

    # ==============================
    # EXCEL
    # ==============================
    def skapa_excel():

        wb = Workbook()

        # ---- schema ----
        ws1 = wb.active
        ws1.title = "Placeringar"

        ws1.append(list(schema_df.columns))
        for cell in ws1[1]:
            cell.font = Font(bold=True)

        for _, row in schema_df.iterrows():
            ws1.append(list(row))

        # ---- rapport ----
        ws2 = wb.create_sheet("Rapport")
        ws2.append(["Student", "Hemort", "Vald ort", "OK"])

        for _, row in stud_df.iterrows():
            student = row["Student"]

            ws2.append([
                student,
                row["Ort"],
                ort_val.get(student, ""),
                "OK" if ok_status.get(student) else ""
            ])

        # ---- kontroll ----
        ws3 = wb.create_sheet("Kontroll")
        ws3.append(["Student", "Status"])

        for _, row in stud_df.iterrows():
            student = row["Student"]

            status = "Klar" if ok_status.get(student) else "Ej klar"
            ws3.append([student, status])

        out = BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    excel_data = skapa_excel()

    st.download_button(
        "Ladda ner Excel",
        data=excel_data,
        file_name=f"VFU_{vald_inriktning}_{vald_kull}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
