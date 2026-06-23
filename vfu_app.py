
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(layout="wide")
st.title("VFU-placering")

# ==============================
# FILUPPLADDNING
# ==============================
stud_file = st.file_uploader("Studentfil", type=["xlsx"])
skol_file = st.file_uploader("Översiktsfil", type=["xlsx"])

if stud_file and skol_file:

    # ==============================
    # AUTO-HITTA RÄTT BLAD
    # ==============================
    def hitta_data_blad(excel_file, typ="student"):

        for sheet in excel_file.sheet_names:

            df_test = pd.read_excel(excel_file, sheet_name=sheet, nrows=5)
            cols = [str(c).lower() for c in df_test.columns]

            if typ == "student":
                if any("förnamn" in c for c in cols) and any("efternamn" in c for c in cols):
                    return sheet

            if typ == "skola":
                if any("skolenhet" in c for c in cols) and any("platser" in c for c in cols):
                    return sheet

        return None

    # ==============================
    # LÄS FILER
    # ==============================
    excel_stud = pd.ExcelFile(stud_file)
    excel_skol = pd.ExcelFile(skol_file)

    stud_sheet = hitta_data_blad(excel_stud, "student")
    skol_sheet = hitta_data_blad(excel_skol, "skola")

    if not stud_sheet:
        st.error("❌ Hittar inget datablad i studentfilen")
        st.write("Tillgängliga blad:", excel_stud.sheet_names)
        st.stop()

    if not skol_sheet:
        st.error("❌ Hittar inget skolblad i översiktsfilen")
        st.write("Tillgängliga blad:", excel_skol.sheet_names)
        st.stop()

    st.success(f"Studentblad: {stud_sheet}")
    st.success(f"Skolblad: {skol_sheet}")

    stud_df = pd.read_excel(excel_stud, sheet_name=stud_sheet)
    skol_df = pd.read_excel(excel_skol, sheet_name=skol_sheet)

    stud_df.columns = stud_df.columns.str.strip()
    skol_df.columns = skol_df.columns.str.strip()

    # ==============================
    # SKAPA STUDENTDATA
    # ==============================
    stud_df["Student"] = stud_df["Förnamn"].astype(str).str.strip() + " " + stud_df["Efternamn"].astype(str).str.strip()

    stud_df = stud_df.rename(columns={
        "Bostadsort": "Ort",
        "Eventuell alternativ bostadsort som du har möjlighet att utgå från under läsåren 26/27 och 27/28": "AltOrt",
        "Jag vill helst utgå från": "Val"
    })

    # ==============================
    # AKTIV ORT
    # ==============================
    def välj_ort(row):
        if "alternativ" in str(row["Val"]).lower():
            return row["AltOrt"]
        return row["Ort"]

    stud_df["AktivOrt"] = stud_df.apply(välj_ort, axis=1)

    # ==============================
    # VAL
    # ==============================
    vald_inriktning = st.selectbox("Inriktning", sorted(skol_df["Inriktning"].dropna().unique()))
    vald_kull = st.selectbox("Kull", sorted(skol_df["Kull"].dropna().unique()))

    region_typ = st.selectbox(
        "Regionmodell (LAFOV/LAGRV)",
        ["Kalmar (ABBC)", "Karlskrona/Oskarshamn (ABAB)"]
    )

    # ==============================
    # FILTRERA SKOLOR
    # ==============================
    skolor_df = skol_df[
        (skol_df["Inriktning"] == vald_inriktning) &
        (skol_df["Kull"] == vald_kull)
    ]

    kapacitet = {
        row["Skolenhet"]: int(row["Antal platser"])
        for _, row in skolor_df.iterrows()
        if pd.notna(row["Antal platser"])
    }

    skolor = list(kapacitet.keys())
    studenter = stud_df["Student"].tolist()

    # ==============================
    # TILLDELNING
    # ==============================
    def tilldela(studenter, skolor, kapacitet, antal):

        idx = 0
        kvar = kapacitet.copy()
        resultat = {}

        for student in studenter:
            val = []

            while len(val) < antal:
                s = skolor[idx % len(skolor)]

                if kvar[s] > 0:
                    val.append(s)
                    kvar[s] -= 1

                idx += 1

            resultat[student] = val

        return resultat

    # ==============================
    # SCHEMA
    # ==============================
    if vald_inriktning == "LGFRI":

        student_skolor = tilldela(studenter, skolor, kapacitet, 2)

        schema = [
            [s, A, A, B]
            for s, (A, B) in student_skolor.items()
        ]

        cols = ["Student", "År1", "År2", "År3"]

    else:

        if "Kalmar" in region_typ:

            student_skolor = tilldela(studenter, skolor, kapacitet, 3)

            schema = [
                [s, A, B, B, C]
                for s, (A, B, C) in student_skolor.items()
            ]

        else:

            student_skolor = tilldela(studenter, skolor, kapacitet, 2)

            schema = [
                [s, A, B, A, B]
                for s, (A, B) in student_skolor.items()
            ]

        cols = ["Student", "År1", "År2", "År3", "År4"]

    schema_df = pd.DataFrame(schema, columns=cols)

    # ==============================
    # VISA RESULTAT
    # ==============================
    st.subheader("Placering")
    st.dataframe(schema_df)

    # ==============================
    # PENDLINGSKONTROLL
    # ==============================
    st.subheader("Pendlingskontroll")

    ort_val = {}
    ok_status = {}

    for i, row in stud_df.iterrows():

        col1, col2, col3 = st.columns([3,3,1])

        student = row["Student"]

        with col1:
            st.write(f"{student} ({row['AktivOrt']})")

        with col2:
            ort_val[student] = st.selectbox(
                "Vald ort",
                sorted(kapacitet.keys()),
                key=f"ort_{i}"
            )

        with col3:
            ok_status[student] = st.checkbox("OK", key=f"ok_{i}")

    # ==============================
    # EXCEL EXPORT
    # ==============================
    def skapa_excel():

        wb = Workbook()

        ws1 = wb.active
        ws1.title = "Placeringar"
        ws1.append(cols)

        for cell in ws1[1]:
            cell.font = Font(bold=True)

        for _, r in schema_df.iterrows():
            ws1.append(list(r))

        ws2 = wb.create_sheet("Rapport")
        ws2.append(["Student", "Aktiv ort", "Vald ort", "OK"])

        for _, r in stud_df.iterrows():
            s = r["Student"]

            ws2.append([
                s,
                r["AktivOrt"],
                ort_val.get(s, ""),
                "OK" if ok_status.get(s) else ""
            ])

        ws3 = wb.create_sheet("Kontroll")
        ws3.append(["Student", "Status"])

        for s in studenter:
            ws3.append([
                s,
                "Klar" if ok_status.get(s) else "Ej klar"
            ])

        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return buffer

    st.download_button(
        "Ladda ner Excel",
        data=skapa_excel(),
        file_name="VFU_resultat.xlsx"
    )
