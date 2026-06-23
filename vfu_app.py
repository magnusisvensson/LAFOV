
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
    # AUTODETEKTERING FUNKTIONER
    # ==============================
    def hitta_blad(excel_file, nyckelord):
        for sheet in excel_file.sheet_names:
            if nyckelord.lower() in sheet.lower():
                return sheet
        return excel_file.sheet_names[0]

    def hitta_kolumn(df, möjliga):
        for col in df.columns:
            for m in möjliga:
                if m.lower() in col.lower():
                    return col
        return None

    # ==============================
    # LÄS FILER
    # ==============================
    excel_stud = pd.ExcelFile(stud_file)
    excel_skol = pd.ExcelFile(skol_file)

    stud_df = pd.read_excel(excel_stud, sheet_name=hitta_blad(excel_stud, "data"))
    skol_df = pd.read_excel(excel_skol, sheet_name=hitta_blad(excel_skol, "skol"))

    stud_df.columns = stud_df.columns.str.strip()
    skol_df.columns = skol_df.columns.str.strip()

    # ==============================
    # HITTA KOLUMNER
    # ==============================
    fnamn = hitta_kolumn(stud_df, ["förnamn"])
    enamn = hitta_kolumn(stud_df, ["efternamn"])
    ort = hitta_kolumn(stud_df, ["bostadsort"])
    alt = hitta_kolumn(stud_df, ["alternativ"])
    val = hitta_kolumn(stud_df, ["helst utgå"])

    skola = hitta_kolumn(skol_df, ["skolenhet"])
    kull = hitta_kolumn(skol_df, ["kull"])
    inriktning = hitta_kolumn(skol_df, ["inriktning"])
    platser = hitta_kolumn(skol_df, ["platser"])

    # ==============================
    # VALIDERA
    # ==============================
    if not all([fnamn, enamn, ort, skola, kull, inriktning, platser]):
        st.error("⚠️ Kunde inte tolka alla kolumner")
        st.write("Studentkolumner:", stud_df.columns.tolist())
        st.write("Skolkolumner:", skol_df.columns.tolist())
        st.stop()

    # ==============================
    # SKAPA STUDENTDATA
    # ==============================
    stud_df["Student"] = stud_df[fnamn].astype(str).str.strip() + " " + stud_df[enamn].astype(str).str.strip()
    stud_df["Ort"] = stud_df[ort]

    if alt:
        stud_df["AltOrt"] = stud_df[alt]
    else:
        stud_df["AltOrt"] = ""

    if val:
        stud_df["Val"] = stud_df[val]
    else:
        stud_df["Val"] = ""

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
    vald_inriktning = st.selectbox("Inriktning", sorted(skol_df[inriktning].dropna().unique()))
    vald_kull = st.selectbox("Kull", sorted(skol_df[kull].dropna().unique()))

    region_typ = st.selectbox(
        "Regionmodell (LAFOV/LAGRV)",
        ["Kalmar (ABBC)", "Karlskrona/Oskarshamn (ABAB)"]
    )

    # ==============================
    # FILTRERA SKOLOR
    # ==============================
    skolor_df = skol_df[
        (skol_df[inriktning] == vald_inriktning) &
        (skol_df[kull] == vald_kull)
    ]

    # kapacitet
    kapacitet = {
        row[skola]: int(row[platser])
        for _, row in skolor_df.iterrows()
        if pd.notna(row[platser])
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
    # VISA
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
    # EXCEL
    # ==============================
    def skapa_excel():

        wb = Workbook()

        ws1 = wb.active
        ws1.title = "Placeringar"
        ws1.append(cols)

        for c in ws1[1]:
            c.font = Font(bold=True)

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
