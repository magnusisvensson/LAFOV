
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(layout="wide")
st.title("VFU-placering (komplett)")

# ==============================
# Ladda upp filer
# ==============================
file1 = st.file_uploader("Fil 1", type=["xlsx"])
file2 = st.file_uploader("Fil 2", type=["xlsx"])

if file1 and file2:

    # ==============================
    # Identifiera studentfil
    # ==============================
    def innehåller_studentdata(excel):
        for sheet in excel.sheet_names:
            df = pd.read_excel(excel, sheet_name=sheet, nrows=5)
            cols = " ".join(df.columns.str.lower())
            if "förnamn" in cols and "efternamn" in cols:
                return True
        return False

    excel1 = pd.ExcelFile(file1)
    excel2 = pd.ExcelFile(file2)

    if innehåller_studentdata(excel1):
        excel_stud, excel_skol = excel1, excel2
    else:
        excel_stud, excel_skol = excel2, excel1

    # ==============================
    # Hitta rätt blad
    # ==============================
    def hitta_blad(excel, typ):

        # prioritera "Data"
        for s in excel.sheet_names:
            if s.lower() == "data":
                return s

        for s in excel.sheet_names:
            df = pd.read_excel(excel, sheet_name=s, nrows=5)
            cols = " ".join(df.columns.str.lower())

            if typ == "student" and "förnamn" in cols:
                return s
            if typ == "skola" and "skolenhet" in cols:
                return s

        return None

    stud_sheet = hitta_blad(excel_stud, "student")
    skol_sheet = hitta_blad(excel_skol, "skola")

    if not stud_sheet or not skol_sheet:
        st.error("❌ Kunde inte identifiera rätt blad")
        st.stop()

    st.success(f"Studentblad: {stud_sheet}")
    st.success(f"Skolblad: {skol_sheet}")

    stud_df = pd.read_excel(excel_stud, sheet_name=stud_sheet)
    skol_df = pd.read_excel(excel_skol, sheet_name=skol_sheet)

    stud_df.columns = stud_df.columns.str.strip()
    skol_df.columns = skol_df.columns.str.strip()

    # ==============================
    # Studentdata
    # ==============================
    stud_df["Student"] = (
        stud_df["Förnamn"].astype(str).str.strip() +
        " " +
        stud_df["Efternamn"].astype(str).str.strip()
    )

    stud_df = stud_df.rename(columns={
        "Bostadsort": "Ort",
        "Eventuell alternativ bostadsort som du har möjlighet att utgå från under läsåren 26/27 och 27/28": "AltOrt",
        "Jag vill helst utgå från": "Val"
    })

    def välj_ort(row):
        if "alternativ" in str(row["Val"]).lower():
            return row["AltOrt"]
        return row["Ort"]

    stud_df["AktivOrt"] = stud_df.apply(välj_ort, axis=1)

    # ==============================
    # Val
    # ==============================
    vald_inriktning = st.selectbox("Inriktning", sorted(skol_df["Inriktning"].dropna().unique()))
    vald_kull = st.selectbox("Kull", sorted(skol_df["Kull"].dropna().unique()))

    region_typ = st.selectbox(
        "Region (LAFOV/LAGRV)",
        ["Kalmar (ABBC)", "Karlskrona/Oskarshamn (ABAB)"]
    )

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
    # Tilldelning
    # ==============================
    def tilldela(studenter, skolor, kap, antal):
        resultat = {}
        idx = 0
        kvar = kap.copy()

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
    # Schema
    # ==============================
    if vald_inriktning == "LGFRI":

        data = tilldela(studenter, skolor, kapacitet, 2)

        schema = [[s, A, A, B] for s, (A, B) in data.items()]
        cols = ["Student", "År1", "År2", "År3"]

    else:

        if "Kalmar" in region_typ:
            data = tilldela(studenter, skolor, kapacitet, 3)
            schema = [[s, A, B, B, C] for s, (A, B, C) in data.items()]
        else:
            data = tilldela(studenter, skolor, kapacitet, 2)
            schema = [[s, A, B, A, B] for s, (A, B) in data.items()]

        cols = ["Student", "År1", "År2", "År3", "År4"]

    schema_df = pd.DataFrame(schema, columns=cols)

    st.subheader("Placering")
    st.dataframe(schema_df)

    # ==============================
    # Pendlingskontroll
    # ==============================
    st.subheader("Pendlingskontroll")

    ort_val, ok_status = {}, {}

    for i, row in stud_df.iterrows():
        c1, c2, c3 = st.columns([3,3,1])

        s = row["Student"]

        with c1:
            st.write(f"{s} ({row['AktivOrt']})")

        with c2:
            ort_val[s] = st.selectbox("Vald ort", skolor, key=f"ort_{i}")

        with c3:
            ok_status[s] = st.checkbox("OK", key=f"ok_{i}")

    # ==============================
    # Excel
    # ==============================
    def skapa_excel():

        wb = Workbook()

        ws1 = wb.active
        ws1.title = "Placeringar"
        ws1.append(cols)

        # 🔧 FIXAT HÄR
        for cell in ws1[1]:
            cell.font = Font(bold=True)

        for _, r in schema_df.iterrows():
            ws1.append(list(r))

        ws2 = wb.create_sheet("Rapport")
        ws2.append(["Student", "AktivOrt", "Vald ort", "OK"])

        # 🔧 FIXAT
        for cell in ws2[1]:
            cell.font = Font(bold=True)

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

        # 🔧 FIXAT
        for cell in ws3[1]:
            cell.font = Font(bold=True)

        for s in studenter:
            ws3.append([s, "Klar" if ok_status.get(s) else "Ej klar"])

        out = BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    st.download_button(
        "Ladda ner Excel",
        data=skapa_excel(),
        file_name="VFU_resultat.xlsx"
    )
