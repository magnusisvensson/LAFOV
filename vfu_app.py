
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# =========================
# UI
# =========================
st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# =========================
# REGION
# =========================
def get_region(text):
    t = str(text).lower()

    if "oskarshamn" in t:
        return "Oskarshamn"
    if any(x in t for x in ["karlskrona","ronneby","rödeby"]):
        return "Karlskrona"
    if "kalmar" in t:
        return "Kalmar"

    return None


if system_file and form_file:

    # =========================
    # SKOLOR
    # =========================
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # =========================
    # KAP
    # =========================
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 2

    # =========================
    # STUDENTER
    # =========================
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostads" in c.lower()][0]
    alt = [c for c in students.columns if "alternativ" in c.lower()][0]
    pref = [c for c in students.columns if "utgå" in c.lower()][0]

    students["Namn"] = (students[fn] + " " + students[ln]).str.strip()

    def choose_loc(row):
        if "alternativ" in str(row[pref]).lower():
            return row[alt]
        return row[bost]

    students["Ort"] = students.apply(choose_loc, axis=1)
    students["Region"] = students["Ort"].apply(get_region)

    # =========================
    # PLACERING (KORREKT)
    # =========================
    usage = defaultdict(int)
    results = []

    for _, s in students.iterrows():

        region = s["Region"]
        skolor_r = skolor[skolor["Region"] == region]["Skolenhet"].tolist()

        if len(skolor_r) < 3:
            skolor_r = skolor["Skolenhet"].tolist()

        skolor_r = sorted(skolor_r, key=lambda x: usage[x])

        A = skolor_r[0]
        B = skolor_r[1]
        C = skolor_r[2] if len(skolor_r) > 2 else B

        if program == "LGFRI":
            år1, år2, år3, år4 = A, A, B, ""
        else:
            if region == "Kalmar":
                år1, år2, år3, år4 = A, B, B, C
            else:
                år1, år2, år3, år4 = A, B, A, B

        usage[A] += 1
        usage[B] += 1
        usage[C] += 1

        results.append({
            "Student": s["Namn"],
            "Ort": s["Ort"],
            "Region": region,
            "År1": år1,
            "År2": år2,
            "År3": år3,
            "År4": år4
        })

    df = pd.DataFrame(results)

    # =========================
    # 🚶 PENDLING (FULL)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Skriv studentens namn")

    if student_input:

        input_name = student_input.strip().lower()

        match = df[
            df["Student"].str.lower() == input_name
        ]

        if len(match) == 0:
            st.warning("Student hittades inte")

        else:
            r = match.iloc[0]

            st.write(f"**Bostadsort:** {r['Ort']}")

            for år in ["År1","År2","År3","År4"]:

                if r[år] != "":
                    st.write(f"{år}: {r[år]}")

                    st.radio(
                        f"Pendling OK för {år}?",
                        ["Ja","Nej"],
                        key=f"{r['Student']}_{år}"
                    )

    # =========================
    # EXCEL (VISUELL)
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fills = {
        "header": PatternFill("solid","CCCCCC"),
        "region": PatternFill("solid","D9EAF7"),
        "skola": PatternFill("solid","E7E7E7")
    }

    ws.append(["Skola","År1","År2","År3","År4"])

    for c in range(1,6):
        ws.cell(1,c).fill = fills["header"]
        ws.cell(1,c).font = Font(bold=True)

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])
        ws.merge_cells(start_row=ws.max_row,start_column=1,end_row=ws.max_row,end_column=5)

        for skola in skolor[skolor["Region"] == region]["Skolenhet"]:

            ws.append([f"{skola} (max {kap[skola]})"])

            s1 = df[df["År1"] == skola]["Student"].tolist()
            s2 = df[df["År2"] == skola]["Student"].tolist()
            s3 = df[df["År3"] == skola]["Student"].tolist()
            s4 = df[df["År4"] == skola]["Student"].tolist()

            max_len = max(len(s1),len(s2),len(s3),len(s4),1)

            for i in range(max_len):
                ws.append([
                    "",
                    s1[i] if i < len(s1) else "",
                    s2[i] if i < len(s2) else "",
                    s3[i] if i < len(s3) else "",
                    s4[i] if i < len(s4) else "",
                ])

            ws.append([])

    # auto width
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
