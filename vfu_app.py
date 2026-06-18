
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


# =========================
# START
# =========================
if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file, engine="openpyxl")
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ===== KAP =====
    kap = {}
    for _, r in skolor.iterrows():
        try:
            val = r["Antal platser"]
            if str(val).strip() in ["", "?", "nan"]:
                raise ValueError
            kap[r["Skolenhet"]] = int(float(val))
        except:
            kap[r["Skolenhet"]] = 2

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data", engine="openpyxl")
    students.columns = students.columns.str.strip()

    def find_col(k):
        return [c for c in students.columns if k in c.lower()][0]

    fn = find_col("förnamn")
    ln = find_col("efternamn")
    bost = find_col("bostads")
    alt = find_col("alternativ")
    pref = find_col("utgå")

    students["Namn"] = (students[fn] + " " + students[ln]).str.strip()

    def choose_loc(row):
        if "alternativ" in str(row[pref]).lower():
            return row[alt]
        return row[bost]

    students["Ort"] = students.apply(choose_loc, axis=1)

    # =========================
    # REGIONVAL (VIKTIG!)
    # =========================
    regions = []

    for _, row in students.iterrows():

        region = get_region(row["Ort"])

        if region is None:
            region = st.selectbox(
                f"Välj region för {row['Namn']} ({row['Ort']})",
                ["Kalmar","Karlskrona","Oskarshamn"],
                key=row["Namn"]
            )

        regions.append(region)

    students["Region"] = regions

    # =========================
    # ✅ PLACERING
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
    # 🚶 PENDLING
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange student")

    if student_input:

        match = df[df["Student"].str.lower() == student_input.strip().lower()]

        if len(match) == 0:
            st.warning("Student hittades inte")

        else:
            r = match.iloc[0]
            st.write(f"Bostadsort: {r['Ort']}")

            for year in ["År1","År2","År3","År4"]:
                if r[year] != "":
                    st.write(f"{year}: {r[year]}")
                    st.radio(
                        f"Pendling OK ({year})?",
                        ["Ja","Nej"],
                        key=f"{r['Student']}_{year}"
                    )

    # =========================
    # 📊 EXCEL
    # =========================
    wb = Workbook()

    # ===== BLAD 1 =====
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])
        ws.merge_cells(start_row=ws.max_row,start_column=1,end_row=ws.max_row,end_column=5)

        skolor_r = skolor[skolor["Region"] == region]["Skolenhet"]

        for skola in skolor_r:

            ws.append([f"{skola} (max {kap[skola]})"])

            subset = df[
                (df["År1"] == skola) |
                (df["År2"] == skola) |
                (df["År3"] == skola) |
                (df["År4"] == skola)
            ]

            if subset.empty:
                ws.append(["","","","",""])
            else:
                for _, s in subset.iterrows():
                    ws.append([
                        "",
                        s["Student"] if s["År1"] == skola else "",
                        s["Student"] if s["År2"] == skola else "",
                        s["Student"] if s["År3"] == skola else "",
                        s["Student"] if s["År4"] == skola else "",
                    ])

            ws.append([])

    # ===== BLAD 2 =====
    ws2 = wb.create_sheet("Studenter")
    ws2.append(["Student","Ort","År1","År2","År3","År4"])

    for _, r in df.iterrows():
        ws2.append([r["Student"],r["Ort"],r["År1"],r["År2"],r["År3"],r["År4"]])

    # ===== BLAD 3 =====
    ws3 = wb.create_sheet("Kontroll")
    ws3.append(["Student","Antal skolor"])

    for _, r in df.iterrows():
        skolset = {r["År1"],r["År2"],r["År3"],r["År4"]}
        skolset.discard("")
        ws3.append([r["Student"],len(skolset)])

    # auto width
    for ws_ in wb.worksheets:
        for col in ws_.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws_.column_dimensions[col[0].column_letter].width = max_len + 2

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
