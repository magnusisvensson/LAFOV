
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook

        region = s["Region"]
        skolor_r = skolor[skolor["Region"] == region]["Skolenhet"].tolist()

        # fallback om för få skolor
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
    # 🚶 PENDLINGSKONTROLL
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange student")

    if student_input:

        name = student_input.strip().lower()

        match = df[df["Student"].str.lower() == name]

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
                        key=f"{name}_{year}"
                    )


    # =========================
    # ✅ EXCEL (KORREKT VISUELL)
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])

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

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
from collections import defaultdict
from openpyxl import Workbook


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
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ===== KAPACITET =====
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
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostads" in c.lower()][0]

    students["Namn"] = (students[fn] + " " + students[ln]).str.strip()
    students["Ort"] = students[bost]

    # =========================
    # ✅ REGIONVAL VID OKLAR ORT
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
    # ✅ PLACERING (KORREKT)
    # =========================
    usage = defaultdict(int)
    results = []

    for _, s in students.iterrows():

