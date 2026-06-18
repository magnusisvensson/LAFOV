
import streamlit as st
import pandas as pd
olor["Kull"] == kull) &from openpyxl import Workbook
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ===== KAP =====
    kap = {}
    kap_osaker = {}

    for _, r in skolor.iterrows():
        try:
            val = r["Antal platser"]
            if str(val).strip() in ["", "?", "nan"]:
                raise ValueError
            kap[r["Skolenhet"]] = int(float(val))
            kap_osaker[r["Skolenhet"]] = False
        except:
            kap[r["Skolenhet"]] = 2
            kap_osaker[r["Skolenhet"]] = True


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
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

    students["ChosenOrt"] = students.apply(choose_loc, axis=1)

    # ===== REGION =====
    regions = []

    for _, row in students.iterrows():
        region = get_region(row["ChosenOrt"])

        if region is None:
            region = st.selectbox(
                f"Välj region för {row['Namn']} ({row['ChosenOrt']})",
                ["Kalmar","Karlskrona","Oskarshamn"],
                key=row["Namn"]
            )

        regions.append(region)

    students["Region"] = regions


    # ===== SKAPA PLATSER =====
    rows = []

    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1":"","År2":"","År3":"","År4":""
            })


    # ===== BÄTTRE PLACERING =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])

        for student in stud_r:

            # välj minst fyllda plats
            sorted_rows = sorted(
                rows_r,
                key=lambda x: (x["År1"] != "", x["Skola"])
            )

            placed = False

            for r in sorted_rows:
                if r["År1"] == "":
                    r["År1"] = student
                    placed = True
                    break

            if not placed:
                rows_r[0]["År1"] = student


    # =========================
    # 🚶 PENDLING (TYDLIG)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange studentens namn")

    if student_input:

        input_name = student_input.strip().lower()

        match = students[
            students["Namn"].str.lower() == input_name
        ]

        if len(match) == 0:
            st.warning("Student hittades inte")

        else:
            sr = match.iloc[0]

            st.write(f"**Bostadsort:** {sr['ChosenOrt']}")

            found = False

            for r in rows:
                if r["År1"] == sr["Namn"]:
                    found = True
                    st.write(f"År1: {r['Skola']}")
                    st.radio(
                        "Pendling OK?",
                        ["Ja","Nej"],
                        key=f"{input_name}_År1"
                    )

            if not found:
                st.error("Ingen placering hittad")


    # =========================
    # EXCEL (FIXAD STRUKTUR)
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3"])

    for skola in skolor["Skolenhet"]:

        label = f"{skola} (max {kap[skola]})"
        if kap_osaker[skola]:
            label += " ⚠ osäker"

        ws.append([label])

        school_rows = [r for r in rows if r["Skola"] == skola]

        if not school_rows:
            ws.append(["","","",""])
        else:
            for r in school_rows:
                ws.append(["", r["År1"], r["År2"], r["År3"]])

        ws.append([])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
``
from openpyxl.styles import Font, PatternFill, Alignment


st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


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

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
