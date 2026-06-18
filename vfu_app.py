import streamlit as st
import pandas as pd
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


if system_file and form_file:

    # =========================
    # LÄS SKOLOR
    # =========================
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)


    # =========================
    # KAPACITET
    # =========================
    kap = {}

    for _, r in skolor.iterrows():
        try:
            val = r["Antal platser"]
            if str(val).strip() in ["", "?", "nan"]:
                raise ValueError
            kap[r["Skolenhet"]] = int(float(val))
        except:
            kap[r["Skolenhet"]] = 2


    # =========================
    # STUDENTER
    # =========================
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    def find_col(keyword):
        return [c for c in students.columns if keyword in c.lower()][0]

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
    # REGIONVAL OM OKLAR
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
    # SKAPA PLATSER
    # =========================
    rows = []

    for _, r in skolor.iterrows():
        for i in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "Student": ""
            })


    # =========================
    # PLACERING (ALLA FÅR PLATS)
    # =========================
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])

        i = 0

        for student in stud_r:

            if len(rows_r) == 0:
                continue

            rows_r[i % len(rows_r)]["Student"] = student
            i += 1


    # =========================
    # PENDLINGSKONTROLL
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Skriv studentens namn")

    if student_input:

        name = student_input.strip().lower()

        match = students[
            students["Namn"].str.lower() == name
        ]

        if len(match) == 0:
            st.warning("Student hittades inte")

        else:
            sr = match.iloc[0]
            st.write(f"Bostadsort: {sr['Ort']}")

            found = False

            for r in rows:
                if r["Student"] == sr["Namn"]:
                    found = True
                    st.write(f"Placering: {r['Skola']}")
                    st.radio(
                        "Fungerar pendling?",
                        ["Ja","Nej"],
                        key=name
                    )

            if not found:
                st.error("Ingen placering hittad")


    # =========================
    # EXCEL EXPORT
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","Student"])

    for skola in skolor["Skolenhet"]:

        ws.append([f"{skola} (max {kap[skola]})"])

        school_rows = [r for r in rows if r["Skola"] == skola]

        if len(school_rows) == 0:
            ws.append(["", ""])
        else:
            for r in school_rows:
                ws.append(["", r["Student"]])

        ws.append([])


    file = "kull_resultat.xlsx"
    wb.save(file)


    # ✅ EXPORTKNAPP (GARANTERAD)
    with open(file,"rb") as f:
        st.download_button(
            "⬇️ Ladda ner Excel",
            f,
            file_name=file
        )
