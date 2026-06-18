
import streamlit as st
import pandas as in t:import pandas as pd
        return "Oskarshamn"

    if any(x in t for x in ["karlskrona","ronneby","rödeby"]):
        return "Karlskrona"

    if "kalmar" in t:
        return "Kalmar"

    return None


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

    students["Ort"] = students.apply(choose_loc, axis=1)

    # ===== REGIONVAL =====
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

    # ===== SKAPA PLATSER =====
    rows = []

    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "Student": ""
            })

    # ===== BALANSERAD PLACERING =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])

        for student in stud_r:

            # välj minst fylld skola
            counts = {
                sk: sum(1 for r in rows_r if r["Skola"] == sk and r["Student"] != "")
                for sk in set(r["Skola"] for r in rows_r)
            }

            skola_vald = min(counts, key=counts.get)

            for r in rows_r:
                if r["Skola"] == skola_vald and r["Student"] == "":
                    r["Student"] = student
                    break

    # =========================
    # 🚶 PENDLING
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
            st.write(f"**Bostadsort:** {sr['Ort']}")

            for r in rows:
                if r["Student"] == sr["Namn"]:
                    st.write(f"Skola: {r['Skola']}")
                    st.radio(
                        "Fungerar pendling?",
                        ["Ja","Nej"],
                        key=f"{name}_{r['Skola']}"
                    )

    # =========================
    # EXCEL
    # =========================
    wb = Workbook()

    # ===== Blad 1 =====
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","Student"])

    header_fill = PatternFill("solid","CCCCCC")

    for c in range(1,3):
        ws.cell(1,c).fill = header_fill
        ws.cell(1,c).font = Font(bold=True)

    for skola in skolor["Skolenhet"]:

        label = f"{skola} (max {kap[skola]})"
        if kap_osakerlabel += " ⚠ osäker"

        ws.append([label])

        school_rows = [r for r in rows if r["Skola"] == skola]

        if not school_rows:
            ws.append(["",""])
        else:
            for r in school_rows:
                ws.append(["", r["Student"]])

        ws.append([])

    # ===== Blad 2 =====
    ws2 = wb.create_sheet("Studenter")
    ws2.append(["Student","Ort","Placering"])

    for _, s in students.iterrows():
        placering = next((r["Skola"] for r in rows if r["Student"]==s["Namn"]), "")
        ws2.append([s["Namn"],s["Ort"],placering])

    # ===== Blad 3 =====
    ws3 = wb.create_sheet("Kontroll")
    ws3.append(["Student","Antal","Status"])

    for _, s in students.iterrows():

        skolset = set(
            r["Skola"] for r in rows if r["Student"] == s["Namn"]
        )

        antal = len(skolset)

        if antal == 0:
            status="SAKNAR"; color="FFCCCC"
        elif antal == 1:
            status="EN"; color="FFD9B3"
        else:
            status="OK"; color="CCFFCC"

        ws3.append([s["Namn"],antal,status])

        for c in range(1,4):
            ws3.cell(ws3.max_row,c).fill = PatternFill("solid",color)

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
``
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


def get_region(text):
    t = str(text).lower()

