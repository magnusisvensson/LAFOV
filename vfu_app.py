
import streamlit as st
import pandas as pd
from openpyxl import Workbook
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

    # =========================
    # DATA
    # =========================
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # kapacitet
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


    # =========================
    # STUDENTER
    # =========================
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

    # region
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


    # =========================
    # PLATSER
    # =========================
    rows = []

    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1":"","År2":"","År3":"","År4":""
            })


    # =========================
    # PLACERING (alla får plats)
    # =========================
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"]==region]
        stud_r = list(students[students["Region"]==region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        for student in stud_r:

            placed = False

            for sk in skolor_r:
                for r in rows_r:
                    if r["Skola"] == sk and r["År1"] == "":
                        r["År1"] = student
                        placed = True
                        break
                if placed:
                    break

            if not placed and len(rows_r) > 0:
                rows_r[0]["År1"] = student


    # =========================
    # 🚶 PENDLING (FIXAD)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange studentens namn")

    if student_input:

        input_name = student_input.strip().lower()

        match = students[
            students["Namn"].str.strip().str.lower() == input_name
        ]

        if len(match) == 0:
            st.warning("Student hittades inte")

        else:
            sr = match.iloc[0]

            st.write(f"**Bostadsort:** {sr['ChosenOrt']}")

            found = False

            for r in rows:
                for year in ["År1","År2","År3","År4"]:
                    if r[year] == sr["Namn"]:

                        found = True

                        st.write(f"{year}: {r['Skola']}")

                        st.radio(
                            f"Pendling OK för {year}?",
                            ["Ja","Nej"],
                            key=f"{input_name}_{year}"
                        )

            if not found:
                st.error("Studenten har ingen placering ännu")


    # =========================
    # EXCEL
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

        school_rows = [r for r in rows if r["Skola"]==skola]

        if len(school_rows) == 0:
            ws.append(["","","",""])
        else:
            for r in school_rows:
                ws.append(["",r["År1"],r["År2"],r["År3"]])

        ws.append([])


    file="kull_resultat.xlsx"
    wb.save(file)

    # ✅ EXPORTKNAPP (GARANTERAD)
    with open(file,"rb") as f:
        st.download_button(
            "⬇️ Ladda ner Excel",
            f,
            file_name=file
        )
