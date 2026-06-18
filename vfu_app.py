
import streamlit as st
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


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

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ===== kapacitet =====
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

    # ===== studenter =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    def find_col(k):
        return [c for c in students.columns if k in c.lower()][0]

    fn = find_col("förnamn")
    ln = find_col("efternamn")
    bost = find_col("bostads")
    alt = find_col("alternativ")
    pref = find_col("utgå")

    students["Namn"] = students[fn] + " " + students[ln]

    def choose_loc(row):
        if "alternativ" in str(row[pref]).lower():
            return row[alt]
        return row[bost]

    students["ChosenOrt"] = students.apply(choose_loc, axis=1)

    # ===== regionval =====
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

    # ===== platser =====
    rows = []

    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1":"","År2":"","År3":"","År4":""
            })

    # ===== placering =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"]==region]
        stud_r = list(students[students["Region"]==region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        i = 0

        for student in stud_r:

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)]

            for r in rows_r:
                if r["Skola"] == A and r["År1"]=="":
                    r["År1"]=student
                    break

            if program=="LGFRI":
                for r in rows_r:
                    if r["Skola"]==A and r["År2"]=="":
                        r["År2"]=student
                        break
                for r in rows_r:
                    if r["Skola"]==B and r["År3"]=="":
                        r["År3"]=student
                        break
            else:
                for r in rows_r:
                    if r["Skola"]==B and r["År2"]=="":
                        r["År2"]=student
                        r["År3"]=student
                        break

            i += 1

    # =========================
    # ✅ PENDLING (FIXAD)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange studentens namn")

    if student_input:

        match = students[students["Namn"].str.lower() == student_input.lower()]

        if len(match)==0:
            st.warning("Student hittades inte")

        else:
            sr = match.iloc[0]

            bostad = sr["ChosenOrt"]
            st.write(f"Bostadsort: {bostad}")

            for r in rows:
                for year in ["År1","År2","År3","År4"]:
                    if r[year] == sr["Namn"]:

                        st.write(f"{year}: {r['Skola']}")

                        st.radio(
                            f"OK pendling?",
                            ["Ja","Nej"],
                            key=f"{student_input}_{year}"
                        )

    # =========================
    # EXCEL (3 FLIKAR)
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3"])

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])

        for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

            label = f"{skola} (max {kap[skola]})"

            if kap_osaker[skola]:
                label += " ⚠ osäker"

            ws.append([label])

            school_rows = [r for r in rows if r["Skola"]==skola]

            for r in school_rows:
                ws.append(["",r["År1"],r["År2"],r["År3"]])

            ws.append([])

    # ===== BLAD 2 =====
    ws2 = wb.create_sheet("Studenter")

    ws2.append(["Student","Ort","År1","År2/3","År4"])

    for _, s in students.iterrows():

        p1=p2=p3=""

        for r in rows:
            if r["År1"]==s["Namn"]: p1=r["Skola"]
            if r["År2"]==s["Namn"]: p2=r["Skola"]
            if r["År4"]==s["Namn"]: p3=r["Skola"]

        ws2.append([s["Namn"],s["ChosenOrt"],p1,p2,p3])

    # ===== BLAD 3 =====
    ws3 = wb.create_sheet("Kontroll")

    ws3.append(["Student","Antal skolor"])

    for _, s in students.iterrows():

        skolset = set()

        for r in rows:
            if s["Namn"] in [r["År1"],r["År2"],r["År3"],r["År4"]]:
                skolset.add(r["Skola"])

        ws3.append([s["Namn"],len(skolset)])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("Ladda ner Excel",f,file_name=file)
