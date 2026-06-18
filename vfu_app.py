
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

    students["Namn"] = students[fn] + " " + students[ln]

    def choose_loc(row):
        if "alternativ" in str(row[pref]).lower():
            return row[alt]
        return row[bost]

    students["ChosenOrt"] = students.apply(choose_loc, axis=1)

    # =========================
    # REGIONVAL (OSÄKER ORT)
    # =========================
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
    # SKAPA PLATSER
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
    # PLACERING (GARANTERAR ALLA)
    # =========================
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        for student in stud_r:

            placed = False

            # ✅ försök hitta ledig plats
            for sk in skolor_r:
                for r in rows_r:
                    if r["Skola"] == sk and r["År1"] == "":
                        r["År1"] = student
                        placed = True
                        break
                if placed:
                    break

            # ✅ fallback (alla får plats)
            if not placed and len(rows_r) > 0:
                rows_r[0]["År1"] = student


    # =========================
    # 🚶 PENDLING (FIXAD)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange studentens namn")

    if student_input:

        match = students[students["Namn"].str.lower() == student_input.lower()]

        if len(match) == 0:
            st.warning("Student hittades inte")

        else:
            sr = match.iloc[0]

            st.write(f"Bostadsort: {sr['ChosenOrt']}")

            for r in rows:
                for year in ["År1","År2","År3","År4"]:
                    if r[year] == sr["Namn"]:

                        st.write(f"{year}: {r['Skola']}")

                        st.radio(
                            "OK pendling?",
                            ["Ja","Nej"],
                            key=f"{student_input}_{year}"
                        )


    # =========================
    # EXCEL
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fills = {
        "Header": PatternFill("solid","CCCCCC"),
        "Region": PatternFill("solid","D9EAF7"),
        "Skola": PatternFill("solid","E7E7E7")
    }

    ws.append(["Skola","År1","År2","År3"])

    for c in range(1,5):
        cell = ws.cell(1,c)
        cell.fill = fills["Header"]
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])
        row = ws.max_row

        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=4)

        for c in range(1,5):
            ws.cell(row,c).fill = fills["Region"]

        ws.append([])

        region_schools = skolor[skolor["Region"]==region]["Skolenhet"]

        for skola in region_schools:

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

    # auto width
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2


    # =========================
    # BLAD 2
    # =========================
    ws2 = wb.create_sheet("Studenter")
    ws2.append(["Student","Ort","År1","År2/3","År4"])

    for _, s in students.iterrows():

        p1=p2=p3=""

        for r in rows:
            if r["År1"]==s["Namn"]: p1=r["Skola"]
            if r["År2"]==s["Namn"]: p2=r["Skola"]
            if r["År4"]==s["Namn"]: p3=r["Skola"]

        ws2.append([s["Namn"],s["ChosenOrt"],p1,p2,p3])
