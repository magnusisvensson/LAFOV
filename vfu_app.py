import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV", "LAGRV", "LGFRI"])


# =========================
# REGION
# =========================
def get_region(text):
    t = str(text).lower()

    if "oskarshamn" in t:
        return "Oskarshamn"

    if any(x in t for x in ["karlskrona", "ronneby", "rödeby"]):
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

    students["Ort"] = students.apply(choose_loc, axis=1)

    # =========================
    # REGIONVAL
    # =========================
    regions = []

    for _, row in students.iterrows():
        region = get_region(row["Ort"])

        if region is None:
            region = st.selectbox(
                f"Välj region för {row['Namn']} ({row['Ort']})",
                ["Kalmar", "Karlskrona", "Oskarshamn"],
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
                "År1": "",
                "År2": "",
                "År3": "",
                "År4": ""
            })

    # =========================
    # PLACERING (KORREKT LOGIK)
    # =========================
    for region in ["Kalmar", "Oskarshamn", "Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list({r["Skola"] for r in rows_r})

        for i, student in enumerate(stud_r):

            if not skolor_r:
                continue

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)]
            C = skolor_r[(i+2) % len(skolor_r)]

            # År1
            for r in rows_r:
                if r["Skola"] == A and r["År1"] == "":
                    r["År1"] = student
                    break

            if program == "LGFRI":
                # A A B
                for r in rows_r:
                    if r["Skola"] == A and r["År2"] == "":
                        r["År2"] = student
                        break
                for r in rows_r:
                    if r["Skola"] == B and r["År3"] == "":
                        r["År3"] = student
                        break

            else:
                if region == "Kalmar":
                    # A B B C
                    for r in rows_r:
                        if r["Skola"] == B and r["År2"] == "":
                            r["År2"] = student
                            r["År3"] = student
                            break

                    for r in rows_r:
                        if r["Skola"] == C and r["År4"] == "":
                            r["År4"] = student
                            break

                else:
                    # ABAB
                    for r in rows_r:
                        if r["Skola"] == B and r["År2"] == "":
                            r["År2"] = student
                            break
                    for r in rows_r:
                        if r["Skola"] == A and r["År3"] == "":
                            r["År3"] = student
                            break
                    for r in rows_r:
                        if r["Skola"] == B and r["År4"] == "":
                            r["År4"] = student
                            break

    # =========================
    # EXCEL
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fills = {
        "header": PatternFill("solid", "CCCCCC"),
        "region": PatternFill("solid", "D9EAF7"),
        "skola": PatternFill("solid", "E7E7E7"),
    }

    ws.append(["Skola", "År1", "År2", "År3", "År4"])

    for c in range(1, 6):
        ws.cell(1, c).fill = fills["header"]
        ws.cell(1, c).font = Font(bold=True)

    for region in ["Kalmar", "Oskarshamn", "Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])

        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=5)

        for c in range(1, 6):
            ws.cell(ws.max_row, c).fill = fills["region"]

        ws.append([])

        for skola in skolor[skolor["Region"] == region]["Skolenhet"]:

            label = f"{skola} (max {kap[skola]})"
            if kap_osaker[skola]:
                label += " ⚠"

            ws.append([label])

            school_rows = [r for r in rows if r["Skola"] == skola]

            if not school_rows:
                ws.append(["", "", "", "", ""])
            else:
                for r in school_rows:
                    ws.append(["", r["År1"], r["År2"], r["År3"], r["År4"]])

            ws.append([])

    # auto width
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file, "rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
