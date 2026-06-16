import streamlit as stimport streamlit asst.title("VFU – Placering")

# ==== UPLOAD ====
system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"


def school_region(partner):
    p = str(partner).lower()
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"


# ===== MAIN =====
if system_file and form_file:

    # ---- SKOLOR ----
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    # kapacitet (robust)
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(r["Antal platser"])
        except:
            kap[r["Skolenhet"]] = 0

    # sorteringsordning
    def school_sort_key(skola):
        region = skolor.loc[skolor["Skolenhet"] == skola, "Region"].values[0]
        order = {"Kalmar":0, "Oskarshamn":1, "Karlskrona":2}
        return (order.get(region,3), skola)

    # ---- STUDENTER ----
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]

    student_names = list(students["Namn"])
    skol_lista = list(skolor["Skolenhet"])

    # ===== ÅR-FÖRDELNING =====
    cap_used = {}
    skol_data = {}

    def has_space(s, year):
        return cap_used.get((s,year),0) < kap.get(s,0)

    def use(s, year):
        cap_used[(s,year)] = cap_used.get((s,year),0)+1

    def add(s, student, year):
        skol_data.setdefault(s,{})
        skol_data[s].setdefault(student,{
            "År1":"","År2":"","År3":"","År4":""
        })
        skol_data[s][student][f"År{year}"] = student

    # ✅ viktig: FÖR VARJE ÅR → placera ALLA studenter
    for year in [1,2,3,4]:

        for student in student_names:

            placed = False

            for s in skol_lista:
                if has_space(s, year):
                    add(s, student, year)
                    use(s, year)
                    placed = True
                    break

            if not placed:
                # om ingen plats finns → hoppa (eller markera)
                pass

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(skol_lista, key=school_sort_key):

        max_platser = int(kap.get(skola,0))

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)

        rows = [
            {"År1":"","År2":"","År3":"","År4":""}
            for _ in range(max_platser)
        ]

        i = 0
        if skola in skol_data:
            for student, data in skol_data[skola].items():
                if i >= max_platser:
                    break
                rows[i] = data
                i += 1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in student_names:
        ws2.append([s,"OK"])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp båda filer")
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font


