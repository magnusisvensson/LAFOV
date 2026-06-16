import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – korrekt årsplacering")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])

# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"

if system_file and form_file:

    # ---- SKOLOR ----
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"]==kull) &
        (skolor["Inriktning"].str.upper()==program)
    ].copy()

    skol_lista = list(skolor["Skolenhet"])

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(r["Antal platser"])
        except:
            kap[r["Skolenhet"]] = 0

    # ---- STUDENTER ----
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn=[c for c in students.columns if "förnamn" in c.lower()][0]
    ln=[c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn]+" "+students[ln]

    student_names = list(students["Namn"])

    # ===== ÅR-FÖRDELNING =====
    year_data = {
        1: [],
        2: [],
        3: [],
        4: []
    }

    cap_used = {}

    def has_space(s, y):
        return cap_used.get((s,y),0) < kap.get(s,0)

    def use(s, y):
        cap_used[(s,y)] = cap_used.get((s,y),0)+1

    # ✅ FÖR VARJE ÅR – PLACERA ALLA
    for year in [1,2,3,4]:

        for student in student_names:

            placed = False

            for s in skol_lista:

                if has_space(s, year):

                    year_data[year].append((student, s))
                    use(s, year)

                    placed = True
                    break

            if not placed:
                year_data[year].append((student, "SAKNAR PLATS"))

    # ===== TRANSFORMERA TILL SKOLSTRUKTUR =====
    skol_data = {}

    for year in [1,2,3,4]:

        for student, skola in year_data[year]:

            if skola == "SAKNAR PLATS":
                continue

            skol_data.setdefault(skola, {})
            skol_data[skola].setdefault(student, {
                "År1":"","År2":"","År3":"","År4":""
            })

            skol_data[skola][student][f"År{year}"] = student


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    ws.append(["Skola","År1","År2","År3","År4"])

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    for skola in skol_lista:

        max_platser = kap.get(skola,0)

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        # skapa rader
        rows = [
            {"År1":"","År2":"","År3":"","År4":""}
            for _ in range(max_platser)
        ]

        placed_rows = 0

        if skola in skol_data:

            for student, data in skol_data[skola].items():

                if placed_rows >= max_platser:
                    break

                rows[placed_rows] = data
                placed_rows += 1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp filer")
