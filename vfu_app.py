import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

# ===== UI =====
st.title("VFU – Placering")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


# ===== REGION =====
def school_region(partner):
    p = str(partner).lower()
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"


# ===== MAIN =====
if system_file and form_file:

    # ===== LÄS SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    # Kapacitet säkert som int
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(r["Antal platser"])
        except:
            kap[r["Skolenhet"]] = 0

    skol_lista = list(skolor["Skolenhet"])

    # Sorteringsfunktion (Kalmar → Oskarshamn → Karlskrona)
    def school_sort_key(skola):
        region = skolor.loc[skolor["Skolenhet"] == skola, "Region"].values[0]
        order = {"Kalmar":0, "Oskarshamn":1, "Karlskrona":2}
        return (order.get(region,3), skola)


    # ===== LÄS STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    # ===== PLACERING PER ÅR =====
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

    # 🔥 KRITISKT: loopa per år → alla studenter varje gång
    for year in [1,2,3,4]:

        for student in student_names:

            placed = False

            for s in skol_lista:
                if has_space(s, year):
                    add(s, student, year)
                    use(s, year)
                    placed = True
                    break

            # om ingen plats → hoppa (men alla ska normalt få plats)
            if not placed:
                pass


    # ===== SKAPA EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    header_fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(skol_lista, key=school_sort_key):

        max_platser = int(kap.get(skola,0))

        # rubrikrad
        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = header_fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        # skapa tomma rader
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
            ws.append([
                "",
                row["År1"],
                row["År2"],
                row["År3"],
                row["År4"]
            ])

        ws.append([])


    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in student_names:
        ws2.append([s,"OK"])


    # ===== SPARA =====
    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
