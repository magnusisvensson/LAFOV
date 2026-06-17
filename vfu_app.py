
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering (din modell)")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"]==kull) &
        (skolor["Inriktning"].str.upper()==program)
    ]

    kap = {
        r["Skolenhet"]: int(r["Antal platser"])
        for _, r in skolor.iterrows()
    }

    skol_lista = list(kap.keys())

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    # ===== TRACKING =====
    cap_used = {}

    def has_space(s,y):
        return cap_used.get((s,y),0) < kap.get(s,0)

    def use(s,y):
        cap_used[(s,y)] = cap_used.get((s,y),0)+1


    # ===== ÅR1 =====
    year = {1:{},2:{},3:{},4:{}}

    i = 0
    for skola, max_p in kap.items():
        for _ in range(max_p):
            if i >= len(student_names):
                break

            s = student_names[i]
            year[1][s] = skola
            use(skola,1)
            i += 1


    # ===== ÅR2 =====
    for s in student_names:

        prev = year[1][s]

        if has_space(prev,2):
            year[2][s] = prev
            use(prev,2)
        else:
            for sk in skol_lista:
                if has_space(sk,2):
                    year[2][s] = sk
                    use(sk,2)
                    break


    # ===== ÅR3 =====
    year[3] = year[2].copy()


    # ===== ÅR4 =====
    for s in student_names:

        prev = year[2][s]
        placed = False

        # försök byta
        for sk in skol_lista:
            if sk != prev and has_space(sk,4):
                year[4][s] = sk
                use(sk,4)
                placed = True
                break

        # annars stanna
        if not placed:
            year[4][s] = prev
            use(prev,4)


    # ===== BYGG SKOL-DATA =====
    school_data = {}

    for s in student_names:
        for y in ["År1","År2","År3","År4"]:
            yr = int(y[-1])
            sk = year[yr][s]

            school_data.setdefault(sk,{})
            school_data[sk].setdefault(s,
                {"År1":"","År2":"","År3":"","År4":""}
            )

            school_data[sk][s][y] = s


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola, max_p in kap.items():

        ws.append([f"{skola} (max {max_p})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows=[{"År1":"","År2":"","År3":"","År4":""}
              for _ in range(max_p)]

        i=0
        if skola in school_data:
            for student,data in school_data[skola].items():
                if i>=max_p:
                    break
                rows[i]=data
                i+=1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)
``
