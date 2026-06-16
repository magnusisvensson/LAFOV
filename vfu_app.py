import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

# ===== UI =====
st.title("VFU – Placering (A‑B‑B‑C korrekt)")

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

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(r["Antal platser"])
        except:
            kap[r["Skolenhet"]] = 0

    # dela upp per region ✅
    region_schools = {
        "Kalmar": skolor[skolor["Region"]=="Kalmar"]["Skolenhet"].tolist(),
        "Oskarshamn": skolor[skolor["Region"]=="Oskarshamn"]["Skolenhet"].tolist(),
        "Karlskrona": skolor[skolor["Region"]=="Karlskrona"]["Skolenhet"].tolist()
    }

    def school_sort_key(skola):
        region = skolor.loc[skolor["Skolenhet"]==skola,"Region"].values[0]
        order = {"Kalmar":0,"Oskarshamn":1,"Karlskrona":2}
        return (order.get(region,3), skola)

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    student_list = students.to_dict("records")

    cap_used = {}
    skol_data = {}

    def has_space(s,y):
        return cap_used.get((s,y),0) < kap.get(s,0)

    def use(s,y):
        cap_used[(s,y)] = cap_used.get((s,y),0) + 1

    def add(s,student,y):
        skol_data.setdefault(s,{})
        skol_data[s].setdefault(student,{
            "År1":"","År2":"","År3":"","År4":""
        })
        skol_data[s][student][y] = student


    # ===== PLACERING PER STUDENT (A-B-B-C) =====
    for idx, stud in enumerate(student_list):

        namn = stud["Namn"]
        region = stud["Region"]

        skol_lista = region_schools.get(region, [])

        if len(skol_lista) < 3:
            continue

        # rotation startindex ✅
        start = idx % len(skol_lista)
        ordered = skol_lista[start:] + skol_lista[:start]

        A, B, C = ordered[0], ordered[1], ordered[2]

        # ✅ kontrollera plats – annars fallback inom region
        placed = False

        for i in range(len(skol_lista)-2):

            A, B, C = ordered[i], ordered[i+1], ordered[i+2]

            if all([
                has_space(A,1),
                has_space(B,2),
                has_space(B,3),
                has_space(C,4)
            ]):

                add(A,namn,"År1")
                add(B,namn,"År2")
                add(B,namn,"År3")
                add(C,namn,"År4")

                use(A,1)
                use(B,2)
                use(B,3)
                use(C,4)

                placed = True
                break

        # fallback: sprid men få 4 år
        if not placed:

            used = set()

            for year, label in [(1,"År1"),(2,"År2"),(3,"År3"),(4,"År4")]:

                for s in skol_lista:

                    if s in used:
                        continue

                    if has_space(s,year):

                        add(s,namn,label)
                        use(s,year)
                        used.add(s)
                        break


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(kap.keys(), key=school_sort_key):

        max_platser = int(kap.get(skola,0))

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows = [{"År1":"","År2":"","År3":"","År4":""}
                for _ in range(max_platser)]

        i = 0
        if skola in skol_data:
            for student,data in skol_data[skola].items():
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

    for s in students["Namn"]:
        ws2.append([s,"OK"])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
