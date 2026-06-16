import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering (Stabil version)")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])

# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

def school_region(partner):
    p = str(partner).lower()
    if "oskarshamn" in p: return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p: return "Karlskrona"
    return "Kalmar"

# ===== MAIN =====
if system_file and form_file:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"]==kull) &
        (skolor["Inriktning"].str.upper()==program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    kap = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn=[c for c in students.columns if "förnamn" in c.lower()][0]
    ln=[c for c in students.columns if "efternamn" in c.lower()][0]
    bost=[c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn]+" "+students[ln]
    students["Region"] = students[bost].apply(get_region)

    cap_used = {}
    skol_data = {}
    logg = {}

    # ===== HELP =====
    def has_space(s,år):
        return cap_used.get((s,år),0) < kap.get(s,999)

    def use(s,år):
        cap_used[(s,år)] = cap_used.get((s,år),0)+1

    def add(s,student,år):
        skol_data.setdefault(s,{})
        skol_data[s].setdefault(student,{
            "År1":"","År2":"","År3":"","År4":""
        })
        skol_data[s][student][år]=student

    # ===== PLACERING =====
    def place_student(stud):

        namn = stud["Namn"]
        region = stud["Region"]

        lista = list(skolor["Skolenhet"])

        # ===== 1. STRIKT ROTATION =====
        for i in range(len(lista)-2):

            A,B,C = lista[i], lista[i+1], lista[i+2]

            if region in ["Oskarshamn","Karlskrona"]:
                if all([has_space(A,1),has_space(B,2),has_space(A,3),has_space(B,4)]):

                    add(A,namn,"År1")
                    add(B,namn,"År2")
                    add(A,namn,"År3")
                    add(B,namn,"År4")

                    use(A,1); use(B,2); use(A,3); use(B,4)
                    logg[namn]={"Status":"OK"}
                    return True

            else:
                if all([has_space(A,1),has_space(B,2),has_space(B,3),has_space(C,4)]):

                    add(A,namn,"År1")
                    add(B,namn,"År2")
                    add(B,namn,"År3")
                    add(C,namn,"År4")

                    use(A,1); use(B,2); use(B,3); use(C,4)
                    logg[namn]={"Status":"OK"}
                    return True


        # ===== 2. FYLL PER ÅR (VIKTIGASTE FIXEN) =====
        år_lista = ["År1","År2","År3","År4"]

        for idx,år in enumerate(år_lista, start=1):
            for s in lista:
                if has_space(s,idx):
                    add(s,namn,år)
                    use(s,idx)
                    break

        logg[namn]={"Status":"OK*"}  # markera fallback
        return True


    # ===== KÖR =====
    for stud in students.to_dict("records"):
        place_student(stud)

    # ===== EXCEL (FORMATERAD) =====
    wb = Workbook()
    ws = wb.active

    header_fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in skol_data:

        ws.append([f"{skola} (max {kap.get(skola,0)})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = header_fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        max_platser = kap.get(skola,0)

        rows = [{"År1":"","År2":"","År3":"","År4":""}
                for _ in range(max_platser)]

        idx = 0
        for student,data in skol_data[skola].items():
            if idx >= max_platser:
                break
            rows[idx] = data
            idx += 1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s,v in logg.items():
        ws2.append([s,v["Status"]])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp filer")
