import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering (stabil version)")

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


def school_region(partner):
    p = str(partner).lower()
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"


# ===== MAIN =====
if system_file and form_file:

    # ----- SKOLOR -----
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    # ✅ säker kapacitet (fixar crash)
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(r["Antal platser"])
        except:
            kap[r["Skolenhet"]] = 0

    # ----- STUDENTER -----
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    skol_data = {}
    cap_used = {}
    logg = {}

    # ===== HELP =====
    def has_space(skola, år):
        return cap_used.get((skola, år), 0) < kap.get(skola, 0)

    def use(skola, år):
        cap_used[(skola, år)] = cap_used.get((skola, år), 0) + 1

    def add(skola, student, år):
        skol_data.setdefault(skola, {})
        if student not in skol_data[skola]:
            skol_data[skola][student] = {
                "År1": "", "År2": "", "År3": "", "År4": ""
            }
        skol_data[skola][student][år] = student


    # ===== PLACERING =====
    def place_student(student):

        namn = student["Namn"]
        region = student["Region"]

        skol_lista = list(skolor["Skolenhet"])

        # ------ 1. PERFEKT ROTATION ------
        for i in range(len(skol_lista) - 2):

            A, B, C = skol_lista[i], skol_lista[i+1], skol_lista[i+2]

            if region in ["Oskarshamn", "Karlskrona"]:

                if all([
                    has_space(A,1),
                    has_space(B,2),
                    has_space(A,3),
                    has_space(B,4)
                ]):

                    add(A,namn,"År1")
                    add(B,namn,"År2")
                    add(A,namn,"År3")
                    add(B,namn,"År4")

                    use(A,1); use(B,2); use(A,3); use(B,4)

                    logg[namn] = {"Status":"OK"}
                    return

            else:

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

                    use(A,1); use(B,2); use(B,3); use(C,4)

                    logg[namn] = {"Status":"OK"}
                    return

        # ------ 2. FALLBACK (FYLL ALLTID) ------
        år_lista = [("År1",1),("År2",2),("År3",3),("År4",4)]

        for årnamn, årnr in år_lista:
            placerad = False
            for skola in skol_lista:
                if has_space(skola, årnr):
                    add(skola, namn, årnamn)
                    use(skola, årnr)
                    placerad = True
                    break

            if not placerad:
                logg[namn] = {"Status":"Får ej plats"}
                return

        logg[namn] = {"Status":"OK*"}  # fallback


    # ===== KÖR ALLA =====
    for stud in students.to_dict("records"):
        place_student(stud)


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    ws.append(["Skola","År1","År2","År3","År4"])

    header_fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    for skola in skol_data:

        max_platser = int(kap.get(skola, 0))

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = header_fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)

        rows = [
            {"År1":"","År2":"","År3":"","År4":""}
            for _ in range(max_platser)
        ]

        idx = 0
        for student, data in skol_data[skola].items():
            if idx >= max_platser:
                break
            rows[idx] = data
            idx += 1

        for rad in rows:
            ws.append([
                "",
                rad["År1"],
                rad["År2"],
                rad["År3"],
                rad["År4"]
            ])

        ws.append([])


    # ---- RAPPORT ----
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s, v in logg.items():
        ws2.append([s, v["Status"]])


    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file, "rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
