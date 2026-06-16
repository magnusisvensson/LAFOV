import streamlit as stimport streamlit as as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])

# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

def school_region(partner, skola):
    p = str(partner).lower()
    s = str(skola).lower()

    if "ljungnäs" in s or "blomstermåla" in s:
        return "Kalmar"
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"

# ===== MAIN =====
if system_file and form_file:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"]==kull) &
        (skolor["Inriktning"].str.upper()==program)
    ].copy()

    skolor["Region"] = skolor.apply(
        lambda r: school_region(r["Partnerområde"], r["Skolenhet"]), axis=1
    )

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
    placed = set()

    def has_space(s,year):
        return cap_used.get((s,year),0) < kap.get(s,999)

    def use(s,year):
        cap_used[(s,year)] = cap_used.get((s,year),0)+1

    def add(s,student,year):
        skol_data.setdefault(s,{})
        skol_data[s].setdefault(student,{
            "År1":"","År2":"","År3":"","År4":""
        })
        skol_data[s][student][year]=student

    # ===== PLACERA EN STUDENT =====
    def place_one(student):

        if student["Namn"] in placed:
            return False

        region = student["Region"]

        regional = skolor[skolor["Region"]==region]["Skolenhet"].tolist()
        alla = list(skolor["Skolenhet"])
        lista = regional + [x for x in alla if x not in regional]

        for i in range(len(lista)-2):

            A,B,C = lista[i],lista[i+1],lista[i+2]

            # specialregion
            if region in ["Oskarshamn","Karlskrona"]:
                if all([
                    has_space(A,1),
                    has_space(B,2),
                    has_space(A,3),
                    has_space(B,4)
                ]):
                    add(A,student["Namn"],"År1")
                    add(B,student["Namn"],"År2")
                    add(A,student["Namn"],"År3")
                    add(B,student["Namn"],"År4")

                    use(A,1); use(B,2); use(A,3); use(B,4)
                    placed.add(student["Namn"])
                    logg[student["Namn"]]={"Status":"OK"}
                    return True

            else:
                if all([
                    has_space(A,1),
                    has_space(B,2),
                    has_space(B,3),
                    has_space(C,4)
                ]):
                    add(A,student["Namn"],"År1")
                    add(B,student["Namn"],"År2")
                    add(B,student["Namn"],"År3")
                    add(C,student["Namn"],"År4")

                    use(A,1); use(B,2); use(B,3); use(C,4)
                    placed.add(student["Namn"])
                    logg[student["Namn"]]={"Status":"OK"}
                    return True

        return False

    # ===== STRATEGI =====
    stud_list = students.to_dict("records")

    # först försök grupper
    i = 0
    while i < len(stud_list):

        group = stud_list[i:i+3]

        if len(group)==3 and all(s["Namn"] not in placed for s in group):
            if all(place_one(s) for s in group):
                i += 3
                continue

        group = stud_list[i:i+2]

        if len(group)==2 and all(s["Namn"] not in placed for s in group):
            if all(place_one(s) for s in group):
                i += 2
                continue

        # fallback individ (VIKTIGASTE!)
        if place_one(stud_list[i]):
            i += 1
        else:
            logg[stud_list[i]["Namn"]]={"Status":"Får ej plats"}
            i += 1

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(skol_data):

        ws.append([skola])

        for student,data in skol_data[skola].items():
            ws.append([
                "",
                data["År1"],
                data["År2"],
                data["År3"],
                data["År4"]
            ])

        ws.append([])

    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s,v in logg.items():
        ws2.append([s,v["Status"]])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
