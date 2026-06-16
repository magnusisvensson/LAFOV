import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering")

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

# ===== SORTERING =====
def school_sort_key(skola):
    region = skolor.loc[skolor["Skolenhet"]==skola,"Region"].values[0]
    order = {"Kalmar":0,"Oskarshamn":1,"Karlskrona":2}
    return (order.get(region,3), skola)

# ===== AVSTÅND (enkel region-baserad) =====
geo = {
    "Kalmar": (56.66,16.36),
    "Oskarshamn": (57.26,16.45),
    "Karlskrona": (56.16,15.59)
}

def dist(a,b):
    if a not in geo or b not in geo:
        return 0
    d=((geo[a][0]-geo[b][0])**2+(geo[a][1]-geo[b][1])**2)**0.5
    return round(d*111,1)


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
    bost=[c for c in students.columns if "bostadsort" in c.lower()][0]

    # ev alternativ bostad
    alt_cols = [c for c in students.columns if "alternativ" in c.lower()]
    alt_col = alt_cols[0] if alt_cols else None

    students["Namn"] = students[fn]+" "+students[ln]
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

        if skol_data[skola][student][år] == "":
            skol_data[skola][student][år] = student

    # ===== PLACERING =====
    def place_student(student):

        namn = student["Namn"]
        region = student["Region"]

        skol_lista = list(skolor["Skolenhet"])

        # --- 1 rotation ---
        for i in range(len(skol_lista)-2):
            A,B,C = skol_lista[i],skol_lista[i+1],skol_lista[i+2]

            if region in ["Oskarshamn","Karlskrona"]:
                if all([has_space(A,1),has_space(B,2),has_space(A,3),has_space(B,4)]):

                    add(A,namn,"År1")
                    add(B,namn,"År2")
                    add(A,namn,"År3")
                    add(B,namn,"År4")

                    use(A,1); use(B,2); use(A,3); use(B,4)
                    logg[namn] = {"Status":"OK"}
                    return
            else:
                if all([has_space(A,1),has_space(B,2),has_space(B,3),has_space(C,4)]):

                    add(A,namn,"År1")
                    add(B,namn,"År2")
                    add(B,namn,"År3")
                    add(C,namn,"År4")

                    use(A,1); use(B,2); use(B,3); use(C,4)
                    logg[namn] = {"Status":"OK"}
                    return

        # --- fallback (fixad) ---
        år_lista = [("År1",1),("År2",2),("År3",3),("År4",4)]
        used_skolor=set()

        for år, nr in år_lista:

            placed=False

            for s in skol_lista:

                if s in used_skolor:
                    continue

                if has_space(s,nr):
                    add(s,namn,år)
                    use(s,nr)
                    used_skolor.add(s)
                    placed=True
                    break

            if not placed:
                logg[namn]={"Status":"Får ej plats"}
                return

        logg[namn]={"Status":"OK*"}

    # ===== KÖR =====
    for stud in students.to_dict("records"):
        place_student(stud)

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(skol_data, key=school_sort_key):

        max_platser = int(kap.get(skola,0))

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows = [{"År1":"","År2":"","År3":"","År4":""}
                for _ in range(max_platser)]

        i=0
        for student,data in skol_data[skola].items():
            if i>=max_platser:
                break
            rows[i]=data
            i+=1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    # ===== NY FLik: SAMMANSTÄLLNING =====
    ws2 = wb.create_sheet("Sammanställning")

    ws2.append(["Student","Bostadsort","Alternativ","Max avstånd (km)"])

    for _, row in students.iterrows():

        namn = row["Namn"]
        bostad = row[bost]
        alt = row[alt_col] if alt_col else ""

        region = get_region(bostad)

        max_dist = 0

        for skola in skol_data:
            if namn in skol_data[skola]:

                sk_region = skolor.loc[
                    skolor["Skolenhet"]==skola, "Region"
                ].values[0]

                d = dist(region, sk_region)
                if d > max_dist:
                    max_dist = d

        ws2.append([namn, bostad, alt, max_dist])

    # ===== RAPPORT =====
    ws3 = wb.create_sheet("Rapport")
    ws3.append(["Student","Status"])

    for s,v in logg.items():
        ws3.append([s,v["Status"]])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp båda filer")
