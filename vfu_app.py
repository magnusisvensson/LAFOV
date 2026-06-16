
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolenheter planerade för kull:", value=26)
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

# ===== GEO approx =====
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

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    cap_used = {}
    skol_data = {}
    logg = {}
    group_id = 0

    def has_space(skola,år,n):
        return cap_used.get((skola,år),0)+n <= kap.get(skola,999)

    def use(skola,år,n):
        cap_used[(skola,år)] = cap_used.get((skola,år),0)+n

    def add(skola,student,col,group):
        skol_data.setdefault(skola,{})
        skol_data[skola].setdefault(student,{
            "År1":"","År2":"","År3":"","År4":"",
            "Group":group
        })
        skol_data[skola][student][col] = student

    # ===== PLACERING =====
    def place(group, gid):

        region = group[0]["Region"]
        names = [s["Namn"] for s in group]
        n = len(names)

        regional = skolor[skolor["Region"]==region]["Skolenhet"].tolist()
        all_schools = list(skolor["Skolenhet"])

        möjliga = regional + [s for s in all_schools if s not in regional]

        # 1. full rotation
        for i in range(len(möjliga)-2):
            A,B,C = möjliga[i],möjliga[i+1],möjliga[i+2]

            if all([
                has_space(A,1,n),
                has_space(B,2,n),
                has_space(B,3,n),
                has_space(C,4,n)
            ]):

                for s in names:
                    add(A,s,"År1",gid)
                    add(B,s,"År2",gid)
                    add(B,s,"År3",gid)
                    add(C,s,"År4",gid)

                use(A,1,n); use(B,2,n); use(B,3,n); use(C,4,n)

                for s in names:
                    logg[s]={"Status":"OK","Dist":dist(region,"Kalmar")}

                return True

        # 2. fallback per år
        for skola in möjliga:
            for år,col in [(1,"År1"),(2,"År2"),(3,"År3"),(4,"År4")]:
                if has_space(skola,år,n):

                    for s in names:
                        add(skola,s,col,gid)

                    use(skola,år,n)

                    for s in names:
                        logg[s]={"Status":"OK","Dist":dist(region,"Kalmar")}

                    return True

        return False

    stud_list = students.to_dict("records")

    i = 0
    while i < len(stud_list):

        if place(stud_list[i:i+3], group_id):
            group_id += 1
            i += 3
        elif place(stud_list[i:i+2], group_id):
            group_id += 1
            i += 2
        elif place([stud_list[i]], group_id):
            group_id += 1
            i += 1
        else:
            logg[stud_list[i]["Namn"]]={"Status":"Får ej plats","Dist":0}
            i += 1

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    ws.column_dimensions["A"].width = 40
    for c in ["B","C","D","E"]:
        ws.column_dimensions[c].width = 25

    fill_header = PatternFill(start_color="DDDDDD",fill_type="solid")
    fill_yellow = PatternFill(start_color="FFE699",fill_type="solid")
    fill_red = PatternFill(start_color="FF9999",fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(skol_data):

        max_platser = int(kap.get(skola,0))

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill_header
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows = [{"År1":"","År2":"","År3":"","År4":""} for _ in range(max_platser)]

        for student,data in skol_data[skola].items():

            idx = data["Group"] % max_platser

            for col in ["År1","År2","År3","År4"]:
                if data[col]:
                    rows[idx][col] = student

        for row_data in rows:

            ws.append(["",row_data["År1"],row_data["År2"],row_data["År3"],row_data["År4"]])
            rr = ws.max_row

            for c in range(2,6):
                namn = ws.cell(rr,c).value

                if namn in logg:
                    km = logg[namn]["Dist"]

                    if km >= 50:
                        ws.cell(rr,c).fill = fill_red
                    elif km >= 30:
                        ws.cell(rr,c).fill = fill_yellow

        ws.append([])

    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s,v in logg.items():
        ws2.append([s,v["Status"]])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp båda filer")
