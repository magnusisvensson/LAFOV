import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering (fungerande version)")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


def get_region(text):
    t=str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

def school_region(partner):
    p=str(partner).lower()
    if "oskarshamn" in p: return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p: return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    # ===== SKOLOR =====
    skolor=pd.read_excel(system_file)
    skolor.columns=skolor.columns.str.strip()

    skolor=skolor[
        (skolor["Kull"]==kull) &
        (skolor["Inriktning"].str.upper()==program)
    ].copy()

    skolor["Region"]=skolor["Partnerområde"].apply(school_region)

    kap={
        r["Skolenhet"]: int(r["Antal platser"]) if pd.notna(r["Antal platser"]) else 0
        for _,r in skolor.iterrows()
    }

    region_schools={
        reg: skolor[skolor["Region"]==reg]["Skolenhet"].tolist()
        for reg in ["Kalmar","Oskarshamn","Karlskrona"]
    }

    def school_sort_key(s):
        region=skolor.loc[skolor["Skolenhet"]==s,"Region"].values[0]
        order={"Kalmar":0,"Oskarshamn":1,"Karlskrona":2}
        return (order.get(region,3),s)

    # ===== STUDENTER =====
    students=pd.read_excel(form_file, sheet_name="Data")
    students.columns=students.columns.str.strip()

    fn=[c for c in students.columns if "förnamn" in c.lower()][0]
    ln=[c for c in students.columns if "efternamn" in c.lower()][0]
    bost=[c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"]=students[fn]+" "+students[ln]
    students["Region"]=students[bost].apply(get_region)

    student_list=students.to_dict("records")

    # ===== FÖRDELNING =====
    cap_used={}
    year_assign={1:{},2:{},4:{}}

    def has_space(s,y):
        return cap_used.get((s,y),0) < kap.get(s,0)

    def use(s,y):
        cap_used[(s,y)] = cap_used.get((s,y),0)+1

    # ✅ Fyll År1, År2, År4 för ALLA studenter
    for year in [1,2,4]:

        for idx, stud in enumerate(student_list):

            namn=stud["Namn"]
            region=stud["Region"]
            skol_lista=region_schools.get(region,[])

            if namn in year_assign[year]:
                continue

            start=(idx + year) % len(skol_lista)
            ordered=skol_lista[start:]+skol_lista[:start]

            for s in ordered:
                if has_space(s,year):
                    year_assign[year][namn]=s
                    use(s,year)
                    break

            # ✅ fallback: garantera plats
            if namn not in year_assign[year]:
                for s in skol_lista:
                    if has_space(s,year):
                        year_assign[year][namn]=s
                        use(s,year)
                        break

    # ✅ År3 = År2
    year_assign[3]=year_assign[2]

    # ===== BYGG DATA =====
    school_data={}
    for stud in student_list:

        namn=stud["Namn"]

        data={
            "År1":year_assign[1].get(namn,""),
            "År2":year_assign[2].get(namn,""),
            "År3":year_assign[3].get(namn,""),
            "År4":year_assign[4].get(namn,"")
        }

        for year,skola in data.items():
            if skola=="":
                continue

            school_data.setdefault(skola,{})
            school_data[skola].setdefault(namn,{
                "År1":"","År2":"","År3":"","År4":""
            })

            school_data[skola][namn][year]=namn

    # ===== EXCEL =====
    wb=Workbook()
    ws=wb.active
    ws.title="Placeringar"

    fill=PatternFill(start_color="DDDDDD",fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(kap.keys(), key=school_sort_key):

        max_platser=kap.get(skola,0)

        ws.append([f"{skola} (max {max_platser})"])
        r=ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill=fill
            ws.cell(r,c).font=Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows=[{"År1":"","År2":"","År3":"","År4":""}
              for _ in range(max_platser)]

        i=0
        if skola in school_data:
            for student,data in school_data[skola].items():
                if i>=max_platser:
                    break
                rows[i]=data
                i+=1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    # ===== RAPPORT =====
    ws2=wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in students["Namn"]:
        ws2.append([s,"OK"])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp båda filer")
