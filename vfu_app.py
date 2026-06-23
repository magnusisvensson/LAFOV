
import streamlit as st
import pandas as pd
from collections import defaultdict, deque
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO

st.set_page_config(layout="wide")

st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# =========================
# SESSION
# =========================
if "step" not in st.session_state:
    st.session_state.step = 1

if "student_regions" not in st.session_state:
    st.session_state.student_regions = {}

if "rejected" not in st.session_state:
    st.session_state.rejected = {}


# =========================
# REGION
# =========================
def get_region(text):
    t = str(text).lower()

    if "oskarshamn" in t:
        return "Oskarshamn"
    if any(x in t for x in ["karlskrona","ronneby","rödeby"]):
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    # =========================
    # SKOLOR
    # =========================
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()
    skolor["Skolenhet"] = skolor["Skolenhet"].astype(str).str.strip()

    skolor = skolor[
        (
            (skolor["Kull"] == kull) |
            (skolor["Kull"].astype(str).str.upper() == "VAKANT")
        ) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 2


    # =========================
    # STUDENTER
    # =========================
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn=[c for c in students.columns if "förnamn" in c.lower()][0]
    ln=[c for c in students.columns if "efternamn" in c.lower()][0]
    bost=[c for c in students.columns if "bostads" in c.lower()][0]

    students["Student"]=(students[fn]+" "+students[ln]).str.strip()
    students["Ort"]=students[bost]


    # =========================
    # STEG 1
    # =========================
    if st.session_state.step == 1:

        st.header("1. Region")

        temp={}

        for _,row in students.iterrows():

            name=row["Student"]
            ort=str(row["Ort"]).lower()

            if "kalmar" in ort:
                region="Kalmar"
            elif "karlskrona" in ort:
                region="Karlskrona"
            elif "oskarshamn" in ort:
                region="Oskarshamn"
            else:
                region=st.selectbox(
                    f"{name} ({row['Ort']})",
                    ["Kalmar","Karlskrona","Oskarshamn"],
                    key=name
                )

            temp[name]=region

        if st.button("✅ Bekräfta"):
            st.session_state.student_regions=temp
            st.session_state.step=2
            st.rerun()


    # =========================
    # STEG 2
    # =========================
    if st.session_state.step == 2:

        usage=defaultdict(lambda:{
            "År1":0,"År2":0,"År3":0,"År4":0
        })

        results=[]

        region_queues={
            r:deque(skolor[skolor["Region"]==r]["Skolenhet"].tolist())
            for r in ["Kalmar","Karlskrona","Oskarshamn"]
        }

        # ✅ ROBUST PICK
        def pick(student,year,options,fallback):

            reject=st.session_state.rejected.get((student,year),set())

            valid=[sk for sk in options if sk not in reject]

            if valid:
                return valid[0]

            if options:
                return options[0]

            return fallback

        for _,s in students.iterrows():

            student=s["Student"]
            region=st.session_state.student_regions.get(student)

            queue=region_queues[region]
            queue.rotate(-1)

            skolor_r=list(queue)

            A = pick(student,"År1",skolor_r,skolor_r[0])

            B = pick(
                student,
                "År2",
                [x for x in skolor_r if x!=A],
                skolor_r[1] if len(skolor_r)>1 else skolor_r[0]
            )

            C = pick(
                student,
                "År4",
                [x for x in skolor_r if x not in [A,B]],
                skolor_r[2] if len(skolor_r)>2 else skolor_r[0]
            )

            if program=="LGFRI":
                y1,y2,y3,y4=A,A,B,""
            else:
                if region=="Kalmar":
                    y1,y2,y3,y4=A,B,B,C
                else:
                    y1,y2,y3,y4=A,B,A,B

            usage[y1]["År1"]+=1
            usage[y2]["År2"]+=1
            usage[y3]["År3"]+=1
            if y4: usage[y4]["År4"]+=1

            results.append({
                "Student":student,
                "Ort":s["Ort"],
                "Region":region,
                "År1":y1,
                "År2":y2,
                "År3":y3,
                "År4":y4
            })

        df=pd.DataFrame(results)


        # =========================
        # PENDLING
        # =========================
        st.header("🚶 Pendling")

        name=st.selectbox("Student",df["Student"])

        r=df[df["Student"]==name].iloc[0]

        for year in ["År1","År2","År3","År4"]:

            sk=r[year]

            if sk:

                st.write(f"{year}: {sk}")

                val=st.radio("OK?",["Ja","Nej"],key=f"{name}_{year}")

                if val=="Nej":

                    key=(name,year)

                    if key not in st.session_state.rejected:
                        st.session_state.rejected[key]=set()

                    st.session_state.rejected[key].add(sk)

                    st.rerun()


        # =========================
        # EXCEL (SNYGG VERSION)
        # =========================
        output=BytesIO()
        wb=Workbook()
        ws=wb.active
        ws.title="Placeringar"

        thin=Side(style="thin")
        border=Border(top=thin,left=thin,right=thin,bottom=thin)

        fill_region=PatternFill("solid","D9EAF7")

        row_i=1

        for region in ["Kalmar","Oskarshamn","Karlskrona"]:

            ws.cell(row_i,1,region)
            ws.merge_cells(start_row=row_i,start_column=1,end_row=row_i,end_column=5)

            for c in range(1,6):
                ws.cell(row_i,c).fill=fill_region

            row_i+=1

            for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

                start=row_i

                ws.cell(row_i,1,f"{skola} (max {kap[skola]})")
                row_i+=1

                ws.append(["","År1","År2","År3","År4"])
                row_i+=1

                subset=df[
                    (df["År1"]==skola)|
                    (df["År2"]==skola)|
                    (df["År3"]==skola)|
                    (df["År4"]==skola)
                ]

                for _,rr in subset.iterrows():
                    ws.append([
                        "",
                        rr["Student"] if rr["År1"]==skola else "",
                        rr["Student"] if rr["År2"]==skola else "",
                        rr["Student"] if rr["År3"]==skola else "",
                        rr["Student"] if rr["År4"]==skola else "",
                    ])
                    row_i+=1

                end=row_i-1

                for r_i in range(start,end+1):
                    for c_i in range(1,6):
                        ws.cell(r_i,c_i).border=border

        # ✅ AUTOBREDD
        for col in ws.columns:
            max_length=0
            col_letter=get_column_letter(col[0].column)

            for cell in col:
                try:
                    if cell.value:
                        max_length=max(max_length,len(str(cell.value)))
                except:
                    pass

            ws.column_dimensions[col_letter].width=max_length+3


        wb.save(output)
        output.seek(0)

        st.download_button(
            "⬇️ Ladda ner Excel",
            data=output,
            file_name="kull_resultat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

