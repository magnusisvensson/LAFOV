
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(layout="wide")

st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# =========================
# SESSION
# =========================
if "rejected" not in st.session_state:
    st.session_state.rejected = {}

if "approved" not in st.session_state:
    st.session_state.approved = {}


# =========================
# REGION
# =========================
def get_region(text):
    t=str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if any(x in t for x in ["karlskrona","ronneby","rödeby"]): return "Karlskrona"
    if "kalmar" in t: return "Kalmar"
    return None


if system_file and form_file:

    # =========================
    # SKOLOR
    # =========================
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (
            (skolor["Kull"]==kull) |
            (skolor["Kull"].astype(str).str.upper()=="VAKANT")
        ) &
        (skolor["Inriktning"].str.upper()==program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    kap = {}
    for _,r in skolor.iterrows():
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

    # regionval
    reg=[]
    for _,row in students.iterrows():
        r=get_region(row["Ort"])
        if r is None:
            r=st.selectbox(
                f"Region för {row['Student']} ({row['Ort']})",
                ["Kalmar","Karlskrona","Oskarshamn"],
                key=row["Student"]
            )
        reg.append(r)
    students["Region"]=reg


    # =========================
    # ✅ KAP PER ÅR
    # =========================
    usage = defaultdict(lambda:{
        "År1":0,"År2":0,"År3":0,"År4":0
    })

    results=[]

    for _,s in students.iterrows():

        student=s["Student"]
        region=s["Region"]

        skolor_r = skolor[skolor["Region"]==region]["Skolenhet"].tolist()

        # ❗ bara samma region
        if not skolor_r:
            continue

        # ta bort rejected
        rejected = st.session_state.rejected.get(student,[])
        skolor_r=[sk for sk in skolor_r if sk not in rejected]

        if len(skolor_r)<3:
            skolor_r = skolor[skolor["Region"]==region]["Skolenhet"].tolist()

        # score funktion
        def score(sk,years):
            return max(usage[sk][y]/kap[sk] for y in years)

        best=None

        for a in skolor_r:
            for b in skolor_r:
                for c in skolor_r:

                    if program=="LGFRI":
                        sc = score(a,["År1","År2"]) + score(b,["År3"])
                        combo=(a,a,b)
                    else:
                        if region=="Kalmar":
                            sc = score(a,["År1"]) + score(b,["År2","År3"]) + score(c,["År4"])
                            combo=(a,b,c)
                        else:
                            sc = score(a,["År1","År3"]) + score(b,["År2","År4"])
                            combo=(a,b,a,b)

                    if best is None or sc<best[0]:
                        best=(sc,combo)

        if program=="LGFRI":
            A,A2,B = best[1]
            år1,år2,år3 = A,A2,B
            år4=""
        else:
            if region=="Kalmar":
                A,B,C = best[1]
                år1,år2,år3,år4 = A,B,B,C
            else:
                A,B,_,_ = best[1]
                år1,år2,år3,år4 = A,B,A,B

        # update usage
        usage[år1]["År1"]+=1
        if år2: usage[år2]["År2"]+=1
        if år3: usage[år3]["År3"]+=1
        if år4: usage[år4]["År4"]+=1

        results.append({
            "Student":student,
            "Ort":s["Ort"],
            "År1":år1,"År2":år2,"År3":år3,"År4":år4
        })

    df=pd.DataFrame(results)

    # =========================
    # 🚶 PENDLING (FIX)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    name=st.text_input("Sök student")

    if name:

        match=df[df["Student"].str.lower()==name.strip().lower()]

        if len(match)==0:
            st.warning("Student hittades inte")
        else:
            r=match.iloc[0]

            for year in ["År1","År2","År3","År4"]:
                sk=r[year]
                if sk:
                    val=st.radio(f"{year}: {sk}",["Ja","Nej"],key=f"{name}_{year}")

                    if val=="Nej":
                        st.session_state.rejected.setdefault(r["Student"],[]).append(sk)
                        st.rerun()

                    if val=="Ja":
                        st.session_state.approved[r["Student"]]=True

    # =========================
    # EXCEL
    # =========================
    wb=Workbook()

    # ---- blad 1 ----
    ws=wb.active
    ws.title="Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    fill_reg=PatternFill("solid","D9EAF7")

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])
        row=ws.max_row
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)

        for c in range(1,6):
            ws.cell(row,c).fill=fill_reg

        skolor_r=skolor[skolor["Region"]==region]["Skolenhet"]

        for skola in skolor_r:

            ws.append([f"{skola} (max {kap[skola]})"])

            subset=df[
                (df["År1"]==skola) |
                (df["År2"]==skola) |
                (df["År3"]==skola) |
                (df["År4"]==skola)
            ].drop_duplicates("Student")

            for _,s in subset.iterrows():
                ws.append([
                    "",
                    s["Student"] if s["År1"]==skola else "",
                    s["Student"] if s["År2"]==skola else "",
                    s["Student"] if s["År3"]==skola else "",
                    s["Student"] if s["År4"]==skola else "",
                ])

            ws.append([])

    # ---- blad 2 ----
    ws2=wb.create_sheet("Studenter")
    ws2.append(["Student","Ort","År1","År2","År3","År4"])

    for _,r in df.iterrows():
        ws2.append([r["Student"],r["Ort"],r["År1"],r["År2"],r["År3"],r["År4"]])

    # ---- blad 3 ----
    ws3=wb.create_sheet("Kontroll")
    ws3.append(["Student","Antal skolor"])

    for _,r in df.iterrows():
        skolset={r["År1"],r["År2"],r["År3"],r["År4"]}
        skolset.discard("")
        ws3.append([r["Student"],len(skolset)])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)
