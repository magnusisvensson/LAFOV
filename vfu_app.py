
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import PatternFill

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

    # ========= SKOLOR =========
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


    # ========= STUDENTER =========
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


    # ========= PLACERING =========
    usage = defaultdict(lambda:{
        "År1":0,"År2":0,"År3":0,"År4":0
    })

    results=[]

    def best_school(years, candidates):
        return min(
            candidates,
            key=lambda sk: max(usage[sk][y]/kap[sk] for y in years)
        )

    for _,s in students.iterrows():

        student=s["Student"]
        region=s["Region"]

        skolor_r = skolor[skolor["Region"]==region]["Skolenhet"].tolist()

        # ta bort rejected
        reject = st.session_state.rejected.get(student,[])
        skolor_r=[sk for sk in skolor_r if sk not in reject]

        if len(skolor_r)<2:
            skolor_r = skolor[skolor["Region"]==region]["Skolenhet"].tolist()

        # ===== LGFRI =====
        if program=="LGFRI":

            A = best_school(["År1","År2"], skolor_r)

            remaining=[sk for sk in skolor_r if sk != A]
            if not remaining:
                remaining=skolor_r

            B = best_school(["År3"], remaining)

            år1,år2,år3,år4 = A,A,B,""

        # ===== LAFOV / LAGRV =====
        else:

            # --- Kalmar ---
            if region=="Kalmar":

                A = best_school(["År1"], skolor_r)

                remaining1=[sk for sk in skolor_r if sk!=A]
                if not remaining1:
                    remaining1=skolor_r

                B = best_school(["År2","År3"], remaining1)

                remaining2=[sk for sk in remaining1 if sk!=B]
                if not remaining2:
                    remaining2=skolor_r

                C = best_school(["År4"], remaining2)

                år1,år2,år3,år4 = A,B,B,C

            # --- ABAB ---
            else:

                A = best_school(["År1","År3"], skolor_r)

                remaining=[sk for sk in skolor_r if sk!=A]
                if not remaining:
                    remaining=skolor_r

                B = best_school(["År2","År4"], remaining)

                år1,år2,år3,år4 = A,B,A,B

        # update usage
        usage[år1]["År1"]+=1
        if år2: usage[år2]["År2"]+=1
        if år3: usage[år3]["År3"]+=1
        if år4: usage[år4]["År4"]+=1

        results.append({
            "Student":student,
            "Ort":s["Ort"],
            "Region":region,
            "År1":år1,
            "År2":år2,
            "År3":år3,
            "År4":år4
        })

    df=pd.DataFrame(results)


    # ========= PENDLING =========
    st.subheader("🚶 Pendlingskontroll")

    name=st.text_input("Sök student")

    if name:

        name_clean=name.strip().lower()
        match=df[df["Student"].str.lower()==name_clean]

        if match.empty:
            st.warning("Student hittades inte")
        else:
            r=match.iloc[0]

            for year in ["År1","År2","År3","År4"]:

                sk=r[year]
                if sk:

                    val=st.radio(
                        f"{year}: {sk}",
                        ["Ja","Nej"],
                        key=f"{name_clean}_{year}"
                    )

                    if val=="Nej":
                        st.session_state.rejected.setdefault(r["Student"],[]).append(sk)
                        st.rerun()

                    if val=="Ja":
                        st.session_state.approved[r["Student"]] = True


    # ========= EXCEL =========
    wb=Workbook()

    ws=wb.active
    ws.title="Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    fill=PatternFill("solid","D9EAF7")

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])

        row=ws.max_row
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)

        for c in range(1,6):
            ws.cell(row,c).fill=fill

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
    ws2.append(["Student","Ort","Region","År1","År2","År3","År4"])

    for _,r in df.iterrows():
        ws2.append([r["Student"],r["Ort"],r["Region"],r["År1"],r["År2"],r["År3"],r["År4"]])

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
