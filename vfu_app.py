
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
# SESSION STATE
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

    # ✅ RENGÖR SKOLNAMN (fix för KeyError)
    skolor["Skolenhet"] = skolor["Skolenhet"].astype(str).str.strip()

    skolor = skolor[
        (
            (skolor["Kull"]==kull) |
            (skolor["Kull"].astype(str).str.upper()=="VAKANT")
        ) &
        (skolor["Inriktning"].str.upper()==program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ✅ ta bort trasiga regioner
    skolor = skolor[skolor["Region"].notna()]

    # ✅ kap
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


    # =========================
    # 🔵 STEG 1 – REGIONVAL
    # =========================
    if st.session_state.step == 1:

        st.header("1. Välj region")

        temp_regions = {}

        for _,row in students.iterrows():

            name = row["Student"]

            default = get_region(row["Ort"])
            saved = st.session_state.student_regions.get(name, default)

            region = st.selectbox(
                f"{name} ({row['Ort']})",
                ["Kalmar","Karlskrona","Oskarshamn"],
                index=["Kalmar","Karlskrona","Oskarshamn"].index(saved) if saved else 0,
                key=f"region_{name}"
            )

            temp_regions[name] = region

        if st.button("✅ Bekräfta regionval"):
            st.session_state.student_regions = temp_regions
            st.session_state.step = 2
            st.rerun()


    # =========================
    # 🟢 STEG 2 – PLACERING
    # =========================
    if st.session_state.step == 2:

        st.header("2. Placering & pendling")

        usage = defaultdict(lambda:{
            "År1":0,"År2":0,"År3":0,"År4":0
        })

        results=[]

        # ✅ ROBUST VERSION (fix)
        def best_school(years, candidates):

            # filtrera bort trasig data
            candidates = [
                sk for sk in candidates
                if isinstance(sk,str)
                and sk.strip() != ""
                and sk in kap
            ]

            if not candidates:
                return None

            return min(
                candidates,
                key=lambda sk: max(usage[sk].get(y,0)/kap.get(sk,1) for y in years)
            )

        # =========================
        # PLACERING
        # =========================
        for i,s in students.iterrows():

            student = s["Student"]
            region = st.session_state.student_regions.get(student)

            skolor_r = skolor[skolor["Region"]==region]["Skolenhet"].tolist()

            # ✅ säkerställ lista
            skolor_r = [sk for sk in skolor_r if isinstance(sk,str) and sk.strip()!=""]

            # ✅ ROTATION (så alla skolor används)
            if skolor_r:
                shift = i % len(skolor_r)
                skolor_r = skolor_r[shift:] + skolor_r[:shift]

            # ✅ reject
            reject = st.session_state.rejected.get(student,[])
            skolor_r = [sk for sk in skolor_r if sk not in reject]

            if len(skolor_r) < 2:
                skolor_r = skolor[skolor["Region"]==region]["Skolenhet"].tolist()

            # =========================
            # LGFRI
            # =========================
            if program == "LGFRI":

                A = best_school(["År1","År2"], skolor_r)

                rest=[sk for sk in skolor_r if sk!=A]
                B = best_school(["År3"], rest if rest else skolor_r)

                y1,y2,y3,y4 = A,A,B,""

            # =========================
            # LAFOV/LAGRV
            # =========================
            else:

                if region == "Kalmar":

                    A = best_school(["År1"], skolor_r)

                    rest1=[sk for sk in skolor_r if sk!=A]
                    B = best_school(["År2","År3"], rest1 if rest1 else skolor_r)

                    rest2=[sk for sk in rest1 if sk!=B]
                    C = best_school(["År4"], rest2 if rest2 else skolor_r)

                    y1,y2,y3,y4 = A,B,B,C

                else:

                    A = best_school(["År1","År3"], skolor_r)

                    rest=[sk for sk in skolor_r if sk!=A]
                    B = best_school(["År2","År4"], rest if rest else skolor_r)

                    y1,y2,y3,y4 = A,B,A,B

            # =========================
            # UPDATE usage
            # =========================
            if y1: usage[y1]["År1"]+=1
            if y2: usage[y2]["År2"]+=1
            if y3: usage[y3]["År3"]+=1
            if y4: usage[y4]["År4"]+=1

            results.append({
                "Student":student,
                "Ort":s["Ort"],
                "Region":region,
                "År1":y1,"År2":y2,"År3":y3,"År4":y4
            })

        df=pd.DataFrame(results)


        # =========================
        # 🚶 PENDLING
        # =========================
        st.subheader("🚶 Pendlingskontroll")

        name = st.text_input("Sök student")

        if name:

            match = df[df["Student"].str.lower()==name.strip().lower()]

            if not match.empty:

                r = match.iloc[0]

                for year in ["År1","År2","År3","År4"]:

                    sk = r[year]

                    if sk:

                        val = st.radio(
                            f"{year}: {sk}",
                            ["Ja","Nej"],
                            key=f"{r['Student']}_{year}"
                        )

                        if val == "Nej":
                            st.session_state.rejected.setdefault(r["Student"],[]).append(sk)
                            st.rerun()


        # =========================
        # 📊 EXCEL
        # =========================
        wb=Workbook()

        ws=wb.active
        ws.title="Placeringar"

        ws.append(["Skola","År1","År2","År3","År4"])

        fill = PatternFill("solid","D9EAF7")

        for region in ["Kalmar","Oskarshamn","Karlskrona"]:

            ws.append([])
            ws.append([region])
            row = ws.max_row
            ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=5)

            for c in range(1,6):
                ws.cell(row,c).fill=fill

            skolor_r = skolor[skolor["Region"]==region]["Skolenhet"]

            for skola in skolor_r:

                ws.append([f"{skola} (max {kap[skola]})"])

                subset=df[
                    (df["År1"]==skola)|
                    (df["År2"]==skola)|
                    (df["År3"]==skola)|
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

        # ---- Studenter ----
        ws2=wb.create_sheet("Studenter")
        ws2.append(["Student","Ort","Region","År1","År2","År3","År4"])

        for _,r in df.iterrows():
            ws2.append([r["Student"],r["Ort"],r["Region"],r["År1"],r["År2"],r["År3"],r["År4"]])

        # ---- Kontroll ----
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
``
