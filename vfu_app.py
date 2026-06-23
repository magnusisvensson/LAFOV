
import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook

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

if "regions_locked" not in st.session_state:
    st.session_state.regions_locked = False

if "rejected" not in st.session_state:
    st.session_state.rejected = {}

if "school_index" not in st.session_state:
    st.session_state.school_index = 0


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

    # =========================
    # STEG 1: REGION
    # =========================
    if st.session_state.step == 1:

        st.header("1. Välj region")

        regions=[]

        for _,row in students.iterrows():

            r=get_region(row["Ort"])

            if r is None:
                r=st.selectbox(
                    f"{row['Student']} ({row['Ort']})",
                    ["Kalmar","Karlskrona","Oskarshamn"],
                    key=row["Student"]
                )

            regions.append(r)

        students["Region"]=regions

        if st.button("✅ Bekräfta regionval"):
            st.session_state.regions_locked=True
            st.session_state.step=2
            st.rerun()

    # =========================
    # STEG 2: PLACERING + PENDLING
    # =========================
    if st.session_state.step == 2:

        st.header("2. Placering & pendling")

        usage = defaultdict(lambda:{
            "År1":0,"År2":0,"År3":0,"År4":0
        })

        results=[]

        for i,s in students.iterrows():

            skolor_r = skolor[skolor["Region"]==students.loc[i,"Region"]]["Skolenhet"].tolist()

            # 🔁 ROTATION (NYCKELN)
            skolor_r = skolor_r[i % len(skolor_r):] + skolor_r[:i % len(skolor_r)]

            def best(sk_list, years):
                return min(
                    sk_list,
                    key=lambda sk: max(usage[sk][y]/kap[sk] for y in years)
                )

            if program=="LGFRI":

                A = best(skolor_r, ["År1","År2"])
                rest=[sk for sk in skolor_r if sk!=A]

                B = best(rest if rest else skolor_r, ["År3"])

                y1,y2,y3,y4=A,A,B,""

            else:
                if students.loc[i,"Region"]=="Kalmar":

                    A = best(skolor_r, ["År1"])

                    rest1=[sk for sk in skolor_r if sk!=A]
                    B = best(rest1 if rest1 else skolor_r, ["År2","År3"])

                    rest2=[sk for sk in rest1 if sk!=B]
                    C = best(rest2 if rest2 else skolor_r, ["År4"])

                    y1,y2,y3,y4=A,B,B,C

                else:

                    A = best(skolor_r, ["År1","År3"])
                    rest=[sk for sk in skolor_r if sk!=A]
                    B = best(rest if rest else skolor_r, ["År2","År4"])

                    y1,y2,y3,y4=A,B,A,B

            usage[y1]["År1"]+=1
            if y2: usage[y2]["År2"]+=1
            if y3: usage[y3]["År3"]+=1
            if y4: usage[y4]["År4"]+=1

            results.append({
                "Student":s["Student"],
                "Ort":s["Ort"],
                "Region":students.loc[i,"Region"],
                "År1":y1,"År2":y2,"År3":y3,"År4":y4
            })

        df=pd.DataFrame(results)

        # =========================
        # PENDLING
        # =========================
        st.subheader("🚶 Pendlingskontroll")

        name=st.text_input("Sök student")

        if name:

            r=df[df["Student"].str.lower()==name.lower()]
            if not r.empty:

                r=r.iloc[0]

                for year in ["År1","År2","År3","År4"]:
                    st.write(f"{year}: {r[year]}")

        # =========================
        # EXCEL
        # =========================
        wb=Workbook()
        ws=wb.active
        ws.title="Placeringar"

        ws.append(["Skola","År1","År2","År3","År4"])

        for region in ["Kalmar","Oskarshamn","Karlskrona"]:

            ws.append([])
            ws.append([region])

            skolor_r=skolor[skolor["Region"]==region]["Skolenhet"]

            for skola in skolor_r:

                ws.append([skola])

                subset=df[
                    (df["År1"]==skola)|
                    (df["År2"]==skola)|
                    (df["År3"]==skola)|
                    (df["År4"]==skola)
                ]

                for _,r in subset.iterrows():
                    ws.append([
                        "",
                        r["Student"] if r["År1"]==skola else "",
                        r["Student"] if r["År2"]==skola else "",
                        r["Student"] if r["År3"]==skola else "",
                        r["Student"] if r["År4"]==skola else "",
                    ])

                ws.append([])

        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner",f,file_name=file)
