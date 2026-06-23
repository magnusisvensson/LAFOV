
import streamlit as st
import pandas as pd
from collections import defaultdict, deque
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
from io import BytesIO

st.set_page_config(layout="wide")

st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

# ✅ tillbaka
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

    students["Student"] = (students[fn] + " " + students[ln]).str.strip()
    students["Ort"] = students[bost]


    # =========================
    # STEG 1 – REGION
    # =========================
    if st.session_state.step == 1:

        st.header("1. Region")

        temp = {}

        for _, row in students.iterrows():

            name = row["Student"]
            ort = str(row["Ort"]).lower()

            if "kalmar" in ort:
                region = "Kalmar"
            elif "karlskrona" in ort:
                region = "Karlskrona"
            elif "oskarshamn" in ort:
                region = "Oskarshamn"
            else:
                region = st.selectbox(
                    f"{name} ({row['Ort']})",
                    ["Kalmar","Karlskrona","Oskarshamn"],
                    key=name
                )

            temp[name] = region

        if st.button("✅ Bekräfta"):
            st.session_state.student_regions = temp
            st.session_state.step = 2
            st.rerun()


    # =========================
    # STEG 2 – PLACERING
    # =========================
    if st.session_state.step == 2:

        usage = defaultdict(lambda:{
            "År1":0,"År2":0,"År3":0,"År4":0
        })

        results = []

        region_queues = {
            r: deque(skolor[skolor["Region"]==r]["Skolenhet"].tolist())
            for r in ["Kalmar","Karlskrona","Oskarshamn"]
        }

        def pick(student, year, options):
            reject = st.session_state.rejected.get((student,year), set())

            for sk in options:
                if sk not in reject:
                    return sk

            # fallback
            return options[0]

        for _, s in students.iterrows():

            student = s["Student"]
            region = st.session_state.student_regions.get(student)

            queue = region_queues[region]
            queue.rotate(-1)
            skolor_r = list(queue)

            A = pick(student,"År1",skolor_r)
            B = pick(student,"År2",[x for x in skolor_r if x!=A])
            C = pick(student,"År4",[x for x in skolor_r if x not in [A,B]])

            if program == "LGFRI":
                y1,y2,y3,y4 = A,A,B,""
            else:
                if region == "Kalmar":
                    y1,y2,y3,y4 = A,B,B,C
                else:
                    y1,y2,y3,y4 = A,B,A,B

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

        df = pd.DataFrame(results)

        # =========================
        # ✅ PENDLING FUNKAR NU
        # =========================
        st.header("🚶 Pendling")

        name = st.selectbox("Välj student", df["Student"])

        r = df[df["Student"]==name].iloc[0]

        for year in ["År1","År2","År3","År4"]:

            sk = r[year]

            if sk:

                st.write(f"{year}: {sk}")

                val = st.radio(
                    f"OK?",
                    ["Ja","Nej"],
                    key=f"{name}_{year}"
                )

                if val == "Nej":

                    key = (name, year)

                    if key not in st.session_state.rejected:
                        st.session_state.rejected[key] = set()

                    st.session_state.rejected[key].add(sk)

                    st.rerun()


        # =========================
        # ✅ EXCEL FUNKAR NU
        # =========================
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Placeringar"

        ws.append(["Student","År1","År2","År3","År4"])

        for _, r in df.iterrows():
            ws.append([r["Student"], r["År1"], r["År2"], r["År3"], r["År4"]])

        wb.save(output)
        output.seek(0)

        st.download_button(
            "⬇️ Ladda ner Excel",
            data=output,
            file_name="kull_resultat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
