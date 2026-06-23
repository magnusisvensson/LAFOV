
import streamlit as st
import pandas as pd
from collections import defaultdict, deque
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side

st.set_page_config(layout="wide")

st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = 26
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
    # STEG 1 – REGION (AUTO)
    # =========================
    if st.session_state.step == 1:

        st.header("1. Välj region")

        temp = {}

        for _, row in students.iterrows():

            name = row["Student"]
            ort_text = str(row["Ort"]).lower()

            # ✅ AUTO
            if "kalmar" in ort_text:
                region = "Kalmar"

            elif "karlskrona" in ort_text:
                region = "Karlskrona"

            elif "oskarshamn" in ort_text:
                region = "Oskarshamn"

            else:
                default = get_region(row["Ort"])
                saved = st.session_state.student_regions.get(name, default)

                region = st.selectbox(
                    f"{name} ({row['Ort']})",
                    ["Kalmar","Karlskrona","Oskarshamn"],
                    index=["Kalmar","Karlskrona","Oskarshamn"].index(saved),
                    key=f"reg_{name}"
                )

            temp[name] = region

        if st.button("✅ Bekräfta regionval"):
            st.session_state.student_regions = temp
            st.session_state.step = 2
            st.rerun()


    # =========================
    # STEG 2 – PLACERING
    # =========================
    if st.session_state.step == 2:

        st.header("2. Placering & pendling")

        usage = defaultdict(lambda:{
            "År1":0,"År2":0,"År3":0,"År4":0
        })

        results = []

        region_queues = {
            r: deque(skolor[skolor["Region"]==r]["Skolenhet"].tolist())
            for r in ["Kalmar","Karlskrona","Oskarshamn"]
        }

        for _, s in students.iterrows():

            student = s["Student"]
            region = st.session_state.student_regions.get(student)

            queue = region_queues[region]
            queue.rotate(-1)

            skolor_r = list(queue)

            reject = st.session_state.rejected.get(student,[])
            skolor_r = [sk for sk in skolor_r if sk not in reject]

            if len(skolor_r) < 3:
                skolor_r = list(queue)

            A = skolor_r[0]
            B = skolor_r[1] if len(skolor_r)>1 else A
            C = skolor_r[2] if len(skolor_r)>2 else B

            if program == "LGFRI":
                y1,y2,y3,y4 = A,A,B,""
            else:
                if region == "Kalmar":
                    y1,y2,y3,y4 = A,B,B,C
                else:
                    y1,y2,y3,y4 = A,B,A,B

            usage[y1]["År1"] += 1
            usage[y2]["År2"] += 1
            usage[y3]["År3"] += 1
            if y4: usage[y4]["År4"] += 1

            results.append({
                "Student":student,
                "Ort":s["Ort"],
                "Region":region,
                "År1":y1,"År2":y2,"År3":y3,"År4":y4
            })

        df = pd.DataFrame(results)


        # =========================
        # PENDLING
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
                            st.session_state.rejected.setdefault(r["Student"], []).append(sk)
                            st.rerun()


        # =========================
        # EXCEL
        # =========================
        wb = Workbook()
        ws = wb.active
        ws.title = "Placeringar"

        thin = Side(style="thin")
        border = Border(top=thin,left=thin,right=thin,bottom=thin)

        fill_region = PatternFill("solid","D9EAF7")

        ws.append(["Skola","År1","År2","År3","År4"])

        for region in ["Kalmar","Oskarshamn","Karlskrona"]:

            ws.append([region])
            r0 = ws.max_row
            ws.merge_cells(start_row=r0,start_column=1,end_row=r0,end_column=5)

            for c in range(1,6):
                ws.cell(r0,c).fill = fill_region

            for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

                start = ws.max_row+1

                ws.append([f"{skola} (max {kap[skola]})"])
                ws.append(["","År1","År2","År3","År4"])

                subset = df[
                    (df["År1"]==skola) |
                    (df["År2"]==skola) |
                    (df["År3"]==skola) |
                    (df["År4"]==skola)
                ].drop_duplicates("Student")

                for _, r in subset.iterrows():
                    ws.append([
                        "",
                        r["Student"] if r["År1"]==skola else "",
                        r["Student"] if r["År2"]==skola else "",
                        r["Student"] if r["År3"]==skola else "",
                        r["Student"] if r["År4"]==skola else "",
                    ])

                end = ws.max_row

                for rr in range(start,end+1):
                    for cc in range(1,6):
                        ws.cell(rr,cc).border = border


        # ----- STUDENTER -----
        ws2 = wb.create_sheet("Studenter")
        ws2.append(["Student","Ort","Region","År1","År2","År3","År4"])

        for _, r in df.iterrows():
            ws2.append(list(r))


        # ----- KONTROLL -----
        ws3 = wb.create_sheet("Kontroll")
        ws3.append(["Student","Antal skolor","Status"])

        fill_ok = PatternFill("solid","C6EFCE")
        fill_warn = PatternFill("solid","FFEB9C")
        fill_bad = PatternFill("solid","FFC7CE")

        for _, r in df.iterrows():

            skolset={r["År1"],r["År2"],r["År3"],r["År4"]}
            skolset.discard("")

            antal=len(skolset)

            if antal < 2:
                status="⚠"
                color=fill_warn
            else:
                status="OK"
                color=fill_ok

            ws3.append([r["Student"], antal, status])

            row = ws3.max_row
            ws3.cell(row,3).fill=color


        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel",f,file_name=file)
