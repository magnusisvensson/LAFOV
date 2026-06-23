
import streamlit as st
import pandas as pd

from collections import defaultdict, deque

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side

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

# ✅ NY: reject per (student, år)
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

        # ✅ hjälpfunktion
        def filter_year(student, year, skol_list):
            rejected = st.session_state.rejected.get((student, year), set())
            return [sk for sk in skol_list if sk not in rejected]

        for _, s in students.iterrows():

            student = s["Student"]
            region = st.session_state.student_regions.get(student)

            queue = region_queues[region]
            queue.rotate(-1)

            skolor_r = list(queue)

            # =========================
            # VAL MED REJECT PER ÅR
            # =========================
            if program == "LGFRI":

                A_list = filter_year(student, "År1", skolor_r)
                A = A_list[0] if A_list else skolor_r[0]

                B_list = filter_year(student, "År3", skolor_r)
                B = B_list[0] if B_list else skolor_r[1]

                y1,y2,y3,y4 = A,A,B,""

            else:

                if region == "Kalmar":

                    A_list = filter_year(student,"År1",skolor_r)
                    A = A_list[0] if A_list else skolor_r[0]

                    B_list = filter_year(student,"År2",skolor_r)
                    B = B_list[0] if B_list else skolor_r[1]

                    C_list = filter_year(student,"År4",skolor_r)
                    C = C_list[0] if C_list else skolor_r[2]

                    y1,y2,y3,y4 = A,B,B,C

                else:

                    A_list = filter_year(student,"År1",skolor_r)
                    A = A_list[0] if A_list else skolor_r[0]

                    B_list = filter_year(student,"År2",skolor_r)
                    B = B_list[0] if B_list else skolor_r[1]

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
        # ✅ PENDLING – ITERATIV
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

                        st.write(f"{year}: {sk}")

                        val = st.radio(
                            f"Pendling OK?",
                            ["Ja","Nej"],
                            key=f"{r['Student']}_{year}"
                        )

                        if val == "Nej":

                            key = (r["Student"], year)

                            if key not in st.session_state.rejected:
                                st.session_state.rejected[key] = set()

                            st.session_state.rejected[key].add(sk)

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


        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button(
                "⬇️ Ladda ner Excel",
                data=f,
                file_name=file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
from openpyxl import Workbook
