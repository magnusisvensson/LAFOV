
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
# SESSION STATE (NYTT)
# =========================
if "rejected" not in st.session_state:
    st.session_state.rejected = {}

if "approved" not in st.session_state:
    st.session_state.approved = {}


def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if any(x in t for x in ["karlskrona","ronneby","rödeby"]):
        return "Karlskrona"
    if "kalmar" in t:
        return "Kalmar"
    return None


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (
            (skolor["Kull"] == kull) |
            (skolor["Kull"].astype(str).str.upper()=="VAKANT")
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


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostads" in c.lower()][0]

    students["Student"] = (students[fn] + " " + students[ln]).str.strip()
    students["Ort"] = students[bost]

    # ===== REGIONVAL =====
    regions = []
    for _, row in students.iterrows():
        region = get_region(row["Ort"])

        if region is None:
            region = st.selectbox(
                f"Välj region för {row['Student']} ({row['Ort']})",
                ["Kalmar","Karlskrona","Oskarshamn"],
                key=row["Student"]
            )

        regions.append(region)

    students["Region"] = regions


    # =========================
    # ✅ PLACERING (BALANS + REJEKT)
    # =========================
    usage = defaultdict(int)
    results = []

    for _, s in students.iterrows():

        student = s["Student"]
        region = s["Region"]

        if st.session_state.approved.get(student):
            continue

        skolor_r = skolor[skolor["Region"] == region]["Skolenhet"].tolist()

        # luta mot alla om få
        if len(skolor_r) < 3:
            skolor_r = skolor["Skolenhet"].tolist()

        # 🟥 ta bort rejected
        reject_list = st.session_state.rejected.get(student, [])
        skolor_r = [sk for sk in skolor_r if sk not in reject_list]

        if len(skolor_r) == 0:
            skolor_r = skolor["Skolenhet"].tolist()

        # ✅ belastning
        scores = [(sk, usage[sk]/kap.get(sk,2)) for sk in skolor_r]
        scores.sort(key=lambda x: x[1])

        A = scores[0][0]
        B = scores[1][0] if len(scores)>1 else A
        C = scores[2][0] if len(scores)>2 else B

        if program == "LGFRI":
            år1, år2, år3, år4 = A, A, B, ""
        else:
            if region == "Kalmar":
                år1, år2, år3, år4 = A, B, B, C
            else:
                år1, år2, år3, år4 = A, B, A, B

        usage[A]+=1
        usage[B]+=1
        usage[C]+=1

        results.append({
            "Student":student,
            "Ort":s["Ort"],
            "År1":år1,
            "År2":år2,
            "År3":år3,
            "År4":år4
        })

    df = pd.DataFrame(results)

    # =========================
    # 🚶 PENDLING (INTERAKTIV)
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange student")

    if student_input:

        match = df[df["Student"].str.lower()==student_input.strip().lower()]

        if len(match)==0:
            st.warning("Student hittades inte")

        else:
            r = match.iloc[0]

            st.write(f"**Ort:** {r['Ort']}")

            for year in ["År1","År2","År3","År4"]:

                if r[year] != "":
                    skola = r[year]

                    val = st.radio(
                        f"{year}: {skola}",
                        ["Ja","Nej"],
                        key=f"{r['Student']}_{year}"
                    )

                    if val == "Nej":

                        if r["Student"] not in st.session_state.rejected:
                            st.session_state.rejected[r["Student"]] = []

                        st.session_state.rejected[r["Student"]].append(skola)

                        st.rerun()

                    if val == "Ja":
                        st.session_state.approved[r["Student"]] = True


    # =========================
    # 📊 EXCEL
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])

        skolor_r = skolor[skolor["Region"]==region]["Skolenhet"]

        for skola in skolor_r:

            ws.append([f"{skola} (max {kap[skola]})"])

            subset = df[
                (df["År1"]==skola) |
                (df["År2"]==skola) |
                (df["År3"]==skola) |
                (df["År4"]==skola)
            ].drop_duplicates(subset="Student")

            if subset.empty:
                ws.append(["","","","",""])
            else:
                for _, s in subset.iterrows():
                    ws.append([
                        "",
                        s["Student"] if s["År1"]==skola else "",
                        s["Student"] if s["År2"]==skola else "",
                        s["Student"] if s["År3"]==skola else "",
                        s["Student"] if s["År4"]==skola else "",
                    ])

            ws.append([])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
