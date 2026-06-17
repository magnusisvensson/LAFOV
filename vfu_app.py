
import streamlit as st
import pandas as pd

from openpyxl import Workbook


# =========================
# UI
# =========================
st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# =========================
# REGION
# =========================
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    # =========================
    # SKOLOR
    # =========================
    skolor = pd.read_excel(system_file, engine="openpyxl")
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ]

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # Kapacitet (säkert)
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 0

    # =========================
    # STUDENTER
    # =========================
    students = pd.read_excel(form_file, sheet_name="Data", engine="openpyxl")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    alt_bost_col = None
    for c in students.columns:
        if "alternativ" in c.lower():
            alt_bost_col = c

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    # =========================
    # PLATSER
    # =========================
    rows = []
    for _, r in skolor.iterrows():
        antal = kap[r["Skolenhet"]]
        for _ in range(antal):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "",
                "År2": "",
                "År3": "",
                "År4": ""
            })

    # =========================
    # PLACERING
    # =========================
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        usage = {sk: {"År1":0,"År2":0,"År3":0,"År4":0} for sk in skolor_r}

        def place(student, year, skola):
            for r in rows_r:
                if r["Skola"] == skola and r[year] == "":
                    r[year] = student
                    usage[skola][year] += 1
                    return True
            return False

        for i, student in enumerate(stud_r):

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)]
            C = skolor_r[(i+2) % len(skolor_r)] if len(skolor_r) > 2 else B

            # ÅR1
            place(student, "År1", A)

            # ÅR2 + ÅR3 (samma)
            for sk in [B] + skolor_r:
                if usage[sk]["År2"] < kap[sk] and usage[sk]["År3"] < kap[sk]:
                    if place(student, "År2", sk):
                        place(student, "År3", sk)
                        break

            # ÅR4 (undvik tidigare skolor)
            used = set()
            for r in rows_r:
                if student in [r["År1"], r["År2"], r["År3"]]:
                    used.add(r["Skola"])

            placed4 = False
            for sk in skolor_r:
                if sk in used:
                    continue
                if usage[sk]["År4"] < kap[sk]:
                    if place(student, "År4", sk):
                        placed4 = True
                        break

            if not placed4:
                for sk in skolor_r:
                    if usage[sk]["År4"] < kap[sk]:
                        place(student, "År4", sk)
                        break

    # =========================
    # EXCEL
    # =========================
    wb = Workbook()

    # Blad 1
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for r in rows:
        ws.append([
            r["Skola"],
            r["År1"],
            r["År2"],
            r["År3"],
            r["År4"]
        ])

    # =========================
    # BLAD 2
    # =========================
    ws2 = wb.create_sheet("Översikt studenter")

    ws2.append([
        "Student",
        "Bostadsort",
        "Alternativ bostadsort",
        "Placering 1",
        "Placering 2",
        "Placering 3"
    ])

    for _, s in students.iterrows():

        namn = s["Namn"]
        bostad = s[bost]
        alt = s[alt_bost_col] if alt_bost_col else ""

        p1 = ""
        p2 = ""
        p3 = ""

        for r in rows:
            if r["År1"] == namn:
                p1 = r["Skola"]
            if r["År2"] == namn:
                p2 = r["Skola"]
            if r["År4"] == namn:
                p3 = r["Skola"]

        ws2.append([namn, bostad, alt, p1, p2, p3])

    # =========================
    # EXPORT
    # =========================
    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file, "rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
