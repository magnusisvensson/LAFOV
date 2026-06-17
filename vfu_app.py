
import streamlit as st
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ===== KAP =====
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = str(int(float(r["Antal platser"])))
        except:
            kap[r["Skolenhet"]] = "?"

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    # ===== PLATSER =====
    rows = []
    for _, r in skolor.iterrows():
        try:
            antal = int(float(r["Antal platser"]))
        except:
            antal = 0

        for _ in range(antal):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })

    # ===== PLACERING =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        if not rows_r:
            continue

        kapasitet = {sk: 0 for sk in skolor_r}
        for r in rows_r:
            kapasitet[r["Skola"]] += 1

        usage = {
            sk: {"År1":0,"År2":0,"År3":0,"År4":0}
            for sk in skolor_r
        }

        # ✅ HUVUDLOGIK (År2 = År3 alltid)
        def try_place(student, start_index):

            A = skolor_r[start_index]

            if len(skolor_r) <= 2:
                B = skolor_r[(start_index+1) % len(skolor_r)]
                schedule = {
                    "År1": A,
                    "År2": B,
                    "År3": B,
                    "År4": B
                }
            else:
                B = skolor_r[(start_index+1) % len(skolor_r)]
                C = skolor_r[(start_index+2) % len(skolor_r)]
                schedule = {
                    "År1": A,
                    "År2": B,
                    "År3": B,
                    "År4": C
                }

            # kontrollera plats (År2+År3 tillsammans)
            if usage[schedule["År2"]]["År2"] >= kapasitet[schedule["År2"]]:
                return False
            if usage[schedule["År3"]]["År3"] >= kapasitet[schedule["År3"]]:
                return False

            if usage[schedule["År1"]]["År1"] >= kapasitet[schedule["År1"]]:
                return False

            if usage[schedule["År4"]]["År4"] >= kapasitet[schedule["År4"]]:
                return False

            # tilldela
            for r in rows_r:
                if r["Skola"] == schedule["År1"] and r["År1"] == "":
                    r["År1"] = student
                    usage[schedule["År1"]]["År1"] += 1
                    break

            for r in rows_r:
                if r["Skola"] == schedule["År2"] and r["År2"] == "" and r["År3"] == "":
                    r["År2"] = student
                    r["År3"] = student
                    usage[schedule["År2"]]["År2"] += 1
                    usage[schedule["År2"]]["År3"] += 1
                    break

            for r in rows_r:
                if r["Skola"] == schedule["År4"] and r["År4"] == "":
                    r["År4"] = student
                    usage[schedule["År4"]]["År4"] += 1
                    break

            return True

        # ✅ FALLBACK (År2+År3 som ETTPAKET)
        def fallback_place(student):

            # År1
            for sk in skolor_r:
                if usage[sk]["År1"] < kapasitet[sk]:
                    for r in rows_r:
                        if r["Skola"] == sk and r["År1"] == "":
                            r["År1"] = student
                            usage[sk]["År1"] += 1
                            break
                    break

            # År2 + År3 SAMMA
            for sk in skolor_r:
                if (
                    usage[sk]["År2"] < kapasitet[sk] and
                    usage[sk]["År3"] < kapasitet[sk]
                ):
                    for r in rows_r:
                        if r["Skola"] == sk and r["År2"] == "" and r["År3"] == "":
                            r["År2"] = student
                            r["År3"] = student
                            usage[sk]["År2"] += 1
                            usage[sk]["År3"] += 1
                            break
                    break

            # År4
            for sk in skolor_r:
                if usage[sk]["År4"] < kapasitet[sk]:
                    for r in rows_r:
                        if r["Skola"] == sk and r["År4"] == "":
                            r["År4"] = student
                            usage[sk]["År4"] += 1
                            break
                    break


        # kör
        for i, student in enumerate(stud_r):

            placed = False

            for shift in range(len(skolor_r)):
                if try_place(student, (i+shift) % len(skolor_r)):
                    placed = True
                    break

            if not placed:
                fallback_place(student)
