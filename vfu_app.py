
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# ✅ VAL AV PROGRAM
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

# =========================
# REGION
# =========================
def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

# =========================
# TEXTMATCHNING
# =========================
def clean_text(text):
    return str(text).lower().replace(" ", "").replace("-", "")

def match_school(a, s):
    return clean_text(a) in clean_text(s) or clean_text(s) in clean_text(a)

if system_file and form_file:

    try:
        # =========================
        # SKOLOR
        # =========================
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        # ✅ filtrera på valt program
        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == program)
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_region)

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # =========================
        # STUDENTER
        # =========================
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()
        students["Region"] = students["Bostadsort"].apply(get_region)

        ank_kol = "Personlig anknytning"

        best_result = None
        best_log = None
        best_unplaced = 999

        # =========================
        # OPTIMERING
        # =========================
        for attempt in range(30):

            result = []
            logg = []
            ej_placerade = []
            capacity_counter = {}

            students_run = students.sample(frac=1).reset_index(drop=True)

            for region in students_run["Region"].unique():

                stud_grp = students_run[students_run["Region"] == region]
                region_skolor = list(skolor[skolor["Region"] == region]["Skolenhet"])

                for i, (_, student) in enumerate(stud_grp.iterrows()):

                    namn = f"{student['Förnamn']} {student['Efternamn']}"
                    ank_raw = str(student.get(ank_kol, "")).strip()

                    # --- filtrera anknytning
                    if ank_raw.lower() in ["", "ingen", "-", "nej"]:
                        skol_lista = region_skolor
                        exkluderade = []
                    else:
                        ank_list = [a.strip() for a in ank_raw.split(",")]

                        skol_lista = []
                        exkluderade = []

                        for s in region_skolor:
                            if any(match_school(a, s) for a in ank_list):
                                exkluderade.append(s)
                            else:
                                skol_lista.append(s)

                        if len(skol_lista) == 0:
                            skol_lista = region_skolor

                    placed = False
                    status = "OK"
                    kommentar = ""

                    # =========================
                    # LGFRI (SPECIAL)
                    # =========================
                    if program == "LGFRI":

                        for shift in range(len(skol_lista)):

                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]

                            if (
                                capacity_counter.get((A,1),0) < kap_map[A] and
                                capacity_counter.get((A,2),0) < kap_map[A] and
                                capacity_counter.get((B,3),0) < kap_map[B]
                            ):
                                placed = True
                                break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({
                                "Student": namn,
                                "Status": "Får ej plats",
                                "Kommentar": ""
                            })
                            continue

                        capacity_counter[(A,1)] = capacity_counter.get((A,1),0)+1
                        capacity_counter[(A,2)] = capacity_counter.get((A,2),0)+1
                        capacity_counter[(B,3)] = capacity_counter.get((B,3),0)+1

                        result.append({"Skola":A,"År 1":namn,"År 2":namn,"År 3":""})
                        result.append({"Skola":B,"År 1":"","År 2":"","År 3":namn})

                    # =========================
                    # LAFOV / LAGRV
                    # =========================
                    else:

                        # FULL ROTATION
                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]
                            C = skol_lista[(i+2+shift)%len(skol_lista)]

                            if (
                                capacity_counter.get((A,1),0) < kap_map[A] and
                                capacity_counter.get((B,2),0) < kap_map[B] and
                                capacity_counter.get((B,3),0) < kap_map[B] and
                                capacity_counter.get((C,4),0) < kap_map[C]
                            ):
                                placed = True
                                break

                        # fallback
                        if not placed:
                            for A in skol_lista:
                                for B in skol_lista:
                                    if A != B:
                                        if (
                                            capacity_counter.get((A,1),0) < kap_map[A] and
                                            capacity_counter.get((B,2),0) < kap_map[B] and
                                            capacity_counter.get((B,3),0) < kap_map[B] and
                                            capacity_counter.get((B,4),0) < kap_map[B]
                                        ):
                                            C = B
                                            placed = True
                                            status = "Avvikelse"
                                            kommentar = "Full rotation ej möjlig"
                                            break
                                if placed:
                                    break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({
                                "Student": namn,
                                "Status": "Får ej plats",
                                "Kommentar": ""
                            })
                            continue

                        capacity_counter[(A,1)] = capacity_counter.get((A,1),0)+1
                        capacity_counter[(B,2)] = capacity_counter.get((B,2),0)+1
                        capacity_counter[(B,3)] = capacity_counter.get((B,3),0)+1
                        capacity_counter[(C,4)] = capacity_counter.get((C,4),0)+1

                        result.append({"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""})
                        result.append({"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""})
                        result.append({"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn})

                    if exkluderade:
                        status = "Avvikelse"
                        kommentar += " Anknytning påverkade."

                    logg.append({
                        "Student": namn,
                        "Status": status,
                        "Kommentar": kommentar.strip()
                    })

            if len(ej_placerade) < best_unplaced:
                best_unplaced = len(ej_placerade)
                best_result = result
                best_log = logg

        df = pd.DataFrame(best_result)

        # =========================
        # EXCEL (enkelt – du kan koppla tillbaka din styling här)
        # =========================
        wb = Workbook()
        ws = wb.active
        ws.title = "Placering"

        ws.append(list(df.columns))

        for _, row in df.iterrows():
            ws.append(list(row))

        # RAPPORT
        ws_log = wb.create_sheet("Rapport")
        ws_log.append(["Student", "Status", "Kommentar"])

        for r in best_log:
            ws_log.append([r["Student"], r["Status"], r["Kommentar"]])

        file = "kull_resultat.xlsx"
        wb.save(file)

        with open(file, "rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

        st.success(f"✅ Klar – {best_unplaced} ej placerade")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
