
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

# ========= REGION =========
def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

# ========= TEXTMATCHNING =========
def clean_text(text):
    return str(text).lower().replace(" ", "").replace("-", "")

def match_school(a, s):
    return clean_text(a) in clean_text(s) or clean_text(s) in clean_text(a)

# ========= AUTO-UPPTÄCK =========
def find_column(columns, keywords):
    for col in columns:
        if any(k.lower() in col.lower() for k in keywords):
            return col
    return None

if system_file and form_file:

    try:
        # ========= SKOLOR =========
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == program)
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_region)

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # ========= STUDENTER =========
        students = pd.read_excel(form_file, sheet_name="Data")
        students.columns = students.columns.str.strip()

        # ✅ AUTO-KOLUMNER
        fn_col = find_column(students.columns, ["förnamn"])
        ln_col = find_column(students.columns, ["efternamn"])
        bostad_col = find_column(students.columns, ["bostadsort"])
        ank_col = find_column(students.columns, ["anknytning"])

        students["Namn"] = students[fn_col] + " " + students[ln_col]
        students["Region"] = students[bostad_col].apply(get_region)

        best_result = None
        best_log = None
        best_unplaced = 999

        # ========= OPTIMERING =========
        for _ in range(30):

            result = []
            logg = []
            ej_placerade = []
            cap = {}

            students_run = students.sample(frac=1)

            for region in students_run["Region"].unique():

                stud_grp = students_run[students_run["Region"] == region]
                region_skolor = list(skolor[skolor["Region"] == region]["Skolenhet"])

                for i, (_, student) in enumerate(stud_grp.iterrows()):

                    namn = student["Namn"]
                    ank_raw = str(student.get(ank_col, "")).strip()

                    # filtrera anknytning
                    if ank_raw.lower() in ["", "ingen", "-", "nej"]:
                        skol_lista = region_skolor
                        exkluderade = []
                    else:
                        skol_lista = []
                        exkluderade = []
                        ank_list = ank_raw.split(",")

                        for s in region_skolor:
                            if any(match_school(a, s) for a in ank_list):
                                exkluderade.append(s)
                            else:
                                skol_lista.append(s)

                        if not skol_lista:
                            skol_lista = region_skolor

                    placed = False
                    status = "OK"
                    kommentar = ""

                    # ===== LGFRI =====
                    if program == "LGFRI":
                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]

                            if (
                                cap.get((A,1),0) < kap_map[A] and
                                cap.get((A,2),0) < kap_map[A] and
                                cap.get((B,3),0) < kap_map[B]
                            ):
                                placed = True
                                break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({"Student": namn, "Status": "Får ej plats", "Kommentar": ""})
                            continue

                        cap[(A,1)] = cap.get((A,1),0)+1
                        cap[(A,2)] = cap.get((A,2),0)+1
                        cap[(B,3)] = cap.get((B,3),0)+1

                        result += [
                            {"Skola":A,"År 1":namn,"År 2":namn,"År 3":""},
                            {"Skola":B,"År 1":"","År 2":"","År 3":namn}
                        ]

                    # ===== ÖVRIGA =====
                    else:
                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]
                            C = skol_lista[(i+2+shift)%len(skol_lista)]

                            if (
                                cap.get((A,1),0) < kap_map[A] and
                                cap.get((B,2),0) < kap_map[B] and
                                cap.get((B,3),0) < kap_map[B] and
                                cap.get((C,4),0) < kap_map[C]
                            ):
                                placed = True
                                break

                        if not placed:
                            for A in skol_lista:
                                for B in skol_lista:
                                    if A != B:
                                        if (
                                            cap.get((A,1),0) < kap_map[A] and
                                            cap.get((B,2),0) < kap_map[B] and
                                            cap.get((B,3),0) < kap_map[B] and
                                            cap.get((B,4),0) < kap_map[B]
                                        ):
                                            C = B
                                            placed = True
                                            status = "Avvikelse"
                                            kommentar = "Fallback använd"
                                            break
                                if placed:
                                    break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({"Student": namn, "Status": "Får ej plats", "Kommentar": ""})
                            continue

                        cap[(A,1)] = cap.get((A,1),0)+1
                        cap[(B,2)] = cap.get((B,2),0)+1
                        cap[(B,3)] = cap.get((B,3),0)+1
                        cap[(C,4)] = cap.get((C,4),0)+1

                        result += [
                            {"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""},
                            {"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""},
                            {"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn}
                        ]

                    if exkluderade:
                        status = "Avvikelse"
                        kommentar += " Anknytning påverkade"

                    logg.append({"Student": namn, "Status": status, "Kommentar": kommentar})

            if len(ej_placerade) < best_unplaced:
                best_unplaced = len(ej_placerade)
                best_result = result
                best_log = logg

        df = pd.DataFrame(best_result)

        # ========= EXCEL (PIXEL LAYOUT) =========
        wb = Workbook()
        ws = wb.active

        fill_green = PatternFill(start_color="CCFFCC", fill_type="solid")
        fill_dark = PatternFill(start_color="99CC66", fill_type="solid")
        fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")

        thin = Side(style="thin")
        thick = Side(style="medium")

        ws.append(df.columns.tolist())

        for _, row in df.iterrows():
            ws.append(row.tolist())

        # ========= RAPPORT =========
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
