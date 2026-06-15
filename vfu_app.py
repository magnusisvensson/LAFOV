
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering (optimerad)")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# =========================
# GRUPPER
# =========================
def get_student_group(bostadsort):
    bostadsort = str(bostadsort)
    if any(x in bostadsort for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Karlskrona" in bostadsort:
        return "Karlskrona"
    if "Oskarshamn" in bostadsort:
        return "Oskarshamn"
    return "Övrigt"

def get_school_group(partnerområde):
    område = str(partnerområde)
    if any(x in område for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"
    if "Karlskrona" in område:
        return "Karlskrona"
    if "Oskarshamn" in område:
        return "Oskarshamn"
    return "Övrigt"

# =========================
# MAIN
# =========================
if system_file and form_file:

    try:
        # === SKOLOR ===
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == "LAFOV")
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_school_group)

        kapacitet_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # === STUDENTER ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()
        students["Region"] = students["Bostadsort"].apply(get_student_group)

        result = []
        ej_placerade = []
        capacity_counter = {}

        # =========================
        # ROTATION OPTIMERAD
        # =========================
        for region in students["Region"].unique():

            stud_grp = students[students["Region"] == region]
            skol_grp = skolor[skolor["Region"] == region]

            skol_lista = list(skol_grp["Skolenhet"])

            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student['Förnamn']} {student['Efternamn']}"
                placed = False

                # ✅ 1. FULL ROTATION
                for shift in range(len(skol_lista)):
                    A = skol_lista[(i+shift)%len(skol_lista)]
                    B = skol_lista[(i+1+shift)%len(skol_lista)]
                    C = skol_lista[(i+2+shift)%len(skol_lista)]

                    if (
                        capacity_counter.get((A,1),0) < kapacitet_map[A] and
                        capacity_counter.get((B,2),0) < kapacitet_map[B] and
                        capacity_counter.get((B,3),0) < kapacitet_map[B] and
                        capacity_counter.get((C,4),0) < kapacitet_map[C]
                    ):
                        placed = True
                        break

                # ✅ 2. fallback B hela slutet
                if not placed:
                    for shift in range(len(skol_lista)):
                        A = skol_lista[(i+shift)%len(skol_lista)]
                        B = skol_lista[(i+1+shift)%len(skol_lista)]
                        C = B

                        if (
                            capacity_counter.get((A,1),0) < kapacitet_map[A] and
                            capacity_counter.get((B,2),0) < kapacitet_map[B] and
                            capacity_counter.get((B,3),0) < kapacitet_map[B] and
                            capacity_counter.get((C,4),0) < kapacitet_map[C]
                        ):
                            placed = True
                            break

                # ✅ 3. fallback samma skola hela tiden
                if not placed:
                    for A in skol_lista:
                        if all(
                            capacity_counter.get((A,a),0) < kapacitet_map[A]
                            for a in [1,2,3,4]
                        ):
                            B = A
                            C = A
                            placed = True
                            break

                if not placed:
                    ej_placerade.append(namn)
                    continue

                # === räkna kapacitet
                capacity_counter[(A,1)] = capacity_counter.get((A,1),0)+1
                capacity_counter[(B,2)] = capacity_counter.get((B,2),0)+1
                capacity_counter[(B,3)] = capacity_counter.get((B,3),0)+1
                capacity_counter[(C,4)] = capacity_counter.get((C,4),0)+1

                # === resultat
                result.append({"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""})
                result.append({"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""})
                result.append({"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn})

        df = pd.DataFrame(result)

        # =========================
        # KOMPAKT DATA
        # =========================
        skol_data = {}

        for _,row in df.iterrows():
            sk = row["Skola"]

            if sk not in skol_data:
                skol_data[sk] = {"År 1":[],"År 2":[],"År 3":[],"År 4":[]}

            for c in ["År 1","År 2","År 3","År 4"]:
                if row[c] != "":
                    skol_data[sk][c].append(row[c])

        # =========================
        # SORTERING (region + alfabet)
        # =========================
        def region_order(skola):
            r = region_map.get(skola,"")
            if r == "Kalmarregion":
                return 1
            elif r == "Oskarshamn":
                return 2
            elif r == "Karlskrona":
                return 3
            else:
                return 0

        sorted_skolors = sorted(skol_data.keys(), key=lambda x: (region_order(x), x))

        # =========================
        # EXCEL
        # =========================
        wb = Workbook()
        ws = wb.active

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        current_region = None

        for skola in sorted_skolors:

            region = region_map.get(skola,"")

            # ✅ regionrubrik
            if region != current_region:
                ws.append([region.upper(),"","","",""])
                current_region = region

            data = skol_data[skola]
            kap = kapacitet_map.get(skola,"-")
            antal = len(set(data["År 1"]+data["År 2"]+data["År 3"]+data["År 4"]))

            ws.append([f"{skola} ({antal}/{kap})","","","",""])

            max_len = max(len(data["År 1"]),len(data["År 2"]),len(data["År 3"]),len(data["År 4"]))

            for i in range(max_len):
                ws.append([
                    "",
                    data["År 1"][i] if i < len(data["År 1"]) else "",
                    data["År 2"][i] if i < len(data["År 2"]) else "",
                    data["År 3"][i] if i < len(data["År 3"]) else "",
                    data["År 4"][i] if i < len(data["År 4"]) else ""
                ])

        # === EJ PLACERAD
        if ej_placerade:
            ws.append(["EJ PLACERADE","","","",""])
            for s in ej_placerade:
                ws.append(["",s,"","",""])

        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
