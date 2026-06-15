
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering (slutversion)")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

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

    # ✅ ÖVRIGT → Kalmar
    return "Kalmarregion"

if system_file and form_file:

    try:
        # === SKOLOR ===
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == "LAFOV")
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_region)

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # === STUDENTER ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()
        students["Region"] = students["Bostadsort"].apply(get_region)

        result = []
        ej_placerade = []
        capacity_counter = {}

        # =========================
        # ROTATION
        # =========================
        for region in students["Region"].unique():

            stud_grp = students[students["Region"] == region]
            skol_lista = list(skolor[skolor["Region"] == region]["Skolenhet"])

            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student['Förnamn']} {student['Efternamn']}"
                placed = False

                # ✅ 1. FULL ROTATION
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

                # ✅ 2. fallback (minst ett byte)
                if not placed:
                    for A in skol_lista:
                        for B in skol_lista:

                            if A == B:
                                continue

                            if (
                                capacity_counter.get((A,1),0) < kap_map[A] and
                                capacity_counter.get((B,2),0) < kap_map[B] and
                                capacity_counter.get((B,3),0) < kap_map[B] and
                                capacity_counter.get((B,4),0) < kap_map[B]
                            ):
                                C = B
                                placed = True
                                break
                        if placed:
                            break

                if not placed:
                    ej_placerade.append(namn)
                    continue

                capacity_counter[(A,1)] = capacity_counter.get((A,1),0)+1
                capacity_counter[(B,2)] = capacity_counter.get((B,2),0)+1
                capacity_counter[(B,3)] = capacity_counter.get((B,3),0)+1
                capacity_counter[(C,4)] = capacity_counter.get((C,4),0)+1

                result.append({"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""})
                result.append({"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""})
                result.append({"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn})

        df = pd.DataFrame(result)

        # =========================
        # KOMPAKT DATA
        # =========================
        skol_data = {}
        for _, row in df.iterrows():
            sk = row["Skola"]
            if sk not in skol_data:
                skol_data[sk] = {"År 1":[],"År 2":[],"År 3":[],"År 4":[]}
            for c in ["År 1","År 2","År 3","År 4"]:
                if row[c] != "":
                    skol_data[sk][c].append(row[c])

        # =========================
        # SORTERING
        # =========================
        def region_order(s):
            r = region_map.get(s,"")
            if r == "Kalmarregion": return 1
            if r == "Oskarshamn": return 2
            if r == "Karlskrona": return 3
            return 0

        sorted_skolor = sorted(skol_data.keys(), key=lambda x: (region_order(x), x))

        # =========================
        # EXCEL
        # =========================
        wb = Workbook()
        ws = wb.active

        ws.append(["Skola","År 1","År 2","År 3","År 4"])

        current_region = None

        for skola in sorted_skolor:

            region = region_map.get(skola,"")

            if region != current_region:
                ws.append([region.upper(),"","","",""])
                current_region = region

            data = skol_data[skola]
            kap = kap_map.get(skola,"-")
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

        if ej_placerade:
            ws.append(["EJ PLACERADE","","","",""])
            for s in ej_placerade:
                ws.append(["",s,"","",""])

        file = "kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
