
import streamlit as st
import pandas as pd
import hashlib
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

# ================= GEO =================
geo = {
    "Kalmar": (56.66, 16.36),
    "Oskarshamn": (57.26, 16.45),
    "Karlskrona": (56.16, 15.59),
    "Ronneby": (56.21, 15.28)
}

def normalize_location(text):
    t = str(text).lower()
    if "påskallavik" in t:
        return "Oskarshamn"
    if "kallinge" in t:
        return "Ronneby"
    if any(x in t for x in ["kalmar","lindsdal","nybro","emmaboda","mönsterås","färjestaden"]):
        return "Kalmar"
    return text

def distance(a, b):
    if a not in geo or b not in geo:
        return None
    return ((geo[a][0]-geo[b][0])**2 + (geo[a][1]-geo[b][1])**2)**0.5

def distance_km_raw(a, b):
    d = distance(a, b)
    if d is None:
        return None
    return round(d * 111, 1)

# ================= MAIN =================
if system_file and form_file:

    skolor_all = pd.read_excel(system_file)
    skolor_all.columns = skolor_all.columns.str.strip()

    skolor = skolor_all[
        (skolor_all["Kull"] == kull) &
        (skolor_all["Inriktning"].str.upper() == program)
    ].copy()

    kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

    # ===== AUTO GEO FÖR ALLA SKOLOR =====
    base_geo = {
        "Kalmar": (56.66, 16.36),
        "Oskarshamn": (57.26, 16.45),
        "Karlskrona": (56.16, 15.59)
    }

    school_geo = {}

    for _, row in skolor.iterrows():
        skola = row["Skolenhet"]
        partner = str(row.get("Partnerområde", ""))

        if any(x in partner for x in ["Kalmar","Nybro","Mönsterås"]):
            base = base_geo["Kalmar"]
        elif "Oskarshamn" in partner:
            base = base_geo["Oskarshamn"]
        elif "Karlskrona" in partner:
            base = base_geo["Karlskrona"]
        else:
            continue

        h = int(hashlib.md5(skola.encode()).hexdigest(), 16)

        offset_lat = ((h % 1000) / 1000 - 0.5) * 0.15
        offset_lon = (((h // 1000) % 1000) / 1000 - 0.5) * 0.15

        school_geo[skola] = (
            base[0] + offset_lat,
            base[1] + offset_lon
        )

    def distance_km(student_ort, skola):
        student_ort = normalize_location(student_ort)

        if student_ort not in geo or skola not in school_geo:
            return None

        lat1, lon1 = geo[student_ort]
        lat2, lon2 = school_geo[skola]

        d = ((lat1 - lat2)**2 + (lon1 - lon2)**2)**0.5
        return round(d * 111, 1)

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]
    alt = [c for c in students.columns if "alternativ" in c.lower()][0]
    val = [c for c in students.columns if "helst" in c.lower()][0]

    def choose_loc(row):
        if "alternativ" in str(row.get(val,"")).lower():
            return row.get(alt)
        return row.get(bost)

    students["Ort"] = students.apply(choose_loc, axis=1)
    students["Ort"] = students["Ort"].apply(normalize_location)
    students["Namn"] = students[fn] + " " + students[ln]

    cap = {}
    result = []
    logg = []

    skol_lista = list(skolor["Skolenhet"])

    for i, (_, student) in enumerate(students.iterrows()):

        namn = student["Namn"]
        ort = student["Ort"]

        # ✅ sortera efter avstånd
        skol_sorted = sorted(
            skol_lista,
            key=lambda s: distance_km(ort, s) if distance_km(ort, s) else 999
        )

        # ✅ GROUP BEHAVIOR (försök hålla ihop)
        if i > 0:
            prev = students.iloc[i-1]["Namn"]
            prev_assignments = [r for r in result if prev in r.values()]

            if prev_assignments:
                skol_guess = [r["Skola"] for r in prev_assignments[:3]]

                if all(cap.get((s,1),0) < kap_map.get(s,999) for s in skol_guess):
                    A, B, C = skol_guess
                else:
                    A = skol_sorted[0]
                    B = skol_sorted[1 % len(skol_sorted)]
                    C = skol_sorted[2 % len(skol_sorted)]
            else:
                A = skol_sorted[0]
                B = skol_sorted[1 % len(skol_sorted)]
                C = skol_sorted[2 % len(skol_sorted)]
        else:
            A = skol_sorted[0]
            B = skol_sorted[1 % len(skol_sorted)]
            C = skol_sorted[2 % len(skol_sorted)]

        # ✅ kapasitet
        if (
            cap.get((A,1),0) >= kap_map.get(A,999) or
            cap.get((B,2),0) >= kap_map.get(B,999) or
            cap.get((B,3),0) >= kap_map.get(B,999) or
            cap.get((C,4),0) >= kap_map.get(C,999)
        ):
            logg.append({
                "Student":namn,
                "Status":"Får ej plats",
                "Kommentar":"",
                "Tips":"Kapacitet slut",
                "Avstånd":"-"
            })
            continue

        # ✅ update cap
        cap[(A,1)] = cap.get((A,1),0)+1
        cap[(B,2)] = cap.get((B,2),0)+1
        cap[(B,3)] = cap.get((B,3),0)+1
        cap[(C,4)] = cap.get((C,4),0)+1

        result += [
            {"Skola":A,"År 1":namn,"År 2":"","År 3":"","År 4":""},
            {"Skola":B,"År 1":"","År 2":namn,"År 3":namn,"År 4":""},
            {"Skola":C,"År 1":"","År 2":"","År 3":"","År 4":namn}
        ]

        # ✅ DISTANS
        distances = []
        for s in [A,B,C]:
            d = distance_km(ort, s)
            if d:
                distances.append((s,d))

        if distances:
            longest = max(distances, key=lambda x: x[1])
            avst = f"Längsta pendling: {longest[0]}, {longest[1]} km"
        else:
            avst = "Okänd"

        logg.append({
            "Student":namn,
            "Status":"OK",
            "Kommentar":"",
            "Tips":"",
            "Avstånd":avst
        })

    df = pd.DataFrame(result)

    # ===== BLOCK-LAYOUT =====
    wb = Workbook()
    ws = wb.active

    fill_header = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År 1","År 2","År 3","År 4"])

    skol_data = {}

    for _, r in df.iterrows():
        s = r["Skola"]
        skol_data.setdefault(s, {"År 1":[],"År 2":[],"År 3":[],"År 4":[]})

        for col in ["År 1","År 2","År 3","År 4"]:
            if r[col]:
                skol_data[s][col].append(r[col])

    for skola in skol_data:

        ws.append([f"{skola} (max {int(kap_map.get(skola,0))})"])

        data = skol_data[skola]
        max_len = max(len(v) for v in data.values())

        for i in range(max_len):
            ws.append([
                "",
                data["År 1"][i] if i<len(data["År 1"]) else "",
                data["År 2"][i] if i<len(data["År 2"]) else "",
                data["År 3"][i] if i<len(data["År 3"]) else "",
                data["År 4"][i] if i<len(data["År 4"]) else ""
            ])

        ws.append([])
        ws.append([])

    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status","Kommentar","Tips","Avstånd"])

    for r in logg:
        ws2.append([
            r["Student"],
            r["Status"],
            r["Kommentar"],
            r["Tips"],
            r["Avstånd"]
        ])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file, "rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp filer")
