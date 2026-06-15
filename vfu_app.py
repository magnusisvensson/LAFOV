
import streamlit as st
import pandas as pd

st.title("VFU-system – Placering och scenario")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
scenario_students = st.number_input("Scenario: antal studenter (valfritt)", value=0)

# === Gruppfunktion ===
def get_group(bostadsort):
    bostadsort = str(bostadsort)
    if "Kalmar" in bostadsort:
        return "Kalmar"
    if "Nybro" in bostadsort:
        return "Nybro"
    if "Karlskrona" in bostadsort:
        return "Karlskrona"
    if "Oskarshamn" in bostadsort:
        return "Oskarshamn"
    return "Övrigt"

if system_file and form_file:

    try:
        # === Läs skolfil ===
        skolor = pd.read_excel(system_file, sheet_name=0)
        skolor.columns = skolor.columns.str.strip()

        st.write("Kolumner i skolfil:", list(skolor.columns))

        # Hitta rätt kolumner automatiskt
        skolkol = [c for c in skolor.columns if "skola" in c.lower()][0]
        gruppkol = [c for c in skolor.columns if "grupp" in c.lower() or "kommun" in c.lower()][0]
        
        kull_kol = [c for c in skolor.columns if "kull" in c.lower()][0]

        kap_kol = None
        for c in skolor.columns:
            if "kapacitet" in c.lower():
                kap_kol = c

        arv_kull = kull - 4

        # Filtrera skolor
        skolor = skolor[skolor[kull_kol] == arv_kull]

        # === Läs studentfil ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()

        st.write("Kolumner i studentfil:", list(students.columns))

        # hitta kolumner dynamiskt
        fnamn_kol = [c for c in students.columns if "förnamn" in c.lower()][0]
        enamn_kol = [c for c in students.columns if "efternamn" in c.lower()][0]
        ort_kol = [c for c in students.columns if "ort" in c.lower() or "bostad" in c.lower()][0]

        students["Grupp"] = students[ort_kol].apply(get_group)

        result = []
        capacity_counter = {}

        st.write("Bearbetar placering...")

        # === Loop per grupp ===
        for grupp in students["Grupp"].unique():

            stud_grp = students[students["Grupp"] == grupp]
            skol_grp = skolor[skolor[gruppkol] == grupp]

            skol_lista = list(skol_grp[skolkol])

            # Kapacitet
            if kap_kol:
                kapacitet_map = dict(zip(skol_grp[skolkol], skol_grp[kap_kol]))
            else:
                kapacitet_map = {s: 999 for s in skol_lista}

            if len(skol_lista) == 0:
                st.warning(f"Inga skolor för grupp: {grupp}")
                continue

            # === Placering ===
            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student[fnamn_kol]} {student[enamn_kol]}"

                for shift in range(len(skol_lista)):
                    A = skol_lista[(i + shift) % len(skol_lista)]
                    B = skol_lista[(i + 1 + shift) % len(skol_lista)]
                    C = skol_lista[(i + 2 + shift) % len(skol_lista)]

                    count = capacity_counter.get((B,2),0)

                    if count < kapacitet_map.get(B, 999):
                        break

                capacity_counter[(B,2)] = capacity_counter.get((B,2),0) + 1
                capacity_counter[(B,3)] = capacity_counter.get((B,3),0) + 1

                result.append([A, namn, "", "", ""])
                result.append([B, "", namn, namn, ""])
                result.append([C, "", "", "", namn])

        df = pd.DataFrame(result, columns=["Skola","År 1","År 2","År 3","År 4"])

        df_final = df.groupby("Skola").agg({
            "År 1": lambda x: ", ".join(filter(None,x)),
            "År 2": lambda x: ", ".join(filter(None,x)),
            "År 3": lambda x: ", ".join(filter(None,x)),
            "År 4": lambda x: ", ".join(filter(None,x)),
        }).reset_index()

        st.subheader("✅ Placering")
        st.dataframe(df_final)

        # === Scenario ===
        if scenario_students > 0:

            st.subheader("Scenario")

            scenario_df = df_final.copy()

            scenario_df["Belastning"] = scenario_df["År 2"].apply(lambda x: len(str(x).split(",")) if x else 0)

            if kap_kol:
                skol_kap = dict(zip(skol_grp[skolkol], skol_grp[kap_kol]))
                scenario_df["Kapacitet"] = scenario_df["Skola"].map(skol_kap)
            else:
                scenario_df["Kapacitet"] = 999

            def status(row):
                if row["Belastning"] > row["Kapacitet"]:
                    return "ÖVER"
                elif row["Belastning"] > row["Kapacitet"] * 0.8:
                    return "HÖG"
                else:
                    return "OK"

            scenario_df["Status"] = scenario_df.apply(status, axis=1)

            st.dataframe(scenario_df)

        # === Export ===
        output_file = "kull_resultat.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=output_file)

    except Exception as e:
        st.error("Fel i appen:")
        st.write(e)

else:
    st.info("Ladda upp båda filer.")
