import streamlit as st
import pandas as pd

st.title("VFU-system – Placering och scenario")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
scenario_students = st.number_input("Scenario: antal studenter (valfritt)", value=0)

def get_group(bostadsort):
    if "Kalmar" in str(bostadsort):
        return "Kalmar"
    if "Nybro" in str(bostadsort):
        return "Nybro"
    if "Karlskrona" in str(bostadsort):
        return "Karlskrona"
    if "Oskarshamn" in str(bostadsort):
        return "Oskarshamn"
    return "Övrigt"

if system_file and form_file:

    skolor = pd.read_excel(system_file, sheet_name="SKOLOR")
    students = pd.read_excel(form_file)

    students["Grupp"] = students["Bostadsort"].apply(get_group)

    arv_kull = kull - 4
    skolor = skolor[skolor["Aktiv kull"] == arv_kull]

    result = []
    capacity_counter = {}

    for grupp in students["Grupp"].unique():

        stud_grp = students[students["Grupp"] == grupp]
        skol_grp = skolor[skolor["Grupp"] == grupp]

        skol_lista = list(skol_grp["Skola"])

        kapacitet_map = dict(zip(skol_grp["Skola"], skol_grp["Kapacitet"]))

        if len(skol_lista) == 0:
            st.warning(f"Inga skolor för grupp: {grupp}")
            continue

        for i, (_, student) in enumerate(stud_grp.iterrows()):

            namn = f"{student['Förnamn']} {student['Efternamn']}"

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

    st.subheader("Placering")
    st.dataframe(df_final)

    # scenario
    if scenario_students > 0:

        st.subheader("Scenario")

        scenario_result = df_final.copy()
        scenario_result["Belastning"] = scenario_result["År 2"].str.count(",") + 1

        skol_kap = dict(zip(skolor["Skola"], skolor["Kapacitet"]))

        scenario_result["Kapacitet"] = scenario_result["Skola"].map(skol_kap)

        def status(row):
            if row["Belastning"] > row["Kapacitet"]:
                return "ÖVER"
            elif row["Belastning"] > row["Kapacitet"] * 0.8:
                return "HÖG"
            else:
                return "OK"

        scenario_result["Status"] = scenario_result.apply(status, axis=1)

        st.dataframe(scenario_result)

    # export
    output_file = "kull_resultat.xlsx"
    df_final.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button("Ladda ner Excel", f, file_name=output_file)

else:
    st.info("Ladda upp båda filer.")