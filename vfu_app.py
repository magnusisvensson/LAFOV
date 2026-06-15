
import streamlit as st
import pandas as pd

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)

# === STUDENTGRUPP ===
def get_student_group(bostadsort):
    bostadsort = str(bostadsort)

    # ✅ Kalmarregion (pendling till Kalmar)
    if any(x in bostadsort for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"

    if "Karlskrona" in bostadsort:
        return "Karlskrona"

    if "Oskarshamn" in bostadsort:
        return "Oskarshamn"

    return "Övrigt"


# === SKOLGRUPP ===
def get_school_group(partnerområde):
    område = str(partnerområde)

    if any(x in område for x in ["Kalmar", "Nybro", "Mönsterås"]):
        return "Kalmarregion"

    if "Karlskrona" in område:
        return "Karlskrona"

    if "Oskarshamn" in område:
        return "Oskarshamn"

    return "Övrigt"


if system_file and form_file:

    try:
        # === SKOLOR ===
        skolor = pd.read_excel(system_file, sheet_name=0)
        skolor.columns = skolor.columns.str.strip()

        st.write("Kolumner i skolfil:", list(skolor.columns))

        skolkol = "Skolenhet"
        gruppkol = "Partnerområde"
        kull_kol = "Kull"
        kap_kol = "Antal platser"

        # ✅ skapa grupp
        skolor["Grupp"] = skolor[gruppkol].apply(get_school_group)

        # ✅ VIKTIGT: använd skolor från samma kull (inte arv just nu)
        skolor = skolor[skolor[kull_kol] == kull]

        st.write("Antal skolor:", len(skolor))

        # === STUDENTER ===
        students = pd.read_excel(form_file)
        students.columns = students.columns.str.strip()

        st.write("Kolumner i studentfil:", list(students.columns))

        fnamn_kol = "Förnamn"
        enamn_kol = "Efternamn"
        ort_kol = "Bostadsort"

        students["Grupp"] = students[ort_kol].apply(get_student_group)

        result = []
        capacity_counter = {}

        st.write("Kör placering...")

        # === LOOP ===
        for grupp in students["Grupp"].unique():

            stud_grp = students[students["Grupp"] == grupp]
            skol_grp = skolor[skolor["Grupp"] == grupp]

            skol_lista = list(skol_grp[skolkol])

            kapacitet_map = dict(zip(skol_grp[skolkol], skol_grp[kap_kol]))

            if len(skol_lista) == 0:
                st.warning(f"Inga skolor för grupp: {grupp}")
                continue

            for i, (_, student) in enumerate(stud_grp.iterrows()):

                namn = f"{student[fnamn_kol]} {student[enamn_kol]}"

                # ✅ rotation + kapacitet
                for shift in range(len(skol_lista)):
                    A = skol_lista[(i + shift) % len(skol_lista)]
                    B = skol_lista[(i + 1 + shift) % len(skol_lista)]
                    C = skol_lista[(i + 2 + shift) % len(skol_lista)]

                    count = capacity_counter.get((B, 2), 0)

                    if count < kapacitet_map.get(B, 999):
                        break

                capacity_counter[(B, 2)] = capacity_counter.get((B, 2), 0) + 1
                capacity_counter[(B, 3)] = capacity_counter.get((B, 3), 0) + 1

                # bygg output
                result.append([A, namn, "", "", ""])
                result.append([B, "", namn, namn, ""])
                result.append([C, "", "", "", namn])

        # === RESULTAT ===
        df = pd.DataFrame(result, columns=["Skola", "År 1", "År 2", "År 3", "År 4"])

        df_final = df.groupby("Skola").agg({
            "År 1": lambda x: ", ".join(filter(None, x)),
            "År 2": lambda x: ", ".join(filter(None, x)),
            "År 3": lambda x: ", ".join(filter(None, x)),
            "År 4": lambda x: ", ".join(filter(None, x)),
        }).reset_index()

        st.subheader("✅ Placering")
        st.dataframe(df_final)

        # === EXPORT ===
        output_file = "kull_resultat.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button("⬇️ Ladda ner Excel", f, file_name=output_file)

    except Exception as e:
        st.error("Fel i appen:")
        st.write(e)

else:
    st.info("Ladda upp båda filer.")
