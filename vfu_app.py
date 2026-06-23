
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ==============================
# FILER
# ==============================
STUDENT_FIL = "kull_resultat (35).xlsx"
OVERSIKT_FIL = "Helhetsbild övningsskolor.xlsx"
OUTPUT_FIL = "VFU_RESULTAT_KLAR.xlsx"

# ==============================
# LÄS IN DATA
# ==============================
stud_df = pd.read_excel(STUDENT_FIL, sheet_name="Studenter", engine="openpyxl")
skol_df = pd.read_excel(OVERSIKT_FIL, sheet_name="SKOLOR", engine="openpyxl")

stud_df.columns = stud_df.columns.str.strip()
skol_df.columns = skol_df.columns.str.strip()

# ==============================
# FILTRERA – LAFOV 26
# ==============================
lafov_df = skol_df[
    (skol_df["Kull"] == 26) &
    (skol_df["Inriktning"] == "LAFOV")
]

# ==============================
# BYGG KAPACITET
# ==============================
kapacitet = {}

for _, row in lafov_df.iterrows():
    skola = row["Skolenhet"]
    platser = row["Antal platser"]

    if pd.notna(skola) and pd.notna(platser):
        kapacitet[skola] = int(platser)

# ==============================
# INITIERA PLACERING
# ==============================
placering = defaultdict(list)

# sortera skolor (störst först = bättre balans)
skolor_sorterade = sorted(kapacitet, key=lambda x: kapacitet[x], reverse=True)

# ==============================
# FÖRDELNINGSALGORITM
# ==============================
for _, row in stud_df.iterrows():

    student = row["Student"]

    placerad = False

    # 1. Försök placera i ordning (balanserat)
    for skola in skolor_sorterade:
        if len(placering[skola]) < kapacitet[skola]:
            placering[skola].append(student)
            placerad = True
            break

    # 2. Om inget funkar (edge case)
    if not placerad:
        minst = min(placering, key=lambda x: len(placering[x]))
        placering[minst].append(student)

# ==============================
# SKAPA EXCEL
# ==============================
wb = Workbook()

# ==============================
# 1. PLACERINGAR (MANUELL LAYOUT)
# ==============================
ws1 = wb.active
ws1.title = "Placeringar"

headers = ["Skola", "Studenter"]
ws1.append(headers)

for cell in ws1[1]:
    cell.font = Font(bold=True)

row_idx = 2

for skola in skolor_sorterade:

    studenter = placering[skola]

    if len(studenter) == 0:
        ws1.append([skola, ""])
        continue

    for i, s in enumerate(studenter):
        if i == 0:
            ws1.append([f"{skola} (max {kapacitet[skola]})", s])
        else:
            ws1.append(["", s])

# ==============================
# 2. RAPPORT (MED OK-FUNKTION)
# ==============================
ws2 = wb.create_sheet("Rapport")

headers = ["Student", "Hemort", "Vald ort", "OK"]
ws2.append(headers)

for cell in ws2[1]:
    cell.font = Font(bold=True)

for _, row in stud_df.iterrows():

    ws2.append([
        row["Student"],
        row["Ort"],
        "",   # väljs manuellt
        ""    # skriv OK här
    ])

# ==============================
# 3. KONTROLL
# ==============================
ws3 = wb.create_sheet("Kontroll")

headers = ["Student", "Status", "Kommentar"]
ws3.append(headers)

for cell in ws3[1]:
    cell.font = Font(bold=True)

for _, row in stud_df.iterrows():

    ws3.append([
        row["Student"],
        "Väntar på OK",
        ""
    ])

# ==============================
# FÄRGER (VISUELLT STÖD)
# ==============================
green_fill = PatternFill(start_color="C6EFCE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFF3CD", fill_type="solid")

# ==============================
# KOPPLA RAPPORT → KONTROLL
# ==============================
for i in range(2, ws2.max_row + 1):

    ok_value = ws2.cell(row=i, column=4).value
    student = ws2.cell(row=i, column=1).value

    kontroll_rad = i

    if ok_value == "OK":
        ws3.cell(row=kontroll_rad, column=2, value="Klar")
        ws3.cell(row=kontroll_rad, column=2).fill = green_fill
    else:
        ws3.cell(row=kontroll_rad, column=2, value="Ej klar")
        ws3.cell(row=kontroll_rad, column=2).fill = yellow_fill

# ==============================
# AUTO-BREDD
# ==============================
for ws in [ws1, ws2, ws3]:
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[col_letter].width = max_length + 2

# ==============================
# SPARA FIL
# ==============================
wb.save(OUTPUT_FIL)

print("✅ KLAR – full version skapad:", OUTPUT_FIL)
