
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ==============================
# INSTÄLLNINGAR
# ==============================
INPUT_FILE = "kull_resultat (35).xlsx"
OUTPUT_FILE = "kull_resultat_FIXAD.xlsx"

# ==============================
# LÄS DATA
# ==============================
df = pd.read_excel(INPUT_FILE, sheet_name="Studenter", engine="openpyxl")
df.columns = df.columns.str.strip()

# Säkerställ kolumner
df = df[["Student", "Ort", "Region", "År1", "År2", "År3", "År4"]]

# ==============================
# SKAPA ALLA SKOLOR
# ==============================
alla_skolor = pd.unique(df[["År1", "År2", "År3", "År4"]].values.ravel())
alla_skolor = [s for s in alla_skolor if pd.notna(s)]
alla_skolor = sorted(alla_skolor)

# ==============================
# BYGG STRUKTUR PER SKOLA
# ==============================
struktur = {}

for skola in alla_skolor:
    struktur[skola] = {
        "År1": [],
        "År2": [],
        "År3": [],
        "År4": []
    }

# Fyll på korrekt (INGEN duplicering)
for _, row in df.iterrows():
    student = row["Student"]

    for år in ["År1", "År2", "År3", "År4"]:
        skola = row[år]
        if pd.notna(skola):
            struktur[skola][år].append(student)

# ==============================
# SKAPA EXCEL
# ==============================
wb = Workbook()

# ==============================
# SHEET 1 – PLACERINGAR
# ==============================
ws1 = wb.active
ws1.title = "Placeringar"

headers = ["Skola", "År1", "År2", "År3", "År4"]

# Skriv header
for col, header in enumerate(headers, start=1):
    cell = ws1.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True)

row_idx = 2

for skola in alla_skolor:

    årdata = struktur[skola]

    max_len = max(
        len(årdata["År1"]),
        len(årdata["År2"]),
        len(årdata["År3"]),
        len(årdata["År4"]),
        1
    )

    for i in range(max_len):
        ws1.cell(row=row_idx, column=1, value=skola if i == 0 else "")

        ws1.cell(row=row_idx, column=2,
                value=årdata["År1"][i] if i < len(årdata["År1"]) else "")

        ws1.cell(row=row_idx, column=3,
                value=årdata["År2"][i] if i < len(årdata["År2"]) else "")

        ws1.cell(row=row_idx, column=4,
                value=årdata["År3"][i] if i < len(årdata["År3"]) else "")

        ws1.cell(row=row_idx, column=5,
                value=årdata["År4"][i] if i < len(årdata["År4"]) else "")

        row_idx += 1

# Justera kolumnbredd
for col in range(1, 6):
    ws1.column_dimensions[get_column_letter(col)].width = 28

# ==============================
# SHEET 2 – RAPPORT
# ==============================
ws2 = wb.create_sheet(title="Rapport")

ws2["A1"] = "Student"
ws2["B1"] = "Status"

ws2["A1"].font = Font(bold=True)
ws2["B1"].font = Font(bold=True)

for idx, row in df.iterrows():

    student = row["Student"]

    status = "OK"

    # Enkel pendling-check (kan byggas ut)
    if row["Ort"] not in str(row["År1"]):
        status = "OK"

    ws2.cell(row=idx + 2, column=1, value=student)
    ws2.cell(row=idx + 2, column=2, value=status)

ws2.column_dimensions["A"].width = 28
ws2.column_dimensions["B"].width = 35

# ==============================
# SHEET 3 – KONTROLL
# ==============================
ws3 = wb.create_sheet(title="Kontroll")

ws3["A1"] = "Student"
ws3["B1"] = "Antal skolor"

ws3["A1"].font = Font(bold=True)
ws3["B1"].font = Font(bold=True)

for idx, row in df.iterrows():

    student = row["Student"]

    skolor = set([
        row["År1"],
        row["År2"],
        row["År3"],
        row["År4"]
    ])

    skolor = [s for s in skolor if pd.notna(s)]

    ws3.cell(row=idx + 2, column=1, value=student)
    ws3.cell(row=idx + 2, column=2, value=len(skolor))

ws3.column_dimensions["A"].width = 28
ws3.column_dimensions["B"].width = 20

# ==============================
# FORMAT (EXTRA)
# ==============================

# Centrera allt vertikalt (snyggt)
for ws in [ws1, ws2, ws3]:
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top")

# ==============================
# SPARA
# ==============================
wb.save(OUTPUT_FILE)

print("KLART ✅ Fil skapad:", OUTPUT_FILE)
``
