
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Välj inriktning", ["LAFOV", "LAGRV", "LGFRI"])

def get_region(text):
    text = str(text)
    if any(x in text for x in ["Kalmar","Nybro","Mönsterås"]):
        return "Kalmarregion"
    if "Oskarshamn" in text:
        return "Oskarshamn"
    if "Karlskrona" in text:
        return "Karlskrona"
    return "Kalmarregion"

def clean_text(text):
    return str(text).lower().replace(" ","").replace("-","")

def match_school(a,s):
    return clean_text(a) in clean_text(s) or clean_text(s) in clean_text(a)

def find_column(cols,keywords):
    for c in cols:
        if any(k.lower() in c.lower() for k in keywords):
            return c
    return None

if system_file and form_file:

    try:
        # ===== SKOLOR =====
        skolor = pd.read_excel(system_file)
        skolor.columns = skolor.columns.str.strip()

        skolor = skolor[
            (skolor["Kull"] == kull) &
            (skolor["Inriktning"].str.upper() == program)
        ].copy()

        skolor["Region"] = skolor["Partnerområde"].apply(get_region)

        kap_map = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))
        region_map = dict(zip(skolor["Skolenhet"], skolor["Region"]))

        # ===== STUDENTER =====
        students = pd.read_excel(form_file, sheet_name="Data")
        students.columns = students.columns.str.strip()

        fn = find_column(students.columns,["förnamn"])
        ln = find_column(students.columns,["efternamn"])
        bost = find_column(students.columns,["bostadsort"])
        ank = find_column(students.columns,["anknytning"])

        students["Namn"] = students[fn] + " " + students[ln]
        students["Region"] = students[bost].apply(get_region)

        best_result, best_log = None, None
        best_unplaced = 999

        # ===== OPTIMERING =====
        for _ in range(30):

            result, logg, ej_placerade = [], [], []
            cap = {}

            s_run = students.sample(frac=1)

            for region in s_run["Region"].unique():

                stud = s_run[s_run["Region"] == region]
                skol_lista_full = list(skolor[skolor["Region"]==region]["Skolenhet"])

                for i,(_, student) in enumerate(stud.iterrows()):

                    namn = student["Namn"]
                    ank_raw = str(student.get(ank,"")).strip()

                    if ank_raw.lower() in ["","ingen","-","nej"]:
                        skol_lista = skol_lista_full
                        exkl = []
                    else:
                        skol_lista, exkl = [], []
                        for s in skol_lista_full:
                            if match_school(ank_raw, s):
                                exkl.append(s)
                            else:
                                skol_lista.append(s)
                        if not skol_lista:
                            skol_lista = skol_lista_full

                    placed = False
                    status = "OK"
                    kommentar = ""

                    # ===== LGFRI =====
                    if program == "LGFRI":

                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]

                            if (
                                cap.get((A,1),0)<kap_map[A] and
                                cap.get((A,2),0)<kap_map[A] and
                                cap.get((B,3),0)<kap_map[B]
                            ):
                                placed=True
                                break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({"Student":namn,"Status":"Får ej plats","Kommentar":""})
                            continue

                        cap[(A,1)] = cap.get((A,1),0)+1
                        cap[(A,2)] = cap.get((A,2),0)+1
                        cap[(B,3)] = cap.get((B,3),0)+1

                        result += [
                            {"Skola":A,"År 1":namn,"År 2":namn,"År 3":""},
                            {"Skola":B,"År 1":"","År 2":"","År 3":namn}
                        ]

                    else:
                        # FULL
                        for shift in range(len(skol_lista)):
                            A = skol_lista[(i+shift)%len(skol_lista)]
                            B = skol_lista[(i+1+shift)%len(skol_lista)]
                            C = skol_lista[(i+2+shift)%len(skol_lista)]

                            if (
                                cap.get((A,1),0)<kap_map[A] and
                                cap.get((B,2),0)<kap_map[B] and
                                cap.get((B,3),0)<kap_map[B] and
                                cap.get((C,4),0)<kap_map[C]
                            ):
                                placed=True
                                break

                        # fallback
                        if not placed:
                            for A in skol_lista:
                                for B in skol_lista:
                                    if A!=B:
                                        if (
                                            cap.get((A,1),0)<kap_map[A] and
                                            cap.get((B,2),0)<kap_map[B] and
                                            cap.get((B,3),0)<kap_map[B] and
                                            cap.get((B,4),0)<kap_map[B]
                                        ):
                                            C=B
                                            placed=True
                                            status="Avvikelse"
                                            kommentar="Fallback använd"
                                            break
                                if placed: break

                        if not placed:
                            ej_placerade.append(namn)
                            logg.append({"Student":namn,"Status":"Får ej plats","Kommentar":""})
                            continue

                        cap[(A,1)] +=1 if (A,1) in cap else cap.setdefault((A,1),1)
                        cap[(B,2)] +=1 if (B,2) in cap else cap.setdefault((B,2),1)
                        cap[(B,3)] +=1 if (B,3) in cap else cap.setdefault((B,3),1)
                        cap[(C,4)] +=1 if (C,4) in cap else cap.setdefault((C,4),1)

                        result += [
                            {"Skola":A,"År1":namn,"År2":"","År3":"","År4":""},
                            {"Skola":B,"År1":"","År2":namn,"År3":namn,"År4":""},
                            {"Skola":C,"År1":"","År2":"","År3":"","År4":namn}
                        ]

                    if exkl:
                        status="Avvikelse"
                        kommentar+=" Anknytning"

                    logg.append({"Student":namn,"Status":status,"Kommentar":kommentar})

            if len(ej_placerade)<best_unplaced:
                best_unplaced=len(ej_placerade)
                best_result=result
                best_log=logg

        df = pd.DataFrame(best_result)

        # ===== GROUP PER SKOLA =====
        skol_data={}
        for _,r in df.iterrows():
            s=r["Skola"]
            if s not in skol_data:
                skol_data[s]={c:[] for c in df.columns if "År" in c}
            for c in skol_data[s]:
                if r.get(c): skol_data[s][c].append(r[c])

        def region_order(s):
            return {"Kalmarregion":1,"Oskarshamn":2,"Karlskrona":3}.get(region_map.get(s,""),0)

        sorted_skolor=sorted(skol_data.keys(), key=lambda x:(region_order(x),x))

        # ===== EXCEL =====
        wb=Workbook()
        ws=wb.active

        # col widths
        ws.column_dimensions["A"].width=40
        for c in ["B","C","D","E"]:
            ws.column_dimensions[c].width=28

        fill_header=PatternFill(start_color="DDDDDD",fill_type="solid")
        fill_green=PatternFill(start_color="CCFFCC",fill_type="solid")
        fill_dark=PatternFill(start_color="99CC66",fill_type="solid")

        thin=Side(style="thin")
        thick=Side(style="medium")

        align=Alignment(vertical="center",wrap_text=True)

        ws.append(df.columns.tolist())

        current_region=None

        for skola in sorted_skolor:

            region=region_map.get(skola,"")

            if region!=current_region:
                ws.append([region.upper()])
                current_region=region

            data=skol_data[skola]
            kap=kap_map.get(skola,"-")
            antal=len(set(sum(data.values(),[])))

            start=ws.max_row+1
            ws.append([f"{skola} ({antal}/{kap})"])

            ws.merge_cells(start_row=start,start_column=1,end_row=start,end_column=5)

            for col in range(1,6):
                ws.cell(start,col).fill=fill_header
                ws.cell(start,col).font=Font(bold=True)

            max_len=max(len(v) for v in data.values())

            cols=list(data.keys())

            for i in range(max_len):
                row=[""]
                for c in cols:
                    row.append(data[c][i] if i<len(data[c]) else "")
                ws.append(row)
                r=ws.max_row

                for c in range(2,len(row)+1):
                    ws.cell(r,c).alignment=align

                if len(row)>=5:
                    ws.cell(r,3).fill=fill_green
                    ws.cell(r,4).fill=fill_green
                    ws.cell(r,5).fill=fill_dark

                for c in range(1,len(row)+1):
                    ws.cell(r,c).border=Border(left=thin,right=thin,top=thin,bottom=thin)

            end=ws.max_row

            for rr in range(start,end+1):
                for cc in range(1,len(cols)+2):
                    ws.cell(rr,cc).border=Border(
                        left=thick if cc==1 else thin,
                        right=thick if cc==len(cols)+1 else thin,
                        top=thick if rr==start else thin,
                        bottom=thick if rr==end else thin
                    )

            ws.append([])
            ws.append([])

        # ===== RAPPORT =====
        ws2=wb.create_sheet("Rapport")
        ws2.append(["Student","Status","Kommentar"])
        for r in best_log:
            ws2.append([r["Student"],r["Status"],r["Kommentar"]])

        file="kull_resultat.xlsx"
        wb.save(file)

        with open(file,"rb") as f:
            st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

        st.success(f"✅ Klar – {best_unplaced} ej placerade")

    except Exception as e:
        st.error(e)

else:
    st.info("Ladda upp filer")
