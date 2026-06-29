import io
import random
import re
import time

import openpyxl
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="GruppenOptimierer", layout="wide")
st.title("GruppenOptimierer")
st.write(
    "Lädt eine Excel-Datei mit Vereinsadressen, berechnet die Fahrtkilometer-Matrix "
    "zwischen den Spielfeldern und teilt die Vereine in ausgewogene Gruppen mit "
    "minimaler interner Fahrtstrecke auf."
)


# ----------------------------------------------------------------------
# Kernlogik (identisch zu mycode.py)
# ----------------------------------------------------------------------
def clean_address(adresse):
    s = adresse.replace("  ", " ").strip()
    teile = [t.strip() for t in s.split(",")]
    strasse = None
    for t in teile:
        if re.search(r"(str\.?|straße|strasse|weg|gasse|ring|allee|wiese)", t, re.I) \
           and not re.search(r"rasenplatz|kunstrasen|hartplatz|stadion", t, re.I):
            strasse = t
            break
    plz_ort = None
    for t in teile:
        m = re.search(r"\b(\d{5})\s+(.+)", t)
        if m:
            ort = re.split(r"[-]| Zentrum| Innenstadt| Stadtgebiet| Süd| Nord| Ost| West",
                           m.group(2))[0].strip()
            plz_ort = f"{m.group(1)} {ort}"
            break
    varianten = []
    if strasse and plz_ort:
        varianten.append(f"{strasse}, {plz_ort}")
    if plz_ort:
        varianten.append(plz_ort)
    if not varianten:
        varianten.append(s)
    return varianten


def geocode(adresse, cache):
    if adresse in cache:
        return cache[adresse]
    headers = {"User-Agent": "spielfeld-abstandsmatrix/1.0"}
    result = None
    for variante in clean_address(adresse):
        r = requests.get("https://photon.komoot.io/api/",
                         params={"q": variante, "limit": 1, "lang": "de"},
                         headers=headers, timeout=30)
        r.raise_for_status()
        feats = r.json().get("features", [])
        time.sleep(1.0)
        if feats:
            lon, lat = feats[0]["geometry"]["coordinates"]
            result = (float(lat), float(lon))
            break
    cache[adresse] = result
    return result


def dist(matrix, i, j):
    a, b = matrix[i][j], matrix[j][i]
    if a is None and b is None:
        return 0.0
    if a is None:
        return b
    if b is None:
        return a
    return (a + b) / 2


def baue_ausgabe_excel(labels, matrix, gruppen):
    out = openpyxl.Workbook()
    sh = out.active
    sh.title = "Abstandsmatrix_km"
    sh.cell(row=1, column=1, value="V.Nr.")
    for j, lab in enumerate(labels):
        sh.cell(row=1, column=2 + j, value=lab)
    for i, lab in enumerate(labels):
        sh.cell(row=2 + i, column=1, value=lab)
        for j in range(len(labels)):
            sh.cell(row=2 + i, column=2 + j, value=matrix[i][j])

    sh2 = out.create_sheet("Gruppen")
    sh2.cell(row=1, column=1, value="Gruppe")
    sh2.cell(row=1, column=2, value="V.Nr.")
    z = 2
    for gi, grp in enumerate(gruppen, 1):
        for i in grp:
            sh2.cell(row=z, column=1, value=gi)
            sh2.cell(row=z, column=2, value=labels[i])
            z += 1

    sh2.cell(row=1, column=4, value="Gruppe")
    sh2.cell(row=1, column=5, value="Anzahl Vereine")
    sh2.cell(row=1, column=6, value="Distanzsumme (km)")
    sh2.cell(row=1, column=7, value="Ø je Vereinspaar (km)")
    sh2.cell(row=1, column=8, value="Ø je Verein (km)")
    for gi, grp in enumerate(gruppen, 1):
        n = len(grp)
        summe = sum(dist(matrix, a, b) for ai, a in enumerate(grp) for b in grp[ai + 1:])
        paare = n * (n - 1) / 2
        sh2.cell(row=1 + gi, column=4, value=gi)
        sh2.cell(row=1 + gi, column=5, value=n)
        sh2.cell(row=1 + gi, column=6, value=round(summe, 1))
        sh2.cell(row=1 + gi, column=7, value=round(summe / paare, 1) if paare else 0)
        sh2.cell(row=1 + gi, column=8, value=round(summe / n, 1) if n else 0)

    buf = io.BytesIO()
    out.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("Eingabe")
    datei = st.file_uploader("Meldelisten-Excel (.xlsx)", type=["xlsx"])
    sheet = st.text_input("Worksheet-Name", value="AZ")
    n_gruppen = st.number_input("Anzahl Gruppen", min_value=2, max_value=20, value=4, step=1)
    with st.expander("Erweiterte Einstellungen"):
        seed = st.number_input("Zufallssaat", value=42, step=1)
        versuche = st.number_input("Zufallsneustarts (Heuristik)", min_value=1, max_value=200, value=40, step=1)
    start = st.button("Berechnen", type="primary", disabled=datei is None)

if not datei:
    st.info("Bitte links eine Excel-Datei hochladen und die Anzahl der Gruppen wählen.")
    st.stop()

if not start:
    st.stop()

# ----------------------------------------------------------------------
# 1) Excel einlesen: Spalte A = V.Nr., Spalte L = Adresse
# ----------------------------------------------------------------------
wb = openpyxl.load_workbook(datei, data_only=True)
if sheet not in wb.sheetnames:
    st.error(f"Worksheet '{sheet}' nicht gefunden. Vorhanden: {', '.join(wb.sheetnames)}")
    st.stop()
ws = wb[sheet]

vereine = []
for row in ws.iter_rows(min_row=2, values_only=True):
    vnr = row[0]
    adr = row[11] if len(row) > 11 else None
    if vnr is None:
        continue
    adr = str(adr).strip() if adr else ""
    vereine.append((str(vnr), adr))

labels = [vnr for vnr, _ in vereine]
N = len(vereine)
st.write(f"**{N} Vereine eingelesen.**")

if N < int(n_gruppen):
    st.error(f"Nur {N} Vereine, aber {n_gruppen} Gruppen angefordert.")
    st.stop()

# ----------------------------------------------------------------------
# 2)+3) Geocoding (Photon)
# ----------------------------------------------------------------------
cache = {}
coords = []
nicht_gefunden = []
progress = st.progress(0.0, text="Geocoding...")
for idx, (vnr, adr) in enumerate(vereine):
    if not adr:
        coords.append(None)
        nicht_gefunden.append((vnr, "(keine Adresse)"))
    else:
        c = geocode(adr, cache)
        coords.append(c)
        if c is None:
            nicht_gefunden.append((vnr, adr))
    progress.progress((idx + 1) / N, text=f"Geocoding {idx + 1}/{N}: {vnr}")
progress.empty()

valid = [c is not None for c in coords]
gueltig = [i for i in range(N) if valid[i]]
M = len(gueltig)

if M < int(n_gruppen):
    st.error(f"Nur {M} Vereine konnten geocodiert werden, aber {n_gruppen} Gruppen angefordert.")
    st.stop()

# ----------------------------------------------------------------------
# 4) Fahrtkilometer-Matrix (OSRM)
# ----------------------------------------------------------------------
with st.spinner("Berechne Fahrtkilometer-Matrix (OSRM)..."):
    gueltige_coords = [coords[i] for i in gueltig]
    coord_str = ";".join(f"{lon},{lat}" for lat, lon in gueltige_coords)
    r = requests.get(f"https://router.project-osrm.org/table/v1/driving/{coord_str}",
                     params={"annotations": "distance"}, timeout=120)
    r.raise_for_status()
    osrm = r.json()["distances"]

matrix = [[None] * N for _ in range(N)]
for a, i in enumerate(gueltig):
    for b, j in enumerate(gueltig):
        d = osrm[a][b]
        matrix[i][j] = round(d / 1000, 1) if d is not None else None

# ----------------------------------------------------------------------
# 5) In N_GRUPPEN ausgewogene Gruppen aufteilen (minimale interne Distanz)
# ----------------------------------------------------------------------
random.seed(int(seed))
n_gruppen = int(n_gruppen)
basis, rest = M // n_gruppen, M % n_gruppen
groessen = [basis + (1 if g < rest else 0) for g in range(n_gruppen)]


def gruppen_kosten(m):
    s = 0.0
    for x in range(len(m)):
        for y in range(x + 1, len(m)):
            s += dist(matrix, m[x], m[y])
    return s


def gesamt_kosten(gr):
    return sum(gruppen_kosten(g) for g in gr)


def startloesung():
    pool = gueltig[:]
    random.shuffle(pool)
    gr, pos = [], 0
    for g in range(n_gruppen):
        gr.append(pool[pos:pos + groessen[g]])
        pos += groessen[g]
    return gr


def optimiere(gr):
    verbessert = True
    while verbessert:
        verbessert = False
        for ga in range(n_gruppen):
            for gb in range(ga + 1, n_gruppen):
                for ia in range(len(gr[ga])):
                    for ib in range(len(gr[gb])):
                        vor = gruppen_kosten(gr[ga]) + gruppen_kosten(gr[gb])
                        gr[ga][ia], gr[gb][ib] = gr[gb][ib], gr[ga][ia]
                        nach = gruppen_kosten(gr[ga]) + gruppen_kosten(gr[gb])
                        if nach < vor - 1e-9:
                            verbessert = True
                        else:
                            gr[ga][ia], gr[gb][ib] = gr[gb][ib], gr[ga][ia]
    return gr


with st.spinner(f"Optimiere Gruppen ({int(versuche)} Zufallsneustarts)..."):
    beste, beste_kosten = None, float("inf")
    alle_kosten = []
    for _ in range(int(versuche)):
        g = optimiere(startloesung())
        k = gesamt_kosten(g)
        alle_kosten.append(k)
        if k < beste_kosten:
            beste, beste_kosten = [grp[:] for grp in g], k

mittlere_kosten = sum(alle_kosten) / len(alle_kosten)
groesste_kosten = max(alle_kosten)

# ----------------------------------------------------------------------
# Ergebnis anzeigen
# ----------------------------------------------------------------------
st.success(
    f"{n_gruppen} Gruppen, Gesamt-Distanzsumme innerhalb: {beste_kosten:.1f} km "
    f"(Schnitt/Verein {beste_kosten / M:.1f} km)"
)
st.caption(
    f"Referenz aus {len(alle_kosten)} optimierten Versuchen: "
    f"Mittelwert {mittlere_kosten:.1f} km, Maximum {groesste_kosten:.1f} km"
)

for gi, grp in enumerate(beste, 1):
    n = len(grp)
    summe = gruppen_kosten(grp)
    paare = n * (n - 1) / 2
    avg_paar = summe / paare if paare else 0.0
    avg_verein = summe / n if n else 0.0
    with st.expander(
        f"Gruppe {gi} ({n} Vereine, Distanzsumme {summe:.1f} km, "
        f"Ø/Paar {avg_paar:.1f} km, Ø/Verein {avg_verein:.1f} km)",
        expanded=True,
    ):
        st.write(", ".join(labels[i] for i in grp))

ohne = [labels[i] for i in range(N) if not valid[i]]
if ohne:
    st.warning("Nicht zugeordnet (keine Koordinate): " + ", ".join(ohne))

st.subheader("Distanzmatrix (km)")
df = pd.DataFrame(matrix, index=labels, columns=labels)
st.dataframe(df)

excel_bytes = baue_ausgabe_excel(labels, matrix, beste)
st.download_button(
    "Abstandsmatrix.xlsx herunterladen",
    data=excel_bytes,
    file_name="Abstandsmatrix.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
