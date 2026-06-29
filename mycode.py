import os
import re
import time
import random
import openpyxl
import requests

EXCEL_IN  = "Meldelisten-mit-Sportstätten_A1.xlsx"
EXCEL_OUT = "Abstandsmatrix.xlsx"
SHEET     = "AZ"
N_GRUPPEN = 4          # <- hier später variieren
SEED      = 42

# ----------------------------------------------------------------------
# 1) Excel einlesen: Spalte A = V.Nr., Spalte L = Adresse
# ----------------------------------------------------------------------
wb = openpyxl.load_workbook(EXCEL_IN, data_only=True)
ws = wb[SHEET]

vereine = []
for row in ws.iter_rows(min_row=2, values_only=True):
    vnr = row[0]
    adr = row[11]
    if vnr is None:
        continue
    adr = str(adr).strip() if adr else ""
    vereine.append((str(vnr), adr))

print(f"{len(vereine)} Vereine eingelesen.")

labels = [vnr for vnr, _ in vereine]
N = len(vereine)

# ----------------------------------------------------------------------
# 2) Adresse bereinigen
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

# ----------------------------------------------------------------------
# 3) Geocoding (Photon)
# ----------------------------------------------------------------------
def geocode(adresse, cache={}):
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

nicht_gefunden = []

if os.path.exists(EXCEL_OUT):
    print(f"\n{EXCEL_OUT} existiert bereits -> Matrix wird daraus geladen (kein erneutes Geocoding/OSRM).")
    cache_wb = openpyxl.load_workbook(EXCEL_OUT, data_only=True)
    cache_sh = cache_wb["Abstandsmatrix_km"]
    cached_labels = [cache_sh.cell(row=1, column=2 + j).value for j in range(N)]
    if cached_labels != labels:
        raise SystemExit(
            f"{EXCEL_OUT} passt nicht zur aktuellen Vereinsliste in {EXCEL_IN}. "
            f"Bitte {EXCEL_OUT} löschen, um die Matrix neu zu berechnen."
        )
    matrix = [[cache_sh.cell(row=2 + i, column=2 + j).value for j in range(N)] for i in range(N)]
    valid = [any(v is not None for v in matrix[i]) for i in range(N)]
    for (vnr, adr), ok in zip(vereine, valid):
        if not ok:
            nicht_gefunden.append((vnr, adr if adr else "(keine Adresse)"))
else:
    print(f"\n{EXCEL_OUT} noch nicht vorhanden -> Matrix wird neu berechnet (Geocoding + OSRM).")
    coords = []
    for vnr, adr in vereine:
        if not adr:
            coords.append(None)
            nicht_gefunden.append((vnr, "(keine Adresse)"))
            print(f"{vnr}: KEINE Adresse")
            continue
        c = geocode(adr)
        coords.append(c)
        if c is None:
            nicht_gefunden.append((vnr, adr))
            print(f"{vnr}: NICHT gefunden -> {adr}")
        else:
            print(f"{vnr}: {c[0]:.5f}, {c[1]:.5f}")

    # ------------------------------------------------------------------
    # 4) Fahrtkilometer-Matrix (OSRM)
    # ------------------------------------------------------------------
    gueltig_idx = [i for i, c in enumerate(coords) if c is not None]
    gueltige_coords = [coords[i] for i in gueltig_idx]

    coord_str = ";".join(f"{lon},{lat}" for lat, lon in gueltige_coords)
    r = requests.get(f"https://router.project-osrm.org/table/v1/driving/{coord_str}",
                     params={"annotations": "distance"}, timeout=120)
    r.raise_for_status()
    osrm = r.json()["distances"]

    matrix = [[None] * N for _ in range(N)]
    for a, i in enumerate(gueltig_idx):
        for b, j in enumerate(gueltig_idx):
            d = osrm[a][b]
            matrix[i][j] = round(d / 1000, 1) if d is not None else None

    valid = [c is not None for c in coords]

# ----------------------------------------------------------------------
# 5) Matrix in Excel (Blatt 1)
# ----------------------------------------------------------------------
out = openpyxl.Workbook()
sh = out.active
sh.title = "Abstandsmatrix_km"
sh.cell(row=1, column=1, value="V.Nr.")
for j, lab in enumerate(labels):
    sh.cell(row=1, column=2 + j, value=lab)
for i, lab in enumerate(labels):
    sh.cell(row=2 + i, column=1, value=lab)
    for j in range(N):
        sh.cell(row=2 + i, column=2 + j, value=matrix[i][j])

# ----------------------------------------------------------------------
# 6) In N_GRUPPEN ausgewogene Gruppen aufteilen (minimale interne Distanz)
# ----------------------------------------------------------------------
random.seed(SEED)
gueltig = [i for i in range(N) if valid[i]]
M = len(gueltig)

def dist(i, j):
    a, b = matrix[i][j], matrix[j][i]
    if a is None and b is None: return 0.0
    if a is None: return b
    if b is None: return a
    return (a + b) / 2

basis, rest = M // N_GRUPPEN, M % N_GRUPPEN
groessen = [basis + (1 if g < rest else 0) for g in range(N_GRUPPEN)]

def gruppen_kosten(m):
    s = 0.0
    for x in range(len(m)):
        for y in range(x + 1, len(m)):
            s += dist(m[x], m[y])
    return s

def gesamt_kosten(gr): return sum(gruppen_kosten(g) for g in gr)

def startloesung():
    pool = gueltig[:]
    random.shuffle(pool)
    gr, pos = [], 0
    for g in range(N_GRUPPEN):
        gr.append(pool[pos:pos + groessen[g]])
        pos += groessen[g]
    return gr

def optimiere(gr):
    verbessert = True
    while verbessert:
        verbessert = False
        for ga in range(N_GRUPPEN):
            for gb in range(ga + 1, N_GRUPPEN):
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

beste, beste_kosten = None, float("inf")
alle_kosten = []
for _ in range(40):
    g = optimiere(startloesung())
    k = gesamt_kosten(g)
    alle_kosten.append(k)
    if k < beste_kosten:
        beste, beste_kosten = [grp[:] for grp in g], k

mittlere_kosten = sum(alle_kosten) / len(alle_kosten)
groesste_kosten = max(alle_kosten)

# Referenz: rein zufällige Gruppierung OHNE lokale Optimierung
zufalls_kosten = [gesamt_kosten(startloesung()) for _ in range(1000)]
mittel_zufall = sum(zufalls_kosten) / len(zufalls_kosten)

print(f"\n{'='*50}")
print(f"{N_GRUPPEN} Gruppen, Gesamt-Distanzsumme innerhalb: {beste_kosten:.1f} km (Schnitt/Verein {beste_kosten / M:.1f} km)")
print(f"  (Referenz aus {len(alle_kosten)} optimierten Versuchen: Mittelwert {mittlere_kosten:.1f} km [{mittlere_kosten / M:.1f} km/Verein], "
      f"Maximum {groesste_kosten:.1f} km [{groesste_kosten / M:.1f} km/Verein])")
print(f"  (Reiner Zufall ohne Optimierung, Mittelwert aus {len(zufalls_kosten)} Versuchen: "
      f"{mittel_zufall:.1f} km [{mittel_zufall / M:.1f} km/Verein])\n")
for gi, grp in enumerate(beste, 1):
    n = len(grp)
    summe = gruppen_kosten(grp)
    paare = n * (n - 1) / 2
    avg_paar = summe / paare if paare else 0.0
    avg_verein = summe / n if n else 0.0
    print(f"Gruppe {gi} ({n} Vereine, Distanzsumme {summe:.1f} km, Schnitt/Paar {avg_paar:.1f} km, Schnitt/Verein {avg_verein:.1f} km):")
    print("   " + ", ".join(labels[i] for i in grp))

ohne = [labels[i] for i in range(N) if not valid[i]]
if ohne:
    print(f"\nNicht zugeordnet (keine Koordinate): {', '.join(ohne)}")

# Gruppen in Blatt 2
sh2 = out.create_sheet("Gruppen")
sh2.cell(row=1, column=1, value="Gruppe")
sh2.cell(row=1, column=2, value="V.Nr.")
z = 2
for gi, grp in enumerate(beste, 1):
    for i in grp:
        sh2.cell(row=z, column=1, value=gi)
        sh2.cell(row=z, column=2, value=labels[i])
        z += 1

# Abstandssummen je Gruppe (Übersicht rechts daneben)
sh2.cell(row=1, column=4, value="Gruppe")
sh2.cell(row=1, column=5, value="Anzahl Vereine")
sh2.cell(row=1, column=6, value="Distanzsumme (km)")
sh2.cell(row=1, column=7, value="Ø je Vereinspaar (km)")
sh2.cell(row=1, column=8, value="Ø je Verein (km)")
for gi, grp in enumerate(beste, 1):
    n = len(grp)
    summe = gruppen_kosten(grp)
    paare = n * (n - 1) / 2
    sh2.cell(row=1 + gi, column=4, value=gi)
    sh2.cell(row=1 + gi, column=5, value=n)
    sh2.cell(row=1 + gi, column=6, value=round(summe, 1))
    sh2.cell(row=1 + gi, column=7, value=round(summe / paare, 1) if paare else 0)
    sh2.cell(row=1 + gi, column=8, value=round(summe / n, 1) if n else 0)

for versuch in range(5):
    try:
        out.save(EXCEL_OUT)
        break
    except PermissionError:
        if versuch == 4:
            raise
        time.sleep(0.5)
print(f"\nGespeichert: {EXCEL_OUT} (Matrix + Gruppen)")
if nicht_gefunden:
    print("\nNicht geocodiert:")
    for vnr, adr in nicht_gefunden:
        print(f"  {vnr}: {adr}")