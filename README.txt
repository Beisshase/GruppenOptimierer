GruppenOptimierer
==================

Beschreibung
------------
Berechnet die Fahrtkilometer (PKW) zwischen den Spielfeldern von
Fussballvereinen und teilt die Vereine in moeglichst ausgewogene
Gruppen auf, deren interne Fahrtstrecken-Summe minimal ist.

Es gibt zwei Varianten mit derselben Kernlogik:
  - mycode.py: Kommandozeilen-Skript fuer eine feste Eingabedatei
  - app.py: interaktive Streamlit-Website (Excel hochladen, Gruppen
            per Drag & Drop anpassen, Kartenansicht, Download)

Installation
-------------
py -3.12 -m pip install -r requirements.txt

Input
-----
Excel-Datei mit Vereinsliste
  Spalte A: Vereinsnummer
  Spalte B: Vereinsname
  Spalte L: Adresse des Spielfelds - entweder eine echte Adresse
            ("Strasse Hausnr., PLZ Ort") oder, falls keine Strasse
            bekannt ist, ein Google Plus Code ("Pluscode, PLZ Ort")

Output
------
<Eingabedatei>-Gruppeneinteilung.xlsx
  Blatt "Abstandsmatrix_km": Fahrtkilometer-Matrix zwischen allen Vereinen
  Blatt "Gruppen": Gruppenzuordnung je Verein (mit Vereinsname), dazu
                   Uebersicht mit Distanzsumme, Schnitt je Vereinspaar
                   und Schnitt je Verein pro Gruppe

Geocoding-Log (nur Website)
----------------------------
Auf Wunsch herunterladbares Protokoll, das auflistet, welche Vereine
nicht praezise (ueber Strasse oder Pluscode) geocodiert werden konnten,
sondern nur ueber die Ortsmitte/PLZ - sowie Vereine ohne gefundene
Koordinate.

Skripte
-------
mycode.py
  Kommandozeilen-Hauptskript: liest die Eingabe-Excel ein, bereinigt die
  Adressen, geocodiert sie ueber Photon (inkl. Pluscode-Unterstuetzung),
  berechnet die Fahrtkilometer-Matrix ueber OSRM und teilt die Vereine
  per Local-Search-Heuristik (Zufallsneustarts) in N_GRUPPEN Gruppen auf.
  Existiert die Ausgabedatei bereits, wird die Matrix daraus geladen statt
  neu berechnet (kein erneutes Geocoding/OSRM) - so laesst sich z.B.
  N_GRUPPEN aendern, ohne alles neu zu rechnen.

app.py
  Interaktive Streamlit-Website mit derselben Kernlogik: Excel hochladen,
  Anzahl Gruppen waehlen, Berechnen. Danach lassen sich die Gruppen per
  Drag & Drop anpassen (mit Vergleich ggue. Optimum und einem selbst
  gespeicherten Vergleichspunkt), die Gruppen auf einer Karte ansehen und
  das Ergebnis als Excel herunterladen.

Konfiguration (oben in mycode.py)
----------------------------------
EXCEL_IN   Pfad zur Eingabedatei
N_GRUPPEN  Anzahl der Gruppen
SEED       Zufallssaat fuer reproduzierbare Ergebnisse

Ausfuehren
----------
Kommandozeile:
  py -3.12 mycode.py

Website (lokal):
  py -3.12 -m streamlit run app.py
