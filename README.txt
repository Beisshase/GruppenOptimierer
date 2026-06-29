GruppenOptimierer
==================

Beschreibung
------------
Berechnet die Fahrtkilometer (PKW) zwischen den Spielfeldern von
Fussballvereinen und teilt die Vereine in moeglichst ausgewogene
Gruppen auf, deren interne Fahrtstrecken-Summe minimal ist.

Input
-----
Meldelisten-mit-Sportstaetten_A1.xlsx (Worksheet "AZ")
  Spalte A: Vereinsnummer
  Spalte L: Adresse des Spielfelds

Output
------
Abstandsmatrix.xlsx
  Blatt "Abstandsmatrix_km": Fahrtkilometer-Matrix zwischen allen Vereinen
  Blatt "Gruppen": Gruppenzuordnung je Verein, dazu Uebersicht mit
                   Distanzsumme, Schnitt je Vereinspaar und Schnitt
                   je Verein pro Gruppe

Skripte
-------
mycode.py
  Hauptskript: liest die Eingabe-Excel ein, bereinigt die Adressen,
  geocodiert sie ueber Photon, berechnet die Fahrtkilometer-Matrix
  ueber OSRM und teilt die Vereine per Local-Search-Heuristik
  (Zufallsneustarts) in N_GRUPPEN Gruppen auf.
  Existiert Abstandsmatrix.xlsx bereits, wird die Matrix daraus
  geladen statt neu berechnet (kein erneutes Geocoding/OSRM) - so
  laesst sich z.B. N_GRUPPEN aendern, ohne alles neu zu rechnen.

exakt.py
  Vergleichsskript: exakte Optimierung der Gruppenbildung per MILP
  (PuLP/CBC) auf Basis der vorhandenen Abstandsmatrix.xlsx.

exakt_cpsat.py
  Vergleichsskript: exakte Optimierung per OR-Tools CP-SAT, inkl.
  Warmstart mit der Heuristik-Loesung als Hint.

Konfiguration (oben in mycode.py)
----------------------------------
N_GRUPPEN  Anzahl der Gruppen
SEED       Zufallssaat fuer reproduzierbare Ergebnisse

Ausfuehren
----------
py -3.12 mycode.py
