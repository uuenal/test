# Windchill BOM-Download-Tool

Automatisiert den bisher manuellen Ablauf: Materialnummer in Windchill
suchen, Multilevel Report der Baugruppenstruktur oeffnen und als Excel
(`.xlsx`) speichern.

## Nutzung

1. `run_bom_download.bat` doppelklicken.
2. Materialnummer eingeben (z. B. `3XAA114646A0001`).
3. Ein Edge-Fenster oeffnet sich automatisch, meldet sich per Windows-SSO
   an Windchill an und laeuft den Ablauf durch.
4. Ergebnis liegt danach unter:
   `C:\Users\DEUFUENA\OneDrive - ABB\Dokumente\Aufgaben Ufuk\Bestandsmanagementtool\rohdaten\BOM auto\<Materialnummer>.xlsx`

Existiert fuer die Materialnummer bereits eine Datei, wird sie automatisch
ueberschrieben (keine Duplikate pro Materialnummer).

## Voraussetzungen

- Python ist installiert (bereits vorhanden).
- Microsoft Edge ist installiert (Standard unter Windows).
- Verbindung zum ABB-Netz (im Buero) bzw. aktives VPN (von zuhause) - das
  Tool prueft die Erreichbarkeit von Windchill zu Beginn und bricht mit
  einer klaren Meldung ab, falls das VPN fehlt.
- Beim ersten Start installiert das `.bat`-Skript automatisch das Python-
  Paket `playwright` (`pip install -r requirements.txt`).

## Bekannte Einschraenkungen

Alle Klick-Schritte (Suchfeld, Suchtreffer, Structure, Reports,
Multilevel Report, Actions, Export List to File/XLSX) wurden inzwischen
per Browser-DevTools am echten Seitenquelltext verifiziert. Falls
Windchill dennoch einmal aktualisiert wird und sich ein Selektor aendert,
fragt das Suchfeld-Handling als letzte Sicherheitsstufe nach einem
einmaligen manuellen Klick, statt einfach abzubrechen.

Bei einem Fehler wird zusaetzlich ein Screenshot des aktuellen
Browser-Zustands unter `letzter_fehler.png` gespeichert - das hilft, den
betroffenen Schritt schnell zu identifizieren und die Selektoren im
Skript gezielt nachzujustieren.

Der erste Lauf sollte beobachtet werden, um eventuelle kleine
Anpassungen direkt vorzunehmen (gleiches Prinzip wie beim bestehenden
SAP-Automatisierungstool).

## Materialnummer direkt als Argument uebergeben

Statt der interaktiven Eingabe kann die Materialnummer auch direkt
mitgegeben werden, z. B. fuer eine eigene Verknuepfung:

```
run_bom_download.bat 3XAA114646A0001
```
