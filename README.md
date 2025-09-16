# Excel Consolidation Tool (Python)

Ein Python-Skript zur automatischen Konsolidierung von Daten aus mehreren Excel-Projektstatusdateien in eine zentrale Übersichtsdatei.

## Das Problem

In vielen Projekten werden Status-Updates in separaten Dateien für verschiedene Teilbereiche oder Zeiträume gepflegt. Diese Daten manuell zu einer Gesamtübersicht zusammenzuführen, ist zeitaufwendig, repetitiv und anfällig für Kopierfehler.

**Ausgangssituation:** Mehrere Excel-Dateien mit ähnlicher Struktur, aber unterschiedlichen Daten.

<img width="1908" height="435" alt="ScreenshotTeilprojekte" src="https://github.com/user-attachments/assets/4d8d136e-eb0f-4a9a-a32b-b529ba519874" />


## Die Lösung

Dieses Skript automatisiert den gesamten Prozess. Es durchsucht einen definierten Ordner nach allen relevanten Projekt-Excel-Dateien, extrahiert die benötigten Spalten und führt alle Daten in einer einzigen, sauberen und übersichtlichen Excel-Datei zusammen.

**Das Ergebnis:** Eine zentrale Statusdatei, die alle Teilprojekte in einer fortlaufenden Liste konsolidiert.

<img width="661" height="520" alt="Screenshot_Zusammenführung" src="https://github.com/user-attachments/assets/945676a1-6126-48bc-8de0-7bee27943ffb" />


## ⚙Funktionsweise im Detail

Das Skript führt die folgenden Schritte aus:

1.  **Definierter Zielordner:** Das Skript durchsucht einen vordefinierten Ordner nach allen Dateien.
2.  **Intelligente Dateisuche:** Es werden nur Excel-Dateien (`.xlsx`) berücksichtigt, die mit dem Präfix `projekt` beginnen, um sicherzustellen, dass nur relevante Statusdateien verarbeitet werden.
3.  **Gezielte Datenextraktion:** Aus jeder gefundenen Datei werden nur die vordefinierten Spalten (`Projekt-ID`, `Status`, `Bearbeiter`, `Letztes Update`) extrahiert.
4.  **
