# ImportWorkflow

Ein RStudio-Addin, das Daten aus der Zwischenablage importiert, aufbereitet und als formatierte Tabelle für wissenschaftliche Publikationen speichert.

## Funktionen

Dieses Addin führt folgende Schritte automatisch aus:

1. Liest Daten aus der Zwischenablage ein
2. Bereinigt Spaltennamen mit dem `janitor`-Paket
3. Erstellt eine Übersicht mit dem `skimr`-Paket
4. Erstellt eine formatierte Tabelle mit dem `flextable`-Paket
5. Speichert die Tabelle als Word-Datei (`Tabelle.docx`)
6. Gibt eine Erfolgsmeldung aus

**Alle benötigten Pakete werden beim ersten Programmstart automatisch installiert, falls diese noch nicht verfügbar sind.**

## Verbesserte Datenerfassung

Das Addin erkennt automatisch verschiedene Datenformate in der Zwischenablage:

- Tabulator-getrennte Werte (TSV)
- Komma-getrennte Werte (CSV)
- Semikolon-getrennte Werte (häufig bei deutschen Excel-Versionen)
- Leerzeichen als Trennzeichen
- Spalten mit fester Breite

Diese intelligente Erkennung sorgt dafür, dass Daten aus verschiedenen Quellen korrekt eingelesen werden.

## Installation

Sie können das Paket direkt von GitHub installieren:

```r
# Installieren Sie zuerst das devtools-Paket, falls noch nicht vorhanden
if (!requireNamespace("devtools", quietly = TRUE)) {
  install.packages("devtools")
}

# Installieren Sie das ImportWorkflow-Paket
devtools::install_github("coderplus007/Import_workflow")
```

## Verwendung

Nach der Installation können Sie das Addin auf folgende Weise verwenden:

1. **Über das RStudio-Addin-Menü:**
   - Klicken Sie auf `Addins` in der oberen Menüleiste von RStudio
   - Wählen Sie `Daten importieren und aufbereiten`

2. **Mit Daten aus der Zwischenablage:**
   - Kopieren Sie Ihre Daten (z.B. aus Excel, Word, CSV-Dateien etc.)
   - Klicken Sie auf `Addins` → `Daten importieren und aufbereiten`
   - Die Daten werden automatisch aus der Zwischenablage eingelesen und das Format erkannt

## Button in RStudio einrichten

Sie können einen eigenen Button in der RStudio-Oberfläche einrichten, um das Addin schneller zu starten:

1. Gehen Sie in RStudio zu `Tools` → `Modify Keyboard Shortcuts...`
2. Suchen Sie nach "Daten importieren und aufbereiten" im Suchfeld
3. Klicken Sie in die Spalte "Shortcut" und weisen Sie eine Tastenkombination zu (z.B. Alt+I)
4. Alternativ können Sie auch einen Button in der Toolbar hinzufügen:
   - Gehen Sie zu `Tools` → `Customize Toolbars...`
   - Klicken Sie auf das Plus-Symbol (+) am unteren Rand der Toolbar
   - Wählen Sie die Kategorie "Addins"
   - Ziehen Sie "Daten importieren und aufbereiten" in Ihre Toolbar
   - Der Button ist nun immer sichtbar und mit einem Klick erreichbar

![Beispiel eines RStudio-Buttons](https://i.imgur.com/uDLPPJT.png)

## Abhängigkeiten

Das Paket verwendet folgende R-Pakete, die bei Bedarf automatisch installiert werden:

- janitor: Für die Bereinigung von Spaltennamen
- skimr: Für die Zusammenfassung der Daten
- flextable: Für die Erstellung formatierter Tabellen
- officer: Für den Export nach Word
- clipr: Für den Zugriff auf die Zwischenablage
- rstudioapi: Für den Zugriff auf RStudio-Funktionen
- readr: Für verbesserte Datenimporte
- utils: Für Hilfsfunktionen

## Kontakt

Bei Fragen oder Problemen wenden Sie sich bitte an: dev@olgameier.ch

## Lizenz

GPL-3