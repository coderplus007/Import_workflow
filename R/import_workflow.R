#' Daten importieren und aufbereiten
#'
#' Diese Funktion importiert und bereitet Daten auf. Sie führt folgende Schritte aus:
#' 1. Daten aus der Zwischenablage einlesen
#' 2. Spaltennamen bereinigen
#' 3. Übersicht mit skimr erstellen
#' 4. Formatierte Tabelle mit flextable erstellen
#' 5. Tabelle als Word-Datei speichern
#' 6. Erfolgsmeldung ausgeben
#'
#' Benötigte Pakete werden automatisch installiert, falls diese noch nicht verfügbar sind.
#'
#' @return Gibt unsichtbar die formatierte Tabelle zurück und speichert "Tabelle.docx"
#' @export
#'
#' @importFrom janitor clean_names
#' @importFrom skimr skim
#' @importFrom flextable flextable autofit theme_vanilla bold fontsize save_as_docx
#' @importFrom clipr read_clip
#' @importFrom rstudioapi getActiveDocumentContext
#'
import_workflow <- function() {
  # Prüfen und Installation der benötigten Pakete
  packages <- c("janitor", "skimr", "flextable", "officer", "clipr", "rstudioapi", "utils", "readr")
  
  for (pkg in packages) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      message(paste0("Installiere benötigtes Paket: ", pkg, "..."))
      install.packages(pkg, quiet = TRUE)
      if (!requireNamespace(pkg, quietly = TRUE)) {
        stop(paste0("Konnte '", pkg, "' nicht installieren. Bitte manuell installieren."))
      }
      message(paste0("Paket '", pkg, "' erfolgreich installiert."))
    }
  }
  
  # Laden der benötigten Pakete
  library(janitor)
  library(skimr)
  library(flextable)
  library(officer)
  library(clipr)
  library(rstudioapi)
  library(utils)
  library(readr)
  
  # Schritt 1: Daten aus der Zwischenablage einlesen
  message("Lese Daten aus der Zwischenablage...")
  data <- read_clipboard_smart()
  
  # Schritt 2: Spaltennamen bereinigen
  message("Bereinige Spaltennamen...")
  data <- janitor::clean_names(data)
  
  # Schritt 3: Übersicht mit skimr erstellen
  message("Erstelle Datenübersicht mit skimr...")
  skim_result <- skimr::skim(data)
  print(skim_result)
  
  # Schritt 4: Formatierte Tabelle mit flextable erstellen
  message("Erstelle formatierte Tabelle mit flextable...")
  ft <- flextable::flextable(data)
  ft <- flextable::autofit(ft)
  ft <- flextable::theme_vanilla(ft)
  ft <- flextable::bold(ft, part = "header")
  ft <- flextable::fontsize(ft, size = 10)
  
  # In der Konsole anzeigen
  print(ft)
  
  # Schritt 5: Tabelle als Word-Datei speichern
  message("Speichere Tabelle als 'Tabelle.docx'...")
  flextable::save_as_docx(ft, path = "Tabelle.docx")
  
  # Schritt 6: Erfolgsmeldung ausgeben
  message("Datenimport und -aufbereitung erfolgreich abgeschlossen! Die Tabelle wurde als 'Tabelle.docx' gespeichert.")
  
  # Unsichtbar die Tabelle zurückgeben
  invisible(ft)
}

#' Intelligente Funktion zum Einlesen von Daten aus der Zwischenablage
#'
#' Diese Hilfsfunktion versucht, Daten aus der Zwischenablage einzulesen und
#' erkennt dabei automatisch das Format und das Trennzeichen.
#'
#' @return Ein Datenrahmen mit den eingelesenen Daten
#' @keywords internal
read_clipboard_smart <- function() {
  # Daten aus der Zwischenablage lesen
  clip_content <- clipr::read_clip()
  
  # Wenn leer, Fehler ausgeben
  if (length(clip_content) == 0 || all(clip_content == "")) {
    stop("Die Zwischenablage ist leer. Bitte kopieren Sie zuerst Daten.")
  }
  
  # Versuchen, verschiedene Trennzeichen zu erkennen und anzuwenden
  
  # 1. Versuch: Als TSV (Tabulator-getrennt) einlesen
  tryCatch({
    data <- readr::read_tsv(paste(clip_content, collapse = "\n"), show_col_types = FALSE)
    # Prüfen, ob sinnvolle Struktur (mindestens 2 Spalten)
    if (ncol(data) >= 2) {
      message("Daten als Tabulator-getrennte Werte (TSV) erkannt.")
      return(as.data.frame(data))
    }
  }, error = function(e) {})
  
  # 2. Versuch: Als CSV (Komma-getrennt) einlesen
  tryCatch({
    data <- readr::read_csv(paste(clip_content, collapse = "\n"), show_col_types = FALSE)
    # Prüfen, ob sinnvolle Struktur (mindestens 2 Spalten)
    if (ncol(data) >= 2) {
      message("Daten als Komma-getrennte Werte (CSV) erkannt.")
      return(as.data.frame(data))
    }
  }, error = function(e) {})
  
  # 3. Versuch: Als CSV mit Semikolon einlesen (häufig bei deutschsprachigen Excel-Dateien)
  tryCatch({
    data <- readr::read_csv2(paste(clip_content, collapse = "\n"), show_col_types = FALSE)
    # Prüfen, ob sinnvolle Struktur (mindestens 2 Spalten)
    if (ncol(data) >= 2) {
      message("Daten als Semikolon-getrennte Werte erkannt.")
      return(as.data.frame(data))
    }
  }, error = function(e) {})
  
  # 4. Versuch: Spalten durch Leerzeichen erkennen (fixe Breite)
  tryCatch({
    # Finde die Positionen der Wörter in der ersten Zeile
    header_line <- clip_content[1]
    
    # Muster von Zeichen und Leerzeichen finden
    positions <- gregexpr("[^ ]+", header_line)[[1]]
    widths <- attr(positions, "match.length")
    
    # Erstelle eine Matrix mit den Daten
    data_matrix <- do.call(rbind, lapply(clip_content, function(line) {
      values <- character(length(positions))
      for (i in seq_along(positions)) {
        start <- positions[i]
        end <- start + widths[i] - 1
        values[i] <- substr(line, start, end)
      }
      return(values)
    }))
    
    # Erstelle einen Datenrahmen
    colnames <- data_matrix[1, ]
    data <- as.data.frame(data_matrix[-1, ], stringsAsFactors = FALSE)
    names(data) <- colnames
    
    message("Daten mit festen Spaltenbreiten erkannt.")
    return(data)
  }, error = function(e) {})
  
  # 5. Versuch: Regulären Ausdruck verwenden, um Zahlen und Texte zu trennen
  tryCatch({
    # Nimm an, dass die Einträge durch ein oder mehrere Leerzeichen getrennt sind
    raw_data <- strsplit(clip_content, "\\s+")
    
    # Prüfe, ob alle Zeilen die gleiche Anzahl von Elementen haben
    element_counts <- sapply(raw_data, length)
    if (length(unique(element_counts)) == 1 && unique(element_counts) > 1) {
      # Erstelle einen Datenrahmen
      data <- as.data.frame(do.call(rbind, raw_data), stringsAsFactors = FALSE)
      colnames(data) <- data[1, ]
      data <- data[-1, ]
      rownames(data) <- NULL
      
      message("Daten mit Leerzeichen als Trennzeichen erkannt.")
      return(data)
    }
  }, error = function(e) {})
  
  # Fallback: Wenn alle Versuche fehlschlagen, versuche einfaches Split bei Whitespace
  warning("Konnte das genaue Format nicht bestimmen. Versuche allgemeine Analyse...")
  
  # Versuche, die erste Zeile als Header zu verwenden und dann zu splitten
  header <- unlist(strsplit(clip_content[1], "\\s+"))
  data <- data.frame(matrix(ncol = length(header), nrow = length(clip_content) - 1))
  colnames(data) <- header
  
  for (i in 2:length(clip_content)) {
    values <- unlist(strsplit(clip_content[i], "\\s+"))
    if (length(values) == length(header)) {
      data[i-1, ] <- values
    } else {
      # Bei ungleicher Länge versuche eine andere Strategie, z.B. Auffüllen mit NA
      row_data <- values
      if (length(row_data) < length(header)) {
        row_data <- c(row_data, rep(NA, length(header) - length(row_data)))
      } else if (length(row_data) > length(header)) {
        row_data <- row_data[1:length(header)]
      }
      data[i-1, ] <- row_data
    }
  }
  
  return(data)
}

#' Import data and prepare (English Version)
#'
#' This function imports and prepares data. It performs the following steps:
#' 1. Read data from the clipboard
#' 2. Clean column names
#' 3. Create an overview with skimr
#' 4. Create a formatted table with flextable
#' 5. Save the table as a Word file
#' 6. Display a success message
#'
#' Required packages will be automatically installed if they are not already available.
#'
#' @return Invisibly returns the formatted table and saves "Table.docx"
#' @export
#'
#' @importFrom janitor clean_names
#' @importFrom skimr skim
#' @importFrom flextable flextable autofit theme_vanilla bold fontsize save_as_docx
#' @importFrom clipr read_clip
#' @importFrom rstudioapi getActiveDocumentContext
#'
import_workflow_en <- function() {
  # Check and install required packages
  packages <- c("janitor", "skimr", "flextable", "officer", "clipr", "rstudioapi", "utils", "readr")
  
  for (pkg in packages) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      message(paste0("Installing required package: ", pkg, "..."))
      install.packages(pkg, quiet = TRUE)
      if (!requireNamespace(pkg, quietly = TRUE)) {
        stop(paste0("Could not install '", pkg, "'. Please install manually."))
      }
      message(paste0("Package '", pkg, "' successfully installed."))
    }
  }
  
  # Load required packages
  library(janitor)
  library(skimr)
  library(flextable)
  library(officer)
  library(clipr)
  library(rstudioapi)
  library(utils)
  library(readr)
  
  # Step 1: Read data from clipboard
  message("Reading data from clipboard...")
  data <- read_clipboard_smart()
  
  # Step 2: Clean column names
  message("Cleaning column names...")
  data <- janitor::clean_names(data)
  
  # Step 3: Create overview with skimr
  message("Creating data overview with skimr...")
  skim_result <- skimr::skim(data)
  print(skim_result)
  
  # Step 4: Create formatted table with flextable
  message("Creating formatted table with flextable...")
  ft <- flextable::flextable(data)
  ft <- flextable::autofit(ft)
  ft <- flextable::theme_vanilla(ft)
  ft <- flextable::bold(ft, part = "header")
  ft <- flextable::fontsize(ft, size = 10)
  
  # Display in console
  print(ft)
  
  # Step 5: Save table as Word file
  message("Saving table as 'Table.docx'...")
  flextable::save_as_docx(ft, path = "Table.docx")
  
  # Step 6: Display success message
  message("Data import and preparation completed successfully! The table has been saved as 'Table.docx'.")
  
  # Return the table invisibly
  invisible(ft)
}