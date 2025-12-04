#' Export Data Frames to Excel
#'
#' Writes the data frame (cross-tabulation or frequency table) to an Excel file
#' with text wrapping and top alignment enabled.
#' Allows adding a title above and a footer (source/note) below the table.
#'
#' @param data The data frame produced by \code{xl_crosstab} or \code{xl_freq}.
#' @param filename String. Path to the output file (.xlsx).
#' @param sheet_name String. Name of the Excel sheet. Default "Table".
#' @param title Optional String. Text to appear in cell A1 above the table (bold).
#' @param footer Optional String. Text to appear below the table (italic, gray).
#'
#' @importFrom openxlsx createWorkbook addWorksheet writeData createStyle addStyle setColWidths saveWorkbook
#' @examples
#' \dontrun{
#' # 1. Prepare data
#' df <- data.frame(x = c("A", "B"), y = c("Y", "N"))
#' my_table <- xl_crosstab(df, x, y)
#'
#' # 2. Export to a file (e.g. "report.xlsx")
#' xl_export(my_table, "report.xlsx", title = "My Table")
#'
#' # Cleanup (delete file for this example)
#' # file.remove("report.xlsx")
#' }
#' @export
xl_export <- function(data, filename, sheet_name = "Table", title = NULL, footer = NULL) {
  # --- 1. Input Checks ---
  check_df(data)
  check_string(filename)
  check_string(sheet_name)
  if (!is.null(title)) check_string(title)
  if (!is.null(footer)) check_string(footer)

  # --- 2. Workbook Setup ---
  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)

  # --- 3. Title Logic ---
  # Wenn Titel da ist: Titel in Zeile 1, Tabelle startet in Zeile 3.
  # Wenn kein Titel: Tabelle startet in Zeile 1.
  table_start_row <- if (!is.null(title)) 3 else 1

  if (!is.null(title)) {
    # Titel schreiben
    writeData(wb, sheet_name, title, startCol = 1, startRow = 1)

    # Style: Fett und etwas größer
    title_style <- createStyle(textDecoration = "bold", fontSize = 12)
    addStyle(wb, sheet_name, title_style, rows = 1, cols = 1)
  }

  # --- 4. Write Data ---
  writeData(wb, sheet_name, data, startRow = table_start_row)

  # Style: Top alignment + Wrap Text (Wichtig für gestapelte Inhalte)
  # Wir wenden das auf den gesamten Tabellenbereich an
  table_rows <- table_start_row:(table_start_row + nrow(data))

  my_style <- createStyle(
    valign = "top",
    wrapText = TRUE
  )

  addStyle(wb, sheet_name, style = my_style,
    rows = table_rows,
    cols = 1:ncol(data),
    gridExpand = TRUE)

  # --- 5. Footer Logic ---
  if (!is.null(footer)) {
    # Footer kommt 2 Zeilen unter das Ende der Tabelle
    footer_row <- max(table_rows) + 2

    writeData(wb, sheet_name, footer, startCol = 1, startRow = footer_row)

    # Style: Kursiv und Grau
    footer_style <- createStyle(fontColour = "#555555", textDecoration = "italic")
    addStyle(wb, sheet_name, footer_style, rows = footer_row, cols = 1)
  }

  # --- 6. Final Polish ---
  # Spaltenbreite setzen
  setColWidths(wb, sheet_name, cols = 1:ncol(data), widths = 15)

  saveWorkbook(wb, filename, overwrite = TRUE)
}
