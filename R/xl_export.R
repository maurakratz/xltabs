#' Export Data Frames to Excel
#'
#' Writes the data frame (cross-tabulation or frequency table) to an Excel file
#' with text wrapping and top alignment enabled. This ensures that the
#' stacked cell content (Counts and Percentages) is displayed correctly.
#'
#' @param data The data frame produced by \code{xl_crosstab} or \code{xl_freq}.
#' @param filename String. Path to the output file (.xlsx).
#' @param sheet_name String. Name of the Excel sheet. Default "Table".
#'
#' @importFrom openxlsx createWorkbook addWorksheet writeData createStyle addStyle setColWidths saveWorkbook
#' @export
xl_export <- function(data, filename, sheet_name = "Table") {
  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)

  writeData(wb, sheet_name, data)

  # Style: Top alignment + Wrap Text (Crucial for stacked content)
  my_style <- createStyle(
    valign = "top",
    wrapText = TRUE
  )

  addStyle(wb, sheet_name, style = my_style,
    rows = 1:(nrow(data) + 1),
    cols = 1:ncol(data),
    gridExpand = TRUE)

  # Set reasonable column width
  setColWidths(wb, sheet_name, cols = 1:ncol(data), widths = 15)

  saveWorkbook(wb, filename, overwrite = TRUE)
}
