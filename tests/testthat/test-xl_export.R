test_that("xl_export creates a file", {
  df <- data.frame(col1 = c("A"), col2 = c("10%"))

  # Wir nutzen eine temporäre Datei (wird von R automatisch gelöscht nach der Session)
  tmp_file <- tempfile(fileext = ".xlsx")

  # Funktion ausführen
  xl_export(df, tmp_file, title = "Test")

  # PRÜFUNG: Existiert die Datei physikalisch?
  expect_true(file.exists(tmp_file))
})

test_that("xl_export validates input", {
  df <- data.frame(a = 1)
  # Fehler bei falschem Dateinamen
  expect_error(xl_export(df, filename = 123))
})
