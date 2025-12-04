test_that("xl_crosstab works with basic input", {
  # Wir erstellen kleine Dummy-Daten direkt im Test
  df <- data.frame(
    a = c("A", "A", "B"),
    b = c("Y", "N", "Y")
  )

  # Ausführen
  res <- xl_crosstab(df, a, b)

  # PRÜFUNGEN (Assertions)
  # 1. Ist es ein Dataframe?
  expect_s3_class(res, "data.frame")

  # 2. Haben wir die erwarteten Spalten? (Total spalte sollte da sein)
  expect_true("Total" %in% names(res))

  # 3. Stimmt die Zeilenanzahl? (A, B + Total = 3 Zeilen)
  expect_equal(nrow(res), 3)
})

test_that("xl_crosstab handles manual labels", {
  df <- data.frame(a = 1:3, b = 1:3)

  # Wir testen das Umbenennen
  res <- xl_crosstab(df, a, b, row_label = "MeineZeile")

  # Die erste Spalte muss jetzt "MeineZeile" heißen
  expect_equal(names(res)[1], "MeineZeile")
})

test_that("xl_crosstab validation works", {
  df <- data.frame(a = 1, b = 1)

  # Wir erwarten einen Fehler, wenn wir Quatsch eingeben
  expect_error(xl_crosstab(df, a, b, show_n = "Falsch"))
})
