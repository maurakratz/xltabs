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

test_that("xl_crosstab handles pre-aggregated data (counts_col)", {
  df_agg <- data.frame(
    a = c("A", "B"),
    b = c("Y", "N"),
    n_count = c(10, 20)
  )

  res <- xl_crosstab(df_agg, a, b, counts_col = n_count)

  # --- ROBUSTER CHECK ---
  # 1. Wir filtern die Zeile, in der die erste Variable "Total" ist
  # (Wir nutzen [[1]], um die erste Spalte sicher zu greifen)
  total_row <- res[res[[1]] == "Total", ]

  # 2. Wir greifen gezielt auf die SPALTE namens "Total" zu
  # (Statt ncol(res))
  total_val_str <- total_row$Total

  # Prüfung
  expect_true(grepl("30", total_val_str))
})
