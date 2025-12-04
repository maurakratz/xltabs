test_that("xl_freq works and handles NAs", {
  df <- data.frame(x = c("A", "A", NA))

  # 1. Ohne NA Anzeige (Default)
  res1 <- xl_freq(df, x, show_na = FALSE)
  # Sollte nur A und Total enthalten (2 Zeilen)
  expect_equal(nrow(res1), 2)

  # 2. Mit NA Anzeige
  res2 <- xl_freq(df, x, show_na = TRUE)
  # Sollte A, Missing und Total enthalten (3 Zeilen)
  expect_equal(nrow(res2), 3)

  # 3. Prüfen ob Missing Label da ist
  expect_true("Missing" %in% res2$x)
})


test_that("xl_freq handles pre-aggregated data (counts_col)", {
  df_agg <- data.frame(
    x = c("High", "Low"),
    my_n = c(50, 50)
  )

  # Test
  res <- xl_freq(df_agg, x, counts_col = my_n)

  # Prüfung: Total Zeile sollte 100 zeigen
  # Wir suchen die Zeile wo x == "Total"
  total_row <- res[res$x == "Total", ]

  expect_true(grepl("100", total_row$cell_content))
})
