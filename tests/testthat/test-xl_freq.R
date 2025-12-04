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
