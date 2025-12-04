# Paket laden (simuliert library(xltabs))
devtools::load_all()
library(dplyr)

# 1. Testdaten erstellen (jedes Mal frisch)
test_df <- tibble(
  edu = sample(c("Low", "High"), 100, replace = TRUE),
  job = sample(c("Yes", "No"), 100, replace = TRUE),
  region = sample(c("North", "South"), 100, replace = TRUE),
  weight = runif(100, 0.5, 1.5)
)

# Label hinzufügen (für den neuen Test)
attr(test_df$edu, "label") <- "Highest Education"

# 2. Funktionen testen
# Einfach Zeile markieren und Strg+Enter drücken
xl_crosstab(test_df, edu, job)

# 3. Export testen
xl_export(xl_crosstab(test_df, edu, job), "test.xlsx")
