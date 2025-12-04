# --- dev_playground.R ---


# xl_crosstab --------------------------------------------------
devtools::load_all()
library(dplyr)

# 1. Testdaten erstellen
test_df <- tibble(
  edu = sample(c("Low", "High"), 100, replace = TRUE),
  job = sample(c("Yes", "No"), 100, replace = TRUE),
  region = sample(c("North", "South"), 100, replace = TRUE),
  weight = runif(100, 0.5, 1.5)
)

# 2. Labels hinzufügen (simuliert "haven"-Import)
attr(test_df$edu, "label")    <- "Höchster Bildungsabschluss"
attr(test_df$region, "label") <- "Geographische Region"
# Variable 'job' hat KEIN Label

# ---------------------------------------------------------
# Test A: Automatische Erkennung
# ---------------------------------------------------------
print("--- TEST A: Automatische Labels ---")
res_auto <- xl_crosstab(test_df, edu, job)
print(res_auto)
# ERWARTUNG: Erste Spalte heißt "Höchster Bildungsabschluss"

# ---------------------------------------------------------
# Test B: Manueller Override
# ---------------------------------------------------------
print("--- TEST B: Manuelles Label (überschreibt auto) ---")
res_manual <- xl_crosstab(test_df, edu, job, row_label = "Bildung (Manuell)")
print(res_manual)
# ERWARTUNG: Erste Spalte heißt "Bildung (Manuell)"

# ---------------------------------------------------------
# Test C: 3-Way mit Labels
# ---------------------------------------------------------
print("--- TEST C: 3-Way Table ---")
res_3way <- xl_crosstab(test_df, edu, job, region)
print(res_3way)
# ERWARTUNG: Spalten heißen "Geographische Region" und "Höchster Bildungsabschluss"


# xl_freq --------------------------------------------------------

# Test A: Automatische Erkennung
# (edu hat immer noch das Attribut "Highest Education" aus deinem vorherigen Test)
xl_freq(test_df, edu)

# Test B: Manuell
xl_freq(test_df, region, label = "Region (Manuell)")


# xl_export ------------------------------------------------------
# Testdaten vorbereiten (falls weg)
test_df <- tibble(
  edu = sample(c("Low", "High"), 100, replace = TRUE),
  job = sample(c("Yes", "No"), 100, replace = TRUE)
)
my_table <- xl_crosstab(test_df, edu, job)

# Der Export mit Titel und Branding
xl_export(
  my_table,
  filename = "report_test.xlsx",
  title = "Table 1: Education by Job Status",
  footer = "Source: Artificial Data 2025 | Created with xltabs package"
)



# pre aggregated data ---------------------------

# Wir bauen uns mal "fertige" Daten
agg_data <- tibble(
  edu = c("High", "Low"),
  job = c("Yes", "No"),
  anzahl = c(500, 200) # Das sind schon gezählte Werte!
)

# agg_data %>% View()

# Test 1: Cross-Tab mit fertigen Counts
# Erwartung: 500 und 200 sollten angezeigt werden (nicht 1 und 1)
xl_crosstab(agg_data, edu, job, counts_col = anzahl)

# Test 2: Frequency mit fertigen Counts
xl_freq(agg_data, edu, counts_col = anzahl)
# Erwartung: High = 500, Low = 200, Total = 700
