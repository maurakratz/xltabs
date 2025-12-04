# --- dev_playground.R ---
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



# Test A: Automatische Erkennung
# (edu hat immer noch das Attribut "Highest Education" aus deinem vorherigen Test)
xl_freq(test_df, edu)

# Test B: Manuell
xl_freq(test_df, region, label = "Region (Manuell)")
