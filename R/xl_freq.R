#' Create Stata-style Frequency Tables for Excel Export
#'
#' Generates a data frame representing a frequency table with counts and percentages,
#' formatted as a single string per cell for easy export.
#' Includes support for weights, missing values, and a total row.
#' Automatically detects variable labels (e.g. from haven).
#'
#' @param data A data frame.
#' @param var The variable to analyze (unquoted).
#' @param w_var Optional. Variable for weights (unquoted). Default is NULL.
#' @param label Optional String. Manually set the display name for the variable.
#' @param show_n Logical. Show counts? Default TRUE.
#' @param show_pct Logical. Show percentages? Default TRUE.
#' @param show_na Logical. If TRUE, missing values are shown as explicit category. Default FALSE.
#' @param na_label String. Label for missing values. Default "Missing".
#' @param add_total Logical. Add a "Total" row at the bottom? Default TRUE.
#'
#' @return A data frame with two columns: The category and the formatted stats.
#'
#' @importFrom dplyr mutate filter count bind_rows select pull if_else arrange group_by ungroup summarise rename
#' @importFrom forcats as_factor fct_na_value_to_level
#' @importFrom scales percent
#' @import rlang
#' @export
xl_freq <- function(data,
                    var,
                    w_var = NULL,
                    # NEU: Label Argument
                    label = NULL,
                    show_n = TRUE,
                    show_pct = TRUE,
                    show_na = FALSE,
                    na_label = "Missing",
                    add_total = TRUE) {
  # --- 1. Input Checks ---
  check_df(data)
  check_bool(show_n)
  check_bool(show_pct)
  check_bool(show_na)
  check_string(na_label)
  check_bool(add_total)
  if (!is.null(label)) check_string(label)

  # --- Variablen einfangen ---
  v_sym <- enquo(var)
  w_sym <- enquo(w_var)

  # --- LABEL LOGIK ---
  final_name <- if (!is.null(label)) {
    label
  } else {
    # Versuche Attribut zu finden
    col_name <- rlang::as_name(v_sym)
    lbl <- attr(data[[col_name]], "label")
    if (!is.null(lbl)) lbl else col_name
  }

  # 2. Prepare Data & Handle NA
  df_prep <- data %>%
    mutate(!!v_sym := as_factor(!!v_sym))

  if (show_na) {
    df_clean <- df_prep %>%
      mutate(!!v_sym := fct_na_value_to_level(!!v_sym, level = na_label))
  } else {
    df_clean <- df_prep %>%
      filter(!is.na(!!v_sym))
  }

  # Convert to character for Total insertion
  df_clean <- df_clean %>%
    mutate(!!v_sym := as.character(!!v_sym))

  # Helper for counts
  calc_counts <- function(d, groups) {
    if (quo_is_null(w_sym)) {
      d %>% count(!!!groups, name = "n")
    } else {
      d %>% count(!!!groups, wt = !!w_sym, name = "n")
    }
  }

  # --- A. Counts ---
  # Categories
  df_counts <- calc_counts(df_clean, quos(!!v_sym))

  # Total Row
  if (add_total) {
    df_total <- calc_counts(df_clean, quos()) %>%
      mutate(!!v_sym := "Total")

    df_counts <- bind_rows(df_counts, df_total)
  }

  # --- B. Percentages ---
  total_n_val <- if (add_total) {
    df_counts %>% filter(!!v_sym == "Total") %>% pull(n)
  } else {
    sum(df_counts$n)
  }

  df_calc <- df_counts %>%
    mutate(
      pct = n / total_n_val,
      is_total = (!!v_sym == "Total")
    )

  # --- C. Formatting ---
  df_final <- df_calc %>%
    mutate(
      cell_content = paste0(
        if (show_n) paste0(round(n, 0)) else "",
        if (show_n && show_pct) "\n" else "",
        if (show_pct) paste0(scales::percent(pct, 0.1)) else ""
      )
    )

  # --- D. Sorting ---
  vals <- df_final %>% pull(!!v_sym)
  specials <- c()

  if (show_na) specials <- c(specials, na_label)
  if (add_total) specials <- c(specials, "Total")

  normals <- setdiff(vals, specials)
  normals <- sort(normals)

  final_levels <- c(normals, specials)

  df_sorted <- df_final %>%
    mutate(!!v_sym := factor(!!v_sym, levels = final_levels)) %>%
    arrange(!!v_sym) %>%
    select(!!v_sym, cell_content)

  # --- E. Final Renaming ---
  # Hier benennen wir die Variable um (z.B. "edu" -> "Highest Education")
  df_sorted <- df_sorted %>%
    rename(!!final_name := !!v_sym)

  return(df_sorted)
}
