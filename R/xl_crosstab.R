#' Create Stata-style Cross-Tabulations for Export to Excel
#'
#' Generates a data frame representing a cross-tabulation, with cell content stacked into a single string for easy export to Excel workbooks.
#'
#' @details This function mimics the visual layout of Stata tables. It handles 2-way and 3-way tables, supports weights and missing values, and allows fully customisable output. You can toggle row/column totals, as well as the visibility of counts, row percentages, column percentages, and total percentages.
#'
#' @param data A data frame.
#' @param row_var The variable to use for rows (unquoted).
#' @param col_var The variable to use for columns (unquoted).
#' @param strat_var Variable to stratify by in a 3-way table (unquoted). Default is NULL.
#' @param w_var Variable for weights (unquoted). Default is NULL.
#' @param add_total_row Logical. Add a "Total" row at the bottom? Default is TRUE.
#' @param add_total_col Logical. Add a "Total" column at the right? Default is TRUE.
#' @param show_n Logical. Show counts? Default is TRUE.
#' @param show_row_pct Logical. Show row percentages? Default is TRUE.
#' @param show_col_pct Logical. Show column percentages? Default TRUE.
#' @param show_tot_pct Logical. Show total percentages? Default is TRUE.
#' @param show_na Logical. Show missing values in row/col variables as explicit category? Default is FALSE.
#' @param na_label String. Label for missing values. Default is "Missing".
#'
#' @return A data frame in wide format, ready for Excel export.
#' @importFrom dplyr mutate filter group_by ungroup count bind_rows arrange select pull if_else
#' @importFrom tidyr pivot_wider
#' @importFrom forcats as_factor fct_relevel fct_na_value_to_level
#' @importFrom scales percent
#' @importFrom openxlsx createWorkbook addWorksheet writeData createStyle addStyle setColWidths saveWorkbook
#' @import rlang
#' @export
xl_crosstab <- function(data,
                        row_var,
                        col_var,
                        strat_var = NULL,
                        w_var = NULL,
                        add_total_row = TRUE,
                        add_total_col = TRUE,
                        show_n = TRUE,
                        show_row_pct = TRUE,
                        show_col_pct = TRUE,
                        show_tot_pct = TRUE,
                        show_na = FALSE,
                        na_label = "Missing") {
  # --- 1. Input Checks ---
  check_df(data)
  check_bool(add_total_row)
  check_bool(add_total_col)
  check_bool(show_n)
  check_bool(show_row_pct)
  check_bool(show_col_pct)
  check_bool(show_tot_pct)
  check_bool(show_na)
  check_string(na_label)

  # --- Variablen einfangen ---
  r_sym <- enquo(row_var)
  c_sym <- enquo(col_var)
  s_sym <- enquo(strat_var)
  w_sym <- enquo(w_var)

  # 1. Prepare Data
  df_prep <- data %>%
    mutate(
      !!r_sym := as_factor(!!r_sym),
      !!c_sym := as_factor(!!c_sym)
    )

  if (!quo_is_null(s_sym)) {
    df_prep <- df_prep %>% mutate(!!s_sym := as_factor(!!s_sym))
  }

  # Handle NA
  if (show_na) {
    df_clean <- df_prep %>%
      mutate(
        !!r_sym := fct_na_value_to_level(!!r_sym, level = na_label),
        !!c_sym := fct_na_value_to_level(!!c_sym, level = na_label)
      )
    if (!quo_is_null(s_sym)) {
      df_clean <- df_clean %>%
        mutate(!!s_sym := fct_na_value_to_level(!!s_sym, level = na_label))
    }
  } else {
    df_clean <- df_prep %>%
      filter(!is.na(!!r_sym), !is.na(!!c_sym))
    if (!quo_is_null(s_sym)) {
      df_clean <- df_clean %>% filter(!is.na(!!s_sym))
    }
  }

  # Convert to character for "Total" insertion
  df_clean <- df_clean %>%
    mutate(
      !!r_sym := as.character(!!r_sym),
      !!c_sym := as.character(!!c_sym)
    )

  if (!quo_is_null(s_sym)) {
    df_clean <- df_clean %>% mutate(!!s_sym := as.character(!!s_sym))
  }

  # Helper for weighted counts
  calc_counts <- function(d, groups) {
    if (quo_is_null(w_sym)) {
      d %>% count(!!!groups, name = "n")
    } else {
      d %>% count(!!!groups, wt = !!w_sym, name = "n")
    }
  }

  # --- A. Core Counts ---
  groups_core <- quos(!!r_sym, !!c_sym)
  if (!quo_is_null(s_sym)) groups_core <- c(quos(!!s_sym), groups_core)
  df_core <- calc_counts(df_clean, groups_core)
  list_parts <- list(df_core)

  # --- B. Marginals (Totals) ---
  if (add_total_col) {
    groups_row <- quos(!!r_sym)
    if (!quo_is_null(s_sym)) groups_row <- c(quos(!!s_sym), groups_row)
    df_col_totals <- calc_counts(df_clean, groups_row) %>% mutate(!!c_sym := "Total")
    list_parts <- append(list_parts, list(df_col_totals))
  }

  if (add_total_row) {
    groups_col <- quos(!!c_sym)
    if (!quo_is_null(s_sym)) groups_col <- c(quos(!!s_sym), groups_col)
    df_row_totals <- calc_counts(df_clean, groups_col) %>% mutate(!!r_sym := "Total")
    list_parts <- append(list_parts, list(df_row_totals))
  }

  if (add_total_row && add_total_col) {
    groups_strat <- quos()
    if (!quo_is_null(s_sym)) groups_strat <- quos(!!s_sym)
    df_grand <- calc_counts(df_clean, groups_strat) %>% mutate(!!r_sym := "Total", !!c_sym := "Total")
    list_parts <- append(list_parts, list(df_grand))
  }

  df_all <- bind_rows(list_parts)

  # --- C. Calculate Percentages ---
  df_calc <- df_all %>%
    group_by(!!!(if (!quo_is_null(s_sym)) s_sym else NULL)) %>%
    mutate(stratum_n = sum(n[as_factor(!!r_sym) != "Total" & as_factor(!!c_sym) != "Total"])) %>%
    group_by(!!!(if (!quo_is_null(s_sym)) c(s_sym, r_sym) else r_sym)) %>%
    mutate(row_denom = sum(n[as_factor(!!c_sym) != "Total"])) %>%
    group_by(!!!(if (!quo_is_null(s_sym)) c(s_sym, c_sym) else c_sym)) %>%
    mutate(col_denom = sum(n[as_factor(!!r_sym) != "Total"])) %>%
    ungroup() %>%
    mutate(
      pct_row = n / row_denom,
      pct_col = n / col_denom,
      pct_tot = n / stratum_n
    )

  # --- D. Formatting (String Build) ---
  df_final <- df_calc %>%
    mutate(
      is_t_row = as.character(!!r_sym) == "Total",
      is_t_col = as.character(!!c_sym) == "Total",

      use_row = show_row_pct & !is_t_col,
      use_col = show_col_pct & !is_t_row,
      use_tot = show_tot_pct & !is_t_row & !is_t_col,

      cell_content = paste0(
        if (show_n) paste0(round(n, 0)) else "",
        if_else(show_n & (use_row | use_col | use_tot), "\n", ""),

        if_else(use_row, paste0(scales::percent(pct_row, 0.1)), ""),
        if_else(use_row & (use_col | use_tot), "\n", ""),

        if_else(use_col, paste0(scales::percent(pct_col, 0.1)), ""),
        if_else(use_col & use_tot, "\n", ""),

        if_else(use_tot, paste0(scales::percent(pct_tot, 0.1)), "")
      )
    ) %>%
    mutate(cell_content = gsub("\n$", "", cell_content))

  # --- E. Sorting & Pivot (Manual Set Logic) ---
  col_vals <- df_final %>% pull(!!c_sym) %>% unique() %>% as.character()

  has_total <- "Total" %in% col_vals
  has_na    <- na_label %in% col_vals

  normal_vals <- setdiff(col_vals, c("Total", na_label))
  normal_vals <- sort(normal_vals)

  final_levels <- normal_vals
  if (has_na) final_levels <- c(final_levels, na_label)
  if (has_total) final_levels <- c(final_levels, "Total")

  df_sorted <- df_final %>%
    mutate(
      !!c_sym := factor(!!c_sym, levels = final_levels),
      !!r_sym := fct_relevel(as_factor(!!r_sym), "Total", after = Inf)
    )

  if (!quo_is_null(s_sym)) {
    df_sorted <- df_sorted %>% arrange(!!s_sym, !!r_sym, !!c_sym)
  } else {
    df_sorted <- df_sorted %>% arrange(!!r_sym, !!c_sym)
  }

  df_pivoted <- df_sorted %>%
    select(!!s_sym, !!r_sym, !!c_sym, cell_content) %>%
    pivot_wider(
      names_from = !!c_sym,
      values_from = cell_content,
      values_fill = "-"
    )

  return(df_pivoted)
}
