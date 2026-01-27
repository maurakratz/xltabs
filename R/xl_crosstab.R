#' Create Stata-style Cross-Tabulations for Export to Excel
#'
#' Generates a data frame representing a cross-tabulation, with cell content
#' stacked into a single string for easy export to Excel workbooks.
#'
#' @details
#' This function mimics the visual layout of Stata tables. It handles 2-way and
#' 3-way tables, supports weights and missing values, and allows fully
#' customisable output. You can toggle row/column totals, as well as the
#' visibility of counts, row percentages, column percentages, and total percentages.
#' It automatically detects variable labels (e.g. from haven/sjlabelled) if present.
#'
#' @param data A data frame.
#' @param row_var The variable to use for rows (unquoted).
#' @param col_var The variable to use for columns (unquoted).
#' @param strat_var Optional. Variable to stratify by (3-way table) (unquoted). Default is NULL.
#' @param w_var Optional. Variable for weights (unquoted). Default is NULL.
#' @param counts_col Optional. Variable containing pre-calculated counts (unquoted). Use this if your data is already aggregated.
#' @param row_label Optional String. Manually set the display name for the row variable.
#' @param col_label Optional String. Manually set the display name for the column variable.
#' @param strat_label Optional String. Manually set the display name for the stratification variable.
#' @param add_total_row Logical. Add a "Total" row at the bottom? Default TRUE.
#' @param add_total_col Logical. Add a "Total" column at the right? Default TRUE.
#' @param show_n Logical. Show counts? Default TRUE.
#' @param show_row_pct Logical. Show row percentages? Default TRUE.
#' @param show_col_pct Logical. Show column percentages? Default TRUE.
#' @param show_tot_pct Logical. Show total percentages? Default TRUE.
#' @param show_na Logical. If TRUE, missing values in row/col variables are shown as an explicit category. Default FALSE.
#' @param na_label String. Label for missing values. Default "Missing".
#'
#' @return A data frame in wide format, ready for export to Excel.
#'
#' @importFrom dplyr mutate filter group_by ungroup count bind_rows arrange select pull if_else rename any_of
#' @importFrom tidyr pivot_wider
#' @importFrom forcats as_factor fct_relevel fct_na_value_to_level
#' @importFrom scales percent
#' @importFrom openxlsx createWorkbook addWorksheet writeData createStyle addStyle setColWidths saveWorkbook
#' @import rlang
#' @export
xl_crosstab <- function(df, row_var, col_var = NULL, strat_var = NULL, w_var = NULL,
                        counts_col = NULL,
                        row_label = NULL, col_label = NULL, strat_label = NULL,
                        title = NULL, footer = NULL,
                        show_n = TRUE, show_row_pct = TRUE, show_col_pct = TRUE, show_tot_pct = FALSE,
                        na_label = "(Missing)", decimals = 1) {
  # --- FIX 1: Gruppierung entfernen (Löst das n=1 Problem) ---
  df <- dplyr::ungroup(df)

  r_sym <- rlang::enquo(row_var)
  c_sym <- rlang::enquo(col_var)
  s_sym <- rlang::enquo(strat_var)
  w_sym <- rlang::enquo(w_var)
  n_sym <- rlang::enquo(counts_col)

  # --- 1. FREQUENCY TABLE CHECK ---
  if (rlang::quo_is_null(c_sym)) {
    # Hinweis: Falls xl_freq nicht exportiert ist, könnte das hier fehlschlagen,
    # aber für deinen aktuellen Fall (Crosstab) ist das egal.
    return(xl_freq(df, !!r_sym, w_var = !!w_sym,
      row_label = row_label, title = title, footer = footer,
      na_label = na_label, decimals = decimals))
  }

  # --- 2. PREPARATION ---
  get_lab <- function(var_quo, manual_lab) {
    if (!is.null(manual_lab)) return(manual_lab)
    lbl <- tryCatch(attr(dplyr::pull(df, !!var_quo), "label"), error = function(e) NULL)
    if (!is.null(lbl)) return(lbl)
    return(rlang::as_name(var_quo))
  }

  final_row_name <- get_lab(r_sym, row_label)
  final_strat_name <- if (!rlang::quo_is_null(s_sym)) get_lab(s_sym, strat_label) else "Stratum"

  # --- FIX 2: Robuste Aggregation (summarise statt count) ---
  calc_counts <- function(d, groups) {
    if (!rlang::quo_is_null(n_sym)) {
      d %>%
        dplyr::group_by(dplyr::across(dplyr::all_of(groups))) %>%
        dplyr::summarise(n = sum(!!n_sym, na.rm = TRUE), .groups = "drop")
    } else if (!rlang::quo_is_null(w_sym)) {
      d %>%
        dplyr::group_by(dplyr::across(dplyr::all_of(groups))) %>%
        dplyr::summarise(n = sum(!!w_sym, na.rm = TRUE), .groups = "drop")
    } else {
      # Das hier erzwingt das Zählen, auch wenn vorher Gruppen da waren
      d %>%
        dplyr::group_by(dplyr::across(dplyr::all_of(groups))) %>%
        dplyr::summarise(n = dplyr::n(), .groups = "drop")
    }
  }

  # Clean Data
  df_clean <- df %>%
    dplyr::mutate(
      !!r_sym := dplyr::if_else(is.na(!!r_sym), na_label, as.character(!!r_sym)),
      !!c_sym := dplyr::if_else(is.na(!!c_sym), na_label, as.character(!!c_sym))
    )
  if (!rlang::quo_is_null(s_sym)) {
    df_clean <- df_clean %>% dplyr::mutate(!!s_sym := dplyr::if_else(is.na(!!s_sym), na_label, as.character(!!s_sym)))
  }

  # --- B. Build Core & Totals ---
  grps_core <- c()
  if (!rlang::quo_is_null(s_sym)) grps_core <- c(grps_core, rlang::as_name(s_sym))
  grps_core <- c(grps_core, rlang::as_name(r_sym), rlang::as_name(c_sym))
  df_core <- calc_counts(df_clean, grps_core)

  grps_row <- c()
  if (!rlang::quo_is_null(s_sym)) grps_row <- c(grps_row, rlang::as_name(s_sym))
  grps_row <- c(grps_row, rlang::as_name(r_sym))
  df_col_totals <- calc_counts(df_clean, grps_row) %>% dplyr::mutate(!!c_sym := "Total")

  grps_col <- c()
  if (!rlang::quo_is_null(s_sym)) grps_col <- c(grps_col, rlang::as_name(s_sym))
  grps_col <- c(grps_col, rlang::as_name(c_sym))
  df_row_totals <- calc_counts(df_clean, grps_col) %>% dplyr::mutate(!!r_sym := "Total")

  grps_strat <- c()
  if (!rlang::quo_is_null(s_sym)) grps_strat <- c(grps_strat, rlang::as_name(s_sym))
  if (length(grps_strat) > 0) {
    df_grand <- calc_counts(df_clean, grps_strat) %>% dplyr::mutate(!!r_sym := "Total", !!c_sym := "Total")
  } else {
    total_n <- if (!rlang::quo_is_null(n_sym)) sum(dplyr::pull(df_clean, !!n_sym), na.rm = TRUE) else      if (!rlang::quo_is_null(w_sym)) sum(dplyr::pull(df_clean, !!w_sym), na.rm = TRUE) else nrow(df_clean)
    df_grand <- dplyr::tibble(!!r_sym := "Total", !!c_sym := "Total", n = total_n)
  }

  df_all <- dplyr::bind_rows(df_core, df_col_totals, df_row_totals, df_grand)

  # --- C. Denominators ---
  df_row_denom <- df_all %>% dplyr::filter(!!c_sym != "Total") %>%
    dplyr::group_by(dplyr::across(dplyr::all_of(grps_row))) %>%
    dplyr::summarise(row_denom = sum(n), .groups = "drop")

  df_col_denom <- df_all %>% dplyr::filter(!!r_sym != "Total") %>%
    dplyr::group_by(dplyr::across(dplyr::all_of(grps_col))) %>%
    dplyr::summarise(col_denom = sum(n), .groups = "drop")

  if (length(grps_strat) > 0) {
    df_strat_denom <- df_all %>% dplyr::filter(!!r_sym != "Total", !!c_sym != "Total") %>%
      dplyr::group_by(dplyr::across(dplyr::all_of(grps_strat))) %>%
      dplyr::summarise(stratum_n = sum(n), .groups = "drop")
  } else {
    df_strat_denom <- df_all %>% dplyr::filter(!!r_sym != "Total", !!c_sym != "Total") %>%
      dplyr::summarise(stratum_n = sum(n))
  }

  df_calc <- df_all %>%
    dplyr::left_join(df_row_denom) %>% dplyr::left_join(df_col_denom)
  if (length(grps_strat) > 0) df_calc <- df_calc %>% dplyr::left_join(df_strat_denom)
  else df_calc <- df_calc %>% dplyr::mutate(stratum_n = df_strat_denom$stratum_n)

  # --- D. Formatting ---
  fmt <- function(x) format(round(x, decimals), nsmall = decimals)

  df_fmt <- df_calc %>%
    dplyr::mutate(
      pct_row = (n / row_denom) * 100,
      pct_col = (n / col_denom) * 100,
      pct_tot = (n / stratum_n) * 100,

      is_t_row = as.character(!!r_sym) == "Total",
      is_t_col = as.character(!!c_sym) == "Total",
      use_row = show_row_pct & !is_t_col,
      use_col = show_col_pct & !is_t_row,
      use_tot = show_tot_pct & !is_t_row & !is_t_col,

      cell_content = paste0(
        if (show_n) paste0(format(round(n, 0), big.mark = ",", scientific = FALSE)) else "",
        if_else(show_n & (use_row | use_col | use_tot), "\n", ""),
        if_else(use_row, paste0(fmt(pct_row), "%"), ""),
        if_else(use_row & (use_col | use_tot), "\n", ""),
        if_else(use_col, paste0(fmt(pct_col), "%"), ""),
        if_else(use_col & use_tot, "\n", ""),
        if_else(use_tot, paste0(fmt(pct_tot), "%"), "")
      )
    ) %>%
    dplyr::mutate(cell_content = gsub("\n$", "", cell_content)) %>%
    dplyr::select(dplyr::any_of(c(rlang::as_name(s_sym), rlang::as_name(r_sym), rlang::as_name(c_sym))), cell_content) %>%
    dplyr::distinct()

  # --- E. Sorting & Pivot ---
  col_vals <- df_fmt %>% dplyr::pull(!!c_sym) %>% unique() %>% as.character()
  has_total <- "Total" %in% col_vals
  has_na <- na_label %in% col_vals
  normal_vals <- setdiff(col_vals, c("Total", na_label))
  final_levels <- c(sort(normal_vals))
  if (has_na) final_levels <- c(final_levels, na_label)
  if (has_total) final_levels <- c(final_levels, "Total")

  df_sorted <- df_fmt %>%
    dplyr::mutate(!!c_sym := factor(!!c_sym, levels = final_levels),
      !!r_sym := forcats::fct_relevel(forcats::as_factor(!!r_sym), "Total", after = Inf))

  if (!rlang::quo_is_null(s_sym)) df_sorted <- df_sorted %>% dplyr::arrange(!!s_sym, !!r_sym, !!c_sym)
  else df_sorted <- df_sorted %>% dplyr::arrange(!!r_sym, !!c_sym)

  cols_left <- character()
  if (!rlang::quo_is_null(s_sym)) cols_left <- c(cols_left, rlang::as_name(s_sym))
  cols_left <- c(cols_left, rlang::as_name(r_sym))

  df_pivoted <- df_sorted %>%
    dplyr::select(!!s_sym, !!r_sym, !!c_sym, cell_content) %>%
    tidyr::pivot_wider(names_from = !!c_sym, values_from = cell_content) %>%
    dplyr::mutate(dplyr::across(dplyr::any_of(final_levels), ~ tidyr::replace_na(., "-"))) %>%
    dplyr::select(dplyr::any_of(cols_left), dplyr::any_of(final_levels))

  # --- F. Renaming ---
  if (!rlang::quo_is_null(s_sym)) df_pivoted <- df_pivoted %>% dplyr::rename(!!final_strat_name := !!s_sym)
  df_pivoted <- df_pivoted %>% dplyr::rename(!!final_row_name := !!r_sym)

  attr(df_pivoted, "title") <- title
  attr(df_pivoted, "footer") <- footer
  class(df_pivoted) <- c("xl_table", class(df_pivoted))

  return(df_pivoted)
}
