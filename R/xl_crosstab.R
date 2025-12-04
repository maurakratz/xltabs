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
#' @importFrom dplyr mutate filter group_by ungroup count bind_rows arrange select pull if_else rename
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
                        # NEU: Pre-aggregated Counts
                        counts_col = NULL,
                        row_label = NULL,
                        col_label = NULL,
                        strat_label = NULL,
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

  if (!is.null(row_label)) check_string(row_label)
  if (!is.null(col_label)) check_string(col_label)
  if (!is.null(strat_label)) check_string(strat_label)

  # --- Variablen einfangen ---
  r_sym <- enquo(row_var)
  c_sym <- enquo(col_var)
  s_sym <- enquo(strat_var)
  w_sym <- enquo(w_var)
  cnt_sym <- enquo(counts_col) # NEU

  # --- LABEL LOGIK ---
  get_var_label <- function(dataset, quo_col, manual_lab) {
    if (!is.null(manual_lab)) return(manual_lab)
    col_name <- rlang::as_name(quo_col)
    lbl <- attr(dataset[[col_name]], "label")
    if (!is.null(lbl)) return(lbl)
    return(col_name)
  }

  final_row_name <- get_var_label(data, r_sym, row_label)
  final_strat_name <- if (!quo_is_null(s_sym)) get_var_label(data, s_sym, strat_label) else NULL

  # --- 2. Prepare Data ---
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

  # Convert to character
  df_clean <- df_clean %>%
    mutate(
      !!r_sym := as.character(!!r_sym),
      !!c_sym := as.character(!!c_sym)
    )

  if (!quo_is_null(s_sym)) {
    df_clean <- df_clean %>% mutate(!!s_sym := as.character(!!s_sym))
  }

  # --- NEU: Helper Logic ---
  # Priorität:
  # 1. Wenn counts_col da ist -> Nimm das als 'wt'
  # 2. Wenn w_var da ist -> Nimm das als 'wt'
  # 3. Sonst -> Einfaches count()

  calc_counts <- function(d, groups) {
    if (!quo_is_null(cnt_sym)) {
      # Fall: Pre-aggregated
      d %>% count(!!!groups, wt = !!cnt_sym, name = "n")
    } else if (!quo_is_null(w_sym)) {
      # Fall: Weighted Raw Data
      d %>% count(!!!groups, wt = !!w_sym, name = "n")
    } else {
      # Fall: Raw Data
      d %>% count(!!!groups, name = "n")
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
  qs_strat <- if (!quo_is_null(s_sym)) quos(!!s_sym) else quos()
  qs_row   <- quos(!!r_sym)
  qs_col   <- quos(!!c_sym)

  df_calc <- df_all %>%
    group_by(!!!qs_strat) %>%
    mutate(stratum_n = sum(n[as_factor(!!r_sym) != "Total" & as_factor(!!c_sym) != "Total"])) %>%
    group_by(!!!c(qs_strat, qs_row)) %>%
    mutate(row_denom = sum(n[as_factor(!!c_sym) != "Total"])) %>%
    group_by(!!!c(qs_strat, qs_col)) %>%
    mutate(col_denom = sum(n[as_factor(!!r_sym) != "Total"])) %>%
    ungroup() %>%
    mutate(
      pct_row = n / row_denom,
      pct_col = n / col_denom,
      pct_tot = n / stratum_n
    )

  # --- D. Formatting ---
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

  # --- E. Sorting & Pivot ---
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

  # FIX: Spalten-Reihenfolge vorab definieren
  # Wir sammeln erst die Namen der "linken" Spalten (Stratum und Row)
  cols_left <- character()

  # Nur wenn Stratum existiert, fügen wir den Namen hinzu
  if (!quo_is_null(s_sym)) {
    cols_left <- c(cols_left, rlang::as_name(s_sym))
  }
  # Row existiert immer
  cols_left <- c(cols_left, rlang::as_name(r_sym))

  df_pivoted <- df_sorted %>%
    select(!!s_sym, !!r_sym, !!c_sym, cell_content) %>%
    pivot_wider(
      names_from = !!c_sym,
      values_from = cell_content,
      values_fill = "-"
    ) %>%
    # HIER ist der Fix: Wir nutzen den vorbereiteten Vektor 'cols_left'
    select(any_of(cols_left), any_of(final_levels))

  # --- F. Final Renaming ---
  if (!quo_is_null(s_sym)) {
    df_pivoted <- df_pivoted %>%
      rename(!!final_strat_name := !!s_sym)
  }

  df_pivoted <- df_pivoted %>%
    rename(!!final_row_name := !!r_sym)

  return(df_pivoted)
}
