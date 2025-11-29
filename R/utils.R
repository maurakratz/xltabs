#' Check if argument is a single string
#' @keywords internal
#' @noRd
check_string <- function(arg, n = 1, call = rlang::caller_env()) {
  if (!rlang::is_character(arg, n = n)) {
    cli::cli_abort(
      c(
        "{.var {rlang::caller_arg(arg)}} must be a character vector of length {n}!",
        "i" = "It was {.type {arg}} of length {length(arg)} instead."
      ),
      call = call
    )
  }
}

#' Check if argument is a single boolean (TRUE/FALSE)
#' @keywords internal
#' @noRd
check_bool <- function(arg, call = rlang::caller_env()) {
  if (!rlang::is_logical(arg, n = 1)) {
    cli::cli_abort(
      c(
        "{.var {rlang::caller_arg(arg)}} must be a single TRUE or FALSE!",
        "i" = "It was {.type {arg}} of length {length(arg)} instead."
      ),
      call = call
    )
  }
}

#' Check if argument is a data frame
#' @keywords internal
#' @noRd
check_df <- function(arg, call = rlang::caller_env()) {
  if (!is.data.frame(arg)) {
    cli::cli_abort(
      c(
        "{.var {rlang::caller_arg(arg)}} must be a data frame!",
        "i" = "It was {.type {arg}} instead."
      ),
      call = call
    )
  }
}
