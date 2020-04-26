#' @importFrom knitr opts_chunk
#' @title knitr hook for figure caption autonumbering
#' @description The function allows you to add a hook when executing
#' knitr to allow to turn figure captions into auto numbered figure
#' captions.
#' @noRd
plot_word_fig_caption <- function(x, options) {

  if(!is.character(options$fig.cap)) options$fig.cap <- NULL
  if(!is.character(options$fig.id)) options$fig.id <- NULL

  cap_str <- pandoc_wml_caption(cap = options$fig.cap, cap.style = options$fig.cap.style,
                                cap.pre = options$fig.cap.pre, cap.sep = options$fig.cap.sep,
                                id = options$fig.id, seq_id = "fig")

  paste("", sprintf("![](%s)", x[1]), cap_str, sep = "\n\n")
}
