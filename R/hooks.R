#' @importFrom knitr opts_chunk
#' @title knitr hook for figure caption autonumbering
#' @description The function allows you to add a hook when executing
#' knitr to allow to turn figure captions into auto numbered figure
#' captions.
#' @noRd
plot_word_fig_caption <- function(x, options) {
  fig.cap.style <- opts_chunk$get("fig.cap.style")
  fig.cap.pre <- opts_chunk$get("fig.cap.pre")
  fig.cap.sep <- opts_chunk$get("fig.cap.sep")

  fig.cap <- opts_current$get("fig.cap")
  fig.id <- opts_current$get("fig.id")

  if(!is.character(fig.cap)) fig.cap <- NULL
  if(!is.character(fig.id)) fig.id <- NULL

  autonum <- run_autonum(seq_id = "fig", pre_label = fig.cap.pre, post_label = fig.cap.sep)

  if( is.null(fig.id) && !is.null(fig.cap)){
    bc <- block_caption(label = fig.cap, style = fig.cap.style, autonum = autonum)
    str <- to_wml(bc, base_document = get_reference_rdocx())
    str <- paste("```{=openxml}", str, "```", sep = "\n")
    paste("", sprintf("![](%s)", x[1]), str, sep = "\n\n")
  } else if( !is.null(fig.id) && !is.null(fig.cap)){
    bc <- block_caption(label = fig.cap, id = fig.id, style = fig.cap.style, autonum = autonum)
    str <- to_wml(bc, base_document = get_reference_rdocx())
    str <- paste("```{=openxml}", str, "```", sep = "\n")
    paste("", sprintf("![](%s)", x[1]), str, sep = "\n\n")
  } else {
    paste("", sprintf("![](%s)", x[1]), sep = "\n\n")
  }
}


