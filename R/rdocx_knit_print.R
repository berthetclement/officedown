#' @export
#' @importFrom officer to_wml
#' @importFrom knitr knit_print asis_output opts_knit
knit_print.run <- function(x, ...){
  knit_print( asis_output(
    paste("`", to_wml(x), "`{=openxml}", sep = "")
  ) )
}

#' @export
knit_print.fp_par <- function(x, ...){
  knit_print( asis_output(
    paste("`", to_wml(x), "`{=openxml}", sep = "")
  ) )
}

#' @export
knit_print.block <- function(x, ...){
  knit_print( asis_output(
    paste("```{=openxml}", to_wml(x), "```", sep = "\n")
  ) )
}

