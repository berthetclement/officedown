#' @importFrom rmarkdown pandoc_version
#' @importFrom knitr knit_print asis_output opts_knit
check_all <- function(){
  if (is.null(opts_knit$get("rmarkdown.pandoc.to")))
    stop("`knit_print.ooxml` needs to be used as a renderer for ",
         "an rmarkdown R code chunk (render by rmarkdown)", call. = FALSE)

  if (!(pandoc_version() >= 2))
    stop("pandoc version >= 2.0 required for ooxml rendering in docx", call. = FALSE)

  if (!(grepl("docx", opts_knit$get("rmarkdown.pandoc.to"))))
    stop("unsupported format for ooxml rendering:", opts_knit$get("rmarkdown.pandoc.to"), call. = FALSE)
  invisible()
}


#' @export
#' @importFrom officer to_wml
knit_print.ooxml_str_chunk <- function(x, ...){
  knit_print( asis_output(
    paste0("`", x, "`{=openxml}")
  ) )
}

#' @export
#' @importFrom officer to_wml
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
  browser()
  knit_print( asis_output(
    paste("```{=openxml}", to_wml(x), "```", sep = "\n")
  ) )
}

