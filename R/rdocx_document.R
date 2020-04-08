# utils ----
#' @importFrom utils getAnywhere getFromNamespace
get_fun <- function(x){
  if( grepl("::", x, fixed = TRUE) ){
    coumpounds <- strsplit(x, split = "::", x, fixed = TRUE)[[1]]
    z <- getFromNamespace(coumpounds[2], ns = coumpounds[1] )
  } else {
    z <- getAnywhere(x)
    if(length(z$objs) < 1){
      stop("could not find any function named ", shQuote(z$name), " in loaded namespaces or in the search path. If the package is installed, specify name with `packagename::function_name`.")
    }
  }
  z
}

file_with_meta_ext <- function(file, meta_ext, ext = tools::file_ext(file)) {
  paste(tools::file_path_sans_ext(file),
        ".", meta_ext, ".", ext,
        sep = ""
  )
}

#' @export
#' @title Convert to an MS Word document
#' @description Format for converting from R Markdown to an MS Word
#' document. The function comes also with improved output options.
#' @param mapstyles a named list of style to be replaced in the generated
#' document. `list("Date"="Author")` will result in a document where
#' all paragraphs styled with stylename "Date" will be styled with
#' stylename "Author".
#' @param base_format a scalar character, format to be used as a base document for
#' officedown. default to [word_document][rmarkdown::word_document] but
#' can also be word_document2 from bookdown
#' @param tab_caption a list that can contain items `style`, `pre` and `sep`:
#'
#' * style: the Word style name to use for table captions.
#' * pre: the prefix for numbering chunk (default to "Table ").
#' * sep: the suffix for numbering chunk (default to ": ").
#'
#' The default is producing a numbering chunk of the form "Table 2: ".
#'
#' Default values are `list(style = "Table Caption", pre = "Table ", sep = ": ")`.
#'
#' Missing items will be replace by default values, e.g. `list(pre="tab.")`
#' will produce "tab. 2: ".
#' @param plot_caption a list that can contain items `style`, `pre` and `sep`:
#'
#' * style: the Word style name to use for figure captions.
#' * pre: the prefix for numbering chunk (default to "Figure ").
#' * sep: the suffix for numbering chunk (default to ": ").
#'
#' The default is producing a numbering chunk of the form "Figure 2: ".
#'
#' Default values are `list(style = "Figure Caption", pre = "Figure ", sep = ": ")`.
#'
#' Missing items will be replace by default values, e.g. `list(pre="fig.")`
#' will produce "fig. 2: ".
#' @param tab.style default table style. This is the name of the Word style to
#' use for tables. It can be read with `officer::styles_info()` or within
#' Word table styles view.
#' @param ... arguments used by [word_document][rmarkdown::word_document]
#' @examples
#' skeleton <- system.file(package = "officedown",
#'   "rmarkdown/templates/word/skeleton/skeleton.Rmd")
#' file.copy(skeleton, to = "officedown.Rmd")
#' library(rmarkdown)
#' render("officedown.Rmd", output_file = "officedown.docx")
#' @importFrom officer change_styles
#' @section R Markdown yaml:
#' The following demonstrates how to pass arguments in the R Markdown yaml:
#'
#' ```
#' ---
#' title: "Word document"
#' output:
#'   bookdown::markdown_document2:
#'   base_format: officedown::rdocx_document
#'   tab_caption:
#'     style: "Table Caption"
#'     pre: "Table "
#'     sep: ": "
#'   plot_caption:
#'     style: "Captioned Figure"
#'     pre: "Figure "
#'     sep: ": "
#'   tab.style: "Light List Accent 2"
#' ---
#' ```
rdocx_document <- function(mapstyles, base_format = "rmarkdown::word_document",
                           tab_caption = list(), plot_caption = list(),
                           tab.style = "Table", ...) {

  base_format_fun <- get_fun(base_format)
  output_formats <- base_format_fun(...)

  plot_caption_ <- list(style = "Table Caption", pre = "Figure ", sep = ": ")
  plot_caption <- modifyList(plot_caption_, plot_caption)
  tab_caption_ <- list(style = "Captioned Figure", pre = "Table ", sep = ": ")
  tab_caption <- modifyList(tab_caption_, tab_caption)

  output_formats$knitr$opts_chunk <- append(
    output_formats$knitr$opts_chunk,
    list(tab.cap.style = tab_caption$style,
         tab.cap.pre = tab_caption$pre,
         tab.cap.sep = tab_caption$sep,
         tab.lp = "tab:",
         tab.style = tab.style,
         fig.cap.style = plot_caption$style,
         fig.cap.pre = plot_caption$pre,
         fig.cap.sep = plot_caption$sep,
         fig.lp = "fig:"
         )
    )

  if( missing(mapstyles) )
    mapstyles <- list()

  output_formats$post_knit <- function(metadata, input_file, runtime, ...){
    output_file <- file_with_meta_ext(input_file, "knit", "md")
    content <- readLines(output_file)

    content <- post_knit_table_captions( content,
      tab.cap.pre = tab_caption$pre, tab.cap.sep = tab_caption$sep)
    content <- post_knit_references(content, lp = "tab:")
    content <- post_knit_references(content, lp = "fig:")

    writeLines(content, output_file)
  }
  output_formats$pre_processor = function(metadata, input_file, runtime, knit_meta, files_dir, output_dir){
    md <- readLines(input_file)
    md <- chunk_macro(md)
    md <- block_macro(md)
    writeLines(md, input_file)
  }

  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_docx(output_file)
    x <- process_images(x)
    x <- process_links(x)
    x <- process_embedded_docx(x)
    x <- process_chunk_style(x)
    x <- process_sections(x)
    x <- process_par_settings(x)
    x <- change_styles(x, mapstyles = mapstyles)

    print(x, target = output_file)
    output_file
  }
  output_formats$bookdown_output_format = 'docx'
  output_formats

}




#' @importFrom officer get_reference_value
get_docx_uncached <- function() {
  ref_docx <- read_docx(get_reference_value(format = "docx"))
  ref_docx
}

#' @importFrom memoise memoise
get_reference_rdocx <- memoise(get_docx_uncached)



