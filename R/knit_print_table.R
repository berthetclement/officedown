sanitize_tab.lp <- function(tab.lp){
  if(is.null(tab.lp)){
    tab.lp <- "tab"
  } else {
    tab.lp <- gsub("[:]{1}$", "", tab.lp)
  }
  tab.lp
}


default_style <- function(type, si){
  si[si$style_type %in% type & si$is_default ,"style_name"]
}
style_id <- function(x, type, si){
  si[
    si$style_type %in% type &
      si$style_name %in% x ,
    "style_id"]
}
validate_style <- function(x, type, si){
  validated_style <- si[si$style_type %in% type & si$style_name %in% x, "style_name"]
  if(length(validated_style) != 1 ){
    validated_style <- default_style(type, si)
    msg <- paste0("could not find ", type, " style ", shQuote(x),
                  ". Switching to default one named ", shQuote(validated_style), ".")
    warning(msg, call. = FALSE)
  }
  validated_style
}

get_table_design_opt <- function(x, default = FALSE){
  x <- opts_current$get(x)
  if(is.null(x)) x <- default
  x
}

#' @importFrom officer styles_info
opts_current_table <- function(){
  tab.cap.style <- opts_chunk$get("tab.cap.style")
  tab.cap.pre <- opts_chunk$get("tab.cap.pre")
  tab.cap.sep <- opts_chunk$get("tab.cap.sep")
  tab.cap <- opts_current$get("tab.cap")
  tab.id <- opts_current$get("tab.id")
  tab.style <- opts_current$get("tab.style")

  doc <- get_reference_rdocx()
  si <- styles_info(doc)

  if(is.null(tab.cap.style)){
    tab.cap.style <- default_style("paragraph", si)
  } else {
    tab.cap.style <- validate_style(x = tab.cap.style, type = "paragraph", si = si)
  }
  tab.cap.style_id <- style_id(tab.cap.style, type = "paragraph", si)

  if(is.null(tab.cap.pre)){
    tab.cap.pre <- "table "
  }
  if(is.null(tab.cap.sep)){
    tab.cap.sep <- ": "
  }

  if(is.null(tab.style)){
    tab.style <- default_style("table", si)
  } else {
    tab.style <- validate_style(x = tab.style, type = "table", si = si)
  }


  list(cap.style = tab.cap.style, cap.style_id = tab.cap.style_id,
       cap.pre = tab.cap.pre, cap.sep = tab.cap.sep,
       id = tab.id, cap = tab.cap,
       style = tab.style, seq_id = "tab"
       )

}

# knit_print.data.frame -----

#' @importFrom officer block_table
#' @importFrom knitr knit_print asis_output opts_current
knit_print.data.frame <- function(x, ...) {

  if( grepl( "docx", knitr::opts_knit$get("rmarkdown.pandoc.to") ) ){
    tab_props <- opts_current_table()

    bt <- block_table(x, style = tab_props$style,
                      header = get_table_design_opt("header", default = TRUE),
                      first_row = get_table_design_opt("first_row", default = TRUE),
                      first_column = get_table_design_opt("first_column"),
                      last_row = get_table_design_opt("last_row"),
                      last_column = get_table_design_opt("last_column"),
                      no_hband = get_table_design_opt("no_hband"),
                      no_vband = get_table_design_opt("no_vband")
                      )

    cap_str <- do.call(pandoc_wml_caption, tab_props)
    res <- paste(cap_str, "```{=openxml}",
                 to_wml(bt, base_document = get_reference_rdocx()),
                 "```\n\n",
                 sep = "\n")
    asis_output(res)
  } else {
    if(is.null( layout <- knitr::opts_current$get("layout") )){
      layout <- officer::ph_location_type()
    }
    location <- get_content_layout(layout)

    tab_props <- opts_current_table()
    bt <- block_table(x)
    res <- paste("```{=openxml}",
                 officer::to_pml(bt, left = location$left, top = location$top,
                                 width = location$width, height = location$height,
                                 label = location$ph_label, ph = location$ph,
                                 rot = location$rotation, bg = location$bg),
                 "```\n\n",
                 sep = "\n")
    asis_output(res)
  }
}

