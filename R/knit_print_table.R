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

  list(tab.cap.style = tab.cap.style, tab.cap.style_id = tab.cap.style_id,
       tab.cap.pre = tab.cap.pre, tab.cap.sep = tab.cap.sep,
       tab.id = tab.id, tab.cap = tab.cap,
       tab.style = tab.style)

}

wml_table_caption <- function(all_properties){

  if( is.null(all_properties$tab.cap)) return("")

  par_style <- paste0("<w:pStyle w:val=\"", all_properties$tab.cap.style_id, "\"/>")

  autonum <- run_autonum(seq_id = "tab",
                         pre_label = all_properties$tab.cap.pre,
                         post_label = all_properties$tab.cap.sep)
  autonum <- to_wml(autonum)
  run_str <- sprintf("<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>",
                     htmlEscapeCopy(all_properties$tab.cap))
  run_str <- paste0(autonum, run_str)

  if(!is.null(all_properties$tab.id)) {
    run_str <- as_bookmark(all_properties$tab.id, run_str)
  }

  paste0("<w:p><w:pPr>", par_style, "</w:pPr>", run_str, "</w:p>")
}

# knit_print.data.frame -----

#' @importFrom officer block_table
#' @importFrom knitr knit_print asis_output opts_current
knit_print.data.frame <- function(x, ...) {
  tab_props <- opts_current_table()
  bt <- block_table(x, style = tab_props$tab.style)
  cap_str <- wml_table_caption(tab_props)
  res <- paste("```{=openxml}", cap_str,
               to_wml(bt, base_document = get_reference_rdocx()),
               "```\n\n",
               sep = "\n")
  asis_output(res)
}

