sanitize_tab.lp <- function(tab.lp){
  if(is.null(tab.lp)){
    tab.lp <- "tab"
  } else {
    tab.lp <- gsub("[:]{1}$", "", tab.lp)
  }
  tab.lp
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
    tab.cap.style <- si[si$style_type %in% "paragraph" & si$is_default ,"style_name"]
  } else {
    tab.cap.style <- si[si$style_type %in% "paragraph" & si$style_name %in% tab.cap.style, "style_name"]
    if(length(tab.cap.style) != 1 ){
      msg <- paste0("could not find paragraph style ", shQuote(tab.cap.style), ". Switching to default one named ")
      tab.cap.style <- si[si$style_type %in% "paragraph" & si$is_default ,"style_name"]
      warning(paste0(msg, shQuote(tab.cap.style), "."))
    }
  }

  if(is.null(tab.cap.pre)){
    tab.cap.pre <- "table "
  }
  if(is.null(tab.cap.sep)){
    tab.cap.sep <- ": "
  }

  if(is.null(tab.style)){
    tab.style <- si[si$style_type %in% "table" & si$is_default ,"style_name"]
  } else {
    tab.style <- si[si$style_type %in% "table" & si$style_name %in% tab.style, "style_name"]
    if(length(tab.style) != 1 ){
      msg <- paste0("could not find table style ", shQuote(tab.style), ". Switching to default one named ")
      tab.style <- si[si$style_type %in% "table" & si$is_default ,"style_name"]
      warning(paste0(msg, shQuote(tab.style), "."))
    }
  }

  tab.cap.style_id <- si[
    si$style_type %in% "paragraph" &
      si$style_name %in% tab.cap.style ,
    "style_id"]

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

