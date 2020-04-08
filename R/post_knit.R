#' @importFrom officer run_reference to_wml
as_reference <- function(z){
  str <- to_wml(run_reference(z))
  paste0("`<w:hyperlink w:anchor=\"", z, "\">", str, "</w:hyperlink>`{=openxml}")
}

#' @importFrom uuid UUIDgenerate
as_bookmark <- function(id, str) {
  new_id <- uuid::UUIDgenerate()
  bm_start_str <- sprintf("<w:bookmarkStart w:id=\"%s\" w:name=\"%s\"/>", new_id, id)
  bm_start_end <- sprintf("<w:bookmarkEnd w:id=\"%s\"/>", new_id)
  paste0("`", bm_start_str, str, bm_start_end, "`{=openxml}")
}

post_knit_table_captions <- function(content, tab.cap.pre, tab.cap.sep){

  is_captions <- grepl("<caption>\\(\\\\#tab:[-[:alnum:]]+\\)(.*)</caption>", content)
  if(any(is_captions)){
    captions <- content[is_captions]
    ids <- gsub("<caption>\\(\\\\#tab:([-[:alnum:]]+)\\)(.*)</caption>", "\\1", captions)
    labels <- gsub("<caption>\\(\\\\#tab:[-[:alnum:]]+\\)(.*)</caption>", "\\1", captions)
    str <- mapply(function(label, id){
      autonum <- run_autonum(seq_id = "tab", pre_label = tab.cap.pre, post_label = tab.cap.sep)
      autonum <- to_wml(autonum)
      run_str <- sprintf("<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>", htmlEscapeCopy(label))
      run_str <- paste0(autonum, run_str)
      as_bookmark(id, run_str)
    }, label = labels, id = ids, SIMPLIFY = FALSE)
    content[is_captions] <- unlist(str)
  }
  content
}

post_knit_references <- function(content, lp){

  if(!lp %in% c("tab:", "fig:")){
    stop("lp must be one of `tab:` or `fig:`.")
  }
  regexpr_str <- paste0('\\\\@ref\\(', lp, '([-[:alnum:]]+)\\)')

  gmatch <- gregexpr(regexpr_str, content)
  result <- regmatches(content,gmatch)

  result <- lapply(result, function(z){
    if(length(z) > 0){
      ids <- gsub(regexpr_str, '\\1', z)
      as_reference(ids)
    } else z
  })
  regmatches(content,gmatch) <- result
  content
}
