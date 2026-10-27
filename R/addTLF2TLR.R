#' Create the appendix of a docx file
#'
#' This function create the appendix of docx file that as been imported into a R object.
#'
#' @param imported_docx_object An imported docx file saved as R object .
#' @param path_TLFs Path of the TLFs
#' @param output_structure_list List of the outputs to be imported to the TLR
#' @noRd
#' @importFrom rlang expr
#' @importFrom rlang eval_tidy


addTLF2TLR <- function(imported_docx_object = imported_docx_object,
                       path_TLFs = path_TLFs,
                       output_structure_list = output_structure_list,
                       keyword = "LISTOFAPPENDICES") {

  # Path of bookmark outputs
  path_docx_bk <- file.path(path_TLFs, "bookmarked")

  # List of bookmark outputs
  list_outputs <- list.files(path_docx_bk)

  # Path of each bookmark outputs
  file_path_outputs <- file.path(path_docx_bk, list_outputs)


  message("Outputs to be loaded into the appendice: ", paste(list_outputs, collapse = ", "))


  # Printing Sections and associated outputs
  purrr::iwalk(output_structure_list, \(output, section){

    msg <- glue::glue( "Section: {section}; Outputs: {paste(paste0('LK', output), collapse = ', ')}" )
    message(msg)

  })


  #
  docx_appendice <- officer::cursor_reach(x = imported_docx_object, keyword = keyword)
  docx_appendice <- officer::cursor_end(docx_appendice)
  docx_appendice <- officer::body_add_break(docx_appendice, pos = "after")

  docx_appendice <- officer::body_end_section_portrait(docx_appendice)

  title1 <- NULL; iduApp <- NULL; dim1 <- NULL; iduj <- NULL


  for (i2 in 1:length(output_structure_list)) {

    docx_appendice=officer::body_add_par(docx_appendice, names( output_structure_list[i2] ), style = "heading 3", pos = "after")
    docx_appendice=officer::body_add_par(docx_appendice, "")

    sub_list = output_structure_list[[i2]]
    i=1; iduj=NULL

    for (i in 1:length(sub_list)) {
      docx_appendice=officer::body_add_par(docx_appendice, "")
      #print( file_path_outputs[ i ] )
      #idui=idu2[i]
      if (i!=1) {
        docx_appendice=officer::body_add_par(docx_appendice, "", pos = "after")
      }

      path_i <- file.path( path_docx_bk, paste0( "LK", sub_list[ i ], ".docx" ) )

      tb1=officer::read_docx( path_i )
      dim1=tb1$sect_dim

      # Detect landscape from orient attribute OR page dimensions (width > height).
      # Some RTF-to-DOCX converters set landscape dimensions but omit the
      # w:orient="landscape" XML attribute, causing officer to report
      # sect_dim$landscape = FALSE even though the page is landscape.
      is_landscape <- dim1$landscape ||
        (dim1$page["width"] > dim1$page["height"])

      ##
      tbi=officer::docx_summary(tb1)
      tbj=tbi[which(tbi$is_header),]
      tbj=tbj[order(tbj$row_id),]
      idujj=strsplit(tbj$text[1],":")[[1]][1]


      iduApp=c(iduApp, idujj)

      tmpi=officer::docx_summary(tb1)
      temi=tmpi$text
      temi=temi[temi!=""][1]
      temi=substr(temi,regexec("[:]", temi)[[1]][1]+1, nchar(temi))
      isim=1
      while(1){
        temi=gsub("^ ","",temi)
        isim=isim+1
        if (isim==5) break
      }

      title1=c(title1,temi)

      ##
      docx_appendice=officer::body_add_docx(docx_appendice, src= path_i)

      # if (!(i2==length(sec1) & i==length(idu2))) {
      if (is_landscape) {
        docx_appendice=officer::body_end_section_landscape(docx_appendice)
      } else {
        docx_appendice=officer::body_end_section_portrait(docx_appendice)
      }

    }
  }

  docx_appendice=officer::cursor_end(docx_appendice)
  docx_appendice=officer::body_remove(docx_appendice)
  docx_appendice=officer::cursor_end(docx_appendice)
  if (!is.null(dim1)) {
    is_landscape <- dim1$landscape ||
      (dim1$page["width"] > dim1$page["height"])
    if (is_landscape) {
      docx_appendice=officer::body_end_section_landscape(docx_appendice)
    }
  }


  docx_appendice=officer::cursor_reach(docx_appendice, keyword)

  for (i in 1:length(iduApp)) {

    idui=iduApp[i]
    idui=gsub(" ","_",idui)

    # Build the cross-reference field expression now that idui is set.
    # Bold formatting is applied to the output ID via fp_text(bold = TRUE).
    def_expr <- expr(officer::run_word_field(sprintf('REF %s \\h ', idui)))

    #if (!idui %in% idu0) next
    docx_appendice=officer::body_add_fpar(docx_appendice, pos="after",
                                          value= officer::fpar(
                                            officer::ftext(iduApp[i], officer::fp_text(bold = TRUE)),
                                            eval_tidy( def_expr ),
                                            paste0(": ", title1[i])),
                                          style = NULL)
  }

  docx_appendice=officer::body_replace_all_text(docx_appendice, keyword, "List of Tables/Figures")

  return(docx_appendice)

}
