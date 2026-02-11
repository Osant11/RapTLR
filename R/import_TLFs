#' Create a copy of TLFs folder within the current environment
#'
#' This function copy the TLFs folder from RapTLR and create a copy called "TLF_outputs" in the current working directory
#'
#' @examples
#' # Just run the import_TLFs() function to create a copy folder TLF_outputs in your current working environment
#' RapTLR::import_TLFs( )
#' 
#' @importFrom utils packageName
#' @export
import_TLFs <- function( ) {

  src <- system.file(
   "extdata/TLF_outputs",
   package = utils::packageName( )
 )
  
 if ( src == " " ) {
   stop( "TLF_outputs not found in inst/extdata" )
 }
  
 dest <- file.path( getwd() )
 file.copy( from = src, to = dest, overwrite = TRUE, recursive = TRUE )
 
 message( "[INFO] TLF_outputs has been successfully imported. File location: ", dest )
  
 invisible( dest )
}
