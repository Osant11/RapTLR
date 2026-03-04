#' Create a copy of RTFs outputs within the current environment
#'
#' This function copy the dummy RTF outputs from the package and create a folder called "rtf_outputs" in the current working directory
#'
#' @examples
#' # Just run the import_RTFs() function to create a copy folder within the RTFs outputs in your current working environment
#' RapTLR::import_RTFs( )
#' 
#' @importFrom utils packageName
#' @export
import_RTFs <- function( ) {

  src <- system.file(
   "extdata/rtf_outputs",
   package = utils::packageName( )
 )
  
 if ( src == " " ) {
   stop( "rtf_outputs not found in inst/extdata" )
 }
  
 dest <- file.path( getwd() )
 file.copy( from = src, to = dest, overwrite = TRUE, recursive = TRUE )
 
 message( "[INFO] rtf_outputs folder has been successfully imported and copied under: ", dest )
  
 invisible( dest )
}
