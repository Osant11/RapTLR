#' Create a copy of demo_TLR.R within the current environment
#'
#' This function copy the example code and create a copy called "demo_TLR.R" in the current working directory
#'
#' @examples
#' # Just run the import_demo() function to create a copy called "demo_TLR.R" in your current working environment
#' import_demo( )
#' 
#' @importFrom utils packageName
#' @export
import_demo <- function( ) {

  src <- system.file(
   "scripts", "demo_TLR.R",
   package = utils::packageName( )
 )
  
 if ( src == " " ) {
   stop( "demo_TLR.R not found in inst/scripts" )
 }
  
 dest <- file.path( getwd(), "demo_TLR.R" )
 file.copy( from = src, to = dest, overwrite = TRUE )
 
 message( "[INFO] demo_TLR.R has been successfully imported. File location: ", dest )
  
 invisible( dest )
}
