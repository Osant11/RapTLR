#' Create a copy of demo_pptx.R within the current environment
#'
#' This function copies the PowerPoint demo script and creates a file called
#' `demo_pptx.R` in the current working directory, ready to run.
#'
#' @examples
#' # Copy the PowerPoint demo script to the current working directory
#' RapTLR::import_pptx_demo()
#'
#' @importFrom utils packageName
#' @export
import_pptx_demo <- function( ) {

  src <- system.file(
    "scripts", "demo_pptx.R",
    package = utils::packageName( )
  )

  if ( src == "" ) {
    stop( "demo_pptx.R not found in inst/scripts" )
  }

  dest <- file.path( getwd(), "demo_pptx.R" )
  file.copy( from = src, to = dest, overwrite = TRUE )

  message( "[INFO] demo_pptx.R has been successfully imported. File location: ", dest )

  invisible( dest )
}
