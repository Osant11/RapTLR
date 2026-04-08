#' Generate Most Common Adverse Events Summary Text
#'
#' This function analyzes adverse event data to identify and summarize the most common
#' treatment-emergent adverse events (TEAEs) by treatment group. It merges AE data with
#' subject-level data, calculates frequencies, filters by a specified threshold, and
#' generates a formatted text summary suitable for clinical study reports.
#'
#' @param arg_data A data frame containing adverse event data (e.g., ADAE dataset).
#'   Must include USUBJID for subject identification.
#' @param arg_data_adsl A data frame containing subject-level data (ADSL dataset).
#'   Must include USUBJID and the grouping variable specified in arg_grp.
#' @param arg_grp A grouping variable (unquoted) for treatment groups (e.g., ARM, TRT01A).
#' @param arg_var A variable (unquoted) containing the adverse event terms to analyze
#'   (e.g., AEDECOD for preferred terms, AEBODSYS for system organ class).
#' @param arg_fl_rest Unquoted expression. Optional filter condition to subset the AE
#'   dataset before analysis. Default is NULL (no filtering).
#' @param arg_trt_flt A treatment group variable (unquoted, can use backticks for names
#'   with spaces) to use for filtering. Only events with frequency >= arg_frq in this
#'   treatment group will be included. Default is NULL.
#' @param arg_frq Numeric threshold for minimum frequency percentage. Only events with
#'   frequency >= this value in the arg_trt_flt treatment group are included. Default is 15.
#' @param arg_freq_num Logical. If TRUE, returns frequency as numeric; if FALSE, as
#'   percentage string. Default is TRUE.
#' @param arg_rounding Number of decimal places for rounding frequencies. Default is 1.
#'
#' @return A character string containing a formatted summary of the most common adverse
#'   events, including frequencies for each treatment group, suitable for inclusion in
#'   clinical study reports.
#'
#' @examples
#'
#' # Find most common AEs (>=15%) in Xanomeline High Dose group
#' fct_mst_aes(
#'   arg_data = tlr_adae,
#'   arg_data_adsl = tlr_adsl,
#'   arg_grp = ARM,
#'   arg_var = AEDECOD,
#'   arg_fl_rest = TRTEMFL == "Y",
#'   arg_trt_flt = `Xanomeline High Dose`
#' )
#'
#' # Find most common AEs (>=10%) with 2 decimal places
#' fct_mst_aes(
#'   arg_data = tlr_adae,
#'   arg_data_adsl = tlr_adsl,
#'   arg_grp = TRT01A,
#'   arg_var = AEDECOD,
#'   arg_fl_rest = TRTEMFL == "Y",
#'   arg_trt_flt = `Xanomeline High Dose`,
#'   arg_frq = 10,
#'   arg_rounding = 2
#' )
#'
#' # Analyze by system organ class instead of preferred term
#' fct_mst_aes(
#'   arg_data = tlr_adae,
#'   arg_data_adsl = tlr_adsl,
#'   arg_grp = ARM,
#'   arg_var = AEBODSYS,
#'   arg_fl_rest = TRTEMFL == "Y",
#'   arg_trt_flt = `Xanomeline High Dose`,
#'   arg_frq = 20
#' )
#'
#' @export


fct_mst_aes <- function( arg_data, arg_data_adsl, arg_grp, arg_var, arg_fl_rest = NULL, arg_trt_flt = NULL, arg_frq = 15, arg_freq_num = TRUE, arg_rounding = 1 ) {

  cond_fl <- rlang::enquo( arg_fl_rest )

  data_grp <- arg_data_adsl %>%
    dplyr::select( USUBJID, {{ arg_grp }} )

  int_col <- dplyr::intersect( colnames( data_grp ), colnames( arg_data ) )

  arg_data %>%
    dplyr::filter( if( !rlang::quo_is_null( cond_fl ) ) !!cond_fl else TRUE ) %>%
    dplyr::left_join( data_grp, by = int_col ) %>%
    dplyr::group_by( {{ arg_grp }}, {{ arg_var }} ) %>% 
    dplyr::summarise( count = dplyr::n_distinct( USUBJID ), .groups = "drop" ) %>% 
    dplyr::left_join( fct_grp_tbl( arg_data = arg_data_adsl, arg_grp = {{ arg_grp }} ), by = dplyr::join_by( {{ arg_grp }} ) ) %>%
    dplyr::mutate( dplyr::across( dplyr::everything(), ~ ifelse( is.na( . ), 0, . ) ),
                   Freq = round( 100 * count / n_rand, arg_rounding ) ) %>% 
    dplyr::arrange( dplyr::desc( count ) ) %>%
    tidyr::pivot_wider( id_cols = -c( n_rand, count ),
                        names_from =  {{ arg_grp }},
                        values_from = c( Freq ), 
                        values_fill = 0 ) %>% 
    dplyr::arrange( dplyr::desc( {{ arg_trt_flt }} ) ) %>% 
    dplyr::filter( {{ arg_trt_flt }} >= arg_frq ) %>%
    dplyr::mutate( dplyr::across( dplyr::where( is.numeric ), ~ paste0( ., "%" ) ),
                   dplyr::across( -{{ arg_var }}, ~ paste( .x, stringr::str_to_sentence( dplyr::cur_column( ) ) ) ) ) %>% 
    tidyr::unite( LABEL_TEXT, -{{ arg_var }}, sep = "; " ) %>% 
    dplyr::mutate( LABEL_TEXT = paste0( " (", LABEL_TEXT, ")" ), 
                   {{ arg_var }} := tolower( {{ arg_var }} ) ) %>% 
    tidyr::unite( LABEL_TEXT, dplyr::everything( ), sep = "" ) %>% 
    dplyr::pull( LABEL_TEXT ) %>% 
    paste( collapse = ", " ) %>% 
    stringr::str_replace( replacement = "), and", pattern = "(\\),)(?!(.|\n)*\\),)" ) %>% 
    paste0( "The most common TEAEs with frequency ≥", 
            arg_frq, 
            "% on ",  stringr::str_to_title( rlang::as_name( rlang::enquo( arg_trt_flt ) ) ),  " were: ", . )
  
}
