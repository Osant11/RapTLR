#' Generate Grouped Variable Summary Text for TLR Reports
#'
#' @description
#' Creates formatted narrative text summarizing the distribution of a categorical variable
#' across different groups. Particularly useful for generating dynamic text in Top-Line
#' Reports (TLR) that describes patient characteristics, or other categorical
#' outcomes by treatment group or other stratification variables.
#'
#' @param arg_text_desc Character string. Optional descriptive text to prepend to the
#'   generated summary. Default is an empty string.
#' @param arg_dataset Data frame. The dataset containing the variables to summarize.
#' @param arg_fl_rest Unquoted expression. Optional filter condition to subset the dataset
#'   before analysis. Default is NULL (no filtering).
#' @param arg_var_rest Unquoted variable name. The categorical variable to summarize
#'   (e.g., adverse event term, discontinuation reason).
#' @param arg_group_var Unquoted variable name. The grouping variable (e.g., treatment arm,
#'   study site).
#' @param arg_group_vals Character vector. Optional subset of group values to include in
#'   the analysis. If NULL, all groups are included. Default is NULL.
#' @param arg_group_lbl Character vector. Optional custom labels for groups, corresponding
#'   to `arg_group_vals`. If NULL, original group values are used as labels. Default is NULL.
#' @param arg_order Logical. If TRUE, orders results by descending overall frequency.
#'   If FALSE, maintains original variable order. Default is TRUE.
#' @param arg_cvt_stg Function. Transformation function to apply to variable labels.
#'   Default is `I` (identity function, no transformation). Use `tolower`, `toupper`, or
#'   custom functions to modify label formatting.
#' @param arg_nbr_obs Numeric. Maximum number of categories to include in the output text.
#'   Default is Inf (all categories).
#'
#' @return A character string containing the formatted summary text with percentages and
#'   group distributions.
#'
#' @details
#' The function calculates:
#' \itemize{
#'   \item Overall percentage of subjects with each category
#'   \item Within-category distribution across groups
#'   \item Formatted narrative text combining these statistics
#' }
#'
#' @examples
#' # Example 1: Basic usage - summarize discontinuation reasons by treatment
#' # Shows all discontinuation reasons ordered by frequency
#' fct_var_text_group(
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P
#' )
#'
#' # Example 2: Add descriptive text prefix
#' # Prepends custom text to the generated summary
#' fct_var_text_group(
#'   arg_text_desc = "The most common reasons for discontinuation were:",
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P
#' )
#'
#' # Example 3: Disable ordering to maintain original variable order
#' # Useful when variable has inherent ordering (e.g., severity grades)
#' fct_var_text_group(
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P,
#'   arg_order = FALSE
#' )
#'
#' # Example 4: Limit to specific treatment groups
#' # Excludes placebo group from the analysis
#' fct_var_text_group(
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P,
#'   arg_group_vals = c("Xanomeline High Dose", "Xanomeline Low Dose")
#' )
#'
#' # Example 5: Use custom group labels for cleaner output
#' # Shortens long treatment names for readability
#' fct_var_text_group(
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P,
#'   arg_group_vals = c("Xanomeline High Dose", "Xanomeline Low Dose", "Placebo"),
#'   arg_group_lbl = c("High Dose", "Low Dose", "Placebo")
#' )
#'
#' # Example 6: Limit output to top N categories
#' # Shows only the 3 most common discontinuation reasons
#' fct_var_text_group(
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P,
#'   arg_group_vals = c("Xanomeline High Dose", "Xanomeline Low Dose", "Placebo"),
#'   arg_nbr_obs = 3
#' )
#'
#' # Example 7: Apply text transformation to variable labels
#' # Converts labels to lowercase for consistent formatting
#' fct_var_text_group(
#'   arg_dataset = tlr_adsl,
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P,
#'   arg_cvt_stg = tolower,
#'   arg_nbr_obs = 2
#' )
#'
#'
#' # Example 8: Combine multiple parameters for precise control
#' # Filtered data, custom labels, limited output, with descriptive text
#' fct_var_text_group(
#'   arg_text_desc = "Among safety population patients,",
#'   arg_dataset = tlr_adsl,
#'   arg_fl_rest = SAFFL == "Y",
#'   arg_var_rest = DCREASCD,
#'   arg_group_var = TRT01P,
#'   arg_group_vals = c("Xanomeline High Dose", "Xanomeline Low Dose", "Placebo"),
#'   arg_group_lbl = c("XAN-H", "XAN-L", "PBO"),
#'   arg_order = TRUE,
#'   arg_nbr_obs = 5
#' )
#'
#' # Example 9: Analyze adverse events by treatment group
#' # Useful for safety summaries in TLR documents
#' fct_var_text_group(
#'   arg_text_desc = "The most frequently reported adverse events were:",
#'   arg_dataset = tlr_adae,
#'   arg_fl_rest = AESEV == "SEVERE",
#'   arg_var_rest = AEDECOD,
#'   arg_group_var = TRTA,
#'   arg_nbr_obs = 5
#' )
#'
#' @seealso
#' \code{\link{textReplace}} for inserting generated text into Word documents
#' \code{\link{run_apdx}} for adding appendices to TLR documents
#' 
#' @importFrom stringr str_replace str_squish
#' @export
fct_var_text_group <- function( arg_text_desc = "", arg_dataset, arg_fl_rest = NULL, arg_var_rest, arg_group_var,  arg_group_vals = NULL,  arg_group_lbl = NULL, arg_order = TRUE, arg_cvt_stg = I, arg_nbr_obs = Inf ) {
  
  cond_fl <- rlang::enquo(arg_fl_rest)

  data_fl <- arg_dataset %>%
    dplyr::filter( if( !rlang::quo_is_null( cond_fl )) {{ cond_fl }} else TRUE ) %>%
    dplyr::select({{ arg_var_rest }}, {{ arg_group_var }}) %>%
    dplyr::mutate( .group_chr = as.character({{ arg_group_var }}))

  if( !is.null( arg_group_vals )) data_fl <- data_fl %>% dplyr::filter( .group_chr %in% arg_group_vals )

  total_n <- nrow( data_fl )

  tab <- data_fl %>%
    dplyr::count( {{ arg_var_rest }}, .group_chr, name = "n" ) %>%
    dplyr::group_by( {{ arg_var_rest }}) %>%
    dplyr::mutate( overall_n = sum( n ), 
            overall_pct = round( overall_n[ 1 ] / total_n * 100, 1 ), 
            pct = round( n / overall_n[ 1 ] * 100, 1 )
    ) %>%
    dplyr::ungroup( )


  if( is.null( arg_group_vals ) ) arg_group_vals <- unique( tab$.group_chr )
  if( is.null( arg_group_lbl )) arg_group_lbl <- arg_group_vals
  lut <- setNames( arg_group_lbl, arg_group_vals )

  tab <- tab %>%
    dplyr::mutate( .group_chr = factor( .group_chr, levels = arg_group_vals ),
            .group_lbl = lut[ as.character( .group_chr )] )

  overall_tab <- tab %>%
    dplyr::arrange( {{ arg_var_rest }}, .group_chr ) %>% 
    dplyr::group_by( {{ arg_var_rest }} ) %>%
    dplyr::summarise( overall_n = dplyr::first( overall_n ),
               overall_pct = dplyr::first( overall_pct ),
               var_lbl = arg_cvt_stg( dplyr::first( {{ arg_var_rest }} )),
               group_txt = { gp <- paste0( pct, "% on ", .group_lbl ) 
                             txt <- paste( gp, collapse = ", ") 
                             stringr::str_replace( txt, ",(?!.*,)", " and" ) }, 
              .groups = "drop" ) %>%
    dplyr::arrange( if( arg_order ) dplyr::desc( overall_n ) else TRUE )  %>%
    dplyr::slice_head(n = arg_nbr_obs) %>%
    dplyr::mutate( lbl = paste0( overall_pct, "% ", var_lbl, " (", group_txt, ")") )

  paste( arg_text_desc, overall_tab$lbl %>% 
                          paste(collapse = ", " ) %>%
                          stringr::str_replace(",(?!.*,)", " and"),
          sep = " " ) %>%
  stringr::str_squish()
}