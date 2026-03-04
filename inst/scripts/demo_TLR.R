library( rlang )
library( dplyr )
library( officer )
library( crosstable )
library( stringr )
library( xml2 )
library( RapTLR )

# Creating the Appendice 
path_TLFs <- system.file( "extdata/TLF_outputs", package = "RapTLR" )
path_docx <- system.file( "extdata/TLR_Shell.docx", package = "RapTLR" )
path_result <- system.file( "docx_with_appendice", package = "RapTLR" )
TLF_list_xlsx <- system.file( "extdata/TLF_list.xlsx", package = "RapTLR" )

run_apdx( path_TLFs = path_TLFs,
          docx_object = path_docx,
          sections_structure = TLF_list_xlsx,
          return_to_file = "TLR_with_appendice.docx" )


# Loading data 
# Subject Level 
data( tlr_adsl )
# Adverse Event
data( tlr_adae )

# Create list for placeholders
PLHD <- list( )

# TLR creation

## Importing TLR in R
TLR <- officer::read_docx( "TLR_with_appendice.docx" )

## Disposition

### General information / Country 

PLHD$TT_geninfo_TT <- fct_smpl_rest( arg_dataset = tlr_adsl, arg_text_desc = "This double-blind, placebo-controlled study randomized ", arg_text_end = " participants to" ) %c%
                      fct_var_text( arg_dataset = tlr_adsl, arg_var_rest = TRT01A, arg_lbl = "nbr-pct" ) %c%
                      fct_smpl_rest( arg_dataset = tlr_adsl, arg_fl_rest = !duplicated( SITEID ), arg_text_desc = "across ", arg_text_end = " sites" ) %c.%
                      fct_var_text( arg_dataset = tlr_adsl, arg_var_rest = SITEID, arg_lbl = "pct", arg_order = TRUE, arg_nbr_obs = 5, arg_text_desc = "The most represented sites were:" ) %c.%
                      fct_var_text( arg_dataset = tlr_adsl, arg_var_rest = RACE, arg_order = TRUE, arg_cvt_stg = stringr::str_to_title, arg_lbl = "pct", arg_nbr_obs = 2, arg_text_desc = "Most were" ) %c.%
                      c( "Table \\@ref(TSIDEM03)" )

PLHD$TT_geninfo_TT

textReplace( TLR, "TT_geninfo_TT", PLHD$TT_geninfo_TT )


### Study disp 
PLHD$TT_studips_TT <- fct_smpl_rest( arg_dataset = tlr_adsl, arg_fl_rest = DISCONFL == "Y", arg_text_desc = "There were ", arg_text_end = " participants who discontinued the study, the main reason for treatment discontinuation were" ) %c%
                      var_text_group( arg_dataset = tlr_adsl, arg_fl_rest = DISCONFL == "Y", arg_group_var = TRT01A, arg_var_rest = DCREASCD, arg_nbr_obs = 3, arg_order = TRUE,  ) %c.%
                      c( "Table \\@ref(TSIDEM03)" )

PLHD$TT_studips_TT
textReplace( TLR, "TT_studips_TT", PLHD$TT_studips_TT )



### Number of Completers ----
PLHD$TT_wkcmplt_TT <- fct_smpl_rest( arg_dataset = tlr_adsl, arg_fl_rest = COMP8FL == "Y", arg_text_desc = "At week 8, the number of completers were: " ) %c.% 
                      fct_smpl_rest( arg_dataset = tlr_adsl, arg_fl_rest = COMP16FL == "Y", arg_text_desc = "At week 16, the number of completers were: " ) %c.%
                      fct_smpl_rest( arg_dataset = tlr_adsl, arg_fl_rest = COMP24FL == "Y", arg_text_desc = "At week 24, the number of completers were: " )
PLHD$TT_wkcmplt_TT
textReplace( TLR, "TT_wkcmplt_TT", PLHD$TT_wkcmplt_TT )



## Demographics and baseline characteristics
### Sex / Age
PLHD$TT_sexage_TT <- fct_smpl_rest( arg_dataset = tlr_adsl, arg_fl_rest = SEX == "F", arg_text_desc = "The study population included ", arg_text_end = " female participants and most participants were" ) %c%
                     fct_var_text( arg_dataset = tlr_adsl, arg_var_rest = RACE, arg_order = TRUE, arg_cvt_stg = stringr::str_to_title, arg_lbl = "pct", arg_nbr_obs = 2 ) %c.%
                     fct_bsl_stats( arg_dataset = tlr_adsl, arg_grp = ARM, arg_var_num = AGE, arg_var1 = "median", arg_var2 = "range", arg_unit = "years", arg_label = "age" ) %c&%
                     fct_bsl_stats( arg_dataset = tlr_adsl, arg_grp = ARM, arg_trt_inc = c( "Xanomeline High Dose", "Xanomeline Low Dose", "Placebo" ), arg_var_num = BMIBL, 
                                    arg_var1 = "mean", arg_var2 = "sd", arg_unit = "kg/m²", arg_label = "baseline BMI", arg_cvt_stg = stringr::str_to_lower ) %c.%
                     c( "Table \\@ref(TSIDEM03)" )

PLHD$TT_sexage_TT 
textReplace( TLR, "TT_sexage_TT", PLHD$TT_sexage_TT )


## Efficacy ----
PLHD$TT_effsum_TT <- c( "On Day 28, the mean (SD) for treatment dummy A was 26.7 (10.35) and for treatment dummy B was 27.1 (10.42).
                        These corresponded to changes from baseline of -9.7 (11.07) and -8.9 (10.22) for treatment A and B, respectively.
                        The least squares (LS) mean difference between treatment dummy A minus B, estimated as -0.6 (1.38) using the MMRM model,
                        was not statistically significant (p = 0.674), Table \\@ref(TEFMAD01A)" )

PLHD$TT_effsum_TT 
textReplace( TLR, "TT_effsum_TT", PLHD$TT_effsum_TT )  
  
## Safety ----

### Most frequent AES ----
PLHD$TT_mosttaes_TT <- fct_mst_aes( arg_data = tlr_adae, arg_data_adsl = tlr_adsl, arg_grp = ARM, arg_var = AEBODSYS, arg_trt_flt = `Xanomeline High Dose`, arg_frq = 15 )
PLHD$TT_mosttaes_TT

textReplace( TLR, "TT_mosttaes_TT", PLHD$TT_mosttaes_TT ) 


### Serious Aes info table ----
PLHD$TT_aesser_TT <- c( "Table \\@ref(TSFAE10) presents an overview of Treatment-emergent Serious Adverse Events" )
textReplace( TLR, "TT_aesser_TT", PLHD$TT_aesser_TT ) 

outTLR1 <- cursor_reach( TLR, "TT_tblTSFAE_TT" ) %>% 
  body_add_docx( file.path( path_TLFs, "tsfae10_c.docx" ), pos = "on" ) 

### Final PRINT ----
print( TLR, 
       "./tlr_final.docx" )



##########################################################
##########################################################
##########################################################
