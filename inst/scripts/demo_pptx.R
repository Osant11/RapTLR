library( officer )
library( RapTLR )

# ── Load example clinical data ─────────────────────────────────────────────────
data( tlr_adsl )
data( tlr_adae )


# ── Generate narrative texts with fct_* functions ──────────────────────────────

## General enrolment information
text_enrolment <- fct_smpl_rest(
  arg_dataset   = tlr_adsl,
  arg_text_desc = "This double-blind, placebo-controlled study randomized ",
  arg_text_end  = " participants to"
) %c%
  fct_var_text(
    arg_dataset  = tlr_adsl,
    arg_var_rest = TRT01A,
    arg_lbl      = "nbr-pct"
  )

## Demographics: sex and age
text_demographics <- fct_smpl_rest(
  arg_dataset   = tlr_adsl,
  arg_fl_rest   = SEX == "F",
  arg_text_desc = "The study population included ",
  arg_text_end  = " female participants and most participants were"
) %c%
  fct_var_text(
    arg_dataset  = tlr_adsl,
    arg_var_rest = RACE,
    arg_order    = TRUE,
    arg_cvt_stg  = stringr::str_to_title,
    arg_lbl      = "pct",
    arg_nbr_obs  = 2
  ) %c.%
  fct_bsl_stats(
    arg_dataset = tlr_adsl,
    arg_grp     = ARM,
    arg_var_num = AGE,
    arg_var1    = "median",
    arg_var2    = "range",
    arg_unit    = "years",
    arg_label   = "age"
  )

## Most common adverse events (>= 15% on high dose arm)
text_aes <- fct_mst_aes(
  arg_data      = tlr_adae,
  arg_data_adsl = tlr_adsl,
  arg_grp       = ARM,
  arg_var       = AEBODSYS,
  arg_trt_flt   = `Xanomeline High Dose`,
  arg_frq       = 15
)

## Study disposition / completers
text_completers <- fct_smpl_rest(
  arg_dataset   = tlr_adsl,
  arg_fl_rest   = COMP24FL == "Y",
  arg_text_desc = "At week 24, the number of completers were: "
) %c.%
  fct_smpl_rest(
    arg_dataset   = tlr_adsl,
    arg_fl_rest   = DISCONFL == "Y",
    arg_text_desc = "Overall, ",
    arg_text_end  = " participants discontinued the study."
  )


# ── Define the slide structure ─────────────────────────────────────────────────
# Each row = one slide.
# 'Content' can be:
#   - A plain text string
#   - A keyword (e.g. "TT_demo_TT") resolved from the 'replacements' list
#   - A path to a PNG/JPG image file (auto-detected as type "image")

slides <- data.frame(
  Title   = c(
    "Study Overview",
    "Demographics & Baseline",
    "Most Common Adverse Events",
    "Study Disposition"
  ),
  Content = c(
    "TT_enrolment_TT",
    "TT_demographics_TT",
    "TT_aes_TT",
    "TT_completers_TT"
  ),
  stringsAsFactors = FALSE
)


# ── Generate the PowerPoint ────────────────────────────────────────────────────
run_pptx(
  slides_structure = slides,
  replacements = list(
    TT_enrolment_TT   = text_enrolment,
    TT_demographics_TT = text_demographics,
    TT_aes_TT         = text_aes,
    TT_completers_TT  = text_completers
  ),
  title          = "Clinical Study — Top-Line Results",
  subtitle       = paste0( "Generated on ", Sys.Date() ),
  return_to_file = "study_toplevel.pptx"
)


# ── Optional: include an image slide ──────────────────────────────────────────
# If you have a figure saved as a PNG, add it as a row with type "image":
#
# slides_with_fig <- rbind(
#   slides,
#   data.frame(
#     Title   = "AE Overview Figure",
#     Content = "path/to/ae_figure.png",
#     stringsAsFactors = FALSE
#   )
# )
#
# run_pptx(
#   slides_structure = slides_with_fig,
#   replacements     = list( ... ),
#   return_to_file   = "study_with_figure.pptx"
# )


# ── Optional: use a CSV/Excel structure file ───────────────────────────────────
# Instead of a data frame, you can pass a path to a CSV or XLSX file.
# Required columns: Title, Content  (optional: Type)
#
# run_pptx(
#   slides_structure = "path/to/slides_plan.csv",
#   replacements     = list( ... ),
#   pptx_template    = "path/to/company_template.pptx",
#   return_to_file   = "branded_results.pptx"
# )
