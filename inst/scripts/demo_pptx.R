library( officer )
library( RapTLR )

# Optional — install for full table rendering in slides:
# install.packages("flextable")

# ── Paths to built-in TLF outputs ─────────────────────────────────────────────
path_outputs <- system.file( "extdata/TLF_outputs", package = "RapTLR" )


# ── Example 1: All outputs, no sections ───────────────────────────────────────
# All .docx and image files found in the folder become slides.

run_pptx(
  path_outputs   = path_outputs,
  title          = "Clinical Study — Top-Line Results",
  subtitle       = paste0( "Generated on ", Sys.Date() ),
  return_to_file = "study_all_outputs.pptx"
)


# ── Example 2: Curated selection with sections (named list) ───────────────────
# Same format as run_apdx():
#   list( "Section Title" = c( "output_stem1", "output_stem2" ), ... )
# A section-divider slide is automatically added before each section.

run_pptx(
  path_outputs = path_outputs,
  sections_structure = list(
    Demographics = "tsidem03",
    Safety       = c( "tsfae10", "lsfae03", "tefmad01a" )
  ),
  title          = "Clinical Study ABC-123",
  subtitle       = "Top-Line Results",
  return_to_file = "study_with_sections.pptx"
)


# ── Example 3: Using a CSV structure file ─────────────────────────────────────
# Two-column CSV with headers "Section" and "Outputs".
# Same format as run_apdx().

TLF_list_csv <- system.file( "extdata", "TLF_list.csv", package = "RapTLR" )

run_pptx(
  path_outputs       = path_outputs,
  sections_structure = TLF_list_csv,
  return_to_file     = "study_from_csv.pptx"
)


# ── Example 4: Using an Excel structure file ──────────────────────────────────
TLF_list_xlsx <- system.file( "extdata", "TLF_list.xlsx", package = "RapTLR" )

run_pptx(
  path_outputs       = path_outputs,
  sections_structure = TLF_list_xlsx,
  return_to_file     = "study_from_xlsx.pptx"
)


# ── Example 5: With a branded .pptx template ──────────────────────────────────
# Pass a path to your company's PowerPoint template.
# Slide layouts and master styles are inherited from the template.

# run_pptx(
#   path_outputs       = path_outputs,
#   sections_structure = TLF_list_xlsx,
#   pptx_template      = "path/to/company_template.pptx",
#   return_to_file     = "branded_results.pptx"
# )


# ── Example 6: Mixing tables (.docx) and figures (images) ─────────────────────
# Put your image files (PNG, JPG, SVG, ...) alongside the .docx files.
# run_pptx() auto-detects the file type:
#   .docx  → table slide (title extracted automatically)
#   image  → figure slide (image embedded full-slide)
#
# run_pptx(
#   path_outputs       = "path/to/mixed_outputs/",
#   sections_structure = list(
#     Efficacy = c( "km_curve.png", "tefmad01a" ),
#     Safety   = c( "ae_barplot.png", "tsfae10" )
#   ),
#   return_to_file = "mixed_outputs.pptx"
# )
