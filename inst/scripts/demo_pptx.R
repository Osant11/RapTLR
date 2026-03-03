library( officer )
library( RapTLR )

# ── System requirements for .docx table rendering ─────────────────────────────
#
# run_pptx() converts .docx tables to images using LibreOffice + pdftools:
#
#   1. LibreOffice — must be installed and on the system PATH.
#      Linux:   sudo apt install libreoffice   (Ubuntu/Debian)
#               sudo yum install libreoffice   (RHEL/CentOS)
#      macOS:   brew install --cask libreoffice
#      Windows: https://www.libreoffice.org/download/
#
#   2. pdftools R package (preferred) — converts PDF pages to PNG.
#      install.packages("pdftools")
#      Alternatively: install.packages("magick")  (uses ImageMagick)
#
# If LibreOffice or pdftools is missing, .docx slides are rendered as
# title-only slides with an install hint message. Image files (.png, .jpg,
# etc.) do not require LibreOffice and work out of the box.

# ── Paths to built-in TLF outputs ─────────────────────────────────────────────
path_outputs <- system.file( "extdata/TLF_outputs", package = "RapTLR" )


# ── Example 1: All outputs, no sections ───────────────────────────────────────
# All .docx and image files in the folder become slides, in alphabetical order.

run_pptx(
  path_outputs   = path_outputs,
  title          = "Clinical Study — Top-Line Results",
  subtitle       = paste0( "Generated on ", Sys.Date() ),
  return_to_file = "study_all_outputs.pptx"
)


# ── Example 2: Curated selection with sections (named list) ───────────────────
# One section-divider slide is added automatically before each group.
# Output names are file stems — no extension needed.

run_pptx(
  path_outputs       = path_outputs,
  sections_structure = list(
    Demographics = "tsidem03",
    Safety       = c( "tsfae10", "lsfae03", "tefmad01a" )
  ),
  title          = "Clinical Study ABC-123",
  subtitle       = "Top-Line Results",
  return_to_file = "study_with_sections.pptx"
)


# ── Example 3: Using a CSV structure file ─────────────────────────────────────
# Two-column CSV with headers "Section" and "Outputs" — same format as
# run_apdx().

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


# ── Example 5: Higher resolution images ───────────────────────────────────────
# Default dpi = 150. Increase for sharper slides (larger file size).

run_pptx(
  path_outputs       = path_outputs,
  sections_structure = TLF_list_xlsx,
  dpi                = 250,
  return_to_file     = "study_high_res.pptx"
)


# ── Example 6: With a branded .pptx template ──────────────────────────────────
# Slide layouts and master styles are inherited from your company template.

# run_pptx(
#   path_outputs       = path_outputs,
#   sections_structure = TLF_list_xlsx,
#   pptx_template      = "path/to/company_template.pptx",
#   return_to_file     = "branded_results.pptx"
# )


# ── Example 7: Mixing tables (.docx) and figures (images) ─────────────────────
# Place your PNG/JPG figures alongside the .docx tables in the same folder.
# run_pptx() auto-detects the type by file extension:
#   .docx  → converted via LibreOffice + pdftools (table slide)
#   image  → embedded directly             (figure slide)
#
# run_pptx(
#   path_outputs       = "path/to/mixed_outputs/",
#   sections_structure = list(
#     Efficacy = c( "km_curve.png", "tefmad01a" ),
#     Safety   = c( "ae_barplot.png", "tsfae10" )
#   ),
#   return_to_file = "mixed_outputs.pptx"
# )
