#' Generate a PowerPoint presentation from clinical study outputs
#'
#' @description
#' Creates a PowerPoint presentation (`.pptx`) from a structured list of
#' slides. Each slide can contain:
#'
#' - **Text** — narrative strings produced by `fct_bsl_stats()`,
#'   `fct_mst_aes()`, `fct_var_text()`, and similar functions.
#' - **Images** — `.png`, `.jpg`, or `.jpeg` figure files.
#'
#' Keywords in the `Content` column (e.g. `TT_baseline_TT`) are resolved
#' against the `replacements` list before rendering — the same pattern used
#' in [textReplace()].
#'
#' @param slides_structure A data frame, or a path to a CSV or Excel (`.xlsx`)
#'   file specifying the presentation structure. Must contain the following
#'   columns:
#'   \describe{
#'     \item{`Title`}{Title displayed at the top of each slide.}
#'     \item{`Content`}{Text to display, a file path to an image
#'       (`.png`, `.jpg`, `.jpeg`), or a keyword that matches a name in
#'       `replacements`.}
#'     \item{`Type`}{*(Optional)* Force the content type: `"text"` or
#'       `"image"`. Auto-detected from `Content` when omitted or `NA`.}
#'   }
#' @param replacements A named list of character strings. Keywords in the
#'   `Content` column that exactly match a name in this list are replaced with
#'   the corresponding value before rendering. Designed to receive outputs from
#'   `fct_bsl_stats()`, `fct_mst_aes()`, `fct_var_text()`, and similar
#'   functions. Default: `NULL`.
#' @param pptx_template Path to a `.pptx` file to use as the design template.
#'   Slide layouts and master styles are inherited from this file. When `NULL`
#'   (default) a blank presentation with the built-in Office Theme is created.
#' @param title Title displayed on an opening title slide. When `NULL`
#'   (default) no title slide is added.
#' @param subtitle Subtitle for the opening title slide. Only used when
#'   `title` is not `NULL`. Default: `NULL`.
#' @param return_to_file Path (with or without `.pptx` extension) where the
#'   file is saved. When `NULL` (default) the officer pptx object is returned
#'   invisibly instead of being written to disk.
#'
#' @return When `return_to_file` is provided, saves the presentation and
#'   returns the file path invisibly. Otherwise returns the officer pptx
#'   object invisibly so it can be piped into further officer calls.
#'
#' @examples
#' \dontrun{
#' library(RapTLR)
#' data(tlr_adsl)
#' data(tlr_adae)
#'
#' # ── 1. Generate narrative texts with existing fct_* functions ──────────────
#' text_age <- fct_bsl_stats(
#'   arg_dataset = tlr_adsl,
#'   arg_grp     = "TRT01A",
#'   arg_var_num = "AGE",
#'   arg_label   = "Age (years)"
#' )
#'
#' text_ae <- fct_mst_aes(
#'   arg_data      = tlr_adae,
#'   arg_data_adsl = tlr_adsl,
#'   arg_grp       = "TRT01A",
#'   arg_var       = "AEDECOD",
#'   arg_trt_flt   = "Xanomeline High Dose"
#' )
#'
#' # ── 2. Define the slide structure as a data frame ──────────────────────────
#' slides <- data.frame(
#'   Title   = c(
#'     "Study Overview",
#'     "Baseline Characteristics",
#'     "Most Common Adverse Events"
#'   ),
#'   Content = c(
#'     "This study enrolled subjects across three treatment arms.",
#'     "TT_age_TT",
#'     "TT_ae_TT"
#'   ),
#'   stringsAsFactors = FALSE
#' )
#'
#' # ── 3. Generate the PowerPoint ─────────────────────────────────────────────
#' run_pptx(
#'   slides_structure = slides,
#'   replacements     = list(TT_age_TT = text_age, TT_ae_TT = text_ae),
#'   title            = "Clinical Study ABC-123 — Top-Line Results",
#'   subtitle         = "Data cut-off: 01 Jan 2025",
#'   return_to_file   = "study_results.pptx"
#' )
#'
#' # ── 4. With a CSV structure file and a branded template ────────────────────
#' run_pptx(
#'   slides_structure = "slides_plan.csv",
#'   replacements     = list(TT_age_TT = text_age),
#'   pptx_template    = "company_template.pptx",
#'   return_to_file   = "branded_results.pptx"
#' )
#'
#' # ── 5. Including a figure image on a slide ─────────────────────────────────
#' slides_with_img <- data.frame(
#'   Title   = c("Key Findings", "AE Overview Figure"),
#'   Content = c("TT_age_TT", "path/to/ae_figure.png"),
#'   stringsAsFactors = FALSE
#' )
#'
#' run_pptx(
#'   slides_structure = slides_with_img,
#'   replacements     = list(TT_age_TT = text_age),
#'   return_to_file   = "results_with_figure.pptx"
#' )
#' }
#'
#' @export
run_pptx <- function(slides_structure,
                     replacements   = NULL,
                     pptx_template  = NULL,
                     title          = NULL,
                     subtitle       = NULL,
                     return_to_file = NULL) {

  # ── Input validation ───────────────────────────────────────────────────────

  if (missing(slides_structure)) {
    stop("Error: 'slides_structure' is required. ",
         "Provide a data frame or a path to a CSV/XLSX file.")
  }

  # ── Load slides structure ──────────────────────────────────────────────────

  if (is.character(slides_structure) && length(slides_structure) == 1) {

    if (!file.exists(slides_structure))
      stop("Error: 'slides_structure' file not found: ", slides_structure)

    df_ext    <- tools::file_ext(slides_structure)
    slides_df <- switch(df_ext,
      csv  = utils::read.csv(slides_structure, stringsAsFactors = FALSE),
      xls  = readxl::read_excel(slides_structure, col_names = TRUE),
      xlsx = readxl::read_excel(slides_structure, col_names = TRUE),
      stop("Error: 'slides_structure' must be a .csv or .xlsx file.")
    )

  } else if (is.data.frame(slides_structure)) {
    slides_df <- slides_structure

  } else {
    stop("Error: 'slides_structure' must be a data frame or a path to a ",
         ".csv or .xlsx file.")
  }

  slides_df <- as.data.frame(slides_df, stringsAsFactors = FALSE)

  missing_cols <- setdiff(c("Title", "Content"), names(slides_df))
  if (length(missing_cols) > 0)
    stop("Error: Missing required column(s) in 'slides_structure': ",
         paste(missing_cols, collapse = ", "))

  # ── Template validation ────────────────────────────────────────────────────

  if (!is.null(pptx_template)) {
    if (!file.exists(pptx_template))
      stop("Error: 'pptx_template' file not found: ", pptx_template)
    if (!grepl("\\.pptx$", pptx_template, ignore.case = TRUE))
      stop("Error: 'pptx_template' must be a .pptx file.")
  }

  # ── Output path ────────────────────────────────────────────────────────────

  if (!is.null(return_to_file)) {
    return_to_file <- file.path(getwd(), return_to_file)
    if (!dir.exists(dirname(return_to_file)))
      stop("Error: Directory for 'return_to_file' does not exist.")
    if (!grepl("\\.pptx$", return_to_file, ignore.case = TRUE))
      return_to_file <- paste0(return_to_file, ".pptx")
  }

  # ── Initialise pptx object ─────────────────────────────────────────────────

  pptx <- if (!is.null(pptx_template)) {
    officer::read_pptx(pptx_template)
  } else {
    officer::read_pptx()
  }

  available_layouts <- officer::layout_summary(pptx)

  # ── Internal helpers ───────────────────────────────────────────────────────

  pick_layout <- function(preferred, fallback = "Blank") {
    if (preferred %in% available_layouts$layout) preferred else fallback
  }

  master_for <- function(layout_name) {
    available_layouts$master[available_layouts$layout == layout_name][1]
  }

  layout_ph_types <- function(layout_name) {
    master <- master_for(layout_name)
    officer::layout_properties(pptx, layout = layout_name,
                               master = master)$type
  }

  detect_type <- function(content) {
    if (grepl("\\.(png|jpg|jpeg|svg)$", content, ignore.case = TRUE) &&
        file.exists(content)) {
      return("image")
    }
    "text"
  }

  resolve_content <- function(content) {
    if (!is.null(replacements) && content %in% names(replacements))
      return(as.character(replacements[[content]]))
    content
  }

  # ── Opening title slide ────────────────────────────────────────────────────

  if (!is.null(title)) {

    layout_title <- pick_layout("Title Slide")
    master_title <- master_for(layout_title)
    ph_types_title <- layout_ph_types(layout_title)

    pptx <- officer::add_slide(pptx, layout = layout_title,
                               master = master_title)

    title_ph <- if ("ctrTitle" %in% ph_types_title) "ctrTitle" else "title"
    pptx <- officer::ph_with(
      pptx, value = title,
      location = officer::ph_location_type(title_ph)
    )

    if (!is.null(subtitle) && "subTitle" %in% ph_types_title) {
      pptx <- officer::ph_with(
        pptx, value = subtitle,
        location = officer::ph_location_type("subTitle")
      )
    }
  }

  # ── Content slides ─────────────────────────────────────────────────────────

  layout_content <- pick_layout("Title and Content")
  master_content <- master_for(layout_content)
  ph_types_content <- layout_ph_types(layout_content)

  for (i in seq_len(nrow(slides_df))) {

    slide_title   <- as.character(slides_df$Title[i])
    slide_content <- resolve_content(as.character(slides_df$Content[i]))

    slide_type <- if ("Type" %in% names(slides_df) &&
                      !is.na(slides_df$Type[i]) &&
                      nchar(trimws(slides_df$Type[i])) > 0) {
      trimws(as.character(slides_df$Type[i]))
    } else {
      detect_type(slide_content)
    }

    pptx <- officer::add_slide(pptx, layout = layout_content,
                               master = master_content)

    # Slide title
    if ("title" %in% ph_types_content) {
      pptx <- officer::ph_with(
        pptx, value = slide_title,
        location = officer::ph_location_type("title")
      )
    }

    # Slide body
    if (slide_type == "image") {

      if (!file.exists(slide_content)) {
        warning("Image file not found for slide '", slide_title,
                "': ", slide_content, ". Skipping image content.")
      } else {
        img  <- officer::external_img(slide_content, width = 8, height = 4.5)
        pptx <- officer::ph_with(
          pptx, value = img,
          location = officer::ph_location(left = 0.5, top = 1.5,
                                          width = 9,   height = 5)
        )
      }

    } else {

      # text — use placeholder if available, otherwise free position
      if ("body" %in% ph_types_content) {
        pptx <- officer::ph_with(
          pptx, value = slide_content,
          location = officer::ph_location_type("body")
        )
      } else {
        pptx <- officer::ph_with(
          pptx, value = slide_content,
          location = officer::ph_location(left = 0.5, top = 1.5,
                                          width = 9,   height = 5)
        )
      }
    }
  }

  # ── Output ─────────────────────────────────────────────────────────────────

  if (!is.null(return_to_file)) {
    print(pptx, target = return_to_file)
    message("PowerPoint saved to: ", return_to_file)
    return(invisible(return_to_file))
  }

  invisible(pptx)
}
