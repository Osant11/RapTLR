#' Generate a PowerPoint presentation from TLF outputs
#'
#' @description
#' Creates a PowerPoint presentation (`.pptx`) from a folder of clinical study
#' outputs (Tables, Listings, and Figures). Follows the same API as
#' [run_apdx()]: accepts the same `sections_structure` formats and discovers
#' outputs from `path_outputs`.
#'
#' Each output becomes one slide:
#' \itemize{
#'   \item **Image files** (`.png`, `.jpg`, `.jpeg`, `.svg`, `.bmp`, `.tiff`)
#'     — embedded as full-slide figures.
#'   \item **Word documents** (`.docx`) — the table title is extracted
#'     automatically (same logic as [run_apdx()]) and the table content is
#'     rendered as a native PowerPoint table. Requires the `flextable` package
#'     for table rendering; a placeholder message is shown if it is not
#'     installed.
#' }
#'
#' When sections are provided a section-divider slide is inserted before each
#' group of outputs.
#'
#' @param path_outputs Path to the folder containing TLF output files
#'   (`.docx`, `.png`, `.jpg`, etc.).
#' @param sections_structure Optional. Controls which outputs to include and
#'   how to group them into sections. Accepts:
#'   \describe{
#'     \item{`NULL`}{All supported files found in `path_outputs` are included
#'       in a single group with no section dividers.}
#'     \item{Named list}{Each name is a section title; values are character
#'       vectors of output file stems (without extension). Same format as
#'       [run_apdx()].}
#'     \item{CSV / XLSX path}{A two-column file with columns \code{Section}
#'       and \code{Outputs}. Same format as [run_apdx()].}
#'   }
#' @param pptx_template Path to a `.pptx` template file for slide layouts and
#'   styling. When `NULL` (default) the built-in Office Theme is used.
#' @param title Title for an optional opening title slide. When `NULL`
#'   (default) no title slide is added.
#' @param subtitle Subtitle for the opening title slide. Only used when
#'   `title` is not `NULL`. Default: `NULL`.
#' @param return_to_file Path (with or without `.pptx` extension) where the
#'   presentation is saved. When `NULL` (default) the officer pptx object is
#'   returned invisibly so it can be piped into further officer calls.
#'
#' @return When `return_to_file` is provided, saves the presentation and
#'   returns the file path invisibly. Otherwise returns the officer pptx
#'   object invisibly.
#'
#' @examples
#' \dontrun{
#' path_outputs <- system.file("extdata/TLF_outputs", package = "RapTLR")
#'
#' # ── All outputs, no sections ─────────────────────────────────────────────
#' run_pptx(path_outputs  = path_outputs,
#'          title          = "Clinical Study — Top-Line Results",
#'          return_to_file = "study_results.pptx")
#'
#' # ── With sections (named list) ───────────────────────────────────────────
#' run_pptx(
#'   path_outputs       = path_outputs,
#'   sections_structure = list(
#'     Demographics = "tsidem03",
#'     Safety       = c("tsfae10", "lsfae03")
#'   ),
#'   title          = "Clinical Study ABC-123",
#'   subtitle       = "Top-Line Results",
#'   return_to_file = "study_sections.pptx"
#' )
#'
#' # ── With a CSV structure file (same format as run_apdx) ──────────────────
#' TLF_list_csv <- system.file("extdata", "TLF_list.csv", package = "RapTLR")
#' run_pptx(
#'   path_outputs       = path_outputs,
#'   sections_structure = TLF_list_csv,
#'   return_to_file     = "study_csv.pptx"
#' )
#'
#' # ── With a branded .pptx template ────────────────────────────────────────
#' run_pptx(
#'   path_outputs       = path_outputs,
#'   sections_structure = TLF_list_csv,
#'   pptx_template      = "company_template.pptx",
#'   return_to_file     = "branded_results.pptx"
#' )
#' }
#'
#' @importFrom readxl read_excel
#' @export
run_pptx <- function(path_outputs,
                     sections_structure = NULL,
                     pptx_template      = NULL,
                     title              = NULL,
                     subtitle           = NULL,
                     return_to_file     = NULL) {

  img_ext <- c("png", "jpg", "jpeg", "svg", "bmp", "tiff", "tif")

  # ── Input validation ───────────────────────────────────────────────────────

  if (missing(path_outputs) || is.null(path_outputs))
    stop("Error: 'path_outputs' cannot be NULL. Please specify a valid folder path.")

  if (!dir.exists(path_outputs))
    stop("Error: 'path_outputs' folder not found: ", path_outputs)

  supported_files <- list.files(path_outputs)[
    grepl(paste0("\\.(", paste(c("docx", img_ext), collapse = "|"), ")$"),
          list.files(path_outputs), ignore.case = TRUE)
  ]

  if (length(supported_files) == 0)
    stop("Error: No supported output files (.docx, .png, .jpg, etc.) found in ",
         "'path_outputs'.")

  if (!is.null(pptx_template)) {
    if (!file.exists(pptx_template))
      stop("Error: 'pptx_template' file not found: ", pptx_template)
    if (!grepl("\\.pptx$", pptx_template, ignore.case = TRUE))
      stop("Error: 'pptx_template' must be a .pptx file.")
  }

  if (!is.null(return_to_file)) {
    return_to_file <- file.path(getwd(), return_to_file)
    if (!dir.exists(dirname(return_to_file)))
      stop("Error: Directory for 'return_to_file' does not exist.")
    if (!grepl("\\.pptx$", return_to_file, ignore.case = TRUE))
      return_to_file <- paste0(return_to_file, ".pptx")
  } else {
    return_to_file <- file.path(getwd(), "presentation_processed.pptx")
  }

  # ── Load sections_structure (same logic as run_apdx) ──────────────────────

  has_explicit_sections <- !is.null(sections_structure)

  if (has_explicit_sections) {

    if (is.character(sections_structure)) {

      if (!file.exists(sections_structure))
        stop("Error: 'sections_structure' file not found: ", sections_structure)

      df_ext       <- tools::file_ext(sections_structure)
      structure_df <- switch(df_ext,
        csv  = utils::read.csv(sections_structure, stringsAsFactors = FALSE),
        xls  = readxl::read_excel(sections_structure, col_names = TRUE),
        xlsx = readxl::read_excel(sections_structure, col_names = TRUE),
        stop("Error: Unsupported file type. Use .csv or .xlsx.")
      )

      if (ncol(structure_df) != 2)
        stop("Error: 'sections_structure' file must have exactly two columns: ",
             "Section and Outputs.")

      structure_df <- structure_df %>%
        dplyr::rename(Section = 1, Outputs = 2)

      output_structure_list <- split(
        structure_df$Outputs,
        factor(structure_df$Section, levels = unique(structure_df$Section)),
        drop = FALSE
      )

    } else if (is.list(sections_structure)) {
      output_structure_list <- sections_structure

    } else {
      stop("Error: 'sections_structure' must be a named list or a path to a ",
           ".csv or .xlsx file.")
    }

  } else {
    # NULL → include all supported files in one unnamed group
    all_stems             <- tools::file_path_sans_ext(supported_files)
    output_structure_list <- list(Outputs = all_stems)
  }

  # ── Initialise pptx ────────────────────────────────────────────────────────

  pptx <- if (!is.null(pptx_template)) {
    officer::read_pptx(pptx_template)
  } else {
    officer::read_pptx()
  }

  available_layouts <- officer::layout_summary(pptx)

  # ── Internal helpers ───────────────────────────────────────────────────────

  master_for <- function(layout_name) {
    available_layouts$master[available_layouts$layout == layout_name][1]
  }

  pick_layout <- function(...) {
    for (candidate in c(...)) {
      if (candidate %in% available_layouts$layout) return(candidate)
    }
    available_layouts$layout[1]
  }

  ph_types_of <- function(layout_name) {
    officer::layout_properties(pptx, layout = layout_name,
                               master = master_for(layout_name))$type
  }

  # Find output file: try with extension as-is, then try each supported ext
  resolve_path <- function(stem) {
    if (tools::file_ext(stem) != "") {
      p <- file.path(path_outputs, stem)
      if (file.exists(p)) return(p)
    }
    for (ext in c("docx", img_ext)) {
      p <- file.path(path_outputs, paste0(stem, ".", ext))
      if (file.exists(p)) return(p)
    }
    NULL
  }

  is_image <- function(path) {
    grepl(paste0("\\.(", paste(img_ext, collapse = "|"), ")$"),
          path, ignore.case = TRUE)
  }

  # Extract human-readable title from a TLF docx — mirrors addLink.R logic
  extract_tlf_title <- function(docx_path) {
    tryCatch({
      tbi <- officer::docx_summary(officer::read_docx(docx_path))

      if (sum(tbi$is_header, na.rm = TRUE) > 0) {
        tbj <- tbi[which(tbi$is_header), ]
        tbj <- tbj[order(tbj$row_id), ]
      } else {
        tbj <- tbi[!is.na(tbi$row_id) & tbi$row_id == 1, ]
      }

      if (nrow(tbj) > 0) {
        full_title <- tbj$text[1]
        colon_pos  <- regexpr(":", full_title)
        if (colon_pos > 0)
          return(trimws(substr(full_title, colon_pos + 1, nchar(full_title))))
        return(full_title)
      }
      tools::file_path_sans_ext(basename(docx_path))
    }, error = function(e) tools::file_path_sans_ext(basename(docx_path)))
  }

  # Extract first table from a docx as a data frame (base R only)
  extract_first_table_df <- function(docx_path) {
    tryCatch({
      tbi   <- officer::docx_summary(officer::read_docx(docx_path))
      cells <- tbi[tbi$content_type == "table cell", ]

      if (nrow(cells) == 0) return(NULL)

      first_idx <- min(cells$doc_index)
      cells     <- cells[cells$doc_index == first_idx,
                         c("row_id", "cell_id", "text", "is_header")]

      # Drop duplicate cells (e.g. from spanning)
      cells <- cells[!duplicated(cells[, c("row_id", "cell_id")]), ]

      # Reshape to wide format (base R — no tidyr needed)
      tbl <- reshape(cells[, c("row_id", "cell_id", "text")],
                     idvar    = "row_id",
                     timevar  = "cell_id",
                     direction = "wide")
      tbl <- tbl[order(tbl[["row_id"]]), ]

      # Derive column names from header rows (or first data row as fallback)
      header_row_ids <- unique(cells$row_id[!is.na(cells$is_header) &
                                              cells$is_header == TRUE])

      if (length(header_row_ids) > 0) {
        hdr_id    <- min(header_row_ids)
        hdr_cells <- cells[cells$row_id == hdr_id, ]
        hdr_cells <- hdr_cells[order(hdr_cells$cell_id), ]
        col_names <- hdr_cells$text
        tbl       <- tbl[!tbl[["row_id"]] %in% header_row_ids, , drop = FALSE]
      } else {
        first_rid <- tbl[1, "row_id"]
        fr_cells  <- cells[cells$row_id == first_rid, ]
        fr_cells  <- fr_cells[order(fr_cells$cell_id), ]
        col_names <- fr_cells$text
        tbl       <- tbl[-1, , drop = FALSE]
      }

      tbl[["row_id"]] <- NULL
      names(tbl)      <- make.names(col_names, unique = TRUE)
      rownames(tbl)   <- NULL

      as.data.frame(tbl, stringsAsFactors = FALSE)
    }, error = function(e) NULL)
  }

  # ── Opening title slide ────────────────────────────────────────────────────

  if (!is.null(title)) {
    lo_title <- pick_layout("Title Slide", "Title, Content", "Blank")
    ph_title <- ph_types_of(lo_title)

    pptx <- officer::add_slide(pptx, layout = lo_title,
                               master = master_for(lo_title))

    title_ph <- if ("ctrTitle" %in% ph_title) "ctrTitle" else "title"
    pptx <- officer::ph_with(pptx, value = title,
                             location = officer::ph_location_type(title_ph))

    if (!is.null(subtitle) && "subTitle" %in% ph_title)
      pptx <- officer::ph_with(pptx, value = subtitle,
                               location = officer::ph_location_type("subTitle"))
  }

  # ── Content slides ─────────────────────────────────────────────────────────

  lo_content   <- pick_layout("Title and Content", "Title, Content", "Blank")
  ph_content   <- ph_types_of(lo_content)
  lo_section   <- pick_layout("Section Header", "Title Slide", lo_content)
  ph_section   <- ph_types_of(lo_section)

  for (section_name in names(output_structure_list)) {

    outputs_in_section <- output_structure_list[[section_name]]

    # Section divider slide (only when caller supplied sections_structure)
    if (has_explicit_sections) {
      pptx <- officer::add_slide(pptx, layout = lo_section,
                                 master = master_for(lo_section))

      sec_ph <- if ("title" %in% ph_section) "title" else "ctrTitle"
      pptx   <- officer::ph_with(pptx, value = section_name,
                                 location = officer::ph_location_type(sec_ph))
    }

    for (stem in outputs_in_section) {

      file_path <- resolve_path(stem)

      if (is.null(file_path)) {
        warning("Output file not found for '", stem, "'. Skipping.")
        next
      }

      pptx <- officer::add_slide(pptx, layout = lo_content,
                                 master = master_for(lo_content))

      if (is_image(file_path)) {
        # ── Figure slide ───────────────────────────────────────────────────

        slide_title <- tools::file_path_sans_ext(basename(file_path))

        if ("title" %in% ph_content)
          pptx <- officer::ph_with(
            pptx, value = slide_title,
            location = officer::ph_location_type("title")
          )

        img  <- officer::external_img(file_path, width = 8.5, height = 4.8)
        pptx <- officer::ph_with(
          pptx, value = img,
          location = officer::ph_location(left = 0.5, top = 1.4,
                                          width = 9,   height = 5.2)
        )

      } else if (grepl("\\.docx$", file_path, ignore.case = TRUE)) {
        # ── Table slide ────────────────────────────────────────────────────

        slide_title <- extract_tlf_title(file_path)

        if ("title" %in% ph_content)
          pptx <- officer::ph_with(
            pptx, value = slide_title,
            location = officer::ph_location_type("title")
          )

        if (requireNamespace("flextable", quietly = TRUE)) {
          tbl_df <- extract_first_table_df(file_path)

          if (!is.null(tbl_df) && nrow(tbl_df) > 0) {
            ft <- flextable::flextable(tbl_df)
            ft <- flextable::fontsize(ft, size = 9, part = "all")
            ft <- flextable::autofit(ft)
            pptx <- officer::ph_with(
              pptx, value = ft,
              location = officer::ph_location(left = 0.3, top = 1.4,
                                              width = 9.4, height = 5.5)
            )
          } else {
            warning("Could not extract table content from '",
                    basename(file_path), "'.")
          }

        } else {
          # flextable not available — show a placeholder
          body_text <- paste0(
            "Source: ", tools::file_path_sans_ext(basename(file_path)),
            "\n\nInstall the 'flextable' package for full table rendering."
          )
          if ("body" %in% ph_content) {
            pptx <- officer::ph_with(
              pptx, value = body_text,
              location = officer::ph_location_type("body")
            )
          }
        }

      } else {
        warning("Unsupported file type for '", basename(file_path),
                "'. Skipping.")
      }
    }
  }

  # ── Output ─────────────────────────────────────────────────────────────────

  print(pptx, target = return_to_file)
  message("PowerPoint saved to: ", return_to_file)
  invisible(return_to_file)
}
