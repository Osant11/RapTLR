#' Generate a PowerPoint presentation from TLF outputs
#'
#' @description
#' Creates a PowerPoint presentation (`.pptx`) from a folder of clinical study
#' outputs (Tables, Listings, and Figures). Follows the same API as
#' [run_apdx()]: accepts the same `sections_structure` formats and discovers
#' outputs from `path_outputs`.
#'
#' Each output becomes one slide (or several slides for multi-page outputs):
#' \itemize{
#'   \item **Image files** (`.png`, `.jpg`, `.jpeg`, `.svg`, `.bmp`, `.tiff`)
#'     — embedded as full-slide figures directly.
#'   \item **Word documents** (`.docx`) — converted to images via LibreOffice
#'     and embedded as slides. Requires **LibreOffice** to be installed on the
#'     system, and either the **pdftools** or **magick** R package for the
#'     PDF-to-image step. Multi-page tables produce one slide per page.
#' }
#'
#' When sections are provided a section-divider slide is inserted before each
#' group of outputs.
#'
#' @section System requirements for `.docx` table rendering:
#' Converting Word tables to slides requires:
#' \enumerate{
#'   \item **LibreOffice** — available as `soffice` or `libreoffice` on the
#'     system `PATH`. Used to convert `.docx` → PDF.
#'   \item **pdftools** R package (preferred) *or* **magick** R package — used
#'     to convert each PDF page to a PNG image.
#' }
#' If LibreOffice is not found, `.docx` outputs are rendered as title-only
#' slides with a placeholder message.
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
#' @param dpi Resolution (dots per inch) used when converting `.docx` tables
#'   to images. Higher values give sharper slides but larger files.
#'   Default: `150`.
#' @param return_to_file Path (with or without `.pptx` extension) where the
#'   presentation is saved. When `NULL` (default) the file is saved as
#'   `presentation_processed.pptx` in the current working directory.
#'
#' @return The file path of the saved presentation, returned invisibly.
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
#' }
#'
#' @importFrom readxl read_excel
#' @export
run_pptx <- function(path_outputs,
                     sections_structure = NULL,
                     pptx_template      = NULL,
                     title              = NULL,
                     subtitle           = NULL,
                     dpi                = 150,
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
    stop("Error: No supported output files (.docx, .png, .jpg, etc.) found ",
         "in 'path_outputs'.")

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

  # ── Check system requirements for .docx conversion ────────────────────────

  lo_path <- Sys.which(c("soffice", "libreoffice"))
  lo_path <- lo_path[nchar(lo_path) > 0]
  has_lo  <- length(lo_path) > 0

  has_pdf_converter <- requireNamespace("pdftools", quietly = TRUE) ||
                       requireNamespace("magick",   quietly = TRUE)

  has_docx_any <- any(grepl("\\.docx$", supported_files, ignore.case = TRUE))

  if (has_docx_any) {
    if (!has_lo) {
      message(
        "[run_pptx] LibreOffice not found on PATH. ",
        ".docx outputs will be rendered as title-only slides.\n",
        "  Install LibreOffice (https://www.libreoffice.org) and ensure ",
        "'soffice' or 'libreoffice' is accessible from R."
      )
    } else if (!has_pdf_converter) {
      message(
        "[run_pptx] Neither 'pdftools' nor 'magick' is installed. ",
        ".docx outputs will be rendered as title-only slides.\n",
        "  Run: install.packages('pdftools')"
      )
    }
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
  slide_w <- officer::slide_size(pptx)$width
  slide_h <- officer::slide_size(pptx)$height

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

  # Find output file: try stem with extension, then try each supported ext
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

  # Trim uniform-colour margins from a PNG in place (uses magick if available).
  trim_png <- function(png_path) {
    if (requireNamespace("magick", quietly = TRUE)) {
      tryCatch({
        img <- magick::image_read(png_path)
        img <- magick::image_trim(img)
        magick::image_write(img, path = png_path, format = "png")
      }, error = function(e) NULL)
    }
  }

  # Return the natural display size of a PNG in inches (pixels / dpi).
  # Returns c(w, h) or NULL when magick is not available.
  png_natural_size <- function(png_path) {
    if (requireNamespace("magick", quietly = TRUE)) {
      tryCatch({
        info <- magick::image_info(magick::image_read(png_path))
        c(w = info$width[1] / dpi, h = info$height[1] / dpi)
      }, error = function(e) NULL)
    } else {
      NULL
    }
  }

  # Convert a .docx to one PNG per page via LibreOffice + pdftools / magick.
  # Returns a character vector of PNG file paths, or NULL on failure.
  docx_to_pngs <- function(docx_path) {
    if (!has_lo || !has_pdf_converter) return(NULL)

    stem   <- tools::file_path_sans_ext(basename(docx_path))
    tmpdir <- tempfile(pattern = "raptlr_pptx_")
    dir.create(tmpdir)
    on.exit(unlink(tmpdir, recursive = TRUE), add = TRUE)

    # docx → PDF
    system2(lo_path[[1]],
            args   = c("--headless", "--convert-to", "pdf",
                       "--outdir", tmpdir, docx_path),
            stdout = FALSE, stderr = FALSE)

    pdf_path <- file.path(tmpdir, paste0(stem, ".pdf"))
    if (!file.exists(pdf_path)) return(NULL)

    # Determine page count
    n_pages <- tryCatch({
      if (requireNamespace("pdftools", quietly = TRUE))
        pdftools::pdf_info(pdf_path)$pages
      else
        1L
    }, error = function(e) 1L)

    # PDF pages → PNG files (kept in system tempdir, persists after cleanup)
    png_outs <- character(0)

    if (requireNamespace("pdftools", quietly = TRUE)) {

      for (p in seq_len(n_pages)) {
        png_tmp <- file.path(tmpdir, sprintf("%s_p%03d.png", stem, p))
        # suppressWarnings: pdf_convert internally calls sprintf(filenames, pages,
        # format) which warns when the path has no % specifiers — harmless.
        tryCatch(
          suppressWarnings(
            pdftools::pdf_convert(pdf_path, format = "png", pages = p,
                                  filenames = png_tmp, dpi = dpi,
                                  verbose = FALSE)
          ),
          error = function(e) NULL
        )
        if (file.exists(png_tmp)) {
          trim_png(png_tmp)
          png_out <- file.path(tempdir(),
                               sprintf("%s_p%03d_%s.png",
                                       stem, p, format(Sys.time(), "%H%M%S%OS3")))
          file.copy(png_tmp, png_out, overwrite = TRUE)
          png_outs <- c(png_outs, png_out)
        }
      }

    } else if (requireNamespace("magick", quietly = TRUE)) {

      tryCatch({
        imgs <- magick::image_read_pdf(pdf_path, density = dpi)
        for (p in seq_along(imgs)) {
          png_tmp <- file.path(tmpdir, sprintf("%s_p%03d.png", stem, p))
          magick::image_write(imgs[p], path = png_tmp, format = "png")
          if (file.exists(png_tmp)) {
            trim_png(png_tmp)
            png_out <- file.path(tempdir(),
                                 sprintf("%s_p%03d_%s.png",
                                         stem, p, format(Sys.time(), "%H%M%S%OS3")))
            file.copy(png_tmp, png_out, overwrite = TRUE)
            png_outs <- c(png_outs, png_out)
          }
        }
      }, error = function(e) NULL)
    }

    if (length(png_outs) == 0) NULL else png_outs
  }

  # Add a content slide with a compact title (18 pt) and a properly sized image.
  # Small images are kept at their natural size and centred; large images are
  # scaled down to fit without ever being upscaled.
  add_image_slide <- function(pptx, img_path, slide_title,
                              lo_content, ph_content) {

    pptx <- officer::add_slide(pptx, layout = lo_content,
                               master = master_for(lo_content))

    # Compact title: 18 pt bold, 0.7" tall — avoids the oversized default font
    title_para <- officer::fpar(
      officer::ftext(slide_title,
                     officer::fp_text(font.size = 18, bold = TRUE))
    )
    pptx <- officer::ph_with(
      pptx, value = title_para,
      location = officer::ph_location(left   = 0.2,
                                      top    = 0.15,
                                      width  = slide_w - 0.4,
                                      height = 0.7)
    )

    # Available content area below the title
    img_area_top <- 0.9
    avail_w      <- slide_w - 0.4
    avail_h      <- slide_h - img_area_top - 0.1

    # Determine display dimensions:
    #   - If natural size (pixels / dpi) is known, fit it within the available
    #     area while preserving aspect ratio and never upscaling.
    #   - Fall back to filling the available area when dimensions are unknown.
    nat <- png_natural_size(img_path)
    if (!is.null(nat)) {
      scale     <- min(avail_w / nat["w"], avail_h / nat["h"], 1.0)
      display_w <- nat["w"] * scale
      display_h <- nat["h"] * scale
      img_left  <- 0.2 + (avail_w - display_w) / 2   # centre horizontally
      img_top   <- img_area_top + (avail_h - display_h) / 2  # centre vertically
    } else {
      display_w <- avail_w
      display_h <- avail_h
      img_left  <- 0.2
      img_top   <- img_area_top
    }

    img  <- officer::external_img(img_path, width = display_w, height = display_h)
    pptx <- officer::ph_with(
      pptx, value = img,
      location = officer::ph_location(left   = img_left,
                                      top    = img_top,
                                      width  = display_w,
                                      height = display_h)
    )
    pptx
  }

  # ── Opening title slide ────────────────────────────────────────────────────

  if (!is.null(title)) {
    lo_title  <- pick_layout("Title Slide", "Title, Content", "Blank")
    ph_title  <- ph_types_of(lo_title)

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

  lo_content <- pick_layout("Title and Content", "Title, Content", "Blank")
  ph_content <- ph_types_of(lo_content)
  lo_section <- pick_layout("Section Header", "Title Slide", lo_content)
  ph_section <- ph_types_of(lo_section)

  for (section_name in names(output_structure_list)) {

    outputs_in_section <- output_structure_list[[section_name]]

    # Section divider slide
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

      if (is_image(file_path)) {
        # ── Figure: embed image directly ───────────────────────────────────
        slide_title <- tools::file_path_sans_ext(basename(file_path))
        pptx        <- add_image_slide(pptx, file_path, slide_title,
                                       lo_content, ph_content)

      } else if (grepl("\\.docx$", file_path, ignore.case = TRUE)) {
        # ── Table: convert docx → PNG(s) and embed ─────────────────────────
        slide_title <- extract_tlf_title(file_path)
        png_paths   <- docx_to_pngs(file_path)

        if (!is.null(png_paths)) {
          for (page_idx in seq_along(png_paths)) {
            page_title <- if (page_idx == 1) {
              slide_title
            } else {
              paste0(slide_title, " (cont'd)")
            }
            pptx <- add_image_slide(pptx, png_paths[[page_idx]], page_title,
                                    lo_content, ph_content)
          }
        } else {
          # Fallback: title-only slide with install hint
          pptx <- officer::add_slide(pptx, layout = lo_content,
                                     master = master_for(lo_content))
          if ("title" %in% ph_content)
            pptx <- officer::ph_with(
              pptx, value = slide_title,
              location = officer::ph_location_type("title")
            )
          hint <- if (!has_lo) {
            "Install LibreOffice to render this table as an image."
          } else {
            "Install the 'pdftools' package to render this table as an image.\nRun: install.packages('pdftools')"
          }
          if ("body" %in% ph_content)
            pptx <- officer::ph_with(
              pptx, value = hint,
              location = officer::ph_location_type("body")
            )
        }

      } else {
        warning("Unsupported file type for '", basename(file_path),
                "'. Skipping.")
      }
    }
  }

  # ── Save ───────────────────────────────────────────────────────────────────

  print(pptx, target = return_to_file)
  message("PowerPoint saved to: ", return_to_file)
  invisible(return_to_file)
}
