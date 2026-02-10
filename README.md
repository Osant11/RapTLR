# RapTLR

**Rapid Top-Line Report** - An R package for streamlined creation of clinical study reports with dynamic text generation and embedded appendices.

## Overview

RapTLR is designed to enhance the generation of time-sensitive formal reports, particularly clinical study reports. It provides tools to:

- Manipulate Microsoft Word (.docx) documents programmatically in R
- Replace template keywords with dynamically generated text and statistics
- Generate formatted narrative text for clinical reports
- Create appendices with linked Tables, Listings, and Figures (TLFs)
- Build natural language summaries from CDISC ADaM datasets

## Installation

```r
# Install from GitHub
# install.packages("devtools")
devtools::install_github("johnsonandjohnson/RapTLR")
```

## Quick Start

### 1. Import the Demo Script

Have a look at the complete demo and process from appendice creation to TLR generation. Run `import_demo()` and a script called `demo_TLR.R` will be created in your current working directory. 

```r
library(RapTLR)

# Copy the demo script to your working directory
import_demo()
```

### 2. Basic Keyword Replacement

Replace placeholder keywords in Word documents with dynamic content:

```r
library(RapTLR)

# Load a template document
docx_path <- system.file("extdata", "Keywords_replace.docx", package = "RapTLR")

# Replace a keyword with new text
textReplace(
  doc_object = docx_path,
  keyword = "TTsubjectTT",
  replacement = "150 participants were enrolled"
)
```

Chain multiple replacements using the pipe operator:

```r
library(officer)

docx_path <- system.file("extdata", "Keywords_replace.docx", package = "RapTLR")

read_docx(docx_path) |>
  textReplace(keyword = "TTsubjectTT", replacement = "150 participants") |>
  textReplace(keyword = "TTsiteTT", replacement = "across 25 sites") |>
  textReplace(keyword = "TTconclusionTT",
              replacement = "The primary endpoint was met",
              return_to_file = "final_report.docx")
```

### 3. Generate Statistical Summaries

Create formatted narrative text from clinical data:

```r
data(tlr_adsl)

# Generate baseline statistics text
fct_bsl_stats(
  arg_dataset = tlr_adsl,
  arg_grp = ARM,
  arg_var_num = AGE,
  arg_var1 = "mean",
  arg_var2 = "sd",
  arg_unit = "years",
  arg_label = "age"
)
# Output: "For age the overall mean was 75.1 years (SD 8.2), for Placebo..."
```

### 4. Create Appendices with TLFs

Automatically build appendices with linked tables, listings, and figures:

```r
path_TLFs <- system.file("extdata/TLF_outputs", package = "RapTLR")
path_docx <- system.file("extdata/Keywords_replace.docx", package = "RapTLR")
TLF_list <- system.file("extdata/TLF_list.xlsx", package = "RapTLR")

run_apdx(
  path_TLFs = path_TLFs,
  docx_object = path_docx,
  sections_structure = TLF_list,
  return_to_file = "report_with_appendix.docx"
)
```

## String Concatenation Operators

RapTLR provides custom infix operators for building narrative text:

| Operator | Separator | Example | Result |
|----------|-----------|---------|--------|
| `%c%` | space | `"Hello" %c% "World"` | `"Hello World"` |
| `%c,%` | comma | `"red" %c,% "blue"` | `"red, blue"` |
| `%c&%` | and | `"efficacy" %c&% "safety"` | `"efficacy and safety"` |
| `%c.%` | period | `"First" %c.% "Second"` | `"First. Second"` |

Build complex sentences naturally:

```r
"The study enrolled" %c% "120 participants" %c.%
"Most were" %c% "White (85%)" %c&% "female (62%)"
# "The study enrolled 120 participants. Most were White (85%) and female (62%)"
```

## Function Reference

### Document Manipulation

| Function | Description |
|----------|-------------|
| `textReplace()` | Replace keywords in Word documents with dynamic text |
| `run_apdx()` | Create appendices with linked TLF files |
| `import_demo()` | Copy overall demo script to working directory |

### Statistical Summary Functions

| Function | Description |
|----------|-------------|
| `fct_bsl_stats()` | Generate baseline statistics text (mean/median with SD/range) |
| `fct_grp_tbl()` | Count unique subjects by grouping variable |
| `fct_mst_aes()` | Summarize most common adverse events by treatment group |
| `fct_smpl_rest()` | Count records with custom text formatting |
| `fct_var_text()` | Create formatted text describing variable distribution |
| `var_text_group()` | Generate grouped variable summary text |

### String Operators

| Operator | Description |
|----------|-------------|
| `%c%` | Concatenate with space separator |
| `%c,%` | Concatenate with comma separator |
| `%c&%` | Concatenate with "and" separator |
| `%c.%` | Concatenate with period separator |

## Included Datasets

RapTLR includes two example clinical trial datasets in CDISC ADaM format:

- **`tlr_adsl`**: Subject-level analysis dataset with demographics and baseline characteristics
- **`tlr_adae`**: Adverse events analysis dataset with treatment-emergent adverse events

```r
data(tlr_adsl)
data(tlr_adae)
```

## Example Templates

The package includes template files accessible via `system.file()`:

```r
# Word template with keywords
system.file("extdata", "Keywords_replace.docx", package = "RapTLR")

# TLR shell document
system.file("extdata", "TLR_Shell.docx", package = "RapTLR")

# Appendix structure files
system.file("extdata", "TLF_list.xlsx", package = "RapTLR")
system.file("extdata", "TLF_list.csv", package = "RapTLR")

# Example TLF outputs
system.file("extdata/TLF_outputs", package = "RapTLR")
```

## Complete Workflow Example

```r
library(RapTLR)
library(officer)
library(dplyr)

# Load clinical data
data(tlr_adsl)
data(tlr_adae)

# Step 1: Create appendix with TLFs
path_TLFs <- system.file("extdata/TLF_outputs", package = "RapTLR")
path_docx <- system.file("extdata/TLR_Shell.docx", package = "RapTLR")
TLF_list <- system.file("extdata/TLF_list.xlsx", package = "RapTLR")

run_apdx(
  path_TLFs = path_TLFs,
  docx_object = path_docx,
  sections_structure = TLF_list,
  return_to_file = "TLR_with_appendix.docx"
)

# Step 2: Generate dynamic text for report sections
TLR <- read_docx("TLR_with_appendix.docx")

# Demographics section
demo_text <- fct_smpl_rest(
  arg_dataset = tlr_adsl,
  arg_text_desc = "This study randomized ",
  arg_text_end = " participants"
) %c%
fct_var_text(
  arg_dataset = tlr_adsl,
  arg_var_rest = TRT01A,
  arg_lbl = "nbr-pct"
) %c.%
fct_bsl_stats(
  arg_dataset = tlr_adsl,
  arg_grp = ARM,
  arg_var_num = AGE,
  arg_var1 = "median",
  arg_var2 = "range",
  arg_unit = "years",
  arg_label = "age"
)

# Step 3: Replace keywords with generated text
textReplace(TLR, "TT_geninfo_TT", demo_text)

# Safety section
safety_text <- fct_mst_aes(
  arg_data = tlr_adae,
  arg_data_adsl = tlr_adsl,
  arg_grp = ARM,
  arg_var = AEBODSYS,
  arg_trt_flt = `Xanomeline High Dose`,
  arg_frq = 10
)

textReplace(TLR, "TT_mosttaes_TT", safety_text,
            return_to_file = "final_TLR.docx")
```

## Dependencies

RapTLR relies on the following packages:

- [officer](https://davidgohel.github.io/officer/) - Word document manipulation
- [dplyr](https://dplyr.tidyverse.org/) - Data transformation
- [stringr](https://stringr.tidyverse.org/) - String manipulation
- [purrr](https://purrr.tidyverse.org/) - Functional programming
- [rlang](https://rlang.r-lib.org/) - Non-standard evaluation
- [xml2](https://xml2.r-lib.org/) - XML parsing
- [crosstable](https://danchaltiel.github.io/crosstable/) - Text formatting
- [readxl](https://readxl.tidyverse.org/) - Excel file reading
- [magrittr](https://magrittr.tidyverse.org/) - Pipe operator
- [uuid](https://cran.r-project.org/package=uuid) - Unique identifiers

## License

MIT License

## Authors

- Xiang Li
- Lauren Crow
- Caesar Li
- Yannick Vandendijck
- Antoine Stos
