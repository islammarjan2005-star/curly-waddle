# word_helpers.R - shared formatters, replacers, and oecd helpers for word output
#
# sourced by both word_output.R (database mode) and manual_word_output.R (excel mode)
# to avoid duplicating these functions across both files.

suppressPackageStartupMessages({
  library(officer)
  library(xml2)
  library(scales)
})

# -- formatters --
# these try to use the helpers.R versions if available, falling back to local logic.
# this lets the word output scripts work standalone or after helpers.R has been sourced.

fmt_one_dec <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  if (x == 0) return(format(0, nsmall = 1, trim = TRUE))
  for (d in 1:4) {
    vr <- round(x, d)
    if (vr != 0) return(format(vr, nsmall = d, trim = TRUE))
  }
  format(round(x, 4), nsmall = 4, trim = TRUE)
}

.format_int <- function(x) {
  if (exists("format_int_unsigned", inherits = TRUE)) return(get("format_int_unsigned", inherits = TRUE)(x))
  if (exists("format_int", inherits = TRUE)) return(get("format_int", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  scales::comma(round(x), accuracy = 1)
}

.format_pct <- function(x) {
  if (exists("format_pct", inherits = TRUE)) return(get("format_pct", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  paste0(fmt_one_dec(x), "%")
}

.format_pp <- function(x) {
  if (exists("format_pp", inherits = TRUE)) return(get("format_pp", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  sign <- if (x > 0) "+" else if (x < 0) "-" else ""
  paste0(sign, fmt_one_dec(abs(x)), "pp")
}

.format_gbp_signed0 <- function(x) {
  if (exists("format_gbp_signed0", inherits = TRUE)) return(get("format_gbp_signed0", inherits = TRUE)(x))
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  sign <- if (x > 0) "+" else if (x < 0) "-" else ""
  paste0(sign, "\u00A3", scales::comma(round(abs(x)), accuracy = 1))
}

fmt_int_signed <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (length(x) == 0 || is.na(x)) return("")
  s <- scales::comma(abs(round(x)), accuracy = 1)
  if (x > 0) paste0("+", s) else if (x < 0) paste0("-", s) else "0"
}

# counts stored as persons; displayed in 000s
fmt_count_000s_current <- function(x) .format_int(x / 1000)
fmt_count_000s_change  <- function(x) fmt_int_signed(x / 1000)

# payroll/vacancies already in 000s
fmt_exempt_current <- function(x) .format_int(x)
fmt_exempt_change  <- function(x) fmt_int_signed(x)

# -- month label --
# converts "mar2026" or "2025-10" to "March 2026"

manual_month_to_label <- function(x) {
  if (is.null(x) || length(x) == 0 || is.na(x)) return("")
  x <- tolower(as.character(x))
  if (grepl("^[0-9]{4}-[0-9]{2}$", x)) {
    parts <- strsplit(x, "-", fixed = TRUE)[[1]]
    d <- as.Date(sprintf("%s-%s-01", parts[1], parts[2]))
    return(format(d, "%B %Y"))
  }
  if (grepl("^[a-z]{3}[0-9]{4}$", x)) {
    mon <- substr(x, 1, 3)
    yr <- substr(x, 4, 7)
    month_map <- c(jan=1,feb=2,mar=3,apr=4,may=5,jun=6,jul=7,aug=8,sep=9,oct=10,nov=11,dec=12)
    if (mon %in% names(month_map)) {
      d <- as.Date(sprintf("%s-%02d-01", yr, month_map[[mon]]))
      return(format(d, "%B %Y"))
    }
  }
  paste0(toupper(substr(x, 1, 1)), substr(x, 2, nchar(x)))
}

# -- docx replacement --
# replaces placeholder text in body, headers, and footers of a word doc

replace_all <- function(doc, key, val) {
  if (is.null(val) || length(val) == 0 || is.na(val)) val <- ""
  val <- as.character(val)

  body_xml <- doc$doc_obj$get()
  ns <- xml2::xml_ns(body_xml)
  text_nodes <- xml2::xml_find_all(body_xml, ".//w:t", ns = ns)
  for (node in text_nodes) {
    txt <- xml2::xml_text(node)
    if (grepl(key, txt, fixed = TRUE)) {
      new_txt <- gsub(key, val, txt, fixed = TRUE)
      xml2::xml_text(node) <- new_txt
      xml2::xml_attr(node, "xml:space") <- "preserve"
    }
  }

  doc <- tryCatch(headers_replace_all_text(doc, key, val, fixed = TRUE), error = function(e) doc)
  doc <- tryCatch(footers_replace_all_text(doc, key, val, fixed = TRUE), error = function(e) doc)
  doc
}

# conditional fill — picks the p/n/z placeholder based on sign of value_num
fill_conditional <- function(doc, base, value_text, value_num, invert = FALSE, neutral = FALSE) {
  value_num <- suppressWarnings(as.numeric(value_num))

  p <- n <- z <- ""

  if (is.na(value_num)) {
    z <- "\u2014"
  } else if (isTRUE(neutral)) {
    z <- value_text
  } else {
    if (value_num > 0) p <- value_text
    if (value_num < 0) n <- value_text
    if (value_num == 0) z <- value_text
    if (isTRUE(invert)) { tmp <- p; p <- n; n <- tmp }
  }

  doc <- replace_all(doc, paste0(base, "p"), p)
  doc <- replace_all(doc, paste0(base, "n"), n)
  doc <- replace_all(doc, paste0(base, "z"), z)
  doc
}

# safe variable lookup — returns the value of a global variable or a default
sv <- function(name, default = NA_real_) {
  if (exists(name, inherits = TRUE)) get(name, inherits = TRUE) else default
}

# -- oecd comparison data --

.oecd_countries <- c(
  "United Kingdom", "United States", "France", "Germany",
  "Italy", "Spain", "Canada", "Japan", "Euro area"
)

.oecd_country_aliases <- list(
  "United Kingdom" = c("United Kingdom", "GBR"),
  "United States"  = c("United States", "USA"),
  "France"         = c("France", "FRA"),
  "Germany"        = c("Germany", "DEU"),
  "Italy"          = c("Italy", "ITA"),
  "Spain"          = c("Spain", "ESP"),
  "Canada"         = c("Canada", "CAN"),
  "Japan"          = c("Japan", "JPN"),
  "Euro area"      = c("Euro area", "Euro area (20 countries)",
                       "Euro area (19 countries)", "EA20", "EA19", "EA")
)

# fills oecd placeholders in the word doc for all 9 countries
.fill_oecd_placeholders <- function(doc, unemp_data, emp_data, inact_data) {
  country_codes <- c("United Kingdom" = "uk", "United States" = "us",
                     "France" = "fr", "Germany" = "de", "Italy" = "it",
                     "Spain" = "es", "Canada" = "ca", "Japan" = "jp",
                     "Euro area" = "ea")

  .get_val <- function(data, country) {
    if (is.null(data)) return(list(period = "", value = ""))
    idx <- match(country, data$country)
    if (is.na(idx)) return(list(period = "", value = ""))
    list(period = data$period[idx],
         value = paste0(fmt_one_dec(data$value[idx]), "%"))
  }

  for (country in names(country_codes)) {
    cc <- country_codes[[country]]
    u <- .get_val(unemp_data, country)
    e <- .get_val(emp_data, country)
    ia <- .get_val(inact_data, country)

    tp <- if (nzchar(u$period)) u$period else if (nzchar(e$period)) e$period else ia$period
    if (cc == "uk" && nzchar(tp)) tp <- paste0(tp, "*")

    doc <- replace_all(doc, paste0("qvzoecd", cc, "tp"), tp)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "ur"), u$value)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "er"), e$value)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "ir"), ia$value)
  }
  doc
}
