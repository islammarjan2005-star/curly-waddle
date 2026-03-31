# word_output.R - generate word output from database calculations

suppressPackageStartupMessages({
  library(officer)
  library(scales)
})

# local fallback formatters (used when helpers.R hasn't been sourced)
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

# "oct2025" or "2025-10" -> "October 2025"
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

# fill_conditional uses no-underscore suffix (p/n/z) to match ManualDB.docx placeholders
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

sv <- function(name, default = NA_real_) {
  if (exists(name, inherits = TRUE)) get(name, inherits = TRUE) else default
}

# oecd comparison data helpers

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
    if (cc == "uk" && nzchar(tp)) tp <- paste0(tp, "*")  # uk row gets asterisk

    doc <- replace_all(doc, paste0("qvzoecd", cc, "tp"), tp)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "ur"), u$value)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "er"), e$value)
    doc <- replace_all(doc, paste0("qvzoecd", cc, "ir"), ia$value)
  }
  doc
}

# fetch OECD data from database
.fetch_oecd_from_db <- function(verbose = TRUE) {
  if (!requireNamespace("DBI", quietly = TRUE) || !requireNamespace("RPostgres", quietly = TRUE)) {
    if (verbose) message("[word_output] DBI/RPostgres not available for OECD data")
    return(list(unemp = NULL, emp = NULL, inact = NULL))
  }

  country_map <- c(GBR = "United Kingdom", USA = "United States", FRA = "France",
                   DEU = "Germany", ITA = "Italy", ESP = "Spain",
                   CAN = "Canada", JPN = "Japan", EA20 = "Euro area")

  .query_table <- function(conn, table_name) {
    tryCatch({
      query <- sprintf(
        'SELECT ref_area, time_period, obs_value FROM "oecd"."%s" WHERE ref_area IN (\'GBR\',\'USA\',\'FRA\',\'DEU\',\'ITA\',\'ESP\',\'CAN\',\'JPN\',\'EA20\') ORDER BY ref_area, time_period DESC',
        table_name
      )
      raw <- DBI::dbGetQuery(conn, query)
      if (nrow(raw) == 0) return(NULL)
      raw <- raw[!duplicated(raw$ref_area), , drop = FALSE]
      raw$country <- country_map[raw$ref_area]
      raw$value <- as.numeric(raw$obs_value)
      raw$period <- raw$time_period
      raw[!is.na(raw$country), c("country", "period", "value"), drop = FALSE]
    }, error = function(e) {
      if (verbose) message("[word_output] OECD table query failed: ", e$message)
      NULL
    })
  }

  conn <- NULL
  result <- list(unemp = NULL, emp = NULL, inact = NULL)
  tryCatch({
    conn <- DBI::dbConnect(RPostgres::Postgres())
    result$unemp <- .query_table(conn, "labour_statistics__unemployment_rate")
    result$emp   <- .query_table(conn, "labour_statistics__employment_rate")
    result$inact <- .query_table(conn, "labour_statistics__inactivity_rate")
    if (verbose) message("[word_output] OECD data fetched from database")
  }, error = function(e) {
    if (verbose) message("[word_output] OECD DB connection failed: ", e$message)
  }, finally = {
    if (!is.null(conn)) try(DBI::dbDisconnect(conn), silent = TRUE)
  })

  result
}

generate_word_output <- function(template_path = "utils/ManualDB.docx",
                                 output_path = "utils/DBoutput.docx",
                                 calculations_path = "utils/calculations.R",
                                 config_path = "utils/config.R",
                                 summary_path = "sheets/summary.R",
                                 top_ten_path = "sheets/top_ten_stats.R",
                                 manual_month_override = NULL,
                                 vacancies_mode_override = NULL,
                                 payroll_mode_override = NULL,
                                 vac_payroll_mode_override = NULL,
                                 contact_names = NULL,
                                 verbose = TRUE) {

  source(config_path, local = FALSE)
  if (!is.null(manual_month_override)) manual_month <<- tolower(manual_month_override)

  # set vac/payroll modes
  if (!is.null(vac_payroll_mode_override) && is.null(vacancies_mode_override) && is.null(payroll_mode_override)) {
    mode <- tolower(as.character(vac_payroll_mode_override))
    mode <- if (mode %in% c("latest", "aligned")) mode else "latest"
    vacancies_mode <<- mode
    payroll_mode   <<- mode
  }
  if (!is.null(vacancies_mode_override)) {
    vacancies_mode <<- if (tolower(vacancies_mode_override) %in% c("latest", "aligned")) tolower(vacancies_mode_override) else "latest"
  }
  if (!is.null(payroll_mode_override)) {
    payroll_mode <<- if (tolower(payroll_mode_override) %in% c("latest", "aligned")) tolower(payroll_mode_override) else "latest"
  }

  source(calculations_path, local = FALSE)
  if (verbose && exists("manual_month", inherits = TRUE)) message("[word_output] manual_month = ", manual_month)

  # save dashboard vac values, then re-run with "latest" for narrative text
  saved_vac       <- list(cur = vac_cur, dq = vac_dq, dy = vac_dy, dc = vac_dc, de = vac_de)
  saved_vac_obj   <- if (exists("vac", inherits = TRUE)) get("vac", inherits = TRUE) else NULL
  saved_vac_label <- if (exists("vacancies_period_short_label", inherits = TRUE)) vacancies_period_short_label else NULL
  tryCatch({
    mm <- if (exists("manual_month", inherits = TRUE)) manual_month else NULL
    vac_lat <- calculate_vacancies(mm, mode = "latest")
    vac_cur <<- vac_lat$cur; vac_dq <<- vac_lat$dq; vac_dy <<- vac_lat$dy
    vac_dc  <<- vac_lat$dc;  vac_de <<- vac_lat$de
    vac     <<- vac_lat
    if (!is.na(vac_lat$end)) vacancies_period_short_label <<- make_lfs_label(vac_lat$end)
  }, error = function(e) NULL)

  source(summary_path, local = FALSE)
  source(top_ten_path, local = FALSE)
  fallback_lines <- function() { stats <- list(); for (i in 1:10) stats[[paste0("line", i)]] <- "(Data unavailable)"; stats }
  summary <- tryCatch(generate_summary(),  error = function(e) { warning("generate_summary() failed: ", e$message);  fallback_lines() })
  top10   <- tryCatch(generate_top_ten(),  error = function(e) { warning("generate_top_ten() failed: ", e$message);  fallback_lines() })

  # restore dropdown-selected vac values for template filling
  vac_cur <<- saved_vac$cur; vac_dq <<- saved_vac$dq; vac_dy <<- saved_vac$dy
  vac_dc  <<- saved_vac$dc;  vac_de <<- saved_vac$de
  if (!is.null(saved_vac_obj))   vac                       <<- saved_vac_obj
  if (!is.null(saved_vac_label)) vacancies_period_short_label <<- saved_vac_label

  doc <- read_docx(template_path)

  # Replace contact names in header
  contact <- if (!is.null(contact_names) && nzchar(contact_names)) contact_names else "Zaynah Asad and Jevan Reynolds"
  doc <- replace_all(doc, "Zaynah Asad and Jevan Reynolds", contact)
  doc <- replace_all(doc, "qvzcontact", contact)

  # header placeholders (qvz convention for ManualDB.docx)
  title_label <- if (exists("manual_month", inherits = TRUE)) manual_month_to_label(manual_month) else ""
  doc <- replace_all(doc, "qvzmonthlabel", title_label)
  doc <- replace_all(doc, "qvzrenderdate", format(Sys.Date(), "%d %B %Y"))
  if (exists("lfs_period_label",            inherits = TRUE)) doc <- replace_all(doc, "qvzlfsperiod",  lfs_period_label)
  if (exists("lfs_period_short_label",      inherits = TRUE)) doc <- replace_all(doc, "qvzlfsquarter", lfs_period_short_label)
  if (exists("vacancies_period_short_label",inherits = TRUE)) {
    doc <- replace_all(doc, "qvzvacquarter", vacancies_period_short_label)
    doc <- replace_all(doc, "qvzvacperiod",  vacancies_period_short_label)
  }
  if (exists("payroll_period_short_label",  inherits = TRUE)) doc <- replace_all(doc, "qozpayperiod", payroll_period_short_label)

  # summary & top ten lines
  for (i in 1:10) doc <- replace_all(doc, sprintf("qvzsl%02d", i), summary[[paste0("line", i)]])
  for (i in 1:10) doc <- replace_all(doc, sprintf("qvztt%02d", i), top10[[paste0("line", i)]])

  # current values
  doc <- replace_all(doc, "qvzempcur",  fmt_count_000s_current(emp16_cur))
  doc <- replace_all(doc, "qvzertcur",  .format_pct(emp_rt_cur))
  doc <- replace_all(doc, "qvzunecur",  fmt_count_000s_current(unemp16_cur))
  doc <- replace_all(doc, "qvzurtcur",  .format_pct(unemp_rt_cur))
  doc <- replace_all(doc, "qvzinacur",  fmt_count_000s_current(inact_cur))
  doc <- replace_all(doc, "qvzifecur",  fmt_count_000s_current(inact5064_cur))
  doc <- replace_all(doc, "qvzirtcur",  .format_pct(inact_rt_cur))
  doc <- replace_all(doc, "qvzifrcur",  .format_pct(inact5064_rt_cur))
  doc <- replace_all(doc, "qvzpaycur",  fmt_exempt_current(payroll_cur))
  doc <- fill_conditional(doc, "qvzvaccur", fmt_exempt_current(vac_cur), 0, neutral = TRUE)
  doc <- replace_all(doc, "qvzwnocur",  .format_pct(latest_wages))
  doc <- replace_all(doc, "qvzwcpcur",  .format_pct(latest_wages_cpi))

  # change on quarter
  doc <- fill_conditional(doc, "qvzempdq",  fmt_count_000s_change(emp16_dq),        emp16_dq)
  doc <- fill_conditional(doc, "qvzertdq",  .format_pp(emp_rt_dq),                  emp_rt_dq)
  doc <- fill_conditional(doc, "qvzunedq",  fmt_count_000s_change(unemp16_dq),      unemp16_dq,      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtdq",  .format_pp(unemp_rt_dq),                unemp_rt_dq,     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinadq",  fmt_count_000s_change(inact_dq),         inact_dq,        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifedq",  fmt_count_000s_change(inact5064_dq),     inact5064_dq,    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtdq",  .format_pp(inact_rt_dq),                 inact_rt_dq,     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrdq",  .format_pp(inact5064_rt_dq),             inact5064_rt_dq, invert = TRUE)
  doc <- fill_conditional(doc, "qvzpaydq",  fmt_exempt_change(payroll_dq),            payroll_dq)
  doc <- fill_conditional(doc, "qvzvacdq",  fmt_exempt_change(vac_dq),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnodq",  .format_gbp_signed0(wages_change_q),     wages_change_q)
  doc <- fill_conditional(doc, "qvzwcpdq",  .format_gbp_signed0(wages_cpi_change_q), wages_cpi_change_q)

  # change on year
  doc <- fill_conditional(doc, "qvzempdy",  fmt_count_000s_change(emp16_dy),        emp16_dy)
  doc <- fill_conditional(doc, "qvzertdy",  .format_pp(emp_rt_dy),                  emp_rt_dy)
  doc <- fill_conditional(doc, "qvzunedy",  fmt_count_000s_change(unemp16_dy),      unemp16_dy,      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtdy",  .format_pp(unemp_rt_dy),                unemp_rt_dy,     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinady",  fmt_count_000s_change(inact_dy),         inact_dy,        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifedy",  fmt_count_000s_change(inact5064_dy),     inact5064_dy,    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtdy",  .format_pp(inact_rt_dy),                 inact_rt_dy,     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrdy",  .format_pp(inact5064_rt_dy),             inact5064_rt_dy, invert = TRUE)
  doc <- fill_conditional(doc, "qvzpaydy",  fmt_exempt_change(payroll_dy),            payroll_dy)
  doc <- fill_conditional(doc, "qvzvacdy",  fmt_exempt_change(vac_dy),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnody",  .format_gbp_signed0(wages_change_y),     wages_change_y)
  doc <- fill_conditional(doc, "qvzwcpdy",  .format_gbp_signed0(wages_cpi_change_y), wages_cpi_change_y)

  # change since covid
  doc <- fill_conditional(doc, "qvzempdc",  fmt_count_000s_change(emp16_dc),        emp16_dc)
  doc <- fill_conditional(doc, "qvzertdc",  .format_pp(emp_rt_dc),                  emp_rt_dc)
  doc <- fill_conditional(doc, "qvzunedc",  fmt_count_000s_change(unemp16_dc),      unemp16_dc,      invert = TRUE)
  doc <- fill_conditional(doc, "qvzurtdc",  .format_pp(unemp_rt_dc),                unemp_rt_dc,     invert = TRUE)
  doc <- fill_conditional(doc, "qvzinadc",  fmt_count_000s_change(inact_dc),         inact_dc,        invert = TRUE)
  doc <- fill_conditional(doc, "qvzifedc",  fmt_count_000s_change(inact5064_dc),     inact5064_dc,    invert = TRUE)
  doc <- fill_conditional(doc, "qvzirtdc",  .format_pp(inact_rt_dc),                 inact_rt_dc,     invert = TRUE)
  doc <- fill_conditional(doc, "qvzifrdc",  .format_pp(inact5064_rt_dc),             inact5064_rt_dc, invert = TRUE)
  doc <- fill_conditional(doc, "qvzpaydc",  fmt_exempt_change(payroll_dc),            payroll_dc)
  doc <- fill_conditional(doc, "qvzvacdc",  fmt_exempt_change(vac_dc),               0, neutral = TRUE)
  doc <- fill_conditional(doc, "qvzwnodc",  .format_gbp_signed0(wages_change_covid),     wages_change_covid)
  doc <- fill_conditional(doc, "qvzwcpdc",  .format_gbp_signed0(wages_cpi_change_covid), wages_cpi_change_covid)

  # change since 2024 election
  if (exists("emp16_de", inherits = TRUE)) {
    doc <- fill_conditional(doc, "qvzempde",  fmt_count_000s_change(emp16_de),        emp16_de)
    doc <- fill_conditional(doc, "qvzertde",  .format_pp(emp_rt_de),                  emp_rt_de)
    doc <- fill_conditional(doc, "qvzunede",  fmt_count_000s_change(unemp16_de),      unemp16_de,      invert = TRUE)
    doc <- fill_conditional(doc, "qvzurtde",  .format_pp(unemp_rt_de),                unemp_rt_de,     invert = TRUE)
    doc <- fill_conditional(doc, "qvzinade",  fmt_count_000s_change(inact_de),         inact_de,        invert = TRUE)
    doc <- fill_conditional(doc, "qvzifede",  fmt_count_000s_change(inact5064_de),     inact5064_de,    invert = TRUE)
    doc <- fill_conditional(doc, "qvzirtde",  .format_pp(inact_rt_de),                 inact_rt_de,     invert = TRUE)
    doc <- fill_conditional(doc, "qvzifrde",  .format_pp(inact5064_rt_de),             inact5064_rt_de, invert = TRUE)
    doc <- fill_conditional(doc, "qvzpayde",  fmt_exempt_change(payroll_de),            payroll_de)
    doc <- fill_conditional(doc, "qvzvacde",  fmt_exempt_change(vac_de),               0, neutral = TRUE)

    if (exists("wages_change_election", inherits = TRUE)) {
      doc <- fill_conditional(doc, "qvzwnode",  .format_gbp_signed0(wages_change_election),     wages_change_election)
    }
    if (exists("wages_cpi_change_election", inherits = TRUE)) {
      doc <- fill_conditional(doc, "qvzwcpde",  .format_gbp_signed0(wages_cpi_change_election), wages_cpi_change_election)
    }
  }

  # workforce jobs
  wfj_period  <- if (exists("workforce_jobs", inherits = TRUE)) workforce_jobs$period else ""
  wfj_tbl     <- if (exists("workforce_jobs", inherits = TRUE)) workforce_jobs$data else NULL
  wfj_top_ind <- if (!is.null(wfj_tbl) && nrow(wfj_tbl) > 0) as.character(wfj_tbl$industry[1]) else ""
  wfj_top_val <- if (!is.null(wfj_tbl) && nrow(wfj_tbl) > 0) .format_int(wfj_tbl$value[1]) else ""

  doc <- replace_all(doc, "WFJ_PERIOD",                  wfj_period)
  doc <- replace_all(doc, "WORKFORCE_JOBS_PERIOD",        wfj_period)
  doc <- replace_all(doc, "WFJ_TOP_INDUSTRY",             wfj_top_ind)
  doc <- replace_all(doc, "WORKFORCE_JOBS_TOP_INDUSTRY",  wfj_top_ind)
  doc <- replace_all(doc, "WFJ_TOP_VALUE",                wfj_top_val)
  doc <- replace_all(doc, "WORKFORCE_JOBS_TOP_VALUE",     wfj_top_val)
  if (!is.null(wfj_tbl) && nrow(wfj_tbl) > 0) {
    for (i in 1:min(5, nrow(wfj_tbl))) {
      line <- paste0(as.character(wfj_tbl$industry[i]), ": ", .format_int(wfj_tbl$value[i]))
      doc <- replace_all(doc, paste0("WFJ_LINE", i), line)
      doc <- replace_all(doc, paste0("WORKFORCE_JOBS_LINE", i), line)
    }
  }

  # unemployment by age
  uage_period <- if (exists("unemployment_by_age", inherits = TRUE)) unemployment_by_age$period else ""
  uage_level  <- if (exists("unemployment_by_age", inherits = TRUE)) unemployment_by_age$level  else NULL
  uage_rate   <- if (exists("unemployment_by_age", inherits = TRUE)) unemployment_by_age$rate   else NULL

  doc <- replace_all(doc, "UNEMP_AGE_PERIOD",           uage_period)
  doc <- replace_all(doc, "UNEMPLOYMENT_BY_AGE_PERIOD", uage_period)
  if (!is.null(uage_level) && nrow(uage_level) > 0) {
    doc <- replace_all(doc, "UNEMP_AGE_TOP_LEVEL_AGE",   as.character(uage_level$age_group[1]))
    doc <- replace_all(doc, "UNEMP_AGE_TOP_LEVEL_VALUE", .format_int(uage_level$value[1]))
    for (i in 1:min(5, nrow(uage_level))) {
      doc <- replace_all(doc, paste0("UNEMP_AGE_LEVEL_LINE", i),
                         paste0(as.character(uage_level$age_group[i]), ": ", .format_int(uage_level$value[i])))
    }
  }
  if (!is.null(uage_rate) && nrow(uage_rate) > 0) {
    doc <- replace_all(doc, "UNEMP_AGE_TOP_RATE_AGE",   as.character(uage_rate$age_group[1]))
    doc <- replace_all(doc, "UNEMP_AGE_TOP_RATE_VALUE", .format_pct(uage_rate$value[1]))
    for (i in 1:min(5, nrow(uage_rate))) {
      doc <- replace_all(doc, paste0("UNEMP_AGE_RATE_LINE", i),
                         paste0(as.character(uage_rate$age_group[i]), ": ", .format_pct(uage_rate$value[i])))
    }
  }

  # payroll by age
  pba_period <- if (exists("payroll_by_age", inherits = TRUE)) payroll_by_age$period else ""
  pba_tbl    <- if (exists("payroll_by_age", inherits = TRUE)) payroll_by_age$data   else NULL

  doc <- replace_all(doc, "PAYROLL_AGE_PERIOD",                  pba_period)
  doc <- replace_all(doc, "PAYROLLED_EMPLOYEES_BY_AGE_PERIOD",   pba_period)
  if (!is.null(pba_tbl) && nrow(pba_tbl) > 0) {
    doc <- replace_all(doc, "PAYROLL_AGE_TOP_AGE",   as.character(pba_tbl$age_group[1]))
    doc <- replace_all(doc, "PAYROLL_AGE_TOP_VALUE", .format_int(pba_tbl$value[1]))
    for (i in 1:min(5, nrow(pba_tbl))) {
      doc <- replace_all(doc, paste0("PAYROLL_AGE_LINE", i),
                         paste0(as.character(pba_tbl$age_group[i]), ": ", .format_int(pba_tbl$value[i])))
    }
  }

  # OECD international comparisons (fetched from database)
  oecd_data <- .fetch_oecd_from_db(verbose = verbose)
  if (!is.null(oecd_data$unemp) || !is.null(oecd_data$emp) || !is.null(oecd_data$inact)) {
    doc <- .fill_oecd_placeholders(doc, oecd_data$unemp, oecd_data$emp, oecd_data$inact)
    if (verbose) message("[word_output] OECD comparison table populated from database")
  }

  # clear any unfilled qvz placeholders with em-dash
  body_xml   <- doc$doc_obj$get()
  ns         <- xml2::xml_ns(body_xml)
  text_nodes <- xml2::xml_find_all(body_xml, ".//w:t", ns = ns)
  for (node in text_nodes) {
    txt <- xml2::xml_text(node)
    if (grepl("qvz", txt, fixed = TRUE)) {
      cleaned <- gsub("qvz[a-z0-9_]+", "\u2014", txt)
      xml2::xml_text(node) <- cleaned
    }
  }

  print(doc, target = output_path)
  invisible(output_path)
}
