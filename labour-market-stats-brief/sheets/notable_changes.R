# ranks each series' latest change by a hybrid z-score + rule-based scorer
# and returns human-readable candidate sentences for the word briefing

.NC_MONTH_MAP <- c(jan=1, feb=2, mar=3, apr=4, may=5, jun=6,
                   jul=7, aug=8, sep=9, oct=10, nov=11, dec=12)

.nc_lfs_label_to_date <- function(lbl) {
  if (is.na(lbl) || !nzchar(lbl)) return(NA_real_)
  mons <- regmatches(lbl, gregexpr("[A-Za-z]{3}", lbl))[[1]]
  yrs  <- regmatches(lbl, gregexpr("[0-9]{4}", lbl))[[1]]
  if (length(mons) < 2 || length(yrs) < 1) return(NA_real_)
  end_m <- .NC_MONTH_MAP[tolower(mons[2])]
  yr    <- suppressWarnings(as.integer(yrs[1]))
  if (is.na(end_m) || is.na(yr)) return(NA_real_)
  as.numeric(as.Date(sprintf("%04d-%02d-01", yr, end_m)))
}

.nc_payroll_label_to_date <- function(lbl) {
  if (is.na(lbl) || !nzchar(lbl)) return(NA_real_)
  d <- suppressWarnings(as.Date(paste0("01 ", lbl), format = "%d %B %Y"))
  if (is.na(d)) return(NA_real_)
  as.numeric(d)
}

.nc_ymd_label_to_date <- function(lbl) {
  if (is.na(lbl) || !nzchar(lbl)) return(NA_real_)
  d <- suppressWarnings(as.Date(substr(lbl, 1, 10)))
  if (is.na(d)) return(NA_real_)
  as.numeric(d)
}

# returns a (date, value) tibble sorted ascending, or NULL
.nc_extract_db <- function(key) {
  if (key == "emp_rt" || key == "unemp_rt" || key == "inact_rt") {
    pg <- get0("lfs_pg_data", envir = globalenv())
    if (is.null(pg) || nrow(pg) == 0) return(NULL)
    code <- switch(key, emp_rt = "LF24", unemp_rt = "MGSX", inact_rt = "LF2S")
    d <- pg[pg$dataset_identifier_code == code, , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    dates <- as.Date(vapply(d$time_period, .nc_lfs_label_to_date, numeric(1)), origin = "1970-01-01")
    vals  <- suppressWarnings(as.numeric(d$value))
    ok    <- !is.na(dates) & !is.na(vals)
    if (sum(ok) < 6) return(NULL)
    res <- data.frame(date = dates[ok], value = vals[ok])
    res[order(res$date), ]
  } else if (key == "vacancies") {
    pg <- get0("vacancies_pg_data", envir = globalenv())
    if (is.null(pg) || nrow(pg) == 0) return(NULL)
    d <- pg[pg$dataset_identifier_code == "AP2Y", , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    dates <- as.Date(vapply(d$time_period, .nc_lfs_label_to_date, numeric(1)), origin = "1970-01-01")
    vals  <- suppressWarnings(as.numeric(d$value))
    ok    <- !is.na(dates) & !is.na(vals)
    if (sum(ok) < 6) return(NULL)
    res <- data.frame(date = dates[ok], value = vals[ok])
    res[order(res$date), ]
  } else if (key == "payroll") {
    pg <- get0("payroll_pg_data", envir = globalenv())
    if (is.null(pg) || nrow(pg) == 0) return(NULL)
    d <- pg[pg$unit_type == "Payrolled employees", , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    dates <- as.Date(vapply(d$time_period, .nc_payroll_label_to_date, numeric(1)), origin = "1970-01-01")
    vals  <- suppressWarnings(as.numeric(d$value))
    ok    <- !is.na(dates) & !is.na(vals)
    if (sum(ok) < 6) return(NULL)
    res <- data.frame(date = dates[ok], value = vals[ok] / 1000)  # thousands
    res[order(res$date), ]
  } else if (key == "wages_nom") {
    pg <- get0("wages_total_pg_data", envir = globalenv())
    if (is.null(pg) || nrow(pg) == 0) return(NULL)
    d <- pg[pg$dataset_identifier_code == "KAC3", , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    dates <- as.Date(vapply(d$time_period, .nc_ymd_label_to_date, numeric(1)), origin = "1970-01-01")
    vals  <- suppressWarnings(as.numeric(d$value))
    ok    <- !is.na(dates) & !is.na(vals)
    if (sum(ok) < 6) return(NULL)
    res <- data.frame(date = dates[ok], value = vals[ok])
    res[order(res$date), ]
  } else if (key == "wages_cpi") {
    pg <- get0("wages_cpi_pg_data", envir = globalenv())
    if (is.null(pg) || nrow(pg) == 0) return(NULL)
    d <- pg[pg$earnings_type == "Total pay, seasonally adjusted" &
              pg$earnings_metric == "% changes year on year" &
              pg$time_basis == "3 month average", , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    dates <- as.Date(vapply(d$time_period, .nc_ymd_label_to_date, numeric(1)), origin = "1970-01-01")
    vals  <- suppressWarnings(as.numeric(d$value))
    ok    <- !is.na(dates) & !is.na(vals)
    if (sum(ok) < 6) return(NULL)
    res <- data.frame(date = dates[ok], value = vals[ok])
    res[order(res$date), ]
  } else if (key == "redund") {
    pg <- get0("redund_pg_data", envir = globalenv())
    if (is.null(pg) || nrow(pg) == 0) return(NULL)
    d <- pg[pg$dataset_identifier_code == "BEIR", , drop = FALSE]
    if (nrow(d) == 0) return(NULL)
    dates <- as.Date(vapply(d$time_period, .nc_lfs_label_to_date, numeric(1)), origin = "1970-01-01")
    vals  <- suppressWarnings(as.numeric(d$value))
    ok    <- !is.na(dates) & !is.na(vals)
    if (sum(ok) < 6) return(NULL)
    res <- data.frame(date = dates[ok], value = vals[ok])
    res[order(res$date), ]
  } else {
    NULL
  }
}

# manual path: each series assigned by calculations_from_excel.R as {key}_manual_series
.nc_extract_manual <- function(key) {
  s <- get0(paste0(key, "_manual_series"), envir = globalenv())
  if (is.null(s) || !is.data.frame(s) || nrow(s) < 6) return(NULL)
  s[order(s$date), c("date", "value")]
}

.nc_get_series <- function(key) {
  m <- .nc_extract_manual(key)
  if (!is.null(m)) return(m)
  .nc_extract_db(key)
}

# z-score of current delta against trailing 24 months of deltas (excluding current)
.nc_score_z <- function(series) {
  if (nrow(series) < 6) return(NULL)
  deltas <- diff(series$value)
  if (length(deltas) < 4) return(NULL)
  cur_delta <- deltas[length(deltas)]
  prior     <- head(tail(deltas, min(24, length(deltas))), -1)
  if (length(prior) < 3) return(NULL)
  m <- mean(prior, na.rm = TRUE)
  s <- sd(prior, na.rm = TRUE)
  if (is.na(s) || s == 0) return(NULL)
  list(
    cur_value  = series$value[nrow(series)],
    cur_date   = series$date[nrow(series)],
    prev_value = series$value[nrow(series) - 1],
    cur_delta  = cur_delta,
    z          = (cur_delta - m) / s
  )
}

# multi-year high/low rule: returns list(tier, months_back, ref_date, direction, boost) or NULL
.nc_rule_multi_year <- function(series) {
  if (nrow(series) < 24) return(NULL)
  cur_v <- series$value[nrow(series)]
  cur_d <- series$date[nrow(series)]
  prior <- series[-nrow(series), ]
  is_high <- cur_v > max(prior$value, na.rm = TRUE)
  is_low  <- cur_v < min(prior$value, na.rm = TRUE)
  if (!is_high && !is_low) return(NULL)
  cmp <- if (is_high) function(a, b) a >= b else function(a, b) a <= b
  hits <- which(cmp(prior$value, cur_v))
  ref_date <- if (length(hits) > 0) prior$date[max(hits)] else prior$date[1]
  months_back <- as.integer(round(as.numeric(difftime(cur_d, ref_date, units = "days")) / 30))
  tier <- if (months_back >= 60) 5L else if (months_back >= 36) 3L else if (months_back >= 24) 2L else NA_integer_
  if (is.na(tier)) return(NULL)
  boost <- c(`2` = 1.5, `3` = 1.8, `5` = 2.2)[as.character(tier)]
  list(tier = tier, months_back = months_back, ref_date = ref_date,
       direction = if (is_high) "high" else "low", boost = unname(boost))
}

# direction reversal: current delta sign flips after >=3 consecutive same-sign deltas
.nc_rule_reversal <- function(series) {
  deltas <- diff(series$value)
  if (length(deltas) < 5) return(NULL)
  cur <- deltas[length(deltas)]
  if (is.na(cur) || cur == 0) return(NULL)
  cur_sign <- sign(cur)
  prior <- rev(head(deltas, -1))
  streak <- 0
  for (d in prior) {
    if (is.na(d) || d == 0) break
    if (sign(d) == cur_sign) break
    streak <- streak + 1
  }
  if (streak < 3) return(NULL)
  list(streak = streak, prev_sign = -cur_sign, boost = 1.5)
}

# threshold crossing since last period
.nc_rule_threshold <- function(series, thresholds) {
  if (length(thresholds) == 0 || nrow(series) < 2) return(NULL)
  cur  <- series$value[nrow(series)]
  prev <- series$value[nrow(series) - 1]
  for (t in thresholds) {
    crossed_up   <- prev <  t && cur >= t
    crossed_down <- prev >= t && cur <  t
    if (crossed_up || crossed_down) {
      return(list(threshold = t,
                  direction = if (crossed_up) "upward" else "downward",
                  boost = 2.0))
    }
  }
  NULL
}

# first significant move after a flat run: |cur z| > 1 after 5+ periods of |z| < 0.5
.nc_rule_first_sig <- function(series) {
  deltas <- diff(series$value)
  if (length(deltas) < 8) return(NULL)
  window <- tail(deltas, min(24, length(deltas)))
  if (length(window) < 8) return(NULL)
  m <- mean(head(window, -1), na.rm = TRUE)
  s <- sd(head(window, -1),   na.rm = TRUE)
  if (is.na(s) || s == 0) return(NULL)
  cur_z <- (window[length(window)] - m) / s
  if (abs(cur_z) <= 1) return(NULL)
  prior_5 <- tail(head(window, -1), 5)
  prior_z <- (prior_5 - m) / s
  if (all(abs(prior_z) < 0.5, na.rm = TRUE)) list(boost = 1.3) else NULL
}

# real-vs-nominal divergence (post-hoc, cross-series)
.nc_rule_real_nominal <- function(nom_delta, cpi_delta) {
  if (is.null(nom_delta) || is.null(cpi_delta) ||
      is.na(nom_delta) || is.na(cpi_delta)) return(NULL)
  if (nom_delta == 0 || cpi_delta == 0) return(NULL)
  if (sign(nom_delta) == sign(cpi_delta)) return(NULL)
  list(boost = 1.5)
}

.nc_fmt_signed <- function(x, digits = 1) {
  if (is.na(x)) return("")
  if (x >= 0) sprintf("+%.*f", digits, x) else sprintf("%.*f", digits, x)
}

.nc_fmt_abs <- function(x, digits = 1) {
  if (is.na(x)) return("") else sprintf("%.*f", digits, abs(x))
}

.nc_period_label <- function(key, d) {
  if (is.na(d)) return("")
  if (key %in% c("emp_rt", "unemp_rt", "inact_rt", "vacancies", "redund")) {
    make_lfs_label(d)
  } else {
    format(d, "%B %Y")
  }
}

.nc_direction_verb <- function(delta, spec) {
  if (is.na(delta)) return("was unchanged")
  if (delta > 0) "rose" else if (delta < 0) "fell" else "was unchanged"
}

.nc_render_sentence <- function(spec, scored, rule_name, rule_ctx) {
  label   <- spec$label
  unit    <- spec$unit
  cur     <- scored$cur_value
  delta   <- scored$cur_delta
  period  <- .nc_period_label(spec$key, scored$cur_date)
  verb    <- .nc_direction_verb(delta, spec)
  abs_d   <- .nc_fmt_abs(delta)
  sgn_d   <- .nc_fmt_signed(delta)
  cur_str <- sprintf("%.1f%s", cur, unit)

  if (rule_name == "multi_year") {
    dir_word <- if (rule_ctx$direction == "high") "highest" else "lowest"
    ref_str  <- .nc_period_label(spec$key, rule_ctx$ref_date)
    sprintf("%s reached %s in %s — its %s level since %s.",
            label, cur_str, period, dir_word, ref_str)
  } else if (rule_name == "reversal") {
    prev_word <- if (rule_ctx$prev_sign > 0) "rises" else "falls"
    sprintf("%s %s by %s%s in %s, ending a run of %d consecutive %s.",
            label, verb, abs_d, unit, period, rule_ctx$streak, prev_word)
  } else if (rule_name == "threshold") {
    dir_word <- rule_ctx$direction
    sprintf("%s crossed %g%s %s in %s, now at %s.",
            label, rule_ctx$threshold, unit, dir_word, period, cur_str)
  } else if (rule_name == "first_sig") {
    sprintf("%s moved notably in %s (%s%s) after five broadly flat periods.",
            label, period, sgn_d, unit)
  } else if (rule_name == "real_nominal") {
    sprintf("Real wages %s while nominal total pay %s in %s — a divergence worth noting.",
            if (rule_ctx$cpi_delta < 0) "fell" else "rose",
            if (rule_ctx$nom_delta < 0) "fell" else "rose",
            period)
  } else {
    mag <- if (abs(scored$z) >= 2) "large" else "notable"
    sprintf("%s %s by %s%s in %s — a %s move vs its recent norm.",
            label, verb, abs_d, unit, period, mag)
  }
}

.NC_SERIES <- list(
  list(key = "emp_rt",    label = "Employment rate 16–64",   unit = "pp",
       thresholds = c(75)),
  list(key = "unemp_rt",  label = "Unemployment rate 16+",        unit = "pp",
       thresholds = c(4, 5, 6)),
  list(key = "inact_rt",  label = "Inactivity rate 16–64",    unit = "pp",
       thresholds = c(20, 22)),
  list(key = "vacancies", label = "Vacancies (000s)",              unit = "",
       thresholds = numeric()),
  list(key = "payroll",   label = "Payrolled employees (000s)",    unit = "",
       thresholds = numeric()),
  list(key = "wages_nom", label = "Wages (total pay, YoY)",        unit = "%",
       thresholds = numeric()),
  list(key = "wages_cpi", label = "Real wages (CPI-adjusted, YoY)", unit = "%",
       thresholds = c(0)),
  list(key = "redund",    label = "Redundancy rate",               unit = "",
       thresholds = numeric())
)

# top-level: scan every series, score, rank, return top candidates
generate_notable_changes <- function(max_n = 10) {
  results <- list()
  scored_by_key <- list()

  for (spec in .NC_SERIES) {
    series <- .nc_get_series(spec$key)
    if (is.null(series)) next
    scored <- .nc_score_z(series)
    if (is.null(scored)) next
    scored_by_key[[spec$key]] <- scored

    candidates <- list()

    r_my <- .nc_rule_multi_year(series)
    if (!is.null(r_my)) candidates[[length(candidates) + 1]] <-
      list(rule = "multi_year", ctx = r_my, boost = r_my$boost)

    r_rv <- .nc_rule_reversal(series)
    if (!is.null(r_rv)) candidates[[length(candidates) + 1]] <-
      list(rule = "reversal", ctx = r_rv, boost = r_rv$boost)

    r_th <- .nc_rule_threshold(series, spec$thresholds)
    if (!is.null(r_th)) candidates[[length(candidates) + 1]] <-
      list(rule = "threshold", ctx = r_th, boost = r_th$boost)

    r_fs <- .nc_rule_first_sig(series)
    if (!is.null(r_fs)) candidates[[length(candidates) + 1]] <-
      list(rule = "first_sig", ctx = r_fs, boost = r_fs$boost)

    if (length(candidates) == 0) {
      if (abs(scored$z) < 1) next
      candidates[[1]] <- list(rule = "plain_z", ctx = list(), boost = 1.0)
    }

    # best candidate for this series = highest score
    for (c in candidates) {
      score <- abs(scored$z) * c$boost
      results[[length(results) + 1]] <- list(
        key      = spec$key,
        score    = score,
        rule     = c$rule,
        sentence = .nc_render_sentence(spec, scored, c$rule, c$ctx)
      )
    }
  }

  # cross-series rule: real-vs-nominal wages divergence
  if (!is.null(scored_by_key$wages_nom) && !is.null(scored_by_key$wages_cpi)) {
    rn <- .nc_rule_real_nominal(scored_by_key$wages_nom$cur_delta,
                                scored_by_key$wages_cpi$cur_delta)
    if (!is.null(rn)) {
      nom_s <- scored_by_key$wages_nom
      cpi_s <- scored_by_key$wages_cpi
      ctx <- list(nom_delta = nom_s$cur_delta, cpi_delta = cpi_s$cur_delta)
      score <- max(abs(nom_s$z), abs(cpi_s$z)) * rn$boost
      spec <- list(key = "wages_nom",
                   label = "Real vs nominal wages",
                   unit = "%", thresholds = numeric())
      sentence <- .nc_render_sentence(spec, nom_s, "real_nominal", ctx)
      results[[length(results) + 1]] <- list(
        key = "wages_divergence", score = score,
        rule = "real_nominal", sentence = sentence
      )
    }
  }

  if (length(results) == 0) return(list())

  # rank, dedupe so each key contributes at most 2 sentences, cap at max_n
  scores <- vapply(results, function(r) r$score, numeric(1))
  ord    <- order(scores, decreasing = TRUE)
  results <- results[ord]

  seen  <- list()
  final <- list()
  for (r in results) {
    k <- r$key
    seen[[k]] <- if (is.null(seen[[k]])) 1L else seen[[k]] + 1L
    if (seen[[k]] > 2) next
    final[[length(final) + 1]] <- r
    if (length(final) >= max_n) break
  }

  final
}
