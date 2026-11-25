# =============================================================================
#  KHIS HIV INTERACTIVE SPATIAL MAP — KENYA
#  Indicators : Currently on ART (as at Oct 2026) | HTS tests done |
#               MOH 731 (2023 version) facilities reporting + reporting rate
#  Source     : KHIS (DHIS2) https://hiskenya.dha.go.ke  — Analytics API
#  Geometry   : KHIS org-unit boundaries (joined by org-unit UID, no name matching)
#  Output     : Self-contained interactive HTML map + Excel table
#  NASCOP — Strategic Information
# =============================================================================

# ---- 0. PACKAGES -------------------------------------------------------------
pkgs <- c("httr2", "jsonlite", "dplyr", "tidyr", "purrr", "stringr", "tibble",
          "sf", "leaflet", "htmlwidgets", "htmltools", "openxlsx", "scales")
options(repos = c(CRAN = "https://cloud.r-project.org"))
# Install into your personal library if the system library is not writable
user_lib <- Sys.getenv("R_LIBS_USER")
if (!dir.exists(user_lib)) dir.create(user_lib, recursive = TRUE, showWarnings = FALSE)
.libPaths(c(user_lib, .libPaths()))

new_pkgs <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(new_pkgs)) {
  message("Installing: ", paste(new_pkgs, collapse = ", "))
  install.packages(new_pkgs, lib = user_lib, dependencies = TRUE, type = getOption("pkgType"))
}
still_missing <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(still_missing)) stop("These packages failed to install: ",
                                paste(still_missing, collapse = ", "),
                                "\nRun install.packages(c('", paste(still_missing, collapse = "','"),
                                "')) and read the error it prints.")
invisible(lapply(pkgs, library, character.only = TRUE))
sf::sf_use_s2(FALSE)

# ---- 1. CONFIG ---------------------------------------------------------------
base_url <- "https://hiskenya.dha.go.ke"

# Credentials: leave blank -> a LOGIN POP-UP asks for username & password.
# (Optional) or fill below / set KHIS_USER, KHIS_PASS in ~/.Renviron.
my_user <- ""          # e.g. "ekirui"
my_pass <- ""          # your KHIS password
my_pat  <- ""          # optional KHIS Personal Access Token

if (file.exists("~/.Renviron")) readRenviron("~/.Renviron")
clean <- function(x) str_trim(gsub('^["\']|["\']$', "", x))   # drop stray quotes/spaces
khis_pat  <- clean(if (nzchar(my_pat))  my_pat  else Sys.getenv("KHIS_PAT"))
khis_user <- clean(if (nzchar(my_user)) my_user else Sys.getenv("KHIS_USER"))
khis_pass <-       if (nzchar(my_pass)) my_pass else clean(Sys.getenv("KHIS_PASS"))
cat("\nKHIS_USER :", if (nzchar(khis_user)) paste0("'", khis_user, "'") else "NOT SET",
    "\nKHIS_PASS :", if (nzchar(khis_pass)) paste(nchar(khis_pass), "characters") else "NOT SET",
    "\nKHIS_PAT  :", if (nzchar(khis_pat)) "set" else "not set", "\n")

# Login pop-up (RStudio); falls back to console prompt outside RStudio
ask_creds <- function(default_user = "") {
  if (!requireNamespace("rstudioapi", quietly = TRUE))
    install.packages("rstudioapi", lib = user_lib)
  has_rs <- requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()
  if (has_rs) {
    u <- rstudioapi::showPrompt("KHIS login", "KHIS username:", default = default_user)
    p <- rstudioapi::askForPassword(paste0("KHIS password for '", u, "'"))
  } else {
    u <- readline("KHIS username: "); p <- readline("KHIS password: ")
  }
  if (is.null(u) || is.null(p) || !nzchar(u) || !nzchar(p)) stop("Login cancelled.")
  list(user = clean(u), pass = p)
}

if (!nzchar(khis_pat) && (!nzchar(khis_user) || !nzchar(khis_pass))) {
  cr <- ask_creds(khis_user); khis_user <- cr$user; khis_pass <- cr$pass
  cat("KHIS_USER :", paste0("'", khis_user, "'"), " | password:", nchar(khis_pass), "characters\n")
}

map_level     <- 2                 # 2 = County, 3 = Sub-county
stock_period  <- "202610"          # Currently on ART = snapshot "as at" Oct 2026
flow_periods  <- c("202610")       # HTS = Oct 2026 only
# For FY 2026/27 to date (Jul–Oct) use instead:
# flow_periods <- c("202607", "202608", "202609", "202610")

out_dir  <- "outputs"
dir.create(out_dir, showWarnings = FALSE)
out_html <- file.path(out_dir, "KHIS_HIV_Map_Oct2026.html")
out_xlsx <- file.path(out_dir, "KHIS_HIV_Map_Oct2026_data.xlsx")

# Indicator list
#   api_term : text sent to KHIS search (name ilike)
#   regex    : optional R filter on the returned names to shorten the list
#   uid      : leave NA = auto-pick best KHIS match; or paste a UID to force it.
#              Several UIDs separated by ";" are SUMMED (e.g. ANC + L&D + PNC).
#   prefer   : names matching this rank first in auto-pick (then shortest name)
indicators <- tribble(
  ~key,         ~label,                         ~period_type, ~api_term, ~regex,           ~prefer,                        ~uid,
  "art_curr",   "Currently on ART (Oct 2026)",  "stock",      "on ART",  "current",        "total|^HV03",                  NA,
  "hts_tested", "HTS: Total tests done",        "flow",       "Tested",  "HV01|HTS|total", "HV01.*total|tested.*total",    NA
  # To add later (put a comma after the row above):
  # "prep_new",   "PrEP: Newly initiated",      "flow",       "PrEP",    NA,               "new|initiat",                  NA,
  # "vtp_tested", "VTP: Pregnant/BF tested",    "flow",       "HV02",    "test",           "total",                        NA
)
indicators$uid   <- as.character(indicators$uid)
indicators$layer <- TRUE          # TRUE = shown as a map layer

# MOH 731 reporting — facilities that submitted the MOH 731 (2023 version) form
add_reporting  <- TRUE
ds_uid         <- NA              # paste the MOH 731 (2023) dataset UID to force it
ds_search      <- "731"           # dataset name search in KHIS
ds_prefer      <- "2023"          # pick the dataset whose name contains this
report_period  <- stock_period    # reporting month (Oct 2026)

navy      <- "#1F4E79"
navy_ramp <- c("#DEEBF7", "#9ECAE1", "#4292C6", "#2171B5", "#1F4E79")

# ---- 2. KHIS API HELPERS -----------------------------------------------------
auth_mode   <- NULL
cookie_file <- tempfile(fileext = ".cookies")

add_auth <- function(req, mode = auth_mode) {
  switch(mode,
    pat     = req_headers(req, Authorization = paste("ApiToken", khis_pat)),
    basic   = req_auth_basic(req, khis_user, khis_pass),
    session = req_cookie_preserve(req, cookie_file))
}

# Returns HTTP status of /api/me for a given auth mode
check_me <- function(mode) {
  request(base_url) |> req_url_path_append("api", "me") |>
    req_url_query(fields = "name") |> add_auth(mode) |>
    req_error(is_error = \(r) FALSE) |> req_timeout(60) |>
    req_perform() |> resp_status()
}

# Session (cookie) login — for KHIS setups that reject Basic auth
session_login <- function() {
  unlink(cookie_file)
  # DHIS2 2.41+ JSON login
  try(request(base_url) |> req_url_path_append("api", "auth", "login") |>
        req_body_json(list(username = khis_user, password = khis_pass)) |>
        req_cookie_preserve(cookie_file) |> req_error(is_error = \(r) FALSE) |>
        req_perform(), silent = TRUE)
  if (check_me("session") == 200) return(200)
  # Older DHIS2 form login
  try(request(base_url) |> req_url_path_append("dhis-web-commons-security", "login.action") |>
        req_body_form(j_username = khis_user, j_password = khis_pass) |>
        req_cookie_preserve(cookie_file) |> req_error(is_error = \(r) FALSE) |>
        req_perform(), silent = TRUE)
  check_me("session")
}

khis_login <- function() {
  res <- c()
  if (nzchar(khis_pat)) {
    res["Token (KHIS_PAT)"] <- check_me("pat")
    if (res[["Token (KHIS_PAT)"]] == 200) return(list(mode = "pat", res = res))
  }
  if (nzchar(khis_user)) {
    res["Basic auth"] <- check_me("basic")
    if (res[["Basic auth"]] == 200) return(list(mode = "basic", res = res))
    res["Session login"] <- session_login()
    if (res[["Session login"]] == 200) return(list(mode = "session", res = res))
  }
  list(mode = NULL, res = res)
}

dhis_get <- function(path, query = list(), simplify = TRUE) {
  request(base_url) |>
    req_url_path_append("api", path) |>
    req_url_query(!!!query, .multi = "explode") |>
    add_auth() |>
    req_timeout(600) |>
    req_retry(max_tries = 3) |>
    req_perform() |>
    resp_body_json(simplifyVector = simplify)
}

# Connection test
lg <- khis_login()
cat("\nLogin attempts (HTTP status):\n"); print(lg$res)
if (is.null(lg$mode) && interactive()) {          # wrong password -> pop-up again
  cat("\nLogin rejected — please re-enter your KHIS username/password.\n")
  khis_pat <- ""; cr <- ask_creds(khis_user); khis_user <- cr$user; khis_pass <- cr$pass
  lg <- khis_login(); print(lg$res)
}
if (is.null(lg$mode)) stop(
  "KHIS rejected every login method (401). The username/password are not being accepted.\n",
  "Confirm you can log in at ", base_url, " in a browser with the SAME username/password,\n",
  "then run again and re-enter them in the pop-up.")
auth_mode <- lg$mode
me <- dhis_get("me", list(fields = "name"))
cat("Connected to KHIS as:", me$name, " | method:", auth_mode, "\n\n")

root_ou <- dhis_get("organisationUnits", list(level = 1, fields = "id,name",
                                              paging = "false"))$organisationUnits$id[1]

# Search data elements + indicators by name
search_dx <- function(term) {
  q <- list(filter = paste0("name:ilike:", term), fields = "id,name", paging = "false")
  de  <- dhis_get("dataElements", q)$dataElements
  ind <- dhis_get("indicators",   q)$indicators
  bind_rows(
    if (length(de))  tibble(type = "dataElement", id = de$id,  name = de$name),
    if (length(ind)) tibble(type = "indicator",   id = ind$id, name = ind$name)
  )
}

# National value for many candidate items in ONE analytics call
national_values <- function(ids, pe) {
  out <- tryCatch({
    res <- dhis_get("analytics.json", list(
      dimension = paste0("dx:", paste(ids, collapse = ";")),
      filter    = c(paste0("ou:", root_ou), paste0("pe:", paste(pe, collapse = ";"))),
      skipMeta  = "true"))
    if (!length(res$rows)) tibble(id = character(), nat = numeric()) else {
      df <- as.data.frame(res$rows, stringsAsFactors = FALSE); names(df) <- res$headers$name
      tibble(id = df$dx, nat = as.numeric(df$value))
    }
  }, error = function(e) NULL)
  if (!is.null(out)) return(out)
  # one bad item breaks a batch -> fall back to one-by-one
  map_dfr(ids, function(i) tryCatch(national_values_one(i, pe), error = function(e) NULL))
}
national_values_one <- function(id, pe) {
  res <- dhis_get("analytics.json", list(dimension = paste0("dx:", id),
          filter = c(paste0("ou:", root_ou), paste0("pe:", paste(pe, collapse = ";"))), skipMeta = "true"))
  if (!length(res$rows)) return(NULL)
  df <- as.data.frame(res$rows, stringsAsFactors = FALSE); names(df) <- res$headers$name
  tibble(id = df$dx, nat = as.numeric(df$value))
}

auto_pick <- function(key, label, api_term, rgx, prefer, period_type) {
  cands <- search_dx(api_term)
  if (nrow(cands) && !is.na(rgx))
    cands <- filter(cands, str_detect(name, regex(rgx, ignore_case = TRUE)))
  if (!nrow(cands)) stop("No KHIS match for '", label, "' (search '", api_term,
                         "'). Change api_term/regex in the config.")
  cands <- cands |>
    mutate(pref = if (is.na(prefer)) FALSE else str_detect(name, regex(prefer, ignore_case = TRUE)),
           de   = type == "dataElement") |>
    arrange(desc(pref), desc(de), nchar(name), name) |>
    mutate(rank = row_number())
  pe <- if (period_type == "stock") stock_period else flow_periods
  top_ids <- head(cands$id, 40)
  nv <- bind_rows(lapply(split(top_ids, ceiling(seq_along(top_ids) / 10)), national_values, pe = pe))
  if (!nrow(nv) || !"id" %in% names(nv)) nv <- tibble(id = character(), nat = numeric())
  cands <- cands |>
    left_join(distinct(nv, id, .keep_all = TRUE), by = "id") |>
    mutate(has_data = !is.na(nat) & nat > 0) |>
    arrange(desc(has_data), rank)
  if (!cands$has_data[1]) stop(
    "None of the KHIS items for '", label, "' has data for period ", paste(pe, collapse = ","),
    ".\nEither KHIS analytics does not yet include that month (try the previous month in",
    " stock_period/flow_periods), or change api_term for this indicator.")
  cands$name <- ifelse(is.na(cands$nat), cands$name,
                       paste0(cands$name, "  [Kenya = ", scales::comma(cands$nat), "]"))
  csv <- file.path(out_dir, paste0("candidates_", key, ".csv"))
  write.csv(select(cands, type, id, name, nat), csv, row.names = FALSE)
  cat("\n", strrep("=", 75), "\n  ", label, " — AUTO-SELECTED:\n     ", cands$name[1],
      "  (", cands$id[1], ")\n  Other options (top 10; full list in ", csv, "):\n", sep = "")
  top <- head(cands[-1, ], 10)
  if (nrow(top)) cat(sprintf("     %s  (%s)\n", top$name, top$id), sep = "")
  cat("  If wrong: paste the right UID into the `uid` column of the config.\n", strrep("=", 75), "\n", sep = "")
  cands$id[1]
}

resolve_uid <- function(key, label, api_term, rgx, prefer, period_type, uid) {
  if (!is.na(uid) && nzchar(uid)) return(uid)
  auto_pick(key, label, api_term, rgx, prefer, period_type)
}

indicators$uid <- pmap_chr(
  list(indicators$key, indicators$label, indicators$api_term, indicators$regex,
       indicators$prefer, indicators$period_type, indicators$uid),
  resolve_uid
)

# ---- 2b. MOH 731 (2023) DATASET — facilities reporting ----------------------
if (add_reporting) {
  if (is.na(ds_uid) || !nzchar(ds_uid)) {
    ds <- dhis_get("dataSets", list(filter = paste0("name:ilike:", ds_search),
                                    fields = "id,name,periodType", paging = "false"))$dataSets
    if (!NROW(ds)) stop("No KHIS dataset found with '", ds_search, "' in its name. Change ds_search.")
    ds <- as_tibble(ds) |>
      mutate(pref    = str_detect(name, regex(ds_prefer, ignore_case = TRUE)),
             monthly = periodType == "Monthly") |>
      arrange(desc(pref), desc(monthly), nchar(name))
    cat("\n", strrep("=", 75), "\n  MOH 731 datasets found in KHIS:\n", sep = "")
    cat(sprintf("   %s %s  (%s, %s)\n", ifelse(seq_len(nrow(ds)) == 1, "->", "  "),
                ds$name, ds$id, ds$periodType), sep = "")
    if (!ds$pref[1]) warning("No dataset name contains '", ds_prefer, "'. Using: ", ds$name[1])
    ds_uid  <- ds$id[1]
    ds_name <- ds$name[1]
    cat("  If wrong: paste the right UID into ds_uid in the config.\n", strrep("=", 75), "\n", sep = "")
  } else {
    ds_name <- dhis_get(paste0("dataSets/", ds_uid), list(fields = "name"))$name
  }
  cat("MOH 731 dataset used:", ds_name, "(", ds_uid, ")\n")

  indicators <- bind_rows(indicators, tibble(
    key         = c("rep_actual", "rep_expected"),
    label       = c("MOH 731 (2023): facilities reporting", "MOH 731 (2023): facilities expected"),
    period_type = "report",
    uid         = paste0(ds_uid, c(".ACTUAL_REPORTS", ".EXPECTED_REPORTS")),
    layer       = c(TRUE, FALSE)))
}

# What is displayed (map layers / popup / title / Excel)
show <- indicators |> select(key, label, layer)
if (add_reporting) show <- bind_rows(show, tibble(
  key = "rep_rate", label = "MOH 731 (2023): reporting rate (%)", layer = TRUE))

# ---- 3. ORG UNITS + BOUNDARIES ----------------------------------------------
ou_dim  <- paste0("ou:", root_ou, ";LEVEL-", map_level)

message("Downloading boundaries (level ", map_level, ") ...")
ous <- dhis_get("organisationUnits",
                list(level = map_level, fields = "id,name,parent[name],geometry", paging = "false"),
                simplify = FALSE)$organisationUnits

features <- compact(lapply(ous, function(o) {
  g <- o$geometry
  if (is.null(g) || !g$type %in% c("Polygon", "MultiPolygon")) return(NULL)
  list(type = "Feature",
       properties = list(ou_uid = o$id,
                         ou_name = o$name,
                         parent  = if (is.null(o$parent$name)) "" else o$parent$name),
       geometry = g)
}))
if (!length(features)) stop("KHIS returned no polygon geometry for level ", map_level)

shp <- st_read(toJSON(list(type = "FeatureCollection", features = features),
                      auto_unbox = TRUE, digits = NA), quiet = TRUE) |>
  st_make_valid() |>
  st_simplify(preserveTopology = TRUE, dTolerance = 0.003) |>
  mutate(ou_name = str_squish(str_remove(ou_name, regex("\\s*County$", ignore_case = TRUE))))

message(nrow(shp), " polygons loaded (", length(ous) - nrow(shp), " org units without geometry).")

# ---- 4. PULL ANALYTICS -------------------------------------------------------
pull_values <- function(dx, pe) {
  if (!length(dx)) return(tibble(dx = character(), ou_uid = character(), value = numeric()))
  res <- dhis_get("analytics.json", list(
    dimension = c(paste0("dx:", paste(unique(dx), collapse = ";")), ou_dim),
    filter    = paste0("pe:", paste(pe, collapse = ";")),   # filter = summed over periods
    displayProperty = "NAME",
    skipMeta = "true"))
  if (!length(res$rows)) return(tibble(dx = character(), ou_uid = character(), value = numeric()))
  df <- as.data.frame(res$rows, stringsAsFactors = FALSE)
  names(df) <- res$headers$name
  tibble(dx = df$dx, ou_uid = df$ou, value = as.numeric(df$value))
}

dx_map <- indicators |>
  select(key, period_type, uid) |>
  separate_longer_delim(uid, ";") |>
  mutate(uid = str_trim(uid))

message("Pulling analytics ...")
pe_for <- list(stock = stock_period, flow = flow_periods, report = report_period)
vals <- map_dfr(unique(dx_map$period_type), function(pt)
  pull_values(dx_map$uid[dx_map$period_type == pt], pe_for[[pt]]) |>
    inner_join(filter(dx_map, period_type == pt), by = c("dx" = "uid")))

data_wide <- vals |>
  group_by(ou_uid, key) |>
  summarise(value = sum(value, na.rm = TRUE), .groups = "drop") |>
  pivot_wider(names_from = key, values_from = value)

for (k in indicators$key) if (!k %in% names(data_wide)) data_wide[[k]] <- NA_real_

map_df <- shp |>
  left_join(data_wide, by = "ou_uid") |>
  arrange(ou_name)

national <- map_df |> st_drop_geometry() |>
  summarise(across(all_of(indicators$key), ~ sum(.x, na.rm = TRUE)))

if (add_reporting) {
  map_df <- map_df |>
    mutate(rep_rate = ifelse(!is.na(rep_expected) & rep_expected > 0,
                             100 * rep_actual / rep_expected, NA_real_))
  national$rep_rate <- with(national, ifelse(rep_expected > 0, 100 * rep_actual / rep_expected, NA_real_))
}
cat("\nNATIONAL TOTALS — sanity check these against what you expect:\n")
print(as.data.frame(setNames(national[show$key], show$label)))

# ---- 5. INTERACTIVE MAP ------------------------------------------------------
fmt   <- function(x) ifelse(is.na(x), "No data", scales::comma(x, accuracy = 1))
fmt_k <- function(x, k) if (grepl("_rate$", k)) ifelse(is.na(x), "No data", sprintf("%.1f%%", x)) else fmt(x)
fmt_v <- function(vals, keys) unlist(Map(fmt_k, vals, keys))

make_pal <- function(x) {
  if (all(is.na(x))) return(NULL)
  if (length(unique(na.omit(x))) == 1) return(colorFactor(navy, domain = unique(na.omit(x)), na.color = "#D9D9D9"))
  br <- unique(quantile(x, probs = seq(0, 1, 0.2), na.rm = TRUE))
  if (length(br) < 3) return(colorNumeric(navy_ramp, x, na.color = "#D9D9D9"))
  br[1] <- floor(br[1]); br[length(br)] <- ceiling(br[length(br)])
  colorBin(navy_ramp, x, bins = br, pretty = FALSE, na.color = "#D9D9D9")
}

area_lab <- ifelse(map_level == 2, "County", "Sub-county")

popup_html <- map_df |> st_drop_geometry() |>
  pmap_chr(function(...) {
    r <- list(...)
    rows <- paste0(
      "<tr><td style='padding:2px 10px 2px 0'>", show$label, "</td>",
      "<td style='text-align:right'><b>", fmt_v(r[show$key], show$key), "</b></td></tr>",
      collapse = "")
    paste0("<div style='font-family:Segoe UI,Arial;font-size:12px'>",
           "<div style='color:", navy, ";font-size:14px;font-weight:700'>", r$ou_name, "</div>",
           if (map_level > 2) paste0("<div style='color:#666'>", r$parent, "</div>") else "",
           "<table style='margin-top:4px'>", rows, "</table></div>")
  })

per_lab <- if (length(flow_periods) == 1) {
  format(as.Date(paste0(flow_periods, "01"), "%Y%m%d"), "%b %Y")
} else {
  paste(format(as.Date(paste0(range(flow_periods), "01"), "%Y%m%d"), "%b %Y"), collapse = " – ")
}

title_html <- paste0(
  "<div style='background:white;padding:8px 12px;border-left:5px solid ", navy,
  ";font-family:Segoe UI,Arial;box-shadow:0 1px 4px rgba(0,0,0,.3)'>",
  "<div style='font-size:16px;font-weight:700;color:", navy, "'>Kenya HIV Programme — KHIS</div>",
  "<div style='font-size:11px;color:#555'>On ART as at ",
  format(as.Date(paste0(stock_period, "01"), "%Y%m%d"), "%b %Y"),
  " | HTS: ", per_lab,
  if (add_reporting) paste0(" | MOH 731 reporting: ",
                            format(as.Date(paste0(report_period, "01"), "%Y%m%d"), "%b %Y")) else "",
  "</div>",
  "<table style='font-size:12px;margin-top:4px'>",
  paste0("<tr><td style='padding-right:10px'>", show$label, "</td><td style='text-align:right'><b>",
         fmt_v(as.list(national[show$key]), show$key), "</b></td></tr>", collapse = ""),
  "</table><div style='font-size:10px;color:#888;margin-top:4px'>Source: KHIS (DHIS2), extracted ",
  format(Sys.Date(), "%d %b %Y"), "</div></div>")

m <- leaflet(map_df, options = leafletOptions(zoomSnap = 0.25, zoomDelta = 0.5)) |>
  addProviderTiles(providers$CartoDB.Positron, group = "Light") |>
  addProviderTiles(providers$OpenStreetMap,    group = "Street")

layers <- filter(show, layer)
for (i in seq_len(nrow(layers))) {
  k   <- layers$key[i]
  lab <- layers$label[i]
  v   <- map_df[[k]]
  pal <- make_pal(v)
  if (is.null(pal)) {
    warning("No county data returned for '", lab, "' — layer skipped. Check its UID/period.")
    next
  }
  m <- m |>
    addPolygons(
      fillColor = pal(v), fillOpacity = 0.8, color = "white", weight = 1,
      group = lab, popup = popup_html,
      label = lapply(sprintf("<b>%s</b><br>%s: %s", map_df$ou_name, lab, fmt_k(v, k)), HTML),
      labelOptions = labelOptions(textsize = "12px", direction = "auto"),
      highlightOptions = highlightOptions(weight = 3, color = navy, fillOpacity = 0.95,
                                          bringToFront = TRUE)) |>
    addLegend(pal = pal, values = v, title = paste0(lab, "<br><small>by ", area_lab, "</small>"),
              position = "bottomright", opacity = 0.9,
              className = paste0("info legend leg-", k),
              labFormat = if (grepl("_rate$", k)) labelFormat(suffix = "%", digits = 0)
                          else labelFormat(big.mark = ",", digits = 0))
}

m <- m |>
  addLayersControl(baseGroups = layers$label, overlayGroups = c("Light", "Street"),
                   options = layersControlOptions(collapsed = FALSE)) |>
  hideGroup("Street") |>
  addControl(HTML(title_html), position = "topleft") |>
  addScaleBar(position = "bottomleft") |>
  addEasyButton(easyButton(icon = "fa-globe", title = "Reset view",
                           onClick = JS("function(btn, map){ map.setView([0.2, 37.9], 6); }"))) |>
  setView(lng = 37.9, lat = 0.2, zoom = 6) |>
  onRender("
    function(el, x, data) {
      var map = this;
      function showLegend(name) {
        el.querySelectorAll('.info.legend').forEach(function(l){ l.style.display = 'none'; });
        var k = data[name];
        if (k) el.querySelectorAll('.leg-' + k).forEach(function(l){ l.style.display = 'block'; });
      }
      showLegend(Object.keys(data)[0]);
      map.on('baselayerchange', function(e){ showLegend(e.name); });
    }", data = setNames(as.list(layers$key), layers$label))

print(m)   # view in RStudio Viewer

tryCatch(
  saveWidget(m, out_html, selfcontained = TRUE, title = "Kenya HIV Map — KHIS"),
  error = function(e) {
    message("Pandoc not found — saving non-self-contained version (keep the _files folder with it).")
    saveWidget(m, out_html, selfcontained = FALSE, title = "Kenya HIV Map — KHIS")
  })
message("Map saved: ", normalizePath(out_html))

# ---- 6. EXCEL TABLE ----------------------------------------------------------
tbl <- map_df |> st_drop_geometry() |>
  select(ou_uid, ou_name, any_of("parent"), all_of(show$key))
if (map_level == 2) tbl <- select(tbl, -any_of("parent"))
names(tbl) <- c("OU UID", area_lab, if (map_level > 2) "County", show$label)
tot_row <- as_tibble(setNames(c(list(NA_character_, "KENYA"),
                                 if (map_level > 2) list(NA_character_),
                                 as.list(unlist(national[show$key]))), names(tbl)))
tbl <- bind_rows(tbl, tot_row)

wb <- createWorkbook()
addWorksheet(wb, "Map data")
hs <- createStyle(fontColour = "white", fgFill = navy, textDecoration = "bold",
                  halign = "center", wrapText = TRUE, border = "TopBottomLeftRight")
writeData(wb, 1, tbl, headerStyle = hs)
num_cols  <- which(names(tbl) %in% show$label)
rate_cols <- which(names(tbl) %in% show$label[grepl("_rate$", show$key)])
addStyle(wb, 1, createStyle(numFmt = "#,##0"), rows = 2:(nrow(tbl) + 1),
         cols = num_cols, gridExpand = TRUE)
addStyle(wb, 1, createStyle(textDecoration = "bold", fgFill = "#DEEBF7", numFmt = "#,##0"),
         rows = nrow(tbl) + 1, cols = seq_along(tbl), gridExpand = TRUE, stack = TRUE)
if (length(rate_cols)) addStyle(wb, 1, createStyle(numFmt = "0.0"), rows = 2:(nrow(tbl) + 1),
                                cols = rate_cols, gridExpand = TRUE, stack = TRUE)
setColWidths(wb, 1, cols = seq_along(tbl), widths = c(14, 20, if (map_level > 2) 18, rep(18, length(num_cols))))
freezePane(wb, 1, firstRow = TRUE)
addWorksheet(wb, "Indicator UIDs")
writeData(wb, 2, indicators |> mutate(periods = vapply(period_type, function(pt)
                                         paste(pe_for[[pt]], collapse = ";"), character(1))),
          headerStyle = hs)
saveWorkbook(wb, out_xlsx, overwrite = TRUE)
message("Table saved: ", normalizePath(out_xlsx))
