# =============================================================================
#  KENYA HIV SERVICE REPORTING — INTERACTIVE MAP (KHIS / DHIS2)
#  Reporting performance of the 7 HIV service forms (Rev 2023), by county:
#    HIV Testing Services & PrEP | EMTCT | HIV & TB Treatment |
#    Medical Male Circumcision   | Post-Exposure Prophylaxis |
#    Key Populations             | Vulnerable Populations
#  Indicators : Reporting rate, Reporting rate on time, Actual reports,
#               Actual reports on time, Expected reports
#  Filters    : Service | Indicator | Reporting month (or last 12 months combined)
#  Output     : Self-contained interactive HTML map + Excel workbook
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


# ---- 1b. MAP SETTINGS --------------------------------------------------------
map_level   <- 2              # 2 = County, 3 = Sub-county
end_period  <- "202610"       # latest reporting month (Oct 2026)
n_months    <- 12             # months available in the time filter
rate_target <- 90             # reporting-rate target (%) shown on charts
map_title   <- "County-Level HIV Service Provision in Kenya"   # main title on the map

# All data AND county boundaries come from KHIS only (no basemap / tile services).
# Water bodies (Lake Victoria, Lake Turkana, other lakes, Indian Ocean) are drawn from
# Natural Earth physical layers (public domain, no attribution needed), downloaded once
# and cached in outputs/natural_earth. Set add_water <- FALSE to switch off.
add_water <- TRUE

out_dir  <- "outputs"
dir.create(out_dir, showWarnings = FALSE)
out_html <- file.path(out_dir, "KHIS_HIV_Map_Service_Reporting.html")
out_xlsx <- file.path(out_dir, "KHIS_HIV_Map_Service_Reporting_data.xlsx")

# Services = the 7 HIV reporting forms (Rev 2023). `regex` finds the KHIS dataset by
# name; paste a dataset UID into `uid` to force it.
services <- tribble(
  ~key,    ~service,                          ~regex,                              ~uid,
  "hts",   "HIV Testing Services & PrEP",     "731-1\\b.*2023",                    NA,
  "emtct", "EMTCT",                           "731-2\\b.*2023",                    NA,
  "art",   "HIV & TB Treatment",              "731-3\\b.*2023",                    NA,
  "vmmc",  "Medical Male Circumcision",       "731-4\\b.*2023",                    NA,
  "pep",   "Post-Exposure Prophylaxis",       "731-5\\b.*2023",                    NA,
  "kp",    "Key Populations",                 "Key Population.*2023",              NA,
  "vp",    "Vulnerable Populations",          "Vulnerable Population.*2023",       NA
)
services$uid <- as.character(services$uid)

metrics <- tribble(
  ~key,         ~label,                    ~type,
  "rate",       "Reporting rate",          "rate",
  "rate_ot",    "Reporting rate on time",  "rate",
  "actual",     "Actual reports",          "count",
  "actual_ot",  "Actual reports on time",  "count",
  "expected",   "Expected reports",        "count"
)

navy      <- "#1F4E79"
navy_ramp <- c("#DEEBF7", "#9ECAE1", "#4292C6", "#2171B5", "#1F4E79")

end_date <- as.Date(paste0(end_period, "01"), "%Y%m%d")
months   <- format(seq(end_date, by = "-1 month", length.out = n_months), "%Y%m")  # latest first
mlab     <- function(pe) format(as.Date(paste0(pe, "01"), "%Y%m%d"), "%b %Y")
all_lab  <- paste0("Last ", n_months, " months (", mlab(tail(months, 1)), " – ", mlab(months[1]), ")")

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
ou_dim  <- paste0("ou:", root_ou, ";LEVEL-", map_level)

# ---- 3. FIND THE 7 SERVICE DATASETS -----------------------------------------
ds_all <- dhis_get("dataSets", list(filter = "name:ilike:731",
                                    fields = "id,name,periodType", paging = "false"))$dataSets
ds_all <- if (NROW(ds_all)) as_tibble(ds_all) else tibble(id = character(), name = character(),
                                                           periodType = character())
ds_name <- character(nrow(services))
for (i in seq_len(nrow(services))) {
  if (!is.na(services$uid[i]) && nzchar(services$uid[i])) {
    ds_name[i] <- dhis_get(paste0("dataSets/", services$uid[i]), list(fields = "name"))$name
    next
  }
  hits <- ds_all |>
    filter(str_detect(name, regex(services$regex[i], ignore_case = TRUE))) |>
    arrange(desc(periodType == "Monthly"), nchar(name))
  if (nrow(hits)) { services$uid[i] <- hits$id[1]; ds_name[i] <- hits$name[1] }
}
services$dataset <- ds_name
cat("\n", strrep("=", 78), "\n  SERVICE DATASETS (KHIS)\n", strrep("=", 78), "\n", sep = "")
cat(sprintf("  %-30s %s  (%s)\n", services$service,
            ifelse(is.na(services$uid), "NOT FOUND", services$dataset),
            ifelse(is.na(services$uid), "-", services$uid)), sep = "")
cat("  If one is wrong/missing: paste its dataset UID into `uid` in the services table.\n",
    strrep("=", 78), "\n", sep = "")
if (any(is.na(services$uid))) warning("Not found and left out: ",
                                      paste(services$service[is.na(services$uid)], collapse = ", "))
services <- filter(services, !is.na(uid))
if (!nrow(services)) stop("None of the service datasets was found in KHIS.")

# ---- 4. BOUNDARIES (GEOSPATIAL LAYERS) --------------------------------------
message("Downloading boundaries (level ", map_level, ") ...")
ous <- dhis_get("organisationUnits",
                list(level = map_level, fields = "id,name,parent[name],geometry", paging = "false"),
                simplify = FALSE)$organisationUnits
ou_tbl <- tibble(ou_uid  = vapply(ous, function(o) o$id, character(1)),
                 ou_name = vapply(ous, function(o) o$name, character(1)),
                 parent  = vapply(ous, function(o) o$parent$name %||% "", character(1)))
features <- compact(lapply(ous, function(o) {
  g <- o$geometry
  if (is.null(g) || !g$type %in% c("Polygon", "MultiPolygon")) return(NULL)
  list(type = "Feature",
       properties = list(ou_uid = o$id, ou_name = o$name,
                         parent = if (is.null(o$parent$name)) "" else o$parent$name),
       geometry = g)
}))
if (length(features) < length(ous))
  warning(length(ous) - length(features), " org unit(s) have no boundary in KHIS and will not be drawn.")
if (!length(features)) stop("KHIS returned no polygon geometry for level ", map_level)
shp <- st_read(toJSON(list(type = "FeatureCollection", features = features),
                      auto_unbox = TRUE, digits = NA), quiet = TRUE)
message("Boundaries: KHIS (", length(features), " of ", length(ous), " org units)")
shp <- shp |>
  st_make_valid() |>
  st_simplify(preserveTopology = TRUE, dTolerance = 0.003) |>
  mutate(ou_name = str_squish(str_remove(ou_name, regex("\\s*County$", ignore_case = TRUE)))) |>
  arrange(ou_name)
ou_order <- shp$ou_uid
lab_pts  <- suppressWarnings(st_coordinates(st_point_on_surface(st_geometry(shp))))  # label positions

# Kenya outline + soft mask outside Kenya (makes the country stand out)
kenya   <- st_make_valid(st_union(st_geometry(shp)))
outline <- st_boundary(kenya)
kbb     <- st_bbox(shp)
mask    <- st_difference(st_as_sfc(st_bbox(c(xmin = kbb[["xmin"]] - 25, ymin = kbb[["ymin"]] - 25,
                                             xmax = kbb[["xmax"]] + 25, ymax = kbb[["ymax"]] + 25),
                                           crs = st_crs(shp))), kenya)

# ---- 4b. WATER: lakes + Indian Ocean (Natural Earth, public domain) ----------
lakes <- NULL; ocean <- NULL
if (add_water) {
  if (!requireNamespace("rnaturalearth", quietly = TRUE)) install.packages("rnaturalearth")
  ne_dir <- file.path(out_dir, "natural_earth"); dir.create(ne_dir, showWarnings = FALSE)
  ne_get <- function(type) {
    for (sc in c(10, 50)) {
      x <- tryCatch(suppressMessages(suppressWarnings(
             rnaturalearth::ne_load(scale = sc, type = type, category = "physical",
                                    destdir = ne_dir, returnclass = "sf"))),
           error = function(e) tryCatch(suppressMessages(suppressWarnings(
             rnaturalearth::ne_download(scale = sc, type = type, category = "physical",
                                        destdir = ne_dir, returnclass = "sf"))),
           error = function(e2) NULL))
      if (!is.null(x)) return(x)
    }
    warning("Could not load Natural Earth '", type, "' — map drawn without it."); NULL
  }
  water_box <- st_as_sfc(st_bbox(c(xmin = kbb[["xmin"]] - 6, ymin = kbb[["ymin"]] - 6,
                                   xmax = kbb[["xmax"]] + 6, ymax = kbb[["ymax"]] + 6),
                                 crs = st_crs(4326)))
  crop_to <- function(x) {
    if (is.null(x)) return(NULL)
    x <- st_make_valid(st_transform(x, 4326))
    x <- suppressWarnings(st_intersection(x, water_box))
    x <- x[!st_is_empty(x), ]
    if (!nrow(x)) return(NULL)
    st_simplify(x, preserveTopology = TRUE, dTolerance = 0.005)
  }
  message("Adding lakes and Indian Ocean ...")
  lakes <- crop_to(ne_get("lakes"))
  ocean <- crop_to(ne_get("ocean"))
  if (!is.null(lakes)) {
    nm <- intersect(c("name_en", "name", "NAME"), names(lakes))[1]
    lakes$lake_name <- if (is.na(nm)) "" else as.character(lakes[[nm]])
    lakes$lake_name[is.na(lakes$lake_name)] <- ""
    lakes <- lakes[, "lake_name"]
  }
}
# Text labels for water bodies and neighbouring countries (fixed positions)
geo_labels <- tribble(
  ~label,            ~lng,  ~lat,  ~type,
  "INDIAN OCEAN",    42.30, -3.40, "water",
  "Lake Victoria",   33.25, -0.75, "water",
  "Lake Turkana",    36.05,  3.55, "water",
  "UGANDA",          33.15,  1.75, "country",
  "SOUTH SUDAN",     33.40,  4.95, "country",
  "ETHIOPIA",        38.80,  5.05, "country",
  "SOMALIA",         42.55,  2.10, "country",
  "TANZANIA",        35.40, -3.75, "country"
)
message(nrow(shp), " polygons loaded.")

# ---- 5. PULL REPORTING DATA --------------------------------------------------
pull_reporting <- function(uid) {
  res <- dhis_get("analytics.json", list(
    dimension = c(paste0("dx:", paste0(uid, c(".ACTUAL_REPORTS", ".ACTUAL_REPORTS_ON_TIME",
                                              ".EXPECTED_REPORTS"), collapse = ";")),
                  paste0("pe:", paste(months, collapse = ";")),
                  ou_dim),
    skipMeta = "true"))
  if (!length(res$rows)) return(tibble(dx = character(), pe = character(),
                                       ou_uid = character(), value = numeric()))
  df <- as.data.frame(res$rows, stringsAsFactors = FALSE); names(df) <- res$headers$name
  tibble(dx = df$dx, pe = df$pe, ou_uid = df$ou, value = as.numeric(df$value))
}

raw <- map_dfr(seq_len(nrow(services)), function(i) {
  message("Pulling ", services$service[i], " ...")
  pull_reporting(services$uid[i]) |> mutate(svc = services$key[i])
})
if (!nrow(raw)) stop("KHIS returned no reporting data for ", mlab(tail(months, 1)), " – ", mlab(months[1]))

sum_na <- function(x) if (all(is.na(x))) NA_real_ else sum(x, na.rm = TRUE)
pct    <- function(a, b) ifelse(!is.na(b) & b > 0, 100 * a / b, NA_real_)
add_rates <- function(d) d |>
  mutate(actual    = ifelse(!is.na(expected) & is.na(actual),    0, actual),
         actual_ot = ifelse(!is.na(expected) & is.na(actual_ot), 0, actual_ot),
         rate      = pct(actual, expected),
         rate_ot   = pct(actual_ot, expected))

county <- raw |>
  mutate(metric = case_when(str_ends(dx, "ACTUAL_REPORTS_ON_TIME") ~ "actual_ot",
                            str_ends(dx, "ACTUAL_REPORTS")         ~ "actual",
                            str_ends(dx, "EXPECTED_REPORTS")       ~ "expected")) |>
  group_by(svc, pe, ou_uid, metric) |>
  summarise(value = sum(value, na.rm = TRUE), .groups = "drop") |>
  pivot_wider(names_from = metric, values_from = value) |>
  complete(svc = services$key, pe = months, ou_uid = ou_order)
for (k in c("actual", "actual_ot", "expected")) if (!k %in% names(county)) county[[k]] <- NA_real_
county <- county |> filter(ou_uid %in% ou_order)

county_all <- county |>                                   # last n months combined
  group_by(svc, ou_uid) |>
  summarise(across(c(actual, actual_ot, expected), sum_na), .groups = "drop") |>
  mutate(pe = "ALL")
county <- bind_rows(county, county_all) |> add_rates()

national <- county |>
  group_by(svc, pe) |>
  summarise(across(c(actual, actual_ot, expected), sum_na), .groups = "drop") |>
  add_rates()

cat("\nNATIONAL REPORTING RATE —", mlab(months[1]), "\n")
print(national |> filter(pe == months[1]) |>
        left_join(select(services, svc = key, service), by = "svc") |>
        transmute(Service = service, Expected = expected, Actual = actual,
                  `Rate %` = round(rate, 1), `On time %` = round(rate_ot, 1)) |>
        as.data.frame())

# ---- 6. INTERACTIVE MAP ------------------------------------------------------
periods  <- c(months, "ALL")
p_labels <- c(mlab(months), all_lab)
area_lab <- ifelse(map_level == 2, "County", "Sub-county")
nv       <- function(v) unname(as.numeric(v))

js_v <- setNames(lapply(services$key, function(s) {
  setNames(lapply(metrics$key, function(m) {
    lapply(periods, function(p) {
      d <- county[county$svc == s & county$pe == p, ]
      nv(d[[m]][match(ou_order, d$ou_uid)])
    })
  }), metrics$key)
}), services$key)

js_nat <- setNames(lapply(services$key, function(s) {
  setNames(lapply(metrics$key, function(m) {
    d <- national[national$svc == s, ]
    nv(d[[m]][match(periods, d$pe)])
  }), metrics$key)
}), services$key)

js_data <- list(
  ous = ou_order, names = shp$ou_name, area = area_lab, title = map_title,
  lat = unname(lab_pts[, "Y"]), lng = unname(lab_pts[, "X"]),
  navy = navy, ramp = navy_ramp, target = rate_target, n_months = n_months,
  services = lapply(seq_len(nrow(services)), function(i)
    list(key = services$key[i], label = services$service[i])),
  metrics  = lapply(seq_len(nrow(metrics)), function(i)
    list(key = metrics$key[i], label = metrics$label[i], type = metrics$type[i])),
  periods  = lapply(seq_along(periods), function(i)
    list(pe = periods[i], label = p_labels[i],
         short = if (periods[i] == "ALL") "All" else format(as.Date(paste0(periods[i], "01"), "%Y%m%d"), "%b"))),
  v = js_v, nat = js_nat)

# Header card: latest-month national reporting rate per service
rag <- function(x) ifelse(is.na(x), "#BDBDBD", ifelse(x >= 95, "#1a9850", ifelse(x >= 90, "#91cf60",
              ifelse(x >= 80, "#fee08b", ifelse(x >= 50, "#fc8d59", "#d73027")))))
latest <- national |> filter(pe == months[1]) |>
  right_join(select(services, svc = key, service), by = "svc")
header_rows <- paste0(
  "<tr><td style='padding:2px 10px 2px 0'>", latest$service, "</td>",
  "<td style='text-align:right'><span style='display:inline-block;min-width:52px;text-align:center;",
  "padding:1px 6px;border-radius:10px;font-weight:700;background:", rag(latest$rate), ";color:",
  ifelse(!is.na(latest$rate) & (latest$rate < 50 | latest$rate >= 95), "white", "#222"), "'>",
  ifelse(is.na(latest$rate), "–", sprintf("%.1f%%", latest$rate)), "</span></td>",
  "<td style='text-align:right;padding-left:8px;color:#666;font-size:11px'>",
  ifelse(is.na(latest$expected), "",
         paste0(scales::comma(latest$actual, accuracy = 1), "/", scales::comma(latest$expected, accuracy = 1))),
  "</td></tr>", collapse = "")

title_html <- paste0(
  "<div class='kh-card' style='max-width:340px'>",
  "<div class='kh-h'>Kenya reporting rate</div>",
  "<div class='kh-sub' style='margin-bottom:6px'>All services · ", mlab(months[1]),
  " · rate (actual/expected reports)</div>",
  "<table style='font-size:12px;border-collapse:collapse'>", header_rows, "</table>",
  "<div style='font-size:10px;color:#999;margin-top:6px'>Source: KHIS (DHIS2) · extracted ",
  format(Sys.Date(), "%d %b %Y"), " · Target ", rate_target, "%</div></div>")

map_css <- "
.kh-card{background:rgba(255,255,255,.97);padding:10px 12px;border-radius:8px;
  box-shadow:0 2px 10px rgba(0,0,0,.18);font:12px 'Segoe UI',Arial,sans-serif;color:#222}
.kh-card label{display:block;font-size:10px;font-weight:700;text-transform:uppercase;
  letter-spacing:.04em;color:#666;margin:8px 0 3px}
.kh-card select{width:100%;padding:5px 6px;border:1px solid #c9d3df;border-radius:5px;
  font:12px 'Segoe UI',Arial;background:#fff;color:#222}
.kh-row{display:flex;gap:4px;align-items:center}
.kh-btn{border:1px solid #c9d3df;background:#f4f7fb;border-radius:5px;width:28px;height:28px;
  cursor:pointer;color:#1F4E79;font-weight:700}
.kh-btn:disabled{opacity:.35;cursor:default}
.kh-h{font-size:13px;font-weight:800;color:#1F4E79}
.kh-sub{font-size:11px;color:#666}
.kh-sw{display:inline-block;width:16px;height:11px;margin-right:6px;border-radius:2px;vertical-align:middle}
.kh-list td{padding:1px 6px 1px 0}
.leaflet-popup-content{margin:10px 12px}
.leaflet-container{background:#e9eff5}
.kh-title{position:absolute;top:12px;left:50%;transform:translateX(-50%);z-index:1000;
  pointer-events:none;text-align:center;background:rgba(255,255,255,.96);padding:9px 22px 8px;
  border-radius:10px;box-shadow:0 2px 12px rgba(0,0,0,.2);border-bottom:4px solid #1F4E79;
  font-family:'Segoe UI',Arial,sans-serif;max-width:44%;}
.kh-title .t1{font-size:20px;font-weight:800;color:#1F4E79;letter-spacing:.01em;line-height:1.2}
.kh-title .t2{font-size:13px;font-weight:600;color:#333;margin-top:3px}
.kh-title .t3{font-size:11px;color:#777;margin-top:1px}
.kh-lab{background:none;border:none}
.kh-lab div{transform:translate(-50%,-50%);white-space:nowrap;text-align:center;
  font:600 10px 'Segoe UI',Arial,sans-serif;color:#1b1b1b;pointer-events:none;
  text-shadow:-1px -1px 0 #fff,1px -1px 0 #fff,-1px 1px 0 #fff,1px 1px 0 #fff,0 0 3px #fff}
.kh-lab div span{display:block;font-weight:700;font-size:10px;color:#1F4E79}
.kh-chk{display:flex;align-items:center;gap:6px;margin-top:9px;font-size:12px;color:#333;cursor:pointer}
"

map_js <- "
function(el, x, D) {
  var map = this, ous = D.ous, names = D.names, NA_COL = '#D9D9D9';
  if (map.attributionControl) map.attributionControl.setPrefix('Source: KHIS (DHIS2) · NASCOP');
  var RB = [0, 50, 80, 90, 95], RC = ['#d73027', '#fc8d59', '#fee08b', '#91cf60', '#1a9850'];
  var RL = ['< 50%', '50 – 79%', '80 – 89%', '90 – 94%', '95% +'];
  var st = {s: 0, m: 0, p: 0}, ALL = D.periods.length - 1;
  function lyr(id) { return map.layerManager.getLayer('shape', id); }
  function ok(v) { return v !== null && v !== undefined && !isNaN(v); }
  function M() { return D.metrics[st.m]; }
  function S() { return D.services[st.s]; }
  function isRate(m) { return (m || M()).type === 'rate'; }
  function fmt(v, m) {
    if (!ok(v)) return '–';
    return isRate(m) ? v.toFixed(1) + '%' : Math.round(v).toLocaleString('en-US');
  }
  function vals(s, m, p) { return D.v[D.services[s].key][D.metrics[m].key][p]; }
  function mi(key) { for (var i = 0; i < D.metrics.length; i++) if (D.metrics[i].key === key) return i; return -1; }
  var I_ACT = mi('actual'), I_AOT = mi('actual_ot'), I_EXP = mi('expected');
  function n0(v) { return ok(v) ? Math.round(v).toLocaleString('en-US') : '–'; }
  /* 'actual/expected' text for a rate metric; '' for count metrics */
  function frac(m, a_on, a, e) {
    var mm = D.metrics[m];
    if (mm.type !== 'rate' || !ok(e)) return '';
    return n0(mm.key === 'rate_ot' ? a_on : a) + '/' + n0(e);
  }
  function fracC(s, m, p, j) { return frac(m, vals(s, I_AOT, p)[j], vals(s, I_ACT, p)[j], vals(s, I_EXP, p)[j]); }
  function fracN(s, m, p)    { return frac(m, natv(s, I_AOT, p), natv(s, I_ACT, p), natv(s, I_EXP, p)); }
  function natv(s, m, p) { var a = D.nat[D.services[s].key][D.metrics[m].key]; return a ? a[p] : null; }
  function qbreaks(vs) {
    var v = vs.filter(ok).sort(function(a, b) { return a - b; });
    if (!v.length) return [];
    var br = [];
    for (var i = 0; i <= 5; i++) {
      var q = i / 5 * (v.length - 1), lo = Math.floor(q), hi = Math.ceil(q);
      br.push(v[lo] + (v[hi] - v[lo]) * (q - lo));
    }
    return br.filter(function(b, i, a) { return i === 0 || b > a[i - 1]; });
  }
  function colour(v, br) {
    if (!ok(v)) return NA_COL;
    if (isRate()) { var c = RC[0]; for (var i = 0; i < RB.length; i++) if (v >= RB[i]) c = RC[i]; return c; }
    if (br.length < 2) return D.ramp[D.ramp.length - 1];
    var nb = br.length - 1;
    for (var k = 1; k < br.length; k++)
      if (v <= br[k] || k === nb) return D.ramp[Math.round((k - 1) * (D.ramp.length - 1) / Math.max(1, nb - 1))];
  }

  /* ---------- trend chart (SVG) ---------- */
  function trend(series, selIdx, rate, w, h, clickable) {
    var n = series.length, mx = 0;
    series.forEach(function(v) { if (ok(v) && v > mx) mx = v; });
    if (rate) mx = Math.max(100, mx); if (!mx) mx = 1;
    var bw = w / n, s = '<svg width=\"' + w + '\" height=\"' + (h + 14) + '\" style=\"display:block\">';
    for (var i = 0; i < n; i++) {
      var v = series[i], bh = ok(v) ? Math.max(1, h * v / mx) : 0;
      var col = (i === selIdx) ? D.navy : (rate && ok(v) ? colourRate(v) : '#9ECAE1');
      s += '<rect ' + (clickable ? 'class=\"kh-bar\" data-i=\"' + i + '\" style=\"cursor:pointer\" ' : '') +
           'x=\"' + (i * bw + 1) + '\" y=\"' + (h - bh) + '\" width=\"' + Math.max(1, bw - 2) +
           '\" height=\"' + bh + '\" rx=\"2\" fill=\"' + col + '\"' +
           (i === selIdx ? ' stroke=\"#000\" stroke-width=\"1\"' : '') + '><title>' +
           D.periods[D.n_months - 1 - i].label + ': ' + fmt(v, {type: rate ? 'rate' : 'count'}) + '</title></rect>';
    }
    if (rate) {
      var ty = h - h * D.target / mx;
      s += '<line x1=\"0\" x2=\"' + w + '\" y1=\"' + ty + '\" y2=\"' + ty +
           '\" stroke=\"#333\" stroke-dasharray=\"3,2\" stroke-width=\"1\"/>';
    }
    s += '<text x=\"0\" y=\"' + (h + 12) + '\" font-size=\"10\" fill=\"#777\">' + D.periods[D.n_months - 1].label + '</text>';
    s += '<text x=\"' + w + '\" y=\"' + (h + 12) + '\" font-size=\"10\" fill=\"#777\" text-anchor=\"end\">' +
         D.periods[0].label + '</text></svg>';
    return s;
  }
  function colourRate(v) { var c = RC[0]; for (var i = 0; i < RB.length; i++) if (v >= RB[i]) c = RC[i]; return c; }
  function monthSeries(getter) {             /* oldest -> latest */
    var a = []; for (var p = D.n_months - 1; p >= 0; p--) a.push(getter(p)); return a;
  }

  /* ---------- hover highlight ---------- */
  ous.forEach(function(id) {
    var l = lyr(id); if (!l) return;
    l.on('mouseover', function() { this.setStyle({weight: 3, color: D.navy}); this.bringToFront(); });
    l.on('mouseout',  function() { this.setStyle({weight: 1, color: 'white'}); });
  });

  /* ---------- controls ---------- */
  function ctl(pos, cls, w) {
    var c = L.control({position: pos});
    c.onAdd = function() {
      var d = L.DomUtil.create('div', 'kh-card ' + (cls || ''));
      if (w) d.style.width = w + 'px';
      L.DomEvent.disableClickPropagation(d); L.DomEvent.disableScrollPropagation(d);
      return d;
    };
    c.addTo(map); return c.getContainer();
  }
  var fBox = ctl('topright', '', 270), lBox = ctl('bottomright', '', 210), tBox = ctl('bottomleft', '', 290);

  var opt = function(arr, lab) { return arr.map(function(o, i) { return '<option value=\"' + i + '\">' + o[lab] + '</option>'; }).join(''); };
  fBox.innerHTML = '<div class=\"kh-h\">Filters</div>' +
    '<label>Service</label><select id=\"kh-s\">' + opt(D.services, 'label') + '</select>' +
    '<label>Indicator</label><select id=\"kh-m\">' + opt(D.metrics, 'label') + '</select>' +
    '<label>Reporting period</label><div class=\"kh-row\">' +
    '<button class=\"kh-btn\" id=\"kh-prev\" title=\"Previous month\">&#9664;</button>' +
    '<select id=\"kh-p\" style=\"flex:1\">' + opt(D.periods, 'label') + '</select>' +
    '<button class=\"kh-btn\" id=\"kh-next\" title=\"Next month\">&#9654;</button></div>' +
    '<label class=\"kh-chk\" style=\"text-transform:none;letter-spacing:0;font-weight:400\">' +
    '<input type=\"checkbox\" id=\"kh-lab\" checked> Show ' + (D.area === 'County' ? 'county' : 'sub-county') + ' names</label>';
  var sS = fBox.querySelector('#kh-s'), sM = fBox.querySelector('#kh-m'), sP = fBox.querySelector('#kh-p');
  var bPrev = fBox.querySelector('#kh-prev'), bNext = fBox.querySelector('#kh-next');
  var cLab = fBox.querySelector('#kh-lab');

  /* ---------- title banner ---------- */
  var tEl = L.DomUtil.create('div', 'kh-title', el);
  function setTitle() {
    tEl.innerHTML = '<div class=\"t1\">' + D.title + '</div>' +
      '<div class=\"t2\">' + M().label + ' &nbsp;·&nbsp; ' + S().label + '</div>' +
      '<div class=\"t3\">' + D.periods[st.p].label + ' &nbsp;·&nbsp; Source: KHIS (DHIS2)</div>';
  }

  /* ---------- area labels ---------- */
  var labLayer = L.layerGroup().addTo(map), labMk = [];
  ous.forEach(function(id, j) {
    if (!ok(D.lat[j]) || !ok(D.lng[j])) { labMk.push(null); return; }
    var mk = L.marker([D.lat[j], D.lng[j]], {interactive: false, keyboard: false,
      icon: L.divIcon({className: 'kh-lab', html: '', iconSize: [0, 0]})});
    mk.addTo(labLayer); labMk.push(mk);
  });
  function setLabels(v) {
    var z = map.getZoom();
    labMk.forEach(function(mk, j) {
      if (!mk) return;
      /* only label areas big enough on screen for their name (avoids clutter) */
      var l = lyr(ous[j]), fits = true;
      if (l && l.getBounds) {
        var b = l.getBounds(), p1 = map.latLngToContainerPoint(b.getSouthWest()),
            p2 = map.latLngToContainerPoint(b.getNorthEast());
        fits = Math.abs(p2.x - p1.x) >= names[j].length * 5.2 && Math.abs(p1.y - p2.y) >= 14;
      }
      var fr = fracC(st.s, st.m, st.p, j);
      var html = fits ? '<div>' + names[j] + (z >= 7 ? '<span>' + fmt(v[j]) +
                 (fr ? ' <small style=\"font-weight:600;color:#555\">(' + fr + ')</small>' : '') + '</span>' : '') +
                 '</div>' : '';
      mk.setIcon(L.divIcon({className: 'kh-lab', html: html, iconSize: [0, 0]}));
    });
  }
  cLab.addEventListener('change', function() {
    if (cLab.checked) labLayer.addTo(map); else map.removeLayer(labLayer);
  });
  map.on('zoomend', function() { setLabels(vals(st.s, st.m, st.p)); });

  /* ---------- render ---------- */
  function render() {
    sS.value = st.s; sM.value = st.m; sP.value = st.p;
    bPrev.disabled = (st.p === ALL || st.p >= D.n_months - 1);
    bNext.disabled = (st.p === ALL || st.p <= 0);
    var v = vals(st.s, st.m, st.p), br = isRate() ? [] : qbreaks(v), plab = D.periods[st.p].label;
    setTitle(); setLabels(v);

    var ranked = v.map(function(val, j) { return {j: j, val: val}; }).filter(function(r) { return ok(r.val); })
                  .sort(function(a, b) { return b.val - a.val; });
    var rank = {}; ranked.forEach(function(r, k) { rank[r.j] = k + 1; });

    ous.forEach(function(id, j) {
      var l = lyr(id); if (!l) return;
      var val = v[j];
      l.setStyle({fillColor: colour(val, br), fillOpacity: 0.88, weight: 1, color: 'white'});
      var fr = fracC(st.s, st.m, st.p, j);
      var tip = '<b>' + names[j] + '</b><br>' + M().label + ': <b>' + fmt(val) + '</b>' +
                (fr ? '<br><span style=\"color:#555\">' + fr + ' reports (actual/expected)</span>' : '');
      if (l.getTooltip()) l.setTooltipContent(tip); else l.bindTooltip(tip, {sticky: true});

      var rows = '';
      D.metrics.forEach(function(mm, k) {
        var f2 = fracC(st.s, k, st.p, j);
        rows += '<tr' + (k === st.m ? ' style=\"background:#eef3f9;font-weight:700\"' : '') +
                '><td style=\"padding:2px 10px 2px 2px\">' + mm.label + '</td><td style=\"text-align:right;padding-right:2px\">' +
                fmt(vals(st.s, k, st.p)[j], mm) +
                (f2 ? ' <span style=\"color:#666;font-weight:400\">(' + f2 + ')</span>' : '') + '</td></tr>';
      });
      var spark = trend(monthSeries(function(p) { return vals(st.s, 0, p)[j]; }),
                        st.p === ALL ? -1 : D.n_months - 1 - st.p, true, 240, 46, false);
      var pop = '<div style=\"font:12px Segoe UI,Arial;width:250px\">' +
        '<div class=\"kh-h\" style=\"font-size:15px\">' + names[j] + '</div>' +
        '<div class=\"kh-sub\">' + S().label + ' · ' + plab + '</div>' +
        (rank[j] ? '<div class=\"kh-sub\" style=\"margin-top:2px\">Rank ' + rank[j] + ' of ' + ranked.length + ' on ' + M().label.toLowerCase() + '</div>' : '') +
        '<table style=\"width:100%;margin:6px 0;border-collapse:collapse\">' + rows + '</table>' +
        '<div class=\"kh-sub\" style=\"margin-bottom:2px\">Reporting rate, last ' + D.n_months + ' months (dashed = ' + D.target + '% target)</div>' +
        spark + '</div>';
      if (l.getPopup()) l.setPopupContent(pop); else l.bindPopup(pop, {maxWidth: 280});
    });

    /* legend */
    var h = '<div class=\"kh-h\">' + M().label + '</div><div class=\"kh-sub\" style=\"margin-bottom:6px\">' +
            S().label + '<br>' + plab + '</div>' +
            '<div style=\"margin-bottom:6px\">Kenya: <b style=\"font-size:14px;color:' + D.navy + '\">' +
            fmt(natv(st.s, st.m, st.p)) + '</b>' +
            (fracN(st.s, st.m, st.p) ? '<br><span class=\"kh-sub\">' + fracN(st.s, st.m, st.p) +
             ' reports (actual/expected)</span>' : '') + '</div>';
    if (isRate()) {
      for (var i = RC.length - 1; i >= 0; i--)
        h += '<div><i class=\"kh-sw\" style=\"background:' + RC[i] + '\"></i>' + RL[i] + '</div>';
    } else {
      for (var k = br.length - 1; k >= 1; k--)
        h += '<div><i class=\"kh-sw\" style=\"background:' + colour(br[k], br) + '\"></i>' +
             fmt(br[k - 1]) + ' – ' + fmt(br[k]) + '</div>';
      if (br.length === 1) h += '<div><i class=\"kh-sw\" style=\"background:' + D.ramp[D.ramp.length - 1] + '\"></i>' + fmt(br[0]) + '</div>';
    }
    h += '<div><i class=\"kh-sw\" style=\"background:' + NA_COL + '\"></i>No data / not expected</div>';
    lBox.innerHTML = h;

    /* trend + county list */
    var series = monthSeries(function(p) { return natv(st.s, st.m, p); });
    var t = '<div class=\"kh-h\">Kenya trend</div><div class=\"kh-sub\" style=\"margin-bottom:4px\">' +
            S().label + ' · ' + M().label + (isRate() ? ' · dashed = ' + D.target + '% target' : '') +
            ' · click a bar</div>' +
            trend(series, st.p === ALL ? -1 : D.n_months - 1 - st.p, isRate(), 266, 60, true);
    var list = ranked.slice();
    var lowFirst = isRate();
    if (lowFirst) list.reverse();
    list = list.slice(0, 5);
    t += '<div class=\"kh-h\" style=\"font-size:12px;margin-top:8px\">' +
         (lowFirst ? 'Lowest 5 ' : 'Highest 5 ') + data_area() + '</div><table class=\"kh-list\">';
    list.forEach(function(r) {
      var f3 = fracC(st.s, st.m, st.p, r.j);
      t += '<tr><td>' + names[r.j] + '</td><td style=\"text-align:right\"><b>' + fmt(r.val) + '</b></td>' +
           (f3 ? '<td style=\"text-align:right;color:#666\">' + f3 + '</td>' : '') + '</tr>';
    });
    t += '</table>';
    if (isRate() && ranked.length) {
      var met = ranked.filter(function(r) { return r.val >= D.target; }).length;
      t += '<div class=\"kh-sub\" style=\"margin-top:4px\"><b>' + met + ' of ' + ranked.length + '</b> ' +
           data_area().toLowerCase() + ' at or above ' + D.target + '%</div>';
    }
    tBox.innerHTML = t;
    tBox.querySelectorAll('.kh-bar').forEach(function(b) {
      b.addEventListener('click', function() { st.p = D.n_months - 1 - (+b.getAttribute('data-i')); render(); });
    });
  }
  function data_area() { return D.area === 'County' ? 'counties' : 'sub-counties'; }

  sS.addEventListener('change', function() { st.s = +sS.value; render(); });
  sM.addEventListener('change', function() { st.m = +sM.value; render(); });
  sP.addEventListener('change', function() { st.p = +sP.value; render(); });
  bPrev.addEventListener('click', function() { if (st.p < D.n_months - 1) { st.p++; render(); } });
  bNext.addEventListener('click', function() { if (st.p > 0) { st.p--; render(); } });
  render();
}"

fit_js <- sprintf(paste0("function(btn, map){ map.fitBounds([[%f,%f],[%f,%f]], ",
                          "{paddingTopLeft:[300,95], paddingBottomRight:[230,30]}); }"),
                  kbb[["ymin"]], kbb[["xmin"]], kbb[["ymax"]], kbb[["xmax"]])
north_html <- paste0("<div class='kh-card' style='padding:4px 8px;text-align:center;line-height:1'>",
                     "<div style='font-size:16px;color:", navy, "'>&#9650;</div>",
                     "<div style='font-weight:800;font-size:12px;color:", navy, "'>N</div></div>")

water_col <- "#BFD9EE"; water_line <- "#7FAFD6"; water_txt <- "#2F6E9F"
m <- leaflet(shp, options = leafletOptions(zoomSnap = 0.25, zoomDelta = 0.5, minZoom = 5)) |>
  addMapPane("bg",    zIndex = 390) |>     # mask + ocean   (below counties)
  addMapPane("water", zIndex = 420) |>     # lakes          (above counties)
  addMapPane("edge",  zIndex = 430) |>     # Kenya border
  addPolygons(data = mask, fillColor = "white", fillOpacity = 0.6, stroke = FALSE,
              options = pathOptions(pane = "bg", pointerEvents = "none"))
if (!is.null(ocean)) m <- m |>
  addPolygons(data = ocean, fillColor = water_col, fillOpacity = 1, color = water_line,
              weight = 0.8, options = pathOptions(pane = "bg", pointerEvents = "none"))
m <- m |>
  addPolygons(data = shp, layerId = ~ou_uid, fillColor = "#D9D9D9", fillOpacity = 0.88,
              color = "white", weight = 1)
if (!is.null(lakes)) m <- m |>
  addPolygons(data = lakes, fillColor = water_col, fillOpacity = 1, color = water_line,
              weight = 0.8, options = pathOptions(pane = "water", pointerEvents = "none"))
m <- m |>
  addPolylines(data = outline, color = navy, weight = 2.5, opacity = 1,
               options = pathOptions(pane = "edge", pointerEvents = "none")) |>
  addLabelOnlyMarkers(data = filter(geo_labels, type == "water"), lng = ~lng, lat = ~lat, label = ~label,
                      labelOptions = labelOptions(noHide = TRUE, textOnly = TRUE, direction = "center",
                        style = list("color" = water_txt, "font-style" = "italic", "font-size" = "12px",
                                     "font-weight" = "600", "letter-spacing" = "0.04em",
                                     "font-family" = "Segoe UI, Arial, sans-serif",
                                     "text-shadow" = "0 0 3px #fff, 0 0 3px #fff"))) |>
  addLabelOnlyMarkers(data = filter(geo_labels, type == "country"), lng = ~lng, lat = ~lat, label = ~label,
                      labelOptions = labelOptions(noHide = TRUE, textOnly = TRUE, direction = "center",
                        style = list("color" = "#8a94a0", "font-size" = "11px", "font-weight" = "700",
                                     "letter-spacing" = "0.12em",
                                     "font-family" = "Segoe UI, Arial, sans-serif"))) |>
  addControl(HTML(title_html), position = "topleft") |>
  addControl(HTML(north_html), position = "topleft") |>
  addScaleBar(position = "bottomleft", options = scaleBarOptions(imperial = FALSE)) |>
  addEasyButton(easyButton(icon = "fa-globe", title = "Reset view", onClick = JS(fit_js))) |>
  fitBounds(kbb[["xmin"]], kbb[["ymin"]], kbb[["xmax"]], kbb[["ymax"]],
            options = list(paddingTopLeft = c(300, 95), paddingBottomRight = c(230, 30))) |>
  setMaxBounds(kbb[["xmin"]] - 12, kbb[["ymin"]] - 10, kbb[["xmax"]] + 12, kbb[["ymax"]] + 10) |>
  htmlwidgets::prependContent(tags$style(HTML(map_css))) |>
  onRender(map_js, data = js_data)

print(m)   # view in RStudio Viewer

tryCatch(
  saveWidget(m, out_html, selfcontained = TRUE, title = "County-Level HIV Service Provision in Kenya"),
  error = function(e) {
    message("Pandoc not found — saving non-self-contained version (keep the _files folder with it).")
    saveWidget(m, out_html, selfcontained = FALSE, title = "County-Level HIV Service Provision in Kenya")
  })
message("Map saved: ", normalizePath(out_html))

# ---- 7. EXCEL WORKBOOK -------------------------------------------------------
wb <- createWorkbook()
hs <- createStyle(fontColour = "white", fgFill = navy, textDecoration = "bold",
                  halign = "center", valign = "center", wrapText = TRUE, border = "TopBottomLeftRight")
tot_style <- createStyle(textDecoration = "bold", fgFill = "#DEEBF7")
svc_name  <- setNames(services$service, services$key)
ou_name   <- setNames(shp$ou_name, shp$ou_uid)

# Sheet 1: latest month, one row per county, rate + on-time per service
latest_c <- county |> filter(pe == months[1])
latest_n <- national |> filter(pe == months[1]) |> mutate(ou_uid = "KENYA")
summ <- bind_rows(latest_c, latest_n) |>
  mutate(rate    = sprintf("%.1f", rate), rate_ot = sprintf("%.1f", rate_ot),
         reports = ifelse(is.na(expected), NA_character_,
                          paste0(format(actual, big.mark = ",", trim = TRUE), "/",
                                 format(expected, big.mark = ",", trim = TRUE))),
         across(c(rate, rate_ot), ~ ifelse(.x == "NA", NA_character_, .x))) |>
  select(svc, ou_uid, rate, rate_ot, reports) |>
  pivot_longer(c(rate, rate_ot, reports)) |>
  mutate(col = paste0(svc_name[svc], c(rate = " — Rate %", rate_ot = " — On time %",
                                       reports = " — Reports (actual/expected)")[name]),
         col = factor(col, levels = as.vector(t(outer(svc_name, c(" — Rate %", " — On time %",
                                                                   " — Reports (actual/expected)"), paste0))))) |>
  arrange(col) |>
  select(ou_uid, col, value) |>
  pivot_wider(names_from = col, values_from = value) |>
  mutate(County = ifelse(ou_uid == "KENYA", "KENYA", ou_name[ou_uid]), .before = 1) |>
  arrange(County == "KENYA", County) |>
  select(-ou_uid) |>
  mutate(across(contains("— Rate %") | contains("— On time %"), as.numeric))
names(summ)[1] <- area_lab
rate_cc <- which(grepl("Rate %|On time %", names(summ)))
sh1 <- paste("Reporting", format(end_date, "%b %Y"))
addWorksheet(wb, sh1)
writeData(wb, sh1, summ, headerStyle = hs)
addStyle(wb, sh1, createStyle(numFmt = "0.0"), rows = 2:(nrow(summ) + 1), cols = rate_cc, gridExpand = TRUE)
addStyle(wb, sh1, createStyle(halign = "center"), rows = 2:(nrow(summ) + 1),
         cols = setdiff(2:ncol(summ), rate_cc), gridExpand = TRUE)
addStyle(wb, sh1, tot_style, rows = nrow(summ) + 1, cols = seq_along(summ), gridExpand = TRUE, stack = TRUE)
for (cc in rate_cc) {
  conditionalFormatting(wb, sh1, cols = cc, rows = 2:(nrow(summ) + 1), type = "between",
                        rule = c(-1e9, 79.999), style = createStyle(bgFill = "#F8CBAD"))
  conditionalFormatting(wb, sh1, cols = cc, rows = 2:(nrow(summ) + 1), type = "between",
                        rule = c(80, 89.999), style = createStyle(bgFill = "#FFE699"))
  conditionalFormatting(wb, sh1, cols = cc, rows = 2:(nrow(summ) + 1), type = "between",
                        rule = c(90, 1e9), style = createStyle(bgFill = "#C6EFCE"))
}
setColWidths(wb, sh1, cols = 1, widths = 20)
setColWidths(wb, sh1, cols = 2:ncol(summ), widths = 14)
setRowHeights(wb, sh1, rows = 1, heights = 48)
freezePane(wb, sh1, firstActiveRow = 2, firstActiveCol = 2)

# Sheet 2: all data, long format (county + Kenya, every month + combined)
long <- bind_rows(county, mutate(national, ou_uid = "KENYA")) |>
  mutate(Service = svc_name[svc],
         Period  = ifelse(pe == "ALL", all_lab, mlab(pe)),
         Month   = pe,
         Area    = ifelse(ou_uid == "KENYA", "KENYA", ou_name[ou_uid])) |>
  arrange(Service, desc(Month), Area == "KENYA", Area) |>
  transmute(Service, Period, Month, Area,
            `Expected reports` = expected, `Actual reports` = actual,
            `Actual reports on time` = actual_ot,
            `Reporting rate %` = round(rate, 1), `Reporting rate on time %` = round(rate_ot, 1))
names(long)[names(long) == "Area"] <- area_lab
addWorksheet(wb, "All data")
writeData(wb, "All data", long, headerStyle = hs)
setColWidths(wb, "All data", cols = 1:ncol(long), widths = c(30, 26, 9, 20, 12, 12, 14, 12, 14))
freezePane(wb, "All data", firstRow = TRUE)
addFilter(wb, "All data", rows = 1, cols = 1:ncol(long))

# Sheet 3: datasets used
addWorksheet(wb, "Datasets")
writeData(wb, "Datasets", services |> transmute(Service = service, `KHIS dataset` = dataset, UID = uid,
                                                Months = paste(mlab(tail(months, 1)), "–", mlab(months[1]))),
          headerStyle = hs)
setColWidths(wb, "Datasets", cols = 1:4, widths = c(30, 70, 14, 22))

saveWorkbook(wb, out_xlsx, overwrite = TRUE)
message("Workbook saved: ", normalizePath(out_xlsx))
