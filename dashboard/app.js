/* OPS KPI Dashboard
 * Reads data/ops_kpi_data.json produced by OPS_KPI.py.
 * Displays only what is in that file: no hardcoded stores, no mock rows,
 * no thresholds beyond the targets the source data already carries.
 */

/* Apply the saved theme immediately so there is no flash of the wrong theme. */
(function () {
  try {
    var t = localStorage.getItem("opsKpiTheme");
    document.documentElement.setAttribute("data-theme", t === "dark" ? "dark" : "light");
  } catch (e) {
    document.documentElement.setAttribute("data-theme", "light");
  }
})();

(function () {
  "use strict";

  var DATA_URL = "../data/ops_kpi_data.json";

  var state = {
    payload: null,
    reportKey: null,
    granularity: null,
    periodKey: null,
    search: "",
    spoc: "",
    rowType: "all",
    onlyBelow: false,
    sortKey: null,
    sortDir: "asc",
    page: 1,
    pageSize: 50
  };

  var $ = function (id) { return document.getElementById(id); };

  /* ---------------- number formatting ---------------- */

  var nf = {
    decimal1: new Intl.NumberFormat("en-IN", { minimumFractionDigits: 1, maximumFractionDigits: 1 }),
    int: new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 })
  };

  function isUptReport(rep) {
    return Boolean(rep) && (/\bupt\b/i.test(String(rep.name || "")) ||
      /\bupt\b/i.test(String(rep._key || "")));
  }

  /* Display formatting only -- underlying values are never changed.
     UPT: exactly 1 decimal (max 1). Everything else: whole numbers,
     Indian digit grouping, no rupee symbol. Percent KPIs keep "%". */
  function formatValue(value, format, rep) {
    if (value === null || value === undefined) return null;
    if (typeof value !== "number") return String(value);
    if (!isFinite(value)) return null;
    rep = rep || (state.payload && state.reportKey ? report() : null);
    if (isUptReport(rep)) return nf.decimal1.format(value);
    var text = nf.int.format(value);
    if (text === "-0") text = "0";
    return /^percent/.test(format || "") ? text + "%" : text;
  }

  function railValue(rep, value) {
    if (value === null || value === undefined) return "\u2014";
    return formatValue(value, rep.value_format, rep);
  }

  /* ---------------- data helpers ---------------- */

  function report() { return state.payload.reports[state.reportKey]; }

  function metricColumns(rep) {
    return rep.columns.filter(function (c) { return c.role === "metric"; });
  }

  var HEADER_OVERRIDES = {
    store_code: "Store",
    store_name: "Store name",
    spoc_name: "SPoC"
  };

  function headerFor(col) {
    if (HEADER_OVERRIDES[col.key]) return HEADER_OVERRIDES[col.key];
    if (col.role === "target") return "Target";
    return col.label;   // the exact Excel header for every metric column
  }

  function columnByKey(rep, key) {
    for (var i = 0; i < rep.columns.length; i++) {
      if (rep.columns[i].key === key) return rep.columns[i];
    }
    return null;
  }

  function miniso(rep) {
    for (var i = 0; i < rep.data.length; i++) {
      if (rep.data[i]._row_type === "miniso_total") return rep.data[i];
    }
    return null;
  }

  function granularities(rep) {
    return ["day", "week", "month"].filter(function (g) {
      return rep.periods[g] && rep.periods[g].length;
    });
  }

  /* Periods are already ordered most-recent-first. The most recent one
     can still be genuinely empty (e.g. yesterday's feed hasn't landed
     yet) -- walk backward to the first one MINISO actually has a value
     for, rather than showing a blank "latest" period. Never invents a
     value: if every period is null, falls back to the newest one so
     the UI can show its existing "No data" state honestly. */
  function latestValidPeriod(rep, totalRow, granularity) {
    var periods = (rep.periods && rep.periods[granularity]) || [];
    if (!periods.length) return null;
    if (!totalRow) return periods[0];
    for (var i = 0; i < periods.length; i++) {
      var v = totalRow[periods[i].key];
      if (v !== null && v !== undefined) return periods[i];
    }
    return periods[0];
  }

  /* Status is derived only from data the report already carries:
     - a per-store target column (abv, upt) -> value vs that store's target
     - an achievement basis of 100 (sales_ach) -> value vs 100%
     Reports with neither are shown without any status colour. */
  function statusFor(rep, row, value) {
    if (value === null || value === undefined) return null;
    var basis = null;
    if (rep.target_column && row._row_type === "store") {
      basis = row[rep.target_column];
    } else if (rep.achievement_basis !== null && rep.achievement_basis !== undefined) {
      basis = rep.achievement_basis;
    }
    if (basis === null || basis === undefined) return null;
    var meets = rep.higher_is_better === false ? value <= basis : value >= basis;
    return { meets: meets, basis: basis };
  }

  function hasStatus(rep) {
    return Boolean(rep.target_column) ||
      (rep.achievement_basis !== null && rep.achievement_basis !== undefined);
  }

  /* Frontend-only: never render a column labelled as a "D-1" /
     "previous day" concept, whatever KPI it belongs to. */
  function isD1Label(text) {
    return /\bd-?1\b/i.test(text) || /previous\s*day/i.test(text);
  }

  function granularityOf(rep, key) {
    var gs = ["day", "week", "month"];
    for (var i = 0; i < gs.length; i++) {
      var list = rep.periods && rep.periods[gs[i]];
      if (list && list.some(function (p) { return p.key === key; })) return gs[i];
    }
    return null;
  }

  function isBlank(v) {
    return v === null || v === undefined || v === "" || (typeof v === "number" && isNaN(v));
  }

  /* #7: a store row is hidden only when EVERY displayed metric column
     is empty. One blank date never hides a store. All-zero rows are
     also treated as empty, except for reports where zero is a real,
     good answer (Negative Stock / lower-is-better). */
  function isEmptyStoreRow(rep, row) {
    if (row._row_type !== "store") return false;
    var cols = rep.columns.filter(function (c) {
      return c.role === "metric" && !isD1Label(c.label) && !isD1Label(c.key);
    });
    if (!cols.length) return false;
    var zeroIsValid = rep.higher_is_better === false || /negative/i.test(String(rep.name || ""));
    return cols.every(function (c) {
      var v = row[c.key];
      if (isBlank(v)) return true;
      return !zeroIsValid && typeof v === "number" && v === 0;
    });
  }

  /* #8: hide "(No SPoC) - TOTAL" rows. Display only. */
  var NO_SPOC_RE = /no[\s_\-]*spoc/i;
  function isNoSpocTotal(row) {
    if (row._row_type !== "spoc_total") return false;
    var spoc = row._spoc;
    if (spoc === null || spoc === undefined || String(spoc).trim() === "") return true;
    return NO_SPOC_RE.test([spoc, row.store_code, row.store_name].join(" "));
  }

  /* #4: label for total rows. "Ashish - TOTAL [SPoC total]" -> "Ashish". */
  function rowLabel(row) {
    var name = row.store_name, code = row.store_code, raw;
    if (row._row_type === "miniso_total") {
      raw = !isBlank(name) ? name : code;
      return isBlank(raw) ? "" : String(raw);
    }
    if (!isBlank(name) && /total/i.test(String(name))) raw = name;
    else if (!isBlank(code) && /total/i.test(String(code))) raw = code;
    else raw = !isBlank(name) ? name : (!isBlank(code) ? code : row._spoc);
    return String(isBlank(raw) ? "" : raw)
      .replace(/\s*[\[(]\s*SPoC\s*total\s*[\])]\s*$/i, "")
      .replace(/\s*[-\u2013\u2014:]?\s*SPoC\s*total\s*$/i, "")
      .replace(/\s*[-\u2013\u2014:]\s*total\s*$/i, "")
      .replace(/\s+total\s*$/i, "")
      .trim();
  }

  /* ---------------- filtering ---------------- */

  function filteredRows() {
    var rep = report();
    var term = state.search.trim().toLowerCase();

    return rep.data.filter(function (row) {
      if (isNoSpocTotal(row)) return false;
      if (isEmptyStoreRow(rep, row)) return false;
      if (state.rowType !== "all" && row._row_type !== state.rowType) return false;
      if (state.spoc && row._spoc !== state.spoc) return false;

      if (term) {
        var code = String(row.store_code || "").toLowerCase();
        var name = String(row.store_name || "").toLowerCase();
        if (code.indexOf(term) === -1 && name.indexOf(term) === -1) return false;
      }

      if (state.onlyBelow) {
        var status = statusFor(rep, row, row[state.periodKey]);
        if (!status || status.meets) return false;
      }
      return true;
    });
  }

  /* Interleave each SPoC's stores with that SPoC's own total row,
     immediately after its stores -- matches the source report's
     hierarchy (SPoC 1 stores, SPoC 1 total, SPoC 2 stores, SPoC 2
     total, ...), MINISO total last. Only re-groups; never reorders
     stores relative to each other within their own SPoC. */
  function groupBySpoc(rows) {
    var stores = rows.filter(function (r) { return r._row_type === "store"; });
    var spocTotals = rows.filter(function (r) { return r._row_type === "spoc_total"; });
    var minisoTotal = rows.filter(function (r) { return r._row_type === "miniso_total"; });
    if (!spocTotals.length) return rows;

    var order = [], bySpoc = {};
    stores.forEach(function (r) {
      var key = r._spoc || "";
      if (!bySpoc[key]) { bySpoc[key] = []; order.push(key); }
      bySpoc[key].push(r);
    });
    var totalBySpoc = {};
    spocTotals.forEach(function (r) { totalBySpoc[r._spoc || ""] = r; });

    var out = [];
    order.forEach(function (key) {
      out = out.concat(bySpoc[key]);
      if (totalBySpoc[key]) out.push(totalBySpoc[key]);
    });
    // a SPoC total whose stores are all filtered out still shows,
    // just ahead of the MINISO total rather than vanishing
    spocTotals.forEach(function (r) {
      if (order.indexOf(r._spoc || "") === -1) out.push(r);
    });
    return out.concat(minisoTotal);
  }

  function sortRows(rows) {
    if (!state.sortKey) return rows;
    var key = state.sortKey;
    var dir = state.sortDir === "asc" ? 1 : -1;
    return rows.slice().sort(function (a, b) {
      var va = a[key], vb = b[key];
      // nulls always sink, whichever direction
      if (va === null || va === undefined) return (vb === null || vb === undefined) ? 0 : 1;
      if (vb === null || vb === undefined) return -1;
      if (typeof va === "number" && typeof vb === "number") return (va - vb) * dir;
      return String(va).localeCompare(String(vb)) * dir;
    });
  }

  /* ---------------- rendering ---------------- */

  function renderRail() {
    var list = $("report-list");
    list.innerHTML = "";
    var order = state.payload.metadata.report_order || Object.keys(state.payload.reports);

    order.forEach(function (key) {
      var rep = state.payload.reports[key];
      if (!rep) return;

      var total = miniso(rep);
      var scopeCol = null;
      var grans = granularities(rep);
      if (grans.length) {
        var g = grans.indexOf("day") >= 0 ? "day" : grans[0];
        scopeCol = rep.periods[g][0];
      }

      var li = document.createElement("li");
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "rail-item";
      btn.setAttribute("aria-current", key === state.reportKey ? "true" : "false");

      var name = document.createElement("span");
      name.className = "rail-name";
      name.textContent = rep.name;

      var value = document.createElement("span");
      value.className = "rail-value";
      value.textContent = total && scopeCol ? railValue(rep, total[scopeCol.key]) : "\u2014";

      var scope = document.createElement("span");
      scope.className = "rail-scope";
      scope.textContent = scopeCol ? "MINISO, " + scopeCol.label : "";

      btn.appendChild(name);
      btn.appendChild(value);
      btn.appendChild(scope);
      btn.addEventListener("click", function () { selectReport(key); });
      li.appendChild(btn);
      list.appendChild(li);
    });

    var meta = state.payload.metadata;
    $("refresh-line").textContent = "Data refreshed " + friendlyTime(meta.last_successful_refresh);
    var excluded = meta.excluded_stores || [];
    $("exclusion-line").textContent = excluded.length
      ? excluded.length + " stores excluded by the reporting rules"
      : "";
  }

  function friendlyTime(iso) {
    if (!iso) return "\u2014";
    var d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    return d.toLocaleString("en-IN", {
      day: "2-digit", month: "short", year: "numeric",
      hour: "2-digit", minute: "2-digit"
    });
  }

  function renderHeader() {
    var rep = report();
    var meta = state.payload.metadata;
    $("report-title").textContent = rep.name;
    $("report-desc").textContent = rep.description || "";
    $("meta-date").textContent = meta.reporting_date || "\u2014";
    $("meta-refresh").textContent = friendlyTime(meta.last_successful_refresh);

    var chip = $("meta-status");
    var refreshed = new Date(meta.last_successful_refresh);
    var ageHours = (Date.now() - refreshed.getTime()) / 36e5;
    var stale = isNaN(ageHours) ? false : ageHours > 36;
    chip.textContent = stale ? "Data is over a day old" : (meta.status || "success");
    chip.setAttribute("data-status", stale ? "stale" : "ok");

    $("foot-source").textContent =
      "Source: " + (rep.excel_file || "OPS KPI") +
      "  \u00B7  " + rep.row_count + " rows  \u00B7  schema " + meta.schema_version;
  }

  function renderSummary() {
    var rep = report();
    var total = miniso(rep);
    var grid = $("summary-cards");
    grid.innerHTML = "";

    if (!total) {
      $("summary-note").textContent = "This report has no MINISO overall row.";
      return;
    }

    var cards = [];
    ["day", "week", "month"].forEach(function (g) {
      if (rep.periods[g] && rep.periods[g].length) {
        var p = latestValidPeriod(rep, total, g);
        if (!p) return;
        var isStale = p.key !== rep.periods[g][0].key;
        cards.push({
          label: p.label,
          key: p.key,
          value: total[p.key],
          sub: isStale
            ? "Latest available \u2014 " + rep.periods[g][0].label + " has no data yet"
            : latestLabel(g)
        });
      }
    });

    // If this report's source data already carries a month-to-date
    // style field (any name containing "mtd" or "month to date"), show
    // it -- this is read directly from the MINISO row, never computed
    // here, so it only appears when the Python/JSON output defines it.
    var mtdCol = rep.columns.filter(function (c) {
      return c.key !== rep.target_column && /mtd|month.?to.?date/i.test(c.key + " " + c.label);
    })[0];
    if (mtdCol) {
      cards.splice(Math.min(1, cards.length), 0, {
        label: mtdCol.label,
        key: mtdCol.key,
        value: total[mtdCol.key],
        sub: "Month to date"
      });
    }

    var selected = columnByKey(rep, state.periodKey);
    var alreadyShown = cards.some(function (c) { return c.key === state.periodKey; });
    if (selected && !alreadyShown) {
      cards.unshift({
        label: selected.label,
        key: selected.key,
        value: total[selected.key],
        sub: "Selected period"
      });    }

    var storeCount = rep.row_counts_by_type.store;
    var spocCount = rep.row_counts_by_type.spoc_total;

    cards.forEach(function (card) {
      var el = document.createElement("div");
      el.className = "card";
      var text = formatValue(card.value, rep.value_format, rep);

      // Target / variance / status -- only when the report defines a
      // basis that is actually valid at the MINISO level. ABV/UPT carry
      // a per-store target that the source data leaves blank on total
      // rows by design, so no target is shown for those here rather
      // than guessing one.
      var status = null, basisText = "";
      if (text !== null && rep.achievement_basis !== null && rep.achievement_basis !== undefined) {
        status = statusFor(rep, total, card.value);
        var variance = card.value - rep.achievement_basis;
        basisText = "Target " + formatValue(rep.achievement_basis, rep.value_format, rep) +
          " \u00B7 Variance " + (variance >= 0 ? "+" : "") + formatValue(variance, rep.value_format, rep) +
          " \u00B7 " + (status && status.meets ? "Target achieved" : "Below target");
      } else if (rep.target_column) {
        basisText = "Target is store-level only \u2014 not defined for the MINISO total";
      }

      el.innerHTML =
        '<div class="card-label"></div>' +
        '<div class="card-value' + (text === null ? " null" : "") + '"></div>' +
        '<div class="card-basis"></div>' +
        '<div class="card-sub"></div>';
      el.querySelector(".card-label").textContent = card.label;
      var valueEl = el.querySelector(".card-value");
      if (status && text !== null) {
        var span = document.createElement("span");
        span.className = "flag " + (status.meets ? "flag-good" : "flag-bad");
        span.textContent = text;
        valueEl.appendChild(span);
      } else {
        valueEl.textContent = text === null ? "No data" : text;
      }
      el.querySelector(".card-basis").textContent = basisText;
      el.querySelector(".card-sub").textContent = card.sub;
      grid.appendChild(el);
    });

    var coverage = document.createElement("div");
    coverage.className = "card";
    coverage.innerHTML =
      '<div class="card-label">Stores in report</div>' +
      '<div class="card-value"></div>' +
      '<div class="card-sub"></div>';
    coverage.querySelector(".card-value").textContent = storeCount;
    coverage.querySelector(".card-sub").textContent = spocCount + " SPoCs";
    grid.appendChild(coverage);

    $("summary-note").textContent =
      "MINISO figures come from the grand-total row the report itself produces. " +
      "They are not re-added from the store rows and are excluded from the table below.";
  }

  function latestLabel(granularity) {
    if (granularity === "day") return "Most recent completed day";
    if (granularity === "week") return "Most recent completed week";
    return "Most recent completed month";
  }

  function renderPeriodControls() {
    var rep = report();
    var segs = $("granularity");
    segs.innerHTML = "";

    granularities(rep).forEach(function (g) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.setAttribute("role", "tab");
      btn.setAttribute("aria-selected", g === state.granularity ? "true" : "false");
      btn.textContent = { day: "Days", week: "Weeks", month: "Months" }[g];
      btn.addEventListener("click", function () {
        state.granularity = g;
        var p = latestValidPeriod(rep, miniso(rep), g);
        state.periodKey = p ? p.key : rep.periods[g][0].key;
        state.page = 1;
        render();
      });
      segs.appendChild(btn);
    });

    var chips = $("periods");
    chips.innerHTML = "";
    (rep.periods[state.granularity] || []).forEach(function (p) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "chip";
      btn.setAttribute("aria-pressed", p.key === state.periodKey ? "true" : "false");
      btn.textContent = p.label;
      btn.addEventListener("click", function () {
        state.periodKey = p.key;
        state.page = 1;
        render();
      });
      chips.appendChild(btn);
    });
  }

  function renderFilters() {
    var rep = report();
    var select = $("spoc");
    var seen = {};
    var spocs = [];
    rep.data.forEach(function (row) {
      if (row._spoc && !seen[row._spoc]) { seen[row._spoc] = true; spocs.push(row._spoc); }
    });
    spocs.sort();

    if (select.getAttribute("data-report") !== state.reportKey) {
      select.innerHTML = "";
      var all = document.createElement("option");
      all.value = "";
      all.textContent = "All SPoCs";
      select.appendChild(all);
      spocs.forEach(function (s) {
        var opt = document.createElement("option");
        opt.value = s;
        opt.textContent = s;
        select.appendChild(opt);
      });
      select.setAttribute("data-report", state.reportKey);
    }
    select.value = state.spoc;

    $("search").value = state.search;
    $("rowtype").value = state.rowType;

    var below = $("only-below");
    below.checked = state.onlyBelow;
    below.disabled = !hasStatus(rep);
    below.parentElement.style.display = hasStatus(rep) ? "" : "none";
  }

  function truncate(text, max) {
    text = String(text === null || text === undefined ? "" : text);
    return text.length > max ? text.slice(0, max - 1) + "\u2026" : text;
  }

  function esc(text) {
    return String(text === null || text === undefined ? "" : text)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function renderTable(rows) {
    var rep = report();
    var headRow = $("thead-row");
    var body = $("tbody");
    headRow.innerHTML = "";
    body.innerHTML = "";

    function columnIsAllNull(rep, col) {
      return rep.data.every(function (row) {
        var v = row[col.key];
        return v === null || v === undefined;
      });
    }

    // #3: SPoC column is never shown (row._spoc still drives
    // filtering, grouping and totals).
    var visible = rep.columns.filter(function (c) {
      if (c.key === "spoc_name") return false;
      if (c.role === "metric" && columnIsAllNull(rep, c)) return false;
      if (isD1Label(c.label) || isD1Label(c.key)) return false;
      return true;
    });
    // Store code first, Store name second, everything else in source order.
    var codeCols = visible.filter(function (c) { return c.key === "store_code"; });
    var nameCols = visible.filter(function (c) { return c.key === "store_name"; });
    var otherCols = visible.filter(function (c) { return c.key !== "store_code" && c.key !== "store_name"; });
    visible = codeCols.concat(nameCols, otherCols);

    // #5: the frozen column is Store name (falls back to the first
    // column only if a report has no store_name column).
    var stickyKey = nameCols.length ? "store_name" : (visible[0] && visible[0].key);
    var labelKey = nameCols.length ? "store_name" : "store_code";

    visible.forEach(function (col) {
      var th = document.createElement("th");
      th.textContent = headerFor(col);
      if (!HEADER_OVERRIDES[col.key] && col.role !== "target" && headerFor(col) !== col.label) {
        th.title = col.label;
      }
      if (col.dtype !== "number") th.classList.add("text");
      if (col.key === "store_code") th.classList.add("col-code");
      if (col.key === "store_name") th.classList.add("col-name");
      if (col.key === stickyKey) th.classList.add("sticky-col");
      if (col.role === "metric") {
        var g = granularityOf(rep, col.key);   // #2 day / week / month header colour
        if (g) th.classList.add("g-" + g);
      }
      if (col.key === state.periodKey) th.classList.add("is-selected");
      th.setAttribute("scope", "col");
      th.tabIndex = 0;
      th.setAttribute("aria-sort",
        state.sortKey === col.key
          ? (state.sortDir === "asc" ? "ascending" : "descending")
          : "none");
      var sort = function () {
        if (state.sortKey === col.key) {
          state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
        } else {
          state.sortKey = col.key;
          state.sortDir = col.dtype === "number" ? "desc" : "asc";
        }
        render();
      };
      th.addEventListener("click", sort);
      th.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); sort(); }
      });
      headRow.appendChild(th);
    });

    // Default (unsorted) view: group each SPoC's stores with that
    // SPoC's own total immediately after them. Clicking a column to
    // sort intentionally switches to a flat sort across all rows.
    var ordered = state.sortKey ? rows : groupBySpoc(rows);

    var isFullscreen = document.querySelector(".table-block").classList.contains("is-fullscreen");
    var pageRows = ordered;
    if (!isFullscreen && state.pageSize > 0) {
      var start = (state.page - 1) * state.pageSize;
      pageRows = ordered.slice(start, start + state.pageSize);
    }

    pageRows.forEach(function (row) {
      var tr = document.createElement("tr");
      var isTotal = row._row_type === "spoc_total" || row._row_type === "miniso_total";
      if (row._row_type === "spoc_total") tr.className = "is-spoc-total";
      if (row._row_type === "miniso_total") tr.className = "is-miniso";

      visible.forEach(function (col) {
        var td = document.createElement("td");
        if (col.key === "store_code") td.classList.add("col-code");
        if (col.key === "store_name") td.classList.add("col-name");
        if (col.key === stickyKey) td.classList.add("sticky-col");
        if (col.key === state.periodKey) td.classList.add("is-selected");

        var value = row[col.key];
        if (col.dtype !== "number") {
          td.classList.add("text");
          if (isTotal && (col.key === "store_code" || col.key === "store_name")) {
            // totals: label lives in the frozen name column, code cell stays empty
            td.textContent = col.key === labelKey ? rowLabel(row) : "";
            if (col.key === labelKey && row._row_type === "miniso_total") {
              var tag = document.createElement("span");
              tag.className = "row-tag";
              tag.textContent = "overall";
              td.appendChild(tag);
            }
          } else {
            td.textContent = value === null || value === undefined ? "\u2014" : value;
            if (value === null || value === undefined) td.classList.add("null-cell");
          }
        } else {
          var text = formatValue(value, rep.value_format, rep);
          if (text === null) {
            td.textContent = "\u2014";
            td.classList.add("null-cell");
            td.title = "No data for this period";
          } else if (col.role === "metric") {
            var status = statusFor(rep, row, value);
            if (status) {
              var span = document.createElement("span");
              span.className = "flag " + (status.meets ? "flag-good" : "flag-bad");
              span.textContent = text;
              span.title = status.meets ? "At or above target" : "Below target";
              td.appendChild(span);
            } else {
              td.textContent = text;
            }
          } else {
            td.textContent = text;
          }
        }
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });

    $("empty").hidden = rows.length > 0;
    document.querySelector(".table-scroll").hidden = rows.length === 0;

    var totalRows = rows.length;
    $("table-count").textContent = totalRows + (totalRows === 1 ? " row" : " rows") +
      (state.rowType === "store" ? ", stores only"
        : state.rowType === "spoc_total" ? ", SPoC totals only"
          : ", stores and SPoC totals");

    // Pagination is bypassed entirely in full screen -- every filtered
    // row is shown at once, grouped, with no page split.
    var pages = (!isFullscreen && state.pageSize > 0)
      ? Math.max(1, Math.ceil(ordered.length / state.pageSize))
      : 1;
    $("page-info").textContent = isFullscreen
      ? "All " + totalRows + (totalRows === 1 ? " row" : " rows")
      : "Page " + state.page + " of " + pages;
    $("prev").disabled = isFullscreen || state.page <= 1;
    $("next").disabled = isFullscreen || state.page >= pages;
  }

  function render() {
    var rows = sortRows(filteredRows());
    var pages = state.pageSize > 0 ? Math.max(1, Math.ceil(rows.length / state.pageSize)) : 1;
    if (state.page > pages) state.page = pages;

    renderRail();
    renderHeader();
    renderSummary();
    renderPeriodControls();
    renderFilters();
    renderTable(rows);
  }

  function selectReport(key) {
    var rep = state.payload.reports[key];
    if (!rep) return;
    state.reportKey = key;
    var grans = granularities(rep);
    state.granularity = grans.indexOf("day") >= 0 ? "day" : (grans[0] || null);
    var total0 = miniso(rep);
    var initialPeriod = state.granularity ? latestValidPeriod(rep, total0, state.granularity) : null;
    state.periodKey = initialPeriod ? initialPeriod.key : null;
    state.sortKey = null;
    state.sortDir = "asc";
    state.spoc = "";
    state.onlyBelow = false;
    state.page = 1;
    render();
  }

  /* ---------------- wiring ---------------- */

  /* ---------------- theme (Light / Dark) ---------------- */

  function currentTheme() {
    return document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "light";
  }

  function setTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    try { localStorage.setItem("opsKpiTheme", theme); } catch (e) { /* storage blocked: still works this session */ }
    var btns = document.querySelectorAll("#theme-switch button");
    for (var i = 0; i < btns.length; i++) {
      btns[i].setAttribute("aria-pressed", btns[i].getAttribute("data-theme") === theme ? "true" : "false");
    }
  }

  function buildThemeSwitch() {
    if ($("theme-switch")) return;
    var wrap = document.createElement("div");
    wrap.id = "theme-switch";
    wrap.className = "theme-switch";
    wrap.setAttribute("role", "group");
    wrap.setAttribute("aria-label", "Colour theme");
    ["light", "dark"].forEach(function (t) {
      var b = document.createElement("button");
      b.type = "button";
      b.setAttribute("data-theme", t);
      b.textContent = t === "light" ? "Light" : "Dark";
      b.addEventListener("click", function () { setTheme(t); });
      wrap.appendChild(b);
    });
    var host = document.querySelector(".head");
    if (host) host.appendChild(wrap);
    else { wrap.classList.add("is-floating"); document.body.appendChild(wrap); }
    setTheme(currentTheme());
  }

  function bind() {
    buildThemeSwitch();
    var searchTimer;
    $("search").addEventListener("input", function (e) {
      clearTimeout(searchTimer);
      var value = e.target.value;
      searchTimer = setTimeout(function () {
        state.search = value;
        state.page = 1;
        render();
      }, 140);
    });

    $("spoc").addEventListener("change", function (e) {
      state.spoc = e.target.value; state.page = 1; render();
    });
    $("rowtype").addEventListener("change", function (e) {
      state.rowType = e.target.value; state.page = 1; render();
    });
    $("only-below").addEventListener("change", function (e) {
      state.onlyBelow = e.target.checked; state.page = 1; render();
    });
    $("page-size").addEventListener("change", function (e) {
      state.pageSize = parseInt(e.target.value, 10) || 0;
      state.page = 1; render();
    });
    $("prev").addEventListener("click", function () { state.page--; render(); });
    $("next").addEventListener("click", function () { state.page++; render(); });

    // Full screen bypasses pagination entirely (see renderTable) and
    // shows every filtered row grouped by SPoC -- no row-count math
    // needed here, just toggle the class and re-render.
    $("fullscreen-toggle").addEventListener("click", function () {
      var block = document.querySelector(".table-block");
      var isFull = block.classList.toggle("is-fullscreen");
      document.body.classList.toggle("has-fullscreen-table", isFull);
      this.textContent = isFull ? "Exit full screen" : "Full screen";
      render();
    });
    document.addEventListener("keydown", function (e) {
      var block = document.querySelector(".table-block");
      if (e.key === "Escape" && block.classList.contains("is-fullscreen")) {
        block.classList.remove("is-fullscreen");
        document.body.classList.remove("has-fullscreen-table");
        $("fullscreen-toggle").textContent = "Full screen";
        render();
      }
    });


    function clearFilters() {
      state.search = ""; state.spoc = ""; state.rowType = "all";
      state.onlyBelow = false; state.page = 1;
      render();
    }
    $("reset").addEventListener("click", clearFilters);
    $("empty-reset").addEventListener("click", clearFilters);
    $("retry").addEventListener("click", load);
  }

  /* ---------------- loading ---------------- */

  function showError(detail) {
    $("loading").hidden = true;
    $("app").hidden = true;
    $("error").hidden = false;
    $("error-detail").textContent = detail;
  }

  function checkPayload(payload) {
    if (!payload || typeof payload !== "object") return "The data file is not valid JSON.";
    if (!payload.metadata || !payload.reports) {
      return "The data file is missing its metadata or reports section.";
    }
    var keys = Object.keys(payload.reports).filter(function (k) {
      var r = payload.reports[k];
      return r && Array.isArray(r.data) && r.data.length && Array.isArray(r.columns);
    });
    if (!keys.length) return "The data file contains no usable reports yet.";

    // drop any report that arrived malformed rather than failing the whole page
    var clean = {};
    keys.forEach(function (k) { clean[k] = payload.reports[k]; clean[k]._key = k; });
    payload.reports = clean;
    payload.metadata.report_order = (payload.metadata.report_order || keys)
      .filter(function (k) { return clean[k]; });
    if (!payload.metadata.report_order.length) payload.metadata.report_order = keys;
    return null;
  }

  /* D-2 reporting rule: the single most recent DAY period (period
     index 1 -- what the source SQL already treats as "yesterday") is
     never shown in the dashboard; the latest displayed/selectable day
     is always the one before it. This is a pure display trim done
     once at load time: the dropped period's column is removed from
     each report's column list (so no table header/cell, no date
     chip, no summary card can ever reference it), but the row data
     itself and week/month periods are untouched -- nothing is deleted
     from the JSON, only what the UI iterates over. Entirely
     index-based, so no date is ever hardcoded: whichever day the
     source data considers freshest is always the one dropped. */
  function applyD2Cutoff(payload) {
    Object.keys(payload.reports || {}).forEach(function (key) {
      var rep = payload.reports[key];
      var days = rep.periods && rep.periods.day;
      if (!days || days.length < 2) return; // nothing left to fall back to
      var dropped = days.shift();
      rep.columns = rep.columns.filter(function (c) { return c.key !== dropped.key; });
    });
  }

  function load() {
    $("error").hidden = true;
    $("loading").hidden = false;

    fetch(DATA_URL, { cache: "no-store" })
      .then(function (res) {
        if (!res.ok) throw new Error("The server returned " + res.status + " for " + DATA_URL);
        return res.json();
      })
      .then(function (payload) {
        var problem = checkPayload(payload);
        if (problem) { showError(problem); return; }

        applyD2Cutoff(payload);
        state.payload = payload;
        var order = payload.metadata.report_order;
        $("loading").hidden = true;
        $("app").hidden = false;
        selectReport(order[0]);
      })
      .catch(function (err) {
        showError(
          err instanceof SyntaxError
            ? "The data file could not be parsed as JSON."
            : err.message
        );
      });
  }

  bind();
  load();
})();