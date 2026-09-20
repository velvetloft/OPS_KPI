/* OPS KPI Dashboard
 * Reads data/ops_kpi_data.json produced by OPS_KPI.py.
 * Displays only what is in that file: no hardcoded stores, no mock rows,
 * no thresholds beyond the targets the source data already carries.
 */

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
    decimal2: new Intl.NumberFormat("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    decimal1: new Intl.NumberFormat("en-IN", { minimumFractionDigits: 1, maximumFractionDigits: 1 }),
    currency0: new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }),
    plain: new Intl.NumberFormat("en-IN", { maximumFractionDigits: 2 })
  };

  function formatValue(value, format) {
    if (value === null || value === undefined) return null;
    if (typeof value !== "number") return String(value);
    switch (format) {
      case "currency0": return "\u20B9" + nf.currency0.format(value);
      case "decimal2": return nf.decimal2.format(value);
      case "decimal1": return nf.decimal1.format(value);
      case "percent2": return nf.decimal2.format(value) + "%";
      default: return nf.plain.format(value);
    }
  }

  function compactCurrency(value) {
    if (value === null || value === undefined) return null;
    var abs = Math.abs(value);
    if (abs >= 1e7) return "\u20B9" + nf.decimal2.format(value / 1e7) + " Cr";
    if (abs >= 1e5) return "\u20B9" + nf.decimal2.format(value / 1e5) + " L";
    return "\u20B9" + nf.currency0.format(value);
  }

  function railValue(report, value) {
    if (value === null || value === undefined) return "\u2014";
    if (report.value_format === "currency0") return compactCurrency(value);
    return formatValue(value, report.value_format);
  }

  /* ---------------- data helpers ---------------- */

  function report() { return state.payload.reports[state.reportKey]; }

  function metricColumns(rep) {
    return rep.columns.filter(function (c) { return c.role === "metric"; });
  }

  var HEADER_OVERRIDES = {
    store_code: "Store code",
    store_name: "Store",
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

  /* ---------------- filtering ---------------- */

  function filteredRows() {
    var rep = report();
    var term = state.search.trim().toLowerCase();

    return rep.data.filter(function (row) {
      if (row._row_type === "miniso_total") return false;
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

    var selected = columnByKey(rep, state.periodKey);
    var alreadyShown = cards.some(function (c) { return c.key === state.periodKey; });
    if (selected && !alreadyShown) {
      cards.unshift({
        label: selected.label,
        key: selected.key,
        value: total[selected.key],
        sub: "Selected period"
      });
    }

    var storeCount = rep.row_counts_by_type.store;
    var spocCount = rep.row_counts_by_type.spoc_total;

    cards.forEach(function (card) {
      var el = document.createElement("div");
      el.className = "card";
      var text = formatValue(card.value, rep.value_format);

      // Target / variance / status -- only when the report defines a
      // basis that is actually valid at the MINISO level. ABV/UPT carry
      // a per-store target that the source data leaves blank on total
      // rows by design, so no target is shown for those here rather
      // than guessing one.
      var status = null, basisText = "";
      if (text !== null && rep.achievement_basis !== null && rep.achievement_basis !== undefined) {
        status = statusFor(rep, total, card.value);
        var variance = card.value - rep.achievement_basis;
        basisText = "Target " + formatValue(rep.achievement_basis, rep.value_format) +
          " \u00B7 Variance " + (variance >= 0 ? "+" : "") + formatValue(variance, rep.value_format) +
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

    var visible = rep.columns.filter(function (c) {
      if (c.key === "spoc_name") return state.rowType === "all" || state.rowType === "store";
      if (c.role === "metric" && columnIsAllNull(rep, c)) return false;
      return true;
    });

    visible.forEach(function (col, index) {
      var th = document.createElement("th");
      th.textContent = headerFor(col);
      if (headerFor(col) !== col.label) th.title = col.label;
      if (col.dtype !== "number") th.classList.add("text");
      if (index === 0) th.classList.add("sticky-col");
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

    var pageRows = rows;
    if (state.pageSize > 0) {
      var start = (state.page - 1) * state.pageSize;
      pageRows = rows.slice(start, start + state.pageSize);
    }

    pageRows.forEach(function (row) {
      var tr = document.createElement("tr");
      if (row._row_type === "spoc_total") tr.className = "is-spoc-total";
      if (row._row_type === "miniso_total") tr.className = "is-miniso";

      visible.forEach(function (col, index) {
        var td = document.createElement("td");
        if (index === 0) td.classList.add("sticky-col");
        if (col.key === state.periodKey) td.classList.add("is-selected");

        var value = row[col.key];
        if (col.dtype !== "number") {
          td.classList.add("text");
          td.textContent = value === null || value === undefined ? "\u2014" : value;
          if (index === 0 && row._row_type !== "store") {
            var tag = document.createElement("span");
            tag.className = "row-tag";
            tag.textContent = row._row_type === "miniso_total" ? "overall" : "SPoC total";
            td.appendChild(tag);
          }
          if (value === null || value === undefined) td.classList.add("null-cell");
        } else {
          var text = formatValue(value, col.role === "target"
            ? rep.value_format : rep.value_format);
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

    var pages = state.pageSize > 0 ? Math.max(1, Math.ceil(totalRows / state.pageSize)) : 1;
    $("page-info").textContent = "Page " + state.page + " of " + pages;
    $("prev").disabled = state.page <= 1;
    $("next").disabled = state.page >= pages;
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

  function bind() {
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

    $("fullscreen-toggle").addEventListener("click", function () {
      var block = document.querySelector(".table-block");
      var isFull = block.classList.toggle("is-fullscreen");
      document.body.classList.toggle("has-fullscreen-table", isFull);
      this.textContent = isFull ? "Exit full screen" : "Full screen";
    });
    document.addEventListener("keydown", function (e) {
      var block = document.querySelector(".table-block");
      if (e.key === "Escape" && block.classList.contains("is-fullscreen")) {
        block.classList.remove("is-fullscreen");
        document.body.classList.remove("has-fullscreen-table");
        $("fullscreen-toggle").textContent = "Full screen";
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
    keys.forEach(function (k) { clean[k] = payload.reports[k]; });
    payload.reports = clean;
    payload.metadata.report_order = (payload.metadata.report_order || keys)
      .filter(function (k) { return clean[k]; });
    if (!payload.metadata.report_order.length) payload.metadata.report_order = keys;
    return null;
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