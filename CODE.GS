/**
 * L2 Daily Data automation.
 *
 * Sheet contracts:
 * - Source sheets: Date/Time(A), Hour(B if present), Order Nr(C), AT(E), PT(F), PKT(G), Fulfillment Time(H), Item Count(I), Picker(J)
 * - RCA sheet: Complaint Date(A), Order(E), SKU ID(F), Defect Category(H), SKU Name(I),
 *   Complaint Reason(J), DA ID(K), DA Name(L), Picker(N), Customer ID(P), RCA/LM Bucket(Q)
 * - RDB manual inputs preserved by Date + Picker:
 *   Bag Deviation(K), Putaway(L)
 */

var CONFIG = {
  SOURCE_ID: "1JgoGkQI2UnY8YtLW9hIY-ltiM9-zGdpUhGp4VVC0Z3w",
  DEST_ID: "1fzm2CVWE8bJed1_OZOsBuMiAZz9E5f4p6QZFXk-0N2c",
  RCA_ID: "1FRvGqkmuA0irw8L5fJv8Q9aCoQH5ViC-eXys9XDAbMo",

  SHEETS: {
    RDB: "RDB",
    RDB_HOURLY: "RDB Hourly",
    MTD_DASHBOARD: "MTD_DASHBOARD",
    COMPLAINT_ANALYSIS: "COMPLAINT_ANALYSIS",
    RCA_DATABASE: "YEARLY_RCA_DATABASE",
    WEEKLY_ANALYTICS: "Weekly Analytics",
    DARK_STORE_GRAPHS: "Dark Store Graphs",
    DARK_STORE_DATA: "DARK_STORE_DATA"
  },

  PROPERTIES: {
    CLOSED_MTD_MONTH_KEY: "CLOSED_MTD_MONTH_KEY",
    MERCHANDISING_SCORES: "MERCHANDISING_SCORES"
  },

  MTD_PICKER_ALIASES: {
    "asik mia": "Saikat Sohhain Bhuiyan",
    "saikat sohhain bhuiyan": "Saikat Sohhain Bhuiyan",
    "tushar imran": "Emamul Hok",
    "emamul hok": "Emamul Hok",
    "yusuf chowdhury": "Sunbim Sarid Moon",
    "sunbim sarid moon": "Sunbim Sarid Moon",
    "mohammad tutul": "Md Apon Ahmed",
    "md apon ahmed": "Md Apon Ahmed",
    "zain hassan": "Jabbarul Islam",
    "jabbarul islam": "Jabbarul Islam"
  }
};

var GRAPH_LAYOUT = {
  DASHBOARD_MIN_ROWS: 130,
  DASHBOARD_MIN_COLUMNS: 14,
  DATA_MIN_ROWS: 220,
  DATA_MIN_COLUMNS: 12,
  ORDER_SOURCE_ROW: 1,
  ITEM_SOURCE_ROW: 35,
  PERFORMANCE_SOURCE_ROW: 70,
  DAILY_KPI_SOURCE_ROW: 105,
  TOP_PICKER_SOURCE_ROW: 125,
  PRODUCTIVITY_SOURCE_ROW: 145,
  PEAK_HOUR_SOURCE_ROW: 160,
  HEATMAP_SOURCE_ROW: 175,
  SUMMARY_ROW: 8,
  PRODUCTIVITY_ROW: 15,
  ORDERS_TABLE_ROW: 24,
  ITEMS_TABLE_ROW: 24,
  SPEED_TABLE_ROW: 53,
  VARIANCE_ROW: 82,
  SLA_ROW: 82,
  PEAK_HOUR_ROW: 96,
  PICKER_ROW: 105
};

const DASHBOARD_THEME = {
  BACKGROUND: "#0B1320",
  CARD: "#172033",
  SECTION: "#22304A",
  BORDER: "#344055",
  PRIMARY_TEXT: "#FFFFFF",
  SECONDARY_TEXT: "#CBD5E1",
  MUTED_TEXT: "#94A3B8",
  SUCCESS: "#A7F3D0",
  WARNING: "#D4A017",
  CRITICAL: "#FECACA",
  INFO: "#7C9BCF",
  NAVY: "#111827",
  BLUE: "#7C9BCF",
  CYAN: "#93C5FD",
  GREEN: "#355E4A",
  RED: "#6B3A3A",
  ORANGE: "#D4A017",
  PURPLE: "#D4A017",
  LIGHT: "#172033",
  WHITE: "#FFFFFF"
};

var MTD_COMPLAINT_SCORING = {
  SAFE_MAX_PERCENT: 0.4,
  WARNING_MAX_PERCENT: 0.7,
  HIGH_MAX_PERCENT: 1,
  SAFE_POINTS: 15,
  WARNING_PENALTY: -5,
  HIGH_PENALTY: -10,
  CRITICAL_PENALTY: -20
};

function runAll() {
  updateRawDatabase();
  updateRdbHourly();
  updateMTDDashboard();
  updateComplaintAnalysis();
  updateWeeklyAnalytics();
  updateDarkStoreDashboard();
}

function createTrigger() {
  return createAutomationTriggers();
}

function createAutomationTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "runAll") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("runAll")
    .timeBased()
    .everyMinutes(15)
    .create();

  ScriptApp.newTrigger("runAll")
    .forSpreadsheet(CONFIG.SOURCE_ID)
    .onEdit()
    .create();

  ScriptApp.newTrigger("runAll")
    .forSpreadsheet(CONFIG.SOURCE_ID)
    .onChange()
    .create();

  ScriptApp.newTrigger("runAll")
    .forSpreadsheet(CONFIG.DEST_ID)
    .onOpen()
    .create();

  return "Automation triggers created: every 15 minutes, source edit/change, and dashboard open.";
}

function closeMtdMonth() {
  var ss = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var rdbSheet = getRequiredSheet_(ss, CONFIG.SHEETS.RDB);
  var data = rdbSheet.getDataRange().getValues();
  var monthToClose = getLatestCompletedMonthKey_(data);

  if (!monthToClose) {
    throw new Error("No completed previous month found in RDB to close.");
  }

  setClosedMtdMonthKey_(monthToClose);
  updateMTDDashboard();

  return "Closed MTD month: " + monthToClose;
}

function reopenMtdMonth() {
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.PROPERTIES.CLOSED_MTD_MONTH_KEY);
  updateMTDDashboard();

  return "MTD month closure has been removed.";
}

function updateRawDatabase() {
  var sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_ID);
  var destSS = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var rcaSS = SpreadsheetApp.openById(CONFIG.RCA_ID);

  var destSheet = getRequiredSheet_(destSS, CONFIG.SHEETS.RDB);
  var rcaSheet = getRequiredSheet_(rcaSS, CONFIG.SHEETS.RCA_DATABASE);

  var existingManualInputs = readExistingRdbManualInputs_(destSheet);
  var existingManualRows = readExistingRdbManualRows_(destSheet);
  var complaints = buildComplaintIndex_(rcaSheet);
  var summary = buildRawSummary_(sourceSS);

  var output = [
    [
      "Date",
      "Picker Name",
      "",
      "Order Count",
      "Item Count",
      "AT",
      "PT",
      "PKT",
      "P95",
      "Complaints",
      "Bag Deviation",
      "Putaway"
    ]
  ];

  var dataRows = [];
  var keys = Object.keys(summary).sort(function(a, b) {
    return summary[a].dateKey === summary[b].dateKey
      ? summary[a].picker.localeCompare(summary[b].picker)
      : summary[a].dateKey.localeCompare(summary[b].dateKey);
  });

  keys.forEach(function(key) {
    var row = summary[key];
    var orderCount = Object.keys(row.orders).length;
    var manual = existingManualInputs[key] || { bagDeviation: 0, putaway: 0 };
    var complaintCount = complaints[key] ? Object.keys(complaints[key]).length : 0;

    dataRows.push([
      row.dateKey,
      row.picker,
      "",
      orderCount,
      row.itemCount,
      orderCount > 0 ? row.totalAcceptance / orderCount : 0,
      row.itemCount > 0 ? row.totalPicking / row.itemCount : 0,
      row.itemCount > 0 ? row.totalPacking / row.itemCount : 0,
      orderCount > 0 ? (row.totalFulfillment / orderCount) / 60 : 0,
      complaintCount,
      manual.bagDeviation,
      manual.putaway
    ]);
  });

  existingManualRows.forEach(function(manualRow) {
    if (!summary[manualRow.key]) {
      dataRows.push(manualRow.row);
    }
  });

  dataRows.sort(function(a, b) {
    var dateCompare = normalizeDate_(a[0]).localeCompare(normalizeDate_(b[0]));
    if (dateCompare !== 0) return dateCompare;

    return normalizeText_(a[1]).localeCompare(normalizeText_(b[1]));
  });

  output = output.concat(dataRows);

  clearSheetContents_(destSheet, 1, 1, Math.max(destSheet.getLastRow(), 1), 12);
  destSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function updateRdbHourly() {
  var sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_ID);
  var destSS = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var hourlySheet = getOrCreateSheet_(destSS, CONFIG.SHEETS.RDB_HOURLY);
  var manualMp = readExistingHourlyManualMp_(hourlySheet);
  var summary = buildHourlySummary_(sourceSS);
  var output = [
    [
      "Date",
      "Hours",
      "Order Count",
      "Item Count",
      "Avg Items/Order",
      "AT",
      "PT",
      "PKT",
      "P95",
      "MP"
    ]
  ];
  var dateKeys = Object.keys(summary).sort();
  var totalRows = [];
  var hourlyBlocks = [];

  dateKeys.forEach(function(dateKey) {
    var dateTotals = createHourlyBucket_();
    var mpTotal = 0;
    var mpCount = 0;
    var hourlyStartRow = output.length + 1;

    for (var hour = 0; hour < 24; hour++) {
      var bucket = summary[dateKey][hour];
      var orderCount = Object.keys(bucket.orders).length;
      var manualValue = manualMp[makeHourlyKey_(dateKey, hour)];

      addBucketToTotal_(dateTotals, bucket);

      if (manualValue !== "" && manualValue !== undefined) {
        var mpNumber = Number(manualValue);

        if (!isNaN(mpNumber)) {
          mpTotal += mpNumber;
          mpCount++;
        }
      }

      output.push([
        dateKey,
        hour,
        orderCount,
        bucket.itemCount,
        orderCount > 0 ? bucket.itemCount / orderCount : 0,
        orderCount > 0 ? bucket.totalAcceptance / orderCount : 0,
        bucket.itemCount > 0 ? bucket.totalPicking / bucket.itemCount : 0,
        bucket.itemCount > 0 ? bucket.totalPacking / bucket.itemCount : 0,
        orderCount > 0 ? (bucket.totalFulfillment / orderCount) / 60 : 0,
        manualValue === undefined ? "" : manualValue
      ]);
    }

    hourlyBlocks.push({
      startRow: hourlyStartRow,
      numRows: 24
    });

    var totalOrderCount = Object.keys(dateTotals.orders).length;
    totalRows.push(output.length + 1);

    output.push([
      dateKey,
      "Total",
      totalOrderCount,
      dateTotals.itemCount,
      totalOrderCount > 0 ? dateTotals.itemCount / totalOrderCount : 0,
      totalOrderCount > 0 ? dateTotals.totalAcceptance / totalOrderCount : 0,
      dateTotals.itemCount > 0 ? dateTotals.totalPicking / dateTotals.itemCount : 0,
      dateTotals.itemCount > 0 ? dateTotals.totalPacking / dateTotals.itemCount : 0,
      totalOrderCount > 0 ? (dateTotals.totalFulfillment / totalOrderCount) / 60 : 0,
      mpCount > 0 ? mpTotal / mpCount : ""
    ]);
  });

  clearSheetContents_(hourlySheet, 1, 1, Math.max(hourlySheet.getLastRow(), 1), Math.max(hourlySheet.getLastColumn(), 10));
  writePaddedRows_(hourlySheet, 1, 1, output);
  formatRdbHourlySheet_(hourlySheet, output, totalRows, hourlyBlocks);
}

function updateMTDDashboard() {
  var ss = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var rdbSheet = getRequiredSheet_(ss, CONFIG.SHEETS.RDB);
  var mtdSheet = getRequiredSheet_(ss, CONFIG.SHEETS.MTD_DASHBOARD);

  var data = rdbSheet.getDataRange().getValues();
  var existingMerchandising = readExistingMerchandisingScores_(mtdSheet);
  saveMerchandisingScores_(existingMerchandising);
  var mtdMonthKey = getActiveMtdMonthKey_(data);
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var dateKey = normalizeDate_(data[i][0]);
    var picker = getMtdPickerName_(data[i][1]);

    if (!dateKey || !picker || !isInMonth_(dateKey, mtdMonthKey)) continue;

    if (!summary[picker]) {
      summary[picker] = {
        activeDays: {},
        orders: 0,
        items: 0,
        complaints: 0,
        putaway: 0,
        atWeightedTotal: 0,
        ptWeightedTotal: 0,
        pktWeightedTotal: 0,
        p95WeightedTotal: 0,
        bagDevTotal: 0
      };
    }

    var orders = toNumber_(data[i][3]);
    var items = toNumber_(data[i][4]);
    var at = toNumber_(data[i][5]);
    var pt = toNumber_(data[i][6]);
    var pkt = toNumber_(data[i][7]);
    var p95 = toNumber_(data[i][8]);
    var complaints = toNumber_(data[i][9]);
    var bagDeviation = toNumber_(data[i][10]);
    var putaway = toNumber_(data[i][11]);

    summary[picker].activeDays[dateKey] = true;
    summary[picker].orders += orders;
    summary[picker].items += items;
    summary[picker].complaints += complaints;
    summary[picker].putaway += putaway;
    summary[picker].atWeightedTotal += at * orders;
    summary[picker].ptWeightedTotal += pt * items;
    summary[picker].pktWeightedTotal += pkt * items;
    summary[picker].p95WeightedTotal += p95 * orders;
    summary[picker].bagDevTotal += bagDeviation;
  }

  var output = [
    [
      "Picker Name",
      "MTD Orders",
      "MTD Items",
      "MTD Complaints",
      "Complaint %",
      "MTD Putaway",
      "Avg Orders/Day",
      "Avg Items/Day",
      "Avg MTD AT",
      "Avg MTD PT",
      "Avg MTD PKT",
      "Avg MTD P95",
      "Avg Bag Deviation",
      "Avg Putaway/Day",
      "Merchandising",
      "Complaint Score",
      "MTD Score",
      "Rank"
    ]
  ];

  var rows = [];

  Object.keys(summary).forEach(function(picker) {
    var r = summary[picker];
    var activeDayCount = Object.keys(r.activeDays).length || 1;
    var complaintPercent = r.orders > 0 ? (r.complaints / r.orders) * 100 : 0;

    var avgOrdersPerDay = r.orders / activeDayCount;
    var avgItemsPerDay = r.items / activeDayCount;
    var avgAT = r.orders > 0 ? r.atWeightedTotal / r.orders : 0;
    var avgPT = r.items > 0 ? r.ptWeightedTotal / r.items : 0;
    var avgPKT = r.items > 0 ? r.pktWeightedTotal / r.items : 0;
    var avgP95 = r.orders > 0 ? r.p95WeightedTotal / r.orders : 0;
    var avgBagDeviation = r.bagDevTotal / activeDayCount;
    var avgPutawayPerDay = r.putaway / activeDayCount;
    var merchandisingScore = existingMerchandising[picker] || 0;
    var complaintScore = calculateMtdComplaintScore_(complaintPercent);
    var finalScore = calculateMtdScore_({
      avgOrdersPerDay: avgOrdersPerDay,
      avgPutawayPerDay: avgPutawayPerDay,
      complaintScore: complaintScore,
      avgBagDeviation: avgBagDeviation,
      avgP95: avgP95,
      avgAT: avgAT,
      avgPT: avgPT,
      merchandisingScore: merchandisingScore
    });

    rows.push([
      picker,
      r.orders,
      r.items,
      r.complaints,
      round2_(complaintPercent),
      r.putaway,
      round2_(avgOrdersPerDay),
      round2_(avgItemsPerDay),
      round2_(avgAT),
      round2_(avgPT),
      round2_(avgPKT),
      round2_(avgP95),
      round2_(avgBagDeviation),
      round2_(avgPutawayPerDay),
      merchandisingScore,
      round2_(complaintScore),
      round2_(finalScore)
    ]);
  });

  rows.sort(function(a, b) {
    return b[16] - a[16] || String(a[0]).localeCompare(String(b[0]));
  });

  rows.forEach(function(row, index) {
    row.push(index + 1);
    output.push(row);
  });

  clearSheetContents_(mtdSheet, 1, 1, Math.max(mtdSheet.getLastRow(), 1), 18);
  mtdSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function updateComplaintAnalysis() {
  var rcaSS = SpreadsheetApp.openById(CONFIG.RCA_ID);
  var destSS = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var rcaSheet = getRequiredSheet_(rcaSS, CONFIG.SHEETS.RCA_DATABASE);
  var analysisSheet = getRequiredSheet_(destSS, CONFIG.SHEETS.COMPLAINT_ANALYSIS);
  var rdbSheet = getRequiredSheet_(destSS, CONFIG.SHEETS.RDB);
  var data = rcaSheet.getDataRange().getValues();
  var orderContext = getComplaintOrderContextFromRdb_(rdbSheet.getDataRange().getValues());
  var output = buildComplaintIntelligenceRows_(data, orderContext);

  analysisSheet.getRange(1, 1, analysisSheet.getMaxRows(), analysisSheet.getMaxColumns()).breakApart();
  analysisSheet.clear();
  ensureSheetSize_(analysisSheet, output.length + 2, getMaxRowWidth_(output) + 1);
  writePaddedRows_(analysisSheet, 1, 1, output);
  formatComplaintIntelligenceSheet_(analysisSheet, output);
}

function updateWeeklyAnalytics() {
  var sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_ID);
  var destSS = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var rcaSS = SpreadsheetApp.openById(CONFIG.RCA_ID);

  var weeklySheet = getOrCreateSheet_(destSS, CONFIG.SHEETS.WEEKLY_ANALYTICS);
  var rdbSheet = getRequiredSheet_(destSS, CONFIG.SHEETS.RDB);
  var rcaSheet = getRequiredSheet_(rcaSS, CONFIG.SHEETS.RCA_DATABASE);

  var sourceRecords = readSourceOrderRecords_(sourceSS);
  var rcaData = rcaSheet.getDataRange().getValues();
  var rdbData = rdbSheet.getDataRange().getValues();
  var weekRanges = getPreviousAndCurrentWeekRanges_();
  var blocks = weekRanges.map(function(weekRange) {
    var pickerPerformance = buildWeeklyPickerPerformance_(sourceRecords, weekRange);
    var complaints = buildWeeklyComplaintAnalytics_(rcaData, weekRange);
    var bagDeviation = buildWeeklyBagDeviation_(rdbData, weekRange);

    return {
      weekKey: weekRange.weekKey,
      rows: buildWeeklyAnalyticsBlock_(weekRange, pickerPerformance, complaints, bagDeviation)
    };
  });

  writeWeeklyAnalyticsBlocks_(weeklySheet, blocks);
}

function updateDarkStoreDashboard() {
  var ss = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var graphsSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.DARK_STORE_GRAPHS);
  var dataSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.DARK_STORE_DATA);
  var hourlySheet = getRequiredSheet_(ss, CONFIG.SHEETS.RDB_HOURLY);
  var mtdSheet = getRequiredSheet_(ss, CONFIG.SHEETS.MTD_DASHBOARD);
  var rdbSheet = getRequiredSheet_(ss, CONFIG.SHEETS.RDB);

  var allHourlyRows = getAllHourlyRows_(hourlySheet);
  var hourlyRows = getLatestHourlyRowsFromRows_(allHourlyRows);
  var latestDate = hourlyRows.length > 0 ? hourlyRows[0].dateKey : "";
  var comparisonDate = latestDate ? addDaysToDateKey_(latestDate, -7) : "";
  var comparisonHourlyRows = getHourlyRowsForDate_(allHourlyRows, comparisonDate);
  var mtdRows = getMtdDashboardRows_(mtdSheet);
  var orderMatrix = buildHourlyMetricComparisonMatrix_(allHourlyRows, "orders", latestDate, comparisonDate);
  var itemMatrix = buildHourlyMetricComparisonMatrix_(allHourlyRows, "items", latestDate, comparisonDate);
  var performanceRows = buildHourlyPerformanceRows_(hourlyRows);
  var dailyComparisonRows = buildDailyComparisonRows_(rdbSheet, latestDate, comparisonDate);
  var dailySummaryRows = buildDailyPerformanceSummaryRows_(dailyComparisonRows);
  var productivityRows = buildProductivityKpiRows_(dailyComparisonRows, hourlyRows, comparisonHourlyRows);
  var peakHourRows = buildPeakHourAnalysisRows_(hourlyRows, comparisonHourlyRows);
  var heatmapRows = buildHourlyHeatmapRows_(orderMatrix, itemMatrix, performanceRows);
  var speedComparisonRows = buildHourlySpeedComparisonRows_(hourlyRows, comparisonHourlyRows);

  resetDashboardSheet_(graphsSheet);
  prepareDarkStoreDataSheet_(dataSheet);
  renderGraphSourceTables_(dataSheet, orderMatrix, itemMatrix, performanceRows, dailyComparisonRows, mtdRows, productivityRows, peakHourRows, heatmapRows);

  renderGraphsHeader_(graphsSheet, latestDate, comparisonDate, dailySummaryRows);
  renderHeaderMetrics_(graphsSheet, latestDate, comparisonDate, dailyComparisonRows);
  renderDailyPerformanceSummary_(graphsSheet, dailySummaryRows);
  renderProductivityKpiCards_(graphsSheet, productivityRows);
  renderDashboardComparisonSections_(graphsSheet);
  renderHourlyMetricComparisonTable_(graphsSheet, GRAPH_LAYOUT.ORDERS_TABLE_ROW, 1, orderMatrix, "Orders");
  renderHourlyMetricComparisonTable_(graphsSheet, GRAPH_LAYOUT.ITEMS_TABLE_ROW, 8, itemMatrix, "Items");
  renderHourlySpeedComparisonTable_(graphsSheet, speedComparisonRows);
  renderVarianceAnalysisTable_(graphsSheet, dailyComparisonRows);
  renderSlaMonitor_(graphsSheet, dailyComparisonRows);
  renderPeakHourAnalysis_(graphsSheet, peakHourRows);
  renderTopPickerMtdRanking_(graphsSheet, mtdRows);
}

function resetDashboardSheet_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  ensureSheetSize_(sheet, GRAPH_LAYOUT.DASHBOARD_MIN_ROWS, GRAPH_LAYOUT.DASHBOARD_MIN_COLUMNS);
  sheet.setFrozenRows(6);
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(DASHBOARD_THEME.BACKGROUND);
  sheet.setColumnWidths(1, GRAPH_LAYOUT.DASHBOARD_MIN_COLUMNS, 126);
  sheet.setRowHeights(1, GRAPH_LAYOUT.DASHBOARD_MIN_ROWS, 34);
  sheet.getRange(1, 1, GRAPH_LAYOUT.DASHBOARD_MIN_ROWS, GRAPH_LAYOUT.DASHBOARD_MIN_COLUMNS)
    .setBackground(DASHBOARD_THEME.BACKGROUND)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontFamily("Arial");
}

function prepareGraphsSheet_(sheet) {
  resetDashboardSheet_(sheet);
}

function prepareDarkStoreDataSheet_(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  ensureSheetSize_(sheet, GRAPH_LAYOUT.DATA_MIN_ROWS, GRAPH_LAYOUT.DATA_MIN_COLUMNS);
  sheet.setHiddenGridlines(true);
  sheet.hideSheet();
}

function getLatestHourlyRows_(hourlySheet) {
  return getLatestHourlyRowsFromRows_(getAllHourlyRows_(hourlySheet));
}

function getAllHourlyRows_(hourlySheet) {
  var data = hourlySheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var dateKey = normalizeDate_(data[i][0]);
    var hour = data[i][1];

    if (!dateKey || hour === "Total" || hour === "" || hour === null || hour === undefined) continue;

    rows.push({
      dateKey: dateKey,
      hour: toNumber_(hour),
      orders: toNumber_(data[i][2]),
      items: toNumber_(data[i][3]),
      ipo: toNumber_(data[i][4]),
      at: toNumber_(data[i][5]),
      pt: toNumber_(data[i][6]),
      pkt: toNumber_(data[i][7]),
      p95: toNumber_(data[i][8]),
      mp: data[i][9]
    });
  }

  rows.sort(function(a, b) {
    return a.dateKey === b.dateKey
      ? a.hour - b.hour
      : a.dateKey.localeCompare(b.dateKey);
  });

  return rows;
}

function getLatestHourlyRowsFromRows_(allHourlyRows) {
  var latestDate = "";
  var rows = [];

  allHourlyRows.forEach(function(row) {
    if (row.dateKey > latestDate) latestDate = row.dateKey;
  });

  if (!latestDate) return rows;

  allHourlyRows.forEach(function(row) {
    if (row.dateKey === latestDate) rows.push(row);
  });

  rows.sort(function(a, b) {
    return a.hour - b.hour;
  });

  return rows;
}

function getHourlyRowsForDate_(allHourlyRows, dateKey) {
  if (!dateKey) return [];

  return allHourlyRows.filter(function(row) {
    return row.dateKey === dateKey;
  }).sort(function(a, b) {
    return a.hour - b.hour;
  });
}

function buildHourlyMetricMatrix_(hourlyRows, metricName) {
  var dates = [];
  var dateLookup = {};
  var matrixByHour = {};
  var output = [];

  hourlyRows.forEach(function(row) {
    if (!dateLookup[row.dateKey]) {
      dateLookup[row.dateKey] = true;
      dates.push(row.dateKey);
    }

    if (!matrixByHour[row.hour]) {
      matrixByHour[row.hour] = {};
    }

    matrixByHour[row.hour][row.dateKey] = row[metricName] || 0;
  });

  dates.sort();
  output.push(["Hour"].concat(dates));

  for (var hour = 0; hour < 24; hour++) {
    var rowValues = [hour];

    dates.forEach(function(dateKey) {
      rowValues.push(matrixByHour[hour] && matrixByHour[hour][dateKey] !== undefined
        ? matrixByHour[hour][dateKey]
        : 0);
    });

    output.push(rowValues);
  }

  return output;
}

function buildHourlyMetricComparisonMatrix_(hourlyRows, metricName, latestDate, comparisonDate) {
  var valuesByDateHour = {};
  var output = [[
    "Hour",
    comparisonDate ? comparisonDate + " Same Day Last Week" : "Same Day Last Week",
    latestDate ? latestDate + " Latest Day" : "Latest Day",
    "Variance"
  ]];

  hourlyRows.forEach(function(row) {
    if (row.dateKey !== latestDate && row.dateKey !== comparisonDate) return;

    if (!valuesByDateHour[row.dateKey]) {
      valuesByDateHour[row.dateKey] = {};
    }

    valuesByDateHour[row.dateKey][row.hour] = row[metricName] || 0;
  });

  for (var hour = 0; hour < 24; hour++) {
    var previousValue = valuesByDateHour[comparisonDate] && valuesByDateHour[comparisonDate][hour] !== undefined
      ? valuesByDateHour[comparisonDate][hour]
      : 0;
    var latestValue = valuesByDateHour[latestDate] && valuesByDateHour[latestDate][hour] !== undefined
      ? valuesByDateHour[latestDate][hour]
      : 0;

    output.push([
      hour,
      previousValue,
      latestValue,
      latestValue - previousValue
    ]);
  }

  return output;
}

function buildDailyComparisonRows_(rdbSheet, latestDate, comparisonDate) {
  var latest = summarizeRdbDate_(rdbSheet, latestDate);
  var previous = summarizeRdbDate_(rdbSheet, comparisonDate);
  var rows = [["KPI", comparisonDate || "Same Day Last Week", latestDate || "Latest Day", "Variance", "Variance %"]];

  [
    ["Orders", previous.orders, latest.orders],
    ["Items", previous.items, latest.items],
    ["Complaints", previous.complaints, latest.complaints],
    ["Complaint %", previous.complaintPercent, latest.complaintPercent],
    ["Avg AT", previous.avgAT, latest.avgAT],
    ["Avg PT", previous.avgPT, latest.avgPT],
    ["Avg PKT", previous.avgPKT, latest.avgPKT],
    ["Avg P95", previous.avgP95, latest.avgP95]
  ].forEach(function(row) {
    var variance = row[2] - row[1];
    var variancePercent = row[1] !== 0 ? (variance / row[1]) * 100 : "";

    rows.push([
      row[0],
      round2_(row[1]),
      round2_(row[2]),
      round2_(variance),
      variancePercent === "" ? "" : round2_(variancePercent)
    ]);
  });

  return rows;
}

function summarizeRdbDate_(rdbSheet, dateKey) {
  var data = rdbSheet.getDataRange().getValues();
  var summary = {
    orders: 0,
    items: 0,
    complaints: 0,
    complaintPercent: 0,
    avgAT: 0,
    avgPT: 0,
    avgPKT: 0,
    avgP95: 0
  };
  var atWeighted = 0;
  var ptWeighted = 0;
  var pktWeighted = 0;
  var p95Weighted = 0;

  if (!dateKey) return summary;

  for (var i = 1; i < data.length; i++) {
    var rowDate = normalizeDate_(data[i][0]);

    if (rowDate !== dateKey) continue;

    var orders = toNumber_(data[i][3]);
    var items = toNumber_(data[i][4]);
    var at = toNumber_(data[i][5]);
    var pt = toNumber_(data[i][6]);
    var pkt = toNumber_(data[i][7]);
    var p95 = toNumber_(data[i][8]);

    summary.orders += orders;
    summary.items += items;
    summary.complaints += toNumber_(data[i][9]);
    atWeighted += at * orders;
    ptWeighted += pt * items;
    pktWeighted += pkt * items;
    p95Weighted += p95 * orders;
  }

  summary.complaintPercent = summary.orders > 0 ? (summary.complaints / summary.orders) * 100 : 0;
  summary.avgAT = summary.orders > 0 ? atWeighted / summary.orders : 0;
  summary.avgPT = summary.items > 0 ? ptWeighted / summary.items : 0;
  summary.avgPKT = summary.items > 0 ? pktWeighted / summary.items : 0;
  summary.avgP95 = summary.orders > 0 ? p95Weighted / summary.orders : 0;

  return summary;
}

function buildHourlyPerformanceRows_(hourlyRows) {
  var rows = [["Hour", "AT", "PT", "PKT", "P95"]];

  hourlyRows.forEach(function(row) {
    rows.push([
      row.hour,
      round2_(row.at),
      round2_(row.pt),
      round2_(row.pkt),
      round2_(row.p95)
    ]);
  });

  return rows;
}

function getMtdDashboardRows_(mtdSheet) {
  var data = mtdSheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var picker = normalizeText_(data[i][0]);

    if (!picker) continue;

    rows.push({
      picker: picker,
      orders: toNumber_(data[i][1]),
      items: toNumber_(data[i][2]),
      complaints: toNumber_(data[i][3]),
      complaintPercent: toNumber_(data[i][4]),
      score: toNumber_(data[i][16]),
      rank: toNumber_(data[i][17])
    });
  }

  rows.sort(function(a, b) {
    return b.score - a.score || a.picker.localeCompare(b.picker);
  });

  return rows;
}

function renderGraphsHeader_(sheet, latestDate, comparisonDate, dailySummaryRows) {
  var refreshedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var overall = getOverallPerformanceStatus_(dailySummaryRows);

  sheet.getRange("A1:N1")
    .merge()
    .setValue("RUHDS48 OPERATIONS PERFORMANCE DASHBOARD")
    .setBackground(DASHBOARD_THEME.NAVY)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(24)
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("A2:N2")
    .merge()
    .setValue("Latest Available Date vs Same Day Last Week  |  " + buildExecutiveBannerText_(dailySummaryRows))
    .setBackground(DASHBOARD_THEME.NAVY)
    .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("A3:N3")
    .merge()
    .setValue("Overall Status: " + overall.label + "     Last Refresh Time: " + refreshedAt + "     Latest Date: " + (latestDate || "-") + "     Comparison Date: " + (comparisonDate || "-"))
    .setBackground(overall.color)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeights(1, 1, 45);
  sheet.setRowHeights(2, 1, 42);
  sheet.setRowHeights(3, 1, 40);
}

function renderHeaderMetrics_(sheet, latestDate, comparisonDate, dailyComparisonRows) {
  var orders = getDailyComparisonValue_(dailyComparisonRows, "Orders");
  var items = getDailyComparisonValue_(dailyComparisonRows, "Items");
  var metrics = [
    ["Last Refresh Time", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")],
    ["Latest Data Date", latestDate || "-"],
    ["Same Day Last Week Date", comparisonDate || "-"],
    ["Total Orders Today", formatKpiValue_(orders.today, "integer")],
    ["Total Orders Last Week", formatKpiValue_(orders.lastWeek, "integer")],
    ["Total Items Today", formatKpiValue_(items.today, "integer")],
    ["Total Items Last Week", formatKpiValue_(items.lastWeek, "integer")]
  ];

  metrics.forEach(function(metric, index) {
    var column = 1 + (index * 2);
    sheet.getRange(4, column, 1, 2)
      .merge()
      .setValue(metric[0])
      .setBackground(DASHBOARD_THEME.SECTION)
      .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
      .setFontWeight("bold")
      .setFontSize(10)
      .setFontFamily("Arial")
      .setHorizontalAlignment("center");
    sheet.getRange(5, column, 1, 2)
      .merge()
      .setValue(metric[1])
      .setBackground(DASHBOARD_THEME.CARD)
      .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
      .setFontWeight("bold")
      .setFontSize(12)
      .setFontFamily("Arial")
      .setHorizontalAlignment("center");
  });

  sheet.getRange("A4:N5").setBorder(true, true, true, true, true, true, DASHBOARD_THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeights(4, 1, 34);
  sheet.setRowHeights(5, 1, 40);
}

function renderDailyPerformanceSummary_(sheet, dailySummaryRows) {
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.SUMMARY_ROW - 1, 1, 14, "DAILY PERFORMANCE SUMMARY", DASHBOARD_THEME.NAVY);
  renderKpiCards_(sheet, GRAPH_LAYOUT.SUMMARY_ROW, dailySummaryRows, 2);
}

function renderProductivityKpiCards_(sheet, productivityRows) {
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.PRODUCTIVITY_ROW - 1, 1, 12, "PRODUCTIVITY KPIS", DASHBOARD_THEME.GREEN);
  renderKpiCards_(sheet, GRAPH_LAYOUT.PRODUCTIVITY_ROW, productivityRows, 3);
}

function renderDashboardComparisonSections_(sheet) {
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.ORDERS_TABLE_ROW - 2, 1, 7, "HOURLY ORDERS COMPARISON", DASHBOARD_THEME.BLUE);
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.ITEMS_TABLE_ROW - 2, 8, 7, "HOURLY ITEMS COMPARISON", DASHBOARD_THEME.GREEN);
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.SPEED_TABLE_ROW - 2, 1, 14, "HOURLY SPEED COMPARISON", DASHBOARD_THEME.WARNING);
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.VARIANCE_ROW - 2, 1, 7, "VARIANCE ANALYSIS TABLE", DASHBOARD_THEME.SECTION);
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.SLA_ROW - 2, 8, 7, "SLA MONITOR", DASHBOARD_THEME.ORANGE);
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.PEAK_HOUR_ROW - 2, 1, 14, "PEAK HOUR ANALYSIS", DASHBOARD_THEME.CYAN);
  renderDashboardSectionTitle_(sheet, GRAPH_LAYOUT.PICKER_ROW - 1, 1, 14, "TOP PICKER MTD RANKING", DASHBOARD_THEME.GREEN);
}

function renderDashboardSectionTitle_(sheet, row, column, numColumns, title, color) {
  sheet.getRange(row, column, 1, numColumns)
    .merge()
    .setValue(title)
    .setBackground(color)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(15)
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 38);
}

function renderKpiCards_(sheet, startRow, kpiRows, cardWidth) {
  kpiRows.forEach(function(kpi, index) {
    var column = 1 + (index * cardWidth);
    var cardRange = sheet.getRange(startRow, column, 5, cardWidth);
    var titleRange = sheet.getRange(startRow, column, 1, cardWidth);
    var todayRange = sheet.getRange(startRow + 1, column, 1, cardWidth);
    var lastWeekRange = sheet.getRange(startRow + 2, column, 1, cardWidth);
    var varianceRange = sheet.getRange(startRow + 3, column, 1, cardWidth);
    var statusRange = sheet.getRange(startRow + 4, column, 1, cardWidth);
    var cardColor = getKpiCardColor_(kpi.metric);

    titleRange.merge().setValue(kpi.metric);
    todayRange.merge().setValue("Today: " + kpi.todayText);
    lastWeekRange.merge().setValue("Last Week: " + kpi.lastWeekText);
    varianceRange.merge().setValue("Variance: " + kpi.varianceText);
    statusRange.merge().setValue("Status: " + kpi.indicator + " " + kpi.status);

    cardRange
      .setBackground(DASHBOARD_THEME.CARD)
      .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
      .setFontFamily("Arial")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, cardColor, SpreadsheetApp.BorderStyle.SOLID_THICK);

    titleRange
      .setBackground(DASHBOARD_THEME.SECTION)
      .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
      .setFontWeight("bold")
      .setFontSize(12);
    todayRange.setFontWeight("bold").setFontSize(14).setFontColor(DASHBOARD_THEME.PRIMARY_TEXT);
    lastWeekRange.setFontSize(12).setFontColor(DASHBOARD_THEME.SECONDARY_TEXT);
    varianceRange.setFontWeight("bold").setFontSize(13).setFontColor(kpi.color);
    applyComparisonCellStyle_(todayRange, kpi.status);
    statusRange
      .setBackground(kpi.background)
      .setFontColor(kpi.color)
      .setFontWeight("bold")
      .setFontSize(13);
  });

  sheet.setRowHeights(startRow, 5, 34);
}

function renderGraphSourceTables_(sheet, orderMatrix, itemMatrix, performanceRows, dailyComparisonRows, mtdRows, productivityRows, peakHourRows, heatmapRows) {
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.ORDER_SOURCE_ROW, 1, "Hourly Orders Source Data", orderMatrix, "#1f4e79");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.ITEM_SOURCE_ROW, 1, "Hourly Items Source Data", itemMatrix, "#548235");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.PERFORMANCE_SOURCE_ROW, 1, "Latest Date Speed / SLA Source Data", performanceRows, "#7030a0");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.DAILY_KPI_SOURCE_ROW, 1, "Daily KPI Source Data", dailyComparisonRows, "#0b1f3a");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.TOP_PICKER_SOURCE_ROW, 1, "Top Picker MTD Source Data", buildTopPickerSourceRows_(mtdRows), "#14532d");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.PRODUCTIVITY_SOURCE_ROW, 1, "Productivity KPI Source Data", kpiRowsToTable_(productivityRows), "#14532d");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.PEAK_HOUR_SOURCE_ROW, 1, "Peak Hour Analysis Source Data", peakHourRows, "#0369a1");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.HEATMAP_SOURCE_ROW, 1, "Hourly Heatmap Source Data", heatmapRows, "#991b1b");
}

function renderDashboardMatrix_(sheet, startRow, startColumn, title, values, color) {
  if (!values || values.length === 0) return;

  ensureSheetSize_(sheet, startRow + values.length + 2, startColumn + values[0].length + 1);
  sheet.getRange(startRow, startColumn, values.length + 1, values[0].length).clearContent();
  sheet.getRange(startRow, startColumn)
    .setValue(title)
    .setBackground(color || "#0b1f3a")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange(startRow + 1, startColumn, values.length, values[0].length).setValues(values);
}

function renderHourlyMetricComparisonTable_(sheet, startRow, startColumn, matrix, metricName) {
  var rows = [["Hour", "Today " + metricName, "Last Week " + metricName, "Difference", "Variance %"]];

  for (var i = 1; i < matrix.length; i++) {
    var lastWeek = toNumber_(matrix[i][1]);
    var today = toNumber_(matrix[i][2]);
    var difference = today - lastWeek;
    var variancePercent = lastWeek !== 0 ? (difference / lastWeek) * 100 : "";

    rows.push([
      matrix[i][0],
      today,
      lastWeek,
      difference,
      variancePercent === "" ? "" : formatVarianceText_(variancePercent)
    ]);
  }

  sheet.getRange(startRow, startColumn, rows.length, rows[0].length).setValues(rows);
  formatDashboardTable_(sheet, startRow, startColumn, rows.length, rows[0].length, DASHBOARD_THEME.SECTION);

  for (var row = 1; row < rows.length; row++) {
    var status = getTrendIndicator_(toNumber_(rows[row][3]), true).status;
    applyComparisonCellStyle_(sheet.getRange(startRow + row, startColumn + 1), status);
  }
}

function buildHourlySpeedComparisonRows_(hourlyRows, comparisonHourlyRows) {
  var latestByHour = mapHourlyRowsByHour_(hourlyRows);
  var comparisonByHour = mapHourlyRowsByHour_(comparisonHourlyRows);
  var rows = [["Hour", "Today AT", "Last Week AT", "Today PT", "Last Week PT", "Today PKT", "Last Week PKT", "Today P95", "Last Week P95"]];

  for (var hour = 0; hour < 24; hour++) {
    var latest = latestByHour[hour] || {};
    var comparison = comparisonByHour[hour] || {};

    rows.push([
      hour,
      round2_(toNumber_(latest.at)),
      round2_(toNumber_(comparison.at)),
      round2_(toNumber_(latest.pt)),
      round2_(toNumber_(comparison.pt)),
      round2_(toNumber_(latest.pkt)),
      round2_(toNumber_(comparison.pkt)),
      round2_(toNumber_(latest.p95)),
      round2_(toNumber_(comparison.p95))
    ]);
  }

  return rows;
}

function mapHourlyRowsByHour_(hourlyRows) {
  var rowsByHour = {};

  hourlyRows.forEach(function(row) {
    rowsByHour[row.hour] = row;
  });

  return rowsByHour;
}

function renderHourlySpeedComparisonTable_(sheet, speedComparisonRows) {
  var startRow = GRAPH_LAYOUT.SPEED_TABLE_ROW;
  sheet.getRange(startRow, 1, speedComparisonRows.length, speedComparisonRows[0].length).setValues(speedComparisonRows);
  formatDashboardTable_(sheet, startRow, 1, speedComparisonRows.length, speedComparisonRows[0].length, DASHBOARD_THEME.SECTION);

  for (var row = 1; row < speedComparisonRows.length; row++) {
    applyComparisonCellStyle_(sheet.getRange(startRow + row, 2), getTrendIndicator_(toNumber_(speedComparisonRows[row][1]) - toNumber_(speedComparisonRows[row][2]), false).status);
    applyComparisonCellStyle_(sheet.getRange(startRow + row, 4), getTrendIndicator_(toNumber_(speedComparisonRows[row][3]) - toNumber_(speedComparisonRows[row][4]), false).status);
    applyComparisonCellStyle_(sheet.getRange(startRow + row, 6), getTrendIndicator_(toNumber_(speedComparisonRows[row][5]) - toNumber_(speedComparisonRows[row][6]), false).status);
    applyComparisonCellStyle_(sheet.getRange(startRow + row, 8), getTrendIndicator_(toNumber_(speedComparisonRows[row][7]) - toNumber_(speedComparisonRows[row][8]), false).status);
  }
}

function renderVarianceAnalysisTable_(sheet, dailyComparisonRows) {
  var values = [["Metric", "Today", "Last Week", "Difference", "Variance %"]];

  for (var i = 1; i < dailyComparisonRows.length; i++) {
    values.push([
      dailyComparisonRows[i][0],
      dailyComparisonRows[i][2],
      dailyComparisonRows[i][1],
      dailyComparisonRows[i][3],
      dailyComparisonRows[i][4] === "" ? "" : formatVarianceText_(dailyComparisonRows[i][4])
    ]);
  }

  var startRow = GRAPH_LAYOUT.VARIANCE_ROW;
  sheet.getRange(startRow, 8, values.length, values[0].length).setValues(values);
  formatDashboardTable_(sheet, startRow, 8, values.length, values[0].length, DASHBOARD_THEME.NAVY);

  for (var row = 1; row < values.length; row++) {
    var metricName = values[row][0];
    var diff = toNumber_(values[row][3]);
    var higherIsBetter = isHigherBetterMetric_(metricName);
    var trend = getTrendIndicator_(diff, higherIsBetter);
    applyComparisonCellStyle_(sheet.getRange(startRow + row, 9), trend.status);
    sheet.getRange(startRow + row, 12).setFontColor(trend.color).setFontWeight("bold");
  }
}

function renderSlaMonitor_(sheet, dailyComparisonRows) {
  var rows = [["Metric", "Today", "Target", "Status"]];
  [
    ["AT", "Avg AT", 5],
    ["PT", "Avg PT", 11],
    ["PKT", "Avg PKT", 5],
    ["P95", "Avg P95", 2]
  ].forEach(function(definition) {
    var metric = getDailyComparisonValue_(dailyComparisonRows, definition[1]);
    var met = metric.today <= definition[2];
    rows.push([
      definition[0],
      round2_(metric.today),
      "<= " + definition[2],
      met ? "GREEN | Target Met" : "RED | Target Missed"
    ]);
  });

  var startRow = GRAPH_LAYOUT.SLA_ROW;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  formatDashboardTable_(sheet, startRow, 1, rows.length, rows[0].length, DASHBOARD_THEME.ORANGE);

  for (var i = 1; i < rows.length; i++) {
    var statusRange = sheet.getRange(startRow + i, 4);
    var metTarget = rows[i][3].indexOf("GREEN") === 0;
    statusRange
      .setValue(metTarget ? "● TARGET MET" : "● TARGET MISSED")
      .setBackground(metTarget ? "#1F3B2D" : "#4B2E2E")
      .setFontColor(metTarget ? "#A7F3D0" : "#FECACA")
      .setFontWeight("bold");
    applyComparisonCellStyle_(sheet.getRange(startRow + i, 2), metTarget ? "Improved" : "Declined");
  }
}

function renderPeakHourAnalysis_(sheet, peakHourRows) {
  var startRow = GRAPH_LAYOUT.PEAK_HOUR_ROW;
  sheet.getRange(startRow, 8, peakHourRows.length, peakHourRows[0].length).setValues(peakHourRows);
  formatDashboardTable_(sheet, startRow, 8, peakHourRows.length, peakHourRows[0].length, DASHBOARD_THEME.CYAN);

  for (var i = 1; i < peakHourRows.length; i++) {
    var status = normalizeText_(peakHourRows[i][5]);
    var color = status.indexOf("Improved") >= 0 ? "#A7F3D0" : status.indexOf("Declined") >= 0 ? "#FECACA" : DASHBOARD_THEME.SECONDARY_TEXT;
    var background = status.indexOf("Improved") >= 0 ? "#1F3B2D" : status.indexOf("Declined") >= 0 ? "#4B2E2E" : DASHBOARD_THEME.SECTION;
    applyComparisonCellStyle_(sheet.getRange(startRow + i, 10), status.indexOf("Improved") >= 0 ? "Improved" : status.indexOf("Declined") >= 0 ? "Declined" : "Stable");
    sheet.getRange(startRow + i, 13)
      .setBackground(background)
      .setFontColor(color)
      .setFontWeight("bold");
  }
}

function renderTopPickerMtdRanking_(sheet, mtdRows) {
  var rows = [["Rank", "Picker", "MTD Score", "Orders", "Items", "Complaints", "Complaint %"]];

  mtdRows.slice(0, 10).forEach(function(row) {
    rows.push([
      row.rank,
      row.picker,
      row.score,
      row.orders,
      row.items,
      row.complaints,
      row.complaintPercent
    ]);
  });

  var startRow = GRAPH_LAYOUT.PICKER_ROW;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  formatDashboardTable_(sheet, startRow, 1, rows.length, rows[0].length, DASHBOARD_THEME.GREEN);
}

function formatDashboardTable_(sheet, row, column, numRows, numColumns, headerColor) {
  var range = sheet.getRange(row, column, numRows, numColumns);
  var backgrounds = [];

  for (var r = 0; r < numRows; r++) {
    var rowBackgrounds = [];
    for (var c = 0; c < numColumns; c++) {
      rowBackgrounds.push(r === 0 ? headerColor : (r % 2 === 0 ? "#1E293B" : DASHBOARD_THEME.CARD));
    }
    backgrounds.push(rowBackgrounds);
  }

  range
    .setBackgrounds(backgrounds)
    .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, DASHBOARD_THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(row, column, 1, numColumns)
    .setBackground(headerColor)
    .setFontColor(DASHBOARD_THEME.WHITE)
    .setFontWeight("bold")
    .setFontSize(11);
  sheet.setRowHeights(row, numRows, 32);
}

function buildDailyPerformanceSummaryRows_(dailyComparisonRows) {
  return [
    { metric: "Orders", sourceMetric: "Orders", higherIsBetter: true, valueType: "integer" },
    { metric: "Items", sourceMetric: "Items", higherIsBetter: true, valueType: "integer" },
    { metric: "AT", sourceMetric: "Avg AT", higherIsBetter: false, valueType: "decimal" },
    { metric: "PT", sourceMetric: "Avg PT", higherIsBetter: false, valueType: "decimal" },
    { metric: "PKT", sourceMetric: "Avg PKT", higherIsBetter: false, valueType: "decimal" },
    { metric: "P95", sourceMetric: "Avg P95", higherIsBetter: false, valueType: "decimal" },
    { metric: "Complaints", sourceMetric: "Complaints", higherIsBetter: false, valueType: "integer" }
  ].map(function(definition) {
    return buildKpiRow_(definition, dailyComparisonRows);
  });
}

function buildProductivityKpiRows_(dailyComparisonRows, latestHourlyRows, comparisonHourlyRows) {
  var orders = getDailyComparisonValue_(dailyComparisonRows, "Orders");
  var items = getDailyComparisonValue_(dailyComparisonRows, "Items");
  var complaints = getDailyComparisonValue_(dailyComparisonRows, "Complaints");
  var latestHours = getActiveHourCount_(latestHourlyRows);
  var comparisonHours = getActiveHourCount_(comparisonHourlyRows);
  var definitions = [
    {
      metric: "Orders per Hour",
      today: latestHours > 0 ? orders.today / latestHours : 0,
      lastWeek: comparisonHours > 0 ? orders.lastWeek / comparisonHours : 0,
      higherIsBetter: true,
      valueType: "decimal"
    },
    {
      metric: "Items per Hour",
      today: latestHours > 0 ? items.today / latestHours : 0,
      lastWeek: comparisonHours > 0 ? items.lastWeek / comparisonHours : 0,
      higherIsBetter: true,
      valueType: "decimal"
    },
    {
      metric: "Average Items per Order",
      today: orders.today > 0 ? items.today / orders.today : 0,
      lastWeek: orders.lastWeek > 0 ? items.lastWeek / orders.lastWeek : 0,
      higherIsBetter: true,
      valueType: "decimal"
    },
    {
      metric: "Complaint %",
      today: orders.today > 0 ? (complaints.today / orders.today) * 100 : 0,
      lastWeek: orders.lastWeek > 0 ? (complaints.lastWeek / orders.lastWeek) * 100 : 0,
      higherIsBetter: false,
      valueType: "percent"
    }
  ];

  return definitions.map(function(definition) {
    return buildKpiFromValues_(definition);
  });
}

function buildPeakHourAnalysisRows_(latestHourlyRows, comparisonHourlyRows) {
  var definitions = [
    ["Peak Orders Hour", "orders", true, true, "integer"],
    ["Peak Items Hour", "items", true, true, "integer"],
    ["Worst AT Hour", "at", true, false, "decimal"],
    ["Worst PT Hour", "pt", true, false, "decimal"]
  ];
  var rows = [["Metric", "Today Hour", "Today Value", "Last Week Hour", "Last Week Value", "Status"]];

  definitions.forEach(function(definition) {
    var todayPeak = findHourlyExtreme_(latestHourlyRows, definition[1], definition[2]);
    var lastWeekPeak = findHourlyExtreme_(comparisonHourlyRows, definition[1], definition[2]);
    var trend = getTrendIndicator_(todayPeak.value - lastWeekPeak.value, definition[3]);
    rows.push([
      definition[0],
      todayPeak.hour === "" ? "-" : todayPeak.hour + ":00",
      formatKpiValue_(todayPeak.value, definition[4]),
      lastWeekPeak.hour === "" ? "-" : lastWeekPeak.hour + ":00",
      formatKpiValue_(lastWeekPeak.value, definition[4]),
      trend.indicator + " " + trend.status
    ]);
  });

  return rows;
}

function buildHourlyHeatmapRows_(orderMatrix, itemMatrix, performanceRows) {
  var rows = [["Hour", "LW Orders", "Today Orders", "Order Var", "LW Items", "Today Items", "Item Var", "AT", "PT", "PKT", "P95"]];
  var performanceByHour = {};

  for (var p = 1; p < performanceRows.length; p++) {
    performanceByHour[performanceRows[p][0]] = performanceRows[p];
  }

  for (var i = 1; i < orderMatrix.length; i++) {
    var hour = orderMatrix[i][0];
    var perf = performanceByHour[hour] || [hour, 0, 0, 0, 0];
    rows.push([
      hour,
      orderMatrix[i][1],
      orderMatrix[i][2],
      orderMatrix[i][3],
      itemMatrix[i] ? itemMatrix[i][1] : 0,
      itemMatrix[i] ? itemMatrix[i][2] : 0,
      itemMatrix[i] ? itemMatrix[i][3] : 0,
      perf[1],
      perf[2],
      perf[3],
      perf[4]
    ]);
  }

  return rows;
}

function buildKpiRow_(definition, dailyComparisonRows) {
  var values = getDailyComparisonValue_(dailyComparisonRows, definition.sourceMetric);
  return buildKpiFromValues_({
    metric: definition.metric,
    today: values.today,
    lastWeek: values.lastWeek,
    variancePercent: values.variancePercent,
    higherIsBetter: definition.higherIsBetter,
    valueType: definition.valueType
  });
}

function buildKpiFromValues_(definition) {
  var difference = definition.today - definition.lastWeek;
  var variancePercent = definition.variancePercent !== undefined
    ? definition.variancePercent
    : definition.lastWeek !== 0
      ? (difference / definition.lastWeek) * 100
      : "";
  var trend = getTrendIndicator_(difference, definition.higherIsBetter);

  return {
    metric: definition.metric,
    today: definition.today,
    lastWeek: definition.lastWeek,
    difference: difference,
    variancePercent: variancePercent,
    todayText: formatKpiValue_(definition.today, definition.valueType),
    lastWeekText: formatKpiValue_(definition.lastWeek, definition.valueType),
    varianceText: formatVarianceText_(variancePercent),
    indicator: trend.indicator,
    status: trend.status,
    color: trend.color,
    background: trend.background
  };
}

function getDailyComparisonValue_(dailyComparisonRows, metricName) {
  var row = getDailyComparisonRow_(dailyComparisonRows, metricName);
  if (!row) {
    return {
      today: 0,
      lastWeek: 0,
      difference: 0,
      variancePercent: ""
    };
  }

  return {
    today: toNumber_(row[2]),
    lastWeek: toNumber_(row[1]),
    difference: toNumber_(row[3]),
    variancePercent: row[4] === "" ? "" : toNumber_(row[4])
  };
}

function getDailyComparisonRow_(dailyComparisonRows, metricName) {
  var target = normalizeForCompare_(metricName);

  for (var i = 1; i < dailyComparisonRows.length; i++) {
    if (normalizeForCompare_(dailyComparisonRows[i][0]) === target) {
      return dailyComparisonRows[i];
    }
  }

  return null;
}

function getTrendIndicator_(difference, higherIsBetter) {
  if (difference === 0) {
    return {
      indicator: "→",
      status: "Stable",
      color: DASHBOARD_THEME.SECONDARY_TEXT,
      background: "#334155",
      border: "#475569"
    };
  }

  var improved = higherIsBetter ? difference > 0 : difference < 0;

  return {
    indicator: difference > 0 ? "▲" : "▼",
    status: improved ? "Improved" : "Declined",
    color: improved ? "#A7F3D0" : "#FECACA",
    background: improved ? "#1F3B2D" : "#4B2E2E",
    border: improved ? "#355E4A" : "#6B3A3A"
  };
}

function isHigherBetterMetric_(metricName) {
  var normalized = normalizeForCompare_(metricName);
  return normalized === "orders" ||
    normalized === "items" ||
    normalized === "orders per hour" ||
    normalized === "items per hour" ||
    normalized === "average items per order";
}

function buildExecutiveBannerText_(dailySummaryRows) {
  return dailySummaryRows.map(function(row) {
    return row.metric + " " + row.indicator + " " + row.varianceText;
  }).join("   |   ");
}

function getOverallPerformanceStatus_(dailySummaryRows) {
  var improved = 0;
  var declined = 0;

  dailySummaryRows.forEach(function(row) {
    if (row.status === "Improved") improved++;
    if (row.status === "Declined") declined++;
  });

  if (improved > declined) {
    return {
      label: "Improved",
      color: "#14532D"
    };
  }

  if (declined > improved) {
    return {
      label: "Declined",
      color: "#7F1D1D"
    };
  }

  return {
    label: "Stable",
    color: DASHBOARD_THEME.SECTION
  };
}

function getActiveHourCount_(hourlyRows) {
  var count = 0;

  hourlyRows.forEach(function(row) {
    if (toNumber_(row.orders) > 0 || toNumber_(row.items) > 0) {
      count++;
    }
  });

  return count || hourlyRows.length || 0;
}

function findHourlyExtreme_(hourlyRows, metricName, highest) {
  var result = {
    hour: "",
    value: 0
  };

  hourlyRows.forEach(function(row) {
    var value = toNumber_(row[metricName]);
    if (value <= 0) return;

    if (result.hour === "" || (highest ? value > result.value : value < result.value)) {
      result = {
        hour: row.hour,
        value: value
      };
    }
  });

  return result;
}

function buildTopPickerSourceRows_(mtdRows) {
  return [["Picker", "MTD Score"]].concat(
    mtdRows.slice(0, 10).map(function(row) {
      return [row.picker, row.score];
    })
  );
}

function kpiRowsToTable_(kpiRows) {
  var output = [["Metric", "Today", "Last Week", "Difference", "Variance %", "Status"]];

  kpiRows.forEach(function(row) {
    output.push([
      row.metric,
      row.today,
      row.lastWeek,
      row.difference,
      row.variancePercent,
      row.indicator + " " + row.status
    ]);
  });

  return output;
}

function getKpiCardColor_(metricName) {
  var normalized = normalizeForCompare_(metricName);

  if (normalized.indexOf("order") >= 0) return DASHBOARD_THEME.BLUE;
  if (normalized.indexOf("item") >= 0) return DASHBOARD_THEME.GREEN;
  if (normalized === "at") return DASHBOARD_THEME.PURPLE;
  if (normalized === "pt") return DASHBOARD_THEME.ORANGE;
  if (normalized === "pkt") return DASHBOARD_THEME.CYAN;
  if (normalized === "p95" || normalized.indexOf("complaint") >= 0) return DASHBOARD_THEME.RED;

  return DASHBOARD_THEME.BLUE;
}

function applyComparisonCellStyle_(range, status) {
  var style = getComparisonStyle_(status);
  range
    .setBackground(style.background)
    .setFontColor(style.text)
    .setFontWeight("bold")
    .setBorder(true, true, true, true, false, false, style.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function getComparisonStyle_(status) {
  if (status === "Improved") {
    return {
      background: "#1F3B2D",
      text: "#A7F3D0",
      border: "#355E4A"
    };
  }

  if (status === "Declined") {
    return {
      background: "#4B2E2E",
      text: "#FECACA",
      border: "#6B3A3A"
    };
  }

  return {
    background: "#334155",
    text: "#CBD5E1",
    border: "#475569"
  };
}

function formatKpiValue_(value, valueType) {
  if (valueType === "integer") return formatNumber_(Math.round(toNumber_(value)), 0);
  if (valueType === "percent") return formatNumber_(round2_(toNumber_(value)), 2) + "%";
  return formatNumber_(round2_(toNumber_(value)), 2);
}

function formatVarianceText_(variancePercent) {
  if (variancePercent === "" || variancePercent === null || variancePercent === undefined) return "n/a";

  var value = round2_(toNumber_(variancePercent));
  return (value > 0 ? "+" : "") + formatNumber_(value, 2) + "%";
}

function formatNumber_(value, decimals) {
  var fixed = decimals > 0 ? Number(value).toFixed(decimals) : String(Math.round(Number(value)));
  var parts = fixed.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return parts.join(".");
}

function buildComplaintIndex_(rcaSheet) {
  var data = rcaSheet.getDataRange().getValues();
  var complaints = {};

  for (var r = 1; r < data.length; r++) {
    var complaintDate = normalizeDate_(data[r][0]);
    var complaintOrder = normalizeText_(data[r][4]);
    var complaintPicker = normalizeText_(data[r][13]);

    if (!complaintDate || !complaintOrder || !complaintPicker) continue;

    var key = makeKey_(complaintDate, complaintPicker);
    if (!complaints[key]) complaints[key] = {};

    complaints[key][complaintOrder] = true;
  }

  return complaints;
}

function buildRawSummary_(sourceSS) {
  var summary = {};
  var sheets = sourceSS.getSheets();

  sheets.forEach(function(sheet) {
    var sheetDateKey = getDateKeyFromSheetName_(sheet.getName());
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var dateKey = sheetDateKey || getValidDateKey_(data[i][0]);
      var orderNr = normalizeText_(data[i][2]);
      var picker = normalizeText_(data[i][9]);

      if (!dateKey || !orderNr || !picker) continue;

      var key = makeKey_(dateKey, picker);
      if (!summary[key]) {
        summary[key] = {
          dateKey: dateKey,
          picker: picker,
          orders: {},
          itemCount: 0,
          totalAcceptance: 0,
          totalPicking: 0,
          totalPacking: 0,
          totalFulfillment: 0
        };
      }

      var fulfillmentTime = toNumber_(data[i][7]);

      summary[key].orders[orderNr] = true;
      summary[key].itemCount += toNumber_(data[i][8]);
      summary[key].totalAcceptance += toNumber_(data[i][4]);
      summary[key].totalPicking += toNumber_(data[i][5]);
      summary[key].totalPacking += toNumber_(data[i][6]);
      summary[key].totalFulfillment += fulfillmentTime;
    }
  });

  return summary;
}

function buildHourlySummary_(sourceSS) {
  var summary = {};
  var sheets = sourceSS.getSheets();

  sheets.forEach(function(sheet) {
    var sheetDateKey = getDateKeyFromSheetName_(sheet.getName());
    var data = sheet.getDataRange().getValues();

    if (sheetDateKey) {
      ensureHourlyDate_(summary, sheetDateKey);
    }

    for (var i = 1; i < data.length; i++) {
      var dateKey = sheetDateKey || getValidDateKey_(data[i][0]);
      var orderNr = normalizeText_(data[i][2]);

      if (!dateKey || !orderNr) continue;

      var hour = extractHourFromRow_(data[i]);
      var bucket = ensureHourlyDate_(summary, dateKey)[hour];

      if (!bucket.orders[orderNr]) {
        bucket.orders[orderNr] = true;
        bucket.totalAcceptance += toNumber_(data[i][4]);
        bucket.totalPicking += toNumber_(data[i][5]);
        bucket.totalPacking += toNumber_(data[i][6]);
        bucket.totalFulfillment += toNumber_(data[i][7]);
      }

      bucket.itemCount += toNumber_(data[i][8]);
    }
  });

  return summary;
}

function readExistingHourlyManualMp_(sheet) {
  var manualMp = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return manualMp;

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 10)).getValues();

  data.forEach(function(row) {
    var dateKey = getValidDateKey_(row[0]);
    var hour = extractHourFromValue_(row[1]);
    var mp = row[9];

    if (!dateKey || hour === null || normalizeText_(row[1]) === "Total") return;

    manualMp[makeHourlyKey_(dateKey, hour)] = mp;
  });

  return manualMp;
}

function formatRdbHourlySheet_(sheet, values, totalRows, hourlyBlocks) {
  if (!values || values.length === 0) return;

  sheet.setConditionalFormatRules([]);

  var totalRowLookup = {};
  totalRows.forEach(function(rowNumber) {
    totalRowLookup[rowNumber] = true;
  });

  var backgrounds = [];
  var fontWeights = [];

  for (var r = 1; r < values.length; r++) {
    var rowNumber = r + 1;
    var rowBackgrounds = [];
    var rowFontWeights = [];

    for (var c = 0; c < 10; c++) {
      rowBackgrounds.push("#ffffff");
      rowFontWeights.push("normal");
    }

    if (totalRowLookup[rowNumber]) {
      for (var totalCol = 0; totalCol < 10; totalCol++) {
        rowBackgrounds[totalCol] = "#fff2cc";
        rowFontWeights[totalCol] = "bold";
      }
    }

    backgrounds.push(rowBackgrounds);
    fontWeights.push(rowFontWeights);
  }

  hourlyBlocks.forEach(function(block) {
    applyGreenScaleToHourlyBlock_(values, backgrounds, block, 2);
    applyGreenScaleToHourlyBlock_(values, backgrounds, block, 3);

    for (var offset = 0; offset < block.numRows; offset++) {
      var valueRowIndex = block.startRow - 1 + offset;
      var backgroundIndex = valueRowIndex - 1;
      var orderCount = toNumber_(values[valueRowIndex][2]);
      var itemCount = toNumber_(values[valueRowIndex][3]);

      if (orderCount > 0) {
        backgrounds[backgroundIndex][4] = thresholdColor_(toNumber_(values[valueRowIndex][4]), 6, true);
        backgrounds[backgroundIndex][5] = thresholdColor_(toNumber_(values[valueRowIndex][5]), 5, true);
        backgrounds[backgroundIndex][8] = thresholdColor_(toNumber_(values[valueRowIndex][8]), 2, true);
      }

      if (itemCount > 0) {
        backgrounds[backgroundIndex][6] = thresholdColor_(toNumber_(values[valueRowIndex][6]), 11, true);
        backgrounds[backgroundIndex][7] = thresholdColor_(toNumber_(values[valueRowIndex][7]), 5, true);
      }
    }
  });

  sheet.getRange(1, 1, 1, 10)
    .setBackground("#002060")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, 10)
      .setBackgrounds(backgrounds)
      .setFontWeights(fontWeights)
      .setFontColor("#000000");
  }
}

function applyGreenScaleToHourlyBlock_(values, backgrounds, block, columnIndex) {
  var valuesForScale = [];

  for (var offset = 0; offset < block.numRows; offset++) {
    var valueRowIndex = block.startRow - 1 + offset;
    valuesForScale.push(toNumber_(values[valueRowIndex][columnIndex]));
  }

  var minValue = Math.min.apply(null, valuesForScale);
  var maxValue = Math.max.apply(null, valuesForScale);

  for (var i = 0; i < valuesForScale.length; i++) {
    var rowIndex = block.startRow - 1 + i;
    var backgroundIndex = rowIndex - 1;

    backgrounds[backgroundIndex][columnIndex] =
      greenScaleColor_(valuesForScale[i], minValue, maxValue);
  }
}

function greenScaleColor_(value, minValue, maxValue) {
  if (maxValue <= minValue) {
    return value > 0 ? "#b7e1cd" : "#ffffff";
  }

  var ratio = (value - minValue) / (maxValue - minValue);
  var start = { r: 232, g: 245, b: 233 };
  var end = { r: 102, g: 187, b: 106 };

  return rgbToHex_({
    r: Math.round(start.r + ((end.r - start.r) * ratio)),
    g: Math.round(start.g + ((end.g - start.g) * ratio)),
    b: Math.round(start.b + ((end.b - start.b) * ratio))
  });
}

function thresholdColor_(value, threshold, higherIsRed) {
  if (value === threshold) return "#fff2cc";

  var isRed = higherIsRed ? value > threshold : value < threshold;

  return isRed ? "#f4cccc" : "#d9ead3";
}

function rgbToHex_(color) {
  return "#" +
    hexPart_(color.r) +
    hexPart_(color.g) +
    hexPart_(color.b);
}

function hexPart_(value) {
  var hex = Math.max(0, Math.min(255, value)).toString(16);
  return hex.length === 1 ? "0" + hex : hex;
}

function ensureHourlyDate_(summary, dateKey) {
  if (!summary[dateKey]) {
    summary[dateKey] = [];

    for (var hour = 0; hour < 24; hour++) {
      summary[dateKey][hour] = createHourlyBucket_();
    }
  }

  return summary[dateKey];
}

function createHourlyBucket_() {
  return {
    orders: {},
    itemCount: 0,
    totalAcceptance: 0,
    totalPicking: 0,
    totalPacking: 0,
    totalFulfillment: 0
  };
}

function addBucketToTotal_(totalBucket, bucket) {
  Object.keys(bucket.orders).forEach(function(orderNr) {
    totalBucket.orders[orderNr] = true;
  });

  totalBucket.itemCount += bucket.itemCount;
  totalBucket.totalAcceptance += bucket.totalAcceptance;
  totalBucket.totalPicking += bucket.totalPicking;
  totalBucket.totalPacking += bucket.totalPacking;
  totalBucket.totalFulfillment += bucket.totalFulfillment;
}

function extractHourFromRow_(row) {
  var candidates = [row[1], row[0], row[3]];

  for (var i = 0; i < candidates.length; i++) {
    var hour = extractHourFromValue_(candidates[i]);

    if (hour !== null) return hour;
  }

  return 0;
}

function extractHourFromValue_(value) {
  if (value === null || value === undefined || value === "") return null;

  if (value instanceof Date) {
    return value.getHours();
  }

  if (typeof value === "number") {
    if (value >= 0 && value < 1) return Math.floor(value * 24);
    if (value >= 0 && value <= 23) return Math.floor(value);
    return null;
  }

  var text = String(value).trim();
  if (!text) return null;

  var numericHour = Number(text);
  if (!isNaN(numericHour) && numericHour >= 0 && numericHour <= 23) {
    return Math.floor(numericHour);
  }

  if (/:|AM|PM/i.test(text)) {
    var timeMatch = text.match(/\b(\d{1,2})(?::\d{2})?\s*(AM|PM)?\b/i);

    if (timeMatch) {
      var hour = Number(timeMatch[1]);
      var meridiem = timeMatch[2] ? timeMatch[2].toUpperCase() : "";

      if (hour >= 0 && hour <= 23) {
        if (meridiem === "PM" && hour < 12) hour += 12;
        if (meridiem === "AM" && hour === 12) hour = 0;

        return hour;
      }
    }
  }

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed.getHours();
  }

  return null;
}

function getDateKeyFromSheetName_(sheetName) {
  return getValidDateKey_(sheetName);
}

function getValidDateKey_(value) {
  var dateKey = normalizeDate_(value);

  return parseDateKey_(dateKey) ? dateKey : "";
}

function makeHourlyKey_(dateKey, hour) {
  return dateKey + "|" + hour;
}

function readSourceOrderRecords_(sourceSS) {
  var records = [];
  var sheets = sourceSS.getSheets();

  sheets.forEach(function(sheet) {
    var sheetDateKey = getDateKeyFromSheetName_(sheet.getName());
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var dateKey = sheetDateKey || getValidDateKey_(data[i][0]);
      var orderNr = normalizeText_(data[i][2]);
      var picker = normalizeText_(data[i][9]);

      if (!dateKey || !orderNr || !picker) continue;

      records.push({
        dateKey: dateKey,
        orderNr: orderNr,
        picker: picker,
        acceptanceTime: toNumber_(data[i][4]),
        pickingTime: toNumber_(data[i][5]),
        packingTime: toNumber_(data[i][6]),
        fulfillmentTime: toNumber_(data[i][7]),
        itemCount: toNumber_(data[i][8])
      });
    }
  });

  return records;
}

function buildWeeklyPickerPerformance_(sourceRecords, weekRange) {
  var summary = {};

  sourceRecords.forEach(function(record) {
    if (!isInDateRange_(record.dateKey, weekRange.startKey, weekRange.endKey)) return;

    if (!summary[record.picker]) {
      summary[record.picker] = {
        picker: record.picker,
        orders: {},
        itemCount: 0,
        totalAcceptance: 0,
        totalPicking: 0,
        totalPacking: 0,
        fulfillmentTimes: []
      };
    }

    var pickerSummary = summary[record.picker];

    if (!pickerSummary.orders[record.orderNr]) {
      pickerSummary.orders[record.orderNr] = true;
      pickerSummary.totalAcceptance += record.acceptanceTime;
      pickerSummary.totalPicking += record.pickingTime;
      pickerSummary.totalPacking += record.packingTime;

      if (record.fulfillmentTime > 0) {
        pickerSummary.fulfillmentTimes.push(record.fulfillmentTime);
      }
    }

    pickerSummary.itemCount += record.itemCount;
  });

  return Object.keys(summary).map(function(picker) {
    var row = summary[picker];
    var orderCount = Object.keys(row.orders).length;
    var p95 = calculatePercentile_(row.fulfillmentTimes, 0.95) / 60;

    return {
      picker: picker,
      uniqueOrderCount: orderCount,
      itemCount: row.itemCount,
      avgAcceptance: orderCount > 0 ? row.totalAcceptance / orderCount : 0,
      pickingPerItem: row.itemCount > 0 ? row.totalPicking / row.itemCount : 0,
      p95: p95,
      efficiency: row.totalPicking > 0 ? row.itemCount / row.totalPicking : 0
    };
  });
}

function buildWeeklyComplaintAnalytics_(rcaData, weekRange) {
  var result = {
    skus: {},
    pickers: {},
    rca: {},
    customers: {},
    deliveryAgents: {}
  };
  var allowedDaBuckets = buildAllowedValues_(["LM Bucket", "DA Mistake", "Transit Damage"]);

  for (var i = 1; i < rcaData.length; i++) {
    var dateKey = normalizeDate_(rcaData[i][0]);
    var orderNr = normalizeText_(rcaData[i][4]);

    if (!dateKey || !orderNr || !isInDateRange_(dateKey, weekRange.startKey, weekRange.endKey)) continue;

    addUniqueComplaint_(result.skus, normalizeText_(rcaData[i][5]), orderNr, normalizeText_(rcaData[i][8]));
    addUniqueComplaint_(result.pickers, normalizeText_(rcaData[i][13]), orderNr);
    addUniqueComplaint_(result.rca, normalizeText_(rcaData[i][16]), orderNr);
    addUniqueComplaint_(result.customers, normalizeText_(rcaData[i][15]), orderNr);

    if (allowedDaBuckets[normalizeForCompare_(rcaData[i][16])]) {
      addUniqueComplaint_(result.deliveryAgents, normalizeText_(rcaData[i][10]), orderNr, normalizeText_(rcaData[i][11]));
    }
  }

  return {
    skus: mapUniqueComplaintRows_(result.skus),
    pickers: mapUniqueComplaintRows_(result.pickers),
    rca: mapUniqueComplaintRows_(result.rca, true),
    customers: mapUniqueComplaintRows_(result.customers),
    deliveryAgents: mapUniqueComplaintRows_(result.deliveryAgents)
  };
}

function buildWeeklyBagDeviation_(rdbData, weekRange) {
  var summary = {};

  for (var i = 1; i < rdbData.length; i++) {
    var dateKey = normalizeDate_(rdbData[i][0]);
    var picker = normalizeText_(rdbData[i][1]);

    if (!dateKey || !picker || !isInDateRange_(dateKey, weekRange.startKey, weekRange.endKey)) continue;

    if (!summary[picker]) {
      summary[picker] = {
        picker: picker,
        orderCount: 0,
        bagDeviationTotal: 0,
        days: 0
      };
    }

    summary[picker].orderCount += toNumber_(rdbData[i][3]);
    summary[picker].bagDeviationTotal += toNumber_(rdbData[i][10]);
    summary[picker].days++;
  }

  return Object.keys(summary).map(function(picker) {
    var row = summary[picker];
    var avgBagDeviation = row.days > 0 ? row.bagDeviationTotal / row.days : 0;

    return {
      picker: picker,
      uniqueOrderCount: row.orderCount,
      avgBagDeviation: avgBagDeviation,
      absBagDeviation: Math.abs(avgBagDeviation)
    };
  }).sort(function(a, b) {
    return b.absBagDeviation - a.absBagDeviation || b.uniqueOrderCount - a.uniqueOrderCount ||
      a.picker.localeCompare(b.picker);
  });
}

function addUniqueComplaint_(summary, key, orderNr, label) {
  if (!key || !orderNr) return;

  if (!summary[key]) {
    summary[key] = {
      key: key,
      label: label || "",
      orders: {}
    };
  }

  if (!summary[key].label && label) {
    summary[key].label = label;
  }

  summary[key].orders[orderNr] = true;
}

function mapUniqueComplaintRows_(summary, includeOrderIds) {
  return Object.keys(summary).map(function(key) {
    var orderIds = Object.keys(summary[key].orders).sort();

    return {
      key: key,
      label: summary[key].label,
      uniqueOrderCount: orderIds.length,
      orderIds: includeOrderIds ? orderIds.join(", ") : ""
    };
  }).sort(function(a, b) {
    return b.uniqueOrderCount - a.uniqueOrderCount || a.key.localeCompare(b.key);
  });
}

function readExistingRdbManualInputs_(sheet) {
  var manualInputs = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return manualInputs;

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getValues();

  data.forEach(function(row) {
    var dateKey = normalizeDate_(row[0]);
    var picker = normalizeText_(row[1]);

    if (!dateKey || !picker) return;

    manualInputs[makeKey_(dateKey, picker)] = {
      bagDeviation: toNumber_(row[10]),
      putaway: toNumber_(row[11])
    };
  });

  return manualInputs;
}

function readExistingRdbManualRows_(sheet) {
  var manualRows = [];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return manualRows;

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getValues();

  data.forEach(function(row) {
    var dateKey = normalizeDate_(row[0]);
    var picker = normalizeText_(row[1]);

    if (!dateKey || !picker) return;

    var preservedRow = row.slice(0, 12);
    while (preservedRow.length < 12) {
      preservedRow.push("");
    }

    preservedRow[0] = dateKey;
    preservedRow[1] = picker;

    manualRows.push({
      key: makeKey_(dateKey, picker),
      row: preservedRow
    });
  });

  return manualRows;
}

function readExistingMerchandisingScores_(sheet) {
  var scores = canonicalizeMtdPickerScores_(getStoredMerchandisingScores_());
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return scores;

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 15)).getValues();

  data.forEach(function(row) {
    var picker = getMtdPickerName_(row[0]);
    var score = row[14];

    if (picker && score !== "" && score !== null && score !== undefined) {
      scores[picker] = Math.max(scores[picker] || 0, toNumber_(score));
    }
  });

  return scores;
}

function canonicalizeMtdPickerScores_(scores) {
  var canonicalScores = {};

  Object.keys(scores || {}).forEach(function(picker) {
    var canonicalPicker = getMtdPickerName_(picker);

    if (!canonicalPicker) return;

    canonicalScores[canonicalPicker] = Math.max(
      canonicalScores[canonicalPicker] || 0,
      toNumber_(scores[picker])
    );
  });

  return canonicalScores;
}

function getStoredMerchandisingScores_() {
  var text = PropertiesService.getScriptProperties()
    .getProperty(CONFIG.PROPERTIES.MERCHANDISING_SCORES);

  if (!text) return {};

  try {
    return JSON.parse(text);
  } catch (error) {
    return {};
  }
}

function saveMerchandisingScores_(scores) {
  PropertiesService.getScriptProperties()
    .setProperty(CONFIG.PROPERTIES.MERCHANDISING_SCORES, JSON.stringify(scores || {}));
}

function buildComplaintIntelligenceRows_(data, orderContext) {
  var records = getComplaintRecords_(data);
  var stats = buildComplaintStats_(records, orderContext);
  var output = [
    ["COMPLAINT INTELLIGENCE CENTER"],
    ["Scope", "All recorded complaints", "Source", CONFIG.SHEETS.RCA_DATABASE, "Generated At", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")],
    [""]
  ];

  addComplaintSection_(output, "EXECUTIVE COMPLAINT SUMMARY", ["Metric", "Value", "Definition"], [
    ["Unique Orders With Complaints", stats.uniqueOrdersWithComplaints, "Distinct Order IDs with at least one complaint"],
    ["Complaint Rate %", formatComplaintPercent_(stats.complaintRate), "Unique Orders With Complaints / Total Orders from RDB"],
    ["Most Common RCA", stats.topRca.key || "No data", "Highest complaint count RCA"],
    ["Most Common RCA %", formatComplaintPercent_(stats.topRca.percent), "Share of unique complaint orders"],
    ["Most Common Complaint Category", stats.topCategory.key || "No data", "Highest complaint count category"],
    ["Most Common Complaint Category %", formatComplaintPercent_(stats.topCategory.percent), "Share of unique complaint orders"],
    ["Top Complaint SKU", stats.topSku.key || "No data", stats.topSku.label || ""],
    ["Top Complaint Picker", stats.topPicker.key || "No data", "Highest complaint count picker"]
  ]);

  addComplaintSection_(output, "MONTHLY COMPARISON", ["Month", "Complaint Count", "Unique Orders Impacted", "Complaint Rate %"],
    buildComplaintMonthlyRows_(records, orderContext));

  addComplaintSection_(output, "RCA ANALYSIS", ["RCA", "Complaint Count", "RCA Complaint %"],
    buildComplaintSummaryRows_(stats.rca, false, 0, false, stats.uniqueOrdersWithComplaints));

  addComplaintSection_(output, "COMPLAINT CATEGORY ANALYSIS", ["Complaint Category", "Complaint Count", "Category Complaint %"],
    buildComplaintSummaryRows_(stats.categories, false, 0, false, stats.uniqueOrdersWithComplaints));

  addComplaintSection_(output, "TOP 5 COMPLAINT SKUs", ["Rank", "SKU", "Complaint Count", "SKU Complaint %"],
    buildComplaintSummaryRows_(stats.skus, true, 5, false, stats.uniqueOrdersWithComplaints));

  addComplaintSection_(output, "TOP 5 PICKERS BY COMPLAINT COUNT", ["Rank", "Picker", "Complaint Count", "Unique Orders Impacted", "Picker Complaint %"],
    buildComplaintSummaryRows_(stats.pickers, true, 5, true, stats.uniqueOrdersWithComplaints, orderContext.ordersByPicker));

  addComplaintSection_(output, "TOP 5 CUSTOMERS BY COMPLAINT COUNT", ["Rank", "Customer", "Complaint Count", "Unique Orders Impacted", "Customer Complaint %"],
    buildComplaintSummaryRows_(stats.customers, true, 5, true, stats.uniqueOrdersWithComplaints));

  addComplaintSection_(output, "TOP 5 DAs BY COMPLAINT COUNT", ["Rank", "DA", "Complaint Count", "Unique Orders Impacted", "DA Complaint %"],
    buildComplaintSummaryRows_(stats.deliveryAgents, true, 5, true, stats.uniqueOrdersWithComplaints));

  addComplaintSection_(output, "KEY INSIGHTS", ["Insight"], buildComplaintInsightRows_(stats));

  addComplaintSection_(output, "RAW COMPLAINT DATA", data[0] || ["No header"], data.slice(1));

  return output;
}

function getComplaintOrderContextFromRdb_(rdbData) {
  var context = {
    totalOrders: 0,
    ordersByMonth: {},
    ordersByPicker: {}
  };

  for (var i = 1; i < rdbData.length; i++) {
    var dateKey = normalizeDate_(rdbData[i][0]);
    var monthKey = getMonthKey_(dateKey);
    var picker = normalizeText_(rdbData[i][1]);
    var orders = toNumber_(rdbData[i][3]);

    context.totalOrders += orders;
    if (monthKey) context.ordersByMonth[monthKey] = (context.ordersByMonth[monthKey] || 0) + orders;
    if (picker) context.ordersByPicker[picker] = (context.ordersByPicker[picker] || 0) + orders;
  }

  return context;
}

function getComplaintRecords_(data) {
  var records = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var orderId = normalizeText_(row[4]);

    if (!orderId && !normalizeDate_(row[0]) && !normalizeText_(row[16]) && !normalizeText_(row[7])) continue;

    records.push({
      raw: row,
      dateKey: normalizeDate_(row[0]),
      monthKey: getMonthKey_(normalizeDate_(row[0])),
      orderId: orderId,
      sku: normalizeText_(row[5]),
      skuName: normalizeText_(row[8]),
      category: normalizeText_(row[7]),
      reason: normalizeText_(row[9]),
      daId: normalizeText_(row[10]),
      daName: normalizeText_(row[11]),
      picker: normalizeText_(row[13]),
      customer: normalizeText_(row[15]),
      rca: normalizeText_(row[16])
    });
  }

  return records;
}

function buildComplaintStats_(records, orderContext) {
  var totalComplaints = records.length;
  var uniqueOrders = {};
  var stats = {
    totalComplaints: totalComplaints,
    rca: {},
    categories: {},
    skus: {},
    pickers: {},
    customers: {},
    deliveryAgents: {}
  };

  records.forEach(function(record) {
    if (record.orderId) uniqueOrders[record.orderId] = true;

    addComplaintSummaryItem_(stats.rca, record.rca || "Unclassified", "", record.orderId);
    addComplaintSummaryItem_(stats.categories, record.category || "Unclassified", "", record.orderId);
    addComplaintSummaryItem_(stats.skus, record.sku || "Unknown SKU", record.skuName, record.orderId);
    addComplaintSummaryItem_(stats.pickers, record.picker || "Unknown Picker", "", record.orderId);
    addComplaintSummaryItem_(stats.customers, record.customer || "Unknown Customer", "", record.orderId);
    addComplaintSummaryItem_(stats.deliveryAgents, record.daId || record.daName || "Unknown DA", record.daName, record.orderId);
  });

  stats.uniqueOrdersWithComplaints = Object.keys(uniqueOrders).length;
  stats.uniqueOrders = stats.uniqueOrdersWithComplaints;
  stats.totalOrders = toNumber_(orderContext.totalOrders);
  stats.complaintRate = stats.totalOrders > 0 ? (stats.uniqueOrdersWithComplaints / stats.totalOrders) * 100 : 0;
  stats.rcaRows = sortComplaintSummary_(stats.rca, stats.uniqueOrdersWithComplaints);
  stats.categoryRows = sortComplaintSummary_(stats.categories, stats.uniqueOrdersWithComplaints);
  stats.skuRows = sortComplaintSummary_(stats.skus, stats.uniqueOrdersWithComplaints);
  stats.pickerRows = sortComplaintSummary_(stats.pickers, stats.uniqueOrdersWithComplaints, orderContext.ordersByPicker);
  stats.customerRows = sortComplaintSummary_(stats.customers, stats.uniqueOrdersWithComplaints);
  stats.daRows = sortComplaintSummary_(stats.deliveryAgents, stats.uniqueOrdersWithComplaints);
  stats.topRca = stats.rcaRows[0] || {};
  stats.topCategory = stats.categoryRows[0] || {};
  stats.topSku = stats.skuRows[0] || {};
  stats.topPicker = stats.pickerRows[0] || {};

  return stats;
}

function addComplaintSummaryItem_(summary, key, label, orderId) {
  if (!summary[key]) {
    summary[key] = {
      key: key,
      label: label || "",
      count: 0,
      orders: {}
    };
  }

  if (!summary[key].label && label) summary[key].label = label;
  summary[key].count++;
  if (orderId) summary[key].orders[orderId] = true;
}

function sortComplaintSummary_(summary, denominator, denominatorLookup) {
  return Object.keys(summary).map(function(key) {
    var uniqueOrders = Object.keys(summary[key].orders).length;
    var percentDenominator = denominatorLookup ? toNumber_(denominatorLookup[key]) : toNumber_(denominator);

    return {
      key: key,
      label: summary[key].label,
      count: summary[key].count,
      uniqueOrders: uniqueOrders,
      percent: percentDenominator > 0 ? (uniqueOrders / percentDenominator) * 100 : 0
    };
  }).sort(function(a, b) {
    return b.count - a.count || a.key.localeCompare(b.key);
  });
}

function buildComplaintMonthlyRows_(records, orderContext) {
  var months = {};

  records.forEach(function(record) {
    var monthKey = record.monthKey || "Unknown";
    if (!months[monthKey]) {
      months[monthKey] = {
        count: 0,
        orders: {}
      };
    }

    months[monthKey].count++;
    if (record.orderId) months[monthKey].orders[record.orderId] = true;
  });

  return Object.keys(months).sort().map(function(monthKey) {
    var uniqueOrders = Object.keys(months[monthKey].orders).length;
    var totalOrders = toNumber_(orderContext.ordersByMonth[monthKey]);
    return [
      monthKey,
      months[monthKey].count,
      uniqueOrders,
      totalOrders > 0 ? formatComplaintPercent_((uniqueOrders / totalOrders) * 100) : "n/a"
    ];
  });
}

function buildComplaintSummaryRows_(summary, ranked, limit, includeUniqueOrders, denominator, denominatorLookup) {
  var rows = sortComplaintSummary_(summary, denominator, denominatorLookup);
  var limitedRows = limit > 0 ? rows.slice(0, limit) : rows;

  return limitedRows.map(function(row, index) {
    var name = row.label ? row.key + " - " + row.label : row.key;
    if (ranked && includeUniqueOrders) {
      return [index + 1, name, row.count, row.uniqueOrders, formatComplaintPercent_(row.percent)];
    }

    if (ranked) {
      return [index + 1, name, row.count, formatComplaintPercent_(row.percent)];
    }

    return [name, row.count, formatComplaintPercent_(row.percent)];
  });
}

function buildComplaintInsightRows_(stats) {
  var top5SkuOrders = {};
  stats.skuRows.slice(0, 5).forEach(function(row) {
    var skuSummary = stats.skus[row.key];
    Object.keys(skuSummary.orders || {}).forEach(function(orderId) {
      top5SkuOrders[orderId] = true;
    });
  });
  var top5SkuUniqueOrders = Object.keys(top5SkuOrders).length;

  var topCustomer = stats.customerRows[0] || {};
  var topPicker = stats.pickerRows[0] || {};

  return [
    [(stats.topCategory.key || "No category") + " contributes " + formatComplaintPercent_(stats.topCategory.percent || 0) + " of unique complaint orders."],
    [(stats.topRca.key || "No RCA") + " contributes " + formatComplaintPercent_(stats.topRca.percent || 0) + " of unique complaint orders."],
    ["Top 5 SKUs contribute " + formatComplaintPercent_(stats.uniqueOrdersWithComplaints > 0 ? (top5SkuUniqueOrders / stats.uniqueOrdersWithComplaints) * 100 : 0) + " of unique complaint orders."],
    ["Complaint Rate is " + formatComplaintPercent_(stats.complaintRate) + " based on unique complaint orders over total RDB orders."],
    ["Most impacted customer has " + (topCustomer.uniqueOrders || 0) + " unique complaint orders."],
    ["Most impacted picker has " + ((topPicker && topPicker.uniqueOrders) || 0) + " unique complaint orders."]
  ];
}

function addComplaintSection_(output, title, header, rows) {
  output.push([title]);
  output.push(header);

  if (!rows || rows.length === 0) {
    output.push(["No data"]);
  } else {
    rows.forEach(function(row) {
      output.push(row);
    });
  }

  output.push([""]);
}

function formatComplaintPercent_(value) {
  return formatNumber_(round2_(toNumber_(value)), 2) + "%";
}

function formatComplaintIntelligenceSheet_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var maxColumns = getMaxRowWidth_(rows);
  var rowCount = rows.length;
  ensureSheetSize_(sheet, rowCount + 2, maxColumns + 1);

  var usedRange = sheet.getRange(1, 1, rowCount, maxColumns);
  usedRange
    .setBackground(DASHBOARD_THEME.BACKGROUND)
    .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(false, false, false, false, false, false);

  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(DASHBOARD_THEME.BACKGROUND);
  sheet.setRowHeights(1, rowCount, 32);

  sheet.getRange(1, 1, 1, maxColumns)
    .merge()
    .setBackground(DASHBOARD_THEME.NAVY)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(22)
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 44);

  for (var i = 1; i < rows.length; i++) {
    if (isBlankRow_(rows[i])) continue;

    if (isComplaintSectionTitleRow_(rows, i)) {
      var endIndex = i;
      while (endIndex + 1 < rows.length && !isBlankRow_(rows[endIndex + 1])) {
        endIndex++;
      }

      formatComplaintSectionRange_(sheet, rows, i + 1, endIndex + 1, maxColumns);
      i = endIndex;
    }
  }

  sheet.autoResizeColumns(1, Math.min(maxColumns, 15));
}

function isComplaintSectionTitleRow_(rows, rowIndex) {
  if (rowIndex + 1 >= rows.length) return false;
  if (isBlankRow_(rows[rowIndex])) return false;
  return !isBlankRow_(rows[rowIndex + 1]);
}

function formatComplaintSectionRange_(sheet, rows, startRow, endRow, maxColumns) {
  var rowCount = endRow - startRow + 1;

  sheet.getRange(startRow, 1, rowCount, maxColumns)
    .setBorder(true, true, true, true, true, true, DASHBOARD_THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(startRow, 1, 1, maxColumns)
    .setBackground(DASHBOARD_THEME.SECTION)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(15);
  sheet.setRowHeight(startRow, 38);

  if (rowCount > 1) {
    sheet.getRange(startRow + 1, 1, 1, maxColumns)
      .setBackground(DASHBOARD_THEME.SECTION)
      .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
      .setFontWeight("bold")
      .setFontSize(11);
  }

  if (rowCount > 2) {
    var bodyRows = rowCount - 2;
    var backgrounds = [];
    for (var r = 0; r < bodyRows; r++) {
      var rowBackgrounds = [];
      for (var c = 0; c < maxColumns; c++) {
        rowBackgrounds.push(r % 2 === 0 ? DASHBOARD_THEME.CARD : "#1E293B");
      }
      backgrounds.push(rowBackgrounds);
    }

    sheet.getRange(startRow + 2, 1, bodyRows, maxColumns)
      .setBackgrounds(backgrounds)
      .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
      .setFontSize(10)
      .setFontWeight("normal")
      .setBorder(true, true, true, true, true, true, DASHBOARD_THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
}

function buildTopList_(data, options) {
  var summary = {};
  var allowedValues = buildAllowedValues_(options.filterValues);

  for (var i = 1; i < data.length; i++) {
    if (options.filterColumn !== undefined) {
      var filterValue = normalizeForCompare_(data[i][options.filterColumn]);
      if (!allowedValues[filterValue]) continue;
    }

    var key = normalizeText_(data[i][options.keyColumn]);
    if (!key) continue;

    if (!summary[key]) {
      summary[key] = {
        label: options.labelColumn === undefined ? null : data[i][options.labelColumn],
        count: 0
      };
    }

    summary[key].count++;
  }

  var rows = Object.keys(summary).map(function(key) {
    return {
      key: key,
      label: summary[key].label,
      count: summary[key].count
    };
  }).sort(function(a, b) {
    return b.count - a.count || a.key.localeCompare(b.key);
  });

  var output = [options.header];

  rows.slice(0, options.limit).forEach(function(row, index) {
    if (options.labelColumn === undefined) {
      output.push([index + 1, row.key, row.count]);
    } else {
      output.push([index + 1, row.key, row.label, row.count]);
    }
  });

  return output;
}

function buildTopUniqueCountList_(data, options) {
  var summary = {};
  var allowedValues = buildAllowedValues_(options.filterValues);

  for (var i = 1; i < data.length; i++) {
    if (options.filterColumn !== undefined) {
      var filterValue = normalizeForCompare_(data[i][options.filterColumn]);
      if (!allowedValues[filterValue]) continue;
    }

    var key = normalizeText_(data[i][options.keyColumn]);
    var uniqueValue = normalizeText_(data[i][options.uniqueColumn]);

    if (!key || !uniqueValue) continue;

    if (!summary[key]) {
      summary[key] = {
        label: options.labelColumn === undefined ? "" : data[i][options.labelColumn],
        values: {}
      };
    }

    if (!summary[key].label && options.labelColumn !== undefined) {
      summary[key].label = data[i][options.labelColumn];
    }

    summary[key].values[uniqueValue] = true;
  }

  var rows = Object.keys(summary).map(function(key) {
    return {
      key: key,
      label: summary[key].label,
      count: Object.keys(summary[key].values).length
    };
  }).sort(function(a, b) {
    return b.count - a.count || a.key.localeCompare(b.key);
  });

  var output = [options.header];

  rows.slice(0, options.limit).forEach(function(row, index) {
    if (options.labelColumn === undefined) {
      output.push([index + 1, row.key, row.count]);
    } else {
      output.push([index + 1, row.key, row.label, row.count]);
    }
  });

  return output;
}

function buildDeliveryAgentComplaintOrderTable_(data, limit) {
  var allowedDaBuckets = buildAllowedValues_(["LM Bucket", "DA Mistake", "Transit Damage"]);
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var bucket = normalizeForCompare_(data[i][16]);
    var orderNr = normalizeText_(data[i][4]);
    var daId = normalizeText_(data[i][10]);
    var daName = normalizeText_(data[i][11]);

    if (!allowedDaBuckets[bucket] || !orderNr || !daId) continue;

    if (!summary[daId]) {
      summary[daId] = {
        daName: daName,
        orders: {}
      };
    }

    if (!summary[daId].daName && daName) {
      summary[daId].daName = daName;
    }

    summary[daId].orders[orderNr] = true;
  }

  var rows = Object.keys(summary).map(function(daId) {
    return {
      daId: daId,
      daName: summary[daId].daName,
      uniqueOrderCount: Object.keys(summary[daId].orders).length
    };
  }).sort(function(a, b) {
    return b.uniqueOrderCount - a.uniqueOrderCount || a.daId.localeCompare(b.daId);
  });

  var output = [["Rank", "DA ID", "Delivery Agent", "Unique Order Count"]];

  rows.slice(0, limit).forEach(function(row, index) {
    output.push([index + 1, row.daId, row.daName, row.uniqueOrderCount]);
  });

  return output;
}

function buildAllowedValues_(values) {
  var allowedValues = {};

  if (!values) return allowedValues;

  values.forEach(function(value) {
    allowedValues[normalizeForCompare_(value)] = true;
  });

  return allowedValues;
}

function buildCountTable_(data, options) {
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var key = normalizeText_(data[i][options.keyColumn]);
    if (!key) continue;
    summary[key] = (summary[key] || 0) + 1;
  }

  var output = [options.header];

  Object.keys(summary).sort(function(a, b) {
    return summary[b] - summary[a] || a.localeCompare(b);
  }).forEach(function(key) {
    output.push([key, summary[key]]);
  });

  return output;
}

function buildUniqueCountTable_(data, options) {
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var key = normalizeText_(data[i][options.keyColumn]);
    var uniqueValue = normalizeText_(data[i][options.uniqueColumn]);

    if (!key || !uniqueValue) continue;

    if (!summary[key]) {
      summary[key] = {};
    }

    summary[key][uniqueValue] = true;
  }

  var output = [options.header];

  Object.keys(summary).sort(function(a, b) {
    return Object.keys(summary[b]).length - Object.keys(summary[a]).length ||
      a.localeCompare(b);
  }).forEach(function(key) {
    output.push([key, Object.keys(summary[key]).length]);
  });

  return output;
}

function buildWeeklyAnalyticsBlock_(weekRange, pickerPerformance, complaints, bagDeviation) {
  var updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var output = [
    [
      "WEEKLY_ANALYTICS_WEEK",
      weekRange.weekKey,
      "Week Status",
      weekRange.status,
      "Week Start",
      weekRange.startKey,
      "Week End",
      weekRange.endKey,
      "Updated At",
      updatedAt
    ],
    [""]
  ];

  addWeeklySection_(output, "Top Orders from the Week", ["Rank", "Picker", "Unique Order Count"],
    topRows_(pickerPerformance, "uniqueOrderCount", false, 5).map(function(row, index) {
      return [index + 1, row.picker, row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Fastest Accepting Time of the Week", ["Rank", "Picker", "Avg AT", "Unique Order Count"],
    topRows_(pickerPerformance, "avgAcceptance", true, 5, true).map(function(row, index) {
      return [index + 1, row.picker, round2_(row.avgAcceptance), row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Fastest Picker", ["Rank", "Picker", "PT per Item", "Unique Order Count", "Item Count"],
    topRows_(pickerPerformance, "pickingPerItem", true, 5, true).map(function(row, index) {
      return [index + 1, row.picker, round2_(row.pickingPerItem), row.uniqueOrderCount, row.itemCount];
    }));

  addWeeklySection_(output, "Fastest P95", ["Rank", "Picker", "P95", "Unique Order Count"],
    topRows_(pickerPerformance, "p95", true, 5, true).map(function(row, index) {
      return [index + 1, row.picker, round2_(row.p95), row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Higher Picker Efficiency", ["Rank", "Picker", "Items per Picking Time", "Unique Order Count", "Item Count"],
    topRows_(pickerPerformance, "efficiency", false, 5, true).map(function(row, index) {
      return [index + 1, row.picker, round2_(row.efficiency), row.uniqueOrderCount, row.itemCount];
    }));

  addWeeklySection_(output, "Top 3 SKUs with Highest Complaint Count", ["Rank", "SKU ID", "SKU Name", "Unique Order Count"],
    complaints.skus.slice(0, 3).map(function(row, index) {
      return [index + 1, row.key, row.label, row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Picker with Highest Complaint Count", ["Rank", "Picker", "Unique Order Count"],
    complaints.pickers.slice(0, 5).map(function(row, index) {
      return [index + 1, row.key, row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Picker with Highest Bag Deviation", ["Rank", "Picker", "Avg Bag Deviation", "Unique Order Count"],
    bagDeviation.slice(0, 5).map(function(row, index) {
      return [index + 1, row.picker, round2_(row.avgBagDeviation), row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Weekly RCA Count with Order IDs", ["RCA", "Unique Order Count", "Order IDs"],
    complaints.rca.map(function(row) {
      return [row.key, row.uniqueOrderCount, row.orderIds];
    }));

  addWeeklySection_(output, "Top 3 Customer IDs with Most Complaints", ["Rank", "Customer ID", "Unique Order Count"],
    complaints.customers.slice(0, 3).map(function(row, index) {
      return [index + 1, row.key, row.uniqueOrderCount];
    }));

  addWeeklySection_(output, "Top 3 DA with Most Complaints under LM Bucket / DA Mistake / Transit Damage", ["Rank", "DA ID", "Delivery Agent", "Unique Order Count"],
    complaints.deliveryAgents.slice(0, 3).map(function(row, index) {
      return [index + 1, row.key, row.label, row.uniqueOrderCount];
    }));

  return output;
}

function addWeeklySection_(output, title, header, rows) {
  output.push([title]);
  output.push(header);

  if (rows.length === 0) {
    output.push(["No data"]);
  } else {
    rows.forEach(function(row) {
      output.push(row);
    });
  }

  output.push([""]);
}

function topRows_(rows, metric, ascending, limit, requirePositive) {
  return rows.filter(function(row) {
    return !requirePositive || row[metric] > 0;
  }).sort(function(a, b) {
    var metricCompare = ascending ? a[metric] - b[metric] : b[metric] - a[metric];

    if (metricCompare !== 0) return metricCompare;

    return String(a.picker || a.key || "").localeCompare(String(b.picker || b.key || ""));
  }).slice(0, limit);
}

function writeWeeklyAnalyticsBlocks_(sheet, blocks) {
  var marker = "WEEKLY_ANALYTICS_WEEK";
  var targetWeekKeys = {};
  var preservedRows = [];
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var skipCurrentBlock = false;

  blocks.forEach(function(block) {
    targetWeekKeys[block.weekKey] = true;
  });

  if (lastRow > 0 && lastColumn > 0) {
    var existingRows = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

    existingRows.forEach(function(row) {
      if (normalizeText_(row[0]) === marker) {
        skipCurrentBlock = Boolean(targetWeekKeys[normalizeText_(row[1])]);
      }

      if (!skipCurrentBlock) {
        preservedRows.push(trimTrailingCells_(row));
      }
    });
  }

  preservedRows = trimTrailingRows_(preservedRows);

  var rowsToWrite = [];

  blocks.forEach(function(block, index) {
    if (index > 0) {
      rowsToWrite.push([""]);
      rowsToWrite.push([""]);
    }

    rowsToWrite = rowsToWrite.concat(block.rows);
  });

  if (preservedRows.length > 0) {
    rowsToWrite.push([""]);
    rowsToWrite.push([""]);
    rowsToWrite = rowsToWrite.concat(preservedRows);
  }

  clearSheetContents_(sheet, 1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1));
  writePaddedRows_(sheet, 1, 1, rowsToWrite);
  formatWeeklyAnalyticsSheet_(sheet, rowsToWrite);
}

function writeWeeklyAnalyticsBlock_(sheet, weekKey, block) {
  var marker = "WEEKLY_ANALYTICS_WEEK";
  var preservedRows = [];
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var skipCurrentBlock = false;

  if (lastRow > 0 && lastColumn > 0) {
    var existingRows = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

    existingRows.forEach(function(row) {
      if (normalizeText_(row[0]) === marker) {
        skipCurrentBlock = normalizeText_(row[1]) === weekKey;
      }

      if (!skipCurrentBlock) {
        preservedRows.push(trimTrailingCells_(row));
      }
    });
  }

  preservedRows = trimTrailingRows_(preservedRows);

  if (preservedRows.length > 0) {
    preservedRows.push([""]);
    preservedRows.push([""]);
  }

  var rowsToWrite = preservedRows.concat(block);

  clearSheetContents_(sheet, 1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1));
  writePaddedRows_(sheet, 1, 1, rowsToWrite);
  formatWeeklyAnalyticsSheet_(sheet, rowsToWrite);
}

function writePaddedRows_(sheet, row, column, values) {
  if (!values || values.length === 0) return;

  var maxColumns = 1;

  values.forEach(function(valueRow) {
    if (valueRow.length > maxColumns) maxColumns = valueRow.length;
  });

  var paddedValues = values.map(function(valueRow) {
    var rowCopy = valueRow.slice();

    while (rowCopy.length < maxColumns) {
      rowCopy.push("");
    }

    return rowCopy;
  });

  sheet.getRange(row, column, paddedValues.length, maxColumns).setValues(paddedValues);
}

function formatWeeklyAnalyticsSheet_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var maxColumns = getMaxRowWidth_(rows);
  var rowCount = rows.length;

  ensureSheetSize_(sheet, rowCount + 2, maxColumns + 1);

  var usedRange = sheet.getRange(1, 1, rowCount, maxColumns);
  usedRange
    .setBackground(DASHBOARD_THEME.BACKGROUND)
    .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
    .setFontWeight("normal")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setFontFamily("Arial")
    .setBorder(false, false, false, false, false, false);

  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(DASHBOARD_THEME.BACKGROUND);
  sheet.setRowHeights(1, rowCount, 32);

  for (var i = 0; i < rows.length; i++) {
    if (isBlankRow_(rows[i])) continue;

    var firstCell = normalizeText_(rows[i][0]);

    if (firstCell === "WEEKLY_ANALYTICS_WEEK") {
      formatWeeklyTableRange_(sheet, i + 1, i + 1, maxColumns, true);
      continue;
    }

    if (isWeeklySectionTitleRow_(rows, i)) {
      var endIndex = i;

      while (endIndex + 1 < rows.length && !isBlankRow_(rows[endIndex + 1])) {
        endIndex++;
      }

      formatWeeklyTableRange_(sheet, i + 1, endIndex + 1, maxColumns, false);
      i = endIndex;
    }
  }

  usedRange
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.autoResizeColumns(1, Math.min(maxColumns, 12));
}

function formatWeeklyTableRange_(sheet, startRow, endRow, maxColumns, isWeekHeader) {
  var rowCount = endRow - startRow + 1;
  var range = sheet.getRange(startRow, 1, rowCount, maxColumns);

  range
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setFontFamily("Arial")
    .setBorder(true, true, true, true, true, true, DASHBOARD_THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  if (isWeekHeader) {
    range
      .setBackground(DASHBOARD_THEME.NAVY)
      .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
      .setFontWeight("bold")
      .setFontSize(12);
    sheet.setRowHeight(startRow, 40);
    return;
  }

  sheet.getRange(startRow, 1, 1, maxColumns)
    .setBackground(DASHBOARD_THEME.SECTION)
    .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
    .setFontWeight("bold")
    .setFontSize(15);
  sheet.setRowHeight(startRow, 38);

  if (rowCount > 1) {
    sheet.getRange(startRow + 1, 1, 1, maxColumns)
      .setBackground(DASHBOARD_THEME.SECTION)
      .setFontColor(DASHBOARD_THEME.PRIMARY_TEXT)
      .setFontWeight("bold")
      .setFontSize(11);
  }

  if (rowCount > 2) {
    var bodyRange = sheet.getRange(startRow + 2, 1, rowCount - 2, maxColumns);
    var backgrounds = [];

    for (var r = 0; r < rowCount - 2; r++) {
      var rowBackgrounds = [];
      for (var c = 0; c < maxColumns; c++) {
        rowBackgrounds.push(r % 2 === 0 ? DASHBOARD_THEME.CARD : "#1E293B");
      }
      backgrounds.push(rowBackgrounds);
    }

    bodyRange
      .setBackgrounds(backgrounds)
      .setFontColor(DASHBOARD_THEME.SECONDARY_TEXT)
      .setFontWeight("normal")
      .setFontSize(10)
      .setBorder(true, true, true, true, true, true, DASHBOARD_THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
}

function isWeeklySectionTitleRow_(rows, rowIndex) {
  if (rowIndex + 1 >= rows.length) return false;
  if (isBlankRow_(rows[rowIndex])) return false;
  if (normalizeText_(rows[rowIndex][0]) === "WEEKLY_ANALYTICS_WEEK") return false;

  return !isBlankRow_(rows[rowIndex + 1]);
}

function getMaxRowWidth_(rows) {
  var maxColumns = 1;

  rows.forEach(function(row) {
    if (row.length > maxColumns) maxColumns = row.length;
  });

  return maxColumns;
}

function calculateMtdScore_(metrics) {
  var ordersScore = scoreHigherIsBetter_(metrics.avgOrdersPerDay, 20, 50, 30);
  var putawayScore = scoreHigherIsBetter_(metrics.avgPutawayPerDay, 200, 500, 20);
  var complaintScore = toNumber_(metrics.complaintScore);
  var bagDevScore = scoreLowerIsBetter_(Math.abs(metrics.avgBagDeviation), 0, 2, 5);
  var p95Score = scoreLowerIsBetter_(metrics.avgP95, 1.8, 3, 10);
  var atScore = scoreLowerIsBetter_(metrics.avgAT, 4, 8, 5);
  var ptScore = scoreLowerIsBetter_(metrics.avgPT, 10, 18, 10);

  return ordersScore +
    putawayScore +
    complaintScore +
    bagDevScore +
    p95Score +
    atScore +
    ptScore +
    metrics.merchandisingScore;
}

function calculateMtdComplaintScore_(complaintPercent) {
  var percent = toNumber_(complaintPercent);

  if (percent <= MTD_COMPLAINT_SCORING.SAFE_MAX_PERCENT) {
    return scoreLowerIsBetter_(
      percent,
      0,
      MTD_COMPLAINT_SCORING.SAFE_MAX_PERCENT,
      MTD_COMPLAINT_SCORING.SAFE_POINTS
    );
  }

  if (percent < MTD_COMPLAINT_SCORING.WARNING_MAX_PERCENT) {
    return MTD_COMPLAINT_SCORING.WARNING_PENALTY;
  }

  if (percent <= MTD_COMPLAINT_SCORING.HIGH_MAX_PERCENT) {
    return MTD_COMPLAINT_SCORING.HIGH_PENALTY;
  }

  return MTD_COMPLAINT_SCORING.CRITICAL_PENALTY;
}

function scoreHigherIsBetter_(value, zeroAt, fullAt, points) {
  if (value >= fullAt) return points;
  if (value <= zeroAt) return 0;
  return ((value - zeroAt) / (fullAt - zeroAt)) * points;
}

function scoreLowerIsBetter_(value, fullAt, zeroAt, points) {
  if (value <= fullAt) return points;
  if (value >= zeroAt) return 0;
  return (1 - ((value - fullAt) / (zeroAt - fullAt))) * points;
}

function calculatePercentile_(values, percentile) {
  if (!values || values.length === 0) return 0;

  var sorted = values.slice().sort(function(a, b) {
    return a - b;
  });

  var index = Math.ceil(percentile * sorted.length) - 1;
  return sorted[Math.max(0, Math.min(index, sorted.length - 1))];
}

function normalizeDate_(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  var text = String(value).trim();
  if (!text) return "";

  var flexibleDate = parseFlexibleDateText_(text);
  if (flexibleDate) {
    return Utilities.formatDate(flexibleDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return text;
}

function getActiveMtdMonthKey_(data) {
  var closedMonthKey = getClosedMtdMonthKey_();
  var activeMonthKey = "";

  for (var i = 1; i < data.length; i++) {
    var monthKey = getMonthKey_(normalizeDate_(data[i][0]));

    if (!monthKey || monthKey <= closedMonthKey) continue;

    if (!activeMonthKey || monthKey > activeMonthKey) {
      activeMonthKey = monthKey;
    }
  }

  return activeMonthKey;
}

function getClosedMtdMonthKey_() {
  var closedMonthKey = PropertiesService.getScriptProperties()
    .getProperty(CONFIG.PROPERTIES.CLOSED_MTD_MONTH_KEY) || "";

  if (closedMonthKey >= getCurrentMonthKey_()) {
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.PROPERTIES.CLOSED_MTD_MONTH_KEY);
    return "";
  }

  return closedMonthKey;
}

function getLatestCompletedMonthKey_(data) {
  var currentMonthKey = getCurrentMonthKey_();
  var latestCompletedMonthKey = "";

  for (var i = 1; i < data.length; i++) {
    var monthKey = getMonthKey_(normalizeDate_(data[i][0]));

    if (!monthKey || monthKey >= currentMonthKey) continue;

    if (!latestCompletedMonthKey || monthKey > latestCompletedMonthKey) {
      latestCompletedMonthKey = monthKey;
    }
  }

  return latestCompletedMonthKey;
}

function setClosedMtdMonthKey_(monthKey) {
  PropertiesService.getScriptProperties()
    .setProperty(CONFIG.PROPERTIES.CLOSED_MTD_MONTH_KEY, monthKey);
}

function isInMonth_(dateKey, monthKey) {
  return monthKey !== "" && getMonthKey_(dateKey) === monthKey;
}

function getMonthKey_(dateKey) {
  if (!dateKey) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
    return dateKey.substring(0, 7);
  }

  var date = new Date(dateKey);
  if (isNaN(date.getTime())) return "";

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM");
}

function getCurrentMonthKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");
}

function getWeekRange_(dateKey) {
  var date = parseDateKey_(dateKey);
  if (!date) return null;

  var start = new Date(date.getTime());
  start.setDate(date.getDate() - date.getDay());

  var end = new Date(start.getTime());
  end.setDate(start.getDate() + 6);

  var startKey = formatDateKey_(start);
  var endKey = formatDateKey_(end);

  return {
    startKey: startKey,
    endKey: endKey,
    weekKey: startKey + "_to_" + endKey
  };
}

function getPreviousAndCurrentWeekRanges_() {
  var currentWeek = getWeekRange_(formatDateKey_(new Date()));
  var previousWeekDate = parseDateKey_(currentWeek.startKey);

  previousWeekDate.setDate(previousWeekDate.getDate() - 1);

  var previousWeek = getWeekRange_(formatDateKey_(previousWeekDate));

  previousWeek.status = "Previous Completed Week";
  currentWeek.status = "Current Running Week";

  return [currentWeek, previousWeek];
}

function parseDateKey_(dateKey) {
  if (!dateKey) return null;

  var date;

  if (/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
    var parts = dateKey.split("-");
    date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  } else {
    date = new Date(dateKey);
  }

  return isNaN(date.getTime()) ? null : date;
}

function parseFlexibleDateText_(text) {
  var normalized = String(text).trim().replace(/(\d)(st|nd|rd|th)\b/gi, "$1");
  var numericMatch = normalized.match(/^(\d{1,2})[\/.\- ](\d{1,2})[\/.\- ](\d{2,4})$/);

  if (numericMatch) {
    return buildDateIfValid_(
      expandYear_(Number(numericMatch[3])),
      Number(numericMatch[2]) - 1,
      Number(numericMatch[1])
    );
  }

  var dayMonthMatch = normalized.match(/^(\d{1,2})[\/.\- ]([A-Za-z]{3,9})(?:[\/.\- ](\d{2,4}))?$/);
  if (dayMonthMatch) {
    return buildDateIfValid_(
      dayMonthMatch[3] ? expandYear_(Number(dayMonthMatch[3])) : new Date().getFullYear(),
      monthIndex_(dayMonthMatch[2]),
      Number(dayMonthMatch[1])
    );
  }

  var monthDayMatch = normalized.match(/^([A-Za-z]{3,9})[\/.\- ](\d{1,2})(?:[\/.\- ](\d{2,4}))?$/);
  if (monthDayMatch) {
    return buildDateIfValid_(
      monthDayMatch[3] ? expandYear_(Number(monthDayMatch[3])) : new Date().getFullYear(),
      monthIndex_(monthDayMatch[1]),
      Number(monthDayMatch[2])
    );
  }

  return null;
}

function buildDateIfValid_(year, monthIndex, day) {
  if (monthIndex < 0 || monthIndex > 11 || day < 1 || day > 31) return null;

  var date = new Date(year, monthIndex, day);

  return date.getFullYear() === year &&
    date.getMonth() === monthIndex &&
    date.getDate() === day
    ? date
    : null;
}

function expandYear_(year) {
  return year < 100 ? 2000 + year : year;
}

function monthIndex_(monthName) {
  var months = {
    jan: 0,
    january: 0,
    feb: 1,
    february: 1,
    mar: 2,
    march: 2,
    apr: 3,
    april: 3,
    may: 4,
    jun: 5,
    june: 5,
    jul: 6,
    july: 6,
    aug: 7,
    august: 7,
    sep: 8,
    sept: 8,
    september: 8,
    oct: 9,
    october: 9,
    nov: 10,
    november: 10,
    dec: 11,
    december: 11
  };

  var key = normalizeForCompare_(monthName);

  return months[key] === undefined ? -1 : months[key];
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function addDaysToDateKey_(dateKey, days) {
  var date = parseDateKey_(dateKey);
  if (!date) return "";

  date.setDate(date.getDate() + days);

  return formatDateKey_(date);
}

function isInDateRange_(dateKey, startKey, endKey) {
  // Inclusive Sunday-through-Saturday boundary check for normalized yyyy-MM-dd values.
  return dateKey >= startKey && dateKey <= endKey;
}

function normalizeText_(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function normalizeForCompare_(value) {
  return normalizeText_(value).toLowerCase().replace(/\s+/g, " ");
}

function getMtdPickerName_(picker) {
  var pickerName = normalizeText_(picker);
  var aliasKey = normalizeForCompare_(pickerName);

  return CONFIG.MTD_PICKER_ALIASES[aliasKey] || pickerName;
}

function toNumber_(value) {
  var number = Number(value);
  return isNaN(number) ? 0 : number;
}

function round2_(value) {
  return Math.round(value * 100) / 100;
}

function makeKey_(dateKey, picker) {
  return dateKey + "|" + picker;
}

function getRequiredSheet_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Missing required sheet: " + sheetName);
  }
  return sheet;
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  return sheet || spreadsheet.insertSheet(sheetName);
}

function ensureSheetSize_(sheet, minRows, minColumns) {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < minColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minColumns - sheet.getMaxColumns());
  }
}

function clearSheetContents_(sheet, row, column, numRows, numColumns) {
  if (numRows > 0 && numColumns > 0) {
    sheet.getRange(row, column, numRows, numColumns).clearContent();
  }
}

function writeTable_(sheet, row, column, values) {
  if (!values || values.length === 0) return;
  sheet.getRange(row, column, values.length, values[0].length).setValues(values);
}

function trimTrailingCells_(row) {
  var trimmed = row.slice();

  while (trimmed.length > 1 && normalizeText_(trimmed[trimmed.length - 1]) === "") {
    trimmed.pop();
  }

  return trimmed;
}

function trimTrailingRows_(rows) {
  var trimmed = rows.slice();

  while (trimmed.length > 0 && isBlankRow_(trimmed[trimmed.length - 1])) {
    trimmed.pop();
  }

  return trimmed;
}

function isBlankRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (normalizeText_(row[i]) !== "") return false;
  }

  return true;
}
