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
  BAU_SOURCE_ID: "1tsolH7eezhIbe5rUgzrROIfh18q6ym_XdXao48R2ra8",

  SHEETS: {
    RDB: "RDB",
    RDB_HOURLY: "RDB Hourly",
    MTD_DASHBOARD: "MTD_DASHBOARD",
    COMPLAINT_ANALYSIS: "COMPLAINT_ANALYSIS",
    RCA_DATABASE: "YEARLY_RCA_DATABASE",
    WEEKLY_ANALYTICS: "Weekly Analytics",
    DARK_STORE_GRAPHS: "Dark Store Graphs",
    BAU_LIVE: "BAU_LIVE",
    BAU_HISTORY: "BAU_HISTORY",
    BAU_RAW: "BAU_RAW"
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

var BAU_TRACKING = {
  SOURCE_SHEET: "Metrics View D-1",
  DS_CODE: "RUHDS48",
  TARGET: 1,
  KPI_NAMES: [
    "Offload AM",
    "Offload CHL",
    "Offload FZR",
    "Offload UF",
    "Putaway AM",
    "Putaway CHL",
    "Putaway FZR",
    "Putaway UF",
    "FSQ",
    "Expiry Adherence",
    "Chiller Movement",
    "Cycle Count"
  ]
};

var GRAPH_LAYOUT = {
  ORDER_BAR_ROW: 4,
  ITEM_BAR_ROW: 40,
  ORDER_SOURCE_ROW: 130,
  ITEM_SOURCE_ROW: 170,
  PERFORMANCE_SOURCE_ROW: 210,
  DAILY_KPI_SOURCE_ROW: 250,
  TOP_PICKER_SOURCE_ROW: 280,
  CHART_COLUMN: 9,
  ORDER_CHART_ROW: 4,
  ITEM_CHART_ROW: 40,
  PERFORMANCE_CHART_ROW: 76,
  DAILY_KPI_CHART_ROW: 112,
  TOP_PICKER_CHART_ROW: 148
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
  updateBauTracking();
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
    .forSpreadsheet(CONFIG.BAU_SOURCE_ID)
    .onEdit()
    .create();

  ScriptApp.newTrigger("runAll")
    .forSpreadsheet(CONFIG.BAU_SOURCE_ID)
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
  var data = rcaSheet.getDataRange().getValues();

  clearSheetContents_(analysisSheet, 1, 1, Math.max(analysisSheet.getLastRow(), 1), 23);

  writeTable_(
    analysisSheet,
    1,
    1,
    buildTopUniqueCountList_(data, {
      header: ["Rank", "SKU ID", "SKU Name", "Unique Order Count"],
      limit: 5,
      keyColumn: 5,
      labelColumn: 8,
      uniqueColumn: 4
    })
  );

  writeTable_(
    analysisSheet,
    1,
    6,
    buildTopUniqueCountList_(data, {
      header: ["Rank", "Customer ID", "Unique Order Count"],
      limit: 5,
      keyColumn: 15,
      uniqueColumn: 4
    })
  );

  writeTable_(
    analysisSheet,
    1,
    10,
    buildDeliveryAgentComplaintOrderTable_(data, 5)
  );

  writeTable_(
    analysisSheet,
    1,
    15,
    buildUniqueCountTable_(data, {
      header: ["RCA", "Unique Order Count"],
      keyColumn: 16,
      uniqueColumn: 4
    })
  );

  writeTable_(
    analysisSheet,
    1,
    18,
    buildUniqueCountTable_(data, {
      header: ["Defect Category", "Unique Order Count"],
      keyColumn: 7,
      uniqueColumn: 4
    })
  );

  writeTable_(
    analysisSheet,
    1,
    21,
    buildTopList_(data, {
      header: ["Rank", "Complaint Reason", "Count"],
      limit: 10,
      keyColumn: 9
    })
  );
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
  var hourlySheet = getRequiredSheet_(ss, CONFIG.SHEETS.RDB_HOURLY);
  var mtdSheet = getRequiredSheet_(ss, CONFIG.SHEETS.MTD_DASHBOARD);
  var rdbSheet = getRequiredSheet_(ss, CONFIG.SHEETS.RDB);

  prepareGraphsSheet_(graphsSheet);

  var allHourlyRows = getAllHourlyRows_(hourlySheet);
  var hourlyRows = getLatestHourlyRowsFromRows_(allHourlyRows);
  var latestDate = hourlyRows.length > 0 ? hourlyRows[0].dateKey : "";
  var comparisonDate = latestDate ? addDaysToDateKey_(latestDate, -7) : "";
  var mtdRows = getMtdDashboardRows_(mtdSheet);
  var orderMatrix = buildHourlyMetricComparisonMatrix_(allHourlyRows, "orders", latestDate, comparisonDate);
  var itemMatrix = buildHourlyMetricComparisonMatrix_(allHourlyRows, "items", latestDate, comparisonDate);
  var performanceRows = buildHourlyPerformanceRows_(hourlyRows);
  var dailyComparisonRows = buildDailyComparisonRows_(rdbSheet, latestDate, comparisonDate);

  renderGraphsHeader_(graphsSheet, latestDate, comparisonDate);
  renderTopBarViews_(graphsSheet, orderMatrix, itemMatrix);
  renderGraphSourceTables_(graphsSheet, orderMatrix, itemMatrix, performanceRows, dailyComparisonRows, mtdRows);
  insertDashboardCharts_(graphsSheet, hourlyRows, mtdRows, orderMatrix, itemMatrix, performanceRows, dailyComparisonRows);
}

function updateBauTracking() {
  var sourceSS = SpreadsheetApp.openById(CONFIG.BAU_SOURCE_ID);
  var destSS = SpreadsheetApp.openById(CONFIG.DEST_ID);
  var sourceSheet = getRequiredSheet_(sourceSS, BAU_TRACKING.SOURCE_SHEET);
  var liveSheet = getOrCreateSheet_(destSS, CONFIG.SHEETS.BAU_LIVE);
  var historySheet = getOrCreateSheet_(destSS, CONFIG.SHEETS.BAU_HISTORY);
  var rawSheet = getOrCreateSheet_(destSS, CONFIG.SHEETS.BAU_RAW);
  var sourceData = sourceSheet.getDataRange().getValues();

  if (!sourceData || sourceData.length === 0) {
    throw new Error("BAU source sheet has no data: " + BAU_TRACKING.SOURCE_SHEET);
  }

  var headers = sourceData[0].map(function(header) {
    return normalizeText_(header);
  });
  var headerMap = buildBauHeaderMap_(headers);

  validateBauSourceHeaders_(headerMap);

  var dsCodeIndex = headerMap[normalizeForCompare_("DS Code")];
  var filteredRows = [];

  for (var i = 1; i < sourceData.length; i++) {
    if (normalizeForCompare_(sourceData[i][dsCodeIndex]) === normalizeForCompare_(BAU_TRACKING.DS_CODE)) {
      filteredRows.push(sourceData[i]);
    }
  }

  if (filteredRows.length === 0) {
    throw new Error("No BAU source rows found for DS Code " + BAU_TRACKING.DS_CODE);
  }

  renderBauRaw_(rawSheet, headers, filteredRows);

  var todayKey = formatDateKey_(new Date());
  var snapshot = buildBauSnapshot_(headerMap, filteredRows, todayKey);

  upsertBauHistorySnapshot_(historySheet, snapshot);
  renderBauLiveDashboard_(liveSheet, readBauHistoryRows_(historySheet), todayKey);
}

function buildBauHeaderMap_(headers) {
  var headerMap = {};

  headers.forEach(function(header, index) {
    var key = normalizeForCompare_(header);

    if (key && headerMap[key] === undefined) {
      headerMap[key] = index;
    }
  });

  return headerMap;
}

function validateBauSourceHeaders_(headerMap) {
  var requiredHeaders = ["DS Code"].concat(BAU_TRACKING.KPI_NAMES);
  var missingHeaders = [];

  requiredHeaders.forEach(function(header) {
    if (headerMap[normalizeForCompare_(header)] === undefined) {
      missingHeaders.push(header);
    }
  });

  if (missingHeaders.length > 0) {
    throw new Error("Missing BAU source columns: " + missingHeaders.join(", "));
  }
}

function renderBauRaw_(sheet, headers, rows) {
  var output = [headers].concat(rows);
  var rowCount = output.length;
  var columnCount = Math.max(headers.length, 1);

  ensureSheetSize_(sheet, Math.max(rowCount + 2, 5), columnCount + 1);
  sheet.clear();
  sheet.getRange(1, 1, rowCount, columnCount).setValues(output);
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, 1, columnCount)
    .setBackground("#0b1f3a")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.autoResizeColumns(1, Math.min(columnCount, 20));
  sheet.hideSheet();
}

function buildBauSnapshot_(headerMap, rows, dateKey) {
  var metrics = {};
  var numericValues = [];
  var missedCount = 0;

  BAU_TRACKING.KPI_NAMES.forEach(function(kpiName) {
    var columnIndex = headerMap[normalizeForCompare_(kpiName)];
    var total = 0;
    var count = 0;

    rows.forEach(function(row) {
      var value = parseBauPercent_(row[columnIndex]);

      if (value !== null) {
        total += value;
        count++;
      }
    });

    var average = count > 0 ? total / count : "";
    metrics[kpiName] = average;

    if (average === "" || average < BAU_TRACKING.TARGET) {
      missedCount++;
    }

    numericValues.push(average === "" ? 0 : Math.min(average, BAU_TRACKING.TARGET));
  });

  return {
    dateKey: dateKey,
    monthKey: getMonthKey_(dateKey),
    dsCode: BAU_TRACKING.DS_CODE,
    metrics: metrics,
    adherence: averageBauValues_(numericValues),
    missedCount: missedCount,
    updatedAt: new Date()
  };
}

function getBauHistoryHeader_() {
  return ["Date", "Month", "DS Code"]
    .concat(BAU_TRACKING.KPI_NAMES)
    .concat(["Today's Adherence", "Missed KPI Count", "Snapshot Updated At"]);
}

function buildBauHistoryRow_(snapshot) {
  var row = [
    snapshot.dateKey,
    snapshot.monthKey,
    snapshot.dsCode
  ];

  BAU_TRACKING.KPI_NAMES.forEach(function(kpiName) {
    row.push(snapshot.metrics[kpiName]);
  });

  row.push(snapshot.adherence);
  row.push(snapshot.missedCount);
  row.push(snapshot.updatedAt);

  return row;
}

function upsertBauHistorySnapshot_(sheet, snapshot) {
  var header = getBauHistoryHeader_();
  var snapshotRow = buildBauHistoryRow_(snapshot);
  var lastRow = sheet.getLastRow();
  var dateRows = [];

  ensureSheetSize_(sheet, Math.max(lastRow, 1) + 2, header.length + 1);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (lastRow >= 2) {
    var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    dates.forEach(function(row, index) {
      if (normalizeDate_(row[0]) === snapshot.dateKey) {
        dateRows.push(index + 2);
      }
    });
  }

  var targetRow = dateRows.length > 0 ? dateRows[0] : Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(targetRow, 1, 1, snapshotRow.length).setValues([snapshotRow]);

  for (var i = dateRows.length - 1; i >= 1; i--) {
    sheet.deleteRow(dateRows[i]);
  }

  formatBauHistorySheet_(sheet);
}

function readBauHistoryRows_(sheet) {
  var lastRow = sheet.getLastRow();
  var headerLength = getBauHistoryHeader_().length;

  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, headerLength).getValues();
  var rows = [];

  data.forEach(function(row) {
    var dateKey = normalizeDate_(row[0]);

    if (!dateKey) return;

    var metrics = {};
    BAU_TRACKING.KPI_NAMES.forEach(function(kpiName, index) {
      metrics[kpiName] = parseBauPercent_(row[3 + index]);
    });

    rows.push({
      dateKey: dateKey,
      monthKey: normalizeText_(row[1]) || getMonthKey_(dateKey),
      dsCode: normalizeText_(row[2]),
      metrics: metrics,
      adherence: parseBauPercent_(row[3 + BAU_TRACKING.KPI_NAMES.length]),
      missedCount: toNumber_(row[4 + BAU_TRACKING.KPI_NAMES.length]),
      updatedAt: row[5 + BAU_TRACKING.KPI_NAMES.length]
    });
  });

  rows.sort(function(a, b) {
    return a.dateKey.localeCompare(b.dateKey);
  });

  return rows;
}

function renderBauLiveDashboard_(sheet, historyRows, todayKey) {
  resetBauLiveSheet_(sheet);

  var currentMonthKey = getMonthKey_(todayKey);
  var currentMonthRows = historyRows.filter(function(row) {
    return row.monthKey === currentMonthKey;
  });
  var todayRow = null;

  historyRows.forEach(function(row) {
    if (row.dateKey === todayKey) todayRow = row;
  });

  var mtdByKpi = buildBauMtdByKpi_(currentMonthRows);
  var missedRows = buildBauMissedRows_(currentMonthRows);
  var todayAdherence = todayRow ? calculateBauRowAdherence_(todayRow.metrics) : 0;
  var monthAdherence = calculateBauMonthAdherence_(currentMonthRows);
  var updatedAt = todayRow && todayRow.updatedAt
    ? todayRow.updatedAt
    : new Date();

  renderBauExecutiveHeader_(sheet, currentMonthKey, updatedAt);
  renderBauSummaryCards_(sheet, todayAdherence, monthAdherence, todayRow ? todayRow.missedCount : BAU_TRACKING.KPI_NAMES.length, missedRows.length);
  renderBauMtdTable_(sheet, mtdByKpi);
  renderBauMissedTracker_(sheet, missedRows);
  formatBauLiveSheet_(sheet, mtdByKpi.length, missedRows.length);
}

function resetBauLiveSheet_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  ensureSheetSize_(sheet, 80, 12);
  sheet.setFrozenRows(3);
  sheet.setHiddenGridlines(true);
}

function renderBauExecutiveHeader_(sheet, monthKey, updatedAt) {
  sheet.getRange("A1:H1").merge().setValue("BAU EXECUTIVE DASHBOARD");
  sheet.getRange("A2:H2").merge().setValue(
    "DS Code: " + BAU_TRACKING.DS_CODE +
    " | Current Month: " + formatBauMonthLabel_(monthKey) +
    " | Target: 100% for every KPI" +
    " | Updated: " + formatBauTimestamp_(updatedAt)
  );
}

function renderBauSummaryCards_(sheet, todayAdherence, monthAdherence, todayMissedCount, monthMissedCount) {
  var cards = [
    ["Today's Adherence", todayAdherence, "Target", BAU_TRACKING.TARGET],
    ["Current Month Adherence", monthAdherence, "Target", BAU_TRACKING.TARGET],
    ["Missed KPIs Today", todayMissedCount, "Open Misses MTD", monthMissedCount]
  ];

  sheet.getRange(4, 1, cards.length, cards[0].length).setValues(cards);
}

function renderBauMtdTable_(sheet, rows) {
  var output = [["KPI", "Target", "Today", "MTD Actual", "Gap", "Status"]];

  rows.forEach(function(row) {
    output.push([
      row.kpi,
      BAU_TRACKING.TARGET,
      row.todayValue,
      row.mtdValue,
      Math.max(BAU_TRACKING.TARGET - row.mtdValue, 0),
      row.mtdValue >= BAU_TRACKING.TARGET ? "MET" : "MISSED"
    ]);
  });

  sheet.getRange(9, 1, output.length, output[0].length).setValues(output);
}

function renderBauMissedTracker_(sheet, rows) {
  var startRow = 25;
  var output = [["Date", "KPI", "Actual", "Gap to 100%", "Status"]];

  if (rows.length === 0) {
    output.push(["Current month", "No missed BAU items", BAU_TRACKING.TARGET, 0, "MET"]);
  } else {
    rows.forEach(function(row) {
      output.push([
        row.dateKey,
        row.kpi,
        row.actual,
        row.gap,
        row.status
      ]);
    });
  }

  sheet.getRange(startRow, 1, output.length, output[0].length).setValues(output);
}

function buildBauMtdByKpi_(monthRows) {
  var todayKey = formatDateKey_(new Date());
  var todayRow = null;

  monthRows.forEach(function(row) {
    if (row.dateKey === todayKey) todayRow = row;
  });

  return BAU_TRACKING.KPI_NAMES.map(function(kpiName) {
    var total = 0;
    var count = monthRows.length;

    monthRows.forEach(function(row) {
      total += getBauScoreValue_(row.metrics[kpiName]);
    });

    return {
      kpi: kpiName,
      todayValue: todayRow ? getBauScoreValue_(todayRow.metrics[kpiName]) : 0,
      mtdValue: count > 0 ? total / count : 0
    };
  });
}

function buildBauMissedRows_(monthRows) {
  var missedRows = [];
  var sortedRows = monthRows.slice().sort(function(a, b) {
    return b.dateKey.localeCompare(a.dateKey);
  });

  sortedRows.forEach(function(row) {
    BAU_TRACKING.KPI_NAMES.forEach(function(kpiName) {
      var actual = row.metrics[kpiName];
      var scoreValue = getBauScoreValue_(actual);

      if (scoreValue < BAU_TRACKING.TARGET) {
        missedRows.push({
          dateKey: row.dateKey,
          kpi: kpiName,
          actual: scoreValue,
          gap: BAU_TRACKING.TARGET - scoreValue,
          status: actual === null ? "NO DATA" : "MISSED"
        });
      }
    });
  });

  return missedRows;
}

function calculateBauRowAdherence_(metrics) {
  var values = BAU_TRACKING.KPI_NAMES.map(function(kpiName) {
    return Math.min(getBauScoreValue_(metrics[kpiName]), BAU_TRACKING.TARGET);
  });

  return averageBauValues_(values);
}

function calculateBauMonthAdherence_(rows) {
  var values = [];

  rows.forEach(function(row) {
    BAU_TRACKING.KPI_NAMES.forEach(function(kpiName) {
      values.push(Math.min(getBauScoreValue_(row.metrics[kpiName]), BAU_TRACKING.TARGET));
    });
  });

  return averageBauValues_(values);
}

function getBauScoreValue_(value) {
  return value === null || value === "" || value === undefined
    ? 0
    : toNumber_(value);
}

function averageBauValues_(values) {
  if (!values || values.length === 0) return 0;

  var total = 0;
  values.forEach(function(value) {
    total += toNumber_(value);
  });

  return total / values.length;
}

function parseBauPercent_(value) {
  if (value === "" || value === null || value === undefined) return null;

  if (value instanceof Date) return null;

  if (typeof value === "string") {
    var text = value.trim();
    if (!text) return null;

    if (/%$/.test(text)) {
      var percentNumber = Number(text.replace("%", "").replace(/,/g, ""));
      return isNaN(percentNumber) ? null : percentNumber / 100;
    }

    var stringNumber = Number(text.replace(/,/g, ""));
    if (isNaN(stringNumber)) return null;

    return stringNumber > 1 ? stringNumber / 100 : stringNumber;
  }

  var number = Number(value);
  if (isNaN(number)) return null;

  return number > 1 ? number / 100 : number;
}

function formatBauHistorySheet_(sheet) {
  var header = getBauHistoryHeader_();
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var columnCount = header.length;

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnCount)
    .setBackground("#0b1f3a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  if (lastRow > 1) {
    sheet.getRange(2, 4, lastRow - 1, BAU_TRACKING.KPI_NAMES.length + 1).setNumberFormat("0.00%");
    sheet.getRange(2, 1, lastRow - 1, columnCount)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.autoResizeColumns(1, Math.min(columnCount, 20));
}

function formatBauLiveSheet_(sheet, mtdRowCount, missedRowCount) {
  var missedDisplayRowCount = Math.max(missedRowCount, 1);

  sheet.getRange("A1:H1")
    .setBackground("#0b1f3a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(18)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.getRange("A2:H2")
    .setBackground("#1f4e79")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 42);
  sheet.setRowHeight(2, 32);

  sheet.getRange(4, 1, 3, 4)
    .setBackground("#edf2f7")
    .setFontColor("#0f172a")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#0b1f3a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(4, 2, 2, 1).setNumberFormat("0.00%");
  sheet.getRange(4, 4, 2, 1).setNumberFormat("0.00%");

  sheet.getRange("A8:F8").merge().setValue("MTD ADHERENCE BY KPI");
  sheet.getRange("A8:F8")
    .setBackground("#0b1f3a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.getRange(9, 1, mtdRowCount + 1, 6)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(9, 1, 1, 6)
    .setBackground("#1f4e79")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange(10, 2, Math.max(mtdRowCount, 1), 4).setNumberFormat("0.00%");

  sheet.getRange("A24:E24").merge().setValue("MISSED BAU TRACKER");
  sheet.getRange("A24:E24")
    .setBackground("#7f1d1d")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.getRange(25, 1, missedDisplayRowCount + 1, 5)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(25, 1, 1, 5)
    .setBackground("#1f4e79")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange(26, 3, missedDisplayRowCount, 2).setNumberFormat("0.00%");

  applyBauStatusColors_(sheet, 10, 6, mtdRowCount);
  applyBauStatusColors_(sheet, 26, 5, missedDisplayRowCount);

  sheet.autoResizeColumns(1, 8);
  sheet.getRange(1, 1, sheet.getMaxRows(), 8).setFontFamily("Arial");
}

function applyBauStatusColors_(sheet, startRow, statusColumn, rowCount) {
  if (rowCount <= 0) return;

  var statusRange = sheet.getRange(startRow, statusColumn, rowCount, 1);
  var values = statusRange.getValues();
  var backgrounds = [];
  var fontColors = [];

  values.forEach(function(row) {
    var status = normalizeForCompare_(row[0]);

    if (status === "met") {
      backgrounds.push(["#dcfce7"]);
      fontColors.push(["#14532d"]);
    } else if (status === "no data") {
      backgrounds.push(["#fef3c7"]);
      fontColors.push(["#92400e"]);
    } else {
      backgrounds.push(["#fee2e2"]);
      fontColors.push(["#7f1d1d"]);
    }
  });

  statusRange
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors)
    .setFontWeight("bold");
}

function formatBauMonthLabel_(monthKey) {
  if (!monthKey) return "";

  var date = parseDateKey_(monthKey + "-01");
  return date
    ? Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy")
    : monthKey;
}

function formatBauTimestamp_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }

  return normalizeText_(value);
}

function resetDashboardSheet_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  ensureSheetSize_(sheet, 145, 45);
  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
}

function prepareGraphsSheet_(sheet) {
  ensureSheetSize_(sheet, 310, 45);
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

function renderGraphsHeader_(sheet, latestDate, comparisonDate) {
  sheet.getRange("A1").setValue("DARK STORE HOURLY GRAPHS");
  sheet.getRange("A2").setValue("Comparing latest day " + (latestDate || "-") + " vs same day last week " + (comparisonDate || "-") + " | Graphs update when runAll/createTrigger runs");
}

function renderTopBarViews_(sheet, orderMatrix, itemMatrix) {
  renderBarComparisonTable_(
    sheet,
    GRAPH_LAYOUT.ORDER_BAR_ROW,
    1,
    "Hourly Orders: Latest Day vs Same Day Last Week",
    orderMatrix,
    "#64748b",
    "#2563eb"
  );
  renderBarComparisonTable_(
    sheet,
    GRAPH_LAYOUT.ITEM_BAR_ROW,
    1,
    "Hourly Items: Latest Day vs Same Day Last Week",
    itemMatrix,
    "#64748b",
    "#16a34a"
  );
}

function renderBarComparisonTable_(sheet, startRow, startColumn, title, matrix, previousColor, latestColor) {
  if (!matrix || matrix.length < 2) return;

  var previousMax = getColumnMax_(matrix, 1);
  var latestMax = getColumnMax_(matrix, 2);
  var rows = [[
    "Hour",
    matrix[0][1],
    "Last Week Bar",
    matrix[0][2],
    "Latest Day Bar",
    "Variance",
    "Variance %"
  ]];

  for (var i = 1; i < matrix.length; i++) {
    var variancePercent = matrix[i][1] !== 0
      ? (matrix[i][3] / matrix[i][1]) * 100
      : "";

    rows.push([
      matrix[i][0],
      matrix[i][1],
      barText_(matrix[i][1], previousMax),
      matrix[i][2],
      barText_(matrix[i][2], latestMax),
      matrix[i][3],
      variancePercent === "" ? "" : round2_(variancePercent)
    ]);
  }

  ensureSheetSize_(sheet, startRow + rows.length + 2, startColumn + rows[0].length + 1);
  sheet.getRange(startRow, startColumn, rows.length + 1, rows[0].length).clearContent();
  sheet.getRange(startRow, startColumn).setValue(title);
  sheet.getRange(startRow + 1, startColumn, rows.length, rows[0].length).setValues(rows);
}

function getColumnMax_(matrix, columnIndex) {
  var maxValue = 0;

  for (var i = 1; i < matrix.length; i++) {
    maxValue = Math.max(maxValue, toNumber_(matrix[i][columnIndex]));
  }

  return maxValue;
}

function barText_(value, maxValue) {
  if (maxValue <= 0 || value <= 0) return "";

  return repeatText_("#", Math.max(1, Math.round((value / maxValue) * 20)));
}

function repeatText_(text, count) {
  var output = "";

  for (var i = 0; i < count; i++) {
    output += text;
  }

  return output;
}

function renderGraphSourceTables_(sheet, orderMatrix, itemMatrix, performanceRows, dailyComparisonRows, mtdRows) {
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.ORDER_SOURCE_ROW, 1, "Hourly Orders Source Data", orderMatrix, "#1f4e79");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.ITEM_SOURCE_ROW, 1, "Hourly Items Source Data", itemMatrix, "#548235");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.PERFORMANCE_SOURCE_ROW, 1, "Latest Date Speed / SLA Source Data", performanceRows, "#7030a0");
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.DAILY_KPI_SOURCE_ROW, 1, "Daily KPI Source Data", dailyComparisonRows, "#0b1f3a");
  var topRows = [["Picker", "MTD Score"]].concat(
    mtdRows.slice(0, 5).map(function(row, index) {
      return [row.picker, row.score];
    })
  );
  renderDashboardMatrix_(sheet, GRAPH_LAYOUT.TOP_PICKER_SOURCE_ROW, 1, "Top Picker MTD Score", topRows, "#14532d");
}

function renderDashboardMatrix_(sheet, startRow, startColumn, title, values, color) {
  if (!values || values.length === 0) return;

  ensureSheetSize_(sheet, startRow + values.length + 2, startColumn + values[0].length + 1);
  sheet.getRange(startRow, startColumn, values.length + 1, values[0].length).clearContent();
  sheet.getRange(startRow, startColumn).setValue(title);
  sheet.getRange(startRow + 1, startColumn, values.length, values[0].length).setValues(values);
}

function insertDashboardCharts_(sheet, hourlyRows, mtdRows, orderMatrix, itemMatrix, performanceRows, dailyComparisonRows) {
  if (orderMatrix.length > 1 && orderMatrix[0].length > 1) {
    upsertDashboardChart_(
      sheet,
      "Hourly Order Flow by Date",
      Charts.ChartType.COLUMN,
      sheet.getRange(GRAPH_LAYOUT.ORDER_SOURCE_ROW + 1, 1, orderMatrix.length, Math.min(orderMatrix[0].length, 3)),
      GRAPH_LAYOUT.ORDER_CHART_ROW,
      GRAPH_LAYOUT.CHART_COLUMN,
      {
        hAxisTitle: "Hour",
        vAxisTitle: "Orders",
        colors: ["#64748b", "#2563eb"]
      }
    );
  }

  if (itemMatrix.length > 1 && itemMatrix[0].length > 1) {
    upsertDashboardChart_(
      sheet,
      "Hourly Item Flow by Date",
      Charts.ChartType.COLUMN,
      sheet.getRange(GRAPH_LAYOUT.ITEM_SOURCE_ROW + 1, 1, itemMatrix.length, Math.min(itemMatrix[0].length, 3)),
      GRAPH_LAYOUT.ITEM_CHART_ROW,
      GRAPH_LAYOUT.CHART_COLUMN,
      {
        hAxisTitle: "Hour",
        vAxisTitle: "Items",
        colors: ["#64748b", "#16a34a"]
      }
    );
  }

  if (performanceRows.length > 1) {
    upsertDashboardChart_(
      sheet,
      "Latest Date AT / PT / PKT / P95 by Hour",
      Charts.ChartType.COLUMN,
      sheet.getRange(GRAPH_LAYOUT.PERFORMANCE_SOURCE_ROW + 1, 1, performanceRows.length, performanceRows[0].length),
      GRAPH_LAYOUT.PERFORMANCE_CHART_ROW,
      GRAPH_LAYOUT.CHART_COLUMN,
      {
        hAxisTitle: "Hour",
        vAxisTitle: "Time",
        colors: ["#2563eb", "#ca8a04", "#ea580c", "#dc2626"]
      }
    );
  }

  if (dailyComparisonRows.length > 1) {
    upsertDashboardChart_(
      sheet,
      "Latest Day vs Same Day Last Week - KPI Comparison",
      Charts.ChartType.COLUMN,
      sheet.getRange(GRAPH_LAYOUT.DAILY_KPI_SOURCE_ROW + 1, 1, dailyComparisonRows.length, 3),
      GRAPH_LAYOUT.DAILY_KPI_CHART_ROW,
      GRAPH_LAYOUT.CHART_COLUMN,
      {
        hAxisTitle: "KPI",
        vAxisTitle: "Value",
        colors: ["#64748b", "#2563eb"]
      }
    );
  }

  if (mtdRows.length > 0) {
    upsertDashboardChart_(
      sheet,
      "Top Picker MTD Score",
      Charts.ChartType.BAR,
      sheet.getRange(GRAPH_LAYOUT.TOP_PICKER_SOURCE_ROW + 1, 1, Math.min(mtdRows.length, 5) + 1, 2),
      GRAPH_LAYOUT.TOP_PICKER_CHART_ROW,
      GRAPH_LAYOUT.CHART_COLUMN,
      {
        hAxisTitle: "MTD Score",
        vAxisTitle: "Picker",
        colors: ["#16a34a"]
      }
    );
  }

  sheet.autoResizeColumns(1, 14);
}

function upsertDashboardChart_(sheet, title, chartType, range, defaultRow, defaultColumn, options) {
  var existingChart = findChartByTitle_(sheet, title);

  if (existingChart) {
    var builder = existingChart.modify().clearRanges().addRange(range);
    var container = existingChart.getContainerInfo();
    builder.setPosition(
      container.getAnchorRow(),
      container.getAnchorColumn(),
      container.getOffsetX(),
      container.getOffsetY()
    );
    sheet.updateChart(builder.build());
  } else {
    var newBuilder = sheet.newChart()
      .setChartType(chartType)
      .addRange(range)
      .setOption("title", title)
      .setOption("legend", { position: "right" })
      .setOption("hAxis", { title: options.hAxisTitle })
      .setOption("vAxis", { title: options.vAxisTitle })
      .setOption("colors", options.colors || ["#2563eb", "#16a34a"])
      .setPosition(defaultRow, defaultColumn, 0, 0);

    sheet.insertChart(newBuilder.build());
  }
}

function findChartByTitle_(sheet, title) {
  var charts = sheet.getCharts();

  for (var i = 0; i < charts.length; i++) {
    if (getChartTitle_(charts[i]) === title) {
      return charts[i];
    }
  }

  return null;
}

function getChartTitle_(chart) {
  try {
    return chart.getOptions().get("title") || "";
  } catch (error) {
    return "";
  }
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
    .setBackground("#ffffff")
    .setFontColor("#000000")
    .setFontWeight("normal")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setFontFamily("Arial")
    .setBorder(false, false, false, false, false, false);

  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  sheet.setRowHeights(1, rowCount, 28);

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
    .setBorder(true, true, true, true, true, true, "#0b1f3a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  if (isWeekHeader) {
    range
      .setBackground("#0b1f3a")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(12);
    sheet.setRowHeight(startRow, 34);
    return;
  }

  sheet.getRange(startRow, 1, 1, maxColumns)
    .setBackground("#0b1f3a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11);

  if (rowCount > 1) {
    sheet.getRange(startRow + 1, 1, 1, maxColumns)
      .setBackground("#1f4e79")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
  }

  if (rowCount > 2) {
    sheet.getRange(startRow + 2, 1, rowCount - 2, maxColumns)
      .setBackground("#edf2f7")
      .setFontColor("#0f172a")
      .setFontWeight("normal")
      .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
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
