/**
 * Repo_FreshnessAuditor.js
 * Sovereign Asset Protocol - Data Freshness Collector & Diagnostics
 */

function buildDataFreshnessReport_(ss, now) {
  const at = now || new Date();
  let s = "\n[V] 資料可信度 (DATA FRESHNESS)\n";

  s += formatDataFreshnessLine_("價格資料", function () {
    return collectPriceFreshness_(ss, at);
  });
  s += formatDataFreshnessLine_("匯率資料", function () {
    return collectFxFreshness_(ss, at);
  });
  s += formatDataFreshnessLine_("交易所同步", function () {
    return collectSyncFreshness_(ss, at);
  });

  return s;
}

function formatDataFreshnessLine_(label, collectFn) {
  try {
    const summary = collectFn();
    const statusIcon = summary.ok ? "✅" : "⚠️";
    let line = "- " + label + ": " + statusIcon + " " + summary.statusText;
    if (summary.latestAt) line += " | 最新 " + formatFreshnessTimestamp_(summary.latestAt);
    if (summary.oldestAt) line += " | 最舊 " + formatFreshnessTimestamp_(summary.oldestAt);
    if (summary.countLabel) line += " | " + summary.countLabel;
    if (summary.note) line += " | " + summary.note;
    line += "\n";
    if (summary.detail) line += "  > " + summary.detail + "\n";
    return line;
  } catch (e) {
    return "- " + label + ": ⚠️ 無法讀取資料可信度 (" + sanitizeFreshnessMessage_(e.message || e) + ")\n";
  }
}

function collectPriceFreshness_(ss, now) {
  const rows = PriceCacheRepo.readRows(ss);
  const staleHours = (Config.THRESHOLDS && Config.THRESHOLDS.PRICE_HEALTH_STALE_HOURS) || 24;
  const stats = buildFreshnessTimestampStats_(rows, function (row) { return row.updatedAt; }, now);
  const blankPriceCount = rows.filter(function (row) { return !isPositiveFreshnessNumber_(row.price); }).length;
  const lockedRows = rows.filter(function (row) {
    const updatedAt = parseFreshnessDate_(row.updatedAt);
    return updatedAt && updatedAt.getTime() > now.getTime() + 1000;
  });
  const stale = isFreshnessStale_(stats.oldestAt, now, staleHours);
  const ok = rows.length > 0 && blankPriceCount === 0 && stats.blankCount === 0 && !stale;
  const notes = [];

  if (blankPriceCount > 0) notes.push("空白價格 " + blankPriceCount);
  if (stats.blankCount > 0) notes.push("缺 timestamp " + stats.blankCount);
  if (lockedRows.length > 0) notes.push("鎖價 " + lockedRows.length);

  return {
    ok: ok,
    statusText: rows.length === 0 ? "沒有價格列" : (ok ? "新鮮" : "需要注意"),
    latestAt: stats.latestAt,
    oldestAt: stats.oldestAt,
    countLabel: rows.length + " rows",
    note: notes.join(", ")
  };
}

function collectFxFreshness_(ss, now) {
  const pairResult = SettingsMatrixRepo.listPairs(ss);
  const pairs = pairResult.pairs || [];
  const staleHours = (Config.THRESHOLDS && Config.THRESHOLDS.FX_RATE_STALE_HOURS) || 24;
  const stats = buildFreshnessTimestampStats_(pairs, function (pair) { return pair.currentTimestamp; }, now);
  const invalidCount = pairs.filter(function (pair) { return !isPositiveFreshnessNumber_(pair.currentValue); }).length;
  const stale = isFreshnessStale_(stats.oldestAt, now, staleHours);
  const ok = pairs.length > 0 && invalidCount === 0 && stats.blankCount === 0 && !stale;
  const notes = [];

  if (invalidCount > 0) notes.push("無效匯率 " + invalidCount);
  if (stats.blankCount > 0) notes.push("缺 timestamp " + stats.blankCount);

  return {
    ok: ok,
    statusText: pairs.length === 0 ? "沒有匯率 pair" : (ok ? "新鮮" : "需要注意"),
    latestAt: stats.latestAt,
    oldestAt: stats.oldestAt,
    countLabel: pairs.length + " pairs",
    note: notes.join(", ")
  };
}

function collectSyncFreshness_(ss, now) {
  const statuses = SyncStatusRepo.readAll(ss);
  const staleHours = (Config.THRESHOLDS && Config.THRESHOLDS.SYNC_STATUS_STALE_HOURS) || 24;
  const stats = buildFreshnessTimestampStats_(statuses, function (status) { return status.lastSuccess; }, now);
  const nonComplete = statuses.filter(function (status) { return status.status !== "COMPLETE"; });
  const staleStatuses = statuses.filter(function (status) {
    return !status.lastSuccess || isFreshnessStale_(status.lastSuccess, now, staleHours);
  });
  const ok = statuses.length > 0 && nonComplete.length === 0 && staleStatuses.length === 0;
  const notes = [];
  const detail = buildSyncFreshnessDetail_(statuses, nonComplete);

  if (nonComplete.length > 0) notes.push("非 COMPLETE " + nonComplete.length);
  if (staleStatuses.length > 0) notes.push("可能過期 " + staleStatuses.length);

  return {
    ok: ok,
    statusText: statuses.length === 0 ? "沒有同步狀態" : (ok ? "全部成功" : "需要注意"),
    latestAt: stats.latestAt,
    oldestAt: stats.oldestAt,
    countLabel: statuses.length + " exchanges",
    note: notes.join(", "),
    detail: detail
  };
}

function buildSyncFreshnessDetail_(statuses, nonComplete) {
  if (nonComplete.length > 0) {
    return "非 COMPLETE: " + nonComplete.slice(0, 6).map(function (status) {
      return status.exchange + " " + status.status;
    }).join(", ") + (nonComplete.length > 6 ? ", ..." : "");
  }

  return statuses.slice(0, 6).map(function (status) {
    const timeText = status.lastSuccess ? formatFreshnessTimeOfDay_(status.lastSuccess) : "no-success";
    return status.exchange + " " + timeText + " " + status.status;
  }).join(", ") + (statuses.length > 6 ? ", ..." : "");
}

function buildFreshnessTimestampStats_(rows, getTimestamp, now) {
  const dates = [];
  let blankCount = 0;
  (rows || []).forEach(function (row) {
    const parsed = parseFreshnessDate_(getTimestamp(row));
    if (!parsed) {
      blankCount++;
      return;
    }
    if (parsed.getTime() > now.getTime() + 1000) return;
    dates.push(parsed);
  });

  dates.sort(function (a, b) { return a.getTime() - b.getTime(); });

  return {
    blankCount: blankCount,
    oldestAt: dates.length > 0 ? dates[0] : null,
    latestAt: dates.length > 0 ? dates[dates.length - 1] : null
  };
}

function isFreshnessStale_(date, now, staleHours) {
  if (!date) return true;
  return now.getTime() - date.getTime() > staleHours * 60 * 60 * 1000;
}

function isPositiveFreshnessNumber_(value) {
  const num = parseFloat(value);
  return isFinite(num) && num > 0;
}

function parseFreshnessDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatFreshnessTimestamp_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  } catch (e) {
    return date && date.toISOString ? date.toISOString() : String(date || "");
  }
}

function formatFreshnessTimeOfDay_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
  } catch (e) {
    return date && date.toISOString ? date.toISOString().substring(11, 16) : String(date || "");
  }
}

function sanitizeFreshnessMessage_(message) {
  return String(message || "").replace(/https?:\/\/\S+/gi, "[url]").substring(0, 160);
}

/**
 * Unified Broadcast Handler (Email + Discord)
 * @private
 */