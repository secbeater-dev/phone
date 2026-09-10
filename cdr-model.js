(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PhoneCdrModel = factory();
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  function raw(value) { return value && typeof value === "object" && "value" in value ? value.value : value; }
  // Expand decimal exponents as text: identifiers must not round through Number.
  function text(value) {
    const item = raw(value);
    let result = item == null ? "" : String(item).trim();
    if (value?.type === "n" && /^[+-]?\d+\.0+$/.test(result)) result = result.replace(/\.0+$/, "");
    const match = result.match(/^([+-]?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/);
    if (!match || Math.abs(Number(match[4])) > 100) return result;
    const digits = match[2] + (match[3] || ""), point = match[2].length + Number(match[4]);
    return match[1] + (point <= 0 ? "0." + "0".repeat(-point) + digits : point >= digits.length ? digits + "0".repeat(point - digits.length) : digits.slice(0, point) + "." + digits.slice(point));
  }
  function canonicalHeader(value) { return text(value).toLowerCase().replace(/[\s_\-\/（）()：:]+/g, ""); }
  function normalizePhone(value) {
    const digits = text(value).split(/[（(]/, 1)[0].replace(/\D/g, "");
    if (digits.length > 32) return '';
    if (digits.startsWith("886") && digits.length >= 11) return "0" + digits.slice(3);
    if (digits.length === 9 && digits.startsWith("9")) return "0" + digits;
    if (digits.length === 8 && /^[2345678]/.test(digits)) return "0" + digits;
    return digits;
  }
  function dictionary(headers, values) {
    const data = Object.create(null);
    headers.forEach((header, index) => { const key = canonicalHeader(header); if (key && data[key] === undefined) data[key] = values[index]; });
    return data;
  }
  function pick(data, ...keys) { for (const key of keys) { const value = data[canonicalHeader(key)]; if (text(value)) return value; } return ""; }
  function numeric(value) { const n = Number(text(value).replace(/,/g, "")); return Number.isFinite(n) ? n : 0; }
  function dateValue(value, context) {
    const original = text(value);
    if (!original) return { value: "", valid: true };
    const number = Number(original);
    const typedNumeric = typeof raw(value) === "number" || (value && value.type === "n");
    const style = value && typeof value === "object" ? context.styles?.[value.style || 0] : null;
    // Date columns also accept unformatted serials emitted by converted exports.
    if (/^[+-]?\d+(?:\.\d+)?$/.test(original) && ((typedNumeric && number >= 0 && number < 2958466) || (number >= 20000 && number <= 100000) || style?.isDate)) {
      const epoch = context.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, number < 60 ? 31 : 30);
      const date = new Date(epoch + Math.round(number * 86400) * 1000);
      if (Number.isFinite(date.getTime())) return { value: date.toISOString().slice(0, 19), valid: true };
    }
    const match = original.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[ T]+)(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$/);
    if (match) {
      const parts = match.slice(1).map(value => Number(value || 0));
      const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], parts[3], parts[4], parts[5]));
      if (date.getUTCFullYear() === parts[0] && date.getUTCMonth() + 1 === parts[1] && date.getUTCDate() === parts[2] && date.getUTCHours() === parts[3] && date.getUTCMinutes() === parts[4] && date.getUTCSeconds() === parts[5]) return { value: date.toISOString().slice(0, 19), valid: true };
    }
    return { value: original, valid: false };
  }
  function isHeader(headers, values) {
    const pairs = headers.map((h, i) => [canonicalHeader(h), canonicalHeader(values[i])]).filter(pair => pair[0]);
    return pairs.length > 0 && pairs.every(pair => pair[0] === pair[1]);
  }
  function addStation(record, role, idValue, addressValue) {
    let cellId = text(idValue), address = text(addressValue);
    if (!cellId && !address) return;
    const virtual = /VOWIFI|WIFI熱點|WI-FI.*通話/i.test((cellId + address).replace(/\s+/g, ""));
    if (virtual) cellId = "VOWIFI";
    const normalized = virtual ? "VOWIFI" : address.replace(/\s+/g, "").replace(/[（(][2345]G[)）]/gi, "");
    const key = cellId + "|" + normalized;
    if (!record.stations.some(station => station.station_key === key)) record.stations.push({ station_key: key, cell_id: cellId, address, normalized_address: normalized, status: virtual ? "not_applicable" : address ? "pending" : "missing_address", is_virtual: virtual });
    if (!record.base_refs.some(ref => ref.role === role && ref.station_key === key)) record.base_refs.push({ role, station_key: key });
  }
  function normalizeRow(kind, headers, values, context = {}) {
    if (!values.some(value => text(value)) || isHeader(headers, values)) return null;
    const d = dictionary(headers, values), get = (...keys) => pick(d, ...keys);
    const sourceQuery = text(get("查詢項目"));
    const sourceTarget = text(kind === "data" ? get("通聯回應-用戶編號/帳號") : get("調閱門號", "目標電話", "調閱號碼"));
    const occurred = dateValue(kind === "data" ? get("開始時間", "啟始時間", "通聯起始時間") : get("始話時間", "始話日期時間"), context);
    const ended = dateValue(get("結束時間", "通聯結束時間"), context);
    const callType = kind === "data" ? "數據" : text(get("通話類別", "CDR類別"));
    const target = normalizePhone(sourceQuery || sourceTarget);
    const calling = normalizePhone(get("主叫號碼", "發話號碼", "callingnumber"));
    const called = normalizePhone(get("受叫號碼", "callednumber"));
    let direction = kind === "data" ? "data" : /受|收|進|^T$|^I$|^1$|^9$/i.test(callType) ? "inbound" : /發|撥|去|^O$|^2$/i.test(callType) ? "outbound" : "other";
    if (kind !== "data" && target && calling && calling !== called) {
      if (calling === target) direction = "outbound";
      else if (called === target || (!called && normalizePhone(get("對象門號", "對象電話", "通話對象")) === calling)) direction = "inbound";
    }
    const ipv4 = text(get("上網IPv4", "external_ipv4", "IP"));
    const ipv6 = text(get("上網IPv6", "external_ipv6"));
    const internal = text(get("使用者內網IP", "internal_ip"));
    const record = {
      record_kind: kind === "data" ? "data" : "call", source_file: context.fileName || "", source_sheet: context.sheetName || "", row_number: context.rowNumber || 0,
      source_query: sourceQuery, source_target: sourceTarget, occurred_at: occurred.value, ended_at: ended.value,
      duration_seconds: numeric(kind === "data" ? get("連線期間", "通聯時間(秒)") : get("通話期間", "通話秒數", "通話時間(秒)")),
      call_type: callType, direction, target_phone: target,
      counterparty_phone: normalizePhone(get("對象門號", "對象電話", "通話對象") || (calling === target ? called : called === target ? calling : "")),
      imei: text(get("IMEI")), imsi: text(get("IMSI", "Q")), external_ip: ipv4 || ipv6 || internal,
      external_ipv4: ipv4, external_ipv6: ipv6, internal_ip: internal,
      upload_bytes: numeric(get("上傳使用量(Byte)", "上行用量", "上傳流量")), download_bytes: numeric(get("下載使用量(Byte)", "下行用量", "下載流量")), total_bytes: numeric(get("全部使用量(Byte)", "全部用量", "總流量")),
      note: text(get("備註", "其他")), base_refs: [], stations: [],
      warning_count: Number(!occurred.valid) + Number(!ended.valid)
    };
    addStation(record, "start", get("開始基地台編號", "起始基地台編號", "基地台 ID"), get("開始基地台", "起始基地台地址", "基地台位址"));
    addStation(record, "end", get("結束基地台編號", "離開基地台編號", "最終基地台 ID"), get("結束基地台", "離開基地台地址", "最終基地台位址"));
    return record;
  }
  function normalizeUser(headers, values, context = {}) {
    if (!values.some(value => text(value)) || isHeader(headers, values)) return null;
    const d = dictionary(headers, values), get = (...keys) => text(pick(d, ...keys));
    const subject = {};
    for (const [name, aliases] of Object.entries({ "用戶名稱": ["用戶名稱"], "申請號碼": ["查詢項目", "用戶回應-用戶編號/帳號"], "身份證字號": ["身份識別碼", "第二身分識別碼"], "生日": ["生日", "出生日期"], "帳寄地址": ["帳寄地址"], "戶籍地址": ["戶籍地址"], "電信業者": ["電信業者"], "區段時間-開始日期時間": ["區段時間-開始日期時間"], "區段時間-結束日期時間": ["區段時間-結束日期時間"] })) {
      const value = get(...aliases); if (value) subject[name] = value;
    }
    return { phone: normalizePhone(get("查詢項目", "用戶回應-用戶編號/帳號")), subject, source_file: context.fileName || "", source_sheet: context.sheetName || "", row_number: context.rowNumber || 0 };
  }
  function kindOf(record) { return record.record_kind === "data" || (!record.record_kind && /數據|上網/.test(record.call_type || "")) ? "data" : "call"; }
  return { normalizePhone, canonicalHeader, normalizeRow, normalizeUser, kindOf };
});
