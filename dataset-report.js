(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PhoneDatasetReport = api;
})(globalThis, function () {
  'use strict';
  function buildVolume({ records, users = [], metadata, summary, volume, label, notes = {} }, app) {
    const legacyMode = metadata.ui_mode === 'legacy';
    const stations = [...new Map(records.flatMap(row => row.stations || []).map(s => [s.station_key, s])).values()];
    const subject = {};
    users.forEach((user, index) => {
      for (const [key, value] of Object.entries(user.subject || {})) subject[users.length > 1 ? `${index + 1}. ${key}` : key] = value;
    });
    const report = app.buildAttachmentReport({ case: { source_file: metadata.source_file, source_files: metadata.source_files, subject }, records, base_stations: stations }, notes, new Date().toISOString(), {
      scope: 'target_volume',
      scope_label: `${label}｜第 ${volume} 卷（統計僅計本卷）｜完整範圍：${legacyMode ? `通聯 ${summary.call_count + summary.data_count} 筆` : `通聯 ${summary.call_count} 筆、網路 ${summary.data_count} 筆`}`,
    });
    report.network = legacyMode ? null : report.calls.filter(row => row.record_kind === 'data');
    report.calls = legacyMode ? report.calls : report.calls.filter(row => row.record_kind !== 'data');
    const addNotes = stats => Object.fromEntries(Object.entries(stats).map(([key, values]) => [key, values.map(row => ({ ...row, note: notes[row.phone] || '' }))]));
    report.stats = { count: addNotes(app.computePhoneStats(report.calls, 'count')), seconds: addNotes(app.computePhoneStats(report.calls, 'seconds')) };
    report.profile.summary['本卷通聯筆數'] = report.calls.length;
    report.profile.summary['通聯筆數'] = report.calls.length;
    if (!legacyMode) {
      report.profile.summary['本卷網路歷程筆數'] = report.network.length;
      delete report.profile.summary['總秒數'];
      report.profile.summary['通話總秒數'] = report.calls.reduce((sum, row) => sum + Number(row.duration_seconds || 0), 0);
      report.profile.summary['連線總秒數'] = report.network.reduce((sum, row) => sum + Number(row.duration_seconds || 0), 0);
    }
    report.profile.summary['完整範圍通聯筆數'] = legacyMode ? summary.call_count + summary.data_count : summary.call_count;
    if(!legacyMode) report.profile.summary['完整範圍網路歷程筆數'] = summary.data_count;
    return report;
  }
  function buildSummary({ summary, metadata, label, volume, users = [], imeis = [], hotspots = [], stats }) {
    const legacyMode = metadata.ui_mode === 'legacy';
    const hourTotal = summary.hours.reduce((sum, item) => sum + item.count, 0);
    const hotspotTotal = summary.call_count + summary.data_count;
    const profileSummary = legacyMode
      ? { '通聯筆數': summary.call_count + summary.data_count, '目標電話': summary.target_count, '對象電話': summary.counterparty_count, '總秒數': summary.call_seconds + summary.data_seconds }
      : { '通聯筆數': summary.call_count, '網路歷程筆數': summary.data_count, '通話秒數': summary.call_seconds, '連線秒數': summary.data_seconds };
    const empty = () => ({ inboundRows: [], outboundRows: [], totalRows: [] });
    return {
      meta: { exported_at: new Date().toISOString(), source_files: metadata.source_files || [], scope: 'target_summary', scope_label: `${label}｜完整範圍統計｜摘要第 ${volume} 卷（排行每卷 500 項；時間明細見明細卷）`, rank_offset: (volume - 1) * 500 },
      calls: [], network: legacyMode ? null : [],
      hours: summary.hours.map(row => ({ ...row, percent: hourTotal ? row.count / hourTotal * 100 : 0 })),
      hotspots: hotspots.map(row => ({ ...row, percent: hotspotTotal ? row.count / hotspotTotal * 100 : 0, times: [] })),
      profile: { summary: { ...profileSummary, '日期異常筆數': summary.invalid_dates, '第一筆時間': summary.first_seen, '最後時間': summary.last_seen }, subject: users.flatMap((user, index) => Object.entries(user.subject || {}).map(([key, value]) => ({ key: `${(volume - 1) * 500 + index + 1}. ${key}`, value }))), imeis },
      stats: { count: { ...empty(), ...stats.count }, seconds: { ...empty(), ...stats.seconds } }
    };
  }
  return { buildVolume, buildSummary };
});
