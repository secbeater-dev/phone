/* Dedicated import worker. Keep identifiers independent of the main UI layer. */
/* global importScripts, PhoneWorkbench */
importScripts(
  "./vendor/xlsx.full.min.js?v=20260911-legacy-ui-v1",
  "./app.js?v=20260911-legacy-ui-v1"
);

self.onmessage = function onImportParserMessage(event) {
  const payload = event.data || {};
  const requestId = payload.requestId;
  const fileName = payload.fileName;
  try {
    const workspace = PhoneWorkbench.parseImportFile(fileName, payload.bytes);
    self.postMessage({ requestId: requestId, ok: true, workspace: workspace });
  } catch (error) {
    self.postMessage({
      requestId: requestId,
      ok: false,
      message: error && error.message ? error.message : "�ѪR����",
    });
  }
};
