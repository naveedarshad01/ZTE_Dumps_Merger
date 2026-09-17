importScripts('vendor/jszip.min.js', 'engine.js');
let session = null;
let lastProgress = 0;
function progress(value) {
  const now = Date.now();
  if (now - lastProgress > 150 || value.fraction === 1) { postMessage({ type: 'progress', value }); lastProgress = now; }
}
self.onmessage = async ({ data }) => {
  const { id, action, payload } = data;
  try {
    let value;
    if (action === 'inspect') { session = await ITBBU.inspect(payload, JSZip, progress); value = ITBBU.publicSummary(session); }
    else {
      if (!session) throw new Error('Inspect your workbooks first.');
      if (action === 'audit') value = ITBBU.summarizeAudit(session, payload, progress);
      else if (action === 'page') value = ITBBU.auditPage(session, payload);
      else if (action === 'merge') value = await ITBBU.merge(session, progress);
      else if (action === 'export') value = await ITBBU.exportAudit(session, payload, progress);
      else throw new Error('Unknown action.');
    }
    if (value instanceof Uint8Array) postMessage({ id, value }, [value.buffer]);
    else postMessage({ id, value });
  } catch (error) { postMessage({ id, error: error.message || String(error) }); }
};
