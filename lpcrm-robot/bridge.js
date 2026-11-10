// Мост между страницей CRM (robot.js) и фоновой частью расширения (background.js).
window.addEventListener('message', function (ev) {
  if (ev.source !== window || !ev.data || ev.data.__lprobot !== 'req') return;
  var id = ev.data.id;
  try {
    chrome.runtime.sendMessage({ kind: 'lprobot-sheet', url: ev.data.url, body: ev.data.body || null }, function (res) {
      var err = chrome.runtime.lastError;
      window.postMessage({ __lprobot: 'res', id: id, res: err ? { ok: false, error: 'расширение перезагружено - обновите страницу CRM' } : res }, location.origin);
    });
  } catch (e) {
    window.postMessage({ __lprobot: 'res', id: id, res: { ok: false, error: 'расширение перезагружено - обновите страницу CRM' } }, location.origin);
  }
});
