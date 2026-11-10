// Запросы к таблице (Apps Script) идут отсюда: со страницы CRM их может не пропустить браузер.
chrome.runtime.onMessage.addListener(function (msg, sender, reply) {
  if (!msg || msg.kind !== 'lprobot-sheet') return;
  var url = String(msg.url || '');
  if (!/^https:\/\/script\.google\.com\/macros\/(u\/\d+\/)?s\/[\w-]+\/exec\?t=\w+$/.test(url)) { reply({ ok: false, error: 'неверная ссылка таблицы' }); return; }
  var opts = msg.body ? { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(msg.body), redirect: 'follow', credentials: 'omit', cache: 'no-store' }
                      : { method: 'GET', redirect: 'follow', credentials: 'omit', cache: 'no-store' };
  var status = 0;
  var finalUrl = '';
  fetch(url, opts).then(function (r) { status = r.status; finalUrl = r.url || ''; return r.text(); }).then(function (t) {
    try { reply(JSON.parse(t)); }
    catch (e) {
      var plain = String(t).replace(/<style[\s\S]*?<\/style>|<script[\s\S]*?<\/script>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 220);
      reply({ ok: false, notJson: true, error: 'Google ответил не данными, а страницей (код ' + status + ', адрес: ' + finalUrl.replace(/^https:\/\//, '').slice(0, 70) + '): «' + plain.slice(0, 120) + '»' });
    }
  }).catch(function (e) { reply({ ok: false, error: 'нет связи с таблицей: ' + e.message }); });
  return true;
});
