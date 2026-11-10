/* Робот LP-CRM. Работает внутри страницы CRM под вашим логином: делает то же, что оператор руками
   (открыть карточку - изменить поле - «Зберегти»), и создаёт заказы через API CRM.
   Никуда, кроме вашей CRM, данные не отправляет. */
(function () {
  'use strict';
  if (window.__lpRobot) return;
  window.__lpRobot = true;

  var CFG = {
    DELIVERY_NP: '1',             // способ доставки «Нова Пошта»
    PAYMENT: '4',                 // вид оплаты «Післяплата»
    COUNTRY: 'UA', OFFICE: '1',
    SITE: 'google.com',           // «сайт» заказа, как в таблице
    COLS: ['Статус', 'ТТН', 'ФИО', 'Телефон', 'Товар', 'Кол-во', 'Сумма', 'Адрес НП', 'ID товара СРМ', 'Call ID', 'Комментарий']
  };
  var JOB_KEY = 'lprobot_job', SHEET_KEY = 'lprobot_sheet', APIKEY_KEY = 'lprobot_apikey', STATUS_KEY = 'lprobot_importstatus';
  function ls(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } }
  // «Вхідний ключ API» вашей LP-CRM (создание заказов) и id статуса «Импорт» вводятся на вкладке «Настройки» и хранятся в браузере
  Object.defineProperty(CFG, 'API_KEY', { get: function () { return ls(APIKEY_KEY); } });
  Object.defineProperty(CFG, 'IMPORT_STATUS', { get: function () { return ls(STATUS_KEY) || '31'; } });
  var stopFlag = false, busy = false;

  // ---------------------------------------------------------------- мелочи
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function waitFor(fn, ms, step) {
    var t0 = Date.now();
    return new Promise(function (resolve) {
      (function tick() {
        var v; try { v = fn(); } catch (e) { v = null; }
        if (v) return resolve(v);
        if (Date.now() - t0 > ms) return resolve(null);
        setTimeout(tick, step || 200);
      })();
    });
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function digits(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }
  function now() { var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; }; return p(d.getDate()) + '.' + p(d.getMonth() + 1) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()); }
  function onOrders() { return /^\/orders\/?$/.test(location.pathname); }
  function curStatus() { var m = location.search.match(/[?&]status=([^&]+)/); return m ? m[1] : ''; }
  function jq() { return window.jQuery || window.$; }
  function rows() { return Array.prototype.slice.call(document.querySelectorAll('tr.order-row[id]')); }
  function saveJob(j) { try { localStorage.setItem(JOB_KEY, JSON.stringify(j)); } catch (e) {} }
  function loadJob() { try { var j = JSON.parse(localStorage.getItem(JOB_KEY) || 'null'); return j && Date.now() - j.t < 120000 ? j : null; } catch (e) { return null; } }
  function clearJob() { try { localStorage.removeItem(JOB_KEY); } catch (e) {} }

  // ---------------------------------------------------------------- таблица
  var pending = {}, seq = 0;
  window.addEventListener('message', function (ev) {
    if (ev.source !== window || !ev.data || ev.data.__lprobot !== 'res') return;
    var cb = pending[ev.data.id]; if (cb) { delete pending[ev.data.id]; cb(ev.data.res || { ok: false, error: 'пустой ответ' }); }
  });
  function sheetUrl() { try { return localStorage.getItem(SHEET_KEY) || ''; } catch (e) { return ''; } }
  function sheet(body) {
    return new Promise(function (resolve) {
      var url = sheetUrl(), id = ++seq;
      if (!url) return resolve({ ok: false, error: 'не указана ссылка таблицы (вкладка «Таблица»)' });
      if (body) body.t = (url.match(/[?&]t=(\w+)/) || [])[1] || '';
      pending[id] = resolve;
      window.postMessage({ __lprobot: 'req', id: id, url: url, body: body || null }, location.origin);
      setTimeout(function () { if (pending[id]) { delete pending[id]; resolve({ ok: false, error: 'таблица не ответила за 40 секунд' }); } }, 40000);
    });
  }

  // ---------------------------------------------------------------- карточка заказа
  function cardEl() { return document.querySelector('#modal-window-order'); }
  function cardReady(id) {
    var c = cardEl(); if (!c) return false;
    var f = c.querySelector('#form-order-id'), t = c.querySelector('#ttn'), b = document.querySelector('#button-save-modal');
    if (!f || !t || !b) return false;
    if (String(f.value) !== String(id)) return false;
    if (window.ORDER_PRODUCTS_DATA_OWNER === '__loading__') return false;
    return c;
  }
  function closeCard() {
    var c = cardEl(); if (!c) return;
    var w = c.closest('.modal-window'); var x = (w || document).querySelector('.modal-window-close');
    if (x) x.click();
  }
  async function openCard(info) {   // info = [id, код, дата, изменён]
    if (cardEl()) { closeCard(); await waitFor(function () { return !cardEl(); }, 5000); }
    if (window.BLOCKED_ORDERS_ID && window.BLOCKED_ORDERS_ID[info[0]]) return 'заказ сейчас открыт другим пользователем';
    if (typeof window.Orders !== 'function') return 'страница заказов не загрузилась';
    window.Orders('edit', info[0], info[1] || '', info[2] || '', info[3] || '');
    var ok = await waitFor(function () { return cardReady(info[0]); }, 20000);
    if (!ok) { closeCard(); return 'карточка не открылась'; }
    await sleep(200);
    return '';
  }
  function setField(el, val) {
    el.value = val;
    ['input', 'keyup', 'change'].forEach(function (t) { el.dispatchEvent(new Event(t, { bubbles: true })); });
  }
  function setSelect(sel, val) {
    if (!sel) return false;
    var has = Array.prototype.some.call(sel.options, function (o) { return o.value === String(val); });
    if (!has) return false;
    sel.value = String(val);
    var $ = jq();
    if ($) { $(sel).trigger('change').trigger('chosen:updated'); } else sel.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }
  /** Нажимает «Зберегти» и ждёт ответ сервера. Возвращает '' или текст ошибки. */
  function saveCard() {
    return new Promise(function (resolve) {
      var $ = jq(), done = false, timer;
      function finish(msg) { if (done) return; done = true; clearTimeout(timer); if ($) $(document).off('ajaxComplete.lprobot'); resolve(msg); }
      if (!$) return resolve('нет jQuery на странице');
      $(document).on('ajaxComplete.lprobot', function (e, xhr, st) {
        if (!st || String(st.url).indexOf('/ajax_orders') < 0) return;
        if (String(st.data || '').indexOf('action=update') < 0 && String(st.data || '').indexOf('action=add') < 0) return;
        var txt = xhr && xhr.responseText != null ? String(xhr.responseText) : '';
        var j = null; try { j = JSON.parse(txt); } catch (err) {}
        if (xhr.status === 200 && (txt === '' || (j && j.success === true))) finish('');
        else finish('CRM не сохранила: ' + (j && (j.message || j.error) ? (j.message || j.error) : (txt || ('код ' + xhr.status))).toString().slice(0, 150));
      });
      timer = setTimeout(function () {
        var note = document.querySelector('.notification-error, .lp-notification, .modal-window .error');
        finish('нет ответа от CRM при сохранении' + (note ? ': ' + note.innerText.trim().slice(0, 120) : '') + ' (вкладка должна быть активной)');
      }, 25000);
      var b = document.querySelector('#button-save-modal');
      if (!b) return finish('нет кнопки «Зберегти»');
      b.click();
    });
  }
  /** Открыть заказ, изменить поля (fn(card) -> '' или ошибка), сохранить. */
  async function editOrder(info, fn) {
    var e = await openCard(info); if (e) return e;
    var err = fn(cardEl());
    if (err) { closeCard(); await waitFor(function () { return !cardEl(); }, 5000); return err; }
    await sleep(150);
    var res = await saveCard();
    await waitFor(function () { return !cardEl(); }, 8000);
    if (cardEl()) closeCard();
    await sleep(250);   // список перечитывается
    return res;
  }

  // ---------------------------------------------------------------- 1. стереть ТТН в статусе
  /** Собирает заказы с ТТН на странице (с прокруткой, если список длинный). */
  async function collectWithTtn() {
    var found = {}, order = [], idle = 0, lastCount = -1;
    for (var step = 0; step < 200 && idle < 3; step++) {
      var rs = rows();
      rs.forEach(function (r) {
        if (found[r.id]) return;
        found[r.id] = true;
        if (digits(r.dataset.ttn)) order.push({ info: (r.dataset.info || (r.id + '|' + (r.dataset.orderid || ''))).split('|'), ttn: digits(r.dataset.ttn), name: r.dataset.buyer || '' });
      });
      var n = Object.keys(found).length;
      if (n === lastCount) idle++; else idle = 0;
      lastCount = n;
      if (rs.length) rs[rs.length - 1].scrollIntoView({ block: 'end' });
      window.scrollTo(0, document.documentElement.scrollHeight);
      await sleep(450);
    }
    window.scrollTo(0, 0);
    return order;
  }
  async function runClear(list, limit) {
    var ok = 0, errs = [];
    var todo = limit > 0 ? list.slice(0, limit) : list;
    for (var i = 0; i < todo.length; i++) {
      if (stopFlag) break;
      var z = todo[i];
      progress('Стираю ТТН', i + 1, todo.length, 'заказ ' + z.info[0]);
      var e = await editOrder(z.info, function (c) {
        var t = c.querySelector('#ttn'), ref = c.querySelector('#ttn_ref');
        if (!t) return 'нет поля ТТН';
        setField(t, '');
        if (ref) ref.value = '';
        return '';
      });
      if (e) errs.push('заказ ' + z.info[0] + ' (ТТН ' + z.ttn + '): ' + e); else ok++;
    }
    return { ok: ok, errs: errs, total: todo.length, stopped: stopFlag };
  }

  // ---------------------------------------------------------------- 2. создать заказы
  function utf8len(s) { return unescape(encodeURIComponent(s)).length; }
  function phpSer(v) {
    if (Array.isArray(v)) return 'a:' + v.length + ':{' + v.map(function (x, i) { return 'i:' + i + ';' + phpSer(x); }).join('') + '}';
    if (v !== null && typeof v === 'object') { var k = Object.keys(v); return 'a:' + k.length + ':{' + k.map(function (x) { return phpSer(x) + phpSer(v[x]); }).join('') + '}'; }
    var s = String(v == null ? '' : v);
    return 's:' + utf8len(s) + ':"' + s + '";';
  }
  /** Строки листа «Заказы на СРМ» (из таблицы) -> заказы с проверкой. */
  function prepOrders(rows) {
    return (rows || []).map(function (r) {
      var g = function (v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); };
      var o = { row: r.row, ttn: digits(r.ttn), name: g(r.name), phone: digits(r.phone), tovar: g(r.tovar),
                sum: parseFloat(g(r.sum).replace(/\s/g, '').replace(',', '.')) || 0, addr: g(r.addr),
                pid: digits(r.pid), code: digits(r.code), comment: g(r.comment), err: '' };
      if (o.phone.length === 10 && o.phone[0] === '0') o.phone = '38' + o.phone;
      if (o.ttn.length !== 14) o.err = 'ТТН должна быть из 14 цифр';
      else if (!o.pid) o.err = 'нет ID товара СРМ';
      else if (!o.code) o.err = 'нет Call ID';
      else if (!o.phone) o.err = 'нет телефона';
      else if (!o.name) o.err = 'нет ФИО';
      return o;
    });
  }
  async function apiCreate(o) {
    if (!CFG.API_KEY) return { id: '', err: 'не указан ключ API (вкладка «Настройки»)' };
    var body = new URLSearchParams();
    var data = {
      key: CFG.API_KEY, order_id: o.code, country: CFG.COUNTRY, office: CFG.OFFICE,
      products: encodeURIComponent(phpSer([{ product_id: o.pid, price: o.sum, count: 1 }])),
      bayer_name: o.name, phone: o.phone, email: '', comment: o.comment || '',
      delivery: CFG.DELIVERY_NP, delivery_adress: o.addr, payment: CFG.PAYMENT,
      sender: encodeURIComponent(phpSer({ SERVER_NAME: CFG.SITE, HTTP_HOST: CFG.SITE, HTTP_REFERER: 'https://' + CFG.SITE + '/', REQUEST_URI: '/' })),
      utm_source: ''
    };
    Object.keys(data).forEach(function (k) { body.append(k, data[k]); });
    try {
      var r = await fetch('/api/addNewOrder.html', { method: 'POST', body: body, credentials: 'omit' });
      var t = await r.text(), j = null;
      try { j = JSON.parse(t); } catch (e) {}
      if (j && j.status === 'ok') {
        var d0 = Array.isArray(j.data) ? j.data[0] : j.data;   // CRM возвращает внутренний ID нового заказа - поиск по списку не нужен
        return { id: d0 && d0.id ? String(d0.id) : '', err: '' };
      }
      return { id: '', err: (j && j.message ? (typeof j.message === 'string' ? j.message : JSON.stringify(j.message)) : t || ('код ' + r.status)).slice(0, 160) };
    } catch (e) { return { id: '', err: 'нет связи с CRM: ' + e.message }; }
  }
  function searchInput() {
    var f = Array.prototype.find.call(document.forms, function (x) { return x.querySelector('[name="order_id"]'); });
    return f ? f.querySelector('[name="order_id"]') : null;
  }
  /** Ищет заказ по коду (order_id) через поиск списка. Возвращает строку таблицы или null. */
  async function findByCode(code) {
    var inp = searchInput();
    if (!inp || typeof window.AJAX_SEARCH !== 'function') return null;
    for (var attempt = 0; attempt < 3; attempt++) {
      inp.value = code;
      window.AJAX_SEARCH();
      var r = await waitFor(function () {
        var m = rows().filter(function (x) { return digits(x.dataset.orderid) === code; });
        return m.length === 1 && rows().length <= 3 ? m[0] : null;
      }, 7000, 300);
      if (r) return r;
      await sleep(1500);
    }
    return null;
  }
  async function resetSearch() {
    var inp = searchInput();
    if (inp && inp.value) { inp.value = ''; if (typeof window.AJAX_SEARCH === 'function') window.AJAX_SEARCH(); await sleep(1500); }
  }
  // статусы, которые не удалось записать в таблицу, хранятся в браузере и отправляются при следующем запуске (иначе заказ мог бы создаться повторно)
  function keepPending(batch) {
    try { var a = JSON.parse(localStorage.getItem('lprobot_pending') || '[]'); localStorage.setItem('lprobot_pending', JSON.stringify(a.concat(batch).slice(-500))); } catch (e) {}
  }
  async function sendPending() {
    var a = []; try { a = JSON.parse(localStorage.getItem('lprobot_pending') || '[]'); } catch (e) {}
    if (!a.length) return;
    var w = await sheet({ updates: a });
    if (w.ok) { try { localStorage.removeItem('lprobot_pending'); } catch (e) {} }
  }
  function stamp() { var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds()); }
  /** Вписать ТТН и поставить статус «Импорт» в заказе (открывается сразу по ID, без поиска). */
  function fillTtn(info, o) {
    return editOrder(info, function (c) {
      var t = c.querySelector('#ttn'); if (!t) return 'нет поля ТТН';
      if (!setSelect(c.querySelector('select[name="status"]'), CFG.IMPORT_STATUS)) return 'нет статуса «Импорт» в списке';
      setSelect(c.querySelector('select[name="payment"]'), CFG.PAYMENT);
      var d = c.querySelector('select[name="delivery"]'); if (d && !d.value) setSelect(d, CFG.DELIVERY_NP);
      var a = c.querySelector('[name="delivery_adress"]'); if (a && !a.value && o.addr) setField(a, o.addr);
      setField(t, o.ttn);
      return '';
    });
  }
  async function runCreate(list, limit) {
    var ok = 0, bad = 0, done = 0, lines = [], queue = [], writes = [];
    var good = list.filter(function (o) { return !o.err; });
    var planned = limit > 0 ? Math.min(good.length, limit) : good.length;
    // статусы в таблицу пишутся фоном, пачками: ответ Google Apps Script идёт 1-3 секунды, ждать его после каждого заказа незачем
    function flush(force) {
      if (!queue.length || (!force && queue.length < 4)) return;
      var batch = queue.splice(0, queue.length);
      writes.push(sheet({ updates: batch }).then(function (w) {
        if (!w.ok) { keepPending(batch); lines.push({ name: 'таблица', status: 'статусы не записаны (' + (w.error || '') + '), заказы в CRM созданы; запишу при следующем запуске', ok: false }); }
      }));
    }
    function mark(o, status, isOk) {
      if (isOk) ok++; else bad++;
      lines.push({ name: o.name, status: status, ok: isOk });
      queue.push({ row: o.row, code: o.code, status: status });
      flush(false);
    }
    for (var i = 0; i < list.length; i++) {
      var o = list[i];
      if (o.err) { mark(o, '⚠ ' + o.err, false); continue; }
      if (stopFlag || (limit > 0 && done >= limit)) continue;
      done++;
      progress('Создаю заказы', done, planned, o.name);
      var created = false, id = '', info = null;
      var r1 = await apiCreate(o);
      if (!r1.err && r1.id) { created = true; id = r1.id; info = [id, o.code, stamp(), stamp()]; }
      else {
        // заказ с таким Call ID уже есть (повторный запуск) или CRM не вернула ID: ищем его по списку
        var row = await findByCode(o.code);
        if (!row) { mark(o, r1.err ? '⚠ CRM не создала заказ: ' + r1.err : '⚠ заказ создан, но не найден - ТТН впишите вручную', false); continue; }
        if (digits(row.dataset.ttn) === o.ttn) { mark(o, 'OK СРМ ' + now() + ' (уже был в CRM, ID ' + row.id + ')', true); continue; }
        if (digits(row.dataset.ttn)) { mark(o, '⚠ заказ с таким Call ID уже есть в CRM с другой ТТН ' + digits(row.dataset.ttn), false); continue; }
        id = row.id; info = (row.dataset.info || (row.id + '|' + o.code)).split('|');
      }
      var e2 = await fillTtn(info, o);
      if (e2) mark(o, '⚠ заказ ' + (created ? 'создан' : 'найден') + ' (ID ' + id + '), но ТТН не вписана: ' + e2, false);
      else mark(o, 'OK СРМ ' + now() + ' (ID ' + id + ')', true);
    }
    flush(true);
    await Promise.all(writes);
    await resetSearch();
    return { ok: ok, bad: bad, lines: lines, stopped: stopFlag, left: good.length - done };
  }

  // ---------------------------------------------------------------- панель
  var panel, progEl, bodyEl, outEl;
  // пиксельный робот 12x12 (каждая клетка - один «пиксель»)
  var BOT = ['....1..1....','....1..1....','..11111111..','.1111111111.','.1..11..11.1','.1..11..11.1','.1111111111.','.1.11..11.1.','.1111111111.','..11111111..','..1.1..1.1..','..11.11.11..'];
  var ICON = (function () {
    var r = '';
    BOT.forEach(function (row, y) { row.split('').forEach(function (c, x) { if (c === '1') r += '<rect x="' + x + '" y="' + y + '" width="1" height="1"/>'; }); });
    return '<svg viewBox="0 0 12 12" width="24" height="24" fill="currentColor" shape-rendering="crispEdges">' + r + '</svg>';
  })();
  var PX = '"Press Start 2P","Courier New",ui-monospace,monospace';
  var css =
    '@import url("https://fonts.googleapis.com/css2?family=Press+Start+2P&subset=cyrillic&display=swap");' +
    '#lpr-btn{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;margin:0 8px;padding:0;border:0;background:#1a1c2c;color:#ffcd75;cursor:pointer;vertical-align:middle;box-shadow:-2px 0 0 #ffcd75,2px 0 0 #ffcd75,0 -2px 0 #ffcd75,0 2px 0 #ffcd75;image-rendering:pixelated}' +
    '#lpr-btn:hover,#lpr-btn.on{background:#5d275d;color:#fff}' +
    '#lpr-btn.float{position:fixed;top:8px;right:210px;z-index:2147483000}' +
    '#lpr,#lpr *{box-sizing:border-box;font-family:' + PX + ';-webkit-font-smoothing:none;font-smooth:never;text-transform:none}' +
    '#lpr{position:fixed;z-index:2147483000;width:440px;max-height:calc(100vh - 70px);overflow:auto;background:#1a1c2c;color:#f4f4f4;font-size:9px;line-height:1.8;display:none;' +
      'box-shadow:0 -4px 0 0 #f4f4f4,0 4px 0 0 #f4f4f4,-4px 0 0 0 #f4f4f4,4px 0 0 0 #f4f4f4,0 -8px 0 -2px #333c57,0 8px 0 -2px #333c57,-8px 0 0 -2px #333c57,8px 0 0 -2px #333c57,12px 12px 0 0 rgba(0,0,0,.45);image-rendering:pixelated}' +
    '#lpr::-webkit-scrollbar{width:10px}#lpr::-webkit-scrollbar-track{background:#333c57}#lpr::-webkit-scrollbar-thumb{background:#ffcd75}' +
    '#lpr .hd{display:flex;align-items:center;gap:12px;padding:16px 18px 12px;font-size:12px;color:#ffcd75;letter-spacing:1px}#lpr .hd svg{color:#73eff7;flex:none}' +
    '#lpr .hd .x{margin-left:auto;border:0;background:#b13e53;color:#fff;width:24px;height:24px;font-size:12px;line-height:24px;padding:0;cursor:pointer;box-shadow:0 -2px 0 #ef7d57,0 2px 0 #5d275d,-2px 0 0 #ef7d57,2px 0 0 #5d275d}#lpr .hd .x:active{transform:translateY(2px)}' +
    '#lpr .tabs{display:flex;gap:8px;margin:0 18px;padding:0}' +
    '#lpr .tabs button{flex:1;border:0;background:#333c57;color:#94b0c2;padding:10px 2px;font-size:8px;cursor:pointer;box-shadow:0 -2px 0 #566c86,0 2px 0 #1a1c2c,-2px 0 0 #566c86,2px 0 0 #566c86;line-height:1.4}' +
    '#lpr .tabs button:hover{color:#fff}#lpr .tabs button.on{background:#ffcd75;color:#1a1c2c;box-shadow:0 -2px 0 #fff,0 2px 0 #ef7d57,-2px 0 0 #fff,2px 0 0 #fff}' +
    '#lpr .pane{padding:16px 18px 18px;display:none}#lpr .pane.on{display:block}' +
    '#lpr p{margin:0 0 12px;color:#94b0c2}#lpr label{display:block;font-size:8px;color:#a7f070;margin:0 0 6px}' +
    '#lpr select,#lpr input{width:100%;padding:9px 8px;border:0;background:#0f1020;color:#f4f4f4;font-size:9px;margin:0 0 12px;outline:none;box-shadow:0 -2px 0 #566c86,0 2px 0 #566c86,-2px 0 0 #566c86,2px 0 0 #566c86;border-radius:0}' +
    '#lpr select:focus,#lpr input:focus{box-shadow:0 -2px 0 #73eff7,0 2px 0 #73eff7,-2px 0 0 #73eff7,2px 0 0 #73eff7}' +
    '#lpr option{background:#0f1020;color:#f4f4f4}' +
    '#lpr .row{display:flex;gap:14px;align-items:flex-end}#lpr .row>div{flex:1}#lpr .row>div.sm{flex:0 0 110px}' +
    '#lpr .btn{display:inline-block;border:0;padding:11px 14px;font-size:9px;cursor:pointer;margin:0 10px 6px 0;color:#fff;line-height:1.3;border-radius:0}#lpr .btn:active{transform:translateY(3px)}' +
    '#lpr .b1{background:#3b5dc9;box-shadow:0 -2px 0 #73eff7,0 4px 0 #29366f,-2px 0 0 #73eff7,2px 0 0 #29366f}' +
    '#lpr .b2{background:#566c86;box-shadow:0 -2px 0 #94b0c2,0 4px 0 #333c57,-2px 0 0 #94b0c2,2px 0 0 #333c57}' +
    '#lpr .b3{background:#b13e53;box-shadow:0 -2px 0 #ef7d57,0 4px 0 #5d275d,-2px 0 0 #ef7d57,2px 0 0 #5d275d}' +
    '#lpr .b4{background:#38b764;box-shadow:0 -2px 0 #a7f070,0 4px 0 #257179,-2px 0 0 #a7f070,2px 0 0 #257179;color:#0f1020}' +
    '#lpr .btn:hover{filter:brightness(1.12)}' +
    '#lpr .out{margin-top:14px}' +
    '#lpr .card{background:#333c57;padding:12px;margin:0 0 14px;box-shadow:0 -2px 0 #566c86,0 2px 0 #566c86,-2px 0 0 #566c86,2px 0 0 #566c86}' +
    '#lpr .card.warn{box-shadow:0 -2px 0 #ffcd75,0 2px 0 #ffcd75,-2px 0 0 #ffcd75,2px 0 0 #ffcd75}#lpr .card.err{background:#3d1f33;box-shadow:0 -2px 0 #b13e53,0 2px 0 #b13e53,-2px 0 0 #b13e53,2px 0 0 #b13e53}' +
    '#lpr .card.ok{box-shadow:0 -2px 0 #38b764,0 2px 0 #38b764,-2px 0 0 #38b764,2px 0 0 #38b764}' +
    '#lpr .big{font-size:20px;color:#ffcd75;margin-bottom:6px}#lpr .mut{color:#94b0c2;font-size:8px}#lpr b{color:#fff;font-weight:normal}' +
    '#lpr .list{max-height:180px;overflow:auto;margin-top:8px;font-size:8px;line-height:1.7}#lpr .list div{padding:5px 0;border-top:2px dashed #566c86;word-break:break-word;color:#f4f4f4}' +
    '#lpr .prog{padding:18px;display:none}#lpr .prog .t{color:#ffcd75;font-size:10px;font-weight:normal}' +
    '#lpr .bar{height:18px;background:#0f1020;margin:12px 0 14px;box-shadow:0 -2px 0 #566c86,0 2px 0 #566c86,-2px 0 0 #566c86,2px 0 0 #566c86;padding:3px}' +
    '#lpr .bar i{display:block;height:100%;width:0;background:repeating-linear-gradient(90deg,#38b764 0,#38b764 10px,#0f1020 10px,#0f1020 12px);transition:width .3s steps(8)}';

  function progress(title, n, total, what) {
    if (!progEl) return;
    progEl.style.display = 'block';
    progEl.querySelector('.t').textContent = title + (total ? ': ' + n + ' из ' + total : '');
    progEl.querySelector('.w').textContent = what || '';
    progEl.querySelector('.bar i').style.width = total ? Math.round((n - 1) / total * 100) + '%' : '100%';
  }
  function setBusy(v) { busy = v; progEl.style.display = v ? 'block' : 'none'; bodyEl.style.display = v ? 'none' : 'block'; }
  function out(tab, html) { panel.querySelector('#lpr-out-' + tab).innerHTML = html; }
  function showTab(name) {
    Array.prototype.forEach.call(panel.querySelectorAll('.tabs button'), function (b) { b.classList.toggle('on', b.dataset.tab === name); });
    Array.prototype.forEach.call(panel.querySelectorAll('.pane'), function (p) { p.classList.toggle('on', p.dataset.tab === name); });
  }
  function statusOptions() {
    var a = Array.prototype.slice.call(document.querySelectorAll('a[id^="tab-status-"]'));
    return a.map(function (x) { return { id: x.id.replace('tab-status-', ''), name: (x.childNodes[0] ? x.childNodes[0].textContent : x.textContent).trim() }; })
      .filter(function (s) { return s.id !== 'all' && s.name; });
  }
  function place(btn) {
    var r = btn.getBoundingClientRect();
    panel.style.top = Math.round(r.bottom + 10) + 'px';
    panel.style.right = Math.max(10, Math.round(window.innerWidth - r.right - 8)) + 'px';
  }
  function build() {
    var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    var btn = document.createElement('button'); btn.id = 'lpr-btn'; btn.type = 'button'; btn.title = 'Робот LP-CRM'; btn.innerHTML = ICON;
    var gear = document.querySelector('.icon-orders-settings');
    var holder = gear && gear.closest('.order-interaction-buttons');
    if (holder) {
      var anchor = gear; while (anchor.parentElement && anchor.parentElement !== holder) anchor = anchor.parentElement;
      holder.insertBefore(btn, anchor);
    } else { btn.className = 'float'; document.body.appendChild(btn); }
    panel = document.createElement('div'); panel.id = 'lpr';
    var sel = statusOptions().map(function (s) { return '<option value="' + esc(s.id) + '"' + (s.id === '69' ? ' selected' : '') + '>' + esc(s.name) + '</option>'; }).join('');
    panel.innerHTML =
      '<div class="hd">' + ICON + '<span>Робот LP-CRM</span><button class="x" type="button" title="Закрыть">×</button></div>' +
      '<div class="prog"><b class="t"></b><div class="mut w"></div><div class="bar"><i></i></div><button class="btn b3" id="lpr-stop" type="button">Остановить</button><div class="mut" style="margin-top:8px">Не переключайте вкладку, пока робот работает.</div></div>' +
      '<div class="body"><div class="tabs"><button type="button" data-tab="create" class="on">Создать заказы</button><button type="button" data-tab="clear">Стереть ТТН</button><button type="button" data-tab="sheet">Настройки</button></div>' +
      '<div class="pane on" data-tab="create"><p>Берёт заказы с листа «Заказы на СРМ», создаёт их в статусе «Импорт» с ТТН и оплатой «Післяплата». Результат сам записывает в колонку «Статус».</p>' +
      '<div class="row"><div class="sm"><label>Не больше (0 = все)</label><input type="number" id="lpr-limit-c" value="0" min="0"></div><div><button class="btn b1" id="lpr-load" type="button" style="margin-bottom:10px">Загрузить из таблицы</button></div></div>' +
      '<div class="out" id="lpr-out-create"></div></div>' +
      '<div class="pane" data-tab="clear"><p>Открывает каждый заказ в выбранном статусе, стирает ТТН и сохраняет.</p>' +
      '<label>Статус</label><select id="lpr-status">' + sel + '</select>' +
      '<div class="row"><div class="sm"><label>Не больше (0 = все)</label><input type="number" id="lpr-limit-d" value="0" min="0"></div><div><button class="btn b1" id="lpr-clear" type="button" style="margin-bottom:10px">Найти заказы с ТТН</button></div></div>' +
      '<div class="out" id="lpr-out-clear"></div></div>' +
      '<div class="pane" data-tab="sheet"><p>1) Адрес: в редакторе скрипта «Развернуть» → «Управление развертываниями» → «Копировать» у блока «Веб-приложение».<br>2) Код: в таблице меню «Заявки» → «Ссылка для расширения».</p>' +
      '<label>Адрес веб-приложения (…/exec)</label><input type="text" id="lpr-url" placeholder="https://script.google.com/macros/s/.../exec">' +
      '<label>Код доступа</label><input type="password" id="lpr-tok" placeholder="код из таблицы">' +
      '<label>Входной ключ API CRM (создание заказов)</label><input type="password" id="lpr-key" placeholder="Налаштування → API → Вхідний ключ">' +
      '<label>ID статуса «Импорт» (по умолчанию 31)</label><input type="text" id="lpr-imp" placeholder="31">' +
      '<button class="btn b1" id="lpr-saveurl" type="button">Сохранить и проверить</button>' +
      '<div class="out" id="lpr-out-sheet"></div></div></div>';
    document.body.appendChild(panel);
    bodyEl = panel.querySelector('.body'); progEl = panel.querySelector('.prog');
    panel.querySelector('#lpr-key').value = ls(APIKEY_KEY); panel.querySelector('#lpr-imp').value = ls(STATUS_KEY);
    (function () { var m = sheetUrl().match(/^(.*?)\?t=(\w+)$/); if (m) { panel.querySelector('#lpr-url').value = m[1]; panel.querySelector('#lpr-tok').value = m[2]; } })();
    function toggle(show) {
      var on = show === undefined ? panel.style.display !== 'block' : show;
      if (on) place(btn);
      panel.style.display = on ? 'block' : 'none'; btn.classList.toggle('on', on);
    }
    btn.onclick = function (e) { e.preventDefault(); e.stopPropagation(); toggle(); };
    panel.querySelector('.x').onclick = function () { toggle(false); };
    window.addEventListener('resize', function () { if (panel.style.display === 'block') place(btn); });
    Array.prototype.forEach.call(panel.querySelectorAll('.tabs button'), function (b) { b.onclick = function () { showTab(b.dataset.tab); }; });
    panel.querySelector('#lpr-stop').onclick = function () { stopFlag = true; progEl.querySelector('.w').textContent = 'Останавливаюсь после текущего заказа...'; };
    panel.querySelector('#lpr-clear').onclick = startClear;
    panel.querySelector('#lpr-load').onclick = function () { loadCreate(false); };
    panel.querySelector('#lpr-saveurl').onclick = saveUrl;
    if (!sheetUrl()) showTab('sheet');
    return toggle;
  }
  function lim(id) { return Math.max(0, parseInt(panel.querySelector(id).value, 10) || 0); }

  // --- таблица
  async function saveUrl() {
    var u = panel.querySelector('#lpr-url').value.replace(/\s+/g, '').replace(/\?.*$/, ''), t = panel.querySelector('#lpr-tok').value.replace(/\s+/g, '');
    if (!/^https:\/\/script\.google\.com\/macros\/(u\/\d+\/)?s\/[\w-]+\/exec$/.test(u)) { out('sheet', '<div class="card err">Адрес должен начинаться с https://script.google.com/macros/s/ и заканчиваться на /exec. Скопируйте его кнопкой «Копировать» в окне развёртывания.</div>'); return; }
    if (!/^\w{16,}$/.test(t)) { out('sheet', '<div class="card err">Вставьте код доступа из таблицы (меню «Заявки» → «Ссылка для расширения»).</div>'); return; }
    var key = panel.querySelector('#lpr-key').value.replace(/\s+/g, ''), imp = panel.querySelector('#lpr-imp').value.replace(/\D/g, '');
    try { localStorage.setItem(SHEET_KEY, u + '?t=' + t); localStorage.setItem(APIKEY_KEY, key); localStorage.setItem(STATUS_KEY, imp); } catch (e) {}
    out('sheet', '<div class="card">Проверяю связь...</div>');
    var r = await sheet(null);
    out('sheet', r.ok ? '<div class="card ok">Таблица подключена. Заказов к созданию на листе: <b>' + r.rows.length + '</b>' + (key ? '' : '<br>⚠ Ключ API не указан - создавать заказы нельзя.') + '</div>'
                      : '<div class="card err">Не получилось: ' + esc(r.error || 'ошибка') + '</div>');
  }

  // --- 1
  function startClear() {
    var id = panel.querySelector('#lpr-status').value, limit = lim('#lpr-limit-d');
    if (!onOrders() || curStatus() !== id) { saveJob({ t: Date.now(), type: 'clear', status: id, limit: limit }); location.href = '/orders/?status=' + id; return; }
    scanClear();
  }
  async function scanClear() {
    stopFlag = false; setBusy(true); progress('Ищу заказы с ТТН', 0, 0, '');
    await waitFor(function () { return typeof window.Orders === 'function' && document.querySelector('.table-data'); }, 15000);
    await sleep(2500);
    var list = await collectWithTtn();
    setBusy(false); showTab('clear');
    var limit = lim('#lpr-limit-d'), n = limit > 0 ? Math.min(limit, list.length) : list.length;
    var name = (statusOptions().filter(function (s) { return s.id === curStatus(); })[0] || {}).name || curStatus();
    if (!list.length) { out('clear', '<div class="card ok">В статусе «' + esc(name) + '» нет заказов с ТТН.</div>'); return; }
    out('clear', '<div class="card warn"><div class="big">' + list.length + '</div>заказов с ТТН в статусе «' + esc(name) + '».<br>ТТН будет стёрта у <b>' + n + '</b>. Отменить это нельзя.</div>' +
      '<button class="btn b3" id="lpr-go" type="button">Стереть ТТН</button><button class="btn b2" id="lpr-no" type="button">Отмена</button>');
    panel.querySelector('#lpr-no').onclick = function () { out('clear', ''); };
    panel.querySelector('#lpr-go').onclick = async function () {
      setBusy(true); out('clear', '');
      var r = await runClear(list, limit);
      setBusy(false);
      out('clear', '<div class="card ' + (r.errs.length ? 'warn' : 'ok') + '"><div class="big">' + r.ok + ' из ' + r.total + '</div>ТТН стёрта' + (r.stopped ? ' (остановлено)' : '') +
        (list.length > r.total ? '<br>Осталось с ТТН: ' + (list.length - r.total) : '') + '</div>' +
        (r.errs.length ? '<div class="card err"><b>Ошибки</b><div class="list">' + r.errs.map(function (x) { return '<div>' + esc(x) + '</div>'; }).join('') + '</div></div>' : ''));
    };
  }

  // --- 2
  async function loadCreate(auto) {
    showTab('create');
    out('create', '<div class="card">Читаю лист «Заказы на СРМ»...</div>');
    await sendPending();
    var r = await sheet(null);
    if (!r.ok) { out('create', '<div class="card err">' + esc(r.error || 'ошибка') + '</div>'); return; }
    var list = prepOrders(r.rows), limit = lim('#lpr-limit-c');
    var good = list.filter(function (x) { return !x.err; }), bad = list.filter(function (x) { return x.err; });
    var n = limit > 0 ? Math.min(limit, good.length) : good.length;
    if (!list.length) { out('create', '<div class="card ok">На листе нет заказов без статуса «OK» - создавать нечего.</div>'); return; }
    if (auto && good.length) { doCreate(list, limit); return; }
    out('create', '<div class="card' + (good.length ? '' : ' warn') + '"><div class="big">' + good.length + '</div>заказов готовы к созданию' + (n !== good.length ? ' (будет создано <b>' + n + '</b>)' : '') + '</div>' +
      (bad.length ? '<div class="card warn"><b>Пропущу, в них ошибки: ' + bad.length + '</b><div class="list">' + bad.map(function (x) { return '<div>строка ' + x.row + ', ' + esc(x.name || x.phone) + ': ' + esc(x.err) + '</div>'; }).join('') + '</div></div>' : '') +
      (good.length ? '<button class="btn b4" id="lpr-go" type="button">Создать ' + n + '</button><button class="btn b2" id="lpr-no" type="button">Отмена</button>' : ''));
    var no = panel.querySelector('#lpr-no'); if (no) no.onclick = function () { out('create', ''); };
    var go = panel.querySelector('#lpr-go');
    if (go) go.onclick = function () {
      if (!onOrders() || curStatus() !== 'all') { saveJob({ t: Date.now(), type: 'create', limit: limit }); location.href = '/orders/?status=all'; return; }
      doCreate(list, limit);
    };
  }
  async function doCreate(list, limit) {
    stopFlag = false; setBusy(true); out('create', ''); progress('Готовлюсь', 0, 0, '');
    await waitFor(function () { return typeof window.AJAX_SEARCH === 'function' && searchInput(); }, 15000);
    await sleep(1500);
    var r = await runCreate(list, limit);
    setBusy(false); showTab('create');
    var errs = r.lines.filter(function (x) { return !x.ok; });
    out('create', '<div class="card ' + (errs.length ? 'warn' : 'ok') + '"><div class="big">' + r.ok + '</div>заказов создано с ТТН' + (r.stopped ? ' (остановлено)' : '') +
      (r.left > 0 ? '<br>Осталось на листе: ' + r.left : '') + '<div class="mut">Статусы записаны в лист «Заказы на СРМ».</div></div>' +
      (errs.length ? '<div class="card err"><b>С ошибками: ' + errs.length + '</b><div class="list">' + errs.map(function (x) { return '<div>' + esc(x.name) + ': ' + esc(x.status) + '</div>'; }).join('') + '</div></div>' : ''));
  }

  // ---------------------------------------------------------------- запуск
  function start() {
    if (!document.body || document.querySelector('#lpr-btn')) return;
    if (!document.querySelector('a[id^="tab-status-"]') && !/^\/orders/.test(location.pathname)) return;
    var toggle = build();
    var job = loadJob();
    if (!job) return;
    clearJob();
    toggle(true);
    if (job.type === 'clear' && onOrders() && curStatus() === job.status) {
      var s = panel.querySelector('#lpr-status'); if (s) s.value = job.status;
      panel.querySelector('#lpr-limit-d').value = job.limit;
      scanClear();
    } else if (job.type === 'create' && onOrders() && curStatus() === 'all') {
      panel.querySelector('#lpr-limit-c').value = job.limit;
      loadCreate(true);
    }
  }
  if (document.readyState === 'complete') setTimeout(start, 800); else window.addEventListener('load', function () { setTimeout(start, 800); });
  window.__lpRobotTest = { prepOrders: prepOrders, phpSer: phpSer, runCreate: runCreate, apiCreate: apiCreate, findByCode: findByCode };
})();
