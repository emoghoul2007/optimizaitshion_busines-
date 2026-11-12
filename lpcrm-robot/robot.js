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

  /** Страница заказов Aveopt под входом пользователя (запрос делает фоновая часть расширения). */
  function aveoptPage(n) {
    return new Promise(function (resolve) {
      var id = ++seq; pending[id] = resolve;
      window.postMessage({ __lprobot: 'req', id: id, kind: 'aveopt', page: n }, location.origin);
      setTimeout(function () { if (pending[id]) { delete pending[id]; resolve({ ok: false, error: 'Aveopt не ответил за 40 секунд' }); } }, 40000);
    });
  }

  // ---------------------------------------------------------------- 3. ТТН от дроп-платформ
  function phone380(p) { var d = digits(p); if (d.length === 10 && d[0] === '0') d = '38' + d; else if (d.length === 9) d = '380' + d; return d; }
  function dayNum(s) {   // «11.11.2026 12:05» или «2026-11-11 09:12:32» -> 20261111
    var m = String(s || '').match(/(\d{2})\.(\d{2})\.(\d{4})/); if (m) return +(m[3] + m[2] + m[1]);
    m = String(s || '').match(/(\d{4})-(\d{2})-(\d{2})/); return m ? +(m[1] + m[2] + m[3]) : 0;
  }
  /** Заказы со страницы Aveopt: [{ttn, phone, name, day, crmId:'', src}] и самый старый день на странице. */
  function parseAveopt(html) {
    var doc = new DOMParser().parseFromString(html, 'text/html'), list = [], minDay = 0, rowsN = 0;
    var head = Array.prototype.slice.call(doc.querySelectorAll('table tr')).filter(function (tr) { return /Номер ТТН/.test(tr.textContent) && /Клієнт|Клиент/.test(tr.textContent); })[0];
    if (!head) return { ok: false };
    Array.prototype.forEach.call(head.parentNode.parentNode.querySelectorAll('tr'), function (tr) {
      if (tr === head || tr.cells.length < 7) return;
      rowsN++;
      var day = dayNum(tr.cells[1].textContent); if (day && (!minDay || day < minDay)) minDay = day;
      var ttn = (digits(tr.cells[5].textContent).match(/^\d{14}$/) || [''])[0];
      var cl = tr.cells[6].innerHTML.replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      var ph = (cl.match(/\d{10,12}\s*$/) || [''])[0].trim();
      var cancel = /Скасовано|Отменен/i.test(tr.cells[2].textContent);
      if (ttn && ph) list.push({ ttn: ttn, phone: phone380(ph), name: cl.replace(/\d{10,12}\s*$/, '').trim(), day: day, crmId: '', cancel: cancel, src: 'Aveopt №' + tr.cells[0].textContent.trim() + (cancel ? ', скасовано' : '') });
    });
    return { ok: true, list: list, minDay: minDay, rows: rowsN };
  }
  /** Минимальное чтение .xlsx (zip + xml) без сторонних библиотек: массив строк первого листа. */
  async function readXlsx(buf) {
    var dv = new DataView(buf), u8 = new Uint8Array(buf), td = new TextDecoder('utf-8'), i = buf.byteLength - 22;
    while (i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
    if (i < 0) throw new Error('это не файл .xlsx');
    var n = dv.getUint16(i + 10, true), off = dv.getUint32(i + 16, true), files = {};
    for (var k = 0; k < n; k++) {
      var method = dv.getUint16(off + 10, true), csize = dv.getUint32(off + 20, true), nl = dv.getUint16(off + 28, true), el = dv.getUint16(off + 30, true), cl = dv.getUint16(off + 32, true), lo = dv.getUint32(off + 42, true);
      files[td.decode(u8.subarray(off + 46, off + 46 + nl))] = { method: method, csize: csize, lo: lo };
      off += 46 + nl + el + cl;
    }
    async function text(name) {
      var f = files[name]; if (!f) return '';
      var st = f.lo + 30 + dv.getUint16(f.lo + 26, true) + dv.getUint16(f.lo + 28, true), raw = u8.subarray(st, st + f.csize);
      if (f.method === 0) return td.decode(raw);
      var ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return await new Response(ds).text();
    }
    function xml(t) { return new DOMParser().parseFromString(t, 'application/xml'); }
    var shared = [];
    var ss = await text('xl/sharedStrings.xml');
    if (ss) Array.prototype.forEach.call(xml(ss).getElementsByTagName('si'), function (si) {
      shared.push(Array.prototype.map.call(si.getElementsByTagName('t'), function (t) { return t.textContent; }).join(''));
    });
    var sheetName = Object.keys(files).filter(function (f) { return /^xl\/worksheets\/sheet\d*\.xml$/.test(f); }).sort()[0];
    if (!sheetName) throw new Error('в файле нет листа');
    var out = [];
    Array.prototype.forEach.call(xml(await text(sheetName)).getElementsByTagName('row'), function (row) {
      var r = [];
      Array.prototype.forEach.call(row.getElementsByTagName('c'), function (c) {
        var col = 0, ref = (c.getAttribute('r') || '').replace(/\d+/g, '');
        for (var j = 0; j < ref.length; j++) col = col * 26 + ref.charCodeAt(j) - 64;
        var t = c.getAttribute('t'), v = c.getElementsByTagName('v')[0], val = '';
        if (t === 's') val = shared[+(v ? v.textContent : -1)] || '';
        else if (t === 'inlineStr') val = Array.prototype.map.call(c.getElementsByTagName('t'), function (x) { return x.textContent; }).join('');
        else if (v) { val = v.textContent; if (/^-?\d+(\.\d+)?E\+?\d+$/i.test(val)) val = Number(val).toFixed(0); }
        r[col - 1] = val;
      });
      out.push(r);
    });
    return out;
  }
  /** Строки файла Opt-Drop -> [{ttn, phone, name, day, crmId, src}] */
  function parseOptDrop(rowsX) {
    // названия колонок у Opt-Drop меняются (укр./рус.): сравниваем без регистра, с латинской I вместо украинской І
    function nz(x) { return String(x == null ? '' : x).toLowerCase().replace(/\u0456/g, 'i').replace(/\s+/g, ' ').trim(); }
    function find(hdr, res) { for (var c = 0; c < hdr.length; c++) { var h = nz(hdr[c]); if (h && res.some(function (re) { return re.test(h); })) return c; } return -1; }
    var hi = -1, H = {};
    for (var i = 0; i < Math.min(rowsX.length, 10) && hi < 0; i++) {
      var hd = rowsX[i].map(nz);
      if (hd.indexOf('ттн') >= 0 && hd.some(function (h) { return /^номер телефон/.test(h); })) hi = i;
    }
    if (hi < 0) throw new Error('не нашёл в файле колонки «ТТН» и «Номер телефона» (первые строки файла: ' + rowsX.slice(0, 3).map(function (r) { return r.filter(Boolean).slice(0, 6).join(' | '); }).join(' // ') + ')');
    var hdr = rowsX[hi];
    H.ttn = find(hdr, [/^ттн$/]); H.phone = find(hdr, [/^номер телефон/]); H.name = find(hdr, [/^(отримувач|получатель|одержувач)/]);
    H.day = find(hdr, [/^дата (i|и) (час|время)/, /^дата/]); H.crm = find(hdr, [/^id (в|у) crm/]); H.id = find(hdr, [/^id$/]); H.status = find(hdr, [/^статус$/]);
    var list = [];
    rowsX.slice(hi + 1).forEach(function (r) {
      var ttn = digits(r[H.ttn]); if (!/^\d{14}$/.test(ttn)) return;
      var cancel = H.status >= 0 && /скасован|відмінен|отменен|отменён|cancel/i.test(String(r[H.status] || ''));
      list.push({ cancel: cancel, ttn: ttn, phone: phone380(r[H.phone]), name: String(r[H.name] || '').trim(), day: dayNum(r[H.day]), crmId: digits(r[H.crm]), src: 'Opt-Drop №' + String(r[H.id >= 0 ? H.id : 0] || '') + (cancel ? ', скасовано' : '') });
    });
    return list;
  }
  /** Заказы статуса без ТТН. */
  function collectNoTtn() {
    return collectRows(function (r) {
      var info = (r.dataset.info || (r.id + '|' + (r.dataset.orderid || ''))).split('|');
      return digits(r.dataset.ttn) ? null : { info: info, phone: phone380(r.dataset.phone), orderid: digits(r.dataset.orderid || info[1]), name: r.dataset.buyer || '', day: dayNum(info[2]) };
    }, 'Ищу заказы без ТТН');
  }
  /**
   * Сопоставление: сначала по «ID в CRM» (Opt-Drop), потом по телефону. Если на телефон несколько заказов дропа с разными ТТН,
   * берётся ближайший по дате не раньше заказа в CRM; если так не выбрать - заказ пропускается.
   */
  function matchDrop(crm, drop) {
    var byId = {}, byPhone = {}, usedTtn = {}, pairs = [], miss = [];
    drop.forEach(function (d) { if (d.crmId) byId[d.crmId] = d; if (d.phone) (byPhone[d.phone] = byPhone[d.phone] || []).push(d); });
    crm.forEach(function (o) {
      var d = o.orderid && byId[o.orderid], why = '';
      if (!d) {
        var c = (byPhone[o.phone] || []).filter(function (x) { return !usedTtn[x.ttn]; });
        var uniq = c.filter(function (x, i) { return c.findIndex(function (y) { return y.ttn === x.ttn; }) === i; });
        var live = uniq.filter(function (x) { return !x.cancel; });
        if (live.length) uniq = live;   // отменённые у дропа берём, только если других ТТН на этот телефон нет
        if (uniq.length === 1) d = uniq[0];
        else if (uniq.length > 1) {
          var later = uniq.filter(function (x) { return x.day && o.day && x.day >= o.day; }).sort(function (a, b) { return a.day - b.day; });
          if (later.length && (later.length === 1 || later[0].day < later[1].day)) d = later[0]; else why = 'на этот телефон несколько ТТН у дропа: ' + uniq.map(function (x) { return x.ttn; }).join(', ');
        } else why = 'у дропа нет ТТН на этот телефон';
      }
      if (d) { usedTtn[d.ttn] = true; pairs.push({ o: o, d: d }); } else miss.push({ o: o, why: why });
    });
    return { pairs: pairs, miss: miss };
  }
  async function runDrop(pairs, limit) {
    var ok = 0, errs = [], todo = limit > 0 ? pairs.slice(0, limit) : pairs;
    for (var i = 0; i < todo.length; i++) {
      if (stopFlag) break;
      var p = todo[i];
      progress('Вписываю ТТН', i + 1, todo.length, 'заказ ' + p.o.info[0] + ' · ' + p.d.ttn);
      var e = await editOrder(p.o.info, function (c) {
        var t = c.querySelector('#ttn'), ref = c.querySelector('#ttn_ref');
        if (!t) return 'нет поля ТТН';
        if (digits(t.value)) return 'в заказе уже есть ТТН ' + digits(t.value);
        setField(t, p.d.ttn);
        if (ref) ref.value = '';
        return '';
      });
      if (e) errs.push('заказ ' + p.o.info[0] + ' (' + p.o.name + '): ' + e); else ok++;
    }
    return { ok: ok, errs: errs, total: todo.length, stopped: stopFlag };
  }

  // ---------------------------------------------------------------- 1. стереть ТТН в статусе
  /** Сколько заказов в текущем статусе по счётчику на вкладке статуса (0, если не видно). */
  function expectedCount() {
    var t = document.querySelector('#tab-status-' + curStatus());
    var sp = t && t.querySelector('span');
    return sp ? parseInt(digits(sp.textContent), 10) || 0 : 0;
  }
  var lastScan = { got: 0, expected: 0 };
  /**
   * Проходит весь список статуса: CRM подгружает строки порциями при прокрутке, поэтому крутим список вниз (и немного
   * назад-вперёд), пока не соберём столько строк, сколько показывает счётчик статуса, или пока список не перестанет расти.
   * pick(row) возвращает объект заказа или null.
   */
  async function collectRows(pick, title) {
    var found = {}, order = [], idle = 0, lastCount = -1, expected = expectedCount();
    var maxIdle = expected ? 30 : 8;
    function scrollers(el) {
      var out = [];
      for (; el && el !== document.documentElement; el = el.parentElement) {
        var oy = getComputedStyle(el).overflowY;
        if (el.scrollHeight > el.clientHeight + 5 && /auto|scroll|overlay/.test(oy)) out.push(el);
      }
      return out;
    }
    for (var step = 0; step < 900 && idle < maxIdle; step++) {
      if (stopFlag) break;
      var rs = rows();
      rs.forEach(function (r) {
        if (found[r.id]) return;
        found[r.id] = true;
        var o = pick(r); if (o) order.push(o);
      });
      var n = Object.keys(found).length;
      if (n === lastCount) idle++; else idle = 0;
      lastCount = n;
      progress(title, 0, 0, 'просмотрено заказов: ' + n + (expected ? ' из ' + expected : ''));
      if (expected && n >= expected) break;
      var last = rs[rs.length - 1], back = idle > 0 && idle % 4 === 0;   // если застряли - чуть назад, потом снова вниз
      if (last) {
        last.scrollIntoView({ block: back ? 'center' : 'end' });
        scrollers(last).forEach(function (c) {
          c.scrollTop = back ? Math.max(0, c.scrollHeight - c.clientHeight * 2) : c.scrollHeight;
          c.dispatchEvent(new Event('scroll', { bubbles: true }));
          try { c.dispatchEvent(new WheelEvent('wheel', { deltaY: back ? -600 : 1200, bubbles: true, cancelable: true })); } catch (e) {}
        });
      }
      window.scrollTo(0, back ? Math.max(0, document.documentElement.scrollHeight - window.innerHeight * 2) : document.documentElement.scrollHeight);
      window.dispatchEvent(new Event('scroll'));
      await sleep(500);
    }
    window.scrollTo(0, 0);
    lastScan = { got: Object.keys(found).length, expected: expected };
    return order;
  }
  function scanNote() {
    return lastScan.expected && lastScan.got < lastScan.expected
      ? '<div class="card err">CRM показала только ' + lastScan.got + ' из ' + lastScan.expected + ' заказов статуса. Покрутите список колёсиком до конца и нажмите кнопку ещё раз.</div>' : '';
  }
  /** Заказы статуса с ТТН. */
  function collectWithTtn() {
    return collectRows(function (r) {
      return digits(r.dataset.ttn) ? { info: (r.dataset.info || (r.id + '|' + (r.dataset.orderid || ''))).split('|'), ttn: digits(r.dataset.ttn), name: r.dataset.buyer || '' } : null;
    }, 'Ищу заказы с ТТН');
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
    '#lpr-btn{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;margin:0 8px;padding:0;border:0;background:#ffd24a;color:#2a2d43;cursor:pointer;vertical-align:middle;box-shadow:0 -2px 0 0 #2a2d43,0 2px 0 0 #2a2d43,-2px 0 0 0 #2a2d43,2px 0 0 0 #2a2d43;image-rendering:pixelated}' +
    '#lpr-btn:hover,#lpr-btn.on{background:#5fd49a}' +
    '#lpr-btn.float{position:fixed;top:8px;right:210px;z-index:2147483000}' +
    '#lpr,#lpr *{box-sizing:border-box;font-family:Roboto,Arial,sans-serif;text-transform:none;letter-spacing:0}' +
    '#lpr{position:fixed;z-index:2147483000;width:480px;max-height:calc(100vh - 70px);overflow:auto;background:#fff6e0;color:#2a2d43;font-size:14px;line-height:1.35;display:none;box-shadow:0 -4px 0 0 #2a2d43,0 4px 0 0 #2a2d43,-4px 0 0 0 #2a2d43,4px 0 0 0 #2a2d43,10px 10px 0 0 rgba(42,45,67,.25)}' +
    '#lpr::-webkit-scrollbar{width:10px}#lpr::-webkit-scrollbar-track{background:#f1e6c8}#lpr::-webkit-scrollbar-thumb{background:#2a2d43}' +
    '#lpr .hd{display:flex;align-items:center;gap:12px;padding:16px 18px 12px;font-family:' + PX + ';font-size:12px;color:#2a2d43}#lpr .hd span{font-family:' + PX + '}#lpr .hd svg{color:#ffb300;flex:none}' +
    '#lpr button,#lpr input,#lpr select{border-radius:0 !important;min-width:0;text-shadow:none}' +
    '#lpr .hd .x{margin-left:auto;flex:none;border:0;background:#ff8a7a;color:#2a2d43;width:24px;height:24px;font-size:16px;font-weight:700;line-height:24px;padding:0;cursor:pointer;box-shadow:0 -2px 0 0 #2a2d43,0 2px 0 0 #2a2d43,-2px 0 0 0 #2a2d43,2px 0 0 0 #2a2d43}#lpr .hd .x:active{transform:translate(2px,2px)}' +
    '#lpr .tabs{display:flex;gap:9px;margin:0 21px;padding:0}' +
    '#lpr .tabs button{flex:1 1 0;white-space:nowrap;overflow:hidden;height:auto;border:0;background:#fff;color:#6b7089;padding:10px 2px;font-family:' + PX + ';font-size:8px;cursor:pointer;box-shadow:0 -3px 0 0 #2a2d43,0 3px 0 0 #2a2d43,-3px 0 0 0 #2a2d43,3px 0 0 0 #2a2d43;line-height:1.4}' +
    '#lpr .tabs button:hover{color:#2a2d43}#lpr .tabs button.on{background:#6cb8ff;color:#2a2d43}' +
    '#lpr .pane{padding:16px 18px 18px;display:none}#lpr .pane.on{display:block}' +
    '#lpr p{margin:0 0 12px;color:#555a75}#lpr label{display:block;font-size:13px;font-weight:700;color:#2a2d43;margin:0 0 6px}' +
    '#lpr select,#lpr input{width:100%;padding:9px 8px;border:0;background:#fff;color:#2a2d43;font-size:14px;margin:3px 0 14px;outline:none;box-shadow:0 -3px 0 0 #2a2d43,0 3px 0 0 #2a2d43,-3px 0 0 0 #2a2d43,3px 0 0 0 #2a2d43;border-radius:0;height:auto}' +
    '#lpr select:focus,#lpr input:focus{box-shadow:0 -3px 0 0 #3b8fe0,0 3px 0 0 #3b8fe0,-3px 0 0 0 #3b8fe0,3px 0 0 0 #3b8fe0}' +
    '#lpr option{background:#fff;color:#2a2d43}' +
    '#lpr .row{display:flex;gap:16px;align-items:flex-end}#lpr .row>div{flex:1}#lpr .row>div.sm{flex:0 0 130px}' +
    '#lpr .btn{display:inline-block;border:0;padding:10px 14px;font-size:14px;font-weight:700;cursor:pointer;margin:3px 14px 9px 3px;color:#2a2d43;line-height:1.3;border-radius:0;box-shadow:0 -3px 0 0 #2a2d43,0 3px 0 0 #2a2d43,-3px 0 0 0 #2a2d43,3px 0 0 0 #2a2d43,inset -3px -3px 0 0 rgba(42,45,67,.18),inset 3px 3px 0 0 rgba(255,255,255,.6)}#lpr .btn:active{transform:translate(2px,2px);box-shadow:0 -3px 0 0 #2a2d43,0 3px 0 0 #2a2d43,-3px 0 0 0 #2a2d43,3px 0 0 0 #2a2d43,inset 3px 3px 0 0 rgba(42,45,67,.22)}' +
    '#lpr .b1{background:#6cb8ff}#lpr .b2{background:#d9dde6}#lpr .b3{background:#ff8a7a}#lpr .b4{background:#5fd49a}' +
    '#lpr .btn:hover{filter:brightness(1.06) saturate(1.1)}' +
    '#lpr .out{margin-top:14px}' +
    '#lpr .card{background:#fff;padding:12px;margin:3px 3px 16px;box-shadow:0 -3px 0 0 #2a2d43,0 3px 0 0 #2a2d43,-3px 0 0 0 #2a2d43,3px 0 0 0 #2a2d43}' +
    '#lpr .card.warn{background:#fff1bf}#lpr .card.err{background:#ffd9d3}#lpr .card.ok{background:#d6f5e4}' +
    '#lpr .big{font-family:' + PX + ';font-size:20px;color:#2a2d43;margin-bottom:8px}#lpr .mut{color:#6b7089;font-size:13px}#lpr b{color:#2a2d43;font-weight:700}' +
    '#lpr .list{max-height:180px;overflow:auto;margin-top:8px;font-size:13px;line-height:1.4}#lpr .list div{padding:5px 0;border-top:2px dashed #b9bfce;word-break:break-word;color:#2a2d43}' +
    '#lpr .prog{padding:18px;display:none}#lpr .prog .t{font-family:' + PX + ';color:#2a2d43;font-size:10px;font-weight:normal}' +
    '#lpr .bar{height:20px;background:#2a2d43;margin:12px 3px 14px;padding:3px;box-shadow:0 -3px 0 0 #2a2d43,0 3px 0 0 #2a2d43,-3px 0 0 0 #2a2d43,3px 0 0 0 #2a2d43}' +
    '#lpr .bar i{display:block;height:100%;width:0;background:repeating-linear-gradient(90deg,#5fd49a 0,#5fd49a 10px,#2a2d43 10px,#2a2d43 12px);transition:width .3s steps(8)}';

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
      '<div class="body"><div class="tabs"><button type="button" data-tab="create" class="on">Создать</button><button type="button" data-tab="clear">Стереть ТТН</button><button type="button" data-tab="drop">ТТН дропа</button><button type="button" data-tab="sheet">Настройки</button></div>' +
      '<div class="pane on" data-tab="create"><p>Берёт заказы с листа «Заказы на СРМ», создаёт их в статусе «Импорт» с ТТН и оплатой «Післяплата». Результат сам записывает в колонку «Статус».</p>' +
      '<div class="row"><div class="sm"><label>Не больше (0 = все)</label><input type="number" id="lpr-limit-c" value="0" min="0"></div><div><button class="btn b1" id="lpr-load" type="button" style="margin-bottom:10px">Загрузить из таблицы</button></div></div>' +
      '<div class="out" id="lpr-out-create"></div></div>' +
      '<div class="pane" data-tab="clear"><p>Открывает каждый заказ в выбранном статусе, стирает ТТН и сохраняет.</p>' +
      '<label>Статус</label><select id="lpr-status">' + sel + '</select>' +
      '<div class="row"><div class="sm"><label>Не больше (0 = все)</label><input type="number" id="lpr-limit-d" value="0" min="0"></div><div><button class="btn b1" id="lpr-clear" type="button" style="margin-bottom:10px">Найти заказы с ТТН</button></div></div>' +
      '<div class="out" id="lpr-out-clear"></div></div>' +
      '<div class="pane" data-tab="drop"><p>Берёт ТТН у дроп-платформы и вписывает их в заказы выбранного статуса, у которых ТТН пустая. Заказы сопоставляются по ID в CRM и телефону.</p>' +
      '<label>Статус в CRM</label><select id="lpr-dstatus">' + sel + '</select>' +
      '<label>Откуда брать ТТН</label><select id="lpr-dsrc"><option value="aveopt">Aveopt - сайт (нужен вход на aveopt.com.ua в этом браузере)</option><option value="optdrop">Opt-Drop - файл Excel с заказами</option></select>' +
      '<div id="lpr-dfile-w" style="display:none"><label>Файл .xlsx из Opt-Drop</label><input type="file" id="lpr-dfile" accept=".xlsx"></div>' +
      '<div class="row"><div class="sm"><label>Не больше (0 = все)</label><input type="number" id="lpr-limit-t" value="0" min="0"></div><div><button class="btn b1" id="lpr-drop" type="button" style="margin-bottom:10px">Найти ТТН</button></div></div>' +
      '<div class="out" id="lpr-out-drop"></div></div>' +
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
    panel.querySelector('#lpr-drop').onclick = startDrop;
    (function () {
      var st = panel.querySelector('#lpr-dstatus'), src = panel.querySelector('#lpr-dsrc'), fw = panel.querySelector('#lpr-dfile-w');
      function sync() { fw.style.display = src.value === 'optdrop' ? 'block' : 'none'; }
      // по названию статуса подставляем источник: «АВЕ ОПТ» - Aveopt, «ОПТ ДРОП» - файл Opt-Drop
      function byStatus() { var nm = (st.options[st.selectedIndex] || {}).text || ''; src.value = /аве/i.test(nm) ? 'aveopt' : /дроп/i.test(nm) ? 'optdrop' : src.value; sync(); }
      Array.prototype.forEach.call(st.options, function (o) { if (/аве\s*опт/i.test(o.text)) st.value = o.value; });
      st.onchange = byStatus; src.onchange = sync; byStatus();
    })();
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
    if (!list.length) { out('clear', scanNote() + '<div class="card ok">В статусе «' + esc(name) + '» нет заказов с ТТН.</div>'); return; }
    out('clear', scanNote() + '<div class="card warn"><div class="big">' + list.length + '</div>заказов с ТТН в статусе «' + esc(name) + '».<br>ТТН будет стёрта у <b>' + n + '</b>. Отменить это нельзя.</div>' +
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

  // --- 3
  var DROP_KEY = 'lprobot_drop';
  async function startDrop() {
    var id = panel.querySelector('#lpr-dstatus').value, src = panel.querySelector('#lpr-dsrc').value, limit = lim('#lpr-limit-t');
    try { localStorage.removeItem(DROP_KEY); } catch (e) {}
    if (src === 'optdrop') {
      var f = panel.querySelector('#lpr-dfile').files[0];
      if (!f) { out('drop', '<div class="card err">Выберите файл .xlsx, скачанный из Opt-Drop.</div>'); return; }
      try {
        var list = parseOptDrop(await readXlsx(await f.arrayBuffer()));
        if (!list.length) { out('drop', '<div class="card err">В файле нет заказов с ТТН.</div>'); return; }
        localStorage.setItem(DROP_KEY, JSON.stringify(list));   // файл переживает переход на страницу статуса
      } catch (e) { out('drop', '<div class="card err">Не удалось прочитать файл: ' + esc(e.message) + '</div>'); return; }
    }
    if (!onOrders() || curStatus() !== id) { saveJob({ t: Date.now(), type: 'drop', status: id, limit: limit, src: src }); location.href = '/orders/?status=' + id; return; }
    scanDrop(src);
  }
  async function scanDrop(src) {
    stopFlag = false; setBusy(true); progress('Ищу заказы без ТТН', 0, 0, '');
    await waitFor(function () { return typeof window.Orders === 'function' && document.querySelector('.table-data'); }, 15000);
    await sleep(2500);
    var crm = await collectNoTtn(), drop = [], note = '';
    var name = (statusOptions().filter(function (s) { return s.id === curStatus(); })[0] || {}).name || curStatus();
    function fin(html) { setBusy(false); showTab('drop'); out('drop', html); }
    var sn = scanNote();
    if (!crm.length) return fin(sn + '<div class="card ok">В статусе «' + esc(name) + '» нет заказов без ТТН.</div>');
    if (src === 'optdrop') {
      try { drop = JSON.parse(localStorage.getItem(DROP_KEY) || '[]'); } catch (e) {}
      if (!drop.length) return fin('<div class="card err">Файл Opt-Drop не загружен: выберите его и нажмите «Найти ТТН» ещё раз.</div>');
    } else {
      // листаем Aveopt, пока не найдём ТТН для всех заказов или не уйдём раньше самого старого заказа CRM
      var oldest = crm.reduce(function (m, o) { return o.day && (!m || o.day < m) ? o.day : m; }, 0), need = {};
      crm.forEach(function (o) { need[o.phone] = true; });
      var past = 0;
      for (var p = 1; p <= 150 && !stopFlag; p++) {
        progress('Читаю заказы Aveopt', 0, 0, 'страница ' + p + ', найдено ТТН: ' + drop.length);
        var r = await aveoptPage(p);
        if (!r.ok) return fin('<div class="card err">Aveopt: ' + esc(r.error || 'ошибка') + '</div>');
        var pg = parseAveopt(r.html);
        if (!pg.ok) return fin('<div class="card err">Не вижу списка заказов Aveopt. Войдите на aveopt.com.ua в этом браузере и повторите.</div>');
        drop = drop.concat(pg.list);
        pg.list.forEach(function (d) { delete need[d.phone]; });
        if (!pg.rows || !Object.keys(need).length) break;
        if (oldest && pg.minDay && pg.minDay < oldest) { if (++past >= 2) break; }   // ещё одна страница с запасом и хватит
        await sleep(300);
      }
      note = 'Прочитано заказов Aveopt с ТТН: ' + drop.length + ' (страниц: ' + p + ').';
    }
    var m = matchDrop(crm, drop), limit = lim('#lpr-limit-t'), n = limit > 0 ? Math.min(limit, m.pairs.length) : m.pairs.length;
    var missHtml = m.miss.length ? '<div class="card warn"><b>Без ТТН останутся: ' + m.miss.length + '</b><div class="list">' +
      m.miss.map(function (x) { return '<div>' + esc(x.o.name) + ' · ' + esc(x.o.phone) + ': ' + esc(x.why) + '</div>'; }).join('') + '</div></div>' : '';
    if (!m.pairs.length) return fin(sn + '<div class="card warn">Заказов без ТТН в «' + esc(name) + '»: <b>' + crm.length + '</b>, но ТТН для них у дропа пока нет. ' + esc(note) + '</div>' + missHtml);
    fin(sn + '<div class="card ok"><div class="big">' + m.pairs.length + ' из ' + crm.length + '</div>заказов без ТТН в «' + esc(name) + '» получат ТТН. ' + esc(note) + '<br>Будет вписано: <b>' + n + '</b>' +
      '<div class="list">' + m.pairs.slice(0, 200).map(function (x) { return '<div>' + esc(x.o.name) + ' · ' + esc(x.o.phone) + ' → ' + x.d.ttn + ' (' + esc(x.d.src) + ')</div>'; }).join('') + '</div></div>' + missHtml +
      '<button class="btn b4" id="lpr-dgo" type="button">Вписать ТТН</button><button class="btn b2" id="lpr-dno" type="button">Отмена</button>');
    panel.querySelector('#lpr-dno').onclick = function () { out('drop', ''); };
    panel.querySelector('#lpr-dgo').onclick = async function () {
      setBusy(true); out('drop', '');
      var r = await runDrop(m.pairs, limit);
      setBusy(false);
      out('drop', '<div class="card ' + (r.errs.length ? 'warn' : 'ok') + '"><div class="big">' + r.ok + ' из ' + r.total + '</div>ТТН вписана' + (r.stopped ? ' (остановлено)' : '') + '</div>' +
        (r.errs.length ? '<div class="card err"><b>Ошибки</b><div class="list">' + r.errs.map(function (x) { return '<div>' + esc(x) + '</div>'; }).join('') + '</div></div>' : '') + missHtml);
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
    } else if (job.type === 'drop' && onOrders() && curStatus() === job.status) {
      panel.querySelector('#lpr-dstatus').value = job.status; panel.querySelector('#lpr-dsrc').value = job.src;
      panel.querySelector('#lpr-limit-t').value = job.limit;
      scanDrop(job.src);
    } else if (job.type === 'create' && onOrders() && curStatus() === 'all') {
      panel.querySelector('#lpr-limit-c').value = job.limit;
      loadCreate(true);
    }
  }
  if (document.readyState === 'complete') setTimeout(start, 800); else window.addEventListener('load', function () { setTimeout(start, 800); });
  window.__lpRobotTest = { prepOrders: prepOrders, phpSer: phpSer, runCreate: runCreate, apiCreate: apiCreate, findByCode: findByCode, readXlsx: readXlsx, parseOptDrop: parseOptDrop, parseAveopt: parseAveopt, matchDrop: matchDrop };
})();
