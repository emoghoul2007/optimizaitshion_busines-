/**
 * Меню «Заявки» для таблицы парсера Binotel.
 *
 * Установка (один раз):
 *   1. В таблице: Расширения -> Apps Script.
 *   2. Слева «+» -> Скрипт, назовите файл «Заявки», вставьте весь этот код, сохраните.
 *   3. Вверху в списке функций выберите ustanovitMenu и нажмите «Выполнить», разрешите доступ.
 *   4. Обновите страницу таблицы: справа от «Справка» появится меню «Заявки».
 *
 * Скрипт не трогает другие скрипты и меню в этой таблице.
 */

var ZAYAVKI_CFG = {
  TZ: 'Europe/Kyiv',            // по какому времени считается «сегодня»
  STATS_SHEET: 'Статистика',
  BIG_SHEET: '1800+',
  BIG_SUM: 1500,                // заказы на эту сумму и больше уходят в лист «1800+»
  ORDER_STATUSES: ['SUCCESS'],  // какие статусы считаются заказом
  COL_ID: 1,                    // A  Call ID
  COL_STATUS: 3,                // C  Статус
  COL_PHONE: 5,                 // E  Телефон
  COL_PRODUCT: 6,               // F  Товар
  COL_SUM: 7,                   // G  Сумма
  WIDTH: 16,                    // сколько колонок переносить в «пул» и «1800+» (A:P)
  COL_LIST: 17,                 // Q  Товар по списку (заполняет парсер)
  COL_OTHER: 18,                // R  Нет в списке (заполняет парсер)
  READ_WIDTH: 18,               // сколько колонок читать с листа дня (A:R)

  // Пул товаров и отправка в LP-CRM
  POOL_SHEET: 'пул',
  COL_PICK: 8,                  // H в листе «Статистика»: названия товаров, которые нужно выпулить
  COL_PICK_ID: 9,               // I в листе «Статистика»: ID товара в СРМ 1, J - ID в СРМ 2
  CRM_SITE: 'google.com',   // «сайт-лендинг покупателя»: под этим названием заказы видны в CRM
  CRM_COUNTRY: 'UA',
  CRM_DELIVERY: '',             // ID способа доставки в LP-CRM (можно оставить пустым)
  CRM_PAYMENT: ''               // ID способа оплаты в LP-CRM (можно оставить пустым)
};

/**
 * Две LP-CRM. Ключи (и адрес второй CRM) хранятся в свойствах скрипта и вводятся через меню.
 * idIndex: 0 - ID товара из колонки I «Статистики» / R «пула», 1 - из колонки J / S.
 */
var CRM_LIST = [
  { name: 'СРМ 1', keyProp: 'LP_CRM_KEY',  urlProp: '', defaultUrl: 'https://deltafund.lp-crm.biz', office: '1', idIndex: 0 },
  { name: 'СРМ 2', keyProp: 'LP_CRM2_KEY', urlProp: 'LP_CRM2_URL', defaultUrl: '', office: '1', idIndex: 1 }
];

/** Запустите один раз из редактора: добавляет меню при каждом открытии таблицы. */
function ustanovitMenu() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pokazatMenuZayavki') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pokazatMenuZayavki').forSpreadsheet(ss).onOpen().create();
  pokazatMenuZayavki();
}

function pokazatMenuZayavki() {
  var ui = SpreadsheetApp.getUi();
  var vse = ui.createMenu('Все кнопки списком')
  .addItem('Собрать статистику за сегодня', 'sobratStatistiku')
      .addItem('Выпулить ' + ZAYAVKI_CFG.BIG_SUM + '+', 'vypulitKrupnye')
      .addSeparator()
      .addItem('Выпулить товары в «' + ZAYAVKI_CFG.POOL_SHEET + '»', 'vypulitTovary')
      .addItem('Выгрузить выделенные заказы для дропа (TXT)', 'vygruzitDlyaDropa')
      .addSeparator()
      .addItem('Перенести выделенные в «НП лиды»', 'perenestiVNpLidy')
      .addItem('Выпулить заказы для Скайлаба', 'vypulitDlyaSkailab')
      .addItem('Выпулить заказы для Щербака', 'vypulitDlyaShcherbak')
      .addItem('Создать ТТН (НП лиды без ТТН)', 'sozdatTtn')
      .addItem('Печать наклеек 100×100', 'pechatNakleek')
      .addItem('Обновить статистику НП лидов', 'obnovitStatistikuNpLidov')
      .addItem('Перенести выделенные в «Заказы на СРМ» (для расширения)', 'perenestiVZakazyNaCrm')
      .addSeparator()
      .addItem('Выпулить заказы в СРМ 1', 'otpravitVCrm1')
      .addItem('Выпулить заказы в СРМ 2', 'otpravitVCrm2')
      .addItem('Выпулить ТТН из статуса «ОПТ ДРОП»', 'vypulitTtnOptDrop')
      .addItem('Выпулить ТТН из статуса «АВЕ ОПТ»', 'vypulitTtnAveOpt')
      .addItem('Удалить ТТН (лист «ТТН на удаление»)', 'udalitTtn')
      .addSeparator()
      .addItem('Проверить пул ИИ: Sonnet 5.5', 'proveritPulSonnet')
      .addItem('Проверить пул ИИ: Haiku 5.5', 'proveritPulHaiku')
      .addItem('Как идёт проверка ИИ', 'statusProverkiIi')
      .addItem('Отсортировать весь пул: ✓, ≈, ⚠', 'sortirovatPulKnopka')
      .addItem('Остановить проверку ИИ', 'ostanovitProverkuIi');
  var teh = ui.createMenu('Технические: ключи и ссылки')
      .addItem('Указать ключ СРМ 1', 'ukazatKlyuchCrm1')
      .addItem('Указать выходной ключ СРМ 1 (чтение заказов)', 'ukazatVyhodnoiKlyuchCrm1')
      .addItem('Указать адрес и ключ СРМ 2', 'ukazatKlyuchCrm2')
      .addItem('Указать ключ Новой Почты', 'ukazatKlyuchNp')
      .addItem('Указать ключ НП из LP-CRM (удаление ТТН)', 'ukazatKlyuchNpCrm')
      .addItem('Ссылка для расширения «Робот LP-CRM»', 'ssylkaDlyaRasshireniya')
      .addItem('Указать ключ OpenRouter', 'ukazatKlyuchOpenRouter');
  ui.createMenu('Заявки')
    .addItem('▶ Открыть панель', 'pokazatPanel')
    .addItem('▶ Открыть панель в широком окне', 'pokazatPanelOkno')
    .addSeparator()
    .addSubMenu(vse)
    .addSeparator()
    .addItem('Установить расширение «Робот LP-CRM»', 'ustanovitRasshirenieCrm')
    .addItem('Установить / обновить скрипт таблицы', 'ustanovitSkriptTablitsy')
    .addSubMenu(teh)
    .addToUi();
}

/* ===================== Панель «Заявки» (боковая, пиксельная) и обновления =====================
 * REPO - ваш репозиторий на GitHub в виде «логин/название». В нём лежат: Zayavki.gs, version.json,
 * а в разделе Releases - архив lpcrm-robot.zip. Ключей в репозитории быть не должно.
 */
var UPDATE_CFG = {
  REPO: 'emoghoul2007/optimizaitshion_busines-',                 // например 'oleg/zayavki' - впишите после создания репозитория
  SCRIPT_VERSION: '3.2.0'
};

function pokazatPanel() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(PANEL_HTML_).setTitle('Автоматизация'));
}

/** То же в отдельном окне: боковая панель Google всегда 300 px, окно - шире и его можно двигать. */
function pokazatPanelOkno() {
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(PANEL_HTML_).setWidth(480).setHeight(720), 'Автоматизация');
}

/** Данные для шапки панели: идёт ли проверка ИИ, сколько осталось, расход. */
function dannyePaneli() {
  var props = PropertiesService.getScriptProperties();
  var pool = SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  var left = 0;
  try { if (pool && pool.getLastRow() > 1) left = strokiDlyaProverki_(pool).length; } catch (e) {}
  var modelKey = props.getProperty('II_MODEL');
  var idet = !!modelKey && ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'prodolzhitProverkuIi'; });
  return { idet: idet, model: idet ? II_MODELS[modelKey].name : '', left: left,
           total: Number(props.getProperty('II_TOTAL') || 0),
           last: props.getProperty('II_LAST_RUN') || '',
           day: props.getProperty('II_DAY') === Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM.yyyy') ? Number(props.getProperty('II_DAY_COUNT') || 0) : 0,
           cost: parseFloat(props.getProperty('II_COST_OR') || '0'), version: UPDATE_CFG.SCRIPT_VERSION };
}

function oknoUstanovki_(title, html) {
  var css = '<style>body{font:14px/1.5 Arial,sans-serif;color:#222;margin:0;padding:4px 2px}h3{margin:0 0 8px}ol{padding-left:20px;margin:8px 0}' +
    'a.b{display:inline-block;background:#3b5dc9;color:#fff;padding:8px 14px;text-decoration:none;margin:6px 0;border-radius:2px}' +
    '.n{background:#fff4d6;padding:8px;margin:8px 0}.ok{background:#e2f5e7;padding:8px;margin:8px 0}code{background:#eee;padding:1px 4px}</style>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(css + html).setWidth(520).setHeight(430), title);
}

function versiiNaGithub_() {
  if (!UPDATE_CFG.REPO) return null;
  try {
    var r = UrlFetchApp.fetch('https://raw.githubusercontent.com/' + UPDATE_CFG.REPO + '/main/version.json', { muteHttpExceptions: true });
    return r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null;
  } catch (e) { return null; }
}

function ustanovitRasshirenieCrm() {
  var v = versiiNaGithub_();
  var link = UPDATE_CFG.REPO ? 'https://github.com/' + UPDATE_CFG.REPO + '/releases/latest/download/lpcrm-robot.zip' : '';
  oknoUstanovki_('Робот LP-CRM: установка и обновление',
    (link ? '<a class="b" href="' + link + '" target="_blank">Скачать lpcrm-robot.zip</a>' + (v && v.extension ? ' &nbsp;последняя версия: <b>' + v.extension + '</b>' : '')
          : '<div class="n">Ссылка на GitHub ещё не указана: впишите репозиторий в <code>UPDATE_CFG.REPO</code> в начале скрипта.</div>') +
    '<h3>Установка</h3><ol><li>Распакуйте архив в постоянную папку, например <code>Документы/lpcrm-robot</code>.</li>' +
    '<li>В Chrome откройте <code>chrome://extensions</code>, включите «Режим разработчика».</li>' +
    '<li>Нажмите «Загрузить распакованное расширение» и выберите эту папку.</li>' +
    '<li>Откройте LP-CRM, нажмите значок робота, вкладка «Настройки»: адрес веб-приложения, код доступа ' +
    '(кнопка «Ссылка и код для робота») и входной ключ API вашей CRM.</li></ol>' +
    '<h3>Обновление</h3><ol><li>Скачайте новый архив и замените файлы в той же папке.</li>' +
    '<li>В <code>chrome://extensions</code> нажмите круглую стрелку у «Робот LP-CRM» и обновите страницу CRM.</li></ol>' +
    '<div class="n">Настройки и ключ при обновлении сохраняются: они хранятся в браузере, а не в файлах.</div>');
}

function ustanovitSkriptTablitsy() {
  var v = versiiNaGithub_(), cur = UPDATE_CFG.SCRIPT_VERSION;
  var link = UPDATE_CFG.REPO ? 'https://raw.githubusercontent.com/' + UPDATE_CFG.REPO + '/main/Zayavki.gs' : '';
  var st = !UPDATE_CFG.REPO ? '<div class="n">Ссылка на GitHub ещё не указана: впишите репозиторий в <code>UPDATE_CFG.REPO</code> в начале скрипта.</div>'
    : !v ? '<div class="n">Не удалось узнать последнюю версию на GitHub. У вас: <b>' + cur + '</b>.</div>'
    : v.script === cur ? '<div class="ok">У вас последняя версия: <b>' + cur + '</b>.</div>'
    : '<div class="n">Доступна версия <b>' + v.script + '</b>, у вас <b>' + cur + '</b>.' + (v.notes ? '<br>' + String(v.notes).replace(/[<>&]/g, '') : '') + '</div>';
  oknoUstanovki_('Скрипт таблицы: установка и обновление', st +
    (link ? '<a class="b" href="' + link + '" target="_blank">Открыть Zayavki.gs</a>' : '') +
    '<h3>Обновление</h3><ol><li>Откройте файл по кнопке, выделите всё (Ctrl+A) и скопируйте.</li>' +
    '<li>В таблице: «Расширения» → «Apps Script», откройте файл «Заявки», выделите всё и вставьте новый код.</li>' +
    '<li>Впишите свой репозиторий в <code>UPDATE_CFG.REPO</code>, сохраните (Ctrl+S) и обновите страницу таблицы.</li></ol>' +
    '<h3>Первая установка в новой таблице</h3><ol><li>«Расширения» → «Apps Script» → «+» → «Скрипт», назовите «Заявки», вставьте код.</li>' +
    '<li>Выберите функцию <code>ustanovitMenu</code>, нажмите «Выполнить» и разрешите доступ.</li>' +
    '<li>Ключи вводятся в панели, раздел «Ключи». Они хранятся в таблице и при обновлении кода не пропадают.</li></ol>');
}

/*PANEL_BEGIN*/
var PANEL_HTML_ = "<!DOCTYPE html>\n<html><head><base target=\"_top\"><meta charset=\"utf-8\">\n<style>\n@import url('https://fonts.googleapis.com/css2?family=Press+Start+2P&display=swap');\n:root{\n  --bg:#fff6e0; --panel:#ffffff; --line:#2a2d43; --text:#2a2d43; --mut:#6b7089;\n  --green:#5fd49a; --yellow:#ffd24a; --red:#ff8a7a; --blue:#6cb8ff; --cyan:#7fe3e0; --violet:#c4a5ff; --gray:#d9dde6;\n  --px:3px;\n}\n*{box-sizing:border-box;margin:0;padding:0}\nhtml,body{background:var(--bg);color:var(--text);font-family:Roboto,Arial,sans-serif;font-size:15px;line-height:1.25}\nbody{padding:10px 9px 24px}\n.wide{max-width:460px;margin:0 auto}\n.title{font-family:'Press Start 2P','Courier New',monospace;font-size:11px;color:var(--text);display:flex;align-items:center;gap:8px;margin:2px 2px 12px;letter-spacing:.5px}\nsvg{shape-rendering:crispEdges;display:block;flex:none}\n.pix{box-shadow:0 calc(var(--px)*-1) 0 0 var(--line),0 var(--px) 0 0 var(--line),calc(var(--px)*-1) 0 0 0 var(--line),var(--px) 0 0 0 var(--line)}\n.box{background:var(--panel);margin:var(--px);padding:9px 9px 10px}\n.status{margin-bottom:14px}\n.status h3{font-family:'Press Start 2P',monospace;font-size:8px;color:var(--text);margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}\n.led{width:10px;height:10px;background:#b9bfce;display:inline-block;margin-right:6px;outline:2px solid var(--line)}\n.led.on{background:var(--green);animation:bl 1s steps(2) infinite}\n@keyframes bl{50%{background:#fff}}\n.bar{height:16px;background:var(--line);padding:3px;display:flex;gap:2px;margin:6px 0 8px}\n.bar i{flex:1;background:#4a4f6e}\n.bar i.f{background:var(--green)}\n.row{display:flex;justify-content:space-between;gap:8px;color:var(--mut);font-size:14px;margin-top:2px}\n.row b{color:var(--text);font-weight:700;text-align:right}\n.sec{margin:14px 0 0}\n.sec>.hd{cursor:pointer;display:flex;align-items:center;gap:8px;padding:9px 8px;margin:var(--px);color:var(--text);font-family:'Press Start 2P',monospace;font-size:9px;user-select:none}\n.sec>.hd .n{margin-left:auto;font-size:8px;opacity:.7}\n.sec>.hd .ar{width:9px;height:9px;background:currentColor;clip-path:polygon(0 0,100% 0,50% 100%)}\n.sec.closed>.hd .ar{clip-path:polygon(0 0,100% 50%,0 100%)}\n.sec.closed>.list{display:none}\n.list{padding:8px 3px 2px;display:grid;gap:9px}\n.btn{all:unset;box-sizing:border-box;cursor:pointer;display:flex;align-items:center;gap:9px;margin:var(--px);padding:9px 9px;background:var(--c,var(--blue));color:var(--text);font:700 15px/1.2 Roboto,Arial,sans-serif;\n  box-shadow:0 calc(var(--px)*-1) 0 0 var(--line),0 var(--px) 0 0 var(--line),calc(var(--px)*-1) 0 0 0 var(--line),var(--px) 0 0 0 var(--line),inset calc(var(--px)*-1) calc(var(--px)*-1) 0 0 rgba(42,45,67,.18),inset var(--px) var(--px) 0 0 rgba(255,255,255,.6)}\n.btn:hover{filter:brightness(1.06) saturate(1.1)}\n.btn:active,.btn.busy{box-shadow:0 calc(var(--px)*-1) 0 0 var(--line),0 var(--px) 0 0 var(--line),calc(var(--px)*-1) 0 0 0 var(--line),var(--px) 0 0 0 var(--line),inset var(--px) var(--px) 0 0 rgba(42,45,67,.22);transform:translate(2px,2px)}\n.btn.busy{pointer-events:none;filter:saturate(.6)}\n.btn.busy .t::after{content:' ...'}\n.btn .t{flex:1;text-align:left}\n.btn small{display:block;font-weight:400;font-size:13px;color:#454a66;margin-top:1px}\n.c-green{--c:var(--green)}.c-blue{--c:var(--blue)}.c-cyan{--c:var(--cyan)}.c-red{--c:var(--red)}.c-yellow{--c:var(--yellow)}.c-violet{--c:var(--violet)}.c-gray{--c:var(--gray)}\n.s1>.hd{background:var(--green)}.s2>.hd{background:var(--blue)}.s3>.hd{background:var(--yellow)}.s4>.hd{background:var(--violet)}.s5>.hd{background:var(--red)}.s6>.hd{background:var(--cyan)}.s7>.hd{background:var(--gray)}\n.log{margin-top:14px;font-size:13px;color:var(--mut);max-height:120px;overflow:auto;white-space:pre-wrap}\n.log .e{color:#c0392b}.log .o{color:#1e8a55}\n.ver{margin-top:12px;text-align:center;color:#a9a48f;font-size:12px}\n\n.pet{display:flex;justify-content:space-between;align-items:center;font-size:14px;margin-bottom:0}\n#pet-name{font-weight:700;cursor:pointer;border-bottom:2px dotted var(--mut)}\n#pet-name input{font:inherit;width:120px;border:2px solid var(--line);padding:1px 4px}\n#pet-lvl{color:var(--mut)}\n.fx{position:fixed;z-index:51;pointer-events:none;font-size:13px;font-weight:700;color:#2a2d43;white-space:nowrap;animation:fx 1s steps(8) forwards}\n@keyframes fx{to{transform:translateY(-46px);opacity:0}}\n#slime{position:fixed;left:0;top:0;width:54px;height:36px;z-index:50;cursor:pointer;will-change:transform;transform-origin:0 0}\n#slime svg{width:54px;height:36px;overflow:visible}\nbody{padding-bottom:64px}\n</style></head>\n<body><div class=\"wide\">\n<div class=\"title\"><svg id=\"logo\" width=\"24\" height=\"24\" viewBox=\"0 0 8 8\"></svg>АВТОМАТИЗАЦИЯ</div>\n\n<div class=\"box pix status\">\n  <h3><span><span class=\"led\" id=\"led\"></span>ПРОВЕРКА ИИ</span><span id=\"st-state\">...</span></h3>\n  <div class=\"bar\" id=\"bar\"></div>\n  <div class=\"row\"><span>Не проверено</span><b id=\"st-left\">-</b></div>\n  <div class=\"row\"><span>Потрачено</span><b id=\"st-cost\">-</b></div>\n  <div class=\"row\"><span>Сегодня проверено</span><b id=\"st-day\">-</b></div>\n  <div class=\"row\"><span>Посл. порция</span><b id=\"st-last\" style=\"font-size:12px\">-</b></div>\n</div>\n\n<div class=\"box pix pet\"><span id=\"pet-name\" title=\"Нажмите, чтобы переименовать\"></span><span id=\"pet-lvl\"></span></div>\n<div id=\"app\"></div>\n<div class=\"log pix box\" id=\"log\" hidden></div>\n<div class=\"ver\" id=\"ver\"></div>\n\n</div><script>\n// ---- пиксельные иконки 8x8 ('#' = пиксель)\nvar IC = {\n  play:['..#.....','..##....','..###...','..####..','..###...','..##....','..#.....','........'],\n  stop:['........','.######.','.######.','.######.','.######.','.######.','.######.','........'],\n  chk:['........','......#.','.....##.','#...##..','##.##...','.###....','..#.....','........'],\n  box:['.######.','#......#','#.####.#','#......#','#......#','#......#','#......#','.######.'],\n  key:['..###...','.#...#..','.#...#..','..###...','...#....','...##...','...#....','...##...'],\n  send:['#.......','###.....','#####...','#######.','#####...','###.....','#.......','........'],\n  trash:['..####..','########','.#.##.#.','.#.##.#.','.#.##.#.','.#.##.#.','.######.','........'],\n  sort:['...#....','..###...','.#####..','...#....','...#....','.#####..','..###...','...#....'],\n  find:['..###...','.#...#..','#.....#.','#.....#.','.#...#..','..###.#.','.....###','......#.'],\n  print:['..####..','..#..#..','########','#......#','#..##..#','########','..#..#..','..####..'],\n  dl:['...##...','...##...','...##...','.######.','..####..','...##...','########','........'],\n  list:['#.######','........','#.######','........','#.######','........','#.######','........'],\n  star:['...##...','...##...','########','.######.','..####..','..#..#..','.#....#.','........']\n};\nfunction icon(name,color){\n  var m=IC[name]||IC.star,r='';\n  m.forEach(function(row,y){for(var x=0;x<8;x++)if(row[x]==='#')r+='<rect x=\"'+x+'\" y=\"'+y+'\" width=\"1\" height=\"1\"/>';});\n  return '<svg width=\"20\" height=\"20\" viewBox=\"0 0 8 8\" fill=\"'+(color||'#2a2d43')+'\">'+r+'</svg>';\n}\ndocument.getElementById('logo').innerHTML=(function(){var r='';IC.star.forEach(function(row,y){for(var x=0;x<8;x++)if(row[x]==='#')r+='<rect x=\"'+x+'\" y=\"'+y+'\" width=\"1\" height=\"1\" fill=\"#ffb300\"/>';});return r;})();\n\n// ---- разделы и кнопки: [название, функция, цвет, иконка, подпись]\nvar SECTIONS = [\n  {id:'s1', title:'1 СБОР', open:true, items:[\n    ['Статистика за сегодня','sobratStatistiku','green','chk'],\n    ['Выпулить крупные','vypulitKrupnye','green','dl','сумма от порога, лист «1800+»'],\n    ['Выпулить товары в «пул»','vypulitTovary','green','dl','товары из «Статистики»']\n  ]},\n  {id:'s2', title:'2 ПРОВЕРКА ИИ', open:true, items:[\n    ['Haiku 5.5 - проверить пул','proveritPulHaiku','blue','play','дёшево, сам разбирает ⚠'],\n    ['Sonnet 5.5 - проверить пул','proveritPulSonnet','blue','play','точнее, для остатка'],\n    ['Как идёт проверка','statusProverkiIi','cyan','list'],\n    ['Отсортировать пул','sortirovatPulKnopka','cyan','sort','✓ ≈ ⚠'],\n    ['Остановить проверку','ostanovitProverkuIi','red','stop']\n  ]},\n  {id:'s3', title:'3 ЛИДЫ НП', open:false, items:[\n    ['Перенести в «НП лиды»','perenestiVNpLidy','yellow','send','выделенные строки пула'],\n    ['Выпулить для Скайлаба','vypulitDlyaSkailab','yellow','box'],\n    ['Выпулить для Щербака','vypulitDlyaShcherbak','yellow','box'],\n    ['Создать ТТН','sozdatTtn','yellow','play','лиды без ТТН'],\n    ['Печать наклеек 100×100','pechatNakleek','yellow','print'],\n    ['Обновить статистику лидов','obnovitStatistikuNpLidov','cyan','list']\n  ]},\n  {id:'s4', title:'4 CRM', open:false, items:[\n    ['В «Заказы на СРМ»','perenestiVZakazyNaCrm','violet','send','для робота LP-CRM, с ТТН'],\n    ['Выпулить заказы в СРМ 1','otpravitVCrm1','violet','send'],\n    ['Выпулить заказы в СРМ 2','otpravitVCrm2','violet','send']\n  ]},\n  {id:'s5', title:'5 ДРОП И ТТН', open:false, items:[\n    ['Выгрузить дропу (TXT)','vygruzitDlyaDropa','red','dl','выделенные строки пула'],\n    ['Выпулить ТТН: «ОПТ ДРОП»','vypulitTtnOptDrop','red','list'],\n    ['Выпулить ТТН: «АВЕ ОПТ»','vypulitTtnAveOpt','red','list'],\n    ['Удалить ТТН','udalitTtn','red','trash','лист «ТТН на удаление»']\n  ]},\n  {id:'s6', title:'6 РАСШИРЕНИЯ', open:false, items:[\n    ['Установить робот LP-CRM','ustanovitRasshirenieCrm','cyan','dl','расширение Chrome'],\n    ['Обновить скрипт таблицы','ustanovitSkriptTablitsy','cyan','dl','проверка версии и ссылка'],\n    ['Ссылка и код для робота','ssylkaDlyaRasshireniya','gray','key']\n  ]},\n  {id:'s7', title:'7 КЛЮЧИ', open:false, items:[\n    ['Ключ СРМ 1','ukazatKlyuchCrm1','gray','key'],\n    ['Выходной ключ СРМ 1','ukazatVyhodnoiKlyuchCrm1','gray','key','чтение заказов'],\n    ['Адрес и ключ СРМ 2','ukazatKlyuchCrm2','gray','key'],\n    ['Ключ Новой Почты','ukazatKlyuchNp','gray','key'],\n    ['Ключ НП из LP-CRM','ukazatKlyuchNpCrm','gray','key','удаление ТТН'],\n    ['Ключ OpenRouter','ukazatKlyuchOpenRouter','gray','key']\n  ]}\n];\n\nvar app=document.getElementById('app');\nSECTIONS.forEach(function(s){\n  var d=document.createElement('div'); d.className='sec '+s.id+(s.open?'':' closed');\n  d.innerHTML='<div class=\"hd pix\"><i class=\"ar\"></i>'+s.title+'<span class=\"n\">'+s.items.length+'</span></div><div class=\"list\"></div>';\n  var list=d.querySelector('.list');\n  s.items.forEach(function(it){\n    var b=document.createElement('button'); b.className='btn c-'+it[2]; b.type='button';\n    b.innerHTML=icon(it[3])+'<span class=\"t\">'+it[0]+(it[4]?'<small>'+it[4]+'</small>':'')+'</span>';\n    b.onclick=function(){run(b,it[1],it[0]);};\n    list.appendChild(b);\n  });\n  d.querySelector('.hd').onclick=function(){d.classList.toggle('closed'); try{localStorage.setItem('zp_'+s.id,d.classList.contains('closed')?'0':'1');}catch(e){}};\n  try{var v=localStorage.getItem('zp_'+s.id); if(v==='1')d.classList.remove('closed'); else if(v==='0')d.classList.add('closed');}catch(e){}\n  app.appendChild(d);\n});\n\nvar logEl=document.getElementById('log');\nfunction log(msg,cls){logEl.hidden=false; var t=new Date(), p=function(n){return (n<10?'0':'')+n;};\n  logEl.insertAdjacentHTML('afterbegin','<div class=\"'+(cls||'')+'\">'+p(t.getHours())+':'+p(t.getMinutes())+' '+msg+'</div>');}\nfunction run(btn,fn,name){\n  if(btn.classList.contains('busy'))return;\n  btn.classList.add('busy');\n  google.script.run\n    .withSuccessHandler(function(){btn.classList.remove('busy'); log(name+' - готово','o'); if(window.pet)pet.hop(); refresh();})\n    .withFailureHandler(function(e){btn.classList.remove('busy'); log(name+': '+(e&&e.message||e),'e'); if(window.pet)pet.sad();})\n    [fn]();\n}\n\nvar prevIdet=null;\nfunction refresh(){\n  google.script.run.withSuccessHandler(function(d){\n    if(!d)return;\n    if(window.pet){ if(prevIdet===true&&!d.idet){ if(d.left===0)pet.happy(); else pet.sad(); } pet.day(d.day||0); } prevIdet=d.idet;\n    document.getElementById('st-day').textContent=d.day||0;\n    document.getElementById('led').className='led'+(d.idet?' on':'');\n    document.getElementById('st-state').textContent=d.idet?d.model:'СТОП';\n    document.getElementById('st-left').textContent=d.left;\n    document.getElementById('st-cost').textContent='$'+d.cost.toFixed(2);\n    document.getElementById('st-last').textContent=d.last||'-';\n    document.getElementById('ver').textContent='v'+d.version;\n    var bar=document.getElementById('bar'),N=20,total=Math.max(d.total||0,d.left||0),done=total?Math.round((1-d.left/total)*N):0;\n    bar.innerHTML='';for(var i=0;i<N;i++)bar.insertAdjacentHTML('beforeend','<i class=\"'+(i<done?'f':'')+'\"></i>');\n  }).dannyePaneli();\n}\nrefresh(); setInterval(refresh,15000);\n\n// ---- кот: ходит по низу панели, садится «булкой», спит, идёт за курсором, реагирует на события, растёт с уровнем\n(function(){\n  var BASE=[\n    '............O...O.',\n    '...........OFO.OFO',\n    'O..........OFFOFFO',\n    'OO.........OFFFFFO',\n    '.O.........OEFFEFO',\n    '.OO........OFFPFFO',\n    '..OOOOOOOOOOCCCCO.',\n    '..OFFSFFSFFFFFFO..',\n    '..OFFFFFFFFFFFO...',\n    '..OFFFFFFFFFFO....'];\n  var LEGS={a:['..OFOFO..OFOFO....','..OO.OO..OO.OO....'],b:['...OFOFOOFOFO.....','...OO.OOOO.OO.....'],loaf:['..OFFFFFFFFFFO....','..OOOOOOOOOOOO....']};\n  var TAIL_UP=['O..','O..','OO.','.O.','.O.','.OO'];           // хвост трубой (строки 0-5)\n  var el=document.createElement('div'); el.id='slime';\n  var PAL={O:'#2a2d43',F:'#ffa94d',S:'#e8842a',P:'#ff8fa3',W:'#fff'};\n  var COLLAR={idle:'#5fd49a',work:'#6cb8ff',busy:'#ffd24a'};\n  var mood='idle',feel='',feelT=0,size=1,legs='loaf',tailUp=false,eyes='open',key='';\n  function paint(){\n    var k=[mood,feel,legs,tailUp,eyes].join('|'); if(k===key)return; key=k;\n    var rows=BASE.concat(LEGS[legs]).map(function(r,y){return (tailUp&&y<6)?TAIL_UP[y]+r.slice(3):((tailUp&&y===0)?r:r);});\n    if(tailUp)rows=rows.map(function(r,y){return y<6?TAIL_UP[y]+r.slice(3):r;});\n    var fur=feel==='sad'?'#c9cdd8':PAL.F, spot=feel==='sad'?'#9aa0b5':PAL.S, s='';\n    rows.forEach(function(r,y){for(var x=0;x<18;x++){var c=r[x],col=null;\n      if(c==='O')col=PAL.O;else if(c==='F')col=fur;else if(c==='S')col=spot;else if(c==='P')col=PAL.P;else if(c==='C')col=COLLAR[mood];\n      else if(c==='E')col=eyes==='open'?PAL.O:eyes==='angry'?'#e03131':spot;\n      if(col)s+='<rect x=\"'+x+'\" y=\"'+y+'\" width=\"1\" height=\"1\" fill=\"'+col+'\"/>';}});\n    if(eyes==='open')s+='<rect x=\"12\" y=\"4\" width=\".5\" height=\".5\" fill=\"#fff\"/><rect x=\"15\" y=\"4\" width=\".5\" height=\".5\" fill=\"#fff\"/>';\n    el.innerHTML='<svg viewBox=\"0 0 18 12\">'+s+'</svg>';\n  }\n  var W=54,H=36,x=30,y=-60,vy=0,dir=1,on=false,t=0,state='sit',stT=120,target=0,lastAct=0,mouseX=-1,mouseT=-9999;\n  function fx(ch){var s=document.createElement('div');s.className='fx';s.textContent=ch;s.style.left=(x+W*size/2-6+(Math.random()*24-12))+'px';s.style.top=(y-10)+'px';document.body.appendChild(s);setTimeout(function(){s.remove();},1000);}\n  // имя и уровень: растёт от числа проверенных за сегодня заказов (каждые 25 - уровень, до 10-го)\n  var nameEl=document.getElementById('pet-name'),lvlEl=document.getElementById('pet-lvl'),lvl=-1;\n  function petName(){try{return localStorage.getItem('zp_cat')||'Мурзик';}catch(e){return 'Мурзик';}}\n  nameEl.textContent=petName();\n  nameEl.onclick=function(){\n    if(nameEl.querySelector('input'))return;\n    var inp=document.createElement('input');inp.value=petName();inp.maxLength=14;nameEl.textContent='';nameEl.appendChild(inp);inp.focus();inp.select();\n    function done(){var v=inp.value.trim()||'Мурзик';try{localStorage.setItem('zp_cat',v);}catch(e){}nameEl.textContent=v;el.title=v;}\n    inp.onblur=done;inp.onkeydown=function(e){if(e.key==='Enter')inp.blur();};\n  };\n  function go(s,dur){state=s;stT=dur||0;}\n  window.pet={\n    happy:function(){feel='happy';feelT=420;lastAct=t;go('sit',9999);},\n    sad:function(){feel='sad';feelT=360;lastAct=t;go('sit',9999);fx('ш-ш!');},\n    hop:function(){if(!feel&&on){vy=-8;on=false;lastAct=t;if(state==='sleep')go('sit',80);}},\n    day:function(n){var L=Math.min(10,Math.floor(n/25));lvlEl.textContent='ур. '+L+(L<10?' · до след. '+(25-n%25):' · макс.');\n      if(L!==lvl){if(lvl>=0&&L>lvl){window.pet.happy();fx('★');}lvl=L;size=1+L*.07;}}\n  };\n  paint(); document.body.appendChild(el); el.title=petName();\n  el.onclick=function(){lastAct=t;if(on){vy=-11;on=false;}fx('мяу');if(state==='sleep')go('sit',100);};\n  document.addEventListener('mousemove',function(e){mouseX=e.clientX;mouseT=t;});\n  function step(){\n    t++;\n    var busy=!!document.querySelector('.btn.busy'), work=document.getElementById('led').classList.contains('on');\n    mood=busy?'busy':work?'work':'idle';\n    var maxX=window.innerWidth-W*size-4, fl=window.innerHeight-H*size-4;\n    if(feelT>0&&--feelT===0){feel='';go('sit',60);}\n    vy+=0.6; y+=vy;\n    if(y>=fl){y=fl;vy=0;on=true;}\n    var jx=0;\n    if(feel==='happy'){                                     // радость: хвост трубой, подпрыгивает, сердечки\n      legs='loaf';eyes='closed';tailUp=true;\n      if(on&&t%30===0){vy=-(7+Math.random()*3);on=false;fx(Math.random()<.5?'♥':'★');}\n    }else if(feel==='sad'){                                  // ошибка: шипит, шерсть серая, дрожит\n      legs='a';eyes='angry';tailUp=true;jx=(t%4<2?1:-1);\n    }else if(state==='walk'){\n      var sp=(mood==='idle'?1.1:1.8);\n      if(Math.abs(target-x)<=sp){go('sit',80+Math.random()*(mood==='idle'?260:90));}\n      else{dir=target>x?1:-1;x+=dir*sp;}\n      legs=(Math.floor(t/7)%2)?'a':'b';eyes=(t%190<8)?'closed':'open';tailUp=Math.floor(t/20)%2===0;\n    }else if(state==='sleep'){\n      legs='loaf';eyes='closed';tailUp=false;\n      if(t%80===0)fx('z');\n      if(mood!=='idle')go('sit',40);\n    }else{                                                   // сидит «булкой»: моргает, иногда машет хвостом\n      legs='loaf';eyes=(t%160<8)?'closed':'open';tailUp=(t%240)<30&&Math.floor(t/8)%2===0;\n      if(--stT<=0){\n        if(mood==='idle'&&t-lastAct>2400&&t-mouseT>1200)go('sleep');       // ~40 секунд тишины - засыпает\n        else{\n          var chase=mouseX>=0&&t-mouseT<240&&Math.random()<.6;             // идёт за курсором\n          target=Math.max(4,Math.min(maxX,chase?mouseX-W*size/2:4+Math.random()*maxX));go('walk');\n          if(Math.random()<.12&&on){vy=-9;on=false;}\n        }\n      }\n    }\n    if(x>maxX)x=maxX; if(x<4)x=4;\n    paint();\n    el.style.transform='translate('+Math.round(x+jx+(dir<0?W*size:0))+'px,'+Math.round(y)+'px) scale('+(dir*size).toFixed(2)+','+size.toFixed(2)+')';\n    requestAnimationFrame(step);\n  }\n  requestAnimationFrame(step);\n})();\n</script>\n</body></html>\n";
/*PANEL_END*/

/** Имя листа за сегодня, например «07.10.2026». */
function imyaListaSegodnya_() {
  return Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM.yyyy');
}

/** Последний по дате лист с именем вида «07.10.2026» или null. */
function posledniyListDnya_() {
  var best = null, bestKey = '';
  SpreadsheetApp.getActive().getSheets().forEach(function (sh) {
    var m = sh.getName().match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (!m) return;
    var key = m[3] + m[2] + m[1];
    if (key > bestKey) { bestKey = key; best = sh; }
  });
  return best;
}

/**
 * Лист дня, с которым работают кнопки.
 * Статистика: лист за сегодня, а если его ещё нет (например, после полуночи) - последний обработанный день.
 * Пул и 1800+ (poStatistike = true): тот день, за который собрана статистика (ячейка F1 листа «Статистика»).
 */
function listDnya_(poStatistike) {
  var ss = SpreadsheetApp.getActive();
  var sheet = null;
  if (poStatistike) {
    var stats = ss.getSheetByName(ZAYAVKI_CFG.STATS_SHEET);
    var day = stats ? String(stats.getRange(1, 6).getDisplayValue()).trim() : '';
    if (day) sheet = ss.getSheetByName(day);
  }
  if (!sheet) sheet = ss.getSheetByName(imyaListaSegodnya_());
  if (!sheet) sheet = posledniyListDnya_();
  if (!sheet) {
    SpreadsheetApp.getUi().alert('В таблице нет листов с днями (например «07.10.2026»). Сначала запустите парсер.');
    return null;
  }
  return sheet;
}

function chislo_(value) {
  if (typeof value === 'number') return value;
  var n = parseFloat(String(value).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

/**
 * Разбирает ячейку «Товар» на позиции.
 * Понимает оба формата парсера: «товар (2 шт); товар (1 шт)» и старый «товар, товар».
 */
function razobratTovary_(cell) {
  var text = String(cell || '').trim();
  if (!text) return [];
  var parts = text.indexOf(';') >= 0 ? text.split(';') : text.split(',');
  var items = [];
  parts.forEach(function (part) {
    var name = part.trim();
    if (!name) return;
    var qty = 1;
    var m = name.match(/\(\s*(\d+)\s*шт\.?\s*\)\s*$/i);
    if (m) {
      qty = parseInt(m[1], 10) || 1;
      name = name.slice(0, m.index).trim();
    }
    if (name) items.push({ name: name, qty: qty });
  });
  return items;
}

/**
 * Товары строки листа дня: {list: [...], other: [...]}.
 * Берутся из колонок «Товар по списку» и «Нет в списке»; у старых строк без них - из колонки «Товар».
 */
function tovaryStroki_(row) {
  var list = razobratTovary_(row[ZAYAVKI_CFG.COL_LIST - 1]);
  var other = razobratTovary_(row[ZAYAVKI_CFG.COL_OTHER - 1]);
  if (!list.length && !other.length) other = razobratTovary_(row[ZAYAVKI_CFG.COL_PRODUCT - 1]);
  return { list: list, other: other };
}

/** Кнопка «Собрать статистику за сегодня». */
function sobratStatistiku() {
  var sheet = listDnya_(false);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var width = Math.min(ZAYAVKI_CFG.READ_WIDTH, sheet.getLastColumn());
  var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];

  var statsList = {}, statsOther = {};   // ключ -> {name, orders, qty}
  var orders = 0, totalSum = 0;
  function dobavit(stats, items) {
    var seen = {};
    items.forEach(function (item) {
      var key = item.name.toLowerCase().replace(/\s+/g, ' ');
      if (!stats[key]) stats[key] = { name: item.name, orders: 0, qty: 0 };
      if (!seen[key]) { stats[key].orders++; seen[key] = true; }  // один заказ считается один раз
      stats[key].qty += item.qty;
    });
  }
  var uzhe = {};   // если несколько серверов случайно записали один звонок дважды - считаем один раз
  rows.forEach(function (row) {
    if (ZAYAVKI_CFG.ORDER_STATUSES.indexOf(String(row[ZAYAVKI_CFG.COL_STATUS - 1])) < 0) return;
    var cid = String(row[ZAYAVKI_CFG.COL_ID - 1]);
    if (uzhe[cid]) return;
    uzhe[cid] = true;
    orders++;
    totalSum += chislo_(row[ZAYAVKI_CFG.COL_SUM - 1]);
    var t = tovaryStroki_(row);
    dobavit(statsList, t.list);
    dobavit(statsOther, t.other);
  });

  function vSpisok(stats) {
    var arr = Object.keys(stats).map(function (k) { return [stats[k].name, stats[k].orders, stats[k].qty]; });
    arr.sort(function (a, b) { return b[1] - a[1] || b[2] - a[2]; });
    return arr;
  }
  var listPart = vSpisok(statsList), otherPart = vSpisok(statsOther);
  var list = listPart.slice();
  var razdelitelRow = 0;
  if (otherPart.length) {
    list.push(['', '', '']);
    list.push(['——— Нет в списке товаров ———', '', '']);
    razdelitelRow = list.length + 1;   // номер строки на листе (с учётом заголовка)
    list = list.concat(otherPart);
  }

  var ss = SpreadsheetApp.getActive();
  var out = ss.getSheetByName(ZAYAVKI_CFG.STATS_SHEET) || ss.insertSheet(ZAYAVKI_CFG.STATS_SHEET);
  out.getRange('A:F').clearContent();   // колонки H и I (список для пула) не трогаем
  out.getRange(1, 1, 1, 3).setValues([['Товар', 'Заказов', 'Штук']]).setFontWeight('bold');
  shapkaPula_(out);
  out.getRange('A:C').setFontWeight('normal');
  out.getRange(1, 1, 1, 3).setFontWeight('bold');
  if (list.length) out.getRange(2, 1, list.length, 3).setValues(list);
  if (razdelitelRow) out.getRange(razdelitelRow, 1).setFontWeight('bold');
  out.getRange(1, 6).setNumberFormat('@');
  out.getRange(1, 5, 4, 2).setValues([
    ['День', sheet.getName()],
    ['Заказов', orders],
    ['Общая сумма', Math.round(totalSum)],
    ['Средний чек', orders ? Math.round(totalSum / orders) : 0]
  ]);
  out.getRange(1, 5, 4, 1).setFontWeight('bold');
  out.setFrozenRows(1);
  ss.toast('День: ' + sheet.getName() + '. Товаров из списка: ' + listPart.length + ', не из списка: ' + otherPart.length +
           ', заказов: ' + orders + ', сумма: ' + Math.round(totalSum), 'Статистика собрана', 8);
}

/** Кнопка «Выпулить 1800+»: копирует крупные заказы за сегодня в лист «1800+» без дублей. */
function vypulitKrupnye() {
  var sheet = listDnya_(true);
  if (!sheet) return;
  var ss = SpreadsheetApp.getActive();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('В листе «' + sheet.getName() + '» нет строк.', ZAYAVKI_CFG.BIG_SHEET, 6); return; }

  var out = ss.getSheetByName(ZAYAVKI_CFG.BIG_SHEET) || ss.insertSheet(ZAYAVKI_CFG.BIG_SHEET);
  if (out.getMaxColumns() < ZAYAVKI_CFG.WIDTH) out.insertColumnsAfter(out.getMaxColumns(), ZAYAVKI_CFG.WIDTH - out.getMaxColumns());
  if (out.getLastRow() === 0) {
    out.getRange(1, 1, 1, ZAYAVKI_CFG.WIDTH).setValues(sheet.getRange(1, 1, 1, ZAYAVKI_CFG.WIDTH).getValues()).setFontWeight('bold');
    out.setFrozenRows(1);
  }

  var existing = {};
  if (out.getLastRow() > 1) {
    out.getRange(2, ZAYAVKI_CFG.COL_ID, out.getLastRow() - 1, 1).getValues().forEach(function (r) {
      existing[String(r[0])] = true;
    });
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, ZAYAVKI_CFG.WIDTH).getValues();
  var toAdd = [];
  rows.forEach(function (row) {
    var id = String(row[ZAYAVKI_CFG.COL_ID - 1]);
    if (!id || existing[id]) return;
    if (ZAYAVKI_CFG.ORDER_STATUSES.indexOf(String(row[ZAYAVKI_CFG.COL_STATUS - 1])) < 0) return;
    if (chislo_(row[ZAYAVKI_CFG.COL_SUM - 1]) < ZAYAVKI_CFG.BIG_SUM) return;
    existing[id] = true;
    toAdd.push(row);
  });

  if (!toAdd.length) { ss.toast('Новых заказов от ' + ZAYAVKI_CFG.BIG_SUM + ' грн нет.', ZAYAVKI_CFG.BIG_SHEET, 6); return; }

  var start = out.getLastRow() + 1;
  // Call ID и телефон храним как текст, иначе у телефона пропадёт первый ноль.
  out.getRange(start, ZAYAVKI_CFG.COL_ID, toAdd.length, 1).setNumberFormat('@');
  out.getRange(start, ZAYAVKI_CFG.COL_PHONE, toAdd.length, 1).setNumberFormat('@');
  out.getRange(start, 1, toAdd.length, ZAYAVKI_CFG.WIDTH).setValues(toAdd);
  ss.toast('День: ' + sheet.getName() + '. Добавлено заказов: ' + toAdd.length, ZAYAVKI_CFG.BIG_SHEET, 8);
}

/* ===================== Пул товаров и отправка в LP-CRM ===================== */

function klyuchTovara_(name) {
  return String(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Заголовки колонок H, I, J в листе «Статистика», если их ещё нет. */
function shapkaPula_(stats) {
  var c = ZAYAVKI_CFG.COL_PICK;
  if (stats.getMaxColumns() < c + 3) stats.insertColumnsAfter(stats.getMaxColumns(), c + 3 - stats.getMaxColumns());
  var head = stats.getRange(1, c, 1, 4);
  var v = head.getValues()[0];
  if (!v[0] && !v[1]) v[0] = 'Выпулить товары';
  if (!v[1] || v[1] === 'ID товара в LP-CRM') v[1] = 'ID товара в СРМ 1';
  if (!v[2]) v[2] = 'ID товара в СРМ 2';
  if (!v[3]) v[3] = 'Артикул дропа';
  head.setValues([v]).setFontWeight('bold');
  stats.getRange(2, c + 3, stats.getMaxRows() - 1, 1).setNumberFormat('@');   // артикул как текст: «0428» не теряет ноль
}

/** Список из колонок H:J листа «Статистика»: ключ товара -> [ID в СРМ 1, ID в СРМ 2]. */
function spisokPula_() {
  var stats = SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.STATS_SHEET);
  var picks = {}, count = 0;
  if (!stats || stats.getLastRow() < 2) return { picks: picks, count: 0 };
  shapkaPula_(stats);
  stats.getRange(2, ZAYAVKI_CFG.COL_PICK, stats.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var key = klyuchTovara_(r[0]);
    if (!key) return;
    if (!(key in picks)) count++;
    picks[key] = [String(r[1] || '').trim(), String(r[2] || '').trim()];
  });
  return { picks: picks, count: count };
}

/* Колонки листа «пул». Город и отделение склеиваются в «Адрес НП» в формате Новой Почты:
   «Кривий Ріг, Відділення №27 (до 30 кг): вул. ...» - этот адрес уходит в поле адреса заказа в CRM. */
// «Проверка ИИ» и «Допы» заполняет кнопка «Проверить пул ИИ». Call ID нужен для защиты от дублей
// (в пуле и в CRM), поэтому он остаётся, но стоит в самом конце.
var POOL_COLUMNS = ['Проверка ИИ', 'сек.', 'ФИО', 'Телефон', 'Товар', 'Шт', 'Сумма', 'Сумма/шт ИИ',
                    'Адрес НП', 'Допы', 'Транскрипт', 'Статус CRM', 'ID товара СРМ 1', 'ID товара СРМ 2', 'Call ID',
                    'Область, район', 'Партия'];
// колонки пула, которые копируются из листа дня: название в пуле -> название в листе дня
var POOL_FROM_DAY = { 'Call ID': 'Call ID', 'сек.': 'Длительность, сек', 'ФИО': 'ФИО', 'Телефон': 'Телефон',
                      'Товар': 'Товар', 'Сумма': 'Сумма', 'Транскрипт': 'Транскрипт' };
// названия из прежних версий листа «пул»
var POOL_ALIASES = { 'ID товара': 'ID товара СРМ 1', 'Длительность, сек': 'сек.' };
var POOL_WIDTHS = { 'сек.': 45, 'Шт': 40, 'Сумма': 60, 'Сумма/шт ИИ': 220 };   // узкие колонки (ширина в пикселях)

/** Ширина узких колонок пула. */
function shirinaKolonokPula_(out) {
  Object.keys(POOL_WIDTHS).forEach(function (name) { out.setColumnWidth(P_(name) + 1, POOL_WIDTHS[name]); });
}

/** Номер колонки пула (с нуля) по названию. */
function P_(name) { return POOL_COLUMNS.indexOf(name); }

/** Колонки «Call ID» и «Телефон» как текст, иначе у телефона пропадёт первый ноль. */
function tekstovyeKolonkiPula_(out, start, count) {
  out.getRange(start, P_('Call ID') + 1, count, 1).setNumberFormat('@');
  out.getRange(start, P_('Телефон') + 1, count, 1).setNumberFormat('@');
}

/** Сколько слов в ФИО (0 - пусто, 1 - только имя или только фамилия). */
function slovVFio_(fio) {
  var t = String(fio || '').replace(/[^A-Za-zА-Яа-яІіЇїЄєҐґЁё'’ʼ-]+/g, ' ').trim();
  return t ? t.split(/\s+/).length : 0;
}

/** Ссылка на Телеграм по номеру: https://t.me/+380XXXXXXXXX (пусто, если номер не украинский мобильный). */
function ssylkaTg_(phone) {
  var d = String(phone || '').replace(/\D/g, '');
  if (d.length === 10 && d[0] === '0') d = '38' + d;
  else if (d.length === 9) d = '380' + d;
  return /^380\d{9}$/.test(d) ? 'https://t.me/+' + d : '';
}

/**
 * Где в «ФИО» пусто или одно слово, телефон становится ссылкой на Телеграм (значение ячейки не меняется).
 * У остальных строк ссылка снимается. Вызывается после выпула, после проверки ИИ и после сортировки.
 */
function ssylkiTg_(pool) {
  try {
    var n = pool.getLastRow() - 1;
    if (n < 1) return;
    var fio = pool.getRange(2, P_('ФИО') + 1, n, 1).getValues();
    var rng = pool.getRange(2, P_('Телефон') + 1, n, 1);
    var tel = rng.getDisplayValues();
    var rich = tel.map(function (t, i) {
      var txt = String(t[0] || ''), b = SpreadsheetApp.newRichTextValue().setText(txt);
      var url = txt && slovVFio_(fio[i][0]) < 2 ? ssylkaTg_(txt) : '';
      if (txt) b.setLinkUrl(url || null);
      return [b.build()];
    });
    rng.setNumberFormat('@').setRichTextValues(rich);
  } catch (e) { console.log('ssylkiTg_: ' + e); }
}

/** Ровные строки: текст обрезается, высота строки 21. */
function rovnyeStroki_(out, start, count) {
  out.getRange(start, 1, count, POOL_COLUMNS.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  out.setRowHeightsForced(start, count, 21);
}

function vOdnuStroku_(v) {
  return typeof v === 'string' ? v.replace(/\s*\n\s*/g, ' | ') : v;
}

/**
 * Готовит лист «пул»: заголовки по POOL_COLUMNS. Лист прежнего вида перестраивается:
 * данные переносятся по названиям колонок, статусы CRM и ID сохраняются, адрес собирается из города и отделения.
 */
function podgotovitPul_(out) {
  var w = POOL_COLUMNS.length;
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('POOL_NOTES_CLEARED')) {   // старые примечания удаляем один раз
    out.getRange(1, 1, out.getMaxRows(), w).clearNote();
    props.setProperty('POOL_NOTES_CLEARED', '1');
  }
  if (out.getMaxColumns() < w) out.insertColumnsAfter(out.getMaxColumns(), w - out.getMaxColumns());
  if (out.getLastRow() === 0) {
    out.getRange(1, 1, 1, w).setValues([POOL_COLUMNS]).setFontWeight('bold');
    out.setFrozenRows(1);
    shirinaKolonokPula_(out);
    return;
  }
  var lastCol = Math.max(out.getLastColumn(), 1);
  var head = out.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head.slice(0, w).join('|') === POOL_COLUMNS.join('|')) return;

  var idx = {};
  head.forEach(function (h, i) { var name = POOL_ALIASES[h] || h; if (name && !(name in idx)) idx[name] = i; });
  var data = out.getLastRow() > 1 ? out.getRange(2, 1, out.getLastRow() - 1, lastCol).getValues() : [];
  data = data.filter(function (r) { return String(r[idx['Call ID'] !== undefined ? idx['Call ID'] : 0]).trim(); });
  var adresa = !('Адрес НП' in idx) ? adresaNp_(data.map(function (r) {
    return { city: 'Город' in idx ? r[idx['Город']] : '', np: 'Отделение НП' in idx ? r[idx['Отделение НП']] : '' };
  })) : null;
  var rows = data.map(function (r, n) {
    return POOL_COLUMNS.map(function (name) {
      if (name === 'Адрес НП' && adresa) return adresa[n];
      return name in idx ? vOdnuStroku_(r[idx[name]]) : '';
    });
  });
  out.clear();
  if (out.getMaxColumns() > w) out.deleteColumns(w + 1, out.getMaxColumns() - w);
  out.getRange(1, 1, 1, w).setValues([POOL_COLUMNS]).setFontWeight('bold');
  out.setFrozenRows(1);
  shirinaKolonokPula_(out);
  if (rows.length) {
    tekstovyeKolonkiPula_(out, 2, rows.length);
    out.getRange(2, 1, rows.length, w).setValues(rows);
    rovnyeStroki_(out, 2, rows.length);
    ssylkiTg_(out);
  }
  SpreadsheetApp.getActive().toast('Лист «' + ZAYAVKI_CFG.POOL_SHEET + '» перестроен под новые колонки, строк: ' + rows.length,
                                   ZAYAVKI_CFG.POOL_SHEET, 8);
}

/** Кнопка «Выпулить товары»: заказы за день с товарами из колонки H уходят в лист «пул». */
function vypulitTovary() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var list = spisokPula_();
  if (!list.count) {
    var st = ss.getSheetByName(ZAYAVKI_CFG.STATS_SHEET);
    if (st) shapkaPula_(st);
    ui.alert('В листе «' + ZAYAVKI_CFG.STATS_SHEET + '» в колонке H нет товаров. ' +
             'Скопируйте туда названия из колонки A (со строки 2) и нажмите кнопку ещё раз.');
    return;
  }
  var sheet = listDnya_(true);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('В листе «' + sheet.getName() + '» нет строк.', ZAYAVKI_CFG.POOL_SHEET, 6); return; }

  var out = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET) || ss.insertSheet(ZAYAVKI_CFG.POOL_SHEET);
  podgotovitPul_(out);
  var existing = {};
  if (out.getLastRow() > 1) {
    out.getRange(2, P_('Call ID') + 1, out.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[String(r[0])] = true; });
  }

  // Колонки листа дня по названиям: так пул не зависит от их порядка.
  var width = sheet.getLastColumn();
  var dayHead = sheet.getRange(1, 1, 1, width).getValues()[0];
  var dayIdx = {};
  dayHead.forEach(function (h, i) { h = String(h).trim(); if (h && !(h in dayIdx)) dayIdx[h] = i; });
  function iz(full, name) { return name in dayIdx ? full[dayIdx[name]] : ''; }

  var toAdd = [], mesta = [];
  var partiya = Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM HH:mm:ss');
  sheet.getRange(2, 1, lastRow - 1, width).getValues().forEach(function (full) {
    var id = String(full[ZAYAVKI_CFG.COL_ID - 1]);
    if (!id || existing[id]) return;
    if (ZAYAVKI_CFG.ORDER_STATUSES.indexOf(String(full[ZAYAVKI_CFG.COL_STATUS - 1])) < 0) return;
    var t = tovaryStroki_(full);
    var items = t.list.concat(t.other, razobratTovary_(full[ZAYAVKI_CFG.COL_PRODUCT - 1]));
    var hit = items.some(function (item) { return klyuchTovara_(item.name) in list.picks; });
    if (!hit) return;
    // ID товара для обеих CRM запоминаем сразу: потом список в колонках H:J можно менять, на отправку это не повлияет.
    var ids = ['', ''];
    items.forEach(function (item) {
      var p = list.picks[klyuchTovara_(item.name)];
      if (!p) return;
      if (!ids[0]) ids[0] = p[0];
      if (!ids[1]) ids[1] = p[1];
    });
    var row = POOL_COLUMNS.map(function (name) {
      if (name === 'Партия') return partiya;   // по ней строки сортируются после проверки ИИ - внутри своего выпула
      if (name === 'Шт') { var it0 = razobratTovary_(iz(full, 'Товар')); return it0.length ? it0[0].qty : 1; }   // уточнит проверка ИИ
      if (name === 'ID товара СРМ 1') return ids[0];
      if (name === 'ID товара СРМ 2') return ids[1];
      return name in POOL_FROM_DAY ? vOdnuStroku_(iz(full, POOL_FROM_DAY[name])) : '';
    });
    existing[id] = true;
    toAdd.push(row);
    // адрес, который уже нашёл сервер (колонка «Адрес НП» листа дня); если его нет - ищем здесь
    mesta.push({ city: iz(full, 'Город'), np: iz(full, 'Отделение НП'), gotov: String(iz(full, 'Адрес НП') || '').trim() });
  });
  if (!toAdd.length) { ss.toast('День: ' + sheet.getName() + '. Новых заказов с этими товарами нет.', ZAYAVKI_CFG.POOL_SHEET, 6); return; }

  var iskat = [], kudaIskat = [];
  mesta.forEach(function (m, n) { if (!m.gotov) { iskat.push(m); kudaIskat.push(n); } });
  if (iskat.length) ss.toast('Ищу адреса отделений в Новой Почте: ' + iskat.length + '...', ZAYAVKI_CFG.POOL_SHEET, 30);
  var naideno = iskat.length ? adresaNp_(iskat) : [];
  var adresa = mesta.map(function (m) { return m.gotov; });
  kudaIskat.forEach(function (n, k) { adresa[n] = naideno[k]; });
  var neNaideno = 0;
  toAdd.forEach(function (row, n) {
    row[P_('Адрес НП')] = adresa[n];
    if (String(adresa[n]).indexOf('⚠') === 0) neNaideno++;
  });

  var start = out.getLastRow() + 1;
  tekstovyeKolonkiPula_(out, start, toAdd.length);
  out.getRange(start, 1, toAdd.length, POOL_COLUMNS.length).setValues(toAdd);
  rovnyeStroki_(out, start, toAdd.length);
  ssylkiTg_(out);
  ss.toast('День: ' + sheet.getName() + '. Добавлено заказов: ' + toAdd.length +
           (neNaideno ? '. Адрес не найден в НП: ' + neNaideno + ' (отмечены ⚠, поправьте перед отправкой)' : ''),
           ZAYAVKI_CFG.POOL_SHEET, 10);
}

/* ===================== Адрес отделения через API Новой Почты ===================== */

var NP_URL = 'https://api.novaposhta.ua/v2.0/json/';

/** Кнопка «Указать ключ Новой Почты»: ключ из кабинета НП (Налаштування -> Безпека -> API-ключ). */
function ukazatKlyuchNp() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Ключ Новой Почты', 'Вставьте API-ключ из бизнес-кабинета Новой Почты:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('NP_KEY', key);
  SpreadsheetApp.getActive().toast('Ключ сохранён.', 'Новая Почта', 5);
}

/** Несколько запросов к API НП параллельно. Возвращает массив ответов (или {success:false, errors:[...]}) */
var NP_OSHIBKI = {};   // тексты ошибок API за текущий запуск (для подсказки)

/**
 * Запросы к API НП небольшими пачками с паузами: на десятки одновременных запросов НП отвечает ошибкой
 * «слишком много запросов», и часть адресов терялась. Неудачные запросы повторяются до 3 раз.
 */
function npZaprosy_(calls) {
  var key = PropertiesService.getScriptProperties().getProperty('NP_KEY') || '';
  var out = new Array(calls.length);
  var todo = calls.map(function (c, i) { return i; });
  for (var attempt = 0; attempt < 4 && todo.length; attempt++) {
    if (attempt) Utilities.sleep(1500 * attempt);
    var retry = [];
    for (var i = 0; i < todo.length; i += 8) {
      var ids = todo.slice(i, i + 8);
      var part = ids.map(function (k) {
        var c = calls[k];
        return { url: NP_URL, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                 payload: JSON.stringify({ apiKey: key, modelName: c.model, calledMethod: c.method, methodProperties: c.props }) };
      });
      var resps;
      try { resps = UrlFetchApp.fetchAll(part); } catch (e) { resps = part.map(function () { return null; }); }
      resps.forEach(function (r, j) {
        var data = null;
        try { data = r ? JSON.parse(r.getContentText()) : null; } catch (e) {}
        if (!data || data.success === false) {
          var err = data && data.errors && data.errors.length ? String(data.errors[0]) : 'нет ответа';
          NP_OSHIBKI[err] = (NP_OSHIBKI[err] || 0) + 1;
          retry.push(ids[j]);
        }
        out[ids[j]] = data || { success: false, data: [], errors: ['нет ответа'] };
      });
      Utilities.sleep(250);
    }
    todo = retry;
  }
  return out;
}

/** Русские названия крупных городов -> как они записаны в НП. */
var NP_GORODA_RU = {
  'киев': 'Київ', 'харьков': 'Харків', 'одесса': 'Одеса', 'днепр': 'Дніпро', 'днепропетровск': 'Дніпро',
  'запорожье': 'Запоріжжя', 'львов': 'Львів', 'кривой рог': 'Кривий Ріг', 'николаев': 'Миколаїв',
  'винница': 'Вінниця', 'херсон': 'Херсон', 'полтава': 'Полтава', 'чернигов': 'Чернігів', 'черкассы': 'Черкаси',
  'житомир': 'Житомир', 'сумы': 'Суми', 'хмельницкий': 'Хмельницький', 'черновцы': 'Чернівці', 'ровно': 'Рівне',
  'ривне': 'Рівне', 'кропивницкий': 'Кропивницький', 'кировоград': 'Кропивницький', 'ивано-франковск': 'Івано-Франківськ',
  'тернополь': 'Тернопіль', 'луцк': 'Луцьк', 'ужгород': 'Ужгород', 'белая церковь': 'Біла Церква',
  'каменское': "Кам'янське", 'днепродзержинск': "Кам'янське", 'кременчуг': 'Кременчук', 'никополь': 'Нікополь',
  'павлоград': 'Павлоград', 'мукачево': 'Мукачево', 'бровары': 'Бровари', 'ирпень': 'Ірпінь', 'буча': 'Буча',
  'борисполь': 'Бориспіль', 'фастов': 'Фастів', 'васильков': 'Васильків', 'измаил': 'Ізмаїл',
  'бердичев': 'Бердичів', 'умань': 'Умань', 'каменец-подольский': "Кам'янець-Подільський", 'конотоп': 'Конотоп',
  'шостка': 'Шостка', 'лозовая': 'Лозова', 'изюм': 'Ізюм', 'чугуев': 'Чугуїв', 'новомосковск': 'Самар',
  'южноукраинск': 'Южноукраїнськ', 'вознесенск': 'Вознесенськ', 'первомайск': 'Первомайськ', 'смела': 'Сміла',
  'александрия': 'Олександрія', 'мелитополь': 'Мелітополь', 'бердянск': 'Бердянськ', 'энергодар': 'Енергодар',
  'краматорск': 'Краматорськ', 'славянск': "Слов'янськ", 'покровск': 'Покровськ', 'днепре': 'Дніпро', 'киеве': 'Київ'
};

/**
 * Разбирает то, что записал парсер в «Город»: «село Миколаївка, Новомосковський район, Дніпропетровська область»
 * -> {names: ['Миколаївка'], hints: ['новомо', 'дніпро']}. hints - начала названий района и области.
 */
function razobratGorod_(city) {
  var raw = String(city || '').replace(/[«»"']/g, ' ');
  var names = [], hints = [];
  raw.split(/[,;()]/).forEach(function (part) {
    var p = (' ' + part + ' ').replace(/(^|[\s.])(місто|город|село|смт|селище|пгт|поселок|посёлок|м|г|с)\.?(?=\s)/gi, ' ')
      .replace(/\s+/g, ' ').trim();
    if (!p || /\d/.test(p) || /вул|улиц|просп|пров|відділ|отдел|пошт/i.test(p)) return;
    if (/обл|район|р-н|округ/i.test(p)) {
      var word = p.replace(/област[ьі]|обл\.?|район|р-н|округ/gi, ' ').trim().split(/\s+/)[0] || '';
      if (word.length >= 4) hints.push(word.toLowerCase().slice(0, 5));
      return;
    }
    var ua = NP_GORODA_RU[p.toLowerCase()];
    if (ua && names.indexOf(ua) < 0) names.push(ua);
    if (names.indexOf(p) < 0) names.push(p);
  });
  return { names: names, hints: hints };
}

/** Для совместимости: варианты названия населённого пункта. */
function variantyGoroda_(city) { return razobratGorod_(city).names; }

function nomerNp_(np) {
  var m = String(np || '').match(/\d+/);
  return m ? m[0] : '';
}

function tekstAdresa_(w) { return w.CityDescription + ', ' + w.Description; }

/** Сколько подсказок (район, область) совпало с вариантом. Районы переименовывали, поэтому хватает и области. */
function ochkiPodskazok_(cand, hints) {
  var text = (String(cand.area || '') + ' ' + String(cand.region || '') + ' ' + String(cand.present || '')).toLowerCase();
  return hints.filter(function (h) { return text.indexOf(h) >= 0; }).length;
}

/**
 * mesta: [{city, np}] -> массив адресов в формате НП «Місто, Відділення №27 (до 30 кг): вул. ...».
 * Населённых пунктов с одним названием много (Миколаївка), поэтому адрес ставится, только если вариант один:
 *   1) ищем все населённые пункты с таким названием, где есть отделения;
 *   2) оставляем те, где есть отделение с нужным номером;
 *   3) если вариантов несколько - отбираем по району/области из разговора.
 * Не нашлось - «⚠ не найдено в НП: ...», несколько вариантов - «⚠ несколько вариантов: ...» (поправить руками).
 */
function adresaNp_(mesta) {
  var cache = CacheService.getScriptCache();
  var result = new Array(mesta.length);
  var zadachi = {};
  mesta.forEach(function (m, n) {
    var city = String(m.city || '').trim(), num = nomerNp_(m.np);
    if (!num) {   // номер иногда попадает в «Город»: «Харків, 15», «Житомир, №2»
      var mm = city.match(/(?:№|відділення|отделение|пошт\S*)?\s*№?\s*(\d{1,5})\s*$/i);
      if (mm) num = mm[1];
    }
    var zapas = '⚠ не найдено в НП: ' + [city, num ? '№' + num : ''].filter(String).join(', ');
    if (!city || !num) { result[n] = (city || num) ? zapas : ''; return; }
    var key = 'np2|' + city.toLowerCase() + '|' + num;
    var hit = cache.get(key);
    if (hit) { result[n] = hit; return; }
    if (!zadachi[key]) {
      var g = razobratGorod_(city);
      zadachi[key] = { key: key, num: num, names: g.names, hints: g.hints, zapas: zapas, cands: [], idx: [] };
    }
    zadachi[key].idx.push(n);
  });
  var list = Object.keys(zadachi).map(function (k) { return zadachi[k]; });
  var errors = {};
  function zapomnitOshibku(resp) { if (resp.errors && resp.errors.length) errors[String(resp.errors[0])] = true; }

  // Шаг 1: все населённые пункты с таким названием (с районом и областью).
  var s1 = [];
  list.forEach(function (z) { z.names.forEach(function (name) { s1.push({ z: z, name: name }); }); });
  var r1 = npZaprosy_(s1.map(function (s) {
    return { model: 'Address', method: 'searchSettlements', props: { CityName: s.name, Limit: '50', Page: '1' } };
  }));
  r1.forEach(function (resp, i) {
    zapomnitOshibku(resp);
    var z = s1[i].z, name = s1[i].name.toLowerCase();
    var addrs = (resp.success && resp.data && resp.data[0] && resp.data[0].Addresses) || [];
    var withWh = addrs.filter(function (a) { return a.DeliveryCity && Number(a.Warehouses) > 0; });
    var exact = withWh.filter(function (a) { return String(a.MainDescription).toLowerCase() === name; });
    (exact.length ? exact : withWh).slice(0, 12).forEach(function (a) {
      if (z.cands.some(function (c) { return c.ref === a.DeliveryCity; })) return;
      z.cands.push({ ref: a.DeliveryCity, area: a.Area, region: a.Region, present: a.Present });
    });
  });

  // Шаг 1б: название не нашлось (например, по-русски «Белая Церковь») - ищем город среди городов НП.
  var s1b = [];
  list.forEach(function (z) {
    if (z.cands.length) return;
    z.names.forEach(function (name) { s1b.push({ z: z, name: name }); });
  });
  var r1b = npZaprosy_(s1b.map(function (s) {
    return { model: 'Address', method: 'getCities', props: { FindByString: s.name, Limit: '10' } };
  }));
  r1b.forEach(function (resp, i) {
    var z = s1b[i].z;
    ((resp.success && resp.data) || []).forEach(function (c) {
      if (z.cands.some(function (x) { return x.ref === c.Ref; })) return;
      z.cands.push({ ref: c.Ref, area: c.AreaDescription, region: '', present: c.Description });
    });
  });

  // Шаг 2: в каких из них есть отделение с нужным номером.
  var s2 = [];
  list.forEach(function (z) { z.cands.forEach(function (c) { s2.push({ z: z, c: c }); }); });
  var r2 = npZaprosy_(s2.map(function (s) {
    return { model: 'AddressGeneral', method: 'getWarehouses', props: { CityRef: s.c.ref, WarehouseId: s.z.num, Limit: '5' } };
  }));
  r2.forEach(function (resp, i) {
    zapomnitOshibku(resp);
    var c = s2[i].c, num = s2[i].z.num;
    var w = ((resp.success && resp.data) || []).filter(function (d) { return String(d.Number) === num; })[0];
    if (w) c.address = tekstAdresa_(w);
  });

  list.forEach(function (z) {
    var ok = z.cands.filter(function (c) { return c.address; });
    if (ok.length > 1 && z.hints.length) {
      var best = Math.max.apply(null, ok.map(function (c) { return ochkiPodskazok_(c, z.hints); }));
      if (best > 0) ok = ok.filter(function (c) { return ochkiPodskazok_(c, z.hints) === best; });
    }
    var value;
    if (ok.length === 1) {
      value = ok[0].address;
      cache.put(z.key, value, 21600);
    } else if (ok.length > 1) {
      value = '⚠ несколько вариантов №' + z.num + ': ' +
        ok.slice(0, 4).map(function (c) { return c.present || c.address; }).join(' / ');
    } else {
      value = z.zapas;
    }
    z.idx.forEach(function (n) { result[n] = value; });
  });

  var errText = Object.keys(NP_OSHIBKI).map(function (k) { return k + ' (' + NP_OSHIBKI[k] + ')'; }).join('; ');
  if (errText && list.length) {
    SpreadsheetApp.getActive().toast('Ответы НП с ошибкой (после повторов часть могла пройти): ' + errText.slice(0, 250) +
      '. Если ошибка про ключ - меню «Заявки» -> «Указать ключ Новой Почты».', 'Новая Почта', 15);
  }
  return result;
}

/**
 * Кнопка «Найти адреса НП заново»: для строк пула с ⚠ (ещё не отправленных) ищет адрес по тексту ячейки.
 * Можно вписать в ячейку своё: «Миколаївка, Новомосковський район, №1» или оставить один из
 * «нескольких вариантов» - и нажать кнопку.
 */
function obnovitAdresaNp() {
  var ss = SpreadsheetApp.getActive();
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || pool.getLastRow() < 2) { SpreadsheetApp.getUi().alert('Лист «' + ZAYAVKI_CFG.POOL_SHEET + '» пуст.'); return; }
  podgotovitPul_(pool);
  var col = P_('Адрес НП') + 1;
  var rows = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();
  var nomera = [], mesta = [];
  rows.forEach(function (r, i) {
    var a = String(r[P_('Адрес НП')] || '').trim();
    if (String(r[P_('Статус CRM')]).indexOf('OK') === 0) return;
    // ищем строки с ⚠ и то, что вписано вручную коротко («Обухів, 37»): в нём нет «Відділення/Поштомат/Пункт»
    var gotovyi = /Відділення|Поштомат|Пункт/i.test(a) && a.indexOf('⚠') !== 0;
    if (!a || gotovyi || adresNeNp_(a)) return;   // «укрпошта» / «нелид» - не адреса НП, не ищем
    var m = a.match(/№\s*(\d+)/) || a.match(/,\s*(\d{1,5})\s*$/);
    var city = a.replace(/^⚠[^:]*:\s*/, '').split('/')[0]
      .replace(/\[([^\]]*)\]/, function (_, x) {   // вариант с сервера: «... [Самарівський, Дніпропетровська]»
        var p = x.split(','); return ', ' + p[0].trim() + ' район' + (p[1] ? ', ' + p[1].trim() + ' область' : '');
      })
      .replace(/Відділення[^,]*|Поштомат[^,]*|Пункт[^,]*/gi, '')
      .replace(/,?\s*№\s*\d+/g, '').replace(/,\s*\d{1,5}\s*$/, '').replace(/^≈\s*/, '').trim();
    nomera.push(i);
    mesta.push({ city: city, np: m ? m[1] : '' });
  });
  if (!mesta.length) { ss.toast('Строк с ⚠ или вписанных вручную адресов нет.', 'Новая Почта', 6); return; }
  ss.toast('Ищу адреса: ' + mesta.length + '...', 'Новая Почта', 30);
  var adresa = adresaNp_(mesta), ok = 0;
  nomera.forEach(function (i, n) {
    pool.getRange(i + 2, col).setValue(adresa[n]);
    if (String(adresa[n]).indexOf('⚠') !== 0) ok++;
  });
  ss.toast('Найдено: ' + ok + ' из ' + mesta.length + '. Остальные с ⚠ поправьте и нажмите ещё раз.', 'Новая Почта', 10);
}

function ukazatKlyuchCrm1() { ukazatKlyuchCrm_(CRM_LIST[0]); }
function ukazatKlyuchCrm2() { ukazatKlyuchCrm_(CRM_LIST[1]); }

/** Кнопки «Указать ключ»: ключ и адрес хранятся в свойствах скрипта, а не в коде. */
function ukazatKlyuchCrm_(crm) {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  if (crm.urlProp) {
    var cur = props.getProperty(crm.urlProp) || '';
    var r1 = ui.prompt(crm.name + ': адрес',
      'Адрес этой LP-CRM, например https://name.lp-crm.biz' + (cur ? '\nСейчас: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    var url = r1.getResponseText().trim();
    if (url) {
      if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
      props.setProperty(crm.urlProp, url.replace(/\/+$/, '').replace(/\/api\/.*$/i, ''));
    } else if (!cur) return;
  }
  var res = ui.prompt(crm.name + ': ключ', 'Вставьте API-ключ из настроек LP-CRM:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  props.setProperty(crm.keyProp, key);
  SpreadsheetApp.getActive().toast('Сохранено.', crm.name, 5);
}

/** Адрес API для отправки заказа или '' если адрес не указан. */
function adresCrm_(crm) {
  var base = crm.urlProp ? (PropertiesService.getScriptProperties().getProperty(crm.urlProp) || crm.defaultUrl) : crm.defaultUrl;
  return base ? base + '/api/addNewOrder.html' : '';
}

/** Длина строки в байтах UTF-8 (нужна для формата PHP serialize). */
function dlinaBait_(s) { return unescape(encodeURIComponent(s)).length; }

/** PHP serialize для строк, чисел и массивов: в таком виде LP-CRM принимает список товаров. */
function phpSerialize_(v) {
  if (Array.isArray(v)) {
    return 'a:' + v.length + ':{' + v.map(function (x, i) { return 'i:' + i + ';' + phpSerialize_(x); }).join('') + '}';
  }
  if (v !== null && typeof v === 'object') {
    var keys = Object.keys(v);
    return 'a:' + keys.length + ':{' + keys.map(function (k) { return phpSerialize_(k) + phpSerialize_(v[k]); }).join('') + '}';
  }
  var s = String(v === null || v === undefined ? '' : v);
  return 's:' + dlinaBait_(s) + ':"' + s + '";';
}

/** Собирает заказ для LP-CRM из строки листа «пул». */
function zakazDlyaCrm_(row, picks, key, crm) {
  var items = razobratTovary_(row[P_('Товар')]);
  var sum = chislo_(row[P_('Сумма')]);
  // В заказ идёт одна строка: первый товар из списка пула (колонка H), 1 штука, цена = сумма из таблицы.
  // В комментарий: полный состав заказа с количеством, сумма и допы (колонка «Допы»).
  var products = [];
  var savedId = String(row[P_('ID товара СРМ 1') + crm.idIndex] || '').trim();
  if (savedId) products.push({ product_id: savedId, price: sum, count: 1 });
  for (var n = 0; n < items.length && !products.length; n++) {
    var p = picks[klyuchTovara_(items[n].name)];
    var id = p ? p[crm.idIndex] : '';
    if (id) { products.push({ product_id: id, price: sum, count: 1 }); break; }
  }
  if (!products.length) return null;   // нет ID товара: такой заказ не отправляем, чтобы в CRM не было пустых
  var sostav = items.map(function (item) { return item.name + ' - ' + item.qty + ' шт'; }).join(', ');
  var comment = (sostav || String(row[P_('Товар')])) + '. Сумма: ' + sum + ' грн.';
  var dopy = String(row[P_('Допы')] || '').trim();
  if (dopy && !/^(нет|немає|-)$/i.test(dopy)) comment += '\nДопы: ' + dopy;
  var address = String(row[P_('Адрес НП')] || '').replace(/^≈\s*/, '').trim();   // ≈ - пометка «проверить», в CRM не нужна

  return {
    key: key,
    order_id: String(row[P_('Call ID')]),      // Call ID: по нему CRM не примет один заказ дважды
    country: ZAYAVKI_CFG.CRM_COUNTRY,
    office: crm.office,
    products: encodeURIComponent(phpSerialize_(products)),
    bayer_name: String(row[P_('ФИО')] || ''),
    phone: String(row[P_('Телефон')] || ''),
    email: '',
    comment: comment,
    delivery: ZAYAVKI_CFG.CRM_DELIVERY,
    delivery_adress: address,
    payment: ZAYAVKI_CFG.CRM_PAYMENT,
    // LP-CRM берёт «сайт-лендинг покупателя» из этого поля (на лендинге это данные сервера $_SERVER)
    sender: encodeURIComponent(phpSerialize_({
      SERVER_NAME: ZAYAVKI_CFG.CRM_SITE,
      HTTP_HOST: ZAYAVKI_CFG.CRM_SITE,
      HTTP_REFERER: 'https://' + ZAYAVKI_CFG.CRM_SITE + '/',
      REQUEST_URI: '/'
    })),
    utm_source: ''
  };
}

function otpravitVCrm1() { otpravitPulVCrm_(CRM_LIST[0]); }
function otpravitVCrm2() { otpravitPulVCrm_(CRM_LIST[1]); }

/**
 * Кнопки «Выпулить заказы в СРМ 1 / 2»: отправляют из листа «пул» строки, которые ещё не отправлены.
 * Заказ, уже отправленный в любую из CRM (статус начинается с «OK»), повторно не отправляется.
 */
function otpravitPulVCrm_(crm) {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var key = PropertiesService.getScriptProperties().getProperty(crm.keyProp);
  var url = adresCrm_(crm);
  var menuItem = crm.urlProp ? '«Указать адрес и ключ ' + crm.name + '»' : '«Указать ключ ' + crm.name + '»';
  if (!key || !url) { ui.alert('Сначала укажите ' + (url ? 'ключ' : 'адрес и ключ') + ': меню «Заявки» -> ' + menuItem + '.'); return; }
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || pool.getLastRow() < 2) { ui.alert('Лист «' + ZAYAVKI_CFG.POOL_SHEET + '» пуст.'); return; }
  podgotovitPul_(pool);

  var rows = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();
  var colCrm = P_('Статус CRM') + 1;
  var todo = [];
  rows.forEach(function (row, i) {
    if (!String(row[P_('Call ID')])) return;
    if (/^(OK|ДРОП)/.test(String(row[P_('Статус CRM')]))) return;   // уже в CRM или отдан дропу
    todo.push(i);
  });
  if (!todo.length) { ss.toast('Все заказы из пула уже отправлены.', 'LP-CRM', 6); return; }
  var answer = ui.alert('Отправка в ' + crm.name, 'Отправить заказов в ' + crm.name + ': ' + todo.length + '?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  var picks = spisokPula_().picks;
  var started = Date.now(), ok = 0, failed = 0, left = 0, noId = 0, noAdr = 0;
  for (var n = 0; n < todo.length; n++) {
    if (Date.now() - started > 4.5 * 60 * 1000) { left = todo.length - n; break; }  // лимит Google: 6 минут на запуск
    var i = todo[n], result;
    if (adresNeNp_(rows[i][P_('Адрес НП')])) {
      pool.getRange(i + 2, colCrm).setValue('НЕ ОТПРАВЛЕН: адрес НП не подтверждён (⚠/укрпошта/нелид). Поправьте адрес и нажмите «Найти адреса НП заново».');
      noAdr++;
      continue;
    }
    var zakaz = zakazDlyaCrm_(rows[i], picks, key, crm);
    if (!zakaz) {
      var names = razobratTovary_(rows[i][P_('Товар')]).map(function (it) { return it.name; }).join('; ');
      result = 'НЕ ОТПРАВЛЕН в ' + crm.name + ': нет ID товара. Впишите его в колонку «ID товара ' + crm.name + '»' +
               ' этой строки или в колонку ' + (crm.idIndex ? 'J' : 'I') + ' листа «' + ZAYAVKI_CFG.STATS_SHEET + '». Товары: ' + names;
      pool.getRange(i + 2, colCrm).setValue(result.slice(0, 500));
      noId++;
      continue;
    }
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'post', payload: zakaz, muteHttpExceptions: true
      });
      var text = resp.getContentText(), data = null;
      try { data = JSON.parse(text); } catch (e) {}
      if (data && data.status === 'ok') { result = 'OK ' + crm.name + ' ' + Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM HH:mm'); ok++; }
      else {
        var msg = data && data.message ? (typeof data.message === 'string' ? data.message : JSON.stringify(data.message)) : text;
        result = 'ОШИБКА ' + crm.name + ': ' + String(msg).slice(0, 300); failed++;
      }
    } catch (e) {
      result = 'ОШИБКА ' + crm.name + ': ' + String(e.message || e).slice(0, 300); failed++;
    }
    pool.getRange(i + 2, colCrm).setValue(result);   // пишем сразу, чтобы не потерять при обрыве
  }
  SpreadsheetApp.flush();
  ui.alert(crm.name + '. Отправлено: ' + ok + ', с ошибкой: ' + failed +
           (noId ? ', без ID товара (не отправлены): ' + noId : '') +
           (noAdr ? ', с неподтверждённым адресом ⚠ (не отправлены): ' + noAdr : '') +
           (left ? '. Не успели: ' + left + ', нажмите кнопку ещё раз.' : '.') +
           (failed ? ' Причины в колонке «Статус CRM» листа «' + ZAYAVKI_CFG.POOL_SHEET + '».' : ''));
}
/* ===================== Проверка пула ИИ (через OpenRouter) =====================
 * Кнопки «Проверить пул ИИ: Sonnet 5.5 / Haiku 5.5». Модели Claude вызываются через OpenRouter (openrouter.ai).
 * Для каждой строки пула, где «Проверка ИИ» пустая и заказ ещё не отправлен в CRM, модель читает транскрипт и:
 *   1) сама ищет отделение Новой Почты через инструменты (поиск населённого пункта, отделения по номеру
 *      или улице) и выбирает его только из реальных результатов НП -> «Адрес НП» в формате НП;
 *   2) определяет допы, на которые клиент согласился -> «Допы» (с цитатой из разговора).
 * Строки с уже заполненной «Проверка ИИ» не проверяются повторно (чтобы перепроверить - очистите ячейку).
 * Проверка идёт порциями по ~4 минуты и сама продолжается каждую минуту, пока не пройдёт весь пул.
 * Когда весь пул проверен, строки каждого выпула сортируются: сначала ✓, потом ≈, потом ⚠.
 */

// id - название модели в OpenRouter (страница модели на openrouter.ai, кнопка копирования рядом с названием)
var II_MODELS = {
  sonnet: { id: 'anthropic/claude-sonnet-5.5', name: 'Sonnet 5.5', price: { inp: 2, out: 10 } },
  haiku:  { id: 'anthropic/claude-haiku-5.5', name: 'Haiku 5.5', price: { inp: 0.1, out: 0.5 } }
};
var II_URL = 'https://openrouter.ai/api/v1/chat/completions';
var II_PARALLEL = 15;          // сколько заказов проверяется одновременно
var II_MAX_STEPS = 10;         // шагов (обращений к модели) на один заказ
var II_RUN_MS = 4 * 60 * 1000;   // лимит Google - 6 минут на запуск, оставляем запас

var II_SYSTEM = [
  'Ты проверяешь заказ интернет-магазина в Украине по расшифровке телефонного разговора оператора с клиентом.',
  'Речь на украинском, русском или суржике, в расшифровке есть ошибки распознавания (искажённые названия, числа).',
  '',
  'Задача 1. Определи, в какое отделение или почтомат Новой Почты доставить заказ, и найди его инструментами:',
  '- search_settlements: найти населённый пункт. Учитывай ошибки распознавания, русские названия',
  '  («Белая Церковь» = «Біла Церква»), район и область из разговора. Сёл с одним названием много - выбирай по району/области.',
  '- find_warehouses: найти отделение в населённом пункте по номеру или по улице.',
  'Ищи как внимательный человек: если номер не находится, он мог быть расслышан неверно - попробуй похожие номера',
  'и поиск по улице, если улица звучала. Если клиент назвал свою домашнюю улицу, а не отделение - не подставляй случайное отделение.',
  'Выбирать можно ТОЛЬКО отделение из результатов find_warehouses (по его warehouse_ref). Ничего не выдумывай.',
  '',
  'Задача 2. Допы: дополнительные товары (кроме основного товара заказа), на которые клиент СОГЛАСИЛСЯ.',
  'Если оператор предложил, а клиент отказался или не ответил согласием - это не доп. Укажи название и количество.',
  'Дополнительные штуки ОСНОВНОГО товара по акции (второй со скидкой, третий в подарок) - это не доп, они идут в units.',
  'Все допы - только в extras, не в problem.',
  'В extras_quote ВСЕГДА дай доказательство из разговора: дословную короткую цитату (до 25 слов), где оператор',
  'предлагает доп и что ответил клиент (например: «а крем не хочете? - ні, дякую»). Если предложений было несколько -',
  'цитаты через « / ». Если оператор допы не предлагал - так и напиши: «допы не предлагались».',
  '',
  'Задача 3. Сумма и количество. В сообщении указано, что записано в системе. Проверь по разговору:',
  '- total_sum: итоговая сумма, которую клиент заплатит при получении (со всеми скидками и допами), числом;',
  '- units: сколько штук ОСНОВНОГО товара клиент физически получит, включая подарочные по акции (1+1=3 -> 3);',
  '- sum_quote: дословная короткая цитата (до 30 слов), где звучат итоговая сумма и количество/акция.',
  'Если в разговоре сумма или количество не звучат - верни то, что записано в системе, а в sum_quote напиши «не звучало».',
  '',
  'Перед итогом ОБЯЗАТЕЛЬНО поищи отделение инструментами (хотя бы search_settlements), даже если адрес кажется неясным.',
  'В конце вызови submit_result. confidence: high - адрес однозначно подтверждается разговором;',
  'medium - выбран наиболее вероятный вариант, стоит глянуть; low или пустой warehouse_ref - не удалось определить.',
  'problem - ОЧЕНЬ коротко, до 12 слов, что не так с адресом (например: «клиент не назвал номер отделения»).',
  'Пусто, если всё однозначно.',
  '',
  'Задача 4. Способ доставки (поля delivery_kind и delivery_quote в submit_result) - как клиент получает посылку:',
  '- np: звучит Нова Пошта/«Новая почта»/НП, или номер отделения/почтомата, или адрес (улица, дом) для доставки;',
  '- ukrposhta: звучит Укрпошта/Укрпочта, или почтовый индекс из 5 цифр;',
  '- other: другой перевозчик (Meest/Міст, Justin, Delivery, САТ), самовывоз, курьер, или вообще не звучит почта/отделение/индекс/адрес;',
  '- unclear: звучат и Нова Пошта, и Укрпошта, или ты не уверен.',
  'delivery_quote - дословная короткая цитата из разговора, доказывающая выбор (для other без упоминаний: «не звучало»).',
  'Выбирай по тому, что клиент выбрал В ИТОГЕ. Это поле заполняй всегда, даже если отделение нашлось.',
  '',
  'Задача 5. Получатель (поля recipient_name и name_quote). Если в разговоре звучат фамилия и/или имя получателя посылки,',
  'запиши их в recipient_name в именительном падеже в порядке «Фамилия Имя Отчество» (что звучало; не выдумывай недостающее).',
  'Имя оператора не подходит. name_quote - дословная цитата, где это звучит. Если не звучало: recipient_name пустая строка, name_quote «не звучало».'
].join('\n');

var II_TOOLS = [
  {
    name: 'search_settlements',
    description: 'Поиск населённых пунктов Украины в базе Новой Почты по названию (можно часть названия). ' +
      'Возвращает до 15 вариантов с районом, областью и settlement_ref для find_warehouses.',
    parameters: { type: 'object', properties: { query: { type: 'string', description: 'Название населённого пункта по-украински' } },
                  required: ['query'] }
  },
  {
    name: 'find_warehouses',
    description: 'Отделения и почтоматы Новой Почты в населённом пункте. Укажите number (номер отделения) или street ' +
      '(часть названия улицы). Без них вернёт первые отделения списка.',
    parameters: { type: 'object', properties: {
      settlement_ref: { type: 'string' }, number: { type: 'string' }, street: { type: 'string' } },
      required: ['settlement_ref'] }
  },
  {
    name: 'submit_result',
    description: 'Итог проверки заказа. Вызвать один раз в конце.',
    parameters: { type: 'object', properties: {
      warehouse_ref: { type: 'string', description: 'warehouse_ref из find_warehouses или пустая строка, если не найдено' },
      confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
      problem: { type: 'string' },
      extras_quote: { type: 'string', description: 'Дословная цитата: предложение допа и ответ клиента, или «допы не предлагались»' },
      total_sum: { type: 'number', description: 'Итоговая сумма к оплате по разговору, грн' },
      units: { type: 'integer', description: 'Штук основного товара в посылке, включая подарочные' },
      sum_quote: { type: 'string', description: 'Цитата, где звучат сумма и количество, или «не звучало»' },
      delivery_kind: { type: 'string', enum: ['np', 'ukrposhta', 'other', 'unclear'], description: 'Способ доставки по разговору' },
      delivery_quote: { type: 'string', description: 'Дословная цитата про способ доставки или «не звучало»' },
      recipient_name: { type: 'string', description: 'Фамилия и/или имя получателя из разговора, или пустая строка' },
      name_quote: { type: 'string', description: 'Дословная цитата, где звучит фамилия/имя, или «не звучало»' },
      extras: { type: 'array', items: { type: 'object', properties: {
        name: { type: 'string' }, qty: { type: 'integer' } }, required: ['name'] } }
    }, required: ['warehouse_ref', 'confidence', 'extras', 'extras_quote', 'total_sum', 'units', 'sum_quote', 'delivery_kind', 'delivery_quote', 'recipient_name', 'name_quote'] }
  }
].map(function (t) { return { type: 'function', function: t }; });

function proveritPulSonnet() { zapustitProverkuIi_('sonnet'); }
function proveritPulHaiku() { zapustitProverkuIi_('haiku'); }

function ukazatKlyuchOpenRouter() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Ключ OpenRouter', 'Вставьте API-ключ с openrouter.ai (начинается с sk-or-):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('OPENROUTER_KEY', key);
  SpreadsheetApp.getActive().toast('Ключ сохранён.', 'Проверка ИИ', 5);
}

function udalitTriggerIi_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'prodolzhitProverkuIi') ScriptApp.deleteTrigger(t);
  });
}

/** Кнопка «Как идёт проверка ИИ»: идёт ли проверка, сколько осталось, когда была последняя порция. */
function statusProverkiIi() {
  var props = PropertiesService.getScriptProperties();
  var pool = SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  var left = pool && pool.getLastRow() > 1 ? (podgotovitPul_(pool), strokiDlyaProverki_(pool).length) : 0;
  var modelKey = props.getProperty('II_MODEL');
  var idet = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'prodolzhitProverkuIi'; });
  var last = props.getProperty('II_LAST_RUN');
  SpreadsheetApp.getUi().alert(
    (idet && modelKey ? 'Проверка идёт: ' + II_MODELS[modelKey].name + '.' : 'Проверка сейчас не запущена.') +
    '\nНе проверено строк в пуле: ' + left + '.' +
    (last ? '\nПоследняя порция: ' + last + '.' : '') +
    '\nПотрачено всего ≈ $' + parseFloat(props.getProperty('II_COST_OR') || '0').toFixed(2) + '.');
}

function ostanovitProverkuIi() {
  udalitTriggerIi_();
  PropertiesService.getScriptProperties().deleteProperty('II_MODEL');
  SpreadsheetApp.getActive().toast('Проверка ИИ остановлена. Уже проверенные строки сохранены.', 'Проверка ИИ', 6);
}

function zapustitProverkuIi_(modelKey) {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OPENROUTER_KEY')) {
    SpreadsheetApp.getUi().alert('Сначала укажите ключ: меню «Заявки» -> «Указать ключ OpenRouter».');
    return;
  }
  props.setProperty('II_MODEL', modelKey);
  props.deleteProperty('II_SKIP');   // новый запуск: строки со снятым ⚠ снова доступны (Sonnet их и берёт)
  PROPUSK_II_CACHE_ = null;
  udalitTriggerIi_();
  pometitPartiyu_();
  try { props.setProperty('II_TOTAL', String(strokiDlyaProverki_(SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.POOL_SHEET)).length)); } catch (e) {}
  ScriptApp.newTrigger('prodolzhitProverkuIi').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getActive().toast('Проверка ' + II_MODELS[modelKey].name + ' запущена. Строки заполняются по мере проверки, ' +
    'дальше она продолжится сама каждую минуту. В конце строки отсортируются: ✓, ≈, ⚠.', 'Проверка ИИ', 10);
  prodolzhitProverkuIi();
}

/** Порция проверки (запускается кнопкой и затем триггером каждую минуту). */
function prodolzhitProverkuIi() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return;   // предыдущая порция ещё работает
  try {
    proverkaPorcii_();
  } finally {
    lock.releaseLock();
  }
}

/** Call ID, у которых Haiku снял ⚠ (адрес НП назван): «Проверка ИИ» пустая для Sonnet, но Haiku их второй раз не берёт. */
var PROPUSK_II_CACHE_ = null;
function propuskIi_() {
  if (!PROPUSK_II_CACHE_) {
    PROPUSK_II_CACHE_ = {};
    String(PropertiesService.getScriptProperties().getProperty('II_SKIP') || '').split(',').forEach(function (c) { if (c) PROPUSK_II_CACHE_[c] = true; });
  }
  return PROPUSK_II_CACHE_;
}
function dobavitPropuskIi_(callId) {
  var props = PropertiesService.getScriptProperties();
  var cur = String(props.getProperty('II_SKIP') || '');
  if (cur.length > 8500) return;
  props.setProperty('II_SKIP', cur + (cur ? ',' : '') + callId);
  PROPUSK_II_CACHE_ = null;
}

/** Адрес НП не является отделением: ⚠, «укрпошта» или «нелид». */
function adresNeNp_(a) {
  var t = String(a || '').trim();
  return t.indexOf('⚠') === 0 || /^(укрпошта|нелид)$/i.test(t);
}

function strokiDlyaProverki_(pool) {
  var rows = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();
  var out = [];
  rows.forEach(function (r, i) {
    if (!String(r[P_('Call ID')]).trim()) return;
    if (String(r[P_('Проверка ИИ')]).trim()) return;                   // уже проверено
    if (/^(OK|ДРОП)/.test(String(r[P_('Статус CRM')]))) return;        // уже в CRM или у дропа
    if (propuskIi_()[String(r[P_('Call ID')]).trim()]) return;          // Haiku уже разобрал и снял ⚠ - дальше Sonnet
    out.push({ row: i + 2, data: r });
  });
  return out;
}

/**
 * Помечает строки, которые сейчас пойдут на проверку, одной «партией» («проверка 23.10 16:05»).
 * После проверки сортируются только строки своей партии, уже разобранные раньше остаются на местах.
 * Строки, которые перепроверяются (очищена «Проверка ИИ»), остаются в своей прежней партии -
 * после перепроверки они переедут в нужную группу (✓, ≈, ⚠) внутри неё.
 */
function pometitPartiyu_() {
  var pool = SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || pool.getLastRow() < 2) return;
  podgotovitPul_(pool);
  var stamp = 'проверка ' + Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM HH:mm');
  var col = P_('Партия') + 1;
  var rng = pool.getRange(2, col, pool.getLastRow() - 1, 1);
  var vals = rng.getValues();
  var todo = {};
  strokiDlyaProverki_(pool).forEach(function (it) { todo[it.row] = true; });
  var changed = false;
  var dopy = pool.getRange(2, P_('Допы') + 1, pool.getLastRow() - 1, 1).getValues();
  vals.forEach(function (v, i) {
    if (!todo[i + 2] || /^проверка/.test(String(v[0]))) return;
    if (String(dopy[i][0]).trim()) return;   // строку уже проверяли раньше (перепроверка) - остаётся в своей группе
    v[0] = stamp;
    changed = true;
  });
  if (changed) rng.setValues(vals);
}

/** Ранг строки для сортировки: ✓ точные, ≈ приблизительные, ⚠ с ошибкой, остальное в конце. */
function rangIi_(text) {
  var t = String(text || '').trim();
  return t.indexOf('✓') === 0 ? 0 : t.indexOf('≈') === 0 ? 1 : t.indexOf('⚠') === 0 ? 2 : 3;
}

/**
 * Сортирует пул внутри каждой партии (колонка «Партия» = запуск проверки): сначала ✓, потом ≈, потом ⚠.
 * Порядок партий не меняется; строки, ещё не бывшие в запуске проверки, сортируются как одна группа.
 */
function sortirovatPul_(pool, vseVmeste) {
  var n = pool.getLastRow() - 1;
  if (n < 2) return;
  var w = POOL_COLUMNS.length;
  var rows = pool.getRange(2, 1, n, w).getValues();
  var groups = [], byKey = {};
  rows.forEach(function (r, i) {
    // группа = запуск проверки («проверка 23.10 16:05»); всё, что ещё не было в таком запуске, - одна общая группа
    var p = String(r[P_('Партия')] || '').trim();
    var key = vseVmeste ? '__all' : (/^проверка/.test(p) ? p : '__old');
    if (!(key in byKey)) { byKey[key] = groups.length; groups.push([]); }
    groups[byKey[key]].push({ r: r, i: i });
  });
  var out = [];
  groups.forEach(function (g) {
    g.sort(function (a, b) { return rangIi_(a.r[P_('Проверка ИИ')]) - rangIi_(b.r[P_('Проверка ИИ')]) || a.i - b.i; });
    g.forEach(function (x) { out.push(x.r); });
  });
  var changed = out.some(function (r, i) { return r !== rows[i]; });
  if (changed) pool.getRange(2, 1, n, w).setValues(out);
}

/** Кнопка «Отсортировать пул»: весь пул целиком (автоматическая сортировка после проверки - по партиям). */
function sortirovatPulKnopka() {
  var pool = SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || pool.getLastRow() < 3) return;
  if (ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'prodolzhitProverkuIi'; })) {
    SpreadsheetApp.getUi().alert('Сейчас идёт проверка ИИ - пул отсортируется сам, когда она закончится.');
    return;
  }
  podgotovitPul_(pool);
  sortirovatPul_(pool, true);   // по кнопке - весь пул целиком
  ssylkiTg_(pool);
  SpreadsheetApp.getActive().toast('Весь пул отсортирован: ✓, ≈, ⚠.', ZAYAVKI_CFG.POOL_SHEET, 6);
}

function proverkaPorcii_() {
  var started = Date.now();
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var modelKey = props.getProperty('II_MODEL');
  var key = props.getProperty('OPENROUTER_KEY');
  if (!modelKey || !key) { udalitTriggerIi_(); return; }
  var model = II_MODELS[modelKey];
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || pool.getLastRow() < 2) { udalitTriggerIi_(); return; }
  podgotovitPul_(pool);

  var queue = strokiDlyaProverki_(pool);
  if (!queue.length) {
    udalitTriggerIi_();
    props.deleteProperty('II_MODEL');
    sortirovatPul_(pool);
    ssylkiTg_(pool);
    ss.toast('Все заказы пула проверены и отсортированы.', 'Проверка ИИ', 8);
    return;
  }

  var cost = 0, done = 0, fatal = '';
  var active = [];
  function vzyat() {
    while (active.length < II_PARALLEL && queue.length) active.push(novyiZakazIi_(queue.shift()));
  }
  vzyat();

  while (active.length && Date.now() - started < II_RUN_MS && !fatal) {
    // 1. Шаг модели для всех активных заказов параллельно
    var reqs = active.map(function (st) {
      return { url: II_URL, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
               headers: { 'Authorization': 'Bearer ' + key, 'X-Title': 'Binotel pool check' },
               payload: JSON.stringify({ model: model.id, max_tokens: 1024, tools: II_TOOLS, tool_choice: 'auto',
                                         messages: st.messages, usage: { include: true } }) };
    });
    var resps;
    try { resps = UrlFetchApp.fetchAll(reqs); } catch (e) { Utilities.sleep(5000); continue; }

    var npTasks = [];   // вызовы инструментов, которые надо выполнить через API НП
    var pauza = false;
    resps.forEach(function (resp, k) {
      var st = active[k];
      var code = resp.getResponseCode(), body = null;
      try { body = JSON.parse(resp.getContentText()); } catch (e) {}
      var err = body && body.error;
      if (err && code === 200) code = Number(err.code) || 500;   // OpenRouter иногда отдаёт ошибку с кодом 200
      var errText = err ? String(err.message || JSON.stringify(err)).slice(0, 150) : '';
      if (code === 401 || code === 403) { fatal = 'ключ OpenRouter не подходит (' + code + '). Меню «Заявки» -> «Указать ключ OpenRouter»'; return; }
      if (code === 402) { fatal = 'на балансе OpenRouter закончились деньги - пополните баланс'; return; }
      if ((code === 400 || code === 404) && /model/i.test(errText)) { fatal = 'модель ' + model.id + ' не найдена в OpenRouter: ' + errText; return; }
      if (code === 408 || code === 429 || code >= 500 || !body || !body.choices) {
        // перегрузка или временная ошибка: ждём и повторяем; после 5 неудач заказ откладывается до следующей порции
        pauza = true; st.sboi++;
        if (st.sboi > 5) { st.finished = true; st.otlozhen = true; }
        return;
      }
      if (code !== 200) { zavershitIi_(st, null, 'ошибка API ' + code + ': ' + errText); return; }
      var u = body.usage || {};
      cost += typeof u.cost === 'number' ? u.cost :
        ((u.prompt_tokens || 0) * model.price.inp + (u.completion_tokens || 0) * model.price.out) / 1e6;
      st.steps++;
      var msg = (body.choices[0] || {}).message || {};
      var calls = (msg.tool_calls || []).map(function (c) {
        var args = {};
        try { args = JSON.parse((c.function && c.function.arguments) || '{}'); } catch (e) {}
        return { id: c.id, name: c.function && c.function.name, input: args || {} };
      });
      var asst = { role: 'assistant', content: msg.content || '' };
      if (msg.tool_calls && msg.tool_calls.length) asst.tool_calls = msg.tool_calls;
      st.messages.push(asst);

      var submit = calls.filter(function (c) { return c.name === 'submit_result'; })[0];
      if (submit) {
        var bezPoiska = !String(submit.input.warehouse_ref || '').trim() && !st.poiskov;
        if (!(bezPoiska && !st.napomnil)) { zavershitIi_(st, submit.input, ''); return; }
        st.napomnil = true;   // итог без единого поиска - просим сначала поискать
      }
      if (!calls.length) {
        if (st.steps >= II_MAX_STEPS) { zavershitIi_(st, null, 'модель не дала итог'); return; }
        st.messages.push({ role: 'user', content: 'Вызови submit_result с итогом.' });
        return;
      }
      if (st.steps >= II_MAX_STEPS) { zavershitIi_(st, null, 'слишком много шагов поиска'); return; }
      // на каждый вызов инструмента нужен ответ (иначе модель не продолжит)
      st.pending = calls.map(function (c) {
        var p = { id: c.id, name: c.name, input: c.input, result: null };
        if (c.name === 'submit_result') {
          p.result = { error: 'Сначала найди отделение инструментами search_settlements и find_warehouses, затем снова вызови submit_result.' };
        } else {
          st.poiskov = (st.poiskov || 0) + 1;
          npTasks.push({ st: st, p: p });
        }
        return p;
      });
    });
    if (fatal) break;

    // 2. Выполняем инструменты (запросы к НП) для всех заказов разом
    if (npTasks.length) vypolnitInstrumentyIi_(npTasks);
    active.forEach(function (st) {
      if (st.finished || !st.pending) return;
      st.pending.forEach(function (p) {
        st.messages.push({ role: 'tool', tool_call_id: p.id, content: JSON.stringify(p.result) });
      });
      st.pending = null;
    });

    // 3. Записываем готовые, добираем новые
    active.filter(function (st) { return st.finished && !st.otlozhen; }).forEach(function (st) { zapisatIi_(pool, st, model); done++; });
    if (active.some(function (st) { return st.otlozhen; })) queue = [];   // API перегружен - остальное в следующей порции
    active = active.filter(function (st) { return !st.finished; });
    vzyat();
    if (pauza) Utilities.sleep(3000);
  }

  // счётчик проверенных за сегодня (для панели)
  var den = Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM.yyyy');
  var bylo = props.getProperty('II_DAY') === den ? Number(props.getProperty('II_DAY_COUNT') || 0) : 0;
  props.setProperty('II_DAY', den); props.setProperty('II_DAY_COUNT', String(bylo + done));
  var total = parseFloat(props.getProperty('II_COST_OR') || '0') + cost;
  props.setProperty('II_COST_OR', String(total));
  if (fatal) {
    udalitTriggerIi_();
    props.deleteProperty('II_MODEL');
    ss.toast('Проверка остановлена: ' + fatal + '.', 'Проверка ИИ', 15);
    return;
  }
  var left = strokiDlyaProverki_(pool).length;
  props.setProperty('II_LAST_RUN', Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM HH:mm') +
                    ', проверено ' + done + ', осталось ' + left);
  if (!left) {
    udalitTriggerIi_();
    props.deleteProperty('II_MODEL');
    sortirovatPul_(pool);   // всё проверено - расставляем ✓, ≈, ⚠ внутри каждого выпула
    ssylkiTg_(pool);
  }
  ss.toast('Проверено: ' + done + (left ? ', осталось: ' + left + ' (продолжится само через минуту)' : ', пул проверен и отсортирован') +
           '. Эта порция ≈ $' + cost.toFixed(2) + ', всего ≈ $' + total.toFixed(2) + '.', 'Проверка ИИ ' + model.name, 10);
}

function novyiZakazIi_(item) {
  var r = item.data;
  var text = String(r[P_('Транскрипт')] || '').slice(0, 30000);
  var adr = String(r[P_('Адрес НП')] || '').trim();
  var tovar = String(r[P_('Товар')] || '').trim();
  var it = razobratTovary_(tovar);
  var zapSum = chislo_(r[P_('Сумма')]);
  var zapSht = chislo_(r[P_('Шт')]) || (it.length ? it[0].qty : 1);
  var intro = 'Основной товар заказа (по данным системы): ' + (tovar || 'не указан') + '.\n' +
    'Записано в системе: сумма ' + zapSum + ' грн, штук основного товара ' + zapSht + '.\n' +
    'Адрес, найденный автоматически (может быть неверным): ' + (adr || 'нет') + '.\n\nРасшифровка разговора:\n' + text;
  return {
    row: item.row, callId: String(r[P_('Call ID')]).trim(), steps: 0, sboi: 0, pending: null, finished: false, result: null, problem: '', transcript: text,
    zapSum: zapSum, zapSht: zapSht, fio: String(r[P_('ФИО')] || '').trim(),
    known: {},   // warehouse_ref -> адрес в формате НП (только то, что реально вернула НП)
    messages: [
      { role: 'system', content: II_SYSTEM },
      // cache_control: OpenRouter передаёт его Claude, повторное чтение транскрипта на следующих шагах дешевле
      { role: 'user', content: [{ type: 'text', text: intro, cache_control: { type: 'ephemeral' } }] }
    ]
  };
}

function zavershitIi_(st, result, problem) {
  st.finished = true;
  st.result = result;
  st.problem = problem;
}

/** Выполняет вызовы search_settlements / find_warehouses через API Новой Почты (пачками, с паузами). */
function vypolnitInstrumentyIi_(tasks) {
  var calls = [], map = [];
  tasks.forEach(function (t, k) {
    var p = t.p, inp = p.input;
    if (p.name === 'search_settlements') {
      calls.push({ model: 'Address', method: 'searchSettlements', props: { CityName: String(inp.query || ''), Limit: '20', Page: '1' } });
      map.push(k);
    } else if (p.name === 'find_warehouses') {
      var props = { CityRef: String(inp.settlement_ref || ''), Limit: '30', Page: '1' };
      if (inp.number) props.WarehouseId = String(inp.number).replace(/\D/g, '');
      else if (inp.street) props.FindByString = String(inp.street);
      calls.push({ model: 'AddressGeneral', method: 'getWarehouses', props: props });
      map.push(k);
    } else {
      p.result = { error: 'неизвестный инструмент' };
    }
  });
  var resps = calls.length ? npZaprosy_(calls) : [];
  var retry = [];   // населённый пункт не нашёлся (например, русское название) - пробуем список городов НП
  resps.forEach(function (resp, j) {
    var t = tasks[map[j]], p = t.p;
    if (!resp || resp.success === false) { p.result = { error: 'Новая Почта не ответила, попробуйте ещё раз' }; return; }
    if (p.name === 'search_settlements') {
      var addrs = (resp.data && resp.data[0] && resp.data[0].Addresses) || [];
      var list = addrs.filter(function (a) { return a.DeliveryCity && Number(a.Warehouses) > 0; }).slice(0, 15)
        .map(function (a) {
          return { settlement_ref: a.DeliveryCity, name: a.Present || a.MainDescription, area: a.Area, district: a.Region,
                   warehouses: Number(a.Warehouses) };
        });
      if (list.length) p.result = { settlements: list };
      else retry.push(t);
    } else {
      var ws = (resp.data || []).map(function (w) {
        var text = w.CityDescription + ', ' + w.Description;
        t.st.known[w.Ref] = text;
        t.st.regions = t.st.regions || {};
        t.st.regions[w.Ref] = oblastRaion_(w.SettlementAreaDescription, w.SettlementRegionsDescription);
        return { warehouse_ref: w.Ref, number: w.Number, address: text, type: w.CategoryOfWarehouse };
      });
      p.result = ws.length ? { warehouses: ws } : { warehouses: [], note: 'ничего не найдено' };
    }
  });
  if (retry.length) {
    var r2 = npZaprosy_(retry.map(function (t) {
      return { model: 'Address', method: 'getCities', props: { FindByString: String(t.p.input.query || ''), Limit: '15' } };
    }));
    r2.forEach(function (resp, j) {
      var list = ((resp && resp.data) || []).map(function (c) {
        return { settlement_ref: c.Ref, name: c.Description, area: c.AreaDescription || '', district: '' };
      });
      retry[j].p.result = list.length ? { settlements: list } :
        { settlements: [], note: 'не найдено; попробуйте другое написание или часть названия' };
    });
  }
}

/** Есть ли цитата в транскрипте (без учёта регистра, знаков и кавычек; хватает любого куска из 4+ слов подряд). */
function citataEst_(quote, transcript) {
  function norm(s) { return String(s || '').toLowerCase().replace(/[^a-zа-яіїєґ0-9]+/g, ' ').trim(); }
  var t = norm(transcript);
  return norm(quote).split(/ ?(?:\s—\s|\/|-) ?/).concat([norm(quote)]).some(function (part) {
    var w = norm(part).split(' ').filter(String);
    for (var i = 0; i + 4 <= w.length; i++) if (t.indexOf(w.slice(i, i + 4).join(' ')) >= 0) return true;
    return false;
  });
}

/**
 * ФИО, которое можно вписать вместо неполного: '' если менять не нужно.
 * Условия: в заказе меньше двух слов; у ИИ слов больше, чем было; цитата дословная; каждое слово ИИ есть в транскрипте
 * (сравнение по основе без окончания: «Савчук Галина» найдётся и в «Савчук Галину»).
 */
function fioIzRazgovora_(name, quote, bylo, transcript) {
  name = String(name || '').replace(/\s+/g, ' ').trim();
  if (!name || name.length > 80) return '';
  var nBylo = slovVFio_(bylo), nNov = slovVFio_(name);
  if (nBylo >= 2 || nNov <= nBylo || nNov > 4) return '';
  if (!citataDoslovno_(quote, transcript)) return '';
  function norm(x) { return String(x || '').toLowerCase().replace(/[’ʼ`]/g, "'").replace(/[^a-zа-яіїєґё' -]+/g, ' '); }
  var t = norm(transcript);
  var ok = norm(name).trim().split(/\s+/).every(function (w) {
    var osnova = w.length > 4 ? w.slice(0, w.length - 2) : w;
    return t.indexOf(osnova) >= 0;
  });
  return ok ? name : '';
}

/** Короткая цитата (даже одно слово) встречается в транскрипте дословно (без регистра и знаков) или проходит правило 4 слов. */
function citataDoslovno_(quote, transcript) {
  function norm(s) { return String(s || '').toLowerCase().replace(/[^a-zа-яіїєґ0-9]+/g, ' ').trim(); }
  var q = norm(quote);
  return !!q && (norm(transcript).indexOf(q) >= 0 || citataEst_(quote, transcript));
}

/** Первое предложение, не длиннее 90 символов: в ячейки пишем коротко, полный текст - в примечание. */
function korotko_(text) {
  var t = String(text || '').split(/(?<=[.!?])\s/)[0].trim();
  return t.length > 90 ? t.slice(0, 87).trim() + '...' : t;
}

/** Записывает итог проверки в строку пула: «Проверка ИИ», «Адрес НП», «Допы». */
function zapisatIi_(pool, st, model) {
  var r = st.result;
  var check, adr = null, dopy = null, note = '', region = '', sumSht = null, sht = null;
  if (!r) {
    check = '⚠ ' + (st.problem || 'проверка не удалась');
  } else {
    var ref = String(r.warehouse_ref || '').trim();
    var problem = String(r.problem || '').trim();
    note = problem;
    problem = korotko_(problem);
    if (ref && st.known[ref]) {
      var uveren = r.confidence === 'high';
      adr = (uveren ? '' : '≈ ') + st.known[ref];
      region = (st.regions || {})[ref] || '';
      check = (uveren ? '✓ ' : '≈ ') + model.name + (problem ? ': ' + problem : '');
    } else {
      adr = '⚠ ' + (problem || 'ИИ не смог определить отделение');
      check = '⚠ ' + model.name + ': ' + (problem || 'адрес не определён');
      // разбор ⚠ по разговору: НП (снять ⚠ - пойдёт Sonnet) / укрпошта / нелид; без дословной цитаты ничего не меняем
      var kind = String(r.delivery_kind || '').trim();
      var dq = String(r.delivery_quote || '').replace(/\s+/g, ' ').trim();
      var nezv = !dq || /не звучал/i.test(dq);
      var dqOk = nezv || citataDoslovno_(dq, st.transcript);
      var dqShort = dq.length > 80 ? dq.slice(0, 77) + '...' : dq;
      // подстраховка по самому транскрипту: модель могла дать неточную цитату или «unclear», хотя почта названа прямо
      var tl = st.transcript.toLowerCase();
      var estNp = /нов\S{0,3}\s+по[шч]т|відділен|отделен|поштомат|почтомат/.test(tl);
      var estUkr = /укрпо[шч]т|індекс|индекс/.test(tl);
      if (kind === 'ukrposhta' && !nezv && dqOk) {   // оператор часто сам говорит «доставка Новой почтой» - верим цитате клиента
        adr = 'укрпошта';
        check = 'укрпошта (' + model.name + '): «' + dqShort + '»';
      } else if (kind === 'other' && dqOk && (!nezv || (!estNp && !estUkr))) {
        adr = 'нелид';
        check = 'нелид (' + model.name + '): ' + (nezv ? 'в разговоре нет почты, отделения, индекса и адреса' : 'другая доставка «' + dqShort + '»');
      } else if ((kind === 'np' || (estNp && !estUkr)) && model === II_MODELS.haiku) {
        adr = '';
        check = '';   // НП / отделение / адрес названы - ⚠ снимаем, строку дальше проверит Sonnet (цитата тут не обязательна)
        st.snyatoVoprosy = true;
      }
    }
    var extras = (r.extras || []).filter(function (e) { return e && e.name; });
    dopy = extras.length ? extras.map(function (e) { return e.name + ' (' + (e.qty || 1) + ' шт)'; }).join('; ') : 'нет';
    var quote = String(r.extras_quote || '').replace(/\s+/g, ' ').trim();
    if (quote) {
      // цитату сверяем с транскриптом: если её там нет, модель пересказала своими словами - помечаем
      var est = /не предлаг/i.test(quote) || citataEst_(quote, st.transcript);
      if (!/не предлаг/i.test(quote) && !/^[«"]/.test(quote)) quote = '«' + quote + '»';
      dopy += ' — ' + quote + (est ? '' : ' (не дословно)');
    }

    // сумма и количество: сверяем с записанным
    var aiSum = Number(r.total_sum) || 0, aiSht = parseInt(r.units, 10) || 0;
    var sq = String(r.sum_quote || '').replace(/\s+/g, ' ').trim();
    var nezvuchalo = !sq || /не звучал/i.test(sq);
    var sumOk = !aiSum || nezvuchalo || Math.abs(aiSum - st.zapSum) < 1;
    var shtOk = !aiSht || nezvuchalo || aiSht === st.zapSht;
    if (aiSht && !nezvuchalo) sht = aiSht;   // количество по разговору (с подарочными) - в колонку «Шт»
    var qEst = nezvuchalo || citataEst_(sq, st.transcript);
    if (sq && !/^[«"]/.test(sq) && !nezvuchalo) sq = '«' + sq + '»';
    sumSht = (sumOk ? (aiSum || st.zapSum) + ' грн ✓' : '⚠ ' + aiSum + ' грн (записано ' + st.zapSum + ')') + ', ' +
             (shtOk ? (aiSht || st.zapSht) + ' шт ✓' : '≈ ' + aiSht + ' шт (было ' + st.zapSht + ')') +
             (sq ? ' — ' + sq + (qEst ? '' : ' (не дословно)') : '');
    // итоговая пометка строки: худшее из «адрес» и «сумма/шт» (по ней сортируется пул)
    if (!sumOk) check = '⚠ ' + model.name + ': сумма ' + aiSum + ', записано ' + st.zapSum + (check.indexOf('⚠') === 0 ? '; адрес' : '');
    else if (!shtOk && check.indexOf('✓') === 0) check = '≈ ' + model.name + ': шт по разговору ' + aiSht;
  }
  // ФИО из разговора: только если в заказе пусто или одно слово, а в разговоре звучит больше; каждое слово должно быть в транскрипте
  var fioNovoe = r ? fioIzRazgovora_(r.recipient_name, r.name_quote, st.fio, st.transcript) : '';
  if (fioNovoe) {
    pool.getRange(st.row, P_('ФИО') + 1).setValue(fioNovoe);
    if (check) check += ' · ФИО из разговора' + (st.fio ? ' (было: ' + st.fio + ')' : '');
  }
  pool.getRange(st.row, P_('Проверка ИИ') + 1).setValue(check.slice(0, 300));
  if (st.snyatoVoprosy) {
    if (st.callId) dobavitPropuskIi_(st.callId);
  }
  if (adr !== null) pool.getRange(st.row, P_('Адрес НП') + 1).setValue(adr.slice(0, 500));
  if (dopy !== null) pool.getRange(st.row, P_('Допы') + 1).setValue(dopy.slice(0, 500));
  if (sumSht !== null) pool.getRange(st.row, P_('Сумма/шт ИИ') + 1).setValue(sumSht.slice(0, 500));
  if (sht !== null) pool.getRange(st.row, P_('Шт') + 1).setValue(sht);
  pool.getRange(st.row, P_('Область, район') + 1).setValue(region);
}

/* ===================== Выгрузка для дропа (TXT) =====================
 * Выделите в листе «пул» строки заказов (можно несколько диапазонов через Ctrl) и нажмите
 * «Выгрузить выделенные заказы для дропа (TXT)». Откроется окно с текстом в формате дропа:
 * его можно скопировать в чат или скачать файлом. У выгруженных строк в «Статус CRM» ставится «ДРОП дата»,
 * такие заказы не отправляются в CRM. Артикулы берутся из колонки K «Артикул дропа» листа «Статистика».
 */

function oblastRaion_(area, district) {
  area = String(area || '').trim();
  district = String(district || '').trim();
  var parts = [];
  if (area) parts.push(/обл/i.test(area) ? area : area + ' обл.');
  if (district) parts.push(/р-н|район/i.test(district) ? district : district + ' р-н');
  return parts.join(', ');
}

/** Артикулы из «Статистики»: ключ товара (колонка H) -> артикул (колонка K). */
function artikulyDropa_() {
  var stats = SpreadsheetApp.getActive().getSheetByName(ZAYAVKI_CFG.STATS_SHEET);
  var out = {};
  if (!stats || stats.getLastRow() < 2) return out;
  shapkaPula_(stats);
  stats.getRange(2, ZAYAVKI_CFG.COL_PICK, stats.getLastRow() - 1, 4).getValues().forEach(function (r) {
    var key = klyuchTovara_(r[0]);
    if (key && String(r[3]).trim()) out[key] = String(r[3]).trim();
  });
  return out;
}

/** Номера строк (с 2-й), выделенных в листе. */
function vydelennyeStroki_(sheet) {
  var list = sheet.getActiveRangeList();
  var rows = {};
  (list ? list.getRanges() : [sheet.getActiveRange()]).forEach(function (rg) {
    if (!rg) return;
    for (var r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r >= 2) rows[r] = true;
  });
  return Object.keys(rows).map(Number).sort(function (a, b) { return a - b; });
}

/** Область и район для адресов, у которых их нет (заказы, проверенные без ИИ): ищем по названию в НП. */
function doiskatRegiony_(gorod) {
  var names = Object.keys(gorod);
  if (!names.length) return {};
  var resps = npZaprosy_(names.map(function (n) {
    return { model: 'Address', method: 'searchSettlements', props: { CityName: n, Limit: '20', Page: '1' } };
  }));
  var out = {};
  resps.forEach(function (resp, i) {
    var name = names[i].toLowerCase();
    var addrs = ((resp && resp.data && resp.data[0] && resp.data[0].Addresses) || []).filter(function (a) {
      return String(a.MainDescription).toLowerCase() === name && Number(a.Warehouses) > 0;
    });
    if (addrs.length === 1) out[names[i]] = oblastRaion_(addrs[0].Area, addrs[0].Region);   // однозначно
  });
  return out;
}

function vygruzitDlyaDropa() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || ss.getActiveSheet().getName() !== pool.getName()) {
    ui.alert('Откройте лист «' + ZAYAVKI_CFG.POOL_SHEET + '», выделите строки заказов для дропа и нажмите кнопку ещё раз.');
    return;
  }
  podgotovitPul_(pool);
  var nomera = vydelennyeStroki_(pool).filter(function (r) { return r <= pool.getLastRow(); });
  if (!nomera.length) {
    ui.alert('Выделите в пуле строки заказов, которые нужно передать дропу (можно мышкой по номерам строк, несколько - с Ctrl).');
    return;
  }
  var art = artikulyDropa_();
  var zakazy = [], propuski = [];
  var vse = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();
  nomera.forEach(function (r) {
    var row = vse[r - 2];
    if (!String(row[P_('Call ID')]).trim()) return;
    var adr = String(row[P_('Адрес НП')] || '').trim();
    if (!adr || adresNeNp_(adr)) { propuski.push(r + ' (адрес не подтверждён)'); return; }
    if (/^OK/.test(String(row[P_('Статус CRM')]))) { propuski.push(r + ' (уже в CRM)'); return; }
    zakazy.push({ r: r, row: row, adr: adr.replace(/^≈\s*/, '') });
  });
  if (!zakazy.length) { ui.alert('Нечего выгружать. Пропущены строки: ' + propuski.join(', ')); return; }

  // область и район: из проверки ИИ, а если нет - ищем по названию населённого пункта
  var nuzhno = {};
  zakazy.forEach(function (z) {
    z.region = String(z.row[P_('Область, район')] || '').trim();
    z.gorod = z.adr.split(',')[0].trim();
    if (!z.region) nuzhno[z.gorod] = true;
  });
  var naideno = doiskatRegiony_(nuzhno);

  var bloki = zakazy.map(function (z, n) {
    var tovar = String(z.row[P_('Товар')] || '').trim();
    var items = razobratTovary_(tovar);
    var a = items.length ? art[klyuchTovara_(items[0].name)] : '';
    var dopy = String(z.row[P_('Допы')] || '').split(' — ')[0].trim();   // без цитаты
    var tovarText = (a ? a + ' ' : '') + tovar + (dopy && !/^нет$/i.test(dopy) ? ' + ' + dopy : '');
    var region = z.region || naideno[z.gorod] || '';
    var mesto = z.gorod + (region ? ', ' + region : '');
    var punkt = z.adr.slice(z.gorod.length).replace(/^,\s*/, '');
    return [
      (n + 1) + '.',
      '- Артикул і назва товару: ' + tovarText,
      '- Ціна продажу: ' + chislo_(z.row[P_('Сумма')]) + ' грн',
      '- ПІБ отримувача: ' + String(z.row[P_('ФИО')] || '').trim(),
      '- Номер телефону: ' + String(z.row[P_('Телефон')] || '').trim(),
      '- Місто та область: ' + mesto,
      '- Відділення Нової пошти: ' + punkt
    ].join('\n');
  });
  var text = bloki.join('\n\n');

  var stamp = Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM HH:mm');
  zakazy.forEach(function (z) { pool.getRange(z.r, P_('Статус CRM') + 1).setValue('ДРОП ' + stamp); });

  var name = 'drop_' + Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM_HH-mm') + '.txt';
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px">' +
    '<p>Заказов: <b>' + zakazy.length + '</b>' + (propuski.length ? '. Пропущены строки: ' + propuski.join(', ') : '') +
    '. Без артикула: ' + zakazy.filter(function (z) {
      var it = razobratTovary_(z.row[P_('Товар')]); return !(it.length && art[klyuchTovara_(it[0].name)]);
    }).length + ' (артикулы - колонка K листа «Статистика»).</p>' +
    '<textarea id="t" style="width:100%;height:330px;font-family:monospace;font-size:12px"></textarea>' +
    '<p><button onclick="dl()">Скачать .txt</button> <button onclick="cp()">Копировать</button> <span id="m"></span></p>' +
    '<script>var T=' + JSON.stringify(text).replace(/</g, '\\u003c') + ';document.getElementById("t").value=T;' +
    'function dl(){var b=new Blob([T],{type:"text/plain;charset=utf-8"});var a=document.createElement("a");' +
    'a.href=URL.createObjectURL(b);a.download=' + JSON.stringify(name) + ';document.body.appendChild(a);a.click();a.remove();}' +
    'function cp(){var t=document.getElementById("t");t.select();document.execCommand("copy");' +
    'document.getElementById("m").textContent="Скопировано";}</script></div>'
  ).setWidth(720).setHeight(470);
  ui.showModalDialog(html, 'Заказы для дропа');
}

/* ===================== ТТН Новой Почты (лист «НП лиды») =====================
 * 1. Лист «Настройки НП»: откуда отправляем (город, номер отделения), телефон отправителя, размеры по умолчанию.
 *    Отправитель берётся из кабинета НП, к которому относится ключ (меню «Указать ключ Новой Почты»).
 * 2. Лист «Товары НП»: товар -> описание для накладной, вес и габариты одной штуки.
 * 3. В «пуле» выделить заказы -> «Перенести выделенные в НП лиды».
 * 4. «Создать ТТН» - для строк «НП лидов» без ТТН: накладная отделение -> отделение, наложенный платёж = сумма,
 *    доставку платит получатель, одно место. Номер ТТН пишется в строку.
 * 5. «Печать наклеек 100×100» - PDF для Zebra по выделенным строкам (или по всем ТТН, созданным сегодня).
 */

var TTN_SHEET = 'НП лиды';
var TTN_SETTINGS = 'Настройки НП';
var TTN_GOODS = 'Товары НП';
var TTN_COLUMNS = ['Статус', 'ТТН', 'ФИО', 'Телефон', 'Товар', 'Кол-во', 'Наложка', 'Адрес НП', 'Описание',
                   'Вес, кг', 'Длина, см', 'Ширина, см', 'Высота, см', 'Call ID', 'Дата ТТН',
                   'Ref отделения', 'Ref города', 'Ref ТТН', 'ТТН в CRM'];
function TC_(name) { return TTN_COLUMNS.indexOf(name); }

var TTN_SETTINGS_ROWS = [
  ['Город отправки', '', 'например: Чайки, Бучанський (район через запятую - если сёл с таким названием несколько)'],
  ['Номер отделения отправки', '', 'например: 5'],
  ['Телефон отправителя', '', 'например: 0961234567 (должен быть у контактного лица в кабинете НП)'],
  ['Описание груза по умолчанию', 'Товар', 'если товара нет в листе «Товары НП»'],
  ['Вес по умолчанию, кг', '0.5', 'за 1 шт'],
  ['Длина по умолчанию, см', '20', 'за 1 шт'],
  ['Ширина по умолчанию, см', '15', 'за 1 шт'],
  ['Высота по умолчанию, см', '10', 'за 1 шт'],
  ['Наложенный платёж', 'контроль оплаты', '«контроль оплаты» или «післяплата» - как подключено у вашего ФОП в НП'],
  ['ID доставки «Нова Пошта» в СРМ 1', '', 'число из LP-CRM: Налаштування -> Способи доставки (ID у «Нова Пошта»)'],
  ['ID доставки «Нова Пошта» в СРМ 2', '', 'то же для второй CRM']
];

var TTN_GOODS_HEADER = ['Товар (как в пуле)', 'Описание для накладной', 'Вес 1 шт, кг', 'Длина 1 шт, см',
                        'Ширина 1 шт, см', 'Высота 1 шт, см', 'Скидка на заказ, грн', 'Название у поставщика', 'Поставщик', 'ID товара в СРМ'];
var GOODS_POSTAV_COL = 9;   // колонка I «Поставщик»
var POSTAVSHCHIKI = ['Скайлаб', 'Щербак'];   // выпадающий список в колонке «Поставщик» листа «Товары НП»
var TTN_FILL = '#fff2cc';     // жёлтым - ячейки, которые заполняете вы
var TTN_HEAD = '#d9d9d9';     // серым - заголовки

/**
 * Лист с заголовком: создаёт, а если лист уже есть, но заголовка нет (например, после сбоя) - добавляет строку
 * заголовка сверху, не трогая данные. widths: ширина колонок, hide: какие колонки спрятать.
 */
function listTtn_(name, header, widths, hide) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getMaxColumns() < header.length) sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
  var first = sh.getRange(1, 1, 1, header.length).getValues()[0].map(function (x) { return String(x).trim(); });
  if (first.join('|') !== header.join('|')) {
    var pusto = first.every(function (x) { return !x; });
    var staryiZag = first[0] && header.indexOf(first[0]) >= 0;   // заголовок прежней версии - просто обновим
    if (!pusto && !staryiZag) sh.insertRowBefore(1);              // в первой строке данные - сдвигаем вниз
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground(TTN_HEAD).setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  (widths || []).forEach(function (wd, i) { if (wd) sh.setColumnWidth(i + 1, wd); });
  (hide || []).forEach(function (col) { sh.hideColumns(col); });
  return sh;
}

/** Настройки НП: {город, отделение, телефон, описание, вес, д, ш, в}. Создаёт/дополняет лист-шаблон. */
function nastroikiNp_() {
  var sh = listTtn_(TTN_SETTINGS, ['Параметр', 'Значение (заполните)', 'Подсказка'], [240, 220, 460]);
  var est = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { est[String(r[0]).trim()] = true; });
  var dobavit = TTN_SETTINGS_ROWS.filter(function (r) { return !est[r[0]]; });
  if (dobavit.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 2, dobavit.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, dobavit.length, 3).setValues(dobavit);
  }
  var n = sh.getLastRow() - 1;
  sh.getRange(2, 1, n, 1).setFontWeight('bold');
  sh.getRange(2, 2, n, 1).setBackground(TTN_FILL);
  sh.getRange(2, 3, n, 1).setFontColor('#777777');
  var v = {};
  sh.getRange(2, 1, n, 2).getValues().forEach(function (r) { v[String(r[0]).trim()] = String(r[1]).trim(); });
  return {
    gorod: v['Город отправки'] || '', otdelenie: (v['Номер отделения отправки'] || '').replace(/\D/g, ''),
    telefon: v['Телефон отправителя'] || '', opisanie: v['Описание груза по умолчанию'] || 'Товар',
    ves: chislo_(v['Вес по умолчанию, кг']) || 0.5, d: chislo_(v['Длина по умолчанию, см']) || 20,
    sh: chislo_(v['Ширина по умолчанию, см']) || 15, v: chislo_(v['Высота по умолчанию, см']) || 10,
    kontrol: !/післяплат|послеплат|наложен/i.test(v['Наложенный платёж'] || 'контроль оплаты'),
    dostavkaCrm: [String(v['ID доставки «Нова Пошта» в СРМ 1'] || '').replace(/\D/g, ''),
                  String(v['ID доставки «Нова Пошта» в СРМ 2'] || '').replace(/\D/g, '')]
  };
}

/** Товары НП: ключ товара -> {opisanie, ves, d, sh, v}. Новые товары из колонки H «Статистики» дописываются сами. */
function tovaryNp_() {
  var ss = SpreadsheetApp.getActive();
  var stary = ss.getSheetByName(TTN_GOODS);   // прежняя статистика стояла в I:K, потом в K:M - теперь там колонки товаров
  if (stary && stary.getMaxColumns() >= 11) {
    var zag = stary.getRange(1, 1, 1, Math.min(stary.getMaxColumns(), 13)).getValues()[0].map(String);
    var gde = /Статистика/.test(zag[8]) ? 8 : (zag[9] !== 'ID товара в СРМ' && /^Товар( \(|$)/.test(zag[10] || '') ? 10 : 0);
    if (gde) stary.getRange(1, gde, stary.getMaxRows(), 4).clearContent().setBackground(null).setFontWeight('normal').clearDataValidations();
  }
  var sh = listTtn_(TTN_GOODS, TTN_GOODS_HEADER, [280, 220, 110, 110, 110, 110, 120, 260, 110, 110]);
  var est = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { est[klyuchTovara_(r[0])] = true; });
  var stats = ss.getSheetByName(ZAYAVKI_CFG.STATS_SHEET);
  var names = stats && stats.getLastRow() > 1 ?
    stats.getRange(2, ZAYAVKI_CFG.COL_PICK, stats.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  var novye = names.filter(function (n) { var k = klyuchTovara_(n); if (!k || est[k]) return false; est[k] = true; return true; });
  if (novye.length) sh.getRange(sh.getLastRow() + 1, 1, novye.length, 1).setValues(novye.map(function (n) { return [n]; }));
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 2, sh.getLastRow() - 1, TTN_GOODS_HEADER.length - 1).setBackground(TTN_FILL);
    sh.getRange(2, GOODS_POSTAV_COL, sh.getLastRow() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(POSTAVSHCHIKI, true).setAllowInvalid(true).build());
  }
  var out = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, TTN_GOODS_HEADER.length).getValues().forEach(function (r) {
      var key = klyuchTovara_(r[0]);
      if (!key) return;
      out[key] = { opisanie: String(r[1]).trim(), ves: chislo_(r[2]), d: chislo_(r[3]), sh: chislo_(r[4]), v: chislo_(r[5]),
                   skidka: chislo_(r[6]), uPostav: String(r[7]).trim(), postav: String(r[8]).trim(),
                   idCrm: String(r[9]).trim() };
    });
  }
  return out;
}

/** «Петренко Іван Сергійович» -> {last, first, middle}. */
/** Лист «НП лиды» с заголовками и оформлением; служебные Ref-колонки спрятаны. */
function listNpLidy_() {
  var widths = [230, 130, 190, 110, 220, 60, 80, 340, 140, 70, 70, 70, 70, 110, 90, 90, 90, 90, 170];
  var hide = [TC_('Ref отделения') + 1, TC_('Ref города') + 1, TC_('Ref ТТН') + 1];
  var sh = listTtn_(TTN_SHEET, TTN_COLUMNS, widths, hide);
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, TTN_COLUMNS.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.getRange(2, TC_('Описание') + 1, sh.getLastRow() - 1, 5).setBackground(TTN_FILL);   // можно поправить перед ТТН
    sh.getRange(2, TC_('ФИО') + 1, sh.getLastRow() - 1, 1).setBackground(TTN_FILL);
  }
  return sh;
}

function razobratFio_(fio) {
  var p = String(fio || '').replace(/[^A-Za-zА-Яа-яІіЇїЄєҐґ'’ʼ\- ]/g, ' ').replace(/\s+/g, ' ').trim().split(' ');
  if (p.length >= 3) return { last: p[0], first: p[1], middle: p.slice(2).join(' ') };
  if (p.length === 2) return { last: p[0], first: p[1], middle: '' };
  return { last: '', first: p[0] || '', middle: '' };
}

/** Телефон в формате НП: 380XXXXXXXXX или '' если не похож на украинский мобильный. */
function telefonNp_(t) {
  var d = String(t || '').replace(/\D/g, '');
  if (d.length === 9) d = '380' + d;
  if (d.length === 10 && d[0] === '0') d = '38' + d;
  if (d.length === 11 && d.indexOf('80') === 0) d = '3' + d;
  return /^380\d{9}$/.test(d) ? d : '';
}

/** Из «Кривий Ріг, Відділення №27 (до 30 кг): вул. ...» -> {gorod, nomer}. */
function razobratAdresNp_(adr) {
  var a = String(adr || '').replace(/^≈\s*/, '').trim();
  var m = a.match(/№\s*(\d+)/);
  return { gorod: a.split(',')[0].trim(), nomer: m ? m[1] : '', tekst: a };
}

/** Находит Ref отделения и города для адресов в формате НП (сверяя текст адреса). */
function refyOtdeleniy_(adresa) {
  var calls = adresa.map(function (a) {
    var r = razobratAdresNp_(a);
    return { model: 'AddressGeneral', method: 'getWarehouses', props: { CityName: r.gorod, WarehouseId: r.nomer, Limit: '20' } };
  });
  function vybrat(list, r) {
    var tochno = list.filter(function (x) { return (x.CityDescription + ', ' + x.Description) === r.tekst; });
    if (tochno.length) return tochno[0];
    var poNomeru = list.filter(function (x) { return String(x.Number) === r.nomer; });
    var vGorode = poNomeru.filter(function (x) { return String(x.CityDescription).toLowerCase() === r.gorod.toLowerCase(); });
    if (vGorode.length === 1) return vGorode[0];
    return poNomeru.length === 1 ? poNomeru[0] : null;
  }
  var resps = npZaprosy_(calls);
  var out = resps.map(function (resp, i) {
    var w = vybrat((resp && resp.data) || [], razobratAdresNp_(adresa[i]));
    return w ? { ref: w.Ref, city: w.CityRef } : null;
  });
  // не нашлось по названию города - ищем город в НП (его код), затем отделение в нём
  var ostalos = [];
  out.forEach(function (x, i) { if (!x) ostalos.push(i); });
  if (ostalos.length) {
    var goroda = npZaprosy_(ostalos.map(function (i) {
      return { model: 'Address', method: 'getCities', props: { FindByString: razobratAdresNp_(adresa[i]).gorod, Limit: '5' } };
    }));
    var shag2 = [], kuda = [];
    goroda.forEach(function (resp, k) {
      var i = ostalos[k], r = razobratAdresNp_(adresa[i]);
      ((resp && resp.data) || []).slice(0, 5).forEach(function (c) {
        shag2.push({ model: 'AddressGeneral', method: 'getWarehouses', props: { CityRef: c.Ref, WarehouseId: r.nomer, Limit: '20' } });
        kuda.push(i);
      });
    });
    npZaprosy_(shag2).forEach(function (resp, k) {
      var i = kuda[k];
      if (out[i]) return;
      var w = vybrat((resp && resp.data) || [], razobratAdresNp_(adresa[i]));
      if (w) out[i] = { ref: w.Ref, city: w.CityRef };
    });
  }
  return out;
}

/** Кнопка «Перенести выделенные в НП лиды»: выделенные строки пула -> лист «НП лиды». */
function perenestiVNpLidy() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || ss.getActiveSheet().getName() !== pool.getName()) {
    ui.alert('Откройте лист «' + ZAYAVKI_CFG.POOL_SHEET + '», выделите строки заказов и нажмите кнопку ещё раз.');
    return;
  }
  podgotovitPul_(pool);
  var nomera = vydelennyeStroki_(pool).filter(function (r) { return r <= pool.getLastRow(); });
  if (!nomera.length) { ui.alert('Выделите в пуле строки заказов (можно по номерам строк, несколько групп - с Ctrl).'); return; }
  perenestiStroki_(pool, nomera, null);
}

function vypulitDlyaSkailab() { vypulitDlyaPostavshchika_(POSTAVSHCHIKI[0]); }
function vypulitDlyaShcherbak() { vypulitDlyaPostavshchika_(POSTAVSHCHIKI[1]); }

/** Все заказы пула, у которых основной товар помечен в «Товары НП» этим поставщиком -> «НП лиды». */
function vypulitDlyaPostavshchika_(postav) {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (!pool || pool.getLastRow() < 2) { ui.alert('Лист «' + ZAYAVKI_CFG.POOL_SHEET + '» пуст.'); return; }
  podgotovitPul_(pool);
  var tovary = tovaryNp_();
  var nash = function (t) { return t && t.postav && t.postav.toLowerCase() === postav.toLowerCase(); };
  if (!Object.keys(tovary).some(function (k) { return nash(tovary[k]); })) {
    ui.alert('В листе «' + TTN_GOODS + '» нет товаров с поставщиком «' + postav + '». Отметьте их в колонке «Поставщик».');
    return;
  }
  var vse = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();
  var nomera = [];
  vse.forEach(function (row, i) {
    var items = razobratTovary_(row[P_('Товар')]);
    if (items.length && nash(tovary[klyuchTovara_(items[0].name)])) nomera.push(i + 2);
  });
  if (!nomera.length) { ss.toast('В пуле нет заказов с товарами «' + postav + '».', postav, 8); return; }
  perenestiStroki_(pool, nomera, postav);
}

/** Перенос строк пула (номера строк листа) в «НП лиды». postav - для подписи в сообщении. */
function perenestiStroki_(pool, nomera, postav) {
  var ss = SpreadsheetApp.getActive();
  var dest = listNpLidy_();
  var est = {};
  if (dest.getLastRow() > 1) {
    dest.getRange(2, TC_('Call ID') + 1, dest.getLastRow() - 1, 1).getValues().forEach(function (r) { est[String(r[0])] = true; });
  }
  var nast = nastroikiNp_(), tovary = tovaryNp_();
  var add = [], propuski = [];
  var vse = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();   // одно чтение вместо сотен
  nomera.forEach(function (r) {
    var row = vse[r - 2];
    var cid = String(row[P_('Call ID')]).trim();
    if (!cid) return;
    if (est[cid]) { propuski.push(r + ' (уже в НП лидах)'); return; }
    var adr = String(row[P_('Адрес НП')] || '').trim();
    if (!adr || adresNeNp_(adr)) { propuski.push(r + ' (адрес не подтверждён)'); return; }
    var items = razobratTovary_(row[P_('Товар')]);
    var qty = chislo_(row[P_('Шт')]) || (items.length ? items[0].qty : 1);   // «Шт» уточняет проверка ИИ (с подарочными)
    var t = items.length ? tovary[klyuchTovara_(items[0].name)] : null;
    if (t && !(t.ves && t.d && t.sh && t.v)) t = null;   // габариты товара не заполнены - берём по умолчанию
    var one = {
      opisanie: (items.length && (tovary[klyuchTovara_(items[0].name)] || {}).opisanie) || nast.opisanie,
      ves: (t && t.ves) || nast.ves,
      d: (t && t.d) || nast.d, sh: (t && t.sh) || nast.sh, v: (t && t.v) || nast.v
    };
    var stroka = TTN_COLUMNS.map(function () { return ''; });
    stroka[TC_('Статус')] = t ? '' : 'проверьте габариты (у товара не заполнены «Товары НП»)';
    stroka[TC_('ФИО')] = row[P_('ФИО')];
    stroka[TC_('Телефон')] = row[P_('Телефон')];
    stroka[TC_('Товар')] = row[P_('Товар')];
    stroka[TC_('Кол-во')] = qty;
    // скидка из «Товары НП» (колонка G) - один раз на заказ, а не на каждую штуку
    var tAll = items.length ? tovary[klyuchTovara_(items[0].name)] : null;
    var skidka = (tAll && tAll.skidka) || 0;
    stroka[TC_('Наложка')] = Math.max(chislo_(row[P_('Сумма')]) - skidka, 0);
    stroka[TC_('Адрес НП')] = adr.replace(/^≈\s*/, '');
    stroka[TC_('Описание')] = one.opisanie;
    stroka[TC_('Вес, кг')] = Math.round(one.ves * qty * 1000) / 1000;
    stroka[TC_('Длина, см')] = one.d;
    stroka[TC_('Ширина, см')] = one.sh;
    stroka[TC_('Высота, см')] = one.v * qty;   // несколько штук - друг на друга
    stroka[TC_('Call ID')] = cid;
    est[cid] = true;
    add.push(stroka);
  });
  if (add.length) {
    var start = dest.getLastRow() + 1;
    dest.getRange(start, TC_('Телефон') + 1, add.length, 1).setNumberFormat('@');
    dest.getRange(start, TC_('Call ID') + 1, add.length, 1).setNumberFormat('@');
    dest.getRange(start, 1, add.length, TTN_COLUMNS.length).setValues(add);
  }
  try { statistikaNpLidov_(); } catch (e) {}
  if (postav) {   // при выпуле по поставщику «уже в НП лидах» - норма, не перечисляем их
    var uzhe = propuski.filter(function (p) { return p.indexOf('уже в НП') >= 0; }).length;
    var drugie = propuski.filter(function (p) { return p.indexOf('уже в НП') < 0; });
    SpreadsheetApp.getUi().alert('«' + postav + '»', 'Перенесено в «' + TTN_SHEET + '»: ' + add.length +
      (uzhe ? '\nУже были в «' + TTN_SHEET + '»: ' + uzhe : '') +
      (drugie.length ? '\nПропущены строки пула: ' + drugie.join(', ') : ''), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  ss.toast('Перенесено в «' + TTN_SHEET + '»: ' + add.length + (propuski.length ? '. Пропущены строки: ' + propuski.join(', ') : ''),
           TTN_SHEET, 10);
}

/** Отправитель из кабинета НП (по ключу) + отделение отправки. Запоминается до смены настроек. */
function otpravitelNp_(nast) {
  var props = PropertiesService.getScriptProperties();
  var sig = ['v2', nast.gorod, nast.otdelenie, nast.telefon].join('|');
  var saved = props.getProperty('NP_SENDER');
  if (saved) { try { var s = JSON.parse(saved); if (s.sig === sig) return s; } catch (e) {} }
  if (!nast.gorod || !nast.otdelenie || !telefonNp_(nast.telefon)) {
    throw new Error('заполните лист «' + TTN_SETTINGS + '»: город, номер отделения и телефон отправителя');
  }
  var r1 = npZaprosy_([
    { model: 'Counterparty', method: 'getCounterparties', props: { CounterpartyProperty: 'Sender', Page: '1' } },
    { model: 'AddressGeneral', method: 'getWarehouses', props: { CityName: nast.gorod.split(',')[0].trim(), WarehouseId: nast.otdelenie, Limit: '30' } }
  ]);
  var cp = r1[0] && r1[0].data && r1[0].data[0];
  if (!cp) throw new Error('НП не вернула отправителя по ключу: ' + JSON.stringify((r1[0] || {}).errors || []));
  var imya = nast.gorod.split(',')[0].trim().toLowerCase();
  var podskazka = nast.gorod.split(',').slice(1).join(' ').toLowerCase().replace(/район|р-н|область|обл\.?/g, ' ').trim();
  var whs = ((r1[1] && r1[1].data) || []).filter(function (w) {
    return String(w.Number) === nast.otdelenie && String(w.CityDescription).toLowerCase().indexOf(imya) === 0;
  });
  if (podskazka) {
    var st = podskazka.split(/\s+/)[0].slice(0, 5);
    var sPodsk = whs.filter(function (w) {
      return (String(w.SettlementRegionsDescription) + ' ' + String(w.SettlementAreaDescription) + ' ' + String(w.CityDescription))
        .toLowerCase().indexOf(st) >= 0;
    });
    if (sPodsk.length) whs = sPodsk;
  }
  if (!whs.length) {
    // сёла часто не находятся по названию напрямую: ищем населённый пункт (с районом), затем отделение в нём
    var rs = npZaprosy_([{ model: 'Address', method: 'searchSettlements', props: { CityName: nast.gorod.split(',')[0].trim(), Limit: '50', Page: '1' } }]);
    var addrs = ((rs[0] && rs[0].data && rs[0].data[0] && rs[0].data[0].Addresses) || []).filter(function (a) {
      return a.DeliveryCity && String(a.MainDescription).toLowerCase() === imya;
    });
    if (podskazka) {
      var st2 = podskazka.split(/\s+/)[0].slice(0, 5);
      var sP = addrs.filter(function (a) { return (String(a.Region) + ' ' + String(a.Area) + ' ' + String(a.Present)).toLowerCase().indexOf(st2) >= 0; });
      if (sP.length) addrs = sP;
    }
    var rw = npZaprosy_(addrs.slice(0, 10).map(function (a) {
      return { model: 'AddressGeneral', method: 'getWarehouses', props: { CityRef: a.DeliveryCity, WarehouseId: nast.otdelenie, Limit: '20' } };
    }));
    rw.forEach(function (resp, k) {
      ((resp && resp.data) || []).forEach(function (w) {
        if (String(w.Number) !== nast.otdelenie) return;
        w.SettlementRegionsDescription = w.SettlementRegionsDescription || addrs[k].Region;
        w.SettlementAreaDescription = w.SettlementAreaDescription || addrs[k].Area;
        whs.push(w);
      });
    });
  }
  if (!whs.length) {
    // последний вариант: адреса отправки, сохранённые у отправителя в кабинете НП
    var ra = npZaprosy_([{ model: 'Counterparty', method: 'getCounterpartyAddresses', props: { Ref: cp.Ref, CounterpartyProperty: 'Sender' } }]);
    ((ra[0] && ra[0].data) || []).forEach(function (a) {
      var txt = String(a.Description || '') + ' ' + String(a.CityDescription || '');
      if (new RegExp('№\\s*' + nast.otdelenie + '\\b').test(txt) && txt.toLowerCase().indexOf(imya) >= 0) {
        whs.push({ Ref: a.Ref, CityRef: a.CityRef, CityDescription: a.CityDescription || imya, Number: nast.otdelenie });
      }
    });
  }
  var raznye = {};
  whs.forEach(function (w) { raznye[w.CityRef] = w; });
  if (!whs.length) throw new Error('не найдено отделение отправки: ' + nast.gorod + ', №' + nast.otdelenie +
    '. Проверьте название населённого пункта (по-украински, как в НП) и номер отделения.');
  if (Object.keys(raznye).length > 1) {
    throw new Error('населённых пунктов «' + nast.gorod + '» с отделением №' + nast.otdelenie + ' несколько: ' +
      Object.keys(raznye).map(function (k) { var w = raznye[k]; return w.CityDescription + ' (' + (w.SettlementRegionsDescription || '') +
        ' р-н, ' + (w.SettlementAreaDescription || '') + ' обл.)'; }).join('; ') +
      '. Допишите район в «Город отправки» через запятую, например «Чайки, Бучанський».');
  }
  var wh = whs[0];
  var r2 = npZaprosy_([{ model: 'Counterparty', method: 'getCounterpartyContactPersons', props: { Ref: cp.Ref, Page: '1' } }]);
  var tel = telefonNp_(nast.telefon);
  var contacts = (r2[0] && r2[0].data) || [];
  var contact = contacts.filter(function (c) { return telefonNp_(c.Phones) === tel; })[0] || contacts[0];
  if (!contact) throw new Error('у отправителя в кабинете НП нет контактного лица');
  var s2 = { sig: sig, sender: cp.Ref, contact: contact.Ref, phone: tel, city: wh.CityRef, address: wh.Ref };
  props.setProperty('NP_SENDER', JSON.stringify(s2));
  return s2;
}

/** Кнопка «Создать ТТН»: для всех строк «НП лидов» без ТТН. */
function sozdatTtn() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sh = listNpLidy_();
  if (sh.getLastRow() < 2) { ui.alert('Лист «' + TTN_SHEET + '» пуст. Сначала перенесите заказы из пула.'); return; }
  var nast = nastroikiNp_();
  var otpr;
  try { otpr = otpravitelNp_(nast); } catch (e) { ui.alert('Не получилось: ' + e.message); return; }

  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, TTN_COLUMNS.length).getValues();
  var todo = [];
  rows.forEach(function (r, i) { if (String(r[TC_('Call ID')]).trim() && !String(r[TC_('ТТН')]).trim()) todo.push({ i: i, r: r }); });
  if (!todo.length) { ss.toast('Все строки уже с ТТН.', TTN_SHEET, 6); return; }
  if (ui.alert('Создание ТТН', 'Создать ТТН: ' + todo.length + '?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  function status(z, text) { sh.getRange(z.i + 2, TC_('Статус') + 1).setValue(String(text).slice(0, 300)); }

  // 1. отделения получателей (Ref) - если ещё не найдены
  var nuzhno = todo.filter(function (z) { return !z.r[TC_('Ref отделения')]; });
  if (nuzhno.length) {
    var refs = refyOtdeleniy_(nuzhno.map(function (z) { return z.r[TC_('Адрес НП')]; }));
    nuzhno.forEach(function (z, k) {
      if (!refs[k]) return;
      z.r[TC_('Ref отделения')] = refs[k].ref;
      z.r[TC_('Ref города')] = refs[k].city;
      sh.getRange(z.i + 2, TC_('Ref отделения') + 1, 1, 2).setValues([[refs[k].ref, refs[k].city]]);
    });
  }
  var gotovy = [];
  todo.forEach(function (z) {
    var fio = razobratFio_(z.r[TC_('ФИО')]);
    var tel = telefonNp_(z.r[TC_('Телефон')]);
    if (!z.r[TC_('Ref отделения')]) return status(z, 'ОШИБКА: отделение не найдено в НП - поправьте «Адрес НП»');
    if (!tel) return status(z, 'ОШИБКА: телефон не похож на украинский мобильный');
    if (!fio.last || !fio.first) return status(z, 'ОШИБКА: нужны фамилия и имя получателя - допишите в колонке «ФИО»');
    z.fio = fio; z.tel = tel;
    gotovy.push(z);
  });

  // 2. получатели (частные лица)
  var r1 = npZaprosy_(gotovy.map(function (z) {
    return { model: 'Counterparty', method: 'save', props: {
      FirstName: z.fio.first, MiddleName: z.fio.middle, LastName: z.fio.last, Phone: z.tel, Email: '',
      CounterpartyType: 'PrivatePerson', CounterpartyProperty: 'Recipient' } };
  }));
  var dlyaTtn = [];
  gotovy.forEach(function (z, k) {
    var d = r1[k] && r1[k].data && r1[k].data[0];
    var contact = d && d.ContactPerson && d.ContactPerson.data && d.ContactPerson.data[0];
    if (!d || !contact) return status(z, 'ОШИБКА НП (получатель): ' + ((r1[k] || {}).errors || []).join('; '));
    z.recipient = d.Ref; z.contact = contact.Ref;
    dlyaTtn.push(z);
  });

  /** «Товар: Сюрприз Бокс*2шт;» - поле «Додаткова інформація», печатается на наклейке.
   *  Только основной товар; количество - из колонки «Кол-во» листа «НП лиды» (его можно поправить руками). */
  var tovaryDlyaNakleek = tovaryNp_();
  function sostavDlyaNakleiki(r) {
    var items = razobratTovary_(r[TC_('Товар')]);
    var name = items.length ? items[0].name : String(r[TC_('Товар')] || '').replace(/\(\s*\d+\s*шт\.?\s*\)/i, '').trim();
    var t = tovaryDlyaNakleek[klyuchTovara_(name)];
    if (t && t.uPostav) name = t.uPostav;   // название, как у поставщика (колонка H «Товары НП»)
    var qty = chislo_(r[TC_('Кол-во')]) || (items.length ? items[0].qty : 1);
    return ('Товар: ' + name + '*' + qty + 'шт;').slice(0, 100);   // у НП ограничение на длину поля
  }

  // 3. накладные
  var data = Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM.yyyy');
  var r2 = npZaprosy_(dlyaTtn.map(function (z) {
    var r = z.r;
    var ves = chislo_(r[TC_('Вес, кг')]) || nast.ves;
    var d = chislo_(r[TC_('Длина, см')]) || nast.d, w = chislo_(r[TC_('Ширина, см')]) || nast.sh, h = chislo_(r[TC_('Высота, см')]) || nast.v;
    var sum = Math.round(chislo_(r[TC_('Наложка')]));
    return { model: 'InternetDocument', method: 'save', props: {
      PayerType: 'Recipient', PaymentMethod: 'Cash', DateTime: data, CargoType: 'Parcel', ServiceType: 'WarehouseWarehouse',
      SeatsAmount: '1', Weight: String(ves), VolumeGeneral: String(Math.max(d * w * h / 1e6, 0.0004)),
      Description: String(r[TC_('Описание')] || nast.opisanie), Cost: String(sum || 1),
      AdditionalInformation: sostavDlyaNakleiki(r),
      CitySender: otpr.city, Sender: otpr.sender, SenderAddress: otpr.address, ContactSender: otpr.contact, SendersPhone: otpr.phone,
      CityRecipient: String(r[TC_('Ref города')]), Recipient: z.recipient, RecipientAddress: String(r[TC_('Ref отделения')]),
      ContactRecipient: z.contact, RecipientsPhone: z.tel,
      // «Контроль оплати» (для ФОП/юрлиц) или обычная «Післяплата» - смотря что подключено у отправителя
      AfterpaymentOnGoodsCost: nast.kontrol ? String(sum) : undefined,
      BackwardDeliveryData: nast.kontrol ? undefined : [{ PayerType: 'Recipient', CargoType: 'Money', RedeliveryString: String(sum) }] } };
  }));
  var ok = 0, oshibki = 0;
  dlyaTtn.forEach(function (z, k) {
    var d = r2[k] && r2[k].data && r2[k].data[0];
    if (d && d.IntDocNumber) {
      sh.getRange(z.i + 2, TC_('ТТН') + 1).setNumberFormat('@').setValue(String(d.IntDocNumber));
      sh.getRange(z.i + 2, TC_('Дата ТТН') + 1).setValue(data);
      sh.getRange(z.i + 2, TC_('Ref ТТН') + 1).setValue(d.Ref);
      status(z, 'OK, доставка ≈ ' + (d.CostOnSite || '?') + ' грн');
      ok++;
    } else {
      status(z, 'ОШИБКА НП: ' + ((r2[k] || {}).errors || []).join('; '));
      oshibki++;
    }
  });
  var bez = todo.length - ok - oshibki;
  try { statistikaNpLidov_(); } catch (e) {}
  ui.alert('Создано ТТН: ' + ok + (oshibki + bez ? ', с ошибкой: ' + (oshibki + bez) + ' (причины в колонке «Статус»)' : '') + '.');
}

/** Кнопка «Печать наклеек 100×100»: по выделенным строкам «НП лидов», иначе - все ТТН, созданные сегодня. */
function pechatNakleek() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sh = listNpLidy_();
  if (sh.getLastRow() < 2) { ui.alert('Лист «' + TTN_SHEET + '» пуст.'); return; }
  var key = PropertiesService.getScriptProperties().getProperty('NP_KEY');
  if (!key) { ui.alert('Укажите ключ: меню «Заявки» -> «Указать ключ Новой Почты».'); return; }
  try { statistikaNpLidov_(); } catch (e) {}   // статистика в «Товары НП» (K:M) - по текущим «НП лидам»
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, TTN_COLUMNS.length).getValues();
  var vybor = ss.getActiveSheet().getName() === TTN_SHEET ? vydelennyeStroki_(sh) : [];
  var segodnya = Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM.yyyy');
  var nomera = [], vydeleno = [];
  rows.forEach(function (r, i) {
    var ttn = String(r[TC_('ТТН')]).trim();
    if (ttn && vybor.indexOf(i + 2) >= 0) vydeleno.push(ttn);
  });
  if (vydeleno.length) nomera = vydeleno;   // выделены строки с ТТН - печатаем их
  else {
    // иначе - ТТН последнего дня, за который они создавались в этом листе
    var daty = rows.map(function (r) {
      return r[TC_('Дата ТТН')] instanceof Date ? Utilities.formatDate(r[TC_('Дата ТТН')], ZAYAVKI_CFG.TZ, 'dd.MM.yyyy')
                                                : String(r[TC_('Дата ТТН')]).trim();
    });
    var posl = '';
    daty.forEach(function (d) { var k = d.split('.').reverse().join(''); if (d && k > posl.split('.').reverse().join('')) posl = d; });
    rows.forEach(function (r, i) { if (String(r[TC_('ТТН')]).trim() && daty[i] === posl) nomera.push(String(r[TC_('ТТН')]).trim()); });
  }
  if (!nomera.length) { ui.alert('Нет ТТН для печати: выделите строки с ТТН.'); return; }
  var k = encodeURIComponent(key), list = nomera.join(','), base = 'https://my.novaposhta.ua/orders/';
  // формат Zebra 100×100: на наклейке печатается и «Товар: ...» из «Додаткової інформації»
  var url = base + 'printMarking100x100/orders/' + list + '/type/pdf/zebra/zebra/apiKey/' + k;
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px"><p>Наклеек: <b>' + nomera.length + '</b></p>' +
    '<p><a href="' + url + '" target="_blank">Открыть PDF для печати (Zebra 100×100)</a></p>' +
    '<p style="color:#888;font-size:12px">Ссылка содержит ключ НП - не пересылайте её.</p></div>').setWidth(420).setHeight(170);
  ui.showModalDialog(html, 'Наклейки Новой Почты');
}

function otpravitTtnVCrm1() { otpravitTtnVCrm_(CRM_LIST[0]); }
function otpravitTtnVCrm2() { otpravitTtnVCrm_(CRM_LIST[1]); }

/**
 * Кнопки «Отправить выделенные с ТТН в СРМ 1 / 2»: выделенные строки «НП лидов» с ТТН создаются заказами в CRM
 * (как из пула: товар, цена, допы в комментарии), номер ТТН передаётся в поле ttn заказа.
 * Если заказ уже был отправлен в CRM из пула, создаётся не новый заказ, а в старый дописывается ТТН (updateOrder).
 */
function otpravitTtnVCrm_(crm) {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sh = listNpLidy_();
  if (ss.getActiveSheet().getName() !== TTN_SHEET || sh.getLastRow() < 2) {
    ui.alert('Откройте лист «' + TTN_SHEET + '», выделите строки с ТТН и нажмите кнопку ещё раз.');
    return;
  }
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty(crm.keyProp), url = adresCrm_(crm);
  if (!key || !url) { ui.alert('Нет ключа или адреса ' + crm.name + ': меню «Заявки» -> «Указать ключ».'); return; }
  var base = url.replace(/\/api\/addNewOrder\.html$/, '');
  var w = TTN_COLUMNS.length;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues();
  var vybor = vydelennyeStroki_(sh).filter(function (r) { return r <= sh.getLastRow(); });
  var todo = vybor.filter(function (r) {
    var x = rows[r - 2];
    return String(x[TC_('ТТН')]).trim() && !/^OK/.test(String(x[TC_('ТТН в CRM')]));
  });
  if (!todo.length) { ui.alert('Выделите строки, у которых есть ТТН и ещё нет «OK» в «ТТН в CRM».'); return; }
  if (ui.alert(crm.name, 'Отправить в ' + crm.name + ' заказов с ТТН: ' + todo.length + '?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // строки пула по Call ID: товар, ID товара, допы и т.п. берём оттуда - как при обычной отправке
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  podgotovitPul_(pool);
  var poolRows = pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues();
  var poId = {};
  poolRows.forEach(function (r, i) { poId[String(r[P_('Call ID')]).trim()] = i; });
  var picks = spisokPula_().picks;
  var stamp = Utilities.formatDate(new Date(), ZAYAVKI_CFG.TZ, 'dd.MM HH:mm');
  var ok = 0, err = 0;

  function otvet(resp) {
    var text = resp.getContentText(), data = null;
    try { data = JSON.parse(text); } catch (e) {}
    if (data && data.status === 'ok') return '';
    var msg = data && data.message ? (typeof data.message === 'string' ? data.message : JSON.stringify(data.message)) : text;
    return '(' + resp.getResponseCode() + '): ' + String(msg).replace(/\s+/g, ' ').slice(0, 200);
  }

  todo.forEach(function (r) {
    var x = rows[r - 2], cid = String(x[TC_('Call ID')]).trim(), ttn = String(x[TC_('ТТН')]).trim(), res;
    var pi = poId[cid];
    try {
      if (pi === undefined) throw new Error('заказа нет в пуле (Call ID ' + cid + ')');
      var prow = poolRows[pi];
      var uzhe = String(prow[P_('Статус CRM')]).match(/^OK\s+СРМ\s*(\d)/);
      if (uzhe) {
        // заказ уже в CRM - дописываем ТТН
        var c2 = CRM_LIST[Number(uzhe[1]) - 1];
        var resp = UrlFetchApp.fetch(adresCrm_(c2).replace(/addNewOrder/, 'updateOrder'), { method: 'post', muteHttpExceptions: true,
          payload: { key: props.getProperty(c2.keyProp), order_id: cid, ttn: ttn } });
        var e1 = otvet(resp);
        if (e1) throw new Error('заказ уже был в ' + c2.name + ', ТТН дописать не удалось ' + e1);
        res = 'OK ТТН дописана в ' + c2.name + ' ' + stamp;
      } else {
        var prow2 = prow.slice();
        prow2[P_('Сумма')] = chislo_(x[TC_('Наложка')]) || prow[P_('Сумма')];   // цена в CRM = наложка в ТТН (со скидкой)
        var zakaz = zakazDlyaCrm_(prow2, picks, key, crm);
        if (!zakaz) throw new Error('нет ID товара для ' + crm.name + ' (колонка «ID товара ' + crm.name + '» в пуле)');
        zakaz.ttn = ttn;   // номер ТТН в поле ТТН заказа
        var dost = nastroikiNp_().dostavkaCrm[crm.idIndex];
        if (dost) zakaz.delivery = dost;   // способ доставки «Нова Пошта» в этой CRM
        // адрес и ФИО - как в «НП лидах» (там их могли поправить перед созданием ТТН)
        zakaz.bayer_name = String(x[TC_('ФИО')] || zakaz.bayer_name);
        zakaz.delivery_adress = String(x[TC_('Адрес НП')] || zakaz.delivery_adress);
        var resp2 = UrlFetchApp.fetch(url, { method: 'post', payload: zakaz, muteHttpExceptions: true });
        var e2 = otvet(resp2);
        if (e2) throw new Error(e2);
        pool.getRange(pi + 2, P_('Статус CRM') + 1).setValue('OK ' + crm.name + ' ' + stamp + ' (с ТТН)');
        res = 'OK ' + crm.name + ' ' + stamp;
      }
      ok++;
    } catch (e) {
      res = 'ОШИБКА ' + crm.name + ': ' + String(e.message || e).slice(0, 250);
      err++;
    }
    sh.getRange(r, TC_('ТТН в CRM') + 1).setValue(res);
  });
  ui.alert(crm.name + ': отправлено ' + ok + (err ? ', с ошибкой ' + err + ' (причины в «ТТН в CRM»)' : '') + '.');
}

/**
 * Статистика по листу «НП лиды» - справа в листе «Товары НП» (колонки I:K): товар, заказов, штук.
 * Считается по всем строкам, которые сейчас лежат в «НП лидах» (товар - из колонки «Товар», штуки - из «Кол-во»).
 * Обновляется при каждой кнопке НП и отдельной кнопкой «Обновить статистику НП лидов».
 */
function statistikaNpLidov_() {
  var ss = SpreadsheetApp.getActive();
  var lidy = ss.getSheetByName(TTN_SHEET);
  var tov = ss.getSheetByName(TTN_GOODS);
  if (!lidy || !tov) return;
  var col = TTN_GOODS_HEADER.length + 2;   // L (K - пустой разделитель после «ID товара в СРМ» в J)
  var tovary = tovaryNp_();
  if (tov.getMaxColumns() < col + 2) tov.insertColumnsAfter(tov.getMaxColumns(), col + 2 - tov.getMaxColumns());
  var stats = {}, poryadok = [], zakazov = 0, shtuk = 0, daty = [];
  if (lidy.getLastRow() > 1) {
    lidy.getRange(2, 1, lidy.getLastRow() - 1, TTN_COLUMNS.length).getValues().forEach(function (r) {
      if (!String(r[TC_('Call ID')]).trim()) return;
      var items = razobratTovary_(r[TC_('Товар')]);
      var name = items.length ? items[0].name : String(r[TC_('Товар')] || '').trim() || '(без товара)';
      var t = tovary[klyuchTovara_(name)];
      if (t && t.uPostav) name = t.uPostav;   // в статистике - название, как у поставщика
      var qty = chislo_(r[TC_('Кол-во')]) || (items.length ? items[0].qty : 1);
      var key = klyuchTovara_(name);
      if (!stats[key]) { stats[key] = { name: name, z: 0, q: 0 }; poryadok.push(key); }
      stats[key].z++; stats[key].q += qty;
      zakazov++; shtuk += qty;
      var d = r[TC_('Дата ТТН')];
      d = d instanceof Date ? Utilities.formatDate(d, ZAYAVKI_CFG.TZ, 'dd.MM.yyyy') : String(d || '').trim();
      if (d) daty.push(d);
    });
  }
  var rows = poryadok.map(function (k) { return [stats[k].name, stats[k].z, stats[k].q]; });
  rows.sort(function (a, b) { return b[2] - a[2] || b[1] - a[1]; });
  var sortD = daty.map(function (d) { return d.split('.').reverse().join(''); }).sort();
  var period = sortD.length ? [sortD[0], sortD[sortD.length - 1]].map(function (k) {
    return k.slice(6, 8) + '.' + k.slice(4, 6) + '.' + k.slice(0, 4); }).filter(function (v, i, a) { return a.indexOf(v) === i; }).join(' - ') : '';

  tov.getRange(1, col - 1, tov.getMaxRows(), 4).clearContent().setBackground(null).setFontWeight('normal');   // K:N
  tov.getRange(1, col, 1, 3).setValues([['Товар' + (period ? ' (ТТН ' + period + ')' : ''), 'Заказов', 'Штук']])
    .setFontWeight('bold').setBackground(TTN_HEAD).setWrap(true);
  if (rows.length) tov.getRange(2, col, rows.length, 3).setValues(rows);
  tov.getRange(rows.length + 2, col, 1, 3).setValues([['Итого', zakazov, shtuk]]).setFontWeight('bold');
  tov.setColumnWidth(col, 280); tov.setColumnWidth(col + 1, 80); tov.setColumnWidth(col + 2, 80);
}

function obnovitStatistikuNpLidov() {
  tovaryNp_();
  statistikaNpLidov_();
  SpreadsheetApp.getActive().toast('Статистика «НП лиды» обновлена (лист «' + TTN_GOODS + '», колонки L:N).', TTN_GOODS, 6);
}

// ======================= ТТН дроп-заказов: выпул из CRM и удаление в НП =======================
// Заказы в статусах «ОПТ ДРОП» / «АВЕ ОПТ» уходят на дроп-платформы, а платформа создаёт свою ТТН.
// ТТН, которую создал колл-центр, больше не нужна: выпуливаем её на лист «ТТН на удаление» и удаляем в кабинете НП.
var DEL_SHEET = 'ТТН на удаление';
var DEL_COLUMNS = ['Статус', 'ТТН', 'ID заказа СРМ', 'ФИО', 'Телефон', 'Сумма', 'Дроп', 'Дата заказа', 'Ref ТТН'];
var DROP_STATUSY = { opt: { name: 'ОПТ ДРОП', id: '69' }, ave: { name: 'АВЕ ОПТ', id: '70' } };
var CRM_OUT_KEY = 'LP_CRM_OUT_KEY';   // «Вихідний ключ API» СРМ 1 (читать заказы)

function DC_(name) { return DEL_COLUMNS.indexOf(name); }

function listUdaleniya_() {
  return listTtn_(DEL_SHEET, DEL_COLUMNS, [200, 130, 110, 200, 110, 70, 90, 130, 10], [DC_('Ref ТТН') + 1]);
}

/** Кнопка «Указать выходной ключ СРМ 1»: в LP-CRM Налаштування -> Система -> «Вихідний ключ API». */
function ukazatVyhodnoiKlyuchCrm1() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('СРМ 1: выходной ключ',
    'Вставьте «Вихідний ключ API» (LP-CRM: Налаштування -> Система). Он нужен, чтобы читать заказы из CRM.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty(CRM_OUT_KEY, key);
  SpreadsheetApp.getActive().toast('Сохранено.', 'СРМ 1', 5);
}

/** POST-запрос к API LP-CRM (СРМ 1) с выходным ключом. Возвращает разобранный JSON. */
function crmChitat_(method, params) {
  var key = PropertiesService.getScriptProperties().getProperty(CRM_OUT_KEY) || '';
  if (!key) throw new Error('Не указан выходной ключ СРМ 1 (меню «Заявки» -> «Указать выходной ключ СРМ 1»).');
  var payload = { key: key };
  Object.keys(params || {}).forEach(function (k) { payload[k] = params[k]; });
  var r = UrlFetchApp.fetch(CRM_LIST[0].defaultUrl + '/api/' + method + '.html',
                            { method: 'post', payload: payload, muteHttpExceptions: true });
  var data;
  try { data = JSON.parse(r.getContentText()); } catch (e) { throw new Error('CRM ответила не JSON: ' + korotko_(r.getContentText(), 150)); }
  if (data.status !== 'ok') {
    var msg = typeof data.message === 'string' ? data.message : JSON.stringify(data.message || data);
    throw new Error('CRM: ' + msg);
  }
  return data;
}

function vypulitTtnOptDrop() { vypulitTtnIzStatusa_(DROP_STATUSY.opt); }
function vypulitTtnAveOpt() { vypulitTtnIzStatusa_(DROP_STATUSY.ave); }

/** Заказы из статуса CRM, у которых есть ТТН, -> лист «ТТН на удаление» (без повторов). */
function vypulitTtnIzStatusa_(st) {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var ids, zakazy = [];
  try {
    ids = (crmChitat_('getOrdersIdByStatus', { status: st.id }).data || []).map(String);
    for (var i = 0; i < ids.length; i += 100) {
      var part = ids.slice(i, i + 100);
      var d = crmChitat_('getOrdersByID', { order_id: part.join(',') }).data || {};
      if (d.order_id) d = [d];                                   // один заказ приходит без обёртки
      Object.keys(d).forEach(function (k) { zakazy.push(d[k]); });
    }
  } catch (e) { ui.alert('«' + st.name + '»', e.message, ui.ButtonSet.OK); return; }

  var sh = listUdaleniya_();
  var est = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, DC_('ТТН') + 1, sh.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { var t = String(r[0]).replace(/\D/g, ''); if (t) est[t] = 1; });
  }
  var bezTtn = 0, povtor = 0, add = [];
  zakazy.forEach(function (z) {
    var ttn = String(z.ttn || '').replace(/\D/g, '');
    if (!ttn) { bezTtn++; return; }
    if (est[ttn]) { povtor++; return; }
    est[ttn] = 1;
    add.push(['', ttn, String(z.order_id || ''), z.bayer_name || '', String(z.phone || ''), chislo_(z.total) || '',
              st.name, String(z.datetime || ''), '']);
  });
  if (add.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, DC_('ТТН') + 1, add.length, 1).setNumberFormat('@');
    sh.getRange(start, DC_('ID заказа СРМ') + 1, add.length, 1).setNumberFormat('@');
    sh.getRange(start, DC_('Телефон') + 1, add.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, add.length, DEL_COLUMNS.length).setValues(add);
  }
  ui.alert('«' + st.name + '»',
    'Заказов в статусе: ' + ids.length + '\nДобавлено на лист «' + DEL_SHEET + '»: ' + add.length +
    (povtor ? '\nУже были на листе: ' + povtor : '') + (bezTtn ? '\nБез ТТН (пропущены): ' + bezTtn : ''),
    ui.ButtonSet.OK);
}

/** Кнопка «Указать ключ НП из LP-CRM»: ключ, через который LP-CRM создаёт ТТН (если он другой, чем основной). */
function ukazatKlyuchNpCrm() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Ключ НП из LP-CRM',
    'Если колл-центр создаёт ТТН через LP-CRM, вставьте API-ключ Новой Почты, который подключён в LP-CRM ' +
    '(LP-CRM: настройки службы доставки «Нова Пошта»). Им будут искаться и удаляться такие ТТН.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('NP_KEY_CRM', key);
  SpreadsheetApp.getActive().toast('Сохранено.', 'Новая Почта', 5);
}

/** Ключи НП для удаления: основной и (если указан) ключ из LP-CRM. */
function klyuchiNp_() {
  var p = PropertiesService.getScriptProperties();
  var out = [];
  [p.getProperty('NP_KEY'), p.getProperty('NP_KEY_CRM')].forEach(function (k) {
    k = String(k || '').trim();
    if (k && out.indexOf(k) < 0) out.push(k);
  });
  return out;
}

/** Кнопка «Удалить ТТН»: удаляет в кабинете НП все ТТН листа «ТТН на удаление», которые ещё не удалены. */
function udalitTtn() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var keys = klyuchiNp_();
  if (!keys.length) { ui.alert('Сначала укажите ключ Новой Почты.'); return; }
  var sh = listUdaleniya_();
  if (sh.getLastRow() < 2) { ui.alert('Лист «' + DEL_SHEET + '» пуст. Сначала выпулите ТТН из статуса.'); return; }
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, DEL_COLUMNS.length).getValues();
  var todo = [];
  data.forEach(function (r, i) {
    var ttn = String(r[DC_('ТТН')]).replace(/\D/g, '');
    if (ttn && String(r[DC_('Статус')]).indexOf('удалена') !== 0) todo.push({ i: i, ttn: ttn, ref: '', k: -1 });
  });
  if (!todo.length) { ss.toast('Нечего удалять: все ТТН на листе уже удалены.', DEL_SHEET, 6); return; }
  if (ui.alert('Удалить ТТН', 'Будет удалено ТТН в кабинете Новой Почты: ' + todo.length +
      '.\nОтменить это нельзя. Продолжить?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var refs = {};   // ТТН -> {ref, k: номер ключа}
  var np = ss.getSheetByName(TTN_SHEET);   // свои ТТН из «НП лиды» созданы основным ключом
  if (np && np.getLastRow() > 1) {
    np.getRange(2, 1, np.getLastRow() - 1, TTN_COLUMNS.length).getValues().forEach(function (r) {
      var t = String(r[TC_('ТТН')]).replace(/\D/g, ''), ref = String(r[TC_('Ref ТТН')] || '');
      if (t && ref) refs[t] = { ref: ref, k: 0 };
    });
  }
  var nenaid = function () { return todo.filter(function (z) { return !refs[z.ttn]; }); };
  var diag = [], oshibka = '';
  var tz = Session.getScriptTimeZone(), now = new Date(), fmt = function (d) { return Utilities.formatDate(d, tz, 'dd.MM.yyyy'); };

  keys.forEach(function (key, k) {
    if (!nenaid().length) return;
    // 1) список ТТН этого ключа: сегодня, вчера, затем по неделям за 30 дней («по дату» у НП не включает сам день)
    var calls = [0, 1].map(function (n) { return { DateTime: fmt(new Date(now.getTime() - n * 864e5)), GetFullList: '1' }; });
    for (var w = 0; w < 5; w++) {
      var to = new Date(now.getTime() + 864e5 - w * 7 * 864e5), from = new Date(to.getTime() - 8 * 864e5);
      calls.push({ DateTimeFrom: fmt(from), DateTimeTo: fmt(to), GetFullList: '1' });
    }
    var vse = {};
    for (var c = 0; c < calls.length && nenaid().length; c++) {
      // НП отдаёт список страницами: листаем, пока приходят новые ТТН
      for (var page = 1; page <= 50 && nenaid().length; page++) {
        var props = JSON.parse(JSON.stringify(calls[c])); props.Page = String(page);
        var resp = npOdin_('InternetDocument', 'getDocumentList', props, key);
        if (resp.success === false && resp.errors && resp.errors.length) { oshibka = String(resp.errors[0]); break; }
        var novyh = 0;
        (resp.data || []).forEach(function (d) {
          var nom = String(d.IntDocNumber || '').replace(/\D/g, '');
          if (nom && d.Ref && !vse[nom]) { vse[nom] = 1; novyh++; refs[nom] = { ref: d.Ref, k: k, otkuda: 'список' }; }
        });
        var vsego = resp.info && resp.info.totalCount ? Number(resp.info.totalCount) : 0;
        if (!novyh || (vsego && page * 100 >= vsego)) break;
      }
    }
    // 2) запасной путь: трекинг по номеру (для своих ТТН НП отдаёт и их Ref)
    var ost = nenaid();
    for (var i = 0; i < ost.length; i += 100) {
      var part = ost.slice(i, i + 100);
      var tr = npOdin_('TrackingDocument', 'getStatusDocuments',
                       { Documents: part.map(function (z) { return { DocumentNumber: z.ttn, Phone: '' }; }) }, key);
      (tr.data || []).forEach(function (d) {
        var nom = String(d.Number || '').replace(/\D/g, ''), ref = d.RefEW || d.Ref || '';
        if (nom && ref && !refs[nom]) refs[nom] = { ref: ref, k: k, otkuda: 'трекинг' };
      });
    }
    diag.push((k ? 'ключ из LP-CRM' : 'основной ключ') + ' видит ТТН: ' + Object.keys(vse).length);
  });

  var ok = 0, err = 0, net = 0;
  var den = Utilities.formatDate(new Date(), tz, 'dd.MM HH:mm');
  todo.forEach(function (z) {
    var f = refs[z.ttn];
    if (!f) {
      net++;
      zapisatUdalenie_(sh, z, '⚠ не найдена (' + diag.join(', ') + ')' +
                              (oshibka ? '; ошибка НП: ' + oshibka : ''));
      return;
    }
    z.ref = f.ref;
    // удаляем ключом, которым нашли; если не вышло - пробуем другие ключи
    var order = [f.k].concat(keys.map(function (x, j) { return j; }).filter(function (j) { return j !== f.k; }));
    var resp, e = '';
    for (var q = 0; q < order.length; q++) {
      resp = npOdin_('InternetDocument', 'delete', { DocumentRefs: z.ref }, keys[order[q]]);
      if (resp && resp.success) break;
      e = resp && resp.errors && resp.errors.length ? String(resp.errors[0]) : 'нет ответа';
    }
    if (resp && resp.success) { ok++; zapisatUdalenie_(sh, z, 'удалена ' + den); }
    else { err++; zapisatUdalenie_(sh, z, '⚠ ' + e + ' (найдена: ' + (f.otkuda || 'НП лиды') + ')'); }
  });
  ui.alert('Удаление ТТН', 'Удалено: ' + ok + (err ? '\nОшибка НП: ' + err : '') +
    (net ? '\nНе найдены: ' + net : '') + (err || net ? '\nПричины - в колонке «Статус».' : ''), ui.ButtonSet.OK);
}

/** Один запрос к API НП с паузой и повторами при «To many requests» (для удаления и списка ТТН). */
function npOdin_(model, method, props, key) {
  key = key || PropertiesService.getScriptProperties().getProperty('NP_KEY') || '';
  var data = null;
  for (var attempt = 0; attempt < 6; attempt++) {
    if (attempt) Utilities.sleep(2000 * attempt);
    try {
      var r = UrlFetchApp.fetch(NP_URL, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ apiKey: key, modelName: model, calledMethod: method, methodProperties: props }) });
      data = JSON.parse(r.getContentText());
    } catch (e) { data = { success: false, data: [], errors: ['нет ответа'] }; continue; }
    var err = data && data.errors && data.errors.length ? String(data.errors.join('; ')) : '';
    if (data && data.success !== false) break;
    if (!/many requests|нет ответа/i.test(err)) break;    // настоящая ошибка - не повторяем
  }
  Utilities.sleep(700);
  return data || { success: false, data: [], errors: ['нет ответа'] };
}

function zapisatUdalenie_(sh, z, status) {
  sh.getRange(z.i + 2, DC_('Статус') + 1).setValue(status);
  if (z.ref) sh.getRange(z.i + 2, DC_('Ref ТТН') + 1).setValue(z.ref);
}


// ======================= Лист «Заказы на СРМ»: заказы с ТТН для расширения браузера =======================
// Расширение «Робот LP-CRM» создаёт по этим строкам заказы в CRM (статус «Импорт») и вписывает ТТН.
// Строки копируются из листа и вставляются в окно расширения, поэтому порядок колонок менять нельзя.
var CRMZ_SHEET = 'Заказы на СРМ';
var CRMZ_COLUMNS = ['Статус', 'ТТН', 'ФИО', 'Телефон', 'Товар', 'Кол-во', 'Сумма', 'Адрес НП', 'ID товара СРМ', 'Call ID', 'Комментарий'];
function ZC_(name) { return CRMZ_COLUMNS.indexOf(name); }

/** Кнопка: выделенные строки «НП лиды» (с ТТН) -> лист «Заказы на СРМ». */
function perenestiVZakazyNaCrm() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  var src = ss.getSheetByName(TTN_SHEET);
  if (!src || ss.getActiveSheet().getName() !== TTN_SHEET) {
    ui.alert('Откройте лист «' + TTN_SHEET + '», выделите строки с ТТН и нажмите кнопку ещё раз.'); return;
  }
  var nomera = vydelennyeStroki_(src).filter(function (r) { return r <= src.getLastRow(); });
  if (!nomera.length) { ui.alert('Выделите строки заказов в «' + TTN_SHEET + '».'); return; }
  var tovary = tovaryNp_();
  var dest = listTtn_(CRMZ_SHEET, CRMZ_COLUMNS, [200, 130, 190, 110, 240, 60, 80, 340, 110, 110, 320]);
  var est = {};
  if (dest.getLastRow() > 1) {
    dest.getRange(2, ZC_('Call ID') + 1, dest.getLastRow() - 1, 1).getValues().forEach(function (r) { est[String(r[0]).trim()] = true; });
  }
  // допы из пула - в комментарий заказа
  var dopy = {};
  var pool = ss.getSheetByName(ZAYAVKI_CFG.POOL_SHEET);
  if (pool && pool.getLastRow() > 1 && P_('Допы') >= 0) {
    pool.getRange(2, 1, pool.getLastRow() - 1, POOL_COLUMNS.length).getValues().forEach(function (r) {
      var d = String(r[P_('Допы')] || '').trim();
      if (d && !/^(нет|немає|-)/i.test(d)) dopy[String(r[P_('Call ID')]).trim()] = d;
    });
  }
  var vse = src.getRange(2, 1, src.getLastRow() - 1, TTN_COLUMNS.length).getValues();
  var add = [], propuski = [], bezId = 0;
  nomera.forEach(function (n) {
    var r = vse[n - 2];
    var cid = String(r[TC_('Call ID')]).trim(), ttn = String(r[TC_('ТТН')]).replace(/\D/g, '');
    if (!cid) return;
    if (!ttn) { propuski.push(n + ' (нет ТТН)'); return; }
    if (est[cid]) { propuski.push(n + ' (уже на листе)'); return; }
    var items = razobratTovary_(r[TC_('Товар')]);
    var t = items.length ? tovary[klyuchTovara_(items[0].name)] : null;
    var id = (t && t.idCrm) || '';
    if (!id) bezId++;
    var qty = chislo_(r[TC_('Кол-во')]) || (items.length ? items[0].qty : 1);
    var sum = chislo_(r[TC_('Наложка')]);
    var comment = String(r[TC_('Товар')] || '') + '. Штук: ' + qty + '. Сумма: ' + sum + ' грн.' + (dopy[cid] ? ' Допы: ' + dopy[cid] : '');
    var z = CRMZ_COLUMNS.map(function () { return ''; });
    z[ZC_('Статус')] = id ? '' : '⚠ нет ID товара (лист «' + TTN_GOODS + '», колонка J)';
    z[ZC_('ТТН')] = ttn; z[ZC_('ФИО')] = r[TC_('ФИО')]; z[ZC_('Телефон')] = String(r[TC_('Телефон')]);
    z[ZC_('Товар')] = r[TC_('Товар')]; z[ZC_('Кол-во')] = qty; z[ZC_('Сумма')] = sum;
    z[ZC_('Адрес НП')] = r[TC_('Адрес НП')]; z[ZC_('ID товара СРМ')] = id; z[ZC_('Call ID')] = cid;
    z[ZC_('Комментарий')] = '';   // в CRM комментарий не пишем; колонку можно заполнить вручную
    est[cid] = true;
    add.push(z);
  });
  if (add.length) {
    var start = dest.getLastRow() + 1;
    [ZC_('ТТН'), ZC_('Телефон'), ZC_('ID товара СРМ'), ZC_('Call ID')].forEach(function (c) {
      dest.getRange(start, c + 1, add.length, 1).setNumberFormat('@');
    });
    dest.getRange(start, 1, add.length, CRMZ_COLUMNS.length).setValues(add).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  ui.alert(CRMZ_SHEET, 'Перенесено: ' + add.length +
    (bezId ? '\nБез ID товара: ' + bezId + ' - впишите ID в «' + TTN_GOODS + '» (колонка J) или прямо в строку.' : '') +
    (propuski.length ? '\nПропущены строки: ' + propuski.join(', ') : ''), ui.ButtonSet.OK);
}


// ======================= Связь с расширением «Робот LP-CRM» =======================
// Расширение само читает лист «Заказы на СРМ» и записывает в него статусы. Для этого скрипт нужно один раз
// опубликовать: «Развернуть» -> «Новое развертывание» -> «Веб-приложение» (от моего имени, доступ: все).
// Доступ защищён секретным кодом в ссылке: без него скрипт ничего не отдаёт.

/** Кнопка: показывает код доступа для расширения (и ссылку, которую отдаёт Google, для сравнения). */
function ssylkaDlyaRasshireniya() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  var token = props.getProperty('ROBOT_TOKEN');
  if (!token) { token = Utilities.getUuid().replace(/-/g, ''); props.setProperty('ROBOT_TOKEN', token); }
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var html = HtmlService.createHtmlOutput(
    '<div style="font:13px Arial;line-height:1.5">' +
    '<b>1. Код доступа</b> (вставьте во вторую строку расширения):<br>' +
    '<input style="width:100%;font:13px monospace;padding:4px" onclick="this.select()" readonly value="' + token + '"><br><br>' +
    '<b>2. Адрес веб-приложения</b> - скопируйте его из окна «Развернуть» → «Управление развертываниями» кнопкой «Копировать» ' +
    'у блока «Веб-приложение» (он начинается с https://script.google.com/macros/s/ и заканчивается на /exec).<br><br>' +
    (url ? 'Адрес, который сообщает сам скрипт (сравните с тем, что скопировали):<br>' +
      '<input style="width:100%;font:11px monospace;padding:4px" onclick="this.select()" readonly value="' + url + '"><br><br>' : '') +
    'Никому не отправляйте код: по нему виден лист «' + CRMZ_SHEET + '».</div>').setWidth(560).setHeight(380);
  ui.showModalDialog(html, 'Подключение расширения');
}

function robotOtvet_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function robotDostup_(t) {
  var token = PropertiesService.getScriptProperties().getProperty('ROBOT_TOKEN') || '';
  return token && String(t || '') === token;
}

/** Расширение запрашивает заказы, которые ещё не созданы в CRM (в «Статус» нет OK). */
function doGet(e) {
  try { return doGet_(e); } catch (err) { return robotOtvet_({ ok: false, error: 'ошибка скрипта таблицы: ' + err.message }); }
}
function doGet_(e) {
  var p = (e && e.parameter) || {};
  if (!robotDostup_(p.t)) return robotOtvet_({ ok: false, error: 'нет доступа' });
  var sh = SpreadsheetApp.getActive().getSheetByName(CRMZ_SHEET);
  var rows = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, CRMZ_COLUMNS.length).getDisplayValues().forEach(function (r, i) {
      var cid = String(r[ZC_('Call ID')]).trim();
      if (!cid || /^OK/i.test(String(r[ZC_('Статус')]).trim())) return;
      rows.push({ row: i + 2, ttn: r[ZC_('ТТН')], name: r[ZC_('ФИО')], phone: r[ZC_('Телефон')], tovar: r[ZC_('Товар')],
                  qty: r[ZC_('Кол-во')], sum: r[ZC_('Сумма')], addr: r[ZC_('Адрес НП')], pid: r[ZC_('ID товара СРМ')],
                  code: cid, comment: r[ZC_('Комментарий')] });
    });
  }
  return robotOtvet_({ ok: true, rows: rows });
}

/** Расширение записывает статусы: {t, updates:[{row, code, status}]}. Строка сверяется по Call ID. */
function doPost(e) {
  try { return doPost_(e); } catch (err) { return robotOtvet_({ ok: false, error: 'ошибка скрипта таблицы: ' + err.message }); }
}
function doPost_(e) {
  var data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) {}
  if (!robotDostup_(data.t)) return robotOtvet_({ ok: false, error: 'нет доступа' });
  var sh = SpreadsheetApp.getActive().getSheetByName(CRMZ_SHEET);
  if (!sh || sh.getLastRow() < 2) return robotOtvet_({ ok: false, error: 'нет листа' });
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return robotOtvet_({ ok: false, error: 'таблица занята' }); }
  var n = 0;
  try {
    var ids = sh.getRange(2, ZC_('Call ID') + 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); });
    (data.updates || []).forEach(function (u) {
      var code = String(u.code || '').trim(), row = Number(u.row) || 0;
      if (!code) return;
      if (!(row >= 2 && ids[row - 2] === code)) { var k = ids.indexOf(code); row = k >= 0 ? k + 2 : 0; }   // строки могли сдвинуться
      if (!row) return;
      sh.getRange(row, ZC_('Статус') + 1).setValue(String(u.status || '').slice(0, 300));
      n++;
    });
  } finally { lock.releaseLock(); }
  return robotOtvet_({ ok: true, written: n });
}
