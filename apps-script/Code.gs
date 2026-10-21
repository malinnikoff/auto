/**
 * Приём заявок с лендинга @dmitry_miami.
 *
 * Что делает:
 *  1. Принимает заявку с сайта (POST, JSON).
 *  2. Проверяет данные и защищается от спама.
 *  3. Сохраняет заявку в лист «Заявки»: номер, время, метку, все поля, источник.
 *  4. Присылает уведомление в Telegram и копию на email.
 *  5. Отвечает сайту {ok:true, id} только после того, как строка записана.
 *
 * Настройки хранятся в «Свойствах скрипта» (Project Settings → Script properties):
 *   TELEGRAM_TOKEN  — токен бота от @BotFather
 *   TELEGRAM_CHAT_ID — ваш chat id (функция getChatId ниже его покажет)
 *   NOTIFY_EMAIL     — почта для копии уведомлений (необязательно)
 *   CRM_WEBHOOK_URL  — адрес CRM, когда она будет готова (необязательно)
 *   CONTRACTS_FOLDER_ID — папка Google Drive для договоров (создаётся сама при первой загрузке)
 */

var SHEET_NAME = 'Заявки';
var STATUSES = ['Новая', 'Связался', 'Не дозвонился', 'Передана партнёрам', 'Предложение отправлено', 'Сделка', 'Отказ', 'Не целевая'];

var COLUMNS = [
  ['id', 'Номер'],
  ['created', 'Дата и время'],
  ['score', 'Метка'],
  ['status', 'Статус'],
  ['name', 'Имя'],
  ['phone', 'Телефон'],
  ['contact', 'Связь'],
  ['state', 'Штат'],
  ['zip', 'ZIP'],
  ['make', 'Марка'],
  ['model', 'Модель'],
  ['year', 'Год'],
  ['deal', 'Покупка / лизинг'],
  ['pay', 'Оплата'],
  ['budgetType', 'Тип бюджета'],
  ['budget', 'Бюджет'],
  ['tradein', 'Trade-in'],
  ['when', 'Когда покупка'],
  ['utm_source', 'utm_source'],
  ['utm_medium', 'utm_medium'],
  ['utm_campaign', 'utm_campaign'],
  ['utm_content', 'utm_content'],
  ['referrer', 'Откуда пришёл'],
  ['page', 'Страница'],
  ['notes', 'Заметки'],
  ['crm', 'CRM']
];

var ALLOWED = {
  deal: ['Покупка', 'Лизинг', 'Пока не решил(а)'],
  pay: ['', 'Кредит', 'Наличные', 'Пока не знаю'],
  when: ['7 дней', '30 дней', '1–3 месяца', 'Позже'],
  budgetType: ['Платёж в месяц', 'Цена машины'],
  tradein: ['Да', 'Нет'],
  contact: ['Звонок', 'WhatsApp', 'Telegram'],
  year: ['', 'Не важно', '2027', '2026', '2025', '2024', '2023', '2022', '2021', '2020', 'Старше 2020']
};

/* ================= Приём заявки ================= */

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Спам-ловушка: настоящий посетитель это поле не видит.
    if (data.website) return json_({ ok: true, id: 'OK' });

    if (data.type === 'contract') return handleContract_(data);

    // Отчёт сайта о сбое отправки договора (без личных данных): виден в «Выполнениях» скрипта.
    if (data.type === 'clientError') {
      console.warn('Сбой на сайте: ' + String(data.where).slice(0, 200) + ' | ' + String(data.ua).slice(0, 200) + ' | ' + JSON.stringify(data.files || []).slice(0, 300));
      return json_({ ok: true });
    }

    var lead = clean_(data);
    var problem = check_(lead);
    if (problem) return json_({ ok: false, error: problem });

    // Не больше одной заявки с одного номера за 2 минуты (двойной клик, повторная отправка).
    var cache = CacheService.getScriptCache();
    var dupKey = 'p' + lead.phone;
    var prevId = cache.get(dupKey);
    if (prevId) return json_({ ok: true, id: prevId, duplicate: true });

    lead.score = score_(lead);
    lead.status = 'Новая';

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sheet = sheet_();
      var props = PropertiesService.getScriptProperties();
      var n = Number(props.getProperty('LAST_ID') || '0') + 1;
      props.setProperty('LAST_ID', String(n));
      lead.id = ('0000' + n).slice(-4);
      lead.created = new Date();
      // Номер и телефон пишем как текст, иначе таблица превратит «0007» в 7, а «+1305…» в число.
      var asText = { id: 1, phone: 1, zip: 1 };
      sheet.appendRow(COLUMNS.map(function (c) {
        var v = lead[c[0]] !== undefined ? lead[c[0]] : '';
        return asText[c[0]] && v !== '' ? "'" + v : v;
      }));
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    cache.put(dupKey, lead.id, 120);

    // Уведомления после записи. Их сбой не отменяет сохранённую заявку.
    try { notifyTelegram_(lead); } catch (err) { console.error('Telegram: ' + err); notifyEmail_(lead, 'Telegram не сработал: ' + err); }
    try { notifyEmail_(lead); } catch (err) { console.error('Email: ' + err); }
    try { pushToCrm_(lead); } catch (err) { console.error('CRM: ' + err); }

    return json_({ ok: true, id: lead.id });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function doGet() {
  return json_({ ok: true, service: 'leads' });
}

/* ================= Проверка договора ================= */

var CONTRACT_SHEET = 'Договоры';
var CONTRACT_COLUMNS = [
  ['id', 'Номер'], ['created', 'Дата и время'], ['status', 'Статус'], ['name', 'Имя'], ['phone', 'Телефон'],
  ['contact', 'Куда прислать разбор'], ['stage', 'Подписан?'],
  ['folder', 'Файлы'], ['fileCount', 'Кол-во файлов'], ['utm_source', 'utm_source'], ['utm_medium', 'utm_medium'],
  ['utm_campaign', 'utm_campaign'], ['utm_content', 'utm_content'], ['referrer', 'Откуда пришёл'], ['notes', 'Заметки / найдено $']
];
var CONTRACT_STATUSES = ['Новая', 'Разбираем', 'Разбор отправлен', 'Перешёл в покупку', 'Не целевая'];
var MAX_FILES = 10, MAX_BYTES = 12 * 1024 * 1024;
var ALLOWED_MIME = /^(image\/(jpeg|png|heic|heif|webp)|application\/pdf)$/;

function handleContract_(d) {
  function s(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 80); }
  var l = { name: s(d.name), phone: s(d.phone, 20), contact: s(d.contact), stage: s(d.stage),
    utm_source: s(d.utm_source), utm_medium: s(d.utm_medium),
    utm_campaign: s(d.utm_campaign), utm_content: s(d.utm_content), referrer: s(d.referrer, 300) };
  if (d.consent !== true) return json_({ ok: false, error: 'consent' });
  if (l.name.length < 2) return json_({ ok: false, error: 'name' });
  if (!/^\+\d{8,15}$/.test(l.phone)) return json_({ ok: false, error: 'phone' });
  if (['WhatsApp', 'Telegram', 'Звонок'].indexOf(l.contact) === -1) return json_({ ok: false, error: 'contact' });
  if (['Ещё нет, это предложение дилера', 'Да, уже подписал(а)'].indexOf(l.stage) === -1) return json_({ ok: false, error: 'stage' });
  var files = Array.isArray(d.files) ? d.files : [];
  if (!files.length || files.length > MAX_FILES) return json_({ ok: false, error: 'files' });

  var blobs = [], total = 0;
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    if (!f || !ALLOWED_MIME.test(String(f.mime))) return json_({ ok: false, error: 'filetype' });
    var bytes = Utilities.base64Decode(String(f.data || ''));
    total += bytes.length;
    if (total > MAX_BYTES) return json_({ ok: false, error: 'size' });
    var safe = String(f.name || ('file' + (i + 1))).replace(/[\\/:*?"<>|]/g, '_').slice(0, 100);
    blobs.push(Utilities.newBlob(bytes, f.mime, safe));
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('LAST_CONTRACT_ID') || '0') + 1;
    props.setProperty('LAST_CONTRACT_ID', String(n));
    l.id = 'D' + ('000' + n).slice(-3);
    l.created = new Date();
    l.status = 'Новая';
    var root = contractsFolder_();
    var folder = root.createFolder(l.id + ' · ' + l.name + ' · ' + Utilities.formatDate(l.created, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    blobs.forEach(function (b) { folder.createFile(b); });
    l.folder = folder.getUrl();
    l.fileCount = blobs.length;
    var sh = contractSheet_();
    var asText = { id: 1, phone: 1 };
    sh.appendRow(CONTRACT_COLUMNS.map(function (c) { var v = l[c[0]] !== undefined ? l[c[0]] : ''; return asText[c[0]] && v !== '' ? "'" + v : v; }));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var text = '📄 <b>ДОГОВОР НА ПРОВЕРКУ</b> · #' + l.id + ' · ' + Utilities.formatDate(l.created, Session.getScriptTimeZone(), 'dd.MM HH:mm') + '\n' +
    '👤 ' + esc_(l.name) + ' · ' + esc_(phoneText_(l.phone)) + ' · прислать разбор: ' + esc_(l.contact) + '\n' +
    (l.stage === 'Да, уже подписал(а)' ? 'Договор уже подписан' : 'Не подписан, это предложение дилера') + ' · файлов: ' + l.fileCount + '\n' +
    'Откуда: ' + esc_(sourceText_(l));
  var digits = l.phone.replace(/\D/g, '');
  try {
    sendTelegram_(text, [[{ text: 'Открыть файлы', url: l.folder }, { text: 'WhatsApp', url: 'https://wa.me/' + digits }], [{ text: 'Таблица', url: sheetLink_(contractSheet_()) }]]);
  } catch (err) { console.error('Telegram: ' + err); }
  try {
    var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
    if (to) MailApp.sendEmail({ to: to, subject: 'Договор на проверку #' + l.id + ': ' + l.name, body: text.replace(/<[^>]+>/g, '') + '\nФайлы: ' + l.folder });
  } catch (err) { console.error('Email: ' + err); }

  return json_({ ok: true, id: l.id });
}

function contractsFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CONTRACTS_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var folder = DriveApp.createFolder('Договоры с сайта');
  props.setProperty('CONTRACTS_FOLDER_ID', folder.getId());
  return folder;
}

function contractSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONTRACT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONTRACT_SHEET);
    sh.getRange(1, 1, 1, CONTRACT_COLUMNS.length).setValues([CONTRACT_COLUMNS.map(function (c) { return c[1]; })])
      .setFontWeight('bold').setBackground('#FFD600');
    sh.setFrozenRows(1);
    sh.getRange('B2:B').setNumberFormat('dd.MM.yyyy HH:mm');
    sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CONTRACT_STATUSES, true).setAllowInvalid(false).build());
  }
  return sh;
}

function esc_(v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* ================= Проверка и оценка ================= */

function clean_(d) {
  function s(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 120); }
  var lead = {};
  ['state', 'make', 'model', 'year', 'deal', 'pay', 'when', 'budgetType', 'budget', 'tradein', 'zip', 'name', 'phone', 'contact',
   'utm_source', 'utm_medium', 'utm_campaign', 'utm_content'].forEach(function (k) { lead[k] = s(d[k], 80); });
  lead.referrer = s(d.referrer, 300);
  lead.page = s(d.page, 300);
  lead.consent = d.consent === true;
  // Защита таблицы от формул: значение не должно начинаться с = + - @
  Object.keys(lead).forEach(function (k) {
    if (typeof lead[k] === 'string' && /^[=+\-@]/.test(lead[k]) && k !== 'phone') lead[k] = "'" + lead[k];
  });
  return lead;
}

function check_(l) {
  if (!l.consent) return 'consent';
  if (l.name.length < 2) return 'name';
  if (!/^\+\d{8,15}$/.test(l.phone)) return 'phone';
  if (!l.state || !l.make) return 'car';
  if (l.zip && !/^\d{5}$/.test(l.zip)) return 'zip';
  if (!l.budget) return 'budget';
  for (var k in ALLOWED) {
    if (ALLOWED[k].indexOf(l[k]) === -1) return k;
  }
  return '';
}

function score_(l) {
  if (l.state === 'Не в США') return 'Холодная';
  if (l.when === '7 дней' || l.when === '30 дней') return 'Горячая';
  if (l.when === '1–3 месяца') return 'Тёплая';
  return 'Холодная';
}

/* ================= Таблица ================= */

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    setupSheet_(sh);
  }
  return sh;
}

function setupSheet_(sh) {
  sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(function (c) { return c[1]; })])
    .setFontWeight('bold').setBackground('#FFD600');
  sh.setFrozenRows(1);
  sh.getRange('B2:B').setNumberFormat('dd.MM.yyyy HH:mm');
  sh.getRange('F2:F').setNumberFormat('@');
  var statusCol = COLUMNS.map(function (c) { return c[0]; }).indexOf('status') + 1;
  sh.getRange(2, statusCol, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  var scoreCol = COLUMNS.map(function (c) { return c[0]; }).indexOf('score') + 1;
  var r = sh.getRange(2, scoreCol, sh.getMaxRows() - 1, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Горячая').setBackground('#FFE0D1').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Тёплая').setBackground('#FFF6BF').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Холодная').setBackground('#E8EEF7').setRanges([r]).build()
  ]);
}

/* ================= Уведомления ================= */

var SCORE_ICON = { 'Горячая': '🔥', 'Тёплая': '🟡', 'Холодная': '⚪' };

// +13055550100 → +1 305 555 0100 (так номер легче прочитать)
function phoneText_(p) {
  var m = /^\+1(\d{3})(\d{3})(\d{4})$/.exec(p || '');
  return m ? '+1 ' + m[1] + ' ' + m[2] + ' ' + m[3] : String(p || '');
}

// Откуда пришёл человек, простыми словами
function sourceText_(l) {
  var src = (l.utm_source || '').toLowerCase(), med = (l.utm_medium || '').toLowerCase(), cnt = l.utm_content || '';
  var t;
  if (src === 'instagram' && med === 'bio') t = 'Instagram, шапка профиля';
  else if (src === 'instagram' && med === 'reels') t = 'Instagram, рилс';
  else if (src === 'instagram') t = 'Instagram' + (med ? ', ' + med : '');
  else if (src === 'dogovor' && med === 'bio') t = 'Instagram, ссылка «проверить договор» в шапке';
  else if (src === 'dogovor') t = 'ответ на слово «договор»';
  else if (src === 'cena') t = 'ответ на слово «цена»';
  else if (src === 'test') t = 'тест';
  else if (src) t = l.utm_source + (med ? ', ' + l.utm_medium : '');
  else if (l.referrer) t = 'переход с ' + l.referrer.replace(/^https?:\/\//, '').split('/')[0];
  else t = 'открыл сайт напрямую';
  return t + (cnt ? ' (' + cnt + ')' : '');
}

function dealText_(l) {
  if (l.deal === 'Лизинг') return 'лизинг';
  if (l.deal === 'Пока не решил(а)') return 'покупка или лизинг, пока не решил(а)';
  if (l.pay === 'Кредит') return 'покупка в кредит';
  if (l.pay === 'Наличные') return 'покупка за наличные';
  return 'покупка, способ оплаты пока не выбран';
}

function whenText_(w) {
  return { '7 дней': 'в течение 7 дней', '30 дней': 'в течение 30 дней', '1–3 месяца': 'через 1–3 месяца', 'Позже': 'позже, не в ближайшие 3 месяца' }[w] || w;
}

function leadText_(l, html) {
  var esc = function (v) { return html ? esc_(v) : String(v); };
  var b = function (v) { return html ? '<b>' + esc(v) + '</b>' : String(v); };
  var tz = Session.getScriptTimeZone();
  var car = [l.make, l.model].filter(String).join(' ') + (l.year && l.year !== 'Не важно' ? ' ' + l.year : '');
  var lines = [
    (SCORE_ICON[l.score] || '') + ' ' + b(l.score.toUpperCase()) + ' · #' + l.id + ' · ' + Utilities.formatDate(l.created, tz, 'dd.MM HH:mm'),
    '👤 ' + esc(l.name) + ' · ' + esc(phoneText_(l.phone)) + ' · связь: ' + esc(l.contact.toLowerCase() === 'звонок' ? 'звонок' : l.contact),
    '🚗 ' + b(car || 'Машина не указана') + ' · ' + esc(dealText_(l)),
    '💵 ' + esc(l.budgetType) + ': ' + esc(l.budget) + ' · ' + (l.tradein === 'Да' ? 'есть машина на обмен' : 'без обмена'),
    '📅 Хочет купить: ' + esc(whenText_(l.when)),
    '📍 ' + esc(l.state) + (l.zip ? ', ZIP ' + esc(l.zip) : ''),
    'Откуда: ' + esc(sourceText_(l))
  ];
  return lines.join('\n');
}

function sheetLink_(sh) {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + sh.getSheetId();
}

function notifyTelegram_(l) {
  var digits = l.phone.replace(/\D/g, '');
  sendTelegram_(leadText_(l, true), [[
    { text: 'WhatsApp', url: 'https://wa.me/' + digits },
    { text: 'Таблица', url: sheetLink_(sheet_()) }
  ]]);
}

function sendTelegram_(html, buttons) {
  var p = PropertiesService.getScriptProperties();
  var token = p.getProperty('TELEGRAM_TOKEN'), chat = p.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chat) throw new Error('не заданы TELEGRAM_TOKEN / TELEGRAM_CHAT_ID');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: html, parse_mode: 'HTML', disable_web_page_preview: true,
      reply_markup: buttons ? { inline_keyboard: buttons } : undefined })
  });
  if (res.getResponseCode() !== 200) throw new Error(res.getResponseCode() + ' ' + res.getContentText());
}

function notifyEmail_(l, prefix) {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (!to) return;
  MailApp.sendEmail({
    to: to,
    subject: (prefix ? '[!] ' : '') + l.score + ' заявка #' + l.id + ': ' + l.make + ' ' + l.model,
    body: (prefix ? prefix + '\n\n' : '') + leadText_(l, false) + '\n\nТаблица: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  });
}

/* ================= CRM (на будущее) ================= */

function pushToCrm_(l) {
  var url = PropertiesService.getScriptProperties().getProperty('CRM_WEBHOOK_URL');
  if (!url) return;
  var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(l) });
  var sh = sheet_();
  var row = findRow_(sh, l.id);
  var crmCol = COLUMNS.map(function (c) { return c[0]; }).indexOf('crm') + 1;
  if (row) sh.getRange(row, crmCol).setValue(res.getResponseCode() < 300 ? 'отправлено' : 'ошибка ' + res.getResponseCode());
}

function findRow_(sh, id) {
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getDisplayValues();
  for (var i = ids.length - 1; i >= 0; i--) if (ids[i][0] === id) return i + 2;
  return 0;
}

/* ================= Служебные функции (запускаются вручную) ================= */

/** 1. Создаёт лист «Заявки» с заголовками, статусами и цветом меток. */
function setup() {
  var sh = sheet_();
  setupSheet_(sh);
  contractSheet_();
  contractsFolder_();
  Logger.log('Лист «' + SHEET_NAME + '» готов: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
}

/** 2. Напишите своему боту любое сообщение, затем запустите эту функцию: в журнале появится chat id. */
function getChatId() {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Сначала добавьте TELEGRAM_TOKEN в свойства скрипта');
  var res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates').getContentText());
  var seen = {};
  (res.result || []).forEach(function (u) {
    var c = (u.message || u.channel_post || {}).chat;
    if (c && !seen[c.id]) { seen[c.id] = 1; Logger.log('chat id: ' + c.id + ' (' + (c.title || c.first_name || '') + ')'); }
  });
  if (!Object.keys(seen).length) Logger.log('Сообщений нет. Напишите боту /start и запустите снова.');
}

/** 3. Проверка: отправляет тестовое уведомление, в таблицу ничего не пишет. */
function testNotify() {
  var l = { id: 'TEST', created: new Date(), score: 'Горячая', make: 'Toyota', model: 'RAV4 XLE', year: '2026', deal: 'Лизинг', pay: '',
    budget: '$400–600', budgetType: 'Платёж в месяц', tradein: 'Да', when: '30 дней', state: 'Florida', zip: '33160',
    name: 'Тест', phone: '+13050000000', contact: 'WhatsApp', utm_source: 'test', referrer: '' };
  notifyTelegram_(l);
  notifyEmail_(l);
  Logger.log('Тестовое уведомление отправлено');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
