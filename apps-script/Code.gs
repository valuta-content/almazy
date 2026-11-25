/**
 * 💎 Мои алмазы — серверная часть (Google Apps Script).
 * Вставляется в таблицу: Расширения → Apps Script.
 * Здесь ничего менять не нужно. Токен бота вводится через меню «💎 Алмазы» в таблице.
 */

const SH = {
  members: 'Участницы',
  log: 'Начисления',
  tasks: 'Задания',
  rewards: 'Награды',
  rules: 'Правила',
  settings: 'Настройки',
};

// ---------- Меню в таблице ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💎 Алмазы')
    .addItem('1. Создать листы', 'setup')
    .addItem('2. Ввести токен бота', 'setToken')
    .addToUi();
}

function setToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Токен бота', 'Вставь токен из BotFather (сообщение с токеном лежит в Избранном):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim();
  if (!/^\d+:[\w-]{30,}$/.test(t)) {
    ui.alert('Это не похоже на токен. Он выглядит так: 1234567890:AAH... Скопируй его целиком из сообщения BotFather.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('BOT_TOKEN', t);
  ui.alert('Готово, токен сохранён.');
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Europe/Warsaw');

  // Задания: что вписывается в «За что» и сколько алмазов это даёт
  const tasks = sheet_(ss, SH.tasks, ['За что', 'Алмазы'], [
    ['ДЗ к уроку', 20],
    ['Фокус недели', 15],
    ['Reels по технике', 10],
    ['Итоги недели', 5],
    ['Участие в эфире', 10],
    ['Итоговая практика месяца', 15],
    ['Обмен: персональный разбор', -100],
    ['Обмен: персональная помощь Оли', -160],
    ['Другое (сумма вручную)', 0],
  ]);

  const members = sheet_(ss, SH.members, ['Имя', 'Ник', 'Telegram ID', 'Первый вход', 'Баланс'], []);
  if (members) {
    members.getRange('C:C').setNumberFormat('@');
    members.getRange('D:D').setNumberFormat('dd.MM.yyyy');
    members.getRange('E2').setFormula(
      "=ARRAYFORMULA(IF(A2:A=\"\",\"\",SUMIF('" + SH.log + "'!B:B,A2:A,'" + SH.log + "'!D:D)))");
    members.setColumnWidth(1, 200);
  }

  const log = sheet_(ss, SH.log, ['Дата', 'Участница', 'За что', 'Алмазы', 'Сумма вручную', 'Уточнение (видит участница)'], []);
  if (log) {
    log.getRange('A2:A').setNumberFormat('dd.MM.yyyy')
      .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
    log.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getSheetByName(SH.members).getRange('A2:A'), true).setAllowInvalid(false).build());
    log.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getSheetByName(SH.tasks).getRange('A2:A'), true).setAllowInvalid(false).build());
    log.getRange('D2').setFormula(
      "=ARRAYFORMULA(IF(C2:C=\"\",\"\",IF(E2:E<>\"\",E2:E,IFERROR(VLOOKUP(C2:C,'" + SH.tasks + "'!A:B,2,FALSE),0))))");
    log.getRange('D1:D').setBackground('#eef3fb');
    log.setColumnWidths(1, 6, 160);
    log.setColumnWidth(3, 240);
    log.setColumnWidth(6, 260);
  }

  sheet_(ss, SH.rewards, ['Награда', 'Алмазы', 'Описание'], [
    ['Персональный разбор', 100, 'Разбор одного Reels или сценария от Оли с конкретными рекомендациями'],
    ['Персональная помощь Оли', 160, 'Мини-консультация или индивидуальный разбор контент-стратегии'],
  ]);

  sheet_(ss, SH.rules, ['Правило (каждая строка — отдельный пункт)'], [
    ['Алмазы начисляются по понедельникам.'],
    ['ДЗ 20, фокус недели 15, эфир 10, Reels по технике 10, итоги 5, практика месяца 15.'],
    ['Задание засчитывается, если сдано в течение 2 недель.'],
    ['После окончания подписки алмазы хранятся 30 дней.'],
  ]);

  sheet_(ss, SH.settings, ['Параметр', 'Значение'], [
    ['Тестовый режим (ДА / НЕТ)', 'ДА'],
    ['Кто видит приложение в тестовом режиме (Telegram ID через запятую)', ''],
    ['Как обменять', 'Чтобы обменять, напиши Оле в личные сообщения, какую награду хочешь.'],
  ]);

  // Лишний пустой «Лист1» убираем
  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  ss.setActiveSheet(ss.getSheetByName(SH.log));
  SpreadsheetApp.getUi().alert('Листы созданы. Следующий шаг: меню «💎 Алмазы» → «2. Ввести токен бота».');
}

// Создаёт лист с шапкой и стартовыми строками, если такого листа ещё нет. Существующий не трогает.
function sheet_(ss, name, header, rows) {
  if (ss.getSheetByName(name)) return null;
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#dfe8f7');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.autoResizeColumns(1, header.length);
  return sh;
}

// ---------- Ответ приложению ----------

function doGet(e) {
  let out;
  try {
    out = handle_((e && e.parameter && e.parameter.initData) || '');
  } catch (err) {
    out = { ok: false, error: 'server', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(initData) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) return { ok: false, error: 'no_token' };
  const user = verify_(initData, token);
  if (!user) return { ok: false, error: 'bad_auth' };

  const ss = SpreadsheetApp.getActive();
  const settings = readSettings_(ss);
  const name = findOrAddMember_(ss, user);

  if (settings.test && settings.allowed.indexOf(String(user.id)) < 0) {
    return { ok: true, soon: true, first_name: user.first_name || '' };
  }

  const tz = ss.getSpreadsheetTimeZone();
  const logRows = values_(ss, SH.log, 6);
  const history = [];
  let balance = 0;
  logRows.forEach(function (r) {
    if (String(r[1]).trim() !== name || r[2] === '') return;
    const amount = Number(r[3]) || 0;
    balance += amount;
    const d = r[0] instanceof Date ? r[0] : null;
    history.push({
      date: d ? Utilities.formatDate(d, tz, 'dd.MM.yyyy') : '',
      ts: d ? d.getTime() : 0,
      title: String(r[2]).replace(' (сумма вручную)', '') + (r[5] ? ' · ' + r[5] : ''),
      amount: amount,
    });
  });
  history.sort(function (a, b) { return b.ts - a.ts; });

  const rewards = values_(ss, SH.rewards, 3)
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { name: String(r[0]), price: Number(r[1]) || 0, desc: String(r[2] || '') }; });
  const rules = values_(ss, SH.rules, 1)
    .map(function (r) { return String(r[0]).trim(); })
    .filter(String);

  return {
    ok: true,
    name: name,
    first_name: user.first_name || '',
    balance: balance,
    history: history,
    rewards: rewards,
    rules: rules,
    howto: settings.howto,
  };
}

// Проверка, что данные действительно пришли от Telegram (подпись токеном бота)
function verify_(initData, token) {
  if (!initData) return null;
  let hash = '';
  const rest = [];
  initData.split('&').forEach(function (p) {
    const i = p.indexOf('=');
    if (i < 0) return;
    const k = decodeURIComponent(p.slice(0, i));
    const v = decodeURIComponent(p.slice(i + 1).replace(/\+/g, '%20'));
    if (k === 'hash') hash = v; else rest.push([k, v]);
  });
  if (!hash) return null;
  rest.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  const dcs = rest.map(function (kv) { return kv[0] + '=' + kv[1]; }).join('\n');
  const bytes = function (s) { return Utilities.newBlob(s).getBytes(); };
  const secret = Utilities.computeHmacSha256Signature(bytes(token), bytes('WebAppData'));
  const sig = Utilities.computeHmacSha256Signature(bytes(dcs), secret);
  const hex = sig.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
  if (hex !== hash) return null;
  const obj = {};
  rest.forEach(function (kv) { obj[kv[0]] = kv[1]; });
  if (Date.now() / 1000 - Number(obj.auth_date) > 86400) return null; // ссылка старше суток
  try { return JSON.parse(obj.user); } catch (e) { return null; }
}

function readSettings_(ss) {
  const rows = values_(ss, SH.settings, 2);
  const get = function (prefix) {
    const r = rows.filter(function (x) { return String(x[0]).indexOf(prefix) === 0; })[0];
    return r ? String(r[1]).trim() : '';
  };
  return {
    test: get('Тестовый режим').toUpperCase() !== 'НЕТ',
    allowed: get('Кто видит').split(/[\s,;]+/).filter(String),
    howto: get('Как обменять'),
  };
}

// Находит участницу по Telegram ID; новую добавляет в лист «Участницы». Возвращает имя из колонки «Имя».
function findOrAddMember_(ss, user) {
  const sh = ss.getSheetByName(SH.members);
  const id = String(user.id);
  const find = function () {
    const rows = values_(ss, SH.members, 3);
    for (let i = 0; i < rows.length; i++) if (String(rows[i][2]).trim() === id) return { name: String(rows[i][0]).trim(), rows: rows };
    return { name: null, rows: rows };
  };
  let f = find();
  if (f.name) return f.name;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    f = find();
    if (f.name) return f.name;
    const names = f.rows.map(function (r) { return String(r[0]).trim(); });
    let name = [user.first_name, user.last_name].filter(Boolean).join(' ').trim() || ('Участница ' + id);
    if (names.indexOf(name) >= 0) name += user.username ? ' (@' + user.username + ')' : ' (' + id + ')';
    const row = f.rows.length + 2;
    sh.getRange(row, 1, 1, 4).setValues([[name, user.username ? '@' + user.username : '', id, new Date()]]);
    return name;
  } finally {
    lock.releaseLock();
  }
}

// Строки листа без шапки, до последней заполненной в первой колонке (формулы ARRAYFORMULA не мешают)
function values_(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const n = sh.getMaxRows() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, cols).getValues();
  let last = -1;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i].slice(0, Math.min(cols, 3)).some(function (v, j) { return j !== 3 && v !== ''; })) last = i;
  }
  return vals.slice(0, last + 1);
}
