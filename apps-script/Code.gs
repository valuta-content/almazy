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

// Задания. Колонки: За что | Алмазы | Отмечает участница | Как часто | Сколько раз
// «Как часто»: ОДИН РАЗ (за всё время), В НЕДЕЛЮ (пн–вс), В МЕСЯЦ, БЕЗ ЛИМИТА. «Сколько раз» — сколько отметок за этот период.
const TASK_HEADER = ['За что', 'Алмазы', 'Отмечает участница (ДА/НЕТ)', 'Как часто (ОДИН РАЗ / В НЕДЕЛЮ / В МЕСЯЦ / БЕЗ ЛИМИТА)', 'Сколько раз за период'];
const TASK_ROWS = [
  ['ДЗ к уроку 1', 20, 'ДА', 'ОДИН РАЗ', 1],
  ['ДЗ к уроку 2', 20, 'ДА', 'ОДИН РАЗ', 1],
  ['ДЗ к уроку 3', 20, 'ДА', 'ОДИН РАЗ', 1],
  ['ДЗ к уроку 4', 20, 'ДА', 'ОДИН РАЗ', 1],
  ['Фокус недели', 15, 'ДА', 'В НЕДЕЛЮ', 1],
  ['Итоги недели', 5, 'ДА', 'В НЕДЕЛЮ', 1],
  ['Участие в эфире', 10, 'ДА', 'В НЕДЕЛЮ', 1],
  ['Reels по технике', 5, 'ДА', 'БЕЗ ЛИМИТА', ''],
  ['Итоговая практика месяца', 15, 'ДА', 'В МЕСЯЦ', 1],
  ['Обмен: персональный разбор', -100, 'НЕТ', '', ''],
  ['Обмен: персональная помощь Оли', -160, 'НЕТ', '', ''],
  ['Другое (сумма вручную)', 0, 'НЕТ', '', ''],
];

// Переделывает лист «Задания» под кнопки для участниц (старый лист сохраняется копией).
function upgradeTasks() {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(SH.tasks);
  if (old) {
    const backup = old.copyTo(ss).setName('Задания (старое) ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM HH:mm'));
    backup.hideSheet();
    old.clear();
    old.getRange(1, 1, 1, TASK_HEADER.length).setValues([TASK_HEADER]).setFontWeight('bold').setBackground('#dfe8f7');
    old.getRange(2, 1, TASK_ROWS.length, TASK_HEADER.length).setValues(TASK_ROWS);
    old.setFrozenRows(1);
    old.autoResizeColumns(1, TASK_HEADER.length);
  } else {
    sheet_(ss, SH.tasks, TASK_HEADER, TASK_ROWS);
  }
  try { SpreadsheetApp.getUi().alert('Готово. Лист «Задания» обновлён. Новые уроки добавляй строкой «ДЗ к уроку 5» и т.д.'); } catch (e) {}
}

// ---------- Меню в таблице ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💎 Алмазы')
    .addItem('1. Создать листы', 'setup')
    .addItem('2. Ввести токен бота', 'setToken')
    .addItem('3. Обновить лист «Задания» (кнопки для участниц)', 'upgradeTasks')
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
  // Формулы пишутся с запятыми; на время настройки ставим английскую локаль, потом возвращаем.
  const locale = ss.getSpreadsheetLocale();
  ss.setSpreadsheetLocale('en_US');
  try { setupSheets_(ss); } finally { ss.setSpreadsheetLocale(locale); }
  try { SpreadsheetApp.getUi().alert('Листы созданы. Следующий шаг: меню «💎 Алмазы» → «2. Ввести токен бота».'); } catch (e) {}
}

function setupSheets_(ss) {

  // Задания: что вписывается в «За что» и сколько алмазов это даёт
  sheet_(ss, SH.tasks, TASK_HEADER, TASK_ROWS);

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
    ['ДЗ 20, фокус недели 15, эфир 10, Reels по технике 5, итоги 5, практика месяца 15.'],
    ['Задание засчитывается, если сдано в течение 2 недель.'],
    ['После окончания подписки алмазы хранятся 30 дней.'],
  ]);

  sheet_(ss, SH.settings, ['Параметр', 'Значение'], [
    ['Тестовый режим (ДА / НЕТ)', 'ДА'],
    ['Кто видит приложение в тестовом режиме (Telegram ID через запятую)', ''],
    ['Как обменять', 'Чтобы обменять, напиши Оле в личные сообщения, какую награду хочешь.'],
  ]);

  // Лишний пустой «Лист1» убираем
  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Arkusz1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  ss.setActiveSheet(ss.getSheetByName(SH.log));
  SpreadsheetApp.flush();
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
    const p = (e && e.parameter) || {};
    out = handle_(p.initData || '', p.action || '', p.task || '');
  } catch (err) {
    out = { ok: false, error: 'server', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(initData, action, taskName) {
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
  let marked = null;
  if (action === 'mark') marked = mark_(ss, name, taskName, tz);
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

  const tasks = taskStatus_(ss, name, logRows, tz);

  return {
    ok: true,
    marked: marked,
    tasks: tasks,
    name: name,
    first_name: user.first_name || '',
    balance: balance,
    history: history,
    rewards: rewards,
    rules: rules,
    howto: settings.howto,
  };
}

// ---------- Отметки участниц ----------

function readTasks_(ss) {
  return values_(ss, SH.tasks, 5)
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      const per = String(r[3]).trim().toUpperCase();
      return {
        name: String(r[0]).trim(),
        amount: Number(r[1]) || 0,
        self: String(r[2]).trim().toUpperCase() === 'ДА',
        period: per.indexOf('БЕЗ') >= 0 ? 'none' : per.indexOf('НЕДЕЛ') >= 0 ? 'week' : per.indexOf('МЕСЯЦ') >= 0 ? 'month' : 'once',
        limit: Math.max(1, Number(r[4]) || 1),
      };
    });
}

function periodKey_(d, period, tz) {
  if (period === 'once') return 'all';
  const ymd = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  if (period === 'month') return ymd.slice(0, 7);
  const p = ymd.split('-').map(Number);
  const u = Date.UTC(p[0], p[1] - 1, p[2]);
  const dow = (new Date(u).getUTCDay() + 6) % 7; // пн = 0
  return new Date(u - dow * 86400000).toISOString().slice(0, 10);
}

function taskStatus_(ss, name, logRows, tz) {
  const now = new Date();
  return readTasks_(ss).filter(function (t) { return t.self && t.amount > 0; }).map(function (t) {
    if (t.period === 'none') return { name: t.name, amount: t.amount, period: 'none', limit: 0, used: 0, available: true };
    const key = periodKey_(now, t.period, tz);
    let used = 0;
    logRows.forEach(function (r) {
      if (String(r[1]).trim() !== name || String(r[2]).trim() !== t.name) return;
      const d = r[0] instanceof Date ? r[0] : null;
      if (t.period === 'once' || (d && periodKey_(d, t.period, tz) === key)) used++;
    });
    return { name: t.name, amount: t.amount, period: t.period, limit: t.limit, used: used, available: used < t.limit };
  });
}

function mark_(ss, name, taskName, tz) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTasks_(ss).filter(function (x) { return x.name === taskName && x.self && x.amount > 0; })[0];
    if (!t) return { ok: false, reason: 'unknown' };
    const logRows = values_(ss, SH.log, 6);
    const st = taskStatus_(ss, name, logRows, tz).filter(function (x) { return x.name === taskName; })[0];
    if (!st || !st.available) return { ok: false, reason: 'limit' };
    const sh = ss.getSheetByName(SH.log);
    const row = logRows.length + 2;
    sh.getRange(row, 1, 1, 3).setValues([[new Date(), name, t.name]]);
    SpreadsheetApp.flush();
    return { ok: true, task: t.name, amount: t.amount };
  } finally {
    lock.releaseLock();
  }
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
