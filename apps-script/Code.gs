/**
 * 💎 Мои алмазы — серверная часть (Google Apps Script).
 * Вставляется в таблицу: Расширения → Apps Script.
 * Здесь ничего менять не нужно. Суммы, награды, правила и настройки меняются в самой таблице.
 */

const SH = {
  members: 'Участницы',
  log: 'Начисления',
  tasks: 'Задания',
  rewards: 'Награды',
  rules: 'Правила',
  settings: 'Настройки',
  bonus: 'Бонус за ноябрь',
};

const LOG_HEADER = ['Дата', 'Участница', 'За что', 'Алмазы', 'Сумма вручную', 'Уточнение (видит участница)', 'Статус обмена (ЗАЯВКА / ОТМЕНА / ВЫПОЛНЕНО)', 'Служебное (не трогать)'];
const LOG_COLS = LOG_HEADER.length;
const CANCEL_HOURS = 24;

// Задания. «Как засчитывается»: ХЭШТЕГ (сообщение в чате с хэштегом), РИЛС (ссылка на рилс в ветке из «Настроек»),
// КНОПКА (участница отмечает сама в приложении), ВРУЧНУЮ (строку вписывают в «Начисления»).
// «Как часто»: ОДИН РАЗ (за всё время), В НЕДЕЛЮ (пн–вс), В МЕСЯЦ, БЕЗ ЛИМИТА. «Сколько раз» — сколько начислений за период.
// ДЗ к новым урокам (#дз_урок5, #дз_урок6 …) распознаются сами: строка «ДЗ к уроку N» появляется автоматически.
const TASK_HEADER = ['За что', 'Алмазы', 'Как засчитывается (ХЭШТЕГ / РИЛС / КНОПКА / ВРУЧНУЮ)', 'Как часто (ОДИН РАЗ / В НЕДЕЛЮ / В МЕСЯЦ / БЕЗ ЛИМИТА)', 'Сколько раз за период', 'Хэштеги (через запятую)'];
const TASK_ROWS = [
  ['ДЗ к уроку 1', 20, 'ХЭШТЕГ', 'ОДИН РАЗ', 1, '#дз_урок1'],
  ['ДЗ к уроку 2', 20, 'ХЭШТЕГ', 'ОДИН РАЗ', 1, '#дз_урок2'],
  ['ДЗ к уроку 3', 20, 'ХЭШТЕГ', 'ОДИН РАЗ', 1, '#дз_урок3'],
  ['ДЗ к уроку 4', 20, 'ХЭШТЕГ', 'ОДИН РАЗ', 1, '#дз_урок4'],
  ['Итоги недели', 10, 'ХЭШТЕГ', 'В НЕДЕЛЮ', 1, '#разбор'],
  ['Отчёт месяца', 20, 'ХЭШТЕГ', 'В МЕСЯЦ', 1, '#отчет_месяца'],
  ['Эфир', 10, 'КНОПКА', 'В МЕСЯЦ', 1, ''],
  ['Рилс во «Взаимном активе»', 2, 'РИЛС', 'БЕЗ ЛИМИТА', '', ''],
  ['Бонус за ноябрь', 80, 'ВРУЧНУЮ', 'ОДИН РАЗ', 1, ''],
  ['Другое (сумма вручную)', 0, 'ВРУЧНУЮ', '', '', ''],
];

const REWARD_HEADER = ['Награда', 'Алмазы', 'Описание', 'Мест в месяц на весь клуб', 'Сколько раз в месяц одной участнице'];
const REWARD_ROWS = [
  ['Проверка домашнего задания', 100, 'Оля лично проверит и прокомментирует твоё ДЗ', 10, 1],
  ['Персональная обратная связь', 300, 'Рекомендации Оли по твоей конкретной работе', 5, 1],
  ['Консультация от Оли', 500, 'Индивидуальная помощь по блогу, контенту или стратегии', 3, 1],
];

const RULE_ROWS = [
  ['ДЗ к уроку — 20 💎. Сдай ДЗ в ветке «Общение» с хэштегом урока, например #дз_урок2.'],
  ['Итоги недели — 10 💎. Отчёт с хэштегом #разбор, раз в неделю.'],
  ['Отчёт месяца — 20 💎. Хэштег #отчет_месяца, раз в месяц.'],
  ['Эфир — 10 💎. Нажми «Отметить» в этом приложении, раз в месяц.'],
  ['Рилс — 2 💎 за каждый рилс, который ты кидаешь в ветку «Взаимный актив».'],
  ['Алмазы приходят сами в течение 5 минут. Если забыла хэштег, просто допиши его в сообщение.'],
  ['Обменять алмазы можно здесь, кнопкой у награды. Передумала — отмени обмен в течение 24 часов.'],
  ['За полученную награду алмазы не возвращаются. Если не смогла прийти или награду не выдали — напиши в «Поддержку».'],
  ['После окончания подписки алмазы хранятся 30 дней.'],
];

const SETTINGS_ROWS = [
  ['Тестовый режим (ДА / НЕТ)', 'ДА'],
  ['Кто видит приложение в тестовом режиме (Telegram ID через запятую)', ''],
  ['Не начислять по хэштегам (Telegram ID через запятую, например Оля)', ''],
  ['Ветка с рилсами (название)', 'Взаимный актив'],
  ['Ник Оли в Telegram (для сообщения при обмене)', ''],
  ['Поддержка (ссылка)', 'https://t.me/kuratorbeautyrosta'],
];

// ---------- Меню в таблице ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💎 Алмазы')
    .addItem('1. Создать листы', 'setup')
    .addItem('2. Ввести токен бота', 'setToken')
    .addItem('3. Обновить таблицу до новой версии', 'upgrade')
    .addItem('4. Включить начисление из чата', 'enableChat')
    .addSeparator()
    .addItem('Проверить чат сейчас', 'pollChat')
    .addItem('Начислить бонус по листу «Бонус за ноябрь»', 'grantBonus')
    .addItem('Выключить начисление из чата', 'disableChat')
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

function withEnLocale_(ss, fn) {
  // Формулы пишутся с запятыми; на время настройки ставим английскую локаль, потом возвращаем.
  const locale = ss.getSpreadsheetLocale();
  ss.setSpreadsheetLocale('en_US');
  try { fn(); } finally { ss.setSpreadsheetLocale(locale); }
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Europe/Warsaw');
  withEnLocale_(ss, function () { setupSheets_(ss); });
  try { SpreadsheetApp.getUi().alert('Листы созданы. Следующий шаг: меню «💎 Алмазы» → «2. Ввести токен бота».'); } catch (e) {}
}

function setupSheets_(ss) {
  sheet_(ss, SH.tasks, TASK_HEADER, TASK_ROWS);
  const members = sheet_(ss, SH.members, ['Имя', 'Ник', 'Telegram ID', 'Первый вход', 'Баланс'], []);
  if (members) {
    members.getRange('C:C').setNumberFormat('@');
    members.getRange('D:D').setNumberFormat('dd.MM.yyyy');
    members.setColumnWidth(1, 200);
  }
  const log = sheet_(ss, SH.log, LOG_HEADER, []);
  if (log) {
    log.getRange('A2:A').setNumberFormat('dd.MM.yyyy HH:mm');
    log.setColumnWidths(1, LOG_COLS, 160);
    log.setColumnWidth(3, 240);
  }
  sheet_(ss, SH.rewards, REWARD_HEADER, REWARD_ROWS);
  sheet_(ss, SH.rules, ['Правило (каждая строка — отдельный пункт)'], RULE_ROWS);
  sheet_(ss, SH.settings, ['Параметр', 'Значение'], SETTINGS_ROWS);
  sheet_(ss, SH.bonus, ['Имя', 'Telegram ID', 'Начислено'], []);
  formulas_(ss);
  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Arkusz1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  ss.setActiveSheet(ss.getSheetByName(SH.log));
  SpreadsheetApp.flush();
}

// Формулы: сумма строки (с учётом отмены обмена) и баланс участницы
function formulas_(ss) {
  const log = ss.getSheetByName(SH.log);
  log.getRange('D2').setFormula(
    "=ARRAYFORMULA(IF(C2:C=\"\",\"\",IF(G2:G=\"ОТМЕНА\",0,IF(E2:E<>\"\",E2:E,IFERROR(VLOOKUP(C2:C,'" + SH.tasks + "'!A:B,2,FALSE),0)))))");
  log.getRange('D1:D').setBackground('#eef3fb');
  // Списки в «Участница» и «За что» — подсказки, а не запрет (чтобы скрипт мог писать «Обмен: …»)
  log.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SH.members).getRange('A2:A'), true).setAllowInvalid(true).build());
  log.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SH.tasks).getRange('A2:A'), true).setAllowInvalid(true).build());
  log.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['ЗАЯВКА', 'ОТМЕНА', 'ВЫПОЛНЕНО'], true).setAllowInvalid(true).build());
  ss.getSheetByName(SH.members).getRange('E2').setFormula(
    "=ARRAYFORMULA(IF(A2:A=\"\",\"\",SUMIF('" + SH.log + "'!B:B,A2:A,'" + SH.log + "'!D:D)))");
}

// Переход на новую версию: старые «Задания», «Награды», «Правила» сохраняются скрытыми копиями.
function upgrade() {
  const ss = SpreadsheetApp.getActive();
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM HH:mm');
  const replace = function (name, header, rows) {
    const old = ss.getSheetByName(name);
    if (!old) { sheet_(ss, name, header, rows); return; }
    old.copyTo(ss).setName(name + ' (старое ' + stamp + ')').hideSheet();
    old.clear();
    old.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#dfe8f7');
    old.getRange(2, 1, rows.length, header.length).setValues(rows);
    old.setFrozenRows(1);
    old.autoResizeColumns(1, header.length);
  };
  replace(SH.tasks, TASK_HEADER, TASK_ROWS);
  replace(SH.rewards, REWARD_HEADER, REWARD_ROWS);
  replace(SH.rules, ['Правило (каждая строка — отдельный пункт)'], RULE_ROWS);
  const log = ss.getSheetByName(SH.log);
  log.getRange(1, 1, 1, LOG_COLS).setValues([LOG_HEADER]).setFontWeight('bold').setBackground('#dfe8f7');
  log.getRange('A2:A').clearDataValidations().setNumberFormat('dd.MM.yyyy HH:mm');
  SETTINGS_ROWS.forEach(function (r) { if (r[0].indexOf('Тестовый') !== 0) addSetting_(ss, r[0], r[1]); });
  sheet_(ss, SH.bonus, ['Имя', 'Telegram ID', 'Начислено'], []);
  withEnLocale_(ss, function () { formulas_(ss); });
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert('Готово, таблица обновлена. Старые листы сохранены скрытыми копиями.'); } catch (e) {}
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
    out = handle_(p);
  } catch (err) {
    out = { ok: false, error: 'server', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) return { ok: false, error: 'no_token' };
  const user = verify_(p.initData || '', token);
  if (!user) return { ok: false, error: 'bad_auth' };

  const ss = SpreadsheetApp.getActive();
  const settings = readSettings_(ss);
  const name = findOrAddMember_(ss, user);

  if (settings.test && settings.allowed.indexOf(String(user.id)) < 0) {
    return { ok: true, soon: true, first_name: user.first_name || '' };
  }

  const tz = ss.getSpreadsheetTimeZone();
  let result = null;
  if (p.action === 'mark') result = mark_(ss, name, p.task || '', tz);
  if (p.action === 'redeem') result = redeem_(ss, name, p.reward || '', tz, settings);
  if (p.action === 'cancel') result = cancel_(ss, name, Number(p.row) || 0);

  const now = new Date();
  const logRows = values_(ss, SH.log, LOG_COLS);
  const history = [];
  let balance = 0;
  logRows.forEach(function (r, i) {
    if (String(r[1]).trim() !== name || r[2] === '') return;
    const amount = Number(r[3]) || 0;
    balance += amount;
    const d = r[0] instanceof Date ? r[0] : null;
    const status = String(r[6]).trim().toUpperCase();
    const isRedeem = String(r[2]).indexOf('Обмен: ') === 0;
    const h = {
      date: d ? Utilities.formatDate(d, tz, 'dd.MM.yyyy') : '',
      ts: d ? d.getTime() : 0,
      title: String(r[2]).replace(' (сумма вручную)', '') + (r[5] ? ' · ' + r[5] : ''),
      amount: amount,
    };
    if (isRedeem) {
      h.row = i + 2;
      h.status = status === 'ОТМЕНА' ? 'cancelled' : status === 'ВЫПОЛНЕНО' ? 'done' : 'pending';
      h.cancellable = h.status === 'pending' && d && (now - d) < CANCEL_HOURS * 3600000;
    }
    history.push(h);
  });
  history.sort(function (a, b) { return b.ts - a.ts; });

  const rewards = rewardStatus_(ss, name, logRows, tz, now, balance);
  const rules = values_(ss, SH.rules, 1).map(function (r) { return String(r[0]).trim(); }).filter(String);

  return {
    ok: true,
    result: result,
    marked: p.action === 'mark' ? result : null,
    tasks: taskStatus_(ss, name, logRows, tz),
    name: name,
    first_name: user.first_name || '',
    balance: balance,
    history: history,
    rewards: rewards,
    rules: rules,
    support: settings.support,
    cancel_hours: CANCEL_HOURS,
  };
}

// ---------- Задания ----------

function readTasks_(ss) {
  return values_(ss, SH.tasks, 6)
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      const how = String(r[2]).trim().toUpperCase();
      const per = String(r[3]).trim().toUpperCase();
      return {
        name: String(r[0]).trim(),
        amount: Number(r[1]) || 0,
        // «ДА» — старое значение колонки, означает кнопку
        how: how.indexOf('КНОП') >= 0 || how === 'ДА' ? 'button' : how.indexOf('РИЛС') >= 0 ? 'reels' : how.indexOf('ХЭШ') >= 0 ? 'tag' : 'manual',
        period: per.indexOf('БЕЗ') >= 0 ? 'none' : per.indexOf('НЕДЕЛ') >= 0 ? 'week' : per.indexOf('МЕСЯЦ') >= 0 ? 'month' : 'once',
        limit: Math.max(1, Number(r[4]) || 1),
        tags: String(r[5] || '').split(/[\s,;]+/).map(normTag_).filter(String),
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

// Сколько раз задание уже засчитано участнице в периоде на дату when
function used_(t, name, logRows, tz, when) {
  if (t.period === 'none') return 0;
  const key = periodKey_(when, t.period, tz);
  let used = 0;
  logRows.forEach(function (r) {
    if (String(r[1]).trim() !== name || String(r[2]).trim() !== t.name) return;
    const d = r[0] instanceof Date ? r[0] : null;
    if (t.period === 'once' || (d && periodKey_(d, t.period, tz) === key)) used++;
  });
  return used;
}

function taskStatus_(ss, name, logRows, tz) {
  const now = new Date();
  return readTasks_(ss).filter(function (t) { return t.how === 'button' && t.amount > 0; }).map(function (t) {
    if (t.period === 'none') return { name: t.name, amount: t.amount, period: 'none', limit: 0, used: 0, available: true };
    const used = used_(t, name, logRows, tz, now);
    return { name: t.name, amount: t.amount, period: t.period, limit: t.limit, used: used, available: used < t.limit };
  });
}

function mark_(ss, name, taskName, tz) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = readTasks_(ss).filter(function (x) { return x.name === taskName && x.how === 'button' && x.amount > 0; })[0];
    if (!t) return { ok: false, reason: 'unknown' };
    const logRows = values_(ss, SH.log, LOG_COLS);
    if (t.period !== 'none' && used_(t, name, logRows, tz, new Date()) >= t.limit) return { ok: false, reason: 'limit' };
    ss.getSheetByName(SH.log).getRange(logRows.length + 2, 1, 1, 3).setValues([[new Date(), name, t.name]]);
    SpreadsheetApp.flush();
    return { ok: true, task: t.name, amount: t.amount };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Награды и обмен ----------

function readRewards_(ss) {
  return values_(ss, SH.rewards, 5)
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return {
        name: String(r[0]).trim(),
        price: Number(r[1]) || 0,
        desc: String(r[2] || ''),
        slots: r[3] === '' ? 0 : Number(r[3]) || 0,      // 0 / пусто = без ограничения
        perPerson: r[4] === '' ? 0 : Number(r[4]) || 0,
      };
    });
}

// Сколько обменов на награду в текущем месяце: всего по клубу и у этой участницы (отменённые не считаются)
function redeemCounts_(logRows, rewardName, name, tz, now) {
  const key = periodKey_(now, 'month', tz);
  const title = 'Обмен: ' + rewardName;
  let all = 0, mine = 0;
  logRows.forEach(function (r) {
    if (String(r[2]).trim() !== title) return;
    if (String(r[6]).trim().toUpperCase() === 'ОТМЕНА') return;
    const d = r[0] instanceof Date ? r[0] : null;
    if (!d || periodKey_(d, 'month', tz) !== key) return;
    all++;
    if (String(r[1]).trim() === name) mine++;
  });
  return { all: all, mine: mine };
}

function rewardStatus_(ss, name, logRows, tz, now, balance) {
  return readRewards_(ss).map(function (w) {
    const c = redeemCounts_(logRows, w.name, name, tz, now);
    const noSlots = w.slots > 0 && c.all >= w.slots;
    const already = w.perPerson > 0 && c.mine >= w.perPerson;
    return {
      name: w.name, price: w.price, desc: w.desc,
      left: w.slots > 0 ? Math.max(0, w.slots - c.all) : null,
      state: already ? 'already' : noSlots ? 'no_slots' : balance >= w.price ? 'can' : 'need',
    };
  });
}

function redeem_(ss, name, rewardName, tz, settings) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const w = readRewards_(ss).filter(function (x) { return x.name === rewardName; })[0];
    if (!w || w.price <= 0) return { ok: false, reason: 'unknown' };
    const now = new Date();
    const logRows = values_(ss, SH.log, LOG_COLS);
    let balance = 0;
    logRows.forEach(function (r) { if (String(r[1]).trim() === name) balance += Number(r[3]) || 0; });
    const c = redeemCounts_(logRows, w.name, name, tz, now);
    if (w.perPerson > 0 && c.mine >= w.perPerson) return { ok: false, reason: 'already' };
    if (w.slots > 0 && c.all >= w.slots) return { ok: false, reason: 'no_slots' };
    if (balance < w.price) return { ok: false, reason: 'balance' };
    const log = ss.getSheetByName(SH.log);
    const row = logRows.length + 2;
    log.getRange(row, 1, 1, 3).setValues([[now, name, 'Обмен: ' + w.name]]);
    log.getRange(row, 5).setValue(-w.price);
    log.getRange(row, 7).setValue('ЗАЯВКА');
    SpreadsheetApp.flush();
    const text = 'Привет! Я обменяла ' + w.price + ' 💎 на «' + w.name + '» в «Моих алмазах».';
    const nick = settings.olya.replace(/^@/, '').replace(/^https?:\/\/t\.me\//, '').trim();
    return {
      ok: true, reward: w.name, price: w.price, text: text,
      link: nick ? 'https://t.me/' + nick + '?text=' + encodeURIComponent(text) : '',
    };
  } finally {
    lock.releaseLock();
  }
}

function cancel_(ss, name, row) {
  if (row < 2) return { ok: false, reason: 'unknown' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(SH.log);
    const r = sh.getRange(row, 1, 1, 7).getValues()[0];
    if (String(r[1]).trim() !== name || String(r[2]).indexOf('Обмен: ') !== 0) return { ok: false, reason: 'unknown' };
    if (String(r[6]).trim().toUpperCase() !== 'ЗАЯВКА') return { ok: false, reason: 'status' };
    const d = r[0] instanceof Date ? r[0] : null;
    if (!d || new Date() - d >= CANCEL_HOURS * 3600000) return { ok: false, reason: 'late' };
    sh.getRange(row, 7).setValue('ОТМЕНА');
    SpreadsheetApp.flush();
    return { ok: true, cancelled: String(r[2]).replace('Обмен: ', '') };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Бонус за ноябрь ----------
// Лист «Бонус за ноябрь»: Имя | Telegram ID. Каждой из списка один раз пишется строка «Бонус за ноябрь».

function grantBonus() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH.bonus);
  if (!sh) return;
  const task = readTasks_(ss).filter(function (t) { return t.name === 'Бонус за ноябрь'; })[0];
  if (!task) { try { SpreadsheetApp.getUi().alert('В листе «Задания» нет строки «Бонус за ноябрь».'); } catch (e) {} return; }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let added = 0;
  try {
    const rows = values_(ss, SH.bonus, 3);
    const log = ss.getSheetByName(SH.log);
    let logRows = values_(ss, SH.log, LOG_COLS);
    rows.forEach(function (r, i) {
      const id = String(r[1]).replace(/\D/g, '');
      if (!id) return;
      const parts = String(r[0]).trim().split(/\s+/);
      const name = findOrAddMember_(ss, { id: id, first_name: parts[0] || '', last_name: parts.slice(1).join(' ') }, true);
      if (used_(task, name, logRows, ss.getSpreadsheetTimeZone(), new Date()) > 0) { sh.getRange(i + 2, 3).setValue('уже было'); return; }
      const when = new Date();
      log.getRange(logRows.length + 2, 1, 1, 3).setValues([[when, name, task.name]]);
      logRows.push([when, name, task.name, task.amount, '', '', '', '']);
      sh.getRange(i + 2, 3).setValue('да');
      added++;
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  try { SpreadsheetApp.getUi().alert('Бонус начислен: ' + added + '.'); } catch (e) {}
}

// ---------- Начисление из чата клуба ----------
// Бот алмазов добавлен в чат клуба. Раз в 5 минут скрипт забирает новые и исправленные сообщения,
// находит хэштеги и ссылки на рилсы и сам записывает строки в «Начисления».

function enableChat() {
  const ui = SpreadsheetApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) { ui.alert('Сначала введи токен бота (меню 💎 Алмазы → 2).'); return; }
  const me = tg_(token, 'getMe', {});
  if (!me.ok) { ui.alert('Telegram не принял токен. Проверь его (меню 💎 Алмазы → 2).'); return; }
  if (!me.result.can_read_all_group_messages) {
    ui.alert('Бот пока не видит сообщения в чате.\n\nВ @BotFather: /setprivacy → выбери @' + me.result.username +
      ' → Disable. Потом удали бота из группы и добавь заново. И снова нажми этот пункт меню.');
    return;
  }
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'pollChat') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pollChat').timeBased().everyMinutes(5).create();
  pollChat();
  ui.alert('Включено. Бот проверяет чат каждые 5 минут и начисляет алмазы сам.');
}

function disableChat() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'pollChat') ScriptApp.deleteTrigger(t); });
  try { SpreadsheetApp.getUi().alert('Начисление из чата выключено.'); } catch (e) {}
}

function tg_(token, method, payload) {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true,
  });
  try { return JSON.parse(r.getContentText()); } catch (e) { return { ok: false }; }
}

function norm_(s) {
  return String(s || '').toLowerCase().replace(/ё/g, 'е');
}
function normTag_(s) {
  return norm_(s).replace(/^#/, '').replace(/[_\s]/g, '');
}

// Что засчитать по тексту сообщения: список { kind: 'lesson', n } / { kind: 'tag', tag }.
// Терпит опечатки: #ДЗ_урок2, #дзурок2, #дз_урк_2, #дз урок 2, #дз2; #Разбор; #отчёт_месяца, #отчетмесяца, #отчет месяца.
function parseText_(text) {
  const t = norm_(text);
  const out = [];
  const lessonRe = /#\s*дз[\s_]*(?:у?р?о?к?)?[\s_]*№?\s*(\d{1,2})(?!\d)/g;
  let m;
  while ((m = lessonRe.exec(t))) out.push({ kind: 'lesson', n: Number(m[1]) });
  if (/#\s*разб[оа]?р/.test(t)) out.push({ kind: 'tag', tag: 'разбор' });
  if (/#\s*отч[её]?т?[\s_]*мес/.test(t)) out.push({ kind: 'tag', tag: 'отчетмесяца' });
  (t.match(/#[0-9a-zа-я_]+/g) || []).forEach(function (x) { out.push({ kind: 'tag', tag: normTag_(x) }); });
  return out;
}

// Ссылки на рилсы в тексте (Instagram reel/p, TikTok, YouTube Shorts) → короткий ключ для защиты от повторов
function reelKeys_(text) {
  const keys = [];
  const re = /(?:instagram\.com\/(?:[\w.]+\/)?(?:reels?|p|tv)\/([\w-]+))|(?:tiktok\.com\/[^\s?#]+)|(?:youtube\.com\/shorts\/([\w-]+))|(?:youtu\.be\/([\w-]+))/gi;
  let m;
  while ((m = re.exec(String(text || '')))) {
    const k = m[1] ? 'ig:' + m[1] : m[2] ? 'yt:' + m[2] : m[3] ? 'yt:' + m[3] : 'tt:' + m[0].toLowerCase().replace(/^.*tiktok\.com\//, '').replace(/\/$/, '');
    if (keys.indexOf(k) < 0) keys.push(k);
  }
  return keys;
}

function lessonTask_(ss, tasks, n) {
  const name = 'ДЗ к уроку ' + n;
  const found = tasks.filter(function (t) { return t.name === name; })[0];
  if (found) return found;
  const sample = tasks.filter(function (t) { return /^ДЗ к уроку \d+$/.test(t.name); })[0];
  const amount = sample ? sample.amount : 20;
  ss.getSheetByName(SH.tasks).appendRow([name, amount, 'ХЭШТЕГ', 'ОДИН РАЗ', 1, '#дз_урок' + n]);
  const t = { name: name, amount: amount, how: 'tag', period: 'once', limit: 1, tags: ['дзурок' + n] };
  tasks.push(t);
  return t;
}

// Ветка «Взаимный актив»: запоминаем её номер, когда видим сообщение в ней
function isReelsThread_(msg, props, threadName) {
  if (!msg.is_topic_message || !msg.message_thread_id) return false;
  const key = 'REELS_THREAD_' + msg.chat.id;
  const ftc = msg.reply_to_message && msg.reply_to_message.forum_topic_created;
  const created = msg.forum_topic_created || ftc;
  if (created && norm_(created.name).indexOf(norm_(threadName)) >= 0) props.setProperty(key, String(msg.message_thread_id));
  if (msg.forum_topic_edited && msg.forum_topic_edited.name) {
    if (norm_(msg.forum_topic_edited.name).indexOf(norm_(threadName)) >= 0) props.setProperty(key, String(msg.message_thread_id));
  }
  return props.getProperty(key) === String(msg.message_thread_id);
}

function pollChat() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('BOT_TOKEN');
  if (!token) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const offset = Number(props.getProperty('TG_OFFSET') || 0);
    const res = tg_(token, 'getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message', 'edited_message'] });
    if (!res.ok || !res.result.length) return;
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    const settings = readSettings_(ss);
    const tasks = readTasks_(ss);
    const reelsTask = tasks.filter(function (t) { return t.how === 'reels' && t.amount > 0; })[0];
    const sh = ss.getSheetByName(SH.log);
    const logRows = values_(ss, SH.log, LOG_COLS);
    const usedReels = {};
    logRows.forEach(function (r) { String(r[7] || '').split(' ').forEach(function (k) { if (k.indexOf('reel=') === 0) usedReels[k.slice(5)] = true; }); });
    let next = offset;

    // Колонку D (формула) не трогаем: пишем A:C и H отдельно
    const write = function (when, name, t, service) {
      const row = logRows.length + 2;
      sh.getRange(row, 1, 1, 3).setValues([[when, name, t.name]]);
      sh.getRange(row, 8).setValue(service);
      logRows.push([when, name, t.name, t.amount, '', '', '', service]);
    };

    res.result.forEach(function (u) {
      next = Math.max(next, u.update_id + 1);
      const msg = u.message || u.edited_message;
      if (!msg || !msg.chat || (msg.chat.type !== 'group' && msg.chat.type !== 'supergroup')) return;
      const inReels = isReelsThread_(msg, props, settings.reelsThread);
      if (!msg.from || msg.from.is_bot) return;
      const uid = String(msg.from.id);
      if (settings.excluded.indexOf(uid) >= 0) return;
      if (settings.test && settings.allowed.indexOf(uid) < 0) return;
      const text = String(msg.text || msg.caption || '');
      const when = new Date(msg.date * 1000);
      const service = 'msg=' + msg.chat.id + ':' + msg.message_id;
      let name = null;
      const member = function () { return name || (name = findOrAddMember_(ss, msg.from, true)); };

      // Хэштеги
      const seen = {};
      parseText_(text).forEach(function (h) {
        let t = null;
        if (h.kind === 'lesson') t = lessonTask_(ss, tasks, h.n);
        else t = tasks.filter(function (x) { return x.how === 'tag' && x.tags.indexOf(h.tag) >= 0; })[0];
        if (!t || t.amount <= 0 || seen[t.name]) return;
        seen[t.name] = true;
        if (t.period !== 'none' && used_(t, member(), logRows, tz, when) >= t.limit) return;
        write(when, member(), t, service);
      });

      // Рилсы во «Взаимном активе»: каждая ссылка один раз
      if (reelsTask && inReels) {
        let entText = text;
        (msg.entities || msg.caption_entities || []).forEach(function (e) { if (e.url) entText += ' ' + e.url; });
        reelKeys_(entText).forEach(function (k) {
          if (usedReels[k]) return;
          usedReels[k] = true;
          if (reelsTask.period !== 'none' && used_(reelsTask, member(), logRows, tz, when) >= reelsTask.limit) return;
          write(when, member(), reelsTask, service + ' reel=' + k);
        });
      }
    });
    props.setProperty('TG_OFFSET', String(next));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

// ---------- Служебное ----------

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

function addSetting_(ss, label, value) {
  const sh = ss.getSheetByName(SH.settings);
  if (!sh) return;
  const rows = values_(ss, SH.settings, 2);
  if (rows.some(function (r) { return String(r[0]).indexOf(label.slice(0, 12)) === 0; })) return;
  sh.getRange(rows.length + 2, 1, 1, 2).setValues([[label, value]]);
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
    excluded: get('Не начислять').split(/[\s,;]+/).filter(String),
    reelsThread: get('Ветка с рилсами') || 'Взаимный актив',
    olya: get('Ник Оли'),
    support: get('Поддержка') || 'https://t.me/kuratorbeautyrosta',
  };
}

// Находит участницу по Telegram ID; новую добавляет в лист «Участницы». Возвращает имя из колонки «Имя».
function findOrAddMember_(ss, user, haveLock) {
  const sh = ss.getSheetByName(SH.members);
  const id = String(user.id);
  const find = function () {
    const rows = values_(ss, SH.members, 3);
    for (let i = 0; i < rows.length; i++) if (String(rows[i][2]).trim() === id) return { name: String(rows[i][0]).trim(), rows: rows };
    return { name: null, rows: rows };
  };
  let f = find();
  if (f.name) return f.name;

  const lock = haveLock ? null : LockService.getScriptLock();
  if (lock) lock.waitLock(20000);
  try {
    f = find();
    if (f.name) return f.name;
    const names = f.rows.map(function (r) { return String(r[0]).trim(); });
    let name = [user.first_name, user.last_name].filter(Boolean).join(' ').trim() || ('Участница ' + id);
    if (names.indexOf(name) >= 0) name += user.username ? ' (@' + user.username + ')' : ' (' + id + ')';
    sh.getRange(f.rows.length + 2, 1, 1, 4).setValues([[name, user.username ? '@' + user.username : '', id, new Date()]]);
    return name;
  } finally {
    if (lock) lock.releaseLock();
  }
}

// Строки листа без шапки, до последней заполненной (по первым трём колонкам; формулы ARRAYFORMULA не мешают)
function values_(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const n = sh.getMaxRows() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, cols).getValues();
  let last = -1;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i].slice(0, Math.min(cols, 3)).some(function (v) { return v !== ''; })) last = i;
  }
  return vals.slice(0, last + 1);
}
