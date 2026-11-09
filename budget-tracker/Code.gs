/**
 * ТРЕКЕР БЮДЖЕТУ — Google Sheets + Apps Script
 * Етап 1: Старт, Налаштування, Операції, довідники, залежний список «Тип → Категорія».
 *
 * Як встановити: Розширення → Apps Script → вставити цей файл у Code.gs → зберегти →
 * оновити таблицю → меню «💜 Трекер» → «Побудувати трекер з нуля».
 */

// ─────────────────────────────────────────────────────────────────────────────
// ПАЛІТРА — змінюй кольори тут, решта коду бере їх звідси
// ─────────────────────────────────────────────────────────────────────────────
const THEME = {
  primary:   '#4F86B5', // основний (заголовки блоків, велика цифра)
  accent:    '#8CC4E8', // акцент (лінії під заголовками, другий ряд графіків)
  headerBg:  '#E5F1FA', // фон заголовків таблиць
  headerTxt: '#2C5677', // текст заголовків таблиць
  input:     '#FFF8D6', // поля вводу (світло-жовті)
  danger:    '#E57373', // червоний — перевищення
  warn:      '#FFD54F', // жовтий — попередження (>80%)
  ok:        '#81C784', // зелений — норма
  okBg:      '#E8F5E9', // світлий фон «норма»
  okTxt:     '#2E7D32', // текст «норма»
  dangerBg:  '#FDECEC', // світлий фон «перевищення»
  warnBg:    '#FFF6D5', // світлий фон «попередження»
  calcBg:    '#F3F8FC', // фон клітинок з формулами
  disabled:  '#CCD9E4', // неактивні поля
  text:      '#2E3A46', // основний текст
  muted:     '#8A9BAA', // підказки
  bg:        '#FFFFFF', // фон аркушів
  font:      'Montserrat',
};

// ─────────────────────────────────────────────────────────────────────────────
// НАЗВИ ТА ФОРМАТИ
// ─────────────────────────────────────────────────────────────────────────────
const SHEETS = {
  start:    'Старт',
  dash:     'Дашборд',
  settings: 'Налаштування',
  ops:      'Операції',
  lists:    '_Списки',
};

const FMT = {
  money: '#,##0 "грн"',
  pct:   '0%',
  date:  'dd.mm.yyyy',
  int:   '0',
  money0: '#,##0 "грн";-#,##0 "грн";""',            // нуль — порожньо
  days:  '0" дн.";"прострочено "0" дн.";"сьогодні"',
  signPct: '+0%;-0%;0%',
};

const HISTORY_MONTHS = 3;   // скільки попередніх циклів беремо для «середніх» (пишеться в «_Списки»)

const NR = 'BT_';             // префікс іменованих діапазонів трекера
const DATA_ROWS = 1000;       // скільки рядків готуємо в «Операціях»

const TYPES = { income: 'Дохід', expense: 'Витрата', debt: 'Борг', saving: 'Заощадження' };

/** Довідкові значення (не розрахункові дані): пишуться у прихований аркуш «_Списки». */
const REF = {
  types:  [TYPES.income, TYPES.expense, TYPES.debt, TYPES.saving],
  freq:   [['Щомісяця', 1], ['Раз на квартал', 3], ['Раз на рік', 12]],
  groups: ['Потреби', 'Бажання'],
  cycle:  ['Від основної виплати', 'Календарний місяць'],
  current: 'Поточний цикл',
  months: ['січень', 'лютий', 'березень', 'квітень', 'травень', 'червень',
           'липень', 'серпень', 'вересень', 'жовтень', 'листопад', 'грудень'],
};

/** Стовпці «Операцій» (1-based). */
const OPS = { date: 1, type: 2, amount: 3, category: 4, account: 5, comment: 6, source: 7 };
const OPS_HEADERS = ['Дата', 'Тип', 'Сума', 'Категорія', 'Рахунок', 'Коментар', 'Джерело'];

/** Для якого типу операції — який стовпець «_Списків» з категоріями. */
const TYPE_LIST = {
  [TYPES.income]:  'L_Inc',
  [TYPES.expense]: 'L_Exp',
  [TYPES.debt]:    'L_Debt',
  [TYPES.saving]:  'L_Sav',
};

// ─────────────────────────────────────────────────────────────────────────────
// СТРУКТУРА «НАЛАШТУВАНЬ»
// t: text | money | pct | date | day | int | check | list:<L_*> | f (формула)
// У формулах {Key} замінюється на клітинку того ж рядка з колонки Key.
// ─────────────────────────────────────────────────────────────────────────────
const OPS_REF = "'" + SHEETS.ops + "'";

const GENERAL_ROWS = [
  { key: 'Year',  label: 'Рік',                 t: 'int',  hint: 'Рік, за який будується вкладка «Рік».' },
  { key: 'Start', label: 'Дата старту',         t: 'date', hint: 'З якого дня ти ведеш трекер.' },
  { key: 'Calc',  label: 'Дата розрахунку',     t: 'date', hint: 'Залиш порожньою — буде «сьогодні».' },
  { key: 'Cycle', label: 'Бюджетний цикл',      t: 'list:L_Cycle', hint: 'Від дня основної виплати до наступної або календарний місяць.' },
  { key: 'Warn',  label: 'Попередження, %',     t: 'pct',  hint: 'Категорія жовта, коли витрачено більше цього.' },
  { key: 'Over',  label: 'Перевищення, %',      t: 'pct',  hint: 'Категорія червона, коли витрачено більше цього.' },
];

const RULE_ROWS = [
  { key: 'Needs', label: 'Потреби',      t: 'pct', hint: 'Житло, їжа, транспорт, борги.' },
  { key: 'Wants', label: 'Бажання',      t: 'pct', hint: 'Кафе, одяг, розваги, підписки.' },
  { key: 'Save',  label: 'Заощадження',  t: 'pct', hint: 'Подушка і цілі.' },
];

const SETTINGS_BLOCKS = [
  {
    key: 'Inc', title: 'ДОХІД', rows: 8,
    hint: 'Усі джерела грошей. Познач одне «Основне» — від його дня виплати рахується цикл. Сума може бути порожньою, якщо дохід плаває.',
    cols: [
      { key: 'Name', h: 'Джерело',        t: 'text',  w: 190 },
      { key: 'Main', h: 'Основне',        t: 'check', w: 80 },
      { key: 'Exp',  h: 'Очікувана сума', t: 'money', w: 130 },
      { key: 'Day',  h: 'День виплати',   t: 'day',   w: 110 },
    ],
  },
  {
    key: 'Reg', title: 'РЕГУЛЯРНІ ВИТРАТИ', rows: 15,
    hint: 'Обов\'язкові платежі й підписки. Для спільних витрат вкажи свою частку: оренда 30 000 × 50% = 15 000.',
    cols: [
      { key: 'Name',  h: 'Назва',                t: 'text',         w: 190 },
      { key: 'Total', h: 'Загальна сума',        t: 'money',        w: 130 },
      { key: 'Share', h: 'Моя частка',           t: 'pct',          w: 100 },
      { key: 'Freq',  h: 'Як часто',             t: 'list:L_Freq',  w: 130 },
      { key: 'Month', h: 'Місяць оплати (для квартал / рік)', t: 'list:L_Month', w: 150 },
      { key: 'Day',   h: 'День оплати',          t: 'day',          w: 100 },
      { key: 'Group', h: 'Група 50/30/20',       t: 'list:L_Group', w: 120 },
      { key: 'Sub',   h: 'Підписка',             t: 'check',        w: 85 },
      { key: 'Use',   h: 'Користуюсь',           t: 'check',        w: 100 },
      { key: 'Mine',  h: 'Моя сума',             t: 'f', fmt: 'money', w: 120,
        f: '=IF(OR({Name}="",{Total}=""),"",{Total}*IF({Share}="",1,{Share}))' },
      { key: 'Reserve', h: 'Резерв на місяць',   t: 'f', fmt: 'money', w: 130,
        f: '=IF({Mine}="","",{Mine}/IFERROR(VLOOKUP({Freq},' + NR + 'L_Freq,2,FALSE),1))' },
    ],
  },
  {
    key: 'Var', title: 'ЗМІННІ ВИТРАТИ', rows: 12,
    hint: 'Те, що щодня: продукти, кава, транспорт. Тижневий бюджет рахується сам.',
    cols: [
      { key: 'Name',   h: 'Підкатегорія',      t: 'text',         w: 190 },
      { key: 'Budget', h: 'Місячний бюджет',   t: 'money',        w: 130 },
      { key: 'Group',  h: 'Група 50/30/20',    t: 'list:L_Group', w: 120 },
      { key: 'Week',   h: 'Тижневий бюджет',   t: 'f', fmt: 'money', w: 130,
        f: '=IF(OR({Name}="",{Budget}=""),"",{Budget}*12/52)' },
    ],
  },
  {
    key: 'Debt', title: 'БОРГИ І РОЗСТРОЧКИ', rows: 8,
    hint: 'Кредити, «оплата частинами», борги друзям. Залишок рахується з «Операцій». Закрила раніше — постав галочку.',
    cols: [
      { key: 'Name',    h: 'Назва',               t: 'text',        w: 190 },
      { key: 'Pay',     h: 'Сума платежу',        t: 'money',       w: 120 },
      { key: 'Freq',    h: 'Як часто',            t: 'list:L_Freq', w: 130 },
      { key: 'Start',   h: 'Дата початку',        t: 'date',        w: 110 },
      { key: 'Count',   h: 'Кількість платежів',  t: 'int',         w: 140 },
      { key: 'Initial', h: 'Початковий залишок',  t: 'money',       w: 150 },
      { key: 'Closed',  h: 'Закрито достроково',  t: 'check',       w: 150 },
      { key: 'Paid',    h: 'Сплачено',            t: 'f', fmt: 'money', w: 120,
        f: '=IF({Name}="","",SUMIFS(' + OPS_REF + '!$C:$C,' + OPS_REF + '!$B:$B,"' + TYPES.debt + '",' + OPS_REF + '!$D:$D,{Name}))' },
      { key: 'Left',    h: 'Залишок',             t: 'f', fmt: 'money', w: 120,
        f: '=IF({Name}="","",IF({Closed},0,MAX(0,IF({Initial}="",{Pay}*{Count},{Initial})-{Paid})))' },
    ],
  },
  {
    key: 'Sav', title: 'ЗАОЩАДЖЕННЯ І ЦІЛІ', rows: 8,
    hint: 'На що відкладаєш: подушка, відпустка, мрія. Накопичено рахується з «Операцій».',
    cols: [
      { key: 'Name',    h: 'Ціль',              t: 'text',  w: 190 },
      { key: 'Target',  h: 'Цільова сума',      t: 'money', w: 130 },
      { key: 'Monthly', h: 'Відкладати щомісяця', t: 'money', w: 150 },
      { key: 'Date',    h: 'Бажана дата',       t: 'date',  w: 110 },
      { key: 'Saved',   h: 'Накопичено',        t: 'f', fmt: 'money', w: 120,
        f: '=IF({Name}="","",SUMIFS(' + OPS_REF + '!$C:$C,' + OPS_REF + '!$B:$B,"' + TYPES.saving + '",' + OPS_REF + '!$D:$D,{Name}))' },
    ],
  },
  {
    key: 'Acc', title: 'РАХУНКИ', rows: 6,
    hint: 'Картки й готівка. Познач основний: операції без рахунку записуються на нього.',
    cols: [
      { key: 'Name',     h: 'Рахунок',             t: 'text',  w: 190 },
      { key: 'Initial',  h: 'Початковий залишок',  t: 'money', w: 150 },
      { key: 'InitDate', h: 'На дату',             t: 'date',  w: 110 },
      { key: 'Main',     h: 'Основний',            t: 'check', w: 90 },
      { key: 'Balance',  h: 'Баланс зараз',        t: 'f', fmt: 'money', w: 130, f: accountBalanceFormula_() },
    ],
  },
];

/** Стартові категорії (без сум): щоб таблиця не була порожньою після побудови. */
const SEEDS = {
  Inc: [{ Name: 'Зарплата', Main: true }, { Name: 'Фріланс' }],
  Reg: [
    { Name: 'Оренда',               Share: 1, Freq: 'Щомісяця', Group: 'Потреби' },
    { Name: 'Комуналка + інтернет', Share: 1, Freq: 'Щомісяця', Group: 'Потреби' },
  ],
  Var: [
    { Name: 'Продукти і кава',  Group: 'Потреби' },
    { Name: 'Транспорт',        Group: 'Потреби' },
    { Name: 'Здоров\'я і краса', Group: 'Потреби' },
    { Name: 'Кафе і доставка',  Group: 'Бажання' },
    { Name: 'Одяг',             Group: 'Бажання' },
    { Name: 'Розваги',          Group: 'Бажання' },
    { Name: 'Подарунки',        Group: 'Бажання' },
    { Name: 'Інше',             Group: 'Бажання' },
  ],
  Debt: [],
  Sav: [{ Name: 'Подушка безпеки' }, { Name: 'Відпустка' }],
  Acc: [{ Name: 'Картка', Main: true }, { Name: 'Готівка' }],
};

// ─────────────────────────────────────────────────────────────────────────────
// МЕНЮ
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💜 Трекер')
    .addItem('Побудувати трекер з нуля', 'buildTracker')
    .addSeparator()
    .addItem('Заповнити демо-даними', 'fillDemoData')
    .addItem('Прибрати демо-дані', 'clearDemoData')
    .addSeparator()
    .addItem('Створити форму для телефону', 'createEntryForm')
    .addItem('Перевірити таблицю', 'selfCheck')
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────────────────
// ПОБУДОВА
// ─────────────────────────────────────────────────────────────────────────────
function buildTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ARG_SEP_ = null; // роздільник визначаємо заново в кожному запуску

  // Якщо трекер уже є — пропонуємо зберегти введені дані.
  const saved = captureUserData_(ss);
  const keep = saved ? askKeepData_(saved) : false;
  if (!keep) resetDemoState_();
  deleteTrackerSheets_(ss);

  // Спочатку створюємо всі аркуші, бо вони посилаються один на одного.
  const order = [SHEETS.start, SHEETS.dash, SHEETS.settings, SHEETS.ops, SHEETS.lists];
  const sh = {};
  order.forEach((name, i) => { sh[name] = ss.insertSheet(name, i); });
  removeTempSheet_(ss);
  detectArgSeparator_(sh[SHEETS.lists]);

  buildListsStatic_(sh[SHEETS.lists]);
  buildSettings_(ss, sh[SHEETS.settings]);
  buildListsDynamic_(ss, sh[SHEETS.lists]);
  buildOps_(ss, sh[SHEETS.ops]);
  buildDashboard_(ss, sh[SHEETS.dash]);
  buildStart_(sh[SHEETS.start]);
  if (keep) restoreUserData_(ss, saved);

  sh[SHEETS.lists].hideSheet();
  ss.setActiveSheet(sh[SHEETS.start]);
  SpreadsheetApp.flush();

  // Локаль і часовий пояс — лише в кінці: формули вже записані в поточній локалі,
  // а Google сам переводить збережені формули під нову (коми → «;»).
  ss.setSpreadsheetLocale('uk_UA');
  ss.setSpreadsheetTimeZone('Europe/Kyiv');
  SpreadsheetApp.flush();

  const errors = scanErrors_(ss, trackerSheetNames_());
  const msg = errors.length
    ? '⚠ Помилок у формулах: ' + errors.length + '. Перша: ' + errors[0]
    : '✓ Трекер побудовано, помилок немає. Почни з вкладки «Налаштування».';
  Logger.log(msg + (errors.length ? '\n' + errors.slice(0, 50).join('\n') : ''));
  ss.toast(msg, '💜 Трекер', 10);
}

/** Знімок полів вводу «Налаштувань» і рядків «Операцій» перед перебудовою. */
function captureUserData_(ss) {
  if (!ss.getSheetByName(SHEETS.settings) && !ss.getSheetByName(SHEETS.ops)) return null;
  const settings = {};
  inputKeys_().forEach(k => {
    const rng = ss.getRangeByName(NR + k);
    if (rng) settings[k] = rng.getValues().map(r => r[0]);
  });
  let ops = [];
  const opsSh = ss.getSheetByName(SHEETS.ops);
  if (opsSh && opsSh.getLastRow() > 1) {
    ops = opsSh.getRange(2, 1, opsSh.getLastRow() - 1, OPS_HEADERS.length).getValues()
      .filter(r => r.slice(0, OPS.comment).some(v => v !== ''));
  }
  return { settings: settings, ops: ops };
}

function askKeepData_(saved) {
  try {
    const ui = SpreadsheetApp.getUi();
    const ans = ui.alert('Перебудувати трекер',
      'Зберегти твої налаштування і ' + saved.ops.length + ' операцій?\n\n' +
      '«Так» — вкладки побудуються заново, а дані повернуться на місце.\n«Ні» — почати з чистого аркуша.',
      ui.ButtonSet.YES_NO);
    return ans === ui.Button.YES;
  } catch (e) {
    return true; // запуск з редактора без діалогу — дані не губимо
  }
}

function restoreUserData_(ss, saved) {
  Object.keys(saved.settings).forEach(k => {
    const rng = ss.getRangeByName(NR + k);
    if (!rng) return;
    const vals = saved.settings[k].slice(0, rng.getNumRows());
    while (vals.length < rng.getNumRows()) vals.push(vals.length && typeof saved.settings[k][0] === 'boolean' ? false : '');
    rng.setValues(vals.map(v => [v]));
  });
  if (saved.ops.length) {
    const sh = ss.getSheetByName(SHEETS.ops);
    const rows = saved.ops.slice(0, DATA_ROWS - 1);
    sh.getRange(2, 1, rows.length, OPS_HEADERS.length).setValues(rows);
    applyCategoryValidation_(sh, 2, rows.length, false);
  }
}

const ERROR_CODES = ['#ERROR!', '#REF!', '#NAME?', '#VALUE!', '#N/A', '#DIV/0!', '#NUM!'];

/** Повертає список «Аркуш!A1 = #ПОМИЛКА» для всіх клітинок з помилками. */
function scanErrors_(ss, names) {
  const found = [];
  names.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const vals = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).getDisplayValues();
    vals.forEach((row, r) => row.forEach((v, c) => {
      if (ERROR_CODES.indexOf(v) !== -1) found.push(name + '!' + colLetter_(c + 1) + (r + 1) + ' = ' + v);
    }));
  });
  return found;
}

function deleteTrackerSheets_(ss) {
  ss.getNamedRanges().forEach(nr => { if (nr.getName().indexOf(NR) === 0) nr.remove(); });
  const ours = Object.values(SHEETS);
  // У таблиці має лишатися хоч один аркуш — тимчасовий.
  ss.insertSheet('__tmp__' + Date.now());
  ss.getSheets().forEach(s => { if (ours.indexOf(s.getName()) !== -1) ss.deleteSheet(s); });
}

function removeTempSheet_(ss) {
  ss.getSheets().forEach(s => { if (s.getName().indexOf('__tmp__') === 0) ss.deleteSheet(s); });
}

// ─── «_Списки» ───────────────────────────────────────────────────────────────
// A типи | B дохід | C витрата | D борг | E заощадження | F рахунки | G-H частота |
// I групи | J цикл | K усі категорії | L місяці
const LISTS_LAYOUT = {
  L_Types: 'A', L_Inc: 'B', L_Exp: 'C', L_Debt: 'D', L_Sav: 'E', L_Acc: 'F',
  L_Freq: 'G', L_Group: 'I', L_Cycle: 'J', L_All: 'K', L_Month: 'L', L_Period: 'M',
};
const LISTS_MAX = 120;

function buildListsStatic_(sh) {
  const ss = sh.getParent();
  sh.getRange('A1:Q1').setValues([['Типи', 'Дохід', 'Витрата', 'Борг', 'Заощадження', 'Рахунки',
    'Частота', 'Місяців', 'Групи', 'Цикл', 'Усі категорії', 'Місяці', 'Період',
    'SPARKLINE', 'значення', '', 'Місяців історії']]).setFontWeight('bold');

  const put = (col, values) => sh.getRange(col + '2:' + col + (values.length + 1)).setValues(values.map(v => [v]));
  put('A', REF.types);
  sh.getRange(2, 7, REF.freq.length, 2).setValues(REF.freq);
  put('I', REF.groups);
  put('J', REF.cycle);
  put('L', REF.months);
  put('M', [REF.current].concat(REF.months));
  sh.getRange('Q2').setValue(HISTORY_MONTHS);
  ss.setNamedRange(NR + 'HistN', sh.getRange('Q2'));
  ss.setNamedRange(NR + 'L_Period', sh.getRange(2, 13, REF.months.length + 1, 1));

  // Параметри SPARKLINE-барів (діапазони замість {масивів} — не залежать від локалі).
  const spark = { SP_Ok: THEME.ok, SP_Warn: THEME.warn, SP_Over: THEME.danger, SP_Main: THEME.primary };
  Object.keys(spark).forEach((k, i) => {
    const top = 2 + i * 4;
    const rng = sh.getRange(top, 14, 3, 2);
    rng.setValues([['charttype', 'bar'], ['max', 1], ['color1', spark[k]]]);
    ss.setNamedRange(NR + k, rng);
  });

  ss.setNamedRange(NR + 'L_Types', sh.getRange(2, 1, REF.types.length, 1));
  ss.setNamedRange(NR + 'L_Freq',  sh.getRange(2, 7, REF.freq.length, 2));
  ss.setNamedRange(NR + 'L_FreqName', sh.getRange(2, 7, REF.freq.length, 1));
  ss.setNamedRange(NR + 'L_Group', sh.getRange(2, 9, REF.groups.length, 1));
  ss.setNamedRange(NR + 'L_Cycle', sh.getRange(2, 10, REF.cycle.length, 1));
  ss.setNamedRange(NR + 'L_Month', sh.getRange(2, 12, REF.months.length, 1));
  ['L_Inc', 'L_Exp', 'L_Debt', 'L_Sav', 'L_Acc', 'L_All'].forEach(k => {
    const c = LISTS_LAYOUT[k];
    ss.setNamedRange(NR + k, sh.getRange(c + '2:' + c + (LISTS_MAX + 1)));
  });
}

function buildListsDynamic_(ss, sh) {
  const nonBlank = expr => '=LET(v,' + expr + ',IFERROR(FILTER(v,v<>""),""))';
  const f = {
    B: nonBlank(NR + 'Inc_Name'),
    C: nonBlank('VSTACK(' + NR + 'Reg_Name,' + NR + 'Var_Name)'),
    D: nonBlank(NR + 'Debt_Name'),
    E: nonBlank(NR + 'Sav_Name'),
    F: nonBlank(NR + 'Acc_Name'),
    K: nonBlank('VSTACK(' + NR + 'Inc_Name,' + NR + 'Reg_Name,' + NR + 'Var_Name,' + NR + 'Debt_Name,' + NR + 'Sav_Name)'),
  };
  Object.keys(f).forEach(c => sh.getRange(c + '2').setFormula(F_(f[c])));
  sh.protect().setDescription('Службовий аркуш трекера').setWarningOnly(true);
}

function listRange_(ss, key) {
  return ss.getRangeByName(NR + key);
}

// ─── «Налаштування» ──────────────────────────────────────────────────────────
function buildSettings_(ss, sh) {
  styleSheet_(sh);
  sh.setColumnWidth(1, 24);
  sh.getRange('B1').setValue(spaced_('НАЛАШТУВАННЯ'))
    .setFontSize(18).setFontWeight('bold').setFontColor(THEME.primary);
  sh.getRange('B2').setValue('Заповни жовті клітинки. Сірі рахуються самі — їх не чіпай.')
    .setFontColor(THEME.muted).setFontStyle('italic');

  let row = 4;
  row = buildKeyValueBlock_(ss, sh, row, 'ЗАГАЛЬНЕ', GENERAL_ROWS);
  row = buildKeyValueBlock_(ss, sh, row, 'ПРАВИЛО 50/30/20', RULE_ROWS);

  // Значення за замовчуванням загального блоку.
  const today = todayParts_(ss);
  ss.getRangeByName(NR + 'Year').setValue(today.y);
  ss.getRangeByName(NR + 'Start').setValue(isoDate_(today.y, today.m, 1));
  ss.getRangeByName(NR + 'Cycle').setValue(REF.cycle[0]);
  ss.getRangeByName(NR + 'Warn').setValue(0.8);
  ss.getRangeByName(NR + 'Over').setValue(1);
  ss.getRangeByName(NR + 'Needs').setValue(0.5);
  ss.getRangeByName(NR + 'Wants').setValue(0.3);
  ss.getRangeByName(NR + 'Save').setValue(0.2);

  // Дата розрахунку, що діє насправді (порожня → сьогодні).
  const calc = ss.getRangeByName(NR + 'Calc');
  const calcNow = calc.offset(0, 1);
  calcNow.setFormula(F_('=IF(' + calc.getA1Notation() + '="",TODAY(),' + calc.getA1Notation() + ')'))
    .setNumberFormat(FMT.date).setFontColor(THEME.primary).setFontWeight('bold');
  ss.setNamedRange(NR + 'CalcDate', calcNow);
  calc.offset(0, 2).setValue('← діє зараз. ' + GENERAL_ROWS[2].hint);
  protect_(calcNow, 'Дата розрахунку');

  // Сума правила 50/30/20 має бути 100%.
  const needs = ss.getRangeByName(NR + 'Needs');
  needs.offset(0, 1).setFormula(F_('=IF(ROUND(SUM(' + needs.offset(0, 0, 3, 1).getA1Notation() + '),4)=1,"✓ разом 100%","⚠ разом має бути 100%")'))
    .setFontColor(THEME.primary);
  needs.offset(0, 2).setValue(RULE_ROWS[0].hint).setFontColor(THEME.muted).setFontStyle('italic');
  protect_(needs.offset(0, 1), 'Перевірка 50/30/20');

  SETTINGS_BLOCKS.forEach(block => { row = buildTableBlock_(ss, sh, row, block); });

  // Ширини колонок: беремо найширшу серед блоків.
  const widths = {};
  SETTINGS_BLOCKS.forEach(b => b.cols.forEach((c, i) => { widths[i] = Math.max(widths[i] || 0, c.w); }));
  Object.keys(widths).forEach(i => sh.setColumnWidth(2 + Number(i), widths[i]));
  sh.setColumnWidth(2, 210);

  // Перевірка «рівно одне Основне».
  const incMain = ss.getRangeByName(NR + 'Inc_Main');
  const accMain = ss.getRangeByName(NR + 'Acc_Main');
  addExactlyOneHint_(sh, incMain, blockWidth_('Inc'), 'рівно одне основне джерело');
  addExactlyOneHint_(sh, accMain, blockWidth_('Acc'), 'рівно один основний рахунок');

  // Валідація дати початку боргу ≥ дати старту.
  const debtStart = ss.getRangeByName(NR + 'Debt_Start');
  const startA1 = absA1_(ss.getRangeByName(NR + 'Start'));
  const first = debtStart.getCell(1, 1).getA1Notation();
  debtStart.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(F_('=AND(ISNUMBER(' + first + '),' + first + '>=' + startA1 + ')'))
    .setAllowInvalid(false)
    .setHelpText('Дата початку не може бути раніше дати старту трекера (Налаштування → Загальне → Дата старту).')
    .build());

  // «Користуюсь» має сенс лише для підписок — для інших рядків поле сіре.
  const sub = ss.getRangeByName(NR + 'Reg_Sub');
  const use = ss.getRangeByName(NR + 'Reg_Use');
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(F_('=$' + colLetter_(sub.getColumn()) + use.getRow() + '<>TRUE'))
      .setBackground(THEME.calcBg).setFontColor(THEME.disabled).setRanges([use]).build(),
  ]));

  sh.setFrozenRows(2);
  trimSheet_(sh, row + 2, 2 + Math.max.apply(null, SETTINGS_BLOCKS.map(b => b.cols.length)) + 1);
}

/** Блок «назва — значення — підказка». Створює іменовані діапазони NR+key. */
function buildKeyValueBlock_(ss, sh, row, title, rows) {
  writeBlockTitle_(sh, row, title, 3);
  row++;
  rows.forEach(r => {
    sh.getRange(row, 2).setValue(r.label).setFontWeight('bold');
    const cell = sh.getRange(row, 3);
    styleInput_(ss, cell, r.t);
    sh.getRange(row, 4).setValue(r.hint).setFontColor(THEME.muted).setFontStyle('italic');
    ss.setNamedRange(NR + r.key, cell);
    row++;
  });
  return row + 1;
}

/** Табличний блок: заголовок, підказка, шапка, N рядків вводу/формул. */
function buildTableBlock_(ss, sh, row, block) {
  const ncol = block.cols.length;
  writeBlockTitle_(sh, row, block.title, ncol);
  sh.getRange(row + 1, 2).setValue(block.hint).setFontColor(THEME.muted).setFontStyle('italic');
  const headRow = row + 2;
  const top = row + 3;
  const n = block.rows;

  sh.getRange(headRow, 2, 1, ncol).setValues([block.cols.map(c => c.h)])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerTxt).setFontWeight('bold').setWrap(true)
    .setVerticalAlignment('middle');

  const letter = {};
  block.cols.forEach((c, i) => { letter[c.key] = colLetter_(2 + i); });

  block.cols.forEach((c, i) => {
    const rng = sh.getRange(top, 2 + i, n, 1);
    ss.setNamedRange(NR + block.key + '_' + c.key, rng);
    if (c.t === 'f') {
      const formulas = [];
      for (let r = 0; r < n; r++) {
        const rr = top + r;
        formulas.push([F_(c.f).replace(/\{(\w+)\}/g, (m, k) => {
          if (!letter[k]) throw new Error('Невідома колонка ' + k + ' у блоці ' + block.key);
          return letter[k] + rr;
        })]);
      }
      rng.setFormulas(formulas).setNumberFormat(FMT[c.fmt] || FMT.money).setFontColor(THEME.text)
        .setBackground(THEME.calcBg);
      protect_(rng, block.title + ': ' + c.h);
    } else {
      styleInput_(ss, rng, c.t);
    }
  });

  // Стартові значення.
  const seeds = SEEDS[block.key] || [];
  if (seeds.length) {
    const values = seeds.map(s => block.cols.map(c => {
      if (c.t === 'f') return null;
      if (s[c.key] !== undefined) return s[c.key];
      return c.t === 'check' ? false : '';
    }));
    writeSkippingFormulas_(sh, top, block, values);
  }

  sh.getRange(top, 2, n, ncol).setBorder(true, true, true, true, true, true,
    THEME.headerBg, SpreadsheetApp.BorderStyle.SOLID);
  return top + n + 1;
}

function writeSkippingFormulas_(sh, top, block, values) {
  block.cols.forEach((c, i) => {
    if (c.t === 'f') return;
    sh.getRange(top, 2 + i, values.length, 1).setValues(values.map(v => [v[i]]));
  });
}

/** Оформлення поля вводу + валідація за типом. */
function styleInput_(ss, rng, t) {
  rng.setBackground(THEME.input);
  const dv = SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  if (t === 'money') {
    rng.setNumberFormat(FMT.money);
    rng.setDataValidation(dv.requireNumberGreaterThanOrEqualTo(0).setHelpText('Сума в гривнях, без знаків — лише число.').build());
  } else if (t === 'pct') {
    rng.setNumberFormat(FMT.pct);
    rng.setDataValidation(dv.requireNumberBetween(0, 1).setHelpText('Відсоток від 0% до 100%. Наприклад, 50%.').build());
  } else if (t === 'date') {
    rng.setNumberFormat(FMT.date);
    rng.setDataValidation(dv.requireDate().setHelpText('Дата у форматі дд.мм.рррр').build());
  } else if (t === 'day') {
    rng.setNumberFormat(FMT.int);
    rng.setDataValidation(dv.requireNumberBetween(1, 31).setHelpText('День місяця: від 1 до 31.').build());
  } else if (t === 'int') {
    rng.setNumberFormat(FMT.int);
    rng.setDataValidation(dv.requireNumberGreaterThanOrEqualTo(0).setHelpText('Ціле число.').build());
  } else if (t === 'check') {
    rng.insertCheckboxes();
  } else if (t.indexOf('list:') === 0) {
    const key = t.split(':')[1];
    const src = listRange_(ss, key === 'L_Freq' ? 'L_FreqName' : key);
    rng.setDataValidation(dv.requireValueInRange(src, true).setHelpText('Обери зі списку.').build());
  }
}

/** Підказка праворуч від шапки блоку, якщо галочок не рівно одна. */
function addExactlyOneHint_(sh, rng, blockCols, what) {
  const hintCell = sh.getRange(rng.getRow() - 1, 2 + blockCols + 1);
  hintCell.setFormula(F_('=IF(COUNTIF(' + rng.getA1Notation() + ',TRUE)=1,"","⚠ Познач ' + what + '")'))
    .setFontColor(THEME.danger).setFontWeight('bold');
  protect_(hintCell, 'Підказка: ' + what);
}

// ─── «Операції» ──────────────────────────────────────────────────────────────
function buildOps_(ss, sh) {
  styleSheet_(sh);
  sh.getRange(1, 1, 1, OPS_HEADERS.length).setValues([OPS_HEADERS])
    .setBackground(THEME.primary).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  [110, 130, 120, 200, 120, 260, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const n = DATA_ROWS - 1;
  const col = c => sh.getRange(2, c, n, 1);
  sh.getRange(2, 1, n, OPS.comment).setBackground(THEME.input);

  col(OPS.date).setNumberFormat(FMT.date).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireDate().setAllowInvalid(false).setHelpText('Дата у форматі дд.мм.рррр. Порожня — поставиться сьогодні.').build());
  col(OPS.type).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange_(ss, 'L_Types'), true).setAllowInvalid(false)
    .setHelpText('Дохід, Витрата, Борг або Заощадження.').build());
  col(OPS.amount).setNumberFormat(FMT.money).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0).setAllowInvalid(false).setHelpText('Сума більше нуля.').build());
  // Запасна валідація: будь-яка відома категорія. Після вибору типу onEdit звужує список.
  col(OPS.category).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange_(ss, 'L_All'), true).setAllowInvalid(false)
    .setHelpText('Спочатку обери тип — тут з\'являться його категорії.').build());
  col(OPS.account).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange_(ss, 'L_Acc'), true).setAllowInvalid(false)
    .setHelpText('Порожньо — основний рахунок.').build());
  col(OPS.source).setFontColor(THEME.muted);

  [['O_Date', OPS.date], ['O_Type', OPS.type], ['O_Amount', OPS.amount], ['O_Category', OPS.category],
   ['O_Account', OPS.account], ['O_Source', OPS.source]].forEach(p => ss.setNamedRange(NR + p[0], col(p[1])));

  sh.getRange(1, 1, DATA_ROWS, OPS_HEADERS.length).createFilter();
  sh.hideColumns(OPS.source);
  protect_(sh.getRange(1, 1, 1, OPS_HEADERS.length), 'Шапка «Операцій»');
  trimSheet_(sh, DATA_ROWS, OPS_HEADERS.length);
}

// ─── «Старт» ─────────────────────────────────────────────────────────────────
const START = { formRow: 14, checkRow: 18, checkLines: 10 };

function buildStart_(sh) {
  styleSheet_(sh);
  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 40);
  sh.setColumnWidth(3, 720);
  sh.getRange('B2').setValue(spaced_('ТРЕКЕР БЮДЖЕТУ')).setFontSize(22).setFontWeight('bold').setFontColor(THEME.primary);
  sh.getRange('B3').setValue('Вписуєш суму — таблиця сама рахує, скільки можна витратити до наступної виплати.')
    .setFontColor(THEME.muted).setFontStyle('italic');

  writeBlockTitle_(sh, 5, 'ЯК ПОЧАТИ', 2);
  const steps = [
    'Відкрий «Налаштування» і заповни жовті клітинки: дохід і день основної виплати, обов\'язкові платежі, бюджети на щодень, борги, цілі, рахунки.',
    'Прийшли гроші або щось купила — додай рядок в «Операції»: сума, тип, категорія. Дата і тип «Витрата» поставляться самі.',
    'З телефона ще швидше — через форму (посилання нижче). Це 10 секунд прямо на касі.',
    'Дивись «Дашборд»: велика цифра — скільки можна витратити до наступної виплати і скільки це на день.',
    'Раз на тиждень заглядай у «Рік», «Цілі», «Борги» і «Підписки». Пропустила кілька днів — нічого страшного, просто допиши.',
  ];
  steps.forEach((s, i) => {
    sh.getRange(6 + i, 2).setValue(i + 1).setFontSize(16).setFontWeight('bold').setFontColor(THEME.accent)
      .setHorizontalAlignment('center').setVerticalAlignment('top');
    sh.getRange(6 + i, 3).setValue(s).setWrap(true).setVerticalAlignment('top');
    sh.setRowHeight(6 + i, 42);
  });

  writeBlockTitle_(sh, START.formRow - 1, 'ФОРМА ДЛЯ ТЕЛЕФОНУ', 2);
  sh.getRange(START.formRow, 3).setValue('Ще не створена: меню «💜 Трекер» → «Створити форму для телефону».')
    .setFontColor(THEME.muted);

  writeBlockTitle_(sh, START.checkRow - 1, 'ПЕРЕВІРКА ТАБЛИЦІ', 2);
  sh.getRange(START.checkRow, 3).setValue('Меню «💜 Трекер» → «Перевірити таблицю».').setFontColor(THEME.muted);
  trimSheet_(sh, 40, 4);
}

// ─────────────────────────────────────────────────────────────────────────────
// ЗАЛЕЖНИЙ СПИСОК І ЗРУЧНОСТІ ПРИ ВВЕДЕННІ
// ─────────────────────────────────────────────────────────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  if (name === SHEETS.ops) handleOpsEdit_(e);
  else if (name === SHEETS.settings) handleSettingsEdit_(e);
}

function handleOpsEdit_(e) {
  const r = e.range;
  const sh = r.getSheet();
  const top = Math.max(2, r.getRow());
  const bottom = r.getLastRow();
  if (bottom < 2) return;
  const c1 = r.getColumn();
  const c2 = r.getLastColumn();
  const touches = c => c >= c1 && c <= c2;
  const n = bottom - top + 1;

  const rng = sh.getRange(top, 1, n, OPS.account);
  const vals = rng.getValues();
  let changed = false;
  let typeSet = touches(OPS.type);

  // Вписала лише суму → сьогоднішня дата і тип «Витрата».
  if (touches(OPS.amount)) {
    const today = todayIso_(sh.getParent());
    vals.forEach(v => {
      if (v[OPS.amount - 1] === '') return;
      if (v[OPS.date - 1] === '') { v[OPS.date - 1] = today; changed = true; }
      if (v[OPS.type - 1] === '') { v[OPS.type - 1] = TYPES.expense; changed = true; typeSet = true; }
    });
  }
  if (changed) {
    sh.getRange(top, OPS.date, n, 2).setValues(vals.map(v => [v[OPS.date - 1], v[OPS.type - 1]]));
  }
  if (typeSet) applyCategoryValidation_(sh, top, n, touches(OPS.type));
}

/**
 * Ставить на клітинки «Категорія» список категорій відповідного типу.
 * clearInvalid=true — стирає категорію, якої немає у списку нового типу.
 */
function applyCategoryValidation_(sh, top, n, clearInvalid) {
  const ss = sh.getParent();
  const types = sh.getRange(top, OPS.type, n, 1).getValues();
  const catRng = sh.getRange(top, OPS.category, n, 1);
  const cats = catRng.getValues();
  const cache = {};
  const rules = [];
  let cleared = false;

  types.forEach((t, i) => {
    const key = TYPE_LIST[t[0]] || 'L_All';
    if (!cache[key]) {
      const src = listRange_(ss, key);
      cache[key] = {
        rule: SpreadsheetApp.newDataValidation().requireValueInRange(src, true).setAllowInvalid(false)
          .setHelpText(t[0] ? 'Категорії для типу «' + t[0] + '».' : 'Спочатку обери тип.').build(),
        values: src.getValues().map(v => v[0]).filter(String),
      };
    }
    rules.push([cache[key].rule]);
    if (clearInvalid && cats[i][0] !== '' && cache[key].values.indexOf(cats[i][0]) === -1) {
      cats[i][0] = ''; cleared = true;
    }
  });
  if (cleared) catRng.setValues(cats);
  catRng.setDataValidations(rules);
}

/** «Основне» джерело і «Основний» рахунок — лише одна галочка в блоці. */
function handleSettingsEdit_(e) {
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  if (e.range.getValue() !== true) return;
  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  ['Inc_Main', 'Acc_Main'].forEach(key => {
    const rng = ss.getRangeByName(NR + key);
    if (!rng || !within_(e.range, rng)) return;
    const idx = e.range.getRow() - rng.getRow();
    rng.setValues(rng.getValues().map((v, i) => [i === idx]));
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ДАШБОРД
// Усі формули пишуться через X_(): «@Name» → іменований діапазон BT_Name.
// Цикл: [@C_Start; @C_Next) — права межа виключна, тож час у даті нічого не ламає.
// ─────────────────────────────────────────────────────────────────────────────
const DASH = {
  cols: 7,              // B..H
  helperLabelCol: 18,   // R — службові розрахунки (приховані)
  helperValCol: 19,     // S
  chartCol: 10,         // J — графіки
  periodRow: 2,
  bigHeadRow: 4, bigTop: 5, sentenceRow: 8, expectRow: 9,
  calcHeadRow: 11, calcTop: 12,
  blocksTop: 20,
};

/** Сума операцій типу type (і категорії cat, якщо задано) у вибраному циклі. */
function sumCycle_(type, cat) {
  return 'SUMIFS(@O_Amount,@O_Type,"' + type + '"' + (cat ? ',@O_Category,' + cat : '') +
    ',@O_Date,">="&@C_Start,@O_Date,"<"&@C_Next)';
}

/** Чи припадає платіж (день g, кожні k місяців, базовий місяць f) на цикл; результат — дата або "". */
function dueDateLet_(dayExpr, freqExpr, monthExpr, minDateExpr) {
  const okMin = minDateExpr ? ',d>=' + minDateExpr : '';
  return 'LET(st,@C_Start,en,@C_End,g,' + dayExpr + ',k,IFERROR(VLOOKUP(' + freqExpr + ',@L_Freq,2,FALSE),1),f,' + monthExpr + ',' +
    'pa,DATE(YEAR(st),MONTH(st),MIN(g,DAY(EOMONTH(st,0)))),' +
    'pb,DATE(YEAR(en),MONTH(en),MIN(g,DAY(EOMONTH(en,0)))),' +
    'IF(AND(pa>=st,pa<=en,MOD(MONTH(pa)-f,k)=0' + okMin.replace(/d>=/, 'pa>=') + '),pa,' +
    'IF(AND(pb>=st,pb<=en,MOD(MONTH(pb)-f,k)=0' + okMin.replace(/d>=/, 'pb>=') + '),pb,"")))';
}

function buildDashboard_(ss, sh) {
  styleSheet_(sh);
  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 230);
  [130, 120, 120, 120, 120, 130].forEach((w, i) => sh.setColumnWidth(3 + i, w));
  sh.setColumnWidth(9, 24);

  const H = (key, row) => sh.getRange(row, DASH.helperValCol);

  // ── 1. Розкладка блоків (спочатку лише рядки, щоб формули могли посилатися одна на одну)
  const blocks = dashBlocks_();
  let row = DASH.blocksTop;
  const pos = {};
  blocks.forEach(b => {
    pos[b.key] = { title: row, head: row + 1, top: row + 2, bottom: row + 1 + b.n, total: b.total ? row + 2 + b.n : null };
    row = row + 2 + b.n + (b.total ? 1 : 0) + 2;
  });
  const lastRow = row;
  const colRange = (bk, letter) => letter + pos[bk].top + ':' + letter + pos[bk].bottom;

  // ── 2. Іменовані клітинки: період, службові розрахунки, рядки «як рахується»
  const period = sh.getRange(DASH.periodRow, 3);
  ss.setNamedRange(NR + 'D_Period', period);
  const helpers = dashHelpers_();
  helpers.forEach((h, i) => ss.setNamedRange(NR + h.key, H(h.key, 1 + i)));
  const calc = dashCalcLines_(colRange);
  calc.forEach((c, i) => { if (c.key) ss.setNamedRange(NR + c.key, sh.getRange(DASH.calcTop + i, 3)); });

  // ── 3. Шапка і вибір періоду
  sh.getRange('B1').setValue(spaced_('ДАШБОРД')).setFontSize(18).setFontWeight('bold').setFontColor(THEME.primary);
  sh.getRange(DASH.periodRow, 2).setValue('Період').setFontWeight('bold').setHorizontalAlignment('right');
  period.setValue(REF.current).setBackground(THEME.input).setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listRange_(ss, 'L_Period'), true)
      .setAllowInvalid(false).setHelpText('Поточний цикл або місяць цього року.').build());
  sh.getRange(DASH.periodRow, 4, 1, 3).merge()
    .setFormula(X_('=TEXT(@C_Start,"dd.mm.yyyy")&" – "&TEXT(@C_End,"dd.mm.yyyy")')).setFontColor(THEME.muted);

  // ── 4. Службові розрахунки (R:S, приховані)
  helpers.forEach((h, i) => {
    sh.getRange(1 + i, DASH.helperLabelCol).setValue(h.label);
    const c = H(h.key, 1 + i).setFormula(X_(h.f));
    if (h.fmt) c.setNumberFormat(h.fmt);
  });
  const helperRow = key => 1 + helpers.findIndex(h => h.key === key);
  const hAbs = key => '$' + colLetter_(DASH.helperValCol) + '$' + helperRow(key);

  // ── 5. ВЕЛИКА ЦИФРА
  writeBlockTitle_(sh, DASH.bigHeadRow, 'МОЖУ ВИТРАТИТИ ДО НАСТУПНОЇ ВИПЛАТИ', DASH.cols);
  const big = sh.getRange(DASH.bigTop, 2, 3, 4).merge();
  big.setFormula(X_('=IF(@C_Status="empty","Внеси перший дохід",@D_Free)'))
    .setNumberFormat(FMT.money).setFontSize(40).setFontWeight('bold').setFontColor(THEME.primary)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(THEME.headerBg);
  const side = [
    ['Наступна виплата', '=@C_Next', FMT.date],
    ['Днів до неї', '=IF(@C_IsCur,@C_Days,"—")', FMT.int],
    ['Ліміт на день', '=IF(@C_IsCur,@C_Limit,"—")', FMT.money],
  ];
  side.forEach((s, i) => {
    sh.getRange(DASH.bigTop + i, 6).setValue(s[0]).setFontColor(THEME.muted);
    sh.getRange(DASH.bigTop + i, 7, 1, 2).merge().setFormula(X_(s[1])).setNumberFormat(s[2])
      .setFontWeight('bold').setFontSize(13).setFontColor(THEME.headerTxt);
    sh.setRowHeight(DASH.bigTop + i, 34);
  });
  sh.getRange(DASH.sentenceRow, 2, 1, DASH.cols).merge().setFormula(X_(
    '=IF(@C_Status="empty","Додай дохід в «Операції» — і тут з\'явиться твоя цифра.",' +
    'IF(@C_IsCur,IF(@D_Free<0,"⚠ Перевитрата на "&FIXED(-@D_Free,0)&" грн. Подивись блок «Змінні» нижче.",' +
    '"≈ "&FIXED(@C_Limit,0)&" грн на день, "&@C_Days&" дн. до виплати"&' +
    'IF(@C_Limit>=@C_Avg," — ти в нормі 💚"," — менше, ніж ти зазвичай витрачаєш ("&FIXED(@C_Avg,0)&" грн/день)")),' +
    'IF(@C_Past,"Цикл завершено — це залишок на його кінець.","Цей цикл ще не почався.")))'))
    .setFontSize(13).setFontWeight('bold').setFontColor(THEME.headerTxt).setHorizontalAlignment('center');
  sh.getRange(DASH.expectRow, 2, 1, DASH.cols).merge().setFormula(X_(
    '=IF(AND(@C_IsCur,SUM(' + colRange('inc', 'F') + ')>0),"Ще очікується доходу: "&FIXED(SUM(' + colRange('inc', 'F') +
    '),0)&" грн — у цифру він увійде, коли прийде.","")'))
    .setFontColor(THEME.muted).setFontStyle('italic').setHorizontalAlignment('center');

  // Колір великої цифри: зелений / червоний / сірий.
  const rules = [];
  const cf = (formula, rng, bg, fg) => rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(F_(formula)).setBackground(bg).setFontColor(fg).setRanges([rng]).build());
  cf('=' + hAbs('C_Status') + '="bad"', big, THEME.dangerBg, THEME.danger);
  cf('=' + hAbs('C_Status') + '="ok"', big, THEME.okBg, THEME.okTxt);
  cf('=' + hAbs('C_Status') + '="empty"', big, THEME.headerBg, THEME.muted);

  // ── 6. ЯК РАХУЄТЬСЯ ЦИФРА
  writeBlockTitle_(sh, DASH.calcHeadRow, 'ЯК РАХУЄТЬСЯ ЦИФРА', 2);
  calc.forEach((c, i) => {
    const r = DASH.calcTop + i;
    sh.getRange(r, 2).setValue(c.label).setFontWeight(c.bold ? 'bold' : 'normal');
    sh.getRange(r, 3).setFormula(X_(c.f)).setNumberFormat(FMT.money).setFontWeight(c.bold ? 'bold' : 'normal')
      .setBackground(THEME.calcBg);
  });
  sh.getRange(DASH.calcTop + calc.length - 1, 2, 1, 2).setBorder(true, null, null, null, null, null,
    THEME.primary, SpreadsheetApp.BorderStyle.SOLID);

  // ── 7. Блоки
  blocks.forEach(b => writeDashBlock_(sh, b, pos[b.key], colRange));

  // Умовне форматування блоків.
  const p = pos.var;
  const varPct = 'F' + p.top;
  const varRows = sh.getRange(p.top, 5, b_n_(blocks, 'var'), 2);
  cf('=AND(ISNUMBER($' + varPct + '),$' + varPct + '>' + hAbs('C_Over') + ')', varRows, THEME.dangerBg, THEME.danger);
  cf('=AND(ISNUMBER($' + varPct + '),$' + varPct + '>' + hAbs('C_Warn') + ')', varRows, THEME.warnBg, THEME.text);
  const reg = pos.reg;
  cf('=$F' + reg.top + '="✓"', sh.getRange(reg.top, 6, b_n_(blocks, 'reg'), 1), THEME.okBg, THEME.okTxt);
  cf('=AND(ISNUMBER($G' + reg.top + '),$G' + reg.top + '<0)', sh.getRange(reg.top, 7, b_n_(blocks, 'reg'), 1), THEME.dangerBg, THEME.danger);
  const debt = pos.debt;
  cf('=$G' + debt.top + '="✓ закрито"', sh.getRange(debt.top, 7, b_n_(blocks, 'debt'), 1), THEME.okBg, THEME.okTxt);
  sh.setConditionalFormatRules(rules);

  // ── 8. Графіки
  const varChart = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(p.head, 2, b_n_(blocks, 'var') + 1, 1))
    .addRange(sh.getRange(p.head, 3, b_n_(blocks, 'var') + 1, 2))
    .setNumHeaders(1)
    .setOption('title', 'Змінні: бюджет vs факт')
    .setOption('colors', [THEME.headerBg, THEME.accent])
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 520).setOption('height', 300)
    .setPosition(p.title, DASH.chartCol, 0, 0)
    .build();
  sh.insertChart(varChart);

  const dist = pos.dist;
  const pie = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange(dist.top, 2, b_n_(blocks, 'dist'), 2))
    .setOption('title', 'Розподіл доходу')
    .setOption('pieHole', 0.45)
    .setOption('colors', [THEME.primary, THEME.accent, THEME.danger, THEME.ok, THEME.headerBg])
    .setOption('legend', { position: 'right' })
    .setOption('width', 520).setOption('height', 300)
    .setPosition(dist.title, DASH.chartCol, 0, 0)
    .build();
  sh.insertChart(pie);

  // ── 9. Захист, приховування, розмір
  sh.hideColumns(DASH.helperLabelCol, 2);
  sh.setFrozenRows(DASH.periodRow);
  protect_(sh.getRange(1, 1, 1, DASH.helperValCol), 'Дашборд: шапка');
  protect_(sh.getRange(DASH.periodRow, 4, 1, DASH.helperValCol - 3), 'Дашборд: дати періоду');
  protect_(sh.getRange(DASH.periodRow + 1, 1, lastRow - DASH.periodRow, DASH.helperValCol), 'Дашборд');
  trimSheet_(sh, lastRow, DASH.helperValCol);
}

function b_n_(blocks, key) {
  return blocks.filter(b => b.key === key)[0].n;
}

/** Службові клітинки в прихованих R:S. */
function dashHelpers_() {
  return [
    { key: 'C_Today', label: 'Дата розрахунку', f: '=@CalcDate', fmt: FMT.date },
    { key: 'C_PayDay', label: 'День основної виплати',
      f: '=IF(@Cycle=INDEX(@L_Cycle,2),1,IFERROR(INDEX(FILTER(@Inc_Day,@Inc_Main=TRUE,@Inc_Day<>""),1),1))' },
    { key: 'C_Start', label: 'Початок циклу', fmt: FMT.date,
      f: '=LET(d,@C_PayDay,t,@C_Today,m,IFERROR(MATCH(@D_Period,@L_Month,0),0),' +
         'thisP,DATE(YEAR(t),MONTH(t),MIN(d,DAY(EOMONTH(t,0)))),pm,EOMONTH(t,-1),' +
         'prevP,DATE(YEAR(pm),MONTH(pm),MIN(d,DAY(pm))),' +
         'IF(m=0,IF(t>=thisP,thisP,prevP),DATE(@Year,m,MIN(d,DAY(EOMONTH(DATE(@Year,m,1),0))))))' },
    { key: 'C_Next', label: 'Наступна виплата', fmt: FMT.date,
      f: '=LET(d,@C_PayDay,n,EOMONTH(@C_Start,0)+1,DATE(YEAR(n),MONTH(n),MIN(d,DAY(EOMONTH(n,0)))))' },
    { key: 'C_End', label: 'Останній день циклу', f: '=@C_Next-1', fmt: FMT.date },
    { key: 'C_IsCur', label: 'Поточний цикл?', f: '=AND(@C_Today>=@C_Start,@C_Today<@C_Next)' },
    { key: 'C_Past', label: 'Цикл минув?', f: '=@C_Today>=@C_Next' },
    { key: 'C_Days', label: 'Днів до виплати', f: '=IF(@C_IsCur,@C_Next-@C_Today,0)' },
    { key: 'C_DaysTotal', label: 'Днів у циклі', f: '=@C_Next-@C_Start' },
    { key: 'C_Warn', label: 'Поріг попередження', f: '=@Warn' },
    { key: 'C_Over', label: 'Поріг перевищення', f: '=@Over' },
    { key: 'C_HistFrom', label: 'Історія з', f: '=MAX(EDATE(@C_Start,-@HistN),@Start)', fmt: FMT.date },
    { key: 'C_Avg', label: 'Середні змінні витрати на день',
      f: '=LET(from,@C_HistFrom,days,@C_Start-from,' +
         's,SUMPRODUCT(N(@O_Amount)*(@O_Type="' + TYPES.expense + '")*(@O_Date>=from)*(@O_Date<@C_Start)*' +
         '(@O_Category<>"")*ISNUMBER(MATCH(@O_Category,@Var_Name,0))),' +
         'IF(AND(days>0,s>0),s/days,IFERROR(SUM(@Var_Budget)/@C_DaysTotal,0)))' },
    { key: 'C_Limit', label: 'Ліміт на день', f: '=IF(@C_IsCur,@D_Free/MAX(1,@C_Days),0)' },
    { key: 'C_Status', label: 'Статус цифри',
      f: '=IF(@D_Income=0,"empty",IF(OR(@D_Free<0,AND(@C_IsCur,@C_Limit<@C_Avg)),"bad","ok"))' },
  ];
}

/** Рядки «як рахується цифра». Останній — сама велика цифра. */
function dashCalcLines_(colRange) {
  return [
    { key: 'D_Income', label: 'Дохід за цикл (+)', f: '=' + sumCycle_(TYPES.income) },
    { key: 'D_Spent', label: '− Уже витрачено (усі операції циклу)',
      f: '=' + sumCycle_(TYPES.expense) + '+' + sumCycle_(TYPES.debt) + '+' + sumCycle_(TYPES.saving) },
    { key: 'D_RegLeft', label: '− Ще сплатити: регулярні', f: '=SUM(' + colRange('reg', 'H') + ')' },
    { key: 'D_DebtLeft', label: '− Ще сплатити: борги', f: '=SUM(' + colRange('debt', 'E') + ')' },
    { key: 'D_SavLeft', label: '− Ще відкласти', f: '=SUM(' + colRange('sav', 'E') + ')' },
    { key: 'D_Free', label: 'Підсумок: можу витратити до наступної виплати', bold: true,
      f: '=@D_Income-@D_Spent-@D_RegLeft-@D_DebtLeft-@D_SavLeft' },
  ];
}

/**
 * Опис блоків дашборду. Кожна колонка: h — заголовок, f(i, r) — формула для i-го рядка
 * (i — номер рядка в блоці «Налаштувань», r — рядок аркуша), fmt — формат, t — тип підсумку.
 */
function dashBlocks_() {
  const inc = TYPES.income, exp = TYPES.expense, debt = TYPES.debt, sav = TYPES.saving;
  const need = 'INDEX(@L_Group,1)';
  return [
    {
      key: 'inc', title: 'ДОХІД', n: 8, total: true,
      hint: 'План — очікувана сума з «Налаштувань» або середнє за попередні цикли.',
      cols: [
        { h: 'Джерело', f: i => '=INDEX(@Inc_Name,' + i + ')', tot: 'label' },
        { h: 'План', fmt: FMT.money0, tot: 'sum',
          f: (i, r) => '=IF(B' + r + '="",0,IF(INDEX(@Inc_Exp,' + i + ')<>"",INDEX(@Inc_Exp,' + i + '),' +
            'LET(m,IFERROR(MIN(@HistN,DATEDIF(@Start,@C_Start,"M")),0),IF(m=0,0,' +
            'SUMIFS(@O_Amount,@O_Type,"' + inc + '",@O_Category,B' + r + ',@O_Date,">="&EDATE(@C_Start,-m),@O_Date,"<"&@C_Start)/m))))' },
        { h: 'Факт', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(inc, 'B' + r) + ')' },
        { h: '%', fmt: FMT.pct, tot: 'ratio:D:C', f: (i, r) => '=IF(OR(B' + r + '="",C' + r + '=0),"",D' + r + '/C' + r + ')' },
        { h: 'Ще очікується', fmt: FMT.money0, tot: 'sum',
          f: (i, r) => '=IF(OR(B' + r + '="",NOT(@C_IsCur)),0,MAX(0,C' + r + '-D' + r + '))' },
        { h: 'Прогрес', f: (i, r) => '=IF(E' + r + '="","",SPARKLINE(MIN(E' + r + ',1),@SP_Main))' },
      ],
    },
    {
      key: 'reg', title: 'РЕГУЛЯРНІ ПЛАТЕЖІ', n: 15, total: true,
      hint: 'Оплатила — внеси операцію з цією категорією, галочка ✓ поставиться сама.',
      cols: [
        { h: 'Назва', f: i => '=INDEX(@Reg_Name,' + i + ')', tot: 'label' },
        { h: 'Моя сума', fmt: FMT.money0, f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Reg_Mine,' + i + ')))' },
        { h: 'Дата оплати', fmt: FMT.date,
          f: (i, r) => '=IF(B' + r + '="","",' + dueDateLet_(
            'IF(INDEX(@Reg_Day,' + i + ')="",DAY(@C_Start),INDEX(@Reg_Day,' + i + '))',
            'INDEX(@Reg_Freq,' + i + ')',
            'IFERROR(MATCH(INDEX(@Reg_Month,' + i + '),@L_Month,0),MONTH(@Start))') + ')' },
        { h: 'Сплачено', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(exp, 'B' + r) + ')' },
        { h: 'Оплачено', f: (i, r) => '=IF(AND(B' + r + '<>"",C' + r + '>0,E' + r + '>=C' + r + '),"✓","")' },
        { h: 'Днів до оплати', fmt: FMT.days,
          f: (i, r) => '=IF(OR(D' + r + '="",F' + r + '="✓",@C_Past),"",D' + r + '-@C_Today)' },
        { h: 'Ще сплатити', fmt: FMT.money0, tot: 'sum',
          f: (i, r) => '=IF(OR(D' + r + '="",@C_Past),0,MAX(0,C' + r + '-E' + r + '))' },
      ],
    },
    {
      key: 'var', title: 'ЗМІННІ ВИТРАТИ', n: 12, total: true,
      hint: 'Жовтий — витрачено понад поріг попередження, червоний — понад бюджет.',
      cols: [
        { h: 'Категорія', f: i => '=INDEX(@Var_Name,' + i + ')', tot: 'label' },
        { h: 'Бюджет', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Var_Budget,' + i + ')))' },
        { h: 'Факт', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(exp, 'B' + r) + ')' },
        { h: 'Залишок', fmt: FMT.money, tot: 'diff:C:D', f: (i, r) => '=IF(B' + r + '="","",C' + r + '-D' + r + ')' },
        { h: '%', fmt: FMT.pct, tot: 'ratio:D:C', f: (i, r) => '=IF(OR(B' + r + '="",C' + r + '=0),"",D' + r + '/C' + r + ')' },
        { h: 'Прогрес', f: (i, r) => '=IF(F' + r + '="","",SPARKLINE(MIN(F' + r + ',1),CHOOSE(1+(F' + r + '>@C_Warn)+(F' + r +
            '>@C_Over),@SP_Ok,@SP_Warn,@SP_Over)))' },
      ],
    },
    {
      key: 'debt', title: 'БОРГИ І РОЗСТРОЧКИ', n: 8, total: true,
      hint: 'Платіж цього циклу береться з графіка в «Налаштуваннях». Закриті достроково — ✓.',
      cols: [
        { h: 'Назва', f: i => '=INDEX(@Debt_Name,' + i + ')', tot: 'label' },
        { h: 'Платіж цього циклу', fmt: FMT.money0, tot: 'sum',
          f: (i, r) => '=IF(B' + r + '="",0,LET(sd,INDEX(@Debt_Start,' + i + '),dd,' + dueDateLet_(
            'IF(sd="",DAY(@C_Start),DAY(sd))', 'INDEX(@Debt_Freq,' + i + ')', 'IF(sd="",MONTH(@C_Start),MONTH(sd))', 'N(sd)') +
            ',IF(AND(dd<>"",INDEX(@Debt_Closed,' + i + ')<>TRUE),MIN(N(INDEX(@Debt_Pay,' + i + ')),N(INDEX(@Debt_Left,' + i + '))+D' + r + '),0)))' },
        { h: 'Сплачено в циклі', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(debt, 'B' + r) + ')' },
        { h: 'Ще сплатити', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(@C_Past,0,MAX(0,C' + r + '-D' + r + '))' },
        { h: 'Залишок боргу', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Debt_Left,' + i + ')))' },
        { h: 'Прогноз закриття',
          f: (i, r) => '=IF(B' + r + '="","",IF(F' + r + '=0,"✓ закрито",LET(p,N(INDEX(@Debt_Pay,' + i + ')),' +
            'k,IFERROR(VLOOKUP(INDEX(@Debt_Freq,' + i + '),@L_Freq,2,FALSE),1),IF(p=0,"—",' +
            'LET(n,ROUNDUP(F' + r + '/p,0),d,EDATE(@C_Start,IF(E' + r + '>0,(n-1)*k,n*k)),INDEX(@L_Month,MONTH(d))&" "&YEAR(d))))))' },
      ],
    },
    {
      key: 'sav', title: 'ЗАОЩАДЖЕННЯ І ЦІЛІ', n: 8, total: true,
      hint: 'План — «відкладати щомісяця» з «Налаштувань». Досягнута ціль більше не просить грошей.',
      cols: [
        { h: 'Ціль', f: i => '=INDEX(@Sav_Name,' + i + ')', tot: 'label' },
        { h: 'План на цикл', fmt: FMT.money0, tot: 'sum',
          f: (i, r) => '=IF(B' + r + '="",0,IF(AND(INDEX(@Sav_Target,' + i + ')<>"",N(INDEX(@Sav_Saved,' + i + '))-D' + r +
            '>=N(INDEX(@Sav_Target,' + i + '))),0,N(INDEX(@Sav_Monthly,' + i + '))))' },
        { h: 'Відкладено в циклі', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(sav, 'B' + r) + ')' },
        { h: 'Ще відкласти', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(@C_Past,0,MAX(0,C' + r + '-D' + r + '))' },
        { h: 'Накопичено', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Sav_Saved,' + i + ')))' },
        { h: 'Ціль', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Sav_Target,' + i + ')))' },
        { h: 'Прогрес', f: (i, r) => '=IF(G' + r + '>0,SPARKLINE(MIN(F' + r + '/G' + r + ',1),@SP_Main),"")' },
      ],
    },
    {
      key: 'dist', title: 'РОЗПОДІЛ ДОХОДУ', n: 5, total: false,
      hint: 'Куди йде дохід цього циклу: уже сплачене + те, що ще заплановано.',
      rowsFixed: true,
      cols: [{ h: 'Куди' }, { h: 'Сума', fmt: FMT.money0 }, { h: '% доходу', fmt: FMT.pct }],
      rows: cr => [
        ['Регулярні платежі', '=SUM(' + cr('reg', 'E') + ')+SUM(' + cr('reg', 'H') + ')'],
        ['Змінні та інше', '=MAX(0,' + sumCycle_(exp) + '-SUM(' + cr('reg', 'E') + '))'],
        ['Борги', '=' + sumCycle_(debt) + '+SUM(' + cr('debt', 'E') + ')'],
        ['Заощадження', '=' + sumCycle_(sav) + '+SUM(' + cr('sav', 'E') + ')'],
        ['Вільний залишок', '=MAX(0,@D_Free)'],
      ],
    },
    {
      key: 'rule', title: 'ПРАВИЛО 50/30/20', n: 3, total: false,
      hint: 'Потреби (разом з боргами) / бажання / заощадження — у % від доходу циклу.',
      rowsFixed: true,
      cols: [{ h: 'Група' }, { h: 'Сума', fmt: FMT.money0 }, { h: '% доходу', fmt: FMT.pct },
             { h: 'Ціль', fmt: FMT.pct }, { h: 'Різниця', fmt: FMT.signPct }],
      rows: cr => [
        ['Потреби', '=SUMPRODUCT((@Reg_Group=' + need + ')*(' + cr('reg', 'E') + '+' + cr('reg', 'H') + '))+' +
          'SUMPRODUCT((@Var_Group=' + need + ')*' + cr('var', 'D') + ')+' + sumCycle_(debt) + '+SUM(' + cr('debt', 'E') + ')', '=@Needs'],
        ['Бажання', '=SUMPRODUCT((@Reg_Group<>' + need + ')*(' + cr('reg', 'E') + '+' + cr('reg', 'H') + '))+' +
          'SUMPRODUCT((@Var_Group<>' + need + ')*' + cr('var', 'D') + ')+' +
          'MAX(0,' + sumCycle_(exp) + '-SUM(' + cr('reg', 'E') + ')-SUM(' + cr('var', 'D') + '))', '=@Wants'],
        ['Заощадження', '=' + sumCycle_(sav) + '+SUM(' + cr('sav', 'E') + ')', '=@Save'],
      ],
    },
    {
      key: 'acc', title: 'РАХУНКИ', n: 6, total: true,
      hint: 'Баланс = початковий залишок + доходи − усі списання. Без рахунку — на основний.',
      cols: [
        { h: 'Рахунок', f: i => '=INDEX(@Acc_Name,' + i + ')', tot: 'label' },
        { h: 'Баланс', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Acc_Balance,' + i + ')))' },
      ],
    },
    {
      key: 'tips', title: '3 ПІДКАЗКИ', n: 3, total: false,
      hint: 'Що зробити просто зараз.',
      rowsFixed: true, merged: true,
      cols: [{ h: 'Порада' }],
      rows: cr => [
        ['=ARRAYFORMULA(LET(n,' + cr('var', 'B') + ',b,' + cr('var', 'C') + ',f,' + cr('var', 'D') + ',o,f-b,mx,MAX(o),' +
          'p,IFERROR(f/b,0),mp,MAX(p),' +
          'IF(mx>0,"⚠ «"&INDEX(n,MATCH(mx,o,0))&"» — понад бюджет на "&FIXED(mx,0)&" грн. Пригальмуй тут до кінця циклу.",' +
          'IF(mp>@C_Warn,"🟡 «"&INDEX(n,MATCH(mp,p,0))&"» — уже "&TEXT(mp,"0%")&" бюджету.",' +
          '"✓ Усі змінні витрати в межах бюджету."))))'],
        ['=ARRAYFORMULA(LET(nm,@Reg_Name,k,IFERROR(VLOOKUP(@Reg_Freq,@L_Freq,2,FALSE),1),' +
          'yr,IFERROR(IF((@Reg_Sub=TRUE)*(@Reg_Use<>TRUE)*(nm<>""),@Reg_Mine*12/k,0),0),mx,MAX(yr),' +
          'IF(mx>0,"💸 Підписка «"&INDEX(nm,MATCH(mx,yr,0))&"» коштує "&FIXED(mx,0)&" грн на рік, а ти нею не користуєшся. Скасуй — і ці гроші твої.",' +
          '"✓ Невикористаних підписок немає.")))'],
        ['=ARRAYFORMULA(LET(nm,@Debt_Name,l,IFERROR(@Debt_Left*1,0),a,(nm<>"")*(l>0)*(@Debt_Closed<>TRUE),' +
          'c,IF(a,l,MAX(l)+1),mn,MIN(c),' +
          'IF(SUM(a)=0,"✓ Активних боргів немає. Так тримати!",' +
          '"❄ Сніжна куля: першим закривай «"&INDEX(nm,MATCH(mn,c,0))&"» — залишок "&FIXED(mn,0)&" грн.")))'],
      ],
    },
  ];
}

function writeDashBlock_(sh, b, p, colRange) {
  const ncol = b.merged ? DASH.cols : b.cols.length;
  writeBlockTitle_(sh, p.title, b.title, DASH.cols);
  sh.getRange(p.title, 2 + Math.min(DASH.cols - 1, 3)).setValue(b.hint).setFontColor(THEME.muted).setFontStyle('italic');
  const head = sh.getRange(p.head, 2, 1, ncol);
  head.setBackground(THEME.headerBg).setFontColor(THEME.headerTxt).setFontWeight('bold').setWrap(true);
  if (b.merged) head.merge().setValue(b.cols[0].h);
  else head.setValues([b.cols.map(c => c.h)]);

  if (b.rowsFixed) {
    const rows = b.rows(colRange);
    rows.forEach((rv, k) => {
      const r = p.top + k;
      if (b.merged) {
        sh.getRange(r, 2, 1, DASH.cols).merge().setFormula(X_(rv[0])).setWrap(true);
        sh.setRowHeight(r, 30);
        return;
      }
      sh.getRange(r, 2).setValue(rv[0]);
      sh.getRange(r, 3).setFormula(X_(rv[1])).setNumberFormat(b.cols[1].fmt);
      sh.getRange(r, 4).setFormula(X_('=IF(@D_Income=0,"",C' + r + '/@D_Income)')).setNumberFormat(FMT.pct);
      if (rv[2]) {
        sh.getRange(r, 5).setFormula(X_(rv[2])).setNumberFormat(FMT.pct);
        sh.getRange(r, 6).setFormula(X_('=IF(D' + r + '="","",D' + r + '-E' + r + ')')).setNumberFormat(FMT.signPct);
      }
    });
  } else {
    b.cols.forEach((c, j) => {
      const formulas = [];
      for (let i = 1; i <= b.n; i++) formulas.push([X_(c.f(i, p.top + i - 1))]);
      const rng = sh.getRange(p.top, 2 + j, b.n, 1).setFormulas(formulas);
      if (c.fmt) rng.setNumberFormat(c.fmt);
    });
  }

  if (b.total) {
    const r = p.total;
    b.cols.forEach((c, j) => {
      const L = colLetter_(2 + j);
      const cell = sh.getRange(r, 2 + j);
      if (c.tot === 'label') cell.setValue('Разом');
      else if (c.tot === 'sum') cell.setFormula(X_('=SUM(' + L + p.top + ':' + L + p.bottom + ')'));
      else if (c.tot && c.tot.indexOf('ratio:') === 0) {
        const a = c.tot.split(':');
        cell.setFormula(X_('=IF(' + a[2] + r + '=0,"",' + a[1] + r + '/' + a[2] + r + ')'));
      } else if (c.tot && c.tot.indexOf('diff:') === 0) {
        const a = c.tot.split(':');
        cell.setFormula(X_('=' + a[1] + r + '-' + a[2] + r));
      }
      if (c.fmt) cell.setNumberFormat(c.fmt === FMT.money0 ? FMT.money : c.fmt);
    });
    sh.getRange(r, 2, 1, ncol).setFontWeight('bold').setBorder(true, null, null, null, null, null,
      THEME.primary, SpreadsheetApp.BorderStyle.SOLID);
  }
  sh.getRange(p.top, 2, b.n, ncol).setBorder(null, null, true, null, null, true,
    THEME.headerBg, SpreadsheetApp.BorderStyle.SOLID);
}

// ─────────────────────────────────────────────────────────────────────────────
// ДЕМО-ДАНІ
// Сценарій з поста для поточного місяця + 3 попередні місяці. Цикл — від 1-го числа.
// Перед заповненням «Налаштування» зберігаються; «Прибрати демо-дані» їх повертає.
// ─────────────────────────────────────────────────────────────────────────────
const DEMO_TAG = 'демо';
const DEMO_EXPECTED = 9400;
const PROP_DEMO = 'BT_DEMO';
const PROP_SNAP = 'BT_SNAP_';

function fillDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.dash)) {
    ss.toast('Спочатку побудуй трекер: меню «💜 Трекер» → «Побудувати трекер з нуля».', '💜 Трекер', 8);
    return;
  }
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(PROP_DEMO) !== '1') snapshotSettings_(ss);
  removeDemoOps_(ss);

  const t = todayParts_(ss);
  const months = REF.months;
  const mo = k => months[((t.m + k) % 12 + 12) % 12];
  const start = isoDate_(t.y, t.m - 3, 1);

  // ── Налаштування
  setKV_(ss, { Year: t.y, Start: start, Calc: '', Cycle: REF.cycle[0], Warn: 0.8, Over: 1, Needs: 0.5, Wants: 0.3, Save: 0.2 });
  setBlockRows_(ss, 'Inc', [
    { Name: 'Зарплата', Main: true, Exp: 30000, Day: 1 },
    { Name: 'Фріланс', Main: false },
  ]);
  setBlockRows_(ss, 'Reg', [
    { Name: 'Оренда', Total: 30000, Share: 0.5, Freq: 'Щомісяця', Day: 1, Group: 'Потреби' },
    { Name: 'Комуналка + інтернет', Total: 1950, Share: 1, Freq: 'Щомісяця', Day: 3, Group: 'Потреби' },
    { Name: 'Фоторедактор', Total: 450, Share: 1, Freq: 'Раз на квартал', Month: mo(1), Day: 12, Group: 'Бажання', Sub: true, Use: false },
    { Name: 'Хмарне сховище', Total: 1200, Share: 1, Freq: 'Раз на рік', Month: mo(2), Day: 20, Group: 'Бажання', Sub: true, Use: false },
  ]);
  setBlockRows_(ss, 'Var', [
    { Name: 'Продукти і кава', Budget: 7000, Group: 'Потреби' },
    { Name: 'Транспорт', Budget: 1200, Group: 'Потреби' },
    { Name: 'Здоров\'я і краса', Budget: 1500, Group: 'Потреби' },
    { Name: 'Кафе і доставка', Budget: 1500, Group: 'Бажання' },
    { Name: 'Одяг', Budget: 1500, Group: 'Бажання' },
    { Name: 'Розваги', Budget: 1000, Group: 'Бажання' },
    { Name: 'Подарунки', Budget: 800, Group: 'Бажання' },
    { Name: 'Інше', Budget: 500, Group: 'Бажання' },
  ]);
  setBlockRows_(ss, 'Debt', [
    { Name: 'Розстрочка за телефон', Pay: 1590, Freq: 'Щомісяця', Start: start, Count: 6, Closed: true },
  ]);
  setBlockRows_(ss, 'Sav', [
    { Name: 'Подушка безпеки', Target: 50000, Monthly: 3000, Date: isoDate_(t.y, t.m + 12, 1) },
    { Name: 'Відпустка', Target: 20000, Monthly: 2000, Date: isoDate_(t.y, t.m + 7, 1) },
  ]);
  setBlockRows_(ss, 'Acc', [
    { Name: 'Картка', Initial: 0, InitDate: start, Main: true },
    { Name: 'Готівка', Initial: 2000, InitDate: start, Main: false },
  ]);

  // ── Операції
  const ops = demoOperations_(t);
  const sh = ss.getSheetByName(SHEETS.ops);
  const first = firstFreeOpsRow_(sh);
  sh.getRange(first, 1, ops.length, OPS_HEADERS.length).setValues(ops);
  applyCategoryValidation_(sh, first, ops.length, false);

  ss.getRangeByName(NR + 'D_Period').setValue(REF.current);
  props.setProperty(PROP_DEMO, '1');
  SpreadsheetApp.flush();

  const free = ss.getRangeByName(NR + 'D_Free').getValue();
  const foreign = countForeignOpsInCycle_(ss);
  ss.toast('Демо-дані додано. «Можу витратити» = ' + fmtMoney_(free) + ' (очікується ' + fmtMoney_(DEMO_EXPECTED) + ').' +
    (foreign ? ' Увага: у поточному циклі є ще ' + foreign + ' твоїх операцій — вони змінюють цифру.' : ''),
    '💜 Трекер', 10);
  ss.setActiveSheet(ss.getSheetByName(SHEETS.dash));
}

/** Нова побудова — старі демо-позначки й знімок налаштувань більше не потрібні. */
function resetDemoState_() {
  const props = PropertiesService.getDocumentProperties();
  const n = Number(props.getProperty(PROP_SNAP + 'N') || 0);
  for (let i = 0; i < n; i++) props.deleteProperty(PROP_SNAP + i);
  props.deleteProperty(PROP_SNAP + 'N');
  props.deleteProperty(PROP_DEMO);
}

function clearDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const removed = removeDemoOps_(ss);
  const restored = restoreSettings_(ss);
  PropertiesService.getDocumentProperties().deleteProperty(PROP_DEMO);
  SpreadsheetApp.flush();
  ss.toast('Видалено демо-операцій: ' + removed + '. ' +
    (restored ? 'Налаштування повернуто як були до демо.' : 'Збережених налаштувань не було — їх не змінено.'), '💜 Трекер', 8);
}

/** Рядки операцій демо: [дата, тип, сума, категорія, рахунок, коментар, джерело]. */
function demoOperations_(t) {
  const I = TYPES.income, E = TYPES.expense, D = TYPES.debt, S = TYPES.saving;
  const rows = [];
  const add = (m, day, type, amount, cat, acc, comment) =>
    rows.push([isoDate_(t.y, t.m + m, day), type, amount, cat, acc || '', comment || '', DEMO_TAG]);

  // Поточний місяць — точно як у пості. Дні не пізніше сьогодні.
  const d = x => Math.min(x, t.d);
  add(0, d(1), I, 30200, 'Зарплата', '', 'Зарплата (офіс)');
  add(0, d(4), I, 12700, 'Фріланс', '', 'Фріланс-проєкт');
  add(0, d(1), D, 4770, 'Розстрочка за телефон', '', 'Закрила повністю, щоб не тягнулася');
  add(0, d(1), E, 15000, 'Оренда', '', 'Половина оренди — ділимо з нареченим');
  add(0, d(3), E, 1950, 'Комуналка + інтернет', '', '');
  add(0, d(1), S, 3000, 'Подушка безпеки', '', '');
  add(0, d(1), S, 2000, 'Відпустка', '', '');
  const food = [412, 89, 645, 120, 538, 95, 760, 148, 384, 75, 920, 130, 466, 88, 702, 118, 1090];
  const cash = { 1: true, 9: true, 13: true }; // кава готівкою
  food.forEach((a, i) => {
    const day = 1 + Math.floor(i * (t.d - 1) / (food.length - 1));
    add(0, day, E, a, 'Продукти і кава', cash[i] ? 'Готівка' : '', cash[i] ? 'кава' : '');
  });

  // 3 попередні місяці: дохід 38–46 тис., плаваючий фріланс.
  const hist = [
    { m: -3, inc: [29800, 8700], util: 1850, save: [3000, 2000], v: [6200, 700, 450, 900, 0, 600, 0, 200] },
    { m: -2, inc: [30200, 15600], util: 2100, save: [3000, 1000], v: [5900, 650, 500, 1100, 1600, 400, 500, 150] },
    { m: -1, inc: [30000, 11900], util: 2300, save: [3000, 2000], v: [6400, 750, 300, 800, 0, 500, 0, 150] },
  ];
  const varNames = ['Продукти і кава', 'Транспорт', 'Здоров\'я і краса', 'Кафе і доставка', 'Одяг', 'Розваги', 'Подарунки', 'Інше'];
  const parts = [8, 3, 2, 3, 1, 2, 1, 1];
  hist.forEach(h => {
    add(h.m, 1, I, h.inc[0], 'Зарплата', '', 'Зарплата (офіс)');
    add(h.m, 18, I, h.inc[1], 'Фріланс', '', 'Фріланс');
    add(h.m, 1, E, 15000, 'Оренда', '', '');
    add(h.m, 3, E, h.util, 'Комуналка + інтернет', '', '');
    add(h.m, 1, D, 1590, 'Розстрочка за телефон', '', 'Плановий платіж');
    add(h.m, 2, S, h.save[0], 'Подушка безпеки', '', '');
    add(h.m, 2, S, h.save[1], 'Відпустка', '', '');
    if (h.m === -2) add(h.m, 12, E, 450, 'Фоторедактор', '', 'Підписка');
    h.v.forEach((total, k) => {
      if (!total) return;
      splitAmount_(total, parts[k], k + h.m * 7).forEach((a, j) => {
        const day = 2 + Math.floor((j + 0.5) * 26 / parts[k]);
        const isCash = varNames[k] === 'Кафе і доставка' && j === 0;
        add(h.m, day, E, a, varNames[k], isCash ? 'Готівка' : '', '');
      });
    });
  });
  return rows;
}

/** Ділить суму на n «правдоподібних» частин (детерміновано), остання бере залишок. */
function splitAmount_(total, n, seed) {
  if (n <= 1) return [total];
  const w = [];
  let x = Math.abs(seed * 9301 + 49297) % 233280;
  for (let i = 0; i < n; i++) { x = (x * 9301 + 49297) % 233280; w.push(0.6 + x / 233280); }
  const sumW = w.reduce((a, b) => a + b, 0);
  const out = w.map(v => Math.round(total * v / sumW));
  out[n - 1] += total - out.reduce((a, b) => a + b, 0);
  return out;
}

function inputKeys_() {
  const keys = GENERAL_ROWS.concat(RULE_ROWS).map(r => r.key);
  SETTINGS_BLOCKS.forEach(b => b.cols.forEach(c => { if (c.t !== 'f') keys.push(b.key + '_' + c.key); }));
  return keys;
}

function snapshotSettings_(ss) {
  const props = PropertiesService.getDocumentProperties();
  const data = {};
  inputKeys_().forEach(k => {
    const rng = ss.getRangeByName(NR + k);
    if (rng) data[k] = rng.getValues().map(r => r[0]);
  });
  // Розбиваємо на частини: властивість документа вміщує до ~9 КБ.
  const json = JSON.stringify(data);
  const size = 8000;
  const n = Math.ceil(json.length / size);
  props.setProperty(PROP_SNAP + 'N', String(n));
  for (let i = 0; i < n; i++) props.setProperty(PROP_SNAP + i, json.substr(i * size, size));
}

function restoreSettings_(ss) {
  const props = PropertiesService.getDocumentProperties();
  const n = Number(props.getProperty(PROP_SNAP + 'N') || 0);
  if (!n) return false;
  let json = '';
  for (let i = 0; i < n; i++) json += props.getProperty(PROP_SNAP + i) || '';
  const data = JSON.parse(json);
  Object.keys(data).forEach(k => {
    const rng = ss.getRangeByName(NR + k);
    if (!rng) return;
    const vals = data[k].map(v => [typeof v === 'string' && /^\d{4}-\d\d-\d\dT/.test(v) ? new Date(v) : v]);
    if (vals.length === rng.getNumRows()) rng.setValues(vals);
  });
  for (let i = 0; i < n; i++) props.deleteProperty(PROP_SNAP + i);
  props.deleteProperty(PROP_SNAP + 'N');
  return true;
}

function setKV_(ss, obj) {
  Object.keys(obj).forEach(k => ss.getRangeByName(NR + k).setValue(obj[k]));
}

/** Перезаписує рядки блоку «Налаштувань» (лише поля вводу, формули не чіпає). */
function setBlockRows_(ss, key, rows) {
  const block = SETTINGS_BLOCKS.filter(b => b.key === key)[0];
  block.cols.forEach(c => {
    if (c.t === 'f') return;
    const rng = ss.getRangeByName(NR + key + '_' + c.key);
    const vals = [];
    for (let i = 0; i < block.rows; i++) {
      const v = rows[i] && rows[i][c.key] !== undefined ? rows[i][c.key] : (c.t === 'check' ? false : '');
      vals.push([v]);
    }
    rng.setValues(vals);
  });
}

/** Видаляє демо-операції, зсуваючи решту вгору. Повертає кількість видалених. */
function removeDemoOps_(ss) {
  const sh = ss.getSheetByName(SHEETS.ops);
  if (!sh) return 0;
  const n = DATA_ROWS - 1;
  const rng = sh.getRange(2, 1, n, OPS_HEADERS.length);
  const vals = rng.getValues();
  const keep = vals.filter(r => r[OPS.source - 1] !== DEMO_TAG && r.slice(0, OPS.comment).some(v => v !== ''));
  const removed = vals.filter(r => r[OPS.source - 1] === DEMO_TAG).length;
  if (!removed) return 0;
  const blank = new Array(OPS_HEADERS.length).fill('');
  const out = keep.concat(Array.from({ length: n - keep.length }, () => blank.slice()));
  rng.setValues(out);
  applyCategoryValidation_(sh, 2, n, false);
  return removed;
}

function firstFreeOpsRow_(sh) {
  const vals = sh.getRange(2, 1, DATA_ROWS - 1, OPS.comment).getValues();
  let last = 0;
  vals.forEach((r, i) => { if (r.some(v => v !== '')) last = i + 1; });
  return last + 2;
}

/** Скільки не-демо операцій потрапляє в поточний цикл. */
function countForeignOpsInCycle_(ss) {
  const st = ss.getRangeByName(NR + 'C_Start').getValue();
  const nx = ss.getRangeByName(NR + 'C_Next').getValue();
  return opsRows_(ss).filter(r => r.source !== DEMO_TAG && r.date >= st && r.date < nx).length;
}

/** Операції як об'єкти (лише заповнені рядки). */
function opsRows_(ss) {
  const sh = ss.getSheetByName(SHEETS.ops);
  return sh.getRange(2, 1, DATA_ROWS - 1, OPS_HEADERS.length).getValues()
    .filter(r => r[OPS.amount - 1] !== '')
    .map(r => ({
      date: r[OPS.date - 1], type: r[OPS.type - 1], amount: Number(r[OPS.amount - 1]) || 0,
      category: r[OPS.category - 1], account: r[OPS.account - 1], source: r[OPS.source - 1],
    }));
}

function fmtMoney_(v) {
  if (typeof v !== 'number') return String(v);
  return Math.round(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' грн';
}

// ─────────────────────────────────────────────────────────────────────────────
// ПЕРЕВІРКА
// ─────────────────────────────────────────────────────────────────────────────
function selfCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  const res = [];
  const ok = (pass, text) => res.push((pass ? '✓ ' : '✗ ') + text);

  // 1. Помилки у формулах.
  const errors = scanErrors_(ss, trackerSheetNames_());
  ok(!errors.length, errors.length
    ? 'Помилки у формулах: ' + errors.length + ' (перші: ' + errors.slice(0, 5).join('; ') + ')'
    : 'Жодної клітинки з #REF!, #N/A, #ERROR!, #VALUE!, #DIV/0!, #NAME?');

  // 2. Велика цифра на демо-даних.
  const props = PropertiesService.getDocumentProperties();
  const periodCell = ss.getRangeByName(NR + 'D_Period');
  if (props.getProperty(PROP_DEMO) === '1' && periodCell) {
    const was = periodCell.getValue();
    periodCell.setValue(REF.current);
    SpreadsheetApp.flush();
    const free = ss.getRangeByName(NR + 'D_Free').getValue();
    const foreign = countForeignOpsInCycle_(ss);
    const parts = dashCalcLines_(() => '').map(c => c.key + '=' + ss.getRangeByName(NR + c.key).getValue()).join(', ');
    ok(free === DEMO_EXPECTED, 'Демо: «Можу витратити» = ' + fmtMoney_(free) + ' (очікується ' + fmtMoney_(DEMO_EXPECTED) + ')' +
      (free === DEMO_EXPECTED ? '' : '. Складові: ' + parts) +
      (foreign ? '. У циклі є ' + foreign + ' не-демо операцій' : ''));
    periodCell.setValue(was);
    SpreadsheetApp.flush();
  } else {
    res.push('• Демо: пропущено — демо-дані не заповнені (меню «Заповнити демо-даними»).');
  }

  // 3. Сума по категоріях = сума в «Операціях» для кожного місяця.
  const ops = opsRows_(ss);
  const lists = {};
  Object.keys(TYPE_LIST).forEach(type => {
    lists[type] = listRange_(ss, TYPE_LIST[type]).getValues().map(v => v[0]).filter(String);
  });
  const byMonth = {};
  const unknown = [];
  ops.forEach(o => {
    if (!(o.date instanceof Date)) { unknown.push('дата «' + o.date + '»'); return; }
    const key = o.date.getFullYear() + '-' + ('0' + (o.date.getMonth() + 1)).slice(-2);
    byMonth[key] = byMonth[key] || { all: 0, cat: 0 };
    byMonth[key].all += o.amount;
    if (lists[o.type] && lists[o.type].indexOf(o.category) !== -1) byMonth[key].cat += o.amount;
    else unknown.push(key + ': ' + (o.type || 'без типу') + ' / ' + (o.category || 'без категорії') + ' ' + fmtMoney_(o.amount));
  });
  const bad = Object.keys(byMonth).filter(k => Math.abs(byMonth[k].all - byMonth[k].cat) > 0.005);
  ok(!bad.length && !unknown.length, bad.length || unknown.length
    ? 'Не всі операції мають відому категорію: ' + unknown.slice(0, 5).join('; ')
    : 'Сума по категоріях = сума в «Операціях» для кожного місяця (' + Object.keys(byMonth).length + ' міс.)');

  // 4. Дашборд збігається з «Операціями» за вибраний цикл.
  const st = ss.getRangeByName(NR + 'C_Start').getValue();
  const nx = ss.getRangeByName(NR + 'C_Next').getValue();
  const inCycle = ops.filter(o => o.date instanceof Date && o.date >= st && o.date < nx);
  const jsIncome = inCycle.filter(o => o.type === TYPES.income).reduce((a, o) => a + o.amount, 0);
  const jsSpent = inCycle.filter(o => o.type && o.type !== TYPES.income).reduce((a, o) => a + o.amount, 0);
  const dInc = ss.getRangeByName(NR + 'D_Income').getValue();
  const dSp = ss.getRangeByName(NR + 'D_Spent').getValue();
  ok(Math.abs(jsIncome - dInc) < 0.005 && Math.abs(jsSpent - dSp) < 0.005,
    'Дашборд: дохід ' + fmtMoney_(dInc) + ' і витрати ' + fmtMoney_(dSp) + ' циклу збігаються з «Операціями»' +
    (Math.abs(jsIncome - dInc) < 0.005 && Math.abs(jsSpent - dSp) < 0.005 ? '' : ' — НІ: в «Операціях» ' + fmtMoney_(jsIncome) + ' / ' + fmtMoney_(jsSpent)));

  // 5. Баланси рахунків.
  const accN = ss.getRangeByName(NR + 'Acc_Name').getValues().map(v => v[0]);
  const accI = ss.getRangeByName(NR + 'Acc_Initial').getValues().map(v => Number(v[0]) || 0);
  const accD = ss.getRangeByName(NR + 'Acc_InitDate').getValues().map(v => v[0]);
  const accM = ss.getRangeByName(NR + 'Acc_Main').getValues().map(v => v[0] === true);
  const accB = ss.getRangeByName(NR + 'Acc_Balance').getValues().map(v => v[0]);
  const accBad = [];
  accN.forEach((name, i) => {
    if (!name) return;
    let bal = accI[i];
    ops.forEach(o => {
      const mine = o.account === name || (o.account === '' && accM[i]);
      if (!mine || !(o.date instanceof Date)) return;
      if (accD[i] instanceof Date && o.date < accD[i]) return;
      bal += o.type === TYPES.income ? o.amount : -o.amount;
    });
    if (Math.abs(bal - (Number(accB[i]) || 0)) > 0.005) accBad.push(name + ': таблиця ' + fmtMoney_(accB[i]) + ', має бути ' + fmtMoney_(bal));
  });
  ok(!accBad.length, accBad.length ? 'Баланси рахунків не сходяться: ' + accBad.join('; ')
    : 'Баланси рахунків = початкові залишки + доходи − списання');

  // 6. Знак $ у значеннях і форматах.
  const dollars = [];
  trackerSheetNames_().forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const rng = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    const dv = rng.getDisplayValues();
    const nf = rng.getNumberFormats();
    dv.forEach((row, r) => row.forEach((v, c) => {
      if (String(v).indexOf('$') !== -1 || String(nf[r][c]).indexOf('$') !== -1) dollars.push(name + '!' + colLetter_(c + 1) + (r + 1));
    }));
  });
  ok(!dollars.length, dollars.length ? 'Знайдено «$»: ' + dollars.slice(0, 10).join(', ') : 'Ніде немає символу $');

  // Запис результату.
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm');
  Logger.log('Перевірка ' + stamp + '\n' + res.join('\n'));
  const start = ss.getSheetByName(SHEETS.start);
  if (start) {
    const area = start.getRange(START.checkRow, 3, START.checkLines, 1);
    area.clearContent().setFontColor(THEME.text);
    const lines = ['Перевірено ' + stamp].concat(res).slice(0, START.checkLines);
    start.getRange(START.checkRow, 3, lines.length, 1).setValues(lines.map(l => [l])).setWrap(true);
    lines.forEach((l, i) => {
      const c = start.getRange(START.checkRow + i, 3);
      if (l.charAt(0) === '✓') c.setFontColor(THEME.okTxt);
      else if (l.charAt(0) === '✗') c.setFontColor(THEME.danger);
      else c.setFontColor(THEME.muted);
    });
  }
  const failed = res.filter(l => l.charAt(0) === '✗').length;
  ss.toast(failed ? '⚠ Не пройдено перевірок: ' + failed + '. Деталі — на вкладці «Старт».' : '✓ Усі перевірки пройдено.', '💜 Трекер', 8);
  return res;
}

// ─────────────────────────────────────────────────────────────────────────────
// ЕТАП 4 (поки заглушка)
// ─────────────────────────────────────────────────────────────────────────────
function createEntryForm() { notYet_('Форма для телефону з\'явиться на етапі 4.'); }

function notYet_(msg) {
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '💜 Трекер', 5);
}

// ─────────────────────────────────────────────────────────────────────────────
// РОЗДІЛЬНИК АРГУМЕНТІВ У ФОРМУЛАХ
// У коді всі формули пишуться з комами (англійський синтаксис). В українській та
// інших «європейських» локалях Google Sheets чекає «;», тому перед записом формула
// проходить через F_(). Роздільник визначається пробною формулою в самій таблиці.
// ─────────────────────────────────────────────────────────────────────────────
let ARG_SEP_ = null;

function detectArgSeparator_(probeSheet) {
  detectArgSeparatorAt_(probeSheet.getRange('A1'));
}

function argSep_() {
  if (ARG_SEP_) return ARG_SEP_;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const probe = ss.getSheetByName(SHEETS.lists) || ss.getSheets()[0];
  const cell = probe.getRange(probe.getMaxRows(), probe.getMaxColumns());
  const keep = cell.getFormula() || cell.getValue();
  detectArgSeparatorAt_(cell);
  if (typeof keep === 'string' && keep.charAt(0) === '=') cell.setFormula(keep); else cell.setValue(keep);
  return ARG_SEP_;
}

/** =SUM(1,2): у «комовій» локалі дасть 3; в українській — 1,2 або помилку. */
function detectArgSeparatorAt_(cell) {
  cell.setFormula('=SUM(1,2)');
  SpreadsheetApp.flush();
  ARG_SEP_ = cell.getValue() === 3 ? ',' : ';';
  cell.clearContent();
}

/** Формула з комами → формула з роздільником цієї таблиці. Коми в "рядках" не чіпає. */
function F_(formula) {
  if (argSep_() === ',') return formula;
  let out = '', inStr = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula.charAt(i);
    if (ch === '"') inStr = !inStr;
    out += (ch === ',' && !inStr) ? ';' : ch;
  }
  return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// ДОПОМІЖНІ
// ─────────────────────────────────────────────────────────────────────────────

/** Баланс рахунку: початковий + доходи − усі інші операції з дати залишку.
 *  Операції без рахунку зараховуються на основний. */
function accountBalanceFormula_() {
  const C = OPS_REF + '!$C:$C', B = OPS_REF + '!$B:$B', E = OPS_REF + '!$E:$E', A = OPS_REF + '!$A:$A';
  const since = '">="&N({InitDate})';
  const sum = (typeCrit, accCrit) => 'SUMIFS(' + C + ',' + B + ',' + typeCrit + ',' + E + ',' + accCrit + ',' + A + ',' + since + ')';
  const inc = '"' + TYPES.income + '"', out = '"<>' + TYPES.income + '"';
  return '=IF({Name}="","",N({Initial})+' + sum(inc, '{Name}') + '-' + sum(out, '{Name}') +
    '+IF({Main},' + sum(inc, '""') + '-' + sum(out, '""') + ',0))';
}

function trackerSheetNames_() {
  return Object.keys(SHEETS).map(k => SHEETS[k]);
}

/** Дата як рядок «рррр-мм-дд»: таблиця сама розбирає її як дату без часу (без зсуву поясів).
 *  Місяць 0-based і може виходити за межі року (−3, 14 тощо). */
function isoDate_(y, m, d) {
  const dt = new Date(Date.UTC(y, m, d));
  return dt.getUTCFullYear() + '-' + ('0' + (dt.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + dt.getUTCDate()).slice(-2);
}

/** Сьогодні в часовому поясі таблиці: { y, m (0-based), d, iso }. */
function todayParts_(ss) {
  const s = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd').split('-').map(Number);
  return { y: s[0], m: s[1] - 1, d: s[2], iso: isoDate_(s[0], s[1] - 1, s[2]) };
}

function todayIso_(ss) {
  return todayParts_(ss).iso;
}

/** «@Name» у формулі → іменований діапазон трекера, далі — роздільники під локаль. */
function X_(formula) {
  return F_(formula.replace(/@(\w+)/g, (m, k) => NR + k));
}

function blockWidth_(key) {
  return SETTINGS_BLOCKS.filter(b => b.key === key)[0].cols.length;
}

function writeBlockTitle_(sh, row, title, width) {
  sh.getRange(row, 2).setValue(spaced_(title)).setFontSize(12).setFontWeight('bold').setFontColor(THEME.primary);
  sh.getRange(row, 2, 1, Math.max(1, width)).setBorder(null, null, true, null, null, null,
    THEME.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(row, 30);
}

/** «ДОХІД» → «Д О Х І Д», слова розділені трьома пробілами. */
function spaced_(s) {
  return String(s).toUpperCase().split(' ').map(w => w.split('').join(' ')).join('   ');
}

function styleSheet_(sh) {
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())
    .setFontFamily(THEME.font).setFontColor(THEME.text).setFontSize(10).setBackground(THEME.bg);
  sh.setTabColor(THEME.primary);
}

function protect_(rng, what) {
  rng.protect().setDescription('Формула: ' + what).setWarningOnly(true);
}

function trimSheet_(sh, rows, cols) {
  if (sh.getMaxRows() > rows) sh.deleteRows(rows + 1, sh.getMaxRows() - rows);
  else if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function absA1_(rng) {
  return '$' + colLetter_(rng.getColumn()) + '$' + rng.getRow();
}

function within_(cell, rng) {
  return cell.getRow() >= rng.getRow() && cell.getRow() <= rng.getLastRow() &&
         cell.getColumn() >= rng.getColumn() && cell.getColumn() <= rng.getLastColumn();
}
