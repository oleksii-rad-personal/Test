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
  primary:   '#8E6BBF', // основний (заголовки блоків, велика цифра)
  accent:    '#E58FB3', // акцент (кнопки, лінії, другий ряд графіків)
  headerBg:  '#F3E8F7', // фон заголовків таблиць
  headerTxt: '#5B3F86', // текст заголовків таблиць
  input:     '#FFF8D6', // поля вводу (світло-жовті)
  danger:    '#E57373', // червоний — перевищення
  warn:      '#FFD54F', // жовтий — попередження (>80%)
  ok:        '#81C784', // зелений — норма
  text:      '#3D3550', // основний текст
  muted:     '#9A8FB0', // підказки
  bg:        '#FFFFFF', // фон аркушів
  font:      'Montserrat',
};

// ─────────────────────────────────────────────────────────────────────────────
// НАЗВИ ТА ФОРМАТИ
// ─────────────────────────────────────────────────────────────────────────────
const SHEETS = {
  start:    'Старт',
  settings: 'Налаштування',
  ops:      'Операції',
  lists:    '_Списки',
};

const FMT = {
  money: '#,##0 "грн"',
  pct:   '0%',
  date:  'dd.mm.yyyy',
  int:   '0',
};

const NR = 'BT_';             // префікс іменованих діапазонів трекера
const DATA_ROWS = 1000;       // скільки рядків готуємо в «Операціях»

const TYPES = { income: 'Дохід', expense: 'Витрата', debt: 'Борг', saving: 'Заощадження' };

/** Довідкові значення (не розрахункові дані): пишуться у прихований аркуш «_Списки». */
const REF = {
  types:  [TYPES.income, TYPES.expense, TYPES.debt, TYPES.saving],
  freq:   [['Щомісяця', 1], ['Раз на квартал', 3], ['Раз на рік', 12]],
  groups: ['Потреби', 'Бажання'],
  cycle:  ['Від основної виплати', 'Календарний місяць'],
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
      { key: 'Month', h: 'Місяць першої оплати', t: 'list:L_Month', w: 150 },
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
  ss.setSpreadsheetLocale('uk_UA');
  ss.setSpreadsheetTimeZone('Europe/Kyiv');

  deleteTrackerSheets_(ss);

  // Спочатку створюємо всі аркуші, бо вони посилаються один на одного.
  const order = [SHEETS.start, SHEETS.settings, SHEETS.ops, SHEETS.lists];
  const sh = {};
  order.forEach((name, i) => { sh[name] = ss.insertSheet(name, i); });
  removeTempSheet_(ss);

  buildListsStatic_(sh[SHEETS.lists]);
  buildSettings_(ss, sh[SHEETS.settings]);
  buildListsDynamic_(ss, sh[SHEETS.lists]);
  buildOps_(ss, sh[SHEETS.ops]);
  buildStart_(sh[SHEETS.start]);

  sh[SHEETS.lists].hideSheet();
  ss.setActiveSheet(sh[SHEETS.start]);
  SpreadsheetApp.flush();
  ss.toast('Трекер побудовано. Почни з вкладки «Налаштування».', '💜 Трекер', 5);
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
  L_Freq: 'G', L_Group: 'I', L_Cycle: 'J', L_All: 'K', L_Month: 'L',
};
const LISTS_MAX = 120;

function buildListsStatic_(sh) {
  const ss = sh.getParent();
  sh.getRange('A1:L1').setValues([['Типи', 'Дохід', 'Витрата', 'Борг', 'Заощадження', 'Рахунки',
    'Частота', 'Місяців', 'Групи', 'Цикл', 'Усі категорії', 'Місяці']]).setFontWeight('bold');

  const put = (col, values) => sh.getRange(col + '2:' + col + (values.length + 1)).setValues(values.map(v => [v]));
  put('A', REF.types);
  sh.getRange(2, 7, REF.freq.length, 2).setValues(REF.freq);
  put('I', REF.groups);
  put('J', REF.cycle);
  put('L', REF.months);

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
  Object.keys(f).forEach(c => sh.getRange(c + '2').setFormula(f[c]));
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
  const today = new Date();
  ss.getRangeByName(NR + 'Year').setValue(today.getFullYear());
  ss.getRangeByName(NR + 'Start').setValue(new Date(today.getFullYear(), today.getMonth(), 1));
  ss.getRangeByName(NR + 'Cycle').setValue(REF.cycle[0]);
  ss.getRangeByName(NR + 'Warn').setValue(0.8);
  ss.getRangeByName(NR + 'Over').setValue(1);
  ss.getRangeByName(NR + 'Needs').setValue(0.5);
  ss.getRangeByName(NR + 'Wants').setValue(0.3);
  ss.getRangeByName(NR + 'Save').setValue(0.2);

  // Дата розрахунку, що діє насправді (порожня → сьогодні).
  const calc = ss.getRangeByName(NR + 'Calc');
  const calcNow = calc.offset(0, 1);
  calcNow.setFormula('=IF(' + calc.getA1Notation() + '="",TODAY(),' + calc.getA1Notation() + ')')
    .setNumberFormat(FMT.date).setFontColor(THEME.primary).setFontWeight('bold');
  ss.setNamedRange(NR + 'CalcDate', calcNow);
  calc.offset(0, 2).setValue('← діє зараз. ' + GENERAL_ROWS[2].hint);
  protect_(calcNow, 'Дата розрахунку');

  // Сума правила 50/30/20 має бути 100%.
  const needs = ss.getRangeByName(NR + 'Needs');
  needs.offset(0, 1).setFormula('=IF(ROUND(SUM(' + needs.offset(0, 0, 3, 1).getA1Notation() + '),4)=1,"✓ разом 100%","⚠ разом має бути 100%")')
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
    .requireFormulaSatisfied('=AND(ISNUMBER(' + first + '),' + first + '>=' + startA1 + ')')
    .setAllowInvalid(false)
    .setHelpText('Дата початку не може бути раніше дати старту трекера (Налаштування → Загальне → Дата старту).')
    .build());

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
        formulas.push([c.f.replace(/\{(\w+)\}/g, (m, k) => {
          if (!letter[k]) throw new Error('Невідома колонка ' + k + ' у блоці ' + block.key);
          return letter[k] + rr;
        })]);
      }
      rng.setFormulas(formulas).setNumberFormat(FMT[c.fmt] || FMT.money).setFontColor(THEME.text)
        .setBackground('#F7F5FA');
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
  hintCell.setFormula('=IF(COUNTIF(' + rng.getA1Notation() + ',TRUE)=1,"","⚠ Познач ' + what + '")')
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

  sh.getRange(1, 1, DATA_ROWS, OPS_HEADERS.length).createFilter();
  sh.hideColumns(OPS.source);
  protect_(sh.getRange(1, 1, 1, OPS_HEADERS.length), 'Шапка «Операцій»');
  trimSheet_(sh, DATA_ROWS, OPS_HEADERS.length);
}

// ─── «Старт» ─────────────────────────────────────────────────────────────────
const START = { formRow: 14, checkRow: 18 };

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
    const today = new Date(); today.setHours(0, 0, 0, 0);
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
// ЕТАПИ 2–4 (поки заглушки)
// ─────────────────────────────────────────────────────────────────────────────
function fillDemoData()    { notYet_('Демо-дані з\'являться на етапі 2.'); }
function clearDemoData()   { notYet_('Демо-дані з\'являться на етапі 2.'); }
function selfCheck()       { notYet_('Перевірка з\'явиться на етапі 2.'); }
function createEntryForm() { notYet_('Форма для телефону з\'явиться на етапі 4.'); }

function notYet_(msg) {
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '💜 Трекер', 5);
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
