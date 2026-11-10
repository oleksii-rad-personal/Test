/**
 * ТРЕКЕР БЮДЖЕТУ — Google Sheets + Apps Script
 * Етап 1: Старт, Налаштування, Операції, довідники, залежний список «Тип → Категорія».
 *
 * Як встановити: Розширення → Apps Script → вставити цей файл у Code.gs → зберегти →
 * оновити таблицю → меню «Трекер» → «Побудувати трекер з нуля».
 */

// ─────────────────────────────────────────────────────────────────────────────
// ПАЛІТРА — змінюй кольори тут, решта коду бере їх звідси
// ─────────────────────────────────────────────────────────────────────────────
const THEMES = {
  // Жінка — м'яка лілово-рожева
  f: {
    primary: '#8E6BBF', accent: '#E58FB3', headerBg: '#F3E8F7', headerTxt: '#5B3F86',
    calcBg: '#F7F5FA', disabled: '#D9D4E3', text: '#3D3550', muted: '#9A8FB0',
  },
  // Чоловік — блідо-блакитна
  m: {
    primary: '#4F86B5', accent: '#8CC4E8', headerBg: '#E5F1FA', headerTxt: '#2C5677',
    calcBg: '#F3F8FC', disabled: '#CCD9E4', text: '#2E3A46', muted: '#8A9BAA',
  },
};

/** Активна палітра. Спільні кольори тут, а «характер» (основний, акцент, фони) — з THEMES за статтю. */
const THEME = Object.assign({
  input:     '#FFF8D6', // поля вводу (світло-жовті)
  danger:    '#E57373', // червоний — перевищення
  warn:      '#FFD54F', // жовтий — попередження (>80%)
  ok:        '#81C784', // зелений — норма
  okBg:      '#E8F5E9',
  okTxt:     '#2E7D32',
  dangerBg:  '#FDECEC',
  warnBg:    '#FFF6D5',
  bg:        '#FFFFFF',
  font:      'Montserrat',
}, THEMES.m);

// ─────────────────────────────────────────────────────────────────────────────
// НАЗВИ ТА ФОРМАТИ
// ─────────────────────────────────────────────────────────────────────────────
const SHEETS = {
  start:    'Старт',
  dash:     'Дашборд',
  settings: 'Налаштування',
  ops:      'Операції',
  year:     'Рік',
  goals:    'Цілі',
  debts:    'Борги',
  subs:     'Підписки',
  cal:      'Календар',
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
  groups: ['Потреби', 'Бажання', 'Освіта', 'Подарунки'],
  savKinds: ['Подушка й цілі', 'Інвестиції'],
  grpCodes: ['need', 'want', 'edu', 'gift', 'debt', 'save', 'invest'],
  cycle:  ['Від основної виплати', 'Календарний місяць'],
  current: 'Поточний цикл',
  currentMonth: 'Поточний місяць',
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

/**
 * Правила розподілу доходу. Кошик: [назва, частка, групи (коди через пробіл), що входить].
 * Коди груп: need — потреби, want — бажання, edu — освіта, gift — подарунки/допомога,
 * debt — платежі по боргах, save — заощадження, invest — інвестиції.
 */
const RULES = [
  { name: '50/30/20', title: '50/30/20 — класика', tag: 'рекомендую',
    about: 'Половина — на потреби, третина — на бажання, пʼята частина — відкласти. Підходить більшості.',
    buckets: [['Потреби', 0.5, 'need debt', 'житло, їжа, транспорт, борги'],
              ['Бажання', 0.3, 'want edu gift', 'кафе, розваги, одяг, підписки'],
              ['Заощадження', 0.2, 'save invest', 'подушка, цілі, інвестиції']] },
  { name: '70/20/10', title: '70/20/10 — просте',
    about: 'Не ділиш на «потреби» і «бажання»: 70% живеш, 20% відкладаєш, 10% — на борги й допомогу іншим.',
    buckets: [['Життя', 0.7, 'need want edu', 'усі щоденні й обовʼязкові витрати'],
              ['Заощадження', 0.2, 'save invest', 'подушка, цілі, інвестиції'],
              ['Борги й допомога', 0.1, 'debt gift', 'погашення боргів, подарунки, благодійність']] },
  { name: '80/20', title: '80/20 — спершу заплати собі',
    about: 'Одразу після зарплати відкладаєш 20%, а рештою користуєшся вільно, без підрахунку категорій.',
    buckets: [['Заощадження', 0.2, 'save invest', 'відкладаєш першим, щойно прийшли гроші'],
              ['Усе інше', 0.8, 'need want edu gift debt', 'витрачаєш як зручно']] },
  { name: '60/20/20', title: '60/20/20 — коли обовʼязкових витрат багато',
    about: 'Для великих міст і оренди: 60% на потреби, по 20% — на бажання і заощадження.',
    buckets: [['Потреби', 0.6, 'need debt', 'житло, їжа, транспорт, борги'],
              ['Бажання', 0.2, 'want edu gift', 'кафе, розваги, одяг, підписки'],
              ['Заощадження', 0.2, 'save invest', 'подушка, цілі, інвестиції']] },
  { name: '6 глечиків', title: '6 глечиків — для системних',
    about: 'Дохід ділиться на 6 «глечиків»: потреби, фінансова свобода, довгострокові цілі, освіта, розваги, подарунки.',
    buckets: [['Потреби', 0.55, 'need debt', 'житло, їжа, транспорт, борги'],
              ['Фінансова свобода', 0.1, 'invest', 'інвестиції, що приносять дохід'],
              ['Довгострокові', 0.1, 'save', 'подушка, великі покупки, відпустка'],
              ['Освіта', 0.1, 'edu', 'курси, книги, навчання'],
              ['Розваги', 0.1, 'want', 'кафе, хобі, подорожі'],
              ['Подарунки', 0.05, 'gift', 'подарунки, благодійність']] },
  { name: 'Без правила', title: 'Без правила — просто стежити',
    about: 'Без цілей розподілу: трекер лише покаже, куди йдуть гроші.', buckets: [] },
];
const RULE_SLOTS = 6;

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
      { key: 'Group', h: 'Група',                t: 'list:L_Group', w: 120 },
      { key: 'Mine',  h: 'Моя сума',             t: 'f', fmt: 'money', w: 120,
        f: '=IF(OR({Name}="",{Total}=""),"",{Total}*IF({Share}="",1,{Share}))' },
      { key: 'Reserve', h: 'Резерв на місяць',   t: 'f', fmt: 'money', w: 130,
        f: '=IF({Mine}="","",{Mine}/IFERROR(VLOOKUP({Freq},' + NR + 'L_Freq,2,FALSE),1))' },
    ],
  },
  {
    key: 'Subs', title: 'ПІДПИСКИ', rows: 12,
    hint: 'Netflix, Spotify, YouTube, iCloud… Сума — за один платіж. Не користуєшся — зніми «Користуюсь», і я покажу, скільки можна звільнити.',
    cols: [
      { key: 'Name',  h: 'Підписка',                          t: 'text',         w: 190 },
      { key: 'Total', h: 'Сума за раз',                       t: 'money',        w: 130 },
      { key: 'Freq',  h: 'Як часто',                          t: 'list:L_Freq',  w: 130 },
      { key: 'Month', h: 'Місяць оплати (для квартал / рік)', t: 'list:L_Month', w: 150 },
      { key: 'Day',   h: 'День списання',                     t: 'day',          w: 110 },
      { key: 'Use',   h: 'Користуюсь',                        t: 'check',        w: 100 },
      { key: 'Mine',  h: 'Сума',                              t: 'f', fmt: 'money', w: 110,
        f: '=IF(OR({Name}="",{Total}=""),"",{Total})' },
      { key: 'Year',  h: 'За рік',                            t: 'f', fmt: 'money', w: 120,
        f: '=IF({Mine}="","",{Mine}*12/IFERROR(VLOOKUP({Freq},' + NR + 'L_Freq,2,FALSE),1))' },
    ],
  },
  {
    key: 'Var', title: 'ЗМІННІ ВИТРАТИ', rows: 12,
    hint: 'Те, що щодня: продукти, кава, транспорт. Тижневий бюджет рахується сам.',
    cols: [
      { key: 'Name',   h: 'Підкатегорія',      t: 'text',         w: 190 },
      { key: 'Budget', h: 'Місячний бюджет',   t: 'money',        w: 130 },
      { key: 'Group',  h: 'Група',             t: 'list:L_Group', w: 120 },
      { key: 'Week',   h: 'Тижневий бюджет',   t: 'f', fmt: 'money', w: 130,
        f: '=IF(OR({Name}="",{Budget}=""),"",{Budget}*12/52)' },
    ],
  },
  {
    key: 'Debt', title: 'БОРГИ І РОЗСТРОЧКИ', rows: 8,
    hint: 'Кредити, «оплата частинами», борги друзям. Залишок рахується з «Операцій». Закрито раніше — постав галочку.',
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
        f: '=IF({Name}="","",IF({Closed},0,IF(AND({Initial}="",{Count}=""),"",MAX(0,IF({Initial}="",{Pay}*{Count},{Initial})-{Paid}))))' },
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
      { key: 'Kind',    h: 'Вид',               t: 'list:L_SavKind', w: 140 },
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
  Subs: [],
  Sav: [{ Name: 'Подушка безпеки' }, { Name: 'Відпустка' }],
  Acc: [{ Name: 'Картка', Main: true }, { Name: 'Готівка' }],
};

// ─────────────────────────────────────────────────────────────────────────────
// МЕНЮ
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  try { ensureOwnDocument_(); } catch (e) { /* простий тригер може не мати доступу */ }
  SpreadsheetApp.getUi()
    .createMenu(menuTitle_())
    .addItem('👋 Мій профіль (квіз)', 'showQuiz')
    .addItem('📖 Як користуватись', 'showGuide')
    .addItem('Побудувати трекер з нуля', 'buildTracker')
    .addSeparator()
    .addItem('Заповнити демо-даними', 'fillDemoData')
    .addItem('Прибрати демо-дані', 'clearDemoData')
    .addSeparator()
    .addItem('Створити / оновити форму для телефону', 'createEntryForm')
    .addItem('Перевірити таблицю', 'selfCheck')
    .addItem('🧩 Відновити категорії з операцій', 'repairCategories')
    .addSeparator()
    .addItem('🏦 Імпорт виписки з банку…', 'showStatementImport')
    .addItem('Імпорт з іншої таблиці…', 'importBudgetSheet')
    .addItem('📱 Підключити застосунок на телефон', 'showConnectApp')
    .addItem('📤 Підготувати копію для поширення', 'makeShareCopy')
    .addToUi();
  showOpenHint_();
}

/** В. Підказка справа внизу при кожному відкритті: де меню і з чого почати. */
function showOpenHint_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const p = profile_();
    const msg = hasProfile_()
      ? 'Привіт, ' + p.name + '! Меню «' + menuTitle_() + '» — вгорі над таблицею, поруч із «Довідка». Кнопки дій — на вкладці «Старт».'
      : '👋 Почни тут: меню «' + menuTitle_() + '» вгорі над таблицею → «👋 Мій профіль (квіз)». Або кнопка «Пройти квіз» на вкладці «Старт».';
    ss.toast(msg, menuTitle_(), 12);
  } catch (e) { /* простий тригер — без підказки, не критично */ }
}

// ─────────────────────────────────────────────────────────────────────────────
// ПОБУДОВА
// ─────────────────────────────────────────────────────────────────────────────
/**
 * opts.keep — з квізу: зберегти дані без діалогу. З меню (без opts) і без профілю — спершу квіз.
 */
function buildTracker(opts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureOwnDocument_();
  const fromQuiz = opts && typeof opts === 'object' && 'keep' in opts;
  if (!fromQuiz && !hasProfile_()) {
    try { showQuiz(); return; } catch (e) { /* запуск з редактора — будуємо без квізу */ }
  }
  applyProfileTheme_();
  ARG_SEP_ = null; // роздільник визначаємо заново в кожному запуску

  // Якщо трекер уже є — пропонуємо зберегти введені дані.
  const saved = captureUserData_(ss);
  const keep = saved ? (fromQuiz ? !!opts.keep : askKeepData_(saved)) : false;
  if (!keep) resetDemoState_();
  deleteTrackerSheets_(ss);

  // Спочатку створюємо всі аркуші, бо вони посилаються один на одного.
  const order = [SHEETS.start, SHEETS.settings, SHEETS.dash, SHEETS.ops, SHEETS.year, SHEETS.goals,
                 SHEETS.debts, SHEETS.subs, SHEETS.cal, SHEETS.lists];
  const sh = {};
  order.forEach((name, i) => { sh[name] = ss.insertSheet(name, i); });
  removeTempSheet_(ss);
  detectArgSeparator_(sh[SHEETS.lists]);

  buildListsStatic_(sh[SHEETS.lists]);
  buildSettings_(ss, sh[SHEETS.settings]);
  buildListsDynamic_(ss, sh[SHEETS.lists]);
  writeProfileCells_(ss);
  buildOps_(ss, sh[SHEETS.ops]);
  buildDashboard_(ss, sh[SHEETS.dash]);
  buildYear_(ss, sh[SHEETS.year]);
  buildGoals_(ss, sh[SHEETS.goals]);
  buildDebts_(ss, sh[SHEETS.debts]);
  buildSubs_(ss, sh[SHEETS.subs]);
  buildCalendar_(ss, sh[SHEETS.cal]);
  buildStart_(sh[SHEETS.start]);
  if (keep) { restoreUserData_(ss, saved); syncCategoriesFromOps_(ss); }

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
    : '✓ ' + nameComma_().replace(/""/g, '"') + 'трекер побудовано, помилок немає. Почни з вкладки «Налаштування».';
  Logger.log(msg + (errors.length ? '\n' + errors.slice(0, 50).join('\n') : ''));
  ss.toast(msg, menuTitle_(), 10);
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
    writeOpsRows_(sh, 2, rows);
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
    'SPARKLINE', 'значення', 'Календар', 'Місяців історії']]).setFontWeight('bold');

  const put = (col, values) => sh.getRange(col + '2:' + col + (values.length + 1)).setValues(values.map(v => [v]));
  put('A', REF.types);
  sh.getRange(2, 7, REF.freq.length, 2).setValues(REF.freq);
  put('I', REF.groups);
  put('J', REF.cycle);
  put('L', REF.months);
  put('M', [REF.current].concat(REF.months));
  put('P', [REF.currentMonth].concat(REF.months));
  ss.setNamedRange(NR + 'L_CalPeriod', sh.getRange(2, 16, REF.months.length + 1, 1));
  sh.getRange('Q2').setValue(HISTORY_MONTHS);
  sh.getRange('T1:W1').setValues([['Правила', 'Коди кошиків', 'Коди груп', 'Види заощаджень']]).setFontWeight('bold');
  put('T', RULES.map(r => r.name));
  put('V', REF.grpCodes);
  put('W', REF.savKinds);
  ss.setNamedRange(NR + 'L_Rule', sh.getRange(2, 20, RULES.length, 1));
  ss.setNamedRange(NR + 'Rule_Codes', sh.getRange(2, 21, RULE_SLOTS, 1));
  ss.setNamedRange(NR + 'L_GrpCodes', sh.getRange(2, 22, REF.grpCodes.length, 1));
  ss.setNamedRange(NR + 'L_SavKind', sh.getRange(2, 23, REF.savKinds.length, 1));
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
    C: nonBlank('VSTACK(' + NR + 'Reg_Name,' + NR + 'Subs_Name,' + NR + 'Var_Name)'),
    D: nonBlank(NR + 'Debt_Name'),
    E: nonBlank(NR + 'Sav_Name'),
    F: nonBlank(NR + 'Acc_Name'),
    K: nonBlank('VSTACK(' + NR + 'Inc_Name,' + NR + 'Reg_Name,' + NR + 'Subs_Name,' + NR + 'Var_Name,' + NR + 'Debt_Name,' + NR + 'Sav_Name)'),
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
  row = buildRuleBlock_(ss, sh, row);

  // Значення за замовчуванням загального блоку.
  const today = todayParts_(ss);
  ss.getRangeByName(NR + 'Year').setValue(today.y);
  ss.getRangeByName(NR + 'Start').setValue(isoDate_(today.y, today.m, 1));
  ss.getRangeByName(NR + 'Cycle').setValue(REF.cycle[0]);
  ss.getRangeByName(NR + 'Warn').setValue(0.8);
  ss.getRangeByName(NR + 'Over').setValue(1);
  applyRule_(ss, (profile_().rule) || RULES[0].name);

  // Дата розрахунку, що діє насправді (порожня → сьогодні).
  const calc = ss.getRangeByName(NR + 'Calc');
  const calcNow = calc.offset(0, 1);
  calcNow.setFormula(F_('=IF(' + calc.getA1Notation() + '="",TODAY(),' + calc.getA1Notation() + ')'))
    .setNumberFormat(FMT.date).setFontColor(THEME.primary).setFontWeight('bold');
  ss.setNamedRange(NR + 'CalcDate', calcNow);
  calc.offset(0, 2).setValue('← діє зараз. ' + GENERAL_ROWS[2].hint);
  protect_(calcNow, 'Дата розрахунку');

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

  sh.setFrozenRows(2);
  trimSheet_(sh, row + 2, 2 + Math.max.apply(null, SETTINGS_BLOCKS.map(b => b.cols.length)) + 1);
}

/** Блок «Правило розподілу»: вибір правила + до 6 кошиків (назва, частка, що входить). */
function buildRuleBlock_(ss, sh, row) {
  writeBlockTitle_(sh, row, 'ПРАВИЛО РОЗПОДІЛУ', 3);
  row++;
  sh.getRange(row, 2).setValue('Правило').setFontWeight('bold');
  const sel = sh.getRange(row, 3);
  sel.setBackground(THEME.input).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange_(ss, 'L_Rule'), true).setAllowInvalid(false)
    .setHelpText('Обери правило — кошики й частки заповняться самі.').build());
  ss.setNamedRange(NR + 'Rule', sel);
  sh.getRange(row, 4).setValue('Обери правило — кошики й частки заповняться самі. Частки можна змінити.')
    .setFontColor(THEME.muted).setFontStyle('italic');
  row++;
  sh.getRange(row, 2, 1, 3).setValues([['Кошик', 'Частка', 'Що входить']])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerTxt).setFontWeight('bold');
  row++;
  const names = sh.getRange(row, 2, RULE_SLOTS, 1);
  const shares = sh.getRange(row, 3, RULE_SLOTS, 1);
  const desc = sh.getRange(row, 4, RULE_SLOTS, 1);
  names.setBackground(THEME.input);
  styleInput_(ss, shares, 'pct');
  desc.setFontColor(THEME.muted).setFontStyle('italic');
  ss.setNamedRange(NR + 'Rule_Names', names);
  ss.setNamedRange(NR + 'Rule_Shares', shares);
  ss.setNamedRange(NR + 'Rule_Desc', desc);
  row += RULE_SLOTS;
  const chk = sh.getRange(row, 3);
  chk.setFormula(F_('=IF(COUNTA(' + names.getA1Notation() + ')=0,"",IF(ROUND(SUM(' + shares.getA1Notation() + '),4)=1,"✓ разом 100%","⚠ разом має бути 100%"))'))
    .setFontColor(THEME.primary).setFontWeight('bold');
  protect_(chk, 'Перевірка суми часток');
  return row + 2;
}

/** Записує кошики обраного правила в «Налаштування» і коди груп у «_Списки». */
function applyRule_(ss, name) {
  const rule = RULES.filter(r => r.name === name)[0] || RULES[0];
  const sel = ss.getRangeByName(NR + 'Rule');
  if (!sel) return;
  if (sel.getValue() !== rule.name) sel.setValue(rule.name);
  const pad = (arr, f) => { const out = arr.map(f); while (out.length < RULE_SLOTS) out.push(['']); return out; };
  ss.getRangeByName(NR + 'Rule_Names').setValues(pad(rule.buckets, b => [b[0]]));
  ss.getRangeByName(NR + 'Rule_Shares').setValues(pad(rule.buckets, b => [b[1]]));
  ss.getRangeByName(NR + 'Rule_Desc').setValues(pad(rule.buckets, b => [b[3]]));
  ss.getRangeByName(NR + 'Rule_Codes').setValues(pad(rule.buckets, b => [b[2]]));
  if (!rule.buckets.length) {
    ss.getRangeByName(NR + 'Rule_Desc').getCell(1, 1)
      .setValue('Кошиків немає — трекер просто показує, куди йдуть гроші. Хочеш ціль — обери правило вище.');
  }
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
  col(OPS.account).setDataValidation(accountRule_(ss));
  col(OPS.source).setFontColor(THEME.muted);

  [['O_Date', OPS.date], ['O_Type', OPS.type], ['O_Amount', OPS.amount], ['O_Category', OPS.category],
   ['O_Account', OPS.account], ['O_Source', OPS.source]].forEach(p => ss.setNamedRange(NR + p[0], col(p[1])));

  sh.getRange(1, 1, DATA_ROWS, OPS_HEADERS.length).createFilter();
  sh.hideColumns(OPS.source);
  protect_(sh.getRange(1, 1, 1, OPS_HEADERS.length), 'Шапка «Операцій»');
  trimSheet_(sh, DATA_ROWS, OPS_HEADERS.length);
}

// ─── «Старт» ─────────────────────────────────────────────────────────────────
const START = { formRow: 14, checkRow: 18, checkLines: 10, profileRow: 31 };

function buildStart_(sh) {
  styleSheet_(sh);
  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 640);
  sh.setColumnWidth(4, 290);
  writeStartBanner_(sh);
  sh.getRange('B2').setValue(spaced_('ТРЕКЕР БЮДЖЕТУ')).setFontSize(22).setFontWeight('bold').setFontColor(THEME.primary);
  sh.getRange('B3').setValue('Вписуєш суму — таблиця сама рахує, скільки можна витратити до наступної виплати.')
    .setFontColor(THEME.muted).setFontStyle('italic');

  writeBlockTitle_(sh, 5, 'ЯК ПОЧАТИ', 2);
  const steps = [
    'Відкрий «Налаштування» і заповни жовті клітинки: дохід і день основної виплати, обов\'язкові платежі, бюджети на щодень, борги, цілі, рахунки.',
    'Прийшли гроші або щось ' + g_('купила', 'купив') + ' — додай рядок в «Операції»: сума, тип, категорія. Дата і тип «Витрата» поставляться самі.',
    'З телефона ще швидше — через форму (посилання нижче). Це 10 секунд прямо на касі.',
    'Дивись «Дашборд»: велика цифра — скільки можна витратити до наступної виплати і скільки це на день.',
    'Раз на тиждень заглядай у «Рік», «Цілі», «Борги», «Підписки» і «Календар». ' + g_('Пропустила', 'Пропустив') + ' кілька днів — нічого страшного, просто допиши.',
  ];
  steps.forEach((s, i) => {
    sh.getRange(6 + i, 2).setValue(i + 1).setFontSize(16).setFontWeight('bold').setFontColor(THEME.accent)
      .setHorizontalAlignment('center').setVerticalAlignment('top');
    sh.getRange(6 + i, 3).setValue(s).setWrap(true).setVerticalAlignment('top');
    sh.setRowHeight(6 + i, 42);
  });

  writeBlockTitle_(sh, START.formRow - 1, 'ФОРМА ДЛЯ ТЕЛЕФОНУ', 2);
  sh.getRange(START.formRow, 3).setValue('Ще не створена: меню «Трекер» → «Створити / оновити форму для телефону».')
    .setFontColor(THEME.muted);
  const formUrl = PropertiesService.getDocumentProperties().getProperty(FORM.propUrl);
  if (formUrl) writeFormLink_(sh, formUrl);

  writeBlockTitle_(sh, START.checkRow - 1, 'ПЕРЕВІРКА ТАБЛИЦІ', 2);
  sh.getRange(START.checkRow, 3).setValue('Меню «Трекер» → «Перевірити таблицю».').setFontColor(THEME.muted);
  writeStartProfile_(sh);
  writeStartMap_(sh);
  trimSheet_(sh, 50, 4);
  insertStartButtons_(sh);
}

/** А. Банер угорі «Старту»: де шукати меню. */
function writeStartBanner_(sh) {
  sh.getRange(1, 2, 1, 3).merge()
    .setValue('👆 Усе керування — у меню «' + menuTitle_() + '» над таблицею, поруч із «Довідка» (Help). ' +
      'Найчастіші дії — кнопками праворуч 👉')
    .setBackground(THEME.headerBg).setFontColor(THEME.headerTxt).setFontWeight('bold').setFontSize(12)
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 44);
}

const START_BUTTONS = [
  ['quiz', 'showQuiz', 'Пройти квіз'],
  ['guide', 'showGuide', 'Як користуватись'],
  ['import', 'showStatementImport', 'Імпорт виписки'],
  ['form', 'createEntryForm', 'Форма для телефону'],
];

/** Б. Кнопки-картинки на «Старті» з привʼязаними діями (працюють на компʼютері). */
function insertStartButtons_(sh) {
  try {
    sh.getImages().forEach(img => img.remove());
    const theme = profile_().gender === 'f' ? 'f' : 'm';
    sh.getRange(4, 4).setValue(spaced_('ШВИДКІ ДІЇ')).setFontWeight('bold').setFontColor(THEME.primary).setFontSize(11);
    START_BUTTONS.forEach((b, i) => {
      const blob = Utilities.newBlob(Utilities.base64Decode(BUTTON_PNG[theme][b[0]]), 'image/png', 'btn-' + b[0] + '.png');
      const img = sh.insertImage(blob, 4, 5 + i * 2, 8, 6);
      img.setWidth(250).setHeight(50).assignScript(b[1]);
      try { img.setAltTextTitle(b[2]); } catch (e) { /* не критично */ }
    });
    sh.getRange(5 + START_BUTTONS.length * 2, 4).setValue('Кнопки працюють на компʼютері. На телефоні — меню «Трекер».')
      .setFontColor(THEME.muted).setFontStyle('italic').setFontSize(9).setWrap(true);
  } catch (e) {
    Logger.log('Кнопки на «Старті» не вставлено: ' + e.message);
  }
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

function accountRule_(ss) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange_(ss, 'L_Acc'), true).setAllowInvalid(false)
    .setHelpText('Порожньо — основний рахунок.').build();
}

/**
 * Записує рядки в «Операції» скриптом. Строгі списки категорій/рахунків стоять на кожній
 * клітинці окремо, і Apps Script відхиляє запис, що їм не відповідає (напр. після зсуву рядків),
 * тому спершу знімаємо ці перевірки з рядків, пишемо, а потім ставимо заново під нові типи.
 */
function writeOpsRows_(sh, top, rows) {
  const n = rows.length;
  if (!n) return;
  sh.getRange(top, OPS.category, n, 2).clearDataValidations(); // категорія + рахунок
  sh.getRange(top, 1, n, OPS_HEADERS.length).setValues(rows);
  sh.getRange(top, OPS.account, n, 1).setDataValidation(accountRule_(sh.getParent()));
  applyCategoryValidation_(sh, top, n, false);
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
  const ssE = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const ruleCell = ssE.getRangeByName(NR + 'Rule');
  if (ruleCell && within_(e.range, ruleCell)) { applyRule_(ssE, e.range.getValue()); return; }
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
  sh.setColumnWidth(2, 330);
  [170, 120, 120, 120, 120, 130].forEach((w, i) => sh.setColumnWidth(3 + i, w));
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
    '"' + nameComma_() + '≈ "&FIXED(@C_Limit,0)&" грн на день, "&@C_Days&" дн. до виплати"&' +
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
  // Іменовані діапазони блоків — їх читає застосунок.
  blocks.forEach(b => ss.setNamedRange(NR + 'DB_' + b.key, sh.getRange(pos[b.key].top, 2, b.n, Math.max(1, b.cols.length))));

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
    { key: 'D_SubsLeft', label: '− Ще сплатити: підписки', f: '=SUM(' + colRange('subs', 'H') + ')' },
    { key: 'D_DebtLeft', label: '− Ще сплатити: борги', f: '=SUM(' + colRange('debt', 'E') + ')' },
    { key: 'D_SavLeft', label: '− Ще відкласти', f: '=SUM(' + colRange('sav', 'E') + ')' },
    { key: 'D_Free', label: 'Підсумок: можу витратити до наступної виплати', bold: true,
      f: '=@D_Income-@D_Spent-@D_RegLeft-@D_SubsLeft-@D_DebtLeft-@D_SavLeft' },
  ];
}

/**
 * Опис блоків дашборду. Кожна колонка: h — заголовок, f(i, r) — формула для i-го рядка
 * (i — номер рядка в блоці «Налаштувань», r — рядок аркуша), fmt — формат, t — тип підсумку.
 */
function dashBlocks_() {
  const inc = TYPES.income, exp = TYPES.expense, debt = TYPES.debt, sav = TYPES.saving;
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
    payBlock_('reg', 'РЕГУЛЯРНІ ПЛАТЕЖІ', 'Reg', 15, exp,
      g_('Оплатила', 'Оплатив') + ' — внеси операцію з цією категорією, галочка ✓ поставиться сама.'),
    payBlock_('subs', 'ПІДПИСКИ', 'Subs', 12, exp, 'Списання підписки — це операція «Витрата» з назвою підписки.'),
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
            ',IF(AND(dd<>"",INDEX(@Debt_Closed,' + i + ')<>TRUE),MIN(N(INDEX(@Debt_Pay,' + i + ')),IF(INDEX(@Debt_Left,' + i + ')="",N(INDEX(@Debt_Pay,' + i + ')),N(INDEX(@Debt_Left,' + i + '))+D' + r + ')),0)))' },
        { h: 'Сплачено в циклі', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(debt, 'B' + r) + ')' },
        { h: 'Ще сплатити', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(@C_Past,0,MAX(0,C' + r + '-D' + r + '))' },
        { h: 'Залишок боргу', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Debt_Left,' + i + ')))' },
        { h: 'Прогноз закриття',
          f: (i, r) => '=IF(B' + r + '="","",IF(AND(INDEX(@Debt_Left,' + i + ')="",INDEX(@Debt_Closed,' + i + ')<>TRUE),"допиши залишок у «Налаштуваннях»",IF(F' + r + '=0,"✓ закрито",LET(p,N(INDEX(@Debt_Pay,' + i + ')),' +
            'k,IFERROR(VLOOKUP(INDEX(@Debt_Freq,' + i + '),@L_Freq,2,FALSE),1),IF(p=0,"—",' +
            'LET(n,ROUNDUP(F' + r + '/p,0),d,EDATE(@C_Start,IF(E' + r + '>0,(n-1)*k,n*k)),INDEX(@L_Month,MONTH(d))&" "&YEAR(d)))))))' },
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
        ['Регулярні й підписки', '=SUM(' + cr('reg', 'E') + ')+SUM(' + cr('reg', 'H') + ')+SUM(' + cr('subs', 'E') + ')+SUM(' + cr('subs', 'H') + ')'],
        ['Змінні та інше', '=MAX(0,' + sumCycle_(exp) + '-SUM(' + cr('reg', 'E') + ')-SUM(' + cr('subs', 'E') + '))'],
        ['Борги', '=' + sumCycle_(debt) + '+SUM(' + cr('debt', 'E') + ')'],
        ['Заощадження', '=' + sumCycle_(sav) + '+SUM(' + cr('sav', 'E') + ')'],
        ['Вільний залишок', '=MAX(0,@D_Free)'],
      ],
    },
    {
      key: 'grp', title: 'КУДИ ЙДУТЬ ГРОШІ — ЗА ГРУПАМИ', n: 7, total: false,
      hint: 'Група категорії задається в «Налаштуваннях». Сплачене + ще заплановане в цьому циклі.',
      rowsFixed: true,
      cols: [{ h: 'Група' }, { h: 'Сума', fmt: FMT.money0 }, { h: '% доходу', fmt: FMT.pct }],
      rows: cr => {
        const regAll = '(' + cr('reg', 'E') + '+' + cr('reg', 'H') + ')';
        const g = k => 'INDEX(@L_Group,' + k + ')';
        const byGroup = (k, blankToo) => '=SUMPRODUCT(((@Reg_Group=' + g(k) + ')' + (k === 1 ? '+(@Reg_Group="")' : '') + ')*' + regAll + ')+' +
          'SUMPRODUCT(((@Var_Group=' + g(k) + ')' + (blankToo ? '+(@Var_Group="")' : '') + ')*' + cr('var', 'D') + ')';
        const savBy = kind => 'SUMPRODUCT((@Sav_Kind' + kind + 'INDEX(@L_SavKind,2))*(' + cr('sav', 'D') + '+' + cr('sav', 'E') + '))';
        return [
          ['Потреби', byGroup(1, false)],
          ['Бажання (з підписками й іншим)', byGroup(2, true) + '+SUM(' + cr('subs', 'E') + ')+SUM(' + cr('subs', 'H') + ')+' +
            'MAX(0,' + sumCycle_(exp) + '-SUM(' + cr('reg', 'E') + ')-SUM(' + cr('subs', 'E') + ')-SUM(' + cr('var', 'D') + '))'],
          ['Освіта', byGroup(3, false)],
          ['Подарунки й допомога', byGroup(4, false)],
          ['Борги', '=' + sumCycle_(debt) + '+SUM(' + cr('debt', 'E') + ')'],
          ['Заощадження', '=' + savBy('<>')],
          ['Інвестиції', '=' + savBy('=')],
        ];
      },
    },
    {
      key: 'rule', title: 'ПРАВИЛО РОЗПОДІЛУ', n: RULE_SLOTS, total: false,
      hint: 'Правило й частки — у «Налаштуваннях». Різниця: + перевищено ціль, − ще є запас.',
      rowsFixed: true,
      cols: [{ h: 'Кошик' }, { h: 'Сума', fmt: FMT.money0 }, { h: '% доходу', fmt: FMT.pct },
             { h: 'Ціль', fmt: FMT.pct }, { h: 'Різниця', fmt: FMT.signPct }],
      rows: cr => Array.from({ length: RULE_SLOTS }, (x, k) => [
        '=IF(INDEX(@Rule_Names,' + (k + 1) + ')="","",INDEX(@Rule_Names,' + (k + 1) + '))',
        '=IF(INDEX(@Rule_Names,' + (k + 1) + ')="","",SUMPRODUCT(ISNUMBER(SEARCH(@L_GrpCodes,INDEX(@Rule_Codes,' + (k + 1) + ')))*' + cr('grp', 'C') + '))',
        '=IF(INDEX(@Rule_Names,' + (k + 1) + ')="","",INDEX(@Rule_Shares,' + (k + 1) + '))',
      ]),
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
        ['=ARRAYFORMULA(LET(nm,@Subs_Name,' +
          'yr,IFERROR(IF((@Subs_Use<>TRUE)*(nm<>""),@Subs_Year*1,0),0),mx,MAX(yr),' +
          'IF(mx>0,"💸 Підписка «"&INDEX(nm,MATCH(mx,yr,0))&"» коштує "&FIXED(mx,0)&" грн на рік, а ти нею не користуєшся. Скасуй — і ці гроші твої.",' +
          '"✓ Невикористаних підписок немає.")))'],
        ['=ARRAYFORMULA(LET(nm,@Debt_Name,l,IFERROR(@Debt_Left*1,0),a,(nm<>"")*(l>0)*(@Debt_Closed<>TRUE),' +
          'c,IF(a,l,MAX(l)+1),mn,MIN(c),' +
          'IF(SUM(a)=0,IF(SUMPRODUCT((@Debt_Name<>"")*(@Debt_Left="")*(@Debt_Closed<>TRUE))>0,"❄ Допиши залишки боргів у «Налаштуваннях» — тоді покажу, який закривати першим.","✓ Активних боргів немає. Так тримати!"),' +
          '"❄ Сніжна куля: першим закривай «"&INDEX(nm,MATCH(mn,c,0))&"» — залишок "&FIXED(mn,0)&" грн.")))'],
      ],
    },
  ];
}

/** Блок «платежі за графіком» (регулярні, підписки): дата, сплачено, ✓, днів, ще сплатити. */
function payBlock_(key, title, src, n, exp, hint) {
  return {
    key: key, title: title, n: n, total: true, hint: hint,
    cols: [
      { h: 'Назва', f: i => '=INDEX(@' + src + '_Name,' + i + ')', tot: 'label' },
      { h: 'Сума', fmt: FMT.money0, f: (i, r) => '=IF(B' + r + '="",0,N(INDEX(@' + src + '_Mine,' + i + ')))' },
      { h: 'Дата оплати', fmt: FMT.date,
        f: (i, r) => '=IF(B' + r + '="","",' + dueDateLet_(
          'IF(INDEX(@' + src + '_Day,' + i + ')="",DAY(@C_Start),INDEX(@' + src + '_Day,' + i + '))',
          'INDEX(@' + src + '_Freq,' + i + ')',
          'IFERROR(MATCH(INDEX(@' + src + '_Month,' + i + '),@L_Month,0),MONTH(@Start))') + ')' },
      { h: 'Сплачено', fmt: FMT.money0, tot: 'sum', f: (i, r) => '=IF(B' + r + '="",0,' + sumCycle_(exp, 'B' + r) + ')' },
      { h: 'Оплачено', f: (i, r) => '=IF(AND(B' + r + '<>"",C' + r + '>0,E' + r + '>=C' + r + '),"✓","")' },
      { h: 'Днів до оплати', fmt: FMT.days,
        f: (i, r) => '=IF(OR(D' + r + '="",F' + r + '="✓",@C_Past),"",D' + r + '-@C_Today)' },
      { h: 'Ще сплатити', fmt: FMT.money0, tot: 'sum',
        f: (i, r) => '=IF(OR(D' + r + '="",@C_Past),0,MAX(0,C' + r + '-E' + r + '))' },
    ],
  };
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
      if (String(rv[0]).charAt(0) === '=') sh.getRange(r, 2).setFormula(X_(rv[0])); else sh.getRange(r, 2).setValue(rv[0]);
      sh.getRange(r, 3).setFormula(X_(rv[1])).setNumberFormat(b.cols[1].fmt);
      sh.getRange(r, 4).setFormula(X_('=IF(OR(@D_Income=0,C' + r + '=""),"",C' + r + '/@D_Income)')).setNumberFormat(FMT.pct);
      if (rv[2]) {
        sh.getRange(r, 5).setFormula(X_(rv[2])).setNumberFormat(FMT.pct);
        sh.getRange(r, 6).setFormula(X_('=IF(OR(D' + r + '="",E' + r + '=""),"",D' + r + '-E' + r + ')')).setNumberFormat(FMT.signPct);
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
    ss.toast('Спочатку побудуй трекер: меню «Трекер» → «Побудувати трекер з нуля».', menuTitle_(), 8);
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
  setKV_(ss, { Year: t.y, Start: start, Calc: '', Cycle: REF.cycle[0], Warn: 0.8, Over: 1 });
  applyRule_(ss, RULES[0].name);
  setBlockRows_(ss, 'Inc', [
    { Name: 'Зарплата', Main: true, Exp: 30000, Day: 1 },
    { Name: 'Фріланс', Main: false },
  ]);
  setBlockRows_(ss, 'Reg', [
    { Name: 'Оренда', Total: 30000, Share: 0.5, Freq: 'Щомісяця', Day: 1, Group: 'Потреби' },
    { Name: 'Комуналка + інтернет', Total: 1950, Share: 1, Freq: 'Щомісяця', Day: 3, Group: 'Потреби' },
  ]);
  setBlockRows_(ss, 'Subs', [
    { Name: 'Фоторедактор', Total: 450, Freq: 'Раз на квартал', Month: mo(1), Day: 12, Use: false },
    { Name: 'Хмарне сховище', Total: 1200, Freq: 'Раз на рік', Month: mo(2), Day: 20, Use: false },
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
  writeOpsRows_(sh, first, ops);

  ss.getRangeByName(NR + 'D_Period').setValue(REF.current);
  props.setProperty(PROP_DEMO, '1');
  SpreadsheetApp.flush();

  const free = ss.getRangeByName(NR + 'D_Free').getValue();
  const foreign = countForeignOpsInCycle_(ss);
  ss.toast('Демо-дані додано. «Можу витратити» = ' + fmtMoney_(free) + ' (очікується ' + fmtMoney_(DEMO_EXPECTED) + ').' +
    (foreign ? ' Увага: у поточному циклі є ще ' + foreign + ' твоїх операцій — вони змінюють цифру.' : ''),
    menuTitle_(), 10);
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
    (restored ? 'Налаштування повернуто як були до демо.' : 'Збережених налаштувань не було — їх не змінено.'), menuTitle_(), 8);
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
  add(0, d(1), D, 4770, 'Розстрочка за телефон', '', g_('Закрила', 'Закрив') + ' повністю, щоб не тягнулася');
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
  const keys = GENERAL_ROWS.map(r => r.key).concat(['Rule', 'Rule_Names', 'Rule_Shares', 'Rule_Desc']);
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
  writeOpsRows_(sh, 2, out);
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
      comment: r[OPS.comment - 1],
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
  applyProfileTheme_();
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
    // Твої (не-демо) операції в циклі змінюють цифру: дохід додає, решта віднімає.
    const cs = ss.getRangeByName(NR + 'C_Start').getValue();
    const cn = ss.getRangeByName(NR + 'C_Next').getValue();
    const own = opsRows_(ss).filter(o => o.source !== DEMO_TAG && o.type && o.date instanceof Date && o.date >= cs && o.date < cn);
    const delta = own.reduce((a, o) => a + (o.type === TYPES.income ? o.amount : -o.amount), 0);
    const expected = DEMO_EXPECTED + delta;
    const pass = Math.abs(free - expected) < 0.005;
    const parts = dashCalcLines_(() => '').map(c => c.key + '=' + ss.getRangeByName(NR + c.key).getValue()).join(', ');
    ok(pass, 'Демо: «Можу витратити» = ' + fmtMoney_(free) +
      (own.length
        ? ' = ' + fmtMoney_(DEMO_EXPECTED) + ' демо ' + (delta < 0 ? '− ' : '+ ') + fmtMoney_(Math.abs(delta)) +
          ' (твоїх операцій у циклі: ' + own.length + ')'
        : ' (очікується ' + fmtMoney_(DEMO_EXPECTED) + ')') +
      (pass ? '' : '. Має бути ' + fmtMoney_(expected) + '. Складові: ' + parts));
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
    ? 'Не всі операції мають відому категорію: ' + unknown.slice(0, 5).join('; ') +
      '. Виправити: меню «Трекер» → «🧩 Відновити категорії з операцій»'
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

  // 4б. «Рік»: сума всіх груп (з «Без категорії») = сума операцій кожного місяця року.
  const year = Number(ss.getRangeByName(NR + 'Year').getValue());
  const yearTot = new Array(12).fill(0);
  YEAR_GROUPS.forEach(g => {
    const rng = ss.getRangeByName(NR + 'Y_Tot_' + g.key);
    if (rng) rng.getValues()[0].forEach((v, m) => { yearTot[m] += Number(v) || 0; });
  });
  const opsTot = new Array(12).fill(0);
  ops.forEach(o => {
    if (o.type && o.date instanceof Date && o.date.getFullYear() === year) opsTot[o.date.getMonth()] += o.amount;
  });
  const yearBad = [];
  for (let m = 0; m < 12; m++) {
    if (Math.abs(yearTot[m] - opsTot[m]) > 0.005) yearBad.push(REF.months[m] + ': «Рік» ' + fmtMoney_(yearTot[m]) + ', «Операції» ' + fmtMoney_(opsTot[m]));
  }
  ok(!yearBad.length, yearBad.length ? '«Рік» не збігається з «Операціями»: ' + yearBad.join('; ')
    : '«Рік»: сума по категоріях = сума в «Операціях» для кожного місяця ' + year);

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

  // 5б. Форма для телефону.
  const formId = props.getProperty(FORM.propId);
  if (formId) {
    let formOk = false;
    try { formOk = !!FormApp.openById(formId).getPublishedUrl(); } catch (e) { formOk = false; }
    const trig = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === FORM.handler).length;
    ok(formOk && trig === 1, formOk && trig === 1 ? 'Форма для телефону працює, тригер один'
      : 'Форма: ' + (formOk ? '' : 'не відкривається; ') + 'тригерів onFormSubmit: ' + trig + ' (має бути 1) — запусти «Створити / оновити форму»');
  } else {
    res.push('• Форма для телефону ще не створена.');
  }

  // 6. Знак $ у значеннях і форматах.
  const dollars = [];
  trackerSheetNames_().forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const rng = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    const dv = rng.getDisplayValues();
    // Власний звіт перевірки на «Старті» не перевіряємо — він описує саму перевірку.
    if (name === SHEETS.start) {
      for (let r = START.checkRow - 1; r < Math.min(dv.length, START.checkRow - 1 + START.checkLines); r++) dv[r] = dv[r].map(() => '');
    }
    const nf = rng.getNumberFormats();
    dv.forEach((row, r) => row.forEach((v, c) => {
      if (String(v).indexOf('$') !== -1 || String(nf[r][c]).indexOf('$') !== -1) dollars.push(name + '!' + colLetter_(c + 1) + (r + 1));
    }));
  });
  ok(!dollars.length, dollars.length ? 'Знайдено знак долара: ' + dollars.slice(0, 10).join(', ') : 'Ніде немає знака долара — лише «грн»');

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
  ss.toast(failed ? '⚠ Не пройдено перевірок: ' + failed + '. Деталі — на вкладці «Старт».' : '✓ Усі перевірки пройдено.', menuTitle_(), 8);
  return res;
}

// ─────────────────────────────────────────────────────────────────────────────
// ЕТАП 3: РІК, ЦІЛІ, БОРГИ, ПІДПИСКИ, КАЛЕНДАР
// ─────────────────────────────────────────────────────────────────────────────

/** Заголовок аркуша: велика назва + підказка. */
function sheetHead_(sh, title, hint) {
  styleSheet_(sh);
  sh.setColumnWidth(1, 24);
  sh.getRange('B1').setValue(spaced_(title)).setFontSize(18).setFontWeight('bold').setFontColor(THEME.primary);
  if (hint) sh.getRange('B2').setValue(hint).setFontColor(THEME.muted).setFontStyle('italic');
}

function headerRow_(sh, row, col, titles) {
  sh.getRange(row, col, 1, titles.length).setValues([titles])
    .setBackground(THEME.headerBg).setFontColor(THEME.headerTxt).setFontWeight('bold').setWrap(true)
    .setVerticalAlignment('middle');
}

/** Записує стовпець формул f(i, r) для n рядків, починаючи з top. */
function formulaCol_(sh, top, col, n, f, fmt) {
  const out = [];
  for (let i = 1; i <= n; i++) out.push([X_(f(i, top + i - 1))]);
  const rng = sh.getRange(top, col, n, 1).setFormulas(out);
  if (fmt) rng.setNumberFormat(fmt);
  return rng;
}

// ─── «Рік» ───────────────────────────────────────────────────────────────────
const YEAR_GROUPS = [
  { key: 'inc',  title: 'ДОХІД',          type: TYPES.income,  names: 'Inc_Name',  n: 8,  other: true },
  { key: 'reg',  title: 'РЕГУЛЯРНІ',      type: TYPES.expense, names: 'Reg_Name',  n: 15, other: false, heat: true },
  { key: 'subs', title: 'ПІДПИСКИ',       type: TYPES.expense, names: 'Subs_Name', n: 12, other: false, heat: true },
  { key: 'var',  title: 'ЗМІННІ',         type: TYPES.expense, names: 'Var_Name',  n: 12, other: true,  heat: true },
  { key: 'debt', title: 'БОРГИ',          type: TYPES.debt,    names: 'Debt_Name', n: 8,  other: true },
  { key: 'sav',  title: 'ЗАОЩАДЖЕННЯ',    type: TYPES.saving,  names: 'Sav_Name',  n: 8,  other: true },
];

function monthSum_(type, catRef, m) {
  return 'SUMIFS(@O_Amount,@O_Type,"' + type + '"' + (catRef ? ',@O_Category,' + catRef : '') +
    ',@O_Date,">="&DATE(@Year,' + m + ',1),@O_Date,"<"&DATE(@Year,' + (m + 1) + ',1))';
}

function buildYear_(ss, sh) {
  sheetHead_(sh, 'РІК', 'Кожна клітинка — сума операцій за календарний місяць року з «Налаштувань». Чим насиченіший колір, тим більша витрата.');
  const head = 4, firstCol = 3, totalCol = 15; // C..N — місяці, O — ВСЬОГО
  sh.setColumnWidth(2, 210);
  for (let c = firstCol; c < totalCol; c++) sh.setColumnWidth(c, 84);
  sh.setColumnWidth(totalCol, 104);
  sh.getRange(3, 2).setFormula(X_('="Рік: "&@Year')).setFontWeight('bold').setFontColor(THEME.headerTxt);
  headerRow_(sh, head, 2, ['Категорія'].concat(REF.months).concat(['ВСЬОГО']));

  let row = head + 1;
  const pos = {};
  const heatRanges = [];
  YEAR_GROUPS.forEach(g => {
    // Заголовок групи
    sh.getRange(row, 2, 1, totalCol - 1).setBackground(THEME.calcBg);
    sh.getRange(row, 2).setValue(spaced_(g.title)).setFontWeight('bold').setFontColor(THEME.primary);
    row++;
    const top = row, bottom = row + g.n - 1;
    formulaCol_(sh, top, 2, g.n, i => '=INDEX(@' + g.names + ',' + i + ')');
    for (let m = 1; m <= 12; m++) {
      formulaCol_(sh, top, firstCol + m - 1, g.n, (i, r) => '=IF($B' + r + '="","",' + monthSum_(g.type, '$B' + r, m) + ')', FMT.money0);
    }
    formulaCol_(sh, top, totalCol, g.n, (i, r) => '=IF($B' + r + '="","",SUM(C' + r + ':N' + r + '))', FMT.money0);
    if (g.heat) heatRanges.push(sh.getRange(top, firstCol, g.n, 12));
    row = bottom + 1;
    pos[g.key] = { top: top, bottom: bottom };

    let otherRow = null;
    if (g.other) {
      otherRow = row;
      sh.getRange(row, 2).setValue('Без категорії / інше').setFontColor(THEME.muted).setFontStyle('italic');
      for (let m = 1; m <= 12; m++) {
        const L = colLetter_(firstCol + m - 1);
        // Для витрат «інше» = усі витрати − регулярні − змінні.
        const minus = g.type === TYPES.expense
          ? '-SUM(' + L + pos.reg.top + ':' + L + pos.reg.bottom + ')-SUM(' + L + pos.subs.top + ':' + L + pos.subs.bottom + ')-SUM(' + L + top + ':' + L + bottom + ')'
          : '-SUM(' + L + top + ':' + L + bottom + ')';
        sh.getRange(row, firstCol + m - 1).setFormula(X_('=' + monthSum_(g.type, null, m) + minus)).setNumberFormat(FMT.money0);
      }
      sh.getRange(row, totalCol).setFormula('=SUM(C' + row + ':N' + row + ')').setNumberFormat(FMT.money0);
      row++;
    }
    // Разом по групі
    sh.getRange(row, 2).setValue('Разом: ' + g.title.toLowerCase()).setFontWeight('bold');
    for (let c = firstCol; c <= totalCol; c++) {
      const L = colLetter_(c);
      sh.getRange(row, c).setFormula('=SUM(' + L + top + ':' + L + (otherRow || bottom) + ')').setNumberFormat(FMT.money0);
    }
    sh.getRange(row, 2, 1, totalCol - 1).setFontWeight('bold').setBorder(true, null, null, null, null, null,
      THEME.primary, SpreadsheetApp.BorderStyle.SOLID);
    ss.setNamedRange(NR + 'Y_Tot_' + g.key, sh.getRange(row, firstCol, 1, 12));
    pos[g.key].total = row;
    row += 2;
  });

  // ── Підсумок року
  writeBlockTitle_(sh, row, 'ПІДСУМОК РОКУ', totalCol - 1);
  row++;
  const sumHead = row;
  headerRow_(sh, sumHead, 2, ['По місяцях'].concat(REF.months).concat(['ВСЬОГО']));
  const t = k => pos[k].total;
  const lines = [
    ['Дохід', L => '=' + L + t('inc')],
    ['Витрати (з боргами)', L => '=' + L + t('reg') + '+' + L + t('subs') + '+' + L + t('var') + '+' + L + t('debt')],
    ['Заощаджено', L => '=' + L + t('sav')],
    ['Залишок', L => '=' + L + (sumHead + 1) + '-' + L + (sumHead + 2) + '-' + L + (sumHead + 3)],
    ['Накопичено з початку року', L => '=SUM($C' + (sumHead + 3) + ':' + L + (sumHead + 3) + ')'],
  ];
  lines.forEach((ln, k) => {
    const r = sumHead + 1 + k;
    sh.getRange(r, 2).setValue(ln[0]).setFontWeight('bold');
    for (let c = firstCol; c <= totalCol; c++) {
      const L = colLetter_(c);
      let f = ln[1](L);
      if (c === totalCol) f = k === 4 ? '=N' + r : '=SUM(C' + r + ':N' + r + ')';
      sh.getRange(r, c).setFormula(f).setNumberFormat(FMT.money0);
    }
  });
  const leftRow = sumHead + 4;
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(THEME.danger)
      .setRanges([sh.getRange(leftRow, firstCol, 1, 13)]).build(),
  ];
  // Теплова карта витрат.
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue(THEME.accent, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(THEME.danger)
    .setRanges(heatRanges).build());
  sh.setConditionalFormatRules(rules);

  // Лінійний графік «Дохід vs витрати».
  const chart = sh.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(sumHead, 2, 3, 13))
    .setTransposeRowsAndColumns(true)
    .setNumHeaders(1)
    .setOption('title', 'Дохід vs витрати по місяцях')
    .setOption('colors', [THEME.primary, THEME.danger])
    .setOption('legend', { position: 'bottom' })
    .setOption('curveType', 'function')
    .setOption('width', 760).setOption('height', 320)
    .setPosition(sumHead + lines.length + 2, 2, 0, 0)
    .build();
  sh.insertChart(chart);

  sh.setFrozenRows(head);
  sh.setFrozenColumns(2);
  sh.protect().setDescription('Рік — рахується автоматично').setWarningOnly(true);
  trimSheet_(sh, sumHead + lines.length + 22, totalCol + 1);
}

// ─── «Цілі» ──────────────────────────────────────────────────────────────────
function buildGoals_(ss, sh) {
  sheetHead_(sh, 'ЦІЛІ', 'Скарбничка бажань. Нову ціль додай у «Налаштування → Заощадження і цілі», а внески — в «Операції» з типом «Заощадження».');
  [190, 120, 120, 70, 140, 110, 90, 130, 300].forEach((w, i) => sh.setColumnWidth(2 + i, w));
  const head = 4, top = 5, n = 8;
  headerRow_(sh, head, 2, ['Ціль', 'Потрібна сума', 'Накопичено', '%', 'Прогрес', 'Бажана дата',
    'Місяців лишилось', 'Відкладати на місяць', 'Статус']);
  const plan = i => 'N(INDEX(@Sav_Monthly,' + i + '))';
  formulaCol_(sh, top, 2, n, i => '=INDEX(@Sav_Name,' + i + ')');
  formulaCol_(sh, top, 3, n, (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Sav_Target,' + i + ')))', FMT.money0);
  formulaCol_(sh, top, 4, n, (i, r) => '=IF(B' + r + '="",0,N(INDEX(@Sav_Saved,' + i + ')))', FMT.money0);
  formulaCol_(sh, top, 5, n, (i, r) => '=IF(OR(B' + r + '="",C' + r + '=0),"",D' + r + '/C' + r + ')', FMT.pct);
  formulaCol_(sh, top, 6, n, (i, r) => '=IF(E' + r + '="","",SPARKLINE(MIN(E' + r + ',1),@SP_Main))');
  formulaCol_(sh, top, 7, n, (i, r) => '=IF(B' + r + '="","",INDEX(@Sav_Date,' + i + '))', FMT.date);
  formulaCol_(sh, top, 8, n, (i, r) => '=IF(OR(B' + r + '="",G' + r + '=""),"",IFERROR(DATEDIF(@C_Today,G' + r + ',"M"),0))', FMT.int);
  formulaCol_(sh, top, 9, n, (i, r) => '=IF(OR(B' + r + '="",G' + r + '="",C' + r + '=0),"",IF(D' + r + '>=C' + r + ',0,(C' + r + '-D' + r + ')/MAX(1,H' + r + ')))', FMT.money);
  formulaCol_(sh, top, 10, n, (i, r) => '=IF(B' + r + '="","",IF(AND(C' + r + '>0,D' + r + '>=C' + r + '),"✓ Ціль досягнута 🎉",' +
    'IF(G' + r + '="","Додай бажану дату — порахую, скільки відкладати",' +
    'IF(' + plan(i) + '>=I' + r + ',"✓ Встигаєш: відкладаєш "&FIXED(' + plan(i) + ',0)&" грн на місяць",' +
    '"⚠ Не встигаєш: додай ще "&FIXED(I' + r + '-' + plan(i) + ',0)&" грн на місяць або перенеси дату"))))');

  const tot = top + n;
  sh.getRange(tot, 2).setValue('Разом');
  sh.getRange(tot, 3).setFormula('=SUM(C' + top + ':C' + (tot - 1) + ')').setNumberFormat(FMT.money);
  sh.getRange(tot, 4).setFormula('=SUM(D' + top + ':D' + (tot - 1) + ')').setNumberFormat(FMT.money);
  sh.getRange(tot, 5).setFormula(X_('=IF(C' + tot + '=0,"",D' + tot + '/C' + tot + ')')).setNumberFormat(FMT.pct);
  sh.getRange(tot, 6).setFormula(X_('=IF(E' + tot + '="","",SPARKLINE(MIN(E' + tot + ',1),@SP_Main))'));
  sh.getRange(tot, 9).setFormula('=SUM(I' + top + ':I' + (tot - 1) + ')').setNumberFormat(FMT.money);
  sh.getRange(tot, 2, 1, 9).setFontWeight('bold').setBorder(true, null, null, null, null, null,
    THEME.primary, SpreadsheetApp.BorderStyle.SOLID);

  const st = sh.getRange(top, 10, n, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✓').setFontColor(THEME.okTxt).setRanges([st]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setFontColor(THEME.danger).setRanges([st]).build(),
  ]);
  sh.setFrozenRows(head);
  sh.protect().setDescription('Цілі — рахуються автоматично').setWarningOnly(true);
  trimSheet_(sh, tot + 3, 11);
}

// ─── «Борги» ─────────────────────────────────────────────────────────────────
function buildDebts_(ss, sh) {
  sheetHead_(sh, 'БОРГИ', 'Метод сніжної кулі: спершу закриваєш найменший борг, а його платіж додаєш до наступного. Без відсотків — прогноз приблизний.');
  [40, 210, 130, 130, 110, 150].forEach((w, i) => sh.setColumnWidth(2 + i, w));
  const n = 8, head = 8, top = 9, bottom = top + n - 1;
  const hid = 12; // L:N — відсортований список (приховано)

  // Відсортовані активні борги: назва | залишок | платіж на місяць.
  sh.getRange(head, hid, 1, 3).setValues([['Борг', 'Залишок', 'Платіж/міс']]);
  sh.getRange(top, hid).setFormula(X_('=ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(@Debt_Name,IFERROR(@Debt_Left*1,0),' +
    'IFERROR(@Debt_Pay/VLOOKUP(@Debt_Freq,@L_Freq,2,FALSE),0)),' +
    '(@Debt_Name<>"")*(IFERROR(@Debt_Left*1,0)>0)*(@Debt_Closed<>TRUE)),2,TRUE),""))'));

  writeBlockTitle_(sh, 6, 'ПОРЯДОК ПОГАШЕННЯ', 6);
  headerRow_(sh, head, 2, ['№', 'Борг', 'Залишок', 'Платіж на місяць', 'Місяців до закриття', 'Прогноз закриття']);
  formulaCol_(sh, top, 2, n, (i, r) => '=IF(L' + r + '="","",' + i + ')');
  formulaCol_(sh, top, 3, n, (i, r) => '=IF(L' + r + '="","",L' + r + ')');
  formulaCol_(sh, top, 4, n, (i, r) => '=IF(L' + r + '="","",M' + r + ')', FMT.money);
  formulaCol_(sh, top, 5, n, (i, r) => '=IF(L' + r + '="","",N' + r + ')', FMT.money);
  formulaCol_(sh, top, 6, n, (i, r) => '=IF(L' + r + '="","",IF(SUM(N$' + top + ':N' + r + ')=0,"—",' +
    'ROUNDUP(SUM(M$' + top + ':M' + r + ')/SUM(N$' + top + ':N' + r + '),0)))', FMT.int);
  formulaCol_(sh, top, 7, n, (i, r) => '=IF(NOT(ISNUMBER(F' + r + ')),"",INDEX(@L_Month,MONTH(EDATE(@C_Today,F' + r + ')))&" "&YEAR(EDATE(@C_Today,F' + r + ')))');

  // Велике повідомлення і загальний прогрес.
  sh.getRange(3, 2, 1, 6).merge().setFormula(X_('=IF(L' + top + '="",IF(SUMPRODUCT((@Debt_Name<>"")*(@Debt_Left="")*(@Debt_Closed<>TRUE))>0,"Допиши залишки боргів у «Налаштуваннях»","✓ Активних боргів немає — так тримати!"),' +
    '"❄ ' + g_('Вільна', 'Вільний') + ' від боргів: "&INDEX(G' + top + ':G' + bottom + ',MATCH(MAX(F' + top + ':F' + bottom + '),F' + top + ':F' + bottom + ',0)))'))
    .setFontSize(16).setFontWeight('bold').setFontColor(THEME.headerTxt).setBackground(THEME.headerBg)
    .setHorizontalAlignment('center');
  sh.setRowHeight(3, 40);
  sh.getRange(5, 2, 1, 6).merge().setFormula(X_('=IFERROR("⚠ Без залишку (не в сніжній кулі): "&TEXTJOIN(", ",TRUE,' +
    'FILTER(@Debt_Name,(@Debt_Name<>"")*(@Debt_Left="")*(@Debt_Closed<>TRUE))),"")')).setFontColor(THEME.danger).setWrap(true);
  sh.getRange(4, 2, 1, 2).merge().setValue('Загальний прогрес').setFontColor(THEME.muted);
  sh.getRange(4, 4).setFormula(X_('=IFERROR(SUM(@Debt_Paid)/(SUM(@Debt_Paid)+SUM(@Debt_Left)),"")')).setNumberFormat(FMT.pct)
    .setFontWeight('bold');
  sh.getRange(4, 5, 1, 2).merge().setFormula(X_('=IF(D4="","",SPARKLINE(D4,@SP_Ok))'));
  sh.getRange(4, 7).setFormula(X_('="сплачено "&FIXED(SUM(@Debt_Paid),0)&" грн"')).setFontColor(THEME.muted);

  // Закриті борги.
  const closedTop = bottom + 3;
  writeBlockTitle_(sh, closedTop - 1, 'ЗАКРИТІ', 6);
  headerRow_(sh, closedTop, 3, ['Борг', 'Сплачено']);
  sh.getRange(closedTop + 1, 3).setFormula(X_('=ARRAYFORMULA(IFERROR(FILTER(HSTACK(@Debt_Name,IFERROR(@Debt_Paid*1,0)),' +
    '(@Debt_Name<>"")*((IFERROR(@Debt_Left*1,-1)=0)+(@Debt_Closed=TRUE)>0)),"Поки немає"))'));
  sh.getRange(closedTop + 1, 3, n, 2).setFontColor(THEME.muted);
  sh.getRange(closedTop + 1, 4, n, 1).setNumberFormat(FMT.money);

  sh.hideColumns(hid, 3);
  sh.protect().setDescription('Борги — рахуються автоматично').setWarningOnly(true);
  trimSheet_(sh, closedTop + n + 3, hid + 2);
}

// ─── «Підписки» ──────────────────────────────────────────────────────────────
function buildSubs_(ss, sh) {
  sheetHead_(sh, 'ПІДПИСКИ', 'Блок «Підписки» в «Налаштуваннях». Не користуєшся — зніми «Користуюсь» і скасуй підписку.');
  [210, 110, 130, 130, 110, 120, 120].forEach((w, i) => sh.setColumnWidth(2 + i, w));
  const n = 15, head = 6, top = 7, idx = 12; // L — номери рядків підписок у «Налаштуваннях» (приховано)
  sh.getRange(head, idx).setValue('№ у налаштуваннях');
  sh.getRange(top, idx).setFormula(X_('=ARRAYFORMULA(IFERROR(FILTER(ROW(@Subs_Name)-ROW(INDEX(@Subs_Name,1))+1,' +
    '@Subs_Name<>""),""))'));

  headerRow_(sh, head, 2, ['Підписка', 'Моя сума', 'Як часто', 'Наступне списання', 'Через', 'За рік', 'Користуюсь']);
  const at = (rng, r) => 'INDEX(@' + rng + ',L' + r + ')';
  formulaCol_(sh, top, 2, n, (i, r) => '=IF(L' + r + '="","",' + at('Subs_Name', r) + ')');
  formulaCol_(sh, top, 3, n, (i, r) => '=IF(L' + r + '="",0,N(' + at('Subs_Mine', r) + '))', FMT.money0);
  formulaCol_(sh, top, 4, n, (i, r) => '=IF(L' + r + '="","",' + at('Subs_Freq', r) + ')');
  formulaCol_(sh, top, 5, n, (i, r) => '=IF(L' + r + '="","",ARRAYFORMULA(LET(t,@C_Today,' +
    'g,IF(' + at('Subs_Day', r) + '="",DAY(t),' + at('Subs_Day', r) + '),' +
    'k,IFERROR(VLOOKUP(D' + r + ',@L_Freq,2,FALSE),1),' +
    'f,IFERROR(MATCH(' + at('Subs_Month', r) + ',@L_Month,0),MONTH(@Start)),' +
    'mo,DATE(YEAR(t),MONTH(t)+SEQUENCE(13,1,0),1),' +
    'dd,DATE(YEAR(mo),MONTH(mo),IF(g>DAY(EOMONTH(mo,0)),DAY(EOMONTH(mo,0)),g)),' +
    'IFERROR(MIN(FILTER(dd,(MOD(MONTH(dd)-f,k)=0)*(dd>=t))),""))))', FMT.date);
  formulaCol_(sh, top, 6, n, (i, r) => '=IF(E' + r + '="","",E' + r + '-@C_Today)', FMT.days);
  formulaCol_(sh, top, 7, n, (i, r) => '=IF(L' + r + '="",0,C' + r + '*12/IFERROR(VLOOKUP(D' + r + ',@L_Freq,2,FALSE),1))', FMT.money0);
  formulaCol_(sh, top, 8, n, (i, r) => '=IF(L' + r + '="","",IF(' + at('Subs_Use', r) + '=TRUE,"✓ так","✗ ні"))');

  const bottom = top + n - 1;
  sh.getRange(3, 2, 1, 7).merge().setFormula(X_(
    '=IF(COUNT(L' + top + ':L' + bottom + ')=0,"Підписок поки немає — познач їх у «Налаштуваннях».",' +
    '"Підписки коштують "&FIXED(SUM(G' + top + ':G' + bottom + '),0)&" грн на рік")'))
    .setFontSize(15).setFontWeight('bold').setFontColor(THEME.headerTxt).setBackground(THEME.headerBg).setHorizontalAlignment('center');
  sh.setRowHeight(3, 36);
  sh.getRange(4, 2, 1, 7).merge().setFormula(X_(
    '=IF(SUMIF(H' + top + ':H' + bottom + ',"✗ ні",G' + top + ':G' + bottom + ')=0,"✓ Усіма підписками ти користуєшся.",' +
    '"💸 Можна звільнити "&FIXED(SUMIF(H' + top + ':H' + bottom + ',"✗ ні",G' + top + ':G' + bottom + '),0)&" грн на рік — скасуй ті, де «✗ ні».")'))
    .setFontColor(THEME.danger).setHorizontalAlignment('center');

  const rows = sh.getRange(top, 2, n, 7);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H' + top + '="✗ ні"')
      .setBackground(THEME.dangerBg).setRanges([rows]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(F_('=AND(ISNUMBER($F' + top + '),$F' + top + '<=3)'))
      .setBackground(THEME.warnBg).setRanges([sh.getRange(top, 5, n, 2)]).build(),
  ]);
  sh.hideColumns(idx);
  sh.setFrozenRows(head);
  sh.protect().setDescription('Підписки — рахуються автоматично').setWarningOnly(true);
  trimSheet_(sh, bottom + 3, idx + 1);
}

// ─── «Календар» ──────────────────────────────────────────────────────────────
const CAL = { selRow: 2, titleRow: 4, dowRow: 5, gridTop: 6, weeks: 6, hidLabel: 11, hidVal: 12 };

function buildCalendar_(ss, sh) {
  sheetHead_(sh, 'КАЛЕНДАР', '');
  for (let c = 2; c <= 8; c++) sh.setColumnWidth(c, 104);
  sh.getRange(CAL.selRow, 2).setValue('Місяць').setFontWeight('bold').setHorizontalAlignment('right');
  const sel = sh.getRange(CAL.selRow, 3);
  sel.setValue(REF.currentMonth).setBackground(THEME.input).setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listRange_(ss, 'L_CalPeriod'), true)
      .setAllowInvalid(false).setHelpText('Поточний місяць або місяць року з «Налаштувань».').build());
  ss.setNamedRange(NR + 'K_Sel', sel);

  // Службові клітинки (K:L, приховані).
  const hidden = [
    ['K_First', 'Перший день місяця', '=IF(@K_Sel=INDEX(@L_CalPeriod,1),DATE(YEAR(@C_Today),MONTH(@C_Today),1),DATE(@Year,MATCH(@K_Sel,@L_Month,0),1))', FMT.date],
    ['K_GridStart', 'Понеділок першого тижня', '=@K_First-WEEKDAY(@K_First,2)+1', FMT.date],
    ['K_Days', 'Днів у місяці', '=DAY(EOMONTH(@K_First,0))', FMT.int],
    ['K_DayBudget', 'Денний бюджет змінних', '=IFERROR(SUM(@Var_Budget)/@K_Days,0)', FMT.money],
    ['K_TodayDay', 'Сьогоднішнє число (якщо цей місяць)', '=IF(EOMONTH(@C_Today,0)=EOMONTH(@K_First,0),DAY(@C_Today),0)', FMT.int],
  ];
  hidden.forEach((h, i) => ss.setNamedRange(NR + h[0], sh.getRange(1 + i, CAL.hidVal)));
  hidden.forEach((h, i) => {
    sh.getRange(1 + i, CAL.hidLabel).setValue(h[1]);
    sh.getRange(1 + i, CAL.hidVal).setFormula(X_(h[2])).setNumberFormat(h[3]);
  });
  const abs = key => '$' + colLetter_(CAL.hidVal) + '$' + (1 + hidden.findIndex(h => h[0] === key));

  sh.getRange(CAL.titleRow, 2, 1, 7).merge()
    .setFormula(X_('=UPPER(INDEX(@L_Month,MONTH(@K_First)))&" "&YEAR(@K_First)'))
    .setFontSize(16).setFontWeight('bold').setFontColor(THEME.primary).setHorizontalAlignment('center');
  headerRow_(sh, CAL.dowRow, 2, ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд']);
  sh.getRange(CAL.dowRow, 2, 1, 7).setHorizontalAlignment('center');

  const dayRows = [], amtRows = [];
  for (let w = 0; w < CAL.weeks; w++) {
    const dr = CAL.gridTop + w * 2, ar = dr + 1;
    dayRows.push(dr); amtRows.push(ar);
    const dayF = [], amtF = [];
    for (let d = 0; d < 7; d++) {
      const off = w * 7 + d;
      dayF.push(X_('=LET(dt,@K_GridStart+' + off + ',IF(MONTH(dt)=MONTH(@K_First),DAY(dt),""))'));
      amtF.push(X_('=LET(dt,@K_GridStart+' + off + ',IF(OR(MONTH(dt)<>MONTH(@K_First),dt>@C_Today),"",' +
        'SUMPRODUCT(N(@O_Amount)*(@O_Type="' + TYPES.expense + '")*(@O_Date>=dt)*(@O_Date<dt+1)*' +
        '(@O_Category<>"")*ISNUMBER(MATCH(@O_Category,@Var_Name,0)))))'));
    }
    sh.getRange(dr, 2, 1, 7).setFormulas([dayF]).setFontColor(THEME.muted).setFontSize(9).setVerticalAlignment('bottom');
    sh.getRange(ar, 2, 1, 7).setFormulas([amtF]).setNumberFormat('#,##0 "грн";-#,##0 "грн";"—"')
      .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('top');
    sh.setRowHeight(dr, 18);
    sh.setRowHeight(ar, 34);
  }
  const gridBottom = CAL.gridTop + CAL.weeks * 2 - 1;
  sh.getRange(CAL.gridTop, 2, CAL.weeks * 2, 7).setBorder(true, true, true, true, true, null,
    THEME.headerBg, SpreadsheetApp.BorderStyle.SOLID);

  // Кольори: більше денного бюджету — червоний, у межах — зелений, 0 — «день без витрат».
  const amtRanges = amtRows.map(r => sh.getRange(r, 2, 1, 7));
  const dayRanges = dayRows.map(r => sh.getRange(r, 2, 1, 7));
  const a1 = 'B' + amtRows[0];
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(F_('=AND(ISNUMBER(' + a1 + '),' + a1 + '>0,' + abs('K_DayBudget') + '>0,' + a1 + '>' + abs('K_DayBudget') + ')'))
      .setBackground(THEME.dangerBg).setFontColor(THEME.danger).setRanges(amtRanges).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(F_('=AND(ISNUMBER(' + a1 + '),' + a1 + '>0)'))
      .setBackground(THEME.okBg).setFontColor(THEME.okTxt).setRanges(amtRanges).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(F_('=AND(ISNUMBER(' + a1 + '),' + a1 + '=0)'))
      .setBackground(THEME.headerBg).setFontColor(THEME.primary).setRanges(amtRanges).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(F_('=AND(ISNUMBER(B' + dayRows[0] + '),B' + dayRows[0] + '=' + abs('K_TodayDay') + ')'))
      .setBackground(THEME.primary).setFontColor('#FFFFFF').setRanges(dayRanges).build(),
  ];
  sh.setConditionalFormatRules(rules);

  // Підсумок під сіткою.
  const amtList = amtRows.map(r => 'B' + r + ':H' + r).join(',');
  const s = gridBottom + 2;
  const summary = [
    ['Змінні витрати за місяць', '=SUM(' + amtList + ')', FMT.money],
    ['Денний бюджет змінних', '=@K_DayBudget', FMT.money],
    ['Днів без витрат', '=' + amtRows.map(r => 'COUNTIF(B' + r + ':H' + r + ',0)').join('+'), FMT.int],
    ['Днів понад бюджет', '=IF(@K_DayBudget=0,0,' + amtRows.map(r => 'COUNTIF(B' + r + ':H' + r + ',">"&@K_DayBudget)').join('+') + ')', FMT.int],
    ['Найдорожчий день', '=MAX(' + amtList + ')', FMT.money],
  ];
  writeBlockTitle_(sh, s, 'ПІДСУМОК МІСЯЦЯ', 7);
  summary.forEach((x, i) => {
    sh.getRange(s + 1 + i, 2, 1, 2).merge().setValue(x[0]);
    sh.getRange(s + 1 + i, 4).setFormula(X_(x[1])).setNumberFormat(x[2]).setFontWeight('bold');
  });
  sh.getRange(s + summary.length + 2, 2, 1, 7).merge()
    .setValue('Зелений — у межах денного бюджету, червоний — понад нього, блакитний «—» — день без змінних витрат.')
    .setFontColor(THEME.muted).setFontStyle('italic').setWrap(true);

  sh.hideColumns(CAL.hidLabel, 2);
  protect_(sh.getRange(1, 1, 1, CAL.hidVal), 'Календар: шапка');
  protect_(sh.getRange(CAL.selRow, 4, 1, CAL.hidVal - 3), 'Календар');
  protect_(sh.getRange(CAL.selRow + 1, 1, s + summary.length + 2 - CAL.selRow, CAL.hidVal), 'Календар');
  trimSheet_(sh, s + summary.length + 3, CAL.hidVal);
}

// ─────────────────────────────────────────────────────────────────────────────
// ЕТАП 4: ФОРМА ДЛЯ ТЕЛЕФОНУ
// Сторінка 1: сума + тип. Тип веде на свою секцію з категоріями цього типу.
// Відповідь тригер onFormSubmit одразу переносить у «Операції».
// ─────────────────────────────────────────────────────────────────────────────
const FORM = {
  title: 'Нова операція',
  amount: 'Сума, грн', type: 'Тип', category: 'Категорія', account: 'Рахунок', comment: 'Коментар', date: 'Дата',
  source: 'форма',
  propId: 'BT_FORM_ID', propUrl: 'BT_FORM_URL', propSs: 'BT_SS_ID',
  responsesSheet: '_Форма',
  handler: 'onFormSubmit',
};

function createEntryForm() {
  applyProfileTheme_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.ops)) {
    ss.toast('Спочатку побудуй трекер: меню «Трекер» → «Побудувати трекер з нуля».', menuTitle_(), 8);
    return;
  }
  const props = PropertiesService.getDocumentProperties();
  let form = null;
  const oldId = props.getProperty(FORM.propId);
  if (oldId) {
    try { form = FormApp.openById(oldId); } catch (e) { form = null; }
  }
  const isNew = !form;
  if (isNew) form = FormApp.create(FORM.title);
  else form.getItems().forEach(it => form.deleteItem(it));

  form.setTitle(FORM.title)
    .setDescription('10 секунд: сума → тип → категорія → «Надіслати». Решта — необов\'язково.')
    .setCollectEmail(false)
    .setAllowResponseEdits(false)
    .setProgressBar(false)
    .setShowLinkToRespondAgain(true)
    .setConfirmationMessage('Записано ✓ Дашборд уже перерахувався.');

  // Сторінка 1.
  form.addTextItem().setTitle(FORM.amount).setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .requireTextMatchesPattern('^ *[0-9][0-9 ]*([.,][0-9]{1,2})? *$')
      .setHelpText('Лише число, наприклад 250 або 1 250,50').build());
  const typeItem = form.addMultipleChoiceItem().setTitle(FORM.type).setRequired(true);

  // Секція для кожного типу.
  const accounts = listValues_(ss, 'L_Acc');
  const pages = [];
  REF.types.forEach((type, k) => {
    const page = form.addPageBreakItem().setTitle(type);
    if (k > 0) page.setGoToPage(FormApp.PageNavigationType.SUBMIT); // попередня секція → «Надіслати»
    pages.push(page);
    const cats = listValues_(ss, TYPE_LIST[type]);
    if (cats.length) {
      form.addMultipleChoiceItem().setTitle(FORM.category).setChoiceValues(cats).setRequired(true);
    } else {
      page.setHelpText('Категорій цього типу ще немає — додай їх у «Налаштуваннях», потім меню «Створити / оновити форму».');
      form.addTextItem().setTitle(FORM.category).setRequired(false);
    }
    if (accounts.length > 1) {
      form.addListItem().setTitle(FORM.account).setHelpText('Порожньо — основний рахунок').setChoiceValues(accounts).setRequired(false);
    }
    form.addTextItem().setTitle(FORM.comment).setRequired(false);
    form.addDateItem().setTitle(FORM.date).setHelpText('Порожньо — сьогодні').setRequired(false);
  });
  typeItem.setChoices(REF.types.map((type, k) => typeItem.createChoice(type, pages[k])));

  // Відповіді — ще й у резервний прихований аркуш.
  if (isNew) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    SpreadsheetApp.flush();
    hideResponsesSheet_(ss, form);
  }

  // Тригер без дублів.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FORM.handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(FORM.handler).forForm(form).onFormSubmit().create();

  let url = form.getPublishedUrl();
  try { url = form.shortenFormUrl(url); } catch (e) { /* лишаємо повне посилання */ }
  props.setProperty(FORM.propId, form.getId());
  props.setProperty(FORM.propUrl, url);
  props.setProperty(FORM.propSs, ss.getId());

  const start = ss.getSheetByName(SHEETS.start);
  if (start) writeFormLink_(start, url);
  ss.toast((isNew ? 'Форму створено' : 'Форму оновлено') + '. Посилання — на вкладці «Старт».', menuTitle_(), 8);
  return url;
}

function listValues_(ss, key) {
  const rng = listRange_(ss, key);
  return rng ? rng.getValues().map(v => v[0]).filter(v => v !== '' && v !== null) : [];
}

function hideResponsesSheet_(ss, form) {
  const formUrl = form.getEditUrl().replace(/\/edit.*$/, '');
  const sh = ss.getSheets().filter(s => {
    const u = s.getFormUrl && s.getFormUrl();
    return u && u.replace(/\/viewform.*$|\/edit.*$/, '') === formUrl;
  })[0];
  if (!sh) return;
  let name = FORM.responsesSheet;
  if (ss.getSheetByName(name) && ss.getSheetByName(name).getSheetId() !== sh.getSheetId()) name += ' ' + Date.now();
  sh.setName(name);
  sh.hideSheet();
}

function writeFormLink_(start, url) {
  start.getRange(START.formRow, 3)
    .setFormula(X_('=HYPERLINK("' + url + '","📱 Відкрити форму «Нова операція»")'))
    .setFontSize(12).setFontWeight('bold').setFontColor(THEME.primary);
  start.getRange(START.formRow + 1, 3)
    .setValue('iPhone: відкрий у Safari → «Поділитися» (квадрат зі стрілкою) → «На початковий екран» → «Додати». ' +
      'Android: відкрий у Chrome → меню ⋮ → «Додати на головний екран». Покроково з картинками: меню «Трекер» → «📖 Як користуватись». ' +
      g_('Змінила', 'Змінив') + ' категорії в «Налаштуваннях» — меню «Створити / оновити форму».')
    .setFontColor(THEME.muted).setFontStyle('italic').setWrap(true);
}

/** Інсталюваний тригер форми: відповідь → новий рядок «Операцій». */
function onFormSubmit(e) {
  const props = PropertiesService.getDocumentProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; }
  if (!ss) ss = SpreadsheetApp.openById(props.getProperty(FORM.propSs));

  const answers = {};
  e.response.getItemResponses().forEach(ir => {
    const title = ir.getItem().getTitle();
    const v = ir.getResponse();
    if (v !== '' && v !== null && answers[title] === undefined) answers[title] = v;
  });

  const amount = Number(String(answers[FORM.amount] || '').replace(/\s/g, '').replace(',', '.'));
  if (!(amount > 0)) {
    Logger.log('Форма: некоректна сума «' + answers[FORM.amount] + '» — рядок не додано.');
    return;
  }
  const tz = ss.getSpreadsheetTimeZone();
  const date = answers[FORM.date]
    ? String(answers[FORM.date]).slice(0, 10)
    : Utilities.formatDate(e.response.getTimestamp(), tz, 'yyyy-MM-dd');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(SHEETS.ops);
    const row = firstFreeOpsRow_(sh);
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
    writeOpsRows_(sh, row, [[
      date, answers[FORM.type] || TYPES.expense, amount, answers[FORM.category] || '',
      answers[FORM.account] || '', answers[FORM.comment] || '', FORM.source,
    ]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ІМПОРТ З ТАБЛИЦІ «БЮДЖЕТ» (аркуші-місяці: Категорія | Сума | Звідки; праворуч — Дохід)
// Останній місяць → план у «Налаштуваннях», попередні → історія в «Операціях».
// Особисті назви категорій не зашиті в код: правила нижче лише загальні.
// ─────────────────────────────────────────────────────────────────────────────
const IMPORT = {
  payDay: { 'Зарплата': 1, 'Аванс': 15 },
  otherIncomeDay: 20,
  avansDay: 15,
  sourceTag: 'імпорт',
  // [ключові слова (початок слова), куди, назва в трекері, група 50/30/20]
  rules: [
    [['квартир', 'оренд'],                 'reg',  'Квартира',  'Потреби'],
    [['комунал'],                          'reg',  'Комуналка', 'Потреби'],
    [['психолог'],                         'reg',  'Психолог',  'Потреби'],
    [['мін. платіж', 'мін платіж'],        'debt', 'Кредитна картка (мін. платіж)'],
    [['розср', 'розстр', 'частин', 'кредит 1'], 'debt', 'Оплата частинами'],
    [['кредит'],                           'debt', 'Кредит'],
    [['відклас', 'заощад', 'подушк'],      'sav',  'Подушка безпеки'],
    [['підписк'],                          'var',  'Підписки (разом)', 'Бажання'],
    [['їжа', 'продукт'],                   'var',  'Їжа',       'Потреби'],
    [['відпоч', 'відпуст'],                'var',  'Відпочинок', 'Бажання'],
    [['здоров', 'стоматолог', 'лік'],      'var',  'Здоров\'я', 'Потреби'],
    [['реч', 'одяг'],                      'var',  'Речі',      'Бажання'],
    [['податк'],                           'var',  'Податки',   'Потреби'],
    [['мак', 'ноутбук', 'projector', 'проектор', 'кондиціон', 'технік'], 'var', 'Техніка', 'Бажання'],
    [['спорт'],                            'var',  'Спорт',     'Бажання'],
    [['барбер', 'перукар'],                'var',  'Барбер',    'Бажання'],
    [['невідомо', 'інше'],                 'var',  'Інше',      'Бажання'],
  ],
};

/** Меню: запитує посилання на таблицю і запускає імпорт. */
function importBudgetSheet() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Імпорт з іншої таблиці',
    'Встав посилання на таблицю з аркушами-місяцями (Січень, Лютий…).\n' +
    'Останній місяць стане планом у «Налаштуваннях», попередні — історією в «Операціях».\n' +
    'Поточні налаштування й операції буде замінено (форма для телефону залишиться).',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = String(r.getResponseText()).match(/[-\w]{25,}/);
  if (!m) { ui.alert('Не бачу ID таблиці в посиланні.'); return; }
  const res = importFromSpreadsheet_(m[0]);
  ui.alert('Імпорт завершено', res, ui.ButtonSet.OK);
}

function importFromSpreadsheet_(sourceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.ops)) throw new Error('Спочатку побудуй трекер.');
  const src = SpreadsheetApp.openById(sourceId);
  const year = Number(ss.getRangeByName(NR + 'Year').getValue()) || todayParts_(ss).y;

  // 1. Зчитуємо місяці.
  const months = [];
  REF.months.forEach((name, mi) => {
    const sh = src.getSheets().filter(s => s.getName().trim().toLowerCase() === name)[0];
    if (sh) months.push(Object.assign({ m: mi }, parseBudgetMonth_(sh)));
  });
  if (!months.length) throw new Error('У таблиці немає аркушів з назвами місяців.');
  const plan = months[months.length - 1];
  const history = months.slice(0, -1);
  const subsSheet = src.getSheets().filter(s => /підписк/i.test(s.getName()))[0];
  const subs = subsSheet ? subsSheet.getDataRange().getValues()
    .filter(r => String(r[0]).trim() && Number(r[1]) > 0)
    .map(r => ({ name: String(r[0]).trim(), amount: Number(r[1]) })) : [];

  // 2. Категорії: правила + найчастіші «свої» назви, поки є місце.
  const varCap = SETTINGS_BLOCKS.filter(b => b.key === 'Var')[0].rows;
  const regCap = SETTINGS_BLOCKS.filter(b => b.key === 'Reg')[0].rows;
  const cats = { reg: {}, var: {}, debt: {}, sav: {} }; // назва → { group, plan, src }
  const unknown = {};                                    // назва → кількість місяців
  months.forEach(mo => mo.expenses.forEach(e => {
    const hit = matchImportRule_(e.name);
    if (hit) cats[hit.kind][hit.target] = cats[hit.kind][hit.target] || { group: hit.group || '' };
    else unknown[e.name] = (unknown[e.name] || 0) + 1;
  }));
  const free = varCap - Object.keys(cats.var).length - (cats.var['Інше'] ? 0 : 1);
  const own = Object.keys(unknown).sort((a, b) => unknown[b] - unknown[a]).slice(0, Math.max(0, free));
  own.forEach(n => { cats.var[n] = { group: 'Бажання' }; });
  cats.var['Інше'] = cats.var['Інше'] || { group: 'Бажання' };
  const resolve = name => {
    const hit = matchImportRule_(name);
    if (hit) return hit;
    if (cats.var[name]) return { kind: 'var', target: name };
    const first = name.split(/\s+/)[0];
    if (cats.var[first]) return { kind: 'var', target: first };
    return { kind: 'var', target: 'Інше' };
  };

  // 3. План (останній місяць).
  const sum = (obj, key, v) => { obj[key] = (obj[key] || 0) + v; };
  const planBy = { reg: {}, var: {}, debt: {}, sav: {} };
  const srcBy = {};
  plan.expenses.forEach(e => { const t = resolve(e.name); sum(planBy[t.kind], t.target, e.amount); srcBy[t.target] = e.src; });
  // Підписки поштучно вже є в регулярних — загальну суму в бюджет змінних не дублюємо.
  if (subs.length) Object.keys(planBy.var).forEach(k => { if (/^підписк/i.test(k)) delete planBy.var[k]; });

  // 4. Налаштування.
  const incNames = ['Зарплата', 'Аванс'];
  months.forEach(mo => mo.income.forEach(x => { if (incNames.indexOf(x.name) === -1) incNames.push(x.name); }));
  const planInc = {};
  plan.income.forEach(x => sum(planInc, x.name, x.amount));
  resetDemoState_();
  setKV_(ss, { Year: year, Start: isoDate_(year, months[0].m, 1), Calc: '', Cycle: REF.cycle[0] });
  setBlockRows_(ss, 'Inc', incNames.slice(0, 8).map(n => ({
    Name: n, Main: n === 'Зарплата', Exp: planInc[n] || '', Day: IMPORT.payDay[n] || '',
  })));
  const regRows = Object.keys(cats.reg).map(n => ({
    Name: n, Total: planBy.reg[n] || '', Share: 1, Freq: REF.freq[0][0], Group: cats.reg[n].group || 'Потреби',
    Day: /аванс/i.test(srcBy[n] || '') ? IMPORT.avansDay : 1,
  }));
  setBlockRows_(ss, 'Reg', regRows.slice(0, regCap));
  setBlockRows_(ss, 'Subs', subs.slice(0, 12).map(s => ({ Name: s.name, Total: s.amount, Freq: REF.freq[0][0], Use: true })));
  setBlockRows_(ss, 'Var', Object.keys(cats.var).map(n => ({ Name: n, Budget: planBy.var[n] || '', Group: cats.var[n].group || 'Бажання' })));
  const firstSeen = {};
  months.forEach(mo => mo.expenses.forEach(e => {
    const t = resolve(e.name);
    if (t.kind === 'debt' && firstSeen[t.target] === undefined) firstSeen[t.target] = mo.m;
  }));
  setBlockRows_(ss, 'Debt', Object.keys(cats.debt).map(n => ({
    Name: n, Pay: planBy.debt[n] || '', Freq: REF.freq[0][0],
    Start: isoDate_(year, firstSeen[n] !== undefined ? firstSeen[n] : months[0].m, 1),
    Closed: !planBy.debt[n],
  })));
  setBlockRows_(ss, 'Sav', Object.keys(cats.sav).map(n => ({ Name: n, Monthly: planBy.sav[n] || '' })));
  setBlockRows_(ss, 'Acc', [{ Name: 'Картка', Main: true }, { Name: 'Готівка' }]);
  SpreadsheetApp.flush();

  // 5. Історія → «Операції».
  const typeOf = { reg: TYPES.expense, var: TYPES.expense, debt: TYPES.debt, sav: TYPES.saving };
  const ops = [];
  history.forEach(mo => {
    mo.income.forEach(x => ops.push([isoDate_(year, mo.m, IMPORT.payDay[x.name] || IMPORT.otherIncomeDay),
      TYPES.income, x.amount, x.name, '', '', IMPORT.sourceTag]));
    mo.expenses.forEach(e => {
      const t = resolve(e.name);
      const day = /аванс/i.test(e.src) ? IMPORT.avansDay : 1;
      ops.push([isoDate_(year, mo.m, day), typeOf[t.kind], e.amount, t.target, '',
        e.name !== t.target ? e.name : '', IMPORT.sourceTag]);
    });
  });
  const opsSh = ss.getSheetByName(SHEETS.ops);
  const n = DATA_ROWS - 1;
  const blank = new Array(OPS_HEADERS.length).fill('');
  const out = ops.slice(0, n).concat(Array.from({ length: Math.max(0, n - ops.length) }, () => blank.slice()));
  writeOpsRows_(opsSh, 2, out);
  ss.getRangeByName(NR + 'D_Period').setValue(REF.current);
  SpreadsheetApp.flush();
  syncCategoriesFromOps_(ss);

  const debtsNoLeft = Object.keys(cats.debt).filter(k => planBy.debt[k]);
  const msg = 'Імпортовано ' + ops.length + ' операцій за ' + REF.months[history[0] ? history[0].m : plan.m] + '–' +
    REF.months[history.length ? history[history.length - 1].m : plan.m] + '. План «' + REF.months[plan.m] +
    '» — у «Налаштуваннях» (дохід, регулярні, бюджети, борги' + (subs.length ? ', ' + subs.length + ' підписок у блоці «Підписки»' : '') + ').\n\n' +
    'Що дописати руками:\n' +
    (debtsNoLeft.length ? '• залишок боргів (' + debtsNoLeft.join(', ') + ') — «Налаштування → Борги», колонка «Початковий залишок»;\n' : '') +
    '• фактичні операції за ' + REF.months[plan.m] + ' (зарплата, оплати) — в «Операції» або через форму;\n' +
    '• цілі заощаджень — «Налаштування → Заощадження і цілі».';
  Logger.log(msg);
  return msg;
}

/** Один аркуш-місяць: витрати (A:C, рядки 3–17) і дохід (E3…, підписи в D). */
function parseBudgetMonth_(sh) {
  const v = sh.getRange(1, 1, Math.min(sh.getMaxRows(), 30), Math.min(sh.getMaxColumns(), 6)).getValues();
  const expenses = [];
  for (let r = 2; r < v.length; r++) {
    const name = String(v[r][0]).trim();
    if (/^всього/i.test(name)) break;
    const amount = Number(v[r][1]);
    if (name && amount > 0) expenses.push({ name: name, amount: amount, src: String(v[r][2] || '').trim() });
  }
  const income = [];
  const defaults = ['Зарплата', 'Аванс'];
  for (let r = 2; r < Math.min(v.length, 8); r++) {
    const amount = Number(v[r][4]);
    if (!(amount > 0)) break;
    const label = String(v[r][3] || '').split('\n')[0].trim();
    let name;
    if (/^зп|зарплат/i.test(label)) name = 'Зарплата';
    else if (/аванс/i.test(label)) name = 'Аванс';
    else if (label) name = label.slice(0, 40);
    else name = defaults[r - 2] || 'Інший дохід';
    income.push({ name: name, amount: amount });
  }
  return { expenses: expenses, income: income };
}

function matchImportRule_(name) {
  const low = String(name).toLowerCase().replace(/ʼ|'/g, '');
  for (let i = 0; i < IMPORT.rules.length; i++) {
    const rule = IMPORT.rules[i];
    if (rule[0].some(k => low.indexOf(k.replace(/'/g, '')) === 0 || low.indexOf(' ' + k) !== -1)) {
      return { kind: rule[1], target: rule[2], group: rule[3] };
    }
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// ПРОФІЛЬ І КВІЗ-ЗНАЙОМСТВО
// Ім'я, стать (тема і рід у текстах), вік. Місце під пенсійні питання — далі.
// ─────────────────────────────────────────────────────────────────────────────
const PROP_PROFILE = 'BT_PROFILE';
const PROFILE_FIELDS = [
  ['P_Name', 'Ім\'я'], ['P_Gender', 'Стать'], ['P_Age', 'Вік'], ['P_AsOf', 'Вік на дату'],
  ['P_RetireAge', 'Вік виходу на пенсію'], ['P_RetireIncome', 'Бажаний дохід на пенсії'], ['P_RetireCurrency', 'Валюта'],
];
let PROFILE_ = null;

function profile_() {
  if (PROFILE_) return PROFILE_;
  let p = {};
  try { p = JSON.parse(PropertiesService.getDocumentProperties().getProperty(PROP_PROFILE) || '{}'); } catch (e) { p = {}; }
  PROFILE_ = p;
  return p;
}

function hasProfile_() {
  const p = profile_();
  return !!(p.name && p.gender);
}

/** Текст у жіночому або чоловічому роді. Без профілю — жіночий (основна аудиторія). */
function g_(female, male) {
  return profile_().gender === 'm' ? male : female;
}

function menuTitle_() {
  return (profile_().gender === 'm' ? '💙' : '💜') + ' Трекер';
}

/** «Ім'я, » для формул (лапки екрановано) або порожньо. */
function nameComma_() {
  const n = profile_().name;
  return n ? String(n).replace(/"/g, '""') + ', ' : '';
}

function applyProfileTheme_() {
  Object.assign(THEME, THEMES[profile_().gender === 'f' ? 'f' : 'm']);
}

/** Поточний вік з урахуванням того, скільки часу минуло від квізу. */
function currentAge_() {
  const p = profile_();
  if (!p.age) return '';
  const asOf = p.asOf ? new Date(p.asOf) : new Date();
  return p.age + Math.floor((Date.now() - asOf.getTime()) / (365.25 * 24 * 3600 * 1000));
}

function yearsWord_(n) {
  const a = n % 10, b = n % 100;
  if (a === 1 && b !== 11) return 'рік';
  if (a >= 2 && a <= 4 && (b < 12 || b > 14)) return 'роки';
  return 'років';
}

/** Меню: відкрити квіз. */
/** Меню: інструкція в боковій панелі — таблиця лишається доступною. */
function showGuide() {
  ensureOwnDocument_();
  applyProfileTheme_();
  const html = HtmlService.createHtmlOutput(guideHtml_()).setTitle('Як користуватись трекером');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Відкрити вкладку (з панелі чи квізу). */
function goToSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) return 'Вкладки «' + name + '» ще немає — спершу побудуй трекер.';
  ss.setActiveSheet(sh);
  sh.getRange('A1').activate();
  return 'ok';
}

/** Крок туру: відкрити вкладку й виділити блок, про який іде мова. */
function tourFocus(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const block = blockKey => {
    const name = ss.getRangeByName(NR + blockKey + '_Name');
    if (!name) return null;
    const sh = name.getSheet();
    return sh.getRange(name.getRow() - 1, 2, name.getNumRows() + 1, blockWidth_(blockKey));
  };
  let rng = null;
  if (key === 'income') rng = block('Inc');
  else if (key === 'regular') rng = block('Reg');
  else if (key === 'subs') rng = block('Subs');
  else if (key === 'variable') rng = block('Var');
  else if (key === 'debts') rng = block('Debt');
  else if (key === 'ops') {
    const sh = ss.getSheetByName(SHEETS.ops);
    if (sh) rng = sh.getRange(firstFreeOpsRow_(sh), 1, 1, OPS.comment);
  } else if (key === 'form') {
    const sh = ss.getSheetByName(SHEETS.start);
    if (sh) rng = sh.getRange(START.formRow - 1, 2, 3, 2);
  } else if (key === 'dash') {
    const sh = ss.getSheetByName(SHEETS.dash);
    if (sh) rng = sh.getRange(DASH.bigHeadRow, 2, DASH.expectRow - DASH.bigHeadRow + 1, DASH.cols);
  } else if (key === 'calc') {
    const sh = ss.getSheetByName(SHEETS.dash);
    if (sh) rng = sh.getRange(DASH.calcHeadRow, 2, 7, 2);
  } else {
    return goToSheet(key);
  }
  if (!rng) return 'Спершу побудуй трекер: меню «Трекер» → «Побудувати трекер з нуля».';
  focusCentered_(rng);
  return 'ok';
}

/**
 * Прокручує аркуш так, щоб блок опинився посередині екрана, і виділяє його.
 * Прямого «scroll to» в Apps Script немає: activate() прокручує, поки клітинка не стане видимою
 * біля краю. Тому: скидаємо прокрутку вгору → активуємо «якір» нижче блоку (він стає біля нижнього
 * краю, а блок — посередині) → виділяємо сам блок (він уже видимий, тож екран не стрибає).
 */
const VISIBLE_ROWS = 26; // приблизна кількість рядків на екрані ноутбука
function focusCentered_(rng) {
  const sh = rng.getSheet();
  const ss = sh.getParent();
  ss.setActiveSheet(sh);
  const top = rng.getRow(), h = rng.getNumRows(), col = rng.getColumn();
  sh.getRange(1, col).activate();
  SpreadsheetApp.flush();
  const anchor = h >= VISIBLE_ROWS - 4 ? top + VISIBLE_ROWS - 3 : top + Math.ceil((VISIBLE_ROWS + h) / 2);
  if (anchor > VISIBLE_ROWS) {
    sh.getRange(Math.min(anchor, sh.getMaxRows()), col).activate();
    SpreadsheetApp.flush();
  }
  rng.activate();
}

function showQuiz() {
  ensureOwnDocument_();
  const html = HtmlService.createHtmlOutput(quizHtml_()).setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Давай познайомимось');
}

/**
 * Виклик з квізу.
 * mode: 'build' — перебудувати трекер (дані зберігаються), 'save' — без перебудови.
 * apply: true — записати дохід / платежі / категорії з квізу в «Налаштування».
 */
function quizSubmit(p, mode, apply) {
  const name = String(p && p.name || '').trim().replace(/\s+/g, ' ').slice(0, 30);
  const gender = p && (p.gender === 'm' || p.gender === 'f') ? p.gender : '';
  const age = Math.round(Number(p && p.age));
  if (name.length < 2) throw new Error('Ім\'я має бути хоча б з 2 літер.');
  if (!gender) throw new Error('Обери стать.');
  if (!(age >= 14 && age <= 100)) throw new Error('Вік — від 14 до 100.');
  const clean = s => String(s || '').trim().replace(/\s+/g, ' ').slice(0, 40);
  const num = v => { const n = Number(String(v === undefined ? '' : v).replace(/[^0-9.,]/g, '').replace(',', '.')); return n > 0 ? Math.round(n * 100) / 100 : ''; };
  const day = v => { const n = Math.round(Number(v)); return n >= 1 && n <= 31 ? n : ''; };
  const income = (p.income || []).map(x => ({ name: clean(x.name), amount: num(x.amount), day: day(x.day), main: !!x.main }))
    .filter(x => x.name).slice(0, 8);
  if (income.length && !income.some(x => x.main)) income[0].main = true;
  const regular = (p.regular || []).map(x => ({ name: clean(x.name), amount: num(x.amount), day: day(x.day),
    half: !!x.half, debt: !!x.debt, sub: !!x.sub })).filter(x => x.name);
  const variable = (p.variable || []).map(x => ({ name: clean(x.name), budget: num(x.budget), need: !!x.need }))
    .filter(x => x.name).slice(0, 12);
  const subs = (p.subs || []).map(x => ({ name: clean(x.name), amount: num(x.amount) })).filter(x => x.name).slice(0, 12);
  const rule = RULES.some(r => r.name === p.rule) ? p.rule : '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prof = Object.assign({}, profile_(), {
    name: name, gender: gender, age: age, asOf: todayIso_(ss),
    income: income, regular: regular, variable: variable, subs: subs, rule: rule, floating: !!(p && p.floating),
  });
  PropertiesService.getDocumentProperties().setProperty(PROP_PROFILE, JSON.stringify(prof));
  PROFILE_ = prof;
  applyProfileTheme_();

  let result = 'saved';
  if (mode === 'build' || !ss.getSheetByName(SHEETS.start)) {
    buildTracker({ keep: true });
    result = 'built';
  } else {
    writeProfileCells_(ss);
    const start = ss.getSheetByName(SHEETS.start);
    if (start) writeStartProfile_(start);
  }
  if (apply) applyQuizSettings_(ss, prof);
  return result;
}

/** Відповіді квізу → блоки «Налаштувань». */
function applyQuizSettings_(ss, p) {
  if (!ss.getRangeByName(NR + 'Inc_Name')) return;
  const used = usedCategories_(ss);
  if (p.income && p.income.length) {
    mergeBlockRows_(ss, 'Inc', p.income.map(x => ({ Name: x.name, Main: x.main, Exp: x.amount, Day: x.day })), used, true);
  }
  const reg = (p.regular || []).filter(x => !x.debt && !x.sub);
  const debts = (p.regular || []).filter(x => x.debt);
  const subs = (p.subs || []).concat((p.regular || []).filter(x => x.sub && !(p.subs || []).length));
  if (reg.length) {
    mergeBlockRows_(ss, 'Reg', reg.map(x => ({
      Name: x.name, Total: x.amount, Share: x.half ? 0.5 : 1, Freq: REF.freq[0][0], Day: x.day, Group: REF.groups[0],
    })), used);
  }
  if (debts.length) mergeBlockRows_(ss, 'Debt', debts.map(x => ({ Name: x.name, Pay: x.amount, Freq: REF.freq[0][0] })), used);
  if (subs.length && ss.getRangeByName(NR + 'Subs_Name')) {
    mergeBlockRows_(ss, 'Subs', subs.map(x => ({ Name: x.name, Total: x.amount, Freq: REF.freq[0][0], Day: x.day, Use: true })), used);
  }
  if (p.variable && p.variable.length) {
    mergeBlockRows_(ss, 'Var', p.variable.map(x => ({ Name: x.name, Budget: x.budget, Group: x.need ? REF.groups[0] : REF.groups[1] })), used);
  }
  if (p.rule && ss.getRangeByName(NR + 'Rule')) applyRule_(ss, p.rule);
  SpreadsheetApp.flush();
  syncCategoriesFromOps_(ss);
}

/** Усі категорії, що вже трапляються в «Операціях» (нижній регістр). */
function usedCategories_(ss) {
  const set = {};
  try { opsRows_(ss).forEach(o => { if (o.category) set[String(o.category).toLowerCase()] = true; }); } catch (e) { /* немає операцій */ }
  return set;
}

/**
 * Обʼєднує нові рядки з наявними в блоці «Налаштувань»:
 * однакова назва — оновлюємо заповнені поля; нова — додаємо; старі рядки лишаються,
 * якщо в них є суми/дати або категорія вже трапляється в «Операціях». Порожні «шаблонні» — прибираються.
 * singleMain — для блоку доходу: галочка «Основне» лише в одного.
 */
function mergeBlockRows_(ss, key, rows, used, singleMain) {
  const block = SETTINGS_BLOCKS.filter(b => b.key === key)[0];
  if (!block || !ss.getRangeByName(NR + key + '_Name')) return;
  const inputs = block.cols.filter(c => c.t !== 'f');
  const data = {};
  inputs.forEach(c => { data[c.key] = ss.getRangeByName(NR + key + '_' + c.key).getValues().map(v => v[0]); });
  const isData = t => ['money', 'int', 'day', 'date'].indexOf(t) !== -1;
  let cur = [];
  for (let i = 0; i < block.rows; i++) {
    const name = String(data.Name[i] || '').trim();
    if (!name) continue;
    const o = {};
    inputs.forEach(c => { o[c.key] = data[c.key][i]; });
    const hasData = inputs.some(c => isData(c.t) && o[c.key] !== '' && o[c.key] !== null);
    if (hasData || used[name.toLowerCase()]) cur.push(o);
  }
  rows.forEach(r => {
    const hit = cur.filter(o => String(o.Name).toLowerCase() === String(r.Name).toLowerCase())[0];
    const clean = {};
    Object.keys(r).forEach(k => { if (r[k] !== undefined && r[k] !== '') clean[k] = r[k]; });
    if (hit) Object.assign(hit, clean); else cur.push(Object.assign({}, r));
  });
  if (singleMain && rows.some(r => r.Main)) {
    const mainName = String(rows.filter(r => r.Main)[0].Name).toLowerCase();
    cur.forEach(o => { o.Main = String(o.Name).toLowerCase() === mainName; });
  }
  setBlockRows_(ss, key, cur.slice(0, block.rows));
}

/**
 * Кожна категорія з «Операцій», якої немає в «Налаштуваннях», додається у відповідний блок
 * (без сум), щоб операції не ставали «невідомими». Повертає список доданих.
 */
function syncCategoriesFromOps_(ss) {
  if (!ss.getRangeByName(NR + 'Inc_Name')) return [];
  const names = key => {
    const r = ss.getRangeByName(NR + key + '_Name');
    return r ? r.getValues().map(v => String(v[0]).trim().toLowerCase()).filter(String) : [];
  };
  const known = {};
  known[TYPES.income] = names('Inc');
  known[TYPES.expense] = names('Reg').concat(names('Var'), names('Subs'));
  known[TYPES.debt] = names('Debt');
  known[TYPES.saving] = names('Sav');
  const target = {};
  target[TYPES.income] = 'Inc'; target[TYPES.expense] = 'Var'; target[TYPES.debt] = 'Debt'; target[TYPES.saving] = 'Sav';
  const added = [], noRoom = [];
  let ops = [];
  try { ops = opsRows_(ss); } catch (e) { ops = []; }
  ops.forEach(o => {
    const cat = String(o.category || '').trim();
    if (!cat || !known[o.type]) return;
    if (known[o.type].indexOf(cat.toLowerCase()) !== -1) return;
    const key = target[o.type];
    const row = key === 'Var' ? { Name: cat, Group: REF.groups[1] } : key === 'Debt' ? { Name: cat, Freq: REF.freq[0][0] } : { Name: cat };
    if (appendBlockRow_(ss, key, row)) { added.push(cat); known[o.type].push(cat.toLowerCase()); }
    else if (noRoom.indexOf(cat) === -1) noRoom.push(cat);
  });
  if (added.length) SpreadsheetApp.flush();
  syncCategoriesFromOps_.noRoom = noRoom;
  return added;
}

/** Меню: повернути в «Налаштування» категорії, що трапляються в «Операціях». */
function repairCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const added = syncCategoriesFromOps_(ss);
  const noRoom = syncCategoriesFromOps_.noRoom || [];
  ss.toast(added.length
    ? 'Повернуто категорій: ' + added.length + ' (' + added.join(', ') + '). Можеш поставити їм бюджети в «Налаштуваннях».'
    : 'Усі категорії з «Операцій» уже є в «Налаштуваннях» ✓', menuTitle_(), 10);
  if (noRoom.length) ss.toast('Не вмістилися (блок заповнений): ' + noRoom.join(', ') + '. Обʼєднай схожі категорії.', menuTitle_(), 10);
  return added;
}

/** Чи є в трекері вже щось своє (тоді квіз питає, чи замінювати налаштування). */
function trackerHasData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const any = key => {
    const r = ss.getRangeByName(NR + key);
    return r ? r.getValues().some(v => v[0] !== '' && v[0] !== null) : false;
  };
  const opsSh = ss.getSheetByName(SHEETS.ops);
  return any('Inc_Exp') || any('Reg_Total') || any('Subs_Total') || any('Var_Budget') || any('Debt_Pay') ||
    (opsSh ? firstFreeOpsRow_(opsSh) > 2 : false);
}

/**
 * Властивості документа копіюються разом зі скриптом. Якщо файл — копія (інший ID),
 * скидаємо профіль, форму й демо-позначки, щоб нова людина почала з квізу.
 */
function ensureOwnDocument_() {
  const props = PropertiesService.getDocumentProperties();
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  const stored = props.getProperty('BT_DOC_ID');
  if (stored && stored !== id) {
    [PROP_PROFILE, FORM.propId, FORM.propUrl, FORM.propSs, PROP_DEMO].forEach(k => props.deleteProperty(k));
    const n = Number(props.getProperty(PROP_SNAP + 'N') || 0);
    for (let i = 0; i < n; i++) props.deleteProperty(PROP_SNAP + i);
    props.deleteProperty(PROP_SNAP + 'N');
    PROFILE_ = null;
  }
  if (stored !== id) props.setProperty('BT_DOC_ID', id);
}

/** Меню: чиста копія трекера для поширення (без операцій, сум і профілю). */
function makeShareCopy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const copy = ss.copy(ss.getName() + ' — чистий шаблон');
  const c = SpreadsheetApp.openById(copy.getId());
  cleanTrackerCopy_(c);
  const url = 'https://docs.google.com/spreadsheets/d/' + copy.getId() + '/edit';
  const shareUrl = 'https://docs.google.com/spreadsheets/d/' + copy.getId() + '/copy';
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial;line-height:1.5;color:#2E3A46">' +
    '<p><b>Готово!</b> Створено чисту копію без твоїх операцій, сум і профілю.</p>' +
    '<p>1. <a href="' + url + '" target="_blank">Відкрий копію</a> → «Поділитися» → «Усі, хто має посилання» → «Читач».</p>' +
    '<p>2. Роздавай це посилання — кожен отримає свою копію:</p>' +
    '<p style="background:#F3F6F9;padding:8px;border-radius:8px;word-break:break-all">' + shareUrl + '</p>' +
    '<p style="color:#8A9BAA">Свої дані людина внесе сама або через «Імпорт з іншої таблиці».</p></div>')
    .setWidth(460).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '📤 Копія для поширення');
  return shareUrl;
}

/** Очищає копію трекера: операції, суми в налаштуваннях, профіль, форма, звіт перевірки. */
function cleanTrackerCopy_(c) {
  c.getSheets().forEach(s => {
    if (s.getFormUrl && s.getFormUrl()) { try { FormApp.openByUrl(s.getFormUrl()).removeDestination(); } catch (e) { /* ок */ } }
    if (s.getName().indexOf(FORM.responsesSheet) === 0) c.deleteSheet(s);
  });
  const ops = c.getSheetByName(SHEETS.ops);
  if (ops) {
    const blank = new Array(OPS_HEADERS.length).fill('');
    writeOpsRows_(ops, 2, Array.from({ length: DATA_ROWS - 1 }, () => blank.slice()));
  }
  if (c.getRangeByName(NR + 'Inc_Name')) {
    SETTINGS_BLOCKS.forEach(b => setBlockRows_(c, b.key, SEEDS[b.key] || []));
    const t = todayParts_(c);
    setKV_(c, { Year: t.y, Start: isoDate_(t.y, t.m, 1), Calc: '', Cycle: REF.cycle[0], Warn: 0.8, Over: 1 });
    applyRule_(c, RULES[0].name);
  }
  const lists = c.getSheetByName(SHEETS.lists);
  if (lists) PROFILE_FIELDS.forEach((f, i) => lists.getRange(2 + i, 19).setValue(''));
  const start = c.getSheetByName(SHEETS.start);
  if (start) {
    start.getRange(3, 2).setValue('Вписуєш суму — таблиця сама рахує, скільки можна витратити до наступної виплати.');
    start.getRange(START.formRow, 3).setValue('Ще не створена: меню «Трекер» → «Створити / оновити форму для телефону».');
    start.getRange(START.formRow + 1, 3).setValue('');
    start.getRange(START.checkRow, 3, START.checkLines, 1).clearContent();
    start.getRange(START.checkRow, 3).setValue('Меню «Трекер» → «Перевірити таблицю».');
    start.getRange(START.profileRow, 3).setValue('Ще не заповнений — меню «Трекер» → «👋 Мій профіль (квіз)».');
  }
  const dash = c.getSheetByName(SHEETS.dash);
  if (dash) {
    const cell = dash.getRange(DASH.sentenceRow, 2);
    const n = nameComma_();
    if (n) cell.setFormula(cell.getFormula().split('"' + n.replace(/""/g, '"')).join('"'));
  }
  SpreadsheetApp.flush();
}

/** Профіль у прихованих «_Списках» — для майбутнього пенсійного калькулятора. */
function writeProfileCells_(ss) {
  const sh = ss.getSheetByName(SHEETS.lists);
  if (!sh) return;
  const p = profile_();
  const col = 18; // R:S
  sh.getRange(1, col, 1, 2).setValues([['Профіль', '']]).setFontWeight('bold');
  const vals = {
    P_Name: p.name || '', P_Gender: p.gender === 'm' ? 'Чоловік' : (p.gender === 'f' ? 'Жінка' : ''),
    P_Age: p.age || '', P_AsOf: p.asOf || '', P_RetireAge: p.retireAge || '',
    P_RetireIncome: p.retireIncome || '', P_RetireCurrency: p.retireCurrency || '',
  };
  PROFILE_FIELDS.forEach((f, i) => {
    sh.getRange(2 + i, col).setValue(f[1]);
    const cell = sh.getRange(2 + i, col + 1).setValue(vals[f[0]]);
    ss.setNamedRange(NR + f[0], cell);
  });
}

/** «Що де»: посилання на кожну вкладку і навіщо вона. */
function writeStartMap_(sh) {
  const ss = sh.getParent();
  const r0 = START.profileRow + 3;
  writeBlockTitle_(sh, r0, 'ЩО ДЕ', 2);
  const map = [
    [SHEETS.dash, 'головна цифра «можу витратити до виплати», ліміт на день, бюджети й підказки'],
    [SHEETS.ops, 'сюди вписуєш кожну покупку й надходження (або через форму з телефона)'],
    [SHEETS.settings, 'дохід і дні виплат, регулярні платежі, бюджети, борги, цілі, рахунки'],
    [SHEETS.year, 'усі категорії по місяцях, графік і теплова карта'],
    [SHEETS.goals, 'скільки накопичено і скільки відкладати до бажаної дати'],
    [SHEETS.debts, 'порядок погашення (сніжна куля) і прогноз закриття'],
    [SHEETS.subs, 'скільки коштують підписки на рік і що можна скасувати'],
    [SHEETS.cal, 'витрати по днях місяця, дні без витрат'],
  ];
  map.forEach((m, i) => {
    const target = ss.getSheetByName(m[0]);
    const cell = sh.getRange(r0 + 1 + i, 2, 1, 2);
    if (target) sh.getRange(r0 + 1 + i, 2).setFormula(X_('=HYPERLINK("#gid=' + target.getSheetId() + '","' + m[0] + ' ↗")'));
    else sh.getRange(r0 + 1 + i, 2).setValue(m[0]);
    sh.getRange(r0 + 1 + i, 3).setValue(m[1]).setFontColor(THEME.muted).setWrap(true);
    cell.setVerticalAlignment('top');
  });
  sh.getRange(r0 + 1, 2, map.length, 1).setFontWeight('bold').setFontColor(THEME.primary);
  sh.getRange(r0 + map.length + 1, 3).setValue('Покрокова інструкція: меню «Трекер» → «📖 Як користуватись».')
    .setFontColor(THEME.muted).setFontStyle('italic');
}

function writeStartProfile_(sh) {
  const p = profile_();
  const r = START.profileRow;
  sh.getRange(2, 2).setValue(spaced_('ТРЕКЕР БЮДЖЕТУ'));
  sh.getRange(3, 2).setValue((p.name ? 'Привіт, ' + p.name + '! ' : '') +
    'Вписуєш суму — таблиця сама рахує, скільки можна витратити до наступної виплати.');
  writeBlockTitle_(sh, r - 1, 'ПРОФІЛЬ', 2);
  const age = currentAge_();
  sh.getRange(r, 3).setValue(hasProfile_()
    ? p.name + ' · ' + (p.gender === 'm' ? 'чоловік' : 'жінка') + (age ? ' · ' + age + ' ' + yearsWord_(age) : '')
    : 'Ще не заповнений.').setFontWeight('bold').setFontColor(THEME.headerTxt);
  sh.getRange(r + 1, 3).setValue('Змінити: меню «Трекер» → «👋 Мій профіль (квіз)».')
    .setFontColor(THEME.muted).setFontStyle('italic');
}

/** HTML квізу: палітри, поточні відповіді й чи є вже дані в трекері. */
function quizHtml_() {
  const p = profile_();
  const data = {
    themes: THEMES,
    rules: RULES.map(r => ({ name: r.name, title: r.title, about: r.about, tag: r.tag || '', buckets: r.buckets.map(b => [b[0], b[1]]) })),
    hasData: (() => { try { return trackerHasData_(); } catch (e) { return false; } })(),
    profile: {
      name: p.name || '', gender: p.gender || '', age: p.age || 30, floating: !!p.floating,
      income: p.income || [], regular: (p.regular || []).map(x => Object.assign({ key: 'r' + Math.random(), e: '📌' }, x)),
      variable: p.variable || [], subs: p.subs || [], rule: p.rule || '',
    },
  };
  return QUIZ_HTML.replace('__DATA__', () => JSON.stringify(data).replace(/</g, '\\u003c'));
}

/** HTML бокової панелі-інструкції. */
function guideHtml_() {
  const p = profile_();
  const data = {
    name: p.name || '', gender: p.gender || '', theme: THEME,
    formUrl: PropertiesService.getDocumentProperties().getProperty(FORM.propUrl) || '',
  };
  return GUIDE_HTML
    .replace('__QRLIB__', () => QR_LIB)
    .replace('__DATA__', () => JSON.stringify(data).replace(/</g, '\\u003c'));
}

// HTML і бібліотека — як «сирі» рядки: зворотні скісні риски не інтерпретуються.
const QUIZ_HTML = String.raw`<!DOCTYPE html>
<html lang="uk"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  :root { --p:#6F8FB3; --a:#A9C6E0; --bg:#EEF3F8; --t:#34506B; --ink:#2E3A46; --muted:#8A9BAA; --err:#D9534F; --line:#E7EBF0; }
  * { box-sizing:border-box; }
  html,body { margin:0; height:100%; font-family:Montserrat, Arial, sans-serif; color:var(--ink); background:#fff; }
  .wrap { display:flex; flex-direction:column; height:100%; padding:18px 24px 18px; }
  .top { display:flex; align-items:center; gap:12px; min-height:22px; margin-bottom:6px; }
  .steps { font-size:13px; font-weight:700; color:var(--muted); white-space:nowrap; }
  .bar { flex:1; height:6px; background:var(--bg); border-radius:99px; overflow:hidden; }
  .bar i { display:block; height:100%; width:0; background:linear-gradient(90deg,var(--p),var(--a)); border-radius:99px; transition:width .45s ease; }
  .screen { flex:1; display:none; flex-direction:column; min-height:0; animation:in .3s ease; }
  .screen.on { display:flex; }
  .screen.mid { justify-content:center; }
  @keyframes in { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
  .body { flex:1; overflow:auto; min-height:0; padding:4px 2px 6px; }
  .emoji { font-size:46px; line-height:1; margin-bottom:10px; }
  h1 { font-size:24px; font-weight:800; margin:0 0 8px; color:var(--t); line-height:1.2; }
  p.sub { font-size:14.5px; color:var(--muted); margin:0 0 16px; line-height:1.45; }
  input[type=text], input.money, select { font:600 16px Montserrat, Arial; padding:12px 14px; border:2px solid var(--line); border-radius:14px; outline:none; color:var(--ink); background:#fff; }
  input:focus, select:focus { border-color:var(--p); box-shadow:0 0 0 3px var(--bg); }
  #name { width:100%; font-size:22px; padding:16px 18px; }
  .err { color:var(--err); font-size:13px; font-weight:600; min-height:18px; margin-top:6px; }
  .cards { display:grid; grid-template-columns:1fr 1fr; gap:14px; }
  .card { border:2px solid var(--line); border-radius:20px; padding:18px 12px 14px; text-align:center; cursor:pointer; background:#fff; user-select:none; transition:transform .15s, border-color .2s; }
  .card:hover { transform:translateY(-2px); }
  .card.sel { border-color:var(--p); box-shadow:0 0 0 4px var(--bg); }
  .card .e { font-size:42px; } .card b { display:block; font-size:17px; margin:8px 0 4px; } .card small { color:var(--muted); font-size:12px; }
  .sw { display:flex; justify-content:center; gap:5px; margin-top:8px; } .sw span { width:16px; height:16px; border-radius:50%; }
  .age { display:flex; align-items:center; justify-content:center; gap:18px; margin:4px 0 12px; }
  .round { width:52px; height:52px; border-radius:50%; border:2px solid var(--bg); background:#fff; font:700 26px Montserrat, Arial; color:var(--p); cursor:pointer; }
  .num { font:800 60px Montserrat, Arial; color:var(--t); width:130px; text-align:center; border:none; outline:none; background:none; }
  .unit { text-align:center; color:var(--muted); font-weight:600; margin-top:-12px; margin-bottom:10px; }
  input[type=range] { width:100%; accent-color:var(--p); }
  /* дохід */
  .src { border:2px solid var(--line); border-radius:16px; padding:10px 12px; margin-bottom:10px; }
  .src.main { border-color:var(--p); }
  .src .r1, .src .r2 { display:flex; gap:8px; align-items:center; }
  .src .r2 { margin-top:8px; }
  .src .nm { flex:1; min-width:0; }
  .src .money { width:150px; }
  .src select { width:150px; }
  .lbl { font-size:12px; color:var(--muted); font-weight:600; }
  .icon { border:none; background:none; font-size:20px; cursor:pointer; padding:4px 6px; color:var(--muted); border-radius:8px; }
  .icon.star.on { color:#F2B300; }
  .badge { font-size:11px; font-weight:700; color:#fff; background:var(--p); border-radius:8px; padding:2px 7px; }
  .chips { display:flex; flex-wrap:wrap; gap:8px; margin:6px 0 12px; }
  .chip { border:2px solid var(--line); background:#fff; border-radius:99px; padding:8px 12px; font:600 13.5px Montserrat, Arial; cursor:pointer; color:var(--ink); user-select:none; }
  .chip.sel { border-color:var(--p); background:var(--bg); color:var(--t); }
  .chip.add { border-style:dashed; color:var(--p); }
  .check { display:flex; gap:8px; align-items:center; font-size:13.5px; color:var(--ink); margin:4px 0 10px; cursor:pointer; }
  .check input { width:18px; height:18px; accent-color:var(--p); }
  .det { border-radius:14px; background:var(--bg); padding:10px 12px; margin-bottom:8px; }
  .det .h { font-weight:700; font-size:14px; margin-bottom:6px; display:flex; justify-content:space-between; align-items:center; }
  .det .r { display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
  .det .money { width:150px; padding:9px 12px; font-size:15px; }
  .det select { padding:9px 10px; font-size:14px; }
  .det .nm { flex:1; min-width:120px; padding:9px 12px; font-size:15px; }
  .hint { font-size:12px; color:var(--muted); margin-top:6px; line-height:1.4; }
  /* підсумок */
  .rule { border:2px solid var(--line); border-radius:16px; padding:12px 14px; margin-bottom:10px; cursor:pointer; }
  .rule.sel { border-color:var(--p); box-shadow:0 0 0 3px var(--bg); }
  .rule .rt { display:flex; justify-content:space-between; align-items:center; font-weight:800; color:var(--t); font-size:15px; }
  .rule .tag { font-size:11px; font-weight:700; color:#fff; background:var(--p); border-radius:8px; padding:2px 7px; }
  .rule .ra { font-size:13px; color:var(--ink); margin:4px 0 6px; line-height:1.45; }
  .rule .rx { display:flex; flex-wrap:wrap; gap:6px; }
  .rule .rx span { font-size:12px; background:var(--bg); color:var(--t); border-radius:8px; padding:3px 8px; font-weight:600; }
  .subs { display:flex; flex-wrap:wrap; gap:6px; margin:6px 0 8px; }
  .subs .chip { padding:6px 10px; font-size:12.5px; }
  .sr { display:flex; gap:8px; align-items:center; margin-bottom:6px; }
  .sr b { flex:1; font-size:14px; }
  .blk { border-radius:16px; background:var(--bg); padding:10px 14px; margin-bottom:10px; cursor:pointer; }
  .blk .bh { display:flex; justify-content:space-between; font-weight:700; color:var(--t); font-size:14px; }
  .blk .bh em { font-style:normal; color:var(--p); font-size:12.5px; }
  .blk .bb { font-size:13.5px; margin-top:4px; line-height:1.5; color:var(--ink); }
  .foot { display:flex; gap:12px; align-items:center; margin-top:12px; }
  .btn { flex:1; border:none; border-radius:16px; padding:15px 18px; font:700 16px Montserrat, Arial; cursor:pointer; }
  .btn.main { background:var(--p); color:#fff; } .btn.main:disabled { opacity:.4; cursor:default; }
  .btn.ghost { flex:0 0 auto; background:none; color:var(--muted); padding:15px 8px; }
  .btn.soft { background:var(--bg); color:var(--t); }
  .link { background:none; border:none; color:var(--muted); font:600 13px Montserrat, Arial; text-decoration:underline; cursor:pointer; margin-top:10px; align-self:center; }
  .spin { width:56px; height:56px; border-radius:50%; border:6px solid var(--bg); border-top-color:var(--p); animation:r 1s linear infinite; margin-bottom:20px; }
  @keyframes r { to { transform:rotate(360deg); } }
  .center { align-items:center; text-align:center; }
</style></head>
<body><div class="wrap">
  <div class="top" id="top"><div class="steps" id="steps"></div><div class="bar"><i id="bar"></i></div></div>

  <section class="screen mid on" id="s0">
    <div class="emoji">👋</div>
    <h1>Привіт! Давай познайомимось</h1>
    <p class="sub">7 коротких питань — і я одразу заповню перший варіант трекера: кольори, звертання, твій дохід і основні витрати. Відповіді можна змінити будь-коли.</p>
    <div class="foot"><button class="btn main" onclick="go(1)">Почати</button></div>
  </section>

  <section class="screen mid" id="s1">
    <div class="emoji">✍️</div>
    <h1>Як тебе звати?</h1>
    <p class="sub">Саме так я до тебе звертатимусь у таблиці.</p>
    <input type="text" id="name" maxlength="30" placeholder="Твоє ім'я" autocomplete="given-name">
    <div class="err" id="e1"></div>
    <div class="foot"><button class="btn ghost" onclick="go(0)">← Назад</button><button class="btn main" onclick="next1()">Далі</button></div>
  </section>

  <section class="screen mid" id="s2">
    <div class="emoji">🎨</div>
    <h1 id="h2">Як до тебе звертатись?</h1>
    <p class="sub">Від цього залежать кольори таблиці й те, як я пишу підказки.</p>
    <div class="cards">
      <div class="card" id="cf" onclick="pick('f')"><div class="e">👩</div><b>Жінка</b><small>лілово-рожева тема</small><div class="sw" id="swf"></div></div>
      <div class="card" id="cm" onclick="pick('m')"><div class="e">👨</div><b>Чоловік</b><small>блакитна тема</small><div class="sw" id="swm"></div></div>
    </div>
    <div class="err" id="e2"></div>
    <div class="foot"><button class="btn ghost" onclick="go(1)">← Назад</button><button class="btn main" id="b2" onclick="next2()" disabled>Далі</button></div>
  </section>

  <section class="screen mid" id="s3">
    <div class="emoji">🎂</div>
    <h1>Скільки тобі років?</h1>
    <p class="sub">Знадобиться пізніше — щоб порахувати, скільки часу є до пенсії і скільки відкладати.</p>
    <div class="age"><button class="round" onclick="bump(-1)">−</button><input class="num" id="age" inputmode="numeric" maxlength="3"><button class="round" onclick="bump(1)">+</button></div>
    <div class="unit" id="unit">років</div>
    <input type="range" id="range" min="14" max="100">
    <div class="err" id="e3"></div>
    <div class="foot"><button class="btn ghost" onclick="go(2)">← Назад</button><button class="btn main" onclick="next3()">Далі</button></div>
  </section>

  <section class="screen" id="s4">
    <div class="body">
      <div class="emoji">💰</div>
      <h1>Звідки приходять гроші?</h1>
      <p class="sub">Вкажи джерела доходу, суму й число місяця, коли вони приходять. ★ — основне: від його дня я рахую «до наступної виплати».</p>
      <div id="srcs"></div>
      <div class="chips" id="srcAdd"></div>
      <label class="check"><input type="checkbox" id="floating"> Сума щоразу різна — впишу фактичні суми, коли прийдуть</label>
      <div class="err" id="e4"></div>
    </div>
    <div class="foot"><button class="btn ghost" onclick="go(3)">← Назад</button><button class="btn main" onclick="next4()">Далі</button></div>
  </section>

  <section class="screen" id="s5">
    <div class="body">
      <div class="emoji">🏠</div>
      <h1>Обовʼязкові платежі</h1>
      <p class="sub">Що ти платиш щомісяця? Обери — і впиши суму та число. Якщо не знаєш точно — пропусти, можна дописати потім.</p>
      <div class="chips" id="regChips"></div>
      <div id="regDet"></div>
    </div>
    <div class="foot"><button class="btn ghost" onclick="go(4)">← Назад</button><button class="btn main" onclick="next5()">Далі</button></div>
  </section>

  <section class="screen" id="s6">
    <div class="body">
      <div class="emoji">🛒</div>
      <h1>На що йдуть гроші щодня?</h1>
      <p class="sub">Обери свої основні категорії витрат. Бюджет на місяць — за бажанням: тоді трекер покаже, скільки ще лишилось.</p>
      <div class="chips" id="varChips"></div>
      <div id="varDet"></div>
    </div>
    <div class="foot"><button class="btn ghost" onclick="go(5)">← Назад</button><button class="btn main" onclick="next6()">Далі</button></div>
  </section>

  <section class="screen" id="s7">
    <div class="body">
      <div class="emoji">🧮</div>
      <h1>Як хочеш розподіляти гроші?</h1>
      <p class="sub">Обери правило — трекер покаже, чи вкладаєшся в нього. Змінити можна будь-коли в «Налаштуваннях».</p>
      <div id="rules"></div>
    </div>
    <div class="foot"><button class="btn ghost" onclick="go(6)">← Назад</button><button class="btn main" onclick="next7()">Далі</button></div>
  </section>

  <section class="screen" id="s8">
    <div class="body">
      <div class="emoji">🎉</div>
      <h1 id="h8">Готово!</h1>
      <p class="sub" id="p8">Перевір — натисни на блок, щоб змінити.</p>
      <div id="sum"></div>
    </div>
    <div id="replaceBox" style="display:none">
      <label class="check"><input type="checkbox" id="replace" checked> Записати ці дохід і витрати в «Налаштування» (замінить те, що там зараз)</label>
    </div>
    <div class="foot"><button class="btn main" onclick="send('build')">Створити мій трекер</button></div>
    <button class="link" onclick="send('save')">Лише зберегти, без перебудови</button>
  </section>

  <section class="screen mid center" id="s9">
    <div class="spin"></div>
    <h1 id="h9">Будую твій трекер…</h1>
    <p class="sub" id="p9">Це до хвилини — можеш поки зробити чай ☕</p>
  </section>

  <section class="screen mid center" id="s10">
    <div class="emoji" id="em10">✅</div>
    <h1 id="h10">Готово!</h1>
    <p class="sub" id="p10"></p>
    <div class="foot" style="width:100%"><button class="btn main" onclick="openGuide()">Як користуватись →</button></div>
    <button class="link" id="closeLink" onclick="closeMe()">Закрити</button>
  </section>
</div>
<script>
  const D = __DATA__;
  const P = D.profile;
  const st = {
    name: P.name || '', gender: P.gender || '', age: P.age || 30,
    income: (P.income && P.income.length) ? P.income : [{ name: 'Зарплата', amount: '', day: 1, main: true }],
    floating: !!P.floating,
    regular: P.regular || [],
    variable: (P.variable && P.variable.length) ? P.variable : [{ name: 'Їжа', budget: '', need: true }, { name: 'Транспорт', budget: '', need: true }],
    subs: P.subs || [],
    rule: P.rule || (D.rules[0] && D.rules[0].name),
  };
  const POPULAR = ['Netflix', 'YouTube Premium', 'Spotify', 'Apple Music', 'iCloud+', 'Google One', 'Megogo', 'Sweet.tv', 'ChatGPT', 'Claude', 'Telegram Premium'];
  const REG = [
    { key: 'rent', e: '🏠', name: 'Оренда / іпотека', half: true },
    { key: 'util', e: '💡', name: 'Комуналка' },
    { key: 'net', e: '📶', name: 'Інтернет і звʼязок' },
    { key: 'credit', e: '💳', name: 'Кредит / розстрочка', debt: true },
    { key: 'subs', e: '📺', name: 'Підписки', sub: true },
    { key: 'kids', e: '🧒', name: 'Садочок / школа' },
    { key: 'ins', e: '🛡️', name: 'Страховка' },
    { key: 'sport', e: '🏋️', name: 'Спортзал' },
  ];
  const VARS = [
    ['🍎', 'Їжа', 1], ['🚌', 'Транспорт', 1], ['☕', 'Кафе й доставка', 0], ['💅', 'Краса й здоровʼя', 1],
    ['👗', 'Одяг', 0], ['🎬', 'Розваги', 0], ['🎁', 'Подарунки', 0], ['🧸', 'Діти', 1], ['🐾', 'Тварини', 1],
    ['⚽', 'Спорт і хобі', 0], ['🚗', 'Авто', 1], ['💊', 'Аптека', 1],
  ];
  let cur = 0;
  const TOTAL = 7;
  const $ = id => document.getElementById(id);
  const G = (f, m) => st.gender === 'm' ? m : f;
  const yw = n => { const a = n % 10, b = n % 100; return a === 1 && b !== 11 ? 'рік' : (a >= 2 && a <= 4 && (b < 12 || b > 14) ? 'роки' : 'років'); };
  const esc = s => String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const money = v => { const n = Number(String(v || '').replace(/[^0-9.,]/g, '').replace(',', '.')); return n > 0 ? n : ''; };
  const fmt = n => n ? Math.round(n).toString().replace(/(\d)(?=(\d{3})+$)/g, '$1 ') + ' грн' : '';
  function dayOptions(sel, allowIrr) {
    let h = allowIrr ? '<option value=""' + (sel ? '' : ' selected') + '>нерегулярно</option>' : '<option value="">число —</option>';
    for (let d = 1; d <= 31; d++) h += '<option value="' + d + '"' + (Number(sel) === d ? ' selected' : '') + '>' + d + '-го</option>';
    return h;
  }

  function paint() {
    const t = st.gender ? D.themes[st.gender] : null;
    const r = document.documentElement.style;
    if (t) { r.setProperty('--p', t.primary); r.setProperty('--a', t.accent); r.setProperty('--bg', t.headerBg); r.setProperty('--t', t.headerTxt); }
    $('cf').classList.toggle('sel', st.gender === 'f');
    $('cm').classList.toggle('sel', st.gender === 'm');
    $('b2').disabled = !st.gender;
  }
  function swatch(id, t) { $(id).innerHTML = [t.primary, t.accent, t.headerBg].map(c => '<span style="background:' + c + '"></span>').join(''); }

  function go(n) {
    document.querySelectorAll('.screen.on').forEach(e => e.classList.remove('on'));
    cur = n; $('s' + n).classList.add('on');
    const q = Math.min(Math.max(n, 0), TOTAL);
    $('top').style.visibility = n >= 1 && n <= TOTAL ? 'visible' : 'hidden';
    $('steps').textContent = 'Крок ' + q + ' / ' + TOTAL;
    $('bar').style.width = (n > TOTAL ? 100 : q / TOTAL * 100) + '%';
    if (n === 1) setTimeout(() => $('name').focus(), 60);
    if (n === 2) $('h2').textContent = st.name + ', як до тебе звертатись?';
    if (n === 3) setAge(st.age);
    if (n === 4) renderSrc();
    if (n === 5) renderReg();
    if (n === 6) renderVar();
    if (n === 7) renderRules();
    if (n === 8) summary();
  }

  // 1–3
  function next1() {
    const v = $('name').value.trim().replace(/ +/g, ' ');
    if (v.length < 2) { $('e1').textContent = 'Напиши імʼя — хоча б 2 літери.'; return; }
    if (/[0-9<>"{}]/.test(v)) { $('e1').textContent = 'Лише літери, пробіл, дефіс чи апостроф.'; return; }
    $('e1').textContent = ''; st.name = v; go(2);
  }
  function pick(g) { st.gender = g; $('e2').textContent = ''; paint(); }
  function next2() { if (!st.gender) { $('e2').textContent = 'Обери один варіант.'; return; } go(3); }
  function setAge(n) { n = Math.max(14, Math.min(100, Math.round(Number(n) || 30))); st.age = n; $('age').value = n; $('range').value = n; $('unit').textContent = yw(n); }
  function bump(d) { setAge(st.age + d); }
  function next3() { const n = Number($('age').value); if (!(n >= 14 && n <= 100)) { $('e3').textContent = 'Вкажи вік від 14 до 100.'; return; } $('e3').textContent = ''; setAge(n); go(4); }

  // 4. Дохід
  function renderSrc() {
    $('srcs').innerHTML = st.income.map((s, i) =>
      '<div class="src' + (s.main ? ' main' : '') + '">' +
        '<div class="r1"><input type="text" class="nm" value="' + esc(s.name) + '" placeholder="Назва джерела" oninput="st.income[' + i + '].name=this.value">' +
        (s.main ? '<span class="badge">основне</span>' : '') +
        '<button class="icon star' + (s.main ? ' on' : '') + '" title="Зробити основним" onclick="mainSrc(' + i + ')">★</button>' +
        (st.income.length > 1 ? '<button class="icon" title="Прибрати" onclick="delSrc(' + i + ')">✕</button>' : '') + '</div>' +
        '<div class="r2"><input class="money" inputmode="decimal" placeholder="сума, грн" value="' + esc(s.amount) + '" oninput="st.income[' + i + '].amount=this.value">' +
        '<select onchange="st.income[' + i + '].day=this.value">' + dayOptions(s.day, true) + '</select></div>' +
      '</div>').join('');
    const used = st.income.map(s => s.name);
    $('srcAdd').innerHTML = ['Аванс', 'Фріланс', 'Підробіток', 'Інше']
      .filter(n => used.indexOf(n) === -1 || n === 'Інше')
      .map(n => '<span class="chip add" onclick="addSrc(&quot;' + n + '&quot;)">+ ' + n + '</span>').join('');
    $('floating').checked = st.floating;
  }
  function addSrc(n) { if (st.income.length >= 8) return; st.income.push({ name: n === 'Інше' ? '' : n, amount: '', day: n === 'Аванс' ? 15 : '', main: false }); renderSrc(); }
  function delSrc(i) { const wasMain = st.income[i].main; st.income.splice(i, 1); if (wasMain && st.income.length) st.income[0].main = true; renderSrc(); }
  function mainSrc(i) { st.income.forEach((s, k) => { s.main = k === i; }); renderSrc(); }
  function next4() {
    st.floating = $('floating').checked;
    st.income = st.income.filter(s => String(s.name).trim());
    if (!st.income.length) { st.income = [{ name: 'Зарплата', amount: '', day: 1, main: true }]; renderSrc(); $('e4').textContent = 'Додай хоча б одне джерело доходу.'; return; }
    if (!st.income.some(s => s.main)) st.income[0].main = true;
    const m = st.income.filter(s => s.main)[0];
    if (!st.floating && !money(m.amount)) { $('e4').textContent = 'Впиши суму основного доходу або познач «сума щоразу різна».'; renderSrc(); return; }
    $('e4').textContent = ''; go(5);
  }

  // 5. Обов'язкові
  function findReg(r) { return st.regular.filter(x => x.key === r.key)[0]; }
  function renderReg() {
    $('regChips').innerHTML = REG.map(r => '<span class="chip' + (findReg(r) ? ' sel' : '') + '" onclick="togReg(&quot;' + r.key + '&quot;)">' + r.e + ' ' + r.name + '</span>').join('') +
      '<span class="chip add" onclick="addOwnReg()">+ Своє</span>';
    $('regDet').innerHTML = st.regular.map((x, i) =>
      '<div class="det"><div class="h">' + (x.own ? '<input type="text" class="nm" placeholder="Назва платежу" value="' + esc(x.name) + '" oninput="st.regular[' + i + '].name=this.value">' : '<span>' + x.e + ' ' + esc(x.name) + '</span>') +
      '<button class="icon" onclick="delReg(' + i + ')">✕</button></div>' +
      (x.sub ? '' : '<div class="r"><input class="money" inputmode="decimal" placeholder="' + (x.debt ? 'платіж на місяць' : 'сума на місяць') + '" value="' + esc(x.amount) + '" oninput="st.regular[' + i + '].amount=this.value">' +
      '<select onchange="st.regular[' + i + '].day=this.value">' + dayOptions(x.day, false) + '</select></div>') +
      (x.key === 'rent' ? '<label class="check"><input type="checkbox"' + (x.half ? ' checked' : '') + ' onchange="st.regular[' + i + '].half=this.checked"> Ділю навпіл (моя частка 50%) — впиши повну суму</label>' : '') +
      (x.debt ? '<div class="hint">Піде в «Борги». Залишок боргу допишеш у «Налаштуваннях» — тоді покажу, коли закриєш.</div>' : '') +
      (x.sub ? renderSubs() : '') +
      '</div>').join('');
  }
  function renderSubs() {
    const has = n => st.subs.some(x => x.name === n);
    return '<div class="hint">Обери свої підписки й впиши суму за місяць:</div><div class="subs">' +
      POPULAR.map(n => '<span class="chip' + (has(n) ? ' sel' : '') + '" onclick="togSub(&quot;' + n + '&quot;)">' + n + '</span>').join('') +
      '<span class="chip add" onclick="addOwnSub()">+ Своя</span></div>' +
      st.subs.map((x, k) => '<div class="sr">' + (x.own ? '<input type="text" class="nm" placeholder="Назва" value="' + esc(x.name) + '" oninput="st.subs[' + k + '].name=this.value">' : '<b>' + esc(x.name) + '</b>') +
        '<input class="money" inputmode="decimal" placeholder="грн / міс" value="' + esc(x.amount) + '" oninput="st.subs[' + k + '].amount=this.value">' +
        '<button class="icon" onclick="delSub(' + k + ')">✕</button></div>').join('');
  }
  function togSub(n) { const k = st.subs.findIndex(x => x.name === n); if (k >= 0) st.subs.splice(k, 1); else st.subs.push({ name: n, amount: '' }); renderReg(); }
  function addOwnSub() { st.subs.push({ name: '', amount: '', own: true }); renderReg(); }
  function delSub(k) { st.subs.splice(k, 1); renderReg(); }
  function togReg(key) {
    const r = REG.filter(x => x.key === key)[0];
    const i = st.regular.findIndex(x => x.key === key);
    if (i >= 0) st.regular.splice(i, 1);
    else st.regular.push({ key: r.key, e: r.e, name: r.name, amount: '', day: '', half: false, debt: !!r.debt, sub: !!r.sub });
    renderReg();
  }
  function addOwnReg() { st.regular.push({ key: 'own' + Date.now(), e: '📌', name: '', own: true, amount: '', day: '', half: false, debt: false, sub: false }); renderReg(); }
  function delReg(i) { st.regular.splice(i, 1); renderReg(); }
  function next5() {
    st.regular = st.regular.filter(x => String(x.name).trim());
    st.subs = st.subs.filter(x => String(x.name).trim());
    if (!st.regular.some(x => x.sub)) st.subs = [];
    go(6);
  }

  // 6. Щоденні
  function findVar(n) { return st.variable.filter(x => x.name === n)[0]; }
  function renderVar() {
    $('varChips').innerHTML = VARS.map(v => '<span class="chip' + (findVar(v[1]) ? ' sel' : '') + '" onclick="togVar(&quot;' + v[1] + '&quot;)">' + v[0] + ' ' + v[1] + '</span>').join('') +
      '<span class="chip add" onclick="addOwnVar()">+ Своя</span>';
    $('varDet').innerHTML = st.variable.map((x, i) =>
      '<div class="det"><div class="r">' + (x.own ? '<input type="text" class="nm" placeholder="Назва категорії" value="' + esc(x.name) + '" oninput="st.variable[' + i + '].name=this.value">' : '<b style="flex:1">' + esc(x.name) + '</b>') +
      '<input class="money" inputmode="decimal" placeholder="бюджет / міс" value="' + esc(x.budget) + '" oninput="st.variable[' + i + '].budget=this.value">' +
      '<button class="icon" onclick="delVar(' + i + ')">✕</button></div></div>').join('');
  }
  function togVar(n) {
    const i = st.variable.findIndex(x => x.name === n);
    const v = VARS.filter(x => x[1] === n)[0];
    if (i >= 0) st.variable.splice(i, 1); else if (st.variable.length < 12) st.variable.push({ name: n, budget: '', need: !!v[2] });
    renderVar();
  }
  function addOwnVar() { if (st.variable.length < 12) { st.variable.push({ name: '', budget: '', need: false, own: true }); renderVar(); } }
  function delVar(i) { st.variable.splice(i, 1); renderVar(); }
  function next6() { st.variable = st.variable.filter(x => String(x.name).trim()); go(7); }
  function next7() { go(8); }

  // 7. Правило розподілу
  function incomeTotal() { return st.income.reduce((a, s) => a + (money(s.amount) || 0), 0); }
  function renderRules() {
    const inc = incomeTotal() || 40000;
    $('rules').innerHTML = D.rules.map(r =>
      '<div class="rule' + (st.rule === r.name ? ' sel' : '') + '" onclick="pickRule(&quot;' + esc(r.name) + '&quot;)">' +
      '<div class="rt"><span>' + esc(r.title) + '</span>' + (r.tag ? '<span class="tag">' + esc(r.tag) + '</span>' : '') + '</div>' +
      '<div class="ra">' + esc(r.about) + '</div>' +
      (r.buckets.length ? '<div class="rx">' + r.buckets.map(b => '<span>' + esc(b[0]) + ' ' + Math.round(b[1] * 100) + '% · ' + fmt(inc * b[1]) + '</span>').join('') + '</div>' : '') +
      '</div>').join('') +
      '<div class="hint">Суми — приклад на ' + (incomeTotal() ? 'твоєму доході ' + fmt(inc) : 'доході 40 000 грн') + ' на місяць.</div>';
  }
  function pickRule(n) { st.rule = n; renderRules(); }

  // 8. Підсумок
  function summary() {
    $('h8').textContent = 'Готово, ' + st.name + '!';
    $('p8').textContent = G('Ти готова? ', 'Ти готовий? ') + 'Перевір — натисни на блок, щоб змінити.';
    const inc = st.income.map(s => (s.main ? '★ ' : '') + esc(s.name) + (money(s.amount) ? ' — ' + fmt(money(s.amount)) : '') + (s.day ? ', ' + s.day + '-го' : ', нерегулярно')).join('<br>');
    const regList = st.regular.filter(x => !x.sub);
    const reg = (regList.length ? regList.map(x => esc(x.name) + (money(x.amount) ? ' — ' + fmt(money(x.amount)) + (x.half ? ' (моя частка 50%)' : '') : '') + (x.day ? ', ' + x.day + '-го' : '')).join('<br>') : 'не вказано — допишеш потім') +
      (st.subs.length ? '<br>Підписки: ' + st.subs.map(x => esc(x.name) + (money(x.amount) ? ' ' + fmt(money(x.amount)) : '')).join(', ') : '');
    const vr = st.variable.length ? st.variable.map(x => esc(x.name) + (money(x.budget) ? ' — ' + fmt(money(x.budget)) : '')).join(', ') : 'не вказано';
    const blk = (step, title, body) => '<div class="blk" onclick="go(' + step + ')"><div class="bh">' + title + '<em>змінити</em></div><div class="bb">' + body + '</div></div>';
    $('sum').innerHTML =
      blk(1, 'Про тебе', esc(st.name) + ' · ' + (st.gender === 'm' ? '👨 чоловік' : '👩 жінка') + ' · ' + st.age + ' ' + yw(st.age)) +
      blk(4, 'Дохід' + (st.floating ? ' (сума плаває)' : ''), inc) +
      blk(5, 'Обовʼязкові платежі', reg) +
      blk(6, 'Щоденні витрати', vr) +
      blk(7, 'Правило розподілу', esc((D.rules.filter(r => r.name === st.rule)[0] || {}).title || st.rule));
    $('replaceBox').style.display = D.hasData ? 'block' : 'none';
  }

  function payload() {
    return {
      name: st.name, gender: st.gender, age: st.age, floating: st.floating,
      income: st.income.map(s => ({ name: s.name, amount: money(s.amount), day: s.day, main: s.main })),
      regular: st.regular.filter(x => !x.sub).map(x => ({ name: x.name, amount: money(x.amount), day: x.day, half: x.half, debt: x.debt })),
      subs: st.subs.map(x => ({ name: x.name, amount: money(x.amount) })),
      rule: st.rule,
      variable: st.variable.map(x => ({ name: x.name, budget: money(x.budget), need: x.need })),
    };
  }
  function send(mode) {
    const apply = D.hasData ? $('replace').checked : true;
    go(9);
    $('h9').textContent = mode === 'build' ? 'Будую твій трекер…' : 'Зберігаю…';
    $('p9').textContent = mode === 'build' ? 'Це до хвилини — можеш поки зробити чай ☕' : 'Секунду.';
    google.script.run
      .withSuccessHandler(r => {
        $('em10').textContent = '✅';
        $('h10').textContent = r === 'built' ? 'Трекер готовий, ' + st.name + '!' : 'Збережено!';
        $('p10').textContent = 'Покажу, де що лежить і як цим користуватись — інструкція відкриється збоку, а таблиця залишиться під рукою.';
        $('closeLink').textContent = 'Закрити — розберусь ' + G('сама', 'сам');
        go(10);
      })
      .withFailureHandler(e => {
        $('em10').textContent = '😕'; $('h10').textContent = 'Щось пішло не так'; $('p10').textContent = (e && e.message) || String(e); go(10);
      })
      .quizSubmit(payload(), mode, apply);
  }
  function openGuide() { google.script.run.withSuccessHandler(closeMe).withFailureHandler(closeMe).showGuide(); }
  function closeMe() { google.script.host.close(); }

  swatch('swf', D.themes.f); swatch('swm', D.themes.m);
  $('name').value = st.name;
  $('name').addEventListener('input', () => { $('e1').textContent = ''; });
  $('age').addEventListener('input', () => { const n = Number($('age').value); if (n >= 14 && n <= 100) { st.age = n; $('range').value = n; $('unit').textContent = yw(n); } });
  $('range').addEventListener('input', e => setAge(e.target.value));
  document.addEventListener('keydown', e => {
    if (e.key !== 'Enter' || e.target.tagName === 'SELECT') return;
    if (cur === 0) go(1); else if (cur === 1) next1(); else if (cur === 2) next2(); else if (cur === 3) next3();
  });
  paint(); go(0);
</script></body></html>`;

const GUIDE_HTML = String.raw`<!DOCTYPE html>
<html lang="uk"><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  :root { --p:#4F86B5; --a:#8CC4E8; --bg:#E5F1FA; --t:#2C5677; --ink:#2E3A46; --muted:#8A9BAA; }
  * { box-sizing:border-box; }
  html,body { margin:0; height:100%; font-family:Montserrat, Arial, sans-serif; color:var(--ink); background:#fff; }
  .wrap { display:flex; flex-direction:column; height:100%; padding:12px 14px; }
  .top { display:flex; align-items:center; gap:8px; margin-bottom:6px; }
  .bar { flex:1; height:5px; background:var(--bg); border-radius:99px; overflow:hidden; }
  .bar i { display:block; height:100%; background:linear-gradient(90deg,var(--p),var(--a)); transition:width .4s; }
  .cnt { font-size:12px; font-weight:700; color:var(--muted); }
  .body { flex:1; overflow:auto; padding:2px 2px 8px; }
  .emoji { font-size:34px; margin:6px 0 6px; }
  h1 { font-size:19px; font-weight:800; color:var(--t); margin:0 0 8px; line-height:1.25; }
  p { font-size:13.5px; line-height:1.5; margin:0 0 8px; }
  ul { list-style:none; padding:0; margin:0 0 8px; }
  li { position:relative; padding:6px 0 6px 18px; font-size:13.5px; line-height:1.45; }
  li:before { content:""; position:absolute; left:2px; top:13px; width:8px; height:8px; border-radius:50%; background:var(--a); }
  .ex { background:#F7F9FB; border:1px dashed #D5DEE7; border-radius:12px; padding:9px 11px; font-size:12.5px; line-height:1.5; margin:4px 0 10px; }
  .ex b { color:var(--t); }
  .tip { background:var(--bg); border-radius:12px; padding:9px 11px; font-size:12.5px; font-weight:600; color:var(--t); margin:4px 0 10px; }
  .here { display:flex; gap:6px; align-items:center; font-size:12px; color:var(--p); font-weight:700; margin:0 0 8px; }
  .chips { display:flex; gap:6px; flex-wrap:wrap; margin:0 0 8px; }
  .chip { font-size:11.5px; font-weight:700; border-radius:9px; padding:5px 8px; }
  /* «термін — пояснення»: у вузькій панелі один під одним */
  table.cols, table.cols tbody, table.cols tr, table.cols td { display:block; width:100%; }
  table.cols { font-size:13.5px; margin:0 0 10px; }
  table.cols tr { border-bottom:1px solid #EEF1F4; padding:7px 0; }
  table.cols td { padding:0; line-height:1.45; }
  table.cols td:first-child { font-weight:700; color:var(--t); margin-bottom:2px; }
  .btn { border:none; border-radius:12px; padding:12px; font:700 14px Montserrat, Arial; cursor:pointer; }
  .btn.main { background:var(--p); color:#fff; }
  .btn.soft { background:var(--bg); color:var(--t); width:100%; margin:4px 0 8px; }
  .btn.ghost { background:none; color:var(--muted); padding:12px 4px; }
  .foot { display:flex; gap:8px; align-items:center; padding-top:8px; border-top:1px solid #EEF1F4; }
  .foot .main { flex:1; }
  .tabs { display:flex; background:#F1F4F7; border-radius:12px; padding:3px; margin:6px 0 10px; }
  .tabs button { flex:1; border:none; background:none; border-radius:10px; padding:8px; font:700 13px Montserrat, Arial; color:var(--muted); cursor:pointer; }
  .tabs button.on { background:#fff; color:var(--t); box-shadow:0 1px 4px rgba(0,0,0,.08); }
  .step { display:flex; gap:10px; align-items:flex-start; margin:0 0 10px; }
  .step .n { flex:0 0 24px; height:24px; border-radius:50%; background:var(--p); color:#fff; font-weight:800; font-size:12px; display:flex; align-items:center; justify-content:center; }
  .step .tx { font-size:13px; line-height:1.45; }
  .step svg { vertical-align:middle; }
  .phone { display:inline-flex; align-items:center; justify-content:center; width:30px; height:30px; border-radius:8px; background:#F1F4F7; margin:2px 4px 0 0; vertical-align:middle; }
  .qr { text-align:center; margin:6px 0 4px; }
  .qr svg { width:150px; height:150px; }
  .url { font-size:11px; color:var(--muted); word-break:break-all; text-align:center; margin-bottom:8px; }
  .tabgrid { display:grid; grid-template-columns:1fr; gap:6px; margin-bottom:8px; }
  .tabgrid button { text-align:left; border:2px solid #EEF1F4; background:#fff; border-radius:12px; padding:8px 10px; cursor:pointer; font:600 12.5px Montserrat, Arial; color:var(--ink); line-height:1.4; }
  .tabgrid button b { color:var(--t); display:block; font-size:13.5px; }
  .msg { font-size:12px; color:#D9534F; min-height:14px; }
</style></head>
<body><div class="wrap">
  <div class="top"><div class="cnt" id="cnt"></div><div class="bar"><i id="bar"></i></div></div>
  <div class="body" id="body"></div>
  <div class="msg" id="msg"></div>
  <div class="foot"><button class="btn ghost" id="back" onclick="show(i - 1)">← Назад</button><button class="btn main" id="next" onclick="nextStep()">Далі →</button></div>
</div>
<script>
__QRLIB__
</script>
<script>
  const D = __DATA__;
  const G = (f, m) => D.gender === 'm' ? m : f;
  const NAME = D.name ? D.name : '';
  let i = 0;
  let os = 'ios';
  let formUrl = D.formUrl || '';

  const ico = {
    share: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#1C7CF4" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/><path d="M8 7l4-4 4 4"/><path d="M5 11v8a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-8"/></svg>',
    plus: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#333" stroke-width="2" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="4"/><path d="M12 8v8M8 12h8"/></svg>',
    dots: '<svg width="18" height="18" viewBox="0 0 24 24" fill="#333"><circle cx="12" cy="5" r="2"/><circle cx="12" cy="12" r="2"/><circle cx="12" cy="19" r="2"/></svg>',
    safari: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#1C7CF4" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M15.5 8.5l-2 5-5 2 2-5z" fill="#1C7CF4"/></svg>',
    chrome: '<svg width="18" height="18" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#DB4437"/><path d="M12 12L3.3 7A10 10 0 0 1 12 2h0z" fill="#DB4437"/><path d="M12 12l-8.7-5A10 10 0 0 0 7 20.7z" fill="#0F9D58"/><path d="M12 12l5 8.7A10 10 0 0 0 20.7 7H12z" fill="#F4B400"/><circle cx="12" cy="12" r="4.2" fill="#fff"/><circle cx="12" cy="12" r="3.2" fill="#4285F4"/></svg>',
    home: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#333" stroke-width="2" stroke-linejoin="round"><path d="M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/></svg>',
    check: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#2E7D32" stroke-width="2.5" stroke-linecap="round"><path d="M5 12l5 5 9-10"/></svg>',
  };
  const ph = k => '<span class="phone">' + ico[k] + '</span>';

  function qrSvg(url) {
    try { const q = qrcode(0, 'M'); q.addData(url); q.make(); return q.createSvgTag({ cellSize: 4, margin: 2, scalable: true }); }
    catch (e) { return ''; }
  }

  const steps = [
    { focus: null, render: () =>
      '<div class="emoji">🧭</div><h1>' + (NAME ? NAME + ', ось' : 'Ось') + ' як це працює</h1>' +
      '<p>Ця панель — твій провідник. Таблиця поруч лишається доступною: я відкриватиму потрібну вкладку й підсвічуватиму блок, про який іде мова.</p>' +
      '<ul><li><b>Один раз</b> налаштуй дохід, обовʼязкові платежі й бюджети.</li><li><b>Щодня</b> вписуй суми — в «Операції» або з телефона.</li>' +
      '<li><b>Дивись велику цифру</b> на «Дашборді» — скільки можна витратити до наступної виплати.</li></ul>' +
      '<div class="tip">💡 Формул писати не треба — усе рахується саме.</div>' },

    { focus: 'income', render: () =>
      '<div class="emoji">⚙️</div><h1>Налаштування: дохід</h1>' +
      '<div class="here">👉 Відкрито вкладку «Налаштування», блок «Дохід» підсвічено</div>' +
      '<div class="chips"><span class="chip" style="background:#FFF8D6">жовте — заповнюєш ти</span><span class="chip" style="background:#F2F4F7">сіре — рахується само</span></div>' +
      '<table class="cols"><tr><td>Джерело</td><td>звідки гроші: Зарплата, Аванс, Фріланс…</td></tr>' +
      '<tr><td>Основне ☑</td><td>галочка на головній виплаті. Від її дня трекер рахує «до наступної виплати».</td></tr>' +
      '<tr><td>Очікувана сума</td><td>скільки зазвичай приходить. Можна лишити порожнім, якщо плаває.</td></tr>' +
      '<tr><td>День виплати</td><td>число місяця, напр. 1 або 15.</td></tr></table>' +
      '<div class="tip">💡 Якщо ти ' + G('пройшла', 'пройшов') + ' квіз — тут уже твої відповіді. Перевір і поправ, якщо треба.</div>' },

    { focus: 'regular', render: () =>
      '<div class="emoji">🏠</div><h1>Обовʼязкові платежі</h1>' +
      '<div class="here">👉 Блок «Регулярні витрати» підсвічено</div>' +
      '<p>Платежі, що повторюються щомісяця чи раз на квартал / рік: оренда, комуналка, інтернет, підписки.</p>' +
      '<table class="cols"><tr><td>Загальна сума</td><td>увесь рахунок, напр. оренда 30 000.</td></tr>' +
      '<tr><td>Моя частка</td><td>яку частину платиш ти. Ділиш навпіл → 50%, і «Моя сума» стане 15 000.</td></tr>' +
      '<tr><td>Як часто</td><td>щомісяця / раз на квартал / раз на рік.</td></tr>' +
      '<tr><td>День оплати</td><td>число місяця — щоб трекер знав, чи платіж уже мав бути.</td></tr>' +
      '<tr><td>Група</td><td>«Потреби» — без цього не обійтись (оренда, комуналка), «Бажання» — можна відмовитись.</td></tr></table>' +
      '<p><b>Підписки</b> (Netflix, Spotify, iCloud…) — в окремому блоці «Підписки» нижче: сума, як часто, день списання і «Користуюсь». ' +
      'Якщо зняти «Користуюсь», трекер покаже, скільки можна звільнити за рік.</p>' },

    { focus: 'variable', render: () =>
      '<div class="emoji">🛒</div><h1>Щоденні витрати</h1>' +
      '<div class="here">👉 Блок «Змінні витрати» підсвічено</div>' +
      '<p>Те, на що гроші йдуть щодня: їжа, транспорт, кафе, краса. Для кожної категорії можна поставити <b>місячний бюджет</b> — ліміт, скільки хочеш на неї витрачати.</p>' +
      '<div class="ex">Бюджет на «Їжа» — <b>7 000 грн</b>. ' + G('Витратила', 'Витратив') + ' 5 600 → на «Дашборді» смужка 80% стане жовтою, понад 7 000 — червоною.</div>' },

    { focus: 'debts', render: () =>
      '<div class="emoji">💳</div><h1>Борги, цілі, рахунки</h1>' +
      '<div class="here">👉 Блок «Борги і розстрочки» підсвічено</div>' +
      '<table class="cols"><tr><td>Сума платежу</td><td>скільки платиш за раз, напр. <b>1 590 грн</b> щомісяця за телефон частинами.</td></tr>' +
      '<tr><td>Початковий залишок</td><td>скільки всього ще ' + G('винна', 'винен') + ' на старті, напр. <b>9 540 грн</b>. Трекер сам віднімає кожен платіж.</td></tr>' +
      '<tr><td>Закрито достроково ☑</td><td>якщо ' + G('погасила', 'погасив') + ' раніше — постав галочку.</td></tr></table>' +
      '<p>Нижче ще два блоки: <b>«Заощадження і цілі»</b> — на що відкладаєш (подушка, відпустка), і <b>«Рахунки»</b> — картки й готівка, щоб бачити баланс кожної.</p>' },

    { focus: 'ops', render: () =>
      '<div class="emoji">🧾</div><h1>Операції — щоденний запис</h1>' +
      '<div class="here">👉 Відкрито «Операції», виділено перший вільний рядок</div>' +
      '<p>Кожна покупка чи надходження — один рядок. Колонки:</p>' +
      '<table class="cols"><tr><td>Дата</td><td>коли. Порожня → поставиться сьогодні.</td></tr>' +
      '<tr><td>Тип</td><td>що сталося з грошима:<br><b>Дохід</b> — гроші прийшли<br><b>Витрата</b> — ' + G('купила', 'купив') + ' чи ' + G('оплатила', 'оплатив') + '<br><b>Борг</b> — платіж по кредиту чи розстрочці<br><b>Заощадження</b> — ' + G('відклала', 'відклав') + '</td></tr>' +
      '<tr><td>Сума</td><td>скільки гривень.</td></tr>' +
      '<tr><td>Категорія</td><td>на що саме. Список залежить від типу: для «Витрата» — Їжа, Оренда, Кафе…; для «Дохід» — Зарплата, Аванс.</td></tr>' +
      '<tr><td>Рахунок</td><td>картка чи готівка. Порожньо — основний рахунок.</td></tr></table>' +
      '<div class="ex"><b>Приклад:</b> ' + G('купила', 'купив') + ' продукти на 250 грн.<br>1. Впиши <b>250</b> у колонку «Сума».<br>2. Дата і тип «Витрата» поставляться самі.<br>3. У «Категорія» обери <b>Їжа</b>. Готово ✓</div>' },

    { focus: 'form', render: renderForm },

    { focus: 'dash', render: () =>
      '<div class="emoji">🎯</div><h1>Дашборд і велика цифра</h1>' +
      '<div class="here">👉 Відкрито «Дашборд», велику цифру підсвічено</div>' +
      '<p><b>Велика цифра</b> — скільки ти ще можеш витратити до наступної виплати. Вона рахується так:</p>' +
      '<div class="ex">Дохід за цикл <b>42 900</b><br>− уже витрачено <b>33 500</b> (усе, що ' + G('купила', 'купив') + ' й ' + G('оплатила', 'оплатив') + ': продукти, оренда, кава…)<br>− ще не оплачені обовʼязкові платежі<br>− платежі по боргах цього місяця<br>− заплановані заощадження<br>= <b>9 400 грн</b> на щоденні витрати до виплати</div>' +
      '<ul><li>Під цифрою — <b>ліміт на день</b>: 9 400 ÷ днів до виплати.</li>' +
      '<li><b>Зелена</b> — ти в нормі. <b>Червона</b> — витрачаєш швидше, ніж зазвичай.</li>' +
      '<li>Блок <b>«Як рахується цифра»</b> показує ці рядки з твоїми сумами.</li>' +
      '<li>Вгорі <b>«Період»</b> — подивитися будь-який місяць.</li></ul>' },

    { focus: null, render: () =>
      '<div class="emoji">🗂️</div><h1>Інші вкладки</h1><p>Натисни — відкрию вкладку:</p><div class="tabgrid">' +
      [['Рік', 'усі категорії по місяцях, графік «дохід vs витрати», теплова карта витрат'],
       ['Цілі', 'скільки накопичено на кожну ціль і скільки відкладати щомісяця'],
       ['Борги', 'у якому порядку закривати (сніжна куля) і коли будеш ' + G('вільна', 'вільний') + ' від боргів'],
       ['Підписки', 'скільки коштують на рік і що можна скасувати'],
       ['Календар', 'витрати по днях: зелений — у межах, червоний — понад бюджет, «—» — день без витрат'],
       ['Старт', 'посилання на форму, перевірка таблиці, твій профіль']]
        .map(t => '<button onclick="focusOn(&quot;' + t[0] + '&quot;)"><b>' + t[0] + ' ↗</b>' + t[1] + '</button>').join('') + '</div>' },

    { focus: null, last: true, render: () =>
      '<div class="emoji">💡</div><h1>Кілька порад</h1><ul>' +
      '<li>Вписуй витрати щодня — це хвилина.</li><li>Раз на тиждень зазирай на «Дашборд» і в «Календар».</li>' +
      '<li>' + G('Пропустила', 'Пропустив') + ' кілька днів — просто допиши, нічого не зламається.</li>' +
      '<li>Щось не так? Меню «Трекер» → <b>«Перевірити таблицю»</b>.</li>' +
      '<li>Повернутися до цієї інструкції: меню «Трекер» → <b>«📖 Як користуватись»</b>.</li></ul>' +
      '<div class="tip">🚀 Почни з «Налаштувань» — перевір дохід і платежі.</div>' },
  ];

  function renderForm() {
    let h = '<div class="emoji">📱</div><h1>З телефона — за 10 секунд</h1>';
    if (!formUrl) {
      h += '<p>Створю для тебе Google Форму «Нова операція». Ти відкриватимеш її на телефоні, а кожна відповідь сама зʼявлятиметься в «Операціях».</p>' +
        '<button class="btn soft" id="mk" onclick="makeForm()">✨ Створити форму</button>';
    } else {
      h += '<p><b>1.</b> Наведи камеру телефона на QR-код (або відкрий посилання з вкладки «Старт»):</p>' +
        '<div class="qr">' + qrSvg(formUrl) + '</div><div class="url">' + formUrl + '</div>';
    }
    h += '<p><b>2.</b> Додай форму на головний екран — тоді вона відкриватиметься як застосунок:</p>' +
      '<div class="tabs"><button id="tIos" class="' + (os === 'ios' ? 'on' : '') + '" onclick="setOs(&quot;ios&quot;)">📱 iPhone</button>' +
      '<button id="tAnd" class="' + (os === 'android' ? 'on' : '') + '" onclick="setOs(&quot;android&quot;)">🤖 Android</button></div>';
    const st = (n, t) => '<div class="step"><div class="n">' + n + '</div><div class="tx">' + t + '</div></div>';
    if (os === 'ios') {
      h += st(1, ph('safari') + 'Відкрий посилання в <b>Safari</b> (саме в Safari, не в іншому браузері).') +
        st(2, ph('share') + 'Натисни кнопку <b>«Поділитися»</b> — квадрат зі стрілкою вгору, внизу екрана.') +
        st(3, ph('plus') + 'Прогорни список вниз і вибери <b>«На початковий екран»</b>.') +
        st(4, ph('check') + 'Натисни <b>«Додати»</b> вгорі справа. На екрані зʼявиться іконка форми.');
    } else {
      h += st(1, ph('chrome') + 'Відкрий посилання в <b>Chrome</b>.') +
        st(2, ph('dots') + 'Натисни меню <b>⋮</b> у правому верхньому куті.') +
        st(3, ph('home') + 'Вибери <b>«Додати на головний екран»</b> (іноді — «Встановити застосунок»).') +
        st(4, ph('check') + 'Підтвердь <b>«Додати»</b>. Іконка зʼявиться на головному екрані.');
    }
    h += '<p><b>3.</b> Вносиш так: <b>сума → тип → категорія → «Надіслати»</b>. Рядок сам зʼявиться в «Операціях».</p>' +
      '<div class="tip">💡 Найзручніше вносити прямо на касі — тоді нічого не забудеш.</div>';
    return h;
  }
  function setOs(o) { os = o; show(i, true); }
  function makeForm() {
    const b = document.getElementById('mk'); b.textContent = 'Створюю… це ~20 секунд'; b.disabled = true;
    google.script.run.withSuccessHandler(url => { formUrl = url || ''; show(i, true); })
      .withFailureHandler(e => { b.disabled = false; b.textContent = '✨ Створити форму'; msg((e && e.message) || String(e)); })
      .createEntryForm();
  }

  function msg(t) { document.getElementById('msg').textContent = t || ''; }
  function focusOn(key) {
    msg('');
    if (!(window.google && google.script && google.script.run)) return;
    google.script.run.withSuccessHandler(r => { if (r !== 'ok') msg(r); }).withFailureHandler(e => msg((e && e.message) || String(e))).tourFocus(key);
  }
  function show(n, keep) {
    i = Math.max(0, Math.min(steps.length - 1, n));
    const s = steps[i];
    document.getElementById('body').innerHTML = s.render();
    document.getElementById('body').scrollTop = 0;
    document.getElementById('cnt').textContent = (i + 1) + ' / ' + steps.length;
    document.getElementById('bar').style.width = ((i + 1) / steps.length * 100) + '%';
    document.getElementById('back').style.visibility = i === 0 ? 'hidden' : 'visible';
    document.getElementById('next').textContent = s.last ? 'Почати з налаштувань 🚀' : 'Далі →';
    if (!keep) { msg(''); if (s.focus) focusOn(s.focus); }
  }
  function nextStep() {
    if (steps[i].last) { focusOn('income'); return; }
    show(i + 1);
  }
  (function theme() {
    const t = D.theme, r = document.documentElement.style;
    r.setProperty('--p', t.primary); r.setProperty('--a', t.accent); r.setProperty('--bg', t.headerBg); r.setProperty('--t', t.headerTxt);
  })();
  show(0);
</script></body></html>`;

/* QR Code Generator for JavaScript — Copyright (c) 2009 Kazuhiko Arase, MIT license.
   https://github.com/kazuhikoarase/qrcode-generator */
const QR_LIB = String.raw`// Copyright (c) 2009 Kazuhiko Arase
// Licensed under the MIT license:
// DENSO WAVE INCORPORATED
var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())},f.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<f.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<f.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=f.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},f.createSvgTag=function(t,r,e,n){var o={};"object"==typeof arguments[0]&&(t=(o=arguments[0]).cellSize,r=o.margin,e=o.alt,n=o.title),t=t||2,r=void 0===r?4*t:r,(e="string"==typeof e?{text:e}:e||{}).text=e.text||null,e.id=e.text?e.id||"qrcode-description":null,(n="string"==typeof n?{text:n}:n||{}).text=n.text||null,n.id=n.text?n.id||"qrcode-title":null;var i,a,u,c,g=f.getModuleCount()*t+2*r,l="";for(c="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",l+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',l+=o.scalable?"":' width="'+g+'px" height="'+g+'px"',l+=' viewBox="0 0 '+g+" "+g+'" ',l+=' preserveAspectRatio="xMinYMin meet"',l+=n.text||e.text?' role="img" aria-labelledby="'+y([n.id,e.id].join(" ").trim())+'"':"",l+=">",l+=n.text?'<title id="'+y(n.id)+'">'+y(n.text)+"</title>":"",l+=e.text?'<description id="'+y(e.id)+'">'+y(e.text)+"</description>":"",l+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',l+='<path d="',a=0;a<f.getModuleCount();a+=1)for(u=a*t+r,i=0;i<f.getModuleCount();i+=1)f.isDark(a,i)&&(l+="M"+(i*t+r)+","+u+c);return l+='" stroke="transparent" fill="black"/>',l+="</svg>"},f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return I(e,e,function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1})},f.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=f.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=f.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=y(e),o+='"'),o+="/>"};var y=function(t){for(var r="",e=0;e<t.length;e+=1){var n=t.charAt(e);switch(n){case"<":r+="&lt;";break;case">":r+="&gt;";break;case"&":r+="&amp;";break;case'"':r+="&quot;";break;default:r+=n}}return r};return f.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*f.getModuleCount()+2*t,u=t,c=a-t,g={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},h="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<c&&u<=r&&r<c&&f.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<c&&u<=r+1&&r+1<c&&f.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",h+=t<1&&r+1>=c?l[i]:g[i];h+="\n"}return a%2&&t>0?h.substring(0,h.length-a-1)+Array(a+1).join("▀"):h.substring(0,h.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=f.getModuleCount()*t+2*r,u=r,c=a-r,g=Array(t+1).join("██"),l=Array(t+1).join("  "),h="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<c&&u<=e&&e<c&&f.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?g:l;for(o=0;o<t;o+=1)h+=s+"\n"}return h.substring(0,h.length-1)},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(n*r,o*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,t.createStringToBytes=function(t,r){var e=function(){for(var e=S(t),n=function(){var t=e.read();if(-1==t)throw"eof";return t},o=0,i={};;){var a=e.read();if(-1==a)break;var u=n(),f=n()<<8|n();i[String.fromCharCode(a<<8|u)]=f,o+=1}if(o!=r)throw o+" != "+r;return i}(),n="?".charCodeAt(0);return function(t){for(var r=[],o=0;o<t.length;o+=1){var i=t.charCodeAt(o);if(i<128)r.push(i);else{var a=e[t.charAt(o)];"number"==typeof a?(255&a)==a?r.push(a):(r.push(a>>>8),r.push(255&a)):r.push(n)}}return r}};var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(){var t=n("友");if(2!=t.length||38726!=(t[0]<<8|t[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},S=function(t){var r=t,e=0,n=0,o=0,i={read:function(){for(;o<8;){if(e>=r.length){if(0==o)return-1;throw"unexpected end of file./"+o}var t=r.charAt(e);if(e+=1,"="==t)return o=0,-1;t.match(/^\s$/)||(n=n<<6|a(t.charCodeAt(0)),o+=6)}var i=n>>>o-8&255;return o-=8,i}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return i},I=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();qrcode.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||n>=57344?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},function(t){"function"==typeof define&&define.amd?define([],t):"object"==typeof exports&&(module.exports=t())}(function(){return qrcode});`;

// ─────────────────────────────────────────────────────────────────────────────
// ІМПОРТ БАНКІВСЬКОЇ ВИПИСКИ (CSV або Google-таблиця)
// Колонки визначаються за назвами, категорії — за вивченими правилами, ключовими словами
// і кодом MCC. Перекази між своїми рахунками й скасовані покупки пропускаються.
// ─────────────────────────────────────────────────────────────────────────────
const STMT = {
  tag: 'виписка',
  propRules: 'BT_RULES',
  // Канонічні категорії: [назва, синоніми, які вже можуть бути в трекері, група «Потреби»?]
  cats: {
    food:    ['Їжа', ['Їжа', 'Продукти', 'Продукти і кава', 'Продукти й кава'], true],
    cafe:    ['Кафе й доставка', ['Кафе й доставка', 'Кафе і доставка', 'Кафе', 'Ресторани'], false],
    transp:  ['Транспорт', ['Транспорт', 'Таксі', 'Авто'], true],
    health:  ['Здоровʼя', ['Здоровʼя', 'Здоров\'я', 'Краса й здоровʼя', 'Здоров\'я і краса', 'Аптека', 'Медицина'], true],
    shop:    ['Покупки', ['Покупки', 'Речі'], false],
    tech:    ['Техніка', ['Техніка', 'Електроніка'], false],
    subs:    ['Підписки', ['Підписки', 'Підписки (разом)'], false],
    fun:     ['Розваги', ['Розваги', 'Кіно'], false],
    beauty:  ['Краса', ['Краса', 'Краса й здоровʼя', 'Здоров\'я і краса', 'Барбер'], false],
    sport:   ['Спорт', ['Спорт', 'Спорт і хобі', 'Спортзал'], false],
    edu:     ['Освіта', ['Освіта', 'Навчання'], false],
    travel:  ['Подорожі', ['Подорожі', 'Відпочинок', 'Відпустка'], false],
    pets:    ['Тварини', ['Тварини'], true],
    util:    ['Комуналка', ['Комуналка', 'Комуналка + інтернет'], true],
    phone:   ['Звʼязок', ['Звʼязок', 'Інтернет і звʼязок', 'Комуналка + інтернет'], true],
    fees:    ['Відсотки й комісії', ['Відсотки й комісії', 'Комісії'], true],
    other:   ['Інше', ['Інше'], false],
    cashback: ['Кешбек', ['Кешбек'], false],
    refund:  ['Повернення', ['Повернення'], false],
    incOther: ['Інший дохід', ['Інший дохід', 'Інше'], false],
  },
  keywords: [
    [/атб|сільпо|silpo|novus|varus|loko|фора|ашан|auchan|metro cash|еко маркет|food market|продукт|маркет|market/i, 'food'],
    [/glovo|bolt food|uber eats|rocket|mcdonald|kfc|кафе|ресторан|піца|pizza|sushi|суші|coffee|кава|starbucks/i, 'cafe'],
    [/uber|bolt|uklon|таксі|metro|метро|укрзалізниц|wog|okko|socar|upg|parking|паркінг/i, 'transp'],
    [/netflix|spotify|youtube|apple\.com|icloud|google|megogo|sweet\.tv|claude|anthropic|openai|chatgpt|adobe|notion/i, 'subs'],
    [/подорожник|анц|аптека|apteka|pharm|клінік|медичн|лаборатор|синево|dila/i, 'health'],
    [/aliexpress|temu|rozetka|prom\.ua|epicentr|епіцентр|zara|h&m|reserved/i, 'shop'],
    [/comfy|фокстрот|citrus|allo|moyo|wayforpay/i, 'tech'],
    [/київстар|kyivstar|vodafone|lifecell|інтернет|internet/i, 'phone'],
    [/відсотк|комісі|перевищення кредитного|штраф/i, 'fees'],
  ],
  mcc: [
    [[5411, 5412, 5422, 5441, 5451, 5462, 5499], 'food'],
    [[5811, 5812, 5813, 5814], 'cafe'],
    [[4111, 4121, 4131, 4784, 4789, 5541, 5542, 5172, 7523, 7512], 'transp'],
    [[5912, 5122, 8011, 8021, 8031, 8041, 8042, 8043, 8049, 8050, 8062, 8071, 8099], 'health'],
    [[5310, 5311, 5331, 5399, 5611, 5621, 5631, 5641, 5651, 5655, 5661, 5691, 5699, 5944, 5945, 5947, 5948, 5200, 5211, 5251, 5712, 5719], 'shop'],
    [[5045, 5722, 5732, 5734, 5946], 'tech'],
    [[4899, 5815, 5816, 5817, 5818, 7841], 'subs'],
    [[7832, 7922, 7929, 7933, 7991, 7994, 7996, 7998, 7999], 'fun'],
    [[7230, 7297, 7298, 5977], 'beauty'],
    [[7941, 7997, 5940, 5941], 'sport'],
    [[8211, 8220, 8241, 8244, 8249, 8299, 5942], 'edu'],
    [[3000, 4411, 4511, 4722, 7011, 7012], 'travel'],
    [[742, 5995], 'pets'],
    [[4900], 'util'],
    [[4812, 4814], 'phone'],
  ],
  ownTransfer: /з (білої|чорної|своєї|іншої)? ?картки|на (білу|чорну|свою) картку|переказ між (своїми|власними)|банк[иа] «|зняття банки|поповнення банки|з банки|на банку|між своїми рахунками|own account|internal transfer/i,
  cashback: /кешбек|cashback/i,
  cancel: /^(скасування|повернення коштів|відміна|refund|cancel)[.:,]?\s*/i,
};

/** Меню: діалог імпорту виписки. */
function showStatementImport() {
  ensureOwnDocument_();
  applyProfileTheme_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.ops)) {
    ss.toast('Спочатку побудуй трекер: меню «Трекер» → «Побудувати трекер з нуля».', menuTitle_(), 8);
    return;
  }
  const data = { theme: THEME, accounts: listValues_(ss, 'L_Acc'), gender: profile_().gender || '' };
  const html = HtmlService.createHtmlOutput(STMT_HTML.replace('__DATA__', () => JSON.stringify(data).replace(/</g, '\\u003c')))
    .setWidth(980).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🏦 Імпорт виписки з банку');
}

/** Рядки з Google-таблиці за посиланням (перший аркуш або той, на який вказує gid). */
function statementRowsFromSheet(url) {
  const id = (String(url).match(/[-\w]{25,}/) || [])[0];
  if (!id) throw new Error('Не бачу ID таблиці в посиланні.');
  const src = SpreadsheetApp.openById(id);
  const gid = (String(url).match(/gid=(\d+)/) || [])[1];
  const sh = (gid && src.getSheets().filter(s => String(s.getSheetId()) === gid)[0]) || src.getSheets()[0];
  return sh.getDataRange().getDisplayValues();
}

/** Розбір і попередня класифікація рядків виписки. rows[0] — заголовок (шукаємо в перших 15 рядках). */
function statementPreview(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const head = findStatementHeader_(rows);
  if (!head) throw new Error('Не знайшов колонок «Дата» і «Сума». Перевір, що це виписка з заголовками.');
  const lists = {};
  Object.keys(TYPE_LIST).forEach(t => { lists[t] = listValues_(ss, TYPE_LIST[t]); });
  const rules = statementRules_();
  const existing = existingStatementKeys_(ss);
  const subsR = ss.getRangeByName(NR + 'Subs_Name');
  lists.subs = subsR ? subsR.getValues().map(v => v[0]).filter(String) : [];

  const items = [];
  for (let r = head.row + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;
    const date = parseStmtDate_(row[head.date]);
    const raw = parseStmtAmount_(row[head.amount]);
    if (!date || !raw) continue;
    const desc = head.desc >= 0 ? String(row[head.desc] || '').trim() : '';
    const mcc = head.mcc >= 0 ? Number(row[head.mcc]) || '' : '';
    const it = { date: date, desc: desc, mcc: mcc, amount: Math.abs(raw), sign: raw < 0 ? -1 : 1, include: true, reason: '' };
    classifyStmtItem_(it, rules, lists);
    it.key = it.date + '|' + it.sign * it.amount + '|' + desc;
    if (existing[it.key]) { it.include = false; it.reason = 'вже імпортовано'; }
    items.push(it);
  }
  pairCancellations_(items);

  const inc = items.filter(x => x.include && x.type === TYPES.income);
  const out = items.filter(x => x.include && x.type !== TYPES.income);
  const sum = a => a.reduce((s, x) => s + x.amount, 0);
  return {
    items: items, lists: lists, columns: head.names,
    summary: { total: items.length, incomes: inc.length, incomeSum: sum(inc), expenses: out.length, expenseSum: sum(out),
      skipped: items.filter(x => !x.include).length },
  };
}

/** Запис вибраних рядків. items: [{date, type, amount, category, desc, mcc, key, learn}] */
function statementImport(items, account) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.ops);
  const chosen = (items || []).filter(x => x && x.date && Number(x.amount) > 0 && x.type && x.category);
  if (!chosen.length) return { added: 0, created: [] };

  // Категорії, яких ще немає, — додаємо в «Налаштування».
  const created = [];
  const lists = {};
  Object.keys(TYPE_LIST).forEach(t => { lists[t] = listValues_(ss, TYPE_LIST[t]); });
  chosen.forEach(x => {
    if (lists[x.type] && lists[x.type].indexOf(x.category) !== -1) return;
    const key = x.type === TYPES.income ? 'Inc' : x.type === TYPES.debt ? 'Debt' : x.type === TYPES.saving ? 'Sav' : 'Var';
    const canon = Object.keys(STMT.cats).map(k => STMT.cats[k]).filter(c => c[0] === x.category)[0];
    const row = key === 'Var' ? { Name: x.category, Group: canon && canon[2] ? REF.groups[0] : REF.groups[1] }
      : key === 'Debt' ? { Name: x.category, Freq: REF.freq[0][0] } : { Name: x.category };
    if (appendBlockRow_(ss, key, row)) { created.push(x.category); lists[x.type].push(x.category); }
    else x.category = key === 'Var' ? 'Інше' : x.category;
  });
  SpreadsheetApp.flush();

  // Вивчені правила: опис → тип і категорія (лише для змінених вручну).
  const rules = statementRules_();
  chosen.filter(x => x.learn && x.desc).forEach(x => { rules[normDesc_(x.desc)] = [x.type, x.category]; });
  PropertiesService.getDocumentProperties().setProperty(STMT.propRules, JSON.stringify(rules).slice(0, 8500));

  const existing = existingStatementKeys_(ss);
  const rows = chosen.filter(x => !existing[x.key]).map(x => [x.date, x.type, Math.round(Number(x.amount) * 100) / 100, x.category,
    account || '', x.desc + (x.mcc ? ' · MCC ' + x.mcc : ''), STMT.tag]);
  if (rows.length) {
    const first = firstFreeOpsRow_(sh);
    if (first + rows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 50);
    writeOpsRows_(sh, first, rows);
  }
  SpreadsheetApp.flush();
  syncCategoriesFromOps_(ss);
  return { added: rows.length, created: created };
}

function findStatementHeader_(rows) {
  const find = (cells, tests) => cells.findIndex(c => tests.some(t => t.test(c)));
  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const cells = (rows[r] || []).map(c => String(c || '').toLowerCase().trim());
    const date = find(cells, [/^дата|date/]);
    let amount = find(cells, [/сума в валюті картки|сума операції в валюті картки|amount.*card/]);
    if (amount < 0) amount = find(cells, [/^сума|^amount|сума$/]);
    if (date < 0 || amount < 0) continue;
    return {
      row: r, date: date, amount: amount,
      desc: find(cells, [/деталі|опис|призначення|description|details|merchant|контрагент/]),
      mcc: find(cells, [/^mcc/]),
      names: { date: rows[r][date], amount: rows[r][amount] },
    };
  }
  return null;
}

/** «09.11.2026 20:55:53», «2026-11-09», «09/11/2026» → «2026-11-09». */
function parseStmtDate_(v) {
  if (v instanceof Date) return isoDate_(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
  if (m) return isoDate_(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return isoDate_(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return '';
}

/** «−1 234,56», «-429.1», «1 000.00 UAH» → число зі знаком. */
function parseStmtAmount_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/[\s ]/g, '').replace(/[−–]/g, '-').replace(/[^0-9,.\-]/g, '');
  if (!s || s === '-') return 0;
  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) s = s.replace(/,/g, '');
  else s = s.replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function normDesc_(d) {
  return String(d || '').toLowerCase().replace(STMT.cancel, '').replace(/[0-9]+/g, '').replace(/[^a-zа-яіїєґ]+/gi, ' ').trim().slice(0, 40);
}

function statementRules_() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(STMT.propRules) || '{}'); } catch (e) { return {}; }
}

function existingStatementKeys_(ss) {
  const keys = {};
  opsRows_(ss).forEach(o => {
    if (o.source !== STMT.tag || !(o.date instanceof Date)) return;
    const d = isoDate_(o.date.getFullYear(), o.date.getMonth(), o.date.getDate());
    const desc = String(o.comment || '').replace(/ · MCC \d+$/, '');
    const sign = o.type === TYPES.income ? 1 : -1;
    keys[d + '|' + sign * o.amount + '|' + desc] = true;
  });
  return keys;
}

/** Назва категорії: та, що вже є в трекері (за синонімами), або канонічна. */
function stmtCategory_(code, existing) {
  const c = STMT.cats[code] || STMT.cats.other;
  const hit = c[1].filter(n => existing.indexOf(n) !== -1)[0];
  return hit || c[0];
}

function classifyStmtItem_(it, rules, lists) {
  const d = it.desc;
  const learned = rules[normDesc_(d)];
  if (STMT.ownTransfer.test(d)) { it.include = false; it.reason = 'переказ між своїми'; }
  if (it.sign > 0) {
    if (STMT.cancel.test(d)) { it.type = TYPES.expense; it.category = ''; it.cancel = true; it.include = false; it.reason = 'повернення / скасування'; return; }
    it.type = TYPES.income;
    if (STMT.cashback.test(d)) it.category = stmtCategory_('cashback', lists[TYPES.income]);
    else it.category = (lists[TYPES.income].filter(n => /зарплат/i.test(n))[0] && /зарплат|salary|заробітн/i.test(d))
      ? lists[TYPES.income].filter(n => /зарплат/i.test(n))[0] : stmtCategory_('incOther', lists[TYPES.income]);
  } else {
    it.type = TYPES.expense;
    const sub = (lists.subs || []).filter(n => n && d.toLowerCase().indexOf(String(n).toLowerCase()) !== -1)[0];
    if (sub) { it.category = sub; if (learned && it.include) { it.type = learned[0]; it.category = learned[1]; it.learned = true; } return; }
    let code = '';
    STMT.keywords.some(k => { if (k[0].test(d)) { code = k[1]; return true; } return false; });
    if (!code && it.mcc) STMT.mcc.some(m => { if (m[0].indexOf(Number(it.mcc)) !== -1) { code = m[1]; return true; } return false; });
    it.category = stmtCategory_(code || 'other', lists[TYPES.expense]);
  }
  if (learned && it.include) { it.type = learned[0]; it.category = learned[1]; it.learned = true; }
}

/** «Скасування. Glovo +402,2» знімає пару «Glovo −402,2». */
function pairCancellations_(items) {
  items.filter(x => x.cancel).forEach(c => {
    const name = normDesc_(c.desc);
    const match = items.filter(x => x !== c && !x.cancel && x.sign < 0 && x.include && Math.abs(x.amount - c.amount) < 0.005 && normDesc_(x.desc) === name)[0];
    if (match) { match.include = false; match.reason = 'скасовано'; c.reason = 'скасування'; c.category = match.category; }
    else { c.type = TYPES.income; c.category = 'Повернення'; c.include = true; c.reason = ''; }
  });
}

/** Додає рядок у перший порожній рядок блоку «Налаштувань». false — якщо місця немає. */
function appendBlockRow_(ss, key, obj) {
  const block = SETTINGS_BLOCKS.filter(b => b.key === key)[0];
  const names = ss.getRangeByName(NR + key + '_Name');
  if (!block || !names) return false;
  const idx = names.getValues().findIndex(v => v[0] === '' || v[0] === null);
  if (idx < 0) return false;
  block.cols.forEach(c => {
    if (c.t === 'f' || obj[c.key] === undefined) return;
    ss.getRangeByName(NR + key + '_' + c.key).getCell(idx + 1, 1).setValue(obj[c.key]);
  });
  return true;
}

const STMT_HTML = String.raw`<!DOCTYPE html>
<html lang="uk"><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
  :root { --p:#4F86B5; --a:#8CC4E8; --bg:#E5F1FA; --t:#2C5677; --ink:#2E3A46; --muted:#8A9BAA; --line:#E7EBF0; }
  * { box-sizing:border-box; }
  html,body { margin:0; height:100%; font-family:Montserrat, Arial, sans-serif; color:var(--ink); background:#fff; }
  .wrap { display:flex; flex-direction:column; height:100%; padding:16px 20px; }
  .screen { display:none; flex-direction:column; flex:1; min-height:0; }
  .screen.on { display:flex; }
  h1 { font-size:22px; font-weight:800; color:var(--t); margin:0 0 6px; }
  p.sub { font-size:14px; color:var(--muted); margin:0 0 14px; line-height:1.45; }
  .drop { border:2px dashed var(--a); border-radius:18px; padding:26px; text-align:center; background:var(--bg); cursor:pointer; transition:background .2s; }
  .drop.over { background:#fff; }
  .drop b { display:block; font-size:16px; color:var(--t); margin:8px 0 4px; }
  .drop small { color:var(--muted); }
  .or { text-align:center; color:var(--muted); font-size:13px; margin:14px 0; }
  .row { display:flex; gap:10px; align-items:center; }
  input[type=text], select { font:600 14px Montserrat, Arial; padding:11px 12px; border:2px solid var(--line); border-radius:12px; outline:none; color:var(--ink); background:#fff; }
  input[type=text] { flex:1; }
  .btn { border:none; border-radius:14px; padding:13px 18px; font:700 15px Montserrat, Arial; cursor:pointer; }
  .btn.main { background:var(--p); color:#fff; } .btn.main:disabled { opacity:.4; }
  .btn.soft { background:var(--bg); color:var(--t); }
  .btn.ghost { background:none; color:var(--muted); }
  .banks { font-size:12.5px; color:var(--muted); line-height:1.6; margin-top:14px; }
  .banks b { color:var(--ink); }
  .err { color:#D9534F; font-weight:600; font-size:13px; min-height:18px; margin-top:8px; }
  .cards { display:flex; gap:10px; margin-bottom:10px; }
  .card { flex:1; border-radius:14px; background:var(--bg); padding:10px 12px; }
  .card b { display:block; font-size:18px; color:var(--t); }
  .card span { font-size:12px; color:var(--muted); font-weight:600; }
  .tbl { flex:1; overflow:auto; border:1px solid var(--line); border-radius:12px; }
  table { width:100%; border-collapse:collapse; font-size:12.5px; }
  th { position:sticky; top:0; background:#F7F9FB; text-align:left; padding:8px; font-size:11.5px; color:var(--muted); border-bottom:1px solid var(--line); z-index:1; }
  td { padding:5px 8px; border-bottom:1px solid #F1F3F6; vertical-align:middle; }
  tr.off td { color:#B5C0CB; }
  tr.off select { opacity:.5; }
  td.amt { text-align:right; font-weight:700; white-space:nowrap; }
  td.amt.in { color:#2E7D32; } td.amt.out { color:#C0392B; }
  td select { padding:5px 6px; font-size:12px; border-radius:8px; border-width:1px; max-width:170px; }
  .tag { display:inline-block; font-size:10.5px; font-weight:700; border-radius:7px; padding:2px 6px; background:#FFF3D6; color:#8A6100; }
  .tag.l { background:var(--bg); color:var(--t); }
  .desc { max-width:260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .foot { display:flex; gap:10px; align-items:center; margin-top:12px; }
  .foot .sp { flex:1; font-size:12.5px; color:var(--muted); }
  .spin { width:48px; height:48px; border-radius:50%; border:5px solid var(--bg); border-top-color:var(--p); animation:r 1s linear infinite; margin:40px auto 16px; }
  @keyframes r { to { transform:rotate(360deg); } }
  .center { text-align:center; }
  .big { font-size:46px; margin:30px 0 10px; }
</style></head>
<body><div class="wrap">

  <section class="screen on" id="s1">
    <h1>Імпорт виписки з банку</h1>
    <p class="sub">Завантаж виписку — я сам розкладу операції по категоріях. Перед записом покажу, що куди піде, і ти зможеш усе виправити.</p>
    <label class="drop" id="drop">
      <div style="font-size:34px">📄</div>
      <b>Перетягни сюди CSV-файл або натисни, щоб вибрати</b>
      <small>Monobank, ПриватБанк, ПУМБ, Sense, Abank та інші — головне, щоб були колонки «Дата» і «Сума»</small>
      <input type="file" id="file" accept=".csv,text/csv,text/plain" style="display:none">
    </label>
    <div class="or">— або —</div>
    <div class="row"><input type="text" id="link" placeholder="Посилання на Google-таблицю з випискою"><button class="btn soft" onclick="fromLink()">Відкрити</button></div>
    <div class="banks">
      <b>Де взяти виписку:</b><br>
      Monobank: застосунок → картка → «Виписка» → період → формат <b>CSV</b> → надішли собі файл.<br>
      ПриватБанк: Приват24 → «Виписки» → період → «Завантажити» → <b>CSV</b> або Excel (Excel відкрий у Google Таблицях і встав посилання).<br>
      Інший банк — будь-який CSV або таблиця з колонками дати й суми.
    </div>
    <div class="err" id="err1"></div>
  </section>

  <section class="screen" id="s2">
    <div class="spin"></div><p class="sub center" id="loadTxt">Читаю виписку…</p>
  </section>

  <section class="screen" id="s3">
    <h1>Перевір, що куди піде</h1>
    <div class="cards">
      <div class="card"><span>Витрати</span><b id="cOut"></b></div>
      <div class="card"><span>Доходи</span><b id="cIn"></b></div>
      <div class="card"><span>Пропущено</span><b id="cSkip"></b></div>
    </div>
    <p class="sub" style="margin-bottom:8px">Знята галочка — рядок не імпортується (перекази між своїми рахунками й скасовані покупки я пропустив сам). Змінена категорія — запамʼятаю для наступних виписок.</p>
    <div class="tbl"><table><thead><tr><th></th><th>Дата</th><th>Опис</th><th style="text-align:right">Сума</th><th>Тип</th><th>Категорія</th><th></th></tr></thead><tbody id="tb"></tbody></table></div>
    <div class="foot">
      <span style="font-size:13px;font-weight:600">Рахунок:</span><select id="acc"></select>
      <div class="sp" id="selInfo"></div>
      <button class="btn ghost" onclick="go(1)">← Інший файл</button>
      <button class="btn main" id="imp" onclick="doImport()">Імпортувати</button>
    </div>
  </section>

  <section class="screen" id="s4">
    <div class="center"><div class="big" id="doneEm">✅</div><h1 id="doneH"></h1><p class="sub" id="doneP"></p>
    <button class="btn main" onclick="google.script.host.close()">Готово</button></div>
  </section>
</div>
<script>
  const D = __DATA__;
  const TYPES = ['Витрата', 'Дохід', 'Борг', 'Заощадження'];
  let P = null;
  const $ = id => document.getElementById(id);
  const esc = s => String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const fmt = n => { const r = Math.round(n * 100) / 100; const [i, f] = r.toFixed(2).split('.');
    return i.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + (f !== '00' ? ',' + f : '') + ' грн'; };
  (function theme() { const t = D.theme, r = document.documentElement.style; r.setProperty('--p', t.primary); r.setProperty('--a', t.accent); r.setProperty('--bg', t.headerBg); r.setProperty('--t', t.headerTxt); })();
  function go(n) { document.querySelectorAll('.screen').forEach(e => e.classList.remove('on')); $('s' + n).classList.add('on'); }
  function fail(e) { go(1); $('err1').textContent = (e && e.message) || String(e); }

  // ── CSV
  function parseCsv(text) {
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/)[0] || '';
    const delim = (first.split(';').length > first.split(',').length) ? ';' : (first.split('\t').length > first.split(',').length ? '\t' : ',');
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(r => r.some(c => String(c).trim() !== ''));
  }
  function readFile(f) {
    const rd = new FileReader();
    rd.onload = () => {
      let t = rd.result;
      if (/Ð|Ã|�/.test(t.slice(0, 2000))) { const r2 = new FileReader(); r2.onload = () => preview(parseCsv(r2.result)); r2.readAsText(f, 'windows-1251'); return; }
      preview(parseCsv(t));
    };
    rd.readAsText(f, 'utf-8');
  }
  $('file').addEventListener('change', e => { if (e.target.files[0]) readFile(e.target.files[0]); });
  const dz = $('drop');
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) readFile(f); });
  function fromLink() {
    const u = $('link').value.trim(); if (!u) return;
    go(2); $('loadTxt').textContent = 'Відкриваю таблицю…';
    google.script.run.withSuccessHandler(preview).withFailureHandler(fail).statementRowsFromSheet(u);
  }
  function preview(rows) {
    go(2); $('loadTxt').textContent = 'Розкладаю по категоріях…';
    google.script.run.withSuccessHandler(show).withFailureHandler(fail).statementPreview(rows);
  }

  // ── Попередній перегляд
  function opts(list, sel) {
    const l = list.slice(); if (sel && l.indexOf(sel) === -1) l.unshift(sel);
    return l.map(c => '<option' + (c === sel ? ' selected' : '') + '>' + esc(c) + '</option>').join('');
  }
  function show(res) {
    P = res;
    P.items.forEach(x => { x.orig = x.type + '|' + x.category; });
    $('acc').innerHTML = '<option value="">основний</option>' + D.accounts.map(a => '<option>' + esc(a) + '</option>').join('');
    render(); go(3);
  }
  function render() {
    $('tb').innerHTML = P.items.map((x, i) =>
      '<tr class="' + (x.include ? '' : 'off') + '">' +
      '<td><input type="checkbox"' + (x.include ? ' checked' : '') + ' onchange="P.items[' + i + '].include=this.checked;render()"></td>' +
      '<td>' + x.date.split('-').reverse().join('.') + '</td>' +
      '<td><div class="desc" title="' + esc(x.desc) + '">' + esc(x.desc) + '</div></td>' +
      '<td class="amt ' + (x.sign > 0 ? 'in' : 'out') + '">' + (x.sign > 0 ? '+' : '−') + fmt(x.amount) + '</td>' +
      '<td><select onchange="setType(' + i + ',this.value)">' + opts(TYPES, x.type) + '</select></td>' +
      '<td><select onchange="P.items[' + i + '].category=this.value;render()">' + opts(P.lists[x.type] || [], x.category) + '</select></td>' +
      '<td>' + (x.reason ? '<span class="tag">' + esc(x.reason) + '</span>' : (x.learned ? '<span class="tag l">як минулого разу</span>' : '')) + '</td></tr>').join('');
    const on = P.items.filter(x => x.include);
    const inc = on.filter(x => x.type === 'Дохід'), out = on.filter(x => x.type !== 'Дохід');
    const sum = a => a.reduce((s, x) => s + x.amount, 0);
    $('cOut').textContent = out.length + ' · ' + fmt(sum(out));
    $('cIn').textContent = inc.length + ' · ' + fmt(sum(inc));
    $('cSkip').textContent = P.items.length - on.length;
    $('selInfo').textContent = 'Буде додано: ' + on.length;
    $('imp').disabled = !on.length;
    $('imp').textContent = 'Імпортувати ' + on.length;
  }
  function setType(i, t) {
    const x = P.items[i]; x.type = t;
    const l = P.lists[t] || []; if (l.indexOf(x.category) === -1) x.category = l[0] || '';
    render();
  }
  function doImport() {
    const items = P.items.filter(x => x.include).map(x => ({ date: x.date, type: x.type, amount: x.amount, category: x.category,
      desc: x.desc, mcc: x.mcc, key: x.key, learn: (x.type + '|' + x.category) !== x.orig }));
    go(2); $('loadTxt').textContent = 'Записую в «Операції»…';
    google.script.run.withSuccessHandler(r => {
      $('doneH').textContent = 'Додано операцій: ' + r.added;
      $('doneP').textContent = (r.created.length ? 'Нові категорії в «Налаштуваннях»: ' + r.created.join(', ') + '. Можеш поставити їм бюджет. ' : '') +
        'Дашборд уже перерахувався.';
      go(4);
    }).withFailureHandler(fail).statementImport(items, $('acc').value);
  }
</script></body></html>`;

// ─────────────────────────────────────────────────────────────────────────────
// API ДЛЯ ЗАСТОСУНКУ (PWA)
// Веб-застосунок Apps Script: Розгорнути → Нове розгортання → Веб-застосунок
// («Виконувати від мого імені», доступ «Усі»). Кожен запит перевіряється ключем.
// ─────────────────────────────────────────────────────────────────────────────
const APP_URL = 'https://oleksii-rad-personal.github.io/Test/';
const APP = { propKey: 'BT_APP_KEY', propApi: 'BT_APP_API', source: 'застосунок', recent: 40,
  blocks: ['Inc', 'Reg', 'Subs', 'Var', 'Debt', 'Sav'] };

function doGet() {
  return ContentService.createTextOutput('Трекер бюджету: API працює ✓ Повернись у таблицю й встав цю адресу в «Підключити застосунок».');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const props = PropertiesService.getDocumentProperties();
    const key = props.getProperty(APP.propKey);
    if (!key || req.key !== key) throw new Error('Невірний ключ. Відкрий у таблиці «Підключити застосунок» і відскануй QR ще раз.');
    out = { ok: true, data: appDispatch_(req) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function appSpreadsheet_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) ss = SpreadsheetApp.openById(PropertiesService.getDocumentProperties().getProperty('BT_DOC_ID'));
  return ss;
}

function appDispatch_(req) {
  const ss = appSpreadsheet_();
  if (!ss.getSheetByName(SHEETS.ops) || !ss.getRangeByName(NR + 'D_Free')) {
    throw new Error('Трекер у таблиці ще не побудовано: меню «Трекер» → «Побудувати трекер з нуля».');
  }
  applyProfileTheme_();
  switch (req.action) {
    case 'state': return appState_(ss);
    case 'addOp': appAddOp_(ss, req.op || {}); return appState_(ss);
    case 'deleteOp': appDeleteOp_(ss, req.row, req.op || {}); return appState_(ss);
    case 'getSettings': return appGetSettings_(ss);
    case 'saveSettings': appSaveSettings_(ss, req.block, req.rows || []); return appGetSettings_(ss);
    case 'quizHtml': return { html: quizHtml_() };
    case 'quiz': quizSubmit(req.profile || {}, 'save', req.apply !== false); return appState_(ss);
    default: throw new Error('Невідома дія: ' + req.action);
  }
}

/** Усе для головного екрана одним запитом. Цифри — з формул дашборду. */
function appState_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const iso = v => (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v === null || v === undefined ? '' : v));
  const period = ss.getRangeByName(NR + 'D_Period');
  if (period && period.getValue() !== REF.current) { period.setValue(REF.current); SpreadsheetApp.flush(); }
  const g = k => { const r = ss.getRangeByName(NR + k); return r ? r.getValue() : ''; };
  const num = v => (typeof v === 'number' ? Math.round(v * 100) / 100 : 0);
  const block = (key, f) => {
    const r = ss.getRangeByName(NR + 'DB_' + key);
    return r ? r.getValues().filter(v => v[0] !== '' && v[0] !== null).map(f) : [];
  };
  const p = profile_();
  const lists = {};
  Object.keys(TYPE_LIST).forEach(t => { lists[t] = listValues_(ss, TYPE_LIST[t]); });

  const sh = ss.getSheetByName(SHEETS.ops);
  const ops = sh.getRange(2, 1, DATA_ROWS - 1, OPS_HEADERS.length).getValues()
    .map((r, i) => ({ row: i + 2, date: iso(r[OPS.date - 1]), type: r[OPS.type - 1], amount: Number(r[OPS.amount - 1]) || 0,
      category: r[OPS.category - 1], account: r[OPS.account - 1], comment: r[OPS.comment - 1], source: r[OPS.source - 1] }))
    .filter(o => o.amount > 0)
    .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.row - a.row))
    .slice(0, APP.recent);

  return {
    profile: { name: p.name || '', gender: p.gender || '' },
    theme: THEMES[p.gender === 'f' ? 'f' : 'm'],
    today: todayIso_(ss),
    free: num(g('D_Free')), status: g('C_Status'), isCurrent: g('C_IsCur') === true,
    nextPay: iso(g('C_Next')), days: num(g('C_Days')), limit: num(g('C_Limit')), avg: num(g('C_Avg')),
    cycle: { start: iso(g('C_Start')), end: iso(g('C_End')) },
    calc: dashCalcLines_(() => '').map(c => ({ label: c.label, value: num(g(c.key)), total: !!c.bold })),
    budgets: block('var', v => ({ name: v[0], budget: num(v[1]), fact: num(v[2]) })),
    regular: block('reg', v => ({ name: v[0], amount: num(v[1]), date: iso(v[2]), paid: num(v[3]), done: v[4] === '✓', left: num(v[6]) })),
    subs: block('subs', v => ({ name: v[0], amount: num(v[1]), date: iso(v[2]), paid: num(v[3]), done: v[4] === '✓', left: num(v[6]) })),
    types: REF.types, lists: lists, accounts: listValues_(ss, 'L_Acc'),
    ops: ops, sheetUrl: ss.getUrl(),
  };
}

function appAddOp_(ss, op) {
  const amount = Math.round(Number(String(op.amount).replace(',', '.')) * 100) / 100;
  if (!(amount > 0)) throw new Error('Сума має бути більшою за нуль.');
  const type = REF.types.indexOf(op.type) !== -1 ? op.type : TYPES.expense;
  const date = /^\d{4}-\d\d-\d\d$/.test(op.date || '') ? op.date : todayIso_(ss);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(SHEETS.ops);
    // Повторна відправка з офлайн-черги не має дублювати рядок.
    if (op.id) {
      const tags = sh.getRange(2, OPS.source, DATA_ROWS - 1, 1).getValues();
      if (tags.some(v => v[0] === APP.source + ' ' + op.id)) return;
    }
    const row = firstFreeOpsRow_(sh);
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
    writeOpsRows_(sh, row, [[date, type, amount, String(op.category || '').slice(0, 60), String(op.account || '').slice(0, 40),
      String(op.comment || '').slice(0, 200), APP.source + (op.id ? ' ' + String(op.id).slice(0, 40) : '')]]);
    SpreadsheetApp.flush();
    syncCategoriesFromOps_(ss);
  } finally {
    lock.releaseLock();
  }
}

/** Видаляє операцію, внесену із застосунку (рядок звіряється за сумою й міткою). */
function appDeleteOp_(ss, row, op) {
  const sh = ss.getSheetByName(SHEETS.ops);
  const n = DATA_ROWS - 1;
  const rng = sh.getRange(2, 1, n, OPS_HEADERS.length);
  const vals = rng.getValues();
  const i = Number(row) - 2;
  const r = vals[i];
  if (!r || String(r[OPS.source - 1]).indexOf(APP.source) !== 0 || Math.abs(Number(r[OPS.amount - 1]) - Number(op.amount)) > 0.005) {
    throw new Error('Цю операцію можна видалити лише в таблиці.');
  }
  vals.splice(i, 1);
  vals.push(new Array(OPS_HEADERS.length).fill(''));
  writeOpsRows_(sh, 2, vals);
  SpreadsheetApp.flush();
}

function appGetSettings_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const iso = v => (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v);
  const options = {};
  const blocks = APP.blocks.map(key => {
    const b = SETTINGS_BLOCKS.filter(x => x.key === key)[0];
    const cols = b.cols.filter(c => c.t !== 'f').map(c => {
      let t = c.t, list = '';
      if (t.indexOf('list:') === 0) {
        list = t.split(':')[1]; t = 'list';
        if (!options[list]) options[list] = listValues_(ss, list === 'L_Freq' ? 'L_FreqName' : list);
      }
      return { key: c.key, h: c.h, t: t, list: list };
    });
    const data = {};
    cols.forEach(c => { data[c.key] = ss.getRangeByName(NR + key + '_' + c.key).getValues().map(v => iso(v[0])); });
    const rows = [];
    for (let i = 0; i < b.rows; i++) {
      if (String(data.Name[i] || '').trim() === '') continue;
      const o = {};
      cols.forEach(c => { o[c.key] = data[c.key][i]; });
      rows.push(o);
    }
    return { key: key, title: b.title, hint: b.hint, max: b.rows, cols: cols, rows: rows };
  });
  return { blocks: blocks, options: options };
}

function appSaveSettings_(ss, key, rows) {
  if (APP.blocks.indexOf(key) === -1) throw new Error('Невідомий блок налаштувань.');
  const b = SETTINGS_BLOCKS.filter(x => x.key === key)[0];
  const clean = rows.filter(r => r && String(r.Name || '').trim() !== '').slice(0, b.rows).map(r => {
    const o = {};
    b.cols.forEach(c => {
      if (c.t === 'f') return;
      let v = r[c.key];
      if (v === undefined || v === null) v = c.t === 'check' ? false : '';
      if (c.t === 'check') v = v === true;
      else if (['money', 'pct', 'int', 'day'].indexOf(c.t) !== -1) v = v === '' ? '' : Number(v);
      else v = String(v).slice(0, 80);
      o[c.key] = v;
    });
    return o;
  });
  if (key === 'Inc' && clean.length && clean.filter(r => r.Main).length !== 1) {
    clean.forEach((r, i) => { r.Main = i === 0; });
  }
  setBlockRows_(ss, key, clean);
  SpreadsheetApp.flush();
  syncCategoriesFromOps_(ss);
}

// ── Підключення застосунку з таблиці
function appKey_(reset) {
  const props = PropertiesService.getDocumentProperties();
  let key = props.getProperty(APP.propKey);
  if (!key || reset) {
    key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
    props.setProperty(APP.propKey, key);
  }
  return key;
}

function appLink_() {
  const api = PropertiesService.getDocumentProperties().getProperty(APP.propApi) || '';
  return api ? APP_URL + '?api=' + encodeURIComponent(api) + '&key=' + appKey_(false) : '';
}

/** З діалогу: зберегти адресу веб-застосунку (reset — заодно новий ключ). */
function appSaveApi(url, reset) {
  const u = String(url || '').trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[-\w]+\/exec$/.test(u)) {
    throw new Error('Адреса має виглядати як https://script.google.com/macros/s/…/exec');
  }
  PropertiesService.getDocumentProperties().setProperty(APP.propApi, u);
  if (reset) appKey_(true);
  return appLink_();
}

/** Меню: підключити застосунок на телефоні. */
function showConnectApp() {
  ensureOwnDocument_();
  applyProfileTheme_();
  appKey_(false);
  const props = PropertiesService.getDocumentProperties();
  let api = props.getProperty(APP.propApi) || '';
  if (!api) { try { const u = ScriptApp.getService().getUrl(); if (u && /\/exec$/.test(u)) api = u; } catch (e) { /* не розгорнуто */ } }
  const data = { api: api, link: appLink_(), theme: THEME };
  const html = HtmlService.createHtmlOutput(CONNECT_HTML
    .replace('__QRLIB__', () => QR_LIB)
    .replace('__DATA__', () => JSON.stringify(data).replace(/</g, '\\u003c'))).setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 Застосунок на телефон');
}

const CONNECT_HTML = String.raw`<!DOCTYPE html><html lang="uk"><head><meta charset="utf-8">
<style>
  :root { --p:#4F86B5; --bg:#E5F1FA; --t:#2C5677; --muted:#8A9BAA; }
  body { font-family:Arial, sans-serif; color:#2E3A46; margin:0; padding:16px 20px; font-size:14px; line-height:1.5; }
  h2 { color:var(--t); font-size:17px; margin:14px 0 6px; }
  ol { margin:0 0 8px; padding-left:20px; } li { margin-bottom:4px; }
  input { width:100%; box-sizing:border-box; padding:10px 12px; border:2px solid #E7EBF0; border-radius:10px; font-size:13px; }
  .btn { border:none; border-radius:10px; padding:10px 16px; font-weight:bold; font-size:14px; cursor:pointer; background:var(--p); color:#fff; margin-top:8px; }
  .btn.soft { background:var(--bg); color:var(--t); }
  .box { background:var(--bg); border-radius:14px; padding:12px 14px; margin-top:10px; }
  .qr { text-align:center; } .qr svg { width:190px; height:190px; }
  .link { font-size:11px; word-break:break-all; color:var(--muted); }
  .err { color:#D9534F; font-weight:bold; min-height:18px; }
  .warn { font-size:12px; color:var(--muted); margin-top:8px; }
</style></head><body>
<h2>1. Увімкни доступ для застосунку (один раз)</h2>
<ol>
  <li>«Розширення» → «Apps Script» → справа вгорі <b>«Розгорнути» → «Нове розгортання»</b>.</li>
  <li>Тип (шестерня) — <b>«Веб-застосунок»</b>. «Виконувати як» — <b>я</b>. «Хто має доступ» — <b>усі</b>.</li>
  <li>«Розгорнути» → дозволь доступ → скопіюй <b>URL веб-застосунку</b> (закінчується на /exec).</li>
</ol>
<input id="api" placeholder="https://script.google.com/macros/s/…/exec">
<div class="err" id="err"></div>
<button class="btn" onclick="save(false)">Зберегти адресу</button>
<div id="step2" style="display:none">
  <h2>2. Відкрий застосунок на телефоні</h2>
  <div class="box"><div class="qr" id="qr"></div>
    <div>Наведи камеру телефона на QR-код. Застосунок відкриється вже підключеним — додай його на головний екран («Поділитися» → «На початковий екран»).</div>
    <div class="link" id="link"></div>
    <button class="btn soft" onclick="copy()">Скопіювати посилання</button>
    <button class="btn soft" onclick="save(true)">Скинути ключ</button>
  </div>
  <div class="warn">Посилання містить ключ доступу до твого бюджету — не публікуй його. Якщо воно потрапило не туди, натисни «Скинути ключ» і відскануй новий QR.</div>
</div>
<script>
__QRLIB__
</script>
<script>
  const D = __DATA__;
  (function () { const r = document.documentElement.style, t = D.theme; r.setProperty('--p', t.primary); r.setProperty('--bg', t.headerBg); r.setProperty('--t', t.headerTxt); })();
  const $ = id => document.getElementById(id);
  $('api').value = D.api || '';
  function show(link) {
    if (!link) return;
    const q = qrcode(0, 'L'); q.addData(link); q.make();
    $('qr').innerHTML = q.createSvgTag({ cellSize: 3, margin: 2, scalable: true });
    $('link').textContent = link; $('step2').style.display = 'block'; D.link = link;
  }
  function save(reset) {
    $('err').textContent = '';
    google.script.run.withSuccessHandler(show).withFailureHandler(e => { $('err').textContent = (e && e.message) || String(e); }).appSaveApi($('api').value, reset);
  }
  function copy() { const t = document.createElement('textarea'); t.value = D.link; document.body.appendChild(t); t.select(); document.execCommand('copy'); t.remove(); }
  show(D.link);
</script></body></html>`;

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

// Кнопки для «Старту» (PNG, base64): дві теми — f (лілово-рожева), m (блакитна).
const BUTTON_PNG = {"f": {"quiz": "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", "guide": "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", "import": "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", "form": "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"}, "m": {"quiz": "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", "guide": "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", "import": "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", "form": "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"}};
