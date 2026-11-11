/* Трекер бюджету — застосунок (PWA) над Google-таблицею.
   Дані живуть у таблиці користувача; сюди приходять готові цифри з її формул. */
(function () {
  'use strict';
  const LS = { conn: 'bt_conn', state: 'bt_state', outbox: 'bt_outbox' };
  const $ = s => document.querySelector(s);
  const main = $('#main');
  const load = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) || d; } catch (e) { return d; } };
  const save = (k, v) => localStorage.setItem(k, JSON.stringify(v));
  const esc = s => String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const money = n => { const r = Math.round((Number(n) || 0) * 100) / 100; const p = r.toFixed(2).split('.');
    return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + (p[1] !== '00' ? ',' + p[1] : '') + ' грн'; };
  const dm = iso => (iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) : '');
  const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
  const standalone = window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;

  let conn = load(LS.conn, null);
  let st = load(LS.state, null);
  let outbox = load(LS.outbox, []);
  let view = 'home';
  let busy = false;
  const draft = { amount: '', type: 'Витрата', category: '', account: '', date: '', comment: '', more: false };
  let settings = null;
  const G = (f, m) => (st && st.profile && st.profile.gender === 'm' ? m : f);

  // ── Підключення: параметри з посилання (QR у таблиці)
  function readLink(href) {
    try {
      const u = new URL(href, location.href);
      const q = new URLSearchParams(u.search || u.hash.replace(/^#/, '?'));
      const api = q.get('api'), key = q.get('key');
      if (api && key && /^https:\/\/script\.google\.com\/macros\/s\/[-\w]+\/exec$/.test(api)) return { api: api, key: key };
    } catch (e) { /* не посилання */ }
    return null;
  }
  const fromUrl = readLink(location.href);
  if (fromUrl) { conn = fromUrl; save(LS.conn, conn); }
  // iPhone зберігає дані встановленого застосунку окремо від Safari, тому іконка має памʼятати
  // адресу з параметрами: на iOS — маніфест без start_url.
  if (isIos && $('#manifest')) $('#manifest').setAttribute('href', 'manifest.ios.webmanifest');

  async function api(action, extra) {
    const res = await fetch(conn.api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ key: conn.key, action: action }, extra || {})) });
    const j = await res.json();
    if (!j.ok) { const e = new Error(j.error || 'Помилка'); e.api = true; throw e; }
    return j.data;
  }
  function setState(s) { st = s; save(LS.state, st); theme(); }
  function theme() {
    if (!st || !st.theme) return;
    const r = document.documentElement.style, t = st.theme;
    r.setProperty('--p', t.primary); r.setProperty('--a', t.accent); r.setProperty('--bg', t.headerBg); r.setProperty('--t', t.headerTxt);
    document.querySelector('meta[name=theme-color]').setAttribute('content', t.primary);
  }
  let toastT = null;
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, msg.length > 60 ? 7000 : 3200); }

  // Інтернет є, а таблиця не відповідає — найчастіше це доступ у розгортанні веб-застосунку.
  function offlineMsg() {
    return navigator.onLine === false ? 'Немає інтернету — показую збережені дані'
      : 'Таблиця не відповідає. У таблиці: «Трекер» → «📱 Підключити застосунок» → «Перевірити звʼязок»';
  }
  async function refresh(silent) {
    if (!conn || busy) return;
    busy = true;
    try { await flush(); setState(await api('state')); if (view !== 'add' && view !== 'settings') render(); }
    catch (e) { if (!silent || e.api) toast(e.api ? e.message : offlineMsg()); }
    finally { busy = false; }
  }
  // Офлайн-черга: операції, які ще не дійшли до таблиці.
  async function flush() {
    while (outbox.length) {
      const s = await api('addOp', { op: outbox[0] });
      outbox.shift(); save(LS.outbox, outbox); setState(s);
    }
  }

  // ── Екрани
  function render() {
    $('#nav').hidden = !conn;
    document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('on', b.dataset.view === view));
    if (!conn) return renderConnect();
    if (!st) { main.innerHTML = '<div class="spin"></div><p class="sub" style="text-align:center">Завантажую твій бюджет…</p>'; return; }
    ({ home: renderHome, ops: renderOps, add: renderAdd, settings: renderSettings }[view] || renderHome)();
    window.scrollTo(0, 0);
  }

  function renderConnect() {
    main.innerHTML = '<h1>Трекер бюджету</h1>' +
      '<p class="sub">Застосунок працює з твоєю Google-таблицею трекера. Підключи його один раз:</p>' +
      '<ol class="steps"><li>Відкрий таблицю на компʼютері → меню <b>«Трекер» → «📱 Підключити застосунок на телефон»</b>.</li>' +
      '<li>Виконай кроки з вікна і <b>наведи камеру телефона на QR-код</b> — застосунок відкриється вже підключеним.</li></ol>' +
      (isIos && standalone ? '<div class="banner">Іконка на iPhone не бачить даних Safari. Відкрий застосунок у Safari → «Налаштування» → «📋 Скопіювати посилання для іконки» і встав його сюди.</div>' : '') +
      '<label class="f">Або встав посилання з того вікна</label><input id="lnk" placeholder="https://…?api=…&key=…" autocomplete="off">' +
      '<div style="height:10px"></div><button class="btn" id="go">Підключити</button>';
    $('#go').onclick = () => {
      const c = readLink($('#lnk').value.trim());
      if (!c) return toast('Не схоже на посилання з таблиці');
      conn = c; save(LS.conn, conn);
      history.replaceState(null, '', location.pathname + '?api=' + encodeURIComponent(c.api) + '&key=' + c.key);
      render(); refresh();
    };
  }

  function renderHome() {
    const s = st;
    const cls = s.status === 'ok' ? 'ok' : s.status === 'bad' ? 'bad' : '';
    const pend = outbox.reduce((a, o) => a + (o.type === 'Дохід' ? o.amount : -o.amount), 0);
    let h = '<h1>' + (s.profile.name ? 'Привіт, ' + esc(s.profile.name) + '!' : 'Мій бюджет') + '</h1>';
    if (outbox.length) h += '<div class="banner">Чекає на звʼязок: ' + outbox.length + ' операц. Цифра вже з їх урахуванням.</div>';
    h += '<div class="hero ' + cls + '"><div class="lbl">Можу витратити до наступної виплати</div>';
    if (s.status === 'empty' && !outbox.length) {
      h += '<div class="num" style="font-size:26px">Внеси перший дохід</div><div class="hint">Натисни ＋ і додай зарплату — тут зʼявиться твоя цифра</div>';
    } else {
      h += '<div class="num">' + money(s.free + pend) + '</div><div class="hint">' +
        (s.free + pend < 0 ? (s.expect ? '⚠ Мінус, бо дохід циклу ще не внесено. Зарплата вже прийшла? Натисни ＋ → Дохід' : '⚠ Перевитрата — пригальмуй до виплати')
          : s.status === 'bad' ? 'Менше, ніж ти зазвичай витрачаєш (' + money(s.avg) + ' / день)' : 'Ти в нормі 💚') + '</div>';
    }
    if (s.expect) h += '<div class="hint" style="font-size:12px;font-weight:600;opacity:.75;margin-top:6px">' + esc(s.expect) + '</div>';
    h += '</div><div class="pills"><div class="pill"><b>' + money(Math.round(s.days ? (s.free + pend) / Math.max(1, s.days) : s.limit)) + '</b><span>на день</span></div>' +
      '<div class="pill"><b>' + (s.days || '—') + '</b><span>днів до виплати</span></div>' +
      '<div class="pill"><b>' + (dm(s.nextPay) || '—') + '</b><span>наступна виплата</span></div></div>';
    h += '<details><summary>Як рахується цифра</summary><div class="calc">' +
      s.calc.map(c => '<div class="' + (c.total ? 'tot' : '') + '"><span>' + esc(c.label) + '</span><span>' + money(c.value) + '</span></div>').join('') + '</div></details>';

    const bud = s.budgets.filter(b => b.budget > 0 || b.fact > 0);
    if (bud.length) {
      h += '<h2>Бюджети</h2>' + bud.map(b => {
        const p = b.budget > 0 ? b.fact / b.budget : (b.fact > 0 ? 1 : 0);
        return '<div class="bar"><div class="top"><b>' + esc(b.name) + '</b><span>' + money(b.fact) + (b.budget ? ' з ' + money(b.budget) : '') + '</span></div>' +
          '<div class="track"><i class="' + (p > 1 ? 'over' : p > 0.8 ? 'warn' : '') + '" style="width:' + Math.min(100, Math.round(p * 100)) + '%"></i></div></div>';
      }).join('');
    }
    const pays = s.regular.concat(s.subs).filter(x => x.amount > 0 && (x.left > 0 || x.done));
    if (pays.length) {
      h += '<h2>Платежі цього циклу</h2>' + pays.map(x => '<div class="item"><div class="g">' + esc(x.name) +
        '<small>' + (x.date ? 'до ' + dm(x.date) : '') + '</small></div><div class="amt">' + money(x.amount) + '</div>' +
        (x.done ? '<span class="tag done">оплачено</span>' : '<span class="tag">ще ' + money(x.left) + '</span>') + '</div>').join('');
    }
    if (!standalone) h += '<details><summary>📲 Додати на головний екран</summary><ol class="steps">' +
      (isIos ? '<li>Відкрий цю сторінку в <b>Safari</b>.</li><li>Натисни <b>«Поділитися»</b> (квадрат зі стрілкою).</li><li>Вибери <b>«На початковий екран»</b> → «Додати».</li>'
        : '<li>Відкрий меню браузера <b>⋮</b>.</li><li>Вибери <b>«Додати на головний екран»</b> або «Встановити застосунок».</li>') + '</ol></details>';
    main.innerHTML = h;
  }

  function renderOps() {
    const list = outbox.map(o => Object.assign({ pending: true }, o)).concat(st.ops);
    if (!list.length) { main.innerHTML = '<h1>Операції</h1><div class="empty">Ще немає операцій. Натисни ＋, щоб додати першу.</div>'; return; }
    let h = '<h1>Операції</h1>', day = null;
    list.forEach((o, i) => {
      const d = o.date || st.today;
      if (d !== day) { day = d; h += '<div class="day">' + (d === st.today ? 'Сьогодні' : dm(d) + '.' + d.slice(0, 4)) + '</div>'; }
      const mine = !o.pending && String(o.source || '').indexOf('застосунок') === 0;
      h += '<div class="item"><div class="g">' + esc(o.category || o.type) + (o.pending ? '<span class="tag">чекає</span>' : '') +
        '<small>' + esc([o.type !== 'Витрата' ? o.type : '', o.account, o.comment].filter(Boolean).join(' · ')) + '</small></div>' +
        '<div class="amt ' + (o.type === 'Дохід' ? 'in' : 'out') + '">' + (o.type === 'Дохід' ? '+' : '−') + money(o.amount) + '</div>' +
        (mine ? '<button class="del" data-i="' + (i - outbox.length) + '" aria-label="Видалити">✕</button>' : '') + '</div>';
    });
    h += '<p class="sub" style="margin-top:14px">Тут останні ' + st.ops.length + ' операцій. Видалити можна внесені із застосунку; решту — у таблиці.</p>';
    main.innerHTML = h;
    main.querySelectorAll('.del').forEach(b => { b.onclick = async () => {
      const o = st.ops[Number(b.dataset.i)];
      if (!confirm('Видалити «' + (o.category || o.type) + '» ' + money(o.amount) + '?')) return;
      try { setState(await api('deleteOp', { row: o.row, op: { amount: o.amount } })); render(); toast('Видалено'); }
      catch (e) { toast(e.api ? e.message : 'Немає звʼязку — спробуй пізніше'); }
    }; });
  }

  function renderAdd() {
    const cats = (st.lists[draft.type] || []);
    if (cats.indexOf(draft.category) === -1) draft.category = '';
    let h = '<h1>Нова операція</h1><div class="amount">' + (draft.amount || '0') + ' <small>грн</small></div>' +
      '<div class="seg">' + st.types.map(t => '<button data-t="' + esc(t) + '" class="' + (t === draft.type ? 'on' : '') + '">' + esc(t) + '</button>').join('') + '</div>' +
      '<div class="chips">' + (cats.length ? cats.map(c => '<button class="chip' + (c === draft.category ? ' on' : '') + '" data-c="' + esc(c) + '">' + esc(c) + '</button>').join('')
        : '<span class="sub">Для цього типу ще немає категорій — додай їх у «Налаштуваннях».</span>') + '</div>' +
      '<div class="pad">' + ['1', '2', '3', '4', '5', '6', '7', '8', '9', ',', '0', '⌫'].map(k => '<button data-k="' + k + '">' + k + '</button>').join('') + '</div>';
    if (draft.more) {
      h += '<div class="row2"><div><label class="f">Дата</label><input type="date" id="dDate" value="' + esc(draft.date || st.today) + '" max="' + st.today + '"></div>' +
        '<div><label class="f">Рахунок</label><select id="dAcc"><option value="">основний</option>' +
        st.accounts.map(a => '<option' + (a === draft.account ? ' selected' : '') + '>' + esc(a) + '</option>').join('') + '</select></div></div>' +
        '<label class="f">Коментар</label><input id="dCom" value="' + esc(draft.comment) + '" placeholder="необовʼязково"><div style="height:10px"></div>';
    } else h += '<button class="btn ghost" id="more">Дата, рахунок, коментар</button>';
    h += '<button class="btn" id="saveOp"' + (Number(String(draft.amount).replace(',', '.')) > 0 && (draft.category || !cats.length) ? '' : ' disabled') + '>Зберегти</button>';
    main.innerHTML = h;
    main.querySelectorAll('.pad button').forEach(b => { b.onclick = () => {
      const k = b.dataset.k; let a = draft.amount;
      if (k === '⌫') a = a.slice(0, -1);
      else if (k === ',') { if (a.indexOf(',') === -1) a = (a || '0') + ','; }
      else if (a.replace(',', '').length < 9 && !/,\d\d$/.test(a)) a = (a === '0' ? '' : a) + k;
      draft.amount = a; keep(); renderAdd();
    }; });
    main.querySelectorAll('.seg button').forEach(b => { b.onclick = () => { draft.type = b.dataset.t; draft.category = ''; keep(); renderAdd(); }; });
    main.querySelectorAll('.chip').forEach(b => { b.onclick = () => { draft.category = b.dataset.c; keep(); renderAdd(); }; });
    if ($('#more')) $('#more').onclick = () => { draft.more = true; renderAdd(); };
    $('#saveOp').onclick = saveOp;
  }
  function keep() { if ($('#dDate')) { draft.date = $('#dDate').value; draft.account = $('#dAcc').value; draft.comment = $('#dCom').value; } }

  async function saveOp() {
    keep();
    const op = { id: Date.now().toString(36) + Math.random().toString(36).slice(2, 7), amount: Number(String(draft.amount).replace(',', '.')),
      type: draft.type, category: draft.category, account: draft.account, date: draft.date || st.today, comment: draft.comment };
    Object.assign(draft, { amount: '', category: '', comment: '', date: '', more: false });
    outbox.push(op); save(LS.outbox, outbox);
    view = 'home'; render(); toast('Записую…');
    try { await flush(); render(); toast('Записано ✓'); }
    catch (e) {
      if (e.api) { outbox = outbox.filter(o => o.id !== op.id); save(LS.outbox, outbox); render(); toast(e.message); }
      else toast('Немає звʼязку — запишу, щойно зʼявиться');
    }
  }

  // ── Налаштування
  async function renderSettings() {
    let h = '<h1>Налаштування</h1>';
    if (!settings) {
      main.innerHTML = h + '<div class="spin"></div>';
      try { settings = await api('getSettings'); } catch (e) {
        main.innerHTML = h + '<div class="empty">' + (e.api ? esc(e.message) : 'Немає звʼязку — налаштування доступні онлайн.') + '</div>' + footer(); bindFooter(); return;
      }
      if (view !== 'settings') return;
    }
    h += settings.blocks.map((b, bi) => '<details data-b="' + bi + '"><summary>' + esc(cap(b.title)) + ' · ' + b.rows.length + '</summary>' +
      '<p class="sub" style="margin-top:8px">' + esc(b.hint) + '</p><div class="rows">' + b.rows.map((r, ri) => rowCard(b, bi, r, ri)).join('') + '</div>' +
      '<button class="btn soft" data-add="' + bi + '">+ Додати</button><button class="btn" data-save="' + bi + '">Зберегти</button></details>').join('');
    main.innerHTML = h + footer();
    bindFooter();
    main.querySelectorAll('[data-add]').forEach(b => { b.onclick = () => {
      const bi = Number(b.dataset.add), blk = settings.blocks[bi]; collect(bi);
      if (blk.rows.length >= blk.max) return toast('У цьому блоці більше немає місця');
      const r = {}; blk.cols.forEach(c => { r[c.key] = c.t === 'check' ? (c.key === 'Use') : ''; }); blk.rows.push(r); reopen(bi);
    }; });
    main.querySelectorAll('[data-save]').forEach(b => { b.onclick = async () => {
      const bi = Number(b.dataset.save); collect(bi); b.disabled = true; b.textContent = 'Зберігаю…';
      try { settings = await api('saveSettings', { block: settings.blocks[bi].key, rows: settings.blocks[bi].rows }); toast('Збережено ✓'); reopen(bi); refresh(true); }
      catch (e) { b.disabled = false; b.textContent = 'Зберегти'; toast(e.api ? e.message : 'Немає звʼязку'); }
    }; });
    main.querySelectorAll('[data-del]').forEach(b => { b.onclick = () => {
      const p = b.dataset.del.split(':').map(Number); collect(p[0]); settings.blocks[p[0]].rows.splice(p[1], 1); reopen(p[0]);
    }; });
  }
  const cap = s => { s = String(s).toLowerCase(); return s.charAt(0).toUpperCase() + s.slice(1); };
  function reopen(bi) { renderSettings().then(() => { const d = main.querySelector('details[data-b="' + bi + '"]'); if (d) d.open = true; }); }
  function rowCard(b, bi, r, ri) {
    const f = c => {
      const id = 'data-f="' + bi + ':' + ri + ':' + c.key + '"', v = r[c.key];
      if (c.t === 'check') return '<label class="chk"><input type="checkbox" ' + id + (v === true ? ' checked' : '') + '>' + esc(c.h) + '</label>';
      if (c.t === 'list') return '<div><label class="f">' + esc(c.h) + '</label><select ' + id + '><option value=""></option>' +
        (settings.options[c.list] || []).map(o => '<option' + (o === v ? ' selected' : '') + '>' + esc(o) + '</option>').join('') + '</select></div>';
      const val = c.t === 'pct' && v !== '' ? Math.round(Number(v) * 100) : v;
      const type = c.t === 'date' ? 'date' : 'text';
      const mode = ['money', 'pct', 'int', 'day'].indexOf(c.t) !== -1 ? ' inputmode="decimal"' : '';
      return '<div><label class="f">' + esc(c.h) + (c.t === 'pct' ? ', %' : '') + '</label><input type="' + type + '"' + mode + ' ' + id + ' value="' + esc(val === null || val === undefined ? '' : val) + '"></div>';
    };
    const cols = b.cols.filter(c => c.key !== 'Name');
    return '<div class="cardrow"><div class="hd"><input data-f="' + bi + ':' + ri + ':Name" value="' + esc(r.Name) + '" placeholder="Назва">' +
      '<button class="del" data-del="' + bi + ':' + ri + '">✕</button></div><div class="grid2">' +
      cols.filter(c => c.t !== 'check').map(f).join('') + '</div>' + cols.filter(c => c.t === 'check').map(f).join('') + '</div>';
  }
  function collect(bi) {
    const blk = settings.blocks[bi];
    main.querySelectorAll('[data-f^="' + bi + ':"]').forEach(el => {
      const p = el.dataset.f.split(':'), c = blk.cols.filter(x => x.key === p[2])[0], row = blk.rows[Number(p[1])];
      if (!c || !row) return;
      let v = el.type === 'checkbox' ? el.checked : el.value.trim();
      if (['money', 'int', 'day'].indexOf(c.t) !== -1) v = v === '' ? '' : Number(String(v).replace(/\s/g, '').replace(',', '.')) || '';
      if (c.t === 'pct') v = v === '' ? '' : (Number(String(v).replace(',', '.')) || 0) / 100;
      row[c.key] = v;
    });
  }
  function footer() {
    return '<h2>Ще</h2><button class="btn soft" id="quiz">👋 Пройти квіз заново</button>' +
      (st && st.sheetUrl ? '<button class="btn soft" id="sheet">Відкрити таблицю</button>' : '') +
      '<button class="btn soft" id="reload">Оновити дані</button>' +
      (isIos && !standalone ? '<button class="btn soft" id="cpl">📋 Скопіювати посилання для іконки</button>' +
        '<p class="sub">Якщо іконка на початковому екрані просить підключення — скопіюй тут і встав у неї. Нікому не пересилай: у посиланні ключ.</p>' : '') +
      '<button class="btn ghost" id="off">Відʼєднати застосунок</button>';
  }
  function bindFooter() {
    $('#quiz').onclick = openQuiz;
    if ($('#sheet')) $('#sheet').onclick = () => window.open(st.sheetUrl, '_blank');
    if ($('#cpl')) $('#cpl').onclick = () => {
      const l = location.origin + location.pathname + '?api=' + encodeURIComponent(conn.api) + '&key=' + conn.key;
      (navigator.clipboard ? navigator.clipboard.writeText(l) : Promise.reject()).then(() => toast('Скопійовано ✓'), () => prompt('Скопіюй посилання:', l));
    };
    $('#reload').onclick = () => { settings = null; refresh().then(() => { toast('Оновлено'); render(); }); };
    $('#off').onclick = () => { if (!confirm('Відʼєднати застосунок від таблиці? Дані в таблиці залишаться.')) return;
      Object.keys(LS).forEach(k => localStorage.removeItem(LS[k])); conn = null; st = null; outbox = []; settings = null;
      history.replaceState(null, '', location.pathname); render(); };
  }

  // ── Квіз: той самий, що в таблиці; відповіді йдуть через API
  async function openQuiz() {
    const ov = $('#overlay');
    ov.hidden = false; ov.innerHTML = '<button class="x" aria-label="Закрити">✕</button><div class="spin"></div>';
    ov.querySelector('.x').onclick = closeQuiz;
    try {
      const q = await api('quizHtml');
      const shim = '<script>(function(){function mk(){var o={s:null,f:null,withSuccessHandler:function(f){o.s=f;return o;},withFailureHandler:function(f){o.f=f;return o;},' +
        'quizSubmit:function(p,m,a){parent.btQuiz.submit(p,a).then(function(){o.s&&o.s("saved");},function(e){o.f&&o.f(e);});},' +
        'showGuide:function(){parent.btQuiz.close();}};return o;}' +
        'window.google={script:{host:{close:function(){parent.btQuiz.close();}}}};Object.defineProperty(window.google.script,"run",{get:mk});})();<\/script>';
      const fr = document.createElement('iframe');
      fr.srcdoc = q.html.replace('<body>', () => '<body>' + shim);
      ov.querySelector('.spin').replaceWith(fr);
    } catch (e) { closeQuiz(); toast(e.api ? e.message : 'Немає звʼязку — квіз доступний онлайн'); }
  }
  function closeQuiz() { const ov = $('#overlay'); ov.hidden = true; ov.innerHTML = ''; settings = null; render(); }
  window.btQuiz = {
    submit: async (profile, apply) => { setState(await api('quiz', { profile: profile, apply: apply })); },
    close: closeQuiz,
  };

  // ── Запуск
  document.querySelectorAll('#nav button').forEach(b => { b.onclick = () => { view = b.dataset.view; render(); }; });
  window.addEventListener('online', () => refresh(true));
  document.addEventListener('visibilitychange', () => { if (!document.hidden) refresh(true); });
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  theme(); render(); refresh(!!st);
})();
