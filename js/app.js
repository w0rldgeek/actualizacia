/**
 * ============================================================
 * Актуализация номеров абонентов
 * ============================================================
 * Рабочий инструмент учёта номеров. Записи группируются по
 * месяцам и разделяются по дням. Спокойная палитра, компактная
 * подача — и на странице, и в выгрузках.
 *
 * Данные можно:
 *   • загрузить из .xlsx (импорт),
 *   • выгрузить в .xlsx (с разбивкой по месяцам и дням),
 *   • выгрузить в .html (готовая статичная таблица),
 *   • писать прямо в привязанный файл (Chrome/Edge, File System
 *     Access API) — каждая новая строка уходит в .xlsx сразу.
 *
 * localStorage хранит данные между перезагрузками.
 * ============================================================
 */

(function () {
    'use strict';

    const STORAGE_KEY = 'subscribers_v1';

    const MONTH_NAMES = [
        'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
        'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь',
    ];
    const MONTH_NAMES_LC = MONTH_NAMES.map(m => m.toLowerCase());

    const WEEKDAYS = [
        'Воскресенье', 'Понедельник', 'Вторник', 'Среда',
        'Четверг', 'Пятница', 'Суббота',
    ];

    // Приглушённые цвета месяцев для Excel-вкладок/заголовков
    const MONTH_COLORS = [
        '5B7FB4', '4F9A98', '5A9E6F', '7C9A52', 'A8923F', 'C0824E',
        'BF6A5C', 'B06A89', '8F6AAD', '6F72B3', '5F86B0', '4F96A8',
    ];
    // Чуть светлее — для тёмной темы (страница и HTML-выгрузка), как в style.css
    const MONTH_COLORS_DARK = [
        '6F8FC4', '59AAA8', '67AD7D', '8AA85F', 'BBA34A', 'CF9159',
        'CF7668', 'C07A98', '9D79BC', '7D80C2', '6E96C0', '5CA6B8',
    ];

    const supportsFS = typeof window.showSaveFilePicker === 'function';
    let fileHandle = null;   // привязанный .xlsx
    let fileName = null;
    let editingId = null;    // id записи в режиме редактирования

    // ============================================================
    // STORE — данные (localStorage)
    // ============================================================
    const Store = {
        all() {
            try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); }
            catch { return []; }
        },
        save(list) { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); },
        add(item) {
            const list = this.all();
            item.id = uid();
            list.push(item);
            this.save(list);
            return item;
        },
        remove(id) { this.save(this.all().filter(s => s.id !== id)); },
        update(id, fields) {
            const list = this.all();
            const i = list.findIndex(s => s.id === id);
            if (i === -1) return null;
            list[i] = Object.assign({}, list[i], fields);
            this.save(list);
            return list[i];
        },
        get(id) { return this.all().find(s => s.id === id) || null; },
        replaceAll(list) {
            this.save(list.map(s => Object.assign({ id: uid() }, s)));
        },
        addMany(list) {
            const cur = this.all();
            for (const s of list) cur.push(Object.assign({ id: uid() }, s));
            this.save(cur);
        },
    };

    // ============================================================
    // IndexedDB — хранение ссылки на файл между перезагрузками
    // ============================================================
    const HandleDB = {
        _db() {
            return new Promise((resolve, reject) => {
                const req = indexedDB.open('abonenty', 1);
                req.onupgradeneeded = () => req.result.createObjectStore('handles');
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            });
        },
        async set(handle) {
            try {
                const db = await this._db();
                const tx = db.transaction('handles', 'readwrite');
                tx.objectStore('handles').put(handle, 'excel');
            } catch (e) { /* not critical */ }
        },
        async get() {
            try {
                const db = await this._db();
                return await new Promise((resolve) => {
                    const r = db.transaction('handles', 'readonly')
                        .objectStore('handles').get('excel');
                    r.onsuccess = () => resolve(r.result || null);
                    r.onerror = () => resolve(null);
                });
            } catch (e) { return null; }
        },
    };

    // ============================================================
    // ГРУППИРОВКА — по месяцам, внутри по дням
    // ============================================================
    function groupByMonth(rows) {
        const months = new Map();
        for (const s of rows) {
            const d = s.date ? new Date(s.date) : new Date();
            const key = `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}`;
            if (!months.has(key)) months.set(key, { y: d.getFullYear(), m: d.getMonth(), items: [] });
            months.get(key).items.push(s);
        }
        return months;
    }

    // Нормализация номера для сравнения (только цифры)
    function normPhone(p) { return (p || '').replace(/\D/g, ''); }

    // Множество нормализованных номеров, встречающихся больше одного раза
    function buildDupPhones(rows) {
        const counts = new Map();
        for (const s of rows) {
            const n = normPhone(s.phone);
            if (!n) continue;
            counts.set(n, (counts.get(n) || 0) + 1);
        }
        const dups = new Set();
        for (const [n, c] of counts) if (c > 1) dups.add(n);
        return dups;
    }

    function groupByDay(items) {
        const days = new Map();
        for (const s of items) {
            const d = s.date ? new Date(s.date) : new Date();
            const key = `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
            if (!days.has(key)) days.set(key, { date: d, items: [] });
            days.get(key).items.push(s);
        }
        return days;
    }

    // ============================================================
    // EXCEL — сборка книги (ExcelJS): месяцы + разделители дней
    // ============================================================
    async function buildWorkbook() {
        const wb = new ExcelJS.Workbook();
        wb.creator = 'Актуализация номеров';
        const rows = Store.all();

        const months = groupByMonth(rows);
        if (months.size === 0) {
            const now = new Date();
            months.set('empty', { y: now.getFullYear(), m: now.getMonth(), items: [] });
        }

        const keys = Array.from(months.keys()).sort(); // хронологически
        for (const key of keys) {
            const g = months.get(key);
            const argb = 'FF' + MONTH_COLORS[g.m];
            const argbSoft = 'FF' + tint(MONTH_COLORS[g.m], 0.85);
            const ws = wb.addWorksheet(`${MONTH_NAMES[g.m]} ${g.y}`, {
                properties: { tabColor: { argb } },
                views: [{ state: 'frozen', ySplit: 2 }],
            });

            ws.columns = [
                { width: 13 }, { width: 40 }, { width: 20 }, { width: 20 },
            ];

            // Заголовок месяца
            ws.mergeCells('A1:D1');
            const title = ws.getCell('A1');
            title.value = `${MONTH_NAMES[g.m]} ${g.y}`;
            title.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
            title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
            title.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            ws.getRow(1).height = 22;

            // Шапка
            const header = ws.getRow(2);
            ['Дата', 'Адрес', 'Логин', 'Номер телефона'].forEach((h, i) => {
                const c = header.getCell(i + 1);
                c.value = h;
                c.font = { bold: true, color: { argb: 'FF333333' } };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F3F5' } };
                c.alignment = { horizontal: 'left', vertical: 'middle' };
            });
            header.height = 18;

            // Дни (новые сверху)
            const days = groupByDay(g.items);
            const dayKeys = Array.from(days.keys()).sort().reverse();
            for (const dk of dayKeys) {
                const dg = days.get(dk);

                // Разделитель дня — объединённая строка с лёгкой заливкой месяца
                const sepRow = ws.addRow([dayLabel(dg.date)]);
                ws.mergeCells(`A${sepRow.number}:D${sepRow.number}`);
                const sc = sepRow.getCell(1);
                sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbSoft } };
                sc.font = { bold: true, color: { argb: 'FF333333' }, size: 11 };
                sc.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
                sepRow.height = 17;

                // Записи дня (по убыванию даты — в пределах дня порядок добавления)
                for (const s of dg.items) {
                    const r = ws.addRow([
                        s.date ? new Date(s.date) : null,
                        s.address, s.login, s.phone,
                    ]);
                    r.getCell(1).numFmt = 'dd.mm.yyyy';
                    r.height = 16;
                }
            }

            // Тонкие рамки на всю заполненную область
            const lastRow = ws.rowCount;
            for (let rr = 2; rr <= lastRow; rr++) {
                for (let cc = 1; cc <= 4; cc++) {
                    ws.getCell(rr, cc).border = {
                        top: { style: 'hair', color: { argb: 'FFDDE1E6' } },
                        left: { style: 'hair', color: { argb: 'FFDDE1E6' } },
                        bottom: { style: 'hair', color: { argb: 'FFDDE1E6' } },
                        right: { style: 'hair', color: { argb: 'FFDDE1E6' } },
                    };
                }
            }
        }

        const buffer = await wb.xlsx.writeBuffer();
        return new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
    }

    // Записать книгу: в привязанный файл (тихо) либо скачиванием
    async function saveExcel(opts) {
        opts = opts || {};
        const blob = await buildWorkbook();

        if (fileHandle) {
            try {
                const w = await fileHandle.createWritable();
                await w.write(blob);
                await w.close();
                if (opts.notify) toast('Сохранено в Excel', 'success');
                return true;
            } catch (e) {
                console.error('Excel write error:', e);
                toast('Не удалось записать в файл — скачиваю копию', 'warning');
            }
        }

        if (opts.download || !supportsFS) {
            downloadBlob(blob, fileName || `Абоненты_номера_${todayISO()}.xlsx`);
            return true;
        }

        if (opts.notify) toast('Файл Excel не привязан — нажмите «Привязать Excel»', 'warning');
        return false;
    }

    // ============================================================
    // EXCEL — импорт (ExcelJS): читаем месяцы/дни обратно
    // ============================================================
    async function importExcel(file) {
        try {
            const buf = await file.arrayBuffer();
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buf);

            const parsed = [];
            wb.eachSheet((ws) => {
                // Год из названия листа ("Январь 2026"), иначе текущий
                const ym = parseSheetTitle(ws.name);
                ws.eachRow((row) => {
                    // Заголовок месяца и разделители дней — объединённые
                    // ячейки (A:D). У строк данных объединения нет.
                    if (row.getCell(2).isMerged) return;

                    const a = cellVal(row.getCell(1));
                    const b = cellVal(row.getCell(2));
                    const c = cellVal(row.getCell(3));
                    const d = cellVal(row.getCell(4));

                    // Пропускаем шапку и пустые строки.
                    if (isHeaderLabel(a)) return;
                    if (!b && !c && !d) return;

                    let dateISO = null;
                    if (a instanceof Date) dateISO = toISO(a);
                    else if (typeof a === 'string') dateISO = parseRuDate(a, ym);
                    if (!dateISO && ym) dateISO = `${ym.y}-${String(ym.m + 1).padStart(2, '0')}-01`;

                    parsed.push({
                        date: dateISO || todayISO(),
                        address: str(b),
                        login: str(c),
                        phone: str(d),
                    });
                });
            });

            if (parsed.length === 0) {
                toast('В файле не найдено записей', 'warning');
                return;
            }

            const hasExisting = Store.all().length > 0;
            let mode = 'replace';
            if (hasExisting) {
                mode = confirm(
                    `Загружено записей: ${parsed.length}.\n\n` +
                    'OK — заменить текущие данные,\n' +
                    'Отмена — добавить к текущим.'
                ) ? 'replace' : 'merge';
            }

            if (mode === 'replace') Store.replaceAll(parsed);
            else Store.addMany(parsed);

            render();
            toast(`Импортировано записей: ${parsed.length}`, 'success');
            saveExcel(); // синхронизируем привязанный файл, если есть
        } catch (e) {
            console.error('Import error:', e);
            toast('Не удалось прочитать файл Excel', 'error');
        }
    }

    // ============================================================
    // HTML — выгрузка статичной таблицы
    // ============================================================
    function exportHtml() {
        const rows = Store.all();
        const generated = new Date().toLocaleString('ru-RU');
        const total = rows.length;
        const monthsCount = groupByMonth(rows).size;
        const dupPhones = buildDupPhones(rows);
        const dupCount = rows.filter(s => dupPhones.has(normPhone(s.phone))).length;

        const doc = `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Номера абонентов — ${escapeHtml(generated)}</title>
<style>${EXPORT_CSS}</style>
</head>
<body>
<div class="wrap">
<header class="head">
<h1>Актуализация номеров абонентов</h1>
<div class="meta">Сформировано: ${escapeHtml(generated)} · Записей: ${total} · Месяцев: ${monthsCount}${dupCount ? ` · Повторов номеров: ${dupCount}` : ''}</div>
</header>
${rows.length ? monthsHtml(rows, { actions: false, dupPhones }) : '<div class="empty">Нет записей.</div>'}
</div>
</body>
</html>`;

        const blob = new Blob([doc], { type: 'text/html;charset=utf-8' });
        downloadBlob(blob, `Абоненты_номера_${todayISO()}.html`);
        toast('HTML выгружен', 'success');
    }

    // CSS для автономной HTML-выгрузки — тёмная тема, как на странице
    const EXPORT_CSS = `
*{box-sizing:border-box;margin:0;padding:0}
@keyframes rise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
body{font-family:'Inter',-apple-system,'Segoe UI',sans-serif;font-size:13.5px;color:#ecedf5;background:#0b0b14;line-height:1.5;padding:28px;
background-image:radial-gradient(60vw 60vw at 12% -10%,rgba(124,92,255,.18),transparent 60%),radial-gradient(50vw 50vw at 100% 0%,rgba(34,211,238,.12),transparent 55%);background-attachment:fixed}
.wrap{max-width:1100px;margin:0 auto;display:flex;flex-direction:column;gap:16px}
.head{padding:2px 2px}
.head h1{font-size:19px;font-weight:700;letter-spacing:-.3px;background:linear-gradient(120deg,#fff,#c8c9e6);-webkit-background-clip:text;background-clip:text;color:transparent}
.head .meta{font-size:12.5px;color:#6c6e88;margin-top:5px}
.empty{padding:44px;text-align:center;color:#6c6e88;background:rgba(255,255,255,.04);border:1px dashed rgba(255,255,255,.18);border-radius:16px}
.month-card{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.10);border-radius:16px;overflow:hidden;box-shadow:0 20px 50px -20px rgba(0,0,0,.7);animation:rise .36s cubic-bezier(.2,.8,.2,1) both}
.month-card-header{display:flex;align-items:center;justify-content:space-between;padding:13px 20px;border-bottom:1px solid rgba(255,255,255,.10);background:linear-gradient(90deg,color-mix(in srgb,var(--mc,#7c5cff) 24%,transparent),transparent 70%)}
.month-card-header h3{font-size:15px;font-weight:700;letter-spacing:-.2px;color:var(--mc,#7c5cff)}
.month-total{padding:3px 11px;border-radius:999px;font-size:11.5px;font-weight:700;background:var(--mc,#7c5cff);color:#0b0b14}
table{width:100%;border-collapse:collapse;font-size:13.5px}
thead th{text-align:left;padding:10px 18px;color:#6c6e88;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid rgba(255,255,255,.10)}
td{padding:9px 18px;border-bottom:1px solid rgba(255,255,255,.10)}
.phone{font-variant-numeric:tabular-nums;font-weight:600;white-space:nowrap}
.day-row td{padding:7px 18px;background:color-mix(in srgb,var(--mc,#7c5cff) 14%,transparent);border-top:1px solid color-mix(in srgb,var(--mc,#7c5cff) 30%,transparent)}
.day-name{font-weight:700;font-size:12px;color:color-mix(in srgb,var(--mc,#7c5cff) 55%,#ecedf5)}
.day-count{margin-left:8px;font-size:11px;font-weight:600;color:#6c6e88}
.phone-val{font-variant-numeric:tabular-nums}
.dup-tag{display:inline-block;margin-left:8px;padding:1px 7px;border-radius:999px;font-size:10px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;color:#fbbf24;background:rgba(251,191,36,.14);border:1px solid rgba(251,191,36,.35);vertical-align:middle}
${MONTH_COLORS_DARK.map((c, i) => `.month-${i}{--mc:#${c}}`).join('')}
`;

    // ============================================================
    // RENDER — общий генератор разметки месяцев/дней
    //   actions:true  — версия для страницы (кнопки удаления)
    //   actions:false — статичная версия (HTML-выгрузка)
    // ============================================================
    function monthsHtml(rows, opts) {
        opts = opts || {};
        const withActions = !!opts.actions;
        const dupPhones = opts.dupPhones || new Set();
        const months = groupByMonth(rows);
        const monthKeys = Array.from(months.keys()).sort().reverse();
        const colSpan = withActions ? 4 : 3;

        return monthKeys.map(mk => {
            const g = months.get(mk);
            const days = groupByDay(g.items);
            const dayKeys = Array.from(days.keys()).sort().reverse();

            const body = dayKeys.map(dk => {
                const dg = days.get(dk);
                const sep = `<tr class="day-row"><td colspan="${colSpan}">` +
                    `<span class="day-name">${escapeHtml(dayLabel(dg.date))}</span>` +
                    `<span class="day-count">${dg.items.length}</span></td></tr>`;
                const rws = dg.items.map(s => {
                    const isDup = dupPhones.has(normPhone(s.phone));
                    const dupTag = isDup
                        ? '<span class="dup-tag" title="Этот номер встречается в базе несколько раз">дубль</span>'
                        : '';
                    const phoneCell = withActions
                        ? `<td class="phone copyable" data-copy="${escapeHtml(s.phone)}" title="Нажмите, чтобы скопировать"><span class="phone-val">${escapeHtml(s.phone)}</span>${dupTag}</td>`
                        : `<td class="phone"><span class="phone-val">${escapeHtml(s.phone)}</span>${dupTag}</td>`;
                    return `
                    <tr${withActions && s.id === editingId ? ' class="is-editing"' : ''}>
                        <td>${escapeHtml(s.address)}</td>
                        <td>${escapeHtml(s.login)}</td>
                        ${phoneCell}
                        ${withActions ? `<td class="sub-actions">` +
                            `<button class="row-btn row-edit" data-id="${s.id}" title="Редактировать">✎</button>` +
                            `<button class="row-btn row-delete" data-id="${s.id}" title="Удалить">✕</button>` +
                        `</td>` : ''}
                    </tr>`;
                }).join('');
                return sep + rws;
            }).join('');

            return `
                <section class="month-card month-${g.m}">
                    <header class="month-card-header">
                        <h3>${MONTH_NAMES[g.m]} ${g.y}</h3>
                        <span class="month-total">${g.items.length}</span>
                    </header>
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Адрес</th>
                                <th>Логин</th>
                                <th class="phone">Номер телефона</th>
                                ${withActions ? '<th></th>' : ''}
                            </tr>
                        </thead>
                        <tbody>${body}</tbody>
                    </table>
                </section>`;
        }).join('');
    }

    function render() {
        const container = document.getElementById('table-container');
        const summary = document.getElementById('summary');
        const filter = (document.getElementById('f-filter').value || '').trim().toLowerCase();

        const allRows = Store.all();
        let rows = allRows;
        if (filter) {
            rows = rows.filter(s =>
                (s.address || '').toLowerCase().includes(filter) ||
                (s.login || '').toLowerCase().includes(filter) ||
                (s.phone || '').toLowerCase().includes(filter)
            );
        }

        // Повторяющиеся номера — считаем по всей базе
        const dupPhones = buildDupPhones(allRows);
        const dupCount = allRows.filter(s => dupPhones.has(normPhone(s.phone))).length;

        // Сводка
        if (allRows.length) {
            const months = groupByMonth(allRows).size;
            summary.hidden = false;
            summary.innerHTML =
                `<span><b>${allRows.length}</b> записей</span>` +
                `<span>в <b>${months}</b> мес.</span>` +
                (dupCount ? `<span class="dup-note">повторов номеров: <b>${dupCount}</b></span>` : '') +
                (filter ? `<span>показано: <b>${rows.length}</b></span>` : '');
        } else {
            summary.hidden = true;
        }

        if (rows.length === 0) {
            container.innerHTML = `<div class="empty-state">${
                allRows.length === 0
                    ? 'Пока пусто. Заполните форму — строка добавится сюда и уйдёт в Excel.'
                    : 'Ничего не найдено по фильтру.'
            }</div>`;
            return;
        }

        container.innerHTML = monthsHtml(rows, { actions: true, dupPhones });

        container.querySelectorAll('.row-edit').forEach(btn => {
            btn.addEventListener('click', () => startEdit(btn.dataset.id));
        });
        container.querySelectorAll('.phone.copyable').forEach(cell => {
            cell.addEventListener('click', () => copyToClipboard(cell.dataset.copy, cell));
        });
        container.querySelectorAll('.row-delete').forEach(btn => {
            btn.addEventListener('click', () => {
                if (confirm('Удалить запись абонента?')) {
                    if (btn.dataset.id === editingId) cancelEdit();
                    Store.remove(btn.dataset.id);
                    toast('Удалено', 'success');
                    render();
                    saveExcel();
                }
            });
        });
    }

    // ============================================================
    // Добавление / редактирование записи → сразу в Excel
    // ============================================================
    function submitForm(opts) {
        opts = opts || {};
        const address = document.getElementById('f-address').value.trim();
        const login = document.getElementById('f-login').value.trim();
        const phone = document.getElementById('f-phone').value.trim();
        let date = document.getElementById('f-date').value;

        if (!address || !login || !phone) {
            if (opts.notify) toast('Заполните адрес, логин и номер телефона', 'warning');
            return;
        }
        if (!date) date = todayISO();

        // Ненавязчивое предупреждение: такой номер уже есть в базе
        // (саму запись всё равно сохраняем — решает пользователь).
        const dupExists = Store.all().some(s =>
            s.id !== editingId && normPhone(s.phone) && normPhone(s.phone) === normPhone(phone)
        );
        if (dupExists) toast('Внимание: такой номер уже есть в базе', 'warning');

        if (editingId) {
            Store.update(editingId, { address, login, phone, date });
            exitEditMode();
            clearForm();
            render();
            saveExcel({ notify: true });
            toast('Запись обновлена', 'success');
            return;
        }

        Store.add({ address, login, phone, date });
        clearForm();
        if (opts.focus) document.getElementById('f-address').focus();
        render();
        saveExcel({ notify: true });
    }

    // Копирование номера в буфер обмена + короткая подсветка ячейки
    function copyToClipboard(text, cell) {
        const done = () => {
            if (cell) {
                cell.classList.add('copied');
                setTimeout(() => cell.classList.remove('copied'), 900);
            }
            toast('Номер скопирован', 'success', 1500);
        };
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopy(text, done));
        } else {
            fallbackCopy(text, done);
        }
    }
    function fallbackCopy(text, done) {
        try {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            done();
        } catch (e) {
            toast('Не удалось скопировать', 'error');
        }
    }

    function clearForm() {
        document.getElementById('f-address').value = '';
        document.getElementById('f-login').value = '';
        document.getElementById('f-phone').value = '';
        document.getElementById('f-date').value = todayISO();
    }

    // Начать редактирование: переносим запись в форму
    function startEdit(id) {
        const s = Store.get(id);
        if (!s) return;
        editingId = id;
        document.getElementById('f-address').value = s.address || '';
        document.getElementById('f-login').value = s.login || '';
        document.getElementById('f-phone').value = s.phone || '';
        document.getElementById('f-date').value = (s.date || todayISO()).slice(0, 10);

        document.querySelector('.card').classList.add('editing');
        document.getElementById('btn-add').textContent = 'Сохранить';
        document.getElementById('f-address').focus();
        window.scrollTo({ top: 0, behavior: 'smooth' });
        render(); // подсветить редактируемую строку
    }

    function exitEditMode() {
        editingId = null;
        document.querySelector('.card').classList.remove('editing');
        document.getElementById('btn-add').textContent = 'Добавить';
    }

    function cancelEdit() {
        exitEditMode();
        clearForm();
        render();
    }

    // ============================================================
    // Привязка Excel-файла (File System Access API)
    // ============================================================
    async function linkExcel() {
        if (!supportsFS) {
            toast('Этот браузер не пишет в файл напрямую. Используйте «Выгрузить Excel».', 'warning');
            return;
        }
        try {
            const stored = await HandleDB.get();
            if (stored) {
                const perm = await stored.requestPermission({ mode: 'readwrite' });
                if (perm === 'granted') {
                    fileHandle = stored;
                    fileName = stored.name;
                    setExcelStatus();
                    await saveExcel({ notify: true });
                    return;
                }
            }
            fileHandle = await window.showSaveFilePicker({
                suggestedName: `Абоненты_номера_${new Date().getFullYear()}.xlsx`,
                types: [{
                    description: 'Excel',
                    accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
                }],
            });
            fileName = fileHandle.name;
            await HandleDB.set(fileHandle);
            setExcelStatus();
            await saveExcel({ notify: true });
        } catch (e) {
            if (e && e.name === 'AbortError') return;
            console.error('Link error:', e);
            toast('Не удалось привязать файл', 'error');
        }
    }

    async function tryRestoreHandle() {
        if (!supportsFS) return;
        const stored = await HandleDB.get();
        if (!stored) return;
        try {
            const perm = await stored.queryPermission({ mode: 'readwrite' });
            fileName = stored.name;
            if (perm === 'granted') fileHandle = stored;
        } catch (e) { /* ignore */ }
        setExcelStatus();
    }

    function setExcelStatus() {
        const el = document.getElementById('excel-status');
        if (fileHandle) {
            el.textContent = `Excel: ${fileName}`;
            el.classList.add('linked');
        } else if (fileName) {
            el.textContent = `Excel: ${fileName} (привяжите заново)`;
            el.classList.remove('linked');
        } else {
            el.textContent = supportsFS ? 'Excel не привязан' : 'Режим выгрузки';
            el.classList.remove('linked');
        }
    }

    // ============================================================
    // UTILS
    // ============================================================
    function uid() { return Date.now() + '-' + Math.random().toString(36).slice(2, 8); }
    function todayISO() { return new Date().toISOString().slice(0, 10); }
    function toISO(d) {
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    function fmtDate(d) {
        return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
    }
    function dayLabel(d) {
        return `${WEEKDAYS[d.getDay()]}, ${fmtDate(d)}`;
    }
    function str(v) {
        if (v == null) return '';
        if (v instanceof Date) return fmtDate(v);
        return String(v).trim();
    }

    // Значение ячейки ExcelJS → примитив (учёт rich text / формул / гиперссылок)
    function cellVal(cell) {
        let v = cell ? cell.value : null;
        if (v == null) return null;
        if (v instanceof Date) return v;
        if (typeof v === 'object') {
            if (v.richText) return v.richText.map(t => t.text).join('');
            if (v.text != null) return v.text;
            if (v.result != null) return v.result;
            if (v.hyperlink != null) return v.text || v.hyperlink;
        }
        return v;
    }

    function isHeaderLabel(a) {
        if (typeof a !== 'string') return false;
        const t = a.trim().toLowerCase();
        if (t === 'дата') return true;
        // Заголовок месяца: "Январь 2026" и т.п.
        return MONTH_NAMES_LC.some(m => t.startsWith(m));
    }

    // "Январь 2026" → {y, m}
    function parseSheetTitle(name) {
        if (!name) return null;
        const t = name.trim().toLowerCase();
        for (let i = 0; i < MONTH_NAMES_LC.length; i++) {
            if (t.startsWith(MONTH_NAMES_LC[i])) {
                const ym = t.match(/(\d{4})/);
                return { m: i, y: ym ? parseInt(ym[1], 10) : new Date().getFullYear() };
            }
        }
        return null;
    }

    // Разбор даты из строки: "13.06.2026", "2026-06-13" и т.п.
    function parseRuDate(s, ymFallback) {
        if (!s) return null;
        s = String(s).trim();
        let m = s.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/); // dd.mm.yyyy
        if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
        m = s.match(/(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);      // yyyy-mm-dd
        if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
        return null;
    }

    function downloadBlob(blob, name) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = name;
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function escapeHtml(text) {
        if (text == null) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Осветление HEX-цвета к белому (amount: 0..1) → HEX без #
    function tint(hex, amount) {
        const r = parseInt(hex.slice(0, 2), 16);
        const g = parseInt(hex.slice(2, 4), 16);
        const b = parseInt(hex.slice(4, 6), 16);
        const mix = (c) => Math.round(c + (255 - c) * amount);
        return [mix(r), mix(g), mix(b)]
            .map(c => c.toString(16).padStart(2, '0')).join('').toUpperCase();
    }

    function toast(message, type = 'info', duration = 3500) {
        const container = document.getElementById('toast-container');
        const el = document.createElement('div');
        el.className = `toast ${type}`;
        el.textContent = message;
        container.appendChild(el);
        setTimeout(() => {
            el.style.opacity = '0';
            el.style.transform = 'translateX(40px)';
            el.style.transition = 'all 0.3s ease';
            setTimeout(() => el.remove(), 300);
        }, duration);
    }

    // ============================================================
    // INIT
    // ============================================================
    function init() {
        document.getElementById('f-date').value = todayISO();

        document.getElementById('btn-add').addEventListener('click', () => submitForm({ focus: true, notify: true }));
        document.getElementById('btn-cancel').addEventListener('click', cancelEdit);
        document.getElementById('btn-link').addEventListener('click', linkExcel);
        document.getElementById('btn-export-xlsx').addEventListener('click', () => saveExcel({ download: true }));
        document.getElementById('btn-export-html').addEventListener('click', exportHtml);
        document.getElementById('f-filter').addEventListener('input', render);

        const fileInput = document.getElementById('file-import');
        document.getElementById('btn-import').addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files && e.target.files[0];
            if (file) importExcel(file);
            e.target.value = ''; // позволяем повторно выбрать тот же файл
        });

        // Быстрый ввод: Enter из любого поля добавляет строку.
        ['f-address', 'f-login', 'f-phone', 'f-date'].forEach(id => {
            document.getElementById(id).addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    submitForm({ focus: true, notify: true });
                } else if (e.key === 'Escape' && editingId) {
                    e.preventDefault();
                    cancelEdit();
                }
            });
        });

        setExcelStatus();
        tryRestoreHandle();
        render();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
