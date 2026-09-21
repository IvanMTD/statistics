/*
 * Универсальный движок формул Excel: токенизатор -> парсер (AST) -> вычислитель.
 * Чистый модуль: не знает про DOM и ruleset.json. На входе текст формулы и книга.
 */
(function (global) {
    'use strict';

    /* ===================== значения ===================== */

    const ERR = {
        NA: '#Н/Д', VALUE: '#ЗНАЧ!', DIV0: '#ДЕЛ/0!', REF: '#ССЫЛКА!',
        NAME: '#ИМЯ?', NUM: '#ЧИСЛО!', NULL: '#ПУСТО!'
    };

    const ERROR_LITERALS = [
        ERR.NA, ERR.VALUE, ERR.DIV0, ERR.REF, ERR.NAME, ERR.NUM, ERR.NULL,
        '#NULL!', '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#N/A'
    ];

    function error(code) { return { __error: true, code: code }; }
    function isError(v) { return !!v && typeof v === 'object' && v.__error === true; }
    function isRange(v) { return !!v && typeof v === 'object' && v.type === 'range'; }
    function scalar(v) { return isRange(v) ? error(ERR.VALUE) : v; }
    function isEmpty(v) { return v === undefined || v === null || v === ''; }

    function toNumber(v) {
        v = scalar(v);
        if (isError(v)) return v;
        if (typeof v === 'number') return v;
        if (typeof v === 'boolean') return v ? 1 : 0;
        if (isEmpty(v)) return 0;
        const n = Number(String(v).trim());
        return isNaN(n) ? error(ERR.VALUE) : n;
    }

    function toText(v) {
        v = scalar(v);
        if (isError(v)) return v;
        if (v === undefined || v === null) return '';
        if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
        return String(v);
    }

    function toBool(v) {
        v = scalar(v);
        if (isError(v)) return v;
        if (typeof v === 'boolean') return v;
        if (typeof v === 'number') return v !== 0;
        if (isEmpty(v)) return false;
        const s = String(v).trim().toUpperCase();
        if (s === 'ИСТИНА' || s === 'TRUE') return true;
        if (s === 'ЛОЖЬ' || s === 'FALSE') return false;
        return error(ERR.VALUE);
    }

    function kindOf(v) {
        v = scalar(v);
        if (isError(v)) return { k: 'err' };
        if (v === undefined || v === null || v === '') return { k: 'empty' };
        if (typeof v === 'number') return { k: 'num', v: v };
        if (typeof v === 'boolean') return { k: 'num', v: v ? 1 : 0 };
        return { k: 'text', v: String(v).toLowerCase() };
    }

    // Сравнение в духе Excel: пусто = 0 и = "", число всегда меньше текста.
    function compareValues(a, b) {
        const A = kindOf(a), B = kindOf(b);
        if (A.k === 'err' || B.k === 'err') return NaN;
        if (A.k === 'empty' && B.k === 'empty') return 0;
        if (A.k === 'empty') return B.k === 'num' ? (B.v === 0 ? 0 : (B.v > 0 ? -1 : 1)) : -1;
        if (B.k === 'empty') return A.k === 'num' ? (A.v === 0 ? 0 : (A.v > 0 ? 1 : -1)) : 1;
        if (A.k === 'num' && B.k === 'num') return A.v === B.v ? 0 : (A.v < B.v ? -1 : 1);
        if (A.k === 'text' && B.k === 'text') return A.v === B.v ? 0 : (A.v < B.v ? -1 : 1);
        return A.k === 'num' ? -1 : 1;
    }

    /* ===================== адреса ячеек ===================== */

    function colToIndex(col) {
        let n = 0;
        col = col.toUpperCase();
        for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
        return n - 1;
    }

    function indexToCol(index) {
        let s = '', i = index + 1;
        while (i > 0) {
            const r = (i - 1) % 26;
            s = String.fromCharCode(65 + r) + s;
            i = Math.floor((i - 1) / 26);
        }
        return s;
    }

    function cellAddress(col, row) { return col.toUpperCase() + row; }

    /* ===================== токенизатор ===================== */

    const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
    const BOOL_NAMES = { 'ИСТИНА': true, 'TRUE': true, 'ЛОЖЬ': false, 'FALSE': false };
    const WORD_START = /[A-Za-zА-Яа-яЁё_$]/;
    const WORD_CHAR = /[A-Za-zА-Яа-яЁё0-9_.$]/;

    function tokenize(input) {
        const src = String(input).replace(/^\s*=/, '');
        const tokens = [];
        let i = 0;
        const n = src.length;

        while (i < n) {
            const ch = src[i];

            if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { i++; continue; }

            if (ch === '"') {
                let j = i + 1, buf = '';
                while (j < n) {
                    if (src[j] === '"') {
                        if (src[j + 1] === '"') { buf += '"'; j += 2; continue; }
                        break;
                    }
                    buf += src[j++];
                }
                tokens.push({ type: 'string', value: buf });
                i = j + 1;
                continue;
            }

            const errLit = ERROR_LITERALS.find(function (e) { return src.startsWith(e, i); });
            if (errLit) { tokens.push({ type: 'err', value: errLit }); i += errLit.length; continue; }

            if (ch === "'") {
                let j = i + 1, buf = '';
                while (j < n && src[j] !== "'") buf += src[j++];
                j++;
                if (src[j] === '!') j++;
                tokens.push({ type: 'sheet', value: buf });
                i = j;
                continue;
            }

            if (ch >= '0' && ch <= '9' || (ch === '.' && /[0-9]/.test(src[i + 1] || ''))) {
                let j = i;
                while (j < n && /[0-9.]/.test(src[j])) j++;
                if (src[j] === 'e' || src[j] === 'E') {
                    let k = j + 1;
                    if (src[k] === '+' || src[k] === '-') k++;
                    if (/[0-9]/.test(src[k] || '')) { j = k; while (j < n && /[0-9]/.test(src[j])) j++; }
                }
                tokens.push({ type: 'number', value: parseFloat(src.slice(i, j)) });
                i = j;
                continue;
            }

            if (src.startsWith('<=', i) || src.startsWith('>=', i) || src.startsWith('<>', i)) {
                tokens.push({ type: 'op', value: src.substr(i, 2) });
                i += 2;
                continue;
            }
            if ('=<>+-*/^&%'.indexOf(ch) !== -1) { tokens.push({ type: 'op', value: ch }); i++; continue; }
            if (ch === '(') { tokens.push({ type: 'lparen' }); i++; continue; }
            if (ch === ')') { tokens.push({ type: 'rparen' }); i++; continue; }
            if (ch === ';' || ch === ',') { tokens.push({ type: 'sep' }); i++; continue; }
            if (ch === ':') { tokens.push({ type: 'colon' }); i++; continue; }
            if (ch === '@') { tokens.push({ type: 'current' }); i++; continue; }

            if (WORD_START.test(ch)) {
                let j = i;
                while (j < n && WORD_CHAR.test(src[j])) j++;
                const word = src.slice(i, j);
                const ref = REF_RE.exec(word);
                if (src[j] === '(') {
                    tokens.push({ type: 'name', value: word });
                } else if (ref) {
                    tokens.push({
                        type: 'ref', col: ref[2].toUpperCase(), row: parseInt(ref[4], 10),
                        absCol: ref[1] === '$', absRow: ref[3] === '$'
                    });
                } else {
                    tokens.push({ type: 'name', value: word });
                }
                i = j;
                continue;
            }

            throw new Error('Непонятный символ «' + ch + '» (позиция ' + i + ')');
        }

        return tokens;
    }

    /* ===================== парсер (AST) ===================== */

    function parseFormula(input) {
        const tokens = tokenize(input);
        let pos = 0;

        const peek = function (offset) { return tokens[pos + (offset || 0)]; };
        const next = function () { return tokens[pos++]; };
        const isOp = function (value) { const t = peek(); return t && t.type === 'op' && value.indexOf(t.value) !== -1; };
        const isType = function (type) { const t = peek(); return t && t.type === type; };

        function readRefOrRange(sheet) {
            const t = next();
            if (!t || t.type !== 'ref') throw new Error('Ожидалась ссылка на ячейку');
            const first = { sheet: sheet || null, col: t.col, row: t.row, absCol: t.absCol, absRow: t.absRow };
            if (isType('colon')) {
                next();
                let rightSheet = sheet;
                if (isType('sheet')) rightSheet = next().value;
                const t2 = next();
                if (!t2 || t2.type !== 'ref') throw new Error('Ожидался конец диапазона');
                const second = { sheet: rightSheet || null, col: t2.col, row: t2.row, absCol: t2.absCol, absRow: t2.absRow };
                return { type: 'range', from: first, to: second };
            }
            return { type: 'ref', sheet: first.sheet, col: first.col, row: first.row, absCol: first.absCol, absRow: first.absRow };
        }

        function parseArgs() {
            const args = [];
            if (isType('rparen')) return args;
            args.push(parseExpression());
            while (isType('sep')) { next(); args.push(parseExpression()); }
            return args;
        }

        function parsePrimary() {
            const t = next();
            if (!t) throw new Error('Неожиданный конец формулы');

            if (t.type === 'number' || t.type === 'string') return { type: 'literal', value: t.value };
            if (t.type === 'err') return { type: 'literal', value: error(t.value) };
            if (t.type === 'current') return { type: 'current' };
            if (t.type === 'ref') { pos--; return readRefOrRange(null); }
            if (t.type === 'sheet') return readRefOrRange(t.value);
            if (t.type === 'lparen') {
                const inner = parseExpression();
                if (!isType('rparen')) throw new Error('Ожидалась закрывающая скобка');
                next();
                return inner;
            }
            if (t.type === 'name') {
                if (isType('lparen')) {
                    next();
                    const args = parseArgs();
                    if (!isType('rparen')) throw new Error('Ожидалась закрывающая скобка');
                    next();
                    return { type: 'call', name: t.value.toUpperCase(), args: args };
                }
                const boolName = BOOL_NAMES[t.value.toUpperCase()];
                if (boolName !== undefined) return { type: 'literal', value: boolName };
                throw new Error('Неизвестное имя «' + t.value + '»');
            }
            throw new Error('Неожиданный элемент формулы');
        }

        function parsePower() {
            const base = parsePrimary();
            if (isOp('^')) { next(); return { type: 'binary', op: '^', left: base, right: parseUnary() }; }
            return base;
        }

        function parseUnary() {
            if (isOp('+-')) { const op = next().value; return { type: 'unary', op: op, operand: parseUnary() }; }
            return parsePower();
        }

        function parseMultiplicative() {
            let left = parseUnary();
            while (isOp('*/')) { const op = next().value; left = { type: 'binary', op: op, left: left, right: parseUnary() }; }
            return left;
        }

        function parseAdditive() {
            let left = parseMultiplicative();
            while (isOp('+-')) { const op = next().value; left = { type: 'binary', op: op, left: left, right: parseMultiplicative() }; }
            return left;
        }

        function parseConcat() {
            let left = parseAdditive();
            while (isOp('&')) { next(); left = { type: 'binary', op: '&', left: left, right: parseAdditive() }; }
            return left;
        }

        function parseComparison() {
            let left = parseConcat();
            while (isOp('=<>') || isOp('<>') || isOp('<=') || isOp('>=')) {
                const op = next().value;
                left = { type: 'binary', op: op, left: left, right: parseConcat() };
            }
            return left;
        }

        function parseExpression() { return parseComparison(); }

        const ast = parseExpression();
        if (pos !== tokens.length) throw new Error('Лишние символы после формулы');
        return ast;
    }

    // Сдвигает строку ссылки, если она не абсолютная ($).
    function shiftRef(ref, row) {
        if (ref.absRow) return ref;
        return { type: 'ref', sheet: ref.sheet, col: ref.col, row: row, absCol: ref.absCol, absRow: false };
    }

    // Сдвигает относительные номера строк (для правил с диапазоном строк).
    function shiftRow(node, row) {
        switch (node.type) {
            case 'ref':
                return shiftRef(node, row);
            case 'range':
                // У концов диапазона нет поля type, поэтому сдвигаем их напрямую.
                return { type: 'range', from: shiftRef(node.from, row), to: shiftRef(node.to, row) };
            case 'unary':
                return { type: 'unary', op: node.op, operand: shiftRow(node.operand, row) };
            case 'binary':
                return { type: 'binary', op: node.op, left: shiftRow(node.left, row), right: shiftRow(node.right, row) };
            case 'call':
                return { type: 'call', name: node.name, args: node.args.map(function (a) { return shiftRow(a, row); }) };
            default:
                return node;
        }
    }

    /* ===================== вычислитель ===================== */

    const COMPARISONS = { '=': 0, '<>': 1, '<': 2, '>': 3, '<=': 4, '>=': 5 };

    function dedupe(cells) {
        const seen = {}, out = [];
        for (const c of cells) {
            const key = (c.sheet || '') + '!' + c.address;
            if (!seen[key]) { seen[key] = true; out.push(c); }
        }
        return out;
    }

    function createContext(workbook, current) {
        const sheets = (workbook && workbook.Sheets) || {};
        const names = Object.keys(sheets);
        const cache = {};

        function resolveSheet(name) {
            if (!name) return null;
            if (sheets[name]) return name;
            const key = String(name).toLowerCase();
            if (key in cache) return cache[key];
            const found = names.find(function (s) { return s.toLowerCase() === key; }) || null;
            cache[key] = found;
            return found;
        }

        function rawCell(sheet, col, row) {
            if (!sheet || !sheets[sheet]) return null;
            return sheets[sheet][cellAddress(col, row)] || null;
        }

        function readCell(sheet, col, row) {
            if (!sheet || !sheets[sheet]) return error(ERR.REF);
            const cell = rawCell(sheet, col, row);
            return cell ? cell.v : undefined;
        }

        const cur = (current && current.sheet) ? { sheet: resolveSheet(current.sheet), col: current.col, row: current.row } : null;

        return { sheets: sheets, resolveSheet: resolveSheet, readCell: readCell, rawCell: rawCell, current: cur };
    }

    function evalBinary(node, ctx) {
        const L = evalNode(node.left, ctx);
        const R = evalNode(node.right, ctx);
        const cells = dedupe(L.cells.concat(R.cells));
        const a = scalar(L.value), b = scalar(R.value);
        if (isError(a)) return { value: a, cells: cells };
        if (isError(b)) return { value: b, cells: cells };

        const op = node.op;
        if (op in COMPARISONS) {
            const c = compareValues(a, b);
            const result = op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
            return { value: result, cells: cells };
        }
        if (op === '&') return { value: toText(a) + toText(b), cells: cells };

        const x = toNumber(a), y = toNumber(b);
        if (isError(x)) return { value: x, cells: cells };
        if (isError(y)) return { value: y, cells: cells };
        if (op === '+') return { value: x + y, cells: cells };
        if (op === '-') return { value: x - y, cells: cells };
        if (op === '*') return { value: x * y, cells: cells };
        if (op === '/') return { value: y === 0 ? error(ERR.DIV0) : x / y, cells: cells };
        return { value: Math.pow(x, y), cells: cells };
    }

    function evalNode(node, ctx) {
        switch (node.type) {
            case 'literal':
                return { value: node.value, cells: [] };

            case 'current': {
                if (!ctx.current) return { value: error(ERR.REF), cells: [] };
                const sheet = ctx.current.sheet, col = ctx.current.col, row = ctx.current.row;
                const value = ctx.readCell(sheet, col, row);
                return { value: value, cells: [{ sheet: sheet, col: col, row: row, address: cellAddress(col, row), value: value }] };
            }

            case 'ref': {
                const sheet = ctx.resolveSheet(node.sheet);
                const value = ctx.readCell(sheet, node.col, node.row);
                return {
                    value: value,
                    cells: [{ sheet: sheet, col: node.col, row: node.row, address: cellAddress(node.col, node.row), value: value }]
                };
            }

            case 'range': {
                const sheet = ctx.resolveSheet(node.from.sheet || node.to.sheet);
                const rows = [], cells = [];
                const r1 = Math.min(node.from.row, node.to.row), r2 = Math.max(node.from.row, node.to.row);
                const c1 = Math.min(colToIndex(node.from.col), colToIndex(node.to.col));
                const c2 = Math.max(colToIndex(node.from.col), colToIndex(node.to.col));
                for (let r = r1; r <= r2; r++) {
                    const rowValues = [];
                    for (let c = c1; c <= c2; c++) {
                        const col = indexToCol(c);
                        const value = ctx.readCell(sheet, col, r);
                        rowValues.push(value);
                        cells.push({ sheet: sheet, col: col, row: r, address: cellAddress(col, r), value: value });
                    }
                    rows.push(rowValues);
                }
                return { value: { type: 'range', rows: rows }, cells: cells };
            }

            case 'unary': {
                const r = evalNode(node.operand, ctx);
                const n = toNumber(r.value);
                return { value: isError(n) ? n : (node.op === '-' ? -n : n), cells: dedupe(r.cells) };
            }

            case 'binary':
                return evalBinary(node, ctx);

            case 'call': {
                const impl = FUNCTIONS[node.name];
                if (!impl) return { value: error(ERR.NAME), cells: [] };
                return impl(node.args, ctx);
            }
        }
        return { value: error(ERR.NAME), cells: [] };
    }

    /* ---- помощники для функций ---- */

    function evalArgs(args, ctx) { return args.map(function (a) { return evalNode(a, ctx); }); }
    function flatCells(results) { return dedupe(results.reduce(function (acc, r) { return acc.concat(r.cells); }, [])); }

    function flatValues(results) {
        const out = [];
        for (const r of results) {
            if (isRange(r.value)) {
                for (const row of r.value.rows) for (const v of row) out.push(v);
            } else {
                out.push(scalar(r.value));
            }
        }
        return out;
    }

    function firstError(values) { return values.find(isError); }

    // Функция, все аргументы которой считаются заранее и сворачиваются в скаляры.
    function eager(fn) {
        return function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const vals = res.map(function (r) { return scalar(r.value); });
            const err = firstError(vals);
            return err ? { value: err, cells: cells } : { value: fn(vals), cells: cells };
        };
    }

    function aggregate(reducer, fallback) {
        return function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const vals = flatValues(res);
            const err = firstError(vals);
            if (err) return { value: err, cells: cells };
            const nums = vals.filter(function (v) { return typeof v === 'number'; });
            return { value: nums.length ? reducer(nums) : fallback, cells: cells };
        };
    }

    function predicate(fn) {
        return function (args, ctx) {
            const r = evalNode(args[0], ctx);
            return { value: fn(r.value), cells: dedupe(r.cells) };
        };
    }

    function numOf(vals, i) { const n = toNumber(vals[i]); return isError(n) ? NaN : n; }
    function digitsOf(vals, i) { const d = numOf(vals, i); return isNaN(d) ? 0 : Math.trunc(d); }
    function roundTo(x, d) { const f = Math.pow(10, d); return Math.round(x * f) / f; }

    /* ===================== реестр функций ===================== */

    const FUNCTIONS = {

        /* --- логические --- */
        'И': function (args, ctx) {
            const res = evalArgs(args, ctx);
            let result = true;
            for (const r of res) {
                const b = toBool(r.value);
                if (isError(b)) return { value: b, cells: flatCells(res) };
                if (!b) result = false;
            }
            return { value: result, cells: dedupe(res.filter(function (r) { return toBool(r.value) === result; }).reduce(function (a, r) { return a.concat(r.cells); }, [])) };
        },

        'ИЛИ': function (args, ctx) {
            const res = evalArgs(args, ctx);
            let result = false;
            for (const r of res) {
                const b = toBool(r.value);
                if (isError(b)) return { value: b, cells: flatCells(res) };
                if (b) result = true;
            }
            return { value: result, cells: dedupe(res.filter(function (r) { return toBool(r.value) === result; }).reduce(function (a, r) { return a.concat(r.cells); }, [])) };
        },

        'НЕ': function (args, ctx) {
            const r = evalNode(args[0], ctx);
            const b = toBool(r.value);
            return { value: isError(b) ? b : !b, cells: dedupe(r.cells) };
        },

        'ЕСЛИ': function (args, ctx) {
            const cond = evalNode(args[0], ctx);
            const b = toBool(cond.value);
            if (isError(b)) return { value: b, cells: dedupe(cond.cells) };
            const branch = b ? args[1] : args[2];
            if (!branch) return { value: b ? true : false, cells: dedupe(cond.cells) };
            const r = evalNode(branch, ctx);
            return { value: scalar(r.value), cells: dedupe(cond.cells.concat(r.cells)) };
        },

        'ЕСЛИОШИБКА': function (args, ctx) {
            const r = evalNode(args[0], ctx);
            if (isError(r.value) && args[1]) {
                const f = evalNode(args[1], ctx);
                return { value: scalar(f.value), cells: dedupe(r.cells.concat(f.cells)) };
            }
            return { value: scalar(r.value), cells: dedupe(r.cells) };
        },

        'ИСТИНА': function () { return { value: true, cells: [] }; },
        'ЛОЖЬ': function () { return { value: false, cells: [] }; },

        /* --- проверки --- */
        'ЕЧИСЛО': predicate(function (v) { return typeof v === 'number'; }),
        'ЕТЕКСТ': predicate(function (v) { return typeof v === 'string'; }),
        'ЕПУСТО': predicate(function (v) { return v === undefined || v === null; }),
        'ЕЛОГИЧ': predicate(function (v) { return typeof v === 'boolean'; }),
        'ЕОШИБКА': predicate(isError),
        'ЕНД': predicate(function (v) { return isError(v) && v.code === ERR.NA; }),
        'ЕФОРМУЛА': function (args, ctx) {
            const arg = args[0];
            let sheet, col, row;
            if (arg && arg.type === 'current') {
                if (!ctx.current) return { value: error(ERR.REF), cells: [] };
                sheet = ctx.current.sheet; col = ctx.current.col; row = ctx.current.row;
            } else if (arg && arg.type === 'ref') {
                sheet = ctx.resolveSheet(arg.sheet); col = arg.col; row = arg.row;
            } else {
                return { value: error(ERR.VALUE), cells: [] };
            }
            const cell = ctx.rawCell(sheet, col, row);
            const has = !!(cell && cell.f != null && cell.f !== '');
            return { value: has, cells: [{ sheet: sheet, col: col, row: row, address: cellAddress(col, row), value: cell ? cell.v : undefined }] };
        },

        /* --- текст --- */
        'ДЛСТР': eager(function (v) { return toText(v[0]).length; }),
        'СЖПРОБЕЛЫ': eager(function (v) { return toText(v[0]).replace(/\s+/g, ' ').trim(); }),
        'ПРОПИСН': eager(function (v) { return toText(v[0]).toUpperCase(); }),
        'СТРОЧН': eager(function (v) { return toText(v[0]).toLowerCase(); }),
        'ЛЕВСИМВ': eager(function (v) {
            const s = toText(v[0]);
            const n = v.length > 1 ? numOf(v, 1) : 1;
            return s.slice(0, isNaN(n) || n < 0 ? s.length : n);
        }),
        'ПРАВСИМВ': eager(function (v) {
            const s = toText(v[0]);
            const n = v.length > 1 ? numOf(v, 1) : 1;
            return isNaN(n) || n < 0 ? s : s.slice(-n);
        }),
        'ПСТР': eager(function (v) {
            const s = toText(v[0]);
            const start = numOf(v, 1), len = numOf(v, 2);
            if (isNaN(start) || isNaN(len) || start < 1 || len < 0) return error(ERR.VALUE);
            return s.substr(start - 1, len);
        }),
        'СЦЕПИТЬ': function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const vals = flatValues(res);
            const err = firstError(vals);
            if (err) return { value: err, cells: cells };
            return { value: vals.map(toText).join(''), cells: cells };
        },
        'ПОДСТАВИТЬ': eager(function (v) {
            const text = toText(v[0]), oldText = toText(v[1]), newText = toText(v[2]);
            if (v.length < 4) return text.split(oldText).join(newText);
            const n = numOf(v, 3);
            if (isNaN(n) || n < 1) return text;
            let count = 0, out = '', i = 0;
            while (i < text.length) {
                if (count < n && text.startsWith(oldText, i)) { out += newText; count++; i += oldText.length; }
                else { out += text[i++]; }
            }
            return out;
        }),
        'ПОВТОР': eager(function (v) {
            const n = numOf(v, 1);
            return isNaN(n) || n < 0 ? error(ERR.VALUE) : toText(v[0]).repeat(Math.trunc(n));
        }),
        'НАЙТИ': eager(function (v) { return findText(v, false); }),
        'ПОИСК': eager(function (v) { return findText(v, true); }),
        'ЗНАЧЕН': eager(function (v) { return toNumber(v[0]); }),

        /* --- математика --- */
        'СУММ': aggregate(function (nums) { return nums.reduce(function (a, b) { return a + b; }, 0); }, 0),
        'ПРОИЗВЕД': aggregate(function (nums) { return nums.reduce(function (a, b) { return a * b; }, 1); }, 0),
        'МИН': aggregate(function (nums) { return Math.min.apply(null, nums); }, 0),
        'МАКС': aggregate(function (nums) { return Math.max.apply(null, nums); }, 0),
        'СРЗНАЧ': aggregate(function (nums) { return nums.reduce(function (a, b) { return a + b; }, 0) / nums.length; }, error(ERR.DIV0)),
        'СЧЁТ': function (args, ctx) {
            const res = evalArgs(args, ctx);
            return { value: flatValues(res).filter(function (v) { return typeof v === 'number'; }).length, cells: flatCells(res) };
        },
        'СЧЁТЗ': function (args, ctx) {
            const res = evalArgs(args, ctx);
            return { value: flatValues(res).filter(function (v) { return !isEmpty(v); }).length, cells: flatCells(res) };
        },
        'ОКРУГЛ': eager(function (v) { return roundTo(numOf(v, 0), digitsOf(v, 1)); }),
        'ОКРУГЛВВЕРХ': eager(function (v) {
            const d = digitsOf(v, 1), f = Math.pow(10, d), x = numOf(v, 0);
            return x < 0 ? -Math.ceil(-x * f) / f : Math.ceil(x * f) / f;
        }),
        'ОКРУГЛВНИЗ': eager(function (v) {
            const d = digitsOf(v, 1), f = Math.pow(10, d), x = numOf(v, 0);
            return x < 0 ? -Math.floor(-x * f) / f : Math.floor(x * f) / f;
        }),
        'ЦЕЛОЕ': eager(function (v) { return Math.floor(numOf(v, 0)); }),
        'ОСТАТ': eager(function (v) {
            const a = numOf(v, 0), b = numOf(v, 1);
            return b === 0 ? error(ERR.DIV0) : a - b * Math.floor(a / b);
        }),
        'ABS': eager(function (v) { return Math.abs(numOf(v, 0)); }),
        'КОРЕНЬ': eager(function (v) { const x = numOf(v, 0); return x < 0 ? error(ERR.NUM) : Math.sqrt(x); }),
        'СТЕПЕНЬ': eager(function (v) { return Math.pow(numOf(v, 0), numOf(v, 1)); }),
        'ЗНАК': eager(function (v) { return Math.sign(numOf(v, 0)); }),

        /* --- поиск и ссылки --- */
        'ВПР': function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const lookup = scalar(res[0].value);
            const table = res[1] && res[1].value;
            if (!isRange(table)) return { value: error(ERR.VALUE), cells: cells };
            const colIdx = toNumber(scalar(res[2] ? res[2].value : undefined));
            if (isError(colIdx)) return { value: colIdx, cells: cells };
            const approx = res[3] ? toBool(scalar(res[3].value)) !== false : true;
            let found;
            for (const row of table.rows) {
                const c = compareValues(row[0], lookup);
                if (c === 0) { found = row[colIdx - 1]; break; }
                if (approx && c < 0) found = row[colIdx - 1];
                else if (approx && c > 0) break;
            }
            return { value: found === undefined ? error(ERR.NA) : scalar(found), cells: cells };
        },
        'ГПР': function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const lookup = scalar(res[0].value);
            const table = res[1] && res[1].value;
            if (!isRange(table)) return { value: error(ERR.VALUE), cells: cells };
            const rowIdx = toNumber(scalar(res[2] ? res[2].value : undefined));
            if (isError(rowIdx)) return { value: rowIdx, cells: cells };
            const head = table.rows[0] || [];
            let found;
            for (let i = 0; i < head.length; i++) if (compareValues(head[i], lookup) === 0) { found = (table.rows[rowIdx - 1] || [])[i]; break; }
            return { value: found === undefined ? error(ERR.NA) : scalar(found), cells: cells };
        },
        'ПОИСКПОЗ': function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const lookup = scalar(res[0].value);
            const table = res[1] && res[1].value;
            if (!isRange(table)) return { value: error(ERR.NA), cells: cells };
            const matchType = res[2] ? numOf([scalar(res[2].value)], 0) : 1;
            const list = table.rows.length > 1 ? table.rows.map(function (r) { return r[0]; }) : table.rows[0];
            let idx = -1;
            if (matchType === 0) {
                idx = list.findIndex(function (v) { return compareValues(v, lookup) === 0; });
            } else {
                for (let i = 0; i < list.length; i++) {
                    const c = compareValues(list[i], lookup);
                    if (matchType > 0 ? c <= 0 : c >= 0) idx = i;
                }
            }
            return { value: idx >= 0 ? idx + 1 : error(ERR.NA), cells: cells };
        },
        'ИНДЕКС': function (args, ctx) {
            const res = evalArgs(args, ctx);
            const cells = flatCells(res);
            const table = res[0] && res[0].value;
            if (!isRange(table)) return { value: error(ERR.VALUE), cells: cells };
            const rowN = res[1] ? numOf([scalar(res[1].value)], 0) : 1;
            const colN = res[2] ? numOf([scalar(res[2].value)], 0) : 1;
            if (rowN < 1 || rowN > table.rows.length) return { value: error(ERR.REF), cells: cells };
            const row = table.rows[rowN - 1];
            if (colN < 1 || colN > row.length) return { value: error(ERR.REF), cells: cells };
            return { value: row[colN - 1] === undefined ? error(ERR.NA) : scalar(row[colN - 1]), cells: cells };
        },
        'СТРОКА': function (args, ctx) {
            const node = args[0];
            if (!node) return { value: error(ERR.VALUE), cells: [] };
            if (node.type === 'ref') return { value: node.row, cells: dedupe(evalNode(node, ctx).cells) };
            if (node.type === 'range') return { value: Math.min(node.from.row, node.to.row), cells: dedupe(evalNode(node, ctx).cells) };
            return { value: error(ERR.VALUE), cells: [] };
        },
        'СТОЛБЕЦ': function (args, ctx) {
            const node = args[0];
            if (!node) return { value: error(ERR.VALUE), cells: [] };
            if (node.type === 'ref') return { value: colToIndex(node.col) + 1, cells: dedupe(evalNode(node, ctx).cells) };
            if (node.type === 'range') return { value: Math.min(colToIndex(node.from.col), colToIndex(node.to.col)) + 1, cells: dedupe(evalNode(node, ctx).cells) };
            return { value: error(ERR.VALUE), cells: [] };
        },
        'ЧСТРОК': function (args, ctx) {
            const node = args[0];
            if (node && node.type === 'range') return { value: Math.abs(node.to.row - node.from.row) + 1, cells: dedupe(evalNode(node, ctx).cells) };
            if (node && node.type === 'ref') return { value: 1, cells: dedupe(evalNode(node, ctx).cells) };
            return { value: error(ERR.VALUE), cells: [] };
        },
        'ЧИСЛСТОЛБ': function (args, ctx) {
            const node = args[0];
            if (node && node.type === 'range') return { value: Math.abs(colToIndex(node.to.col) - colToIndex(node.from.col)) + 1, cells: dedupe(evalNode(node, ctx).cells) };
            if (node && node.type === 'ref') return { value: 1, cells: dedupe(evalNode(node, ctx).cells) };
            return { value: error(ERR.VALUE), cells: [] };
        }
    };

    function findText(v, ignoreCase) {
        let haystack = toText(v[1]), needle = toText(v[0]);
        const start = v.length > 2 ? numOf(v, 2) : 1;
        if (isNaN(start) || start < 1) return error(ERR.VALUE);
        if (ignoreCase) { haystack = haystack.toLowerCase(); needle = needle.toLowerCase(); }
        const idx = haystack.indexOf(needle, start - 1);
        return idx === -1 ? error(ERR.VALUE) : idx + 1;
    }

    /* ===================== английские псевдонимы ===================== */

    const ALIASES = {
        AND: 'И', OR: 'ИЛИ', NOT: 'НЕ', IF: 'ЕСЛИ', IFERROR: 'ЕСЛИОШИБКА',
        TRUE: 'ИСТИНА', FALSE: 'ЛОЖЬ',
        ISNUMBER: 'ЕЧИСЛО', ISTEXT: 'ЕТЕКСТ', ISBLANK: 'ЕПУСТО', ISLOGICAL: 'ЕЛОГИЧ',
        ISERROR: 'ЕОШИБКА', ISNA: 'ЕНД', ISFORMULA: 'ЕФОРМУЛА',
        LEN: 'ДЛСТР', TRIM: 'СЖПРОБЕЛЫ', UPPER: 'ПРОПИСН', LOWER: 'СТРОЧН',
        LEFT: 'ЛЕВСИМВ', RIGHT: 'ПРАВСИМВ', MID: 'ПСТР', CONCAT: 'СЦЕПИТЬ', CONCATENATE: 'СЦЕПИТЬ',
        SUBSTITUTE: 'ПОДСТАВИТЬ', REPT: 'ПОВТОР', FIND: 'НАЙТИ', SEARCH: 'ПОИСК', VALUE: 'ЗНАЧЕН',
        SUM: 'СУММ', PRODUCT: 'ПРОИЗВЕД', MIN: 'МИН', MAX: 'МАКС', AVERAGE: 'СРЗНАЧ',
        COUNT: 'СЧЁТ', СЧЕТ: 'СЧЁТ', COUNTA: 'СЧЁТЗ',
        ROUND: 'ОКРУГЛ', ROUNDUP: 'ОКРУГЛВВЕРХ', ROUNDDOWN: 'ОКРУГЛВНИЗ',
        INT: 'ЦЕЛОЕ', MOD: 'ОСТАТ', SQRT: 'КОРЕНЬ', POWER: 'СТЕПЕНЬ', SIGN: 'ЗНАК', ABS: 'ABS',
        VLOOKUP: 'ВПР', HLOOKUP: 'ГПР', MATCH: 'ПОИСКПОЗ', INDEX: 'ИНДЕКС',
        ROW: 'СТРОКА', COLUMN: 'СТОЛБЕЦ', ROWS: 'ЧСТРОК', COLUMNS: 'ЧИСЛСТОЛБ'
    };

    for (const alias in ALIASES) {
        if (!FUNCTIONS[alias]) FUNCTIONS[alias] = FUNCTIONS[ALIASES[alias]];
    }

    /* ===================== публичный интерфейс ===================== */

    function evaluateAst(ast, workbook, current) {
        const ctx = createContext(workbook, current);
        const r = evalNode(ast, ctx);
        const value = scalar(r.value);
        const err = isError(value);
        const b = err ? false : toBool(value);
        return { value: value, cells: dedupe(r.cells), error: err, truthy: b === true };
    }

    function evaluateFormula(workbook, formula) {
        return evaluateAst(parseFormula(formula), workbook);
    }

    // Подстановка ${key} в текст сообщения об ошибке.
    function renderTemplate(template, data) {
        return String(template).replace(/\$\{(\w+)\}/g, function (match, key) {
            return Object.prototype.hasOwnProperty.call(data, key) ? String(data[key]) : match;
        });
    }

    const API = {
        tokenize: tokenize,
        parseFormula: parseFormula,
        shiftRow: shiftRow,
        evaluateAst: evaluateAst,
        evaluateFormula: evaluateFormula,
        renderTemplate: renderTemplate,
        colToIndex: colToIndex,
        indexToCol: indexToCol,
        isError: isError,
        ERROR: ERR
    };

    if (typeof module !== 'undefined' && module.exports) module.exports = API;
    else global.ExcelFormula = API;

})(typeof window !== 'undefined' ? window : globalThis);
