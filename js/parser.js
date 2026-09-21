/*
 * Проверка книги Excel по правилам из внешнего ruleset.json.
 * Вся работа с формулами делегируется движку js/formula-engine.js.
 */
(function (global) {
    'use strict';

    const RULESET_URL = 'https://cdn.s3.fcpsr.ru/stat/ruleset.json';

    const ENGINE = (typeof module !== 'undefined' && module.exports)
        ? require('./formula-engine.js')
        : global.ExcelFormula;

    // "11-323" -> [11..323], "11" -> [11]
    function parseRowList(rows) {
        const out = [];
        if (!rows) return out;
        for (const item of rows) {
            const range = String(item).match(/^(\d+)\s*-\s*(\d+)$/);
            if (range) {
                for (let r = +range[1]; r <= +range[2]; r++) out.push(r);
            } else if (/^\d+$/.test(String(item).trim())) {
                out.push(+item);
            }
        }
        return out;
    }

    // "Раздел 2!I11:J16" или "I11:J16" -> { sheet, from:{col,row}, to:{col,row} }
    function parseRangeSpec(spec) {
        const s = String(spec).trim();
        const bang = s.lastIndexOf('!');
        let sheet = null, rangeStr = s;
        if (bang !== -1) {
            sheet = s.slice(0, bang).replace(/^'/, '').replace(/'$/, '');
            rangeStr = s.slice(bang + 1);
        }
        const m = rangeStr.match(/^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/);
        if (!m) return null;
        return {
            sheet: sheet,
            from: { col: m[1].toUpperCase(), row: +m[2] },
            to: m[3] ? { col: m[3].toUpperCase(), row: +m[4] } : { col: m[1].toUpperCase(), row: +m[2] }
        };
    }

    function unique(list) { return Array.from(new Set(list)); }

    function dedupeCells(cells) {
        const seen = {}, out = [];
        for (const c of cells) {
            const key = (c.sheet || '') + '!' + c.address;
            if (!seen[key]) { seen[key] = true; out.push(c); }
        }
        return out;
    }

    function formatValue(value) {
        if (value === undefined || value === null || value === '') return '(пусто)';
        if (ENGINE.isError(value)) return value.code;
        if (typeof value === 'boolean') return value ? 'ИСТИНА' : 'ЛОЖЬ';
        return String(value);
    }

    // Собирает данные для плейсхолдеров ${section}, ${row}, ${col}, ${value}, ${cells}.
    function buildError(rule, result, row) {
        const cells = result.cells;
        const section = rule.section || (cells[0] && cells[0].sheet) || '';
        const sameSheet = cells.every(function (c) { return c.sheet === section; });

        // Имя листа добавляется только когда правило тянет данные из разных листов.
        const labels = cells.map(function (c) { return sameSheet ? c.address : ((c.sheet || '') + '!' + c.address); });

        const data = {
            section: section,
            row: unique(cells.map(function (c) { return c.row; })).join(', ') || (row === null ? '' : String(row)),
            col: unique(cells.map(function (c) { return c.col; })).join(', '),
            value: cells.map(function (c) { return formatValue(c.value); }).join(', '),
            cells: labels.join(', ')
        };

        return {
            section: section,
            cell: labels.join(', '),
            cellList: labels,
            cells: cells.map(function (c) { return c.address; }),
            description: ENGINE.renderTemplate(rule.error || '', data)
        };
    }

    async function validateWorkbook(workbook, rulesInput) {
        let rules = rulesInput;
        try {
            if (!rules) rules = await (await fetch(RULESET_URL)).json();
        } catch (e) {
            console.error('Не удалось загрузить ' + RULESET_URL + ':', e);
            return [];
        }

        const errors = [];

        for (const rule of rules) {
            // Запись без formula — служебный шаблон-описание, а не правило.
            if (!rule || !rule.formula) continue;

            let ast;
            try {
                ast = ENGINE.parseFormula(rule.formula);
            } catch (e) {
                console.error('Правило ' + (rule.id === undefined ? '?' : rule.id) + ': ошибка разбора формулы «' + rule.formula + '»', e);
                continue;
            }

            if (rule.ranges && rule.ranges.length) {
                const found = [];
                for (const spec of rule.ranges) {
                    const r = parseRangeSpec(spec);
                    if (!r) {
                        console.error('Правило ' + (rule.id === undefined ? '?' : rule.id) + ': не удалось разобрать диапазон «' + spec + '»');
                        continue;
                    }
                    const c1 = ENGINE.colToIndex(r.from.col), c2 = ENGINE.colToIndex(r.to.col);
                    for (let row = r.from.row; row <= r.to.row; row++) {
                        for (let c = c1; c <= c2; c++) {
                            const result = ENGINE.evaluateAst(ast, workbook, { sheet: r.sheet, col: ENGINE.indexToCol(c), row: row });
                            if (result.truthy) found.push.apply(found, result.cells);
                        }
                    }
                }
                const cells = dedupeCells(found);
                if (cells.length) errors.push(buildError(rule, { cells: cells }, null));
                continue;
            }

            const rows = parseRowList(rule.rows);
            const targets = rows.length ? rows : [null];

            for (const row of targets) {
                const node = row === null ? ast : ENGINE.shiftRow(ast, row);
                const result = ENGINE.evaluateAst(node, workbook);
                if (result.truthy) errors.push(buildError(rule, result, row));
            }
        }

        return errors;
    }

    if (typeof module !== 'undefined' && module.exports) module.exports = { validateWorkbook: validateWorkbook };
    else global.validateWorkbook = validateWorkbook;

})(typeof window !== 'undefined' ? window : globalThis);
