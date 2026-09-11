/*
 * Проверка движка формул и правил. Запуск: node tests/engine.test.js
 */
const assert = require('assert');
const engine = require('../js/formula-engine.js');
const { validateWorkbook } = require('../js/parser.js');

let passed = 0;
function check(name, fn) {
    fn();
    passed++;
    console.log('  ok - ' + name);
}

// Книга: лист -> адрес -> { v: значение }
function book(sheets) {
    const out = {};
    for (const name in sheets) {
        out[name] = {};
        for (const addr in sheets[name]) out[name][addr] = { v: sheets[name][addr] };
    }
    return { Sheets: out };
}

console.log('Адреса столбцов');
check('colToIndex/indexToCol', () => {
    assert.strictEqual(engine.colToIndex('A'), 0);
    assert.strictEqual(engine.colToIndex('Z'), 25);
    assert.strictEqual(engine.colToIndex('AA'), 26);
    assert.strictEqual(engine.colToIndex('AF'), 31);
    assert.strictEqual(engine.indexToCol(26), 'AA');
});

console.log('Пустая ячейка (правила 2-4, 6, 9-10, 12)');
check('пустая -> истина', () => {
    const r = engine.evaluateFormula(book({ 'Раздел 0': {} }), "'Раздел 0'!$H$20=\"\"");
    assert.strictEqual(r.truthy, true);
    assert.deepStrictEqual(r.cells.map(c => c.address), ['H20']);
});
check('заполненная -> ложь', () => {
    const r = engine.evaluateFormula(book({ 'Раздел 0': { H20: 'ООО Ромашка' } }), "'Раздел 0'!$H$20=\"\"");
    assert.strictEqual(r.truthy, false);
});

console.log('ДЛСТР (правила 8, 11, 14)');
check('9 цифр не ошибка (ОКПО 8/9/14)', () => {
    const b = book({ 'Раздел 0': { E27: '123456789' } });
    const r = engine.evaluateFormula(b, "И(ДЛСТР('Раздел 0'!$E$27)<>8; ДЛСТР('Раздел 0'!$E$27)<>9; ДЛСТР('Раздел 0'!$E$27)<>14)");
    assert.strictEqual(r.truthy, false);
});
check('10 цифр -> ошибка', () => {
    const b = book({ 'Раздел 0': { E27: '1234567890' } });
    const r = engine.evaluateFormula(b, "И(ДЛСТР('Раздел 0'!$E$27)<>8; ДЛСТР('Раздел 0'!$E$27)<>9; ДЛСТР('Раздел 0'!$E$27)<>14)");
    assert.strictEqual(r.truthy, true);
});
check('11 цифр ОКТМО не ошибка (правило 14)', () => {
    const b = book({ 'Раздел 0': { P27: '12345678901' } });
    const r = engine.evaluateFormula(b, "И(ДЛСТР('Раздел 0'!$P$27)<>8; ДЛСТР('Раздел 0'!$P$27)<>11)");
    assert.strictEqual(r.truthy, false);
});
check('правило 11 работает (было мёртвым)', () => {
    const bad = engine.evaluateFormula(book({ 'Раздел 0': { M27: '123' } }), "ИЛИ(ДЛСТР('Раздел 0'!$M$27)<1; ДЛСТР('Раздел 0'!$M$27)>2)");
    const good = engine.evaluateFormula(book({ 'Раздел 0': { M27: '12' } }), "ИЛИ(ДЛСТР('Раздел 0'!$M$27)<1; ДЛСТР('Раздел 0'!$M$27)>2)");
    assert.strictEqual(bad.truthy, true);
    assert.strictEqual(good.truthy, false);
});

console.log('ЕЧИСЛО (правила 5, 7, 13)');
check('число в населённом пункте -> ошибка, только для числовой ячейки', () => {
    const b = book({ 'Раздел 0': { B23: 'Москва', I23: 12, M23: 'Тверь', S23: 'Псков' } });
    const r = engine.evaluateFormula(b, "ИЛИ(ЕЧИСЛО('Раздел 0'!$B$23); ЕЧИСЛО('Раздел 0'!$I$23); ЕЧИСЛО('Раздел 0'!$M$23); ЕЧИСЛО('Раздел 0'!$S$23))");
    assert.strictEqual(r.truthy, true);
    assert.deepStrictEqual(r.cells.map(c => c.address), ['I23']);
});

console.log('И + сравнение (правила 1, 18 — ранее мёртвые)');
check('правило 1 срабатывает при незаполненном Разделе 2', () => {
    const b = book({ 'Раздел 1': { H5: 1 }, 'Раздел 2': { I323: 0 } });
    const r = engine.evaluateFormula(b, "=И('Раздел 1'!$H$5=1; СУММ('Раздел 2'!$I$323:$AF$323)>0)");
    assert.strictEqual(r.truthy, false);
    const b2 = book({ 'Раздел 1': { H5: 1 }, 'Раздел 2': { I323: 5 } });
    assert.strictEqual(engine.evaluateFormula(b2, "=И('Раздел 1'!$H$5=1; СУММ('Раздел 2'!$I$323:$AF$323)>0)").truthy, true);
});
check('правило 18 срабатывает', () => {
    const b = book({ 'Раздел 8': { J11: 0, K11: 0 }, 'Раздел 2': { I11: 1 } });
    const r = engine.evaluateFormula(b, "=И(СУММ('Раздел 8'!$J11:$K11)=0;'Раздел 2'!$I11=1)");
    assert.strictEqual(r.truthy, true);
});

console.log('Суммовые правила и плейсхолдеры');
async function checkRule15() {
    const rules = [{
        id: 15,
        formula: "=И(СУММ('Раздел 2'!$Q11:$R11)<'Раздел 2'!$L11)",
        error: "Раздел ${section}, строка ${row}: графы 11 и 12 (${value}) | ${cells}"
    }];
    const b = book({ 'Раздел 2': { Q11: 3, R11: 4, L11: 10 } });
    const res = await validateWorkbook(b, rules);
    assert.strictEqual(res.length, 1);
    assert.strictEqual(res[0].section, 'Раздел 2');
    assert.strictEqual(res[0].cell, 'Q11, R11, L11');
    assert.deepStrictEqual(res[0].cellList, ['Q11', 'R11', 'L11']);
    assert.strictEqual(res[0].description, 'Раздел Раздел 2, строка 11: графы 11 и 12 (3, 4, 10) | Q11, R11, L11');
    passed++;
    console.log('  ok - правило 15: сумма < L -> ошибка, подстановки корректны');
}

console.log('Шаблоны');
check('renderTemplate оставляет неизвестные ключи как есть', () => {
    assert.strictEqual(engine.renderTemplate('${row}/${col}/${nope}', { row: 5, col: 'A' }), '5/A/${nope}');
});

// Асинхронные проверки через validateWorkbook
(async function () {
    await checkRule15();

    console.log('Интеграция с ruleset.json (диапазон строк)');
    const rules = require('../ruleset.json');

    // Полностью корректно заполненная книга: ни одной ошибки.
    // ОКПО и ОКТМО — числовые ячейки (как при обычном вводе цифр в Excel).
    const header = {
        H20: 'ООО Пример', O20: 'Пример', E21: 'г. Москва',
        B23: 'Москва', I23: 'Москва', M23: 'Тверь', S23: 'Псков',
        E27: 12345678, I27: 'Л-1234', M27: '12', P27: 12345678
    };
    const goodRaw = {
        'Раздел 0': header,
        'Раздел 1': { H5: 1 },
        'Раздел 2': { Q11: 10, R11: 10, L11: 10, AE11: 5, AF11: 5, K11: 10, N11: 0 },
        'Раздел 3': { M11: 0, N11: 0, O11: 0, P11: 0, Q11: 0, R11: 0 },
        'Раздел 8': { J11: 1, K11: 1 }
    };
    const withRows = function (sheet, cells) {
        const copy = Object.assign({}, goodRaw);
        copy[sheet] = Object.assign({}, goodRaw[sheet], cells);
        return book(copy);
    };

    const errorsGood = await validateWorkbook(book(goodRaw), rules);
    assert.deepStrictEqual(errorsGood.map(e => e.description), [], 'корректная книга должна быть без ошибок');
    passed++; console.log('  ok - корректная книга -> 0 ошибок');

    // Пустая книга: срабатывают только проверки шапки (правила 2-14), не диапазонные.
    const emptyBook = book({
        'Раздел 0': {}, 'Раздел 1': {}, 'Раздел 2': {}, 'Раздел 3': {}, 'Раздел 8': {}
    });
    const errorsEmpty = await validateWorkbook(emptyBook, rules);
    assert.ok(errorsEmpty.length > 0, 'пустая книга: ожидали ошибки шапки');
    assert.ok(errorsEmpty.every(e => e.section === 'Раздел 0'),
        'диапазонные правила 15-18 не должны срабатывать на пустом листе');
    passed++; console.log('  ok - пустая книга -> ' + errorsEmpty.length + ' ошибок только по шапке');

    // Строка 12: ломаем правило 15 и проверяем подстановки ${section}/${row}/${value}.
    const errors12 = await validateWorkbook(
        withRows('Раздел 2', { Q12: 1, R12: 1, L12: 10 }), rules);
    assert.strictEqual(errors12.length, 1, 'ожидали ровно одну ошибку (правило 15, строка 12)');
    assert.strictEqual(errors12[0].description, 'Раздел 2, строка 12: заполните графы 11 и 12 (сейчас 1, 1, 10)');
    passed++; console.log('  ok - строка 12 из диапазона: ' + errors12[0].description);

    // Строка 13: правило 17 тянет данные из двух листов -> ${cells} с именами листов.
    const errors13 = await validateWorkbook(withRows('Раздел 2', { N13: 5 }), rules);
    assert.strictEqual(errors13.length, 1);
    assert.strictEqual(errors13[0].section, 'Раздел 3');
    assert.strictEqual(errors13[0].description,
        'Раздел 3, строка 13: неправильно заполнены графы 7 - 12 (ячейки Раздел 3!M13, Раздел 3!N13, Раздел 3!O13, Раздел 3!P13, Раздел 3!Q13, Раздел 3!R13, Раздел 2!N13)');
    assert.deepStrictEqual(errors13[0].cellList,
        ['Раздел 3!M13', 'Раздел 3!N13', 'Раздел 3!O13', 'Раздел 3!P13', 'Раздел 3!Q13', 'Раздел 3!R13', 'Раздел 2!N13']);
    passed++; console.log('  ok - межлистовое правило 17: ' + errors13[0].cell);

    // Отсутствующие листы -> #ССЫЛКА! -> правила не срабатывают (нет лавины ложных ошибок).
    const errorsMissing = await validateWorkbook(book({ 'Раздел 0': header }), rules);
    assert.deepStrictEqual(errorsMissing.map(e => e.description), [], 'отсутствующие листы не должны давать ошибок');
    passed++; console.log('  ok - отсутствующие листы -> 0 ложных ошибок');

    console.log('\nВсе проверки пройдены: ' + passed);
})().catch(e => { console.error(e); process.exit(1); });
