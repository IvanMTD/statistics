// Parser for Excel-like formulas in ruleset.json

/**
 * Извлекает имя листа (раздела) и адрес ячейки из формулы
 * Пример: 'Раздел 0'!$H$20 -> { section: 'Раздел 0', cell: 'H20' }
 * Пример: 'Раздел 1'!$H$5=1 -> { section: 'Раздел 1', cell: 'H5' }
 */
function extractCellReference(formula) {
    // Удаляем знак = в начале если есть
    const cleanFormula = formula.replace(/^=/, '');
    
    // Регулярное выражение для поиска ссылок на ячейки вида 'Раздел X'!$A$1 или A1
    const regex = /'([^']+)'!\$?([A-Z]+)\$?(\d+)|\$?([A-Z]+)\$?(\d+)/g;
    const matches = [];
    let match;
    
    while ((match = regex.exec(cleanFormula)) !== null) {
        if (match[1]) {
            // Формат 'Раздел X'!$A$1
            matches.push({
                section: match[1],
                col: match[2],
                row: match[3],
                address: match[2] + match[3]
            });
        } else if (match[4]) {
            // Формат A1 (без указания листа - используем текущий)
            matches.push({
                section: null,
                col: match[4],
                row: match[5],
                address: match[4] + match[5]
            });
        }
    }
    
    return matches;
}

/**
 * Получает значение ячейки из книги
 */
function getCellValue(workbook, section, cellAddress) {
    const sheet = workbook.Sheets[section];
    if (!sheet) {
        return null;
    }
    const cell = sheet[cellAddress];
    return cell ? cell.v : null;
}

/**
 * Вычисляет результат формулы
 * Возвращает true если формула истинна (ошибка найдена), false если ложна
 */
function evaluateFormula(workbook, formula, errorDescription) {
    const cleanFormula = formula.replace(/^=/, '');
    const cells = extractCellReference(formula);
    
    if (cells.length === 0) {
        console.warn('Не удалось извлечь ссылки на ячейки из формулы:', formula);
        return false;
    }
    
    // Простая реализация для базовых проверок
    // Поддерживаемые функции: И, ИЛИ, НЕ, ЕЧИСЛО, ДЛСТР, СУММ
    
    // Проверка на пустую ячейку: 'Раздел 0'!$H$20=""
    const emptyCheckRegex = /'([^']+)'!\$?([A-Z]+)\$?(\d+)=""/;
    const emptyMatch = cleanFormula.match(emptyCheckRegex);
    if (emptyMatch) {
        const section = emptyMatch[1];
        const cellAddress = emptyMatch[2] + emptyMatch[3];
        const value = getCellValue(workbook, section, cellAddress);
        const isEmpty = (value === null || value === '' || value === undefined);
        
        if (isEmpty) {
            return {
                isError: true,
                section: section,
                cells: [cellAddress],
                description: errorDescription
            };
        }
        return null;
    }
    
    // Проверка на числовое значение: НЕ(ЕЧИСЛО('Раздел 0'!$E$27))
    const notNumberRegex = /НЕ\(ЕЧИСЛО\('([^']+)'!\$?([A-Z]+)\$?(\d+)\)\)/;
    const notNumberMatch = cleanFormula.match(notNumberRegex);
    if (notNumberMatch) {
        const section = notNumberMatch[1];
        const cellAddress = notNumberMatch[2] + notNumberMatch[3];
        const value = getCellValue(workbook, section, cellAddress);
        const isNotNumber = (value !== null && value !== '' && !/^\d+$/.test(String(value)));
        
        if (isNotNumber) {
            return {
                isError: true,
                section: section,
                cells: [cellAddress],
                description: errorDescription
            };
        }
        return null;
    }
    
    // Проверка длины строки: ИЛИ(ДЛСТР('Раздел 0'!$E$27)<>8; ...)
    const lengthCheckRegex = /ИЛИ\(([^)]+)\)/;
    const lengthMatch = cleanFormula.match(lengthCheckRegex);
    if (lengthMatch) {
        const innerContent = lengthMatch[1];
        // Извлекаем все проверки ДЛСТР
        const dlstrRegex = /ДЛСТР\('([^']+)'!\$?([A-Z]+)\$?(\d+)\)<>(\d+)/g;
        let dlstrMatch;
        const checks = [];
        
        while ((dlstrMatch = dlstrRegex.exec(innerContent)) !== null) {
            checks.push({
                section: dlstrMatch[1],
                cellAddress: dlstrMatch[2] + dlstrMatch[3],
                expectedLength: parseInt(dlstrMatch[4])
            });
        }
        
        if (checks.length > 0) {
            // Проверяем первую ячейку (для упрощения)
            const check = checks[0];
            const value = getCellValue(check.section, check.cellAddress);
            if (value !== null && value !== '') {
                const actualLength = String(value).length;
                const allLengths = checks.map(c => {
                    const val = getCellValue(c.section, c.cellAddress);
                    return val !== null && val !== '' ? String(val).length : null;
                });
                
                // Если хотя бы одна длина не совпадает с ожидаемой
                const hasError = allLengths.some(len => len !== null && len !== check.expectedLength);
                
                if (hasError) {
                    const cells = checks.map(c => c.cellAddress);
                    return {
                        isError: true,
                        section: checks[0].section,
                        cells: cells,
                        description: errorDescription
                    };
                }
            }
            return null;
        }
    }
    
    // Проверка на наличие цифр в тексте (для населенных пунктов)
    const numberInTextRegex = /ИЛИ\(ЕЧИСЛО\('([^']+)'!\$?([A-Z]+)\$?(\d+)\);/;
    const numberInTextMatch = cleanFormula.match(numberInTextRegex);
    if (numberInTextMatch) {
        // Это сложная проверка с несколькими ячейками
        const cellsInFormula = extractCellReference(formula);
        const section = cellsInFormula[0]?.section;
        
        if (section) {
            const errorCells = [];
            cellsInFormula.forEach(cellRef => {
                const value = getCellValue(workbook, cellRef.section || section, cellRef.address);
                if (value !== null && value !== '' && typeof value === 'number') {
                    errorCells.push(cellRef.address);
                }
            });
            
            if (errorCells.length > 0) {
                return {
                    isError: true,
                    section: section,
                    cells: errorCells,
                    description: errorDescription
                };
            }
        }
        return null;
    }
    
    // Сложная формула с И и СУММ (для Раздел 2, строки с диапазоном)
    // Пример: =И(СУММ('Раздел 2'!$Q11:$R11)<'Раздел 2'!$L11)
    const sumCompareRegex = /=И\(СУММ\('([^']+)'!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)([><]=?)'([^']+)'!\$?([A-Z]+)\$?(\d+)\)/;
    const sumMatch = cleanFormula.match(sumCompareRegex);
    if (sumMatch) {
        const sumSection = sumMatch[1];
        const sumStartCol = sumMatch[2];
        const sumStartRow = parseInt(sumMatch[3]);
        const sumEndCol = sumMatch[4];
        const sumEndRow = parseInt(sumMatch[5]);
        const operator = sumMatch[6];
        const compareSection = sumMatch[7];
        const compareCol = sumMatch[8];
        const compareRow = parseInt(sumMatch[9]);
        
        // Вычисляем сумму диапазона
        const sumResult = calculateSumRange(workbook, sumSection, sumStartCol + sumStartRow, sumEndCol + sumEndRow);
        // Получаем значение для сравнения
        const compareValue = getCellValue(workbook, compareSection, compareCol + compareRow);
        
        let conditionMet = false;
        if (operator === '<') conditionMet = sumResult < compareValue;
        else if (operator === '<=') conditionMet = sumResult <= compareValue;
        else if (operator === '>') conditionMet = sumResult > compareValue;
        else if (operator === '>=') conditionMet = sumResult >= compareValue;
        else if (operator === '=') conditionMet = sumResult === compareValue;
        else if (operator === '<>') conditionMet = sumResult !== compareValue;
        
        // Если формула возвращает TRUE -> это ошибка (по логике наших правил)
        if (conditionMet) {
            // Определяем ячейки для ошибки
            const errorCells = [];
            for (let c = colToIndex(sumStartCol); c <= colToIndex(sumEndCol); c++) {
                errorCells.push(indexToCol(c) + sumStartRow);
            }
            errorCells.push(compareCol + compareRow);
            
            return {
                isError: true,
                section: sumSection,
                cells: errorCells,
                description: errorDescription
            };
        }
        return null;
    }
    
    // Сложная формула с И, СУММ и <> (не равно)
    // Пример: =И(СУММ('Раздел 2'!$AE11:$AF11)<>'Раздел 2'!$K11)
    const sumNotEqualRegex = /=И\(СУММ\('([^']+)'!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)<>'([^']+)'!\$?([A-Z]+)\$?(\d+)\)/;
    const sumNeMatch = cleanFormula.match(sumNotEqualRegex);
    if (sumNeMatch) {
        const sumSection = sumNeMatch[1];
        const sumStartCol = sumNeMatch[2];
        const sumStartRow = parseInt(sumNeMatch[3]);
        const sumEndCol = sumNeMatch[4];
        const sumEndRow = parseInt(sumNeMatch[5]);
        const compareSection = sumNeMatch[6];
        const compareCol = sumNeMatch[7];
        const compareRow = parseInt(sumNeMatch[8]);
        
        const sumResult = calculateSumRange(workbook, sumSection, sumStartCol + sumStartRow, sumEndCol + sumEndRow);
        const compareValue = getCellValue(workbook, compareSection, compareCol + compareRow);
        
        // Ошибка если сумма НЕ равна значению
        if (sumResult !== compareValue) {
            const errorCells = [];
            for (let c = colToIndex(sumStartCol); c <= colToIndex(sumEndCol); c++) {
                errorCells.push(indexToCol(c) + sumStartRow);
            }
            errorCells.push(compareCol + compareRow);
            
            return {
                isError: true,
                section: sumSection,
                cells: errorCells,
                description: errorDescription
            };
        }
        return null;
    }
    
    return null;
}

/**
 * Вычисляет сумму значений в диапазоне ячеек
 */
function calculateSumRange(workbook, section, startCell, endCell) {
    const sheet = workbook.Sheets[section];
    if (!sheet) return 0;
    
    const startCol = colToIndex(startCell.replace(/\d/, ''));
    const startRow = parseInt(startCell.replace(/[A-Z]/, '')) - 1;
    const endCol = colToIndex(endCell.replace(/\d/, ''));
    const endRow = parseInt(endCell.replace(/[A-Z]/, '')) - 1;
    
    let sum = 0;
    
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = sheet[cellAddress];
            if (cell && typeof cell.v === 'number') {
                sum += cell.v;
            }
        }
    }
    
    return sum;
}

/**
 * Преобразует букву столбца в индекс
 */
function colToIndex(colStr) {
    let result = 0;
    for (let i = 0; i < colStr.length; i++) {
        result = result * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1;
}

/**
 * Преобразует индекс столбца обратно в букву
 */
function indexToCol(index) {
    let result = '';
    while (index >= 0) {
        result = String.fromCharCode((index % 26) + 'A'.charCodeAt(0)) + result;
        index = Math.floor(index / 26) - 1;
    }
    return result;
}

/**
 * Основная функция проверки всех правил из ruleset.json
 */
async function validateWorkbook(workbook) {
    try {
        const response = await fetch('ruleset.json');
        const rules = await response.json();
        
        const errors = [];
        
        for (const rule of rules) {
            // Если есть диапазон строк (rows), обрабатываем каждую строку отдельно
            if (rule.rows && rule.rows.length > 0) {
                // Парсим диапазон строк (например "11-323")
                const rowRange = rule.rows[0]; // Берем первый диапазон
                const [startRow, endRow] = rowRange.split('-').map(Number);
                
                for (let row = startRow; row <= endRow; row++) {
                    // Заменяем номер строки в формуле и сообщении об ошибке
                    let adaptedFormula = adaptFormulaForRow(rule.formula, row);
                    let adaptedError = adaptErrorForRow(rule.error, row);
                    
                    const result = evaluateFormula(workbook, adaptedFormula, adaptedError);
                    if (result && result.isError) {
                        errors.push({
                            section: result.section,
                            cell: result.cells.length === 1 ? result.cells[0] : result.cells.join(', '),
                            cells: result.cells,
                            description: result.description
                        });
                    }
                }
            } else {
                // Обычная проверка без диапазона
                const result = evaluateFormula(workbook, rule.formula, rule.error);
                if (result && result.isError) {
                    errors.push({
                        section: result.section,
                        cell: result.cells.length === 1 ? result.cells[0] : result.cells.join(', '),
                        cells: result.cells,
                        description: result.description
                    });
                }
            }
        }
        
        return errors;
    } catch (error) {
        console.error('Ошибка при загрузке или обработке ruleset.json:', error);
        return [];
    }
}

/**
 * Адаптирует формулу для конкретной строки
 * Заменяет все вхождения номеров строк в формуле на указанный номер
 */
function adaptFormulaForRow(formula, rowNum) {
    // Заменяем все номера строк в формуле (после $ или без)
    // Например: $Q11:$R11 -> $Q25:$R25, L11 -> L25
    return formula.replace(/\$(\d+)|(\d+)/g, (match, dollarRow, plainRow) => {
        if (dollarRow) {
            return '$' + rowNum;
        }
        if (plainRow) {
            // Проверяем, что это действительно номер строки (после буквы столбца)
            const precedingChar = formula[formula.indexOf(match) - 1];
            if (precedingChar && /[A-Z]/i.test(precedingChar)) {
                return rowNum;
            }
        }
        return match;
    });
}

/**
 * Адаптирует сообщение об ошибке для конкретной строки
 * Заменяет ??? на номер строки
 */
function adaptErrorForRow(errorTemplate, rowNum) {
    return errorTemplate.replace('???', rowNum);
}
