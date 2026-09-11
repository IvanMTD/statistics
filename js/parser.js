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
    
    // Сложная формула с И и СУММ
    const complexFormulaRegex = /=И\('([^']+)'!\$?([A-Z]+)\$?(\d+)=(\d+);\s*СУММ\('([^']+)'!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)([><]=?)(\d+)\)/;
    const complexMatch = cleanFormula.match(complexFormulaRegex);
    if (complexMatch) {
        const conditionSection = complexMatch[1];
        const conditionCell = complexMatch[2] + complexMatch[3];
        const conditionValue = parseInt(complexMatch[4]);
        
        const sumSection = complexMatch[5];
        const sumStart = complexMatch[6] + complexMatch[7];
        const sumEnd = complexMatch[8] + complexMatch[9];
        const sumOperator = complexMatch[10];
        const sumThreshold = parseFloat(complexMatch[11]);
        
        const condValue = getCellValue(workbook, conditionSection, conditionCell);
        
        if (condValue === conditionValue) {
            // Условие выполнено, проверяем сумму
            const sumResult = calculateSumRange(workbook, sumSection, sumStart, sumEnd);
            
            let sumConditionMet = false;
            if (sumOperator === '>') sumConditionMet = sumResult > sumThreshold;
            else if (sumOperator === '>=') sumConditionMet = sumResult >= sumThreshold;
            else if (sumOperator === '<') sumConditionMet = sumResult < sumThreshold;
            else if (sumOperator === '<=') sumConditionMet = sumResult <= sumThreshold;
            else if (sumOperator === '=') sumConditionMet = sumResult === sumThreshold;
            else if (sumOperator === '<>') sumConditionMet = sumResult !== sumThreshold;
            
            if (!sumConditionMet) {
                return {
                    isError: true,
                    section: conditionSection,
                    cells: [conditionCell],
                    description: errorDescription
                };
            }
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
 * Основная функция проверки всех правил из ruleset.json
 */
async function validateWorkbook(workbook) {
    try {
        const response = await fetch('ruleset.json');
        const rules = await response.json();
        
        const errors = [];
        
        for (const rule of rules) {
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
        
        return errors;
    } catch (error) {
        console.error('Ошибка при загрузке или обработке ruleset.json:', error);
        return [];
    }
}
