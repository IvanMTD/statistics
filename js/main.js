// Main JavaScript file for statistics project

let loadedWorkbook = null;

document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const checkBtn = document.getElementById('checkBtn');
    
    fileInput.addEventListener('change', function(event) {
        const file = event.target.files[0];
        
        if (file) {
            const fileName = file.name;
            const fileExtension = fileName.split('.').pop().toLowerCase();
            
            // Проверка расширения файла
            if (fileExtension !== 'xlsx' && fileExtension !== 'xls') {
                alert('Пожалуйста, выберите файл Excel (.xlsx или .xls)');
                fileInput.value = '';
                return;
            }
            
            // Чтение файла с помощью SheetJS
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                loadedWorkbook = XLSX.read(data, { type: 'array' });
                
                // Показываем кнопку "Проверить" после успешной загрузки
                checkBtn.classList.remove('d-none');
                
                // Скрываем результаты при новой загрузке
                document.getElementById('resultArea').classList.add('d-none');
            };
            reader.readAsArrayBuffer(file);
        }
    });
    
    checkBtn.addEventListener('click', function() {
        if (!loadedWorkbook) {
            alert('Файл не загружен');
            return;
        }
        
        checkSection0(loadedWorkbook);
    });
});

// Функция для преобразования буквы столбца в индекс (A=0, B=1, ..., Z=25, AA=26, ...)
function colToIndex(colStr) {
    let result = 0;
    for (let i = 0; i < colStr.length; i++) {
        result = result * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1;
}

// Функция для получения значения ячейки
function getCellValue(sheet, cellAddress) {
    const cell = sheet[cellAddress];
    return cell ? cell.v : null;
}

// Проверка Раздела 0
function checkSection0(workbook) {
    const sectionName = 'Раздел 0';
    const sheet = workbook.Sheets[sectionName];
    
    if (!sheet) {
        alert(`Лист "${sectionName}" не найден в файле`);
        return;
    }
    
    const errors = [];
    
    // H20 - Строки наименование отчитывающейся организации обязательны к заполнению
    const h20 = getCellValue(sheet, 'H20');
    if (h20 === null || h20 === '' || h20 === undefined) {
        errors.push('В Разделе 0 в ячейке H20 "Строки наименование отчитывающейся организации обязательны к заполнению"');
    }
    
    // O20 - Строки краткое наименование отчитывающейся организации обязательны к заполнению
    const o20 = getCellValue(sheet, 'O20');
    if (o20 === null || o20 === '' || o20 === undefined) {
        errors.push('В Разделе 0 в ячейке O20 "Строки краткое наименование отчитывающейся организации обязательны к заполнению"');
    }
    
    // E21 - Строка почтовый адрес обязательна к заполнению
    const e21 = getCellValue(sheet, 'E21');
    if (e21 === null || e21 === '' || e21 === undefined) {
        errors.push('В Разделе 0 в ячейке E21 "Строка почтовый адрес обязательна к заполнению"');
    }
    
    // B23, I23, M23, S23 - Строки населенного пункта – буквенные значения, цифрами заполнять должно быть запрещено
    const settlementCells = ['B23', 'I23', 'M23', 'S23'];
    settlementCells.forEach(cell => {
        const value = getCellValue(sheet, cell);
        if (value !== null && value !== '' && value !== undefined) {
            const strValue = String(value);
            // Проверяем, есть ли цифры в значении
            if (/\d/.test(strValue)) {
                errors.push(`В Разделе 0 в ячейке ${cell} "Строки населенного пункта – буквенные значения, цифрами заполнять должно быть запрещено"`);
            }
        }
    });
    
    // E27 -multiple checks
    const e27 = getCellValue(sheet, 'E27');
    if (e27 === null || e27 === '' || e27 === undefined) {
        errors.push('В Разделе 0 в ячейке E27 "В графе 2 прописывается код по Общероссийскому классификатору предприятий и организаций (ОКПО) или идентификационный номер. Графа обязательна к заполнению"');
    } else {
        const e27Str = String(e27);
        // Проверка на только цифровые значения
        if (!/^\d+$/.test(e27Str)) {
            errors.push('В Разделе 0 в ячейке E27 "В графе 2 только цифровые значения (текстовый формат – возможность вносить значения начинающиеся на 0)"');
        } else {
            // Проверка длины ОКПО (8, 9 или 14 цифр)
            const len = e27Str.length;
            if (len !== 8 && len !== 9 && len !== 14) {
                errors.push('В Разделе 0 в ячейке E27 "В графе 2 ОКПО должно равняться 8; 9 или 14 цифрам"');
            }
        }
    }
    
    // I27 - В графе 3 лицензия обязательна к заполнению
    const i27 = getCellValue(sheet, 'I27');
    if (i27 === null || i27 === '' || i27 === undefined) {
        errors.push('В Разделе 0 в ячейке I27 "В графе 3 кодовой части формы юридическое лицо проставляет регистрационный номер лицензии на осуществление образовательной деятельности в соответствии с реестром лицензий. Графа обязательна к заполнению"');
    }
    
    // M27 - В графе 4 код субъекта обязателен и должен быть однозначным или двухзначным числом
    const m27 = getCellValue(sheet, 'M27');
    if (m27 === null || m27 === '' || m27 === undefined) {
        errors.push('В Разделе 0 в ячейке M27 "В графе 4 кодовой части формы юридическое лицо проставляет код субъекта на основании классификатора адресов Российской Федерации. Графа обязательна к заполнению"');
    } else {
        const m27Str = String(m27);
        if (/^\d+$/.test(m27Str)) {
            const len = m27Str.length;
            if (len !== 1 && len !== 2) {
                errors.push('В Разделе 0 в ячейке M27 "В графе 4 только однозначное или двухзначное число"');
            }
        }
    }
    
    // P27 - В графе 5 ОКТМО обязателен, 8 или 11 цифр, только цифры
    const p27 = getCellValue(sheet, 'P27');
    if (p27 === null || p27 === '' || p27 === undefined) {
        errors.push('В Разделе 0 в ячейке P27 "В графе 5 кодовой части формы юридическое лицо проставляет Код ОКТМО. Графа обязательна к заполнению"');
    } else {
        const p27Str = String(p27);
        // Проверка на только цифровые значения
        if (!/^\d+$/.test(p27Str)) {
            errors.push('В Разделе 0 в ячейке P27 "В графе 5 текстовый формат ячейки с возможность вносить значения начинающиеся на 0, только цифровые значения"');
        } else {
            // Проверка длины ОКТМО (8 или 11 цифр)
            const len = p27Str.length;
            if (len !== 8 && len !== 11) {
                errors.push('В Разделе 0 в ячейке P27 "В графе 5, только восьмизначное число или одиннадцатизначное"');
            }
        }
    }
    
    displayResults(errors);
}

// Отображение результатов
function displayResults(errors) {
    const resultArea = document.getElementById('resultArea');
    const errorList = document.getElementById('errorList');
    
    errorList.innerHTML = '';
    
    if (errors.length === 0) {
        const li = document.createElement('li');
        li.className = 'list-group-item list-group-item-success';
        li.textContent = 'Ошибок не найдено. Раздел 0 заполнен корректно.';
        errorList.appendChild(li);
    } else {
        errors.forEach(error => {
            const li = document.createElement('li');
            li.className = 'list-group-item list-group-item-danger';
            li.textContent = error;
            errorList.appendChild(li);
        });
    }
    
    resultArea.classList.remove('d-none');
}
