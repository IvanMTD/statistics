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
        
        // Показываем оверлей и блокируем взаимодействие
        showLoadingOverlay();
        
        // Имитируем небольшую задержку для демонстрации оверлея
        setTimeout(function() {
            checkSection0(loadedWorkbook);
            hideLoadingOverlay();
        }, 500);
    });
});

// Показать оверлей загрузки
function showLoadingOverlay() {
    document.getElementById('loadingOverlay').classList.remove('d-none');
    document.body.classList.add('loading');
}

// Скрыть оверлей загрузки
function hideLoadingOverlay() {
    document.getElementById('loadingOverlay').classList.add('d-none');
    document.body.classList.remove('loading');
}

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
        errors.push({
            section: 'Раздел 0',
            cell: 'H20',
            description: 'Строки наименование отчитывающейся организации обязательны к заполнению'
        });
    }
    
    // O20 - Строки краткое наименование отчитывающейся организации обязательны к заполнению
    const o20 = getCellValue(sheet, 'O20');
    if (o20 === null || o20 === '' || o20 === undefined) {
        errors.push({
            section: 'Раздел 0',
            cell: 'O20',
            description: 'Строки краткое наименование отчитывающейся организации обязательны к заполнению'
        });
    }
    
    // E21 - Строка почтовый адрес обязательна к заполнению
    const e21 = getCellValue(sheet, 'E21');
    if (e21 === null || e21 === '' || e21 === undefined) {
        errors.push({
            section: 'Раздел 0',
            cell: 'E21',
            description: 'Строка почтовый адрес обязательна к заполнению'
        });
    }
    
    // B23, I23, M23, S23 - Строки населенного пункта – буквенные значения, цифрами заполнять должно быть запрещено
    const settlementCells = ['B23', 'I23', 'M23', 'S23'];
    settlementCells.forEach(cell => {
        const value = getCellValue(sheet, cell);
        if (value !== null && value !== '' && value !== undefined) {
            const strValue = String(value);
            // Проверяем, есть ли цифры в значении
            if (/\d/.test(strValue)) {
                errors.push({
                    section: 'Раздел 0',
                    cell: cell,
                    description: 'Строки населенного пункта – буквенные значения, цифрами заполнять должно быть запрещено'
                });
            }
        }
    });
    
    // E27 - multiple checks
    const e27 = getCellValue(sheet, 'E27');
    if (e27 === null || e27 === '' || e27 === undefined) {
        errors.push({
            section: 'Раздел 0',
            cell: 'E27',
            description: 'В графе 2 прописывается код по Общероссийскому классификатору предприятий и организаций (ОКПО) или идентификационный номер. Графа обязательна к заполнению'
        });
    } else {
        const e27Str = String(e27);
        // Проверка на только цифровые значения
        if (!/^\d+$/.test(e27Str)) {
            errors.push({
                section: 'Раздел 0',
                cell: 'E27',
                description: 'В графе 2 только цифровые значения (текстовый формат – возможность вносить значения начинающиеся на 0)'
            });
        } else {
            // Проверка длины ОКПО (8, 9 или 14 цифр)
            const len = e27Str.length;
            if (len !== 8 && len !== 9 && len !== 14) {
                errors.push({
                    section: 'Раздел 0',
                    cell: 'E27',
                    description: 'В графе 2 ОКПО должно равняться 8; 9 или 14 цифрам'
                });
            }
        }
    }
    
    // I27 - В графе 3 лицензия обязательна к заполнению
    const i27 = getCellValue(sheet, 'I27');
    if (i27 === null || i27 === '' || i27 === undefined) {
        errors.push({
            section: 'Раздел 0',
            cell: 'I27',
            description: 'В графе 3 кодовой части формы юридическое лицо проставляет регистрационный номер лицензии на осуществление образовательной деятельности в соответствии с реестром лицензий. Графа обязательна к заполнению'
        });
    }
    
    // M27 - В графе 4 код субъекта обязателен и должен быть однозначным или двухзначным числом
    const m27 = getCellValue(sheet, 'M27');
    if (m27 === null || m27 === '' || m27 === undefined) {
        errors.push({
            section: 'Раздел 0',
            cell: 'M27',
            description: 'В графе 4 кодовой части формы юридическое лицо проставляет код субъекта на основании классификатора адресов Российской Федерации. Графа обязательна к заполнению'
        });
    } else {
        const m27Str = String(m27);
        if (/^\d+$/.test(m27Str)) {
            const len = m27Str.length;
            if (len !== 1 && len !== 2) {
                errors.push({
                    section: 'Раздел 0',
                    cell: 'M27',
                    description: 'В графе 4 только однозначное или двухзначное число'
                });
            }
        }
    }
    
    // P27 - В графе 5 ОКТМО обязателен, 8 или 11 цифр, только цифры
    const p27 = getCellValue(sheet, 'P27');
    if (p27 === null || p27 === '' || p27 === undefined) {
        errors.push({
            section: 'Раздел 0',
            cell: 'P27',
            description: 'В графе 5 кодовой части формы юридическое лицо проставляет Код ОКТМО. Графа обязательна к заполнению'
        });
    } else {
        const p27Str = String(p27);
        // Проверка на только цифровые значения
        if (!/^\d+$/.test(p27Str)) {
            errors.push({
                section: 'Раздел 0',
                cell: 'P27',
                description: 'В графе 5 текстовый формат ячейки с возможность вносить значения начинающиеся на 0, только цифровые значения'
            });
        } else {
            // Проверка длины ОКТМО (8 или 11 цифр)
            const len = p27Str.length;
            if (len !== 8 && len !== 11) {
                errors.push({
                    section: 'Раздел 0',
                    cell: 'P27',
                    description: 'В графе 5, только восьмизначное число или одиннадцатизначное'
                });
            }
        }
    }
    
    displayResults(errors);
}

// Отображение результатов
function displayResults(errors) {
    const resultArea = document.getElementById('resultArea');
    const mainCard = document.getElementById('mainCard');
    const uploadArea = document.getElementById('uploadArea');
    const checkBtn = document.getElementById('checkBtn');
    
    // Очищаем предыдущие результаты
    resultArea.innerHTML = '';
    
    // Скрываем элементы загрузки и кнопку
    uploadArea.classList.add('d-none');
    checkBtn.classList.add('d-none');
    
    // Увеличиваем карточку если есть ошибки
    if (errors.length > 0) {
        mainCard.classList.add('expanded');
    }
    
    if (errors.length === 0) {
        // Нет ошибок - показываем зеленую галочку
        resultArea.innerHTML = `
            <div class="result-success">
                <svg xmlns="http://www.w3.org/2000/svg" width="80" height="80" fill="currentColor" class="bi bi-check-circle" viewBox="0 0 16 16">
                    <path d="M8 15A7 7 0 1 1 8 1a7 7 0 0 1 0 14zm0 1A8 8 0 1 0 8 0a8 8 0 0 0 0 16z"/>
                    <path d="M10.97 4.97a.235.235 0 0 0-.02.022L7.477 9.417 5.384 7.323a.75.75 0 0 0-1.06 1.06L6.97 11.03a.75.75 0 0 0 1.079-.02l3.992-4.99a.75.75 0 0 0-1.071-1.05z"/>
                </svg>
                <h4 class="mt-3">Ошибок нет</h4>
            </div>
        `;
    } else {
        // Есть ошибки - выводим список с тегами
        let errorsHtml = '<h6 class="text-danger fw-bold mb-3">Найдены ошибки:</h6>';
        errorsHtml += '<div class="error-list">';
        
        errors.forEach(error => {
            errorsHtml += `
                <div class="error-item">
                    <span class="error-tag tag-section">${error.section}</span>
                    <span class="error-tag tag-cell">${error.cell}</span>
                    <span class="error-description">${error.description}</span>
                </div>
            `;
        });
        
        errorsHtml += '</div>';
        resultArea.innerHTML = errorsHtml;
    }
    
    resultArea.classList.remove('d-none');
}
