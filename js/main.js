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
        
        // Запускаем проверку через парсер
        setTimeout(async function() {
            const errors = await validateWorkbook(loadedWorkbook);
            displayResults(errors);
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
