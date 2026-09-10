// Main JavaScript file for statistics project

document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    
    fileInput.addEventListener('change', function(event) {
        const file = event.target.files[0];
        
        if (file) {
            const fileName = file.name;
            const fileExtension = fileName.split('.').pop().toLowerCase();
            
            // Проверка расширения файла (уже ограничено через accept, но добавим дополнительную проверку)
            if (fileExtension !== 'xlsx' && fileExtension !== 'xls') {
                alert('Пожалуйста, выберите файл Excel (.xlsx или .xls)');
                fileInput.value = '';
                return;
            }
            
            console.log('Выбран файл:', fileName);
            // Здесь будет дальнейшая логика обработки файла
        }
    });
});
