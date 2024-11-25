function addArticle(button) {
    const articlesContainer = button.previousElementSibling;
    const articleBlock = document.createElement('div');
    articleBlock.className = 'article-block';
    articleBlock.innerHTML = `
        <label>Артикул/Наименование:</label>
        <input type="text" class="article" placeholder="Введите артикул/наименование" required>
    `;
    articlesContainer.appendChild(articleBlock);
}

function addPlace() {
    const placesContainer = document.getElementById('places-container');
    const placeBlock = document.createElement('div');
    placeBlock.className = 'place-block';
    placeBlock.innerHTML = `
        <label>Место №:</label>
        <input type="text" class="place-number" placeholder="Введите номер места" required>
        <div class="articles-container">
            <div class="article-block">
                <label>Артикул/Наименование:</label>
                <input type="text" class="article" placeholder="Введите артикул/наименование" required>
            </div>
        </div>
        <button type="button" onclick="addArticle(this)">Добавить артикул</button>
    `;
    placesContainer.appendChild(placeBlock);
}

function generateReport() {
    const orderNumber = document.getElementById('order-number').value;
    const recipient = document.getElementById('recipient').value;
    const places = document.querySelectorAll('.place-block');

    // Проверка, что ключевые данные заполнены
    if (!orderNumber || !recipient) {
        alert("Пожалуйста, заполните номер заказа и грузополучателя.");
        return;
    }

    // Подготовка данных для Excel
    const data = [
        ["Номер заказа:", orderNumber],
        ["Грузополучатель:", recipient],
        [],
        ["Место №", "Артикул/Наименование"]
    ];

    // Добавление данных мест и артикулов
    places.forEach((place, index) => {
        const placeNumber = `Место №${index + 1}`;
        const articles = place.querySelectorAll('.article');

        articles.forEach((article, articleIndex) => {
            data.push([articleIndex === 0 ? placeNumber : "", article.value]);
        });
    });

    // Генерация Excel-файла
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);

    XLSX.utils.book_append_sheet(wb, ws, 'Отчет');

    // Формируем имя файла на основе номера заказа
    const fileName = `${orderNumber}.xlsx`;
    XLSX.writeFile(wb, fileName);
}
