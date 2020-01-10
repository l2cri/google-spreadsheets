function OrderForm() {
    var sheetForm = SpreadsheetApp.getActiveSheet();

    var readRecord = function (range) {
        var data = sheetForm.getRange(range).getValues();
        if (!data || !data[0]) return [];

        var res = data[0]; // первый столбец

        return {
            order: res[0], // ЯЛ-x/XXXXX
            num: res[1], // 1
            client: res[2], // Кондратьев Иван Сергеевич
            payment: res[3], // 1-й ПЛАТЕЖ
            sum: res[4], // 5000
            day: res[5], // 12.12.1990
            type: res[6], // наличные
        };
    }

    var findOrder = function (orderNum) {
        var column = 20; // индекс колонки поиска  (T)
        var columnValues = sheetForm.getRange(2, column, sheetForm.getLastRow()).getValues(); //1st is header row
        var searchResult = columnValues.findIndex(orderNum); // Row Index - 2

        if (searchResult !== -1) {
            return searchResult + 2;
        }

        return false;
    }

    // поиск по платежам
    var detectTypePayment = function (record) {
        var col = findOrder(record.order);

        var paymentRegular = record.payment;

        // куда вставлять диапозон начало:конец
        if (paymentRegular === 'Авансовый платеж(50%)') return 'Z' + col + ':AB' + col;
        if (paymentRegular === '1-й ПЛАТЕЖ') return 'AD' + col + ':AF' + col;
        if (paymentRegular === '2-й ПЛАТЕЖ') return 'AI' + col + ':AK' + col;
        if (paymentRegular === '3-й ПЛАТЕЖ') return 'AN' + col + ':AP' + col;
        if (paymentRegular === '4-й ПЛАТЕЖ') return 'AS' + col + ':AU' + col;
    }

    // запись в ячейки
    var writeData = function (record) {
        var range = detectTypePayment(record);

        try {
            sheetForm.getRange(range).setValues([[record.sum, record.day, record.type]]);
        } catch (e) {
            Browser.msgBox("Ошибка обработки заказа №" + record.order);
        }

    }

    // очистка формы
    var clearTable = function () {
        sheetForm.getRange('C4:G13').clearContent();
    }


    // Go!
    for (var i = 4; i < 14; i++) {
        var record = readRecord('A' + i + ':G' + i);

        if (!record || !record.order) continue;

        writeData(record);
    }

    clearTable();
}

Array.prototype.findIndex = function (search) {
    if (search == "") return false;
    for (var i = 0; i < this.length; i++)
        if (this[i] == search) return i;

    return -1;
}
