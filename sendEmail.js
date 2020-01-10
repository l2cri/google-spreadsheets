var sendObject = {
    flagSendEmail : 'Почта отправлена',
    // формат даты
    formatDate: function (date) {
        return Utilities.formatDate((date), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yy");
    },
    // условие пропуска отправки сообщения
    skipFields: function(date, user) {
        return date.getValue() === '' || user === '' || date.getNote() === this.flagSendEmail;
    },
    // Возвращает массив строки с найденным мпенеджером
    findManager: function (name) {
        var sheetForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Для рассылки уведомлений");
        var column = 3; // индекс колонки поиска  (T)
        var columnValues = sheetForm.getRange(2, column, sheetForm.getLastRow()).getValues(); //1st is header row
        var searchResult = columnValues.findIndex(name); // Row Index - 2

        if (searchResult !== -1) {
            return sheetForm.getRange(searchResult + 2, 2, 1, 6).getValues()[0];
        }

        return false;
    },
    // получать инфо о заказе
    getOrderInfo: function(sheet, row) {
        var nameManager = sheet.getRange(row, 2).getValue();
        var order = sheet.getRange(row, 3).getValue();

        if (nameManager === '') {
            throw 'Не заполненно имя менеджера';
        }

        if (order === '') {
            throw 'Не указан номер заказа';
        }

        var manager = this.findManager(nameManager);

        return {manager : manager, order : order};
    },
    // получен новый запрос на замер
    zamerEvent: function (sheet, col, row) {
        var date = sheet.getRange(row, 8); // дата замера
        var user = sheet.getRange(row, 7).getValue(); // замерщик
        var orderInfo;

        if (this.skipFields(date, user)) {
            return;
        }

        try {
            orderInfo = this.getOrderInfo(sheet, row);
        } catch (e) {
            Browser.msgBox(e);
        }

        this.sendZamer(orderInfo.manager, orderInfo.order, date, user);
    },
    // отправка письма на замер
    sendZamer: function (manager, order, date, user) {
        GmailApp.sendEmail(
            manager[5],
            "Запланирован новый замер",
            'Здравствуйте, ' + manager[0] + '!' + "\n" + "\n" +
            'По вашему заказу №' + order + ' запланирован замер на ' + this.formatDate(date.getValue()) + ' сотрудник ' + user
        );
        date.setNote(this.flagSendEmail);
    },
    // отправка письма на установку
    sendUstanovka: function (manager, order, date, user) {
        GmailApp.sendEmail(
            manager[5],
        "Запланирован новая установка",
        'Здравствуйте, ' + manager[0] + '!' + "\n" + "\n" +
        'По вашему заказу №' + order + ' запланирована установка на ' + this.formatDate(date.getValue()) + ' сотрудник ' + user
    );
        date.setNote(this.flagSendEmail);

    },
    // получен новый запрос на установку
    ustanovkaEvent: function (sheet, col, row) {
        var date = sheet.getRange(row, 34);  // дата установки
        var user = sheet.getRange(row, 32).getValue(); // сотрудник установки
        var orderInfo;

        if (this.skipFields(date, user)) {
            return;
        }

        try {
            orderInfo = this.getOrderInfo(sheet, row);
        } catch (e) {
            Browser.msgBox(e);
        }

        this.sendUstanovka(orderInfo.manager, orderInfo.order, date, user);
    }
};

// add this function on trigger edit sheet
function onEditTriggered(e){
    var row = e.range.getRow();
    var col = e.range.getColumn();
    var sheet = e.source.getActiveSheet();

    if(sheet.getName() === 'Манипуляция') {
        // уведоление о замере
        if(col === 7 || col === 8) {
            sendObject.zamerEvent(sheet, col, row);
        }
        // уведоление о установки
        if(col === 32 || col === 34) {
            sendObject.ustanovkaEvent(sheet, col, row);
        }
    }
}


Array.prototype.findIndex = function (search) {
    if (search == "") return false;
    for (var i = 0; i < this.length; i++)
        if (this[i] == search) return i;

    return -1;
};
