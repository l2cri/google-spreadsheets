function OrderForm() {
    var sheetForm = SpreadsheetApp.getActive().getSheetByName('Отчет о движ товара')
    var sheetData = SpreadsheetApp.getActive().getSheetByName('Свод данных')
    var curDate = new Date()
    var idOrder = sheetForm.getRange('C2').getValue()
    var dateOrder = sheetForm.getRange('C4').getValue()
    var lastRow = lastRowForColumn(sheetData, 4)
    var needHeadRow = false
    var colStartCategories = 6
    var columnValues = sheetData.getRange(1, colStartCategories, 1, sheetData.getLastColumn()).getValues() // для поиска интервалы

    var readRecord = function(range) {
        var data = sheetForm.getRange(range).getValues()
        if (!data || !data[0]) return []
        var res = data[0] // первый столбец
        return {
            name: res[0], // ПарфюмКод NL 1
            operation: res[3], // Поступление
            comment: res[4] || ' ', // ...
            num: res[5] // 100
        }
    }

    var findOrder = function(orderNum) {
        var searchResult = columnValues.findIndex(orderNum) // Row Index - 2

        if (searchResult !== -1) {
            return searchResult + colStartCategories // смещение относительно вставки
        }

        return false
    }

    // поиск по платежам
    var detectTypePayment = function(record) {
        var col = findOrder(record.name)

        return col
    }

    // запись в ячейки
    var writeData = function(record) {
        var col = detectTypePayment(record)

        if (!col) {
            return false
        }

        try {
            // row, column, numRows, numColumns
            sheetData.getRange(lastRow, col, 1, 3).setValues([[record.operation, record.comment, record.num]])
            return record.name
        } catch (e) {
            Browser.msgBox('Ошибка обработки заказа №' + record.name)
        }

    }

    function lastRowForColumn(sheet, column) {
        // Get all data for the given column
        var data = sheet.getRange(3, column, sheet.getLastRow(), 1)
            .getValues().filter(String).length

        return data + 1
    }

    // очистка формы
    var clearTable = function() {
        sheetForm.getRange('D7:F' + sheetForm.getLastRow()).clearContent()
    }

    // Go!
    for (var i = 7; i < sheetForm.getLastRow() + 1; i = i + 2) { // нечетные 7.9.11.13...
        var record = readRecord('A' + i + ':F' + i)
        if (!record || !record.operation || !record.name || !record.num) continue
        needHeadRow = writeData(record)
        if (needHeadRow !== false) {
            sheetForm.getRange('D' + i + ':F' + i).clearContent()
        }
    }
    if (needHeadRow) {
        sheetData.getRange('C' + lastRow + ':E' + lastRow).setValues([[curDate, idOrder, dateOrder]])
    }
    lastRow++
    needHeadRow = false
    for (var i = 8; i < sheetForm.getLastRow() + 1; i = i + 2) { // четные 8,10,12,14 ...
        var record = readRecord('A' + i + ':F' + i)
        if (!record || !record.operation || !record.name || !record.num) continue
        needHeadRow = writeData(record)
        if (needHeadRow !== false) {
            sheetForm.getRange('D' + i + ':F' + i).clearContent()
        }
    }
    if (needHeadRow) {
        sheetData.getRange('C' + lastRow + ':E' + lastRow).setValues([[curDate, idOrder, dateOrder]])
    }
    // clearTable();
}

Array.prototype.findIndex = function(search) {
    if (search == '') return false
    for (var i = 0; i < this[0].length; i++)
        if (this[0][i] == search) return i

    return -1
}