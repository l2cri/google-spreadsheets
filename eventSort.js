function onEdit(e){

    // ______ первый скрипт по АвтоДате
    var row = e.range.getRow();
    var col = e.range.getColumn();
    var ws = e.source.getActiveSheet().getName();
    var curDate = new Date()
    // Col>16 && Col<18 это определяется 17 колонка где будут менятся значения
    if(row > 1 && col > 16 && col < 18 && ws === "Заказы (нов.)"){
        e.source.getActiveSheet().getRange(row,19).setValue(curDate)
    }
    if(e.source.getActiveSheet().getRange(row,18).getValue() == "" && ws === "Заказы (нов.)"){
        e.source.getActiveSheet().getRange(row,18).setValue(curDate)}
    //19 это колонка куда записывается дата и время изменения первой записи даты. 18 это в случае изменения ее

    // ________ начинается второй скрипт по автофильтру.
    var range = e.range;
    // 12 это десятая колонка M
    if (range.getColumn() === 12) {
        updateSort();
    }
}

function updateSort() {
    var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Заказы (нов.)');
    var criteria = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues([''])
        .build();
    // номер колонки для фильтра (3)
    spreadsheet.getFilter().setColumnFilterCriteria(3, criteria);

    var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Входящие кл.');
    var criteria = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(['Договор'])
        .build();
    // номер колонки для фильтра (12)
    spreadsheet.getFilter().setColumnFilterCriteria(12, criteria);
}
