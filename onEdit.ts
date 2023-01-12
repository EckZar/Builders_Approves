function onEdit(e: any){
    if(e.range.getA1Notation().indexOf('W')>=0 && e.value == 'TRUE'){    
        showDialogBox(e);
    };
};

function showDialogBox(e: any){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Подтверждаем правильность заполнения строки и отправку уведомления?', ui.ButtonSet.YES_NO);
    let cell = e.range.getA1Notation();

    if (response == ui.Button.YES) {       
        let active = SpreadsheetApp.getActiveSheet();
        let editor_email = e.user;
        active.getRange(cell.replace('W', 'Y')).setValue(editor_email);
    };

    if (response == ui.Button.NO) {       
        SpreadsheetApp.getActiveSheet().getRange(cell).setValue(false);
    };
  };