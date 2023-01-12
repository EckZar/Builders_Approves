function restore(){    
    if(!MAIN_REQUESTS_SHEET){
        throw Error('');
    };

    if(checkForRestore()){
        MAIN_REQUESTS_SHEET.getRange(1, 1, MAIN_REQUESTS_SHEET.getLastRow(), MAIN_REQUESTS_SHEET.getLastColumn()).getValues()
                            .map((row, i) => [i, ...row])
                            .filter(row => !row[33] && row[34])
                            .forEach(row => {

                                let directionId = getDirectionSheetData(row)[0][2];
                                
                                restoreRow(row[1], 'Отмененные поручения', directionId);
                              
                                MAIN_REQUESTS_SHEET.getRange(row[0]+1, 33).setValue(false);

                            });
    };
};

function restoreRow(rowKey: string, sheetName: string, directionId: string){

    let rowObj = findRow(rowKey, sheetName, directionId);
    let tempMain = SpreadsheetApp.openById(directionId);    

    copyToRequests('Запрос поручений', tempMain, rowObj.range);  
    deleteRow('Отмененные поручения', tempMain, rowObj.rowNum+1);

};

function copyToRequests(sheetName: string, tempMain: any, row: Array<Array<string>>){
    tempMain = tempMain.getSheetByName(sheetName);
    row[0][comment_col] = '';
    tempMain.getRange(tempMain.getLastRow()+1, 1, 1, row[0].length).setValues(row);
};

function checkForRestore() {
    if (!MAIN_REQUESTS_SHEET) {
        throw Error('');
    };
    let rows = MAIN_REQUESTS_SHEET.getRange(1, 33, MAIN_REQUESTS_SHEET.getLastRow(), 2).getValues()
        .filter(row => !row[0] && row[1]);
    if (rows.length > 0) {
        return true;
    }
    else {
        return false;
    };    
};