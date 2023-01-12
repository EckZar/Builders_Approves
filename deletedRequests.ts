const deletionComment_col = 35;
const comment_col = 21;

function remove(){    
    if(!MAIN_REQUESTS_SHEET){
        throw Error('');
    };

    if(checkRows()){
        MAIN_REQUESTS_SHEET.getRange(1, 1, MAIN_REQUESTS_SHEET.getLastRow(), MAIN_REQUESTS_SHEET.getLastColumn()).getValues()
                            .map((row, i) => [i, ...row])
                            .filter(row => row[33] && !row[34])
                            .forEach(row => {

                                let directionId = getDirectionSheetData(row)[0][2];
                                
                                let reason = row[deletionComment_col];

                                moveRowToRemoved(row[1], 'Запрос поручений', directionId, reason);
                              
                                MAIN_REQUESTS_SHEET.getRange(row[0]+1, 34).setValue(true);

                            });
    };
};

function moveRowToRemoved(rowKey: string, sheetName: string, directionId: string, reason: string){
    
    let rowObj = findRow(rowKey, sheetName, directionId);
    let tempMain = SpreadsheetApp.openById(directionId);    
    copyToRemoved('Отмененные поручения', tempMain, rowObj.range, reason);  
    deleteRow('Запрос поручений', tempMain, rowObj.rowNum+1);
      

};

function checkRows(){
    if(!MAIN_REQUESTS_SHEET){
        throw Error('');
    };
    let rows = MAIN_REQUESTS_SHEET.getRange(1, 32, MAIN_REQUESTS_SHEET.getLastRow(), 2).getValues()
                                  .filter(row => row[0] && !row[1]);
    if(rows.length>0){
        return true;
    } else {
        return false;
    };
};

function findRow(rowKey: string, sheetName: string, directionId: string){

    let tempMain = SpreadsheetApp.openById(directionId);
    let searchSheet = tempMain.getSheetByName(sheetName);

    if(!searchSheet){
        throw Error();
    };

    let row = searchSheet.getRange(1, 1, searchSheet.getLastRow(), searchSheet.getLastColumn()).getValues()
    .map((row, i) => [i, ...row])
    .filter(row => row[6] == rowKey)[0];

    return {
        rowNum: row[0],
        range: searchSheet.getRange(row[0]+1, 1, 1, searchSheet.getLastColumn()).getValues()
    };
};

function deleteRow(sheetName: string, tempMain: any, rowNum: number){
    tempMain.getSheetByName(sheetName).deleteRow(rowNum);
};

function copyToRemoved(sheetName: string, tempMain: any, row: Array<Array<string>>, reason: string){
    tempMain = tempMain.getSheetByName(sheetName);
    row[0][comment_col-1] = reason;
    tempMain.getRange(tempMain.getLastRow()+1, 1, 1, row[0].length).setValues(row);
};

