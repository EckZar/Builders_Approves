function resendRows(){

    let CONFIGS = getCONFIGS();

    if(!CONFIGS.CONFIG_SHEETS_RELATION_SHEET){
        throw Error('CONFIG_SHEETS_RELATION_SHEET !!!');
    };

    if(!MAIN_REQUESTS_SHEET){
        throw Error('MAIN_REQUESTS_SHEET !!!');
    };

    let rows = MAIN_REQUESTS_SHEET.getRange(3, 1, MAIN_REQUESTS_SHEET.getLastRow()-2, MAIN_REQUESTS_SHEET.getLastColumn()).getValues()
                                  .map((row, i) => [i, ...row])
                                  .filter(row => row[24] && row[29] === 'Направляется на согласование заказчику');

    CONFIGS.CONFIG_SHEETS_RELATION_SHEET.getRange(2, 1, CONFIGS.CONFIG_SHEETS_RELATION_SHEET.getLastRow()-1, 3).getValues()
    .forEach(relations => {

        Logger.log(relations);
        let direction = SpreadsheetApp.openById(relations[2]);

        if(!direction){
            throw Error(`Direction ${relations[0]} !!!`);
        };

        let directionRequestSheet = direction.getSheetByName('Запрос поручений');

        try{
          var directionRequestSheet_rows = directionRequestSheet.getRange(2, 6, directionRequestSheet.getLastRow()-1, 1).getValues()
                                         .map(row => row[0]);
        } catch(e){
          var directionRequestSheet_rows = []
        };

        let activeSheetId = MAIN.getId();
        let headKeys = ['empty_col_name', ...MAIN_REQUESTS_SHEET.getRange(2, 1, 1, MAIN_REQUESTS_SHEET.getLastColumn()).getValues()[0]];
        

        rows
        .filter(row => row[32] === relations[0])
        .forEach(row => {
            if(directionRequestSheet_rows.indexOf(row[1]) == -1){
                let directionSheetData = getDirectionSheetData(row);
                let directionId = directionSheetData[0][2];

                moveRowToDirection(headKeys, row, directionId, activeSheetId);

                Logger.log(row);
            };
        });
    });
};