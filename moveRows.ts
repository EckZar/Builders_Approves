const project_col = 1;

const isSendedToApproval_col = 23;
const isApprovedSended_col = 24;

const sender_email_col = 25;

const isApprovedAccepted_col = 26;
const isApprovedDeclined_col = 27;
const isDeletingRow_col = 33;
const isDeletedRow_col = 34;

const object_col = 31;
const directionName_col = 32;

const DIRECT_SHEET_NAME = 'Запрос поручений';

function sendRowToSheet(){
    if(checkNewRequests()){
        if(!MAIN_REQUESTS_SHEET){
            throw Error('');
        };
        let newRows = MAIN_REQUESTS_SHEET.getRange(2, 1, MAIN_REQUESTS_SHEET.getLastRow()-1, MAIN_REQUESTS_SHEET.getLastColumn()).getValues()
                                         .map((row, i) => [i, ...row])
                                         .filter(row => row[isSendedToApproval_col] && !row[isApprovedSended_col]);
       
        let activeSheetId = MAIN.getId();

        let headKeys = ['empty_col_name', ...MAIN_REQUESTS_SHEET.getRange(2, 1, 1, MAIN_REQUESTS_SHEET.getLastColumn()).getValues()[0]];
        
        newRows.forEach(row=>{
            if(row[29] == "Направляется на согласование заказчику" && row[25] != ''){

                let directionSheetData = getDirectionSheetData(row);            
                let directionId: string = directionSheetData[0][2];

                moveRowToDirection(headKeys, row, directionId, activeSheetId);
                MAIN_REQUESTS_SHEET.getRange(row[0]+2, isApprovedSended_col).setValue(true);

                protectRows(row[0]+2);
                try{
                    sendTelegramMessageToManager(activeSheetId, row[1]);
                } catch(e){};
                
                sendMailMessageToManager(activeSheetId, row[1]);

                try{
                    sendTelegramMessageToCustomer(row, directionSheetData[0][3]);
                } catch(e){};
                
                sendMailMessageToCustomer(row, directionSheetData[0][3]);

                try{
                    gipNotifications(row[sender_email_col], row[1]);
                } catch(e){};
                
            };
        });
        
    };
};

function updateFormulas(){

    if(!MAIN_REQUESTS_SHEET){
        throw Error('');
    };

    let formulas = MAIN_REQUESTS_SHEET.getRange("AC:AF").getFormulas()

    MAIN_REQUESTS_SHEET.getRange("W:Y").getValues()
                        .map((item, i) => [i, ...item])
                        .filter(item => item[1] && !item[2] && item[3])
                        .forEach(item => {
                            Logger.log(item);
                            Logger.log(formulas[item[0]]); 

                            try{
                            MAIN_REQUESTS_SHEET.getRange(`AC${item[0]+1}`).setValue(formulas[item[0]][0])
                            MAIN_REQUESTS_SHEET.getRange(`AF${item[0]+1}`).setValue(formulas[item[0]][3])
                            }catch(e){}
                        });

};

function  checkNewRequests(): boolean{

    if(!MAIN_REQUESTS_SHEET){
        throw Error('');
    };

    let check = MAIN_REQUESTS_SHEET.getRange(2, 1, MAIN_REQUESTS_SHEET.getLastRow()-1, MAIN_REQUESTS_SHEET.getLastColumn()).getValues()
                                   .filter(row => row[isSendedToApproval_col-1] && !row[isApprovedSended_col-1]);

    if(check.length>0){
        return true;
    } else {
        return false;
    };

};

function getDirectionSheetData(activeRowValues: Array<string|number>): Array<Array<string>>{
    
    let configs = getCONFIGS();

    if(!configs.CONFIG_SHEETS_RELATION_SHEET){
        throw Error('CONFIG_SHEETS_RELATION_SHEET!!!');
    };

    return configs.CONFIG_SHEETS_RELATION_SHEET.getRange(2, 1, configs.CONFIG_SHEETS_RELATION_SHEET.getLastRow()-1, configs.CONFIG_SHEETS_RELATION_SHEET.getLastColumn())
                                                .getValues()
                                                .filter(row => row[0] == activeRowValues[directionName_col]);
};

function moveRowToDirection(headKeys_from: Array<string|number|Date>, activeRowValues: Array<string|number|Date>, directionId: string, activeSheetId: string){

    let tempMain = SpreadsheetApp.openById(directionId).getSheetByName(DIRECT_SHEET_NAME);
    
    let rate = getDirectionRate(MAIN.getId());

    if(!tempMain){
        throw Error('Таблица согласования для указанной в строке дирекции не найдена!');
    };

    let headKeys_dest = tempMain.getRange(1, 1, 1, tempMain.getLastColumn()).getValues()[0];
    
    let build: Array<string|number|Date> = [];

    headKeys_dest.forEach((key, i) => {

        if(key == 'Прогнозное окончание'){

            let date = activeRowValues[headKeys_from.indexOf(key)];
            let type = typeof(date);

            if(type == 'string'){
                date = date.split('.');
                date = new Date(`${date[2]}.${date[1]}.${date[0]}`);                
            };
            
            let day = date.getDay();
                
            if(day == 1 || day == 2 || day == 0){
                date = date.addDays(3);
            };
            
            if(day == 6){
                date = date.addDays(4);
            };

            if(day >= 3 && day <= 5){
                date = date.addDays(5);
            };

            build[i] = date;

        } else if(key == 'Прогнозная стоимость работ, руб с НДС'){

            build[i] = (Number(activeRowValues[11]) + Number(activeRowValues[10])) * Number(rate);

        } else {

            build[i]= activeRowValues[headKeys_from.indexOf(key)];

        };

        
    });    

    build.map(item => {
        if(item===null){
            return '';
        };
    });

    build[2] = activeSheetId;
    build[3] = new Date();
    tempMain.getRange(tempMain.getLastRow()+1, 1, 1, build.length).setValues([build]);

};

