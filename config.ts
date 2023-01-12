const MAIN = SpreadsheetApp.getActiveSpreadsheet();
const MAIN_REQUESTS_SHEET = MAIN.getSheetByName('Запрос поручений');
const MAIN_OBJECT_LIST_SHEET = MAIN.getSheetByName('Список Объектов');

// CONFIGS.INI ========================================================================

function getCONFIGS(){

    const CONFIG_SHEET_ID = '14rknnYzXvHpC_9Utk3iyEWaZtCDDxDfMq6yBw7Sb8qY';
    const CONFIG_TELEGRAM_CHATS_DATA_SHEET_NAME = 'telegram_chats_data';
    const CONFIG_SHEETS_RELATION_SHEET_NAME = 'sheets_relation';
    const CONFIG_MANAGERS_LIST_SHEET_NAME = 'managers_list';
    const CONFIG_GIPS_LIST_SHEET_NAME = 'gip_list';
    const CONFIG_UPO_LIST_SHEET_NAME = 'upo_list';

    const CONFIG_SHEET = SpreadsheetApp.openById(CONFIG_SHEET_ID);

    return {
        CONFIG_TELEGRAM_CHATS_SHEET: CONFIG_SHEET.getSheetByName(CONFIG_TELEGRAM_CHATS_DATA_SHEET_NAME),
        CONFIG_SHEETS_RELATION_SHEET: CONFIG_SHEET.getSheetByName(CONFIG_SHEETS_RELATION_SHEET_NAME),
        CONFIG_MANAGERS_LIST_SHEET: CONFIG_SHEET.getSheetByName(CONFIG_MANAGERS_LIST_SHEET_NAME),
        CONFIG_GIPS_LIST_SHEET: CONFIG_SHEET.getSheetByName(CONFIG_GIPS_LIST_SHEET_NAME),
        CONFIG_UPO_LIST_SHEET: CONFIG_SHEET.getSheetByName(CONFIG_UPO_LIST_SHEET_NAME)
    };
};

function getRequestSheetManagerData(sheetId: string){
    let configs = getCONFIGS();
    if(!configs.CONFIG_MANAGERS_LIST_SHEET){
        throw Error('CONFIG_MANAGERS_LIST_SHEET!');
    };
    return configs.CONFIG_MANAGERS_LIST_SHEET.getRange(2, 1, configs.CONFIG_MANAGERS_LIST_SHEET.getLastRow()-1, 6).getValues()
    .filter(row => row[1] == sheetId);
};

function getApproveSheetManagerData(project: string, object: string){
    Logger.log(project)
    Logger.log(object)
    if(!MAIN_OBJECT_LIST_SHEET){
        throw Error('MAIN_OBJECT_LIST_SHEET!');
    };
    return MAIN_OBJECT_LIST_SHEET.getRange(2, 1, MAIN_OBJECT_LIST_SHEET.getLastRow()-1, MAIN_OBJECT_LIST_SHEET.getLastColumn()).getValues()
    .filter(row => row[0].replace(/ /g, '').toLowerCase() == project.replace(/ /g, '').toLowerCase() && row[1].replace(/ /g, '').toLowerCase() == object.replace(/ /g, '').toLowerCase())[0];
};

function getGIPData(gipEmail: string){
    let configs = getCONFIGS();
    if(!configs.CONFIG_GIPS_LIST_SHEET){
        throw Error('CONFIG_GIPS_LIST_SHEET!');
    };
    return configs.CONFIG_GIPS_LIST_SHEET.getRange(2, 1, configs.CONFIG_GIPS_LIST_SHEET.getLastRow()-1, 3).getValues()
    .filter(row => row[1] == gipEmail)[0];
};

function getUPOData(upoEmail: string){
    let configs = getCONFIGS();
    if(!configs.CONFIG_UPO_LIST_SHEET){
        throw Error('CONFIG_UPO_LIST_SHEET!');
    };
    return configs.CONFIG_UPO_LIST_SHEET.getRange(2, 1, configs.CONFIG_UPO_LIST_SHEET.getLastRow()-1, 3).getValues()
    .filter(row => row[1] == upoEmail)[0];
};

function getChatId(userName: string){
    let configs = getCONFIGS();
    if(!configs.CONFIG_TELEGRAM_CHATS_SHEET){
        throw Error('CONFIG_TELEGRAM_CHATS_SHEET!');
    };
    return configs.CONFIG_TELEGRAM_CHATS_SHEET.getRange(2, 1, configs.CONFIG_TELEGRAM_CHATS_SHEET.getLastRow()-1, 2).getValues()
    .filter(row => row[1].replace('@', '').toLowerCase() == userName.replace('@', '').toLowerCase())[0];
};

function getDirectionRate(directionId: string){
    let configs = getCONFIGS();
    if(!configs.CONFIG_MANAGERS_LIST_SHEET){
        throw Error('CONFIG_MANAGERS_LIST_SHEET!');
    };
    return configs.CONFIG_MANAGERS_LIST_SHEET.getRange(2, 2, configs.CONFIG_MANAGERS_LIST_SHEET.getLastRow()-1, 6).getValues()
    .filter(row => row[0] == directionId)[0][5];
};

Date.prototype.addDays = function(days: number) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

function deleteAllProtections(){

    MAIN_REQUESTS_SHEET.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  
    MAIN_REQUESTS_SHEET.getRange("1:1").protect();
    MAIN_REQUESTS_SHEET.getRange("B:S").protect();
  
  };