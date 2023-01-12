function sendTelegramMessageToManager(sheetId: string, requestCode: string){

    let managerData = getRequestSheetManagerData(sheetId);
    let chat_id = getRequestSheetManagerTelegramChactId(managerData[0][5])
    let messageText = tSendedForApprove(requestCode);
    sendNotification(chat_id, messageText);
};

function getRequestSheetManagerTelegramChactId(userName: string){
    let configs = getCONFIGS();
    if(!configs.CONFIG_TELEGRAM_CHATS_SHEET){
        throw Error('CONFIG_MANAGERS_LIST_SHEET!');
    };
    return configs.CONFIG_TELEGRAM_CHATS_SHEET.getRange(2, 1, configs.CONFIG_TELEGRAM_CHATS_SHEET.getLastRow()-1, 5).getValues()
    .filter(row => row[1].replace(/@/g, '').toLowerCase() == userName.replace(/@/g, '').toLowerCase())[0][0];
};


function sendTelegramMessageToCustomer(activeRow: Array<string>, sheet_url: string){

    let user_name_col = 16;
    let user_name_col2 = 19;
    let user_name_col3 = 22;

    let requestCode = activeRow[1];

    let customerData = getApproveSheetManagerData(activeRow[project_col+1], activeRow[object_col]);
    let messageText = tSendedForCustomerApprove(requestCode, sheet_url);

    let telegram_user_name = customerData[user_name_col-1];
    let telegram_user_name2 = customerData[user_name_col2-1];
    let telegram_user_name3 = customerData[user_name_col3-1];

    let chat_id = getRequestSheetManagerTelegramChactId(telegram_user_name);
    
    sendNotification(chat_id, messageText);     

    if(telegram_user_name2!=''){
        let chat_id2 = getRequestSheetManagerTelegramChactId(telegram_user_name2);
        sendNotification(chat_id2, messageText);
    
    };

    if(telegram_user_name3!=''){
        let chat_id3 = getRequestSheetManagerTelegramChactId(telegram_user_name3);
        sendNotification(chat_id3, messageText);
    };

};