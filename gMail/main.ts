function sendMailMessageToManager(sheetId: string, requestCode: string){
    let managerData = getRequestSheetManagerData(sheetId);
    let userEmail = managerData[0][4];
    let messageText = gSendedForApprove(requestCode);
    gSendNotification(userEmail, messageText);
};

function sendMailMessageToCustomer(activeRow: Array<string>, sheet_url: string){
    
    let user_mail_col = 15;
    let user_mail_col2 = 15;
    let user_mail_col3 = 15;

    let requestCode = activeRow[1];
    let messageText = gSendedForCustomerApprove(requestCode, sheet_url);
    let customerData = getApproveSheetManagerData(activeRow[project_col+1], activeRow[object_col]);
    
    let user_mail = customerData[user_mail_col-1];
    let user_mail2 = customerData[user_mail_col2-1];
    let user_mail3 = customerData[user_mail_col3-1];
    
    gSendNotification(user_mail, messageText);
    

    if(user_mail2 != ''){
        gSendNotification(user_mail2, messageText);
    };

    if(user_mail2 != ''){
        gSendNotification(user_mail3, messageText);
    };
};