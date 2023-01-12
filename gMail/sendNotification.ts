function gSendNotification(managerEmail: string, messageText: string){
    GmailApp.sendEmail(managerEmail, 'Заявка на согласование', messageText);
};