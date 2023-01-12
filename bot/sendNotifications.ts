function sendNotification(manager_chat_id:number, messageText: string){    
    let url = `https://api.telegram.org/bot${TOKEN}/sendMessage?chat_id=${manager_chat_id}&text=${messageText}`;
    UrlFetchApp.fetch(url);
};