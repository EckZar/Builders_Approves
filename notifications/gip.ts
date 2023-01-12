function gipNotifications(gipEmail: string, requestCode: string){

    let gipData = getGIPData(gipEmail);
    let gipChatData = getChatId(gipData[2]);

    let messageText = tSendedForGip(requestCode);

    sendNotification(gipChatData[0], messageText);
    gSendNotification(gipData[1], messageText);
};