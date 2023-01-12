function gSendedForApprove(requestCode: string){
    return `Строка с кодом ${requestCode} отправлена на согласование. `;
};

function gSendedForCustomerApprove(requestCode: string, sheet_url: string){
    return `Пришел запрос на согласование с кодом ${requestCode}. Ссылка на таблицу согласования – ${sheet_url}`;
};

function gSendedForGip(requestCode: string){
    return `Строка с кодом ${requestCode} успешно отправлена на согласование.`;
};