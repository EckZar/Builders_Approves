function tSendedForApprove(requestCode: string){
    return `Строка с кодом ${requestCode} отправлена на согласование.`;
};

function tSendedForCustomerApprove(requestCode: string, sheet_url: string){
    return `Пришел запрос на согласование с кодом ${requestCode}. Ссылка на таблицу согласования - ${sheet_url}`;
};

function tSendedForGip(requestCode: string){
    return `Строка с кодом ${requestCode} успешно отправлена на согласование.`;
};