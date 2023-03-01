function protectRows(rowNum: number){
    if(!MAIN_REQUESTS_SHEET){
        throw Error('');
    };
    let protection = MAIN_REQUESTS_SHEET.getRange(rowNum, 1, 1, 32).protect();

    try{
        protection.removeTargetAudience(protection.getTargetAudiences());
    } catch(e){};
        
    protection.removeEditors(protection.getEditors());
};