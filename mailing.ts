function sendEmails() {

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("test")

  if(!sheet){
    throw Error('');
  };

  try{
    
    var arrData = sheet.getRange("A3:A" + sheet.getLastRow()).getDisplayValues()
    var arrEmails = []
    
    for(var i = 0; i< arrData.length; i++){
      var email = arrData[i][0]
      if(email.length > 1 && email.search("@") > 0){
        arrEmails.push(email)
      }
    }
    
    var emailCount = arrEmails.length
    if(emailCount < 1){
      Browser.msgBox("Внимание", "Нет валидных Emails в списке", Browser.Buttons.OK)
      return
    }
    
    var subject = 'Согласование на выставление поручения'+ "_" + new Date();
    var tableName = sheet.getName()
    var dopText1 = sheet.getRange("A29").getDisplayValue()
    var dopText2 = sheet.getRange("A30").getDisplayValue()
    var dopText3 = sheet.getRange("A31").getDisplayValue()
    var dopText4 = sheet.getRange("A32").getDisplayValue()
    var dopText5 = sheet.getRange("A33").getDisplayValue()
    var dopText6 = sheet.getRange("A34").getDisplayValue()
    var dopText7 = sheet.getRange("A35").getDisplayValue()
    var dopText0 = sheet.getRange("A28").getDisplayValue()
    var isNotOneMail = Boolean(sheet.getRange("C3").getDisplayValue() == "ДА")
    
    var sendEmailCount = 0
    if(isNotOneMail){
      sendEmailCount = sendEmailsSeparate_(emailCount, arrEmails, subject, tableName, dopText0, dopText1, dopText2, dopText3, dopText4, dopText5, dopText6, dopText7)
    } else {
      sendEmailCount = sendEmailsInOne_(emailCount, arrEmails, subject, tableName, dopText0, dopText1, dopText2, dopText3, dopText4, dopText5, dopText6, dopText7)
    }
    
    var remainingEmails = MailApp.getRemainingDailyQuota()
    Browser.msgBox("Внимание!", "Отправлено " + sendEmailCount + 
                   " Emails\\nОстаток квоты на сегодня: " + remainingEmails, Browser.Buttons.OK)
  } catch(error){
    Browser.msgBox(error)
  }
};

function sendEmailsSeparate_(emailCount, arrEmails, subject, tableName, dopText0, dopText1, dopText2, dopText3, dopText4, dopText5, dopText6, dopText7){
  var sendEmailCount = 0
  for(var i = 0; i < emailCount; i++){
    MailApp.sendEmail({
      to: arrEmails[i],
      subject: subject,
      htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +
      " - Проект: " + dopText0 + "<br>" +
      " - Задача: " + dopText1 + "<br>" +
      " - Поручение: " + dopText2 + "<br>" +
      " - Вид операции: " + dopText3 + "<br>" +
      " - Описание причины: " + dopText4 + "<br>" +
      " - Диапазон дат: " + dopText5 
    })
    sendEmailCount++
  }
  return sendEmailCount
};

function sendEmailsInOne_(emailCount, arrEmails, subject, tableName, dopText0, dopText1, dopText2, dopText3, dopText4, dopText5, dopText6, dopText7){
  if(emailCount == 1) {
    MailApp.sendEmail({
      to: arrEmails[0],
      subject: subject,
      htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +
      " - Проект: " + dopText0 + "<br>" +
      " - Задача: " + dopText1 + "<br>" +
      " - Поручение: " + dopText2 + "<br>" +
      " - Вид операции: " + dopText3 + "<br>" +
      " - Описание причины: " + dopText4 + "<br>" +
      " - Диапазон дат: " + dopText5 
    })
  } else {
    var to = arrEmails.shift()
    var cc = arrEmails.join(",")
    MailApp.sendEmail({
      to: to,
      cc: cc,
      subject: subject,
      htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +
      " - Проект: " + dopText0 + "<br>" +
      " - Задача: " + dopText1 + "<br>" +
      " - Поручение: " + dopText2 + "<br>" +
      " - Вид операции: " + dopText3 + "<br>" +
      " - Описание причины: " + dopText4 + "<br>" +
      " - Диапазон дат: " + dopText5 
    })
  }  
  return 1
};




function send_Two(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения'+ "_№ " + range[0][0];
  
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: range[0][16],
    cc: range[0][17] + " , " + range[0][18] + " , " + range[0][19] + " , " + range[0][21],
    subject: subject,
    htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][7] + " - " + range[0][8] + "<br>"  +
    " - Плановые ТРЗ: " + range[0][9]
  })
  
  
};




function send_notification(){
      
  var subject = 'Новая заявка на поручение в 1С';
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("Запрос поручений")
  var mail = sheet.getRange("x1").getDisplayValue()
  
 
  MailApp.sendEmail({
    to: mail,
    subject: subject,
    htmlBody: "Добрый день, коллеги" + "<br><br>" +
    " В таблице по поручениям 1С добавлена новая заявка, прошу обработать  " 
  })
  
  
};

function send_3(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения (у ГИПа)'+ "_№ " + range[0][0];
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("Запрос поручений")
  var mail = range[0][13]
  var mail2 = range[0][14] + " , " + range[0][15] + " , " + range[0][16]
  var mail3 = range[0][16]
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: mail,
    cc: mail2,
    subject: subject,
    htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][6] + " - " + range[0][7] 
    })
  
  
};

function send_4(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения'+ "_№ " + range[0][0];
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("Запрос поручений")
  var mail = range[0][18]
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: mail,
    subject: subject,
    htmlBody: "Здравствуйте!" + "<br><br>" +
    " Вам выдана задача: ''" + range[0][2] + "''  " + "в проекте ''" + range[0][1] + "''" + "<br>"  +
    "ID задачи в ПИК Планере: " + range[0][15]
    })
  
  
};



function s_koment_tech_derikc(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'КОМ_'+ range[0][20] + "_" + range[0][19] + "_" + range[0][1] + " " + 'Согласование на выставление поручения'+ " № " + range[0][0];
  
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: range[0][12],
    cc: cc,
    subject: subject,
    htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][6] + " - " + range[0][7] + "<br>"  +
    " - Плановые ТРЗ: " + range[0][8]
  })
  
  
};




function soglas_400_kod_s_iniciatorom(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения'+ "_№ " + range[0][0];
  
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: range[0][16],
    cc: range[0][17] + " , " + range[0][18] + " , " + range[0][19] + " , " + range[0][21],
    subject: subject,
    htmlBody: "Здравствуйте! Прошу согласовать выставление поручения по внутреннему изму" + "<br>" +
    range[0][12] + ", прошу вас также согласовать, как ГС-инициатор по изму:" + "<br><br>" +
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][7] + " - " + range[0][8] + "<br>"  +
    " - Плановые ТРЗ: " + range[0][9] + "<br>"  +
    " - ГС - инициатор ИЗМа: " + range[0][12]
  })
  
  
};





function soglas_100_kod(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения'+ "_№ " + range[0][0];
  
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: range[0][16],
    cc: range[0][17] + " , " + range[0][18] + " , " + range[0][19] + " , " + range[0][20] + " , " + range[0][21],
    subject: subject,
    htmlBody: "Здравствуйте! Прошу согласовать выставление поручения:" + "<br><br>" +    
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Типовое описание причины: " + range[0][6] + "<br>" +
    " - Детальное описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][7] + " - " + range[0][8] + "<br>"  +
    " - Плановые ТРЗ: " + range[0][9] + "<br>"  +
    "   - - Из них в выходные: " + range[0][10]
  })
  
  
};




function soglas_400_kod_bez_iniciatora(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения'+ "_№ " + range[0][0];
  
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: range[0][16],
    cc: range[0][17] + " , " + range[0][18] + " , " + range[0][19] + " , " + range[0][21],
    subject: subject,
    htmlBody: "Здравствуйте! Прошу согласовать выставление поручения по внутреннему изму" + "<br><br>" +
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][7] + " - " + range[0][8] + "<br>"  +
    " - Плановые ТРЗ: " + range[0][9] + "<br>"  +
    " - ГС - инициатор ИЗМа = ГС - Исполнитель" 
  })
  
  
};


function soglas_111_kod(){
  
  
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Запрос поручений')
  var range = data.getActiveRange().getValues()
  var subject = 'Согласование на выставление поручения'+ "_№ " + range[0][0] + " (код 111)";
  
  
  var arrData = range[0]
  var arrEmails = []
  
  for(var i = 0 ; i< arrData.length; i++){
    var email = arrData[i]
    if(email.length > 1 && email.search("@") > 0){
      arrEmails.push(email)
    }
  }

 
  var cc = arrEmails.join(",")
  MailApp.sendEmail({
    to: range[0][16],
    cc: range[0][17] + " , " + range[0][18] + " , " + range[0][19] + " , " + range[0][20] + " , " + range[0][21],
    subject: subject,
    htmlBody: "Здравствуйте! Пришёл запрос на выставление задачи по СБЦ (МРР):"  + "<br>" +
    "Прошу согласовать задачу:"  + "<br>" + 
    "ГИПа;"  + "<br>" +
    "Лидера (РП) после согласования с заказчиком (приложить скрин)." + "<br><br>" +    
    " - Проект: " + range[0][1] + "<br>" +
    " - Задача: " + range[0][2] + "<br>" +
    " - Поручение: " + range[0][3] + "<br>" +
    " - Вид операции: " + range[0][4] + "<br>" +
    " - Типовое описание причины: " + range[0][6] + "<br>" +
    " - Детальное описание причины: " + range[0][5] + "<br>" +
    " - Плановые даты: " + range[0][7] + " - " + range[0][8] + "<br>"  +
    " - Плановые ТРЗ: " + range[0][9] + "<br>"  +
    "   - - Из них в выходные: " + range[0][10]
  })
  
  
};
