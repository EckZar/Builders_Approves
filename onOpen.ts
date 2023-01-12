function onOpen(e: any){

    SpreadsheetApp.getUi().createMenu("__MENU__")
    .addItem("Просто согласование", "send_Two")
    .addItem("согласование 400го кода (с виновным ГСом)", "soglas_400_kod_s_iniciatorom")
    .addItem("согласование 400го кода (без виновного ГСа)", "soglas_400_kod_bez_iniciatora")
    .addItem("согласование 100го кода (с типовым описанием)", "soglas_100_kod")
    .addItem("согласование 111го кода (СБЦ/МРР)", "soglas_111_kod")
    .addItem("оповещение о выдачи задачи (с ID в Планере)", "send_4")
    .addToUi()
  
  };