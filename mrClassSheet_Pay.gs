function menuMakeCopy_оплаты() {
  // let classSheet_Buy = new ClassSheet_Buy();
  let classSheet_Pay = getContext().getSheetPay();
  classSheet_Pay.menuMakeCopy();
  Logger.log(`menuMakeCopy UrlExternalSpreadSheet=${classSheet_Pay.getUrlExternalSpreadSheet()}`);
  addNewTrigger();
}




class ClassSheet_Pay {
  constructor() { // class constructor
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Оплаты);

    // this.folderId = getSettings().folderId;
    this.folderId = getSettings().folderId_оплаты;
    this.url_patternSpreadSheet = getSettings().url_таблица_оплаты_шаблон;

    this.rowHads = 1;

    this.rowBodyFirst = 2;

    // this.rowBodyLast = this.findRowBodyLast();
    this.sheetName_Исполнение_Внешние = getContext().settings.sheetName_Исполнение_Внешние;

    this.makeCol();

  }


  getMapPercenOfPayment() {

    if (!this.mapPercenOfPayment) {

      this.mapPercenOfPayment = new Map();
      let url = getContext().getSpreadSheet().getUrl();
      let sheetName = this.sheet.getSheetName();
      let rowConf = {
        head: { first: 2, last: 2, key: 2, },
        body: { first: 3, last: 3, },
      }
      let sheetModel = MrLib_Midlle.makeSheetModelBy(url, sheetName, rowConf);

      sheetModel.getMap().forEach(v => {
        if (!v["Номенклатура"]) { return; }
        this.mapPercenOfPayment.set(`${v["№"]} | ${v["Номенклатура"]}`, v["% оплаты"])
      });



    }
    return this.mapPercenOfPayment;
  }



  makeFormulaUrlFolderForExterlanSheet() {
    // let row = getContext().getSheetList7().rowBodyLast;
    let url = SpreadsheetApp.getActive().getUrl();

    // let col = getContext().getSheetList7().columnContragentEmail;
    // return `IMPORTRANGE(
    // "${url}"
    // ;
    // "'${getSettings().sheetName_Лист_7}'!A1:${nc(col)}${row}"
    // )
    // `;

    return `=VLOOKUP("Счета сохранять в"; IMPORTRANGE("${url}";"'5 Закупка товара'!B:C");2;FALSE)`


  }



  getПлательщики() {
    let ret = new Array();

    if (!this.getExternalSpreadSheet()) {
      return ret;
    }
    let sheet_Оплаты_Внешняя = this.getExternalSpreadSheet().getSheetByName(getContext().settings.sheetName_Оплаты_Внешняя);
    if (!sheet_Оплаты_Внешняя) {
      return ret;
    }

    let rowBodyFirst = 2;
    let rowBodyLast = sheet_Оплаты_Внешняя.getLastRow();
    let colПлательщик = sheet_Оплаты_Внешняя.getRange(1, 1, 1, sheet_Оплаты_Внешняя.getLastColumn()).getValues()[0].indexOf("Плательщик") + 1;
    if (!colПлательщик) { return ret; }
    let retSet = new Set(
      sheet_Оплаты_Внешняя.getRange(rowBodyFirst, colПлательщик, rowBodyLast - rowBodyFirst + 1, 1).getValues().filter(v => `${v[0]}`).map(v => v[0])
    );

    ret = [...retSet.values()];
    return ret;
  }

  getПосредник() {
    let ret = undefined;

    try {

      let urlExternalSpreadSheet = this.getUrlExternalSpreadSheet();
      // Logger.log(` getContextТаблицаОплатВнешняя urlExternalSpreadSheet =${urlExternalSpreadSheet}`);
      if (urlExternalSpreadSheet) {
        let MrContextConstructor = MrLib_Midlle.getMrContextConstructor();
        this.contextТаблицаОплатВнешняя = new MrContextConstructor(urlExternalSpreadSheet, MrLib_Midlle.getSettings());
        let MrClassSheetModel = MrLib_Midlle.getMrClassSheetModel();
        let sheetModelСводнаяПрочее = new MrClassSheetModel("Сводная Прочее", this.contextТаблицаОплатВнешняя);
        ret = sheetModelСводнаяПрочее.getMap().get("Посредник")["Посредник"];
      }
    } catch (err) {
      // ret = mrErrToString(err); 
    }
    return ret;
  }

  getСводнаяПрочее(ret = new Object()) {
    // let ret = new Object();
    // return "DSDSD";


    try {
      let urlExternalSpreadSheet = this.getUrlExternalSpreadSheet();
      // Logger.log(` getContextТаблицаОплатВнешняя urlExternalSpreadSheet =${urlExternalSpreadSheet}`);
      if (urlExternalSpreadSheet) {
        let MrContextConstructor = MrLib_Midlle.getMrContextConstructor();
        this.contextТаблицаОплатВнешняя = new MrContextConstructor(urlExternalSpreadSheet, MrLib_Midlle.getSettings());
        let MrClassSheetModel = MrLib_Midlle.getMrClassSheetModel();
        let sheetModelСводнаяПрочее = new MrClassSheetModel("Сводная Прочее", this.contextТаблицаОплатВнешняя);
        // ret =[...sheetModelСводнаяПрочее.getMap().values()];

        /** @type {Map} */
        let mapСП = sheetModelСводнаяПрочее.getMap();
        mapСП.forEach((v, key) => {
          for (let k in v) {
            if (`${v[k]}` != "") { continue; }
            v[k] = undefined;
          }
          ret[key] = v;
        });

      }
    } catch (err) {
      ret = mrErrToString(err);
    }
    return ret;
  }



  makeFormulaForExterlanSheet() {
    let row = getContext().getSheetList7().rowBodyLast;
    let url = SpreadsheetApp.getActive().getUrl();

    // let col = getContext().getSheetList7().columnContragentEmail;
    let col = getContext().getSheetList7().columnStaus;
    // columnStaus
    return `IMPORTRANGE(
    "${url}"
    ;
    "'${getSettings().sheetName_Лист_7}'!A1:${nc(col)}${row}"
    )
    `;
  }

  makeFormulaForThisSheet(lastColumn = "Z") {
    let row = getContext().getSheetList7().rowBodyLast;

    // return `IMPORTRANGE(
    // $${nc(this.col_C)}$${this.getRowUrlExternalSpreadSheet()}
    // ;
    // "'${getSettings().sheetName_Оплаты_Внешняя}'!A1:${row + 2}"
    // )
    // `;

    return `IMPORTRANGE(
    $${nc(this.col_C)}$${this.getRowUrlExternalSpreadSheet()}
    ;
    "'${getSettings().sheetName_Оплаты_Внешняя}'!A:${lastColumn}"
    )
    `;


  }




  findRowBodyLast() {
    let ret = getContext().getRowSobachkaBySheetName(this.sheet.getSheetName());
    if (!ret) { ret = this.sheet.getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }





  makeCol() {
    this.col_A = nr("A");
    this.col_B = nr("B");
    this.col_C = nr("C");
    this.col_D = nr("D");
  }



  menuMakeCopy() {
    this.makePatternCopy();


    // this.sheet.getRange(this.rowBodyFirst, this.col_A).setFormula(this.makeFormulaForThisSheet());

    try {
      let ssa = this.getExternalSpreadSheet();
      if (!ssa) { return; }

      let sheet_Оплаты_Внешняя = ssa.getSheetByName(getSettings().sheetName_Оплаты_Внешняя);
      sheet_Оплаты_Внешняя.getRange("A1").setFormula(this.makeFormulaForExterlanSheet());

      // let lastColumn = sheet_Оплаты_Внешняя.getLastColumn();
      let lastColumn = sheet_Оплаты_Внешняя.getMaxColumns();
      this.sheet.getRange(this.rowBodyFirst, this.col_A).setFormula(this.makeFormulaForThisSheet(nc(lastColumn)));


      let col_key = "ссылка на папку со счетами";
      let col_url = sheet_Оплаты_Внешняя.getRange("A1:1").getValues()[0].map((v) => { return fl_str(v) }).indexOf(fl_str(col_key)) + 1;
      // Logger.log("|||||||||||||||||||||||||||" + col_url);

      if (col_url) {
        col_url = col_url + 1;
      } else { return; }
      sheet_Оплаты_Внешняя.getRange(1, col_url).setFormula(this.makeFormulaUrlFolderForExterlanSheet());

      // let row_key = fl_str("Счета сохранять в");
      // let urls = getContext().getSheetBuy().sheet.getRange("B1:C").getValues().filter(v => {
      //   if (fl_str(v[0]) == row_key) {
      //     return true;
      //   }
      //   return false;
      // }).map(v => v[1]);

      // // Logger.log("|||||||||||||||||||||||||||" + urls);
      // if (urls.length) {
      //   sheet_Оплаты_Внешняя.getRange(1, col_url).setValue(urls[0]);
      // }


    } catch (err) {
      mrErrToString(err);
    }
  }


  makePatternCopy() {
    // let formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd' 'HH:mm:ss");
    let formattedDate = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd' 'HH:mm:ss");
    let name = SpreadsheetApp.getActiveSpreadsheet().getName() + " | ТАБЛИЦА ОПЛАТ ВНЕШНЯЯ | " + formattedDate;
    let destination = DriveApp.getFolderById(this.folderId);
    let file = DriveApp.getFileById(SpreadsheetApp.openByUrl(this.url_patternSpreadSheet).getId());
    let patternCopy = file.makeCopy(name, destination);
    patternCopy.addEditor("script@ss-postavka.ru");
    let url_patternCopy = patternCopy.getUrl();
    Logger.log(`url_patternCopy = ${url_patternCopy}`);
    this.setUrlExternalSpreadSheet(url_patternCopy);

    let msg = `ТАБЛИЦА ОПЛАТ ВНЕШНЯЯ созданна. Название: ${name}. Ссылка: ${url_patternCopy}. `;
    Browser.msgBox(msg);

  }




  getExternalSpreadSheet() {
    if (!this.externalSpreadSheet) {
      let urlExternalSpreadSheet = this.getUrlExternalSpreadSheet();
      if (!urlExternalSpreadSheet) { this.externalSpreadSheet = undefined; }
      try {
        this.externalSpreadSheet = SpreadsheetApp.openByUrl(urlExternalSpreadSheet);
      } catch (err) {
        Logger.log(mrErrToString(err));
        this.externalSpreadSheet = undefined;
      }
    }
    return this.externalSpreadSheet;
  }


  getUrlExternalSpreadSheet() {
    let row = this.getRowUrlExternalSpreadSheet()
    let col = this.col_C;
    let url = this.sheet.getRange(row, col).getValue();
    return url;
  }

  getRowUrlExternalSpreadSheet() {
    return 1;
  }

  setUrlExternalSpreadSheet(url) {
    let row = this.getRowUrlExternalSpreadSheet()
    let col = this.col_C;
    this.sheet.getRange(row, col).setValue(url);
  }


  triggerOnEditHelperPayExternal(duration) {
    Logger.log(`ClassSheet_Pay triggerOnEditHelperPayExternal duration=${duration}`);

    let externalSpreadSheet = this.getExternalSpreadSheet();
    if (externalSpreadSheet != undefined) {
      try {
        MrLib_External.triggerMrClassSheetКонтрольИсполненияВнешний(externalSpreadSheet, getContext().timeConstruct, duration);
      } catch (err) {
        Logger.log(mrErrToString(err));
      }

    } else {
      Logger.log(`ClassSheet_Pay triggerOnEditHelperPayExternal ClassSheet_PayExternal Внешней таблици нет `);
    }



  }

  тампусто(duration) {

    let externalSpreadSheet = this.getExternalSpreadSheet();
    if (externalSpreadSheet != undefined) {
      try {
        return MrLib_External.тампусто(externalSpreadSheet, getContext().timeConstruct, duration);
      } catch (err) {
        Logger.log(mrErrToString(err));
        return "Ошибка\n" + mrErrToString(err);
      }

    } else {

      Logger.log(`ClassSheet_Pay triggerOnEditHelperPayExternal ClassSheet_PayExternal Внешней таблици нет `);
      return "Нет" + JSON.stringify(getContext().getUrls());
    }

    return "Нет должно быть";
  }



}
