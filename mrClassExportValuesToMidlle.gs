
class MrClassExportProject {
  /** @param {MrContext} context*/
  constructor(context, item) { // class constructor
    if (!context) { throw new Error("context"); }
    /** @private  */
    this.context = context;
    /** @private  */
    this.item = item;
    /** @private  */
    this.sobachkaMap = undefined;
    this.sheetName_Проекты = "Проекты";
    this.sheetName_Сводная = "Сводная";
  }


  getContextТаблицаОплатВнешняя() {
    if (!this.contextТаблицаОплатВнешняя) {
      let urlExternalSpreadSheet = this.context.getSheetPay().getUrlExternalSpreadSheet();
      // Logger.log(` getContextТаблицаОплатВнешняя urlExternalSpreadSheet =${urlExternalSpreadSheet}`);
      if (urlExternalSpreadSheet) {
        // Logger.log(` getContextТаблицаОплатВнешняя urlExternalSpreadSheet  =${urlExternalSpreadSheet}`);
        let MrContextConstructor = MrLib_Midlle.getMrContextConstructor();
        // this.contextТаблицаОплатВнешняя = MrLib_Midlle.getMrContextConstructor()(urlExternalSpreadSheet, MrLib_Midlle.getSettings());
        this.contextТаблицаОплатВнешняя = new MrContextConstructor(urlExternalSpreadSheet, MrLib_Midlle.getSettings());
        // Logger.log(`  this.contextТаблицаОплатВнешняя =${this.contextТаблицаОплатВнешняя}`);
      }
    }
    // Logger.log(`  this.contextТаблицаОплатВнешняя =${this.contextТаблицаОплатВнешняя}`);
    return this.contextТаблицаОплатВнешняя;
  }

  getProjectNumber(url) {
    let retProjectNumber = undefined;
    let sheetModelСводная = MrLib_Midlle.getSheetModelBySheetName(this.sheetName_Сводная);
    if (sheetModelСводная) {
      let mapСводная = sheetModelСводная.getMap();
      mapСводная.forEach(v => {
        if (url != v.Таблица) { return; }
        if (retProjectNumber) { return; }
        retProjectNumber = v.key;
      });

    }
    return retProjectNumber;
  }


  triggerExportValuesToMidlle() {

    // return
    if (getContext().timeConstruct.getMinutes() % 30 > 5) { return; }
    Logger.log(`triggerExportValuesToMidlle`);
    try {
      let exportValues = this.getExportValues();

      let projectNumber = this.getProjectNumber(exportValues.key);
      // exportValues.key = projectNumber;
      exportValues.Тб_date = getContext().timeConstruct;

      let ФайлПроекта = (() => {
        try { return MrLib_Midlle.getItemByKey(MrLib_Midlle.getSettings().sheetNames.Файлы, `${projectNumber}`, "ФайлПроекта", undefined); }
        catch (err) { mrErrToString(err); return undefined; }
      })();

      if (ФайлПроекта) {
        let obj = MrLib_Midlle.getNewFileObjData();
        obj.value = exportValues;
        obj.ТаблицаПроекта = exportValues.key;
        obj.НомерПроекта = `${projectNumber}`;
        obj.ФайлПроекта = ФайлПроекта;
        obj.key = `${projectNumber}`;
        obj.info = "update " + obj.key;
        MrLib_Midlle.getMrClassFile().update(obj);
      } else {
        Logger.log(`triggerExportValuesToMidlle Нет Файла `);
      }

      // if (JSON.stringify(exportValues.Тб).length < 50000) {
      MrLib_Midlle.updateItems(this.sheetName_Проекты, [exportValues]);
      exportValues.key = projectNumber;
      MrLib_Midlle.updateItems(this.sheetName_Сводная, [exportValues]);
      // }

    } catch (err) { mrErrToString(err); }
  }


  /** @private  */
  findRowBodyLast(sheetName) {
    // Logger.log(JSON.stringify(sheetName));

    let ret = getContext().getRowSobachkaBySheetName(sheetName)
    // let ret = this.getRowSobachkaBySheetName(sheetName);
    if (!ret) { ret = this.context.getSheetByName(sheetName).getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }


  /** @private  */
  getValuesFromВыбор_Поставщиков(item) {

    let mapСводная = new Map();
    this.getValuesFromСводнаяПоНоменклатуре().forEach(item => {
      let key = item["№"];
      if (!key) { return; }
      mapСводная.set(key, item);
    });


    let rowBodyFirst = 2;
    let retArr = new Array();
    let sheet = this.context.getSheetByName(this.context.settings.sheetName_Лист_7);
    let heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    heads = ["row"].concat(heads).map(h => `${h}`.trim());
    let vls = sheet.getRange(rowBodyFirst, 1, item.Тб.countProd, sheet.getLastColumn()).getValues();
    vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
    let col_name = nr("B");
    let col_list = ["№", "Номенклатура", "Ед", "Кол", "Цена", "Сумма", "Срок поставки", "Поставщик", "Статус", "Предполагаемая дата оплаты",
    ].map(c => { return heads.indexOf(c); }).filter(c => (c >= 0));
    vls.forEach(v => {
      if (item.Тб.ИсключитьСтроки.includes(v[0])) { return; }
      if (!v[col_name]) { return; }
      let obj = new Object();
      col_list.forEach(col => {
        obj[heads[col]] = v[col];

      });
      obj["Сводка"] = {
        Поставки: (() => {
          // if (getContext().timeConstruct.getHours() % 2 > 0) { return; }
          return mapСводная.get(obj["№"])
        })(),  // тут проблема

      };
      // retArr.push(heads);
      retArr.push(obj);
    });

    // return retArr;
    return {
      тело: retArr,
      шапка: heads.filter((h, i, arr) => { return col_list.includes(i); }),
    }
  }
  /** @private  */
  getValuesFromОплаты(item) {

    let rowBodyFirst = 3;
    let retArr = new Array();
    let sheet = this.context.getSheetByName(this.context.settings.sheetName_Оплаты);
    // let heads = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    let heads = sheet.getRange(2, 1, 1, nr("W")).getValues()[0];
    // heads = ["row"].concat(heads).map(h => fl_str(h));
    heads = ["row"].concat(heads).map(h => `${h}`.trim());
    heads = heads.map(v => {
      return `${v}`.replace("\n", " ").replace(/ +/g, ' ').trim();   // /[\n\r]/g,' '
    });
    let vls = sheet.getRange(rowBodyFirst, 1, item.Тб.countProd, sheet.getLastColumn()).getValues();
    vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
    let col_name = nr("B");

    // let col_list = ["№", "Номенклатура", "Юр лицо поставщика", "№ счета", "Статус", "Плательщик", "Дата оплаты", "% оплаты",].map(c => { return heads.indexOf(c); }).filter(c => (c >= 0));
    let col_names = ["№", "Номенклатура", "Юр лицо поставщика", "№ счета", "Статус", "Плательщик", "Дата оплаты", "% оплаты", "Статус оплаты"];

    let col_list = new Array();
    let col_set = new Set();

    heads = heads.map((h, i, arr) => {
      if (!col_names.includes(h)) { return h; }
      col_list.push(i);

      if (col_set.has(h)) {
        h = `${h}_${nc(i)}`;
      }
      col_set.add(h);
      return h;
    });


    vls.forEach(v => {
      if (!v[col_name]) { return; }
      let obj = new Object();
      col_list.forEach(col => {
        obj[heads[col]] = v[col];
        if (heads[col] == "Дата оплаты") {

          if (v[col] instanceof Date) {
            obj["Дата оплаты"] = `${Utilities.formatDate(v[col], "Europe/Moscow", "dd.MM.YYYY")}`;
            // obj["Дата оплаты📆"] = `📆\n${Utilities.formatDate(v[col], "Europe/Moscow", "dd.MM.YYYY")}`;
          }

        }

      });
      retArr.push(obj);
    });
    return {
      тело: retArr,
      шапка: heads.filter((h, i, arr) => { return col_list.includes(i); }),
      шапка_вся: heads,
    }

  }
  /** @private  */
  getValuesFromНастройки_писем(item) {
    let retObj = new Object();
    let sheet_Настройки_писем = this.context.getSheetByName(getSettings().sheetName_Настройки_писем);
    let vls_AJ1_AK = sheet_Настройки_писем.getRange("AJ1:AK").getValues();
    for (let i = 0; i < vls_AJ1_AK.length; i++) {
      let fild_name = vls_AJ1_AK[i][0];
      let fild_value = vls_AJ1_AK[i][1];
      if (!fild_name) { continue; }
      retObj[fild_name] = fild_value;
    }
    retObj["Число товаров"] = undefined;
    retObj["Таблица в 1 строке часть 1"] = undefined;
    retObj["Название компаний и номера счетов"] = undefined;
    // тампусто

    retObj["тампусто"] = "&";
    // retObj["тампусто"] = "&" + (() => { try { return sheet_Настройки_писем.getRange("AR1:AT6").getValues().flat(2).join("") } catch (err) { return mrErrToString(err); } })();
    // retObj["тампусто"] = getContext().getSheetList7().тампусто();
    // retObj["тампусто"] = getContext().getSheetPay().тампусто(1 / 24 / 60 * 4.9);

    return retObj;
  }

  /** @private  */
  getListOfИсключитьСтрокиИзТаблицыЛогиста(item) {
    let str_key = fl_str("Исключить строки из таблицы логиста");
    let ret_arr = new Array();

    let sheet_Настройки_писем = this.context.getSheetByName(getSettings().sheetName_Настройки_писем);
    let vls_AJ1_AK = sheet_Настройки_писем.getRange("AJ1:AK").getValues();
    for (let i = 0; i < vls_AJ1_AK.length; i++) {
      let fild_name = vls_AJ1_AK[i][0];
      let fild_value = vls_AJ1_AK[i][1];
      if (!fild_name) { continue; }
      if (fl_str(fild_name) != str_key) { continue; }
      ret_arr = ret_arr.concat(
        `${fild_value}`.split(" ").map(v => Number.parseInt(v))
        // .filter(v => {
        //   if (!Number.isInteger(v)) { false; }
        //   if (!Number.isFinite(v)) { false; }
        //   if (!Number.isSafeInteger(v)) { false; }
        //   true;
        // })
      );
      break;
    }
    return ret_arr;
  }

  getValuesFromУПД(item) {
    let retArr = new Array();
    try {
      let context = this.getContextТаблицаОплатВнешняя();
      // Logger.log(` getValuesFromУПД contest =${context}`);
      if (context) {
        // Logger.log(` getValuesFromУПД contest =${context}`);
        let MrClassSheetModel = MrLib_Midlle.getMrClassSheetModel();
        let sheetModelУПД = new MrClassSheetModel("Сводная по УПД", context);
        /** @type {Map} */
        let map = sheetModelУПД.getMap();
        retArr = [...map.values()];
      }
    } catch (err) { mrErrToString(err); }
    return retArr;
  }



  getValuesFromДОК(item) {
    let retArr = new Array();
    try {
      let context = this.getContextТаблицаОплатВнешняя();
      if (context) {
        let MrClassSheetModel = MrLib_Midlle.getMrClassSheetModel();
        let sheetModelУПД = new MrClassSheetModel("Сводная по Договорам", context);
        /** @type {Map} */
        let map = sheetModelУПД.getMap();
        retArr = [...map.values()];
      }
    } catch (err) { mrErrToString(err); }
    // Сводная по Договорам
    return retArr;
  }

  getValuesFromСводнаяПоНоменклатуре(item) {
    let retArr = new Array();
    try {
      let context = this.getContextТаблицаОплатВнешняя();
      if (context) {
        let MrClassSheetModel = MrLib_Midlle.getMrClassSheetModel();
        let sheetModelСводная = new MrClassSheetModel("Сводная по Номенклатуре", context);
        /** @type {Map} */
        let map = sheetModelСводная.getMap();
        retArr = [...map.values()];
      }
    } catch (err) { mrErrToString(err); }
    // Сводная по Номенклатуре
    return retArr;
  }

  getUrls(item) {
    return this.context.getUrls();
  }




  getСводкаПроекта() {
    let Плательщик = getContext().getSheetPay().getПлательщики().join("\n");
    let Посредник = getContext().getSheetPay().getПосредник();

    let СводнаяПрочее = new Object();
    // СводнаяПрочее["🧩"] = "🧩";
    СводнаяПрочее = getContext().getSheetPay().getСводнаяПрочее(СводнаяПрочее);
    let СводкаПроекта = {
      "🧩": "🧩",
      Плательщик,
      Посредник,
      СводнаяПрочее,
    }
    return СводкаПроекта;
  }


  getExportValues() {
    this.item.ver = "1";
    this.item.key = this.context.getSpreadSheet().getUrl();

    this.item.info = {
      // key: SpreadsheetApp.getActive().getUrl(),
      email: Session.getEffectiveUser().getEmail(),
      // date: new Date(),
      date: timeConstruct,
      // name: SpreadsheetApp.getActive().getName(),
    }

    this.item.Тб = new Object();


    this.item.Тб.name = this.context.getSpreadSheet().getName();
    this.item.Тб.rowBodyLastСбор_КП = this.findRowBodyLast(this.context.settings.sheetName_Лист_7);
    this.item.Тб.countProd = this.item.Тб.rowBodyLastСбор_КП - 2 + 1;
    this.item.Тб.ИсключитьСтроки = this.tryAddErr(this, this.getListOfИсключитьСтрокиИзТаблицыЛогиста.name, this.item);
    this.item.Тб.Настройки_писем = this.getValuesFromНастройки_писем(this.item);
    this.item.Тб["👨‍👩‍👧‍👦"] = "👨‍👩‍👧‍👦";
    this.item.Тб.Выбор_Поставщиков = this.getValuesFromВыбор_Поставщиков(this.item);
    this.item.Тб["💎"] = "💎";
    this.item.Тб.Оплаты = this.tryAddErr(this, this.getValuesFromОплаты.name, this.item);
    this.item.Тб["🧾"] = "🧾";
    this.item.Тб.УПД = this.tryAddErr(this, this.getValuesFromУПД.name, this.item);
    this.item.Тб.ДОК = this.tryAddErr(this, this.getValuesFromДОК.name, this.item);
    this.item.Тб.URLS = this.tryAddErr(this, this.getUrls.name, this.item);
    this.item.Тб.СводкаПроекта = this.tryAddErr(this, this.getСводкаПроекта.name, this.item);
    // this.item.Тб["ПоставокиПоНоменклатуре"] = this.tryAddErr(this, this.getValuesFromСводнаяПоставокПоНоменклатуре.name, this.item);
    // Logger.log("getExportValues " + JSON.stringify(this.item));
    return this.item;
  }
  /** @private  */
  tryAddErr(tt, colable, item) {
    try {
      // Logger.log(colable)
      return tt[colable](item);
    } catch (err) {
      item.errors = [].concat(mrErrToString(err), item.errors);
      mrErrToString(err);
    }
  }
}



//

// ----------------------------------------------



function triggerExportValuesToMidlle() {
  // return;
  // Logger.log(`triggerExportValuesToMidlle  S`);
  try { new MrClassExportProject(getContext(), new Object()).triggerExportValuesToMidlle(); } catch (err) { mrErrToString(err); }
  // try { new MrClassExportProject(getContext(), new Object()).triggerExportValuesToMidlle(); } catch (err) { mrErrToString(err); }
  // MrLib_Midlle.updateItems(this.sheetName_Проекты, [this.getExportValues()]);
  Logger.log(`triggerExportValuesToMidlle  F`);
}


