// MisterVova@mail.ru 23.12.2020 добавил файл  
let testUrls = [

];
function isTest() {
  Logger.log("--==:isTest:==--")
  let testUrls = [

  ];
  let url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  if (testUrls.includes(url)) { return true; }
  return false;
}
function isEmail(email) {
  if (!email) { return false; }
  if (!email.toString().includes("@")) { return false; }

  return true;
}
let DeyMilliseconds = 24 * 60 * 60 * 1000;
let CacheExpirationInSeconds = 60 * 70; // 70 минут

class MrClassSettings {

  constructor() {
    this.ver = "по умолчанию "
    // this.sheetName = "";

    //    Лист1
    //    Лист7
    //    черновикG
    //    Лист18
    //    Вопросы
    //    Доставка
    //    Экспертиза
    //    Экспертиза(копия)
    //    Доставка(копия)
    //    Группы товаров
    //    Лист3
    //    Отзывы о поставщиках
    //    Товары
    //    Настройки писем

    this.contextTypes = {
      pay: "pay",
      buy: "buy",
      product: "product",
      summary: "summary",
    }

    this.sheetName_Паспорт_проекта = "0-Паспорт проекта";

    // this.sheetName_Лист_1 = "Лист1"; //
    this.sheetName_Лист_1 = "1-1 Сбор КП"; //
    this.sheetName_Группы_товаров = "1-2 Запросы по товарным группам"; //
    this.sheetName_Актуальный_Сбор_КП = "1-1-1 Актуальный Сбор КП из внешней"; // 
    // this.sheetName_Вопросы = "2 Вопросы от Поставщиков"; //
    this.sheetName_Вопросы = "2-1 Вопросы от Поставщиков"; //
    this.sheetName_ВопросыО = "2 Вопросы от Поставщиков"; //
    this.sheetName_Проблемы = "2-2 Проблемы по проекту";
    this.sheetName_Лист_3 = "4 Подготовка КП";  //

    this.sheetName_Доставка = "3-1 Расчет стоимости доставки"; //
    this.sheetName_Экспертиза = "3-2 Экспертиза предложений"; //
    this.sheetName_Лист_7 = "3-3 Выбор поставщиков"; //
    this.sheetName_Закупка_товара = "5 Закупка товара"; //
    this.sheetName_Закупка_товара_Внешняя = "5 Закупка товара Внешняя"; //
    this.sheetName_Оплаты = "6 Оплаты"; //
    // this.sheetName_Оплаты_Внешняя = "6 Оплаты"; //  6 Оплаты Внешняя
    this.sheetName_Оплаты_Внешняя = "6 Оплаты Внешняя"; //  6 Оплаты Внешняя
    // this.sheetName_черновик = "черновик";
    // this.sheetName_Лист_18 = "Лист18";
    // this.sheetName_Экспертиза(копия) = "Экспертиза(копия)";
    // this.sheetName_Доставка(копия) = "Доставка(копия)";

    this.sheetName_Отзывы_о_поставщиках = "Отзывы о поставщиках"; //
    this.sheetName_Товары = "Товары"; //
    this.sheetName_Настройки_писем = "Настройки писем"; //
    this.sheetName_Исполнение_Внешние = "7 Исполнение Внешние"
    this.sheetName_Технический_лист = "Технический лист";
    this.sheetName_Звонки = "1-3 Звонки";

    this.sheetName_Отпр_письма = "Отправленные письма";
    this.sheetName_Отпр_запросы = "Отправленные запросы";
    this.sheetName_Экспорт_в_битрикс = "Экспорт в битрикс";

    this.sheetName_Расс_Зак = "Расс Зак";
    this.sheetName_Расс_Прод = "Расс Прод";
    this.sheetName_mem_ОплатыВнешняя = "mem";
    this.sheetName_mem_Оплаты = "mem";


    this.sheetName_РасчетРентабельностиКП = "4-1 Расчет рентабельности КП";

    // Лист1 = 1-1 Предложения поставщиков
    // Группы товаров  = 1-2 Группы товаров
    // Вопросы = 2 Вопросы от Поставщиков

    // Доставка = 3-1 Расчет стоимости доставки
    // Экспертиза = 3-2 Экспертиза предложений
    // Лист7 = 3-3 Выбор поставщиков

    // Закупка = 5 Закупка товара
    // Лист3 = 4 Подготовка КП


    this.str_Ссылка_на_таблицу_закупки = "Ссылка на таблицу закупки"; //
    
    this.url_таблица_закупки_шаблон = ""; // рабочая

    // this.url_таблица_оплаты_шаблон = ""; // тест
    this.url_таблица_оплаты_шаблон = ""; //
    this.folderId = "1R4u9zhnnajoZSZvP6"; // папка для хранения внешних Таблиц "Закупка Товара Внешняя"
    

    // this.folderId_оплаты = this.folderId; // папка для хранения внешних Таблиц "Оплата Товара Внешняя"
    this.folderId_оплаты = "1KPzN5uB1xG7a6_"; // папка для хранения внешних Таблиц "Оплата Товара Внешняя"



    this.settingsBuyExternal = {
      "адрес таблици": SpreadsheetApp.getActive().getUrl(),
      "имя листа": "5 Закупка товара",
      "имя листа внешний": "5 Закупка товара Внешняя",
      "шапка начало": 2,
      "шапка конец": 2,
      "тело начало": 3,
      "тело конец": 5,
      "диапозон1 начало": "A",
      "диапозон1 конец": "Q",
      "диапозон2 начало": "Z",
      "диапозон2 конец": "AI",
      "диапозон3 начало": "R",
      "диапозон3 конец": "W",
      "диапозон4 начало": "AJ",
      "диапозон4 конец": "AO",
      "@ Столбец с @": "A",
      "диапозон5 начало": "AX",
      "диапозон5 конец": "AX",
    }

    this.storedRanges = [
      // { first: nr("R"), last: nr("W") },
      { first: nr(this.settingsBuyExternal["диапозон3 начало"]), last: nr(this.settingsBuyExternal["диапозон3 конец"]) },
      // { first: nr("AJ"), last: nr("AO") },
      { first: nr(this.settingsBuyExternal["диапозон4 начало"]), last: nr(this.settingsBuyExternal["диапозон4 конец"]) },
    ];

    this.cacheName = "ProgectDocumentCache";

    // this.log()


    this.colorNames = {
      DEFAULT: "DEFAULT",
      ЗаливкаПриОтправке: "Заливка при отправке",
      ЗаливкаПоставщика: "Заливка поставщика",
      ЗаливкаГруппы: "Заливка группы",
      Официал: "Официал",
      EmailИзБазыТоваров: "E-mail из базы товаров",
      Аналог: "Аналог",
      ПервоеСовпадение: "Первое совпадение",
      ВтороеСовпадение: "Второе совпадение",
      ТретьеМовпадение: "Третье совпадение",
      ТриггерыДобавлены: "Триггеры Добавлены",



    }

  }

  setSettings(newSettings) {
    for (let key in newSettings) {
      this[key] = newSettings[key];
      Logger.log(`key ${key} =  ${newSettings[key]};`)
    }
    Logger.log(`Версия настройки =  ${this.ver}`)
  }


  log() {
    // let email = Session.getEffectiveUser().getEmail();
    // let owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
    // Logger.log(`email ${email}`);
    // Logger.log(`owner ${owner}`);

    // for (let key in this) {
    //   Logger.log(`key ${key} = ${JSON.stringify(this[key])};`);
    // }
    // let url = getContext().getSpreadSheet().getUrl();

  }




  getSheetTemplateByName(sheetName) {
    let url = "https://docs.google.com/spreadsheets/d/1oPY/edit";
    let tms = {}
    tms["Отправленные письма"] = { url, sheetName: "Отправленные письма", };
    tms["Отправленные запросы"] = { url, sheetName: "Отправленные запросы", };
    tms["mem"] = { url, sheetName: "mem", };
    tms["99-0 Проект"] = { url, sheetName: "99-0 Проект" };
    tms["99-1 Номенклатуры"] = { url, sheetName: "99-1 Номенклатуры" };
    tms["99-2 Поставшики"] = { url, sheetName: "99-2 Поставшики" };
    tms["99-3 Цены"] = { url, sheetName: "99-3 Цены" };
    tms["99-4 Коментарий"] = { url, sheetName: "99-4 Коментарий" };
    tms["99-4 Модели"] = { url, sheetName: "99-4 Модели" };




    return tms[sheetName];
  }

}

function setSettings(newSettings) {
  return getSettings().setSettings(newSettings);
}


// let mrSettings = new MrClassSettings();
/** @type  {MrClassSettings} */
let mrSettings = undefined;

/** @returns  {MrClassSettings} */
function getSettings() {
  if (!mrSettings) {
    mrSettings = new MrClassSettings();
  }
  return mrSettings;
}






class MrContext {
  constructor(spreadSheetURL = SpreadsheetApp.getActive().getUrl(), settings = getSettings()) { // class constructor
    //     Logger.log(`MrContext  constructor spreadSheetURL=${spreadSheetURL}\nEffectiveUser:${Session.getEffectiveUser().getEmail()}`);
    Logger.log(`MrContext  constructor spreadSheetURL=${spreadSheetURL}`);

    if (!spreadSheetURL) { throw new Error("spreadSheetURL id not Valid"); }
    if (!`${spreadSheetURL}`.includes("https://docs.google.com/spreadsheets/")) { throw new Error("spreadSheetURL id not Valid"); }

    this.spreadSheetURL = spreadSheetURL;
    this.spreadSheet = undefined;
    if (!settings) { throw new Error("settings"); }
    /** @type {MrClassSettings} */
    this.settings = settings;
    this.contextType = this.settings.contextTypes.product;

    this.counteragentMap = undefined;
    this.colorMap = undefined;
    this.statusMap = undefined;
    this.productMap = undefined;
    this.sheetProductGroups = undefined;
    this.sheetList7 = undefined;
    this.sheetList1 = undefined;
    this.sheetVoprosi = undefined;
    this.sheetExpertise = undefined;

    // 20.02.2021 MrVova
    this.sobachkaMap = undefined;

    this.tMin = 5000;
    this.timeConstruct = timeConstruct;
    Logger.log(`MrContext constructor Finish`);

  }

  setContextType(contextType = getSettings().contextTypes.product) {
    this.contextType = contextType;
  }

  getSpreadSheet() {
    // if (!this.spreadSheet) {
    //   this.spreadSheet = SpreadsheetApp.openByUrl(this.spreadSheetURL);
    // }
    // return this.spreadSheet;

    return SpreadsheetApp.getActive();

  }



  getUserEmail() {
    return Session.getEffectiveUser().getEmail();
  }



  // 20.02.2021 MrVova
  getRowSobachkaBySheetName(sheetName) {
    if (!this.sobachkaMap) { this.sobachkaMap = new SobachkaMap() }
    return this.sobachkaMap.getRowSobachkaBySheetName(sheetName);
  }


  // 14.07.2022 MrVova
  getSobachkaMap() {
    if (!this.sobachkaMap) { this.sobachkaMap = new SobachkaMap() }
    return this.sobachkaMap
  }



  isSheetExpertise(sheetName) {
    return this.getSheetNamesExpertise().includes(fl_str(sheetName));
  }

  addRowsForAllSheetExpertise(rowArr) {
    let sheetNamesExpertise = getContext().getSheetNamesExpertise();
    sheetNamesExpertise.forEach(sheetName => new ClassSheet_Expertise(sheetName).addRows(rowArr));
  }

  getSheetNamesExpertise() {
    // let strExpertise = fl_str("Экспертиза");
    let strExpertise = fl_str(getSettings().sheetName_Экспертиза);
    let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    let sheetNames = sheets.map(s => fl_str(s.getName()));
    sheetNames = sheetNames.filter(sn => sn.includes(strExpertise));
    return sheetNames;
  }

  getSheetExpertise(sheetName) {
    return new ClassSheet_Expertise(sheetName);


    if (!this.sheetExpertise) { this.sheetExpertise = new ClassSheet_Expertise(); }
    return this.sheetExpertise;
  }

  getSheetVoprosi() {
    if (!this.sheetVoprosi) { this.sheetVoprosi = new ClassSheet_Voprosi(); }
    return this.sheetVoprosi;
  }


  getSheetПроблемы() {
    if (!this.sheetПроблемы) { this.sheetПроблемы = new ClassSheet_Проблемы(); }
    return this.sheetПроблемы;
  }

  getSheetList1() {
    if (!this.sheetList1) { this.sheetList1 = new ClassSheet_List1(); }
    return this.sheetList1;
  }

  getSheetList7() {
    if (!this.sheetList7) { this.sheetList7 = new ClassSheet_List7(); }
    return this.sheetList7;
  }

  getSheetBuy() {
    if (!this.sheetBuy) { this.sheetBuy = new ClassSheet_Buy(); }
    return this.sheetBuy;
  }


  getSheetPay() {
    if (!this.sheetPay) { this.sheetPay = new ClassSheet_Pay(); }
    return this.sheetPay;
  }


  getSheetProductGroups() {
    if (!this.sheetProductGroups) { this.sheetProductGroups = new ClassSheet_ProductGroups(); }
    return this.sheetProductGroups;
  }


  getSheetНастройки_писем() {
    if (!this.sheetНастройки_писем) { this.sheetНастройки_писем = new ClassSheet_НастройкиПисем(); }
    return this.sheetНастройки_писем;
  }

  getSheetЗвонкиПроект() {
    if (!this.sheetЗвонкиПроект) { this.sheetЗвонкиПроект = new MrClassSheetЗвонки(); }
    return this.sheetЗвонкиПроект;
  }



  getCounteragentMap() {
    if (!this.counteragentMap) { this.counteragentMap = new CounteragentMap(); }
    return this.counteragentMap;
  }

  getColorMap() {
    if (!this.colorMap) { this.colorMap = new ColorMap(); }
    return this.colorMap;
  }

  getStatusMap() {
    if (!this.statusMap) { this.statusMap = new StatusMap(); }
    return this.statusMap;
  }

  getProductMap() {
    if (!this.productMap) { this.productMap = new ProductMap(); }
    return this.productMap;
  }

  getCounteragent(email) {
    return this.getCounteragentMap().getCounteragent(email);
  }

  getDefaultColor() {
    return this.getColorMap().def;
  }

  getOfficialColor() {
    return this.getColorMap().getOfficialColor();
  }


  getAnalogColor() {
    return this.getColorMap().getAnalogColor();
  }

  getIsFromProductListColor() {
    return this.getColorMap().getIsFromProductListColor();
  }


  isOfficial(email) {
    let ca = this.getCounteragentMap().getCounteragent(email);

    if (!ca) return false;
    if (!(ca instanceof Counteragent)) return false;
    return ca.isOfficial();
    // return this.getCounteragentMap().getCounteragent(email).isOfficial();
  }


  getCounteragentColor(email) {
    email = fl_str(email);

    let ca = this.getCounteragent(email);

    if (!ca) { return undefined; }

    if (this.isOfficial(email)) return this.getOfficialColor();
    if (ca.getStatus().toString() == fl_str("Нет статуса")) {
      if (ca.isFromProductList()) { return getContext().getIsFromProductListColor(); }
    }
    return ca.getStatus().getColor();

  }

  getProduct(prodStr) {
    return this.getProductMap().getProduct(prodStr);

  }


  getColor(colorStr) {
    return this.getColorMap().getColor(colorStr);

  }

  getStatus(statusStr) {
    return this.getStatusMap().getStatus(statusStr);
  }


  upDate() {
    Logger.log("MrContext.upDate() начало");

    if (this.colorMap) { this.getColorMap().upDate(); }
    if (this.statusMap) { this.getStatusMap().upDate(); }
    if (this.counteragentMap) { this.getCounteragentMap().upDate(); }
    if (this.productMap) { this.getProductMap().upDate(); }


    Logger.log("MrContext.upDate() конец");
    return this;
  }

  canSendLetter(email) {
    if (!email) { return true; }
    email = fl_str(email);
    let ca = this.getCounteragent(email);
    if (!ca) { return true; }


    // return this.getStatusMap().canSendLetter(ca.getStatus());
    return ca.canSendLetter();
  }



  logMapElements(value, key, map) {
    console.log(`m[${key}] = ${value.logString()}`);
  }

  logMapElements2(value, key, map) {
    console.log(`m[${key}] = ${value}`);
  }

  logToConsole() { // class method

    this.getCounteragentMap().counteragentMap.forEach(this.logMapElements);


    this.getStatusMap().statusMap.forEach(this.logMapElements);

    this.getColorMap().colorMap.forEach(this.logMapElements2);
  }

  getProductGroup(row) {

    return this.getSheetProductGroups().getGroupByRow(row);

    // throw "nenene устарело ";
    // return new ProductGroup(row);
  }

  makeProduct(text) {
    return this.getProductMap().makeProduct(text);
  }



  /*листы */

  // getSheetByName(sheetName) {

  //   let ss = SpreadsheetApp.getActive();
  //   let sh = ss.getSheetByName(sheetName);
  //   if (!sh) { throw `нет листа с именем ${sheetName}`; }
  //   return sh;
  // }


  getSheetByName(sheetName, ssa = this.getSpreadSheet()) {
    let sh = ssa.getSheetByName(sheetName);
    if (!sh) {
      try {
        sh = this.makeSheetByName(sheetName);
      } catch (err) { mrErrToString(err); }
    }
    if (!sh) { throw new Error(`нет листа с именем ${sheetName} в таблице ${ssa.getUrl()}  |  ${this.spreadSheetURL} `); }
    return sh;
  }


  makeSheetByName(sheetName) {
    let sheetTemplate = this.settings.getSheetTemplateByName(sheetName);
    Logger.log(`sheetTemplate = ${sheetTemplate}`);
    if (!sheetTemplate) { return undefined; }
    /** @type {[[]]} */
    let vls = new Array();
    try {
      vls = SpreadsheetApp.openByUrl(sheetTemplate.url).getSheetByName(sheetTemplate.sheetName).getDataRange().getValues();
    } catch (err) { mrErrToString(err); return undefined; }

    let retSh = this.getSpreadSheet().insertSheet(sheetName);
    retSh.hideSheet();

    if (vls.length > 0) {
      retSh.getRange(1, 1, vls.length, vls[0].length).setValues(vls);
    }
    return retSh;
  }





  getFormulaR1C1(sheetName, row, col) {
    return makeFormulaR1C1(sheetName, row, col);
  }

  hasTime(duration, tMin = this.tMin) {
    let tDey = 24 * 60 * 60 * 1000;
    let tDuration = duration * tDey;
    let tStart = this.timeConstruct;
    let tNow = new Date();
    let tDef = tNow - tStart;
    let tHas = tDuration - tDef;

    // Logger.log(` 
    // duration=${duration}
    // tDuration=${tDuration}
    // tMim=${tMin}
    // tHas=${tHas}
    // ret=${(tHas < tMin ? false : true)}
    // `);

    if (tHas < tMin) { return false; }
    return true;
  }


  getUrls() {
    if (!this.urls) {
      let pay = (() => { try { return this.getSheetPay().getUrlExternalSpreadSheet(); } catch { } })();
      let buy = (() => { try { return this.getSheetBuy().getUrlExternalSpreadSheet(); } catch { } })();
      let product = (() => { try { return this.getSpreadSheet().getUrl(); } catch { } })();
      let summary = undefined;
      let callCoord = "https://docs.google.com/spreadsheets/d/1xYfT9Ie4/edit"
      this.urls = {
        pay,
        buy,
        product,
        // summary,
        callCoord,
      }
    }
    // Logger.log(`  this.contextТаблицаОплатВнешняя =${this.contextТаблицаОплатВнешняя}`);
    return this.urls;
  }

  getНомерПроекта() {
    if (!this.НомерПроекта) {
      let url = this.getSpreadSheet().getUrl();
      this.НомерПроекта = MrLib_Midlle.getItemByKey("Проекты", url, "НомерПроекта", undefined);
      if (!this.НомерПроекта) {
        this.НомерПроекта = MrLib_Midlle.getItemByKey("НомераПроектов", url, "НомерПроекта", undefined);
      }
    }
    return this.НомерПроекта;
  }


  getСтатусПроекта() {
    if (!this.СтатусПроекта) {
      let НомерПроекта = this.getНомерПроекта();

      if (НомерПроекта) {
        this.СтатусПроекта = MrLib_Midlle.getItemByKey("Сводная", `${НомерПроекта}`, "Статус проекта", undefined);
      } else {
        this.СтатусПроекта = "Проект не найден!!!";
      }
      // this.СтатусПроекта = `№ ${НомерПроекта} Статус: ${this.СтатусПроекта}.`;
    }
    return this.СтатусПроекта;
  }



  getСтатусыЗакрытияПроекта() {
    if (!this.СтатусыЗакрытияПроекта) {
      let mapDictionarys = MrLib_Midlle.getMapDictionarys();
      this.СтатусыЗакрытияПроекта = [...mapDictionarys.values()].map(v => v["Статусы остановить скрипт"]).filter(v => v);
      // Logger.log(JSON.stringify(arr));
    }
    return this.СтатусыЗакрытияПроекта;
  }



  getИсполнение_date() {
    if (!this.Исполнение_date) {
      let НомерПроекта = this.getНомерПроекта();

      if (НомерПроекта) {
        this.Исполнение_date = MrLib_Midlle.getItemByKey("Сводная", `${НомерПроекта}`, "Исполнение_date", undefined);
      } else {
        this.Исполнение_date = undefined;
      }
      // this.СтатусПроекта = `№ ${НомерПроекта} Статус: ${this.СтатусПроекта}.`;
    }
    return this.Исполнение_date;
  }



  getMemCell() {
    if (!this.memCell) {
      this.memCell = this.getSheetByName(getSettings().sheetName_Настройки_писем).getRange("B1");
    }
    return this.memCell;
  }

  getMemObj() {
    // let ret = this.getMemCell().getValue();
    let ret = this.getMemCell().getNote();
    if (!ret) { ret = JSON.stringify(new Object()) };
    ret = (() => { try { return JSON.parse(ret); } catch { return new Object(); } })();
    return ret;
  }

  getMem(key) {
    return this.getMemObj()[key];
  }

  setMem(key, value) {
    let mem = this.getMemObj();
    mem[key] = value;
    // this.getMemCell().setValue(JSON.stringify(mem));
    this.getMemCell().setNote(JSON.stringify(mem, null, 2));
  }


  generateNextTimeId() {
    if (this.memGenerateIdKode >= 999) { this.memGenerateIdKode = undefined; Utilities.sleep(1000); this.memGenerateIdTime = undefined; }
    if (!this.memGenerateIdKode) { this.memGenerateIdKode = 1; } else { this.memGenerateIdKode++; }
    if (!this.memGenerateIdTime) { this.memGenerateIdTime = new Date(); }
    let kode = `000${this.memGenerateIdKode}`.slice(-3);
    return `${this.memGenerateIdTime.getTime()}_${kode}`;
  }






  getCacheDocument() {
    if (!this.cacheDocument) {
      this.cacheDocument = CacheService.getDocumentCache().get(this.settings.cacheName);
      // Logger.log(`getScriptCache 1  this.scriptCache  = ${JSON.stringify(this.scriptCache)}`);
      this.cacheDocument = (() => { try { return JSON.parse(this.cacheDocument); } catch (err) { mrErrToString(err); return new Object() } })();
      // Logger.log(`getScriptCache 2  this.scriptCache  = ${JSON.stringify(this.scriptCache)}`);
    }

    if (!this.cacheDocument) {
      this.cacheDocument = new Object();
    }

    // Logger.log(`getScriptCache  this.scriptCache  = ${JSON.stringify(this.scriptCache)}`);
    return this.cacheDocument;
  }


  saveCacheDocument() {
    // Logger.log(`saveCache  this.scriptCache  = ${JSON.stringify(this.scriptCache)}`);
    CacheService.getDocumentCache().put(this.settings.cacheName, JSON.stringify(this.getCacheDocument()), CacheExpirationInSeconds);
  }

  removeCacheDocument() {
    // Logger.log(`removeCache  this.scriptCache  = ${JSON.stringify(this.scriptCache)}`);
    CacheService.getDocumentCache().remove(this.settings.cacheName);
    // CacheService.getDocumentCache()
  }





















}



let timeConstruct = new Date();
// let mrContext = new MrContext();
let mrContext = undefined;

/** @returns  {MrContext} */
function getContext() {
  if (!mrContext) { mrContext = new MrContext(); }
  return mrContext;
}



/** @returns  {MrContext} */
function makeContext(spreadSheetURL, settings = getSettings()) {
  let mrContext = new MrContext(spreadSheetURL, settings);
  return mrContext;
}



function makeSheetModelBy(url, sheetName) {
  return MrLib_Midlle.makeSheetModelBy(url, sheetName);
}

