class MrClassSheetЗвонки {

  constructor(url = getContext().getSpreadSheet().getUrl(), prefix = "") {
    Logger.log(`MrClassSheetЗвонки constructor START`);


    this.url = url;
    this.prefix = prefix;

    // this.id = getContext().getSpreadSheet().getId(); // todo

    let ph = "https://docs.google.com/spreadsheets/d/";
    let id = undefined;
    id = (() => {
      let idm = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
      Logger.log(`MrClassSheetЗвонки constructor | ph=${ph} | url=${this.url}`);
      if (!this.url.includes(ph)) { Logger.log(1); Logger.log(!this.url.includes(ph)); return; }
      this.url = this.url.trim();
      if (!this.url) { Logger.log(2); return undefined; }
      if (this.url.length < (ph.length + idm.length)) { Logger.log(3); return undefined; }
      if (this.url.indexOf(ph) != 0) { Logger.log(4); return undefined; }
      let ret = this.url.slice(ph.length, ph.length + idm.length);
      Logger.log(ret);
      return ret;
    })();

    if (id) {
      if (this.url.includes(id)) { this.id = id; }
    }
    this.isValid = false;
    if (this.id) { this.isValid = true; }


    this.sheetName = getSettings().sheetName_Звонки;
    this.rowConf = {
      head: { first: 1, last: 1, key: 1, },
      body: { first: 2, last: 2, },
    };
    this.projNr = getContext().getНомерПроекта();
    // this.sheetModel_Звонки = undefined;
    this.projStatus = getContext().getСтатусПроекта();

    this.strs = {
      СутьВопросаСтатус: `Был отправлен первичный запрос.`,
      СутьВопроса: "Суть вопроса",
      ТекущийСтатус: "Текущий статус",
      НаКакуюПочтуПисали: "На какую почту писали",
      ПереченьТоваров: "Перечень товаров по которым требуется уточнение",
      Приоритет: "Приоритет",
    }

    Logger.log(`MrClassSheetЗвонки constructor ${id} | id=${this.id} | isValid=${this.isValid} | url=${this.url}`);

    // Logger.log(JSON.stringify({ n: this.projNr, s: this.projStatus }));
    this.requiredColNames();
    this.requiredKey();
    Logger.log(`MrClassSheetЗвонки constructor FINIS `);


  }

  requiredColNames() {
    let requiredColNamesArr = [
      "№ вопроса в проекте",
      "Перечень товаров по которым требуется уточнение",
      "Возможна ли поставка аналога",
      "С какой почты писали",
      "На какую почту писали",
      "Суть вопроса",
      "Комментарий по итогу звонка",
      "Дата след контакта",
      "Почта с которой получен ответ",
      "Текущий статус",
      // "Текущий статус проекта",
      "Ссылка на таблицу проекта",
      "№ проекта",
      "Письмо",
      "key",
      "Дата_Время_Добавления_Звонка",
      "Добавить Вопрос Проблему",
      this.strs.Приоритет,
    ];
    this.getSheetModel_Звонки().requiredColNames(requiredColNamesArr);


    let head_key = this.getSheetModel_Звонки().head_key;

    let col = head_key.indexOf(this.strs.Приоритет);
    if (col > 0) {
      let rule = SpreadsheetApp.newDataValidation().requireValueInList(["A", "B", "C", "D", "E", "Z", "-"]).build();
      this.getSheetModel_Звонки()
        .sheet
        .getRange(this.rowConf.body.first, col, this.getSheetModel_Звонки().sheet.getMaxRows() - this.rowConf.body.first + 1, 1)
        .setDataValidation(rule);
    }

    this.sheetModel_Звонки = undefined;
  }

  requiredKey() {

    // Logger.log(`MrClassSheetЗвонки requiredKey START`);
    /** @type {SpreadsheetApp.Sheet} */


    if (!this.isValid) { return; }
    let sheet = this.getSheetModel_Звонки().sheet;
    let colKey = this.getSheetModel_Звонки().col.key;
    let colNrQ = this.getSheetModel_Звонки().col["№ вопроса в проекте"];
    let colПроект = this.getSheetModel_Звонки().col["№ проекта"];
    let colТаблица = this.getSheetModel_Звонки().col["Ссылка на таблицу проекта"];
    // let colСтатусПроекта = this.getSheetModel_Звонки().col["Текущий статус проекта"];

    if (!this.projNr) { return; }
    // Logger.log(`MrClassSheetЗвонки requiredKey this.projN ${this.projN}`);
    // let rowBodyFirst = this.getSheetModel_Звонки().row.body.first;
    // let rowBodyLast = this.getSheetModel_Звонки().row.body.last;

    /** @type {[[]]} */
    let vls = this.getSheetModel_Звонки().getValues();
    // sheet.getRange().getValues()
    vls.forEach(v => {
      // Logger.log(v[0]);
      // Logger.log(v.slice(1).join(""))
      if (v.slice(1).join("") == "") { return; }

      if (v[colKey]) { return; }
      if (!v[colNrQ]) { v[colNrQ] = v[0] - 1; sheet.getRange(v[0], colNrQ).setValue(v[colNrQ]); }
      let key = `${this.prefix}${`00000${this.projNr}`.slice(-5)} № ${`000${v[colNrQ]}`.slice(-3)} Id:${this.id}`;
      sheet.getRange(v[0], colKey).setValue(key);
      sheet.getRange(v[0], colПроект).setValue(`${this.prefix}${this.projNr}`);
      if (v[colТаблица]) { return; }
      sheet.getRange(v[0], colТаблица).setValue(this.url);
    });

    // vls.forEach(v => {
    //   if (!v[colNrQ]) { return; }
    //   // sheet.getRange(v[0], colСтатусПроекта).setValue(this.projStatus);
    // });

    this.sheetModel_Звонки = undefined;

  }



  getSheetModel_Звонки() {
    // if (!this.isValid) { return; }
    // Logger.log(this.url);
    if (!this.sheetModel_Звонки) {
      this.sheetModel_Звонки = MrLib_Midlle.makeSheetModelBy(this.url, this.sheetName, this.rowConf);
    }
    return this.sheetModel_Звонки;
  }

  getStatistics() {

    let mapСтатусыЗвонков = getContext().getSheetНастройки_писем().getСтатусыЗвонков();

    let statusOnArr = new Array();
    let statusOfArr = new Array();

    mapСтатусыЗвонков.forEach((v, k) => {
      if (v === true) { statusOfArr.push(k); return; }
      statusOnArr.push(k);
    });

    /** @type */
    let map = this.getSheetModel_Звонки().getMap();
    let ret = {
      total: 0,
      statusOn: 0,
      statusOf: 0,
      statusNone: 0,
      statusOther: 0,
      empty: 0,
    }
    map.forEach(item => {
      if (!item[this.strs.ПереченьТоваров]) { ret.empty++; return; }
      ret.total++;
      if (!item[this.strs.ТекущийСтатус]) { ret.statusNone++; return; }
      if (statusOnArr.includes(item[this.strs.ТекущийСтатус])) { ret.statusOn++; return; }
      if (statusOfArr.includes(item[this.strs.ТекущийСтатус])) { ret.statusOf++; return; }
      if (item[this.strs.ТекущийСтатус]) { ret.statusOther++; return; }
      // if (item[this.strs.ТекущийСтатус]) { ret.empty++; return; }
    });
    Logger.log(JSON.stringify({ statusOnArr, statusOfArr }, null, 2));
    Logger.log(JSON.stringify(ret, null, 2));
    return ret;
  }

  update() {

    if (!this.isValid) { return; }

    // if (!this.projNr) { Logger.log("Не Установлен номер проекта"); return; }
    this.requiredKey();
    // return;

    let dataTaskSyncЗвонки = MrLib_Midlle.getDataTaskSyncЗвонки();
    dataTaskSyncЗвонки.param.urls.proj = this.url;

    // Logger.log(JSON.stringify(dataTaskSyncЗвонки, undefined, 2));

    MrLib_Midlle.ВыполненитьМасивЗадач("MrClassSheetЗвонки update ВыполненитьМасивЗадач", [dataTaskSyncЗвонки,], "MrClassSheetЗвонки");

  }

  servicesПоискЗвонков() {
    return;
    // this.requiredKey();
    // let limitDey = 1; /// вывести в твблицу на лист настнойки писем
    let limitDey = getContext().getSheetНастройки_писем().getDaysBetweenLetterAndCall();
    // Logger.log(`limitDey = ${limitDey}`);
    // return

    let dateNow = getContext().timeConstruct;
    Logger.log("Время: " + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
    Logger.log(`H${dateNow.getHours()} `);
    if (dateNow.getHours() != 8) {
      // if (dateNow.getHours() != 22) {
      // 
      Logger.log(`Час ${dateNow.getHours()} ` + "Не подходящее Время: " + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy hh:mm:ss")}`); return;
    }
    this.requiredKey();

    let daysAgo = (() => {
      let lDey = 0;
      // for (let i = 0; i < 15; i++) {
      // }
      let dAgo = 0;
      let date = getContext().timeConstruct;
      while (lDey < limitDey) {
        if (dAgo > 100) { break; }
        dAgo++;
        date = new Date(date.getTime() - DeyMilliseconds);
        if (date.getDay() == 6) { continue; }
        if (date.getDay() == 0) { continue; }
        // Logger.log(JSON.stringify(date).slice(0, 11));
        lDey++;
      }
      // Logger.log(lDey);
      return dAgo;
    })();
    // Logger.log(daysAgo);
    let timeAgo = daysAgo * DeyMilliseconds;
    let timeLeft = dateNow.getTime() - timeAgo;
    // Logger.log(daysAgo);
    // return


    let url_Письма = MrLib_Midlle.getSettings().urls.Письма;
    // let sheetName = "Отправленные";
    let sheetName_Отправленные = MrLib_Midlle.getSettings().sheetNames.Отправленные;
    let sheetModel_Отправленные = MrLib_Midlle.makeSheetModelBy(url_Письма, sheetName_Отправленные);
    /** @type {Map} */
    let map_Отправленные = sheetModel_Отправленные.getMap();
    let url_product = getContext().getUrls().product;
    let url_buy = getContext().getUrls().buy;
    let arrОтправленные = [...map_Отправленные.values()];


    arrОтправленные = arrОтправленные.filter(item => {
      if (item.Звонок === true) { return false };

      if ((item.url !== url_product) && (item.url !== url_buy)) { return false };

      if (!item.ДатаОтправки) { return false };
      if (typeof item.ДатаОтправки == "object") { item.ДатаОтправки = JSON.stringify(item.ДатаОтправки) }
      if (typeof item.ДатаОтправки == "string") { item.ДатаОтправки = new Date(JSON.parse(item.ДатаОтправки)) }
      if (item.ДатаОтправки.getTime() > timeLeft) { return false; }
      // Logger.log(item.ДатаОтправки);
      return true;
    });

    // Logger.log("Номенклатура");
    // Logger.log(arrОтправленные.map(item => item.Номенклатура).join("\n"));



    if (arrОтправленные.length == 0) {
      return;
    }


    let classSheetList1 = getContext().getSheetList1();



    let vlsProduct = classSheetList1.sheet.getRange(classSheetList1.rowBodyFirst, classSheetList1.columnProduct, classSheetList1.rowBodyLast - classSheetList1.rowBodyFirst + 1, 1).getValues();
    let vlsId = classSheetList1.sheet.getRange(classSheetList1.rowBodyFirst, classSheetList1.columnId, classSheetList1.rowBodyLast - classSheetList1.rowBodyFirst + 1, 1).getValues();
    let vlsКол_во = classSheetList1.sheet.getRange(classSheetList1.rowBodyFirst, classSheetList1.columnКол_во, classSheetList1.rowBodyLast - classSheetList1.rowBodyFirst + 1, 1).getValues();
    let vlsЕдИзм = classSheetList1.sheet.getRange(classSheetList1.rowBodyFirst, classSheetList1.columnЕдИзм, classSheetList1.rowBodyLast - classSheetList1.rowBodyFirst + 1, 1).getValues();

    let productMap = new Map();

    for (let i = 0; i < vlsProduct.length; i++) {
      // numberOfOffers
      let row = classSheetList1.rowBodyFirst + i;
      let numOffers = classSheetList1.numberOfOffers(row, 3);
      // Logger.log("productArr");
      // // Logger.log(productArr);
      if (numOffers.priceAll > 3) { continue; }
      productMap.set(fl_str(vlsProduct[i][0]), {
        name: vlsProduct[i][0],
        numOffers,
        id: vlsId[i][0],
        ЕдИзм: vlsЕдИзм[i][0],
        КолВо: vlsКол_во[i][0],
      });
    }

    Logger.log("productMap");
    Logger.log(productMap.size);


    // productMap.forEach(item => Logger.log(JSON.stringify(item, null, 2)));
    if (productMap.size == 0) { return; }

    let mapНовыеЗадачиНаЗвонок = new Map();
    // let arr_ОтправленныеtoUpdate = new Array();

    arrОтправленные.forEach(item => { mapНовыеЗадачиНаЗвонок.set(item.emailAddress, { from: item.from, "Номенклатуры": new Array() }); });

    arrОтправленные.forEach(item => {

      if (!productMap.has(fl_str(item.Номенклатура))) { return; }

      // Logger.log("item | " + JSON.stringify(item));
      // mapНовыеЗадачиНаЗвонок.get(item.emailAddress).Номенклатуры.push(item.Номенклатура);
      mapНовыеЗадачиНаЗвонок.get(item.emailAddress).Номенклатуры.push((() => {
        let obj = productMap.get(fl_str(item.Номенклатура));
        // name
        // numOffers
        // id
        // ЕдИзм
        // КолВо
        let ret = [
          `${`     ${obj.id}`.slice(-4)}№`,
          `${obj.КолВо} ${obj.ЕдИзм}`,
          obj.name,
        ];

        // Logger.log(ret.join("\t"))
        // return obj.Номенклатура;
        return ret.join("  ");
        // return `${obj.id}\t${obj.ЕдИзм}\t${obj.КолВо}\t${obj.name}`;
      })());
      item.Звонок = true;
    });

    // return;

    mapНовыеЗадачиНаЗвонок.forEach((v, k, m) => {
      // Logger.log(JSON.stringify({ k, v }));
      if (v.Номенклатуры.length == 0) { return; }
      let obj = {
        "Перечень товаров по которым требуется уточнение": v.Номенклатуры.join("\n"),
        "На какую почту писали": k,
        "Суть вопроса": this.strs.СутьВопросаСтатус,
        "С какой почты писали": v.from,
      }

      // MrLib_Midlle.makeSheetModelBy("https://docs.google.com/spreadsheets/d/1_kCo/", this.sheetName, this.rowConf).appendObj(obj);
      this.getSheetModel_Звонки().appendObj(obj);
    });

    arrОтправленные = arrОтправленные.filter(item => item.Звонок).map(item => { return { key: item.key, Звонок: item.Звонок, } });
    sheetModel_Отправленные.updateItems(arrОтправленные);

  }



  /**
  Если есть частичный ответ от поставщика (например запрашивали 10 позиций а получили ответ только на 5 шт) 
  В текст телефонного запроса вставляем краткую сводку 
  “Изначально был отправлен запрос на 10 позиций, на 5 позиций цены получены, а на следующие позиции цен нет:” Если не на одну позицию нет цены то краткую сводку не надо
   */
  servicesАктуализациЗвонков() {

    Logger.log(`servicesАктуализациЗвонков url = ${this.url} | vvvvvvvv `);
    if (!this.isValid) { return; }

    /**@type {Map} */
    let mapЗвонки = this.getSheetModel_Звонки().getMap();
    if (mapЗвонки.size == 0) { return; }
    /**@type {Map} */
    let mapЗапросы = getContext().getSheetНастройки_писем().getSheetModelОтпрЗапросы(this.url).getMap();
    let mapСтатусыЗвонков = getContext().getSheetНастройки_писем().getСтатусыЗвонков();

    Logger.log(` mapЗвонки.size = ${mapЗвонки.size}`);

    let arrЗвонки = [...mapЗвонки.values()].filter((item) => {
      if (!item.Письмо) { return false; }
      if (item["Текущий статус"]) {
        let tt = mapСтатусыЗвонков.get(item["Текущий статус"]);
        if (tt == true) { return false; }
      }
      return true;
    });
    Logger.log(` arrЗвонки.length = ${arrЗвонки.length}`);
    if (arrЗвонки.length == 0) {
      // Logger.log(` arrЗвонки.length = ${arrЗвонки.length}`);
      return;
    }

    // let arrПисьма = arrЗвонки.map(item => { return item.Письмо; });
    let mapЗапросыВПисьмах = new Map();
    arrЗвонки.forEach(item => {
      mapЗапросыВПисьмах.set(item.Письмо, { countPrice: 0, hasPrice: new Map(), notPrice: new Map() });
    });

    let shLt1 = getContext().getSheetList1();

    let vlsProduct = shLt1.sheet.getRange(shLt1.rowBodyFirst, shLt1.columnProduct, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1).getValues();
    let vlsId = shLt1.sheet.getRange(shLt1.rowBodyFirst, shLt1.columnId, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1).getValues();
    let vlsКол_во = shLt1.sheet.getRange(shLt1.rowBodyFirst, shLt1.columnКол_во, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1).getValues();
    let vlsЕдИзм = shLt1.sheet.getRange(shLt1.rowBodyFirst, shLt1.columnЕдИзм, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1).getValues();

    let mapProduct = new Map();
    let mapEmailProducts = shLt1.getEmailProductMap();

    for (let i = 0; i < vlsProduct.length; i++) {
      // numberOfOffers
      let row = shLt1.rowBodyFirst + i;
      let numOffers = shLt1.numberOfOffers(row, 3);
      // Logger.log("productArr");
      // // Logger.log(productArr);
      // if (numOffers.priceAll > 3) { continue; }
      mapProduct.set(fl_str(vlsProduct[i][0]), {
        name: vlsProduct[i][0],
        numOffers,
        id: vlsId[i][0],
        ЕдИзм: vlsЕдИзм[i][0],
        КолВо: vlsКол_во[i][0],
        info: (() => {
          let ret = [
            `${`   ${numOffers.priceAll}`.slice(-3)} цен`,
            `${`   ${vlsId[i][0]}`.slice(-3)}№`,
            `${vlsКол_во[i][0]} ${vlsЕдИзм[i][0]}`,
            vlsProduct[i][0],
          ];
          return ret.join("  ");
        })(),
      });
    }

    Logger.log(`mapЗапросы.size ${mapЗапросы.size}
    mapProduct.size ${mapProduct.size}`);
    mapЗапросы.forEach(item => {
      if (!mapЗапросыВПисьмах.has(item.Письмо)) { return; }
      // Logger.log(JSON.stringify(item, null, 2));
      mapЗапросыВПисьмах.get(item.Письмо).countPrice++;

      let Номенклатура = mapProduct.get(fl_str(item.Номенклатура));
      if (!Номенклатура) { return; }

      let it = { Запрос: item, Номенклатура, };
      let hasP = (() => {
        if (!mapEmailProducts.has(fl_str(item.to))) { return false; }
        if (!mapEmailProducts.get(fl_str(item.to)).includes(fl_str(it.Номенклатура.name))) { return false; }
        return true;
      })();


      if (hasP) { mapЗапросыВПисьмах.get(item.Письмо).hasPrice.set(Номенклатура.name, it); }
      else { mapЗапросыВПисьмах.get(item.Письмо).notPrice.set(Номенклатура.name, it); }
    });

    let upArrЗвоки = arrЗвонки.map(item => {
      let ret = {
        key: item.key,
        dd: item["На какую почту писали"]
      };
      ((obj) => {
        // obj["Перечень товаров по которым требуется уточнение"] = "";
        let arrStr = new Array()
        let rp = mapЗапросыВПисьмах.get(item.Письмо); if (!rp) { return undefined; }

        // arrStr.push(`Всего ${rp.countPrice}/ есть ${rp.hasPrice.size} / нет ${rp.notPrice.size}`);
        arrStr.push(`${rp.notPrice.size} нет / ${rp.hasPrice.size} есть / ${rp.countPrice}`);
        if (rp.hasPrice.size >= rp.countPrice) {
          obj["Текущий статус"] = "Ответы получены";
          // obj["Перечень товаров по которым требуется уточнение"] =
          arrStr.push(`Цены получены на ${rp.hasPrice.size} из ${rp.countPrice} запросов.`);
          // arrStr.push(`Цены получены на ${rp.hasPrice.size} из ${rp.countPrice} запросов.`);
        }

        if (rp.notPrice.size) {
          // arrStr.push(`На следующие ${rp.notPrice.size} запросы нет цен.`);
          //  arrStr.push( `На следующие позиции цен нет:” Если не на одну позицию нет цены то краткую сводку не надо`)
          arrStr = arrStr.concat([...rp.notPrice.values()].sort((a, b) => {
            if (a.Номенклатура.numOffers.priceAll < b.Номенклатура.numOffers.priceAll) { return -1; }
            if (a.Номенклатура.numOffers.priceAll > b.Номенклатура.numOffers.priceAll) { return 1; }
            return 0;
          }).map(item => `${item.Номенклатура.info}`));
        }

        obj["Перечень товаров по которым требуется уточнение"] = arrStr.join("\n");

      })(ret);
      return ret;
    });
    // Logger.log(JSON.stringify(upArrЗвоки, null, 2));

    this.getSheetModel_Звонки().updateItems([...upArrЗвоки.values()]);

  }

  servicesПоискЗвонковПоПисьмам() {
    Logger.log(`servicesПоискЗвонковПоПисьмам`);
    if (!this.isValid) { return; }
    // let limitDey = 1; /// вывести в твблицу на лист настнойки писем
    let limitDey = getContext().getSheetНастройки_писем().getDaysBetweenLetterAndCall();
    // Logger.log(`limitDey = ${limitDey}`);
    // return

    let dateNow = getContext().timeConstruct;
    Logger.log("Время: " + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
    Logger.log(`H${dateNow.getHours()} `);
    if (dateNow.getHours() != 8) {
      // if (dateNow.getHours() != 22) {
      Logger.log(`Час ${dateNow.getHours()} `
        + "Не подходящее Время: "
        + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
      return;
    }
    this.requiredKey();

    let daysAgo = (() => {
      let lDey = 0;
      // for (let i = 0; i < 15; i++) {
      // }
      let dAgo = 0;
      let date = getContext().timeConstruct;
      while (lDey < limitDey) {
        if (dAgo > 100) { break; }
        dAgo++;
        date = new Date(date.getTime() - DeyMilliseconds);
        if (date.getDay() == 6) { continue; }
        if (date.getDay() == 0) { continue; }
        // Logger.log(JSON.stringify(date).slice(0, 11));
        lDey++;
      }
      // Logger.log(lDey);
      return dAgo;
    })();
    // Logger.log(daysAgo);
    let timeAgo = daysAgo * DeyMilliseconds;
    let timeLeft = dateNow.getTime() - timeAgo;
    // Logger.log(daysAgo);
    // return


    let shMlОтпрПисьма = getContext().getSheetНастройки_писем().getSheetModelОтпрПисьма(this.url);
    let mapОтпрПисьма = shMlОтпрПисьма.getMap();


    // let url_product = getContext().getUrls().product;
    // let url_buy = getContext().getUrls().buy;
    let arr_Отпр_письма = [...mapОтпрПисьма.values()];

    Logger.log(arr_Отпр_письма.map(item => item.to).join("\n"));

    arr_Отпр_письма = arr_Отпр_письма.filter(item => {

      // if ((item.url !== url_product) && (item.url !== url_buy)) { return false };
      if (item.url !== this.url) { return false };
      if (item.Звонок) { return false };
      if (!item.ДатаОтправки) { return false };

      if (typeof item.ДатаОтправки == "object") { item.ДатаОтправки = JSON.stringify(item.ДатаОтправки) }
      if (typeof item.ДатаОтправки == "string") { item.ДатаОтправки = new Date(JSON.parse(item.ДатаОтправки)) }
      if (item.ДатаОтправки.getTime() > timeLeft) { return false; }
      // Logger.log(item.ДатаОтправки);
      return true;
    });

    // Logger.log("Письма");
    // Logger.log(arr_Отпр_письма.map(item => item.to).join("\n"));

    if (arr_Отпр_письма.length == 0) {
      Logger.log("Не найдены письма для добавление на Звонок");
      return;
    }

    arr_Отпр_письма.forEach((v) => {
      // Logger.log(JSON.stringify({ k, v }));
      // if (v.Номенклатуры.length == 0) { return; }

      let obj = {
        // "Перечень товаров по которым требуется уточнение": v.Номенклатуры.join("\n"),
        "На какую почту писали": v.to,
        // "Суть вопроса": this.strs.СутьВопросаСтатус,
        "С какой почты писали": v.from,
        "Суть вопроса": `${Utilities.formatDate(v.ДатаОтправки, "Europe/Moscow", "dd.MM.yyyy HH:mm")} Мы отправели вам письмо.`,
        "Письмо": v.key,
      }
      this.getSheetModel_Звонки().appendObj(obj);
      v.objЗвонок = obj;
      // Logger.log(obj.row);
    });


    this.requiredKey();
    this.getSheetModel_Звонки().reset();
    let mRow = new Map();
    this.getSheetModel_Звонки().getMap().forEach(v => {
      mRow.set(v.row, v);
    });

    arr_Отпр_письма = arr_Отпр_письма
      .map(item => {
        return {
          key: item.key,
          Звонок: (() => {
            try {
              let a = mRow.get(item.objЗвонок.row);
              // Logger.log(`${JSON.stringify(a)}`);
              return a.key;
            } catch (err) { mrErrToString(err); }
          })(),
        }
      }).filter(item => item.Звонок);

    shMlОтпрПисьма.updateItems(arr_Отпр_письма);
    return;

  }



  servicesПоискЗвонковПоПисьмамV2() {
    Logger.log(`servicesПоискЗвонковПоПисьмамV2 url = ${this.url} | vvvvvvvv `);
    if (!this.isValid) { return; }
    // let limitDey = 1; /// вывести в твблицу на лист настнойки писем
    let limitDey = getContext().getSheetНастройки_писем().getDaysBetweenLetterAndCall();
    // Logger.log(`limitDey = ${limitDey}`);
    // return

    let dateNow = getContext().timeConstruct;
    Logger.log("Время: " + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
    Logger.log(`H${dateNow.getHours()} `);
    if (dateNow.getHours() != 8) {
      // if (dateNow.getHours() != 22) {
      Logger.log(`Час ${dateNow.getHours()} `
        + "Не подходящее Время servicesПоискЗвонковПоПисьмамV2: "
        + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
      if (!testUrls.includes(this.url)) {
        return;
      }
    }
    this.requiredKey();

    let daysAgo = (() => {
      let lDey = 0;
      // for (let i = 0; i < 15; i++) {
      // }
      let dAgo = 0;
      let date = getContext().timeConstruct;
      while (lDey < limitDey) {
        if (dAgo > 100) { break; }
        dAgo++;
        date = new Date(date.getTime() - DeyMilliseconds);
        // if (date.getDay() == 6) { continue; }  // пропускает субботу
        // if (date.getDay() == 0) { continue; } // пропускает воскресенье
        // Logger.log(JSON.stringify(date).slice(0, 11));
        lDey++;
      }
      // Logger.log(lDey);
      return dAgo;
    })();
    // Logger.log(daysAgo);
    let timeAgo = daysAgo * DeyMilliseconds;
    let timeLeft = dateNow.getTime() - timeAgo;
    // Logger.log(daysAgo);
    // return


    let shMlОтпрПисьма = getContext().getSheetНастройки_писем().getSheetModelОтпрПисьма(this.url);
    let mapОтпрПисьма = shMlОтпрПисьма.getMap();


    // let url_product = getContext().getUrls().product;
    // let url_buy = getContext().getUrls().buy;
    let arr_Отпр_письма = [...mapОтпрПисьма.values()];

    // Logger.log(arr_Отпр_письма.map(item => item.to).join("\n"));

    arr_Отпр_письма = arr_Отпр_письма.filter(item => {

      // if ((item.url !== url_product) && (item.url !== url_buy)) { return false };
      if (item.url !== this.url) { return false };
      if (item.Звонок) { return false };
      if (!item.ДатаОтправки) { return false };

      if (typeof item.ДатаОтправки == "object") { item.ДатаОтправки = JSON.stringify(item.ДатаОтправки) }
      if (typeof item.ДатаОтправки == "string") { item.ДатаОтправки = new Date(JSON.parse(item.ДатаОтправки)) }
      if (item.ДатаОтправки.getTime() > timeLeft) { return false; }
      // Logger.log(item.ДатаОтправки);
      return true;
    });

    Logger.log("Письма");
    Logger.log(arr_Отпр_письма.map(item => item.to).join("\n"));

    if (arr_Отпр_письма.length == 0) {
      Logger.log("Не найдены письма для добавление на Звонок");
      return;
    }

    arr_Отпр_письма.forEach((v) => {
      // Logger.log(JSON.stringify({ k, v }));
      // if (v.Номенклатуры.length == 0) { return; }
      let obj = (() => {
        let zvArr = [...this.getSheetModel_Звонки().getMap().values()];
        // Logger.log(JSON.stringify(zvArr, null, 2));
        zvArr = zvArr.filter(obj => {
          if (fl_str(obj["На какую почту писали"]) != fl_str(v.to)) { return false; }
          if (fl_str(obj["С какой почты писали"]) != fl_str(v.from)) { return false; }
          if (obj["Текущий статус"]) {
            if (getContext().getSheetНастройки_писем().getСтатусыЗвонков().has(obj["Текущий статус"])) {
              if (getContext().getSheetНастройки_писем().getСтатусыЗвонков().get(obj["Текущий статус"]) == true) { return false; }
            }
          }
          return true;
        });
        if (zvArr.length == 0) { return undefined; }
        Logger.log(zvArr[0].row);
        // Logger.log(JSON.stringify(zvArr, null, 2));
        return zvArr[0];
      })();


      if (obj) { v.objЗвонок = obj; return; }
      obj = {
        // "Перечень товаров по которым требуется уточнение": v.Номенклатуры.join("\n"),
        "На какую почту писали": v.to,
        // "Суть вопроса": this.strs.СутьВопросаСтатус,
        "С какой почты писали": v.from,
        "Суть вопроса": `${Utilities.formatDate(v.ДатаОтправки, "Europe/Moscow", "dd.MM.yyyy HH:mm")} Мы отправели вам письмо.`,
        "Письмо": v.key,
      }
      this.getSheetModel_Звонки().appendObj(obj);
      this.getSheetModel_Звонки().reset();
      v.objЗвонок = obj;
      // Logger.log(obj.row);
    });


    this.requiredKey();
    this.getSheetModel_Звонки().reset();
    let mRow = new Map();
    this.getSheetModel_Звонки().getMap().forEach(v => {
      mRow.set(v.row, v);
    });

    arr_Отпр_письма = arr_Отпр_письма
      .map(item => {
        return {
          key: item.key,
          Звонок: (() => {
            try {
              let a = mRow.get(item.objЗвонок.row);
              // Logger.log(`${JSON.stringify(a)}`);
              return a.key;
            } catch (err) { mrErrToString(err); }
          })(),
        }
      }).filter(item => item.Звонок);

    shMlОтпрПисьма.updateItems(arr_Отпр_письма);
    return;

  }

  /** @param {Array} arrЗвонкиЗакрытые */
  doServicesАктуализациЗакрытыхЗвонков(arrЗвонкиЗакрытые) {
    // Logger.log("doServicesАктуализациЗакрытыхЗвонков S")
    // if (!isTest()) { return; }
    // Logger.log("doServicesАктуализациЗакрытыхЗвонков A1")
    if (arrЗвонкиЗакрытые.length == 0) { return; }
    let arrUpdateЗвонки = new Array();

    arrЗвонкиЗакрытые.forEach(item => {
      if (item[this.strs.Приоритет] == "-") { return; }
      if (item[this.strs.Приоритет] == "A") { return; }
      let itemT = { key: item.key, };
      itemT[this.strs.Приоритет] = "-";
      arrUpdateЗвонки.push(itemT);
    });
    // Logger.log("doServicesАктуализациЗакрытыхЗвонков A2")
    if (arrUpdateЗвонки.length == 0) { return; }
    this.getSheetModel_Звонки().updateItems(arrUpdateЗвонки);
    // Logger.log("doServicesАктуализациЗакрытыхЗвонков F")
  }

  servicesАктуализациЗвонковV2() {

    Logger.log(`servicesАктуализациЗвонковV2 url = ${this.url} | vvvvvvvv `);
    if (!this.isValid) { return; }

    /**@type {Map} */
    let mapЗвонки = this.getSheetModel_Звонки().getMap();
    if (mapЗвонки.size == 0) { return; }
    /**@type {Map} */
    let mapЗапросы = getContext().getSheetНастройки_писем().getSheetModelОтпрЗапросы(this.url).getMap();
    let mapПисьма = getContext().getSheetНастройки_писем().getSheetModelОтпрПисьма(this.url).getMap();
    let maloB8 = getContext().getSheetНастройки_писем().getMaloB8();
    let mapСтатусыЗвонков = getContext().getSheetНастройки_писем().getСтатусыЗвонков();

    mapЗвонки.forEach(item => { item.arrПисьма = new Array(); item.arrЗапросы = new Array(); });
    mapПисьма.forEach(item => { item.arrЗапросы = new Array(); });
    // mapЗвонки.forEach(item => { item.arrПисьма = new Array(); });

    mapЗапросы.forEach(item => {
      if (!mapПисьма.has(item.Письмо)) { return; }
      mapПисьма.get(item.Письмо).arrЗапросы.push(item);

      if (!mapЗвонки.has(mapПисьма.get(item.Письмо).Звонок)) { return; }
      mapЗвонки.get(mapПисьма.get(item.Письмо).Звонок).arrЗапросы.push(item);
    });

    mapПисьма.forEach(item => {
      if (!mapЗвонки.has(item.Звонок)) { return; }
      mapЗвонки.get(item.Звонок).arrПисьма.push(item);
    });


    let arrЗвонкиЗакрытые = new Array()


    Logger.log(` mapЗвонки.size = ${mapЗвонки.size}`);

    let arrЗвонки = [...mapЗвонки.values()].filter((item) => {
      // if (!item.Письмо) { return false; }
      if (item.arrПисьма.length == 0) { arrЗвонкиЗакрытые.push(item); return false; }
      if (item["Текущий статус"]) {
        if (mapСтатусыЗвонков.has(item["Текущий статус"])) {
          let tt = mapСтатусыЗвонков.get(item["Текущий статус"]);
          if (tt == true) { arrЗвонкиЗакрытые.push(item); return false; }
        }
      }
      return true;
    });

    Logger.log(` arrЗвонки.length = ${arrЗвонки.length}`);
    Logger.log(` arrЗвонкиЗакрытые.length = ${arrЗвонкиЗакрытые.length}`);

    if (arrЗвонкиЗакрытые.length != 0) {
      this.doServicesАктуализациЗакрытыхЗвонков(arrЗвонкиЗакрытые);
    }



    if (arrЗвонки.length == 0) {
      // Logger.log(` arrЗвонки.length = ${arrЗвонки.length}`);
      return;
    }
    // return;
    // let arrПисьма = arrЗвонки.map(item => { return item.Письмо; });

    let shLt1 = getContext().getSheetList1();

    let vlsProduct = shLt1
      .sheet
      .getRange(shLt1.rowBodyFirst, shLt1.columnProduct, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1)
      .getValues();
    let vlsId = shLt1.sheet
      .getRange(shLt1.rowBodyFirst, shLt1.columnId, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1)
      .getValues();
    let vlsКол_во = shLt1.sheet
      .getRange(shLt1.rowBodyFirst, shLt1.columnКол_во, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1)
      .getValues();
    let vlsЕдИзм = shLt1.sheet
      .getRange(shLt1.rowBodyFirst, shLt1.columnЕдИзм, shLt1.rowBodyLast - shLt1.rowBodyFirst + 1, 1)
      .getValues();

    let mapProduct = new Map();
    let mapEmailProducts = shLt1.getEmailProductMap();

    for (let i = 0; i < vlsProduct.length; i++) {
      // numberOfOffers
      let row = shLt1.rowBodyFirst + i;
      let numOffers = shLt1.numberOfOffers(row, maloB8);

      mapProduct.set(fl_str(vlsProduct[i][0]), {
        name: vlsProduct[i][0],
        numOffers,
        id: vlsId[i][0],
        ЕдИзм: vlsЕдИзм[i][0],
        КолВо: vlsКол_во[i][0],
        info: (() => {
          let ret = [
            `${`   ${numOffers.priceAll}`.slice(-3)} цен`,
            `${`   ${vlsId[i][0]}`.slice(-3)}№`,
            `${vlsКол_во[i][0]} ${vlsЕдИзм[i][0]}`,
            vlsProduct[i][0],
          ];
          return ret.join("  ");
        })(),
      });
    }



    arrЗвонки.forEach(item => {
      item.countЗапросы = { countPrice: item.arrЗапросы.length, hasPrice: new Map(), notPrice: new Map() };
    });


    Logger.log(`mapЗапросы.size ${mapЗапросы.size}    mapProduct.size ${mapProduct.size}`);

    arrЗвонки.forEach(itemЗвонк => {
      itemЗвонк.arrЗапросы.forEach(itemЗапрос => {

        let Номенклатура = mapProduct.get(fl_str(itemЗапрос.Номенклатура));
        if (!Номенклатура) { return; }

        let it = { Запрос: itemЗапрос, Номенклатура, };
        let hasP = (() => {
          if (!mapEmailProducts.has(fl_str(itemЗапрос.to))) { return false; }
          if (!mapEmailProducts.get(fl_str(itemЗапрос.to)).includes(fl_str(it.Номенклатура.name))) { return false; }
          return true;
        })();

        if (hasP) { itemЗвонк.countЗапросы.hasPrice.set(Номенклатура.name, it); }
        else { itemЗвонк.countЗапросы.notPrice.set(Номенклатура.name, it); }

      });
    });




    let upArrЗвоки = arrЗвонки.map(itemЗвонк => {
      let ret = {
        key: itemЗвонк.key,
        dd: itemЗвонк["На какую почту писали"],
      };

      ((obj) => {

        let rp = itemЗвонк.countЗапросы;
        if (!rp) { return undefined; }

        let arrStr = new Array()
        arrStr.push(`${rp.notPrice.size} нет / ${rp.hasPrice.size} есть / ${rp.countPrice}`);
        if (rp.hasPrice.size >= rp.countPrice) {
          obj["Текущий статус"] = "Ответы получены";
          arrStr.push(`Цены получены на ${rp.hasPrice.size} из ${rp.countPrice} запросов.`);
        }

        if (rp.notPrice.size) {
          arrStr = arrStr.concat([...rp.notPrice.values()].sort((a, b) => {
            if (a.Номенклатура.numOffers.priceAll < b.Номенклатура.numOffers.priceAll) { return -1; }
            if (a.Номенклатура.numOffers.priceAll > b.Номенклатура.numOffers.priceAll) { return 1; }
            return 0;
          }).map(item => `${item.Номенклатура.info}`));
        }
        obj["Перечень товаров по которым требуется уточнение"] = arrStr.join("\n");


        if (rp.notPrice.size) {
          let min = maloB8 + 1;
          min = Math.min(maloB8 + 1, ...[...rp.notPrice.values()].map(a => { return a.Номенклатура.numOffers.priceAll }));
          if (min > maloB8) {
            obj["Текущий статус"] = "Ответов достаточно";

          }
        }

        obj["Суть вопроса"] = itemЗвонк.arrПисьма.map(v => {
          return `${Utilities.formatDate(v.ДатаОтправки, "Europe/Moscow", "dd.MM.yyyy HH:mm")} Мы отправели вам письмо.`
        }).join("\n");

        obj[this.strs.Приоритет] = (() => {
          try {
            Logger.log(itemЗвонк[this.strs.Приоритет]);
            if (itemЗвонк[this.strs.Приоритет] == "A") { return undefined; }
          } catch (err) {
            mrErrToString(err);
          }

          try {
            if (rp.notPrice.size) {
              let min = Math.min(...[...rp.notPrice.values()].map(p => p.Номенклатура.numOffers.priceAll));
              // MrLib_Midlle.setFildForItem("Лог", obj.key, "min", min);

              if (isFinite(min)) {
                if (min == 0) return "B";
                if (min == 1) return "C";
                if (min == 2) return "D";
                if (min == 3) return "E";
                if (min > 3) return "-";
              }
            }

            if (rp.hasPrice.size >= rp.countPrice) {
              return "-";
            }

          } catch (err) {
            mrErrToString(err);
            // MrLib_Midlle.setFildForItem("Лог", obj.key, "err", mrErrToString(err));
          }
          return "-";
        }
        )();

        // MrLib_Midlle.setFildForItem("Лог", obj.key, "Приоритет", obj[this.strs.Приоритет]);
        // Logger.log([obj.key, "Приоритет", Math.min(...[...rp.notPrice.values()].map(p => p.Номенклатура.numOffers.priceAll)), obj[this.strs.Приоритет], ...[...rp.notPrice.values()].map(p => p.Номенклатура.numOffers.priceAll)]);

      })(ret);
      return ret;
    });

    // this.getSheetModel_Звонки().updateItems([...upArrЗвоки.values()]);
    this.getSheetModel_Звонки().updateItems(upArrЗвоки);
  }


  servicesПоискЗвонковПоПисьмамV3() {
    Logger.log(`servicesПоискЗвонковПоПисьмамV3 url = ${this.url} | vvvvvvvv `);
    if (!this.isValid) { return; }
    // let limitDey = 1; /// вывести в твблицу на лист настнойки писем
    let limitDey = getContext().getSheetНастройки_писем().getDaysBetweenLetterAndCall();
    // Logger.log(`limitDey = ${limitDey}`);
    // return

    let dateNow = getContext().timeConstruct;
    Logger.log("Время: " + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
    Logger.log(`H${dateNow.getHours()} `);
    // if (dateNow.getHours() != 8) {
    //   // if (dateNow.getHours() != 22) {
    //   Logger.log(`Час ${dateNow.getHours()} `
    //     + "Не подходящее Время servicesПоискЗвонковПоПисьмамV3: "
    //     + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
    //   if (!testUrls.includes(this.url)) {
    //     return;
    //   }
    // }
    this.requiredKey();

    let daysAgo = (() => {
      let lDey = 0;
      // for (let i = 0; i < 15; i++) {
      // }
      let dAgo = 0;
      let date = getContext().timeConstruct;
      while (lDey < limitDey) {
        if (dAgo > 100) { break; }
        dAgo++;
        date = new Date(date.getTime() - DeyMilliseconds);
        if (date.getDay() == 6) { continue; }
        if (date.getDay() == 0) { continue; }
        // Logger.log(JSON.stringify(date).slice(0, 11));
        lDey++;
      }
      // Logger.log(lDey);
      return dAgo;
    })();
    // Logger.log(daysAgo);
    let timeAgo = daysAgo * DeyMilliseconds;
    let timeLeft = dateNow.getTime() - timeAgo;
    // Logger.log(daysAgo);
    // return


    let shMlОтпрПисьма = getContext().getSheetНастройки_писем().getSheetModelОтпрПисьма(this.url);
    let mapОтпрПисьма = shMlОтпрПисьма.getMap();


    // let url_product = getContext().getUrls().product;
    // let url_buy = getContext().getUrls().buy;
    let arr_Отпр_письма = [...mapОтпрПисьма.values()];
    let setEm = new Set();
    // Logger.log(arr_Отпр_письма.map(item => item.to).join("\n"));

    arr_Отпр_письма = arr_Отпр_письма.filter(item => {

      // if ((item.url !== url_product) && (item.url !== url_buy)) { return false };
      if (item.url !== this.url) { return false };
      if (item.Звонок) { return false };
      if (!item.ДатаОтправки) { return false };

      if (typeof item.ДатаОтправки == "object") { item.ДатаОтправки = JSON.stringify(item.ДатаОтправки) }
      if (typeof item.ДатаОтправки == "string") { item.ДатаОтправки = new Date(JSON.parse(item.ДатаОтправки)) }
      if (item.ДатаОтправки.getTime() > timeLeft) { return false; }
      // Logger.log(item.ДатаОтправки);
      if (setEm.has(fl_str(item.to))) { return false; }
      setEm.add(fl_str(item.to));
      return true;
    });

    Logger.log("Письма");
    Logger.log(arr_Отпр_письма.map(item => item.to).join("\n"));

    if (arr_Отпр_письма.length == 0) {
      Logger.log("Не найдены письма для добавление на Звонок");
      return;
    }

    // return;
    arr_Отпр_письма = arr_Отпр_письма.slice(0, 20);

    arr_Отпр_письма.forEach((v) => {
      // Logger.log(JSON.stringify({ k, v }));
      // if (v.Номенклатуры.length == 0) { return; }


      let obj = (() => {
        let zvArr = [...this.getSheetModel_Звонки().getMap().values()];
        // Logger.log(JSON.stringify(zvArr, null, 2));
        zvArr = zvArr.filter(obj => {

          if (fl_str(obj["На какую почту писали"]) != fl_str(v.to)) { return false; }
          if (fl_str(obj["С какой почты писали"]) != fl_str(v.from)) { return false; }
          if (obj["Текущий статус"]) {
            if (getContext().getSheetНастройки_писем().getСтатусыЗвонков().has(obj["Текущий статус"])) {
              if (getContext().getSheetНастройки_писем().getСтатусыЗвонков().get(obj["Текущий статус"]) == true) { return false; }
            }
          }
          return true;
        });
        if (zvArr.length == 0) {
          // Logger.log(`${v.to} = "net"`);
          return undefined;
        }
        // Logger.log(`${v.to} = ${zvArr[0].row}`);
        // Logger.log(JSON.stringify(zvArr, null, 2));
        return zvArr[0];
      })();


      if (obj) { v.objЗвонок = obj; return; }
      obj = {
        // "Перечень товаров по которым требуется уточнение": v.Номенклатуры.join("\n"),
        "На какую почту писали": v.to,
        // "Суть вопроса": this.strs.СутьВопросаСтатус,
        "С какой почты писали": v.from,
        "Суть вопроса": `${Utilities.formatDate(v.ДатаОтправки, "Europe/Moscow", "dd.MM.yyyy HH:mm")} Мы отправели вам письмо.`,
        "Письмо": v.key,
        "Дата_Время_Добавления_Звонка": getContext().timeConstruct,

      }
      this.getSheetModel_Звонки().appendObj(obj);

      // this.getSheetModel_Звонки().reset();

      // this.sheetModel_Звонки = undefined;
      v.objЗвонок = obj;
      // Logger.log(obj.row);
      // this.requiredKey();
    });
    // return;

    this.requiredKey();
    this.getSheetModel_Звонки().reset();
    let mRow = new Map();
    this.getSheetModel_Звонки().getMap().forEach(v => {
      mRow.set(v.row, v);
    });

    arr_Отпр_письма = arr_Отпр_письма
      .map(item => {
        return {
          key: item.key,
          Звонок: (() => {
            try {
              let a = mRow.get(item.objЗвонок.row);
              // Logger.log(`${JSON.stringify(a)}`);
              return a.key;
            } catch (err) { mrErrToString(err); }
          })(),
        }
      }).filter(item => item.Звонок);

    shMlОтпрПисьма.updateItems(arr_Отпр_письма);
    return;

  }

}

function triggerUpdateЗвонки() {

  if (getContext().timeConstruct.getHours() >= 9) {
    Logger.log("В рабочее время не обновляем");
    return;
  }

  Logger.log("triggerUpdateЗвонки");
  let звонки = new MrClassSheetЗвонки();
  if (звонки.isValid) { звонки.update(); }
  Logger.log("triggerUpdateЗвонкиВнешняя");

  let urlbuy = getContext().getUrls().buy;
  if (!urlbuy) { return; }

  let звонкиВ = new MrClassSheetЗвонки(urlbuy, "В-");
  if (звонкиВ.isValid) { звонкиВ.update(); }

}























