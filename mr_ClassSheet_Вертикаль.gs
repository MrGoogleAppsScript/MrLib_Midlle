class MrClassВертикаль {
  constructor() {


    this.url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    this.id = SpreadsheetApp.getActiveSpreadsheet().getId();
    this.sheetNames = {
      Проект: "99-0 Проект",
      Номенклатуры: "99-1 Номенклатуры",
      // Поставшики: "99-2 Поставшики",
      Поставшики: "99-2 Поставщики",
      Цены: "99-3 Цены",
      Коментарий: "99-4 Коментарий",
      Модели: "99-4 Модели",
      Сбор_КП: getSettings().sheetName_Лист_1,
    }

    getContext().getSheetByName(this.sheetNames.Номенклатуры);
    // getContext().getSheetByName(this.sheetNames.Номенклатуры);
    getContext().getSheetByName(this.sheetNames.Поставшики);
    getContext().getSheetByName(this.sheetNames.Цены);
    getContext().getSheetByName(this.sheetNames.Коментарий);
    getContext().getSheetByName(this.sheetNames.Проект);
    getContext().getSheetByName(this.sheetNames.Модели);

    this.Проект = getContext().getНомерПроекта();
    this.sheetModelsMap = new Map();
    this.rowConf = {
      head: { first: 1, last: 1, key: 1, },
      body: { first: 2, last: 2, },
    };
  }

  getSheetModel(url, sheetName, rowConf) {
    let k = JSON.stringify({ url, sheetName }, null, 2);
    if (!this.sheetModelsMap.has(k)) {
      this.sheetModelsMap.set(k, MrLib_Midlle.makeSheetModelBy(url, sheetName, rowConf));
    }
    return this.sheetModelsMap.get(k);
  }


  собратьНоменклатуру() {
    // getContext().getSheetByName(getSettings().sheetName_Лист_1);

    // getContext().getSheetByName(this.sheetNames.Номенклатуры);
    let smНо = this.getSheetModel(this.url, this.sheetNames.Номенклатуры);
    /** @type {Map} */
    let mapНо = smНо.getMap();
    let mapKeyНо = new Map();

    mapНо.forEach(item => mapKeyНо.set(JSON.stringify({ "№": `${item["№"]}`, Номенклатура: item.Номенклатура }), item.key));

    let colProduct = getClassColRow().list1_col_Product;
    let smL1 = this.getSheetModel(this.url, getSettings().sheetName_Лист_1, this.rowConf);
    /** @type {Map} */
    let mapL1 = smL1.getMap();
    let items = [...mapL1.values()]
      .filter(item => { if (!item["Номенклатура"]) { return false; } if (!item["№"]) { return false; } return true; })
      .map(item => {
        // let key = getContext().generateNextTimeId();
        // key = { Проект: this.Проект, id: this.id, Номенклатура: key };
        // key = JSON.stringify(key);
        return {
          // key,  // Номенклатура	Характеристики	Ед. Изм	Кол-во	Цена	Итого	Предложений	НМЦ	Запросить		Группа
          key: mapKeyНо.get(JSON.stringify({ "№": `${item["№"]}`, Номенклатура: item.Номенклатура })),
          "№": item["№"],
          Номенклатура: item["Номенклатура"],
          Наименование: item["Номенклатура"],
          Характеристики: item["Характеристики"],
          "Ед. Изм": item["Ед. Изм"],
          "Кол-во": item["Кол-во"],
          Цена: item["Цена"],
          Итого: item["Итого"],
          Предложений: item["Предложений"],
          НМЦ: item["НМЦ"],
          Группа: item["Группа"],
          ДатаДобавления: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"),
          Нет: false,
          НетДата: "",
          //  :item[""],

          range: {
            str: `'${this.sheetNames.Сбор_КП}'!$${nc(colProduct)}$${item.row}`,
            row: item.row,
            col: colProduct,
            sheetName: this.sheetNames.Сбор_КП,
          },
          unique: JSON.stringify({ "№": item["№"], Номенклатура: item.Номенклатура }),
          // МодельАналог_unique: `'${this.sheetNames.Сбор_КП}'!$${nc(colProduct)}$${v.row}`,
        }
      });


    let mapHas = new Map();
    let mapAdd = new Map();
    let mapNon = new Map();

    items.forEach(item => {
      if (!item.key) {
        let key = getContext().generateNextTimeId();
        key = JSON.stringify({ Проект: this.Проект, id: this.id, Номенклатура: key });
        item.key = key;
        item.ДатаДобавления = Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy")
        mapAdd.set(item.key, item);
      }

      if ((() => { try { return (mapНо.get(item.key).Нет === true) } catch (err) { /** mrErrToString(err); */ return false; } })()) {
        mapAdd.set(item.key, { key: item.key, Нет: false, ДатаДобавления: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"), });
      }
      mapHas.set(item.key, item);
    });

    mapНо.forEach(item => {
      if (mapHas.has(item.key)) { return; }
      mapNon.set(item.key, { key: item.key, Нет: true, НетДата: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"), });
    })
    let ap = [].concat(...mapNon.values(), ...mapAdd.values());
    // Logger.log(ap)
    smНо.updateItems(ap);

  }

  собратьПоставщиков() {
    Logger.log("собратьПоставщиков S");
    // getContext().getSheetByName(this.sheetNames.Номенклатуры);
    let smПо = this.getSheetModel(this.url, this.sheetNames.Поставшики);

    /** @type {Map} */
    let mapПо = smПо.getMap();
    let mapKeyПо = new Map();

    mapПо.forEach(item => mapKeyПо.set(item.email, item.key));



    let smL1 = this.getSheetModel(this.url, getSettings().sheetName_Лист_1, this.rowConf);
    /** @type {Map} */
    let mapL1 = smL1.getMap();
    let headsL1 = smL1.head_key.join(" ");
    let m = getContext().getSheetProductGroups().sheet.getRange("J4:BK").getValues().flat().join(" ");
    let emails_str = [headsL1, m].join(" ");
    // Logger.log(emails_str);
    let emails_arr = emails_str.split(" ").filter(e => {
      if (!e.includes("@")) { return false; }
      if (!e.includes(".")) { return false; }
      return true;
    });
    let emails_set = new Set();
    emails_arr.forEach(e => emails_set.add(e.toLowerCase()))

    emails_arr = [...emails_set.values()].sort((a, b) => {
      let da = a.split("@");
      let db = b.split("@");

      if (da[1] > db[1]) { return 1; }
      if (da[1] < db[1]) { return -1; }

      if (da[0] > db[0]) { return 1; }
      if (da[0] < db[0]) { return -1; }
      return 0;
    });
    Logger.log(emails_arr.length);

    let items = emails_arr
      .map(item => {
        return {
          key: mapKeyПо.get(item),  // Номенклатура	Характеристики	Ед. Изм	Кол-во	Цена	Итого	Предложений	НМЦ	Запросить		Группа
          domen: item.split("@")[1],
          email: item,
          Почта: item,
          unice: fl_str(item),
        }
      });


    let mapHas = new Map();
    let mapAdd = new Map();
    let mapNon = new Map();

    items.forEach(item => {
      if (!item.key) {
        let key = getContext().generateNextTimeId();

        key = { Проект: this.Проект, id: this.id, Почта: item.email, };
        key = JSON.stringify(key);
        item.key = key;
        item.ДатаДобавления = Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy");
        mapAdd.set(item.key, item);
      }

      if ((() => { try { return (mapПо.get(item.key).Нет === true) } catch (err) { return false; } })()) {
        mapNon.set(item.key, { key: item.key, Нет: false, ДатаДобавления: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"), });
      }
      mapHas.set(item.key, item);
    });

    mapПо.forEach(item => {
      if (mapHas.has(item.key)) { return; }
      mapNon.set(item.key, { key: item.key, Нет: true, НетДата: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"), });
    })
    let ap = [].concat(
      ...mapNon.values(),
      // ...mapAdd.values(),
    );

    // Logger.log(ap)
    [...mapAdd.values()].forEach(item => smПо.appendObj(item))
    smПо.updateItems(ap);

    // Logger.log("собратьПоставщиков S");
    Logger.log(JSON.stringify(ap, undefined, 2));
    Logger.log("собратьПоставщиков F");

  }



  собратьЦены() {
    Logger.log("собратьЦены S");

    /** @type {Map} */
    let smПо = this.getSheetModel(this.url, this.sheetNames.Поставшики);
    /** @type {Map} */
    let mapНо = this.getSheetModel(this.url, this.sheetNames.Номенклатуры).getMap();
    let mapПо = smПо.getMap();
    // let arrНо = [...mapНо.values()].map(e => e["Номенклатура"]);
    let mapНоЛ = new Map();
    mapНо.forEach(e => { mapНоЛ.set(e.unique, e.key) });



    let mapПоЛ = new Map();
    mapПо.forEach(e => { mapПоЛ.set(e["email"], e.key) });

    let smL1 = this.getSheetModel(this.url, getSettings().sheetName_Лист_1, this.rowConf);
    let vls = getContext().getSheetByName(getSettings().sheetName_Лист_1).getDataRange().getValues();
    /** @type {Map} */
    let mapL1 = smL1.getMap();
    //  mapL1.forEach(e=>{Logger.log(JSON.stringify(e,undefined,2))})
    let items = new Array();


    // сбор текущих Цен
    mapL1.forEach((v, k) => {
      // Logger.log("_________________________________________________________________");
      // Logger.log(v["Номенклатура"]);
      if (!v["Номенклатура"]) { return }
      let pu = JSON.stringify({ "№": v["№"], Номенклатура: v.Номенклатура });
      if (!mapНоЛ.has(pu)) { return; }
      let ном = mapНо.get(mapНоЛ.get(pu));


      // Logger.log(ном);
      // Logger.log(v.row);

      for (let почта in v) {
        // Logger.log("+++++++++++++");
        // Logger.log(Постовщик);
        if (!почта.includes("@")) { continue; }
        let col = smL1.head_key.indexOf(почта);
        let Цена = v[почта];
        if (!Цена) { continue; }
        // Logger.log(Цена);
        // Logger.log(col);

        // Logger.log(vls[v.row-1][col-1+0]);
        // Logger.log(vls[v.row-1][col-1+1]);
        // Logger.log(vls[v.row-1][col-1+2]);
        // Logger.log(vls[v.row-1][col-1]);

        почта.split(" ").forEach(p => {
          if (!p.includes("@")) {
            // Logger.log(p);
            return;
          }
          let key = getContext().generateNextTimeId();
          // key = `Проект:T-тест id:${this.id} Цена:${key}`
          key = { Проект: this.Проект, id: this.id, Цена: key };
          key = JSON.stringify(key);
          let пос = mapПо.get(mapПоЛ.get(p));
          items.push({
            key,
            // Модель: "",
            Постовщик: пос,
            Номенклатура: ном,
            Цена: vls[v.row - 1][col - 1 + 0],
            Срок: vls[v.row - 1][col - 1 + 1],
            // Модель_Наименование: fl_str(vls[v.row - 1][col - 1 + 2]),
            Модель_Наименование: vls[v.row - 1][col - 1 + 2],
            ДатаЦены: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"),

            range: {
              str: `'${this.sheetNames.Сбор_КП}'!$${nc(col)}$${v.row}`,
              row: v.row,
              col: col,
              sheetName: this.sheetNames.Сбор_КП,
            },
            unique: `'${this.sheetNames.Сбор_КП}'!$${nc(col)}$${v.row}`,
            // Модель_unique: `'${this.sheetNames.Сбор_КП}'!$${nc(col + 2)}$${v.row}`,
            Модель_unique: fl_str(vls[v.row - 1][col - 1 + 2]),
            ТипЦены: ((п) => {
              // ♣️♥️♦️♠️
              if (`${п}`.includes("♠")) { return "♠"; }
              if (`${п}`.includes("♦️")) { return "♦️"; }
              if (`${п}`.includes("♣️")) { return "♣️"; }
              if (`${п}`.includes("♥️")) { return "♥️"; }
              return "";
            })(почта),

          });
        })
      }
    });

    // let mapSЦены = this.getSheetModel(this.url, this.sheetNames.Цены).getMap();
    let mapFromSheetМодели = this.getSheetModel(this.url, this.sheetNames.Модели).getMap();

    // let mapUЦены = new Map(); mapSЦены.forEach(e => { mapUЦены.set(e.unique, e.key) });
    let mapUniqueМодели = new Map(); mapFromSheetМодели.forEach(e => { mapUniqueМодели.set(e.unique, e.key) });



    let mapFromЦеныМодели = new Map();

    // сбор текущих Моделей
    items.forEach(item => {
      // Logger.log(JSON.stringify(item, null, 2));
      if (!item.Модель_Наименование) {
        item.Модель_Наименование = item.Номенклатура.Наименование;
        item.Модель_unique = fl_str(item.Модель_Наименование);
        item.Аналог = false;
      } else {
        item.Аналог = true;
      }

      let Модель = mapFromЦеныМодели.get(item.Модель_unique);
      if (!Модель) {
        // let key = getContext().generateNextTimeId();
        // key = { Проект: this.Проект, id: this.id, Модель: key };
        // key = JSON.stringify(key);

        let k = mapUniqueМодели.get(item.Модель_unique);
        if (k) { Модель = mapFromSheetМодели.get(k); }
        if (!Модель) {
          Модель = {
            Наименование: item.Модель_Наименование,
            Аналог: item.Аналог,
            range: {
              str: `'${this.sheetNames.Сбор_КП}'!$${nc(item.range.col + 2)}$${item.range.row}`,
              row: item.range.row,
              col: item.range.col + 2,
              sheetName: this.sheetNames.Сбор_КП,
            },
            // unique:  item.Аналог === true? `'${this.sheetNames.Сбор_КП}'!$${nc(item.range.col + 2)}$${item.range.row}`:"",
            unique: item.Модель_unique,
          }
        }
        // Модель.Нет = false;
        mapFromЦеныМодели.set(item.Модель_unique, Модель);
      }

      if (!Модель.key) {
        let key = getContext().generateNextTimeId();
        key = { Проект: this.Проект, id: this.id, Модель: key };
        Модель.key = JSON.stringify(key);
      }


      let { key, Наименование, Аналог, unigue } = mapFromЦеныМодели.get(item.Модель_unique);
      item.Модель = { key, Наименование, Аналог, unigue };
      // item.Модель_Наименование = undefined; 

    });


    let smЦн = this.getSheetModel(this.url, this.sheetNames.Цены);
    // smЦн.updateItems(items);
    smЦн.setItems(items);

    Logger.log("собратьЦены F");


    let mapHasМодели = new Map();
    let mapAddМодели = new Map();
    let mapNotМодели = new Map();

    Logger.log("собратьЦены Модели S");

    [...mapFromЦеныМодели.values()].forEach(item => {
      if (!item.ДатаДобавления) {
        item.ДатаДобавления = Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy")
        item.Нет = false;
        mapAddМодели.set(item.key, item);
      }

      if ((() => { try { return (mapFromSheetМодели.get(item.key).Нет === true) } catch (err) { /** mrErrToString(err); */ return false; } })()) {
        mapNotМодели.set(item.key, { key: item.key, Нет: false, ДатаДобавления: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"), });
      }

      mapHasМодели.set(item.key, item);
    });

    mapFromSheetМодели.forEach(item => {
      if (mapHasМодели.has(item.key)) {

        return;
      }
      mapNotМодели.set(item.key, { key: item.key, Нет: true, НетДата: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy"), });
    })

    let ap = [].concat(
      ...mapNotМодели.values(),
      // ...mapAddМодели.values()
    );

    // Logger.log(ap)


    let smМд = this.getSheetModel(this.url, this.sheetNames.Модели);

    [...mapAddМодели.values()].forEach(item => { smМд.appendObj(item); })


    // smМд.updateItems([...mapFromЦеныМодели.values()]);

    smМд.updateItems(ap);





    Logger.log("собратьЦены Модели F");
  }


  собратьКомментарий() {

  }

  run() {
    // if (!this.ttUrls.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { return; }
  }

  собрать() {
    // if (!this.ttUrls.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { return; }
    this.собратьНоменклатуру();
    this.собратьПоставщиков();
    this.собратьЦены();
  }


}

function servisВертикаль() {
  Logger.log("servisВертикаль");
  if (!isTest()) { return; }
  let mrClassВертикаль = new MrClassВертикаль();
  mrClassВертикаль.собрать();
}


function testMrClassВертикаль() {
  // let r = getContext().getSheetЗвонкиПроект().getStatistics();

  // return;

  Logger.log("testMrClassВертикаль");
  let mrClassВертикаль = new MrClassВертикаль();
  mrClassВертикаль.собрать();
}