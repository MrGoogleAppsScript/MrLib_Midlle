class DataTaskJoinProgectInfo {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_JoinProgectInfo.name;
    /** @constant */
    this.name = undefined;
    /** @constant */
    this.актуально_дней = 1 / 24 * 3;
    this.off = false;
    this.param = {

      sheetNames: {
        from: "",
        to: "Сводная",
      },

      urls: {
        from: "",
        to: getSettings().urls.Проекты,
      },

      prefix: "",
      fild: "",
      mem: "обновлено",
      rows: {
        headFirst: 1,
        bodyFirst: 2,
      },
    };
  }
}




class ExeTask_JoinProgectInfo extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_ExportReestr.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",
    }
    this.mapValuesFromРеестр = new Map();
    this.mapValuesFromИсполнение = new Map();
  }

  /** @param {DataTaskJoinProgectInfo} task*/
  taskGetРеестр(task) {
    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModel_to = new MrClassSheetModel(task.param.sheetNames.to, context_to);


    if ((new Date().getTime() - new Date(sheetModel_to.getMem(task.param.mem)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    ) { Logger.log("да ещё актуальны"); return; }



    // let sheetName = task.param.sheetName;
    let rows = task.param.rows;

    let rowheadFirst = rows.headFirst;
    let rowBodyFirst = rows.bodyFirst;


    let itemArr = new Array();
    let sheet_from = context_from.getSheetByName(task.param.sheetNames.from);

    let heads = sheet_from.getRange(rowheadFirst, 1, 1, sheet_from.getLastColumn()).getValues()[0];
    heads = ["row"].concat(heads);//.map(h => fl_str(h));
    let vls = sheet_from.getRange(rowBodyFirst, 1, sheet_from.getLastRow() - rowBodyFirst + 1, sheet_from.getLastColumn()).getValues();
    vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
    let col_Filter = nr("A");

    let col_list = [
      "row",
      "№",
      "Заказчик",
      "Название",
      "Таблица",
      "Ответственный",
      "Статус проекта",
      "Статус",
      // "Вид подписания договора",

    ].concat(heads).map(c => {
      return heads.indexOf(c);
    }).filter(c => (c >= 0));


    vls.forEach(v => {
      if (!v[col_Filter]) { return; }

      let obj = new Object();
      col_list.forEach(col => {
        obj[heads[col]] = v[col];
      });

      // if (fl_str(obj["Ответственный"]) != fl_str(task.param.Ответственный)) { return; }

      obj["key"] = `${task.param.prefix}${obj["№"]}`;
      obj["row_строка"] = obj["row"];
      obj["row"] = undefined;

      // this.mapValuesFromРеестр.set(obj["key"], obj);

      let ph = [{ key: `${obj["key"]}` }, obj];
      itemArr.push(ph);

      // Logger.log(JSON.stringify(ph));
    });

    // let context_to = new MrContext(task.param.url, getSettings());
    itemArr = itemArr.map(([item, obj]) => {
      // let obj = this.mapValuesFromРеестр.get(item.key);
      let url = `${obj["Таблица"]}`.slice(0, 88);

      if (url.length < 88) {
        let url_patern = "https://docs.google.com/spreadsheets/d/ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ/edit";
        let ok_url = `${url}${url_patern.slice(url.length)}`;
        if (!ok_url.includes("Я")) {
          url = ok_url;
        }
      }

      let ret = {
        key: item.key,
        "№": obj["№"],
        "Заказчик": obj["Заказчик"],
        "Название": obj["Название"],
        "Ссылка": obj["Ссылка"],
        "Таблица": url,
        // "Статус проекта": obj["Статус проекта"],
        "Статус проекта": obj["Статус"],
        "Ответственный": obj["Ответственный"],
        // "obj":obj,
        // Р: this.getExportValues(context_from, item),
      }
      ret[task.param.fild] = obj;
      ret[`${task.param.fild}_date`] = getContext().timeConstruct;
      return ret;
    });


    sheetModel_to.updateItems(itemArr);
    sheetModel_to.setMem(task.param.mem, new Date());
    this.saveInFile(itemArr);

  }

  /** @param {DataTaskJoinProgectInfo} task*/
  taskGetИсполнение(task) {
    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModel_to = new MrClassSheetModel(task.param.sheetNames.to, context_to);

    // Logger.log("ещё актуальны?");
    // Logger.log(sheetModelProjects.getMem(this.names.обновлено));
    // Logger.log("не актуальны | " + `${new Date().getTime()}  | ${new Date(sheetModelProjects.getMem(this.names.обновлено)).getTime() + task.актуально_дней * DeyMilliseconds} | 
    // ${(new Date().getTime() - new Date(sheetModelProjects.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds)} `);

    if ((new Date().getTime() - new Date(sheetModel_to.getMem(task.param.mem)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    ) { Logger.log("да ещё актуальны"); return; }



    // let sheetName = task.param.sheetName;
    let rows = task.param.rows;

    let rowheadFirst = rows.headFirst;
    let rowBodyFirst = rows.bodyFirst;


    let itemArr = new Array();
    let sheet_from = context_from.getSheetByName(task.param.sheetNames.from);

    let heads = sheet_from.getRange(rowheadFirst, 1, 1, sheet_from.getLastColumn()).getValues()[0];
    heads = ["row"].concat(heads);//.map(h => fl_str(h));
    let vls = sheet_from.getRange(rowBodyFirst, 1, sheet_from.getLastRow() - rowBodyFirst + 1, sheet_from.getLastColumn()).getValues();
    vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
    let col_Filter = nr("A");

    let col_list = [
      "row",
      "№",
      "Заказчик",
      "Название",
      "Таблица",
      "Ответственный",
      "Статус проекта",
      "Вид подписания договора",
    ].concat(heads).map(c => {
      return heads.indexOf(c);
    }).filter(c => (c >= 0));


    vls.forEach(v => {
      if (!v[col_Filter]) { return; }

      let obj = new Object();
      col_list.forEach(col => {
        obj[heads[col]] = v[col];
      });

      // if (fl_str(obj["Ответственный"]) != fl_str(task.param.Ответственный)) { return; }

      obj["key"] = `${task.param.prefix}${obj["№"]}`;
      obj["row_строка"] = obj["row"];
      obj["row"] = undefined;

      // this.mapValuesFromРеестр.set(obj["key"], obj);

      let ph = [{ key: `${obj["key"]}` }, obj];
      itemArr.push(ph);

      // Logger.log(JSON.stringify(ph));
    });

    // let context_to = new MrContext(task.param.url, getSettings());
    itemArr = itemArr.map(([item, obj]) => {
      // let obj = this.mapValuesFromРеестр.get(item.key);
      let url = `${obj["Таблица"]}`.slice(0, 88);

      if (url.length < 88) {
        let url_patern = "https://docs.google.com/spreadsheets/d/ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ/edit";
        let ok_url = `${url}${url_patern.slice(url.length)}`;
        if (!ok_url.includes("Я")) {
          url = ok_url;
        }
      }

      let ret = {
        key: item.key,
        "№": obj["№"],
        "Заказчик": obj["Заказчик"],
        "Название": obj["Название"],
        "Ссылка": obj["Ссылка"],
        "Таблица": url,
        "Статус проекта": obj["Статус проекта"],
        "Ответственный": obj["Ответственный"],
        "Вид подписания договора": obj["Вид подписания договора"],

      }
      ret[task.param.fild] = obj;
      ret[`${task.param.fild}_date`] = getContext().timeConstruct;
      return ret;
    });


    sheetModel_to.updateItems(itemArr);
    sheetModel_to.setMem(task.param.mem, new Date());
    this.saveInFile(itemArr);

  }

  /** @param {Array} itemArr */
  saveInFile(itemArr) {

    itemArr.reverse();

    itemArr.forEach(item => {
      try {
        let obj = new FileObjData();
        obj.value = item;
        obj.НомерПроекта = item.key;
        obj.ТаблицаПроекта = item.Таблица;
        getMrClassFile().update(obj);
      } catch (err) {
        mrErrToString(err);
      }
    });


  }


  /** @param {DataTaskJoinProgectInfo} task*/
  taskSyncОтчетыПроектовИзСводнойвРеестр(task) {
    Logger.log(`vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv-=taskSyncОтчетыПроектовИзСводнойвРеестр=-vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv`);

    Logger.log(`${JSON.stringify(task)} `);

    let theContext = {
      from: new MrContext(task.param.urls.from, getSettings()),
      to: new MrContext(task.param.urls.to, getSettings()),

    }

    let theSheetModel = {
      from: new MrClassSheetModel(task.param.sheetNames.from, theContext.from),
      to: new MrClassSheetModel(task.param.sheetNames.to, theContext.to),
    }

    if ((new Date().getTime() - new Date(theSheetModel.to.getMem(task.param.mem)).getTime() - task.актуально_дней * DeyMilliseconds * 0.25) < 0
    ) { Logger.log("да ещё актуальны"); return; }

    let theMapFrom = theSheetModel.from.getMap();

    let theItems = [...theMapFrom.values()].filter(item => {
      if (task.param.prefix) {
        if (!`${item[task.param.fild]}`.includes(task.param.prefix)) { return false; }
        return true;
      }
      return true;
    }).map(item => {
      if (task.param.prefix) {
        item[task.param.fild] = `${item[task.param.fild]}`.replace(task.param.prefix, "");
      }
      return item;
    })
    Logger.log({ l: theItems.length });
    // return;
    theSheetModel.to.updateItems(theItems);
    theSheetModel.to.setMem(task.param.mem, new Date());



    Logger.log(`^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^-=taskSyncОтчетыПроектовИзСводнойвРеестр=-^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^`);
  }



  /** @param {DataTaskJoinProgectInfo} task*/
  taskSyncОтчетыПроектовСинхронизцияШапки(task) {
    Logger.log(`vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv-=taskSyncОтчетыПроектовСинхронизцияШапки=-vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv`);

    Logger.log(`${JSON.stringify(task)} `);

    let theContext = {
      from: new MrContext(task.param.urls.from, getSettings()),
      tos: task.param.urls.tos
        .map(to => {
          return new MrContext(to, getSettings());
        }),

    }

    let theSheetModel = {
      from: new MrClassSheetModel(task.param.sheetNames.from, theContext.from),
      /** @type {[MrClassSheetModel]} */
      tos: theContext.tos
        .map(context => {
          return new MrClassSheetModel(task.param.sheetNames.to, context);
        }),
    }

    let theSetHeadKeyFrom = new Set(theSheetModel.from.head_key);
    Logger.log([...theSetHeadKeyFrom.values()]);


    theSheetModel.tos.forEach(theModelTo => {
      let theAddNames = [];
      let theSetHeadKeyTo = new Set(theModelTo.head_key);

      [...theSetHeadKeyFrom.values()].forEach(theHead => {
        if (!theHead) { return false; }
        if (theSetHeadKeyTo.has(theHead)) { return false }
        theAddNames.push(theHead);
        return true;
      });

      Logger.log(theAddNames);
      if (theAddNames.length == 0) { return; }

      theModelTo.requiredColNames(theAddNames);
    });
    Logger.log(`^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^-=taskSyncОтчетыПроектовСинхронизцияШапки=-^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^`);
  }



  tryTask(self, colable_name, item) {
    try {
      return self[colable_name](item);
    } catch (err) {
      item.errors = [].concat(mrErrToString(err), item.errors);
    }
  }







}




