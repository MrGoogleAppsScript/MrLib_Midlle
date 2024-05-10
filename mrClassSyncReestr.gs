

class DataTaskSyncReestrПроцентОплаты {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncReestr.name;
    /** @constant */
    this.name = ExeTask_SyncReestr.prototype.taskExportSyncПроцентОплаты.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {

      // sheetName_from: "",
      sheetNames: {
        from: "Проекты в Исполнение",
        to: "Процент оплаты",
      },

      urls: {
        from: "",
        to: "",
      },

    };
  }
}



class ExeTask_SyncReestr extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_SyncLogist.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",
    }
    this.mapValuesFromРеестр = new Map();
    this.mapValuesFromИсполнение = new Map();
  }


  taskExportSyncПроцентОплаты(task) {

    let f_key_i = ((key, n) => { let nnn = `000${n}`.slice(-3); return `${key} № ${nnn}` });

    Logger.log(`${JSON.stringify(task)} `);
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModel_РабочийЛист = new MrClassSheetModel(task.param.sheetNames.to, context_to);
    if ((new Date().getTime() - new Date(sheetModel_РабочийЛист.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
      Logger.log("да ещё актуальны"); return;
    }

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let sheetModel_ПроектыВРаботе = new MrClassSheetModel(task.param.sheetNames.from, context_from);

    let mapПроектыВРаботе = sheetModel_ПроектыВРаботе.getMap();
    // let mapРабочийЛист = sheetModel_РабочийЛист.getMap();

    let arr = [...mapПроектыВРаботе.values()].filter((obj) => {
      return true;
      return obj["Рабочий Лист Логиста"] === true
    });

    // Logger.log(`arr=${arr.length} |       // key= |${arr.map(v => v["key"])}      `);

    let date = new Date();
    let itemArr = new Array();
    arr.forEach(project => {

      // itemArr.push({
      //   key: f_key_i(project.key, 0),
      //   r: 0,
      //   Проект: project.key,
      //   Р: (() => { try { return JSON.parse(project.Р); } catch { return new Object(); } })(),
      //   date: date,
      // });

      let productMap = new Map();
      /** @type {Array} */
      let Тб_Выбор_Поставщиков_тело = (() => { try { return JSON.parse(project.Тб).Выбор_Поставщиков.тело; } catch { return new Array() } })();
      // Logger.log(`arr=${Тб_Выбор_Поставщиков_тело.length} |      // key= |${Тб_Выбор_Поставщиков_тело.map(v => [v["№"], v["НОМЕНКЛАТУРА"], "\n"])}      `);
      Тб_Выбор_Поставщиков_тело.forEach(pr => {
        // let key_i = `${project.key}№${pr["№"]}`;
        let key_i = f_key_i(project.key, pr["№"]);
        let item = new Object();

        item.Проект = project.key;
        item.Р = (() => { try { return JSON.parse(project.Р); } catch { return new Object(); } })();
        item["key"] = key_i;
        item["r"] = pr["№"];
        item["Тб"] = new Object();
        item["Тб"]["Выбор_Поставщиков"] = pr;
        item["Проект"] = project.key;
        item.date = date;
        productMap.set(item["key"], item);
      });

      // Тб.Оплаты.тело
      let Тб_Оплаты_тело = (() => { try { return JSON.parse(project.Тб).Оплаты.тело; } catch { return new Array() } })();
      let Есть_Оплаты_тело = (() => { try { return JSON.parse(project.Тб).Оплаты; } catch { return undefined } })();

      Тб_Оплаты_тело.forEach(pr => {
        let key_i = f_key_i(project.key, pr["№"]);
        let item = productMap.get(key_i);
        if (!item) { return; }
        item["Тб"]["Оплата"] = pr;
      });

      let productArr = [...productMap.values()]
        .filter(item => {
          return (() => {
            try {
              // if (Тб_Оплаты_тело.length == 0) { return false; }
              if (!Есть_Оплаты_тело) { return false; }
              if (!item["Тб"]["Оплата"]) { return true; }

              return (item["Тб"]["Оплата"]["% оплаты"] != 1)
            } catch (err) { return true; }
          })();
        })
        .sort((a, b) => {
          if (Number.parseInt(a["r"]) > Number.parseInt(b["r"])) { return 1; }
          if (Number.parseInt(a["r"]) < Number.parseInt(b["r"])) { return -1; }
          return 0;
        });

      itemArr = itemArr.concat(productArr);

      // itemArr.forEach(it => {         Logger.log(`${it.key} | ${JSON.stringify(it)}`);      });
    });

    // sheetModel_РабочийЛист.updateItems(itemArr);
    sheetModel_РабочийЛист.setItems(itemArr);
    sheetModel_РабочийЛист.optimize();
    sheetModel_РабочийЛист.setMem(this.names.обновлено, new Date());
  }



  /** @param {DataTaskSyncCopyИсполнениеЮ} task */
  taskSyncCopyИсполнениеЮ(task) {


    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, new Object());
    let context_to = new MrContext(task.param.urls.to, new Object());

    let sheet_from = context_from.getSheetByName(task.param.sheetNames.from);
    let sheet_to = context_to.getSheetByName(task.param.sheetNames.to);

    let vls_from = sheet_from.getRange(task.param.range.from).getValues();

    if (task.param.filter_empty != -1) {
      vls_from = vls_from.filter(v => `${v[task.param.filter_empty]}`);
    }


    let heads_from = sheet_from.getRange(task.param.range.heads).getValues()[0];

    vls_from = vls_from.map(arr => {
      return task.param.changeProgectID.function(heads_from, arr, task.param.changeProgectID.fild, task.param.changeProgectID.prefix);
    });



    let mapSinonim = new Map();
    getMapDictionarys().forEach(value => {
      let vfrom = value[task.param.changeSinonim.from];
      let vto = value[task.param.changeSinonim.to];

      if (!vfrom) { return; }
      if (!vto) { return; }
      if (vfrom == vto) { return; }
      //  Logger.log({ vfrom, vto });
      mapSinonim.set(vfrom, vto);
    });

    // mapSinonim.forEach((v, k) => Logger.log({ v, k }));


    let colInd = task.param.changeSinonim.filds.map(fild => {
      return heads_from.indexOf(fild);
    }).filter(ind => { return ind >= 0 });


    vls_from = vls_from.map(arr => {

      colInd.forEach(ind => {
        let v = arr[ind];
        let s = mapSinonim.get(v);
        if (!`${s}`) { return; }
        arr[ind] = s;
      });
      return arr;
    });


    // vls_from.forEach(v => Logger.log(v));
    // return

    // vls_from = [].concat([heads_from], vls_from);

    // vls_from = [].concat([heads_from], vls_from);

    let col_list = [].concat(heads_from).map(c => {
      if (!`${c}`) { return -1; }
      return heads_from.indexOf(c);
    }).filter(c => (c >= 0));


    let items = vls_from.map(v => {


      let obj = new Object();
      col_list.forEach(col => {
        obj[heads_from[col]] = v[col];
      });

      // if (fl_str(obj["Ответственный"]) != fl_str(task.param.Ответственный)) { return; }

      obj["key"] = `${obj["№"]}`;
      // obj["row_строка"] = obj["row"];
      // obj["row"] = undefined;

      // this.mapValuesFromРеестр.set(obj["key"], obj);
      // let ph = [{ key: `${obj["key"]}` }, obj];


      return obj;

    });

    // items.forEach(v => Logger.log(JSON.stringify(v, null, 2)));
    items = items.filter(item => {
      // Logger.log(JSON.stringify(v, null, 2))
      // Logger.log(JSON.stringify({
      //   item,
      //   "Статус проекта": fl_str(item["Статус проекта"]),
      //   "ЗАВЕРШЕН": fl_str(item["Статус проекта"]) == fl_str("ЗАВЕРШЕН"),
      //   "ERROR": fl_str(JSON.stringify(item)).includes(fl_str("#ERROR!")),
      // }, null, 2));

      if (fl_str(item["Статус проекта"]) == fl_str("ЗАВЕРШЕН")) { Logger.log([item["key"], item["Статус проекта"]]); return false; }  // Статус проекта ЗАВЕРШЕН

      for (let k in item) {
        if (item[k] == "#ОШИБКА!" || item[k] == "#ERROR!") {
          item[k] = undefined;
        }
      }


      // if (fl_str(JSON.stringify(item)).includes(fl_str("#ОШИБКА!"))) { Logger.log([item["key"],"#ОШИБКА!"]); return false; }
      // if (fl_str(JSON.stringify(item)).includes(fl_str("#ERROR!"))) { Logger.log([item["key"],"#ERROR!"]); return false; }


      Logger.log([item["key"], item["Статус проекта"]]); return true;
    });

    Logger.log(`items = ${items.length}`);
    // items.forEach(v => Logger.log(JSON.stringify(v, null, 2)));


    let sheetModelTo = new MrClassSheetModel(task.param.sheetNames.to, context_to, task.param.rowConf.to);
    sheetModelTo.updateItemsObj(items);
    sheetModelTo.setMem(task.param.обновлено, new Date());

    return;






    let cell_left_top = sheet_to.getRange(task.param.range.to).getCell(1, 1);
    let col_start = cell_left_top.getColumn();
    let row_start = cell_left_top.getRow();

    if (vls_from.length) {
      try {
        // sheet_to.getRange(row_start, col_start, sheet_to.getLastRow() - row_start + 1, vls_from[0].length).clearContent();
        sheet_to.getRange(task.param.range.to).clearContent();
      } catch (err) { }
      sheet_to.getRange(row_start, col_start, vls_from.length, vls_from[0].length).setValues(vls_from);
      // sheet_to.getRange(row_start, col_start).setNote(JSON.stringify({ date: new Date, name: "Задача: Cинхронизация таблиц", task }));
    }






  }



}


class DataTaskSyncCopyИсполнениеЮ {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncReestr.name;
    /** @constant */
    this.name = ExeTask_SyncReestr.prototype.taskSyncCopyИсполнениеЮ.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      обновлено: "обновлено",
      urls: {
        from: getSettings().urls.РеестрЮ,
        to: getSettings().urls.РеестрК,

      },

      sheetNames: {
        from: "Исполнение",
        to: "Исполнение Ю",
      },

      range: {
        heads: "A2:AU2",
        from: "A3:AU",
        to: "A10:AU",
      },
      rowConf: {
        from: {
          head: { first: 1, last: 1, key: 1, },
          //    head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
          body: { first: 2, last: 2, },
        },
        to: {
          head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
          body: { first: 11, last: 11, },
        }
      },

      filter_empty: 0,

      changeProgectID: {
        fild: `№`, prefix: "Ю-",
        function: ((heads, arr, fild, prefix) => {
          let col = heads.indexOf(fild);
          if (`${arr[col]}`.includes(prefix)) { return arr; }
          arr[col] = `${prefix}${arr[col]}`;
          return arr;
        }),
      },

      changeSinonim: {
        url: getSettings().urls.РеестрЮ,
        sheetName: "Технический",
        range: "A1:AK",
        from: "Полное название",
        to: "Дубль названия",
        filds: [`Победил`, `Закупает`],
      },

    };
  }
}






function text_triggerИсполнениеЮ(trigerrInfo = "texttriggerПродолжитьВыполнениеЗадачь") {
  Logger.log(JSON.stringify(trigerrInfo));
  triggerLookCache();
  // return;
  let tasks = [].concat(
    [

      (() => {
        let syncCopy = new DataTaskSyncCopyИсполнениеЮ();
        return syncCopy;
      })(),

    ]
  );

  let mrTaskExe = new MrTaskExecutor(getContext(), tasks, getContext().getScriptCache());
  try {
    mrTaskExe.triggerПродолжитьВыполнениеЗадачь();
  } catch (err) { mrErrToString(err); }

  triggerLookCache();
}





