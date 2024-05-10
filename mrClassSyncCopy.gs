

class DataTaskSyncCopyRange {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncCopyRange.name;
    /** @constant */
    this.name = ExeTask_SyncCopyRange.prototype.taskExportSyncCopyRange.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      urls: {
        from: "",
        to: "",
      },

      sheetNames: {
        from: "",
        to: "",
      },

      range: {
        from: "",
        to: "",
      },

      filter_empty: -1,
    };


    this.note = false;
  }
}



class ExeTask_SyncCopyRange extends MrTaskExecutorBaze {
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

  /** @param {DataTaskSyncCopyRange} task */
  taskExportSyncCopyRange(task) {

    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, new Object());
    let context_to = new MrContext(task.param.urls.to, new Object());

    let sheet_from = context_from.getSheetByName(task.param.sheetNames.from);
    let sheet_to = context_to.getSheetByName(task.param.sheetNames.to);

    let vls_from = sheet_from.getRange(task.param.range.from).getValues();

    if (task.param.filter_empty != -1) {
      vls_from = vls_from.filter(v => `${v[task.param.filter_empty]}`);
    }


    let cell_left_top = sheet_to.getRange(task.param.range.to).getCell(1, 1);
    let col_start = cell_left_top.getColumn();
    let row_start = cell_left_top.getRow();

    if (vls_from.length) {
      try {
        // sheet_to.getRange(row_start, col_start, sheet_to.getLastRow() - row_start + 1, vls_from[0].length).clearContent();
        sheet_to.getRange(task.param.range.to).clearContent();
      } catch (err) { }
      sheet_to.getRange(row_start, col_start, vls_from.length, vls_from[0].length).setValues(vls_from);
      if (task.note) { sheet_to.getRange(row_start, col_start).setNote(JSON.stringify({ date: new Date, name: "Задача: Cинхронизация таблиц", task })); }
    }

  }

}




class DataTaskUpdateFilds {
  constructor() {



    /** @constant */
    this.exe_class = ExeTask_SyncFilds.name;
    /** @constant */
    this.name = ExeTask_SyncFilds.prototype.taskExportUpdateFilds.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;


    this.strategies = {
      overwrite: true,
    };



    this.off = false;
    this.param = {
      urls: {
        from: "",
        to: "",
      },

      sheetNames: {
        from: "",
        to: "",
      },

      strategy: "",

      fildCouples: {
        key: { from: "", to: "" },
      },

      syncFilds: [],

      rowConf: {
        from: null,
        to: null,
      },
    };
    this.note = false;
  }
}




class ExeTask_SyncFilds extends MrTaskExecutorBaze {
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

  /** @param {DataTaskSyncCopyRange} task */
  taskExportSyncCopyRange(task) {

    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, new Object());
    let context_to = new MrContext(task.param.urls.to, new Object());

    let sheet_from = context_from.getSheetByName(task.param.sheetNames.from);
    let sheet_to = context_to.getSheetByName(task.param.sheetNames.to);

    let vls_from = sheet_from.getRange(task.param.range.from).getValues();

    if (task.param.filter_empty != -1) {
      vls_from = vls_from.filter(v => `${v[task.param.filter_empty]}`);
    }


    let cell_left_top = sheet_to.getRange(task.param.range.to).getCell(1, 1);
    let col_start = cell_left_top.getColumn();
    let row_start = cell_left_top.getRow();

    if (vls_from.length) {
      try {
        // sheet_to.getRange(row_start, col_start, sheet_to.getLastRow() - row_start + 1, vls_from[0].length).clearContent();
        sheet_to.getRange(task.param.range.to).clearContent();
      } catch (err) { }
      sheet_to.getRange(row_start, col_start, vls_from.length, vls_from[0].length).setValues(vls_from);
      if (task.note) { sheet_to.getRange(row_start, col_start).setNote(JSON.stringify({ date: new Date, name: "Задача: Cинхронизация таблиц", task })); }
    }

  }


  /** @param {DataTaskUpdateFilds} task */
  taskExportUpdateFilds(task) {
    Logger.log(`${JSON.stringify(task, null, 2)} `);

    let theContext = {
      from: new MrContext(task.param.urls.from, new Object()),
      to: new MrContext(task.param.urls.to, new Object()),
    }

    let theSheetModel = {
      from: new MrClassSheetModel(task.param.sheetNames.from, theContext.from, task.param.rowConf.from),
      to: new MrClassSheetModel(task.param.sheetNames.to, theContext.to, task.param.rowConf.to),
    };

    let couplesMap = {
      from: new Map(),
      to: new Map(),
    };

    for (let f in task.param.fildCouples) {
      couplesMap.from.set(task.param.fildCouples[f].from, f);
      couplesMap.to.set(task.param.fildCouples[f].to, f);
    }

    theSheetModel.from.head_key = theSheetModel.from.head_key.map(head => {
      return (couplesMap.from.has(head) ? couplesMap.from.get(head) : head);
    });

    theSheetModel.to.head_key = theSheetModel.to.head_key.map(head => {
      return (couplesMap.to.has(head) ? couplesMap.to.get(head) : head);
    });


    let theMap = {
      from: theSheetModel.from.getMap(),
      to: theSheetModel.to.getMap(),
    };


    // [...theMap.from.values()].slice(-3).forEach(item => {
    //   Logger.log(JSON.stringify(item, null, 2));
    // });


    // [...theMap.to.values()].filter(item => {
    //   return item.key;
    // }).slice(-3).forEach(item => {
    //   Logger.log(JSON.stringify(item, null, 2));
    // });

    let colMap = new Map();

    task.param.syncFilds.forEach(fildName => {
      if (theSheetModel.to.head_key.indexOf(fildName) >= 0) {
        colMap.set(fildName, theSheetModel.to.head_key.indexOf(fildName));
      }
    });

    // colMap.forEach((v, k) => {
    //   Logger.log(JSON.stringify({ v, k }, null, 2));
    // });
    // colMap.set("notExist", 199);

    let updateItems = [...theMap.to.values()].filter(item => { return item.key; });

    updateItems = updateItems.map(item => {

      let ret = { key: item.key, row: item.row, };
      let itemFrom = theMap.from.get(ret.key);
      if (!itemFrom) {
        // ret["Присутствует в таблице проекта"] = false;
      }
      else {
        task.param.syncFilds.forEach(fildName => {
          if (item[fildName] !== itemFrom[fildName]) {
            ret[fildName] = itemFrom[fildName];
            if (ret[fildName] === undefined) {
              ret[fildName] = "";
            }
          }
        });
      }
      return ret;
    });

    updateItems
      // .slice(-5)
      .reverse()
      .forEach(item => {
        // Logger.log(JSON.stringify({ item }, null, 2));
        colMap.forEach((col, fildName) => {
          // Logger.log({ col, fildName, ifi: item[fildName] });
          if (fildName in item) {
            // Logger.log(true);
            if (typeof item[fildName] == "object") {
              item[fildName] = Utilities.formatDate(item[fildName], "Europe/Moscow", "dd.MM.yyyy HH:mm:ss");
            }
            theSheetModel.to.sheet.getRange(item.row, col).setValue( item[fildName]);
          } else {
            // Logger.log(false);
          }
        });
      });




  }


}





