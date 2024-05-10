// Verification of contracts

// VerifContracts

class DataTask_ExeTask_ExportPayments {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_ExportPayments.name;
    /** @constant */
    this.name = ExeTask_ExportPayments.prototype.taskExportProgectListFromReestr.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {

      Ответственный: "",
      sheetNames: {
        // from: "Исполнение",
        // Реестр: "Реестр",
        // Исполнение: "Исполнение",
        // to: "Проекты в Исполнение",
      },

      urls: {
        from: "",
        to: "",
      },

      rows: {
        headFirst: 2,
        bodyFirst: 3,
      },
    };
  }
}


class ExeTask_ExportPayments extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_VerifContracts.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",
    }
    // this.mapValuesFromРеестр = new Map();
    // this.mapValuesFromИсполнение = new Map();
  }

  
  taskExportProgectListFromReestr(task) {
    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModelTo = new MrClassSheetModel(task.param.sheetNames.to, context_to);


    if ((new Date().getTime() - new Date(sheetModelTo.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) { Logger.log("да ещё актуальны"); return; }



    // let sheetName = task.param.sheetName;
    let rows = task.param.rows;

    let rowBodyFirst = rows.bodyFirst;
    let rowheadFirst = rows.headFirst;

    let itemArr = new Array();
    let sheet_from = context_from.getSheetByName(task.param.sheetNames.from);

    let heads = sheet_from.getRange(rowheadFirst, 1, 1, sheet_from.getLastColumn()).getValues()[0];
    heads = ["row"].concat(heads);//.map(h => fl_str(h));
    let vls = sheet_from.getRange(rowBodyFirst, 1, sheet_from.getLastRow() - rowBodyFirst + 1, sheet_from.getLastColumn()).getValues();
    vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });

    let col_Filter = nr("B");

    let col_list = [
      "row",
      "Партнерская УПД",
      "ID",
      "Проект",
      "Название проекта",
      "Продавец",
      "Пок-ль",
      "Статус",
      "Сумма",
      "Сумма оплаты",

    ].map(c => {
      return heads.indexOf(c);
    }).filter(c => (c >= 0));

    // Logger.log(JSON.stringify(heads));
    vls.forEach(v => {
      // Logger.log(JSON.stringify(v));
      if (!v[col_Filter]) { return; }
      let obj = new Object();
      col_list.forEach(col => {
        obj[heads[col]] = v[col];
      });

      obj["key"] = `${obj["ID"]}`;
      obj["row_строка"] = obj["row"];
      obj["row"] = undefined;
      obj["date"] = context_to.timeConstruct;
      itemArr.push(obj);

    });

    // itemArr.forEach(v => Logger.log(JSON.stringify(v)));


    itemArr = itemArr.filter(item => {
      if (item["Сумма оплаты"]) { return true; }
      // if (item["Статус"]) { return true; }
      return false;
    });



    // return 


    // let vlsTo = itemArr.map(item => {
    //   return [item["ID"], item["Сумма"]];
    // });


    if (itemArr.length == 0) { return; }
    // Logger.log(JSON.stringify(itemArr));

    sheetModelTo.setItems(itemArr);
    sheetModelTo.setMem(this.names.обновлено, new Date());

  }


  tryTask(self, colable_name, item) {
    try {
      return self[colable_name](item);
    } catch (err) {
      item.errors = [].concat(mrErrToString(err), item.errors);
    }
  }
}




