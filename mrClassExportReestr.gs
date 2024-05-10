class DataTaskExportProgectListFromReestr {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_ExportReestr.name;
    /** @constant */
    this.name = ExeTask_ExportReestr.prototype.taskExportProgectListFromReestr.name;
    /** @constant */
    this.актуально_дней = 1 / 24;
    this.off = false;
    this.param = {

      Ответственный: "",
      sheetNames: {
        from: "Исполнение",
        Реестр: "Реестр",
        Исполнение: "Исполнение",
        to: "Проекты в Исполнение",
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


class TaskListExportReestr {
  constructor() {
    let reestr_K = new DataTaskExportProgectListFromReestr();
    reestr_K.param.Ответственный = "Наталья";
    reestr_K.param.urls.from = getSettings().urls.РеестрК;
    reestr_K.param.sheetNames.from = "Исполнение";

    reestr_K.param.urls.to = getSettings().urls.ПромежуткиЛогиста;
    reestr_K.param.sheetNames.to = getSettings().sheetNames.Проекты_в_Исполнение;


    let taskGetProjects = new DataTaskGetExportValuesFromProjects();
    taskGetProjects.param.sheetNames.to = getSettings().sheetNames.Проекты_в_Исполнение;
    taskGetProjects.param.urls.to = getSettings().urls.ПромежуткиЛогиста;


    /** @type {[DataTaskExportProgectListFromReestr]} */
    this.list = [
      reestr_K,
      // taskLoadProjects,
      taskGetProjects,
    ];
  }


  getTaskList() {
    return this.list;
  }
}

// TaskListExportReestr.bind

class ExeTask_ExportReestr extends MrTaskExecutorBaze {
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

  /** @param {DataTaskExportProgectListFromReestr} task*/
  taskExportProgectListFromReestr(task) {
    Logger.log(`${JSON.stringify(task)} `);



    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModelProjects = new MrClassSheetModel(task.param.sheetNames.to, context_to);

    // Logger.log("ещё актуальны?");
    // Logger.log(sheetModelProjects.getMem(this.names.обновлено));
    // Logger.log("не актуальны | " + `${new Date().getTime()}  | ${new Date(sheetModelProjects.getMem(this.names.обновлено)).getTime() + task.актуально_дней * DeyMilliseconds} | 
    // ${(new Date().getTime() - new Date(sheetModelProjects.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds)} `);

    if ((new Date().getTime() - new Date(sheetModelProjects.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    ) { Logger.log("да ещё актуальны"); return; }



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
    let col_Filter = nr("A");

    let col_list = [
      "row",
      "№",
      "Заказчик",
      "Название",
      "Таблица",
      "Ответственный",
      "Статус проекта",

    ].concat(heads).map(c => {
      // return c;
      // return heads.indexOf(fl_str(c));
      return heads.indexOf(c);
    }).filter(c => (c >= 0));

    // Logger.log(JSON.stringify(heads));
    vls.forEach(v => {
      if (!v[col_Filter]) { return; }

      let obj = new Object();
      col_list.forEach(col => {
        obj[heads[col]] = v[col];
      });

      // if (fl_str(obj["Ответственный"]) != fl_str(task.param.Ответственный)) { return; }

      obj["key"] = `${obj["№"]}`;
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
        if (!ok_url.includes("Я")) { url = ok_url; }
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

        Р: this.getExportValues(context_from, item),
        // Р: this.getExportValuesFromFile(context_from, item.key, ["Реестр", "Исполнение"]),
        // Тб: this.getExportValuesFromFile(context_from, item.key, ["Тб",]),
      }

      return ret;
    });


    sheetModelProjects.updateItems(itemArr);

    let keyOkArr = itemArr.map(item => item.key);
    let delRow = [...sheetModelProjects.getMap().values()].filter((obj) => {
      if (keyOkArr.includes(`${obj.key}`)) { return false; }
      return true
    }).map(obj => obj.row).sort((a, b) => {
      if (a > b) { return -1; }
      if (a < b) { return 1; }
      return 0;
    });

    delRow.forEach(r => sheetModelProjects.sheet.deleteRow(r));


    sheetModelProjects.optimize();
    sheetModelProjects.setMem(this.names.обновлено, new Date());
  }

  getValuesFromSheetРеестр(context, item) {
    if (this.mapValuesFromРеестр.size == 0) {
      let context_from = context;

      let rowBodyFirst = 2;
      let rowheadFirst = 1;

      let sheet_from = context_from.getSheetByName("Реестр"); // Реестр

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
        "род",
        "Тип",
        "ЭТП",
        "Ссылка",
        "Таблица",
        "Контрольная дата",
        "Последние Изм-я",
        "Статус",
        "Дата след контакта",
        "Наши участники",
        "Наши цены",
        "не используется",
        "Коментарий А",
        "Кооментарий К",
        "Состояние расчета",
        "НМЦ",
        "Закупка",
        "НМЦ / Закуп",
        "Срок поставки ТЗ",
        "Реальный срок",
        "Срок оплаты",
        "Групп",
        "Обесп-ие заявки",
        "Обесп-ие дог-ра",
        "Участники",
        "Коэф на выход",
        "Цена выхода",
        "Победил",
        "Цена побед-ля",
        "Коэф Побед-ля",
        "Результаты",
        "Конкуренция",
        "Допущено",
        "Отклонено",
        "Дата доб-ния",
        "Дата разм-ния",
        "Ссылка на папку с файлами",
        "Комментарий А",
        "Комментарий К",
        // "Строка не найдена в Дог Уч",
        // "Обьединенное название",
        // "JSON_OBSERVER",
      ].concat(heads).map(c => {
        // return c;
        // return heads.indexOf(fl_str(c));
        return heads.indexOf(c);
      }).filter(c => (c >= 0));

      // Logger.log(JSON.stringify(heads));
      vls.forEach(v => {
        if (!v[col_Filter]) { return; }
        let obj = new Object();
        col_list.forEach(col => {
          obj[heads[col]] = v[col];
        });
        obj["key"] = `${obj["№"]}`;
        obj["row_строка"] = obj["row"];
        obj["row"] = undefined;

        this.mapValuesFromРеестр.set(obj["key"], obj);

      });

    }


    return this.mapValuesFromРеестр.get(item.key);
  }

  getValuesFromSheetИсполнение(context, item) {
    if (this.mapValuesFromИсполнение.size == 0) {

      // let rows = task.param.rows;

      // let rowBodyFirst = rows.bodyFirst;
      // let rowheadFirst = rows.headFirst;

      let rowheadFirst = 2;
      let rowBodyFirst = 3;


      let sheet_from = context.getSheetByName("Исполнение");

      let heads = sheet_from.getRange(rowheadFirst, 1, 1, sheet_from.getLastColumn()).getValues()[0];
      heads = ["row"].concat(heads);//.map(h => fl_str(h));
      let vls = sheet_from.getRange(rowBodyFirst, 1, sheet_from.getLastRow() - rowBodyFirst + 1, sheet_from.getLastColumn()).getValues();
      vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
      let col_Filter = nr("A");

      let col_list = heads.map(c => {
        // return c;
        // return heads.indexOf(fl_str(c));
        return heads.indexOf(c);
      }).filter(c => (c >= 0));


      // Logger.log(JSON.stringify(heads));
      vls.forEach(v => {
        if (!v[col_Filter]) { return; }
        let obj = new Object();
        col_list.forEach(col => {
          obj[heads[col]] = v[col];
        });
        obj["key"] = `${obj["№"]}`
        obj["row_строка"] = obj["row"];
        obj["row"] = undefined;

        this.mapValuesFromИсполнение.set(obj["key"], obj);
        // Logger.log(JSON.stringify(obj));
      });

    }
    return this.mapValuesFromИсполнение.get(item.key);
  }

  getExportValues(context, item) {

    item.name = context.getSpreadSheet().getName();
    item.Реестр = this.getValuesFromSheetРеестр(context, item);
    item.Исполнение = this.getValuesFromSheetИсполнение(context, item);
    return item;
  }



  tryTask(self, colable_name, item) {
    try {
      return self[colable_name](item);
    } catch (err) {
      item.errors = [].concat(mrErrToString(err), item.errors);
    }
  }
}




