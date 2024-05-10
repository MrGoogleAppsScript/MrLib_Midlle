

class DataTaskSyncРабочийЛист {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncLogist.name;
    /** @constant */
    this.name = ExeTask_SyncLogist.prototype.taskExportSyncРабочийЛист.name;
    /** @constant */
    this.актуально_дней = 1 / 24 * 2 / 1;
    this.off = false;
    this.param = {

      // sheetName_from: "",
      sheetNames: {
        from: "Проекты в Исполнение",
        to: "Рабочий Лист",
        a: "Проекты в работе",
      },

      urls: {
        from: "",
        to: "",
      },
      // // sheetName_to: "",
      // rows: {
      //   headFirst: 1,
      //   bodyFirst: 2,
      // },
    };
  }
}



class DataTaskРабочийЛистОбработка {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncLogist.name;
    /** @constant */
    this.name = ExeTask_SyncLogist.prototype.taskExportРабочийЛистОбработка.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 1;
    this.off = false;
    this.param = {

      // sheetName_from: "",
      sheetNames: {
        // from: "Проекты в Исполнение",
        to: "Рабочий Лист",
        // a: "Проекты в работе",
      },

      urls: {
        // from: "",
        to: "",
      },
      // // sheetName_to: "",
      // rows: {
      //   headFirst: 1,
      //   bodyFirst: 2,
      // },
    };
  }
}



class DataTaskSyncПроектыВРаботе {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncLogist.name;
    /** @constant */
    this.name = ExeTask_SyncLogist.prototype.taskExportSyncПроектыВРаботе.name;
    /** @constant */
    this.актуально_дней = 1 / 24;
    this.off = false;
    this.param = {

      sheetNames: {
        from: "Проекты в Исполнение",
        to: "Проекты в работе",
      },

      urls: {
        from: "",
        to: "",
      },

      cols: {
        key: nr("A"),
      }

    };
  }
}



class DataTaskSyncУПД_ДОК {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncLogist.name;
    /** @constant */
    this.name = ExeTask_SyncLogist.prototype.taskExportSyncУПД_ДОК.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 1;
    this.off = false;
    this.param = {

      sheetNames: {
        from: "Проекты в Исполнение",
        to_УПД: "Отгрузочные документы (копия)",
        to_ДОК: "Договора (копия) 1",
        to_Рабочий_Лист: "Рабочий Лист (копия)",
      },

      urls: {
        from: getSettings().urls.ПромежуткиЛогиста,
        to: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
      },

      // cols: {
      //   key: nr("A"),
      // }
      rows: {
        body: { first: 2, },
      }

    };
  }
}

class DataTaskClearSelectRowInУПД_ДОК {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncLogist.name;
    /** @constant */
    this.name = ExeTask_SyncLogist.prototype.taskClearSelectRowInУПД_ДОК.name;
    /** @constant */
    this.актуально_дней = -1;
    this.off = false;
    this.param = {

      sheetNames: {
        from: "Проекты в Исполнение",
        to_УПД: "Отгрузочные документы",
        to_ДОК: "Договора",
        mem_Рабочий_Лист: "Рабочий Лист",
      },

      urls: {
        to: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
      },

      rows: {
        body: { first: 2, },
      },
      rowConf: {
        to: {
          head: { first: 1, last: 1, key: 1, },
          body: { first: 2, last: 2, },
        },

      }

    };
  }
}


class DataTaskCopyРабочийЛистToЛогистика {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncLogist.name;
    /** @constant */
    this.name = ExeTask_SyncLogist.prototype.taskCopyРабочийЛистToЛогистика.name;
    /** @constant */
    this.актуально_дней = 1 / 24 * 1;
    this.off = false;
    this.param = {


      simplifyHeader: [
        "Проверка обязательного заполнения полей",
        "Результат проверки",
        "Автоматическая Проверка",
      ],


      sheetNames: {
        from: "Рабочий Лист",
        to: "7-2 Номенклатуры",
      },

      urls: {
        from: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
        to: getSettings().urls.Письма,
      },

      // theSheetModel.from.row.head.ru_name,
      row: {
        from: {
          head: {
            first: 1,
            last: 10,
            key: 9,
            type: 8,
            ru_name: 10,
            changed: 7,
            off: 10,
            mem: 7,
          },

          body: {
            first: 11,
            last: 11,
          }
        },
        to: {
          head: {
            first: 1,
            last: 10,
            key: 9,
            type: 8,
            ru_name: 10,
            changed: 7,
            off: 10,
            mem: 7,
          },

          body: {
            first: 11,
            last: 11,
          }
        }

      }

    };
  }
}




class TaskListSyncLogist {
  constructor() {

    let sheetNames = {
      Проекты_в_работе: "Проекты в работе",
      Импорт_Антон: "Импорт Антон",
      Процент_оплаты: "Процент оплаты",
      Договора: "Договора",
      Технический_лист: "Технический лист",
      Проверка_Документов: "Проверка Документов",
      Исполнение: "Исполнение",
      Оплаты: "Оплаты",
      Сведения_об_оплате_УПД_из_Реестр_К: "Сведения об оплате УПД из Реестр К",
      ОтгрузочныеДокументы: "Отгрузочные документы",

      Логистика_7_2: "7-2 Логистика",
      Номенклатуры_7_2: "7-2 Номенклатуры",
      УПД_7_2: "7-2 УПД",
      Договора_7_2: "7-2 Договора",
      Оплаты_7_2: "7-2 Оплаты",
      Оплаты: "Оплаты",
    }


    let urls = {
      ОтслеживаниеПоставкиАнтонЛизякин: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
      ПромежуткиРеестр: getSettings().urls.ПромежуткиРеестр,
      РаскладК: getSettings().urls.РеестрК,

      Промежуточный_Логистики: getSettings().urls.ПромежуткиЛогиста,
    };

    this.list = [
      // syncLogist,

      (() => {
        let syncРабочийЛист = new DataTaskSyncРабочийЛист();
        syncРабочийЛист.param.urls.from = getSettings().urls.ПромежуткиЛогиста;
        syncРабочийЛист.param.urls.to = getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин;
        return syncРабочийЛист;
      })(),


      (() => {
        let taskРабочийЛистОбработка = new DataTaskРабочийЛистОбработка();
        taskРабочийЛистОбработка.param.urls.to = getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин;
        return taskРабочийЛистОбработка;
      })(),



      (() => {
        let syncРабочийЛист = new DataTaskSyncПроектыВРаботе();
        syncРабочийЛист.param.urls.from = getSettings().urls.ПромежуткиЛогиста;
        syncРабочийЛист.param.urls.to = getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин;
        syncРабочийЛист.param.sheetNames.to = "Проекты в работе";
        return syncРабочийЛист;
      })(),


      (() => {
        let syncУПД_ДОК = new DataTaskSyncУПД_ДОК();
        syncУПД_ДОК.param.urls = {
          from: getSettings().urls.ПромежуткиЛогиста,
          to: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
        };
        syncУПД_ДОК.param.sheetNames = {
          from: "Проекты в Исполнение",
          to_УПД: "Отгрузочные документы",
          to_ДОК: "Договора",
          to_Рабочий_Лист: "Рабочий Лист",
        }
        return syncУПД_ДОК;
      })(),



      (() => {
        let task = new DataTask_ExeTask_ExportPayments();
        task.param.urls.from = urls.РаскладК;
        task.param.urls.to = urls.ОтслеживаниеПоставкиАнтонЛизякин;
        task.param.sheetNames.from = sheetNames.Оплаты;
        task.param.sheetNames.to = sheetNames.Сведения_об_оплате_УПД_из_Реестр_К;
        return task;
      })(),

      (() => {
        let task = new DataTaskCopyРабочийЛистToЛогистика();
        return task;
      })(),


      (() => {
        let task = new DataTaskCopyРабочийЛистToЛогистика();
        task.param.sheetNames.from = sheetNames.Договора;
        task.param.sheetNames.to = sheetNames.Договора_7_2;
        task.param.row.from = {
          head: {
            first: 1,
            last: 1,
            key: 1,
            ru_name: 1,
          },
          body: {
            first: 2,
            last: 2,
          }
        };
        return task;
      })(),



      (() => {
        let task = new DataTaskCopyРабочийЛистToЛогистика();
        task.param.sheetNames.from = sheetNames.ОтгрузочныеДокументы;
        task.param.sheetNames.to = sheetNames.УПД_7_2;
        task.param.row.from = {
          head: {
            first: 1,
            last: 1,
            key: 1,
            ru_name: 1,
          },
          body: {
            first: 2,
            last: 2,
          }

        };
        return task;
      })(),





      (() => {
        let task = new DataTaskCopyРабочийЛистToЛогистика();

        task.param.urls.from = getSettings().urls.РеестрК;
        task.param.urls.to = getSettings().urls.Письма;


        task.param.sheetNames.from = sheetNames.Оплаты;
        task.param.sheetNames.to = sheetNames.Оплаты_7_2;
        task.param.row.from = {
          head: {
            first: 1,
            last: 2,
            key: 2,
            ru_name: 2,
          },
          body: {
            first: 3,
            last: 3,
          }

        };
        return task;
      })(),








    ];
  }

  getTaskList() {
    return this.list;
  }
}


class ExeTask_SyncLogist extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_SyncLogist.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",
      обработкаВыполнена: "обработкаВыполнена",
      обновлено_УПД_ДОК: "обновлено_УПД_ДОК",
    }
    this.mapValuesFromРеестр = new Map();
    this.mapValuesFromИсполнение = new Map();


    this.СтатусыДоплаты = {
      НеТребуется: "Не требуется",
      Доплачено: "Доплачено",
      Ждем: "Ждем",
    };

    this.col_CheckName = "для проверки расхождения данных с логистом";
    this.col_Name_Has = "Присутствует в таблице проекта";

    this.cols_update_УПД = [
      // this.col_CheckName,
      "№ документа",
      "Дата документа",
      "Сумма документа",
      "№ УПД о закупке",
      "Тип",
      // "Ошибки при заполнение УПД во внешней",
    ];


    this.cols_update_ДОК = [
      // this.col_CheckName,
      "№ Договора о закупке",
      "Юр. лицо Продавец",
      "Юр. лицо Покупатель",
      "Номер Договора",
      "Дата подписи продавцом",
      "Сумма договора",
      "Дата подписи Покупателем",
      "Частичная оплата за товар",
      "Частичная поставка товара",
      "ТипДоговора",
      "Дата договора от",
      "Вид подписания договора",
      // "Ошибки при заполнение договора во внешней",
    ];


    // Юр. лицо Продавец	Юр. лицо Покупатель	Номер Договора



  }


  /** @param {DataTaskSyncПроектыВРаботе} task*/
  taskExportSyncПроектыВРаботе(task) {

    Logger.log(`${JSON.stringify(task)} `);
    try {


      // Logger.log(`${JSON.stringify(task)} `);
      let context_to = new MrContext(task.param.urls.to, getSettings());
      let sheet = context_to.getSheetByName(task.param.sheetNames.to);

      let sheetModel_ПроектыВРаботе = new MrClassSheetModel(task.param.sheetNames.to, context_to);
      if ((new Date().getTime() - new Date(sheetModel_ПроектыВРаботе.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
        Logger.log("да ещё актуальны"); return;
      }


      let range_str = `${nc(task.param.cols.key)}11:${nc(task.param.cols.key)}`;
      let vls = sheet.getRange(range_str).getValues();
      vls = vls.filter(v => v[0]).map(v => v[0]);
      // Logger.log(`vls = ${vls}`);

      let context_from = new MrContext(task.param.urls.from, getSettings());
      // let sheetModel_ПроектыВИсполнение = new MrClassSheetModel(task.param.sheetNames.from, context_from);
      let sheetModel_ПроектыВИсполнение = new MrClassSheetModelПроектыВИсполнение(task.param.sheetNames.from, context_from);

      let mapПроектыВИсполнение = sheetModel_ПроектыВИсполнение.getMap();
      let projectKeysMemSet = new Set(vls);
      let projectKeys = [...mapПроектыВИсполнение.keys()];

      let projectKeysAdd = projectKeys.filter(k => !projectKeysMemSet.has(k));

      Logger.log(` ${JSON.stringify({ vls, projectKeysMemSet: [...projectKeysMemSet.values()], projectKeysAdd: projectKeysAdd })}`);

      if (projectKeysAdd.length != 0) {
        vls = sheet.getRange(range_str).getValues();
        vls = vls.map((v, i, arr) => [i + 11].concat(v));
        let p = 0;
        let row = 0;
        for (let i = 0; i < vls.length; i++) {
          if (vls[i][1]) { continue; }
          row = vls[i][0];

          sheet.getRange(row, task.param.cols.key).setValue(projectKeysAdd[p]);
          p++;
          if (p >= projectKeysAdd.length) { break; }
        }

        for (; p < projectKeysAdd.length; p++) {
          row++;
          sheet.getRange(row, task.param.cols.key).setValue(projectKeysAdd[p]);
        }

      }
      sheetModel_ПроектыВРаботе.setMem(this.names.обновлено, new Date());
    } catch (err) { mrErrToString(err); }

    //  Полное название	Дубль названия
    try {
      let mapDictionarys = getMapDictionarys(getSettings().sheetNames.Словари);
      // Logger.log(mapDictionarys.size);
      // mapDictionarys.forEach(v=>Logger.log(JSON.stringify(v)));
      // Logger.log()
      let context_to = new MrContext(task.param.urls.to, getSettings());
      let sheetModel_ПроектыВРаботе = new MrClassSheetModel(task.param.sheetNames.to, context_to);
      let head_key = sheetModel_ПроектыВРаботе.head_key;

      // "Словарь.Полное название"
      // "Словарь.Дубль названия"
      // Logger.log(nc(head_key.indexOf("Словарь.Полное название")));
      // Logger.log(nc(head_key.indexOf("Словарь.Дубль названия")));

      let colПолноеНазвание = head_key.indexOf("Словарь.Полное название");
      let colДубльНазвания = head_key.indexOf("Словарь.Дубль названия");

      let items = new Array();

      mapDictionarys.forEach(v => {
        if (!v["Полное название"]) { return; }
        if (!v["Дубль названия"]) { return; }
        let Словарь = new Object();
        Словарь["Полное название"] = v["Полное название"];
        Словарь["Дубль названия"] = v["Дубль названия"];
        items.push(Словарь);
      });

      // Logger.log(Словарь);

      if (colПолноеНазвание > 0) {
        sheetModel_ПроектыВРаботе.sheet.getRange(sheetModel_ПроектыВРаботе.row.body.first, colПолноеНазвание, sheetModel_ПроектыВРаботе.row.body.last - sheetModel_ПроектыВРаботе.row.body.first + 1, 1).clearContent();
        sheetModel_ПроектыВРаботе.sheet.getRange(sheetModel_ПроектыВРаботе.row.body.first, colПолноеНазвание, items.length, 1).setValues(items.map(Словарь => [Словарь["Полное название"]]));
      }

      if (colДубльНазвания > 0) {
        sheetModel_ПроектыВРаботе.sheet.getRange(sheetModel_ПроектыВРаботе.row.body.first, colДубльНазвания, sheetModel_ПроектыВРаботе.row.body.last - sheetModel_ПроектыВРаботе.row.body.first + 1, 1).clearContent();
        sheetModel_ПроектыВРаботе.sheet.getRange(sheetModel_ПроектыВРаботе.row.body.first, colДубльНазвания, items.length, 1).setValues(items.map(Словарь => [Словарь["Дубль названия"]]));
      }

      sheetModel_ПроектыВРаботе.setMem(this.names.обновлено, new Date());
    } catch (err) { mrErrToString(err); }

    // Logger.log(`${JSON.stringify(task)} `);
    try {
      let context_to = new MrContext(task.param.urls.to, getSettings());
      let context_from = new MrContext(task.param.urls.from, getSettings());
      // let sheetModel_ПроектыВИсполнение = new MrClassSheetModel(task.param.sheetNames.from, context_from);
      let sheetModel_ПроектыВИсполнение = new MrClassSheetModelПроектыВИсполнение(task.param.sheetNames.from, context_from);
      let mapПроектыВИсполнение = sheetModel_ПроектыВИсполнение.getMap();

      let sheetModel_ПроектыВРаботе = new MrClassSheetModel(task.param.sheetNames.to, context_to);
      let mapПроектыВРаботе = sheetModel_ПроектыВРаботе.getMap();

      let mapItems = new Map();

      mapПроектыВРаботе.forEach((item, key) => {
        try {

          let Р = (() => {
            try {
              let ret = mapПроектыВИсполнение.get(key).Р;
              if (typeof ret != "object") {
                ret = JSON.parse(ret)
              }
              return ret;
            } catch (err) { return new Object() }
          })();


          let Тб = (() => {
            try {
              let ret = mapПроектыВИсполнение.get(key).Тб
              if (typeof ret != "object") { ret = JSON.parse(ret) }
              return ret;
            } catch (err) { return new Object(); }
          })();


          let Таблица = mapПроектыВИсполнение.get(key).Таблица;

          mapItems.set(key, { key, Р, Таблица, Тб });
        } catch (err) { Logger.log(key); mrErrToString(err); }
      });
      sheetModel_ПроектыВРаботе.updateItems([...mapItems.values()]);
      sheetModel_ПроектыВРаботе.setMem(this.names.обновлено, new Date());
    } catch (err) { mrErrToString(err); }
  }


  /** @param {DataTaskSyncРабочийЛист} task*/
  taskExportSyncРабочийЛист(task) {

    let f_key_i = ((key, n) => { let nnn = `000${n}`.slice(-3); return `${key} № ${nnn}` });

    Logger.log(`${JSON.stringify(task)} `);
    let context_to = new MrContext(task.param.urls.to, getSettings());

    // let sheetModel_РабочийЛист = new MrClassSheetModel(task.param.sheetNames.to, context_to);
    let sheetModel_РабочийЛист = new MrClassSheetModelРабочийЛист(task.param.sheetNames.to, context_to);
    if ((new Date().getTime() - new Date(sheetModel_РабочийЛист.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
      Logger.log("да ещё актуальны"); return;
    }

    let context_from = new MrContext(task.param.urls.from, getSettings());

    // let sheetModel_ПроектыВИсполнение = new MrClassSheetModel(task.param.sheetNames.from, context_from);
    let sheetModel_ПроектыВИсполнение = new MrClassSheetModelПроектыВИсполнение(task.param.sheetNames.from, context_from);



    let mapПроектыВИсполнение = sheetModel_ПроектыВИсполнение.getMap();
    // let mapРабочийЛист = sheetModel_РабочийЛист.getMap();

    let arr = [...mapПроектыВИсполнение.values()].filter((obj) => {
      return true;
      return obj["Рабочий Лист Логиста"] === true
    });
    arr.reverse();
    // Logger.log(`arr=${arr.length} |       // key= |${arr.map(v => v["key"])}      `);

    let date = new Date();
    let itemArr = new Array();
    arr.forEach(project => {
      // Logger.log(project.key);
      // Logger.log(`project=${project.key} |   `);
      // itemArr.push({
      //   key: f_key_i(project.key, 0),
      //   r: 0,
      //   Проект: project.key,
      //   Р: (() => { try { return JSON.parse(project.Р); } catch { return new Object(); } })(),
      //   date: date,
      // });

      if (fl_str(project["Р.Исполнение.Статус проекта"]) == fl_str("ЗАВЕРШЕН")) {
        Logger.log([project.key, project["Р.Исполнение.Статус проекта"]]);
        return;
      }

      if (fl_str(project["Р.Исполнение.Статус проекта"]) == fl_str("Завершен")) {
        Logger.log([project.key, project["Р.Исполнение.Статус проекта"]]);
        return;
      }

      Logger.log(project.key);
      let productMap = new Map();
      /** @type {Array} */
      let Тб_Выбор_Поставщиков_тело = (() => {
        try {
          let Тб = project.Тб;
          if (typeof Тб != "object") { Тб = JSON.parse(Тб); }
          return Тб.Выбор_Поставщиков.тело;
        } catch (err) { Logger.log(project.key); mrErrToString(err); return new Array() }
      })();
      // Logger.log(`arr=${Тб_Выбор_Поставщиков_тело.length} |      // key= |${Тб_Выбор_Поставщиков_тело.map(v => [v["№"], v["НОМЕНКЛАТУРА"], "\n"])}      `);
      Тб_Выбор_Поставщиков_тело.forEach(pr => {
        // let key_i = `${project.key}№${pr["№"]}`;
        let key_i = f_key_i(project.key, pr["№"]);
        let item = new Object();

        item.Проект = project.key;
        item.Р = (() => {
          try {
            let Р = project.Р;
            if (typeof Р != "object") { Р = JSON.parse(Р); }
            return Р;
          } catch (err) { mrErrToString(err); return new Object(); }
        })();
        item["key"] = key_i;
        item["r"] = pr["№"];
        item["Тб"] = new Object();
        item["Тб"]["Выбор_Поставщиков"] = pr;
        item["Проект"] = project.key;
        // Тб.Настройки_писем.Требования к составу технической документации
        item["Тб"]["Настройки_писем"] = project.Тб["Настройки_писем"];
        item.date = date;
        productMap.set(item["key"], item);
      });

      // Тб.Оплаты.тело
      let Тб_Оплаты_тело = (() => {
        try {
          let Тб = project.Тб;
          if (typeof Тб != "object") { Тб = JSON.parse(Тб); }
          return Тб.Оплаты.тело;
        } catch (err) {
          //  mrErrToString(err); 
          return new Array()
        }
      })();


      Тб_Оплаты_тело.forEach(pr => {
        let key_i = f_key_i(project.key, pr["№"]);
        let item = productMap.get(key_i);
        if (!item) { return; }
        item["Тб"]["Оплата"] = pr;
      });

      itemArr = itemArr.concat([...productMap.values()].sort((a, b) => {
        if (Number.parseInt(a["r"]) > Number.parseInt(b["r"])) { return 1; }
        if (Number.parseInt(a["r"]) < Number.parseInt(b["r"])) { return -1; }
        return 0;
      }));

      // itemArr.forEach(it => {         Logger.log(`${it.key} | ${JSON.stringify(it)}`);      });
    });
    // itemArr.forEach(it => { Logger.log(`${it.key} | ${JSON.stringify(it.Р.Исполнение,null,2)}`); });
    // itemArr.forEach(it => { Logger.log(`${it.key} | `) });



    itemArr.forEach(item => {
      if (!item.Р.Исполнение) {
        Logger.log("Внимание");
        Logger.log(JSON.stringify(item, null, 2));

        return;
      }


    });


    sheetModel_РабочийЛист.updateItems(itemArr);
    sheetModel_РабочийЛист.setMem(this.names.обновлено, new Date());

  }





  /** @param {DataTaskSyncУПД_ДОК} task*/
  taskExportSyncУПД_ДОК(task) {
    let dateFildsName = [
      fl_str("Дата подписи продавцом"),
      fl_str("Дата подписи Покупателем"),
      fl_str("Дата документа"),
      fl_str("Дата договора от"),

    ];

    // let col_CheckName = "для проверки расхождения данных с логистом";
    let col_CheckName = this.col_CheckName;
    // let col_Name_Has = "Присутствует в таблице проекта";
    let col_Name_Has = this.col_Name_Has;

    let col_Name_N_УПД_о_закупке = "№ УПД о закупке";
    let col_Name_N_Договора_о_закупке = "№ Договора о закупке";

    let cols_update_УПД = this.cols_update_УПД;

    let cols_update_ДОК = this.cols_update_ДОК;

    Logger.log(`${JSON.stringify(task)} `);
    try {


      // Logger.log(`${JSON.stringify(task)} `);
      let context_to = new MrContext(task.param.urls.to, getSettings());
      // let sheet_ДОК = context_to.getSheetByName(task.param.sheetNames.to_ДОК);

      let sheetModel_Рабочий_Лист = new MrClassSheetModel(task.param.sheetNames.to_Рабочий_Лист, context_to);
      if ((new Date().getTime() - new Date(sheetModel_Рабочий_Лист.getMem(this.names.обновлено_УПД_ДОК)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
        Logger.log("да ещё актуальны"); return;
      }


      // let range_str = `${nc(task.param.cols.key)}11:${nc(task.param.cols.key)}`;
      // let vls = sheet_ДОК.getRange(range_str).getValues();
      // vls = vls.filter(v => v[0]).map(v => v[0]);
      // // Logger.log(`vls = ${vls}`);

      let context_from = new MrContext(task.param.urls.from, getSettings());
      // let sheetModel_ПроектыВИсполнение = new MrClassSheetModel(task.param.sheetNames.from, context_from);
      let sheetModel_ПроектыВИсполнение = new MrClassSheetModelПроектыВИсполнение(task.param.sheetNames.from, context_from);

      let mapПроектыВИсполнение = sheetModel_ПроектыВИсполнение.getMap();




      let mapПроектыВИсполнение_ДОК = new Map();
      let mapПроектыВИсполнение_УПД = new Map();

      Logger.log(JSON.stringify({ mapПроектыВИсполнение_size: mapПроектыВИсполнение.size }));


      // mapПроектыВИсполнение.forEach((item, key) => {

      //   try {

      //     // Logger.log( JSON.parse( item.Тб) );
      //     // let Тб = JSON.parse(item.Тб)
      //     let Тб = (() => {
      //       try {
      //         let Тб = item.Тб;
      //         if (typeof Тб != "object") { Тб = JSON.parse(Тб); }
      //         return Тб;
      //       } catch (err) {
      //         Logger.log(`${key} | ${item.ver} | ${typeof item.Тб}`);
      //         // Logger.log(typeof Тб);
      //         mrErrToString(err);
      //         return new Object();
      //       }
      //     })()


      //     let УПД = Тб.УПД;
      //     let ДОК = Тб.ДОК;
      //     // Logger.log(JSON.stringify({ key, ДОК, УПД }));
      //     if (Array.isArray(УПД)) { УПД.forEach(v => { if (!v.key) { return; } if (!v["Проект"]) { return; } mapПроектыВИсполнение_УПД.set(v.key, v); }); }
      //     if (Array.isArray(ДОК)) { ДОК.forEach(v => { if (!v.key) { return; } if (!v["Проект №"]) { return; } mapПроектыВИсполнение_ДОК.set(v.key, v); }); }
      //   } catch (err) { mrErrToString(err); }
      // });

      [...mapПроектыВИсполнение.values()]
        .reverse()
        .forEach(element => {
          ((item, key) => {
            try {

              // Logger.log( JSON.parse( item.Тб) );
              // let Тб = JSON.parse(item.Тб)
              let Тб = (() => {
                try {
                  let Тб = item.Тб;
                  if (typeof Тб != "object") { Тб = JSON.parse(Тб); }
                  return Тб;
                } catch (err) {
                  Logger.log(`${key} | ${item.ver} | ${typeof item.Тб}`);
                  // Logger.log(typeof Тб);
                  mrErrToString(err);
                  return new Object();
                }
              })()


              let УПД = Тб.УПД;
              let ДОК = Тб.ДОК;
              // Logger.log(JSON.stringify({ key, ДОК, УПД }));
              if (Array.isArray(УПД)) { УПД.forEach(v => { if (!v.key) { return; } if (!v["Проект"]) { return; } mapПроектыВИсполнение_УПД.set(v.key, v); }); }
              if (Array.isArray(ДОК)) { ДОК.forEach(v => { if (!v.key) { return; } if (!v["Проект №"]) { return; } mapПроектыВИсполнение_ДОК.set(v.key, v); }); }
            } catch (err) { mrErrToString(err); }
          })(element, element.key)
        }
        );

      Logger.log(JSON.stringify({ ДОК_size: mapПроектыВИсполнение_ДОК.size, УПД_size: mapПроектыВИсполнение_УПД.size }));

      // return;

      let sheet_to_УПД = context_to.getSheetByName(task.param.sheetNames.to_УПД);
      let sheet_to_ДОК = context_to.getSheetByName(task.param.sheetNames.to_ДОК);

      let max_row_УПД = Math.max(2, ...sheet_to_УПД.getRange("B1:B").getValues().map((v, i) => [i + 1].concat(v)).filter(v => v[1]).map(v => v[0]));
      let max_row_ДОК = Math.max(2, ...sheet_to_ДОК.getRange("B1:B").getValues().map((v, i) => [i + 1].concat(v)).filter(v => v[1]).map(v => v[0]));



      let heads_УПД = sheet_to_УПД.getRange("A1:AD1").getValues()[0];
      let heads_ДОК = sheet_to_ДОК.getRange("A1:AP1").getValues()[0];

      let col_ID_УПД = heads_УПД.indexOf("ID") + 1;
      let col_ID_ДОК = heads_ДОК.indexOf("ID") + 1;
      let col_Check_УПД = heads_УПД.indexOf(col_CheckName) + 1;
      let col_Check_ДОК = heads_ДОК.indexOf(col_CheckName) + 1;
      // let col_Has_УПД = heads_УПД.indexOf(col_Name_Has) + 1;
      // let col_Has_ДОК = heads_ДОК.indexOf(col_Name_Has) + 1;

      let col_УПД_о_закупке = heads_УПД.indexOf(col_Name_N_УПД_о_закупке) + 1;
      let col_ДОК_о_закупке = heads_ДОК.indexOf(col_Name_N_Договора_о_закупке) + 1;

      if (col_ID_ДОК) {

        // Logger.log({ col_ID_ДОК });
        let map_ДОК_ID = new Map();
        let rowBodyFirst = task.param.rows.body.first;

        // todo
        // очистка col_ДОК_о_закупке
        // sheet_to_ДОК.getRange(rowBodyFirst, col_ДОК_о_закупке, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).clearContent();


        sheet_to_ДОК
          .getRange(rowBodyFirst, col_ID_ДОК, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1)
          .getValues()
          .map((v, i) => [rowBodyFirst + i].concat(v))
          .filter(v => v[1])
          .forEach(v => { map_ДОК_ID.set(v[1], v) });





        mapПроектыВИсполнение_ДОК.forEach((item, key) => {
          if (map_ДОК_ID.has(key)) {
            let row = map_ДОК_ID.get(key)[0];
            let col = col_Check_ДОК;
            // Logger.log(`${(row && (col > 0))} | ${row} | ${col}`)
            // if (row && (col > 0)) {
            //   sheet_to_ДОК.getRange(row, col).setValue(item[col_CheckName]);
            // }
            // обновление данных
            cols_update_ДОК.forEach((col_Name) => {
              col = heads_ДОК.indexOf(col_Name) + 1;
              if (row && (col > 0)) {
                let curent_value = sheet_to_ДОК.getRange(row, col).getValue();
                if (`${curent_value}` != "") {
                  if (!["№ Договора о закупке",].includes(col_Name)) {
                    return;
                  }
                }
                if (`${item[col_Name]}` == "") {
                  if (!["№ Договора о закупке",].includes(col_Name)) {
                    return;
                  }
                  // return;
                }

                let newValue = item[col_Name];

                if (dateFildsName.includes(fl_str(col_Name))) {
                  if (newValue) {
                    newValue = new Date(newValue);
                    newValue = Utilities.formatDate(newValue, "Europe/Moscow", "dd.MM.yyyy")
                  }
                }
                Logger.log(`${newValue} добавить = ${JSON.stringify({ row, col, col_Name, newValue })}`);
                sheet_to_ДОК.getRange(row, col).setValue(newValue);
              }
            });


            return;
          }
          let row = Math.max(max_row_ДОК, ...[...map_ДОК_ID.values()].map(v => v[0])) + 1;
          let vl = new Array(heads_ДОК.length);
          item["ID"] = key;
          heads_ДОК.forEach((h, i) => {
            // if (!h) { return; }
            vl[i] = item[h];

            if (dateFildsName.includes(fl_str(h))) {
              if (vl[i]) {
                vl[i] = new Date(vl[i]);
              }
            }

            if (h == "Проект №") {
              if (!vl[i]) { return; }
              vl[i] = `${vl[i]}`;
            }

          });
          try {
            sheet_to_ДОК.getRange(row, 1, 1, vl.length).setValues([vl]);
          } catch (err) { }
          // map_ДОК_ID.set(key, [row].concat(vl));
          map_ДОК_ID.set(key, [row, key]);
        });


        // let vls = sheet_to_ДОК.getRange(rowBodyFirst, col_Has_ДОК, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).getValues();
        // vls = vls.map(v => [undefined]);
        // map_ДОК_ID.forEach(([row, key], keyM) => {
        //   // Logger.log({ row, key, keyM });
        //   if (!key) { return; }
        //   if (!mapПроектыВИсполнение_ДОК.has(key)) { return; }
        //   vls[row - rowBodyFirst][0] = true;
        // });
        // sheet_to_ДОК.getRange(rowBodyFirst, col_Has_ДОК, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).setValues(vls);



        // Вставка Номеров сткоки 
        try {
          let maxNumber = 0;
          maxNumber = Math.max(maxNumber, ...sheet_to_ДОК.getRange(rowBodyFirst, nr("A"), sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).getValues().flat(2)) + 1;
          Logger.log(`maxNumber = ${maxNumber}`);

          let id_ДОКs = sheet_to_ДОК.getRange(rowBodyFirst, col_ID_ДОК, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).getValues();
          let vlsNumber = sheet_to_ДОК.getRange(rowBodyFirst, nr("A"), sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).getValues();

          vlsNumber = vlsNumber.map((v, i) => [i + rowBodyFirst].concat(v));
          // Logger.log(`vlsNumber = ${vlsNumber}`);
          vlsNumber.forEach((v, i) => {
            if (`${v[1]}` != "") { return; }

            let row = v[0];
            let id_ДОК = id_ДОКs[i][0];
            if (!id_ДОК) { return; }
            // Logger.log(JSON.stringify({ maxNumber, row, id_ДОК }, null, 2)); maxNumber++;
            sheet_to_ДОК.getRange(row, nr("A")).setValue(maxNumber++);
          });
        } catch (err) { mrErrToString(err); }


        // Вставка Формулы 

        try {
          let col_FormulaName = "№ строки с Договором о закупке";
          let col_Formula = heads_ДОК.indexOf(col_FormulaName) + 1;
          // let formula = `${sheet_to_ДОК.getRange(1, col_Formula).getNote()}`;
          let formula = ``;
          let id_ДОКs = sheet_to_ДОК.getRange(rowBodyFirst, col_ID_ДОК, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).getValues();

          let formulas = sheet_to_ДОК.getRange(rowBodyFirst, col_Formula, sheet_to_ДОК.getLastRow() - rowBodyFirst + 1, 1).getFormulasR1C1();

          formulas.forEach((f, i) => {
            let row = i + rowBodyFirst;

            let id_ДОК = id_ДОКs[i][0];
            if (!id_ДОК) { return; }
            if (`${id_ДОК}`.slice(0, 1) != "П") { return; }

            if (`${f[0]}` != "") {
              formula = f[0];
              return;
            }
            if (f[0] == `${formula}`) { return; }
            if (`${formula}` == "") { return; }

            Logger.log(JSON.stringify({ row, id_ДОК, f, formula }, null, 2));

            // let ff = sheet_to_ДОК.getRange(row - 1, col_Formula, 1, 1).getFormulasR1C1();
            sheet_to_ДОК.getRange(row, col_Formula).clear();
            sheet_to_ДОК.getRange(row, col_Formula).setFormulaR1C1(formula);
            // sheet_to_ДОК.getRange(row, col_Formula, 1, 1).setFormulaR1C1(ff[0][0]);
            sheet_to_ДОК.getRange(row, col_Formula + 1).setValue(undefined)

          });
        } catch (err) { mrErrToString(err); }


      }


      if (col_ID_УПД) {
        // Logger.log({ col_ID_УПД });
        // mapПроектыВИсполнение_УПД.forEach(v => Logger.log(JSON.stringify(v, undefined, 2)));


        // Logger.log(JSON.stringify({ ДОК_size: mapПроектыВИсполнение_ДОК.size, УПД_size: mapПроектыВИсполнение_УПД.size }));




        let map_УПД_ID = new Map();
        let rowBodyFirst = task.param.rows.body.first;
        // очистка УПД_о_закупке
        // sheet_to_УПД.getRange(rowBodyFirst, col_УПД_о_закупке, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).clearContent();

        sheet_to_УПД
          .getRange(rowBodyFirst, col_ID_УПД, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1)
          .getValues()
          .map((v, i) => [rowBodyFirst + i].concat(v))
          .filter(v => v[1])
          .forEach(v => { map_УПД_ID.set(v[1], v) });

        mapПроектыВИсполнение_УПД.forEach((item, key) => {
          if (map_УПД_ID.has(key)) {
            // Logger.log(JSON.stringify({ key, item }, undefined, 2));

            let row = map_УПД_ID.get(key)[0];
            let col = col_Check_УПД;
            // Logger.log(`${(row && (col > 0))} | ${row} | ${col}`)
            // if (row && (col > 0)) {
            //   sheet_to_УПД.getRange(row, col).setValue(item[col_CheckName]);
            // }


            // обновление данных
            cols_update_УПД.forEach((col_Name) => {
              col = heads_УПД.indexOf(col_Name) + 1;
              if (row && (col > 0)) {

                let curent_value = sheet_to_УПД.getRange(row, col).getValue();
                // sheet_to_УПД.getRange(rowBodyFirst, col_УПД_о_закупке, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).clearContent();

                if (col_УПД_о_закупке != col) {
                  if (`${curent_value}` != "") {
                    // if (!["№ УПД о закупке",].includes(col_Name)) {
                    return;
                    // }
                  }
                }

                if (col_УПД_о_закупке != col) {
                  if (`${item[col_Name]}` == "") { return; }
                }

                let newValue = item[col_Name];

                if (col_УПД_о_закупке == col) {
                  if (curent_value == newValue) {
                    // Logger.log(`НЕ  добавить = ${JSON.stringify({ row, col, col_Name, newValue })}`);
                    return;
                  }
                }



                if (dateFildsName.includes(fl_str(col_Name))) {
                  if (newValue) {
                    newValue = new Date(newValue);
                    newValue = Utilities.formatDate(newValue, "Europe/Moscow", "dd.MM.yyyy")
                  }
                }
                Logger.log(`  ${newValue} добавить = ${JSON.stringify({ row, col, col_Name, newValue })}`);
                // if (fl_str(col_Name) == fl_str("№ УПД о закупке")) {

                // }
                sheet_to_УПД.getRange(row, col).setValue(newValue);

              }
            });

            return;
          }

          let row = Math.max(max_row_УПД, ...[...map_УПД_ID.values()].map(v => v[0])) + 1;
          let vl = new Array(heads_УПД.length);
          item["ID"] = key;
          heads_УПД.forEach((h, i) => {
            // if (!h) { return; }
            vl[i] = item[h];
            if (dateFildsName.includes(fl_str(h))) {
              if (vl[i]) {
                vl[i] = new Date(vl[i]);
              }
            }

            if (h == "Проект") {
              if (!vl[i]) { return; }
              vl[i] = `${vl[i]}`;
            }


          });
          try {
            sheet_to_УПД.getRange(row, 1, 1, vl.length).setValues([vl]);
          } catch (err) { }
          // map_УПД_ID.set(key, [row].concat(vl));
          map_УПД_ID.set(key, [row, key]);
        });

        // let vls = sheet_to_УПД.getRange(rowBodyFirst, col_Has_УПД, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).getValues();
        // vls = vls.map(v => [undefined]);
        // map_УПД_ID.forEach(([row, key], keyM) => {
        //   // Logger.log({ row, key, keyM });
        //   if (!key) { return; }
        //   if (!mapПроектыВИсполнение_УПД.has(key)) { return; }
        //   vls[row - rowBodyFirst][0] = true;
        // });
        // sheet_to_УПД.getRange(rowBodyFirst, col_Has_УПД, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).setValues(vls);


        // Вставка Номеров сткоки 
        try {
          let maxNumber = 0;
          maxNumber = Math.max(maxNumber, ...sheet_to_УПД.getRange(rowBodyFirst, nr("A"), sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).getValues().flat(2)) + 1;
          // Logger.log(`maxNumber = ${maxNumber}`);

          let id_УПДs = sheet_to_УПД.getRange(rowBodyFirst, col_ID_УПД, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).getValues();
          let vlsNumber = sheet_to_УПД.getRange(rowBodyFirst, nr("A"), sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).getValues();

          vlsNumber = vlsNumber.map((v, i) => [i + rowBodyFirst].concat(v));
          // Logger.log(`vlsNumber = ${vlsNumber}`);


          vlsNumber.forEach((v, i) => {
            if (`${v[1]}` != "") { return; }

            let row = v[0];
            let id_УПД = id_УПДs[i][0];
            if (!id_УПД) { return; }
            // Logger.log(JSON.stringify({ maxNumber, row, id_УПД }, null, 2));
            // maxNumber++;
            sheet_to_УПД.getRange(row, nr("A")).setValue(maxNumber++);
          });
        } catch (err) { mrErrToString(err); }



        // Вставка Формулы 

        try {
          let col_FormulaName = "Требуется ли Партнерская УПД";
          let col_Formula = heads_УПД.indexOf(col_FormulaName) + 1;
          // let formula = `=${sheet_to_УПД.getRange(1, col_Formula).getNote()}`;
          let formula = "";
          let id_УПДs = sheet_to_УПД.getRange(rowBodyFirst, col_ID_УПД, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).getValues();

          let formulas = sheet_to_УПД.getRange(rowBodyFirst, col_Formula, sheet_to_УПД.getLastRow() - rowBodyFirst + 1, 1).getFormulasR1C1();

          formulas.forEach((f, i) => {
            let row = i + rowBodyFirst;

            let id_УПД = id_УПДs[i][0];
            if (!id_УПД) { return; }
            if (`${id_УПД}`.slice(0, 1) != "П") { return; }
            if (`${f[0]}` != "") {
              // if (`${formula}` == "") { }
              formula = f[0];
              return;
            }
            if (f[0] == formula) { return; }
            if (`${formula}` == "") { return; }

            Logger.log(JSON.stringify({ row, id_УПД, f, formula }, null, 2));
            // sheetNames_to_УПД.getRange(row, col_Formula, 1, 2).setFormulas([[formula, undefined]]);
            sheet_to_УПД.getRange(row, col_Formula).clear();
            sheet_to_УПД.getRange(row, col_Formula).setFormulaR1C1(formula);
            sheet_to_УПД.getRange(row, col_Formula + 1).setValue(undefined);

          });
        } catch (err) { mrErrToString(err); }


      }

      sheetModel_Рабочий_Лист.setMem(this.names.обновлено_УПД_ДОК, new Date());
      // this.checkRow([sheet_to_ДОК, context_to.getSheetByName("Ошибки Листа Договора")]);

    } catch (err) { mrErrToString(err); }

  }




  checkRow(sheets) {
    Logger.log(`ExeTask_SyncLogist checkRow`);


    if (!Array.isArray(sheets)) { Logger.log(`не задан массив sheets ${sheets}`); return; }
    if (sheets.length < 2) { Logger.log(`массив мал sheets.length=${sheets.length}`); return; }
    try {
      let maxRows = sheets[0].getMaxRows();
      sheets.forEach(sheet => {
        if (sheet.getMaxRows() != maxRows) {
          let howManyRows = sheet.getMaxRows() - maxRows;
          if (howManyRows < 0) {
            sheet.insertRowsAfter(sheet.getMaxRows() - 1, -howManyRows);
          } else {
            sheet.deleteRows(sheet.getMaxRows() - 1 - howManyRows, howManyRows);
          }
        }
      });
    } catch (err) { mrErrToString(err); }
  }




  /** @param {DataTaskClearSelectRowInУПД_ДОК} task*/
  taskClearSelectRowInУПД_ДОК(task) {

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getUrl().includes(getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин)) { Logger.log(`Не подходящяя таблица`); return; }
    let sheet = ss.getActiveSheet();
    if (![getSettings().sheetNames.УПД, getSettings().sheetNames.ДОК].includes(sheet.getSheetName())) { Logger.log(`Не подходящяя лист ${sheet.getSheetName()}`); return; }



    // if (!in_sheet) { return; }
    // if (![task.param.sheetNames.to_ДОК, task.param.sheetNames.to_УПД].includes(in_sheet)) { Logger.log(`Не подходящяя лист ${in_sheet}`); return; }






    let rowBodyFirst = task.param.rows.body.first;
    let rowBodyLast = sheet.getLastRow();
    let activeRangeList = sheet.getActiveRangeList();
    if (!activeRangeList) { return; }

    let ranges = activeRangeList.getRanges();
    let rows = new Set();

    for (let i = 0; i < ranges.length; i++) {
      let range = ranges[i];
      for (let r = range.getRow(); r <= range.getLastRow(); r++) {
        if (r < rowBodyFirst) { continue; }
        // if (rowBodyLast < r) { continue; }

        rows.add(r);
      }
    }

    if (rows.size == 0) { return; }
    let rowsArr = [...rows.keys()];
    Logger.log(rowsArr.join(", "));

    let in_sheet = sheet.getSheetName();
    let clearColsName = [];
    if (in_sheet == task.param.sheetNames.to_ДОК) {
      clearColsName = this.cols_update_ДОК;
    }

    if (in_sheet == task.param.sheetNames.to_УПД) {
      clearColsName = this.cols_update_УПД;
    }

    if (clearColsName.length == 0) { return; }


    // Logger.log(`${JSON.stringify(task)} `);
    let context_to = new MrContext(getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин, getSettings());
    // let sheetModel_Рабочий_Лист = new MrClassSheetModel(task.param.sheetNames.to_Рабочий_Лист, context_to);
    // if ((new Date().getTime() - new Date(sheetModel_Рабочий_Лист.getMem(this.names.обновлено_УПД_ДОК)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
    //   Logger.log("да ещё актуальны"); return;
    // }

    let sheetModelTo = new MrClassSheetModel(in_sheet, context_to, task.param.rowConf.to);
    let mapTo = sheetModelTo.getMap();
    // mapTo.forEach(item => Logger.log(JSON.stringify(item)));

    let arrClear = [...mapTo.values()].filter(item => {
      if (`${item.ID}` == "") { return false; }
      if (item[this.col_Name_Has] != true) { return false; }
      if (!rowsArr.includes(item.row)) { return false; }
      return true;
    });

    // arrClear..forEach(item => Logger.log(JSON.stringify(item)));

    let heads = sheetModelTo.sheet.getRange(task.param.rowConf.to.head.key, 1, 1, sheetModelTo.sheet.getLastColumn()).getValues()[0];


    Logger.log(arrClear.map(item => item.row).join(", "));
    arrClear.forEach(item => {
      let row = item.row;
      clearColsName.forEach(col_Name => {
        let col = heads.indexOf(col_Name) + 1;
        if (col < 1) { return; }
        Logger.log(JSON.stringify({ row, col, col_Name }));
        sheetModelTo.sheet.getRange(row, col).clearContent();
      });

    });


    let sheetModel_Рабочий_Лист = new MrClassSheetModel(task.param.sheetNames.mem_Рабочий_Лист, context_to);
    // if ((new Date().getTime() - new Date(sheetModel_Рабочий_Лист.getMem(this.names.обновлено_УПД_ДОК)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
    //   Logger.log("да ещё актуальны"); return;
    // }
    sheetModel_Рабочий_Лист.setMem(this.names.обновлено_УПД_ДОК, new Date(new Date().getTime() + task.актуально_дней * DeyMilliseconds));
  }





  /** @param {DataTaskРабочийЛистОбработка} task*/
  taskExportРабочийЛистОбработка(task) {





    Logger.log(`${JSON.stringify(task)} `);
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModel_РабочийЛист = new MrClassSheetModel(task.param.sheetNames.to, context_to);
    if ((new Date().getTime() - new Date(sheetModel_РабочийЛист.getMem(this.names.обработкаВыполнена)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
      Logger.log("да ещё актуальны"); return;
    }


    let colNames = {
      Проц_оплаты: "Тб.Оплата.% оплаты",
      Статус_доплаты: "Статус_доплаты",
    };
    let mapРабочийЛист = sheetModel_РабочийЛист.getMap();
    // let itemArr = new Array();
    let dp = 0.0001;

    let itemArr = [...mapРабочийЛист.values()].filter(item => {
      let po = item[colNames.Проц_оплаты];
      let so = item[colNames.Статус_доплаты];
      if (po == "") { return false; }
      if (so == this.СтатусыДоплаты.Доплачено) { return false; }
      if (so == this.СтатусыДоплаты.НеТребуется) { return false; }
      if (Number.isNaN(po)) { return false; }
      if (Math.abs(1 - po) >= dp) { return false; }
      // Logger.log(JSON.stringify({
      //   k: item.key,
      //   p: item[colNames.Проц_оплаты],
      //   s: item[colNames.Статус_доплаты],
      //   // i: item,
      // }, undefined, 2));
      return true;
    }).map(item => {
      let ret = { key: item.key, }
      ret[colNames.Статус_доплаты] = this.СтатусыДоплаты.Доплачено;
      return ret;
    });

    // Logger.log(JSON.stringify(itemArr, null, 2));

    if (itemArr.length > 0) {
      // return;
      sheetModel_РабочийЛист.updateItems(itemArr);
    }

    sheetModel_РабочийЛист.setMem(this.names.обработкаВыполнена, new Date());

  }






  /** @param {DataTaskCopyРабочийЛистToЛогистика} task*/
  taskCopyРабочийЛистToЛогистика(task) {
    // Logger.log(`taskCopyРабочийЛистToЛогистика`);
    Logger.log(`${JSON.stringify(task, null, 2)} `);

    let theContext = {
      from: new MrContext(task.param.urls.from, getSettings()),
      to: new MrContext(task.param.urls.to, getSettings()),
    }
    let theSheetModel = {
      from: new MrClassSheetModel(task.param.sheetNames.from, theContext.from, task.param.row.from),
      to: new MrClassSheetModel(task.param.sheetNames.to, theContext.to, task.param.row.to),
    }

    try {

      if ((new Date().getTime() - new Date(theSheetModel.to.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
        Logger.log("да ещё актуальны"); return;
      }

      let theRegExpDot = new RegExp(/\./, "gu");
      Logger.log(theSheetModel.from.head_key);
      theSheetModel.from.head_key = (() => {
        try {
          return []
            .concat(
              "row",
              theSheetModel.from.sheet.getRange(theSheetModel.from.row.head.ru_name, 1, 1, theSheetModel.from.col.last).getValues()[0],
            ).map(theName => {

              theName = ((h, sh) => {
                sh.forEach(s => {
                  if (`${h}`.includes(s)) {
                    h = s;
                  }
                })
                return h;
              })(theName, task.param.simplifyHeader);

              return `${theName}`.replace(theRegExpDot, "");
            }

            );
        }
        catch (err) {
          mrErrToString(err); return [];
        }

      })();
      Logger.log(theSheetModel.from.head_key);
      // theSheetModel.to.requiredColNames(theSheetModel.from.head_key);

      let mapFrom = theSheetModel.from.getMap();

      let itemArr = [...mapFrom.values()];

      itemArr.forEach(item => {
        item["ДатаВремяАктуальности"] = theContext.from.timeConstruct;
      });
      // Logger.log(JSON.stringify(itemArr.slice(3,6), null, 2));

      if (itemArr.length > 0) {
        // return;
        theSheetModel.to.setItems(itemArr);
      }
    } catch (err) {
      theSheetModel.to.setMem("err", mrErrToString(err));
    }
    theSheetModel.to.setMem(this.names.обновлено, new Date());

  }


}

function menu_ClearSelectRowInУПД_ДОК() {
  // let ss = SpreadsheetApp.getActiveSpreadsheet();
  // if (!ss.getUrl().includes(getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин)) { Logger.log(`Не подходящяя таблица`); return; }
  // if (![getSettings().sheetNames.УПД, getSettings().sheetNames.ДОК].includes(ss.getActiveSheet().getSheetName())) { Logger.log(`Не подходящяя лист ${ss.getActiveSheet().getSheetName()}`); return; }

  // НаборЗадачьClearSelectRowInУПД_ДОК
  ВыполненитьЗадачи("из menu_ClearSelectRowInУПД_ДОК", "НаборЗадачьClearSelectRowInУПД_ДОК")

}


function text_triggerПродолжитьВыполнениеЗадачь(trigerrInfo = "texttriggerПродолжитьВыполнениеЗадачь") {
  Logger.log(JSON.stringify(trigerrInfo));
  triggerLookCache();
  // return;
  let tasks = [].concat(
    [

      (() => {
        let syncРабочийЛист = new DataTaskSyncРабочийЛист();
        syncРабочийЛист.param.urls.from = getSettings().urls.ПромежуткиЛогиста;
        syncРабочийЛист.param.urls.to = "https://docs.google.com/spreadsheets/d//";
        return syncРабочийЛист;
      })(),

    ]
  );

  let mrTaskExe = new MrTaskExecutor(getContext(), tasks, getContext().getScriptCache());
  try {
    mrTaskExe.triggerПродолжитьВыполнениеЗадачь();
  } catch (err) { mrErrToString(err); }

  triggerLookCache();
}








