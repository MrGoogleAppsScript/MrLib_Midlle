/// отслеживания для Юлии
/*
Задача 4.8 - Сделать новую таблицу отслеживания для Юлии
Берем за основу таблицу Антона
Делаем чтобы данные которые относятся к проектам с префиксом Ю копировались в эту таблицу. Копируем скриптом Нужны вкладки 
Рабочий лист
Текущие доставки
Деньги курьеры
Деньги транспортные
Отгрузочные документы
Договора
Оплата Логисту
Проекты в работе
Делаем вкладку прочие расходы, данные берем из Реестра
На вкладке проекты в работе столбцы B,C,D заполняем формулами
*/

class DataTaskSyncОтслеживанияДляЮли {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncОтслеживанияДляЮли.name;
    /** @constant */
    this.name = ExeTask_SyncОтслеживанияДляЮли.prototype.taskSyncОтслеживанияДляЮлии.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      обновлено: "обновлено",
      renameHeads: [],
      sheetNames: {
        from: "",
        to: "",
      },

      urls: {
        from: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,   // Отслеживание поставки   
        to: getSettings().urls.ОтслеживаниеПоставкиЮлии,     // Сделать новую таблицу отслеживания для Юлии 
      },

      // filter: {
      //   fild: "Проект",
      //   prefix: "Ю-",

      //   function: ((obj, fild, prefix) => {
      //     if (`${obj[fild]}`.slice(0, prefix.length) === prefix) { return true; }
      //     return false;
      //   }),
      // },


      filter: {
        fild: undefined,
        prefix: undefined,

        function: ((obj, fild, prefix) => {
          // if (`${obj[fild]}`.slice(0, prefix.length) === prefix) { return true; }
          return true;
        }),
      },

      // rowConf: {
      //   from: {
      //     head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
      //     body: { first: 11, last: 11, },
      //   },
      //   to: {
      //     head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
      //     body: { first: 11, last: 11, },
      //   }
      // }

      rowConf: {
        from: undefined,
        to: undefined,
      }
    };
  }
}



class ExeTask_SyncОтслеживанияДляЮли extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_SyncОтслеживанияДляЮли.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",

    }
    // this.mapValuesFromРеестр = new Map();
    // this.mapValuesFromИсполнение = new Map();
  }


  /** @param {DataTaskSyncОтслеживанияДляЮли} task*/
  taskSyncОтслеживанияДляЮлии(task) {
    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModelАктуально = new MrClassSheetModel(task.param.sheetNames.to, context_to);

    if ((new Date().getTime() - new Date(sheetModelАктуально.getMem(task.param.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    ) { Logger.log("да ещё актуальны"); return; }

    let sheetModelFrom = new MrClassSheetModel(task.param.sheetNames.from, context_from, task.param.rowConf.from);
    // Logger.log(JSON.stringify(sheetModelFrom.head_key));
    if (task.param.renameHeads.length != 0) {
      task.param.renameHeads.forEach(r => {
        if (!r.ind) { r.ind = sheetModelFrom.head_key.indexOf(r.oldName); }
        if (r.ind < 1) { return; }
        sheetModelFrom.head_key[r.ind] = r.newName;
      });
    }
    // Logger.log(JSON.stringify(sheetModelFrom.head_key));

    let items = [...sheetModelFrom.getMap().values()].filter(obj => {
      return task.param.filter.function(obj, task.param.filter.fild, task.param.filter.prefix);
    });


    items = items.map(item => {
      item[task.param.filter.fild] = `${item[task.param.filter.fild]}`.slice(task.param.filter.prefix.length);
      return item;
    })


    let sheetModelTo = new MrClassSheetModel(task.param.sheetNames.to, context_to, task.param.rowConf.to);
    sheetModelTo.setItems(items);
    sheetModelАктуально.setMem(task.param.обновлено, new Date());

    let sheetFrom = sheetModelFrom.sheet;
    let sheetTo = sheetModelTo.sheet;
    let rules = sheetFrom.getConditionalFormatRules();
    let rdif = task.param.rowConf.to.body.first - task.param.rowConf.from.body.first;
    // if (rdif) {
    rules = rules.map(rule => {
      let retRule = rule.copy();
      let ranges = rule.getRanges().map(range => {
        /** todo */
        let cf = range.getCell(1, 1).getColumn();
        let rf = range.getCell(1, 1).getRow();
        let cl = range.getLastColumn();
        let rl = range.getLastRow();
        // Logger.log(`>-| r = ${rf} | rl = ${rl} | c = ${cf} | cl = ${cl} | getA1Notation =  ${range.getA1Notation()} `);
        if (rf == 2) {
          rf = rf + rdif;
          if (sheetFrom.getMaxRows() == rl) {
            rl = sheetTo.getMaxRows();
          } else {
            rl = rl + rdif;
          }
        }
        if (rl > sheetTo.getMaxRows()) {
          rl = sheetTo.getMaxRows();
        }
        if (rl < rf) { rl = rf; }
        let retRange = sheetTo.getRange(rf, cf, rl - rf + 1, cl - cf + 1);
        // retRange = sheetTo.getRange( range.getA1Notation());
        // Logger.log(`->| r = ${rf} | rl = ${rl} | c = ${cf} | cl = ${cl} | getA1Notation =  ${retRange.getA1Notation()} `);
        return retRange;
      });

      if (retRule.getBooleanCondition().getCriteriaType() == SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
        Logger.log(`${retRule.getBooleanCondition().getCriteriaType()}`);
        Logger.log(`${retRule.getBooleanCondition().getCriteriaValues()}`);
        let strRule = retRule.getBooleanCondition().getCriteriaValues()[0];
        // Logger.log(`>-|${Array.isArray(strRule)}`);
        let searchString = /([A-Z])([0-9])\b/g;
        // let searchString = /([A-Z])\b([0-9])\b/g;
        // let replaceString = 11;
        let replaceFunc = (match, letter, number) => {
          // Logger.log(`match-|${match}`);
          let num = Number(number);
          return letter + (num + rdif);
        };
        Logger.log(`>-|${strRule}`);
        strRule = strRule.replace(searchString, replaceFunc);
        Logger.log(`->|${strRule}`);
        retRule.whenFormulaSatisfied(strRule);
      }


      retRule.setRanges(ranges);
      return retRule;
    });


    // }


    sheetTo.clearConditionalFormatRules();
    if (rules.length) {
      sheetTo.setConditionalFormatRules(rules);
    }
    let mincol = Math.min(sheetTo.getMaxColumns(), sheetFrom.getMaxColumns());
    for (let c = 1; c <= mincol; c++) {
      sheetTo.setColumnWidth(c, sheetFrom.getColumnWidth(c));
    }

  }

  taskSyncПроектыВРаботеКопияДаных(task) {
    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    let sheetModelАктуально = new MrClassSheetModel(task.param.sheetNames.to, context_to);

    if ((new Date().getTime() - new Date(sheetModelАктуально.getMem(task.param.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    ) { Logger.log("да ещё актуальны"); return; }

    let sheetModelFrom = new MrClassSheetModel(task.param.sheetNames.from, context_from, task.param.rowConf.from);
    // Logger.log(JSON.stringify(sheetModelFrom.head_key));
    if (task.param.renameHeads.length != 0) {
      task.param.renameHeads.forEach(r => {
        if (!r.ind) { r.ind = sheetModelFrom.head_key.indexOf(r.oldName); }
        if (r.ind < 1) { return; }
        sheetModelFrom.head_key[r.ind] = r.newName;
      });
    }
    // Logger.log(JSON.stringify(sheetModelFrom.head_key));

    let items = [...sheetModelFrom.getMap().values()].filter(obj => {
      return task.param.filter.function(obj, task.param.filter.fild, task.param.filter.prefix);
    });


    items = items.map(item => {
      item[task.param.filter.fild] = `${item[task.param.filter.fild]}`.slice(task.param.filter.prefix.length);
      item["key"] = item[task.param.filter.fild];
      return item;
    })

    // items.forEach(item => {
    //   Logger.log(JSON.stringify(item));
    // });


    items.forEach(item => {
      // Logger.log(JSON.stringify(item));

      item["Безнал"] = undefined;
      item["К"] = undefined;
      item["А"] = undefined;
      item["Статус"] = undefined;
      item["План дата окончания"] = undefined;
      // К	А	Статус	План дата окончания
    });

    let sheetModelTo = new MrClassSheetModel(task.param.sheetNames.to, context_to, task.param.rowConf.to);
    // sheetModelTo.setItems(items);
    sheetModelTo.updateItems(items);

    this.taskSyncПроектыВРаботеФормулы(task);

    sheetModelАктуально.setMem(task.param.обновлено, new Date());

    //  копирование условного форматирования
    let sheetFrom = sheetModelFrom.sheet;
    let sheetTo = sheetModelTo.sheet;
    let rules = sheetFrom.getConditionalFormatRules();
    rules = rules.map(rule => {
      let retRule = rule.copy();
      let ranges = rule.getRanges().map(range => {
        return sheetTo.getRange(range.getA1Notation());
      });
      retRule.setRanges(ranges);
      return retRule;
    });
    sheetTo.clearConditionalFormatRules();
    if (rules.length) {
      sheetTo.setConditionalFormatRules(rules);
    }
    let mincol = Math.min(sheetTo.getMaxColumns(), sheetFrom.getMaxColumns());
    for (let c = 1; c <= mincol; c++) {
      sheetTo.setColumnWidth(c, sheetFrom.getColumnWidth(c));
    }

  }


  /** @param {DataTaskSyncОтслеживанияДляЮли} task*/
  taskSyncПроектыВРаботеФормулы(task) {


    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());

    // let sheetModelАктуально = new MrClassSheetModel(task.param.sheetNames.to, context_to);
    // if ((new Date().getTime() - new Date(sheetModelАктуально.getMem(task.param.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    // ) { Logger.log("да ещё актуальны"); return; }

    let sheetModelFrom = new MrClassSheetModel(task.param.sheetNames.from, context_from, task.param.rowConf.from);
    let sheetModelTo = new MrClassSheetModel(task.param.sheetNames.to, context_to, task.param.rowConf.to);

    let sheetFrom = sheetModelFrom.sheet;
    let sheetTo = sheetModelTo.sheet;


    // let headsColForFormulas = [
    //   "Безнал",
    //   "К",
    //   "А",
    //   "Статус",
    //   "План дата окончания",
    // ];



    let cols = [nr("B"), nr("C"), nr("D"), nr("E"), nr("F")];
    // let cols = [nr("B"), nr("E"), nr("F")];
    let rowFirst = task.param.rowConf.to.body.first;

    let rowLast = sheetTo.getLastRow();
    cols.forEach(col => {
      let formula = sheetFrom.getRange(rowFirst, col).getFormulaR1C1();
      // Logger.log(`Formula ${col} | ${formula}`);
      for (let row = rowFirst; row <= rowLast; row++) {
        let formulaNow = sheetTo.getRange(row, col).getFormula();
        if (`${formulaNow}`) { continue; }
        sheetTo.getRange(row, col).setFormula(formula);
      }
    });
  }



  /** @param {DataTaskSyncОтслеживанияДляЮли} task*/
  taskSyncПрочиеРасходы(task) {
    Logger.log(`${JSON.stringify(task)} `);

    let context_from = new MrContext(task.param.urls.from, getSettings());
    let context_to = new MrContext(task.param.urls.to, getSettings());
    let sheetModelАктуально = new MrClassSheetModel(task.param.sheetNames.to, context_to);

    if ((new Date().getTime() - new Date(sheetModelАктуально.getMem(task.param.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0
    ) { Logger.log("да ещё актуальны"); return; }

    let sheetFrom = context_from.getSheetByName(task.param.sheetNames.from);
    let heads = sheetFrom.getRange(task.param.range.head).getValues()[0];

    if (task.param.renameHeads.length != 0) {
      task.param.renameHeads.forEach(r => {
        if (!r.ind) { r.ind = heads.indexOf(r.oldName); }
        if (r.ind < 0) { return; }
        heads[r.ind] = r.newName;
      });
    }

    let vls = sheetFrom.getRange(task.param.range.body).getValues();
    let items = new Array();
    vls.forEach(v => {
      if (v.join("") == "") { return; }
      let item = new Object();
      heads.forEach((h, c, arr) => {
        if (!h) { return; }
        item[h] = v[c];
      });
      items.push(item);
    });


    items = items.filter(obj => {
      return task.param.filter.function(obj, task.param.filter.fild, task.param.filter.prefix);
    });

    items = items.map(item => {
      item[task.param.filter.fild] = `${item[task.param.filter.fild]}`.slice(task.param.filter.prefix.length);
      return item;
    })

    // items.forEach(item => {
    //   Logger.log(JSON.stringify(item));
    // });


    let sheetModelTo = new MrClassSheetModel(task.param.sheetNames.to, context_to, task.param.rowConf.to);
    sheetModelTo.setItems(items);
    sheetModelАктуально.setMem(task.param.обновлено, new Date());

  }

}



class TaskListSyncОтслеживанияДляЮли {
  constructor() {

    let sheetNames = {
      Рабочий_лист: "Рабочий лист",
      Текущие_доставки: "Текущие доставки",
      Проекты_в_работе: "Проекты в работе",
      Деньги_курьеры: "Деньги курьеры",
      Деньги_транспортные: "Деньги транспортные",
      Отгрузочные_документы: "Отгрузочные документы",
      Оплата_Логисту: "Оплата Логисту",
      Договора: "Договора",
      Технический_лист: "Технический лист",
      Прочие_расходы: "Прочие расходы",
      Экспорт_Финансы_для_Юлии: "Экспорт Финансы для Юлии",
      Импорт_Антон: "Импорт Антон",
      тех_сводка: "тех. сводка",
      Проекты_в_исполнении: "Проекты в исполнении",

    }




    let urls = {
      ОтслеживаниеПоставкиАнтонЛизякин: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
      ОтслеживаниеПоставкиЮлии: getSettings().urls.ОтслеживаниеПоставкиЮлии,
      РаскладК: getSettings().urls.РеестрК,
      РеестрЮ: getSettings().urls.РеестрЮ,
    };


    this.list = [

      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = task.param.urls.from;
        syncCopy.param.urls.to = task.param.urls.to;
        syncCopy.param.sheetNames.from = sheetNames.Технический_лист;
        syncCopy.param.sheetNames.to = sheetNames.Технический_лист;
        syncCopy.param.range.from = "A1:T";
        syncCopy.param.range.to = "A1:T";
        return syncCopy;
      })(),


      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = task.param.urls.from;
        syncCopy.param.urls.to = task.param.urls.to;
        syncCopy.param.sheetNames.from = sheetNames.тех_сводка;
        syncCopy.param.sheetNames.to = sheetNames.тех_сводка;
        syncCopy.param.range.from = "A1:AN";
        syncCopy.param.range.to = "A1:AN";
        return syncCopy;
      })(),



      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        // task.актуально_дней = 1 / 24 / 2/1000;
        task.param.sheetNames.from = sheetNames.Рабочий_лист;
        task.param.sheetNames.to = sheetNames.Рабочий_лист;
        //task.names.обновлено = `обновлено${ sheetNames.Проекты_в_работе } `;

        // task.param.renameHeads.push({ ind: nr("AL"), newName: "Проверка" });
        task.param.renameHeads.push({ ind: nr("AW"), newName: "Проверка обязательного заполнения полей" });
        task.param.filter = {
          fild: "Проект", prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),


      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Проекты_в_работе;
        task.param.sheetNames.to = sheetNames.Проекты_в_работе;
        //task.names.обновлено = `обновлено${ sheetNames.Проекты_в_работе } `;
        task.name = ExeTask_SyncОтслеживанияДляЮли.prototype.taskSyncПроектыВРаботеКопияДаных.name;
        task.param.filter = {
          fild: `"№"`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),


      // (() => {
      //   let task = new DataTaskSyncОтслеживанияДляЮли();
      //   task.param.sheetNames.from = sheetNames.Проекты_в_работе;
      //   task.param.sheetNames.to = sheetNames.Проекты_в_работе;
      //   task.name = ExeTask_SyncОтслеживанияДляЮли.prototype.taskSyncПроектыВРаботеФормулы.name;
      //   task.param.rowConf = {
      //     from: {
      //       head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
      //       body: { first: 11, last: 11, },
      //     },
      //     to: {
      //       head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
      //       body: { first: 11, last: 11, },
      //     }
      //   };
      //   return task;
      // })(),



      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Текущие_доставки;
        task.param.sheetNames.to = sheetNames.Текущие_доставки;
        //task.names.обновлено = `обновлено${ sheetNames.Текущие_доставки } `;

        task.param.filter = {
          fild: `Проект`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 1, key: 1, },
            body: { first: 2, last: 2, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),


      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Деньги_курьеры;
        task.param.sheetNames.to = sheetNames.Деньги_курьеры;
        //task.names.обновлено = `обновлено${ sheetNames.Деньги_курьеры } `;
        task.param.filter = {
          fild: `Проект`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 1, key: 1, },
            // head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 2, last: 2, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),


      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Деньги_транспортные;
        task.param.sheetNames.to = sheetNames.Деньги_транспортные;
        // this.names.обновлено = `обновлено${ sheetNames.Деньги_транспортные } `;
        task.param.filter = {
          fild: `Проект`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 1, key: 1, },
            //  head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 2, last: 2, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),



      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Отгрузочные_документы;
        task.param.sheetNames.to = sheetNames.Отгрузочные_документы;
        //task.names.обновлено = `обновлено${ sheetNames.Отгрузочные_документы } `;
        // Автоматическая Проверка
        task.param.renameHeads.push({ ind: nr("P"), newName: "Автоматическая Проверка" });
        task.param.filter = {
          fild: `Проект`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 1, key: 1, },
            //  head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 2, last: 2, },
          },
          // to: {
          //   head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
          //   body: { first: 11, last: 11, },
          // }

          to: {
            head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }

        };
        return task;
      })(),



      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Оплата_Логисту;
        task.param.sheetNames.to = sheetNames.Оплата_Логисту;
        //task.names.обновлено = `обновлено${ sheetNames.Оплата_Логисту } `;
        task.param.filter = {
          fild: `Проект`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 1, key: 1, },
            //    head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 2, last: 2, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),

      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Договора;
        task.param.sheetNames.to = sheetNames.Договора;
        //task.names.обновлено = `обновлено${ sheetNames.Договора } `;

        task.param.renameHeads.push({ ind: nr("U"), newName: "Автоматическая Проверка" });
        task.param.renameHeads.push({ oldName: "Юр. лицо Продавец", newName: "Юр лицо Продавец" });
        task.param.renameHeads.push({ oldName: "Юр. лицо Покупатель", newName: "Юр лицо Покупатель" });

        task.param.filter = {
          fild: `Проект №`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 1, key: 1, },
            //   head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 2, last: 2, },
          },
          // to: {
          //   head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
          //   body: { first: 11, last: 11, },
          // },
          to: {
            head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }


        };
        return task;
      })(),

      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Экспорт_Финансы_для_Юлии;
        task.param.sheetNames.to = sheetNames.Прочие_расходы;
        task.param.urls.from = getSettings().urls.РеестрК; // Расклад К
        task.name = ExeTask_SyncОтслеживанияДляЮли.prototype.taskSyncПрочиеРасходы.name;
        task.param.renameHeads.push({ oldName: "Кто оплатил", newName: "Кто оплатил" });

        task.param.filter = {
          fild: `Проект`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };

        task.param.range = {
          head: "S1:Y1",
          body: "S2:Y",
        };
        task.param.rowConf = {
          from: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          },
          to: {
            head: { first: 1, last: 10, key: 10, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),


      (() => {
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = urls.ОтслеживаниеПоставкиЮлии;
        syncCopy.param.urls.to = urls.РеестрЮ;
        syncCopy.param.sheetNames.from = sheetNames.Проекты_в_работе;
        syncCopy.param.sheetNames.to = sheetNames.Импорт_Антон;
        syncCopy.param.range.from = "A11:F";
        syncCopy.param.range.to = "A2:F";
        syncCopy.param.filter_empty = 0;
        syncCopy.note = true;
        return syncCopy;
      })(),

      (() => {
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = urls.ОтслеживаниеПоставкиЮлии;
        syncCopy.param.urls.to = urls.РеестрЮ;
        syncCopy.param.sheetNames.from = sheetNames.Договора;
        syncCopy.param.sheetNames.to = sheetNames.Импорт_Антон;
        syncCopy.param.range.from = "A10:AK";
        syncCopy.param.range.to = "AA1:BK";
        syncCopy.param.filter_empty = 0;
        syncCopy.note = true;
        return syncCopy;
      })(),



      (() => {
        let task = new DataTaskSyncОтслеживанияДляЮли();
        task.param.sheetNames.from = sheetNames.Проекты_в_исполнении;
        task.param.sheetNames.to = sheetNames.Проекты_в_исполнении;
        // Автоматическая Проверка
        // task.param.renameHeads.push({ ind: nr("P"), newName: "Автоматическая Проверка" });
        task.param.filter = {
          // fild: `"№"`, prefix: "Ю-",
          fild: `1`, prefix: "Ю-",
          function: ((obj, fild, prefix) => {
            if (`${obj[fild]}`.slice(0, prefix.length).toUpperCase() === `${prefix}`.toUpperCase()) { return true; }
            return false;
          }),
        };
        task.param.rowConf = {
          from: {
            // head: { first: 1, last: 1, key: 1, },
            head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          },
          to: {
            head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
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


