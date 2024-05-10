class НаборЗадачь {
  constructor() {
    /** @private */
    this.map = new Map();


  }

  getНабор(Имя) {
    if (!this.map.has(Имя)) {
      this.map.set(Имя, this[Имя]());
    }
    return this.map.get(Имя);
  }

  /** @private */
  НаборЗадачьРеестрК() {

    let urls = {
      ПромежуткиРеестр: getSettings().urls.ПромежуткиРеестр,
      РаскладК: getSettings().urls.РеестрК,
    };


    let reestr_K = new DataTaskExportProgectListFromReestr();
    // reestr_K.param.Ответственный = "Наталья";
    reestr_K.param.urls.from = urls.РаскладК;
    reestr_K.param.sheetNames.from = "Исполнение";
    reestr_K.param.urls.to = urls.ПромежуткиРеестр;
    reestr_K.param.sheetNames.to = "Проекты в Исполнение";


    let taskGetProjects = new DataTaskGetExportValuesFromProjects();
    taskGetProjects.param.sheetNames.to = "Проекты в Исполнение";
    taskGetProjects.param.urls.to = urls.ПромежуткиРеестр;

    let list = [
      reestr_K,
      taskGetProjects,

      // (() => {
      //   let ret = new DataTaskSyncReestrПроцентОплаты();
      //   ret.param.urls.from = urls.ПромежуткиРеестр;
      //   ret.param.urls.to = urls.МусорРеестрК;
      //   return ret;
      // })(),

      (() => {
        let ret = new DataTaskSyncReestrПроцентОплаты();
        ret.param.urls.from = urls.ПромежуткиРеестр;
        ret.param.urls.to = urls.ПромежуткиРеестр;
        return ret;
      })(),


    ];

    return list;



  }


  /** @private */
  НаборЗадачь_РеестрК_ИмпортАнтон() {
    let sheetNames = {
      Проекты_в_работе: "Проекты в работе",
      Импорт_Антон: "Импорт Антон",
      // Импорт_Антон: "Импорт Антон (копия)",
      Процент_оплаты: "Процент оплаты",
      Договора: "Договора",
      Технический_лист: "Технический лист",
      Технический: "Технический",
      Словари: "Словари",
    }


    let urls = {
      ОтслеживаниеПоставкиАнтонЛизякин: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
      ПромежуткиРеестр: getSettings().urls.ПромежуткиРеестр,
      РаскладК: getSettings().urls.РеестрК,
      Промежуточный_экспорт_таблица_логистики_Антон_Лизякин: getSettings().urls.ПромежуткиЛогиста,
      Проекты: getSettings().urls.Проекты,
    };



    let list = [
      (() => {
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = urls.ОтслеживаниеПоставкиАнтонЛизякин;
        syncCopy.param.urls.to = urls.РаскладК;
        syncCopy.param.sheetNames.from = sheetNames.Проекты_в_работе;
        syncCopy.param.sheetNames.to = sheetNames.Импорт_Антон;
        syncCopy.param.range.from = "A10:F";
        syncCopy.param.range.to = "A1:F";
        syncCopy.note = true;
        return syncCopy;

      })(),

      // (() => {
      //   let syncCopy = new DataTaskSyncCopyRange();
      //   syncCopy.param.urls.from = urls.ПромежуткиРеестр;
      //   syncCopy.param.urls.to = urls.РаскладК;
      //   syncCopy.param.sheetNames.from = sheetNames.Процент_оплаты;
      //   syncCopy.param.sheetNames.to = sheetNames.Импорт_Антон;
      //   syncCopy.param.range.from = "C10:U";
      //   syncCopy.param.range.to = "G1:Z";
      //   return syncCopy;
      // })(),

      (() => {  // Технический лист'!W1:AO200
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = urls.ОтслеживаниеПоставкиАнтонЛизякин;
        syncCopy.param.urls.to = urls.РаскладК;
        syncCopy.param.sheetNames.from = sheetNames.Технический_лист;
        syncCopy.param.sheetNames.to = sheetNames.Импорт_Антон;
        syncCopy.param.range.from = "W1:AO";
        syncCopy.param.range.to = "G1:Z";
        syncCopy.note = true;
        return syncCopy;
      })(),


      (() => {
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = urls.ОтслеживаниеПоставкиАнтонЛизякин;
        syncCopy.param.urls.to = urls.РаскладК;
        syncCopy.param.sheetNames.from = sheetNames.Договора;
        syncCopy.param.sheetNames.to = sheetNames.Импорт_Антон;
        syncCopy.param.range.from = "A1:AI";
        syncCopy.param.range.to = "AA1:BI";
        syncCopy.param.filter_empty = 0;
        syncCopy.note = true;
        return syncCopy;
      })(),


      // (() => {
      //   let syncCopy = new DataTaskSyncCopyRange();
      //   syncCopy.param.urls.from = urls.РаскладК;
      //   syncCopy.param.urls.to = urls.Промежуточный_экспорт_таблица_логистики_Антон_Лизякин;
      //   syncCopy.param.sheetNames.from = sheetNames.Технический;
      //   syncCopy.param.sheetNames.to = sheetNames.Технический;
      //   syncCopy.param.range.from = "AD1:AE";
      //   syncCopy.param.range.to = "AD1:AE";
      //   return syncCopy;
      // })(),


      (() => {
        let syncCopy = new DataTaskSyncCopyRange();
        syncCopy.param.urls.from = urls.РаскладК;
        syncCopy.param.urls.to = urls.Проекты;
        syncCopy.param.sheetNames.from = sheetNames.Технический;
        syncCopy.param.sheetNames.to = sheetNames.Словари;
        syncCopy.param.range.from = "AD1:AE";
        syncCopy.param.range.to = "D10:E10";
        return syncCopy;
      })(),


      // ((theTask) => {
      //   theTask.param.urls.from = getSettings().urls.Письма;
      //   theTask.param.urls.to = urls.РаскладК;
      //   theTask.param.sheetNames.from = getSettings().sheetNames.ОтчетыПроектов;
      //   theTask.param.sheetNames.to = getSettings().sheetNames.ОтчетыПроектов;
      //   theTask.param.range.from = "A1:Z";
      //   theTask.param.range.to = "A1:Z1";
      //   theTask.note = true;
      //   return theTask;
      // })(new DataTaskSyncCopyRange()),



    ];
    //  return [];
    return list;
  }



  /** @private */
  НаборЗадачь_РеестрК_ИмпортИсполнениеИзРеестрПроцедурЮ() {
    let sheetNames = {
      // Проекты_в_работе: "Проекты в работе",
      // Импорт_Антон: "Импорт Антон",
      // // Импорт_Антон: "Импорт Антон (копия)",
      // Процент_оплаты: "Процент оплаты",
      // Договора: "Договора",
      // Технический_лист: "Технический лист",
      Импорт_Юля: "Импорт Юля",
      Исполнение: "Исполнение",
      Отгрузочные_документы_Экспорт: "Отгрузочные документы Экспорт",
      Оплаты: "Оплаты",

    }


    let urls = {
      РеестрПроцедурЮ: getSettings().urls.РеестрЮ,
      Промежуточный_экспорт_таблица_логистики_Антон_Лизякин: getSettings().urls.ПромежуткиЛогиста,
      РаскладК: getSettings().urls.РеестрК,

    };


    let list = [


      (() => {
        let syncCopy = new DataTaskSyncCopyИсполнениеЮ();
        return syncCopy;
      })(),




    ];
    //  return [];
    return list;
  }





  /** @private */
  НаборЗадачьСводная() {
    let sheetNames = {

      Реестр: "Реестр",
      Исполнение: "Исполнение",
      Оплаты: "Оплаты",
      Сводная: "Сводная",
      ОтчетыПроектов: getSettings().sheetNames.ОтчетыПроектов,
    }


    let urls = {
      РаскладК: getSettings().urls.РеестрК,
      РеестрЮ: getSettings().urls.РеестрЮ,
      Проекты: getSettings().urls.Проекты,
      Письма: getSettings().urls.Письма,

    };


    // DataTaskJoinProgectInfo;


    let list = [



      (() => {
        let task = new DataTaskJoinProgectInfo();
        task.name = ExeTask_JoinProgectInfo.prototype.taskGetРеестр.name;
        task.param.urls.from = urls.РеестрЮ;
        task.param.urls.to = urls.Проекты;
        task.param.sheetNames.from = sheetNames.Реестр;
        task.param.sheetNames.to = sheetNames.Сводная;

        task.param.prefix = "Ю-";
        task.param.fild = "Реестр";
        task.param.mem = "РеестрЮ";
        task.param.rows = {
          headFirst: 1,
          bodyFirst: 2,
        };
        return task;
      })(),


      (() => {
        let task = new DataTaskJoinProgectInfo();
        task.name = ExeTask_JoinProgectInfo.prototype.taskGetРеестр.name;
        task.param.urls.from = urls.РаскладК;
        task.param.urls.to = urls.Проекты;
        task.param.sheetNames.from = sheetNames.Реестр;
        task.param.sheetNames.to = sheetNames.Сводная;

        task.param.prefix = "";
        task.param.fild = "Реестр";
        task.param.mem = "РеестрK";
        task.param.rows = {
          headFirst: 1,
          bodyFirst: 2,
        };
        return task;
      })(),






      (() => {
        let task = new DataTaskJoinProgectInfo();
        task.name = ExeTask_JoinProgectInfo.prototype.taskGetИсполнение.name;
        task.param.urls.from = urls.РеестрЮ;
        task.param.urls.to = urls.Проекты;
        task.param.sheetNames.from = sheetNames.Исполнение;
        task.param.sheetNames.to = sheetNames.Сводная;

        task.param.prefix = "Ю-";
        task.param.fild = "Исполнение";
        task.param.mem = "ИспЮ";
        task.param.rows = {
          headFirst: 2,
          bodyFirst: 3,
        };
        return task;
      })(),



      (() => {
        let task = new DataTaskJoinProgectInfo();
        task.name = ExeTask_JoinProgectInfo.prototype.taskGetИсполнение.name;
        task.param.urls.from = urls.РаскладК;
        task.param.urls.to = urls.Проекты;
        task.param.sheetNames.from = sheetNames.Исполнение;
        task.param.sheetNames.to = sheetNames.Сводная;

        task.param.prefix = "";
        task.param.fild = "Исполнение";
        task.param.mem = "ИспK";
        task.param.rows = {
          headFirst: 2,
          bodyFirst: 3,
        };
        return task;
      })(),





      // taskSyncОтчетыПроектовИзСводнойвРеестр

      (() => {
        let task = new DataTaskJoinProgectInfo();
        task.актуально_дней = 1 / 24 / 2;
        task.name = ExeTask_JoinProgectInfo.prototype.taskSyncОтчетыПроектовИзСводнойвРеестр.name;
        task.param.urls.from = urls.Письма;
        task.param.urls.to = urls.РеестрЮ;
        task.param.sheetNames.from = sheetNames.ОтчетыПроектов;
        task.param.sheetNames.to = sheetNames.ОтчетыПроектов;
        task.param.prefix = "Ю-";
        task.param.fild = "Проект";
        return task;
      })(),

      // taskSyncОтчетыПроектовИзСводнойвРеестр

      (() => {
        let task = new DataTaskJoinProgectInfo();
        task.актуально_дней = 1 / 24 / 1;
        task.name = ExeTask_JoinProgectInfo.prototype.taskSyncОтчетыПроектовИзСводнойвРеестр.name;
        task.param.urls.from = urls.Письма;
        task.param.urls.to = urls.РаскладК;
        task.param.sheetNames.from = sheetNames.ОтчетыПроектов;
        task.param.sheetNames.to = sheetNames.ОтчетыПроектов;
        task.param.fild = "Проект";
        return task;
      })(),



      ((theTask) => {
        theTask.name = ExeTask_JoinProgectInfo.prototype.taskSyncОтчетыПроектовСинхронизцияШапки.name;
        theTask.param.urls.from = urls.РаскладК;
        theTask.param.urls.tos = [urls.Письма, urls.РеестрЮ];
        theTask.param.sheetNames.from = sheetNames.ОтчетыПроектов;
        theTask.param.sheetNames.to = sheetNames.ОтчетыПроектов;
        return theTask;
      })(new DataTaskJoinProgectInfo()),

    ];
    return list;
  }



  /** @private */
  НаборЗадачьОтслеживанияДляЮли() {
    return (new TaskListSyncОтслеживанияДляЮли()).getTaskList();
  }





  /** @private */
  НаборЗадачьClearSelectRowInУПД_ДОК() {

    let list = [

      (() => {
        let task = new DataTaskClearSelectRowInУПД_ДОК();
        return task;
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

    ];
    return list;
  }





  /** @private */
  НаборЗадачьТест() {

    let sheetNames = {

      Реестр: "Реестр",
      Исполнение: "Исполнение",
      Оплаты: "Оплаты",
      Сводная: "Сводная",
      ОтчетыПроектов: getSettings().sheetNames.ОтчетыПроектов,
    }


    let urls = {
      РаскладК: getSettings().urls.РеестрК,
      РеестрЮ: getSettings().urls.РеестрЮ,
      Проекты: getSettings().urls.Проекты,
      Письма: getSettings().urls.Письма,

    };

    let list = [
      ((theTask) => {
        theTask.name = ExeTask_JoinProgectInfo.prototype.taskSyncОтчетыПроектовСинхронизцияШапки.name;
        theTask.param.urls.from = urls.РаскладК;
        theTask.param.urls.tos = [urls.Письма, urls.РеестрЮ];
        theTask.param.sheetNames.from = sheetNames.ОтчетыПроектов;
        theTask.param.sheetNames.to = sheetNames.ОтчетыПроектов;
        return theTask;
      })(new DataTaskJoinProgectInfo()),

    ];
    return list;

  }
  /** @private */
  НаборЗадачь_SyncУПД_ДОК() {

    let list = [

      // (() => {
      //   let syncУПД_ДОК = new DataTaskSyncУПД_ДОК();
      //   syncУПД_ДОК.param.urls = {
      //     from: getSettings().urls.ПромежуткиЛогиста,
      //     to: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
      //   };
      //   syncУПД_ДОК.param.sheetNames = {
      //     from: "Проекты в Исполнение",
      //     to_УПД: "Отгрузочные документы",
      //     to_ДОК: "Договора",
      //     to_Рабочий_Лист: "Рабочий Лист",
      //   }
      //   return syncУПД_ДОК;
      // })(),


      // DataTaskUpdateFilds
      (() => {
        let task = new DataTaskUpdateFilds();
        task.param.urls = {
          from: getSettings().urls.Документы,
          to: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
          // to: "https://docs.google.com/spreadsheets/d/1???????????????????????????????????????????/edit", // test
        };

        task.param.sheetNames = {
          from: getSettings().sheetNames.СводнаяПоДоговорам,
          to: getSettings().sheetNames.Договора,
        }


        task.param.fildCouples.key.from = "key";
        task.param.fildCouples.key.to = "ID";

        // task.param.fildCouples["Ошибки при заполнение договора во внешней"] = {
        //   from: "Ошибки при заполнение договора во внешней",
        //   to: "Ошибки при заполнение договора во внешней",
        // }

        // // task.param.fildCouples["Ошибки при заполнение договора во внешней"] = {
        // //   from: "Ошибки при заполнение договора во внешней",
        // //   to: "Ошибки при заполнение договора во внешней",
        // // }
        // // Присутствует в таблице проекта
        // task.param.fildCouples["Присутствует в таблице проекта"] = {
        //   from: "Присутствует в таблице проекта",
        //   to: "Присутствует в таблице проекта",
        // }

        task.param.syncFilds = [
          "Ошибки при заполнение договора во внешней",
          "Присутствует в таблице проекта",
          "ДатаПолученияСвединийИзПроекта",
          "для проверки расхождения данных с логистом"
        ];

        task.param.rowConf = {
          to: {
            head: { first: 1, last: 1, key: 1, },
            body: { first: 2, last: 2, },
          },
          from: {
            head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),




      (() => {
        let task = new DataTaskUpdateFilds();
        task.param.urls = {
          from: getSettings().urls.Документы,
          to: getSettings().urls.ОтслеживаниеПоставкиАнтонЛизякин,
            };

        task.param.sheetNames = {
          from: getSettings().sheetNames.СводнаяПоУПД,
          to: getSettings().sheetNames.ОтгрузочныеДокументы,
        }


        task.param.fildCouples.key.from = "key";
        task.param.fildCouples.key.to = "ID";

        task.param.syncFilds = [
          "Ошибки при заполнение УПД во внешней",
          "Присутствует в таблице проекта",
          "ДатаПолученияСвединийИзПроекта",
          "для проверки расхождения данных с логистом"
        ];

        task.param.rowConf = {
          to: {
            head: { first: 1, last: 1, key: 1, },
            body: { first: 2, last: 2, },
          },
          from: {
            head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
            body: { first: 11, last: 11, },
          }
        };
        return task;
      })(),
    ];
    return list;

  }


}






function ВыполненитьЗадачи(trigerrInfo = "ВыполненитьЗадачи", Имя) {
  Logger.log(JSON.stringify({ Имя, trigerrInfo }));
  triggerLookCache();
  // return;
  let tasks = (() => { try { return new НаборЗадачь().getНабор(Имя); } catch (err) { mrErrToString(err); return new Array() } })();

  let mrTaskExe = new MrTaskExecutor(getContext(), tasks, getContext().getScriptCache(), Имя);
  try {
    mrTaskExe.triggerПродолжитьВыполнениеЗадачь();
  } catch (err) { mrErrToString(err); }

  triggerLookCache();
}


function ВыполненитьМасивЗадач(trigerrInfo = "ВыполненитьМасивЗадач", tasks, Имя,) {
  Logger.log(JSON.stringify({ Имя, trigerrInfo, tasks }));
  triggerLookCache();
  // return;
  // let tasks = (() => { try { return new НаборЗадачь().getНабор(Имя); } catch (err) { mrErrToString(err); return new Array() } })();
  if (!Array.isArray(tasks)) { return; }

  let mrTaskExe = new MrTaskExecutor(getContext(), tasks, getContext().getScriptCache(), Имя);
  try {
    mrTaskExe.triggerПродолжитьВыполнениеЗадачь();
  } catch (err) { mrErrToString(err); }

  triggerLookCache();
}

