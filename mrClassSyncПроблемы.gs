class DataTaskSyncПроблемы {
  constructor() {
    /** @constant */
    this.date = new Date();
    this.exe_class = ExeTask_SyncПроблемы.name;
    /** @constant */
    this.name = ExeTask_SyncПроблемы.prototype.taskSyncПроблемы.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      projNumber: undefined,
      urls: {
        proj: "",
        central: getSettings().urls.ВопросыПроблемыПроектов,
        settings: getSettings().urls.ВопросыПроблемыПроектов,
      },

      sheetNames: {
        proj: getSettings().sheetNames.Проблемы,
        central: getSettings().sheetNames.Проблемы,
        settings: getSettings().sheetNames.Настройки,
      },

      range: {
        proj: {
          head: { first: 1, last: 1, key: 1, },
          body: { first: 2, last: 2, },
        },
        // central: {
        //   head: { first: 1, last: 10, key: 9, type: 8, ru_name: 10, changed: 7, off: 9, mem: 7, },
        //   body: { first: 11, last: 11, },
        // },
        central: {
          head: { first: 1, last: 1, key: 1, },
          body: { first: 2, last: 2, },
        },
        settings: {
          head: { first: 1, last: 1, key: 1, },
          body: { first: 2, last: 2, },
        },

      },
    };
  }
}

function getDataTaskSyncПроблемы() {
  return new DataTaskSyncПроблемы();
}



class ExeTask_SyncПроблемы extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    this.myCacheName = ExeTask_SyncПроблемы.name;
    this.names = {
      обновлено: "обновлено",
    }

  }

  /** @param {DataTaskSyncПроблемы} task */
  taskSyncПроблемы(task) {

    Logger.log(`vvvvvvvvvvvvvvvvvvvvvvv-taskSyncПроблемы-vvvvvvvvvvvvvvvvvvvvvvvv`);
    let context = {
      proj: new MrContext(task.param.urls.proj, new Object()),
      central: new MrContext(task.param.urls.central, new Object()),
      // settings: new MrContext(task.param.urls.settings, new Object()),
    }

    let sheetModel = {
      project: new MrClassSheetModel(task.param.sheetNames.proj, context.proj, task.param.range.proj),
      central: new MrClassSheetModel(task.param.sheetNames.central, context.central, task.param.range.central),
      // settings: new MrClassSheetModel(task.param.sheetNames.settings, context.settings, task.param.range.settings),
    }

    let mapProjectПроблемы = sheetModel.project.getMap();
    // let mapCentralВопросы = sheetModel.central.getMap();
    // let mapSettings = sheetModel.settings.getMap();


    let items = [...mapProjectПроблемы.values()].map(item => {
      // № проекта	
      item["№ проекта"] = task.param.projNumber;
      item["Ссылка на таблицу проекта"] = task.param.urls.proj;
      return item;
    });

    // Logger.log(items.length);

    // items.forEach(item => {
    //   Logger.log(JSON.stringify(item, null, 2));
    // });


    sheetModel.central.updateItemsObj(items);


    Logger.log(`^^^^^^^^^^^^^^^^^^^^-taskSyncПроблемы-^^^^^^^^^^^^^^^^^^^`);
  }



}

