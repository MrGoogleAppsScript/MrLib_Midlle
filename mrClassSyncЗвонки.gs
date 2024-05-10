class DataTaskSyncЗвонки {
  constructor() {
    /** @constant */
    this.date = new Date();
    this.exe_class = ExeTask_SyncЗвонки.name;
    /** @constant */
    this.name = ExeTask_SyncЗвонки.prototype.taskSyncЗвонки.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      urls: {
        proj: "",
     
        central: "https://docs.google.com/spreadsheets/d/1/edit",
        settings: "https://docs.google.com/spreadsheets/d/1/edit",
      },

      sheetNames: {
        proj: getSettings().sheetNames.Звонки,
        central: getSettings().sheetNames.Звонки,
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

function getDataTaskSyncЗвонки() {
  return new DataTaskSyncЗвонки();
}



class DataTaskOnEditЗвонки {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_SyncCopyRange.name;
    /** @constant */
    this.name = ExeTask_SyncCopyRange.prototype.taskExportSyncCopyRange.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      edit: {},

      urls: {
        proj: "",

      },

      sheetNames: {
        proj: "",

      },

      range: {
        proj: "",

      },

    };

  }
}




class ExeTask_SyncЗвонки extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_SyncLogist.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",
    }

  }

  /** @param {DataTaskSyncЗвонки} task */
  taskSyncЗвонки(task) {
    // return;
    // Logger.log(`taskSyncЗвонки  - ${JSON.stringify(task)} `);
    Logger.log(`vvvvvvvvvvvvvvvvvvvvvvv-taskSyncЗвонки-vvvvvvvvvvvvvvvvvvvvvvvv`);
    let context = {
      proj: new MrContext(task.param.urls.proj, new Object()),
      central: new MrContext(task.param.urls.central, new Object()),
      settings: new MrContext(task.param.urls.settings, new Object()),
    }

    // let sheet_proj = context_proj.getSheetByName(task.param.sheetNames.proj);
    // let sheet_central = context_central.getSheetByName(task.param.sheetNames.central);

    let sheetModel = {
      project: new MrClassSheetModel(task.param.sheetNames.proj, context.proj, task.param.range.proj),
      central: new MrClassSheetModel(task.param.sheetNames.central, context.central, task.param.range.central),
      settings: new MrClassSheetModel(task.param.sheetNames.settings, context.settings, task.param.range.settings),
    }


    let ТехническийлистСтатусыЗакрытия = context.central.getSheetByName("Технический лист").getRange("A2:B").getValues()
      .filter(v => {
        if (!v[0]) { return false; }
        if (v[1] !== true) { return false; }
        return true;
      }).map(v => {
        return v[0];
      })

    // Logger.log(ТехническийлистСтатусыЗакрытия);

    let mapProjectЗвонки = sheetModel.project.getMap();
    let mapCentralЗвонки = sheetModel.central.getMap();
    let mapSettings = sheetModel.settings.getMap();


    // mapSettings.forEach((item,key) => Logger.log(`  ${key} - ${JSON.stringify(item, undefined, 2)}`));


    let changed_key = new Array(sheetModel.central.head_key.length);
    let changed_from_project = mapSettings.get("project");
    let changed_from_central = mapSettings.get("central");

    sheetModel.central.head_key.forEach((h, i) => {
      changed_key[i] = (() => {
        if (changed_from_project[h] === true) { return true; }
        if (changed_from_central[h] === true) { return true; }
        return false;
      })()
    });

    sheetModel.central.changed_key = changed_key;





    // let filterEmtyFild = [
    //   // "№",
    //   "С какой почты писали",
    //   "Кому писали",
    //   "Комментарий по звонку",
    //   "Дата след контакта",
    // ];

    let filterEmtyFild = new Array();
    let filterEmtyObj = mapSettings.get("filterEmty");
    for (let k in filterEmtyObj) {
      if (filterEmtyObj[k] === true) {
        filterEmtyFild.push(k);
      }
    }


    // Logger.log(` arrЗвонки ${JSON.stringify(sheetModelProject, undefined, 2)}`);

    // mapProjectЗвонки.forEach((v, k) => {
    //   Logger.log(`k = ${k} v= ${JSON.stringify(v,undefined,2)}`);
    // });

    let arrЗвонки = [...mapProjectЗвонки.values()].filter(item => {
      // Logger.log(` arrЗвонки ${JSON.stringify(item)}`);

      if (filterEmtyFild.map(k => { return item[k] }).join("") == "") { return false; }
      if (item["Ссылка на таблицу проекта"] != task.param.urls.proj) { return false; }
      if (ТехническийлистСтатусыЗакрытия.includes(item["Текущий статус"])) {
        return false;
      }

      if (`${item["Перечень товаров по которым требуется уточнение"]}` == "") { return false; }
      // Текущий статус
      return true;
    });

    /** @typedef  SyncЗвонки
     * @property {Object} project
     * @property {Object} central
     * @property {Object} syncObj
     * @property {boolean} emtyCentral
     * @property {boolean} emtyProject
     * @property {boolean} changCentral
     * @property {boolean} changProject
     * @property {object} centralSnapshot
     * @property {object} projectSnapshot
    */

    /**  @type {[SyncЗвонки]}  */
    let syncMap = new Map();
    arrЗвонки.forEach(itemЗвонк => {

      let project = itemЗвонк;
      let central = mapCentralЗвонки.get(project.key);

      syncMap.set(project.key, { project, central, });

      // Logger.log(` arrЗвонки ${JSON.stringify({ project, central, })}`);

    });

    Logger.log(`syncMap size = ${syncMap.size} `);

    arrЗвонки = [...mapCentralЗвонки.values()].filter(item => {
      // Logger.log(`1 arrЗвонки ${JSON.stringify(task.param.urls)}`);
      // Logger.log(`1 arrЗвонки ${JSON.stringify(item)}`);
      if (item["Ссылка на таблицу проекта"] != task.param.urls.proj) { return false }

      // Logger.log(`2 arrЗвонки ${JSON.stringify(item)}`);  
      if (filterEmtyFild.map(k => { return item[k] }).join("") == "") { return false; }
      // Logger.log(`3 arrЗвонки ${JSON.stringify(item)}`);
      return true;
    });

    arrЗвонки.forEach(itemЗвонк => {
      let central = itemЗвонк;
      if (syncMap.has(central.key)) { return; }
      let project = mapCentralЗвонки.get(central.key);
      syncMap.set(central.key, { project, central, });
      // Logger.log(` arrЗвонки ${JSON.stringify({ project, central, })}`);

    });






    // Logger.log(` Требуется синхронизация ${JSON.stringify(syncArr)}`);

    let mapUpdate = new Map();


    syncMap.forEach(/** @type {SyncЗвонки} */syncЗвонк => {

      let hasChanged = false;

      // Logger.log(` Требуется? синхронизация ${JSON.stringify(syncЗвонк)}`);

      if (!syncЗвонк.central) { syncЗвонк.emtyCentral = true; syncЗвонк.central = new Object(); hasChanged = true; }
      if (!syncЗвонк.project) { syncЗвонк.emtyProject = true; syncЗвонк.project = new Object(); hasChanged = true; }

      if (hasChanged == false) {
        if (sheetModel.central.isChanged(syncЗвонк.central)) { syncЗвонк.changCentral = true; hasChanged = true; };
        if (sheetModel.central.isChanged(syncЗвонк.project)) { syncЗвонк.changProject = true; hasChanged = true; };
      }
      if (hasChanged == false) { return; }



      // Logger.log(`ДА Требуется синхронизация ${JSON.stringify(syncЗвонк, null, 2)}`);

      // makeSnapshot 
      syncЗвонк.syncObj = new Object();

      syncЗвонк.centralSnapshot = (() => {
        if (!syncЗвонк.central) { return sheetModel.central.makeSnapshot(syncЗвонк.central); } return undefined;
      });
      syncЗвонк.projectSnapshot = (() => {
        if (!syncЗвонк.project) { return sheetModel.central.makeSnapshot(syncЗвонк.project); } return undefined;
      });


      if (syncЗвонк.emtyProject) {
        syncЗвонк.syncObj = syncЗвонк.central;
      }

      if (syncЗвонк.emtyCentral) {
        syncЗвонк.syncObj = syncЗвонк.project;
      }

      // Logger.log(`Вот Требуется синхронизация ${JSON.stringify(syncЗвонк, null, 2)}`);
      //  определяем кто последниий изменился
      if (syncЗвонк.changCentral && syncЗвонк.changProject) {
        syncЗвонк.changCentral = true;
        syncЗвонк.changProject = false;
        Logger.log(`ДА Требуется синхронизация есть изменения везде`);
      }

      if (syncЗвонк.changCentral) {
        syncЗвонк.syncObj = syncЗвонк.central;
      }


      if (syncЗвонк.changProject) {
        syncЗвонк.syncObj = syncЗвонк.project;
      }

      syncЗвонк.syncObj.last = sheetModel.central.makeSnapshot(syncЗвонк.syncObj);
      mapUpdate.set(syncЗвонк.syncObj.key, syncЗвонк.syncObj);

      // Logger.log(`-------------------------------------------------------------`);
    });
    Logger.log(`-------------------------------------------------------------`);

    Logger.log(`mapUpdateCentral size = ${mapUpdate.size} `);


    let head_type_settings = mapSettings.get("type");
    let head_type = new Array(sheetModel.central.head_key.length);
    sheetModel.central.head_key.forEach((h, i) => {
      head_type[i] = (() => { return fl_str(head_type_settings[h]); })();
    });



    // mapUpdate.forEach(v => Logger.log(JSON.stringify(v, null, 2)));
    sheetModel.central.head_type = head_type;
    // Logger.log(JSON.stringify(sheetModel.central.head_type, undefined, 2));
    sheetModel.central.updateItemsObj([...mapUpdate.values()].map(v => { v.task = task; return v }));  /// временно закоментировал.

    let items = [...mapUpdate.values()].map(item => {
      for (let k in changed_from_central) {
        if (changed_from_central[k] === true) { continue; }
        if (["last", "key"].includes(k)) { continue; }

        item[k] = undefined;
      }
      return item;
    })

    head_type = new Array(sheetModel.project.head_key.length);
    sheetModel.project.head_key.forEach((h, i) => {
      head_type[i] = (() => { return fl_str(head_type_settings[h]); })();
    });

    sheetModel.project.head_type = head_type;
    // Logger.log(JSON.stringify(sheetModel.project.head_type, undefined, 2));
    // items.forEach(item => Logger.log(JSON.stringify(item, undefined, 2)));
    sheetModel.project.updateItemsObj(items);




    Logger.log(`^^^^^^^^^^^^^^^^^^^^-taskSyncЗвонки-^^^^^^^^^^^^^^^^^^^`);
  }




  /** @param {DataTaskOnEditЗвонки} task */
  taskOnEditЗвонки(task) {

  }


}

