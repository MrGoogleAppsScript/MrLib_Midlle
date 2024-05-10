class DataTaskSyncЭкспортВБитрикс { // mrClassЭкспортВБитрикс
  constructor() {
    /** @constant */
    this.date = new Date();
    this.exe_class = ExeTask_SyncЭкспортВБитрикс.name;
    /** @constant */
    this.name = ExeTask_SyncЭкспортВБитрикс.prototype.taskSyncЭкспортВБитрикс.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 2;
    this.off = false;
    this.param = {
      urls: {
        proj: "",
        central: "https://docs.google.com/spreadsheets/d//edit",
      },

      sheetNames: {
        proj: getSettings().sheetNames.Экспорт_в_битрикс,
        centralВлияетДА: getSettings().sheetNames.Прием_данных_из_Таблиц.Влияет_на_стадию,
        centralВлияетНЕ: getSettings().sheetNames.Прием_данных_из_Таблиц.Не_влияет_на_стадию,
      },

      range: {
        proj: {
          head: { first: 2, last: 2, key: 2, },
          body: { first: 4, last: 4, },
        },

        central: {
          head: { first: 1, last: 1, key: 1, },
          body: { first: 2, last: 2, },
        },
        // settings: {
        //   head: { first: 1, last: 1, key: 1, },
        //   body: { first: 2, last: 2, },
        // },

      },
    };
  }
}

function getDataTaskSyncЭкспортВБитрикс() {
  return new DataTaskSyncЭкспортВБитрикс();
}






class ExeTask_SyncЭкспортВБитрикс extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    // this.cache = sharedCache;
    // this.cache.abc = 1;
    this.myCacheName = ExeTask_SyncLogist.name;
    // cache.abc = MrClassExportReestr.name;
    this.names = {
      обновлено: "обновлено",
      СсылкаНаТаблицу: "Ссылка на таблицу"
    }

  }

  /** @param {DataTaskSyncЭкспортВБитрикс} task */
  taskSyncЭкспортВБитрикс(task) {


    Logger.log(`taskSyncЭкспортВБитрикс  - ${JSON.stringify(task)} `);

    let context = {
      proj: new MrContext(task.param.urls.proj, new Object()),
      central: new MrContext(task.param.urls.central, new Object()),
    }



    let sheetModel = {
      project: new MrClassSheetModel(task.param.sheetNames.proj, context.proj, task.param.range.proj),
      centralВлияетДА: new MrClassSheetModel(task.param.sheetNames.centralВлияетДА, context.central, task.param.range.central),
      centralВлияетНЕ: new MrClassSheetModel(task.param.sheetNames.centralВлияетНЕ, context.central, task.param.range.central),
    }


    let mapProject = sheetModel.project.getMap();

    [
      sheetModel.centralВлияетНЕ,
      sheetModel.centralВлияетДА,
    ].forEach((sm) => {
      let mapCentral = sm.getMap();

      // Logger.log(mapCentral.size);
      // Logger.log(mapProject.size);
      // return;
      let mapKeyRowCentral = new Map();

      mapCentral.forEach(obj => {
        let url = `${obj[this.names.СсылкаНаТаблицу]}`.slice(0, 88);
        if (url.length < 88) {
          let url_patern = "https://docs.google.com/spreadsheets/d/ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ/edit";
          let ok_url = `${url}${url_patern.slice(url.length)}`;
          if (!ok_url.includes("Я")) { url = ok_url; } else {
          }
        }
        // Logger.log({ row: obj.row, key: obj.key, url, });
        mapKeyRowCentral.set(url, { row: obj.row, key: obj.key, url, obj, });
      });

      // Logger.log([...mapKeyRowCentral.keys()]);

      let keyProjectUrl = task.param.urls.proj;
      let objCentral = mapKeyRowCentral.get(keyProjectUrl);
      if (!objCentral) { return; }
      // Logger.log(JSON.stringify(keyProjectUrl));
      // Logger.log(JSON.stringify(objCentral));

      let objProject = mapProject.get("4");
      objProject["EffectiveUser"] = Session.getEffectiveUser().getEmail();
      objProject["Owner"] = getContext().getSpreadSheet().getOwner().getEmail();
      // objProject["Дата и Время Обновления"] = getContext().timeConstruct;

      // let Дата_и_Время_Изменения=((objProjectData, objCentral) ()
      let theHasCanged = false;
      let setObj = new Object();
      ((theObj) => {
        // Logger.log(`Дата_и_Время_Изменения`);
        try {
          for (let k in objProject) {
            if (k == "row") { continue; }
            // Logger.log(`${objCentral.obj[k]} | ${objProjectData[k]}  | ${objCentral.obj[k] != objProjectData[k]}`);
            if (!(k in objCentral.obj)) { continue; }
            if (objCentral.obj[k] != objProject[k]) {
              theObj[k] = objProject[k];
              theHasCanged = true;
            }


          }
        } catch (err) { mrErrToString(err); }
      })(setObj);


      // objProject["EffectiveUser"] = Session.getEffectiveUser();
      // objProject["Owner"] = getContext().getSpreadSheet().getOwner();
      // objProject["Дата и Время Изменения"] = (() => { if (theHasCanged) { return getContext().timeConstruct; } })();
      // objProject["Дата и Время Обновления"] = getContext().timeConstruct;

      setObj["Дата и Время Изменения"] = (() => { if (theHasCanged) { return getContext().timeConstruct; } })();
      setObj["Дата и Время Обновления"] = getContext().timeConstruct;

      // objProject["Дата и Время Изменения"] = (() => {
      //   // Logger.log(`Дата_и_Время_Изменения`);
      //   try {
      //     let ff = false;
      //     if (!objCentral) { return; }
      //     // Logger.log(objCentral);
      //     for (let k in objProject) {
      //       if (k == "row") { continue; }

      //       // Logger.log(`${objCentral.obj[k]} | ${objProjectData[k]}  | ${objCentral.obj[k] != objProjectData[k]}`);
      //       if (!(k in objCentral.obj)) { continue; }
      //       if (objCentral.obj[k] != objProject[k]) {
      //         sst.push(k);
      //         ff = true;
      //       }
      //       if (ff) { return getContext().timeConstruct; }

      //     }
      //   } catch (err) { mrErrToString(err); }
      // })();


      setObj.key = objCentral.key;
      setObj.row = objCentral.row;
      // sheetModel.central.saveObj(objProjectData);
      sm.updateObjInRow(setObj, setObj.row);
      // sheetModel.centralВлияетНЕ.updateObjInRow(objProject, objProject.row);
      Logger.log(JSON.stringify({ row: setObj.row }, null, 2));
      Logger.log(JSON.stringify(setObj, null, 2));




    });

    try {
      sheetModel.project.sheet.getRange("A1").setValue("№");
      sheetModel.project.sheet.getRange("A1").setNote(Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.YYYY HH:mm:ss"));
    } catch (err) {
      mrErrToString(err);
    }

  }


}

