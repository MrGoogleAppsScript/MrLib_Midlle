// class DataTaskLoadExportValuesFromProjects {
//   constructor() {
//     /** @constant */
//     this.exe_class = ExeTask_ExportValuesFromProjects.name;
//     /** @constant */
//     this.name = ExeTask_ExportValuesFromProjects.prototype.taskLoadExportValuesFromProjects.name;
//     this.off = false;
//     this.актуально_дней = 1 / 8 / 3 / 2;
//     this.param = {
//       sheetNames: {
//         to: "",
//       },

//       urls: {
//         to: "",
//       },

//     };
//   }
// }

class DataTaskGetExportValuesFromProjects {
  constructor() {
    /** @constant */
    this.exe_class = ExeTask_ExportValuesFromProjects.name;
    /** @constant */
    this.name = ExeTask_ExportValuesFromProjects.prototype.taskGetExportValuesFromProjects.name;
    this.off = false;
    this.актуально_дней = 1 / 8 / 3 / 1;
    this.param = {
      sheetNames: {
        to: "",
      },

      urls: {
        to: "",
      },

    };
  }
}





class ExeTask_ExportValuesFromProjects extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    this.myCacheName = ExeTask_ExportValuesFromProjects.name;
  }

  /** @param {DataTaskGetExportValuesFromProjects} task */
  taskGetExportValuesFromProjects(task) {
    Logger.log(`${JSON.stringify(task)} `);
    let schema = {
      Р: {
        Реестр: "Реестр",
        Исполнение: "Исполнение",
      },
      Тб: "Тб",
      Инфо: {
        // "№": "№",
        "Заказчик": "Заказчик",
        "Название": "Название",
        "Таблица": "Таблица",
        "Исполнение_date": "Исполнение_date",
        "Реестр_date": "Реестр_date",
        "Тб_date": "Тб_date",
        "Статус проекта": "Статус проекта",
        "Ответственный": "Ответственный",
        "ver": "ver",
      }
    }

    // let url = task.param.url;

    let context_to = new MrContext(task.param.urls.to, getSettings());
    let sheetModelProjects = new MrClassSheetModel(task.param.sheetNames.to, context_to);
    let projectArr = [...sheetModelProjects.getMap().values()].reverse();

    projectArr = projectArr.filter(project => {
      if (fl_str(project["Р.Исполнение.Статус проекта"]) == fl_str("ЗАВЕРШЕН")) {
        Logger.log([project.key, project["Р.Исполнение.Статус проекта"]]);
        return false;
      }
      if (fl_str(project["Р.Исполнение.Статус проекта"]) == fl_str("Завершен")) {
        Logger.log([project.key, project["Р.Исполнение.Статус проекта"]]);
        return false;
      }
      return true;
    });

    let i = (this.getCacheValue("i") ? this.getCacheValue("i") : 0);
    if (i >= projectArr.length) { i = 0; }
    Logger.log(` ${this.myCacheName} | i=${i} | this.tasks.length.=${projectArr.length} `);



    let fileObjПроекты_в_Исполнение = getNewFileObjData();
    fileObjПроекты_в_Исполнение.НомерПроекта = task.param.sheetNames.to;
    fileObjПроекты_в_Исполнение.ТаблицаПроекта = task.param.urls.to;
    fileObjПроекты_в_Исполнение.value = new Object();
    // fileObjПроекты_в_Исполнение.value.itemArr = itemArr;
    getMrClassFile().readObj(fileObjПроекты_в_Исполнение);
    // fileObjПроекты_в_Исполнение.value.itemArr = "";

    getMrClassFile().moveTo(fileObjПроекты_в_Исполнение, "");


    // let b = 0;
    for (; i < projectArr.length; i++) {
      this.setCacheValue("i", i);
      if (!context_to.hasTime()) { Logger.log(`${this.myCacheName} taskGetExportValuesFromProjects | мало времени break`); break; }
      // if (b++ > 30) { Utilities.sleep(30 * 1000); continue; }


      let project = projectArr[i];
      let lastUpdate = (() => { try { return JSON.parse(project["lastUpdate"]); } catch { return 0; } })();
      // Logger.log(project["тб.date"]);
      // Logger.log(lastUpdate);
      // // Logger.log(JSON.stringify(project));

      // Logger.log("ExeTask_ExportValuesFromProjects ещё актуальны?");
      // Logger.log(lastUpdate);
      // Logger.log(` ${new Date(lastUpdate).getTime()} `);
      // Logger.log(` ${task.актуально_дней} `);

      if ((new Date().getTime() - new Date(lastUpdate).getTime() - task.актуально_дней * DeyMilliseconds) < 0) { Logger.log("да ещё актуальны | " + `${project.key}`); continue; }
      lastUpdate = context_to.timeConstruct;

      let url = project["Таблица"];

      let projectValues = {};

      let error = undefined;
      try {
        // let context = new MrContext(url, getSettings());
        // projectValues = new MrClassExportProject(context, projectValues).getExportValues();
        projectValues = (() => { try { return JSON.parse(getItemByKey(getSettings().sheetNames.Проекты, url, "Тб", undefined)); } catch (err) { return undefined } })();

      } catch (err) {
        error = mrErrToString(err);
      }

      if (!projectValues) {
        projectValues = new Object();
        // projectValues.key = project.key;
        projectValues.errors = [];
      }

      projectValues.spreadSheetURL = url;
      projectValues.email = Session.getEffectiveUser().getEmail();
      projectValues.date = new Date();
      projectValues.errors = new Array();


      projectValues.key = project.key;
      if (error) { projectValues.errors = [].concat(error, projectValues.errors); }

      let ver = (() => { try { return JSON.parse(getItemByKey(getSettings().sheetNames.Проекты, url, "ver", undefined)); } catch (err) { return undefined } })();


      let ФайлПроекта = (() => { try { return getItemByKey(getSettings().sheetNames.Файлы, `${project.key}`, "ФайлПроекта", undefined); } catch (err) { mrErrToString(err); return undefined } })();
      // let ФайлПроекта = undefined;
      Logger.log(JSON.stringify({ key: project.key, ФайлПроекта }));
      sheetModelProjects.updateItems([{ key: project.key, Тб: projectValues, ver: ver, ФайлПроекта, lastUpdate }]);
      // sheetModelProjects.updateItems([{ key: project.key, ФайлПроекта, lastUpdate }]);

      let fileObj = getNewFileObjData();
      fileObj.ФайлПроекта = ФайлПроекта;
      fileObj.НомерПроекта = `${project.key}`;
      fileObj.key = `${project.key}`;

      getMrClassFile().moveTo(fileObj, "1???????????????????????????????????????????");  

      // getMrClassFile().readObj(fileObj);

      let mobj = new Object();
      mobj.key = `${project.key}`;
      getMrClassFile().fillObj(mobj, schema);
      fileObjПроекты_в_Исполнение.value[`${project.key}`] = mobj;

    }

    if (i >= projectArr.length) {
      i = undefined;
    }
    this.setCacheValue("i", i);
    getMrClassFile().update(fileObjПроекты_в_Исполнение);

  }



  // /** @param {DataTaskLoadExportValuesFromProjects} task */
  // taskLoadExportValuesFromProjects(task) {

  //   // let url = task.param.url;

  //   let context_to = new MrContext(task.param.urls.to, getSettings());
  //   let sheetModelProjects = new MrClassSheetModel(task.param.sheetNames.to, context_to);
  //   let projectArr = [...sheetModelProjects.getMap().values()];

  //   let i = (this.getCacheValue("i") ? this.getCacheValue("i") : 0);
  //   if (i >= projectArr.length) { i = 0; }
  //   Logger.log(` ${this.myCacheName} | i=${i} | this.tasks.length.=${projectArr.length} `);

  //   // let b = 0;
  //   for (; i < projectArr.length; i++) {
  //     this.setCacheValue("i", i);
  //     if (!context_to.hasTime()) { Logger.log(`${this.myCacheName} taskLoadExportValuesFromProjects | мало времени break`); break; }
  //     // if (b++ > 30) { Utilities.sleep(30 * 1000); continue; }


  //     let project = projectArr[i];
  //     let lastUpdate = (() => { try { return JSON.parse(project["Тб.date"]); } catch { return 0; } })();
  //     // Logger.log(project["тб.date"]);
  //     // Logger.log(lastUpdate);
  //     // // Logger.log(JSON.stringify(project));

  //     // Logger.log("ExeTask_ExportValuesFromProjects ещё актуальны?");
  //     // Logger.log(lastUpdate);
  //     // Logger.log(` ${new Date(lastUpdate).getTime()} `);
  //     // Logger.log(` ${task.актуально_дней} `);

  //     if ((new Date().getTime() - new Date(lastUpdate).getTime() - task.актуально_дней * DeyMilliseconds) < 0) { Logger.log("да ещё актуальны | " + `${project.key}`); continue; }


  //     let url = project["Таблица"];

  //     let projectValues = {
  //       spreadSheetURL: url,
  //       email: Session.getEffectiveUser().getEmail(),
  //       date: new Date(),
  //       errors: new Array(),
  //     };

  //     let error = undefined;
  //     try {
  //       let context = new MrContext(url, getSettings());
  //       projectValues = new MrClassExportProject(context, projectValues).getExportValues();
  //     } catch (err) {
  //       error = mrErrToString(err);
  //     }

  //     if (!projectValues) {
  //       projectValues = new Object();
  //       // projectValues.key = project.key;
  //       projectValues.errors = [];
  //     }
  //     projectValues.key = project.key;
  //     if (error) { projectValues.errors = [].concat(error, projectValues.errors); }


  //     sheetModelProjects.updateItems([{ key: project.key, Тб: projectValues }]);
  //   }

  //   if (i >= projectArr.length) {
  //     i = undefined;
  //   }
  //   this.setCacheValue("i", i);
  // }






}


// class MrClassExportProject {
//   /** @param {MrContext} context*/
//   constructor(context, item) { // class constructor
//     if (!context) { throw new Error("context"); }
//     this.context = context;
//     this.item = item;
//   }

//   getRowSobachkaBySheetName(sheetName) {
//     if (!this.sobachkaMap) { this.sobachkaMap = new SobachkaMap(this.context); }
//     return this.sobachkaMap.getRowSobachkaBySheetName(sheetName);
//   }


//   findRowBodyLast(sheetName) {
//     let ret = this.getRowSobachkaBySheetName(sheetName);
//     if (!ret) { ret = this.context.getSheetByName(sheetName).getLastRow(); }
//     else { ret = ret - 1; }
//     return ret;
//   }



//   getValuesFromВыбор_Поставщиков(item) {

//     let rowBodyFirst = 2;
//     let retArr = new Array();
//     let sheet = this.context.getSheetByName(this.context.settings.sheetNames.Выбор_поставщиков);
//     let heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//     // heads = ["row"].concat(heads).map(h => fl_str(h));
//     heads = ["row"].concat(heads).map(h => `${h}`.trim());
//     let vls = sheet.getRange(rowBodyFirst, 1, item.countProd, sheet.getLastColumn()).getValues();
//     vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
//     let col_name = nr("B");

//     let col_list = ["№", "Номенклатура", "Ед", "Кол", "Цена", "Сумма", "Срок поставки", "Поставщик", "Статус", "Предполагаемая дата оплаты",
//       // ].map(c => { return heads.indexOf(fl_str(c)); }).filter(c => (c >= 0));
//     ].map(c => { return heads.indexOf(c); }).filter(c => (c >= 0));


//     vls.forEach(v => {
//       if (!v[col_name]) { return; }
//       let obj = new Object();
//       col_list.forEach(col => {
//         obj[heads[col]] = v[col];
//       });
//       // retArr.push(heads);
//       retArr.push(obj);
//     });

//     // return retArr;
//     return {
//       тело: retArr,
//       шапка: heads.filter((h, i, arr) => { return col_list.includes(i); }),
//     }
//   }

//   getValuesFromОплаты(item) {


//     let rowBodyFirst = 3;
//     let retArr = new Array();
//     let sheet = this.context.getSheetByName(this.context.settings.sheetNames.Оплаты);
//     let heads = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
//     // heads = ["row"].concat(heads).map(h => fl_str(h));
//     heads = ["row"].concat(heads).map(h => `${h}`.trim());
//     let vls = sheet.getRange(rowBodyFirst, 1, item.countProd, sheet.getLastColumn()).getValues();
//     vls = vls.map((v, i, arr) => { return [i + rowBodyFirst].concat(v) });
//     let col_name = nr("B");


//     let col_list = [
//       "№", "Номенклатура", "Юр лицо поставщика", "№ счета", "Статус", "Плательщик", "Дата оплаты", "% оплаты", "Статус оплаты",

//       // №	Номенклатура	Ед	Кол	Цена	Сумма	Группа	Срок поставки	Экспертиза	Повторная рассылка	Поставщик	Резервный поставщик		НМЦ	Статус	Юр лицо поставщика	№ счета	Статус	Плательщик	Дата оплаты	% оплаты


//     ].map(c => { return heads.indexOf(c); }).filter(c => (c >= 0));

//     // retArr.push(
//     //   heads.filter((h, i, arr) => { return col_list.includes(i); })
//     // );


//     vls.forEach(v => {
//       if (!v[col_name]) { return; }
//       let obj = new Object();
//       col_list.forEach(col => {
//         obj[heads[col]] = v[col];
//         // if (heads[col] == "Дата оплаты") {
//         //   // if (v[col] instanceof Date) {
//         //   //   obj[heads[col]] = `📆 ${v[col]}`;
//         //   // }

//         //   // obj["Дата оплаты📆"] = `📆 ${v[col]}`;
//         //   // ${v[col]}
//         //   // obj[heads[col]] = `📆 ${v[col]}`
//         // }
//         // obj["Дата оплаты qq"] = `📆 ${new Date()}`;
//       });
//       // retArr.push(heads);
//       retArr.push(obj);
//     });

//     // return retArr;
//     return {
//       тело: retArr,
//       шапка: heads.filter((h, i, arr) => { return col_list.includes(i); }),
//     }

//   }

//   getValuesFromНастройки_писем(item) {
//     let retObj = new Object();
//     let sheet_Настройки_писем = this.context.getSheetByName(getSettings().sheetNames.Настройки_писем);
//     let vls_AJ1_AK = sheet_Настройки_писем.getRange("AJ1:AK").getValues();
//     for (let i = 0; i < vls_AJ1_AK.length; i++) {
//       let fild_name = vls_AJ1_AK[i][0];
//       let fild_value = vls_AJ1_AK[i][1];
//       if (!fild_name) { continue; }
//       retObj[fild_name] = fild_value;
//     }
//     return retObj;
//   }


//   getExportValues() {
//     this.item.name = this.context.getSpreadSheet().getName();
//     this.item.rowBodyLastСбор_КП = this.findRowBodyLast(this.context.settings.sheetNames.Выбор_поставщиков);
//     this.item.countProd = this.item.rowBodyLastСбор_КП - 2 + 1;
//     this.item.Выбор_Поставщиков = this.getValuesFromВыбор_Поставщиков(this.item);
//     // this.item.Оплаты = this.getValuesFromОплаты(this.item);
//     this.item.Настройки_писем = this.getValuesFromНастройки_писем(this.item);
//     this.item.Оплаты = this.tryAddErr(this, this.getValuesFromОплаты.name, this.item);
//     return this.item;
//   }

//   tryAddErr(tt, colable, item) {
//     try {
//       // Logger.log(colable)
//       return tt[colable](item);
//     } catch (err) {
//       item.errors = [].concat(mrErrToString(err), item.errors);
//     }
//   }
// }


