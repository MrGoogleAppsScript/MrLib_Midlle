class DataTaskSyncПроета {
  constructor() {
    /** @constant */
    this.date = new Date();
    this.exe_class = ExeTask_SyncПроекта.name;
    /** @constant */
    this.name = ExeTask_SyncПроекта.prototype.taskSyncItemsFromProjToCentral.name;
    /** @constant */
    this.актуально_дней = 1 / 24 / 4;
    this.off = false;
    this.param = {
      projNumber: undefined,
      projHead: undefined,
      urlHead: undefined,
      urls: {
        proj: "",
        central: getSettings().urls.Письма,
        sheetTemplate: getSettings().urls.ЧистаяВсеПочты,
      },

      sheetNames: {
        proj: undefined,
        central: undefined,
      },


      range: {
        proj: undefined,
        central: undefined,
      },
    };
  }
}

// function getDataTaskSyncОплатыВнешней() {
//   return new DataTaskSyncПроета();
// }

function getDataTaskSyncПроектаСЦентральной() {
  let ret = new DataTaskSyncПроета();
  ret.name = ExeTask_SyncПроекта.prototype.taskSyncItemsFromProjToCentral.name;
   return ret;
}



class ExeTask_SyncПроекта extends MrTaskExecutorBaze {
  constructor(sharedCache) {
    super(sharedCache)
    this.myCacheName = ExeTask_SyncПроекта.name;
    this.names = {
      обновлено: "обновлено",
    }

  }

  /** @param {DataTaskSyncПроета} task 
   
  */
  syncHeaders(task) {
    let context = {
      proj: new MrContext(task.param.urls.proj, getSettings()),
      template: new MrContext(task.param.urls.sheetTemplate, getSettings()),
    }
    let sheets = {
      template: context.template.getSheetByName(task.param.sheetNames.proj),
      proj: context.proj.getSheetByName(task.param.sheetNames.proj),
    }
    let range = {
      template: sheets.template.getRange("A1:10"),
      proj: sheets.proj.getRange("A1:10"),
    }

    let vls = {
      template: range.template.getValues(),
      proj: range.proj.getValues(),
    }

    let maxCol = {
      template: sheets.template.getMaxColumns(),
      proj: sheets.proj.getMaxColumns(),
    }


    let formulas = {
      template: range.template.getFormulasR1C1(),
      proj: range.proj.getFormulasR1C1(),
    }

    let ifs = {
      vls: (JSON.stringify(vls.template) == JSON.stringify(vls.proj)),
      // vls: true,
      form: (JSON.stringify(formulas.template) == JSON.stringify(formulas.proj)),
    }


    if (
      !((v1, v2) => {
        if (v1 == false) { return true; }
        if (v2 == false) { return true; }
        return false;
      })(ifs.vls, ifs.form)
    ) { return false; }

    if (maxCol.proj > maxCol.template) {
      sheets.proj.deleteColumns(maxCol.template + 1, maxCol.proj - maxCol.template);
    }
    let syncRange = sheets.proj.getRange(1, 1, formulas.template.length, formulas.template[0].length)
    // syncRange.setValues(vls.template);
    sheets.proj.clearContents();
    syncRange.setFormulas(formulas.template);
    return true;
  }

  /** @param {DataTaskSyncПроета} task */


  /** @param {DataTaskSyncПроета} task */
  taskSyncItemsFromProjToCentral(task) {

    Logger.log(`vvvvvvvvvvvvvvvvvvvvvvv-taskSyncItemsFromProjToCentral-vvvvvvvvvvvvvvvvvvvvvvvv`);
    Logger.log(`${JSON.stringify(task, null, 2)} `);


    try {
      let aa = this.syncHeaders(task);
      let sheetName = "Лог";
      let itemKey = task.param.projNumber;
      let fild_name = "ok"
      let value = aa;
      setFildForItem(sheetName, itemKey, fild_name, value)
    }
    catch (err) {
      let sheetName = "Лог";
      let itemKey = task.param.projNumber;
      let fild_name = "err"
      let value = mrErrToString(err);
      value = JSON.stringify(value) + "\n" + JSON.stringify(task);
      setFildForItem(sheetName, itemKey, fild_name, value)
    }

    let context = {
      proj: new MrContext(task.param.urls.proj, getSettings()),
      central: new MrContext(task.param.urls.central, getSettings()),
    }


    let theSheetModel = {
      project: new MrClassSheetModel(task.param.sheetNames.proj, context.proj, task.param.range.proj),
      central: new MrClassSheetModel(task.param.sheetNames.central, context.central, task.param.range.central),
    }

    // if ((new Date().getTime() - new Date(theSheetModel.project.getMem(this.names.обновлено)).getTime() - task.актуально_дней * DeyMilliseconds) < 0) {
    //   Logger.log("да ещё актуальны");
    //   Logger.log(`^^^^^^^^^^^^^^^^^^^^-taskSyncSheet-^^^^^^^^^^^^^^^^^^^`);
    //   return;
    // }


    let map = {
      // central: theSheetModel.central.getMap(),
      project: theSheetModel.project.getMap(),
    }


    // task.param.projNumber = `${task.param.projNumber}`;
    let items = [...map.project.values()]
      .map(item => {
        item.key=task.param.projNumber;
        item[task.param.projHead]=task.param.projNumber;
        item[task.param.urlHead]=task.param.urls.proj;
        return item;
      });

    Logger.log(items.length)
    theSheetModel.central.updateItems(items);

    Logger.log(`^^^^^^^^^^^^^^^^^^^^-taskSyncItemsFromProjToCentral-^^^^^^^^^^^^^^^^^^^`);
  }

}

