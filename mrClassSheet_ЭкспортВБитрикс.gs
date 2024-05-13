
class MrClassSheetЭкспортВБитрикс {
  constructor(url = getContext().getSpreadSheet().getUrl()) { // class constructor
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Экспорт_в_битрикс);

    // this.folderId = getSettings().folderId;
    // this.folderId = getSettings().folderId_оплаты;
    // this.url_patternSpreadSheet = getSettings().url_таблица_оплаты_шаблон;

    this.rowHads = 1;
    this.url = url;
    this.rowBodyFirst = 2;

    // this.rowBodyLast = this.findRowBodyLast();
    // this.sheetName_Исполнение_Внешние = getContext().settings.sheetName_Исполнение_Внешние;

    // this.makeCol();

  }

  update() {
    // this.requiredKey();
    // return;
    let dataTaskSyncЭкспортВБитрикс = MrLib_Midlle.getDataTaskSyncЭкспортВБитрикс();
    dataTaskSyncЭкспортВБитрикс.param.urls.proj = this.url;
    // Logger.log(JSON.stringify(dataTaskSyncЭкспортВБитрикс, undefined, 2));
    MrLib_Midlle.ВыполненитьМасивЗадач("MrClassSheetЭкспортВБитрикс update ВыполненитьМасивЗадач", [dataTaskSyncЭкспортВБитрикс,], "MrClassSheetЭкспортВБитрикс");
  }

}


function triggerUpdateЭкспортВБитрикс_() {
  Logger.log("triggerUpdateЭкспортВБитрикси");
  let ЭкспортВБитрикс = undefined;
  try {
    ЭкспортВБитрикс = new MrClassSheetЭкспортВБитрикс();
  } catch (err) { mrErrToString(err); return }
  ЭкспортВБитрикс.update();

}


function menuЭкспортВБитрикс() {
  Logger.log("menuЭкспортВБитрикс");
  triggerUpdateЭкспортВБитрикс_();

}


function ttriggerUpdateЭкспортВБитрикс() {
  Logger.log("triggerUpdateЭкспортВБитрикси");
  triggerUpdateЭкспортВБитрикс_();

}

