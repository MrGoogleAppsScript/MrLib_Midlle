class FileObjData {
  constructor() {
    this.classType = "FileObjData";
    this.НомерПроекта = undefined;
    this.key = undefined;
    this.value = undefined;

    this.ФайлПроекта = undefined;
    this.ИмяФайлаПроекта = undefined;
    this.fileId = undefined;

    this.ТаблицаПроекта = undefined;
  }
}


class MrClassFile {
  constructor(folderId = getSettings().folderId) {
    /**  @private  */
    this.folderId = folderId;  
    /**  @private  */
    this.filePatch = "https://drive.google.com/file/d/ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ";
    /**  @private  */
    this.fileIdL = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ";
    /**  @private  */
    this.sheetModel = undefined;
  }
  /**  @private  */
  getSheetModel() {
    if (!this.sheetModel) {
      let context_to = new MrContext(getSettings().urls.Проекты, getSettings());
      this.sheetModel = new MrClassSheetModel(getSettings().sheetNames.Файлы, context_to);
    }
    return this.sheetModel;
  }


  // getValue(obj) {
  //   if (!obj.value) {
  //     this.readObj(obj);
  //   }
  //   return obj.value;
  // }
  /**   @private  */
  isUrl(url) {
    if (!url) { return false; }
    // if (`${url}`.length != this.filePatch.length) { return false; }
    if (`${url}`.indexOf(`${url}`.slice(0, this.filePatch.indexOf("Я") - 1)) != 0) { return false; }
    return true;
  }

  /**   @private  @param {FileObjData} obj  */
  isValidObj(obj) {
    let ret = false;

    if (obj.НомерПроекта) { ret = true; }
    if (obj.ФайлПроекта) { ret = true; }
    if (obj.ТаблицаПроекта) { ret = true; }
    if (ret == false) { return ret; }

    return ret;
  }



  /**   @private  @param {FileObjData} obj  */
  getFileId(obj) {
    if (!obj.fileId) {
      let sm = this.getSheetModel();


      if (!obj.ФайлПроекта) {
        if (obj.ТаблицаПроекта) {
          let mapN = new Map();
          sm.getMap().forEach(v => { mapN.set(v.ТаблицаПроекта, v) });
          let item = sm.getMap().get(obj.ТаблицаПроекта);
          if (item) {
            obj.ФайлПроекта = item.ФайлПроекта;
            obj.НомерПроекта = item.НомерПроекта;
            obj.key = item.key;
          }
        }
      }

      if (!obj.ФайлПроекта) {
        if (obj.НомерПроекта) {
          // sm.getMap().forEach((v, k) => Logger.log(`${k} | ${JSON.stringify(v)}`));
          let item = sm.getMap().get(`${obj.НомерПроекта}`);
          // Logger.log(`${obj.НомерПроекта} | ${JSON.stringify(item)}`);
          if (item) {
            obj.ФайлПроекта = item.ФайлПроекта;
            obj.key = item.key;
          } else {
            this.makeFile(obj);
          }
        }
      }


      if (!this.isUrl(obj.ФайлПроекта)) {
        this.makeFile(obj)
        return this.getFileId(obj);
      } else {
        obj.fileId = `${obj.ФайлПроекта}`.slice(this.filePatch.indexOf("Я"), this.filePatch.lastIndexOf("Я") + 1);
        // Logger.log(`${obj.fileId}`);
      }
    }
    return obj.fileId;
  }


  // /** @param {FileObjData} obj , @param {FileObjData} objUp */
  // updateValue(obj, objUp) {
  //   if (typeof objUp.value != "object") { return; }
  //   if (!obj.value) { obj.value = new Object(); }
  //   if (typeof obj.value != "object") { return; }
  //   obj.value = this.getValue(obj);
  //   for (let k in objUp.value) {
  //     obj.value[k] = objUp.value[k];
  //   }
  //   // this.writeObj(obj);
  // }


  /**  @private @param {FileObjData} obj , @param {FileObjData} objUp */
  joinObj(obj, objUp) {
    if (typeof objUp != "object") { return; }
    if (typeof obj != "object") { return; }
    for (let k in objUp) {
      if (k = "value") { this.joinObjValue(obj, objUp); continue }
      obj[k] = objUp[k];
    }

  }


  /**  @private @param {FileObjData} obj , @param {FileObjData} objUp */
  joinObjValue(obj, objUp) {
    if (typeof objUp.value != "object") { return; }
    if (!obj.value) { obj.value = new Object(); }
    if (typeof obj.value != "object") { return; }

    // obj.value = this.getValue(obj);
    for (let k in objUp.value) {
      obj.value[k] = objUp.value[k];
    }
  }





  /**   @param {FileObjData} obj  */
  readObj(obj) {
    let fileId = this.getFileId(obj);
    let file = undefined;
    try {
      file = DriveApp.getFileById(fileId);
      let content = file.getBlob().getDataAsString();
      let fObj = (() => { try { return JSON.parse(content); } catch (err) { return new FileObjData(); } })();
      this.joinObj(obj, fObj);
    } catch (err) { }
    // Logger.log([file.getMimeType(), file.getId(), file.getBlob().getDataAsString()]);
    // Logger.log(JSON.stringify(obj, null, 2));
  }

  /**   @private  @param {FileObjData} obj  */
  writeObj(obj) {
    let fileId = this.getFileId(obj);
    let file = undefined;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (err) {
      this.makeFile(obj);
      this.writeObj(obj);
      return;
    }

    let content = JSON.stringify(obj, null, 4);
    file.setContent(content);
    // obj.key = obj.НомерПроекта;
    obj.key = `${obj.НомерПроекта}`;
    Logger.log(`writeObj obj.key ${obj.key} | obj.ФайлПроекта ${obj.НомерПроекта} \n ${obj.ФайлПроекта} `);
    obj.info = "writeObj " + obj.key;
    obj.ИмяФайлаПроекта = file.getName();
    this.getSheetModel().updateItems([obj]);

    // Logger.log([file.getMimeType(), file.getId(), file.getBlob().getDataAsString()])
  }

  /**   @private  @param {FileObjData} obj   */
  makeFile(obj) {
    let destination = DriveApp.getFolderById(this.folderId);
    let objJson = JSON.stringify(new Object(), null, 4);
    let file = DriveApp.createFile(`${obj.НомерПроекта}`, objJson, "text/plain")
    file.moveTo(destination);
    obj.ФайлПроекта = file.getUrl();
    obj.fileId = file.getId();
    // obj.key = obj.НомерПроекта;
    obj.key = `${obj.НомерПроекта}`;
    Logger.log(`makeFile obj.key ${obj.key} | obj.ФайлПроекта ${obj.НомерПроекта} \n ${obj.ФайлПроекта} `);
    obj.info = "makeFile " + obj.key;
    obj.ИмяФайлаПроекта = file.getName();
    this.getSheetModel().updateItems([obj]);
  }

  /**  @param {FileObjData} obj   */

  update(obj) {
    if (!this.isValidObj(obj)) { throw new Error(`"Is not valid FileObjData" | ${JSON.stringify(obj)}`); }
    this.getFileId(obj);
    if (!obj.ФайлПроекта) {
      Logger.log(`"obj.ФайлПроекта Ненайден" | \n ${JSON.stringify(obj)}`);
      throw new Error(`"obj.ФайлПроекта Ненайден" |`);
    }
    let fObj = {
      ФайлПроекта: obj.ФайлПроекта,
      НомерПроекта: obj.НомерПроекта,
      key: obj.key,
    };

    this.readObj(fObj);
    this.joinObj(fObj, obj);
    this.writeObj(fObj);
  }

  moveTo(obj, folderId) {
    // Logger.log(`"moveTo" ${folderId} | `);
    try {
      let fileId = this.getFileId(obj);
      let destination = DriveApp.getFolderById(folderId);
      let file = DriveApp.getFileById(fileId);
      file.moveTo(destination);
    } catch (err) {
      mrErrToString(err);
      return;
    }

  }


  /** @param {Map}  objsMap */
  fillObjs(objsMap, objShema = {}) {

    objsMap.forEach(o => {
      if (typeof o != "object") { return; }
      this.fillObj(o, objShema);
    });

  }


  fillObj(obj, objShema = {}) {

    let fill = ((obj, value, valueShema) => {
      for (let k in valueShema) {
        // Logger.log({ k, shema: valueShema[k] });

        if (typeof valueShema[k] == "object") {
          obj[k] = new Object();
          fill(obj[k], value, valueShema[k]);
          continue;
        }
        if (value[valueShema[k]]) {
          obj[k] = value[valueShema[k]];
        }
      }
    });

    if (typeof obj != "object") { return; }
    let key = `${obj.key}`; // 
    let objR = new FileObjData();
    objR.key = key;
    objR.НомерПроекта = key;
    this.readObj(objR);

    if (!objR.value) { objR.value = new Object(); }
    if (typeof objR.value != "object") { return; }
    fill(obj, objR.value, objShema);
  }

}

function getNewFileObjData() {
  return new FileObjData();
}
/** @type {MrClassFile} */
let mrClassFile = undefined;
function getMrClassFile() {
  if (!mrClassFile) { mrClassFile = new MrClassFile(); }
  return mrClassFile;
}



class Test_MrClassFileModel {
  constructor() {
    this.folderId = ".........................";  
    this.fileUrl_01 = "https://drive.google.com/file/d/.........................../";
  }

  test_01() {
    // let obj= new Object();
    let fileObjData = new FileObjData();
    fileObjData.НомерПроекта = "key"
    fileObjData.folderId = this.folderId;
    fileObjData.value = { date: new Date(), mesage: `mesage= ${new Date()}`, text: "Это проверка на запись и чтение" }
    fileObjData.name = `test | ${new Date()}`;
    fileObjData.ФайлПроекта = undefined;
    let fm = new MrClassFile(this.folderId);

    fm.writeObj(fileObjData);
    this.fileUrl_02 = fileObjData.ФайлПроекта
    this.fileObjData_01 = fileObjData;
  }


  test_02() {
    // let obj= new Object();
    let fileObjData = new FileObjData();
    fileObjData.НомерПроекта = "key"
    fileObjData.folderId = this.folderId;
    fileObjData.value = { date: new Date() }
    fileObjData.ФайлПроекта = this.fileUrl_02;
    let fm = new MrClassFile(this.folderId);
    fm.readObj(fileObjData);
    this.fileObjData_02 = fileObjData;

  }


  test_03() {

    Logger.log(JSON.stringify(this.fileObjData_02.value));
    Logger.log(JSON.stringify(this.fileObjData_01.value));

    if (JSON.stringify(this.fileObjData_02.value) == JSON.stringify(this.fileObjData_01.value)) {
      Logger.log("равны");
    } else { Logger.log(" НЕ равны"); }

  }


  test_04() {

    let fileObjData = new FileObjData();
    fileObjData.НомерПроекта = "Т_9999"
    fileObjData.value = { date: new Date(), mesage: `mesage= ${Utilities.formatString("dd.MM.yyyy", new Date())}`, text: "Это проверка на запись и чтение   В связке с таблицей" };
    let fm = new MrClassFile(this.folderId);
    fm.update(fileObjData);
  }


  test_05() {

    let fileObjData = new FileObjData();
    fileObjData.НомерПроекта = "Т_9998"
    fileObjData.value = { date: new Date(), mesage: `mesage= ${Utilities.formatString("dd.MM.yyyy", new Date())}`, text: "Это проверка на запись и чтение   В связке с таблицей" };
    let fm = new MrClassFile(this.folderId);
    fm.update(fileObjData);
  }


  getTests() {
    return [
      // this.test_01.name,
      // this.test_02.name,
      // this.test_03.name,
      this.test_04.name,
      this.test_05.name,
    ]
  }

  testRun() {
    this.getTests().forEach(t => {
      try {
        this[t]();
      } catch (err) { mrErrToString(err); }
    });
  }
}


function test_MrClassFileModel() {
  let tcfm = new Test_MrClassFileModel();
  tcfm.testRun()


}







