/** @typedef {Object} Obj
 * @property {number} row
 * @property {string} key
 * @property {boolean} on
 * @property {boolean} off
 * @property {string} last
 * @property {boolean} isValid	
 * @property {string} value
 */




function getMrClassSheetModel() {
  return MrClassSheetModel;
}

class MrClassSheetModel {

  constructor(sheetName, context, rowConf = undefined) {
    if (!context) { throw new Error("context"); }
    this.sheetName = sheetName;
    /** @type {MrContext} */
    this.context = context;
    this.sheet = this.context.getSheetByName(this.sheetName);
    this.isStoped = false;


    this.row = {
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
    this.row = undefined;


    this.setRowConf(rowConf);

    this.init();


  }

  log() {
    this.getMap();
    Logger.log(``);
    Logger.log(`vvvvvvvvvvvvvvv${this.sheetName}vvvvvvvvvvvvvvv`);
    Logger.log(`${this.sheetName} | ${JSON.stringify(this)}`);
    Logger.log(`this.head_key=${JSON.stringify(this.head_key)}`);
    Logger.log(`this.col     =${JSON.stringify(this.col)}`);
    // Logger.log(`this.vls     =${JSON.stringify(this.vls)}`);
    Logger.log(`this.map.size=${JSON.stringify(this.map.size)}`);
    Logger.log(`this.getKeys=${JSON.stringify(this.getKeys())}`);
    for (let [key, value] of this.getMap()) {
      Logger.log(`${key} | ${JSON.stringify(value)} `);

    }

    // for (let [key, value] of this.getMap()) {
    //   Logger.log(`${key} | ${JSON.stringify(this.toObj(value))} `);
    // }

    // for (let [key, value] of this.getMap()) {
    //   Logger.log(`${key} | ${this.toArr((this.toObj(value)))} `);
    // }


    Logger.log(`^^^^^^^^^^^^^^^${this.sheetName}^^^^^^^^^^^^^^^`);
    Logger.log(``);

  }

  setRowConf(rowConf) {
    this.row = rowConf;
    if (!this.row) {
      this.row = {
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
    this.init();
    this.reset();
  }


  init() {
    // this.row = {
    //   head: {
    //     first: 1,
    //     last: 10,
    //     key: 9,
    //     type: 8,
    //     ru_name: 10,
    //     changed: 7,
    //     off: 10,
    //     mem: 7,
    //   },

    //   body: {
    //     first: 11,
    //     last: 11,
    //   }
    // }

    this.col = {
      first: 1,
      row: undefined,
      on: undefined,
      mem: undefined,
      isValid: undefined,
      key: undefined,
      value: undefined,
      last: this.sheet.getLastColumn(),
    }
    // this.head_key = ["row"];   

    this.row.body.last = this.sheet.getLastRow();
    if (this.row.body.last < this.row.body.first) { this.row.body.last = this.row.body.first; }
    if (this.col.last == 0) { this.col.last = 1; }
    this.head_key = (() => { try { return [].concat("row", this.sheet.getRange(this.row.head.key, 1, 1, this.col.last).getValues()[0]); } catch (err) { mrErrToString(err); return []; } })();
    this.head_type = (() => { try { return [].concat("row", this.sheet.getRange(this.row.head.type, 1, 1, this.col.last).getValues()[0]).map(v => fl_str(v)); } catch (err) { return []; } })();
    this.changed_key = (() => { try { return [].concat(false, this.sheet.getRange(this.row.head.changed, 1, 1, this.col.last).getValues()[0]); } catch (err) { return []; } })();

    for (let i = this.head_key.length - 1; i >= 0; i--) {
      let key = this.head_key[i];
      if (!key) { continue; }
      this.col[key] = i;
    }
    if (!this.col.key) { this.col.key = this.col.row; }

    if (!this.col.mem) {
      this.col.mem = this.col.last;
    }


    this.isStoped = (() => {
      if (!this.col.mem) { return false; }
      let off = (() => { try { return this.sheet.getRange(this.row.head.off, this.col.mem).getValue() } catch (err) { return false; } })();
      if (off !== true) { return false; }
      return true;
    })();

  }

  getMemCell() {
    if (!this.memCell) {
      this.memCell = this.sheet.getRange(this.row.head.mem, this.col.mem);
    }
    return this.memCell;
  }

  getMemObj() {
    let ret = this.getMemCell().getValue();
    if (!ret) { ret = JSON.stringify(new Object()) };
    ret = (() => { try { return JSON.parse(ret); } catch { return new Object(); } })();
    return ret;
  }

  getMem(key) {
    return this.getMemObj()[key];
  }

  setMem(key, value) {
    let mem = this.getMemObj();
    mem[key] = value;
    this.getMemCell().setValue(JSON.stringify(mem));
  }


  reset() {
    this.context.flush();
    this.row.body.last = this.sheet.getLastRow();
    if (this.row.body.last < this.row.body.first) { this.row.body.last = this.row.body.first; }

    this.vls = undefined;
    this.map = undefined;
  }

  getValues() {
    if (!this.vls) {
      let vls = this.sheet.getRange(this.row.body.first, 1, this.row.body.last - this.row.body.first + 1, this.col.last).getValues();
      vls = vls.map((v, i, arr) => { return [this.row.body.first + i].concat(v); });
      this.vls = vls;
    }
    return this.vls;
  }


  /** @returns {Map} */
  getMap() {
    if (!this.map) {
      this.map = new Map();

      let vls = this.sheet.getRange(this.row.body.first, 1, this.row.body.last - this.row.body.first + 1, this.col.last).getValues();

      // JSON.stringify({ rowFirst: this.row.body.first, colFirst: 1, rowCount: this.row.body.last - this.row.body.first + 1, colCount: this.col.last });


      vls.forEach((v, i, arr) => {
        // Logger.log(` v ${JSON.stringify(v, null, 2)}`);
        v = [this.row.body.first + i].concat(v)
        let key = v[this.col.key];
        if (`${key}` == "") { return; }
        this.map.set(`${key}`, this.toObj(v));
      });
    }
    return this.map;
  }

  getKeys() {
    return [...this.getMap().keys()];
  }

  getRowByKey(key) {
    /** @type {Obj} */
    let obj = this.getObj(key);
    if (!obj) { return; }
    let row = obj.row;
    if (!row) { return; }
    if (row < this.row.body.first) { return; }
    return row;
  }

  /** @returns {Obj} */
  getObj(key) {
    return this.getMap().get(key);
  }

  getValueByKey(key, value_key = "value", defvalue = undefined) {
    if (!this.head_key.includes(value_key)) { return defvalue; }
    let obj = this.getObj(key);
    if (!obj) { return defvalue; }
    return obj[value_key];
  }

  /** @param {Obj} obj*/
  saveObj(obj) {
    if (Array.isArray(obj)) {
      throw new Error(`is not object ${obj}`);
    }

    if (typeof obj != "object") {
      throw new Error(`is not object ${obj}`);

    }


    let row = 0;
    if (!row) { row = obj.row; }
    if (!row) { row = this.getRowByKey(obj.key); }
    if (!row) { row = 0; }
    if (typeof row != "number") { row = 0; };

    let objArr = this.toArr(obj);
    objArr.splice(0, 1);
    // obl = obj.splice(1);
    if (row <= 0) {
      // this.sheet.appendRow(obj);
      row = Math.max(this.sheet.getLastRow() + 1, this.row.body.first);
    }

    this.setValues(row, 1, [objArr]);
    obj.row = row;
  }

  appendObj(obj) {
    obj.row = undefined;
    this.saveObj(obj);
  }

  delObj(key) {
    let row = this.getRowByKey(key);
    if (!row) { return; }
    if (row < this.row.body.first) { return; }
    this.sheet.getRange(row, 1, 1, this.col.last).clearContent();
    // this.sheet.deleteRow(row);
    this.getMap().delete(key);
    this.reset();
  }

  offByKey(key) {
    let row = this.getRowByKey(key);
    this.setValue(row, this.col.on, false);
    this.getObj(key).on = false;
    // this.reset();
  }

  onByKey(key) {
    let row = this.getRowByKey(key);
    this.setValue(row, this.col.on, true);
    this.getObj(key).on = true;
    // this.reset();
  }


  setValue(row, col, value) {
    if (!row) { return; }
    if (row < this.row.body.first) { return; }
    if (!col) { return; }
    if (this.isStoped) { return; }


    if (this.head_type[col] == fl_str("date")) {
      // let ЫЕК = ` ${row}, ${col},  ${value}| `;
      // Logger.log(ЫЕК);
      if (value) {
        // if (typeof value != "object") {
        // let d_value = new Date(value);
        // Logger.log(ЫЕК);
        let d_value = (() => { try { return JSON.parse(value) } catch (err) { return value; } })();
        d_value = new Date(d_value);
        if (!Number.isNaN(d_value.getTime())) {
          if (d_value.getTime() != 0) {
            // value = `${d_value} | ${d_value.getTime()} | ${new Date(0).getTime()}  `;
            value = d_value
          }
        }
        // }
        // ЫЕК = ` ${row}, ${col},  ${value}|2 `;
        // Logger.log(ЫЕК);
      }
      // this.sheet.getRange(row, col).clear();

    }


    try {
      if (this.head_type[col] != fl_str("formula")) {
        if (`${value}`.length < 50000) {
          this.sheet.getRange(row, col).setValue(value);
          // this.context.flush();
        }
      } else {
        this.sheet.getRange(row, col).setFormula(value);
      }
    } catch (err) { mrErrToString(err); }
  }


  setValues(row, col, vls) {
    // Logger.log(`this.isStoped = ${this.isStoped}`);
    if (this.isStoped) { return; }

    if (vls.length <= 0) { return; }
    try {
      // if (`${value}`.length < 50000) {
      this.sheet.getRange(row, col, vls.length, vls[0].length).setValues(vls);
      // }
    } catch (err) { mrErrToString(err); }
    this.reset();
  }

  /** @param {Obj} obj , @returns {[]} */
  toArr(obj) {

    let retArr = new Array(this.head_key.length);
    for (let i = this.head_key.length - 1; i >= 0; i--) {
      let key = this.head_key[i];
      // Logger.log(`i=${i}  key=${key} = val=${obj[key]}`);
      if (!key) { continue; }
      retArr[i] = obj[key];
      if (key == "value") {
        if (typeof retArr[i] != "string") {
          retArr[i] = JSON.stringify(retArr[i], null, 2);
        }
      }
    }
    return retArr;
  }

  /** @returns {Obj} */
  toObj(arr) {
    let retObj = new Object();
    // Logger.log()
    for (let i = this.head_key.length - 1; i >= 0; i--) {
      let key = this.head_key[i];
      if (!key) { continue; }
      retObj[key] = arr[i];
    }
    return retObj;
  }

  optimize() {
    // return
    SpreadsheetApp.flush();
    // Utilities.sleep(1000 * 10)
    this.reset();
    this.map = undefined;
    let map = this.getMap();
    let vls = [...map.values()].map((obj, i, arr) => { let ret = this.toArr(obj); ret.splice(0, 1); return ret; });
    // Logger.log(vls);

    this.sheet.getRange(this.row.body.first, 1, this.row.body.last - this.row.body.first + 1, this.col.last).clearContent();
    this.setValues(this.row.body.first, 1, vls);

    SpreadsheetApp.flush();
    this.reset();
    SpreadsheetApp.flush();

    let maxRows = this.sheet.getMaxRows();
    let minEmpty = 2;

    // Logger.log(maxRows - this.row.body.last);
    if (maxRows - this.row.body.last > minEmpty) {
      this.sheet.deleteRows(this.row.body.last + minEmpty, maxRows - this.row.body.last - minEmpty);
    }

  }

  /** @param {Obj} obj */
  isChanged(obj) {
    let snapshotObj = this.makeSnapshot(obj);
    if (JSON.stringify(snapshotObj, null, 2) === obj.last) { return false; }
    return true;
  }

  /** @param {Obj} obj */
  makeSnapshot(obj) {
    let retObj = new Object();
    for (let i = this.head_key.length - 1; i >= 0; i--) {
      if (this.changed_key[i] !== true) { continue; }
      let key = this.head_key[i];
      if (!key) { continue; }
      retObj[key] = obj[key];
    }
    return retObj;
  }


  // extract a value from an object by the full key

  // extractValueFromObjectByFullKey
  extractValueFromObjectByFullKey(obj, key) {
    if (`${key}` == "value") {
      return this.extractValueFromObjectByFullKey(obj, "");
    }



    if (`${key}` == "") {

      if (Array.isArray(obj)) {
        return JSON.stringify(obj, null, 2);
        // throw new Error(`is not object ${obj}`);
      }

      if (typeof obj == "object") {
        return JSON.stringify(obj, null, 2);
        // throw new Error(`is not object ${obj}`);
      }
      return obj;
    }

    let ret = undefined;

    let key_arr = `${key}`.split(".");
    // Logger.log(key_arr);
    let key_loc = key_arr[0];


    try {
      ret = obj[key_loc];
    } catch {
      // return ret;
      return this.extractValueFromObjectByFullKey(obj, "");
    }

    key_arr = key_arr.slice(1);
    // Logger.log(key_arr.length);
    if (key_arr.length != 0) {
      // Logger.log(key_arr.join("."));
      ret = this.extractValueFromObjectByFullKey(ret, key_arr.join("."));
    } else {
      if (key_loc.includes("price")) {
        if (ret) {
          ret = Number.parseFloat(ret);
        }
      }

    }
    return this.extractValueFromObjectByFullKey(ret, "");
    return ret;
  }

  /** @param {Obj} obj , @returns {[]} */
  toArrByFullKey(obj) {
    // Logger.log(` ${obj.key} `);

    let retArr = new Array(this.head_key.length);
    for (let i = this.head_key.length - 1; i >= 0; i--) {
      let key = this.head_key[i];
      // Logger.log(`type ${this.head_type[i]} | i=${i}  key=${key} = val=${obj[key]} `);
      if (!key) { continue; }
      retArr[i] = this.extractValueFromObjectByFullKey(obj, key);

      try {
        if (this.head_type[i] == fl_str("date")) {
          // Logger.log(`Как Дату ${this.head_key[i]}; ${retArr[i]}`);
          if (retArr[i]) {
            let d_value = (() => { try { return JSON.parse(retArr[i]) } catch (err) { return retArr[i]; } })();
            d_value = new Date(d_value);
            // Logger.log(`type ${this.head_type[i]} | ${d_value} i=${i}  key=${key} = val=${retArr[i]} `);
            if (!Number.isNaN(d_value.getTime())) {
              // Logger.log(`type ${this.head_type[i]} | i=${i}  key=${key} = val=${retArr[i]} `);
              if (d_value.getTime() != 0) {
                // Logger.log(`type ${this.head_type[i]} | i=${i}  key=${key} = val=${retArr[i]} `);
                retArr[i] = new Date(d_value);
              } else {
                // retArr[i]
              }
            } else {
              // Logger.log(` ${obj.key} | type ${this.head_type[i]} | i=${i}  key=${key} = val=${retArr[i]} | d_value="${d_value}"`);
            }
            // retArr[i] = d_value;
          }
          // Logger.log(`Как Дату ${this.head_key[i]}; ${retArr[i]}`);
        }
      } catch (err) {
        Logger.log(`Не Дата; ${retArr[i]}`);
        mrErrToString(err)
      }


    }
    return retArr;
  }



  clearBodyContents() {
    this.sheet.getRange(this.row.body.first, this.col.first, this.row.body.last - this.row.body.first + 1, this.col.last - this.col.first + 1).clearContent();
    this.reset();
  }

  /** @param {[Object]} items */
  setItems(items) {
    this.clearBodyContents();
    if (!items.length) { return; }
    let vls = items.map((obj, i, arr) => {
      // Logger.log(JSON.stringify(obj));
      return this.toArrByFullKey(obj).slice(1);
    });

    this.setValues(this.row.body.first, this.col.first, vls);

  }


  /** @param {[Obj]} items */
  updateItems(items) {
    if (!items.length) { return; }
    // this.getMap();

    items.forEach(item => {
      if (!item.key) { return; }
      item.key = `${item.key}`;
      if (!this.getMap().has(item.key)) {
        // this.getMap().set(item.key, { key: item.key });
        this.appendObj({ key: item.key });
      }
    });


    let vls = items.map((obj, i, arr) => {
      // Logger.log(JSON.stringify(obj));
      // return this.toArrByFullKey(obj).slice(1);
      return this.toArrByFullKey(obj);
    });

    this.reset();
    // Logger.log(vls);

    vls.forEach(v => {
      // Logger.log(v);
      let obj = this.toObj(v);

      obj = JSON.stringify(obj);
      obj = JSON.parse(obj);

      // let arr = this.toArr(obj);
      // Logger.log(JSON.stringify(obj));
      // Logger.log(arr);
      // let row = this.getMap().get(`${obj.key}`);

      if (!this.getMap().has(`${obj.key}`)) { return; }
      let row = this.getMap().get(`${obj.key}`).row;


      for (let i = this.head_key.length - 1; i >= 0; i--) {
        let key = this.head_key[i];

        if (!key) { continue; }
        // Logger.log(`i=${i}  key=${key}  val=${obj[key]}`);
        if (this.head_key[i] == "key") { continue; }
        if (obj[this.head_key[i]] == undefined) { continue; }
        // if (arr[i] == undefined) { continue; }
        let col = i;
        // this.setValue(row, col, arr[i]);
        // Logger.log(`i=${i}  key=${key}  val=${obj[this.head_key[i]]} `);
        this.setValue(row, col, obj[this.head_key[i]]);
      }
    });
  }

  /** @param {[Obj]} items */
  updateItemsObj(items) {
    if (!items.length) { return; }
    // this.getMap();

    let map = new Map();
    this.getMap().forEach((v, k) => map.set(k, v))

    items.forEach(item => {
      if (!item.key) { return; }
      item.key = `${item.key}`;
      // if (!this.getMap().has(item.key)) {
      if (!map.has(item.key)) {
        map.set(item.key, item);

        this.appendObj({ key: item.key });
      }
    });


    let vls = items.map((obj, i, arr) => {
      // Logger.log(JSON.stringify(obj));
      // return this.toArrByFullKey(obj).slice(1);
      return this.toArrByFullKey(obj);
    });

    this.reset();
    // Logger.log(vls);

    vls.forEach(v => {
      // Logger.log(v);
      let obj = this.toObj(v);

      // obj = JSON.stringify(obj);
      // obj = JSON.parse(obj);

      // let arr = this.toArr(obj);
      // Logger.log(JSON.stringify(obj));
      // Logger.log(arr);
      // let row = this.getMap().get(`${obj.key}`);

      if (!this.getMap().has(`${obj.key}`)) { return; }
      let row = this.getMap().get(`${obj.key}`).row;


      for (let i = this.head_key.length - 1; i >= 0; i--) {
        let key = this.head_key[i];

        if (!key) { continue; }
        // Logger.log(`i=${i}  key=${key}  val=${obj[key]}`);
        if (this.head_key[i] == "key") { continue; }
        if (obj[this.head_key[i]] == undefined) { continue; }
        // if (arr[i] == undefined) { continue; }
        let col = i;
        // this.setValue(row, col, arr[i]);
        // Logger.log(`i=${i}  key=${key}  val=${obj[this.head_key[i]]}  | arr[i]=${arr[i]} `);
        this.setValue(row, col, obj[this.head_key[i]]);
      }
    });
  }

  /** @param {Array} colNamesArr */
  requiredColNames(colNamesArr) {
    let colLast = (() => {
      if (this.head_key.includes("last")) { return this.head_key.indexOf("last") } else {
        this.sheet.insertColumnAfter(this.sheet.getMaxColumns());
        this.sheet.getRange(this.row.head.key, this.sheet.getMaxColumns()).setValue("last");
        return this.sheet.getMaxColumns();
      }
    })();

    colNamesArr.forEach(colName => {
      if (!colName) { return; }
      if (this.head_key.includes(colName)) { return; }

      this.sheet.insertColumnBefore(colLast);
      this.sheet.getRange(this.row.head.key, colLast).setValue(colName);
      this.head_key.push(colName);
      colLast++;

    });

    this.init();
    this.reset();
  }



  /** @param {Obj} obj */
  updateObjInRow(obj, row) {

    // if (!this.getMap().has(`${obj.key}`)) { return; }
    // let row = this.getMap().get(`${obj.key}`).row;


    for (let i = this.head_key.length - 1; i >= 0; i--) {
      let key = this.head_key[i];
      if (!key) { continue; }
      // Logger.log(`i=${i}  key=${key}  val=${obj[key]}`);
      if (this.head_key[i] == "key") { continue; }
      if (obj[this.head_key[i]] == undefined) { continue; }

      let col = i;

      this.setValue(row, col, obj[this.head_key[i]]);
    }

  }


}


class MrClassSheetModelПроектыВИсполнение extends MrClassSheetModel {
  constructor(sheetName, context, rowConf = undefined) {
    super(sheetName, context, rowConf);

  }

  getMap() {
    if (!this.map) {
      super.getMap();

      let fileObjПроектыВИсполнение = getNewFileObjData();
      fileObjПроектыВИсполнение.НомерПроекта = getSettings().sheetNames.Проекты_в_Исполнение;
      fileObjПроектыВИсполнение.key = fileObjПроектыВИсполнение.НомерПроекта;
      getMrClassFile().readObj(fileObjПроектыВИсполнение);

      this.map.forEach(item => {
        let key = `${item.key}`;
        let fileValue = fileObjПроектыВИсполнение.value[key];



        if (typeof fileValue != "object") { return; }
        // Logger.log(key);
        // Logger.log(fileValue);
        for (let k in fileValue) {
          // Logger.log(k);
          // Logger.log(fileValue[k]);
          item[k] = fileValue[k];
        }

        // for (let k in item) {
        //   Logger.log(k);
        //   // Logger.log(fileValue[k]);
        //   // item[k] = fileValue[k];
        //   Logger.log(item[k]);
        // }


      });
    }
    return this.map;
  }


}









function getMrClassSheetModelРабочийЛист() {
  return MrClassSheetModelРабочийЛист;
}

class MrClassSheetModelРабочийЛист extends MrClassSheetModel {
  constructor(sheetName, context, rowConf = undefined) {
    super(sheetName, context, rowConf);

  }


  appendObj(obj) {
    obj.row = undefined;
    obj.key_add_data = new Date();
    this.saveObj(obj);
  }

}

