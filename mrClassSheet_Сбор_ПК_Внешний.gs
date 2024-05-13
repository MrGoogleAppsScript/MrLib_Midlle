
function serviceCopyMemFromBuyExternalToProduct() {
  Logger.log(`serviceCopyMemFromBuyExternalToProduct vvvvvvv`);
  let sheetBuy = getContext().getSheetBuy();
  let sheet_mem_buy = undefined;
  let sheet_mem_product = getContext().getSheetByName(getSettings().sheetName_mem_Оплаты);
  sheet_mem_product.hideSheet();
  try {
    sheet_mem_buy = sheetBuy.getExternalSpreadSheet().getSheetByName(getSettings().sheetName_mem_ОплатыВнешняя);
  } catch (err) {
    // return
    // mrErrToString(err);
  }
  if (!sheet_mem_buy) { return; }
  let rows = { body: { first: 3, last: sheet_mem_buy.getLastRow(), } };
  let cols = {
    body: {
      Поставщик: nr("B"),
      Номенклатура: nr("C"),
      РеальнаяЦена: nr("I"),
      РеальныйСрок: nr("J"),
      first: 1,
      last: sheet_mem_buy.getLastColumn(),
    }
  };
  let vls = sheet_mem_buy
    .getRange(rows.body.first, cols.body.first, rows.body.last - rows.body.first + 1, cols.body.last - cols.body.first + 1)
    .getValues()
    // .map((v, i, arr) => { return [i + rows.body.first].concat(v) })
    .filter(v => {
      if (!v[cols.body.Номенклатура - 1]) { return false; }
      if (!v[cols.body.Поставщик - 1]) { return false; }
      if (!v[cols.body.РеальнаяЦена - 1]) { return false; }
      // if (!v[cols.body.РеальныйСрок - 1]) { return false; }
      return true;
    });

  if (vls.length == 0) { return; }

  // let sheet_mem_product = getContext().getSheetByName(getSettings().sheetName_mem_Оплаты);
  sheet_mem_product.getRange(rows.body.first, cols.body.first, vls.length, vls[0].length).setValues(vls);
  // let rowLast = sheet_mem_product.getLastRow();
  let rowM = sheet_mem_product.getMaxRows();
  if (rowM <= vls.length + rows.body.first - 1) { return; }
  let rl = vls.length + rows.body.first - 1;
  sheet_mem_product.deleteRows(rl + 1, rowM - rl);
  Logger.log(`serviceCopyMemFromBuyExternalToProduct ^^^^^^^`)

}



function Сбор_КП_внешней() {


  // if (ttUrls.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) {
  //   return Сбор_КП_внешней_И_ИЗ_Закупка_Товара();
  // }

  // if (isTest()) {
    return Сбор_КП_внешней_И_ИЗ_Закупка_Товара();
  // }

  return Сбор_КП_внешней_test();
  // Utilities.sleep(30*1000);
  return Сбор_КП_внешней_bag();
}





function Сбор_КП_внешней_test() {
  // Logger.log("&&&&Сбор_КП_внешней_test&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&")
  let sheetBuy = getContext().getSheetBuy();

  let sheet_Сбор_ПК_Внешний = undefined;
  try {
    sheet_Сбор_ПК_Внешний = sheetBuy.getExternalSpreadSheet().getSheetByName(getSettings().sheetName_Лист_1);
  } catch (err) {
    mrErrToString(err);
  }
  if (!sheet_Сбор_ПК_Внешний) { return; }
  let classSheet_Сбор_ПК_Актуальный = new ClassSheet_Сбор_ПК_Актуальный_2(sheet_Сбор_ПК_Внешний);
  let { retArr, retBgArr } = classSheet_Сбор_ПК_Актуальный.Сбор_КП_внешней();

  return { retArr, retBgArr }
}

function Сбор_КП_внешней_И_ИЗ_Закупка_Товара() {

  Logger.log("Сбор_КП_внешней_И_ИЗ_Закупка_Товара");
  let sheetBuy = getContext().getSheetBuy();

  let sheet_Сбор_ПК_Внешний = undefined;
  try {
    sheet_Сбор_ПК_Внешний = sheetBuy.getExternalSpreadSheet().getSheetByName(getSettings().sheetName_Лист_1);
  } catch (err) {
    mrErrToString(err);
  }
  if (!sheet_Сбор_ПК_Внешний) {return sheetBuy.Сбор_КП_ИЗ_Закупка_Товара(); }

  let classSheet_Сбор_ПК_Актуальный = new ClassSheet_Сбор_ПК_Актуальный_2(sheet_Сбор_ПК_Внешний);

  let { retArr, retBgArr, len } = classSheet_Сбор_ПК_Актуальный.Сбор_КП_внешней();
  let Сбор_КП_ИЗ_Закупка_Товара = sheetBuy.Сбор_КП_ИЗ_Закупка_Товара();

  retArr = classSheet_Сбор_ПК_Актуальный.joinArr(len, retArr, Сбор_КП_ИЗ_Закупка_Товара.retArr);
  retBgArr = classSheet_Сбор_ПК_Актуальный.joinArr(len, retBgArr, Сбор_КП_ИЗ_Закупка_Товара.retBgArr);

  return { retArr, retBgArr };
}


class ClassSheet_Сбор_ПК_Актуальный {  // для два поля
  constructor(sheet = getContext().getSheetByName(getSettings().sheetName_Актуальный_Сбор_КП)) { // class constructor

    this.sheet = sheet;

    this.rowBodyFirst = 2;



    this.row_Header = 1
    this.make();
    this.rowBodyLast = this.findRowBodyLast();
  }

  make() {

    this.ok = true;


    let vls = this.sheet.getSheetValues(this.row_Header, 1, 1, this.sheet.getMaxColumns())[0];
    this.col_MAX_COL = vls.length;
    // Logger.log(`vls | ${vls}`);
    for (let i = 0; i < vls.length; i++) {
      if (!vls[i]) { continue; }
      vls[i] = fl_str(vls[i]);
    }


    let col = 0;

    col = 1 + vls.indexOf(fl_str("Лучшие цены")); if (!col) { this.ok = false; } else { this.col_TopFormula = col; }
    col = 1 + vls.indexOf(fl_str("Комментарий")); if (!col) { this.ok = false; } else { this.col_Coment = col; }
    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.col_FirstPrice = col + 3 + 3; }
    col = 1 + vls.indexOf(fl_str("Цены Конец")); if (!col) { this.ok = false; } else {
      this.col_LastPrice = col;
      this.col_LastPrice = (((this.col_LastPrice - this.col_FirstPrice) % 2) == 0 ? this.col_LastPrice : this.col_LastPrice - 1)


    }

    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.col_S = col + 6; }
    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.col_M = col; }



    col = 1 + vls.indexOf(fl_str("Номенклатура")); if (!col) { this.ok = false; } else { this.col_Product = col; }
    col = 1 + vls.indexOf(fl_str("Группа")); if (!col) { this.ok = false; } else { this.col_Group = col; }

    col = 1 + vls.indexOf(fl_str("№")); if (!col) { this.ok = false; } else { this.col_id = col; }

    col = 1 + vls.indexOf(fl_str("CommentBlokStart")); if (!col) { this.ok = false; } else { this.col_CommentBlokStart = col; }
    col = 1 + vls.indexOf(fl_str("CommentBlokEnd")); if (!col) { this.ok = false; } else { this.col_CommentBlokEnd = col; }
    col = 1 + vls.indexOf(fl_str("Кол-во")); if (!col) { this.ok = false; } else { this.col_Кол_во = col; }

    col = 1 + vls.indexOf(fl_str(getSettings().sheetName_Актуальный_Сбор_КП)); if (!col) { } else { this.col_Актуальный_Сбор_КП = col; } // Актуальный_Сбор_КП


    // Logger.log(getContext().getSobachkaMap().toArr());
    let sobachka = getContext().getSobachkaMap().map.get(fl_str(this.sheet.getSheetName()));
    // Logger.log(`sobachka = ${JSON.stringify(sobachka)}`)
    this.colSobachka = (sobachka ? sobachka.colSobachka : nr("С"));
    // this.colSobachka = nr("С");

  }

  findRowBodyLast() {
    // let ret = getContext().getRowSobachkaBySheetName(this.sheet.getSheetName());
    let ret = this.findRowSobachka()
    if (!ret) { ret = this.sheet.getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }


  findRowSobachka() {
    let sheet = this.sheet

    let rowSobachka = 0; // @ в строке с номером rowSobachka 
    let vcs = sheet.getRange(1, this.colSobachka, sheet.getLastRow(), 1).getValues();
    for (; rowSobachka < vcs.length; rowSobachka++) {
      if (vcs[rowSobachka].includes("@")) { break; }
    }
    rowSobachka++; // @ в строке с номером rowSobachka 
    return rowSobachka;
  }


  getProductMap() {
    let retMap = new Map();
    if (!this.col_LastPrice) {
      return retMap;
    }

    let vls_head = this.sheet.getRange(1, 1, 1, this.col_LastPrice).getValues();
    vls_head = vls_head.map((v, i, arr) => { return [i + 1].concat(v); })[0];
    let vls_prod = this.sheet.getRange(this.rowBodyFirst, 1, this.rowBodyLast - this.rowBodyFirst + 1, this.col_LastPrice).getValues();
    vls_prod = vls_prod.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    // Logger.log(vls_head);
    // Logger.log(vls_prod);
    vls_prod.forEach((v, i, arr) => {
      if (!v[this.col_Product]) { return; }

      let pr_name = v[this.col_Product];
      let pr_price = new Map();
      for (let i = this.col_FirstPrice; i < this.col_LastPrice; i += 3) {


        if (!`${v[i]}${v[i + 1]}`) { continue; }
        if (!vls_head[i]) { continue; }
        // Logger.log(`${vls_head[i]}=${v[i]}|${v[i + 1]}`);
        pr_price.set(vls_head[i], [v[i], v[i + 1]]);
      }
      retMap.set(pr_name, pr_price);
    });
    return retMap;
  }

  getEmailsArr() {
    let retArr = new Array();

    if (!this.col_FirstPrice) {
      return retArr;
    }

    if (!this.col_LastPrice) {
      return retArr;
    }

    let vls_head = this.sheet.getRange(1, this.col_FirstPrice, 1, this.col_LastPrice - this.col_FirstPrice - 3).getValues()[0];

    retArr = vls_head.filter((v, i, arr) => {
      if (!v) { return false; }
      if (!isEmail(v)) { return false; }
      return true;
    });

    return retArr;
  }


}


// 



class ClassSheet_Сбор_ПК_Актуальный_2 { // для три поля
  constructor(sheet) { // class constructor

    this.sheet = sheet;

    this.rowBodyFirst = 2;



    this.row_Header = 1
    this.make();
    this.rowBodyLast = this.findRowBodyLast();
  }

  make() {

    this.ok = true;


    let vls = this.sheet.getSheetValues(this.row_Header, 1, 1, this.sheet.getMaxColumns())[0];
    this.col_MAX_COL = vls.length;
    // Logger.log(`vls | ${vls}`);
    for (let i = 0; i < vls.length; i++) {
      if (!vls[i]) { continue; }
      vls[i] = fl_str(vls[i]);
    }


    let col = 0;

    col = 1 + vls.indexOf(fl_str("Лучшие цены")); if (!col) { this.ok = false; } else { this.col_TopFormula = col; }
    col = 1 + vls.indexOf(fl_str("Комментарий")); if (!col) { this.ok = false; } else { this.col_Coment = col; }
    col = 1 + vls.indexOf(fl_str("Цены Начало")); if (!col) { this.ok = false; } else { this.col_FirstPrice = col + 1; }
    col = 1 + vls.indexOf(fl_str("Цены Конец")); if (!col) { this.ok = false; } else {
      this.col_LastPrice = col;
      // this.col_LastPrice = (((this.col_LastPrice - this.col_FirstPrice) % 2) == 0 ? this.col_LastPrice : this.col_LastPrice - 1)
      while (((this.col_LastPrice - this.col_FirstPrice + 1) % 3) != 0) {
        this.col_LastPrice = this.col_LastPrice - 1;
      }

    }

    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.col_S = col + 6; }
    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.col_M = col; }



    col = 1 + vls.indexOf(fl_str("Номенклатура")); if (!col) { this.ok = false; } else { this.col_Product = col; }
    col = 1 + vls.indexOf(fl_str("Группа")); if (!col) { this.ok = false; } else { this.col_Group = col; }

    col = 1 + vls.indexOf(fl_str("№")); if (!col) { this.ok = false; } else { this.col_id = col; }

    col = 1 + vls.indexOf(fl_str("CommentBlokStart")); if (!col) { this.ok = false; } else { this.col_CommentBlokStart = col; }
    col = 1 + vls.indexOf(fl_str("CommentBlokEnd")); if (!col) { this.ok = false; } else { this.col_CommentBlokEnd = col; }
    col = 1 + vls.indexOf(fl_str("Кол-во")); if (!col) { this.ok = false; } else { this.col_Кол_во = col; }

    col = 1 + vls.indexOf(fl_str(getSettings().sheetName_Актуальный_Сбор_КП)); if (!col) { } else { this.col_Актуальный_Сбор_КП = col; } // Актуальный_Сбор_КП


    // Logger.log(getContext().getSobachkaMap().toArr());
    let sobachka = getContext().getSobachkaMap().map.get(fl_str(this.sheet.getSheetName()));
    // Logger.log(`sobachka = ${JSON.stringify(sobachka)}`)
    this.colSobachka = (sobachka ? sobachka.colSobachka : nr("С"));
    // this.colSobachka = nr("С");

  }

  findRowBodyLast() {
    // let ret = getContext().getRowSobachkaBySheetName(this.sheet.getSheetName());
    let ret = this.findRowSobachka()
    if (!ret) { ret = this.sheet.getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }


  findRowSobachka() {
    let sheet = this.sheet

    let rowSobachka = 0; // @ в строке с номером rowSobachka 
    let vcs = sheet.getRange(1, this.colSobachka, sheet.getLastRow(), 1).getValues();
    for (; rowSobachka < vcs.length; rowSobachka++) {
      if (vcs[rowSobachka].includes("@")) { break; }
    }
    rowSobachka++; // @ в строке с номером rowSobachka 
    return rowSobachka;
  }


  getProductMap() {
    let retMap = new Map();
    if (!this.col_LastPrice) {
      return retMap;
    }

    let vls_head = this.sheet.getRange(1, 1, 1, this.col_LastPrice).getValues();
    vls_head = vls_head.map((v, i, arr) => { return [i + 1].concat(v); })[0];

    let bg_head = this.sheet.getRange(1, 1, 1, this.col_LastPrice).getBackgrounds();
    bg_head = bg_head.map((v, i, arr) => { return [i + 1].concat(v); })[0];


    let vls_prod = this.sheet.getRange(this.rowBodyFirst, 1, this.rowBodyLast - this.rowBodyFirst + 1, this.col_LastPrice).getValues();
    vls_prod = vls_prod.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    // Logger.log(vls_head);
    // Logger.log(vls_prod);
    vls_prod.forEach((v, j, arr) => {
      if (!v[this.col_Product]) { return; }

      let pr_name = v[this.col_Product];
      let pr_price = new Map();
      for (let i = this.col_FirstPrice; i < this.col_LastPrice; i += 3) {
        // Logger.log(`${i}|${v[i]}${v[i + 1]}${v[i + 2]}`);
        if (!`${v[i]}${v[i + 1]}${v[i + 2]}`) { continue; }

        if (!vls_head[i]) { continue; }
        // Logger.log(`${vls_head[i]} = ${v[i]} | ${v[i + 1]} | ${v[i + 2]}`);

        let email_key = JSON.stringify({
          email: vls_head[i],
          bg: bg_head[i],
          col: i - this.col_FirstPrice,
        });

        pr_price.set(email_key, [v[i], v[i + 1], v[i + 2]]);
        // pr_price.set(vls_head[i], [v[i], v[i + 1]]);
      }
      retMap.set(pr_name, pr_price);
    });
    return retMap;
  }

  getEmailsArr() {
    let retArr = new Array();

    if (!this.col_FirstPrice) {
      return retArr;
    }

    if (!this.col_LastPrice) {
      return retArr;
    }


    let vls_head = this.sheet.getRange(1, this.col_FirstPrice, 1, this.col_LastPrice - this.col_FirstPrice - 3).getValues()[0];
    let bg_head = this.sheet.getRange(1, this.col_FirstPrice, 1, this.col_LastPrice - this.col_FirstPrice - 3).getBackgrounds()[0];

    let head = vls_head.map((v, i, arr) => {
      return {
        email: vls_head[i],
        bg: bg_head[i],
        col: i,
      }
    });

    // retArr = vls_head.filter((v, i, arr) => {
    //   if (!v) { return false; }
    //   if (!isEmail(v)) { return false; }
    //   return true;
    // });

    retArr = head.filter((v, i, arr) => {
      if (!v.email) { return false; }
      if (!isEmail(v.email)) { return false; }
      return true;
    });

    retArr = retArr.map((v, i, arr) => {
      return JSON.stringify(v);
    });

    return retArr;
  }




  /** @returns {[[]]} */
  Сбор_КП_внешней() {
    // // count = count + 1;

    // // Logger.log(retArr);
    // let sheetBuy = getContext().getSheetBuy();

    // let sheet_Сбор_ПК_Внешний = undefined;
    // try {
    //   sheet_Сбор_ПК_Внешний = sheetBuy.getExternalSpreadSheet().getSheetByName(getSettings().sheetName_Лист_1);
    // } catch {

    // }

    // let this = new ClassSheet_Сбор_ПК_Актуальный(sheet_Сбор_ПК_Внешний);

    let sheetList1 = getContext().getSheetList1();

    let productMap = this.getProductMap();
    // productMap.forEach((v, k, map) => { Logger.log(`${k}=${v.size}=${[...v.values()]}`) })
    let emailsArr = this.getEmailsArr();
    // Logger.log(emailsArr);


    let productArr = sheetList1.sheet.getRange(sheetList1.rowBodyFirst, sheetList1.columnProduct, sheetList1.rowBodyLast - sheetList1.rowBodyFirst + 1, 1).getValues();
    productArr = productArr.map((v, i, arr) => { return v[0] });
    // Logger.log(productArr);

    // let retArr = new Array(productArr.length+1);
    // let num_row = productArr.length + 1;
    let num_col = Math.max(emailsArr.length * 3 + 6, sheetList1.columnLastPrice - sheetList1.columnАктуальный_Сбор_КП);

    // let num_col = ;


    let retArr = new Array();

    // let h = [].concat(emailsArr.map((v, i, arr) => { return [v, undefined] }).flat());
    let h = new Array(num_col);
    for (let e = 0, i = 0; e < emailsArr.length; e++, i += 3) {
      h[i] = emailsArr[e];
    }



    // Logger.log(h)
    retArr.push(h);

    for (let i = 0; i < productArr.length; i++) {
      let r = new Array(num_col);
      retArr.push(r);
      let pr_name = productArr[i];
      if (!pr_name) { continue; }
      if (!productMap.has(pr_name)) { continue; }
      /** @type {Map} */
      let email_map = productMap.get(pr_name);
      if (email_map.size == 0) { continue; }
      // Logger.log("____________________________________")
      // Logger.log(JSON.stringify([...email_map.keys()]));
      for (let e = 0; e < h.length; e += 3) {
        // Logger.log(h[e]);
        if (!email_map.has(h[e])) { continue; }
        let v = email_map.get(h[e]);
        r[e] = v[0];
        r[e + 1] = v[1];
        r[e + 2] = v[2];
      }
      // Logger.log(r);
    }

    let retBgArr = [new Array(num_col)];

    retArr[0].forEach((r, i, arr) => {
      let email_key = JSON.parse(r);
      arr[i] = email_key.email;
      retBgArr[0][i] = email_key.bg;
    });


    // let знак = sheetList1.sheet.getRange("A1").getValue();
    // if (!знак) { знак = "(♠)"; }


    let знак = "(♠)";


    for (let e = 0; e < retArr[0].length; e += 3) {
      if (!retArr[0][e]) { continue; }
      // retArr[0][e] = `${retArr[0][e]} (* ${Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyy-MM-dd HH:mm:ss")}) `;
      retArr[0][e] = `${retArr[0][e]} ${знак}`;
    }

    let len = emailsArr.length * 3 + 6
    // let Сбор_КП_ИЗ_Закупка_Товара = { retArr, retBgArr } = getContext().getSheetBuy().Сбор_КП_ИЗ_Закупка_Товара

    return { retArr, retBgArr, len };
  }

  /** 
   * @param {[[]]} arrLeft
   * @param {[[]]} arrRight
   * @parem {number} len
  */
  joinArr(len, arrLeft, arrRight) {

    let retArr = arrLeft.map((r, i, arr) => {
      let l = r.length;
      return [].concat(r.slice(0, len), arrRight[i], r.slice(len + arrRight[i].length, l));
    });
    return retArr;
  }


}


















