class ClassSheet_ProductGroups {
  constructor() { // class constructor
    // this.sheet = getContext().getSheetByName("Группы товаров");
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Группы_товаров);

    this.names = {
      Номер: "№",
      Номенклатура: "Номенклатура",
      Характеристики: "Характеристики",
      ЕдИзм: "Ед. Изм",
      КолВо: "Кол-во",
      Файлы: "Файлы",
      Группа: "Группа",
      Доп1: "Ключ фразы1",
      Доп2: "Ключ фразы2",
      Доп3: "Ключ фразы3",
      // Ключ фразы1	Ключ фразы2	Ключ фразы3
      НеразосланныхПисем: "Неразосланных Писем",
    }

    this.row = {
      head: {
        first: 1,
        last: 1,
        key: 1,
      },
      body: {
        first: 4,
        last: this.sheet.getLastRow(),
      }
    }

    this.col = {
      first: 1,
      row: 0,
      Номер: undefined,
      Номенклатура: undefined,
      Характеристики: undefined,
      ЕдИзм: undefined,
      КолВо: undefined,
      Файлы: undefined,
      Группа: undefined,
      Доп1: undefined,
      Доп2: undefined,
      Доп3: undefined,
      НеразосланныхПисем: undefined,
      email: undefined,
      last: this.sheet.getLastColumn(),
    }

    this.head_key = (() => {
      try {
        return [].concat("row", this.sheet.getRange(this.row.head.key, 1, 1, this.col.last).getValues()[0]);
      } catch (err) { mrErrToString(err); return []; }
    })();

    let getColNumByName = /** @param {[]} keys*/ (keys, name,) => { if (!keys.includes(name)) { return undefined; } return keys.indexOf(name) }

    this.col.Номер = getColNumByName(this.head_key, this.names.Номер);
    this.col.Номенклатура = getColNumByName(this.head_key, this.names.Номенклатура);
    this.col.Характеристики = getColNumByName(this.head_key, this.names.Характеристики);
    this.col.ЕдИзм = getColNumByName(this.head_key, this.names.ЕдИзм);
    this.col.КолВо = getColNumByName(this.head_key, this.names.КолВо);
    this.col.Файлы = getColNumByName(this.head_key, this.names.Файлы);
    this.col.Группа = getColNumByName(this.head_key, this.names.Группа);
    this.col.Доп1 = getColNumByName(this.head_key, this.names.Доп1);
    this.col.Доп2 = getColNumByName(this.head_key, this.names.Доп2);
    this.col.Доп3 = getColNumByName(this.head_key, this.names.Доп3);
    this.col.НеразосланныхПисем = (() => {
      let r = getColNumByName(this.head_key, this.names.НеразосланныхПисем);
      if (r) { return r; }
      if (this.col.Доп3) { return this.col.Доп3 + 1; }
      return nr("K");
    })();
    this.col.email = (() => {
      if (this.col.НеразосланныхПисем) { return this.col.НеразосланныхПисем; }
      return nr("K");
    })();


    // this.columnProduct = nr("B");
    // this.columnProductG = nr("G");
    // this.columnNameGr = nr("F"); // столбец с имнами групп
    // this.columnEmail = nr("J");//столбик откуда начинают писать email


    this.columnProduct = (this.col.Номенклатура ? this.col.Номенклатура : nr("B"));
    this.columnProductG = (this.col.Доп1 ? this.col.Доп1 : nr("G"));
    this.columnNameGr = (this.col.Группа ? this.col.Группа : nr("F"));
    this.columnEmail = (this.col.email ? this.col.email : nr("K"));








    this.groupMap = undefined;

    this.emailBgMap = undefined;  // цвета которыми закрашенны емайлы ключь = email  значение = цвет массив
    this.emailGrupMap = undefined;  // группы в которых есть эти эмайлы  ключь = email  значение = группы массив


  }

  getGroupMap() {
    if (this.groupMap == undefined) {
      this.groupMap = this.makeGroupMap();
    }
    return this.groupMap;
  }


  getEmailBgMap() {
    if (this.emailBgMap == undefined) { this.emailBgMap = this.makeEmailBgMap(); }


    // this.emailBgMap.forEach(logMap);
    return this.emailBgMap;
  }


  getInfoForCounteragent(email) {

    let ret = new Object();
    ret["isEmailSent"] = -1;
    ret["bgArr"] = undefined;

    if (!email) { return ret; }
    email = fl_str(email);


    let bgMap = this.getEmailBgMap();
    // Logger.log(`ClassSheet_ProductGroups getInfoForCounteragent bgArr=${email}`);
    // bgMap.forEach(logMap);

    let bgArr = bgMap.get(email);

    // Logger.log(`ClassSheet_ProductGroups getInfoForCounteragent bgArr=${bgArr}`);


    if (bgArr != undefined) {
      let colorBgSentEmail = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
      if (bgArr.includes(colorBgSentEmail)) {
        ret["isEmailSent"] = true;
      }
      ret["bgArr"] = bgArr;
    }


    // Logger.log(`ClassSheet_ProductGroups getInfoForCounteragent ret.bgArr=${ret.bgArr}  ret.isEmailSent=${ret.isEmailSent}`);
    return ret;
  }


  getColorContragentForUnucRow(row, email) {
    let colorDef = getContext().getColor(getSettings().colorNames.DEFAULT);
    let colorGr = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);
    let colorCo = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);




    let retColor = colorDef;
    let contragent = getContext().getCounteragent(fl_str(email));
    if (contragent.getIsEmailResponse()) { retColor = colorCo; }
    // let prodGr = this.getGroupByRow(row);
    // let info = prodGr.getInfoForCounteragent(fl_str(email));


    return retColor;
  }

  makeEmailBgMap() {
    let retMap = new Map();
    let rangeEmail = this.sheet.getRange(4, this.columnEmail, this.sheet.getLastRow() - 4 + 1, this.sheet.getLastColumn() - this.columnEmail + 1);

    let vl = rangeEmail.getValues();
    let bg = rangeEmail.getBackgrounds();
    for (let i = 0; i < vl.length; i++) {
      for (let j = 0; j < vl[i].length; j++) {
        if (!vl[i][j]) { continue; }
        if (!isEmail(vl[i][j])) { continue; }
        let email = fl_str(vl[i][j]);
        let colorBg = bg[i][j];

        let arBg = retMap.get(email);
        if (!arBg) { arBg = new Array(); }

        if (arBg.includes(colorBg)) { continue; }
        arBg.push(colorBg);
        retMap.set(email, arBg);
      }
    }
    return retMap;
  }

  makeGroupMap() {
    let retMap = new Map();
    let nGr = new Array();
    try {
      nGr = this.sheet.getRange(1, this.columnNameGr, this.sheet.getLastRow(), 1).getValues();
    } catch {
      nGr = []
    }
    for (let i = 0; i < nGr.length; i++) {
      if (!nGr[i][0]) { continue; }
      let namesGr = fl_str(nGr[i][0]);


      if (retMap.has(namesGr)) { continue; }
      // this.productGroupMap.set(namesGr , new ProductGroup(row+1));
      //  this.groupMap.set(namesGr, undefined);   //  undefined  -- чтобы не нагружать вызов чаже всего нужна только одна из групп на каждый вызов 
      retMap.set(namesGr, i + 1);  // найменьшая строка в которой встречается впервые имя группы
    }
    return retMap;
  }



  getGroupByName(namesGr) {
    if (!namesGr) { return undefined; }
    // if (!namesGr) {    throw `аяяяй пустое имя гуппы  namesGr=${namesGr}`; }


    namesGr = fl_str(namesGr);
    if (!this.getGroupMap().has(namesGr)) { return undefined; }

    let gr = this.getGroupMap().get(namesGr);
    if (!(gr instanceof ProductGroup)) {  //зачеить номер строки 

      gr = new ProductGroup(gr, this.sheet);

      if (gr.isProductGroup) { this.groupMap.set(fl_str(gr.name), gr); }
    }
    return gr;
  }

  getGroupByRow(row) {
    let namesGr = this.sheet.getRange(row, this.columnNameGr).getValue();
    // Logger.log(`getGroupByRow() namesGr=${namesGr}`);
    if (!namesGr) {
      let gr = new ProductGroup(row, this.sheet);
      return gr;
    }

    return this.getGroupByName(fl_str(namesGr));
  }

  updateAllRows() {
    // this.updateProduct(1, this.sheet.getLastRow());
    this.updateProductAll(1, this.sheet.getLastRow());
    // updateProductAll
  }



  updateAll() {
    // updateSquare(rowStart, colStart, rowEnd, colEnd);
  }




  updateGroups(edit) {
    let range = edit.range;

    let rowStart = range.rowStart;
    let rowEnd = range.rowEnd;

    let colStart = range.columnStart;
    let colEnd = range.columnEnd;
    this.updateSquare(rowStart, colStart, rowEnd, colEnd);

  }



  updateSquare(rowStart, colStart, rowEnd, colEnd) {
    // let range = edit.range;

    // let rowStart = range.rowStart;
    // let rowEnd = range.rowEnd;

    // let colStart = range.columnStart;
    // let colEnd = range.columnEnd;

    let range = this.sheet.getRange(rowStart, colStart, rowEnd - rowStart + 1, colEnd - colStart + 1);
    let vl = range.getValues();
    let bg = range.getBackgrounds();


    let officialColor = getContext().getOfficialColor();
    let defaultColor = getContext().getDefaultColor();
    let colorFromProd = getContext().getIsFromProductListColor();


    for (let i = 0, row = rowStart; (i < vl.length && row <= rowEnd); i++, row++) {

      // Logger.log(`ClassSheet_ProductGroups.updateGroups row=${row}`);

      let gr = this.getGroupByRow(row);

      // Logger.log(`ClassSheet_ProductGroups.updateGroups gr=${gr.logString()}`);

      if (!gr) { continue; }
      if (!gr.isProductGroup) { continue; }
      if (row != gr.getFirstRow()) { continue; }


      let product = getContext().makeProduct(gr.bodyText);
      if (!product) { return; }
      for (let j = 0, col = colStart; ((j < vl[i].length) && (col <= colEnd)); j++, col++) {
        if (col < this.columnEmail) { continue; }
        let email = fl_str(vl[i][j]);
        if (!isEmail(email)) { continue; }

        let ca = getContext().getCounteragent(email);
        let productInfo = product.getInfoMap().get(email);

        let isOfficial = false;
        if (productInfo instanceof ProductInfo) { isOfficial = productInfo.isOfficial() }

        // логика раскразки дублируется в функции  pushEmailListInRow 
        let colorBg = defaultColor; // раскраска контрагента (поставщика)  цвет по умолчанию белый #ffffff
        colorBg = (isOfficial ? officialColor : ca.getColor());  //  цвет официального или статуса
        if (colorBg == defaultColor) { if (ca.isFromProductList()) { colorBg = colorFromProd; } } // если цвет остался по умолчанию то пробуем покрасить в цвет из базы товаров 
        // Logger.log(`ClassSheet_ProductGroups.updateGroups ${email}=${colorBg}`);
        bg[i][j] = colorBg;

      }

    }


    range.setBackgrounds(bg);
    // this.sheet.getRange().ser
  }





  updateProduct(rowStart, rowEnd) {
    //  при изменения в колонке nr("B") и "GHI"
    //  для каждой строчки создаем масив емайлов из базы товаров обединяем результаты в MAP
    // передаем МАР в функцию добовления емейлов на лист  
    let vlB = this.sheet.getRange(rowStart, this.columnProduct, rowEnd - rowStart + 1, 1).getValues();
    let vlGHI = this.sheet.getRange(rowStart, this.columnProductG, rowEnd - rowStart + 1, 3).getValues();
    let nGr = this.sheet.getRange(rowStart, this.columnNameGr, rowEnd - rowStart + 1, 1).getValues();

    // let mapRowProd = new Map(); //сюда будут собранны емайлы из листа тавары  для продукта из строки row листа Группы товаров   
    for (let i = 0; i < vlB.length; i++) {

      if (!nGr[i][0]) { continue; }
      let namesGr = fl_str(nGr[i][0]);
      let productBGHI = `${vlB[i][0]} ${vlGHI[i][0]} ${vlGHI[i][1]} ${vlGHI[i][2]}`;
      if (!productBGHI) { continue; }

      let pr = getContext().makeProduct(productBGHI);
      if (!pr) { continue; }
      if (pr.getInfoMap().size == 0) { continue; }

      let gr = this.getGroupByName(namesGr);
      // Logger.log(`ClassSheet_ProductGroups.updateProduct gr=${gr.logString()}`);
      if (gr.isProductGroup) {
        for (let [key, info] of pr.getInfoMap()) {
          gr.addInfo(info);
        }
      }

      // mapRowProd.set(row, pr);
    }

    let ar = new Array();
    // Logger.log(`ClassSheet_ProductGroups.updateProduct this.getGroupMap()`);
    // this.getGroupMap().forEach(logMap);// логи 
    // Logger.log(`ClassSheet_ProductGroups.updateProduct this.getGroupMap()`);


    for (let [nameGr, gr] of this.getGroupMap()) {
      if (!(gr instanceof ProductGroup)) { continue; }
      // Logger.log(`gr.newInfoMap. `);
      // gr.newInfoMap.forEach(logMap);// логи 

      let p = this.pushEmailListInRow(gr.getFirstRow(), gr, this.sheet);

      if (p > 0) { ar.push(gr.getFirstRow()) }

    }

    for (let i = 0; i < ar.length; i++) {
      unicEmailRow(ar[i]);
    }
  }



  pushEmailListInRow(row, gr, sheetProductGroups) {
    let infoMap = gr.newInfoMap;
    let ret = 0; // количество добавленных 
    let sec = this.columnEmail;

    let nrCol = sheetProductGroups.getLastColumn() - sec + 1;
    nrCol = (nrCol < 1 ? 1 : nrCol);


    let curentEmailList = sheetProductGroups.getSheetValues(row, sec, 1, nrCol);
    let curentEmailListForCompare = sheetProductGroups.getSheetValues(row, sec, 1, nrCol);
    let curentBackgrounds = sheetProductGroups.getRange(row, sec, 1, nrCol).getBackgrounds();
    for (let i = 0; i < curentEmailListForCompare[0].length; i++) {
      curentEmailListForCompare[0][i] = fl_str(curentEmailListForCompare[0][i]);
    }

    let j = 1 + this.getIndexOflastCompleted(curentEmailList[0]); //следующий за последним заполненным 

    let officialColor = getContext().getOfficialColor();
    let defaultColor = getContext().getDefaultColor();
    let colorFromProd = getContext().getIsFromProductListColor();
    // let defaultColor = getContext().getColor();

    for (var [keyI, info] of infoMap) {
      if (!curentEmailListForCompare[0].includes(info.nameForCompare)) {

        if (j >= curentEmailList[0].length) {
          curentEmailList[0].push("");
          curentBackgrounds[0].push(defaultColor);
          curentEmailListForCompare[0].push("")
        }


        // логика раскразки дублируется в функции  updateGroups(edit)
        let colorBg = defaultColor; // раскраска контрагента (поставщика)  цвет по умолчанию белый #ffffff
        colorBg = (info.isOfficial() ? officialColor : getContext().getCounteragent(info.nameForCompare).getColor());  //  цвет официального или статуса, если нет статуса то будет белым #ffffff
        if (colorBg == defaultColor) { if (getContext().getCounteragent(info.nameForCompare).isFromProductList()) { colorBg = colorFromProd; } } // если цвет остался по умолчанию то пробуем покрасить в цвет из базы товаров 




        curentEmailList[0][j] = info.name;
        curentBackgrounds[0][j] = colorBg;
        curentEmailListForCompare[0][j] = info.nameForCompare;
        j++;
        ret++;
      }
    }
    sheetProductGroups.getRange(row, sec, 1, curentBackgrounds[0].length).setBackgrounds(curentBackgrounds);
    sheetProductGroups.getRange(row, sec, 1, curentEmailList[0].length).setValues(curentEmailList);
    return ret;
  }




  getIndexOflastCompleted(curentEmailList) {
    // последний заполненный 
    for (let i = curentEmailList.length - 1; i >= 0; i--) {
      if (curentEmailList[i]) { return i; }
    }
    return -1;

  }


  getRowForProductGroup(row) {
    let group = getContext().getProductGroup(row);
    let ret = group.getFirstRow();
    return ret;
  }



  updateProductAll(rowStart, rowEnd) {
    //  при изменения в колонке nr("B") и "GHI"
    //  для каждой строчки создаем масив емайлов из базы товаров обединяем результаты в MAP
    // передаем МАР в функцию добовления емейлов на лист  
    let vlB = this.sheet.getRange(rowStart, this.columnProduct, rowEnd - rowStart + 1, 1).getValues();
    let vlGHI = this.sheet.getRange(rowStart, this.columnProductG, rowEnd - rowStart + 1, 3).getValues();
    let nGr = this.sheet.getRange(rowStart, this.columnNameGr, rowEnd - rowStart + 1, 1).getValues();

    // let mapRowProd = new Map(); //сюда будут собранны емайлы из листа тавары  для продукта из строки row листа Группы товаров   
    for (let i = 0; i < vlB.length; i++) {

      if (!nGr[i][0]) { continue; }
      let namesGr = fl_str(nGr[i][0]);
      let productBGHI = `${vlB[i][0]} ${vlGHI[i][0]} ${vlGHI[i][1]} ${vlGHI[i][2]}`;
      if (!productBGHI) { continue; }

      let pr = getContext().makeProduct(productBGHI);
      if (!pr) { continue; }
      // if (pr.getInfoMap().size == 0) { continue; }

      let gr = this.getGroupByName(namesGr);
      // Logger.log(`ClassSheet_ProductGroups.updateProduct gr=${gr.logString()}`);
      if (gr.isProductGroup) {
        for (let [key, info] of pr.getInfoMap()) {
          gr.addInfo(info);
        }
      }

      // mapRowProd.set(row, pr);
    }

    let ar = new Array();
    // Logger.log(`ClassSheet_ProductGroups.updateProduct this.getGroupMap()`);
    // this.getGroupMap().forEach(logMap);// логи 
    // Logger.log(`ClassSheet_ProductGroups.updateProduct this.getGroupMap()`);


    for (let [nameGr, gr] of this.getGroupMap()) {
      if (!(gr instanceof ProductGroup)) { continue; }
      // Logger.log(`gr.newInfoMap. `);
      // gr.newInfoMap.forEach(logMap);// логи 

      let p = this.pushEmailListInRow(gr.getFirstRow(), gr, this.sheet);

      if (p >= 0) { ar.push(gr.getFirstRow()) }

    }

    for (let i = 0; i < ar.length; i++) {
      unicEmailRow(ar[i]);
    }
  }


}







class ProductGroup {
  constructor(row, sheetPG) { // class constructor
    this.forRow = row;
    this.sheet = sheetPG;

    // this.columnProductB = nr("B");
    // this.columnProductG = nr("G");
    // this.columnNameGr = nr("F");

    this.columnProductB = getContext().getSheetProductGroups().columnProduct;
    this.columnProductG = getContext().getSheetProductGroups().columnProductG;
    this.columnNameGr = getContext().getSheetProductGroups().columnNameGr;



    this.name = undefined;
    this.isProductGroup = false;
    this.firstRow = undefined;
    this.secondRow = undefined;
    this.lastRow = undefined;
    this.allRows = new Array();

    this.bodyText = "";
    this.newInfoMap = new Map(); // данные  добавненные из вне возможно из надо будет вывести в на лист
    this.makeProductGroup(this.forRow, this.sheet);

    this.emailSentMap = undefined;
    this.emailSentMap = undefined;

    // this.logToConsole() ;
  }

  addInfo(info) {

    if (!(info instanceof ProductInfo)) { return; }

    let infoHas = this.newInfoMap.get(info.nameForCompare);

    if (!infoHas) {
      this.newInfoMap.set(info.nameForCompare, info);
    } else {
      infoHas.setOfficial(info.isOfficial());
    }

    return;


  }

  isEmailSent(email) {
    // email = fl_str(email);

    // return this.a;
  }

  getFirstRow() {
    return this.firstRow;
  }

  getSecondRow() {
    return this.secondRow;
  }

  toString() { return this.name; }
  logString() { return `ProductGroup(name="${this.name}", isProductGroup=${this.isProductGroup}, firstRow=${this.firstRow}, secondRow=${this.secondRow}, lastRow=${this.lastRow}, forRow=${this.forRow}, bodyText=${this.bodyText} )`; }

  logToConsole() { // class method
    console.log(this.logString());

  }


  makeProductGroup(row, sheetPG) {
    // if (!( row instanceof Number)) {  this.isProductGroup = false; return; }  //throw ( " эээ ");

    // let ss = SpreadsheetApp.getActive();
    // let sheetPG = ss.getSheetByName(getSettings().sheetName_Группы_товаров);

    if (!sheetPG.getRange(row, this.columnNameGr).getValue()) { this.isProductGroup = false; return; }   // в этой строке нет заголовка группы 
    this.name = fl_str(sheetPG.getRange(row, this.columnNameGr).getValue()); // заголовок группы 
    this.isProductGroup = true;
    let namesGroup = sheetPG.getRange(1, this.columnNameGr, sheetPG.getLastRow(), 1).getValues();



    let vlB = this.sheet.getRange(1, this.columnProductB, sheetPG.getLastRow(), 1).getValues();
    let vlGHI = this.sheet.getRange(1, this.columnProductG, sheetPG.getLastRow(), 3).getValues();





    // ищем все стоки для группы

    for (let i = 0; i < namesGroup.length; i++) {
      if (fl_str(namesGroup[i][0]) != this.name) { continue; }
      this.allRows.push(i + 1);

      let rowText = `${(vlB[i][0] ? vlB[i][0] : "")}${(vlGHI[i][0] ? " " + vlGHI[i][0] : "")}${(vlGHI[i][1] ? " " + vlGHI[i][1] : "")}${(vlGHI[i][2] ? " " + vlGHI[i][2] : "")}`;
      if (rowText) { this.bodyText = `${this.bodyText}, ${rowText}`; }
    }

    if (this.allRows.length > 0) {
      this.firstRow = this.allRows[0];
      this.secondRow = this.firstRow + 1;
      this.lastRow = this.allRows[this.allRows.length - 1];
    }

  }



}


/**Сделать формулу по подсчету неразосланных писем 
 * из листа "1-2 Запросы по товарным группам" 
 * Считаем кол-во емейлов которые отмечены как неразосланные 
 * */
function serviceПодсчетНеразосланныхПисем_() {
  let shГруппыТоваров = getContext().getSheetProductGroups().sheet;
  let nn = numberWithoutLetters();
  let vls = [[getContext().getSheetProductGroups().names.НеразосланныхПисем], [nn]];
  let colEmail = nc(getContext().getSheetProductGroups().col.email);
  shГруппыТоваров.getRange(`${colEmail}1:${colEmail}2`).setValues(vls);
  shГруппыТоваров.getRange(`${colEmail}1`).setNote(` ${nn} ${getContext().getSheetProductGroups().names.НеразосланныхПисем} ` + JSON.stringify(getContext().timeConstruct));
  shГруппыТоваров.getRange(`${colEmail}3:${nc(getContext().getSheetProductGroups().col.email + 1)}3`)
    .setValues([[getContext().getSheetProductGroups().names.НеразосланныхПисем, nn]]);

}


// без писем 
function numberWithoutLetters() {

ramoveAllProtection();

  // Logger.log("начало numberWithoutLetters");
  var ss = SpreadsheetApp.getActive();

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);

  // Logger.log([
  //   colorBrown,
  //   colorProd,
  //   colorGroup,

  // ])

  var sh = getContext().getSheetProductGroups().sheet;//  ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var lRow = sh.getLastRow() + 1;
  var lCol = sh.getLastColumn();
  var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();

  // Logger.log(JSON.stringify({ lRow, lCol, arrSh }));



  // var sec = nr("J");//столбик откуда начинают писать email (на листе)
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)

  var mapMail = new Map();
  // for (let i = 3; i < arrSh.length; i++) {
  for (let i = getContext().getSheetProductGroups().row.body.first - 1; i < arrSh.length; i++) {

    if (arrSh[i][0] != "" && arrSh[i - 1][0] == "") {
      // Logger.log(JSON.stringify(arrSh[i][0]));
      for (let j = sec - 1; j < arrSh[i + 1].length; j++) {
        if (arrSh[i + 1][j].toString().length > 0) {

          let bgc = sh.getRange(i + 1 + 1, j + 1).getBackground();
          // Logger.log(JSON.stringify([arrSh[i + 1][j], bgc]));
          if (
            bgc != colorBrown
            && bgc != colorProd
            && bgc != colorGroup
          ) {
            let key = arrSh[i + 1][j];
            mapMail.set(fl_str(key), undefined);
          }
        }
      }
    }
  }




  let retCount = 0;
  mapMail.forEach((v, emailAddress, m) => {

    let check = checkCounteragenth(emailAddress); // вернет emailAddress или цвет
    let isApprovedEmail = (check == emailAddress)  // одобрена отправка писем 
    if (isApprovedEmail) {
      retCount++;
    }

  });

  return retCount;
}






function unicEmailRow(unicRowIn) {
  //  Logger.log("unicEmailRow unicRowIn= " + unicRowIn);

  let firstRow = getContext().getProductGroup(unicRowIn).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
  if (firstRow != unicRowIn) { return; }


  var unicRow = unicRowIn - 1;
  var ss = SpreadsheetApp.getActive();
  // var sh = ss.getSheetByName("Группы товаров");
  // var sh = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var sh = getContext().getSheetProductGroups().sheet;
  var lRow = sh.getLastRow();
  var lCol = sh.getLastColumn();

  var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
  Logger.log(`"unicEmailRow" ${unicRowIn} arrSh=${JSON.stringify(arrSh, null, 2)}`);


  // var sec = nr("J");//столбик откуда начинают писать email
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)
  // let sheetProductGroups = getContext().getSheetProductGroups(); //



  if (arrSh[unicRow][0] != "" && arrSh[unicRow - 1][0] == "") {
    var arrEmail = [];
    var arrBG = [];
    arrBG.push([]);
    arrEmail.push([]);
    for (let j = sec - 1; j < arrSh[unicRow].length; j++) {
      if (arrSh[unicRow][j].toString().indexOf("@") > 0) {
        if (arrEmail[0].indexOf(arrSh[unicRow][j]) == -1) {
          arrEmail[0].push(arrSh[unicRow][j]);

          // let email = fl_str(arrSh[unicRow][j]);  // Todo
          // let color = sheetProductGroups.getColorContragentForUnucRow(unicRowIn, email);
          // Logger.log(`${email} | ${color}`);
          // // let color =   contragent.getColor();     
          // arrBG[0].push(color);
        }
      }
    }
    sh.getRange(unicRow + 2, sec, 1, lCol).protect().remove();
    sh.getRange(unicRow + 2, sec, 1, lCol).clearContent();   // добовляет лишнии пустые колонки 
    // sh.getRange(unicRow + 2, sec, 1, lCol - sec + 1).clearContent();   // Не  добовляет лишнии пустые колонки 
    if (arrEmail[0].length > 0) {
      var r = sh.getRange(unicRow + 2, sec, arrEmail.length, arrEmail[0].length)

      r.setValues(arrEmail);
      // r.setBackgrounds(arrBG);
      // Logger.log("unicEmailRow arrEmail= " + arrEmail);
      sh.getRange(unicRow + 2, sec, 1, lCol).protect().setWarningOnly(true);
    }
  }
}


function unicEmail() {
  Logger.log("начало");

  ramoveAllProtection();

  var ss = SpreadsheetApp.getActive();
  // var sh = ss.getSheetByName("Группы товаров");
  var sh = getContext().getSheetProductGroups().sheet;


  var lRow = sh.getLastRow();
  var lCol = sh.getLastColumn();
  var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
  var flag = false;
  // var sec = nr("J");//столбик откуда начинают писать email
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)
  for (let i = getContext().getSheetProductGroups().row.body.first - 1; i < arrSh.length; i++) {
    if (arrSh[i][0] != "") {
      if (flag == false) {
        flag = true;
        var arrEmail = [];
        arrEmail.push([]);
        for (let j = sec - 1; j < arrSh[i].length; j++) {
          if (arrSh[i][j].toString().indexOf("@") > 0) {
            if (arrEmail[0].indexOf(arrSh[i][j]) == -1) {
              arrEmail[0].push(arrSh[i][j]);
            }
          }
        }
        Logger.log(arrEmail);

        sh.getRange(i + 2, sec, 1, lCol).clearContent();
        if (arrEmail[0].length > 0) {

          // var r = sh.getRange(i + 2, sec, arrEmail.length, arrEmail[0].length)
          // r.setValues(arrEmail);
          // r.protect().setWarningOnly(true);


          //var protection =r.protect();
          //var me = Session.getEffectiveUser();
          //protection.addEditor(me);
          //.setWarningOnly(true); ;//.setDescription('Отфильтованные email запрещено редактировать.');
          //r.protect().setWarningOnly(true);            
        }
      }//if (flag == false)
    } else {
      flag = false;
    }
  }
  Logger.log("конец");
}

function ramoveAllProtection() {
  // var ss = SpreadsheetApp.getActive();
  // var sh = ss.getSheetByName("Группы товаров");
  var sh = getContext().getSheetProductGroups().sheet;

  var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
}




function sendEmail(map = new Map()) { // без файлов
  return;
  var ui = SpreadsheetApp.getUi();
  let url = (() => { return getContext().getSpreadSheet().getUrl(); })();

  let НомерПроекта = getContext().getНомерПроекта();
  var shSet = getContext().getSheetНастройки_писем().sheet

  let tm = getContext().getSheetНастройки_писем().getTemplateForSendMails();
  if (!tm) {
    throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю. или нет доступа к файлу`;
    throw `Для ${getContext().getUserEmail()} нет шаблона`;
  }
  if (tm["Заявка"] !== true) {
    throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`
  }
  Logger.log(tm["Заявка"]);


  // Заявка
  // Заявка Тема письма
  // Заявка Текст до таблицы
  // Заявка Текст после таблицы
  // Заявка Вложение
  // Заявка Адрес ответа

  let mailSubject = tm["Заявка Тема письма"];//тема письма
  let mailText1 = tm["Заявка Текст до таблицы"];//текст До
  let mailText2 = tm["Заявка Текст после таблицы"];//Текст После
  let mailAttachmentsId = tm["Заявка Вложение"];//Прикрепленный файл
  let mailReplyTo = tm["Заявка Адрес ответа"];//Куда отвечать
  let mailTextProject = (() => { try { return `<br><br>Проект № ${SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]}` } catch (err) { return "" } })();

  // Заявка Таблица №
  // Заявка Таблица Наименование
  // Заявка Таблица Описание
  // Заявка Таблица Ед. изм.
  // Заявка Таблица Кол-во
  let f1f5 = [[
    tm["Заявка Таблица №"],
    tm["Заявка Таблица Наименование"],
    tm["Заявка Таблица Описание"],
    tm["Заявка Таблица Ед. изм."],
    tm["Заявка Таблица Кол-во"],
  ]];

  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);
  if (emailQuotaRemaining < map.size) {
    var result = ui.alert(
      'Скрипт рассылки',
      'Количество писем к рассылке ' + map.size + '. Лимит рассылки ' + emailQuotaRemaining,
      ui.ButtonSet.OK
    );
    return false;
  }
  var arrKeys = [];


  var arrHead = shSet.getRange(1, 15, 1, 5).getValues();

  //ui.alert("arrHeadlen=" + arrHead[0].length);

  for (var [key, value] of map) {
    arrKeys.push(key.toString());
  }

  // Mistervova@mail.ru 25.12.2020 фильтрация значений согласно тз 22.12.2020 //Задача№5 Доработать отправку писем
  getContext().upDate();


  let logEmailCounter = 0;

  // let toMemArrMails = new Array();

  let toMemMailsMap = new Map();


  for (let i = 0; i < arrKeys.length; i++) {

    //Добавляем заголовки
    /** @type {Array} */
    var arrNoHead = map.get(arrKeys[i]);
    let arrNotFiltred = JSON.parse(JSON.stringify(arrNoHead));


    arrNoHead = mrFilterArrNoHead(arrNoHead);


    arrNotFiltred.unshift(arrHead[0]);
    arrNoHead.unshift(arrHead[0]);

    if (f1f5[0][3] == true && f1f5[0][4] == true) {//меняем Ед. изм.	и Кол-во
      var len = arrNoHead[0].length;

      for (let f = 0; f < arrNoHead.length; f++) {


        var a = arrNoHead[f][len - 1];
        arrNoHead[f][len - 1] = arrNoHead[f][len - 2];
        arrNoHead[f][len - 2] = a;

        a = arrNotFiltred[f][len - 1];
        arrNotFiltred[f][len - 1] = arrNotFiltred[f][len - 2];
        arrNotFiltred[f][len - 2] = a;

      }
    }



    var mailHtml = dataToHtmlTable(arrNoHead);




    let mail = {
      to: arrKeys[i],
      replyTo: mailReplyTo,
      subject: mailSubject,
      htmlBody: mailText1 + "<br>" + mailHtml + mailText2 + mailTextProject,
      // attachments: [pdfFileBlob]
    };

    if (mailAttachmentsId.toString().length > 0) {
      var pdfFileBlob = DriveApp.getFileById(mailAttachmentsId).getBlob();
      mail["attachments"] = [pdfFileBlob];

    } else {

    }

    let mailColor = undefined; // цвет при отказе если undefined значеть отправили письмо 
    let emailAddress = arrKeys[i];
    let check = checkCounteragenth(emailAddress); // вернет emailAddress или цвет
    // console.log("emailAddress= " + emailAddress + " check=" + check);
    // getContext().getCounteragent(emailAddress).logToConsole();



    let isApprovedEmail = (check == emailAddress)  // одобрена отправка писем 
    if (isApprovedEmail) {    // вернулся emailAddress значить шлем письмо 

      getContext().getSheetНастройки_писем().sendMail(mail);
      try {
        toMemMailsMap.set(emailAddress, {
          url,
          mail,
          ДатаОтправки: getContext().timeConstruct,
          from: Session.getEffectiveUser().getEmail(),
          to: emailAddress,
          key: JSON.stringify({
            from: Session.getEffectiveUser().getEmail(),
            to: emailAddress,
            ДатаОтправки: getContext().timeConstruct,
            url,
          }),
          Номенклатуры: arrNoHead.map((rr, i) => {
            try {
              if (i == 0) { return; }
              let r = arrNotFiltred[i];
              return {
                НомерПроекта,
                to: emailAddress,
                url,
                Номенклатура: r[1],
                ДатаОтправки: getContext().timeConstruct,
                // mail,
                from: Session.getEffectiveUser().getEmail(),
                key: `${JSON.stringify({
                  НомерПроекта,
                  to: emailAddress,
                  url,
                  Номенклатура: r[1],
                  ДатаОтправки: getContext().timeConstruct,
                })}`,
                r,
              };
            } catch (err) { mrErrToString(err); }

          }),
        });

      } catch (err) { mrErrToString(err); }

      // SpreadsheetApp.getUi().alert(` отправить ${arrKeys[i]} ` );
      logEmailCounter++;
    } else { // не отправлять письма, check -это  цвет причины отказа 
      mailColor = check;
    }


    showMail(shSet, mail, isApprovedEmail, mailColor);
    setMailBackground(arrKeys[i], mailColor);

  }

  getContext().getSheetНастройки_писем().addToMemSentMails(toMemMailsMap, url);
  //логируем кол-во писем
  logEmail(logEmailCounter);
  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);
}



function showMail(shSet, mail, isApproved, color) {

  rowContinue++;


  let arr = [
    [(isApproved ? "одобрено" : "отказано"), "", "", "", ""],
    ["to", "replyTo", "subject", "htmlBody", "attachments"],
    [mail["to"], mail["replyTo"], mail["subject"], mail["htmlBody"], mail["attachments"]],
  ]

  // let range = shSet.getRange(rowContinue, nr("N"), 3, 5);
  // let value = range.get
  shSet.getRange(rowContinue, nr("N"), 3, 5).setValues(arr);
  shSet.getRange(rowContinue, nr("N"), 3, 5).setBackground(color);

  rowContinue += 3;

}


let rowContinue = 2;

function createEmail() { // без файлов
  return;
  Logger.log("начало");
  var ui = SpreadsheetApp.getUi();
  // var ss = SpreadsheetApp.getActive();
  var userMail = Session.getEffectiveUser().getEmail();
  //Отсылать ли письма с этого аккаунтв
  var result = ui.alert(
    'Скрипт рассылки',
    'Аккаунт для отправки писем: \n' + userMail.toString().toUpperCase(), // Сообщение
    ui.ButtonSet.YES_NO // Кнопки
  );
  if (result == ui.Button.YES) {
  } else {
    return;
  }

  var shSet = getContext().getSheetНастройки_писем().sheet;

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);

  // var f1f5 = shSet.getRange("C6:G6").getValues();

  let tm = getContext().getSheetНастройки_писем().getTemplateForSendMails();
  if (!tm) { throw `Для ${getContext().getUserEmail()} нет шаблона`; }
  if (tm["Заявка"] !== true) { `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`; }
  let f1f5 = [[
    tm["Заявка Таблица №"],
    tm["Заявка Таблица Наименование"],
    tm["Заявка Таблица Описание"],
    tm["Заявка Таблица Ед. изм."],
    tm["Заявка Таблица Кол-во"],
  ]];



  //Logger.log("colorDrown = " + colorDrown + " colorGreen = " + colorGreen);
  //Logger.log("f1 = " + f1 + " f2 = " + f2 + " f3 = " + f3 + " f4 = " + f4 + " f5 = " + f5);
  //Logger.log("f1f5 = " + f1f5);
  // var sh = ss.getSheetByName("Группы товаров");
  // var sh = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var sh = getContext().getSheetProductGroups().sheet;
  var lRow = sh.getLastRow() + 1;
  var lCol = sh.getLastColumn();
  var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
  // var sec = 10;//столбик откуда начинают писать email (на листе)
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)

  var mapMail = new Map();
  // for (let i = 3; i < arrSh.length; i++) {
  for (let i = getContext().getSheetProductGroups().row.body.first - 1; i < arrSh.length; i++) {

    if (arrSh[i][0] != "" && arrSh[i - 1][0] == "") {
      var arrForGroup = [];
      let k = 0;
      try {
        while (arrSh[i + k][0] != "") {
          var arrTemp = takeFromRow(arrSh[i + k], f1f5);
          //Logger.log((i+k)*1 + " = " + arrTemp);
          arrForGroup.push(arrTemp);
          k++;
          if ((i + k) >= arrSh.length) { break; }
        }
      } catch (err) { mrErrToString(err); }
      //Logger.log(i+1);
      //Logger.log(arrForGroup);
      //Адреса
      for (let j = sec - 1; j < arrSh[i + 1].length; j++) {
        if (arrSh[i + 1][j].toString().length > 0) {
          if (sh.getRange(i + 1 + 1, j + 1).getBackground() != colorBrown && sh.getRange(i + 1 + 1, j + 1).getBackground() != colorProd && sh.getRange(i + 1 + 1, j + 1).getBackground() != colorGroup) {
            let key = arrSh[i + 1][j];
            let val = arrForGroup;

            if (mapMail.has(key)) {
              let tempArr = mapMail.get(key);
              let newTempArr = tempArr.concat(val);
              mapMail.set(key, newTempArr);
            } else {
              mapMail.set(key, val)
            }
          }
        }
      }
    }
  }


  //Проставляем нумерацию позиций 
  var mapNew = new Map();
  var arrKeys = [];
  for (var [key, value] of mapMail) {
    arrKeys.push(key.toString());
  }
  for (let k = 0; k < arrKeys.length; k++) {
    var arrFromMapMail = mapMail.get(arrKeys[k]);
    var arrNew = copyArr(arrFromMapMail);
    if (f1f5[0][0] == true) {
      for (let i = 0; i < arrNew.length; i++) {
        arrNew[i][0] = i + 1;
      }
    }
    mapNew.set(arrKeys[k], arrNew);
  }

  //V) Выводим софомированные шаблоны для писем
  var row = 2;
  var col = 14;
  //  MisterVova@mail.ru 22.12.2020  заменил строку // данные  после колонк T включительно  сохранить 
  //  shSet.getRange(row-1, col,shSet.getLastRow(),shSet.getLastColumn()).clearContent();
  // очистка шести колонок после col =14
  shSet.getRange(row - 1, col, shSet.getLastRow(), 6).clearContent();
  shSet.getRange(row - 1, col, shSet.getLastRow(), 6).setBackground("#ffffff");
  //логируем кол-во писем
  // logEmail(arrKeys.length);// перемещино конец функции sendEmail  // MisterVova@mail.ru 22.12.2020 
  //заголовок
  var f1f5Head = shSet.getRange("C5:G5").getValues();
  var arrTempHead = takeFromRow(f1f5Head.flat(), f1f5);
  shSet.getRange(row - 1, col + 1, 1, arrTempHead.length).setValues([arrTempHead]);



  for (var [keyM, valM] of mapNew) {
    shSet.getRange(row, col).setValue(keyM);
    shSet.getRange(row, col + 1, valM.length, valM[0].length).setValues(valM);
    row = row + valM.length + 1;
  }


  rowContinue = row;
  sendEmail(mapNew);

  Logger.log("конец");
}


function addNewTriggerMenuCreateEmailРассылкаЗапросов(triggerName = "MrLib_top.triggerMenuCreateEmailРассылкаЗапросов") {
  Logger.log(`Триггер проверяем на наличие триггера ${triggerName}`);
  if (typeof triggerName != "string") { return; }
  if (triggerName == "") { return; }
  // let triggers = ScriptApp.getScriptTriggers();
  let triggers = ScriptApp.getProjectTriggers();
  // triggers.forEach(trigger => Logger.log(trigger.getHandlerFunction()));
  let triggersName = triggers.map(trigger => trigger.getHandlerFunction());

  if (!triggersName.includes(triggerName)) {
    Logger.log(`Триггер ${triggerName} не найден. Добовляем.`);
    // ScriptApp.newTrigger(triggerName).forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
    ScriptApp.newTrigger(triggerName).timeBased().everyMinutes(10).create();

    getContext().getSheetProductGroups().sheet
      .getRange(`${nc(getContext().getSheetProductGroups().col.email + 4)}2:${nc(getContext().getSheetProductGroups().col.email + 5)}3`)
      .setValues([[true, new Date()], ["Триггер Добовлен", new Date()]]);


  } else {
    Logger.log(`Триггер ${triggerName} найден.`);

    getContext().getSheetProductGroups().sheet
      .getRange(`${nc(getContext().getSheetProductGroups().col.email + 4)}2:${nc(getContext().getSheetProductGroups().col.email + 5)}3`)
      .setValues([[true, new Date()], ["Триггер Запущен", getContext().timeConstruct]]);
  }
}
function menuTriggerMenuCreateEmailРассылкаЗапросов() {
  triggerMenuCreateEmailРассылкаЗапросов("Из меню menuTriggerMenuCreateEmailРассылкаЗапросов", true)
}


function triggerMenuCreateEmailРассылкаЗапросов(triggerInfo, ПодтвержденияАккаунта = false) {
  // return
  // if (!testUrls.includes(this.url)) { return; }
  let triggerName = "MrLib_top.triggerMenuCreateEmailРассылкаЗапросов";
  // let ПодтвержденияАккаунта = true;
  Logger.log(JSON.stringify(triggerInfo));
  Logger.log(JSON.stringify(ПодтвержденияАккаунта));
  Logger.log(JSON.stringify(Session.getEffectiveUser().getEmail()));
  // if (Session.getEffectiveUser().getEmail() == "zakupka.steklosvet@gmail.com") {
  //   deleteTriggerByName(triggerName); return;
  // }
  if (ПодтвержденияАккаунта) {
    var userMail = Session.getEffectiveUser().getEmail();
    var ui = SpreadsheetApp.getUi();
    Logger.log("Подтверждения Аккаунта " + userMail.toString().toUpperCase());
    var result = ui.alert(
      'Создать триггер рассылки',
      'Аккаунт для отправки писем: \n' + userMail.toString().toUpperCase() + '\n после отправки писем триггер сам удалится', // Сообщение
      ui.ButtonSet.YES_NO // Кнопки
    );
    if (result == ui.Button.YES) {
    } else {
      return;
    }
  }

  addNewTriggerMenuCreateEmailРассылкаЗапросов(triggerName);

  ПодтвержденияАккаунта = false;
  try {
    serviceEmailQuotaRemaining_();
    if (menuCreateEmailРассылкаЗапросов(ПодтвержденияАккаунта) != true) { return };
  } catch (err) {
    let m = mrErrToString(err);
    if (m.includes(" нет шаблона")) {
      deleteNewTriggerMenuCreateEmail(triggerName);
      return;
    }
  }


  try {
    serviceПодсчетНеразосланныхПисем_();
    serviceEmailQuotaRemaining_()
    // var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    // let colEmail = nc(getContext().getSheetProductGroups().col.email + 1);
    // let email_QuotaRemaining = [[Session.getEffectiveUser().getEmail() + "\nЛимит рассылки"], [emailQuotaRemaining]];
    // getContext().getSheetProductGroups().sheet.getRange(`${colEmail}1:${colEmail}2`).setValues(email_QuotaRemaining);
    // // let c1= nc(getContext().getSheetProductGroups().col.email + 2) 
    // getContext().getSheetProductGroups().sheet
    //   .getRange(`${nc(getContext().getSheetProductGroups().col.email + 2)}3:${nc(getContext().getSheetProductGroups().col.email + 3)}3`)
    //   .setValues([[Session.getEffectiveUser().getEmail() + "\nЛимит рассылки", emailQuotaRemaining]]);
  } catch (err) {
    mrErrToString(err);
  }


  deleteNewTriggerMenuCreateEmail(triggerName);
}
function serviceEmailQuotaRemaining_() {
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  let colEmail = nc(getContext().getSheetProductGroups().col.email + 1);
  let email_QuotaRemaining = [[Session.getEffectiveUser().getEmail() + "\nЛимит рассылки"], [emailQuotaRemaining]];
  getContext().getSheetProductGroups().sheet.getRange(`${colEmail}1:${colEmail}2`).setValues(email_QuotaRemaining);
  // let c1= nc(getContext().getSheetProductGroups().col.email + 2) 
  getContext().getSheetProductGroups().sheet
    .getRange(`${nc(getContext().getSheetProductGroups().col.email + 2)}3:${nc(getContext().getSheetProductGroups().col.email + 3)}3`)
    .setValues([[Session.getEffectiveUser().getEmail() + "\nЛимит рассылки", emailQuotaRemaining]]);
}


function deleteNewTriggerMenuCreateEmail(triggerName) {
  serviceПодсчетНеразосланныхПисем_();
  getContext().getSheetProductGroups().sheet
    .getRange(`${nc(getContext().getSheetProductGroups().col.email + 4)}2:${nc(getContext().getSheetProductGroups().col.email + 5)}3`)
    .setValues([
      [false, new Date()],
      ["Триггер Удалён", new Date()]
    ]);
  deleteTriggerByName(triggerName);
}



function menuCreateEmailРассылкаЗапросов(ПодтвержденияАккаунта = true) {
  Logger.log("начало");

  // var ss = SpreadsheetApp.getActive();
  var userMail = Session.getEffectiveUser().getEmail();
  //Отсылать ли письма с этого аккаунтв

  if (ПодтвержденияАккаунта) {
    var ui = SpreadsheetApp.getUi();
    Logger.log("Подтверждения Аккаунта " + userMail.toString().toUpperCase());
    var result = ui.alert(
      'Скрипт рассылки',
      'Аккаунт для отправки писем: \n' + userMail.toString().toUpperCase(), // Сообщение
      ui.ButtonSet.YES_NO // Кнопки
    );
    if (result == ui.Button.YES) {
    } else {
      return;
    }
  }

  var shSet = getContext().getSheetНастройки_писем().sheet;

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);

  let tm = getContext().getSheetНастройки_писем().getTemplateForSendMails();
  if (!tm) { throw new Error(`Для ${getContext().getUserEmail()} нет шаблона`); }
  if (tm["Заявка"] !== true) { `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`; }
  let f1f5 = [[
    tm["Заявка Таблица №"],
    tm["Заявка Таблица Наименование"],
    tm["Заявка Таблица Описание"],
    tm["Заявка Таблица Ед. изм."],
    tm["Заявка Таблица Кол-во"],
    tm["Заявка Таблица Файлы"],
  ]];

  let f1f5Head = [[
    "№",
    "Наименование",
    "Описание",
    "Ед. изм.",
    "Кол-во",
    "Файлы",
  ]];


  var sh = getContext().getSheetProductGroups().sheet;
  var lRow = sh.getLastRow() + 1;
  var lCol = sh.getLastColumn();
  var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
  // var sec = 10;//столбик откуда начинают писать email (на листе)
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)
  //  Logger.log(arrSh)
  var mapMail = new Map();
  // for (let i = 3; i < arrSh.length; i++) {
  for (let i = getContext().getSheetProductGroups().row.body.first - 1; i < arrSh.length; i++) {

    if (arrSh[i][0] != "" && arrSh[i - 1][0] == "") {
      var arrForGroup = [];
      let k = 0;
      try {
        while (arrSh[i + k][0] != "") {
          var arrTemp = takeFromRow(arrSh[i + k], f1f5);
          // Logger.log((i + k) * 1 + " = " + arrTemp);
          arrForGroup.push(arrTemp);
          k++;
          if ((i + k) >= arrSh.length) { break; }
        }
      } catch (err) { mrErrToString(err); }
      // Logger.log(i + 1);
      // Logger.log(arrForGroup);
      //Адреса
      for (let j = sec - 1; j < arrSh[i + 1].length; j++) {
        if (arrSh[i + 1][j].toString().length > 0) {
          if (sh.getRange(i + 1 + 1, j + 1).getBackground() != colorBrown && sh.getRange(i + 1 + 1, j + 1).getBackground() != colorProd && sh.getRange(i + 1 + 1, j + 1).getBackground() != colorGroup) {
            let key = arrSh[i + 1][j];
            let val = arrForGroup;

            if (mapMail.has(key)) {
              let tempArr = mapMail.get(key);
              let newTempArr = tempArr.concat(val);
              mapMail.set(key, newTempArr);
            } else {
              mapMail.set(key, val)
            }
          }
        }
      }
    }
  }
  Logger.log(mapMail.size)

  //Проставляем нумерацию позиций 
  var mapNew = new Map();
  var arrKeys = [];
  for (var [key, value] of mapMail) {
    arrKeys.push(key.toString());
  }
  for (let k = 0; k < arrKeys.length; k++) {
    var arrFromMapMail = mapMail.get(arrKeys[k]);
    var arrNew = copyArr(arrFromMapMail);
    if (f1f5[0][0] == true) {
      for (let i = 0; i < arrNew.length; i++) {
        arrNew[i][0] = i + 1;
      }
    }
    mapNew.set(arrKeys[k], arrNew);
  }

  //V) Выводим софомированные шаблоны для писем
  var row = 2;
  var col = 14;

  // очистка шести колонок после col =14
  shSet.getRange(row - 1, col, shSet.getLastRow(), f1f5Head.flat().length + 1).setBackground("#ffffff");
  shSet.getRange(row - 1, col, shSet.getLastRow(), f1f5Head.flat().length + 1).clearContent();


  var arrTempHead = takeFromRow(f1f5Head.flat(), f1f5);
  shSet.getRange(row - 1, col + 1, 1, arrTempHead.length).setValues([arrTempHead]);



  for (var [keyM, valM] of mapNew) {
    shSet.getRange(row, col).setValue(keyM);
    shSet.getRange(row, col + 1, valM.length, valM[0].length).setValues(valM);
    row = row + valM.length + 1;
  }


  let dellEmailArr = new Array();
  mapNew.forEach((v, e, m) => {
    let check = checkCounteragenth(e);
    if (check != e) {
      dellEmailArr.push(e);
    }
  });

  dellEmailArr.forEach(e => {
    mapNew.delete(e);
  });



  if (mapNew.size == 0) {
    return true;
  }
  rowContinue = row;
  let limitQuotaRemaining = sendEmailРассылкаЗапросов(mapNew);

  if (limitQuotaRemaining == true && mapNew.size > 0) {
    return false;
  }

  return true;

}

function sendEmailРассылкаЗапросов(map = new Map()) {

  let url = (() => { return getContext().getSpreadSheet().getUrl(); })();

  let НомерПроекта = getContext().getНомерПроекта();
  var shSet = getContext().getSheetНастройки_писем().sheet

  let tm = getContext().getSheetНастройки_писем().getTemplateForSendMails();
  if (!tm) {
    throw new Error(`В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю. или нет доступа к файлу`);

  }
  if (tm["Заявка"] !== true) {
    throw new Error(`В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`);
  }
  // Logger.log(tm["Заявка"]);


  // Заявка
  // Заявка Тема письма
  // Заявка Текст до таблицы
  // Заявка Текст после таблицы
  // Заявка Вложение
  // Заявка Адрес ответа
  // Заявка Текст С Номером Проекта
  // Заявка Тема письма Добавить Номер проекта
  let mailSubject = tm["Заявка Тема письма"];//тема письма

  // if (isTest()) {
  //   mailSubject = 
  // }

  let mailText1 = tm["Заявка Текст до таблицы"];//текст До
  let mailText2 = tm["Заявка Текст после таблицы"];//Текст После
  let mailAttachmentsId = tm["Заявка Вложение"];//Прикрепленный файл
  let mailReplyTo = tm["Заявка Адрес ответа"];//Куда отвечать
  let projectNumber = fl_str(SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]).slice(0, 8)
  let mailTextProject = (() => { try { return `<br><br>Проект № ${projectNumber}` } catch (err) { return "" } })();

  let mailAddProgectNumber = tm["Заявка Тема письма Добавить Номер проекта"];
  let mailTextAndProgectNumber = tm["Заявка Текст С Номером Проекта"];
  if (mailAddProgectNumber) {
    let t = `${mailTextAndProgectNumber}`.replace("{ВставитьНомерПроекта}", projectNumber);
    mailSubject = `${mailSubject} ${t}`;
  }

  // let testmailSubject = `${mailSubject} ${mailTextAndProgectNumber}`.replace("{ВставитьНомерПроекта}", projectNumber);
  // mailSubject = tm["Заявка Тема письма"];//тема письма

  // Заявка Таблица №
  // Заявка Таблица Наименование
  // Заявка Таблица Описание
  // Заявка Таблица Ед. изм.
  // Заявка Таблица Кол-во
  let f1f5 = [[
    tm["Заявка Таблица №"],
    tm["Заявка Таблица Наименование"],
    tm["Заявка Таблица Описание"],
    tm["Заявка Таблица Ед. изм."],
    tm["Заявка Таблица Кол-во"],
    tm["Заявка Таблица Файлы"],
  ]];

  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);

  let limitQuotaRemaining = false;
  Logger.log('Скрипт рассылки ' + 'Количество писем к рассылке ' + map.size + '. Лимит рассылки ' + emailQuotaRemaining);

  getContext().getSheetProductGroups().sheet
    .getRange(`${nc(getContext().getSheetProductGroups().col.email + 6)}3:${nc(getContext().getSheetProductGroups().col.email + 7)}3`)
    .setValues([[new Date(), `Дата запуска (раз в 10 минут).\n${'Кол. писем к рассылке: ' + map.size + '.\nЛимит рассылки: ' + emailQuotaRemaining} \nТема Письма: ${mailSubject}`]]);

  if (emailQuotaRemaining < map.size) {
    limitQuotaRemaining = true;
    try {
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'Скрипт рассылки',
        'Количество писем к рассылке ' + map.size + '. Лимит рассылки ' + emailQuotaRemaining,
        ui.ButtonSet.OK
      );
      // return false;
    } catch (err) {

    } finally {
      if (emailQuotaRemaining == 0) {
        // throw new Error("Достигнут лимит отправки писем");
        return limitQuotaRemaining;
      }

      while (map.size - emailQuotaRemaining != 0) {
        map.delete([...map.keys()][0])
      }
    }
  }
  var arrKeys = [];


  // var arrHead = shSet.getRange(1, 15, 1, 5).getValues();
  let arrHead = [[
    "№",
    "Наименование",
    "Описание",
    "Ед. изм.",

    "Кол-во",
    "Файлы",

  ]];


  //ui.alert("arrHeadlen=" + arrHead[0].length);

  for (var [key, value] of map) {
    arrKeys.push(key.toString());
  }

  // Mistervova@mail.ru 25.12.2020 фильтрация значений согласно тз 22.12.2020 //Задача№5 Доработать отправку писем
  // getContext().upDate();


  let logEmailCounter = 0;

  // let toMemArrMails = new Array();

  let toMemMailsMap = new Map();


  for (let i = 0; i < arrKeys.length; i++) {
    // getContext().getSheetProductGroups().sheet
    //   .getRange(`${nc(getContext().getSheetProductGroups().col.email + 4)}3:${nc(getContext().getSheetProductGroups().col.email + 5)}3`)
    //   .setValues([["в очреди на рассылку", arrKeys.length - (i)]]);

    Logger.log(arrKeys[i]);
    //Добавляем заголовки
    /** @type {[[]]} */
    var arrNoHead = map.get(arrKeys[i]);
    let arrNotFiltred = JSON.parse(JSON.stringify(arrNoHead));
    // Logger.log(arrNoHead[0]);

    arrNoHead = mrFilterArrNoHead(arrNoHead);


    arrNotFiltred.unshift(arrHead[0]);
    arrNoHead.unshift(arrHead[0]);
    // Logger.log(arrNoHead[0]);
    // Logger.log(arrNoHead[1]);
    if (f1f5[0][3] == true && f1f5[0][4] == true) {//меняем Ед. изм.	и Кол-во
      var len = arrNoHead[0].length - 1;

      for (let f = 0; f < arrNoHead.length; f++) {


        var a = arrNoHead[f][len - 1];
        arrNoHead[f][len - 1] = arrNoHead[f][len - 2];
        arrNoHead[f][len - 2] = a;

        a = arrNotFiltred[f][len - 1];
        arrNotFiltred[f][len - 1] = arrNotFiltred[f][len - 2];
        arrNotFiltred[f][len - 2] = a;

      }
    }

    // Logger.log(arrNoHead[0]);
    // Logger.log(arrNoHead[1]);

    let filesBlob = (() => {

      if (!arrNoHead[0].includes("Файлы")) { return []; }
      let colФайлы = arrNoHead[0].indexOf("Файлы");
      let ret = arrNoHead
        .map(it => { return it[colФайлы]; })
        .filter((it, i) => {
          // if (i == 0) { return false };
          if (!`${it}`) { return false };
          if (`${it}` == "Файлы") { return false };
          return true;
        })
        .join("\n")
        .split("\n")
        .map(urls => {

          let ph = undefined;

          let id = undefined;
          (() => {
            if (urls.includes("https://drive.google.com/file/d/")) {
              ph = "https://drive.google.com/file/d/";
              id = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
              return
            }




            if (urls.includes("https://docs.google.com/document/d/")) {
              ph = "https://docs.google.com/document/d/";
              id = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
              return
            }

            if (urls.includes("https://docs.google.com/spreadsheets/d/")) {
              ph = "https://docs.google.com/spreadsheets/d/";
              id = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
              return
            }
            Logger.log(`бед фииле`);
            Logger.log(`${urls}`);
          })();




          if (!ph) { return undefined; }

          urls = urls.trim()
          if (!urls) { return undefined; }
          if (urls.length < (ph.length + id.length)) { return undefined; }
          if (urls.indexOf(ph) != 0) { return undefined; }
          return urls.slice(ph.length, ph.length + id.length);
        });

      let st = new Set();
      ret.forEach(id => st.add(id));

      ret = [...st.values()].map(id => {
        if (!id) { return; }
        return (() => { try { return DriveApp.getFileById(id).getBlob(); } catch (err) { return undefined } })();
      }).filter(id => id);

      // ret.forEach((f, i, arr) => { Logger.log(`  ${i}/${arr.length}  ${f.getName()}`) });
      return ret;
    })();

    arrNoHead = arrNoHead.map(r => { return r.slice(0, -1) })
    // Logger.log(arrNoHead[0]);
    // Logger.log(arrNoHead[1]);
    var mailHtml = dataToHtmlTable(arrNoHead);


    let mail = {
      to: arrKeys[i],
      replyTo: mailReplyTo,
      subject: mailSubject,
      htmlBody: mailText1 + "<br>" + mailHtml + mailText2 + mailTextProject,
      // attachments: [pdfFileBlob]
    };

    if (mailAttachmentsId.toString().length > 0) {
      var pdfFileBlob = DriveApp.getFileById(mailAttachmentsId).getBlob();
      mail["attachments"] = filesBlob.concat(pdfFileBlob);

    } else {

    }

    let mailColor = undefined; // цвет при отказе если undefined значеть отправили письмо 
    let emailAddress = arrKeys[i];
    let check = checkCounteragenth(emailAddress); // вернет emailAddress или цвет
    // console.log("emailAddress= " + emailAddress + " check=" + check);
    // getContext().getCounteragent(emailAddress).logToConsole();



    let isApprovedEmail = (check == emailAddress)  // одобрена отправка писем 

    if (isApprovedEmail) { } else { // не отправлять письма, check -это  цвет причины отказа 

      mailColor = check;
    }
    // setMailBackground(arrKeys[i], mailColor);
    if (isApprovedEmail) {    // вернулся emailAddress значить шлем письмо 
      try {
        toMemMailsMap.set(emailAddress, {
          url,
          mail,
          ДатаОтправки: getContext().timeConstruct,
          from: Session.getEffectiveUser().getEmail(),
          to: emailAddress,
          key: JSON.stringify({
            from: Session.getEffectiveUser().getEmail(),
            to: emailAddress,
            ДатаОтправки: getContext().timeConstruct,
            url,
          }),
          Номенклатуры: arrNoHead.map((rr, i) => {
            try {
              if (i == 0) { return; }
              let r = arrNotFiltred[i];
              return {
                НомерПроекта,
                to: emailAddress,
                url,
                Номенклатура: r[1],
                ДатаОтправки: getContext().timeConstruct,
                // mail,
                from: Session.getEffectiveUser().getEmail(),
                key: `${JSON.stringify({
                  НомерПроекта,
                  to: emailAddress,
                  url,
                  Номенклатура: r[1],
                  ДатаОтправки: getContext().timeConstruct,
                })}`,
                r,
              };
            } catch (err) { mrErrToString(err); }

          }),
        });

      } catch (err) { mrErrToString(err); }
      // getContext().getSheetНастройки_писем().addToMemSentMails(toMemMailsMap, url);
      getContext().getSheetНастройки_писем().sendMail(mail);

      try {

        let colEmail = nc(getContext().getSheetProductGroups().col.email + 1);
        let email_QuotaRemaining = [[Session.getEffectiveUser().getEmail() + "\nЛимит рассылки"], [MailApp.getRemainingDailyQuota()]];
        getContext().getSheetProductGroups().sheet.getRange(`${colEmail}1:${colEmail}2`).setValues(email_QuotaRemaining);
      } catch (err) {
        mrErrToString(err);
      }

      getContext().getSheetНастройки_писем().addToMemSentMails(toMemMailsMap, url);
      // SpreadsheetApp.getUi().alert(` отправить ${arrKeys[i]} ` );
      // logEmailCounter++;
    }

    toMemMailsMap.clear();
    // toMemMailsMap.clear();

    // showMail(shSet, mail, isApprovedEmail, mailColor);
    setMailBackground(arrKeys[i], mailColor);

    // getContext().getSheetProductGroups().sheet
    //   .getRange(`${nc(getContext().getSheetProductGroups().col.email + 4)}3:${nc(getContext().getSheetProductGroups().col.email + 5)}3`)
    //   .setValues([["в очреди на рассылку", arrKeys.length - (i + 1)]]);

  }

  // getContext().getSheetНастройки_писем().addToMemSentMails(toMemMailsMap, url);
  //логируем кол-во писем
  // logEmail(logEmailCounter);
  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);

  // if (limitQuotaRemaining) throw new Error("Достигнут лимит отправки писем");

  return limitQuotaRemaining;

}

function sendEmailРассылкаЗапросовМедлено(map = new Map()) {


  let url = (() => { return getContext().getSpreadSheet().getUrl(); })();

  let НомерПроекта = getContext().getНомерПроекта();
  var shSet = getContext().getSheetНастройки_писем().sheet

  let tm = getContext().getSheetНастройки_писем().getTemplateForSendMails();
  if (!tm) {
    throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю. или нет доступа к файлу`;
    throw `Для ${getContext().getUserEmail()} нет шаблона`;
  }
  if (tm["Заявка"] !== true) {
    throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`
  }
  // Logger.log(tm["Заявка"]);


  // Заявка
  // Заявка Тема письма
  // Заявка Текст до таблицы
  // Заявка Текст после таблицы
  // Заявка Вложение
  // Заявка Адрес ответа

  let mailSubject = tm["Заявка Тема письма"];//тема письма
  let mailText1 = tm["Заявка Текст до таблицы"];//текст До
  let mailText2 = tm["Заявка Текст после таблицы"];//Текст После
  let mailAttachmentsId = tm["Заявка Вложение"];//Прикрепленный файл
  let mailReplyTo = tm["Заявка Адрес ответа"];//Куда отвечать
  let mailTextProject = (() => { try { return `<br><br>Проект № ${SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]}` } catch (err) { return "" } })();

  // Заявка Таблица №
  // Заявка Таблица Наименование
  // Заявка Таблица Описание
  // Заявка Таблица Ед. изм.
  // Заявка Таблица Кол-во
  let f1f5 = [[
    tm["Заявка Таблица №"],
    tm["Заявка Таблица Наименование"],
    tm["Заявка Таблица Описание"],
    tm["Заявка Таблица Ед. изм."],
    tm["Заявка Таблица Кол-во"],
    tm["Заявка Таблица Файлы"],
  ]];

  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);

  let limitQuotaRemaining = false;
  Logger.log('Скрипт рассылки ' + 'Количество писем к рассылке ' + map.size + '. Лимит рассылки ' + emailQuotaRemaining,)
  if (emailQuotaRemaining < map.size) {
    limitQuotaRemaining = true;
    try {
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'Скрипт рассылки',
        'Количество писем к рассылке ' + map.size + '. Лимит рассылки ' + emailQuotaRemaining,
        ui.ButtonSet.OK
      );
      // return false;
    } catch (err) {

    } finally {
      if (emailQuotaRemaining == 0) { return false; }
      while (map.size - emailQuotaRemaining != 0) {
        map.delete([...map.keys()][0])
      }
    }
  }
  var arrKeys = [];


  // var arrHead = shSet.getRange(1, 15, 1, 5).getValues();
  let arrHead = [[
    "№",
    "Наименование",
    "Описание",
    "Ед. изм.",

    "Кол-во",
    "Файлы",

  ]];


  //ui.alert("arrHeadlen=" + arrHead[0].length);

  for (var [key, value] of map) {
    arrKeys.push(key.toString());
  }

  // Mistervova@mail.ru 25.12.2020 фильтрация значений согласно тз 22.12.2020 //Задача№5 Доработать отправку писем
  getContext().upDate();


  let logEmailCounter = 0;

  // let toMemArrMails = new Array();

  let toMemMailsMap = new Map();


  for (let i = 0; i < arrKeys.length; i++) {
    Logger.log(arrKeys[i]);
    //Добавляем заголовки
    /** @type {[[]]} */
    var arrNoHead = map.get(arrKeys[i]);
    let arrNotFiltred = JSON.parse(JSON.stringify(arrNoHead));
    // Logger.log(arrNoHead[0]);

    arrNoHead = mrFilterArrNoHead(arrNoHead);


    arrNotFiltred.unshift(arrHead[0]);
    arrNoHead.unshift(arrHead[0]);
    // Logger.log(arrNoHead[0]);
    // Logger.log(arrNoHead[1]);
    if (f1f5[0][3] == true && f1f5[0][4] == true) {//меняем Ед. изм.	и Кол-во
      var len = arrNoHead[0].length - 1;

      for (let f = 0; f < arrNoHead.length; f++) {


        var a = arrNoHead[f][len - 1];
        arrNoHead[f][len - 1] = arrNoHead[f][len - 2];
        arrNoHead[f][len - 2] = a;

        a = arrNotFiltred[f][len - 1];
        arrNotFiltred[f][len - 1] = arrNotFiltred[f][len - 2];
        arrNotFiltred[f][len - 2] = a;

      }
    }

    // Logger.log(arrNoHead[0]);
    // Logger.log(arrNoHead[1]);

    let filesBlob = (() => {

      if (!arrNoHead[0].includes("Файлы")) { return []; }
      let colФайлы = arrNoHead[0].indexOf("Файлы");
      let ret = arrNoHead
        .map(it => { return it[colФайлы]; })
        .filter((it, i) => {
          // if (i == 0) { return false };
          if (!`${it}`) { return false };
          return true;
        })
        .join("\n")
        .split("\n")
        .map(urls => {
          let ph = undefined;

          let id = undefined;
          (() => {
            if (urls.includes("https://drive.google.com/file/d/")) {
              ph = "https://drive.google.com/file/d/";
              id = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
              return
            }



            if (urls.includes("https://docs.google.com/document/d/")) {
              ph = "https://docs.google.com/document/d/";
              id = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
              return
            }

            if (urls.includes("https://docs.google.com/spreadsheets/d/")) {
              ph = "https://docs.google.com/spreadsheets/d/";
              id = "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ"
              return
            }
            Logger.log(`бед фииле`);
            Logger.log(`${urls}`);
          })();

          // "https://drive.google.com/file/d/";
          // "ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ";



          if (!ph) { return undefined; }

          urls = urls.trim()
          if (!urls) { return undefined; }
          if (urls.length < (ph.length + id.length)) { return undefined; }
          if (urls.indexOf(ph) != 0) { return undefined; }
          return urls.slice(ph.length, ph.length + id.length);
        });

      let st = new Set();
      ret.forEach(id => st.add(id));

      ret = [...st.values()].map(id => {
        if (!id) { return; }
        return (() => { try { return DriveApp.getFileById(id).getBlob(); } catch (err) { return undefined } })();
      }).filter(id => id);

      // ret.forEach((f, i, arr) => { Logger.log(`  ${i}/${arr.length}  ${f.getName()}`) });
      return ret;
    })();

    arrNoHead = arrNoHead.map(r => { return r.slice(0, -1) })
    // Logger.log(arrNoHead[0]);
    // Logger.log(arrNoHead[1]);
    var mailHtml = dataToHtmlTable(arrNoHead);


    let mail = {
      to: arrKeys[i],
      replyTo: mailReplyTo,
      subject: mailSubject,
      htmlBody: mailText1 + "<br>" + mailHtml + mailText2 + mailTextProject,
      // attachments: [pdfFileBlob]
    };

    if (mailAttachmentsId.toString().length > 0) {
      var pdfFileBlob = DriveApp.getFileById(mailAttachmentsId).getBlob();
      mail["attachments"] = filesBlob.concat(pdfFileBlob);

    } else {

    }

    let mailColor = undefined; // цвет при отказе если undefined значеть отправили письмо 
    let emailAddress = arrKeys[i];
    let check = checkCounteragenth(emailAddress); // вернет emailAddress или цвет
    // console.log("emailAddress= " + emailAddress + " check=" + check);
    // getContext().getCounteragent(emailAddress).logToConsole();



    let isApprovedEmail = (check == emailAddress)  // одобрена отправка писем 
    if (isApprovedEmail) {    // вернулся emailAddress значить шлем письмо 

      getContext().getSheetНастройки_писем().sendMail(mail);
      try {
        toMemMailsMap.set(emailAddress, {
          url,
          mail,
          ДатаОтправки: getContext().timeConstruct,
          from: Session.getEffectiveUser().getEmail(),
          to: emailAddress,
          key: JSON.stringify({
            from: Session.getEffectiveUser().getEmail(),
            to: emailAddress,
            ДатаОтправки: getContext().timeConstruct,
            url,
          }),
          Номенклатуры: arrNoHead.map((rr, i) => {
            try {
              if (i == 0) { return; }
              let r = arrNotFiltred[i];
              return {
                НомерПроекта,
                to: emailAddress,
                url,
                Номенклатура: r[1],
                ДатаОтправки: getContext().timeConstruct,
                // mail,
                from: Session.getEffectiveUser().getEmail(),
                key: `${JSON.stringify({
                  НомерПроекта,
                  to: emailAddress,
                  url,
                  Номенклатура: r[1],
                  ДатаОтправки: getContext().timeConstruct,
                })}`,
                r,
              };
            } catch (err) { mrErrToString(err); }

          }),
        });

      } catch (err) { mrErrToString(err); }

      // SpreadsheetApp.getUi().alert(` отправить ${arrKeys[i]} ` );
      logEmailCounter++;
    } else { // не отправлять письма, check -это  цвет причины отказа 
      mailColor = check;
    }
    getContext().getSheetНастройки_писем().addToMemSentMails(toMemMailsMap, url); toMemMailsMap.clear();
    // toMemMailsMap.clear();

    // showMail(shSet, mail, isApprovedEmail, mailColor);
    setMailBackground(arrKeys[i], mailColor);

  }

  // getContext().getSheetНастройки_писем().addToMemSentMails(toMemMailsMap, url);
  //логируем кол-во писем
  logEmail(logEmailCounter);
  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);

  if (limitQuotaRemaining) throw "Достигнут лимит отправки писем";
}




function takeFromRow(arrIn, f1f5) {
  //Logger.log("arrIn=" + arrIn);
  //Logger.log("f1f5=" + f1f5 + " f1f5[0].length=" + f1f5[0].length);
  var arrTemp = [];
  for (let j = 0; j < f1f5[0].length; j++) {
    if (f1f5[0][j] == true) {
      arrTemp.push(arrIn[j]);
    }
  }
  //Logger.log("arrTemp=" + arrTemp);
  return arrTemp;
}


function copyArr(arrIn) {
  //Logger.log("arrIn=" + arrIn)
  //Logger.log("arrIn.length=" + arrIn.length)
  //Logger.log("arrIn[0].length=" + arrIn[0].length)  
  var arrOut = [];
  for (let i = 0; i < arrIn.length; i++) {
    var arrTemp = [];
    for (let j = 0; j < arrIn[i].length; j++) {
      arrTemp.push(arrIn[i][j]);
    }
    //Logger.log(i + " arrTemp=" + arrTemp)   
    arrOut.push(arrTemp);
  }

  //Logger.log("arrOut=" + arrOut)
  return arrOut;
}









// MisterVova@mail.ru 23.12.2020 добавил файл  

// MisterVova@mail.ru 23.12.2020 добавил функцию 
function onEditProductGroups(edit) {

  // Получаем диапазон ячеек, в которых произошли изменения
  let range = edit.range;

  // Лист, на котором производились изменения
  let sheet = range.getSheet();

  // Проверяем, нужный ли это нам лист
  Logger.log("onEditProductGroups Начало");
  // if (sheet.getName() != 'Группы товаров') {
  if (sheet.getName() != getSettings().sheetName_Группы_товаров) {
    return false;
  }

  let colB = getContext().getSheetProductGroups().columnProduct;
  let colI = getContext().getSheetProductGroups().columnEmail;
  let colG = getContext().getSheetProductGroups().columnProductG;

  if (
    ((range.columnStart <= colB) && (colB <= range.columnEnd))
    ||
    ((range.columnStart <= colI) && (colG <= range.columnEnd))
  ) {
    let cProductGroups = getContext().getSheetProductGroups();
    cProductGroups.updateProduct(range.rowStart, range.rowEnd);

    Logger.log("onEditProductGroups Конец cProductGroups.updateProduct");
  }

  if ((colI <= range.columnEnd)) {
    // Logger.log(` Группы товаров (nr("J") <= range.columnEnd)`);
    Logger.log(` ${getSettings().sheetName_Группы_товаров} (nr("K") <= range.columnEnd)`);
    let cProductGroups = getContext().getSheetProductGroups();
    cProductGroups.updateGroups(edit);
    Logger.log("onEditProductGroups Конец cProductGroups.updateGroups");
  }
  Logger.log("onEditProductGroups Конец ");
}









function ProdBackground(e) {
  // старая версия // не
  // var ui = SpreadsheetApp.getUi();

  // throw  new Error(" старая версия ");

  // try {
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var val = e.value;
  // var colNameGroup = 12;//столбик где пишется название группы на первом листе
  var colNameGroup = getContext().getSheetList1().columnGroup;//столбик где пишется название группы на первом листе
  Logger.log("ProdBackground edit= " + JSON.stringify(e));
  Logger.log("ProdBackground sheetName=" + sheetName);
  Logger.log("ProdBackground colNameGroup=" + colNameGroup);
  // let colI = getContext().getSheetProductGroups().columnEmail;
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)
  // if (sheetName == "Лист1") {
  if (sheetName == getSettings().sheetName_Лист_1) {
    // ui.alert(JSON.stringify(e));

    if (col > colNameGroup) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      // var sh = ss.getSheetByName("Группы товаров");
      // var sh = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
      var sh = getContext().getSheetProductGroups().sheet
      // var shSet = ss.getSheetByName("Настройки писем");
      // var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);

      // var colorBrown = shSet.getRange("K2").getBackground();
      // var colorProd = shSet.getRange("K3").getBackground();
      // var colorGroup = shSet.getRange("K4").getBackground();


      var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
      var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
      var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);


      var nameGroup = sheet.getRange(row, colNameGroup).getValue();
      var nameEmail = sheet.getRange(1, col).getValue();
      if (!val) { return; } //  если новое значение пусто то ничего не делаем 
      if (!nameEmail) { return; } //  если в строке 1  нет адреса  то ничего не делаем

      //ui.alert("nameEmail = " + nameEmail + "\n nameGroup = " + nameGroup)

      var arr = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();

      for (let i = 0; i < arr.length; i++) {  //  i+1=это строка номер строки в таблице 
        // if (i == 0) { continue; }
        // Logger.log(`"ProdBackground continue= i=" ${i}  row=${i + 1}  SecondRow= ${getContext().getProductGroup(i + 1).getSecondRow()} `);
        // Logger.log(`"ProdBackground continue= i=" ${i}  row=${i + 1}   FirstRow= ${getContext().getProductGroup(i).getFirstRow()} SecondRow= ${getContext().getProductGroup(i + 1).getSecondRow()} `);

        let secondRow = getContext().getProductGroup(i + 1).getSecondRow();// вторая строка для этой группы 
        if (!secondRow) {  // не нашли группу  для этой строки и соответственно нет и второй строки
          if (i == 0) { continue; }
          let firstRow = getContext().getProductGroup(i).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
          if (firstRow == i) {
            secondRow = i + 1;
          }

        }

        // Logger.log(` if =${secondRow == (i + 1)}  `);


        if (secondRow != (i + 1)) { continue; }  // не вторая строка в группе 
        // continue;
        for (let j = sec - 1; j < arr[0].length; j++) {

          // Logger.log(`${arr[i][j]}`)
          if (fl_str(arr[i][j]) == fl_str(nameEmail)) {
            // let bg = sh.getRange(i + 1, j + 1).getBackground();
            // if (bg == colorBrown || bg == colorProd) {//если закраска, что отправил
            // Logger.log(`${nameGroup} | ${arr[i - 1][getContext().getSheetProductGroups().columnNameGr - 1]} | ${arr[i][j]}`);


            if (true) {//  для всех  в это й строчке 
              if (arr[i - 1][getContext().getSheetProductGroups().columnNameGr - 1] == nameGroup) {//если совпала группа
                sh.getRange(i + 1, j + 1).setBackground(colorGroup);
              } else { //если совпала группа
                let bg = sh.getRange(i + 1, j + 1).getBackground();
                if (bg == colorGroup) { continue; } //если совпала группа а уже закрашет в цвет группы colorGroup
                sh.getRange(i + 1, j + 1).setBackground(colorProd);
              }
            }
          }
        }
      }

    }
  }
  // } catch (err) {
  //   ui.alert(mrErrToString(err));
  // }
}






function logEmail(count) {
  var ss = SpreadsheetApp.getActive();
  // var shSet = ss.getSheetByName("Настройки писем");
  var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);
  var row = shSet.getLastRow() + 1;
  shSet.getRange(row, 2).setValue(new Date());
  shSet.getRange(row, 3).setValue(count);
}

function delBackgroundMailGreen() {
  if (!delBackgroundMailAlert()) { return; }
  var ss = SpreadsheetApp.getActive();
  // var shSet = ss.getSheetByName("Настройки писем");
  var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);
  // var colorProd = shSet.getRange("K3").getBackground();
  // var colorGroup = shSet.getRange("K4").getBackground();

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);



  delBackgroundMail(colorProd);
  delBackgroundMail(colorGroup);
}


function delBackgroundMailBrown() {

  if (!delBackgroundMailAlert()) { return; }
  var ss = SpreadsheetApp.getActive();
  // var shSet = ss.getSheetByName("Настройки писем");
  var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);

  delBackgroundMail(colorBrown, true);

}

function delBackgroundMailAlert() {

  var ui = SpreadsheetApp.getUi();
  return result = ui.alert(
    'Скрипт рассылки',
    'Вы уверены, что хотите удалить заливку?',
    ui.ButtonSet.YES_NO
  );
}



function delBackgroundMail(color, delNotes = false) {
  // var ss = SpreadsheetApp.getActive();

  // var sh = ss.getSheetByName("Группы товаров");
  // var sh = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var sh = getContext().getSheetProductGroups().sheet
  var lRow = sh.getLastRow();
  var lCol = sh.getLastColumn();
  // var sec = 10;//столбик откуда начинают писать email (на листе)
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)
  for (let i = 1; i < lRow + 1; i++) {
    for (let j = sec; j < lCol + 1; j++) {
      if (sh.getRange(i, j).getBackground() == color) {
        sh.getRange(i, j).setBackground(null);
        if (delNotes) { sh.getRange(i, j).setNote(""); }
      }
    }
  }

}


// MisterVova@mail.ru 23.12.2020 согласно тз от 22.12.2020 //Задача№6 //  добавелен параметр otherColor=undefined
// otherColor - другой цвет - otherColor=undefined параметер по умолчанию 
//function setMailBackground(email){

function setMailBackground(email, otherColor = undefined) {

  var ss = SpreadsheetApp.getActive();
  // var shSet = ss.getSheetByName("Настройки писем");
  // var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);
  // var shSet = getContext().getSheetНастройки_писем().sheet;
  // var colorBrown = shSet.getRange("K2").getBackground();
  // var colorProd = shSet.getRange("K3").getBackground();
  // var colorGroup = shSet.getRange("K4").getBackground();

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);

  // var sh = ss.getSheetByName("Группы товаров");
  // var sh = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var sh = getContext().getSheetProductGroups().sheet;



  var lRow = sh.getLastRow();
  var lCol = sh.getLastColumn();
  var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
  // var sec = 10;//столбик откуда начинают писать email (на листе)
  var sec = getContext().getSheetProductGroups().columnEmail;//столбик откуда начинают писать email (на листе)
  // MisterVova@mail.ru 23.12.2020 согласно тз от 22.12.2020 Задача№6 
  // если otherColor определен (задан) то делаем подмену цвета 
  let colorBg = colorBrown;
  // if (otherColor) { colorBg = otherColor; }  // цвет токаза 
  if (otherColor) { colorBg = "#ffffff"; }  //цвет отказа белый 

  let fromEmail = Session.getEffectiveUser().getEmail();
  for (let i = arrSh.length - 1; i > 1; i--) {
    for (let j = sec - 1; j < arrSh[0].length; j++) {
      if (email == arrSh[i][j]) {
        if (arrSh[i - 2][0] == "" && arrSh[i - 1][0] != "") {
          let bg = sh.getRange(i + 1, j + 1).getBackground();
          if (bg != colorProd && bg != colorGroup) {
            if (colorBg == colorBrown) {
              sh.getRange(i + 1, j + 1).setNote(`${fromEmail}\n>>>\n${email} ||| \n${Utilities.formatDate(new Date(), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}|||\nда`)
            } else {
              sh.getRange(i + 1, j + 1).setNote(`${email}||| \n${getContext().getCounteragent(email).getStatus().toString()}||| \nнет`);
            }  // ставим удаляем заметку  
            sh.getRange(i + 1, j + 1).setBackground(colorBg);
          }
        }
      }
    }
  }
}


function addCSS(inText) {
  var style = '<style type="text/css"> TABLE { border-collapse: collapse; } TD, TH { padding: 3px; border: 1px solid black;} TH {    background: #b0e0e6;    } </style>'
  var css = style + '\n' + inText;
  return css;
}


function dataToHtmlTable(data) {
  return JSON.stringify(data, null, "  ")
    .replace(/^\[/g, '<table border = "1" rules="all" cellpadding="2">')
    .replace(/\]$/g, "</table>")
    .replace(/^\s\s\[$/mg, '<tr>')
    .replace(/^\s\s\],{0,1}$/mg, "</tr>")
    .replace(/^\s{4}"{0,1}(.*?)"{0,1},{0,1}$/mg, "<td>$1</td>");
}


function createHTML(arr) {
  Logger.log("createHTML");
  Logger.log(arr);
  var html = "<table><tbody>";
  for (let i = 0; i < arr.length; i++) {
    html = html + "<tr>"
    for (let j = 0; j < arr[0].length; j++) {
      html = html + "<br>" + arr[i][j] + "</br>";
    }
    html = html + "</tr>"
  }
  html = html + "</tbody></table>"
  Logger.log(html);
  return html;
}



function checkIfIsSeendMeilInNete(strNote) {
  if (!strNote) { return false; }
  // strNote = fl_str(strNote);

  let arrNote = strNote.split("|||");

  if (arrNote.length == 3) { if (fl_str(arrNote[2]) == fl_str("да")) { return true; } }
  if (arrNote.length == 3) { if (fl_str(arrNote[1]) == fl_str("НЕЛЬЗЯ РАБОТАТЬ")) { return true; } }
  return false;
}


function updateBackground() {  // старая версия 
  // return
  throw new Error(" старая версия ");
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  // var shLit = ss.getSheetByName("Лист1");
  var shLit = ss.getSheetByName(getSettings().sheetName_Лист_1);
  var arrLit = shLit.getDataRange().getValues();
  // var shTov = ss.getSheetByName("Группы товаров");
  var shTov = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var arrTov = shTov.getDataRange().getValues();
  var arrTovBack = shTov.getDataRange().getBackgrounds();
  var arrTovNote = shTov.getDataRange().getNotes();
  // var shSet = ss.getSheetByName("Настройки писем");
  var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);
  // var colorBrown = shSet.getRange("K2").getBackground();
  // var colorProd = shSet.getRange("K3").getBackground();
  // var colorGroup = shSet.getRange("K4").getBackground();


  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);

  var colGroupLit = 11;//[11]L
  var colGroupTov = 5;//F столбик групп на листе "Группы товаров"  // getSettings().sheetName_Группы_товаров

  let arrSecondRowS = new Array();
  let colorDef = getContext().getDefaultColor() // наверно белый ж



  // 0)Сбрасываем все в корчневый для отправленных и в белых для не отправленных
  for (let iTov = 1; iTov < arrTov.length; iTov++) {
    // проверяем это не вторая строка тогда пропускаем 
    let secondRow = getContext().getProductGroup(iTov + 1).getSecondRow();// вторая строка для этой группы 
    if (!secondRow) {  // не нашли группу  для этой строки и соответственно нет и второй строки
      if (iTov == 0) { continue; } // не вторая строка в группе 
      let firstRow = getContext().getProductGroup(iTov).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
      if (firstRow == iTov) {
        secondRow = iTov + 1;
      }
    }
    if (secondRow != (iTov + 1)) { continue; }  // не вторая строка в группе 

    arrSecondRowS.push(iTov);
    for (let jTov = colGroupTov + 4; jTov < arrTov[0].length; jTov++) {
      if (!arrTov[iTov][jTov]) { arrTovBack[iTov][jTov] = colorDef; continue; }  // если пусто то белый 
      if (!checkIfIsSeendMeilInNete(arrTovNote[iTov][jTov])) { arrTovBack[iTov][jTov] = colorDef; continue; }  // не  пометки об отправки письма  
      arrTovBack[iTov][jTov] = colorBrown;
    }
  }


  for (let jL = colGroupLit + 1; jL < arrLit[0].length; jL++) {

    var flag = false;
    for (let iL = 0; iL < arrLit.length; iL++) {

      if (arrLit[iL][jL] > 0) {//нашли ячейку с введенной ценой
        var mailLit = arrLit[0][jL];
        var groupLit = arrLit[iL][colGroupLit];
        if (mailLit.toString().indexOf('@') > 0 && groupLit.toString().length > 0) {
          // Logger.log("arrLit[iL][jL]=" + arrLit[iL][jL] + " mailLit =" + mailLit + " groupLit =" + groupLit);
          //1)При сопадении мыла - краим в синий
          for (let iTov = 1; iTov < arrTov.length; iTov++) {

            if (!arrSecondRowS.includes(iTov)) { continue; }


            for (let jTov = colGroupTov + 4; jTov < arrTov[0].length; jTov++) {
              // if (arrTov[iTov][jTov] == mailLit && arrTovBack[iTov][jTov] == colorBrown && flag == false) {
              //   arrTovBack[iTov][jTov] = colorProd;
              // }

              if (arrTov[iTov][jTov] == mailLit && flag == false) {
                arrTovBack[iTov][jTov] = colorProd;
              }


            }
          }
          flag = true;//уже закасили все в синий
          //2)Те, что сними при наличии группы - красим в зеленый
          for (let iTov = 1; iTov < arrTov.length; iTov++) {
            for (let jTov = colGroupTov + 4; jTov < arrTov[0].length; jTov++) {
              // Logger.log(arrTov[iTov - 1][colGroupTov]);
              // Logger.log("groupLit=" + groupLit);
              if (arrTov[iTov - 1][colGroupTov] == groupLit && arrTovBack[iTov][jTov] == colorProd) {
                arrTovBack[iTov][jTov] = colorGroup;
              }
            }
          }

        }
      }
    }
  }

  shTov.getDataRange().setBackgrounds(arrTovBack);


  // теперь первые строчки группы



  let cProductGroups = getContext().getSheetProductGroups();
  // cProductGroups.updateSquare(rowStart, colStart, rowEnd, colEnd);



  let colStart = nr("J");
  let colEnd = shTov.getLastColumn();
  if (colEnd < nr("J")) { colEnd = nr("J"); }
  let rowStart = 3;
  let rowEnd = shTov.getLastRow();

  for (let row = rowStart; row <= rowEnd; row++) {

    let firstRow = getContext().getProductGroup(row).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
    if (firstRow != row) { continue; }  // не первая строка в группе 

    cProductGroups.updateSquare(row, colStart, row, colEnd);
  }






}


function updateBackgroundV2() {

  Logger.log("updateBackgroundV2 vvvvvvvvvvvvvvvvvvvvvvv  ");

  var ss = SpreadsheetApp.getActive();

  var shLit = ss.getSheetByName(getSettings().sheetName_Лист_1);
  var arrLit = shLit.getDataRange().getValues();

  var shTov = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
  var arrTov = shTov.getDataRange().getValues();
  var arrTovBack = shTov.getDataRange().getBackgrounds();
  var arrTovNote = shTov.getDataRange().getNotes();

  // var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);

  var colorBrown = getContext().getColor(getSettings().colorNames.ЗаливкаПриОтправке);
  var colorProd = getContext().getColor(getSettings().colorNames.ЗаливкаПоставщика);
  var colorGroup = getContext().getColor(getSettings().colorNames.ЗаливкаГруппы);

  let rowLastShTov = shTov.getLastRow();
  // getContext().getSheetProductGroups().col.Группа

  // var colGroupLit = 11;//[11]L
  // var colGroupTov = 5;//F столбик групп на листе "Группы товаров"  // getSettings().sheetName_Группы_товаров

  var colGroupLit = getContext().getSheetList1().columnGroup - 1;
  var colGroupTov = getContext().getSheetProductGroups().col.Группа - 1;


  let arrSecondRowS = new Array();
  let colorDef = getContext().getDefaultColor() // наверно белый ж


  Logger.log("Сбрасываем все в корчневый для отправленных и в белых для не отправленных ----------------------------");
  // 0)Сбрасываем все в корчневый для отправленных и в белых для не отправленных
  for (let iTov = 4; iTov < arrTov.length; iTov++) {
    // Logger.log(JSON.stringify({ iTov, rowLastShTov }, undefined, 2));
    if (rowLastShTov + 1 < iTov) { break; }
    // проверяем это не вторая строка тогда пропускаем 
    let secondRow = getContext().getProductGroup(iTov + 1).getSecondRow();// вторая строка для этой группы 
    if (!secondRow) {  // не нашли группу  для этой строки и соответственно нет и второй строки
      if (iTov == 0) { continue; } // не вторая строка в группе 
      let firstRow = getContext().getProductGroup(iTov).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
      if (firstRow == iTov) {
        secondRow = iTov + 1;
      }
    }
    if (secondRow != (iTov + 1)) { continue; }  // не вторая строка в группе 

    arrSecondRowS.push(iTov);
    for (let jTov = getContext().getSheetProductGroups().col.email - 1; jTov < arrTov[0].length; jTov++) {
      // Logger.log(JSON.stringify({ iTov, jTov, color: arrTovBack[iTov][jTov], da: checkIfIsSeendMeilInNete(arrTovNote[iTov][jTov]), email: arrTov[iTov][jTov], note: arrTovNote[iTov][jTov] }, undefined, 2));
      if (!arrTov[iTov][jTov]) { arrTovBack[iTov][jTov] = colorDef; continue; }  // если пусто то белый 
      if (!checkIfIsSeendMeilInNete(arrTovNote[iTov][jTov])) { arrTovBack[iTov][jTov] = colorDef; continue; }  // не  пометки об отправки письма  
      arrTovBack[iTov][jTov] = colorBrown;

    }
  }
  Logger.log("собираем ответы и листа 1-1 ----------------------------");
  // return
  for (let jL = colGroupLit + 1; jL < arrLit[0].length; jL++) {

    var flag = false;
    for (let iL = 0; iL < arrLit.length; iL++) {

      if (arrLit[iL][jL] > 0) {//нашли ячейку с введенной ценой
        var mailLit = arrLit[0][jL];
        mailLit = `${mailLit}`.split(" ")[0];
        var groupLit = arrLit[iL][colGroupLit];



        if (mailLit.toString().indexOf('@') > 0 && groupLit.toString().length > 0) {
          // Logger.log(JSON.stringify({ p: arrLit[iL][jL], mailLit, groupLit }));
          // Logger.log("arrLit[iL][jL]=" + arrLit[iL][jL] + " mailLit =" + mailLit + " groupLit =" + groupLit);
          //1)При сопадении мыла - краим в синий
          for (let iTov = 1; iTov < arrTov.length; iTov++) {

            if (!arrSecondRowS.includes(iTov)) { continue; }


            for (let jTov = getContext().getSheetProductGroups().col.email - 1; jTov < arrTov[0].length; jTov++) {
              if (arrTov[iTov][jTov] == mailLit && flag == false) {
                arrTovBack[iTov][jTov] = colorProd;
              }
            }
          }
          flag = true;//уже закасили все в синий
          //2)Те, что сними при наличии группы - красим в зеленый
          for (let iTov = 1; iTov < arrTov.length; iTov++) {
            for (let jTov = getContext().getSheetProductGroups().col.email - 1; jTov < arrTov[0].length; jTov++) {
              if (arrTov[iTov - 1][colGroupTov] == groupLit && arrTovBack[iTov][jTov] == colorProd) {
                arrTovBack[iTov][jTov] = colorGroup;
              }
            }
          }

        }
      }
    }
  }

  shTov.getDataRange().setBackgrounds(arrTovBack);

  Logger.log("теперь первые строчки группы ----------------------------");
  // теперь первые строчки группы
  // return;
  let cProductGroups = getContext().getSheetProductGroups();

  let colStart = getContext().getSheetProductGroups().col.email;
  let colEnd = shTov.getLastColumn();
  if (colEnd < colStart) { colEnd = colStart; }
  let rowStart = 3;
  let rowEnd = shTov.getLastRow();

  for (let row = rowStart; row <= rowEnd; row++) {

    let firstRow = getContext().getProductGroup(row).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
    if (firstRow != row) { continue; }  // не первая строка в группе 

    cProductGroups.updateSquare(row, colStart, row, colEnd);
  }
  Logger.log("updateBackgroundV2 ^^^^^^^^^^^^^^^^^^^ ");
}



function mrFilterArrNoHead(arrNoHead) {
  // Mistervova@mail.ru фильтрация значений согласно тз от 22.12.2020 //Задача№5 Доработать отправку писем
  // поиск фильтруемых символов и замена их в текстах Наименования и Описание
  let ss = SpreadsheetApp.getActive();
  // let sheetSetings = ss.getSheetByName("Настройки писем");
  let sheetSetings = ss.getSheetByName(getSettings().sheetName_Настройки_писем);

  let setings = sheetSetings.getRange(1, nr("H"), sheetSetings.getLastRow(), 2).getValues();

  let repStr = new Array();
  let newStr = new Array();


  // console.log("setings =" + setings );
  for (let i = setings.length - 1; i >= 0; i--) {
    // console.log( "i= "+ i+ " лист= " + setings [i][0] + " колонка =" + setings [i][1] );
    if (!setings[i][0]) { continue; }
    repStr.push(setings[i][0]);
    newStr.push(setings[i][1]);
  }

  for (let f = 0; f < arrNoHead.length; f++) {
    for (let i = 0; i < repStr.length; i++) {
      // console.log('mrReplace :' + arrNoHead[f][1] + " = "  +mrReplace(arrNoHead[f][1],repStr[i],newStr[i]));
      arrNoHead[f][1] = mrReplace(arrNoHead[f][1], repStr[i], newStr[i]);
      arrNoHead[f][2] = mrReplace(arrNoHead[f][2], repStr[i], newStr[i]);
    }

    // arrNoHead[f][1] = mrReplace(arrNoHead[f][1],"\n","<br>");
    // arrNoHead[f][2] = mrReplace(arrNoHead[f][2],"\n","<br>");
    // arrNoHead[f][1] = mrReplace(arrNoHead[f][1],'\"','&quot;');
    // arrNoHead[f][2] = mrReplace(arrNoHead[f][2],'\"','&quot;');
  }
  return arrNoHead;
}

//   /\n/gi
//   /"/gi

function mrReplace(str, searchValue, replaceValue) {


  // Logger.log( mrFilterEmail mrFilterArrNoHead mrReplace str=${str}   ${typeof str !="string"});
  if (!str) { return str; }
  if (typeof str != "string") { return str; }


  var sv = new RegExp(searchValue, "gi");
  // console.log(sv);
  let ret = str.replace(sv, replaceValue);
  // console.log( (" " +str + " = "  +ret));
  return ret;
}





// Mistervova@mail.ru 23.12.2020 согласно тз от 22.12.2020 //Задача№6 Перед отправкой писем сделать проверку по Отзывам о поставщиках
// проверка по отзывам 
/* из тз от 22.12.2020 
Если письмо поставщику не отправлено (поставщик не выделен цветом во второй строке группы) 
то перед отправкой проверяем его емейл на листе “Отзывы о поставщиках” столбец B, 
если есть такой емейл в списке то смотрим его статус (столбец С) далее переходим на лист “Настройки писем” столбцы V:W. 
смотрим если в столбце напротив статуса стоит “да” то емайл этому поставщику не отправляем. 
Сам емейл во второй строке группы заливаем таким же цветом как и цвет ячейки со статусом на листе  “Настройки писем”
*/



/* за дача функции вернуть вернуть emailAddress контрагента, если этому контрагенту отправляем письма 
иначе цвет отказа   

*/
function checkCounteragenth(email) {

  // Logger.log("email=" + email);

  if (getContext().canSendLetter(email)) {
    // Logger.log("email=" + email + " отправлять  ");
    return email;
  }



  let ca = getContext().getCounteragent(email);



  // Logger.log("email=" + email + " сa =  " + ca);
  if (!ca) { return email; }


  // Logger.log("email=" + email + " статус  " + ca.getStatus());
  // Logger.log("email=" + email + " статус цвет " + getContext().getColor(ca.getStatus()));
  return ca.getStatus().getColor();
}



