
function onEditList_Buy(edit) {
  // return;
  new ClassSheet_Buy().onEdit(edit);
}



function menuMakeCopy() {
  // let classSheet_Buy = new ClassSheet_Buy();
  let classSheet_Buy = getContext().getSheetBuy();
  classSheet_Buy.menuMakeCopy();
  Logger.log(`menuMakeCopy UrlExternalSpreadSheet=${classSheet_Buy.getUrlExternalSpreadSheet()}`);
  addNewTrigger();
}

function testClassSheet_Buy() {

  Logger.log(`testClassSheet_Buy=`);
  let classSheet_Buy = new ClassSheet_Buy();
  // classSheet_Buy.menuMakeCopy();
  // Logger.log(`UrlExternalSpreadSheet=${classSheet_Buy.getUrlExternalSpreadSheet()}`);

  // classSheet_Buy.onEditHelper();
  let { retArr, retBgArr } = classSheet_Buy.Сбор_КП_ИЗ_Закупка_Товара();
  Logger.log()
  retArr.forEach(r => Logger.log(ret));
}



class ClassSheet_Buy {
  constructor() { // class constructor
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Закупка_товара);

    this.folderId = getSettings().folderId;
    this.url_patternSpreadSheet = getSettings().url_таблица_закупки_шаблон;

    this.rowHadsFirst = 1;
    this.rowHadsLast = 2;


    this.rowBodyFirst = 3;

    this.rowBodyLast = this.findRowBodyLast();

    this.makeCol();


    // this.mapTest = new Map();

  }


  onEdit(edit) {

  }

  triggerOnEditHelperBuyExternal(duration) {
    Logger.log(`ClassSheet_Buy triggerOnEditHelperBuyExternal duration=${duration}`);

    let classSheet_BuyExternal = this.getClassSheet_BuyExternal();
    if (classSheet_BuyExternal != undefined) {
      try {
        classSheet_BuyExternal.setRowBodyFirst(this.rowBodyFirst);
        classSheet_BuyExternal.setRowBodyLast(this.rowBodyLast);
        classSheet_BuyExternal.setStoredRanges(getSettings().storedRanges);
        classSheet_BuyExternal.triggerOnEditHelperBuyExternal(duration);

      } catch (err) {
        Logger.log(mrErrToString(err));
      }
    } else {
      Logger.log(`ClassSheet_Buy triggerOnEditHelperBuyExternal classSheet_BuyExternal Внешней таблици нет `);
    }

  }



  onEditHelper(duration = 1 / 24 / 60 * 4.9) {
    Logger.log(`ClassSheet_Buy onEditHelper=`);

    this.onEditHelperChekCommenrs(duration);



    let classSheet_BuyExternal = this.getClassSheet_BuyExternal();
    if (classSheet_BuyExternal != undefined) {
      try {

        classSheet_BuyExternal.setRowBodyFirst(this.rowBodyFirst);
        classSheet_BuyExternal.setRowBodyLast(this.rowBodyLast);
        classSheet_BuyExternal.setStoredRanges(getSettings().storedRanges);
        classSheet_BuyExternal.onEditHelper(duration);

      } catch (err) {
        Logger.log(mrErrToString(err));
      }
    }

  }

  makeFormulaCommentПоставщикОсновной() {
    return `{""\\OFFSET('3-3 Выбор поставщиков'!R[-1]C[10];0;MATCH(R[0]C[-2];'3-3 Выбор поставщиков'!R[-1]C[10]:R[-1]C[62];0)+1)}`;
  }
  makeFormulaCommentПоставщикРезерв() {
    return `{""\\OFFSET('3-3 Выбор поставщиков'!R[-1]C[-8];0;MATCH(R[0]C[-1];'3-3 Выбор поставщиков'!R[-1]C[-8]:R[-1]C[44];0)+1)}`;
  }

  helperПоставщикОсновнойFormula(duration) {

    let formulaCommentПоставщикОсновной = `=${this.makeFormulaCommentПоставщикОсновной()}`;
    let vlsFormuls = this.sheet.getRange(this.rowBodyFirst, this.col_FormulaCommentПоставщикОсновной, this.rowBodyLast - this.rowBodyFirst + 1, 1).getFormulasR1C1();
    vlsFormuls = vlsFormuls.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });  // lj,jdhztv 
    // vlsFormuls = vlsFormuls.filter(f => { return f[1] == "" });
    vlsFormuls = vlsFormuls.filter((f, i, arr) => {
      // if (f[1] != formulaCommentПоставщикОсновной) {
      //   Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | row${f[1]}  | 
      //  ${f[1] != formulaCommentПоставщикОсновной}
      //  ${formulaCommentПоставщикОсновной}
      //  ${f[1]}
      //  `);
      // }
      return f[1] != formulaCommentПоставщикОсновной
    });

    if (vlsFormuls.length == 0) {
      Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | нет пустых формул return`);
      return;
    }

    Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | vlsFormuls.length=${vlsFormuls.length}`);
    for (let i = 0; i < vlsFormuls.length; i++) {
      let row = vlsFormuls[i][0];
      let col = this.col_FormulaCommentПоставщикОсновной;
      Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | row${row}  | col ${nc(col)} 

       ${formulaCommentПоставщикОсновной}
       ${vlsFormuls[i][1]}
       `);

      if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | мало времени break`); break; }
      // this.sheet.getRange(row, this.col_CommentОсновной).setFormula(formulaCommentПоставщикОсновной);
      this.sheet.getRange(row, col).setFormula(formulaCommentПоставщикОсновной);
    }
  }


  helperПоставщикРезервFormula(duration) {

    let formulaCommentПоставщикРезерв = `=${this.makeFormulaCommentПоставщикРезерв()}`;
    let vlsFormuls = this.sheet.getRange(this.rowBodyFirst, this.col_FormulaCommentПоставщикРезерв, this.rowBodyLast - this.rowBodyFirst + 1, 1).getFormulasR1C1();
    vlsFormuls = vlsFormuls.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });  // lj,jdhztv 
    // vlsFormuls = vlsFormuls.filter(f => { return f[1] == "" });
    vlsFormuls = vlsFormuls.filter((f, i, arr) => {
      // if (f[1] != formulaCommentПоставщикОсновной) {
      //   Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | row${f[1]}  | 
      //  ${f[1] != formulaCommentПоставщикОсновной}
      //  ${formulaCommentПоставщикОсновной}
      //  ${f[1]}
      //  `);
      // }
      return f[1] != formulaCommentПоставщикРезерв
    });

    if (vlsFormuls.length == 0) {
      // Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | нет пустых формул return`);
      return;
    }

    // Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | vlsFormuls.length=${vlsFormuls.length}`);
    for (let i = 0; i < vlsFormuls.length; i++) {
      let row = vlsFormuls[i][0];
      let col = this.col_FormulaCommentПоставщикРезерв;
      // Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | row${row}  | col ${nc(col)} 

      //  ${formulaCommentПоставщикРезерв}
      //  ${vlsFormuls[i][1]}
      //  `);

      if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy helperПоставщикОсновнойFormula | мало времени break`); break; }
      // this.sheet.getRange(row, this.col_CommentОсновной).setFormula(formulaCommentПоставщикОсновной);
      this.sheet.getRange(row, col).setFormula(formulaCommentПоставщикРезерв);
    }

  }


  helperПоставщикОсновнойVelue(duration) {
    // this.col_FormulaComment // формула Комментарий к поставщику
    // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs `);
    let vlsFormuls = this.sheet.getRange(this.rowBodyFirst, this.col_FormulaCommentПоставщикОсновной, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();

    vlsFormuls = vlsFormuls.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    vlsFormuls = vlsFormuls.filter(f => { return f[1] != "" });


    if (vlsFormuls.length == 0) {
      // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | нет новых коментариев return`);
      return;
    }

    // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | vlsFormuls.length=${vlsFormuls.length}`);


    for (let i = 0; i < vlsFormuls.length; i++) {
      let row = vlsFormuls[i][0];
      let vls = this.sheet.getRange(row, 1, 1, this.col_CommentОсновной).getValues();
      if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | мало времени break`); break; }
      let comment = vls[0][this.col_CommentОсновной - 1];
      let tovar = vls[0][this.col_Tovar - 1];
      let agent = vls[0][this.col_AgentОсновной - 1];
      let price = vls[0][this.col_PriceОсновной - 1];
      let id = vls[0][this.col_IdProduct - 1];

      if (getContext().getSheetList7().insertComment(comment, tovar, agent, price, id)) {
        // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | удалось перенисти коментарий на Лист7 ${JSON.stringify([comment, tovar, agent, price, id])}`);
        this.sheet.getRange(row, this.col_CommentОсновной).clearContent();
      } else {
        // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | не удалось перенисти коментарий на Лист7 ${JSON.stringify([comment, tovar, agent, price, id])}`);
      }
    }
    // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs finis`);


  }


  helperПоставщикРезервVelue(duration) {

    // this.col_FormulaComment // формула Комментарий к поставщику
    // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs `);
    let vlsFormuls = this.sheet.getRange(this.rowBodyFirst, this.col_FormulaCommentПоставщикРезерв, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();

    vlsFormuls = vlsFormuls.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    vlsFormuls = vlsFormuls.filter(f => { return f[1] != "" });


    if (vlsFormuls.length == 0) {
      // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | нет новых коментариев return`);
      return;
    }
    // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | vlsFormuls.length=${vlsFormuls.length}`);
    for (let i = 0; i < vlsFormuls.length; i++) {
      let row = vlsFormuls[i][0];
      let vls = this.sheet.getRange(row, 1, 1, this.col_CommentРезерв).getValues();
      if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | мало времени break`); break; }
      let comment = vls[0][this.col_CommentРезерв - 1];
      let tovar = vls[0][this.col_Tovar - 1];
      let agent = vls[0][this.col_AgentРезерв - 1];
      let price = vls[0][this.col_PriceРезерв - 1];
      let id = vls[0][this.col_IdProduct - 1];

      if (getContext().getSheetList7().insertComment(comment, tovar, agent, price, id)) {
        // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | удалось перенисти коментарий на Лист7 ${JSON.stringify([comment, tovar, agent, price, id])}`);
        this.sheet.getRange(row, this.col_CommentРезерв).clearContent();
      } else {
        // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | не удалось перенисти коментарий на Лист7 ${JSON.stringify([comment, tovar, agent, price, id])}`);
      }
    }
    // Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs finis`);



  }




  onEditHelperChekCommenrs(duration) {
    //----------------------------------------------------------------------
    if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | мало времени return`); return; }
    this.helperПоставщикОсновнойFormula(duration);

    //----------------------------------------------------------------------  

    if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | мало времени return`); return; }
    this.helperПоставщикРезервFormula(duration);

    //----------------------------------------------------------------------

    if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | мало времени return`); return; }
    this.helperПоставщикОсновнойVelue(duration);
    //----------------------------------------------------------------------

    if (!getContext().hasTime(duration)) { Logger.log(`ClassSheet_Buy onEditHelperChekCommenrs | мало времени return`); return; }
    this.helperПоставщикРезервVelue(duration);

    //----------------------------------------------------------------------

  }


  getClassSheet_BuyExternal() {
    // Logger.log(`ClassSheet_Buy getClassSheet_BuyExternal=`);
    let urlExternalSpreadSheet = this.getUrlExternalSpreadSheet();

    // Logger.log(`ClassSheet_Buy urlExternalSpreadSheet=${urlExternalSpreadSheet}`);

    let classSheet_BuyExternal = undefined;
    if (urlExternalSpreadSheet) {
      try {
        let ss = SpreadsheetApp.openByUrl(urlExternalSpreadSheet);
        classSheet_BuyExternal = new ClassSheet_BuyExternal(ss);
      } catch (err) {
        Logger.log(mrErrToString(err));
        classSheet_BuyExternal = undefined;
      }

    }
    return classSheet_BuyExternal;
  }

  findRowBodyLast() {
    let ret = getContext().getRowSobachkaBySheetName(this.sheet.getSheetName());
    if (!ret) { ret = this.sheet.getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }



  makeFormulaForComment() {

  }


  makeCol() {
    this.col_A = nr("A");
    this.col_B = nr("B");
    this.col_C = nr("C");
    this.col_D = nr("D");
    this.col_FormulaCommentПоставщикОсновной = nr("L");  // формула Комментарий к поставщику
    this.col_FormulaCommentПоставщикРезерв = nr("AD");  // формула Комментарий к поставщику
    this.col_IdProduct = nr("A")//№ 
    this.col_Tovar = nr("B")// Номенклатура
    this.col_PriceОсновной = nr("G")// Цена
    this.col_PriceРезерв = nr("Z")// Цена
    this.col_AgentОсновной = nr("J")// Поставщик 
    this.col_AgentРезерв = nr("AC")// Поставщик 
    this.col_CommentОсновной = nr("M");  // Комментарий к поставщику
    this.col_CommentРезерв = nr("AE");  // Комментарий к поставщику

    this.col_Реальная_цена_Основной = nr("U");
    this.col_Реальный_срок_Основной = nr("V");
    this.col_Реальная_цена_Резерв = nr("AM");
    this.col_Реальный_срок_Резерв = nr("AN");
  }



  menuMakeCopy() {
    if (!true) { return; }
    this.makePatternCopy();

    // this.getClassSheet_BuyExternal()
    let classSheet_BuyExternal = this.getClassSheet_BuyExternal();
    if (classSheet_BuyExternal) {
      classSheet_BuyExternal.onEditHelper(1 / 24 / 60 * 3.9);
    }
    let duration = 1 / 24 / 60 * 2
    this.triggerOnEditHelperBuyExternal(duration);
    // this.addNewTriggeronEditTrigger();

  }


  makePatternCopy() {
    // let formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd' 'HH:mm:ss");
    let formattedDate = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd' 'HH:mm:ss");
    let name = SpreadsheetApp.getActiveSpreadsheet().getName() + " | ТАБЛИЦА ЗАКУПКИ ВНЕШНЯЯ | " + formattedDate;
    let destination = (() => { try { return DriveApp.getFolderById(this.folderId); } catch (err) { mrErrToString(err); return undefined } })();

    let file = DriveApp.getFileById(SpreadsheetApp.openByUrl(this.url_patternSpreadSheet).getId());
    let patternCopy = (() => { if (destination) { return file.makeCopy(name, destination); } else { return file.makeCopy(name); } })();
    patternCopy.addEditor("script@ss-postavka.ru");
    let url_patternCopy = patternCopy.getUrl();
    Logger.log(`url_patternCopy = ${url_patternCopy}`);
    this.setUrlExternalSpreadSheet(url_patternCopy);

    let shNamesAcc = ["НАСТРОЙКИ", "mem"];
    let msg = `ТАБЛИЦА ЗАКУПКИ ВНЕШНЯЯ созданна. Проверьте права доступа у пользователя "${Session.getEffectiveUser().getEmail()}" к листам: ${shNamesAcc.map(v => `"${v}"`).join(", ")} 
    \n  Название: ${name}. Ссылка: ${url_patternCopy}. `;
    Browser.msgBox(msg);
  }



  getExternalSpreadSheet() {
    if (!this.externalSpreadSheet) {
      let urlExternalSpreadSheet = this.getUrlExternalSpreadSheet();
      if (!urlExternalSpreadSheet) { this.externalSpreadSheet = undefined; }
      try {
        this.externalSpreadSheet = SpreadsheetApp.openByUrl(urlExternalSpreadSheet)
      } catch (err) {

        // Logger.log(mrErrToString(err));
        // Logger.log("");
        this.externalSpreadSheet = undefined;
      }
    }
    return this.externalSpreadSheet;
  }


  getUrlExternalSpreadSheet() {
    let row = this.getRowUrlExternalSpreadSheet()
    let col = this.col_C;
    let url = this.sheet.getRange(row, col).getValue();
    return url;
  }

  getRowUrlExternalSpreadSheet() {
    let str = getSettings().str_Ссылка_на_таблицу_закупки;
    let col = this.col_B;
    let ret = undefined;

    let rowStart = this.rowBodyLast + 1;
    let rowEnd = this.sheet.getLastRow();
    if (rowStart > rowEnd) {
      rowStart = rowEnd;
    }

    let vls = this.sheet.getRange(rowStart, col, rowEnd - rowStart + 1, 1).getValues();
    vls = [].concat(vls.flat());
    let ind = vls.lastIndexOf(str);
    if (ind != -1) { ret = ind + rowStart; }
    if (ret == undefined) {
      let arr = new Array(col);
      arr[col - 1] = str;
      this.sheet.appendRow(arr);
      return this.getRowUrlExternalSpreadSheet();
    }

    Logger.log(`getRowUrlExternalSpreadSheet = ${ret}`);
    return ret;
  }

  setUrlExternalSpreadSheet(url) {
    let row = this.getRowUrlExternalSpreadSheet()
    let col = this.col_C;
    this.sheet.getRange(row, col).setValue(url);
  }

  /** @param {Map} retMap */
  getProductMapFromMem(retMap = new Map()) {
    // Logger.log(`getProductMapFromMem S`);
    // MrLib_Midlle.getMrClassSheetModel()  
    let rowConf = {
      head: { first: 2, last: 2, key: 2, },
      body: { first: 3, last: 3, },
    }
    let sheetModel = MrLib_Midlle.makeSheetModelBy(getContext().getUrls().product, getSettings().sheetName_mem_Оплаты, rowConf);
    sheetModel.getMap().forEach(item => {
      // Logger.log(JSON.stringify(item, null, 2));
      let pr_name = item["Номенклатура"];

      if (!retMap.has(pr_name)) {
        retMap.set(pr_name, new Map());
      }
      let цена = item["Реальная цена"];
      let срок = item["Реальный срок"];
      let agent = item["Поставщик"];
      if (!agent) { return; }
      if (!цена) { return; }
      // Поставщик	Номенклатура	Реальная цена	Реальный срок


      this.filterEmails(agent).forEach(e => {
        // pr_price.set(v[this.col_AgentОсновной], [v[this.col_Реальная_цена_Основной], v[this.col_Реальный_срок_Основной], ""]);
        // this.emailsArr.push(v[this.col_AgentОсновной]);
        this.emailsSets.mem.add(e);
        let ee = `${e} (♣️)`;
        retMap.get(pr_name).set(ee, [цена, срок, "", "(♣️)"]);
        this.emailsArr.push(ee);
      });

      // retMap.get(pr_name).set(agent, [цена, срок, "(♥️♥️♥️)"]);
      // this.emailsArr.push(agent);
    });
    // retMap.forEach((v, k) => { Logger.log( JSON.stringify({ k, ks:[...v.keys()],vs:[...v.values()] }, null, 2)) });
    return retMap
  }

  getProductMap(retMap = new Map()) {
    // Logger.log(`getProductMap S`);
    // let retMap = new Map();

    let vls_prod = this.sheet.getRange(this.rowBodyFirst, 1, this.rowBodyLast - this.rowBodyFirst + 1, this.sheet.getLastColumn()).getValues();
    vls_prod = vls_prod.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    // Logger.log(vls_head);
    // Logger.log(vls_prod);

    this.col_Product = this.col_B;
    vls_prod.forEach((v, j, arr) => {
      if (!v[this.col_Product]) { return; }
      let pr_name = v[this.col_Product];
      if (!retMap.has(pr_name)) {
        retMap.set(pr_name, new Map());
      }
      let pr_price = retMap.get(pr_name)

      if (v[this.col_Реальная_цена_Основной]) {
        if (v[this.col_AgentОсновной]) {
          // let email_key_Основной = JSON.stringify({
          //   email: v[this.col_AgentОсновной],
          //   bg: getContext().getDefaultColor(),
          //   col: this.col_Реальная_цена_Основной,
          // });

          this.filterEmails(v[this.col_AgentОсновной]).forEach(e => {
            this.emailsSets.bay.add(e);
            let ee = `${e} (♥️)`;
            pr_price.set(ee, [v[this.col_Реальная_цена_Основной], v[this.col_Реальный_срок_Основной], "", "(♥️)"]);
            this.emailsArr.push(ee);
          });
          // pr_price.set(v[this.col_AgentОсновной], [v[this.col_Реальная_цена_Основной], v[this.col_Реальный_срок_Основной], ""]);
          // this.emailsArr.push(v[this.col_AgentОсновной]);
        }
      }

      if (v[this.col_Реальная_цена_Резерв] && v[this.col_AgentРезерв]) {
        // let email_key_Резерв = JSON.stringify({
        //   email: v[this.col_AgentРезерв],
        //   bg: getContext().getDefaultColor(),
        //   col: this.col_Реальная_цена_Резерв,
        // });

        this.filterEmails(v[this.col_AgentРезерв]).forEach(e => {
          this.emailsSets.bay.add(e);
          let ee = `${e} (♥️)`;
          pr_price.set(ee, [v[this.col_Реальная_цена_Резерв], v[this.col_Реальный_срок_Резерв], "", "(♥️)"]);
          this.emailsArr.push(ee);
        });
        // pr_price.set(v[this.col_AgentРезерв], [v[this.col_Реальная_цена_Резерв], v[this.col_Реальный_срок_Резерв], ""]);
        // this.emailsArr.push(v[this.col_AgentРезерв]);
      }
      retMap.set(pr_name, pr_price);
    });

    let emailsSet = new Set();
    // this.emailsArr
    // .join(" ")
    // .split(" ")
    // .filter(e => {

    //   if (`${e}`.includes("@") || `${e}`.includes(".")) { return true; }
    //   return false;
    // })

    this.emailsSets.bay.forEach(e => {
      if (this.emailsSets.mem.has(e)) {
        this.emailsSets.mem.delete(e);
      }
    });

    // this.emailsArr.forEach(e => emailsSet.add(e));
    //  this.emailsArr = [...emailsSet.values()];

    this.emailsArr = [...this.emailsSets.mem.values()]
      .map(e => `${e} (♣️)`)
      .concat([...this.emailsSets.bay.values()].map(e => `${e} (♥️)`));

    return retMap;
  }

  /** @returns {[]} */
  filterEmails(emailStr) {
    return this.filterEmails_n(emailStr);
    let emailsSet = new Set();
    `${emailStr}`.split(" ")
      .filter(e => {
        if (`${e}`.includes("@") || `${e}`.includes(".")) { return true; }
        return false;
      })
      .forEach(e => emailsSet.add(e));
    // this.filterEmails_n(emailStr);
    return [...emailsSet.values()];
  }


  /** @returns {[]} */
  filterEmails_n(emailStr) {
    // Logger.log("filterEmails_n");
    try {
      let r = new RegExp("[♥️♠♣️♦️)(]", "gim");
      let emailsSet = new Set();

      [`${emailStr}`]
        .map(e => {
          return e.replace(r, "").split("\n").map(s => { return s.trim() }).filter(s => s).join("\n");
        })
        .filter(e => {
          return true;
          if (`${e}`.includes("@") || `${e}`.includes(".")) { return true; }
          return false;
        })
        .forEach(e => emailsSet.add(e));

      // this.mapTest(emailStr, JSON.stringify([...emailsSet.values()]));
      // MrLib_Midlle.setFildForItem("Лог", emailStr, "ret", JSON.stringify([ ...emailsSet.values()]) ); 
      return [...emailsSet.values()];
    } catch (err) {
      MrLib_Midlle.setFildForItem("Лог", emailStr, "ret", mrErrToString(err));
    }

    return [emailStr];
  }



  /** @returns {[[]]} */
  Сбор_КП_ИЗ_Закупка_Товара() {
    // Logger.log("Сбор_КП_ИЗ_Закупка_Товара S");
    let sheetList1 = getContext().getSheetList1();
    let productMap = new Map();
    this.emailsSets = {
      mem: new Set(),
      bay: new Set(),
    };
    this.emailsArr = new Array();
    productMap = this.getProductMapFromMem(productMap);
    productMap = this.getProductMap(productMap);
    // productMap = this.getProductMapFromMem(new Map());

    // productMap.forEach((v, k) => { Logger.log(JSON.stringify({ k, ks: [...v.keys()], vs: [...v.values()] }, null, 2)) });

    // productMap.forEach((v, k, map) => { Logger.log(`${k} | ${v.size} | ${[...v.values()]} ${[...v.keys()]}`) })
    this.emailsArr.forEach(e => Logger.log(JSON.stringify(e)));
    let emailsSet = new Set();
    this.emailsArr.forEach(e => emailsSet.add(e));
    let emailsArr = [...emailsSet.values()];

    // let emailsArr = this.emailsArr;



    // Logger.log(emailsArr);


    let productArr = sheetList1.sheet.getRange(sheetList1.rowBodyFirst, sheetList1.columnProduct, sheetList1.rowBodyLast - sheetList1.rowBodyFirst + 1, 1).getValues();
    productArr = productArr.map((v, i, arr) => { return v[0] });
    // Logger.log(productArr);

    // let num_col = Math.max(emailsArr.length * 3 + 6, sheetList1.columnLastPrice - sheetList1.columnАктуальный_Сбор_КП);
    let num_col = Math.max(emailsArr.length * 3, 6);
    // Logger.log(num_col);
    // return
    let retArr = new Array();

    // let h = [].concat(emailsArr.map((v, i, arr) => { return [v, undefined] }).flat());
    let h = new Array(num_col);
    for (let e = 0, i = 0; e < emailsArr.length; e++, i += 3) {
      h[i] = emailsArr[e];
    }

    // Logger.log(h);

    // Logger.log(h)
    retArr.push(h);

    for (let i = 0; i < productArr.length; i++) {
      // Logger.log("---------------------------------------------------------------");
      let r = new Array(num_col);
      retArr.push(r);
      let pr_name = productArr[i];
      if (!pr_name) { continue; }
      if (!productMap.has(pr_name)) { continue; }
      // Logger.log(pr_name)
      /** @type {Map} */
      let email_map = productMap.get(pr_name);
      if (email_map.size == 0) { continue; }
      // Logger.log([...email_map.keys()]);
      for (let e = 0; e < h.length; e += 3) {
        if (!email_map.has(h[e])) { continue; }
        let v = email_map.get(h[e]);
        r[e] = v[0];
        r[e + 1] = v[1];
        r[e + 2] = v[2];
      }
      // Logger.log(r);
    }

    let retBgArr = [new Array(num_col)];



    // let знак = "(♥️)";


    // for (let e = 0; e < retArr[0].length; e += 3) {
    //   if (!retArr[0][e]) { continue; }
    //   // retArr[0][e] = `${retArr[0][e]} (* ${Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyy-MM-dd HH:mm:ss")}) `;
    //   retArr[0][e] = `${retArr[0][e]} ${знак}`;
    // }





    // Logger.log(JSON.stringify({ retArr, retBgArr }));
    return { retArr, retBgArr };
  }






}


class ClassSheet_BuyExternal {
  /**@param {SpreadsheetApp.Spreadsheet} spreadsheet */
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.settings = getSettings().settingsBuyExternal;
  }

  getSheetByName(sheetName) {
    let ss = this.spreadsheet;
    let sh = ss.getSheetByName(sheetName);
    if (!sh) { throw `нет листа с именем ${sheetName}`; }
    return sh;
  }


  onEditHelper(duration = 1 / 24 / 60 * 4.9) {
    Logger.log(`ClassSheet_BuyExternal onEditHelper= Start`);

    this.getSheetByName("НАСТРОЙКИ").getRange("A1").setBackground("#12DaFF");
    this.getSheetByName("НАСТРОЙКИ").getRange("A1").setValue(`${JSON.stringify(this.settings)}`);
    // this.triggerOnEditHelperBuyExternal(duration);
    // this.onEditHelperBuyExternal(duration);

    Logger.log(`ClassSheet_BuyExternal onEditHelper= Finish`);
  }



  triggerOnEditHelperBuyExternal(duration) {
    Logger.log(`ClassSheet_BuyExternal triggerOnEditHelperBuyExternal=`);
    // MrLib_External.triggerMrOnEditHelper("из отновной таблицы",duration, this.spreadsheet, getContext().timeConstruct, this.settings);
    MrLib_External.triggerMrOnEditHelperExternal(duration, this.spreadsheet, getContext().timeConstruct);
  }

  onEditHelperBuyExternal(duration) {
    Logger.log(`ClassSheet_BuyExternal onEditHelperBuyExternal`);
    MrLib_External.onEditHelper(duration, this.spreadsheet, getContext().timeConstruct);
  }

  setRowBodyFirst(row) {
    this.settings.rowBodyFirst = row;
    this.settings["тело начало"] = row;
  }
  setRowBodyLast(row) {
    this.settings.rowBodyLast = row;
    this.settings["тело конец"] = row;
  }

  setStoredRanges(storedRanges) {
    this.settings.storedRanges = storedRanges;
  }

}









