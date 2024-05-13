function onEditList_Expertise(edit, sheetName) {
  // return;
  new ClassSheet_Expertise(sheetName).onEdit(edit);
}


class ClassSheet_Expertise {
  // constructor(sheetName = "Экспертиза") {
  constructor(sheetName = getSettings().sheetName_Экспертиза) {
    this.sheetName = sheetName
    this.sheet = getContext().getSheetByName(this.sheetName);


    this.row_Header = getClassColRow().expertise_row_Header;
    this.row_BodyFirst = getClassColRow().expertise_row_BodyFirst;
    this.list1_row_BodyFirst = getClassColRow().list1_row_BodyFirst;


    this.names = {
      Номер: "№",
      Номенклатура: "Номенклатура",
      Характеристики: "Характеристики",
      Ед_Изм: "Ед. Изм",
      Кол_во: "Кол-во",
      Комментарий: "Комментарий",
      Цена: "Цена",
      Сумма: "Сумма",
      БлокТопНачало: "БлокТопНачало",
      БлокТопКонец: "БлокТопКонец",
      ТопЭксперта: "ТопЭксперта",
      НоменклатураПроверка: "НоменклатураПроверка",
      Лист1НомерСтроки: "Лист1НомерСтроки",
      Цена: "Цена",
      Поставщик: "Поставщик",
      Срок: "Срок",
      Модель: "Модель",
      Изготовитель: "Изготовитель",
      КомментарийЭксперта: "Комментарий эксперта",
      Подходит: "Подходит",
      КомментарийМенеджера: "Комментарий менеджера",
      Формула: "Формула",
      row: "row",

    }

    this.cols = {
      Номер: 0,
      Номенклатура: 0,
      Характеристики: 0,
      Ед_Изм: 0,
      Кол_во: 0,
      Комментарий: 0,
      Цена: 0,
      Сумма: 0,
      БлокТопНачало: 0,
      БлокТопКонец: 0,
      ТопЭксперта: 0,
      НоменклатураПроверка: 0,
      Лист1НомерСтроки: 0,
      row: 0,
      // Цена: "Цена",
      // Поставщик: "Поставщик",
      // Срок: "Срок",
      // Модель: "Модель",
      // Изготовитель: "Изготовитель",
      // КомментарийЭксперта: "Комментарий эксперта",
      // Подходит: "Подходит",
      // КомментарийМенеджера: "Комментарий менеджера",
      // Формула: "Формула",

    }


    this.rows = {
      headr: {
        key: 1,
      },
      body: {
        first: 2,
        last: 3,
      }
    }


    this.sheetNameFrom = getSettings().sheetName_Лист_1;
    this.map = undefined;
    this.makeRows();
    this.makeCols();

    this.col_blok_formul_start = this.cols.ТопЭксперта;
    this.col_blok_formul_end = this.col_blok_formul_start + 2;
    this.col_list1_row = this.cols.Лист1НомерСтроки; //Номер строки на Листе1

    this.col_blok_koment_start = this.cols.БлокТопНачало + 1;
    this.col_blok_koment_end = this.cols.БлокТопКонец - 1;

    this.col_top_join = this.cols.ТопЭксперта;

    this.col_commentBlokStart = getClassColRow().list1_col_CommentExpertBlokStart;
    this.col_commentBlokEnd = getClassColRow().list1_col_CommentExpertBlokEnd;

    this.make();
    Logger.log(JSON.stringify(this, null, 2));
    Logger.log(` ClassSheet_Expertise =${this.sheet.getSheetName()}`);

  }

  getNewEdit(row, col) {

    Logger.log(`getNewEdit=${nc(col)}${row}`);
    let range = this.sheet.getRange(row, col, 1, 1);
    range["rowStart"] = row;
    range["rowEnd"] = row;
    range["columnStart"] = col;
    range["columnEnd"] = col;

    let edit = {
      "range": range
    }
    return edit;
  }



  fixSheet(duration = 1 / 24 / 60 * 4.9) {
    if (getContext().timeConstruct.getHours() != 3) {
      // if (!isTest())
        return;
    }
    // return;
    Logger.log("ClassSheet_Expertise fixSheet S");
    // this.row_BodyFirst
    // let rowLast = this.sheet.getLastRow();
    // let rowLast = getContext().getRowSobachkaBySheetName(this.sheetName) - 1; 
    let rowDif = this.list1_row_BodyFirst - this.rows.body.first;
        // let rowLast = getContext().getRowSobachkaBySheetName(this.sheetNameFrom) - 1 ;
    let rowLast = getContext().getRowSobachkaBySheetName(this.sheetNameFrom) - 1 + (rowDif * -1);
    if (rowLast < this.row_BodyFirst) { return; }


    Logger.log("ClassSheet_Expertise fixSheet rowLast " + rowLast);


    let isMmodified = false;
    let formulaNr = this.makeFormulaNr();
    let rangeNr = this.sheet.getRange(this.row_BodyFirst, nr("A"), rowLast - this.row_BodyFirst + 1 + 1, 1);
    rangeNr.getFormulasR1C1().flat().forEach(f => { if (f != "=" + formulaNr) { isMmodified = true } });
    if (isMmodified) {
      Logger.log(`isMmodified rangeNr ${isMmodified}`);
      rangeNr.setFormulaR1C1(formulaNr);
    }


    let f = [
      // ТопЭксперта	
      // НоменклатураПроверка	
      // Лист1НомерCтроки
      this.makeFormulaТопЭксперта(),
      this.makeFormulaНоменклатураПроверка(),
      this.makeFormulaЛист1НомерСтроки(),

      // this.makeFormula_nr(),
      // this.makeFormula_name(),
      // this.makeFormula_Top_14_join(), 
      // this.makeFormula_coment_manager(),
      // this.makeFormula_coment_Expert(),
      // this.makeFormula_coment_Podhodit(),
    ]

    f = f.map(v => `=${v}`);
    // let formulaJoin = `${f}`
    // let formulaJoin = f.join(" | ");
    let strJ = "@@|@@";

    let formulaJoin = f.join(strJ);
    // let formulaEmptyJoin = new Array(f.length).join(strJ);



    let rangeF = this.sheet.getRange(this.row_BodyFirst, this.col_blok_formul_start, rowLast - this.row_BodyFirst + 1, this.col_blok_formul_end - this.col_blok_formul_start + 1);
    let list1Values = this.sheet.getRange(this.row_BodyFirst, this.col_list1_row, rowLast - this.row_BodyFirst + 1, 1).getValues();

    let formulas = rangeF.getFormulasR1C1();
    // Logger.log(`formulas ${formulas}`);
    for (let i = 0; i < list1Values.length; i++) {
      let isMmodified = false
      let rowF = formulas[i].join(strJ);
      // Logger.log(`заполнено  ${rowF != formulaJoin} | \n|||${rowF}||| \n|||${formulaJoin}|||`);
      if (rowF != formulaJoin) { isMmodified = true }
      if (isMmodified) {
        Logger.log(`formulas top isMmodified   row = ${this.row_BodyFirst + i}`);
        // this.onEdit(this.getNewEdit(i + this.row_BodyFirst, this.col_blok_formul_start));
        this.sheet.getRange(this.row_BodyFirst + i, this.col_blok_formul_start, 1, this.col_blok_formul_end - this.col_blok_formul_start + 1).setFormulasR1C1([f]);
      }
    }


    // -------------------------------


    let rangeHeads = this.sheet.getSheetValues(this.rows.headr.key, this.col_blok_koment_start, 1, this.col_blok_koment_end - this.col_blok_koment_start + 1);

    for (let i = 0, col = this.col_blok_koment_start; col <= this.col_blok_koment_end; col++, i++) {
      if (!((this.col_blok_koment_start <= col) && (col <= this.col_blok_koment_end))) { continue; }
      let h = fl_str(rangeHeads[0][i]);
      if (!h.includes(fl_str(this.names.Формула))) { continue; }

      let hn = parseInt(h);
      if (isNaN(hn)) { continue; }
      hn = hn - 1;

      let formula = this.getFormulaForCol(hn + 1);
      let range = this.sheet.getRange(this.row_BodyFirst, col, rowLast - this.row_BodyFirst + 1, 1);
      let formulas = range.getFormulasR1C1();
      let isMmodified = false;
      //  Logger.log(`formulas col isMmodified chek = ${col}`);
      formulas = formulas.flat();
      formulas.forEach(f => { if (f.slice(1) != formula) { isMmodified = true } });
      if (isMmodified) {
        Logger.log(`formulas col isMmodified col = ${col}`);
        range.setFormulaR1C1(formula);
      }

      // this.sheet.hideColumn(range);

      if (getContext().timeConstruct.getHours() != 3) { continue; }
      if (!this.sheet.isColumnHiddenByUser(col)) { this.sheet.hideColumn(range); }


    }

    // -------------------------------


    this.fixSheetCompact()
    Logger.log("ClassSheet_Expertise fixSheet F");
  }


  fixSheetCompact() {
    if (getContext().timeConstruct.getHours() != 3) { return; }
    // this.col_blok_koment_start
    // this.col_blok_koment_end

    // if (getContext().timeConstruct.getHours() != 1) { return }
    let columnLast = this.sheet.getMaxColumns();
    this.sheet.hideColumn(this.sheet.getRange(1, this.cols.БлокТопКонец, 1, columnLast - this.cols.БлокТопКонец + 1))


    let ww = new Array(16)
    for (let i = 0; i < 16; i++) {
      ww[i] = this.sheet.getColumnWidth(this.col_blok_koment_start + i);

    }

    let f = (this.col_blok_koment_end - this.col_blok_koment_start);
    for (let c = 0; c < f; c++) {
      let col = c + this.col_blok_koment_start;
      let w = ww[0];
      let i = c % 16;
      w = ww[i];
      this.sheet.setColumnWidth(col, w)
    }
  }




  onEditHelper(duration = 1 / 24 / 60 * 4.9) {
    // return;
    // let colStart = this.col_blok_koment_start;
    Logger.log(` onEditHelper= ${this.sheetName}`);

    // return;
    // if (getContext().getUrls().product != "https://docs.google.com/spreadsheets/d/1A/edit") { return; }
    // let rowLast = this.sheet.getLastRow();
    let rowLast = getContext().getRowSobachkaBySheetName(this.sheetName) - 1;
    if (rowLast < this.row_BodyFirst) { return; }


    Logger.log(` onEditHelper= ${JSON.stringify({
      row_BodyFirst: this.row_BodyFirst, rowLast, col_blok_koment_start: this.col_blok_koment_start, col_blok_koment_end: this.col_blok_koment_end
    }, null, 2)}`);

    let headsValues = this.sheet.getRange(this.rows.headr.key, this.col_blok_koment_start, 1, this.col_blok_koment_end - this.col_blok_koment_start + 1).getValues();
    let bodyFormula = this.sheet.getRange(this.row_BodyFirst, this.col_blok_koment_start, rowLast - this.row_BodyFirst + 1, this.col_blok_koment_end - this.col_blok_koment_start + 1).getFormulas();

    bodyFormula = bodyFormula.map((v, i, arr) => { return [i + this.row_BodyFirst].concat(v); });
    headsValues = headsValues.map((v, i, arr) => { return [i + this.row_BodyFirst].concat(v); });

    // let list1Values = this.sheet.getRange(this.row_BodyFirst, this.col_list1_row, rowLast - this.row_BodyFirst + 1, 1).getValues();
    // bodyFormula = bodyFormula.filter((f, i, arr) => list1Values[i][0] != "");


    let strformula = fl_str(this.names.Формула);

    // for (let i = 0; i < headsValues[0].length; i++) {
    //   if (!fl_str(headsValues[0][i]).includes(strformula)) { continue; }
    //   let bf = bodyFormula.map(f => { return f[i] });
    //   if (bf.filter(f => f == "").length != 0) {
    //     let problerBf = bodyFormula.filter(f => f[i] == "");
    //     Logger.log(` Сол= ${nc(i + this.col_blok_koment_start - 1)} = ${problerBf.map(f => f[0])}`);

    //     let col = i + this.col_blok_koment_start - 1;
    //     for (let p = 0; p < problerBf.length; p++) {
    //       if (!getContext().hasTime(duration)) { break; }
    //       this.onEdit(this.getNewEdit(problerBf[p][0], col));
    //       // вызиваем от едит для кажндой проблемной формулы

    //     }
    //   }
    // }

    let bodyValues = this.sheet.getRange(this.row_BodyFirst, this.col_blok_koment_start, rowLast - this.row_BodyFirst + 1, this.col_blok_koment_end - this.col_blok_koment_start + 1).getValues();
    bodyValues = bodyValues.map((v, i, arr) => { return [i + this.row_BodyFirst].concat(v); });

    // bodyValues = bodyValues.filter((f, i, arr) => list1Values[i][0] != "");
    // Logger.log(` проверка формул этап 2`);
    for (let i = 0; i < headsValues[0].length; i++) {
      if (!fl_str(headsValues[0][i]).includes(strformula)) { continue; }
      let bv = bodyValues.map(f => { return f[i] });
      if (bv.filter(v => v != "").length != 0) {
        let problerBv = bodyValues.filter(f => f[i] != "");
        Logger.log(` Сол= ${nc(i + this.col_blok_koment_start - 1)} = ${problerBv.map(f => f[0])}`);

        let col = i + this.col_blok_koment_start - 1 + 1;
        for (let p = 0; p < problerBv.length; p++) {
          if (!getContext().hasTime(duration)) { break; }
          // this.onEdit(this.getNewEdit(problerBv[p][0], col));
          this.onEdit_blok_koment(this.getNewEdit(problerBv[p][0], col));
          // вызиваем от едит для кажндой проблемной формулы
        }
      }
    }



  }



  onEdit(edit) {
    // return
    // Получаем диапазон ячеек, в которых произошли изменения
    let range = edit.range;

    // Лист, на котором производились изменения
    let sheet = range.getSheet();

    // Проверяем, нужный ли это нам лист
    Logger.log("onEditList_Экспертиза");
    if (fl_str(sheet.getName()) != fl_str(this.sheetName)) {

      Logger.log("не наш лист " + `${sheet.getName()} = ${(fl_str(sheet.getName()) != fl_str(this.sheetName))} = ${this.sheetName}`);

      return false;
    }

    let rowStart = range.rowStart;
    let rowEnd = range.rowEnd;
    if (rowStart < this.row_BodyFirst) { rowStart = this.row_BodyFirst; }
    if (rowEnd < rowStart) { return; }

    Logger.log("onEditList_Экспертиза_конец");
  }

  onEdit_col_nr_list1(edit) {
    Logger.log("onEdit_col_nr_list1 =" + `${JSON.stringify(edit)}`)
    let range = edit.range;
    let sheet = range.getSheet();
    if (fl_str(sheet.getName()) != fl_str(this.sheetName)) {

      Logger.log("не наш лист " + `${sheet.getName()} = ${(fl_str(sheet.getName()) != fl_str(this.sheetName))} = ${this.sheetName}`);

      return false;
    }

    let rowStart = range.rowStart;
    let rowEnd = range.rowEnd;
    if (rowStart < this.row_BodyFirst) { rowStart = this.row_BodyFirst; }
    if (rowEnd < rowStart) { return; }

    let vls = this.sheet.getRange(rowStart, this.col_list1_row, rowEnd - rowStart + 1, 1).getValues();
    for (let r = 0, row = rowStart; r < vls.length; r++, row++) {

      let f = [
        this.makeFormulaТопЭксперта(),
        this.makeFormulaНоменклатураПроверка(),
        this.makeFormulaЛист1НомерСтроки(),
      ]
      this.sheet.getRange(row, this.col_blok_formul_start, 1, this.col_blok_formul_end - this.col_blok_formul_start + 1).setFormulas([f]);


    }

    Logger.log("onEdit_col_nr_list1 =")

  }

  makeRows() {
    let arrFirstCol = ["col"].concat(this.sheet.getRange(`A:A`).getValues().flat()).map(v => fl_str(v));
    let f = (arr, str, def) => {
      let ret = -1; ret = arr.indexOf(fl_str(str));
      if (ret == -1) { throw  new Error(`Нет ${str} в ${arr}`); }
      return ret;
    };

    this.rows.headr.key = f(arrFirstCol, "№", 1);
    this.rows.body.first = f(arrFirstCol, "№", 1) + 1;
    this.rows.body.last = (()=>{ try{ return f(arrFirstCol, "@", 302) - 1}catch (err){ mrErrToString(err); return this.sheet.getLastRow()} })();

    this.row_Header = this.rows.headr.key;
    this.row_BodyFirst = this.rows.body.first;


    getClassColRow().expertise_row_Header = this.row_Header;
    getClassColRow().expertise_row_BodyFirst = this.row_BodyFirst;


    Logger.log("makeRows");
    Logger.log(JSON.stringify(this.rows, null, 2));

  }


  makeCols() {
    this.arrHeardNames = ["row"].concat(this.sheet.getRange(`${this.row_Header}:${this.row_Header}`).getValues()[0]).map(v => fl_str(v));
    for (let k in this.cols) {
      let name = fl_str(this.names[k]);
      let c = this.arrHeardNames.indexOf(name);
      if (c > 0) { this.cols[k] = c; }
    }

    this.mapHeardNames = new Map();
    this.arrHeardNames.forEach((n, i) => {
      if (!n) { return; }
      this.mapHeardNames.set(n, i);
    });




  }




  make() {
    if (!this.map) { this.map = new Map() }
    let rowLast = this.sheet.getLastRow();

    if (rowLast < this.row_BodyFirst) { return; }
    let vls = this.sheet.getRange(1, this.col_list1_row, rowLast, 1).getValues();

    for (let i = 0, row = 1; i < vls.length; i++, row++) {
      let rowId = vls[i][0];
      if (this.map.has(rowId)) {
        continue;
      }
      this.map.set(rowId, row);
    }

  }


  getFormulasForBlokComment() {
    let retArr = new Array();

    for (let i = 0; i < 14 * 8; i++) {
      retArr.push(this.getFormulaForCol(i + 1));
      retArr.push(undefined);
    }
    return retArr;
  }

  getFormulaForCol(number) {

    if (typeof number != "number") { throw Error("не является числом " + number); }
    if (number < 1 || 8 * 14 < number) { throw Error("нет формулы с порядковым номером " + number); }


    let x = number % 8;

    if (x == 1 || x == 2 || x == 3 || x == 4) {
      return `{""\\IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.col_top_join};"|||";FALSE;FALSE);1;${number})));"")}`;
    }
    else {
      return `{""\\IFERROR(INDEX(SPLIT(R[0]C${this.col_top_join};"|||";FALSE;FALSE);1;${number});"")}`;
    }

  }


  makeFormulaNr() {
    let sheetName = getSettings().sheetName_Лист_1;
    let rowDif = this.list1_row_BodyFirst - this.rows.body.first;
    return `'${sheetName}'!R[${rowDif}]C3`;
  }


  // ТопЭксперта	НоменклатураПроверка	Лист1НомерСтроки
  makeFormulaТопЭксперта() {
    let ce = `C${getClassColRow().list1_col_ТопЭксперта}`;
    let cr = `C${this.col_list1_row}`; // Лист1НомерСтроки
    let sheetName = getSettings().sheetName_Лист_1
    let ret = `INDIRECT( ADDRESS(R[0]${cr}; COLUMN('${sheetName}'!R1${ce}); 1; TRUE; "${sheetName}"))`
    return ret;
  }

  makeFormulaНоменклатураПроверка() {
    let ce = `C${getClassColRow().list1_col_ТопЭксперта}`;
    let cr = `C${this.col_list1_row}`; // Лист1НомерСтроки
    let sheetName = getSettings().sheetName_Лист_1

    let ret = `INDIRECT( ADDRESS(R[0]${cr}; COLUMN('${sheetName}'!R1C4); 1; TRUE; "${sheetName}"))`
    return ret;
  }
  makeFormulaЛист1НомерСтроки() {
    let ret = `ROW(INDIRECT(MIDB(FORMULATEXT(INDIRECT(ADDRESS(ROW();1;1;TRUE)));2;LENB(FORMULATEXT(INDIRECT(ADDRESS(ROW();1;1;TRUE)))))))`
    return ret;
  }


  onEdit_blok_koment(edit) {
    // Получаем диапазон ячеек, в которых произошли изменения
    let range = edit.range;

    // Лист, на котором производились изменения
    let sheet = range.getSheet();

    // Проверяем, нужный ли это нам лист
    // Logger.log("onEditList_Экспертиза 3");
    if (fl_str(sheet.getName()) != fl_str(this.sheetName)) {
      Logger.log("не наш лист " + `${sheet.getName()} = ${(fl_str(sheet.getName()) != fl_str(this.sheetName))} = ${this.sheetName}`);
      return false;
    }



    let rowStart = range.rowStart;
    let rowEnd = range.rowEnd;
    if (rowStart < this.row_BodyFirst) { rowStart = this.row_BodyFirst; }
    if (rowEnd < rowStart) { return; }

    let colStart = range.columnStart;
    let colEnd = range.columnEnd
    let rangeHeads = this.sheet.getSheetValues(this.rows.headr.key, colStart, 1, colEnd - colStart + 1);
    let rangeValues = range.getValues();

    // this.sheet.getRange(rowStart, this.col_blok_koment_start, rowEnd - rowStart + 1, this.col_blok_koment_end - this.col_blok_koment_start + 1).clearContent();


    let indИзготовитель = 1;
    let indКомментарийЭксперта = 2;
    let indПодходит = 3;


    let vlsId = this.sheet.getRange(rowStart, this.col_list1_row, rowEnd - rowStart + 1, 1).getValues();
    // =====================  
    let j = -1;
    for (let row = rowStart; row <= rowEnd; row++) {
      /** todo */
      Logger.log(`onEdit_blok_koment Row=${row}=========================================`);
      j++;
      if (!vlsId[j][0]) {
        // Logger.log(`onEdit_blok_koment continue ${vlsId[j][0]}`); 
        continue;
      }

      // this.sheet.getRange(row, this.col_blok_koment_start, 1, this.col_blok_koment_end - this.col_blok_koment_start + 1).clearContent();


      let comentRow = parseInt(this.sheet.getRange(row, this.cols.Лист1НомерСтроки).getValue().toString());
      // let comentRow = row;
      Logger.log(`comentRow=${comentRow}`);

      // let commentMap = (isNaN(comentRow) ? new Map() : this.makeCommentMap(comentRow, this.sheet, this.col_com_experta));
      let comentBul = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,];           // массив если коментарий эксперта изменен 
      let comentArr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];    // массив введеных коментариев

      // let okMap = (isNaN(comentRow) ? new Map() : this.makeCommentMap(comentRow, this.sheet, this.col_ok_experta));
      let okBul = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,];           // массив если Подходит изменен 
      let okArr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];   // массив введенных Подходит  

      let countAgArr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];   // массив введенных поставщиков  // email


      let Изготовительbul = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,];           // массив если Изготовитель изменен 
      let Изготовительarr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];   // массив введенных Изготовитель  

      let Модельbul = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,];           // массив если Модель изменен 
      let Модельarr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];   // массив введенных Модель  





      for (let i = 0, col = range.columnStart; col <= range.columnEnd; col++, i++) {

        // Logger.log(`j=${j};  i=${i};  row=${row};  col=${col};  head=${rangeHeads[0][i]}; value=${rangeValues[j][i]};`);


        if (!((this.col_blok_koment_start <= col) && (col <= this.col_blok_koment_end))) {
          Logger.log(`onEdit_blok_koment col_blok_koment_start continue  ${!((this.col_blok_koment_start <= col) && (col <= this.col_blok_koment_end))}`);
          continue;
        }
        this.sheet.getRange(row, col, 1, 1).clearContent();
        let h = fl_str(rangeHeads[0][i]);
        let v = rangeValues[j][i];

        let hn = parseInt(h);
        if (isNaN(hn)) {
          Logger.log(`onEdit_blok_koment isNaN(hn) continue  ${isNaN(hn)}`);
          continue;
        }
        hn = hn - 1;
        Logger.log(`onEdit_blok_koment hn  ${hn}`);
        if (h.includes(fl_str(this.names.КомментарийЭксперта))) {
          comentBul[hn] = 1;
          comentArr[hn] = v;
        }

        if (h.includes(fl_str(this.names.Поставщик))) {
          countAgArr[hn] = v;
        }

        if (h.includes(fl_str(this.names.Подходит))) {
          okBul[hn] = 1;
          okArr[hn] = v;
        }

        if (h.includes(fl_str(this.names.Изготовитель))) {
          Изготовительbul[hn] = 1;
          Изготовительarr[hn] = v;
        }


        if (h.includes(fl_str(this.names.Модель))) {
          Модельbul[hn] = 1;
          Модельarr[hn] = v;
        }



        if (h.includes(fl_str(this.names.Формула))) {
          this.sheet.getRange(row, col).setFormula([this.getFormulaForCol(hn + 1)]);
        }

        Logger.log(`hn=${hn};  h=${h};   v=${v};  head=${rangeHeads[0][i]}; value=${rangeValues[j][i]};
       
        // if(Комментарий эксперта)=${h.includes(fl_str("Комментарий эксперта"))};  if(поставщ)=${h.includes(fl_str("поставщ"))}; if(Подходит)=${h.includes(fl_str("Подходит"))};`);

      }




      let topAg = this.getTopAg(row, this.sheet, this.col_blok_koment_start, this.col_blok_koment_end);      // массив email поставщиков на листе
      let topAl = this.getTopAg_link(row, this.sheet, this.col_blok_koment_start, this.col_blok_koment_end, this.col_top_join); // массив адресов яцеек поставщиков в топе link


      Logger.log(JSON.stringify({
        row,
        col_blok_koment_start: this.col_blok_koment_start,
        col_blok_koment_end: this.col_blok_koment_end,
        col_top_join: this.col_top_join,
        comentBul,
        comentArr,
        okBul,
        okArr,
        countAgArr,
        Изготовительbul,
        Изготовительarr,
        Модельbul,
        Модельarr,
        topAg,
        topAl,
      }, null, 2));




      for (let i = 0; i < comentArr.length; i++) {
        if (comentBul[i] != 1) { continue; }  // не меняли 
        let coment = comentArr[i]; // коментарий введенный
        let cAg = countAgArr[i];  // email введенный 
        let tAg = topAg[i];  // email
        let sAg = topAl[i];  // link
        if (!sAg) {
          if (cAg) { sAg = this.getLinkForEmail(cAg); }  // поиск link по email введенным пользователем 
          else if (tAg) { sAg = this.getLinkForEmail(tAg) } // поиск link по email до начало редоктирования 
        }
        if (!sAg) { continue; }  // ничего не нашли адрес яцейки постовщика -=беда=-  
        if (coment) { coment = `${coment}`.replace(/\|/g, '_') }
        this.addComment_separate(sAg, coment, comentRow, this.col_commentBlokStart + i, indКомментарийЭксперта);

      }



      for (let i = 0; i < okArr.length; i++) {
        if (okBul[i] != 1) { continue; }  // не меняли 

        let coment = okArr[i]; // коментарий введенный
        let cAg = countAgArr[i];  // email введенный 
        let tAg = topAg[i];  // email
        let sAg = topAl[i];  // link
        if (!sAg) {
          if (cAg) { sAg = this.getLinkForEmail(cAg); }  // поиск link по email введенным пользователем 
          else if (tAg) { sAg = this.getLinkForEmail(tAg) } // поиск link по email до начало редоктирования 
        }
        if (!sAg) { continue; }  // ничего не нашли адрес яцейки постовщика -=беда=-  
        if (coment) { coment = `${coment}`.replace(/\|/g, '_') }
        this.addComment_separate(sAg, coment, comentRow, this.col_commentBlokStart + i, indПодходит);
      }


      //  Изготовительbul
      //  Изготовительarr


      for (let i = 0; i < Изготовительarr.length; i++) {
        if (Изготовительbul[i] != 1) { continue; }  // не меняли 
        let coment = Изготовительarr[i]; // коментарий введенный
        let cAg = countAgArr[i];  // email введенный 
        let tAg = topAg[i];  // email
        let sAg = topAl[i];  // link
        if (!sAg) {
          if (cAg) { sAg = this.getLinkForEmail(cAg); }  // поиск link по email введенным пользователем 
          else if (tAg) { sAg = this.getLinkForEmail(tAg) } // поиск link по email до начало редоктирования 
        }
        if (!sAg) { continue; }  // ничего не нашли адрес яцейки постовщика -=беда=-  
        if (coment) { coment = `${coment}`.replace(/\|/g, '_') }
        this.addComment_separate(sAg, coment, comentRow, this.col_commentBlokStart + i, indИзготовитель);

      }


      for (let i = 0; i < Модельarr.length; i++) {
        if (Модельbul[i] != 1) { continue; }  // не меняли 
        let coment = Модельarr[i]; // коментарий введенный
        let cAg = countAgArr[i];  // email введенный 
        let tAg = topAg[i];  // email
        let sAg = topAl[i];  // link
        if (!sAg) {
          if (cAg) { sAg = this.getLinkForEmail(cAg); }  // поиск link по email введенным пользователем 
          else if (tAg) { sAg = this.getLinkForEmail(tAg) } // поиск link по email до начало редоктирования 
        }
        if (!sAg) { continue; }  // ничего не нашли адрес яцейки постовщика -=беда=-  
        if (coment) { coment = `${coment}`.replace(/\|/g, '_') }
        // Logger.log(` Модель Agent=  ${sAg} | row=${comentRow} Coment= ${coment}`);
        // this.addComment_separate(sAg, coment, comentRow, this.col_commentBlokStart + i, indИзготовитель);
        getContext().getSheetList1().setModelForContragent(sAg, comentRow, coment);
      }

    }

  }

  // getLinkForEmail(email, sheet = getContext().getSheetByName("Лист 1")) {
  getLinkForEmail(email, sheet = getContext().getSheetByName(getSettings().sheetName_Лист_1)) {

    let p = getClassColRow().list1_col_FirstPrice;
    let hp = getClassColRow().list1_col_LastPrice;

    let rangeEmails = sheet.getSheetValues(getClassColRow().list1_row_Header, p, 1, hp - p + 1)[0];

    for (let i = 0; i < rangeEmails.length; i++) { rangeEmails[i] = fl_str(rangeEmails[i]) }

    let col = rangeEmails.indexOf(fl_str(email));

    if (col == -1) { return ""; }  // нет такого у нас емайла 

    col = col + p;
    // console.log(`getLinkForEmail email=${email}  col=${col}  ret=$${nc(col)}$${1}  `);
    return `$${nc(col)}$${1}`;


  }




  getTopAg(row, sheet, colStart, colEnd) {   // поставцики их email 
    let retArr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];

    let rangeHeads = sheet.getSheetValues(this.rows.headr.key, colStart, 1, colEnd - colStart + 1);
    let rangeValues = sheet.getSheetValues(row, colStart, 1, colEnd - colStart + 1);

    // Logger.log("топ7 heads=" + rangeHeads);
    // Logger.log("топ7 value=" + rangeValues);

    for (let col = 0; col < rangeValues[0].length; col++) {
      let h = fl_str(rangeHeads[0][col]);
      // Logger.log("топ7 H=" + h);
      if (!h.includes(fl_str("поставщ"))) { continue; }
      let hn = parseInt(h);
      if (isNaN(hn)) { continue; }
      hn = hn - 1;
      retArr[hn] = rangeValues[0][col];
    }
    // Logger.log("топ7 retArr=" + retArr);
    return retArr;
  }




  getTopAg_link(row, sheet, colStart, colEnd, columnTopSplit) {   // поставцики  ссылка на поставцика в getSettings().sheetName_Лист_1

    let retArr = ["", "", "", "", "", "", "", "", "", "", "", "", "", "",];

    let rangeHeads = sheet.getSheetValues(this.rows.headr.key, colStart, 1, colEnd - colStart + 1)[0];
    let rangeValues = sheet.getSheetValues(row, columnTopSplit, 1, 1)[0][0].toString().split("|||");



    let length = Math.min(rangeValues.length * 3, rangeHeads.length);
    // let length = Math.min(rangeValues.length, rangeHeads.length);


    // Logger.log("топ7 heads=" + rangeHeads + `  ${rangeHeads.length}`);
    // Logger.log("топ7 value=" + rangeValues + `  ${rangeValues.length}`);


    for (let col = 0; col < length; col++) {
      let h = fl_str(rangeHeads[col]);
      if (!h.includes(fl_str("поставщ"))) { continue; }
      let hn = parseInt(h);
      if (isNaN(hn)) { continue; }
      hn = hn - 1;
      retArr[hn] = rangeValues[hn * 8 + 1];
    }
    // Logger.log("топ7 retArr=" + retArr);
    return retArr;
  }



  addComment_separate(sAg, coment, row, col, ind) {

    let commentMap_separate = undefined;
    try {
      // commentMap_separate = (isNaN(row) ? new Map() : this.makeCommentMap_separate(row, getContext().getSheetList1().sheet, col, ind));
      commentMap_separate = this.makeCommentMap_separate(row, getContext().getSheetList1().sheet, col);
    } catch (err) { }

    if (!commentMap_separate) { commentMap_separate = new Map(); }

    /** @type {[]} */
    let arr = commentMap_separate.get(fl_str(sAg));
    if (!Array.isArray(arr)) {
      Logger.log("нет коментария");
      Logger.log(sAg);
      Logger.log(arr);
      arr = [sAg, "", "", "", new Date().getTime()];
    }

    arr[ind] = coment;
    arr[4] = new Date().getTime();
    // commentMap_separate.set(fl_str(sAg), `${fl_str(sAg)}|${arr.join("|")}|${new Date().valueOf()}`);
    commentMap_separate.set(fl_str(sAg), arr);
    this.saveCommentMap_separate(commentMap_separate, row, getContext().getSheetList1().sheet, col, ind);
  }

  /** @param {SpreadsheetApp.Sheet} sheet */
  makeCommentMap_separate(row, sheet, col) {

    // console.log(`makeCommentMap_separate +++++++++++++++++`);
    // let joinStr = sheet.getRange(row, col).getValue();
    let ce = getClassColRow().list1_col_CommentExpertBlokEnd;
    let cs = getClassColRow().list1_col_CommentExpertBlokStart;

    let joinStrs = sheet.getRange(row, cs, 1, ce - cs + 1).getValues();
    let joinStr = joinStrs[0].filter(v => v).join("|||");
    let joinArr = new Array();
    let commentMap = new Map();
    // console.log(`log map-- joinStr = ${joinStr}`);
    if (joinStr) {
      joinArr = joinStr.split("|||");
    }

    Logger.log(JSON.stringify({ joinStr }, null, 2));
    Logger.log(JSON.stringify({ joinArr }, null, 2));

    for (let i = 0; i < joinArr.length; i++) {
      /** @type {[]} */
      let arr = joinArr[i].split("📎");
      if (commentMap.has(fl_str(arr[0]))) {
        let arrh = commentMap.get(fl_str(arr[0]));
        if (arr[4] < arrh[4]) { continue; }
      }
      // commentMap.set(fl_str(arr[0]), joinArr[i]);
      commentMap.set(fl_str(arr[0]), arr);
      // i = i + 1;
    }
    Logger.log("makeCommentMap_separate");
    commentMap.forEach((v, k) => { Logger.log(JSON.stringify({ k, v })) });

    return commentMap;
  }



  saveCommentMap_separate(commentMap, row, sheet, col) {
    Logger.log("saveCommentMap_separate");
    commentMap.forEach((v, k) => { Logger.log(JSON.stringify({ k, v })) });
    let joinArr = new Array();
    for (let [keyAg, vComm] of commentMap) {
      if (!keyAg) { continue; }
      joinArr.push(vComm.join("📎"));
    }
    Logger.log(JSON.stringify({ row, col }));
    sheet.getRange(row, col).setValue(joinArr.join("|||"));
    SpreadsheetApp.flush();
  }


}
























