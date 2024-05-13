// 20.02.2021 MisterVova   фактический переработай весь файл  сдесь функции из mrOnEdit_List_7 // mrOnEdit_List_7 модно удалять или за коментировать   
function onEditList_7(edit) {
  // return;
  getContext().getSheetList7().onEdit(edit);
}


function getReMailingIdProducts() {
  return getContext().getSheetList7().getReMailingIdProducts();
}



class ClassSheet_List7 {
  constructor() { // class constructor
    Logger.log("ClassSheet_List7 constructor Start");
    // this.sheet = getContext().getSheetByName("Лист7");
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Лист_7);
    this.sheetName = this.sheet.getSheetName();
    this.columnIdProduct = nr("A");
    this.columnNameProduct = nr("B");
    this.columnContragentEmail = nr("K"); // колонка с емайлом контрагента 
    this.columnContragentRezEmail = nr("L"); // колонка с емайлом контрагента Резерв
    this.columnStaus = nr("O"); // колонка со Статусом
    this.columnReMailing = undefined//  nr("J"); // Повторная рассылка
    // this.columnDeviation = 0;  
    this.column_Tsena_1 = nr("T");  // фактический номер колонки к воторой написанно "цена 1" в первой строки таблици  

    this.hesd = this.findColHead();  // функция переопрделит значения  this.columnNameProduct ,  this.columnContragentEmail  this.column_Tsena_1  если будут найдены ключевые слова;


    // this.columnDeviation = nr("T") - this.column_Tsena_1;
    // this.columnTopFirst = nr("T") - this.columnDeviation - 1;
    // this.columnTopLast = nr("BV") - this.columnDeviation;
    // this.columnNameDuplicate = nr("CB") - this.columnDeviation;  // колоннка с наименованиями  для проверки  columnNameDuplicate
    // this.columnTopSplit = nr("CC") - this.columnDeviation; // колоннка с топами columnTopStr
    // this.columnLineNumberSheet_1 = nr("CD") - this.columnDeviation;  // колоннка с номерами строк Лист1 // column with the line numbers of sheet 1 line numbers of sheet  columnLineNumberSheet_1


    this.columnTopFirst = this.columnБлокТопНачало + 1;
    this.columnTopLast = this.columnБлокТопКонец - 1;

    this.columnNameDuplicate = this.columnЛист1_Номер_строки - 2;  // колоннка с наименованиями  для проверки  columnNameDuplicate
    this.columnTopSplit = this.columnЛист1_Номер_строки - 1; // колоннка с топами columnTopStr
    this.columnLineNumberSheet_1 = this.columnЛист1_Номер_строки  // колоннка с номерами строк Лист1 // column with the line numbers of sheet 1 line numbers of sheet  columnLineNumberSheet_1





    this.col_I = getClassColRow().list7_col_I;



    // this.cellsNum = 4 * 2;  // число ячеек на один топ 
    this.cellsNum = 5 * 2;  // число ячеек на один топ 
    this.topNum = 7;  // число торов 

    this.backgroundsForTop = undefined;   // 4*7  = 28 клеток   7 цветов по четыре клктки  
    this.defaultBackgroundsForTop = undefined;  // 7 цветов 
    this.defaultBackgroundsForSamePrices = undefined // 7 цветов // 3 цвета из настроики писем 


    this.rowBodyFirst = 2;
    this.rowBodyLast = this.findRowBodyLast();


    this.summaryArr = undefined; //25.02.2021 mrVova
    this.columnGroup = nr("G");  //25.02.2021 mrVova


    this.sheetNameFrom = getSettings().sheetName_Лист_1;
    this.da = fl_str("да");


    Logger.log("ClassSheet_List7 constructor Finish");
  }


  getSummary(maloOtvetov, colorAnalog) { //25.02.2021 mrVova
    if (this.summaryArr) { return this.summaryArr; }
    let mapCountGroup = new Map();
    let arrCountProd = new Array();

    let vlsProd = this.sheet.getRange(this.rowBodyFirst, this.columnNameProduct, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues()
    let vlsGroup = this.sheet.getRange(this.rowBodyFirst, this.columnGroup, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues()

    for (let i = 0; i < vlsProd.length; i++) {
      if (!vlsProd[i][0]) { continue; }  //наименование пусто 
      let pr = vlsProd[i][0] = fl_str(vlsProd[i][0]);
      let gr = vlsGroup[i][0] = fl_str(vlsGroup[i][0]);
      if (!mapCountGroup.has(gr)) { mapCountGroup.set(gr, new Counter(gr)) }
    }


    var vlsPrice = this.sheet.getRange(this.rowBodyFirst, this.columnTopFirst, this.rowBodyLast - this.rowBodyFirst + 1, this.columnTopLast - this.columnTopFirst).getValues();
    var vlsBG = this.sheet.getRange(this.rowBodyFirst, this.columnTopFirst, this.rowBodyLast - this.rowBodyFirst + 1, this.columnTopLast - this.columnTopFirst).getBackgrounds();
    var vlsHead = vlsPrice;

    for (let i = 0; i < vlsProd.length; i++) {
      if (!vlsProd[i][0]) { continue; }  //наименование пусто 
      let pr = vlsProd[i][0];
      let gr = vlsGroup[i][0];

      let prCounter = new Counter(pr);
      let grCounter = mapCountGroup.get(gr);
      for (let j = 0; j < vlsPrice[i].length; j = j + 4) {
        if (!vlsHead[i][j + 1 * 2 + 1]) { continue; }  // нет контрагента пропускаем
        if (!vlsPrice[i][j + 0 * 2 + 1]) { continue; } // цена пусто пропускаем // нет значения 
        prCounter.increment(true, vlsBG[i][j + 0 * 2 + 1] == colorAnalog);
        grCounter.increment(true, vlsBG[i][j + 0 * 2 + 1] == colorAnalog);
      }
      arrCountProd.push(prCounter);
    }

    // Брать данные с листа 1
    let grCoun_1 = 0;// Кол-во групп у которых нет цены

    let prCoun_1 = 0// Кол-во товаров у которых нет цены


    for (let [gr, counter] of mapCountGroup) {
      if (counter.priceAll == 0) { grCoun_1++; }

    }

    // Logger.log(arrCountProd);

    for (let i = 0; i < arrCountProd.length; i++) {
      let counter = arrCountProd[i];
      if (counter.priceAll == 0) { prCoun_1++; }

    }


    this.summaryArr = new Array();
    this.summaryArr.push([`Всего групп ${getSettings().sheetName_Лист_7}`, mapCountGroup.size]);
    this.summaryArr.push([`Кол-во групп у которых нет цены`, grCoun_1]);

    this.summaryArr.push([`Всего товаров ${getSettings().sheetName_Лист_7}`, arrCountProd.length]);
    this.summaryArr.push([`Кол-во товаров у которых нет цены`, prCoun_1]);

    return this.summaryArr;

  }







  findColHead() {
    let vlsHead = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues();
    let strB = fl_str("Номенклатура");
    let strI = fl_str("Поставщик");
    let str_rez = fl_str("Резервный поставщик");
    let strQ = fl_str("1 цена");
    let str_re_mailing = fl_str("Повторная рассылка");
    let str_status = fl_str("Статус");
    let str_БлокТопНачало = fl_str("БлокТопНачало");
    let str_БлокТопКонец = fl_str("БлокТопКонец");
    let str_Лист1_Номер_строки = fl_str("Лист1 Номер строки");
    let str_ТопЭксперта = fl_str("ТопЭксперта");

    // let strBA = fl_str("1 цена");

    for (let i = 0; i < vlsHead[0].length; i++) {
      vlsHead[0][i] = `${fl_str(vlsHead[0][i])}`.trim();
      if (vlsHead[0][i] == strB) { this.columnNameProduct = i + 1; }
      if (vlsHead[0][i] == strI) { this.columnContragentEmail = i + 1; }
      if (vlsHead[0][i] == str_rez) { this.columnContragentRezEmail = i + 1; }
      if (vlsHead[0][i] == strQ) { this.column_Tsena_1 = i + 1; }
      if (vlsHead[0][i] == str_re_mailing) { this.columnReMailing = i + 1; }
      if (vlsHead[0][i] == str_status) { this.columnStaus = i + 1; }
      if (vlsHead[0][i] == str_БлокТопНачало) { this.columnБлокТопНачало = i + 1; }
      if (vlsHead[0][i] == str_БлокТопКонец) { this.columnБлокТопКонец = i + 1; }
      if (vlsHead[0][i] == str_Лист1_Номер_строки) { this.columnЛист1_Номер_строки = i + 1; }
      if (vlsHead[0][i] == str_ТопЭксперта) { this.columnТопЭксперта = i + 1; }
    }
    return vlsHead;
  }


  findRowBodyLast() {
    let ret = getContext().getRowSobachkaBySheetName(this.sheet.getSheetName());
    if (!ret) { ret = this.sheet.getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }

  updateRow(row) {
    this.updateRows(row, row);
  }

  updateAssociatedRowsWith_List1(rowStart, rowEnd) {
    let rangeAz = this.sheet.getRange(1, this.columnLineNumberSheet_1, this.sheet.getLastRow(), 1);
    let vlAz = rangeAz.getValues();


    // for (let i = 2; i < vlAz.length; i++) { // так было 
    for (let i = 1; i < vlAz.length; i++) { // так работает  во второй   строке
      if (!vlAz[i][0]) { continue; }
      if ((rowStart <= vlAz[i][0]) && (vlAz[i][0] <= rowEnd)) { this.updateRow(i + 1); }
    }

  }



  updateRows(rowStart, rowEnd) {
    // Logger.log(`ClassSheet_List7.updateRows rowStart=${rowStart}, rowEnd=${rowEnd}`);
    if (rowStart == 1) { rowStart = 2; }
    if ((rowEnd - rowStart + 1) <= 0) return;

    this.updateBackgroundRows(rowStart, rowEnd);

  }

  updateBackgroundProduct() {

    // let paintRange = this.sheet.getRange(2, nr("I"), this.sheet.getLastRow() - 2 + 1, nr("I") - nr("I") + 1);
    // //  paintRange.setBackgrounds(paintCounteragentEmail_List_BG(paintRange));
    // paintRange.setBackgrounds(this.paintRangeEmail(paintRange));

    let rowStart = 2;
    let lastRow = this.sheet.getLastRow();
    let ink = 50;
    for (; rowStart <= lastRow; rowStart = rowStart + ink) {
      let rowEnd = rowStart + ink;
      rowEnd = (rowEnd <= lastRow ? rowEnd : lastRow);
      this.topBackgroundClear(rowStart, rowEnd);
      this.topBackgroundIdenticalPrices(rowStart, rowEnd);
    }
  }

  updateBackgroundContragent() {

    let paintRange = this.sheet.getRange(2, this.columnContragentEmail, this.sheet.getLastRow() - 2 + 1, 1);
    //  paintRange.setBackgrounds(paintCounteragentEmail_List_BG(paintRange));
    paintRange.setBackgrounds(this.paintRangeEmail(paintRange));

    let rowStart = 2;
    let lastRow = this.sheet.getLastRow();
    let ink = 50;
    for (; rowStart <= lastRow; rowStart = rowStart + ink) {
      let rowEnd = rowStart + ink;
      rowEnd = (rowEnd <= lastRow ? rowEnd : lastRow);
      this.topBackgroundCounteragentColor(rowStart, rowEnd);
      this.topBackgroundCounteragentAnalog(rowStart, rowEnd);
    }
  }


  updateBackgroundAll() {

    let paintRange = this.sheet.getRange(2, this.columnContragentEmail, this.sheet.getLastRow() - 2 + 1, 1);
    //  paintRange.setBackgrounds(paintCounteragentEmail_List_BG(paintRange));
    paintRange.setBackgrounds(this.paintRangeEmail(paintRange));

    let rowStart = 2;
    let lastRow = this.sheet.getLastRow();
    let ink = 50;
    for (; rowStart <= lastRow; rowStart = rowStart + ink) {
      let rowEnd = rowStart + ink;
      rowEnd = (rowEnd <= lastRow ? rowEnd : lastRow);
      this.updateBackgroundRows(rowStart, rowEnd);
    }
  }

  getBackgroundsForTop() {
    let bgTop = this.getDefaultBackgroundsForTop();
    if (!this.backgroundsForTop) {
      this.backgroundsForTop = new Array(this.cellsNum * this.topNum);

      for (let i = 0; i < 7; i++) {
        this.backgroundsForTop[col + 0 * 2 + 1] = bgTop[i];
        this.backgroundsForTop[col + 1 * 2 + 1] = bgTop[i];
        this.backgroundsForTop[col + 2 * 2 + 1] = bgTop[i];
        this.backgroundsForTop[col + 3 * 2 + 1] = bgTop[i];

      }

    }

    return this.backgroundsForTop;
  }


  // цвет заливки топ по умолчанию как в шапке 
  topBackgroundClear(rowStart, rowEnd) {
    // Logger.log(`ClassSheet_List7.clearBackground() rowStart=${rowStart}, rowEnd=${rowEnd}`);
    // let paintRange = this.sheet.getRange(rowStart, this.firstColumnTop, rowEnd - rowStart + 1, this.lastColumnTop - this.firstColumnTop + 1);

    let bgTop = this.getDefaultBackgroundsForTop();
    let colStart = this.columnTopFirst;
    let rowNum = rowEnd - rowStart + 1;
    for (let i = 0; i < bgTop.length; i++) {
      let paintRange = this.sheet.getRange(rowStart, colStart, rowNum, this.cellsNum);
      paintRange.setBackground(bgTop[i]);
      colStart = colStart + this.cellsNum;
    }

    return;
  }

  // цвет заливки одинаковых цен  fill color of the same prices
  topBackgroundIdenticalPrices(rowStart, rowEnd) {
    // Logger.log(`ClassSheet_List7.topBackgroundIdenticalPrices rowStart=${rowStart}, rowEnd=${rowEnd}`);
    let paintRange = this.sheet.getRange(rowStart, this.columnTopFirst, rowEnd - rowStart + 1, this.columnTopLast - this.columnTopFirst + 1);
    let bgSam = this.getDefaultBackgroundsForSamePrices();

    let bg = paintRange.getBackgrounds();
    let vl = paintRange.getValues();

    let pr = 1;  //смещение позиции цены  относительно начана Ячеек относящихся к поставщику
    for (let row = 0; row < bg.length; row++) {

      let groupPrice = this.findGroupsOfIdenticalPrices_v2(vl[row]);
      // цвет заливки одинаковых цен  
      // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices_v2 groupPrice=${groupPrice}`);
      for (let col = 0, i = 0; (col < bg[row].length && i < bgSam.length); col = col + this.cellsNum, i++) {
        // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices groupPrice=${groupPrice}`);


        if (!vl[row][col + pr]) { continue; }
        if (groupPrice[i] == 0) { continue; }


        let ci = 0;
        for (; ci < this.cellsNum; ci++) { bg[row][col + ci] = bgSam[groupPrice[i] - 1]; }
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];  
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];
        // bg[row][col + ci++] = bgSam[groupPrice[i] - 1];


      }
    }

    paintRange.setBackgrounds(bg);
    return;
  }

  // цвет заливки поставщиков и коментариев
  topBackgroundCounteragentColor(rowStart, rowEnd) {

    // Logger.log(`ClassSheet_List7.topBackgroundCounteragentColor rowStart=${rowStart}, rowEnd=${rowEnd}`);

    let paintRange = this.sheet.getRange(rowStart, this.columnTopFirst, rowEnd - rowStart + 1, this.columnTopLast - this.columnTopFirst + 1);
    let bg = paintRange.getBackgrounds();
    let vl = paintRange.getValues();

    let colorDef = getContext().getColor("DEFAULT");
    let colorOfficial = getContext().getOfficialColor();
    let colorFromProduct = getContext().getIsFromProductListColor();


    for (let row = 0; row < bg.length; row++) {
      // цвет заливки поставщиков 
      let pp = 1 * 2 + 1;  //смещение позиции e-mail контрагента (поставщика)  относительно начана  группы Ячеек относящихся к поставщику
      let pc = 2 * 2 + 1;  //смещение позиции коментария контрагента (поставщика)  относительно начана  группы Ячеек относящихся к поставщику
      for (let col = 0; (col < bg[row].length); col = col + this.cellsNum) {



        // bg[row][col + pc] = "#AA00CC";  // для теста позиционирования
        // bg[row][col + pp] = "#FFCC00"; // для теста позиционирования
        // continue; // для теста позиционирования
        if (!vl[row][col + pp]) { continue; }

        let email = fl_str(vl[row][col + pp]);
        if (email) {
          let emails = `${email}`.split(" ", 2)
          if (emails.length > 0) { email = emails[0] }
        }
        // if (!isEmail(email)) { continue; }


        // let color = getContext().getCounteragentColor(fl_str(email));


        let counteragent = getContext().getCounteragent(email);
        // Logger.log(`ClassSheet_List7.updateBackgroundRows() email=${email},  counteragent=${counteragent}`);

        let colorKomment = getContext().getCounteragent(email).getColor();
        let colorCounteragent = colorDef;

        if (counteragent.isFromProductList()) { colorCounteragent = colorFromProduct; }
        if (counteragent.isOfficial()) { colorCounteragent = colorOfficial; }

        // let colorC = getContext().getCounteragent(email).getStatus().getColor();



        // Logger.log(`ClassSheet_List7.updateBackgroundRows() email=${email},  color=${color}`);
        // if (!color) { continue; }
        if (colorCounteragent != colorDef) { bg[row][col + pp] = colorCounteragent; }
        if (colorKomment != colorDef) { bg[row][col + pc] = colorKomment; }

      }

    }
    // Logger.log(`ClassSheet_List7.updateBackgroundRows() 2bg =${bg}`);
    paintRange.setBackgrounds(bg);
    return;
  }


  // цвет заливки поставщиков Аналог
  topBackgroundCounteragentAnalog(rowStart, rowEnd) {
    // return;
    // Logger.log(`ClassSheet_List7.topBackgroundCounteragentAnalog rowStart=${rowStart}, rowEnd=${rowEnd}`);

    let clSh_1 = new ClassSheet_List1();
    let analogArrLink = clSh_1.getAnalogArrLink(); // список адресов ячеек помеченные цаетом аналог
    if (analogArrLink.length == 0) { return; } // нечего красить 

    let paintRange = this.sheet.getRange(rowStart, this.columnTopFirst, rowEnd - rowStart + 1, this.columnTopLast - this.columnTopFirst + 1);
    let topSplitRange = this.sheet.getRange(rowStart, this.columnTopSplit, rowEnd - rowStart + 1, 1);

    let bg = paintRange.getBackgrounds();
    let sr = topSplitRange.getValues();

    let analogColor = getContext().getAnalogColor();

    let pPrice = -1 * 2 + 1;  //смещение позиции цены контрагента (поставщика)  относительно EMAIL  группы Ячеек относящихся к поставщику 
    for (let row = 0; row < bg.length; row++) {


      let sp = sr[row][0].toString().split("|||"); // 

      for (let i = 0; i < analogArrLink.length; i++) {
        let col = sp.indexOf(analogArrLink[i]);
        if (col == -1) { continue; }
        bg[row][col * 2 + pPrice] = analogColor;
      }

    }
    paintRange.setBackgrounds(bg);
    return;
  }


  topBackgroundCounteragentGoodExpert(rowStart, rowEnd) {
    Logger.log(`topBackgroundCounteragentGoodExpert start`);
    // f


    // this.columnТопЭксперта


    // let clSh_1 = new ClassSheet_List1();
    // let analogArrLink = clSh_1.getAnalogArrLink(); // список адресов ячеек помеченные цаетом аналог
    // if (analogArrLink.length == 0) { return; } // нечего красить 

    let paintRange = this.sheet.getRange(rowStart, this.columnTopFirst, rowEnd - rowStart + 1, this.columnTopLast - this.columnTopFirst + 1);
    let topSplitRange = this.sheet.getRange(rowStart, this.columnTopSplit, rowEnd - rowStart + 1, 1);
    let topExpertaRange = this.sheet.getRange(rowStart, this.columnТопЭксперта, rowEnd - rowStart + 1, 1);

    let bg = paintRange.getBackgrounds();
    let srM = topSplitRange.getValues();
    let srE = topExpertaRange.getValues();

    // Если подходит то эту ячейку подсвечивать зеленым
    // Если нет то красным

    let colorDa = getContext().getColor("ПодходитДа");
    let colorNet = getContext().getColor("ПодходитНет");

    let pPodh = 3 * 2 + 1;  //смещение позиции цены контрагента (поставщика)  относительно EMAIL  группы Ячеек относящихся к поставщику 
    for (let row = 0; row < bg.length; row++) {
      let spE = srE[row][0].toString().split("|||"); // 
      // Logger.log(`${spE}`);

      // let spM = srM[row][0].toString().split("|||"); // 


      // for (let i = 0; i < analogArrLink.length; i++) {
      //   // let col = spM.indexOf(analogArrLink[i]);
      //   if (col == -1) { continue; }
      //   bg[row][col * 2 + pPrice] = analogColor;
      // }

      // for (let col = pPodh; col < bg[row].length; col = col + pPodh) {
      for (let t = 0; t < 7; t++) {
        // if (true) { continue; }

        let col = t * 10 + 10 - 1;
        let i_sp = t * 8 + 7 - 1
        // Logger.log(`${nc(col + this.columnTopFirst)}   ${i_sp}`);
        // Logger.log(`${spE[i_sp]}`);

        if (fl_str(`${spE[i_sp]}`).includes(fl_str("Подходит")) || fl_str(`${spE[i_sp]}`).includes(fl_str("Да"))) {
          bg[row][col] = colorDa;
          continue;
        }

        if (fl_str(`${spE[i_sp]}`).includes(fl_str("Нет"))) {
          bg[row][col] = colorNet;
          continue;
        }

        // bg[row][col] = getContext().getColor("default");
      }

      // break;
    }
    paintRange.setBackgrounds(bg);
    return;

    // Logger.log(`topBackgroundCounteragentGoodExpert start`);
  }


  updateBackgroundRows(rowStart, rowEnd) {
    Logger.log(`ClassSheet_List7.updateBackgroundRows() rowStart=${rowStart}, rowEnd=${rowEnd}`);
    if (rowStart == 1) { rowStart = 2; }
    if ((rowEnd - rowStart + 1) <= 0) return;

    rowEnd = (rowEnd <= this.rowBodyLast ? rowEnd : this.rowBodyLast);
    if ((rowEnd - rowStart + 1) <= 0) return;


    // *
    this.topBackgroundClear(rowStart, rowEnd);
    this.topBackgroundIdenticalPrices(rowStart, rowEnd);
    this.topBackgroundCounteragentColor(rowStart, rowEnd);  ///  &&
    this.topBackgroundCounteragentAnalog(rowStart, rowEnd);
    this.topBackgroundCounteragentGoodExpert(rowStart, rowEnd);
    return;

  }



  updateBackgroundRows_v1(rowStart, rowEnd) { // медленно 
    // Logger.log(`ClassSheet_List7.updateBackgroundRows()_v1  rowStart=${rowStart}, rowEnd=${rowEnd}`);

    let paintRange = this.sheet.getRange(rowStart, this.columnTopFirst, rowEnd - rowStart + 1, this.columnTopLast - this.columnTopFirst + 1);
    let bgTop = this.getDefaultBackgroundsForTop();
    let bgSam = this.getDefaultBackgroundsForSamePrices();

    let bg = paintRange.getBackgrounds();
    let vl = paintRange.getValues();
    let analogColor = getContext().getAnalogColor();
    // Logger.log(`ClassSheet_List7.updateBackgroundRows() analogColor =${analogColor}`);

    // Logger.log(`ClassSheet_List7.updateBackgroundRows() 1bg=${bg}`);
    // Logger.log(`ClassSheet_List7.updateBackgroundRows() vl=${vl}`);
    for (let row = 0; row < bg.length; row++) {
      let price = new Array();


      // цвет заливки топ по умолчанию как в шапке 
      for (let col = 0, i = 0; (col < bg[row].length && i < bgTop.length); col = col + this.cellsNum, i++) {
        bg[row][col + 0 * 2 + 1] = bgTop[i];
        bg[row][col + 1 * 2 + 1] = bgTop[i];
        bg[row][col + 2 * 2 + 1] = bgTop[i];
        bg[row][col + 3 * 2 + 1] = bgTop[i];
        price.push(vl[row][col + 0]);
        // Logger.log(`ClassSheet_List7.updateBackgroundRows() row=${row}, col=${col}, i=${i}`);
      }


      let groupPrice = this.findGroupsOfIdenticalPrices(price);
      // цвет заливки одинаковых цен  
      for (let col = 0, i = 0; (col < bg[row].length && i < bgSam.length); col = col + this.cellsNum, i++) {
        // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices groupPrice=${groupPrice}`);
        if (!vl[row][col + 0]) { continue; }
        if (groupPrice[i] == 0) { continue; }
        bg[row][col + 0 * 2 + 1] = bgSam[groupPrice[i] - 1];
        bg[row][col + 1 * 2 + 1] = bgSam[groupPrice[i] - 1];
        bg[row][col + 2 * 2 + 1] = bgSam[groupPrice[i] - 1];
        bg[row][col + 3 * 2 + 1] = bgSam[groupPrice[i] - 1];
      }

      // цвет заливки поставщиков 
      let pp = (1 + 1) * 2;  //смещение позиции e-mail контрагента (поставщика)  относительно начана Ячеек относящихся к поставщику
      let emailArr = new Array();
      for (let col = 0; (col < bg[row].length); col = col + this.cellsNum) {
        let email = fl_str(vl[row][col + pp]);
        emailArr.push(email);
        // if (!isEmail(email)) { continue; }
        let color = getContext().getCounteragentColor(fl_str(email));
        // Logger.log(`ClassSheet_List7.updateBackgroundRows() email=${email},  color=${color}`);
        if (!color) { continue; }
        if (color == getContext().getColor("DEFAULT")) { continue; }
        bg[row][col + pp] = color;
      }


      // цвет заливки поставщиков Аналог
      let clSh_1 = new ClassSheet_List1();
      let analogMap = clSh_1.getAnalogMap();

      if (analogMap.size > 0) {
        for (let col = 0, i = 0; (col < bg[row].length && i < bgTop.length); col = col + this.cellsNum, i++) {
          let email = fl_str(vl[row][col + pp]);
          if (!analogMap.has(email)) { continue; }
          // Logger.log(`ClassSheet_List7.updateBackgroundRows() email =${email}   is analog!!!!`);
          // analogColor
          let analog = analogMap.get(email);
          let rowInfo = rowStart + row - 1;
          for (let [key, info] of analog.map) {
            try {
              // Logger.log(`vl=${vl[row][col + 0]}, vl=${vl[row][col + 3]},  vl=${vl[row][col + 1]}`);
              // Logger.log(`info=${info[rowInfo][0]},   info=${info[rowInfo][1]}`);

              if (info[rowInfo][0] != vl[row][col + 0]) { continue; }
              if (info[rowInfo][1] != vl[row][col + 3]) { continue; }
              bg[row][col + pp] = analogColor;
            }
            catch (err) {
              // Logger.log(`ERROR key =${key} == info${info}`);
              continue;
            }
          }
        }
      }
    }
    // Logger.log(`ClassSheet_List7.updateBackgroundRows() 2bg =${bg}`);
    paintRange.setBackgrounds(bg);



  }



  /* 
    0 - нет совпадений
    1 - Первое совпадение
    2 - Второе совпадение
    3 - Третье совпадение
  */

  findGroupsOfIdenticalPrices_v2(price) {
    let grId = 0; //0 - нет совпадений; 1 - Первое совпадение; 2 - Второе совпадени; 3 - Третье совпадение; и т.д. ; 

    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices price=${price}`);

    let pr = 1;  //смещение позиции цены  относительно начана Ячеек относящихся к поставщику
    let ret = new Array();
    let grMap = new Map();
    for (let i = 0; i < this.topNum; i++) {
      let ii = i * this.cellsNum + pr;

      ret.push(grId); //0 - нет совпадений
      if (!price[ii]) { continue; }

      let priceKey = price[ii].toString();
      let ar = grMap.get(priceKey);

      if (!ar) { ar = new Array(); }

      ar.push(i);
      grMap.set(priceKey, ar);
    }
    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices price=${price}`);


    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices grMap=${grMap}`);

    for (var [priceKey, ar] of grMap) {
      // Logger.log(`key=${priceKey}, ar=${ar}`);
      if (ar.length < 2) { continue; }
      grId++;
      for (let i = 0; i < ar.length; i++) {
        ret[ar[i]] = grId;
      }
    }

    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices ret=${ret}`);


    return ret;

  }





  findGroupsOfIdenticalPrices(price) {
    let grId = 0; //0 - нет совпадений; 1 - Первое совпадение; 2 - Второе совпадени; 3 - Третье совпадение; и т.д. ; 

    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices price=${price}`);


    let ret = new Array();
    let grMap = new Map();
    for (let i = 0; i < price.length; i++) {
      ret.push(grId); //0 - нет совпадений
      if (!price[i]) { continue; }
      price[i] = price[i].toString();
      // price[i] = Number(price[i]);
      // if (!price[i]) { continue; }


      let ar = grMap.get(price[i]);

      if (!ar) { ar = new Array(); }

      ar.push(i);
      grMap.set(price[i], ar);
    }
    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices price=${price}`);


    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices grMap=${grMap}`);

    for (var [key, ar] of grMap) {
      // Logger.log(`key=${key}, ar=${ar}`);
      if (ar.length < 2) { continue; }
      grId++;
      for (let i = 0; i < ar.length; i++) {
        ret[ar[i]] = grId;
      }
    }

    // Logger.log(`ClassSheet_List7.findGroupsOfIdenticalPrices ret=${ret}`);


    return ret;

  }


  getDefaultBackgroundsForTop() {

    if (!this.defaultBackgroundsForTop) {
      this.defaultBackgroundsForTop = new Array();
      this.defaultBackgroundsForTop.push(getContext().getColor("top_1"));
      this.defaultBackgroundsForTop.push(getContext().getColor("top_2"));
      this.defaultBackgroundsForTop.push(getContext().getColor("top_3"));
      this.defaultBackgroundsForTop.push(getContext().getColor("top_4"));
      this.defaultBackgroundsForTop.push(getContext().getColor("top_5"));
      this.defaultBackgroundsForTop.push(getContext().getColor("top_6"));
      this.defaultBackgroundsForTop.push(getContext().getColor("top_7"));
    }
    return this.defaultBackgroundsForTop;
  }


  getDefaultBackgroundsForSamePrices() {

    if (!this.defaultBackgroundsForSamePrices) {
      this.defaultBackgroundsForSamePrices = new Array();
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("Первое совпадение"));
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("Второе совпадение"));
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("Третье совпадение"));
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("top_4"));
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("top_5"));
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("top_6"));
      this.defaultBackgroundsForSamePrices.push(getContext().getColor("top_7"));
    }
    return this.defaultBackgroundsForSamePrices;
  }

  paintRangeEmail(paintRange) {

    let analogColor = getContext().getAnalogColor();


    let bg = paintRange.getBackgrounds();
    let vl = paintRange.getValues();

    for (let row = 0; row < bg.length; row++) {
      // for (let col = 0; col < bg[row].length; col++) {
      for (let col = 1; col < bg[row].length; col = col + 2) {

        if (!vl[row][col]) { continue; }
        if (bg[row][col] == analogColor) { continue; }

        let email = fl_str(vl[row][col]);

        if (!isEmail(email)) { continue }

        let color = getContext().getCounteragentColor(fl_str(email));
        if (!color) continue;

        bg[row][col] = color;
      }
    }

    return bg;
  }

  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
    // return
    Logger.log("ClassSheet_List7 fixSheet");
    if (this.rowBodyFirst > this.rowBodyLast) { return; }

    let cols = [
      this.columnLineNumberSheet_1,
      this.columnTopSplit,
      this.columnNameDuplicate,
    ];

    // let isMmodifiedFormula = false;

    for (let i = 0; i < cols.length; i++) {
      // let col = this.columnLineNumberSheet_1;
      let col = cols[i];
      let formula = this.makeFormula_List7(this.rowBodyFirst, col);
      let range = this.sheet.getRange(this.rowBodyFirst, col, this.rowBodyLast - this.rowBodyFirst + 1, 1);
      let formulas = range.getFormulasR1C1();
      let isMmodified = false;

      formulas = formulas.flat();
      // formulas.forEach(f => { Logger.log(`${f != formula} | \n|||${f}||| \n|||${formula}|||`); });
      formulas.forEach(f => { if (f != formula) { isMmodified = true } });
      if (isMmodified) {
        range.setFormulaR1C1(formula);
        // isMmodifiedFormula = true;
      }
    }


    let rangeHeads = this.sheet.getSheetValues(1, this.columnTopFirst, 1, this.columnTopLast - this.columnTopFirst + 1);

    for (let i = 0, col = this.columnTopFirst; col <= this.columnTopLast; col++, i++) {
      if (!((this.columnTopFirst <= col) && (col <= this.columnTopLast))) { continue; }
      let h = fl_str(rangeHeads[0][i]);
      if (!h.includes(fl_str("формула"))) { continue; }
      let hn = parseInt(h);
      if (isNaN(hn)) { continue; }
      hn = hn - 1;

      let formula = this.getFormulaForCol(hn + 1);
      let range = this.sheet.getRange(this.rowBodyFirst, col, this.rowBodyLast - this.rowBodyFirst + 1, 1);
      let formulas = range.getFormulasR1C1();
      let isMmodified = false;

      formulas = formulas.flat();
      // formulas.forEach(f => { Logger.log(`${f.slice(1) != formula} | \n|||${f.slice(1)}||| \n|||${formula}|||`); });
      // formulas.forEach(f => { Logger.log(`${f != formula} | \n|||${f}||| \n|||${formula}|||`); });
      formulas.forEach(f => { if (f.slice(1) != formula) { isMmodified = true } });
      // formulas.forEach(f => { if (f != formula) { isMmodified = true } });
      if (isMmodified) {

        // Logger.log("isMmodified = true")
        range.setFormulaR1C1(formula);

      }
      if (getContext().timeConstruct.getHours() != 3) { continue; }
      if (!this.sheet.isColumnHiddenByUser(col)) { this.sheet.hideColumn(range); }
    }

    this.fixSheetCompact();

    // this.tempupdateНастройкиAbkmnhjd();
    this.fixSheetCompactHideColumn();
  }


  fixSheetCompact() {


    // if (getContext().timeConstruct.getHours() != 2) { return; }
    let columnLast = this.sheet.getMaxColumns();
    this.sheet.hideColumn(this.sheet.getRange(1, this.columnNameDuplicate, 1, columnLast - this.columnNameDuplicate + 1))


    let ww = new Array(10);
    for (let i = 0; i < 10; i++) {
      ww[i] = this.sheet.getColumnWidth(this.columnTopFirst + i);
    }

    let f = (this.columnTopLast - this.columnTopFirst) + 2;
    for (let c = 0; c < f; c++) {
      let col = c + this.columnTopFirst;
      let w = ww[0];
      let i = c % 10;
      w = ww[i];
      this.sheet.setColumnWidth(col, w)
    }
  }

  fixSheetCompactHideColumn() {
    try {


      if (this.fl_fixSheetCompactHideColumn) { return; }
      if (this.columnБлокТопНачало != nr("Z")) { return; }
      let r = this.columnБлокТопНачало;
      this.sheet.hideColumns(r++);
      for (; r <= this.columnБлокТопКонец; r = r + 2) {
        this.sheet.hideColumns(r);
      }
      this.fl_fixSheetCompactHideColumn = true;
    } catch (err) { mrErrToString(err); }
  }


  insertComment(comment, tovar, contragentEmail, price, id) {
    // Logger.log(`ClassSheet_List7 insertComment | START `);
    Logger.log(`ClassSheet_List7 insertComment | START insertComment(comment, tovar, contragentEmail, price, id) =${JSON.stringify([comment, tovar, contragentEmail, price, id])}`);
    // let headsValues = this.sheet.getRange(1, 1, 1, this.columnTopLast).getValues();
    let bodyFormula = this.sheet.getRange(this.rowBodyFirst, 1, this.rowBodyLast - this.rowBodyFirst + 1, this.columnTopLast).getValues();
    bodyFormula = bodyFormula.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    // headsValues = headsValues.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });


    bodyFormula = bodyFormula.filter((v, i, arr) => { return v[this.columnNameProduct] == tovar });
    if (bodyFormula.length > 1) {
      bodyFormula = bodyFormula.filter((v, i, arr) => { return v[this.columnIdProduct] == id });
    }

    if (bodyFormula.length > 1) {
      bodyFormula = bodyFormula.filter((v, i, arr) => { return v[this.columnContragentEmail] == contragentEmail });
    }

    if (bodyFormula.length == 0) {
      Logger.log(`ClassSheet_List7 insertComment | FINISH | не нашли строку куда встаить  bodyFormula.length == 0 | return flse`);
      return false;
    }



    Logger.log(`ClassSheet_List7 insertComment | bodyFormula.length=${bodyFormula.length}`);
    // Logger.log(`ClassSheet_List7 insertComment | bodyFormula=${bodyFormula}`);
    let flag_insert = false;
    for (let r = 0; r < bodyFormula.length; r++) {
      let bf = bodyFormula[r];
      let row = bf[0];
      Logger.log(`ClassSheet_List7 insertComment | bodyFormula[${r}]=${JSON.stringify(bf)}`);

      let col_bf = bf.map((v, i, arr) => { return [i].concat(v); });
      // Logger.log(`ClassSheet_List7 insertComment | col_bf=${JSON.stringify(col_bf)}`);
      col_bf = col_bf.slice(this.columnTopFirst);
      // Logger.log(`ClassSheet_List7 insertComment | col_bf=${JSON.stringify(col_bf)}`);
      col_bf = col_bf.filter((v, i, arr) => { return v[1] == contragentEmail });
      Logger.log(`ClassSheet_List7 insertComment | col_bf=${JSON.stringify(col_bf)}`);
      if (col_bf.length == 0) {
        Logger.log(`ClassSheet_List7 insertComment| FINISH | не нашли столбец куда встаить  | col_bf.length == 0 |  return`);
        return false;
      }


      for (let i = 0; i < col_bf.length; i++) {
        let col_price = col_bf[0][0] - 2;  // 
        let col_comment = col_price + 2 * 2  // 

        // Logger.log(`ClassSheet_List7 insertComment| ADD | this.sheet.getRange(row=${row}, col_comment=${col_comment}).setValue(${comment})`);


        Logger.log(`ClassSheet_List7 insertComment|  цены | ${bf[col_price] == price} | ${col_price} | bf[col_price]= ${bf[col_price]} |  price=${price} `);
        if (bf[col_price] != price) {
          Logger.log(`ClassSheet_List7 insertComment| не совпали цены | Continue`);
          continue;
        }

        Logger.log(`ClassSheet_List7 insertComment| ADD | this.sheet.getRange(row=${row}, col_comment=${col_comment}).setValue(${comment})`);
        this.sheet.getRange(row, col_comment).setValue(comment);
        flag_insert = true;
        // break
      }
    }
    if (!flag_insert) {
      Logger.log(`ClassSheet_List7 insertComment | FINISH | не удалось добавить | !flag_insert | return`);
      return false;
    }
    return true;
  }





  onEditHelper(duration = 1 / 24 / 60 * 4.9) {

    // let colStart = this.columnTopFirst;

    let headsValues = this.sheet.getRange(1, this.columnTopFirst, 1, this.columnTopLast - this.columnTopFirst + 1).getValues();
    let bodyFormula = this.sheet.getRange(this.rowBodyFirst, this.columnTopFirst, this.rowBodyLast - this.rowBodyFirst + 1, this.columnTopLast - this.columnTopFirst + 1).getFormulas();

    bodyFormula = bodyFormula.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    headsValues = headsValues.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });


    let strformula = fl_str("формула")
    // Logger.log(` проверка формул этап 1`);
    for (let i = 0; i < headsValues[0].length; i++) {
      if (!fl_str(headsValues[0][i]).includes(strformula)) { continue; }
      let bf = bodyFormula.map(f => { return f[i] });

      if (bf.filter(f => f == "").length != 0) {
        let problerBf = bodyFormula.filter(f => f[i] == "");
        Logger.log(`ClassSheet_List7 onEditHelper Сол= ${nc(i + this.columnTopFirst - 1)} = ${problerBf.map(f => f[0])}`);

        let col = i + this.columnTopFirst - 1;
        for (let p = 0; p < problerBf.length; p++) {
          if (!getContext().hasTime(duration)) { break; }
          this.onEdit(this.getNewEdit(problerBf[p][0], col));
          // вызиваем от едит для кажндой проблемной формулы

        }
      }
    }

    let bodyValues = this.sheet.getRange(this.rowBodyFirst, this.columnTopFirst, this.rowBodyLast - this.rowBodyFirst + 1, this.columnTopLast - this.columnTopFirst + 1).getValues();
    bodyValues = bodyValues.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });
    // Logger.log(` проверка формул этап 2`);
    for (let i = 0; i < headsValues[0].length; i++) {



      if (!fl_str(headsValues[0][i]).includes(strformula)) { continue; }
      let bv = bodyValues.map(f => { return f[i] });
      if (bv.filter(f => f != "").length != 0) {
        let problerBv = bodyValues.filter(f => f[i] != "");
        Logger.log(`ClassSheet_List7 onEditHelpe Сол= ${nc(i + this.columnTopFirst - 1)} = ${problerBv.map(f => f[0])}`);

        let col = i + this.columnTopFirst - 1 + 1;
        for (let p = 0; p < problerBv.length; p++) {
          if (!getContext().hasTime(duration)) { break; }
          this.onEdit(this.getNewEdit(problerBv[p][0], col));
          // вызиваем от едит для кажндой проблемной формулы
        }
      }
    }



  }


  // function onEditList_7(edit) {
  onEdit(edit) {

    // Получаем диапазон ячеек, в которых произошли изменения
    let range = edit.range;

    // Лист, на котором производились изменения
    let sheet = range.getSheet();

    // Проверяем, нужный ли это нам лист 
    Logger.log("onEditList_7 начало");
    // if (sheet.getName() != 'Лист7') {
    if (sheet.getName() != getSettings().sheetName_Лист_7) {
      return false;
    }


    // if ((range.columnStart <= this.columnTopFirst - 1) && (this.columnTopFirst - 1 <= range.columnEnd)) {
    //   this.onEdit_L7_v4_f(edit);
    // }

    if ((range.columnStart <= this.columnTopLast) && (this.columnTopFirst <= range.columnEnd)) {
      // this.onEdit_L7_v3(edit);
      this.onEdit_L7_v4(edit);
    }

    if ((range.columnStart <= this.columnTopLast) && (this.columnTopFirst <= range.columnEnd)) {
      this.updateRows(range.rowStart, range.rowEnd);
    } else if ((range.columnStart <= this.columnNameProduct) && (this.columnNameProduct <= range.columnEnd)) {
      this.updateRows(range.rowStart, range.rowEnd);
    }

    if ((range.columnStart <= this.columnLineNumberSheet_1) && (this.columnNameDuplicate <= range.columnEnd)) {
      this.onEdit_L7_AX_AZ(edit);
    }


    if ((range.columnStart <= this.col_I) && (this.col_I <= range.columnEnd)) {
      this.onEdit_col_Expertise(edit);
    }

    Logger.log("onEditList_7_конец");
  }

  onEdit_col_Expertise(edit) {
    Logger.log("onEdit_col_Экспертиза edit=" + JSON.stringify(edit));
    let range = edit.range;
    let rowStart = (range.rowStart < this.rowBodyFirst ? this.rowBodyFirst : range.rowStart);
    let rowEnd = range.rowEnd;
    rowEnd = (rowEnd <= this.rowBodyLast ? rowEnd : this.rowBodyLast);
    if ((rowEnd - rowStart + 1) <= 0) return;



    let vlsI = this.sheet.getRange(rowStart, this.col_I, rowEnd - rowStart + 1, 1).getValues();
    // let vlsRow = this.sheet.getRange(rowStart, this.columnLineNumberSheet_1, rowEnd - rowStart + 1, 1).getValues();
    let vlsRow = this.sheet.getRange(rowStart, this.columnLineNumberSheet_1, rowEnd - rowStart + 1, 1).getValues();



    let rowArr = new Array();

    for (let i = 0; i < vlsI.length; i++) {
      if (fl_str(vlsI[i][0]) != this.da) { continue; }
      rowArr.push(vlsRow[i][0]);
    }

    if (rowArr.length > 0) {
      // getContext().getSheetExpertise().addRows(rowArr);
      getContext().addRowsForAllSheetExpertise(rowArr);
    }

  }


  // function onEdit_L7_AX_AZ(edit) {
  onEdit_L7_AX_AZ(edit) {
    Logger.log("onEdit_L7_AX_AZ edit= " + JSON.stringify(edit));
    let range = edit.range;
    let sheet_L7 = edit.range.getSheet();

    let rowStart = (range.rowStart < this.rowBodyFirst ? this.rowBodyFirst : range.rowStart);
    let rowEnd = range.rowEnd;
    rowEnd = (rowEnd <= this.rowBodyLast ? rowEnd : this.rowBodyLast);
    if ((rowEnd - rowStart + 1) <= 0) return;


    for (let row = rowStart; row <= rowEnd; row++) {

      let col = this.columnLineNumberSheet_1;  // порядок важен
      sheet_L7.getRange(row, col).setFormulaR1C1(this.makeFormula_List7(row, col));

      col = this.columnTopSplit;  // порядок важен
      sheet_L7.getRange(row, col).setFormulaR1C1(this.makeFormula_List7(row, col));

      col = this.columnNameDuplicate  // порядок важен
      sheet_L7.getRange(row, col).setFormulaR1C1(this.makeFormula_List7(row, col));

      this.sheet.getRange(row, this.columnTopFirst, 1, this.columnTopLast - this.columnTopFirst + 1).setFormulas([this.getFormulasForBlokComment()]);

    }


  }

  getFormulasForBlokComment() {
    let retArr = new Array();

    for (let i = 0; i < this.cellsNum * this.topNum / 2; i++) {
      retArr.push(this.getFormulaForCol(i + 1));
      retArr.push(undefined);
    }
    return retArr;
  }

  getFormulaForCol(number) {
    return this.getFormulaForColОктябрь2023(number);
    // if (isTest()) { return this.getFormulaForColОктябрь2023(number); }

    if (typeof number != "number") { throw Error("не является числом " + number); }
    if (number < 1 || this.cellsNum / 2 * this.topNum < number) { throw Error("нет формулы с порядковым номером " + number); }

    let x = number % (this.cellsNum / 2);
    // последняя в группе
    if (x == 2) {
      // if (x == 1 || x == 2 || x == 4) {
      // if (x == 1 || x == 2 || x == 0) {
      return `{"" \\ IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number})));"")}`;
    }
    if (x == 1 || x == 4) {
      return `{""  \\ LAMBDA(пост;ном;мем;знач;
  IFERROR( INDEX(QUERY(мем;"select * where Col1='"&пост&"' and Col2='"&ном&"'");1;${(x == 1 ? 3 : 4)});знач)
  )

  (
  IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number + (x == 1 ? 1 : -2)})));"");
  '3-3 Выбор поставщиков'!R[0]C2;
  SORT({ 'mem'!R3C2:C2 \\ 'mem'!R3C3:C3 \\ 'mem'!R3C9:C9 \\ 'mem'!R3C10:C10 \\ 'mem'!R3C1:C1 };5;1;1;1);
  IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number})));"")
  )
}`;

      return `{"" \\ IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number})));"")}`;
    }


    return `{"" \\ IFERROR(INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number});"")}`;



  }

  getFormulaForColОктябрь2023(number) {
    if (typeof number != "number") { throw Error("не является числом " + number); }
    if (number < 1 || this.cellsNum / 2 * this.topNum < number) { throw Error("нет формулы с порядковым номером " + number); }

    let x = number % (this.cellsNum / 2);
    // последняя в группе
    if (x == 2) {
      // if (x == 1 || x == 2 || x == 4) {
      // if (x == 1 || x == 2 || x == 0) {
      return `{"" \\ IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number})));"")}`;
    }
    if (x == 1 || x == 4) {

      return `{"" \\ IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number})));"")}`;
    }


    return `{"" \\ IFERROR(INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${number});"")}`;


  }

  makeFormula_List7(row, col) {

    if (col == this.columnLineNumberSheet_1) { return `=ROW(INDIRECT(MIDB(FORMULATEXT(INDIRECT(ADDRESS(ROW();${this.columnNameProduct};1;TRUE)));2;LENB(FORMULATEXT(INDIRECT(ADDRESS(ROW();${this.columnNameProduct};1;TRUE)))))))`; }
    let ct = `C${getClassColRow().list1_col_TopFormula}`;
    if (col == this.columnTopSplit) { return `=INDIRECT(ADDRESS(R[0]C${this.columnLineNumberSheet_1};COLUMN('${this.sheetNameFrom}'!R1${ct});1;TRUE;"${this.sheetNameFrom}"))`; }



    let cp = `C${getClassColRow().list1_col_Product}`;
    if (col == this.columnNameDuplicate) { return `=INDIRECT(ADDRESS(R[0]C${this.columnLineNumberSheet_1};COLUMN('${this.sheetNameFrom}'!R1${cp});1;TRUE;"${this.sheetNameFrom}"))`; }
  }


  makeFormulaR1C1_V3D(r, c) {

    return `={IFERROR(   INDIRECT( CONCATENATE("'${this.sheetNameFrom}'!";   INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C${this.columnTopFirst})+1)));"")\\IFERROR(INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C${this.columnTopFirst})+2);"")}`;
  }


  makeFormulaR1C1_V3(r, c) {

    return `=IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C${this.columnTopFirst})+1)));"")`;

  }


  makeFormulaR1C1_V4_4(t) {
    let i = 1;
    let formulaP = `IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${t * 4 + i++})));"")`;
    let formulaA = `IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${t * 4 + i++})));"")`;
    let formulaC = `IFERROR(INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${t * 4 + i++});"")`;
    let formulaT = `IFERROR( INDIRECT(  CONCATENATE("'${this.sheetNameFrom}'!";  INDEX(SPLIT(R[0]C${this.columnTopSplit};"|||";FALSE;FALSE);1;${t * 4 + i++})));"")`;

    let f4 = `{

    ${formulaP}\\
    ${formulaA}\\
    ${formulaC}\\
    ${formulaT}

    }`;



    return f4;


  }

  makeFormulaR1C1_V4(r, c) {
    let t = 0;




    let f4_7 = `{ "Формула"\\
       ${this.makeFormulaR1C1_V4_4(t++)}\\
      ${this.makeFormulaR1C1_V4_4(t++)}\\
      ${this.makeFormulaR1C1_V4_4(t++)}\\
      ${this.makeFormulaR1C1_V4_4(t++)}\\
      ${this.makeFormulaR1C1_V4_4(t++)}\\
      ${this.makeFormulaR1C1_V4_4(t++)}\\
      ${this.makeFormulaR1C1_V4_4(t++)}
      
    }`

    return f4_7;

  }



  onEdit_L7_v4_f(edit) {
    Logger.log("onEdit_L7_v4_f edit= " + JSON.stringify(edit));
    let range = edit.range;
    let sheet_L7 = edit.range.getSheet();

    let rowStart = (range.rowStart < this.rowBodyFirst ? this.rowBodyFirst : range.rowStart);
    let rowEnd = range.rowEnd;
    rowEnd = (rowEnd <= this.rowBodyLast ? rowEnd : this.rowBodyLast);
    if ((rowEnd - rowStart + 1) <= 0) return;


    // sheet_L7.getRange(rowStart, this.columnTopFirst - 1, rowEnd - rowStart + 1, 1).setBackground("#CCCCCC");
    sheet_L7.getRange(rowStart, this.columnTopFirst - 1, rowEnd - rowStart + 1, 1).setFormulaR1C1(this.makeFormulaR1C1_V4());

    // for (let row = rowStart; row <= rowEnd; row++) {
    //     sheet_L7.getRange(row, this.columnTopFirst - 1).setBackground("#CCCCCC");
    // }
  }

  // function onEdit_L7_v3(edit) {
  onEdit_L7_v4(edit) {

    Logger.log("onEdit_L7_v4 edit= " + JSON.stringify(edit));
    let range = edit.range;
    let sheet_L7 = edit.range.getSheet();
    // let sheet_L1 = getContext().getSheetByName("Лист1");
    let sheet_L1 = getContext().getSheetByName(getSettings().sheetName_Лист_1);


    let rowStart = (range.rowStart < this.rowBodyFirst ? this.rowBodyFirst : range.rowStart);
    let rowEnd = range.rowEnd;
    rowEnd = (rowEnd <= this.rowBodyLast ? rowEnd : this.rowBodyLast);
    if ((rowEnd - rowStart + 1) <= 0) return;


    let rangeHeads = sheet_L7.getSheetValues(1, range.columnStart, 1, range.columnEnd - range.columnStart + 1);
    let rangeValues = range.getValues();

    let j = -1;
    for (let row = rowStart; row <= rowEnd; row++) {
      // console.log(`onEdit_L7_v3 Row=${row}=========================================`);
      j++;


      let comentRow_L1 = parseInt(sheet_L7.getRange(row, this.columnLineNumberSheet_1).getValue().toString());
      // let commentMap = (isNaN(comentRow_L1) ? new Map() : makeCommentMap_L7(comentRow_L1, sheet_L1, getClassColRow().list1_col_Coment));
      let comentBul = [0, 0, 0, 0, 0, 0, 0,];           // массив если коментарий изменен 
      let comentArr = ["", "", "", "", "", "", "",];    // массив введеных коментариев
      let countAgArr = ["", "", "", "", "", "", "",];   // массив введенных поставщиков  // email

      // console.log(`onEdit_L7_v3 comentRow_L1=${comentRow_L1}`);
      // console.log(`onEdit_L7_v3 comentRow_L1=${sheet_L7.getRange(row, nr("AZ")).getValue().toString()}`);

      for (let i = 0, col = range.columnStart; col <= range.columnEnd; col++, i++) {

        // console.log(`j=${j};  i=${i};  row= ${row};  col= ${col};  head= ${rangeHeads[0][i]}; value= " + ${rangeValues[j][i]};`);
        if (!((this.columnTopFirst <= col) && (col <= this.columnTopLast))) { continue; }
        this.sheet.getRange(row, col, 1, 1).clearContent();
        let h = fl_str(rangeHeads[0][i]);
        let v = rangeValues[j][i];
        //rangeValues[j][i] = "";

        // if (!v) { continue; }
        let hn = parseInt(h);
        if (isNaN(hn)) { continue; }
        hn = hn - 1;

        if (h.includes(fl_str("Комментарий"))) {
          comentBul[hn] = 1;
          comentArr[hn] = v;
        }

        if (h.includes(fl_str("Поставщик"))) {
          countAgArr[hn] = v;
        }

        if (h.includes(fl_str("формула"))) {
          this.sheet.getRange(row, col).setFormula([this.getFormulaForCol(hn + 1)]);
        }


      }


      let topAg = getTopAg_L7(row, sheet_L7, this.columnTopFirst, this.columnTopLast);      // массив email поставщиков на листе7
      let topAl = getTopAg_L7_link(row, sheet_L7, this.columnTopFirst, this.columnTopLast, this.columnTopSplit); // массив адресов яцеек поставщиков в топе link
      /*
            console.log("-comentBul=" + comentBul);
            console.log("-comentArr=" + comentArr);
            console.log("-countAgArr=" + countAgArr);
            console.log("-topAg=" + topAg);
            console.log("-topAl=" + topAl);
      */
      for (let i = 0; i < comentArr.length; i++) {

        if (comentBul[i] != 1) { continue; }  // не меняли 

        let coment = comentArr[i]; // коментарий введенный

        let cAg = countAgArr[i];  // email введенный 
        let tAg = topAg[i];  // email
        let sAg = topAl[i];  // link

        // console.log(`onEdit_L7_v3 sAg=${sAg}`);
        if (!sAg) {
          if (cAg) { sAg = getLinkForEmail(cAg); }  // поиск link по email введенным пользователем 
          else if (tAg) { sAg = getLinkForEmail(tAg) } // поиск link по email до начало редоктирования 
        }

        if (!sAg) { continue; }  // ничего не нашли адрес яцейки постовщика -=беда=-  
        // console.log(`onEdit_L7_v3 sAg=${sAg}`);
        // console.log(`onEdit_L7_v3 coment=${coment}`);
        if (coment) { coment = `${coment}`.replace(/\|/g, '_') }
        let commentMap_separate = (isNaN(comentRow_L1) ? new Map() : this.makeCommentMap_separate(comentRow_L1, sheet_L1, getClassColRow().list1_col_CommentBlokStart + i));

        commentMap_separate.set(fl_str(sAg), `${fl_str(sAg)}|${coment}|${new Date().valueOf()}`);

        this.saveCommentMap_separate(commentMap_separate, comentRow_L1, sheet_L1, getClassColRow().list1_col_CommentBlokStart + i)

        // commentMap.set(fl_str(sAg), coment);

        // this.addToDB(comentRow_L1, sAg, coment, new Date())

      }
      // if (!isNaN(comentRow_L1)) { saveCommentMap_L7(commentMap, comentRow_L1, sheet_L1, getClassColRow().list1_col_Coment); }
    }

  }

  makeCommentMap_separate(row, sheet, col) {
    if (!(getClassColRow().list1_col_CommentBlokStart <= col && col <= getClassColRow().list1_col_CommentBlokEnd)) { throw new Error("Не веро задана колонка для сохранения кометрария" + `${col}`); }
    // console.log(`makeCommentMap_separate +++++++++++++++++`);
    let joinStr = sheet.getRange(row, col).getValue();
    let joinArr = new Array();
    let commentMap = new Map();
    // console.log(`log map-- joinStr = ${joinStr}`);
    if (joinStr) {
      joinArr = joinStr.split("|||");
    }

    // console.log(`log map-- joinStr = ${joinArr}`);
    for (let i = 0; i < joinArr.length;) {
      let arr = joinArr[i].split("|");
      commentMap.set(fl_str(arr[0]), joinArr[i])
      // console.log(`log map ${i} ${arr[0]} , ${joinArr[i]} `);
      i = i + 1;
    }

    // // console.log(`log map-----------------`);
    // for (let [keyAg, vComm] of commentMap) {
    // console.log(`log map row= ${row}; keyAg=${keyAg};   vComm=${vComm};`);
    // }
    // // console.log(`log map^^^^^^^^^^^^^^^^^^^^`);

    return commentMap;
  }

  saveCommentMap_separate(commentMap, row, sheet, col) {
    // console.log(`saveCommentMap_L7 +++++++++++++++++`);

    if (!(getClassColRow().list1_col_CommentBlokStart <= col && col <= getClassColRow().list1_col_CommentBlokEnd)) { throw new Error("Не веро задана колонка для сохранения кометрария" + `${col}`); }

    let joinArr = new Array();
    for (let [keyAg, vComm] of commentMap) {
      // console.log(` save row= ${row}; keyAg=${keyAg};   vComm=${vComm};`);
      // joinArr.push(keyAg);
      joinArr.push(vComm);

    }
    sheet.getRange(row, col).setValue(joinArr.join("|||"));
    // console.log(`save row= ${row}; joinArr=${joinArr.join("|||")};`);

    // console.log(`saveCommentMap_L7 ^^^^^^^^^^^^^^^^^^^^^^^`);

  }

  getReMailingIdProducts() {

    // this.columnIdProduct 
    // this.columnNameProduct 
    // this.columnRe_mailing

    let retIdProducts = new Array();
    if (retIdProducts) {
      let vlsIdProduct = this.sheet.getRange(this.rowBodyFirst, this.columnIdProduct, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();
      let vlsReMailing = this.sheet.getRange(this.rowBodyFirst, this.columnReMailing, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();


      for (let i = 0; i < vlsReMailing.length; i++) {
        if (!vlsReMailing[i][0]) { continue; }
        retIdProducts.push(vlsIdProduct[i][0])
      }

    }
    // Logger.log(`retIdProducts=${retIdProducts}` );
    return retIdProducts;


  }

  getContragentEmail() {// колонка с емайлом контрагента  
    if (!this.contragentEmailArr) {
      this.contragentEmailArr = this.sheet.getRange(this.rowBodyFirst, this.columnContragentEmail, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues().flat(2);
    }
    return this.contragentEmailArr;
  }

  getContragentRezEmail() {// колонка с емайлом контрагента Резерв 
    if (!this.contragentRezEmailArr) {
      this.contragentRezEmailArr =
        this.sheet.getRange(this.rowBodyFirst, this.columnContragentRezEmail, this.rowBodyLast - this.rowBodyFirst + 1, 1)
          .getValues().flat(2);
    }
    return this.contragentRezEmailArr;
  }


  // тампусто() {

  //   let sheet = this.sheet;
  //   let col = this.columnTopFirst + 1;
  //   let row = getContext().getRowSobachkaBySheetName(sheet.getSheetName());

  //   let тампусто = "&" + (() => {
  //     try {
  //       if (!row) { throw "Нет @" }
  //       row = row + 1;
  //       return sheet.getRange(row, col, 6, 7).getValues().flat(2).join("");
  //     }
  //     catch (err) {
  //       return mrErrToString(err);
  //     }
  //   })();

  //   return тампусто;
  // }

  // tempupdateНастройкиAbkmnhjd() {


  //   let sheet = this.sheet;
  //   let col = this.columnTopFirst + 1;
  //   let row = getContext().getRowSobachkaBySheetName(sheet.getSheetName());
  //   // return
  //   // let тампусто = "&" + (() => { try { return sheet.getRange("AQ1:AT6").getValues().flat(2).join("") } catch (err) { return mrErrToString(err); } })();
  //   let тампусто = this.тампусто();
  //   if (тампусто != "&") { return; }

  //   row = row + 1;
  //   Logger.log(["row", row, "col", col, nc(col)]);

  //   // return
  //   // sheet.getRange("AQ1:AS6").clear();
  //   sheet.getRange(row, col, 6, 7).clear();


  //   // `Описание`,
  //   //   `Цены с листа "5 Закупка товара"`,
  //   //   `Цены из внешней таблицы закупки`,
  //   //   `Не используется`,
  //   //   `Не используется`,
  //   //   `Цены с листа "1-1 Сбор КП"`,

  //   let vls = [
  //     ["Признак", undefined, "Показать", undefined, "Порядок", undefined, `Описание`,],
  //     ["♥️", undefined, true, undefined, 5, undefined, `Цены с листа "5 Закупка товара"`,],
  //     ["♠", undefined, true, undefined, 5, undefined, `Цены из внешней таблицы закупки`,],
  //     ["♣️", undefined, true, undefined, 5, undefined, `Не используется`,],
  //     ["♦️", undefined, true, undefined, 5, undefined, `Не используется`,],
  //     ["-", undefined, true, undefined, 5, undefined, `Цены с листа "1-1 Сбор КП"`,],
  //   ]
  //   sheet.getRange(row, col, 6, 7).setValues(vls);
  //   let rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  //   sheet.getRange(row + 1, col + 3 - 1, 5, 1).setDataValidation(rule);
  // }




}




function saveCommentMap_L7(commentMap, row, sheet, col = getClassColRow().list1_col_Coment) {
  // console.log(`saveCommentMap_L7 +++++++++++++++++`);
  let joinArr = new Array();
  for (let [keyAg, vComm] of commentMap) {
    // console.log(` save row= ${row}; keyAg=${keyAg};   vComm=${vComm};`);

    joinArr.push(keyAg);
    joinArr.push(vComm);

  }

  sheet.getRange(row, col).setValue(joinArr.join("|||"));
  // console.log(`save row= ${row}; joinArr=${joinArr.join("|||")};`);

  // console.log(`saveCommentMap_L7 ^^^^^^^^^^^^^^^^^^^^^^^`);

}


function makeCommentMap_L7(row, sheet, col) {

  // console.log(`makeCommentMap_L7 +++++++++++++++++`);
  let joinStr = sheet.getRange(row, col).getValue();
  let joinArr = new Array();
  let commentMap = new Map();
  // console.log(`log map-- joinStr = ${joinStr}`);
  if (joinStr) {
    joinArr = joinStr.split("|||")
  }

  // console.log(`log map-- joinStr = ${joinArr}`);
  for (let i = 0; i < joinArr.length;) {
    commentMap.set(fl_str(joinArr[i]), joinArr[i + 1])
    // console.log(`log map ${i} ${joinArr[i]} , ${joinArr[i + 1]} `);
    i = i + 2;
  }

  // console.log(`log map-----------------`);
  // for (let [keyAg, vComm] of commentMap) {
  // console.log(`log map row= ${row}; keyAg=${keyAg};   vComm=${vComm};`);
  // }
  // console.log(`log map^^^^^^^^^^^^^^^^^^^^`);

  return commentMap;
}



// function getTopAg_L7(row, sheet) {   // поставцики их email 
function getTopAg_L7(row, sheet, colStart, colEnd) {   // поставцики их email 
  let retArr = ["", "", "", "", "", "", "",];
  // let retArr = ["_1_", "_2_", "_3_", "_4_", "_5_", "_6_", "_7_",];

  let rangeHeads = sheet.getSheetValues(1, colStart, 1, colEnd - colStart + 1);
  let rangeValues = sheet.getSheetValues(row, colStart, 1, colEnd - colStart + 1);

  // console.log("топ7 heads=" + rangeHeads);
  // console.log("топ7 value=" + rangeValues);

  for (let col = 0; col < rangeValues[0].length; col++) {
    h = fl_str(rangeHeads[0][col]);
    // console.log("топ7 H=" + h);
    if (!h.includes(fl_str("поставщ"))) { continue; }
    let hn = parseInt(h);
    if (isNaN(hn)) { continue; }
    hn = hn - 1;
    retArr[hn] = rangeValues[0][col];
  }
  // console.log("топ7 retArr=" + retArr);
  return retArr;
}



// function getTopAg_L7_link(row, sheet) {   // поставцики  ссылка на поставцика в Лист1 

function getTopAg_L7_link(row, sheet, colStart, colEnd, columnTopSplit) {   // поставцики  ссылка на поставцика в Лист1 

  let retArr = ["", "", "", "", "", "", "",];
  // let retArr = ["_1_", "_2_", "_3_", "_4_", "_5_", "_6_", "_7_",];

  let rangeHeads = sheet.getSheetValues(1, colStart, 1, colEnd - colStart + 1)[0];
  let rangeValues = sheet.getSheetValues(row, columnTopSplit, 1, 1)[0][0].toString().split("|||");



  let length = Math.min(rangeValues.length * 2, rangeHeads.length);


  // console.log("топ7 heads=" + rangeHeads);
  // console.log("топ7 value=" + rangeValues);

  for (let col = 0; col < length; col++) {
    h = fl_str(rangeHeads[col]);
    if (!h.includes(fl_str("поставщ"))) { continue; }
    let hn = parseInt(h);
    if (isNaN(hn)) { continue; }
    hn = hn - 1;
    retArr[hn] = rangeValues[hn * 5 + 1];
  }
  // console.log("топ7 retArr=" + retArr);
  return retArr;
}



// это лист 1

// function getLinkForEmail(email, sheet = getContext().getSheetByName("Лист1")) {
function getLinkForEmail(email, sheet = getContext().getSheetByName(getSettings().sheetName_Лист_1)) {

  let p = getClassColRow().list1_col_FirstPrice;
  let hp = getClassColRow().list1_col_LastPrice;

  let rangeEmails = sheet.getSheetValues(1, p, 1, hp - p + 1)[0];

  for (let i = 0; i < rangeEmails.length; i++) { rangeEmails[i] = fl_str(rangeEmails[i]) }

  let col = rangeEmails.indexOf(fl_str(email));

  if (col == -1) { return ""; }  // нет такого у нас емайла 

  col = col + p;
  // console.log(`getLinkForEmail email=${email}  col=${col}  ret=$${nc(col)}$${1}  `);
  return `$${nc(col)}$${1}`;


}









function copyReMailingPriduct() {

  let sheetBuy = getContext().getSheetBuy();

  let urlExternalSpreadSheet = sheetBuy.getUrlExternalSpreadSheet();

  if (!urlExternalSpreadSheet) {
    Logger.log("Не создана нешная таблица");
    return;
  }


  let sheetList7 = getContext().getSheetList7();

  let idProducts = sheetList7.getReMailingIdProducts();
  if (idProducts.length == 0) { Logger.log("Не Продуктов для Повторной рассылки "); return; }
  let sheetList1 = getContext().getSheetList1();


  let vls_prod = sheetList1.sheet.getRange(sheetList1.rowBodyFirst, 1, sheetList1.rowBodyLast - sheetList1.rowBodyFirst + 1, sheetList1.columnLastPrice).getValues();
  // let vv = sheetList1


  let vl = sheetList1.sheet.getRange(1, 1, 1, sheetList1.columnLastPrice).getValues();
  let emailRowArr = new Array(vl[0].length);
  for (let i = 0; i < vl[0].length; i++) {
    if (!isEmail(vl[0][i])) { continue; }
    // emailRowArr[i] = vl[0][i];
    emailRowArr[i] = `${vl[0][i]}`.trim().split(" ", 1)[0];
  }
  // Logger.log(emailRowArr);

  // Logger.log(idProducts);
  // vls_prod.forEach((v,i, arr) => { Logger.log(v); });

  vls_prod = vls_prod.filter((v, i, arr) => {
    if (!v[sheetList1.columnProduct - 1]) { return false; }
    if (idProducts.includes(v[2])) { return true; }
    return false;

  });

  // vls_prod.forEach((v,i, arr) => { Logger.log(v); });

  // let nameProducts = vls_prod.map((v) => v[sheetList1.columnProduct - 1]);

  let products = new Map()
  vls_prod.forEach((v, i, arr) => {
    let emails = new Array()
    // let ret = {name:}

    for (let j = 0; j < v.length; j++) {
      if (!v[j]) { continue; }
      if (!isEmail(emailRowArr[j])) { continue; }
      emails.push(emailRowArr[j]);
    }

    let ret = {
      "id": v[sheetList1.columnProduct - 1 - 1],
      "name": v[sheetList1.columnProduct - 1],
      // "group": v[sheetList1.columnGroup - 1],

      "character": v[sheetList1.columnProduct + 0],
      "measur": v[sheetList1.columnProduct + 1],
      "quantity": v[sheetList1.columnProduct + 2],

      "emails": emails,
    }
    products.set(ret.name, ret);
  });
  // Logger.log(products);

  // Logger.log(nameProducts);

  let sheetProductGroups = getContext().getSheetProductGroups();
  let vls_prod_g = sheetProductGroups.sheet.getRange(1, 1, sheetProductGroups.sheet.getLastRow(), nr("I")).getValues();

  // vls_prod_g.forEach((v, i, arr) => {
  //   if (!products.has(v[1])) { 

  //     // prod["description"] = [v[nr("G") - 1], v[nr("G")], v[nr("G") + 1]];

  //     return; 
  //   }
  //   prod = products.get(v[1]);
  //   prod["description"] = [v[nr("G") - 1], v[nr("G")], v[nr("G") + 1]];
  // });


  products.forEach(product => {
    product["groups"] = vls_prod_g.filter(v => {
      if (v[nr("B") - 1] == product.name) return true;
      return false;
    }).map(v => v[nr("F") - 1]);
  });


  products.forEach(product => {
    product["specifications"] = vls_prod_g.filter(v => {
      if (v[nr("B") - 1] == product.name) return true;
      return false;
    }).map(v => v[nr("C") - 1]).join(" ");

  });


  let groups = new Map();

  let col_F = nr("F")
  vls_prod_g.forEach((v) => {

    let group_name = v[col_F - 1];
    if (!group_name) { return; }

    if (!groups.has(group_name)) {
      groups.set(group_name, new Array());
    }

    let descriptions = [v[col_F + 0], v[col_F + 1], v[col_F + 2]].filter(e => `${e}`);
    if (descriptions.length == 0) { return; }

    /** @type {[]} */
    let group_descriptions = groups.get(group_name);
    group_descriptions = group_descriptions.concat(descriptions)
    groups.set(group_name, group_descriptions);
  });


  products.forEach((v, k, m) => { Logger.log(`k=${k} | v=${JSON.stringify(v)}`) });

  groups.forEach((v, k, m) => { Logger.log(`k=${k} | v=${JSON.stringify(v)}`) });

  let url = urlExternalSpreadSheet;
  MrLib_External.addReMailingPriduct(products, groups, url);
}








function copyReMailingPriductV2() {

  let sheetBuy = getContext().getSheetBuy();

  let urlExternalSpreadSheet = sheetBuy.getUrlExternalSpreadSheet();

  if (!urlExternalSpreadSheet) {
    Logger.log("Не создана внешная таблица");
    return;
  }


  let sheetList7 = getContext().getSheetList7();

  let idProducts = sheetList7.getReMailingIdProducts();
  if (idProducts.length == 0) { Logger.log("Не Продуктов для Повторной рассылки "); return; }

  idProducts.forEach(p => Logger.log(JSON.stringify(p, undefined, 2)));

  let sheetList1 = getContext().getSheetList1();


  let vls_prod = sheetList1.sheet.getRange(sheetList1.rowBodyFirst, 1, sheetList1.rowBodyLast - sheetList1.rowBodyFirst + 1, sheetList1.columnLastPrice).getValues();
  // let vv = sheetList1


  let vl = sheetList1.sheet.getRange(1, 1, 1, sheetList1.columnLastPrice).getValues();
  let emailRowArr = new Array(vl[0].length);
  for (let i = 0; i < vl[0].length; i++) {
    if (!isEmail(vl[0][i])) { continue; }
    // emailRowArr[i] = vl[0][i];
    emailRowArr[i] = `${vl[0][i]}`.trim().split(" ", 1)[0];
  }
  // Logger.log(emailRowArr);

  // Logger.log(idProducts);
  // vls_prod.forEach((v,i, arr) => { Logger.log(v); });

  vls_prod = vls_prod.filter((v, i, arr) => {
    if (!v[sheetList1.columnProduct - 1]) { return false; }
    if (idProducts.includes(v[2])) { return true; }
    return false;

  });

  // vls_prod.forEach((v,i, arr) => { Logger.log(v); });

  // let nameProducts = vls_prod.map((v) => v[sheetList1.columnProduct - 1]);

  let products = new Map()
  vls_prod.forEach((v, i, arr) => {
    let emails = new Array()
    // let ret = {name:}

    for (let j = 0; j < v.length; j++) {
      if (!v[j]) { continue; }
      if (!isEmail(emailRowArr[j])) { continue; }
      emails.push(emailRowArr[j]);
    }

    let ret = {
      "id": v[sheetList1.columnProduct - 1 - 1],
      "name": v[sheetList1.columnProduct - 1],
      // "group": v[sheetList1.columnGroup - 1],

      "character": v[sheetList1.columnProduct + 0],
      "measur": v[sheetList1.columnProduct + 1],
      "quantity": v[sheetList1.columnProduct + 2],

      "emails": emails,
    }
    products.set(ret.name, ret);
  });
  // Logger.log(products);

  // Logger.log(nameProducts);

  let sheetProductGroups = getContext().getSheetProductGroups();
  let vls_prod_g = sheetProductGroups.sheet.getRange(1, 1, sheetProductGroups.sheet.getLastRow(), nr("J")).getValues();

  // vls_prod_g.forEach((v, i, arr) => {
  //   if (!products.has(v[1])) { 

  //     // prod["description"] = [v[nr("G") - 1], v[nr("G")], v[nr("G") + 1]];

  //     return; 
  //   }
  //   prod = products.get(v[1]);
  //   prod["description"] = [v[nr("G") - 1], v[nr("G")], v[nr("G") + 1]];
  // });


  products.forEach(product => {
    product["groups"] = vls_prod_g.filter(v => {
      if (v[nr("B") - 1] == product.name) return true;
      return false;
    }).map(v => v[nr("G") - 1]);
  });


  products.forEach(product => {
    product["specifications"] = vls_prod_g.filter(v => {
      if (v[nr("B") - 1] == product.name) return true;
      return false;
    }).map(v => v[nr("C") - 1]).join(" ");

  });


  let groups = new Map();

  let col_G = nr("G")
  vls_prod_g.forEach((v) => {

    let group_name = v[col_G - 1];
    if (!group_name) { return; }

    if (!groups.has(group_name)) {
      groups.set(group_name, new Array());
    }

    let descriptions = [v[col_G + 0], v[col_G + 1], v[col_G + 2]].filter(e => `${e}`);
    if (descriptions.length == 0) { return; }

    /** @type {[]} */
    let group_descriptions = groups.get(group_name);
    group_descriptions = group_descriptions.concat(descriptions)
    groups.set(group_name, group_descriptions);
  });


  products.forEach((v, k, m) => { Logger.log(`k=${k} | v=\n${JSON.stringify(v, null, 2)}`) });

  groups.forEach((v, k, m) => { Logger.log(`k=${k} | v=\n${JSON.stringify(v, null, 2)}`) });
  //  sheetList1.getEmailProductGroupMap().forEach((v,k,m)=>{Logger.log(`k=${k} | v=${JSON.stringify(v)}`)});
  


  let url = urlExternalSpreadSheet;

  // if (!isTest()) { return; }
  // if (!isTest()) { return; }
  // return;
  MrLib_External.addReMailingPriduct(products, groups, url);

}

















