function report(colorArr) {
  return getContext().getSheetList1().report(colorArr);
}

function recalculate_report_malo() {

  getContext().getSheetList1().recalculate_report_malo();

}

function menuAddEmtyColsBeforeАктуальный_Сбор_КП() {
  getContext().getSheetList1().addEmtyColsBeforeАктуальный_Сбор_КП(60);
  classColRow = undefined;
  getContext().sheetList1 = undefined;
  getContext().getSheetList1().fixSheetCompact();
}

function menuFixSheetCompact() {
  getContext().getSheetList1().fixSheetCompact();
}

function minimumPriceForRow(rows) {
  return resultPriceForRow(undefined, rows);
}

/**
 * @param {[[number ]]} rows
 * @param {[[number ]]} quantities
 */
function resultPriceForRow(quantities, rows, cols = undefined) {
  let colorArr = [[getContext().getAnalogColor()]];
  let minPrices = getContext().getSheetList1().chooseMinimumPricesByColors(quantities, rows, cols, colorArr).minPrices;
  let ret = minPrices.map(v => v);
  ret = ret.map(v => v.filter(vi => !isNaN(vi)));
  // ret = ret.map(v => v.filter(vi => vi != undefined));
  // ret = ret.map((v, i, arr) => { return v.map(vi => vi * quantities[i][0]); })
  ret = ret.map(v => v.filter(vi => vi != 0));
  let myNumberFormat = new Intl.NumberFormat('ru-RU', { minimumIntegerDigits: 1, minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: true })
  ret = ret.map(v => {
    // return [].concat(v).filter(vi => vi != undefined).map(vi => { return myNumberFormat.format(vi); }).join("\n");
    return v.map(vi => { return myNumberFormat.format(vi); }).join("\n");
  });
  return ret;
}

function totalPrice(quantities, rows, cols, colorArr) {
  // let colorArr = [[getContext().getAnalogColor()]];
  let minPrices = getContext().getSheetList1().chooseMinimumPricesByColors(quantities, rows, cols, colorArr).minPrices;

  Logger.log(`${JSON.stringify(minPrices)}`);
  let ret = [new Array(minPrices[0].length)];
  ret[0].fill(0);
  // minPrices.forEach(r => r.forEach((v, i, arr) => { ret[0][i] = ret[0][i] + v }));
  for (let i = 0; i < minPrices.length; i++) {
    for (let j = 0; j < minPrices[i].length; j++) {
      if (!minPrices[i][j]) { continue; }

      minPrices[i][j];
      ret[0][j] = ret[0][j] + minPrices[i][j];
    }
  }



  // ret = ret.map(v => v.filter(vi => !isNaN(vi)));
  // ret = ret.map(v => v.filter(vi => vi != 0));

  let myNumberFormat = new Intl.NumberFormat('ru-RU', { minimumIntegerDigits: 1, minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: true })
  ret = ret.map(v => {
    return v.map(vi => { return myNumberFormat.format(vi); }).join("\n");
  });

  Logger.log(`${JSON.stringify(ret)}`);

  return ret;
  // return minPrices;
}




/**
 * @param {[[number ]]} rows
 * @param {[[number ]]} quantities
 * @param {[[number ]]} cols
 * @param {[[number ]]} colorArr
 */
function chooseMinimumPricesByColors(quantities, rows, cols, colorArr) {
  return getContext().getSheetList1().chooseMinimumPricesByColors(quantities, rows, cols, colorArr);
}




function onEditList_1(edit) {
  getContext().getSheetList1().onEdit(edit);
}





class Counter {
  constructor(name) {
    this.name = fl_str(name);
    this.priceAll = 0;
    this.priceAnalog = 0;
  }

  increment(isPrice, isAnalogPrice) {
    if (isPrice) { this.priceAll++; }
    if (isAnalogPrice) { this.priceAnalog++; }

  }

}


class CounterGroup {
  constructor(name) {
    this.name = fl_str(name);
    this.priceAll = 0;
    this.priceAnalog = 0;
    this.mapProduct = new Map();
  }

  increment(isPrice, isAnalogPrice) {
    if (isPrice) { this.priceAll++; }
    if (isAnalogPrice) { this.priceAnalog++; }
  }
  /** @param {Counter} prCounter */
  addProduct(prCounter) {
    this.mapProduct.set(prCounter.name, prCounter);
  }

  getCouts(maloOtvetov) {
    if (!this.counts) {
      this.counts = {
        grCoun_1: 0,
        grCoun_2: 0,
        grCoun_3: 0,
      }

      this.mapProduct.forEach(/** @param {Counter} item */ item => {
        if (item.priceAll == 0) { this.counts.grCoun_1++; }
        if (item.priceAll < maloOtvetov) { this.counts.grCoun_2++; }
        if ((item.priceAll - item.priceAnalog) < maloOtvetov) { this.counts.grCoun_3++; }
      });
    }
    return this.counts;
  }



}






/**
 * @typedef PriceAndQuantities
 * @type {object}
 * @property {[[]]} minPrices
 * @property {[[]]} quantities
 * @property {[]} colorNameArr
 */
class ClassSheet_List1 {
  constructor() { // class constructor
    // this.sheet = getContext().getSheetByName("Лист1");
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Лист_1);
    // this.columnFirstPrice = nr("P");
    // this.columnLastPrice = nr("HP");
    // this.columnProduct = nr("D");
    // this.columnGroup = nr("L");
    // this.columnComent = nr("HX");
    // this.columnTopFormula = nr("HY");
    // this.columnS = nr("S");
    // this.columnM = nr("M");

    this.columnFirstPrice = getClassColRow().list1_col_FirstPrice;
    this.columnLastPrice = getClassColRow().list1_col_LastPrice;
    this.columnProduct = getClassColRow().list1_col_Product;
    this.columnGroup = getClassColRow().list1_col_Group;
    this.columnComent = getClassColRow().list1_col_Coment;
    this.columnTopFormula = getClassColRow().list1_col_TopFormula;
    // this.columnS = getClassColRow().list1_col_S;
    // this.columnM = getClassColRow().list1_col_M;
    this.columnS = this.columnFirstPrice;
    this.columnM = this.columnFirstPrice;


    this.columnКол_во = getClassColRow().list1_col_Кол_во;
    this.columnЕдИзм = getClassColRow().list1_col_ЕдИзм;
    this.columnId = getClassColRow().list1_col_id;
    this.columnАктуальный_Сбор_КП = getClassColRow().list1_col_Актуальный_Сбор_КП;

    this.columnПредложений = getClassColRow().list1_col_Предложений;
    // list1_col_Предложений

    this.rowBodyFirst = 2;

    this.rowBodyLast = this.findRowBodyLast();

    this.analogMap = undefined;
    this.analogArrLink = undefined;


    this.emailRowArr = undefined;
    this.productGroupColumnArr = undefined;
    this.productColumnArr = undefined;


    this.emailProductGroupMap = undefined;
    this.emailProductMap = undefined;
    this.emailResponseMap = undefined;

    this.summaryArr = undefined;



  }


  report(colorArr) {
    // Logger.log(`  report colorArr =${JSON.stringify(colorArr)}`);



    // let minPrices = getContext().getSheetList1().chooseMinimumPricesByColors(quantities, rows, cols, colorArr);
    let pq = getContext().getSheetList1().chooseMinimumPricesByColors(undefined, undefined, undefined, colorArr);

    let minPrices = pq.minPrices;
    let quantities = pq.quantities;
    let colorNameArr = pq.colorNameArr;


    for (let i = 0; i < minPrices.length; i++) {
      for (let j = 0; j < minPrices[i].length; j++) {
        if (!isNaN(minPrices[i][j])) continue;
        minPrices[i][j] = 0;
      }
    }


    let price = new Array(minPrices.length);
    for (let i = 0; i < minPrices.length; i++) {
      price[i] = new Array(minPrices[i].length);
      price[i].fill(0);
      if (!quantities[i][0]) { continue; }
      for (let j = 0; j < minPrices[i].length; j++) {
        if (!minPrices[i][j]) { continue; }
        price[i][j] = minPrices[i][j] * quantities[i][0];
      }
    }


    // let jOther = price[0].length - 1;
    let total = [new Array(price[0].length)];
    total[0].fill(0);
    for (let i = 0; i < price.length; i++) {

      let pOther = 0;
      // let tt = price[i].filter(v => (!isNaN(v) || v != 0));
      let tt = price[i].filter(v => v != 0);
      // let tt = price[i].filter(v => !isNaN(v));
      if (tt.length > 0) {
        pOther = Math.min(...tt);
        // Logger.log(pOther);
      }


      for (let j = 0; j < price[0].length; j++) {
        // if (isNaN(price[i][j])) {
        if (price[i][j] == 0) {
          total[0][j] = total[0][j] + pOther;
        } else {
          total[0][j] = total[0][j] + price[i][j];
        }
      }

    }


    price.push(total[0]);


    // let myNumberFormat = new Intl.NumberFormat('ru-RU', { minimumIntegerDigits: 1, minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: true })
    // total = total.map(v => {
    //   return v.map(vi => { return myNumberFormat.format(vi); }).join("\n");
    // });

    // Logger.log(`${JSON.stringify(total)}`);


    let myNumberFormat = new Intl.NumberFormat('ru-RU', { minimumIntegerDigits: 1, minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: true })

    // price = price.map(v => { return v.map(vi => { if (isNaN(vi)) { return undefined } return myNumberFormat.format(vi); }).join("\n"); });
    // minPrices = minPrices.map(v => { return v.map(vi => { if (isNaN(vi)) { return undefined } return myNumberFormat.format(vi); }).join("\n"); });

    price = price.map(v => { if (`${v.join("")}` != "000") return v.map(vi => { return myNumberFormat.format(vi); }).join("\n"); return undefined; });
    minPrices = minPrices.map(v => { if (`${v.join("")}` != "000") return v.map(vi => { return myNumberFormat.format(vi); }).join("\n"); return undefined; });


    minPrices.push(colorNameArr.concat(["Остальные"]).flat().join("\n"));

    // return minPrices;

    let ret = new Array(minPrices.length);

    // ret =[["ddasasd","dasdre"]].concat( ret.map((v, i, arr) => {return [`${minPrices[i][0]}`, `${price[i][0]}`]}));

    for (let i = 0; i < ret.length; i++) {
      ret[i] = new Array(2);
      ret[i][0] = minPrices[i];
      ret[i][1] = price[i];
    }
    return ret;
    // return price;
    // return minPrices;
    // return total;

  }

  /**
   * @param {[[number ]]} rows
   * @param {[[number ]]} quantities
   * @param {[[number ]]} cols
   * @param {[[number ]]} colorArr
   * @returns  {PriceAndQuantities}
   */
  chooseMinimumPricesByColors(quantities, rows, cols, colorArr) {
    // Logger.log(`quantities =${JSON.stringify(quantities)}`);
    // Logger.log(`rows =${JSON.stringify(rows)}`);
    // Logger.log(`cols =${JSON.stringify(cols)}`);
    // Logger.log(`colorArr =${JSON.stringify(colorArr)}`);



    // проверяем цвета
    // if (!colorArr) { colorArr = [[getContext().getAnalogColor()]]; }


    if (!colorArr) { colorArr = [["Аналог"]]; }


    if (!Array.isArray(colorArr)) { colorArr = [[colorArr]] }
    let colorNameArr = colorArr.flat();
    colorArr = colorArr.flat().map(colorName => { return (`${colorName}`.includes("#") ? colorName : getContext().getColorMap().getColor(colorName)) });
    let colorDef = getContext().getColorMap().def;
    colorArr = colorArr.map(c => { return (c == colorDef ? undefined : c); });
    colorArr = [colorArr];
    // Logger.log(`${JSON.stringify(colorArr)}`);



    // проверяем строки 
    if (!rows) {
      let rowBodyFirst = this.rowBodyFirst;
      let rowBodyLast = this.rowBodyLast;
      rows = new Array(rowBodyLast - rowBodyFirst + 1);
      rows.fill(rowBodyFirst);
      rows = rows.map((r, i, arr) => { return [r + i] });
      // Logger.log(`def rows=${JSON.stringify(rows)}`);
    }
    if (!Array.isArray(rows)) {
      if (!isNaN(parseInt(rows))) { rows = [[parseInt(rows)]]; } else { throw new Error("значение  аргумента Строки не число"); }
    }
    if (rows.length == 0) { throw new Error("Пустой масив данных  'Строка'"); }
    if (rows[0].length == 0) { throw new Error("нет значений 'Строка'"); }
    // if (!Array.isArray(rows[0])) { rows = [rows] }



    // проверяем столбцы
    if (!cols) {
      let columnS = this.columnS;
      let columnLastPrice = this.columnLastPrice;
      cols = new Array(columnLastPrice - columnS + 1);
      cols.fill(columnS);
      cols = [cols.map((c, i, arr) => { return c + i })];
      // Logger.log(`def cols=${JSON.stringify(cols)}`);
    }
    if (!Array.isArray(cols)) {
      if (!isNaN(parseInt(cols))) { cols = [[parseInt(cols)]]; } else { throw new Error("значение аргумента Столбца не число"); }
    }

    if (cols.length == 0) { throw new Error("Пустой масив данных  'Столбц'"); }
    if (cols[0].length == 0) { throw new Error("нет значений 'Столбц'"); }



    // проверяем количевство 
    if (!quantities) {
      quantities = this.sheet.getRange(this.rowBodyFirst, this.columnКол_во, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();
      // quantities = new Array(rows.length);
      // quantities.fill(10);
      // quantities = quantities.map(v => [v]);
    }

    if (!Array.isArray(quantities)) {
      if (!isNaN(parseFloat(quantities))) { quantities = [[parseFloat(quantities)]]; } else { throw new Error("значение  аргумента количевство не число"); }
    }

    if (quantities.length == 0) { throw new Error("Пустой масив данных Количевства"); }
    if (quantities[0].length == 0) { throw new Error("нет значений  Количевства"); }
    // if (!Array.isArray(rows[0])) { rows = [rows] }

    if (quantities.length != rows.length) { throw new Error("Разные длины массивов Количевства и Строк"); }




    // Logger.log(`quantities =${JSON.stringify(quantities)}`);
    // Logger.log(`rows =${JSON.stringify(rows)}`);
    // Logger.log(`cols =${JSON.stringify(cols)}`);
    // Logger.log(`colorArr =${JSON.stringify(colorArr)}`);


    let minPrices = this.minimumPriceForRow(rows, cols, colorArr);
    // let price = new Array(minPrices.length);

    // for (let i = 0; i < minPrices.length; i++) {
    //   price[i] = new Array(minPrices[i].length);
    //   if (!quantities[i][0]) { continue; }
    //   // price[i] = new Array(minPrices[i].length);
    //   for (let j = 0; j < minPrices[i].length; j++) {
    //     if (!minPrices[i][j]) { continue; }
    //     price[i][j] = minPrices[i][j] * quantities[i][0];
    //   }
    // }
    // return price;
    let ret = { minPrices, quantities, colorNameArr };
    Logger.log(`ret =${JSON.stringify(ret)}`);
    return ret;

    // return "ttt";
  }






  /**
   * @param {[[number ]]} rows
   * @param {[[number ]]} cols
   * @param {[[]]} colorArr
   */
  minimumPriceForRow(rows, cols, colorArr) {

    // this.columnS
    // this.columnLastPrice

    // this.rowBodyFirst
    // this.rowBodyLast

    // Logger.log(`${JSON.stringify(rows)}`);

    colorArr = colorArr.flat();

    let retArr = new Array(rows.length);


    let rangeHead = this.sheet.getRange(1, this.columnS, 1, this.columnLastPrice - this.columnS + 1);
    let rangeBody = this.sheet.getRange(this.rowBodyFirst, this.columnS, this.rowBodyLast - this.rowBodyFirst + 1, this.columnLastPrice - this.columnS + 1);

    let vlsHead = rangeHead.getValues();
    let vlsBody = rangeBody.getValues();
    let vlsBg = rangeHead.getBackgrounds();




    vlsHead = vlsHead.map((v, i, arr) => { return [i + 1].concat(v); });
    vlsBg = vlsBg.map((v, i, arr) => { return [i + 1].concat(v); });
    vlsBody = vlsBody.map((v, i, arr) => { return [i + this.rowBodyFirst].concat(v); });


    // Logger.log(`vlsHead =${JSON.stringify(vlsHead)}`);
    // Logger.log(`vlsBg =${JSON.stringify(vlsBg)}`);
    // Logger.log(`vlsBody =${JSON.stringify(vlsBody)}`);




    let rowArr = [].concat(rows.flat());
    rowArr = rowArr.map(r => parseInt(r));
    // Logger.log(`rowArr = ${rowArr}`);
    // Logger.log(`colorArr = ${JSON.stringify(colorArr)}`);

    for (let i = 0; i < vlsBody.length; i++) {
      let row_i = vlsBody[i][0];
      let ind = rowArr.indexOf(row_i);
      // Logger.log(`ind=${ind} | row_i=${row_i}`);

      retArr[ind] = new Array(colorArr.length + 1);
      if (ind == -1) { continue; }
      // Logger.log(`ind=${ind} | row_i=${row_i}`);

      // Logger.log(`row_i=${row_i}`);
      retArr[ind] = this.minPriceFor(retArr[ind], cols[0], colorArr, vlsHead[0], vlsBody[i], vlsBg[0]);
      // Logger.log(` retArr[${ind}]=${JSON.stringify(retArr[ind])} `);
    }

    // retArr = retArr.map(v => {      let ret = `${JSON.stringify(v)}`;      return ret;    });

    // retArr = retArr.map(v => [v]);
    return retArr;
  }

  /**
   * @param {[number]} retArr
   * @param {[number]} cols
   * @param {[number]} colorArr
   * @param {[number]} vlsHead
   * @param {[number]} vlsBody
   * @param {[number]} vlsBg
   */
  minPriceFor(retArr, cols, colorArr, vlsHead, vlsBody, vlsBg) {
    // Logger.log(`vlsHead=${JSON.stringify(vlsHead)} |  vlsBg=${JSON.stringify(vlsBg)} | vlsBody=${JSON.stringify(vlsBody)} `);
    // Logger.log(`vlsHead =${JSON.stringify(vlsHead)}`);
    // Logger.log(`vlsBg =${JSON.stringify(vlsBg)}`);
    // Logger.log(`vlsBody =${JSON.stringify(vlsBody)}`);




    // Logger.log(`colorArr.length=${colorArr.length}`);
    // let retArr = new Array(colorArr.length);
    let retOth = undefined;
    for (let i = 1; i < vlsHead.length; ++i) {
      // Logger.log(`  vlsHead[${i}]=${JSON.stringify(vlsHead[i])} |  vlsBg[${i}]=${JSON.stringify(vlsBg[i])} | vlsBody[${i}]=${JSON.stringify(vlsBody[i])} `);
      if (`${vlsHead[i]}` == "") { continue; }


      if (!cols.includes(i + this.columnS)) { continue; }

      if (`${vlsHead[i]}` == "") { continue; }  // в шапке пусто  пропускаем
      if (`${vlsBody[i]}` == "") { continue; }  // цена пусто пропускаем // нет значения 
      if (isNaN(parseFloat(vlsBody[i]))) { continue; }  // не число


      // Logger.log(`vlsHead[${i}]=${JSON.stringify(vlsHead[i])} |  vlsBg[${i}]=${JSON.stringify(vlsBg[i])} | vlsBody[${i}]=${JSON.stringify(vlsBody[i])} `);
      // Logger.log(`vlsBody[${i}]=${JSON.stringify(vlsBody[i])}`);
      let fOth = true;
      for (let color = 0; color < colorArr.length; color++) {
        if (vlsBg[i] == colorArr[color]) {
          // Logger.log(` parseFloat(vlsBody[i])=${parseFloat(vlsBody[i])}`)

          // if (retArr[c] == undefined) { retArr[c] = parseFloat(vlsBody[i]); } else { retArr[c] += parseFloat(vlsBody[i]); }
          if (retArr[color] == undefined) { retArr[color] = parseFloat(vlsBody[i]); } else { retArr[color] = Math.min(retArr[color], parseFloat(vlsBody[i])); }

          // Logger.log(`retArr[${c}]=${retArr[c]} | ${vlsBg[i] == colorArr[c]} | ${vlsBg[i]} = ${colorArr[c]} | vlsBody[${i}]=${vlsBody[i]}`);
          fOth = false;
        }
      }

      // if (fOth) { if (retOth == undefined) { retOth = parseFloat(vlsBody[i]); } else { retOth = retOth + parseFloat(vlsBody[i]); } }
      if (fOth) { if (retOth == undefined) { retOth = parseFloat(vlsBody[i]); } else { retOth = Math.min(retOth, parseFloat(vlsBody[i])) } }
    }

    // Logger.log(`retArr=${JSON.stringify(retArr)}  |  retOth=${JSON.stringify(retOth)}`);

    retArr[retArr.length - 1] = retOth;

    // retArr = [].concat(1234567.123, retArr, 1234567.123);
    // retArr = retArr.concat(1234567.123);
    return retArr;
  }


  getSummary(maloOtvetov, colorAnalog) {
    if (this.summaryArr) { return this.summaryArr; }
    let mapCountGroup = new Map();
    // let arrProd = new Array();
    let arrCountProd = new Array();

    let vlsProd = this.sheet.getRange(this.rowBodyFirst, this.columnProduct, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues()
    let vlsGroup = this.sheet.getRange(this.rowBodyFirst, this.columnGroup, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues()

    for (let i = 0; i < vlsProd.length; i++) {
      if (!vlsProd[i][0]) { continue; }  //наименование пусто 
      let pr = vlsProd[i][0] = fl_str(vlsProd[i][0]);
      let gr = vlsGroup[i][0] = fl_str(vlsGroup[i][0]);
      if (!mapCountGroup.has(gr)) { mapCountGroup.set(gr, new CounterGroup(gr)) }
    }


    var vlsPrice = this.sheet.getRange(this.rowBodyFirst, this.columnS, this.rowBodyLast - this.rowBodyFirst + 1, this.columnLastPrice - this.columnS).getValues();
    var vlsBG = this.sheet.getRange(1, this.columnS, 1, this.columnLastPrice - this.columnS + 1).getBackgrounds();
    var vlsHead = this.sheet.getRange(1, this.columnS, 1, this.columnLastPrice - this.columnS + 1).getValues();

    for (let i = 0; i < vlsProd.length; i++) {
      if (!vlsProd[i][0]) { continue; }  //наименование пусто 
      let pr = vlsProd[i][0];
      let gr = vlsGroup[i][0];

      let prCounter = new Counter(pr);
      /** @type {CounterGroup} */
      let grCounter = mapCountGroup.get(gr);
      grCounter.addProduct(prCounter);

      for (let j = 0; j < vlsPrice[i].length; j++) {

        if (!vlsHead[0][j]) { continue; }  // в шапке пусто  пропускаем
        if (!vlsPrice[i][j]) { continue; } // цена пусто пропускаем // нет значения 
        prCounter.increment(true, vlsBG[0][j] == colorAnalog);
        grCounter.increment(true, vlsBG[0][j] == colorAnalog);
      }
      arrCountProd.push(prCounter);
    }

    // Брать данные с листа 1
    let grCoun_1 = 0;// Кол-во групп у которых нет цены
    let grCoun_2 = 0// Кол-во групп в которых кол-во всех предложений (обычных + аналог) предложений меньше чем $B$8
    let grCoun_3 = 0// Кол-во групп в которых основных предложений (без учета аналогов) меньше чем $B$8
    let prCoun_1 = 0// Кол-во товаров у которых нет цены
    let prCoun_2 = 0// Кол-во товаров в которых кол-во всех предложений (обычных + аналог) предложений меньше чем $B$8
    let prCoun_3 = 0// Кол-во товаров в которых основных предложений (без учета аналогов) меньше чем $B$8
    // Logger.log(mapCountGroup.size);
    for (let [gr, counter] of mapCountGroup) {
      // Logger.log(gr);
      let counts = counter.getCouts(maloOtvetov);
      // grCoun_1 = counts.grCoun_1
      // grCoun_2 = counts.grCoun_2
      // grCoun_3 = counts.grCoun_3

      if (counts.grCoun_1 != 0) { grCoun_1++; }
      if (counts.grCoun_2 != 0) { grCoun_2++; }
      if (counts.grCoun_3 != 0) { grCoun_3++; }

      // if (counts.priceAll == 0) { grCoun_1++; }
      // if (counter.priceAll < maloOtvetov) { grCoun_2++; }
      // if ((counter.priceAll - counter.priceAnalog) < maloOtvetov) { grCoun_3++; }


    }

    // Logger.log(arrCountProd);

    for (let i = 0; i < arrCountProd.length; i++) {
      let counter = arrCountProd[i];
      if (counter.priceAll == 0) { prCoun_1++; }
      if (counter.priceAll <= maloOtvetov) { prCoun_2++; }
      if ((counter.priceAll - counter.priceAnalog) <= maloOtvetov) { prCoun_3++; }
    }




    this.summaryArr = new Array();
    // this.summaryArr.push([`Всего групп ${getSettings().sheetName_Лист_1}`, mapCountGroup.size]);

    // this.summaryArr.push([`Кол-во групп у которых нет цены`, grCoun_1]);
    // this.summaryArr.push([`Кол-во групп в которых кол-во всех предложений (обычных + аналог) предложений меньше чем ${maloOtvetov}`, grCoun_2]);
    // this.summaryArr.push([`Кол-во групп в которых основных предложений (без учета аналогов) меньше чем ${maloOtvetov}`, grCoun_3]);

    // this.summaryArr.push([`Всего товаров ${getSettings().sheetName_Лист_1}`, arrCountProd.length]);
    // this.summaryArr.push([`Кол-во товаров у которых нет цены`, prCoun_1]);
    // this.summaryArr.push([`Кол-во товаров в которых кол-во всех предложений (обычных + аналог) предложений меньше чем ${maloOtvetov}`, prCoun_2]);
    // this.summaryArr.push([`Кол-во товаров в которых основных предложений (без учета аналогов) меньше чем ${maloOtvetov}`, prCoun_3]);


    this.summaryArr.push([`Всего групп ${getSettings().sheetName_Лист_1}`, mapCountGroup.size]);

    // this.summaryArr.push([`Кол-во групп у которых нет цены`, grCoun_1]);
    // this.summaryArr.push([`Кол-во групп в которых кол-во всех предложений (обычных + аналог) меньше чем ${maloOtvetov + 1}`, grCoun_2]);
    // this.summaryArr.push([`Кол-во групп в которых основных предложений (без учета аналогов) меньше чем ${maloOtvetov + 1}`, grCoun_3]);

    this.summaryArr.push([`Кол-во групп у которых имеются товары у которых нет цены`, grCoun_1]);
    this.summaryArr.push([`Кол-во групп в которых имеются товары у которых всех предложений (обычных + аналог) меньше чем ${maloOtvetov + 1}`, grCoun_2]);
    this.summaryArr.push([`Кол-во групп в которых имеются товары у которых основных предложений (без учета аналогов) меньше чем ${maloOtvetov + 1}`, grCoun_3]);



    this.summaryArr.push([`Всего товаров ${getSettings().sheetName_Лист_1}`, arrCountProd.length]);
    this.summaryArr.push([`Кол-во товаров у которых нет цены`, prCoun_1]);
    this.summaryArr.push([`Кол-во товаров в которых кол-во всех предложений (обычных + аналог) меньше чем ${maloOtvetov + 1}`, prCoun_2]);
    this.summaryArr.push([`Кол-во товаров в которых основных предложений (без учета аналогов) меньше чем ${maloOtvetov + 1}`, prCoun_3]);



    return this.summaryArr;

  }




  fixSheet(duration = 1 / 24 / 60 * 4.9) {
    // return
    Logger.log("ClassSheet_List1 fixSheet");
    let max_col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheet.getSheetName()).getMaxColumns();

    if (max_col != getClassColRow().list1_col_MAX_COL) {
      Logger.log("ClassSheet_List1 fixSheet колличество столбцов разное " + `
      list1_col_MAX_COL=${getClassColRow().list1_col_MAX_COL} | max_col=${max_col}`);
      return;
    } else {
      Logger.log("ClassSheet_List1 fixSheet колличество столбцов одинаковое " + `
      list1_col_MAX_COL=${getClassColRow().list1_col_MAX_COL} | max_col=${max_col}`);
    }


    if (this.rowBodyFirst > this.rowBodyLast) { return; }

    let formulaComent = this.formulaR1C1_ForComments();
    let rangeComent = this.sheet.getRange(this.rowBodyFirst, this.columnComent, this.rowBodyLast - this.rowBodyFirst + 1, 1);
    let formulasComent = rangeComent.getFormulasR1C1();

    let isMmodified = false
    formulasComent = formulasComent.flat();
    // formulasComent.forEach(f => { Logger.log(`${f.slice(1) != formulaComent} | |||${f}||| = |||${formulaComent}|||`); });
    formulasComent.forEach(f => { if (f.slice(1) != formulaComent) { isMmodified = true } });
    if (isMmodified) {
      rangeComent.setFormulaR1C1(formulaComent);
    }

    isMmodified = false;
    let formulaTop = getContext().getFormulaR1C1(this.sheet.getName(), this.rowBodyFirst, this.columnTopFormula);
    let rangeTop = this.sheet.getRange(this.rowBodyFirst, this.columnTopFormula, this.rowBodyLast - this.rowBodyFirst + 1, 1);
    let formulasTop = rangeTop.getFormulasR1C1();

    formulasTop = formulasTop.flat();
    // formulasTop.forEach(f => { Logger.log(`${f!= formulaTop} | \n|||${f}||| = \n|||${formulaTop}|||`); });
    formulasTop.forEach(f => {
      if (f != "=" + formulaTop) {
        // if (f.replace(/\\/g, '') != formulaTop.replace(/\\/g, '')) {
        // Logger.log("-------------------------")
        // Logger.log(f)
        // Logger.log(formulaTop)
        isMmodified = true
      }
    });

    if (isMmodified) {
      max_col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheet.getSheetName()).getMaxColumns();
      if (max_col != getClassColRow().list1_col_MAX_COL) {
        Logger.log("ClassSheet_List1 fixSheet колличество столбцов разное " + `
      list1_col_MAX_COL=${getClassColRow().list1_col_MAX_COL} | max_col=${max_col}`);
        return;
      } else {
        Logger.log("ClassSheet_List1 fixSheet колличество столбцов одинаковое " + `
      list1_col_MAX_COL=${getClassColRow().list1_col_MAX_COL} | max_col=${max_col}`);
      }

      rangeTop.setFormulaR1C1(formulaTop);
    }




    isMmodified = false;
    let formulaТопЭксперта = getContext().getFormulaR1C1(this.sheet.getName(), this.rowBodyFirst, getClassColRow().list1_col_ТопЭксперта);
    let rangeТопЭксперта = this.sheet.getRange(this.rowBodyFirst, getClassColRow().list1_col_ТопЭксперта, this.rowBodyLast - this.rowBodyFirst + 1, 1);
    rangeТопЭксперта.getFormulasR1C1().flat().forEach(f => { if (f != "=" + formulaТопЭксперта) { isMmodified = true } });
    if (isMmodified) { rangeТопЭксперта.setFormulaR1C1(formulaТопЭксперта); }

    isMmodified = false;
    let formulaКомментарийЭксперта = getContext().getFormulaR1C1(this.sheet.getName(), this.rowBodyFirst, getClassColRow().list1_col_КомментарийЭксперта);
    let rangeКомментарийЭксперта = this.sheet.getRange(this.rowBodyFirst, getClassColRow().list1_col_КомментарийЭксперта, this.rowBodyLast - this.rowBodyFirst + 1, 1);
    rangeКомментарийЭксперта.getFormulasR1C1().flat().forEach(f => { if (f != "=" + formulaКомментарийЭксперта) { isMmodified = true } });
    if (isMmodified) { rangeКомментарийЭксперта.setFormulaR1C1(formulaКомментарийЭксперта); }




    if (this.columnАктуальный_Сбор_КП) {
      Logger.log(`this.columnАктуальный_Сбор_КП=${this.columnАктуальный_Сбор_КП}`);
      try {
        this.addColForСбор_КП_внешней();
      } catch (err) {

        Logger.log(`this.columnАктуальный_Сбор_КП=${this.columnАктуальный_Сбор_КП} this.columnLastPrice=${this.columnLastPrice} `);
        Logger.log(`this.columnАктуальный_Сбор_КП=Ошибка`);
        // Logger.log(mrErrToString)
        // this.addColForСбор_КП_внешней_ok = true;
        mrErrToString(err);
      }

    }

    // this.recalculate_report_malo();


    if (getContext().timeConstruct.getHours() == 1) { this.fixSheetCompact(); }
    if(this.needAddEmtyColsBeforeАктуальный_Сбор_КП()){
      this.addEmtyColsBeforeАктуальный_Сбор_КП();
    }
  }

  fixSheetCompact() {
    // if (getContext().timeConstruct.getHours() != 1) { return }

    // let columnLast = this.sheet.getMaxColumns();
    // this.sheet.hideColumn(this.sheet.getRange(1, this.columnComent, 1, columnLast - this.columnComent + 1))

    let wp = this.sheet.getColumnWidth(this.columnFirstPrice + 1);
    let wd = this.sheet.getColumnWidth(this.columnFirstPrice + 2);
    let wm = this.sheet.getColumnWidth(this.columnFirstPrice + 3);

    let f = this.columnLastPrice - this.columnFirstPrice - 2 + 9;
    for (let i = 0; i < f; i++) {
      let col = i + this.columnFirstPrice + 1;
      let w = wp;
      switch (i % 3) {
        case 0: w = wp; break;
        case 1: w = wd; break;
        case 2: w = wm; break;
        default: w = wp;
      }
      this.sheet.setColumnWidth(col, w)
    }
  }



  recalculate_report_malo() {


    Logger.log("&&&& recalculate_report_malo START &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");

    try {
      if (!this.sheet.getRange(1, this.columnКол_во + 1).getFormula()) {
        let vls_report = this.report();
        // Logger.log(`vls_report=${vls_report}`);



        if (vls_report.length > 0) {
          let info_arr = new Array(vls_report[0].length);
          info_arr[0] = new Date();
          info_arr[1] = "Меню > Email рассылка > Пересчитать Цена Итого Предложений.\nЕсли дата меньше текущей более чем на пол часа, либо поля 'Цена', 'Итого' и 'Предложений' пусты или вы считаете, что они устарели, то выполните команду из меню.";

          // let url = getContext().getSpreadSheet().getUrl();
          // Logger.log(`recalculate_report_malo URL ${url}`);
          // if (  url == "https://docs.google.com/spreadsheets/d/1t5Cz7TGeg/edit" ){
          vls_report.push(info_arr);
          // }
          this.sheet.getRange(2, this.columnКол_во + 1, vls_report.length, vls_report[0].length).setValues(vls_report);
        }
      }

    } catch (err) {
      Logger.log(`vls_report this.columnКол_во =${this.columnКол_во}`);
      Logger.log(`vls_report=Ошибка`);
      // Logger.log(mrErrToString)
      mrErrToString(err);
    }

    try {

      // Logger.log("&&&& vls_malo &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
      let maloB8 = this.sheet.getRange("B8").getValue();

      let vls_malo = this.sheet.getRange(this.rowBodyFirst, this.columnПредложений, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();
      let vls_name = this.sheet.getRange(this.rowBodyFirst, this.columnProduct, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();
      // Logger.log(vls_malo);
      // malo(row, maloB8);
      let mapMalo = new Map();
      vls_malo = vls_malo.map((v, i, arr) => {
        if (!vls_name[i][0]) { return [undefined]; }
        let row = i + this.rowBodyFirst;
        let ret = malo(row, maloB8);
        mapMalo.set(row, { row, malo: ret, name: vls_name[i][0] });
        return [ret];//*
      });

      // Logger.log(vls_malo);
      // Logger.log("vls_malo | " + vls_malo);
      this.sheet.getRange(this.rowBodyFirst, this.columnПредложений, this.rowBodyLast - this.rowBodyFirst + 1, 1).setValues(vls_malo);
      getContext().getSheetПроблемы().setMaloOtvetov(mapMalo);
      // Logger.log("&&&& vls_malo &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");

    } catch (err) {
      // Logger.log(`vls_malo this.columnПредложений =${this.columnПредложений}`);
      Logger.log(`vls_malo=Ошибка`);
      // Logger.log(mrErrToString)
      mrErrToString(err);
    }

    Logger.log("&&&& recalculate_report_malo FINISH &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");

  }

  addColForСбор_КП_внешней() {
    if (this.addColForСбор_КП_внешней_ok) { return; }

    // let formula_str = this.sheet.getRange(1, this.columnАктуальный_Сбор_КП + 1, 1, 1).getValue();
    // formula_str = fl_str(formula_str);
    let { retArr, retBgArr } = Сбор_КП_внешней();
    let vls = retArr;
    let vls_bg = retBgArr;

    // Logger.log("retArr | " + retArr);
    // Logger.log("retBgArr | " + retBgArr);

    let necessary = vls[0].length;
    let have = this.columnLastPrice - this.columnАктуальный_Сбор_КП;

    // let howMany = (Number.parseInt((necessary - have) / 2) + 1) * 2;
    let howMany = necessary - have;
    Logger.log(`necessary=${necessary} have=${have} add=${howMany}`);

    // if ((howMany + necessary)%2 ==0) {}

    if (howMany > 0) {
      if (howMany < 0) {
        throw new Error(`Невозможно добавить отрецательное количевство строк howMany=${howMany}`);
      }
      // this.sheet.insertColumnsAfter(this.columnАктуальный_Сбор_КП, howMany);
      this.sheet.insertColumnsBefore(this.columnLastPrice, howMany);
      // let formula_footer = this.sheet.getRange(this.rowBodyLast + 1, this.columnАктуальный_Сбор_КП - 1).getFormulaR1C1();
      // this.sheet.getRange(this.rowBodyLast + 1, this.columnАктуальный_Сбор_КП, 1, (this.columnLastPrice - 1) - this.columnАктуальный_Сбор_КП + 1).setFormulaR1C1(formula_footer);
      // Logger.log(`formula_footer=${formula_footer} `);
      getContext().sheetList1 = undefined;
      SpreadsheetApp.flush();
      return;
    }

    this.sheet.getRange(1, this.columnАктуальный_Сбор_КП, vls.length, vls[0].length).setValues(vls);
    this.sheet.getRange(1, this.columnАктуальный_Сбор_КП, vls_bg.length, vls_bg[0].length).setBackgrounds(vls_bg);

    this.addColForСбор_КП_внешней_ok = true;
    SpreadsheetApp.flush();
  }


  needAddEmtyColsBeforeАктуальный_Сбор_КП() {
    let vl = this.sheet.getRange(1, 1, 1, this.columnLastPrice).getValues()[0];
    vl = ["col"].concat(vl);

    let lastEmail = ((arr, colBefore, colLastPrice) => {

      if (!colBefore) { colBefore = colLastPrice; }
      let res = colBefore - 1;
      for (; res >= this.columnFirstPrice; res--) {
        if (arr[res]) { break; }
      }
      return res;
    })(vl, this.columnАктуальный_Сбор_КП - 3, this.columnLastPrice);


    this.minEmptyCol = 15;
    let res = ((this.columnАктуальный_Сбор_КП ? this.columnАктуальный_Сбор_КП : this.columnLastPrice) - lastEmail) < this.minEmptyCol;
    // Logger.log(JSON.stringify(
    //   {
    //     lastEmail,
    //     columnАктуальный_Сбор_КП: this.columnАктуальный_Сбор_КП,
    //     columnLastPrice: this.columnLastPrice,
    //     columnFirstPrice: this.columnFirstPrice,
    //     res,
    //   }, null, 2))
    return res;
  }

  addEmtyColsBeforeАктуальный_Сбор_КП(numCols=60) {
    
    // return;
    let vl = this.sheet.getRange(1, 1, 1, this.columnLastPrice).getValues()[0];
    vl = ["col"].concat(vl);


    let colAddBefore = ((arr, colBefore, colLastPrice) => {
      let theRes = arr.indexOf(getSettings().sheetName_Актуальный_Сбор_КП);
      if (theRes < 0) { theRes = colLastPrice; }
      return theRes;
    })(vl, this.columnАктуальный_Сбор_КП, this.columnLastPrice);

    numCols = ((num, d) => {
      if ((num % d) != 0) { num = num - (num % d) }
      return num;
    })(numCols, 3);

    this.sheet.insertColumnsBefore(colAddBefore, numCols);
    this.fixSheetCompact();
  }

  onEdit(edit) {
    // Получаем диапазон ячеек, в которых произошли изменения
    let range = edit.range;

    // Лист, на котором производились изменения
    let sheet = range.getSheet();

    // Проверяем, нужный ли это нам лист
    Logger.log("onEditList_1");
    // if (sheet.getName() != 'Лист1') {
    if (sheet.getName() != getSettings().sheetName_Лист_1) {
      return false;
    }


    // раскраска поставциков в первой строке лист1
    if (range.rowStart == 1) {
      let cList1 = getContext().getSheetList1();
      cList1.updateBackgroundAll();
    }

    // раскраска листа7
    if ((range.columnStart <= this.columnLastPrice) && (this.columnM <= range.columnEnd)) {
      try {
        let cList7 = getContext().getSheetList7();
        cList7.updateAssociatedRowsWith_List1(range.rowStart, range.rowEnd);
      } catch (err) { mrErrToString(err); }
    }



    // востанавливаем формулы в столбец HY лист1
    if ((range.columnStart <= this.columnTopFormula) && (this.columnTopFormula - 1 <= range.columnEnd)) {
      let rowStart = (range.rowStart == 1 ? 2 : range.rowStart);
      let rowEnd = range.rowEnd;
      let col = this.columnTopFormula;

      if (rowStart < this.rowBodyFirst) { rowStart = this.rowBodyFirst; }
      if (rowEnd > this.rowBodyLast) { rowEnd = this.rowBodyLast; }

      if (rowEnd - rowStart + 1 > 0) {
        sheet.getRange(rowStart, col, rowEnd - rowStart + 1, 1).setFormulaR1C1(getContext().getFormulaR1C1(sheet.getName(), rowStart, this.columnTopFormula));
        sheet.getRange(rowStart, col - 1, rowEnd - rowStart + 1, 1).setFormulaR1C1(this.formulaR1C1_ForComments());
      } else {
        Logger.log(`диапозон не входит в зону востановления формул`);
      }

      // for (let row = rowStart; row <= rowEnd; row++) {
      //   if (row <= 1) { continue; }
      //   if (row > this.rowBodyLast) { break; } // 25.02.2021 MrVova
      //   sheet.getRange(row, col).setFormulaR1C1(getContext().getFormulaR1C1(sheet.getName(), row, this.columnTopFormula));
      //   sheet.getRange(row, col - 1).setFormulaR1C1(this.formulaR1C1_ForComments());

      //   // sheet.getRange(row, col).setFormulaR1C1(getContext().getFormulaR1C1(sheet.getName(), row, col));
      // }

    }



  }



  findRowBodyLast() {
    let ret = getContext().getRowSobachkaBySheetName(this.sheet.getSheetName());
    if (!ret) { ret = this.sheet.getLastRow(); }
    else { ret = ret - 1; }
    return ret;
  }




  getProductColumnArr() {
    if (!this.productColumnArr) {
      this.emailColumnArr = new Array();
      let vl = this.sheet.getRange(this.rowBodyFirst, this.columnProduct, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();
      this.productColumnArr = new Array(vl.length);

      for (let i = 0; i < vl.length; i++) {
        this.productColumnArr[i] = fl_str(vl[i][0]);
      }
    }

    return this.productColumnArr;
  }


  getProductGroupColumnArr() {
    if (!this.productGroupColumnArr) {
      let vl = this.sheet.getRange(this.rowBodyFirst, this.columnGroup, this.rowBodyLast - this.rowBodyFirst + 1, 1).getValues();
      this.productGroupColumnArr = new Array(vl.length);

      for (let i = 0; i < vl.length; i++) {
        if (isError(vl[i][0])) { continue; }
        if (!vl[i][0]) { continue; }
        this.productGroupColumnArr[i] = fl_str(vl[i][0]);
      }
    }
    return this.productGroupColumnArr;
  }





  getEmailRowArr() {
    if (!this.emailRowArr) {
      let vl = this.sheet.getRange(1, this.columnFirstPrice, 1, this.columnLastPrice - this.columnFirstPrice + 1).getValues();
      // Logger.log(`ClassSheet_List1 getInfoForCounteragent vl=${vl}`);
      this.emailRowArr = new Array(vl[0].length);
      for (let i = 0; i < vl[0].length; i++) {
        if (!isEmail(vl[0][i])) { continue; }
        this.emailRowArr[i] = fl_str(vl[0][i]);
      }
    }

    // Logger.log(`ClassSheet_List1 getInfoForCounteragent this.emailRowArr=${this.emailRowArr}`);
    return this.emailRowArr;
  }




  getInfoForCounteragent(email) {
    // Logger.log(`ClassSheet_List1 getInfoForCounteragent email=${email}`);


    let ret = new Object();
    ret["productGroupArr"] = ["нет"];
    ret["productArr"] = ["нет"];
    ret["isEmailResponse"] = -1;

    // Logger.log(`ClassSheet_List1 getEmailProductGroupMap ${this.getEmailProductGroupMap().size}`);
    // this.getEmailProductGroupMap().forEach(logMap);
    // Logger.log(`ClassSheet_List1 getEmailProductMap ${this.getEmailProductMap().size}`);
    // this.getEmailProductMap().forEach(logMap);
    // Logger.log(`ClassSheet_List1 getEmailResponseMap ${this.getEmailResponseMap().size}`);
    // this.getEmailResponseMap().forEach(logMap);

    if (!email) { return ret; }

    // Logger.log(`ClassSheet_List1 getInfoForCounteragent email=${email}`);

    let productGroupArr = this.getEmailProductGroupMap().get(fl_str(email));

    // Logger.log(`ClassSheet_List1 getInfoForCounteragent productGroupArr=${productGroupArr}`);

    // ret["productGroupArr"] = this.getEmailProductGroupMap().get(fl_str(email));
    ret["productGroupArr"] = productGroupArr;
    ret["productArr"] = this.getEmailProductMap().get(fl_str(email));
    ret["isEmailResponse"] = this.getEmailResponseMap().get(fl_str(email));
    // Logger.log(`======ClassSheet_List1 getInfoForCounteragent \nret.productGroupArr=${ret.productGroupArr}  \nret.productArr=${ret.productArr}  \nret.isEmailResponse=${ret.isEmailResponse}`);
    return ret;
  }


  getEmailResponseMap() {
    if (!this.emailResponseMap) { this.emailResponseMap = this.makeEmailResponseMap(); }
    return this.emailResponseMap;

  }

  getEmailProductGroupMap() {
    if (!this.emailProductGroupMap) { this.emailProductGroupMap = this.makeEmailProductGroupMap(); }
    return this.emailProductGroupMap;

  }
  getEmailProductMap() {

    if (!this.emailProductMap) { this.emailProductMap = this.makeEmailProductMap(); }
    return this.emailProductMap;


  }



  getAnalogMap() {
    if (!this.analogMap) { this.analogMap = this.makeAnalogMap(); }
    return this.analogMap;
  }

  getAnalogArrLink() {
    if (!this.analogArrLink) { this.analogArrLink = this.makeAnalogArrLink(); }
    return this.analogArrLink;
  }

  getAnalogEmailArr() {
    let retArr = new Array();
    let range = this.sheet.getRange(1, this.columnFirstPrice, 1, this.columnLastPrice - this.columnFirstPrice + 1);
    let vl = range.getValues();
    let bg = range.getBackgrounds();

    let analogColor = getContext().getAnalogColor();

    for (let col = 0; col < vl[0].length; col++) {
      if (bg[0][col] != analogColor) { continue; }
      if (!vl[0][col]) { continue; }

      retArr.push(fl_str(vl[0][col]));

    }
    return retArr;
  }


  makeEmailResponseMap() {
    let retMap = new Map();

    let emailArr = this.getEmailRowArr();

    for (let i = 0; i < emailArr.length; i++) {
      if (!emailArr[i]) { continue; }
      if (!retMap.has(emailArr[i])) { retMap.set(emailArr[i], false); }
    }
    let vl = this.sheet.getRange(this.rowBodyFirst, this.columnFirstPrice, this.rowBodyLast - this.rowBodyFirst + 1, this.columnLastPrice - this.columnFirstPrice + 1).getValues();
    for (let j = 0; j < vl[0].length; j++) {
      if (!emailArr[j]) { continue; }
      if (!isEmail(emailArr[j])) { continue; }
      for (let i = 0; i < vl.length; i++) {
        if (!vl[i][j]) { continue; }
        retMap.set(emailArr[j], true);
        break;
      }
    }
    return retMap;

  }

  makeEmailProductGroupMap() {
    let retMap = new Map();

    let emailArr = this.getEmailRowArr();
    let prGrArr = this.getProductGroupColumnArr();

    for (let i = 0; i < emailArr.length; i++) {
      if (!emailArr[i]) { continue; }
      if (!retMap.has(emailArr[i])) { retMap.set(emailArr[i], new Array()); }
    }

    let vl = this.sheet.getRange(this.rowBodyFirst, this.columnFirstPrice, this.rowBodyLast - this.rowBodyFirst + 1, this.columnLastPrice - this.columnFirstPrice + 1).getValues();
    for (let i = 0; i < vl.length; i++) {
      for (let j = 0; j < vl[i].length; j++) {
        if (!emailArr[j]) { continue; }
        if (!isEmail(emailArr[j])) { continue; }
        if (!vl[i][j]) { continue; }

        if (!retMap.get(emailArr[j]).includes(prGrArr[i])) {
          retMap.get(emailArr[j]).push(prGrArr[i]);
        }
      }
    }

    // Logger.log("makeEmailProductGroupMap__________________________");
    // retMap.forEach(logMap);
    // Logger.log("makeEmailProductGroupMap^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^");
    return retMap;
  }


  makeEmailProductMap() {
    let retMap = new Map();

    let emailArr = this.getEmailRowArr();


    let prArr = this.getProductColumnArr();

    for (let i = 0; i < emailArr.length; i++) {
      if (!emailArr[i]) { continue; }
      if (!retMap.has(emailArr[i])) { retMap.set(emailArr[i], new Array()); }
    }

    let vl = this.sheet.getRange(this.rowBodyFirst, this.columnFirstPrice, this.rowBodyLast - this.rowBodyFirst + 1, this.columnLastPrice - this.columnFirstPrice + 1).getValues();
    for (let i = 0; i < vl.length; i++) {
      for (let j = 0; j < vl[i].length; j++) {
        if (!emailArr[j]) { continue; }
        if (!isEmail(emailArr[j])) { continue; }
        if (!vl[i][j]) { continue; }
        if (!retMap.get(emailArr[j]).includes(prArr[i])) {
          retMap.get(emailArr[j]).push(prArr[i]);
        }
      }
    }
    return retMap;
  }


  makeAnalogArrLink() {
    let retArr = new Array();
    let range = this.sheet.getRange(1, this.columnFirstPrice, 1, this.columnLastPrice - this.columnFirstPrice + 1);
    let vl = range.getValues();
    let bg = range.getBackgrounds();

    let analogColor = getContext().getAnalogColor();

    for (let col = 0; col < vl[0].length; col++) {
      if (bg[0][col] != analogColor) { continue; }
      let link = `$${nc(col + this.columnFirstPrice)}$${1}`;
      retArr.push(link);

    }
    return retArr;
  }


  makeAnalogMap() {
    let retMap = new Map();
    let range = this.sheet.getRange(1, this.columnFirstPrice, 1, this.columnLastPrice - this.columnFirstPrice + 1);
    let vl = range.getValues();
    let bg = range.getBackgrounds();

    let analogColor = getContext().getAnalogColor();

    for (let col = 0; col < vl[0].length; col++) {
      if (bg[0][col] != analogColor) { continue; }
      if (!vl[0][col]) { continue; }
      let email = fl_str(vl[0][col]);
      let analog = retMap.get(email);

      if (!analog) { analog = new Analog(email); }


      let arInfo = this.sheet.getRange(1, this.columnFirstPrice + col, this.sheet.getLastRow(), 2).getValues();
      analog.addInfo(this.columnFirstPrice + col, arInfo);

      retMap.set(analog.getEmail(), analog);

    }
    return retMap;
  }



  updateBackgroundAll() {

    let paintRange = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn());
    paintRange.setBackgrounds(this.paintRangeEmail(paintRange));



  }

  paintRangeEmail(paintRange) {

    let bg = paintRange.getBackgrounds();
    let vl = paintRange.getValues();
    let analogColor = getContext().getAnalogColor();
    for (let row = 0; row < bg.length; row++) {
      for (let col = 0; col < bg[row].length; col++) {

        if (bg[row][col] == analogColor) { continue }

        let email = fl_str(vl[row][col]);

        if (!isEmail(email)) { continue }
        let color = getContext().getCounteragentColor(fl_str(email));
        if (!color) continue;
        bg[row][col] = color;

      }
    }
    return bg;
  }


  // =ЕСЛИОШИБКА(  TEXTJOIN("|||";ЛОЖЬ ; ARRAY_CONSTRAIN( SORT( ARRAYFORMULA(Split(ТРАНСП (Split(TEXTJOIN("|||";ИСТИНА;'Лист1'!$IB4:$IH4);"|||";ЛОЖЬ;ЛОЖЬ));"|";ЛОЖЬ;ЛОЖЬ));3;ЛОЖЬ );100;2));)

  formulaR1C1_ForComments() {
    let c236 = `C${getClassColRow().list1_col_CommentBlokStart}`;
    let c242 = `C${getClassColRow().list1_col_CommentBlokEnd}`;

    return `
    IFERROR(  
      TEXTJOIN("|||";FALSE ; 
        ARRAY_CONSTRAIN( SORT( ARRAYFORMULA(Split(TRANSPOSE 
          (Split(TEXTJOIN("|||";TRUE;'${getSettings().sheetName_Лист_1}'!R[0]${c236}:R[0]${c242});"|||";FALSE;FALSE)
        );"|";FALSE;FALSE));3;FALSE );100;2)
      )
    ;)`

  }




  numberOfOffers(row, maloB8) {
    let colorAnalog = getContext().getAnalogColor();
    // let colFirst = "S";
    // let colLast = "HP";
    let colFirst = nc(getClassColRow().list1_col_FirstPrice);
    let colLast = nc(getClassColRow().list1_col_LastPrice);

    let rHead = `$${colFirst}$${1}:$${colLast}$${1}`;
    let rBg = `$${colFirst}$${1}:$${colLast}$${1}`;
    let rPrice = `$${colFirst}$${row}:$${colLast}$${row}`;

    return this.countCellsWithBackgroundColor(maloB8, colorAnalog, rHead, rBg, rPrice);
  }

  countCellsWithBackgroundColor(maloB8 = 3, colorAnalog = "#ffff00", rHead, rBg, rPrice) {
    // let sheetName = "Лист1";
    // let sheetName = getSettings().sheetName_Лист_1;
    // var ss = SpreadsheetApp.getActive();
    // var sheet = ss.getSheetByName(sheetName);
    var sheet = this.sheet

    var vlPrice = sheet.getRange(rPrice).getValues();
    var vlBG = sheet.getRange(rBg).getBackgrounds();
    var vlHead = sheet.getRange(rHead).getValues();

    var priceAll = 0;
    var priceAnalog = 0;

    for (var i = 0; i < vlPrice.length; i++) {
      for (var j = 0; j < vlPrice[i].length; j++) {

        if (!vlHead[i][j]) { continue; }  // в шапке пусто  пропускаем
        if (!vlPrice[i][j]) { continue; } // цена пусто пропускаем // нет значения 
        priceAll++;
        if (vlBG[i][j] == colorAnalog) {
          priceAnalog++;
        }
      }
    }
    return { priceAll, priceAnalog };
    // var retStr = `${priceAll}`;
    // if (priceAll <= maloB8) { retStr = `Мало, всего ${retStr}`; }
    // if (priceAnalog != 0) { retStr = `${retStr}, из них аналог ${priceAnalog}` }

    // return retStr;

  }


  setModelForContragent(sAg, row, model) {
    // Logger.log(` Модель Agent=  ${sAg} | row=${row} Coment= ${model}`);
    let r = this.sheet.getRange(sAg);
    // r.offset(row, 3).setNote(model);
    // Logger.log(` Модель Agent=  ${sAg} | row=${row} Coment= ${model} | ${r.getA1Notation()} | ${r.offset(row - 1, 2).getA1Notation()}`);

    r.offset(row - 1, 2).setValue(model);

  }





}



class Analog {
  constructor(email) {
    this.email = fl_str(email);
    this.map = new Map();

  }
  addInfo(col, arInfo) {
    this.map.set(col, arInfo);
  }

  getEmail() {
    return this.email;
  }

  toString() {
    return this.email;
  }

}


