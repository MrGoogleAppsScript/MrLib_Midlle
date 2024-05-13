// MisterVova@mail.ru 23.12.2020 добавил файл  

class Product {
  constructor(name) { // class constructor
    this.name = name;
    this.nameForCompare = fl_str(name);
    this.infoMap = new Map();
  }


  toString() { return this.name; }

  logString() { return `Product(name="${this.name}", nameForCompare=${this.nameForCompare}, infoList={${this.logMapString()}\n} )`; }

  logMapString() {
    let ret = '';
    for (let [key, value] of this.infoMap) {
      ret = `${ret}\n\t  Info=("${key}" = "${value.logString()}")`;
    }
    return ret;
  }



  logMapElements(value, key, map) {
    console.log(`m[${key}] = ${value.logString()}`);
  }

  logToConsole() {
    console.log(this.logString());
  }

  addInfo(info) {
    // console.log(` ${info instanceof ProductInfo} addInfo  ${typeof info}  info =  ${info.logString()}`);
    if (!(info instanceof ProductInfo)) { return; }

    let infoHas = this.infoMap.get(info.nameForCompare);

    if (!infoHas) {
      this.infoMap.set(info.nameForCompare, info);
    } else {
      infoHas.setOfficial(info.isOfficial());
    }
  }

  getInfoArray() {
    let ret = new Array();
    for (var [key, value] of this.infoMap) {
      ret.push(value);
    }
    return ret;
  }

  getEmeils() {
    let ret = new Array();
    for (var [key, value] of this.infoMap) {
      ret.push(value.name);
    }
    return ret;
  }

  getInfoMap() { return this.infoMap; }

}



class ProductInfo {
  constructor(email, official) { // class constructor
    this.name = email;
    this.nameForCompare = fl_str(email);
    this.official = false;
    this.setOfficial(official);

    // this.logToConsole();
  }

  toString() {
    return this.name;
  }


  setOfficial(official) {

    if (this.official === true) { return; }
    if (official === true) { this.official = true; return; }
    if (fl_str(official) == fl_str("ДА")) { this.official = true; return; }
    this.official = false;

  }

  isOfficial() {
    return this.official;
  }

  logString() { return `EmailInfo(name="${this.name}", nameForCompare=${this.nameForCompare}, official=${this.isOfficial()})`; }

  logToConsole() {
    console.log(this.logString());

  }

}





class ProductMap {
  constructor(nameSet = "ProductMap") { // class constructor

    this.nameSet = nameSet;
    this.productMap = undefined;
    this.lastUpdate = new Date();
    this.upDate();
    // this.logToConsole();  // вывод на консоль результата 
  }

  logToConsole() {
    console.log(this.logString());
  }


  getProduct(name) {
    if (!name) { return undefined; }
    nameStr = fl_str(nameStr);
    if (!this.productMap.has(nameStr)) { return this.makeProduct(name); }
    return this.productMap.get(nameStr);
  }

  upDate() {
    this.productMap = this.makeMap();
    this.lastUpdate = new Date();
  }


  logString() { return `ProductMap(name="${this.nameSet}", lastUpdate=${this.lastUpdate}, \n productMap={${this.logMapString()} \n})`; }
  logMapString() {
    let ret = '';
    for (let [key, value] of this.productMap) {
      ret = `${ret}\n\t "${key}" = "${value.logString()}"`;
    }
    return ret;
  }


  makeMap() {
    // Logger.log("makeMap");
    // let ss = SpreadsheetApp.getActive();

    let sheetProducts = getContext().getSheetByName(getSettings().sheetName_Товары);
    // let sheetProducts = getContext().getSheetByName("Товары");
    // let sheetProducts = ss.getSheetByName("Товары");
    let retMap = new Map();   //ProductMap

    let lastRow = sheetProducts.getLastRow();
    if (lastRow <= 1) { // если на листе товары нет даных либо одна строчка толька 
      SpreadsheetApp.getUi().alert(`Лист "${getSettings().sheetName_Товары}" пуст.\nВ нём заполнено ${lastRow} строк.`);
      return retMap;
    }



    let vl = sheetProducts.getSheetValues(2, nr("A"), sheetProducts.getLastRow() - 2 + 1, 3);

    // let vl = sheetProducts.getSheetValues(1, nr("A"), sheetProducts.getLastRow(), 3);
    for (let row = 0; row < vl.length; row++) {
      let name = vl[row][0];
      let email = vl[row][1];
      let official = vl[row][2];

      if (!name) { continue; }
      let prod = retMap.get(fl_str(name));
      if (!prod) {
        prod = new Product(name);
        retMap.set(prod.nameForCompare, prod);
      }

      if (!official) { official = false; }

      let info = new ProductInfo(email, official);
      prod.addInfo(info);
    }

    return retMap;
  }




  makeProduct(name) {
    if (!name) { return undefined; } // нет наименования 
    let ret = new Product(name);

    for (var [keyP, product] of this.productMap) {
      if (!ret.nameForCompare.includes(product.nameForCompare)) { continue; }
      for (var [keyI, info] of product.getInfoMap()) {
        ret.addInfo(info);
      }
    }

    // ret.logToConsole();
    return ret;

  }




}