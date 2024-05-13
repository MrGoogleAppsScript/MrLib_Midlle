// MisterVova@mail.ru 23.12.2020 добавил файл  
class CounteragentMap {
  constructor(nameSet = "MrCounteragentMap") { // class constructor

    this.nameSet = nameSet;
    this.counteragentMap = undefined;
    this.lastUpdate = new Date();
    this.upDate();
  }


  getCounteragent(email) {
    if (!email) { return undefined; }
    email = fl_str(email);
    if (!this.counteragentMap.has(email)) return new Counteragent(email);
    return this.counteragentMap.get(email);
  }

  upDate() {
    this.counteragentMap = this.makeCounteragentMap();
    this.lastUpdate = new Date();
  }





  makeCounteragentMap() {
    // Logger.log("makeCounteragentMap");
    let ss = SpreadsheetApp.getActive();
    // let sheetReviews = ss.getSheetByName("Отзывы о поставщиках");
    let sheetReviews = ss.getSheetByName(getSettings().sheetName_Отзывы_о_поставщиках);
    // let sheetProducts = ss.getSheetByName("Товары");
    let sheetProducts = ss.getSheetByName(getSettings().sheetName_Товары);

    let retMap = new Map();   //CounteragentMap

    if (sheetReviews.getLastRow() - 2 < 1) { 
      return retMap;
      throw new Error(`Лист ${sheetReviews.getSheetName()} пуст`) 
    }

    let vl = sheetReviews.getSheetValues(2, nr("B"), sheetReviews.getLastRow() - 1, 2);
    for (let row = 0; row < vl.length; row++) {

      let email = vl[row][0];
      // Logger.log("eww="+email);
      let status = vl[row][1];
      if (!email) { continue; }  // если емайл не емейл 
      email = fl_str(email);

      let cAgent = retMap.get(email);

      if (!cAgent) {
        cAgent = new Counteragent(email); // нет еще в retMap 
      } else {
        // continue; если берем статус первого вхождения
      }

      cAgent.setStatus(fl_str(status));
      retMap.set(cAgent.email, cAgent);
    }


    if (sheetProducts.getLastRow() - 2 < 1) { throw new Error(`Лист ${sheetProducts.getSheetName()} пуст`) }
    vl = sheetProducts.getSheetValues(2, nr("B"), sheetProducts.getLastRow() - 1, 2);
    // let officialColor = mrColor.getColor("Официал");
    for (let row = 0; row < vl.length; row++) {
      let email = vl[row][0];
      // Logger.log("eww="+email);
      let official = fl_str(vl[row][1]);

      if (!email) { continue; }  // если емайл не емейл 
      email = fl_str(email);

      let cAgent = retMap.get(email);
      if (!cAgent) {
        cAgent = new Counteragent(email); // нет еще в retMap 
      } else {
        // continue; если берем статус первого вхождения
      }
      cAgent.setFromProductList();
      cAgent.setOfficial(fl_str(official));
      retMap.set(cAgent.email, cAgent);
    }

    return retMap;
  }
}


// V8 runtime
class Counteragent {
  constructor(email) { // class constructor
    // if (!email) trows "";
    this.email = fl_str(email);
    this.status = getContext().getStatusMap().def; // 
    this.statusStr = this.status.name;
    this.official = false;  // Официал
    this.fromProductList = false;  // E-mail из базы товаров


    this.isEmailSent = undefined;// отправлялось письмо // undefined - не известно ; false - нет ; true - да  
    this.isEmailResponse = undefined; // получен ответ   // undefined - не известно ; false - нет ; true - да  
    this.responseProduct = undefined; //   имена Товаров для которых есть ответ;   
    this.responseProductGroup = undefined;//  имена ГруппТоваров для которых есть ответ;
  }




  collectInforFromOtherSheets() {


    // Logger.log(`Counteragent collectInforFromOtherSheets this.email =${this.email}  `);

    let sheetProductGroups = getContext().getSheetProductGroups();
    let infoFromProductGroups = sheetProductGroups.getInfoForCounteragent(this.email);
    this.isEmailSent = infoFromProductGroups["isEmailSent"];


    let sheetList1 = getContext().getSheetList1();
    let retS1 = sheetList1.getInfoForCounteragent(this.email);


    // Logger.log(`======Counteragent collectInforFromOtherSheets \nret.productGroupArr=${retS1.productGroupArr}  \nret.productArr=${retS1.productArr}  \nret.isEmailResponse=${retS1.isEmailResponse}`);



    this.isEmailResponse = retS1["isEmailResponse"];
    this.responseProduct = retS1["productArr"];
    this.responseProductGroup = retS1["productGroupArr"];


    // Logger.log(`infoFromSheetList1=${infoFromSheetList1}`);
    // Logger.log(`infoFromProductGroups=${infoFromProductGroups}`);

  }

  getIsEmailSent() { // отправлялось письмо // undefined - не известно ; false - нет ; true - да  
    if (!this.isEmailSent) {
      this.collectInforFromOtherSheets();
    }

    return this.isEmailSent;
  }

  getIsEmailResponse() {  // получен ответ   // undefined - не известно ; false - нет ; true - да  
    if (!this.isEmailResponse) {
      this.collectInforFromOtherSheets();
    }
    return this.isEmailResponse;
  }




  getResponseProduct() { //   имена Товаров для которых есть ответ;   
    if (!this.responseProduct) {
      collectInforFromOtherSheets();
    }
    return this.responseProduct;
  }

  getResponseProductGroup() {//  имена ГруппТоваров для которых есть ответ;
    if (!this.responseProductGroup) {
      collectInforFromOtherSheets();

    }
    return this.responseProductGroup;
  }




  setFromProductList() {
    this.fromProductList = true;
    return this;
  }

  isFromProductList() {
    return this.fromProductList
  }

  setEmail(email) {
    this.email = fl_str(email);
    return this;
  }
  setStatus(status) {
    this.statusStr = fl_str(status);

    this.status = getContext().getStatus(this.statusStr);
    return this;
  }
  setOfficial(official) {
    if (!official) { return this; }
    if (official instanceof Boolean) { this.official = (official || this.official); return this; }
    this.official = (fl_str(official) == fl_str("ДА"));
    return this;
  }

  canSendLetter() {
    return this.status.canSendLetter();
  }

  getEmail() { return this.email; }
  getStatus() {
    if (!this.status) { this.status = new Status() }
    return this.status;
  }


  getColor() {
    return this.getStatus().getColor()
  }


  getOfficial() { return this.official; }

  toString() { return this.email; }
  logString() { return `Counteragent(email=${this.email}, status=${this.status.logString()}  , official=${this.official}  )`; }

  isOfficial() {
    return this.official;
  }

  logToConsole() { // class method
    console.log(this.logString());

  }

  logArray() {
    let ret = Array();
    this.collectInforFromOtherSheets();

    ret.push(["email", this.email]);
    ret.push(["status", this.status.logString()]);
    ret.push(["statusStr", this.statusStr]);
    ret.push(["canSendLetter", this.canSendLetter()]);
    ret.push(["canSendLetter", this.getColor()]);
    ret.push(["official", this.official]);
    ret.push(["fromProductList", this.fromProductList]);

    ret.push(["isEmailSent", this.isEmailSent]);
    ret.push(["isEmailResponse", this.isEmailResponse]);
    ret.push(["responseProduct", this.responseProduct.toString()]);
    ret.push(["responseProductGroup", this.responseProductGroup.toString()]);

    // ret.push(["logArray", `${ret}`]);


    // Logger.log(`logArray=${ret}`);

    return ret;
  }



}

