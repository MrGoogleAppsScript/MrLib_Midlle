/**   справочник статусов поставщиков  */
// MisterVova@mail.ru 23.12.2020 добавил файл  
class Status {
  constructor(name = "Нет статуса", color = "#ffffff", notSendEmail = false) { // class constructor
    this.name = fl_str(name);
    this.color = color;
    this.notSendEmail = false;
    this.setNotSendEmail(notSendEmail);
  }
  getColor() {
    return this.color;
  }

  toString() {
    return this.name;
  }

  logString() {
    return `MrStatus(name="${this.name}", color=${this.color}  , notSendEmail=${this.notSendEmail}  )`;
  }

  setNotSendEmail(str) {

    if (str instanceof Boolean) { this.notSendEmail = (str == true); return }

    // if (!str) { this.notSendEmail = true; return}
    this.notSendEmail = fl_str(str) == fl_str("ДА");

  }


  canSendLetter() {
    return !this.notSendEmail;
  }

  logToConsole() { 
    console.log(this.logString());
  }

}


class StatusMap {
  constructor(nameSet = "MrStatusMap") { // class constructor
    this.nameSet = nameSet; // имя набора Статусов 
    this.statusMap = this.makeStatusMap();
    this.def = new Status();
    this.lastUpdate = new Date();

  }

  // canSendLetter(statusStr) {
  //   if (!statusStr) { return this.def.canSendLetter(); }
  //   if (!this.statusMap.has(statusStr)) { return this.def.canSendLetter(); }
  //   this.statusMap.get(statusStr).canSendLetter();
  // }

  upDate() {
    this.statusMap = this.makeStatusMap();
    this.lastUpdate = new Date();
  }

  getStatus(statusStr) {

    if (!statusStr) { return this.def; }
    statusStr = fl_str(statusStr);

    if (!this.statusMap.has(statusStr)) { return this.def; }
    return this.statusMap.get(fl_str(statusStr));
  }


  getColor(statusStr) {

    if (!statusStr) return this.def.getColor();

    statusStr = fl_str(statusStr);
    if (!this.statusMap.has(statusStr)) { return this.def.getColor(); }
    return this.statusMap.get(statusStr).getColor();
  }

  get(statusStr) {
    return this.getStatus(statusStr);
  }


  makeStatusMap() {
    let retMap = new Map();
    let ss = SpreadsheetApp.getActive();
    // let sheetSetings = ss.getSheetByName("Настройки писем");
    let sheetSetings = ss.getSheetByName(getSettings().sheetName_Настройки_писем);

    let names = sheetSetings.getRange(2, nr("V"), sheetSetings.getLastRow(), 1).getValues();
    let sendEmails = sheetSetings.getRange(2, nr("W"), sheetSetings.getLastRow(), 1).getValues();
    let colorBG = sheetSetings.getRange(2, nr("V"), sheetSetings.getLastRow(), 1).getBackgrounds();


    for (let i = 0; i < names.length; i++) {
      if (!names[i][0]) {
        continue;
      }
      let name = fl_str(names[i][0]);
      if (retMap.has(name)) { continue; }

      let mrStatus = new Status(fl_str(name), colorBG[i][0], fl_str(sendEmails[i][0]));
      // console.log(sendEmails[i][0]);
      // mrStatus.logToConsole();
      retMap.set(mrStatus.name, mrStatus);

    }

    return retMap;

  }

}
