// 20.02.2021 MrVova
class Sobachka {
  // Название листа//	Столбец для поиска пустых строк //	@ Столбец с @	 // Отсавить на листе строки с первой по [значение] включительно//

  constructor(context, sheetName, colEmpty, colSobachka, rowBodyMinimum) {
    if (!context) { throw new Error("context"); }
    this.context = context;
    this.sheetName = sheetName;
    this.colEmpty = colEmpty;
    this.colSobachka = colSobachka;
    this.rowBodyMinimum = rowBodyMinimum;
    this.rowSobachka = undefined;
  }

  getRowSobachka() {
    if (!this.rowSobachka) { this.rowSobachka = this.findRowSobachka() }
    return this.rowSobachka;
  }

  // поиск Строки тотал  20.02.2021
  findRowSobachka() {
    let sheet = this.context.getSheetByName(this.sheetName);

    let rowSobachka = 0; // @ в строке с номером rowSobachka 
    let vcs = sheet.getRange(1, this.colSobachka, sheet.getLastRow(), 1).getValues();
    for (; rowSobachka < vcs.length; rowSobachka++) {
      if (vcs[rowSobachka].includes("@")) { break; }
    }
    rowSobachka++; // @ в строке с номером rowSobachka 
    return rowSobachka;
  }

}


class SobachkaMap {
  // constructor(sheetNameSettings = "Настройки писем", rowStart = 2, colStart = nr("X")) {  // имя листа настроек // левый верхний угол  начало строк 
  constructor(context, sheetNameSettings = getSettings().sheetNames.Настройки_писем, rowStart = 2, colStart = nr("X")) {  // имя листа настроек // левый верхний угол  начало строк 
    if (!context) { throw new Error("context"); }
    this.context = context;
    this.sheetNameSettings = fl_str(sheetNameSettings);
    this.rowStart = rowStart;
    this.colStart = colStart;
    this.map = this.makeMap();
  }


  getRowSobachkaBySheetName(sheetName) {
    let valueSobachka = this.map.get(fl_str(sheetName));
    if (!valueSobachka) { return undefined; }
    if (!(valueSobachka instanceof Sobachka)) { return undefined; }
    return valueSobachka.getRowSobachka();
  }


  makeMap() {
    let retMap = new Map();
    let sheetSetings = this.context.getSheetByName(this.sheetNameSettings);
    // let ui = SpreadsheetApp.getUi();

    let vlsSetings = sheetSetings.getRange(this.rowStart, this.colStart, sheetSetings.getLastRow() - this.rowStart + 1, 4).getValues();

    // console.log("setings =" + setings );
    for (let i = vlsSetings.length - 1; i >= 0; i--) {
      // console.log( "i= "+ i+ " лист= " + setings [i][0] + " колонка =" + setings [i][1] );
      if (!vlsSetings[i][0]) {
        // setings.splice(i, 1);
        continue;
      }

      if (!vlsSetings[i][1]) {
        // setings.splice(i, 1);
        continue;
      }
      if (!vlsSetings[i][2]) {
        // setings.splice(i, 1);
        continue;
      }


      if (!vlsSetings[i][3]) {
        // setings.splice(i, 1);
        continue;
      }

      if (!parseInt(vlsSetings[i][3].toString(), 10)) {
        // setings.splice(i, 1);
        continue;
      }

      let sheetName = fl_str(vlsSetings[i][0]);
      let colEmpty = nr(vlsSetings[i][1])
      let colSobachka = nr(vlsSetings[i][2]);
      let rowBodyMinimum = parseInt(vlsSetings[i][3].toString(), 10);
      rowBodyMinimum = (rowBodyMinimum < 1 ? 1 : rowBodyMinimum);
      retMap.set(sheetName, new Sobachka(this.context, sheetName, colEmpty, colSobachka, rowBodyMinimum));

    }
    return retMap;
  }
  toArr() {
    let retArr = new Array();
    retArr.push(["Название листа КЕУ", "Название листа", "Столбец для поиска пустых строк", "@ Столбец с @", "Отсавить на листе строки с первой по [значение] включительно", "строка c @"]);

    for (let [key, value] of this.map) {
      retArr.push([key, value.sheetName, value.colEmpty, value.colSobachka, value.rowBodyMinimum, value.getRowSobachka()]);
    }

    return retArr;
  }

}

