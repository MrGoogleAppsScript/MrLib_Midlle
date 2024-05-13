/**  справочник цветов  */
// MisterVova@mail.ru 23.12.2020 добавил файл  
class ColorMap {
  constructor(nameSet = "MrColorMap") { // class constructor
    this.nameSet = nameSet; // имя набора цветов 
    this.def = "#ffffff";
    this.colorMap = this.makeColorMap();
    this.lastUpdate = new Date();

  }


  upDate() {
    this.colorMap = this.makeColorMap();
    this.lastUpdate = new Date();
  }

  getOfficialColor() {
    return this.getColor("Официал");
  }
  getIsFromProductListColor() {
    return this.getColor("E-mail из базы товаров");
  }

  getAnalogColor() {
    return this.getColor("Аналог");
  }



  getColor(colorKey) {
    if (!colorKey) return this.def;
    colorKey = fl_str(colorKey);
    if (!this.colorMap.has(colorKey)) { return this.def; }
    return this.colorMap.get(colorKey);
  }

  get(colorKey) {
    return this.getColor(colorKey);
  }


  logArray() {

    let ret = new Array();
    for (var [key, value] of this.colorMap) {
      ret.push([key, value]);
    }
    return ret;

  }

  makeColorMap() {
    let retColorMap = new Map();
    retColorMap.set(fl_str("DEFAULT"), this.def);
    retColorMap.set(fl_str("ПодходитДа"), "#00ff00");
    retColorMap.set(fl_str("ПодходитНет"),"#ff0000");

    let ss = SpreadsheetApp.getActive();
    // let sheetSetings = ss.getSheetByName("Настройки писем");
    let sheetSetings = ss.getSheetByName(getSettings().sheetName_Настройки_писем);


    let colorNames = sheetSetings.getRange(2, nr("J"), sheetSetings.getLastRow(), 1).getValues();
    let colorBG = sheetSetings.getRange(2, nr("K"), sheetSetings.getLastRow(), 1).getBackgrounds();


    for (let i = 0; i < colorNames.length; i++) {
      if (!colorNames[i][0]) {
        continue;
      }
      let colorStr = fl_str(colorNames[i][0]);
      if (retColorMap.has(colorStr)) { continue; }
      retColorMap.set(colorStr, colorBG[i][0]);

    }

    colorNames = sheetSetings.getRange(2, nr("V"), sheetSetings.getLastRow(), 1).getValues();
    colorBG = sheetSetings.getRange(2, nr("v"), sheetSetings.getLastRow(), 1).getBackgrounds();

    for (let i = 0; i < colorNames.length; i++) {
      if (!colorNames[i][0]) { continue; }
      let colorStr = fl_str(colorNames[i]);
      if (retColorMap.has(colorStr)) { continue; }
      retColorMap.set(colorStr, colorBG[i][0]);
    }
    try {
      // цвета для топ 7 
      let colorNamesTop = [
        "top_1",
        "top_2",
        "top_3",
        "top_4",
        "top_5",
        "top_6",
        "top_7",
      ]
      let cList7 = new ClassSheet_List7();
      let sheet_List7 = cList7.sheet;
      colorBG = sheet_List7.getRange(1, cList7.columnTopFirst, 1, cList7.columnTopLast - cList7.columnTopFirst + 1).getBackgrounds();
      for (let i = 0; i < colorNamesTop.length; i++) {
        if (!colorNamesTop[i]) { continue; }
        let colorStr = fl_str(colorNamesTop[i]);
        if (retColorMap.has(colorStr)) { continue; }
        retColorMap.set(colorStr, colorBG[0][i * cList7.cellsNum + 1]);
      }
    } catch { }

    // retColorMap.forEach((v, k) => Logger.log( [v, k]));
    // Logger.log(JSON.stringify([[...retColorMap.values()], [...retColorMap.keys()]]))
    return retColorMap;
  }

}
