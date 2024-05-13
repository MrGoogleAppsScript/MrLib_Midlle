
class MrClassSheetПаспортПроекта {

  constructor() {
    Logger.log(`MrClassSheetЗвонки constructor START`);

    this.sheet = getContext().getSheetByName(getSettings().sheetName_Паспорт_проекта);
    Logger.log(`MrClassSheetЗвонки constructor FINIS `);

    this.url = getContext().getSpreadSheet().getUrl();
    this.id = getContext().getSpreadSheet().getId();
    this.sheetName = this.sheet.getSheetName();
    this.rowConf = {
      head: { first: 1, last: 1, key: 1, },
      body: { first: 2, last: 2, },
    };
    this.projNr = getContext().getНомерПроекта();
    // this.sheetModel_Звонки = undefined;
    this.projStatus = getContext().getСтатусПроекта();

    Logger.log(JSON.stringify({ n: this.projNr, s: this.projStatus }));

    let names = {
      Статус_расчета: "Статус расчета",
      Текущий_статус_проекта: "Текущий статус проекта",
    }

  }

  getMap() {
    if (!this.map) {
      this.map = new Map();
      let vls = this.sheet.getRange("A1:D").getValues();
      vls = vls.map((v, i) => { return [1 + i].concat(v) });
      vls.forEach(v => {
        if (`${v[2]}` == "") { return; }
        let obj = {
          key: v[2],
          row: v[0],
          value: v[3],
          nr: v[1],
          comment: v[4],
        }
        this.map.set(obj.key, obj);
      });

      // this.map.forEach(item => Logger.log(JSON.stringify(item, null, 2)));
    }
    return this.map;
  }
}

