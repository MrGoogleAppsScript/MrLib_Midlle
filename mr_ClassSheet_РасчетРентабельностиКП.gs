function serviceРасчетРентабельностиКП() {
  new ClassSheet_РасчетРентабельностиКП().serviceРасчетРентабельностиКП();
}


class ClassSheet_РасчетРентабельностиКП {
  constructor() { // class constructor
    this.sheet = getContext().getSheetByName(getSettings().sheetName_РасчетРентабельностиКП);
    this.range = {
      checkBox: "L20",
      value: "L21",
      data: "L22",
    }
  }


  serviceРасчетРентабельностиКП() {

    Logger.log(`ClassSheet_Pay serviceРасчетРентабельностиКП Start`);
    if (this.sheet.getRange(this.range.checkBox).getValue() !== true) {
      return;
    }
    let rangeValue = this.sheet.getRange(this.range.value);
    if (rangeValue.getFormula() === "") {
      return;
    }

    let v = rangeValue.getValue();
    this.sheet.getRange(this.range.data).setValue(new Date());
    rangeValue.setValue(v);
    Logger.log(`ClassSheet_Pay serviceРасчетРентабельностиКП OK`);
  }

}
