class ClassColumnRow {
  constructor() { // class constructor


    this.list1_col_FirstPrice = nr("AA");
    this.list1_col_LastPrice = nr("HP");
    this.list1_col_Product = nr("D");

    this.list1_col_Group = nr("N");
    this.list1_col_Coment = nr("HX");
    this.list1_col_TopFormula = nr("IK");
    this.list1_col_S = nr("S");
    this.list1_col_M = nr("M");
    this.list1_col_id = nr("C");

    this.list1_row_Header = 1;
    this.list1_row_BodyFirst = 2;


    this.voprosi_col_N = nr("N");  // Ответ разослан
    this.voprosi_col_O = nr("O"); // Дата 

    this.voprosi_row_Header = 2;
    this.voprosi_row_BodyFirst = 3;


    this.list1_col_CommentBlokStart = nr("IB");
    this.list1_col_CommentBlokEnd = nr("IH");
    this.list1_col_Кол_во = nr("G");
    this.list1_col_Предложений = nr("J");
    this.list1_col_MAX_COL = 0;
    this.list1_col_Актуальный_Сбор_КП = undefined;
    this.list7_col_I = nr("I"); // Экспертиза



    this.list1_col_ЕдИзм = nr("F");


    // this.expertise_col_list1row = nr("CP");
    // №	Номенклатура	Ед. Изм	Кол-во

    this.expertise_row_Header = 1;
    this.expertise_row_BodyFirst = 2;

    this.make();

  }
  make() {

    this.ok = true;

    // this.sheetList1 = getContext().getSheetByName("Лист1");
    this.sheetList1 = getContext().getSheetByName(getSettings().sheetName_Лист_1);
    let vls = this.sheetList1.getSheetValues(this.list1_row_Header, 1, 1, this.sheetList1.getMaxColumns())[0];
    this.list1_col_MAX_COL = vls.length;
    // Logger.log(`vls | ${vls}`);
    for (let i = 0; i < vls.length; i++) {
      if (!vls[i]) { continue; }
      vls[i] = fl_str(vls[i]);
    }


    let col = 0;

    col = 1 + vls.indexOf(fl_str("Лучшие цены")); if (!col) { this.ok = false; } else { this.list1_col_TopFormula = col; }
    col = 1 + vls.indexOf(fl_str("Комментарий")); if (!col) { this.ok = false; } else { this.list1_col_Coment = col; }
    // col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.list1_col_FirstPrice = col + 3; }
    col = 1 + vls.indexOf(fl_str("Цены Начало")); if (!col) { this.ok = false; } else { this.list1_col_FirstPrice = col + 1; }
    col = 1 + vls.indexOf(fl_str("Цены Конец")); if (!col) { this.ok = false; } else {
      this.list1_col_LastPrice = col;
      // this.list1_col_LastPrice = (((this.list1_col_LastPrice - this.list1_col_FirstPrice) % 2) == 0 ? this.list1_col_LastPrice : this.list1_col_LastPrice - 1)

      while (((this.list1_col_LastPrice - this.list1_col_FirstPrice + 1) % 3) != 0) {
        this.list1_col_LastPrice = this.list1_col_LastPrice - 1;
      }
    }

    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.list1_col_S = col + 6; }
    col = 1 + vls.indexOf(fl_str("Цена 1")); if (!col) { this.ok = false; } else { this.list1_col_M = col; }



    col = 1 + vls.indexOf(fl_str("Номенклатура")); if (!col) { this.ok = false; } else { this.list1_col_Product = col; }
    col = 1 + vls.indexOf(fl_str("Группа")); if (!col) { this.ok = false; } else { this.list1_col_Group = col; }

    col = 1 + vls.indexOf(fl_str("№")); if (!col) { this.ok = false; } else { this.list1_col_id = col; }

    col = 1 + vls.indexOf(fl_str("CommentBlokStart")); if (!col) { this.ok = false; } else { this.list1_col_CommentBlokStart = col; }
    col = 1 + vls.indexOf(fl_str("CommentBlokEnd")); if (!col) { this.ok = false; } else { this.list1_col_CommentBlokEnd = col; }
    col = 1 + vls.indexOf(fl_str("Кол-во")); if (!col) { this.ok = false; } else { this.list1_col_Кол_во = col; }

    col = 1 + vls.indexOf(fl_str("Ед. Изм")); if (!col) { this.ok = false; } else { this.list1_col_ЕдИзм = col; }




    // Предложений  // мало
    col = 1 + vls.indexOf(fl_str("Предложений")); if (!col) { this.ok = false; } else { this.list1_col_Предложений = col; }


    col = 1 + vls.indexOf(fl_str(getSettings().sheetName_Актуальный_Сбор_КП)); if (!col) { } else { this.list1_col_Актуальный_Сбор_КП = col; } // Актуальный_Сбор_КП

    //  КомментарийЭксперта	ТопЭксперта	CommentExpertBlokStart	CommentExpertBlokEnd
    col = 1 + vls.indexOf(fl_str("КомментарийЭксперта")); if (!col) { this.ok = false; } else { this.list1_col_КомментарийЭксперта = col; }
    col = 1 + vls.indexOf(fl_str("ТопЭксперта")); if (!col) { this.ok = false; } else { this.list1_col_ТопЭксперта = col; }
    col = 1 + vls.indexOf(fl_str("CommentExpertBlokStart")); if (!col) { this.ok = false; } else { this.list1_col_CommentExpertBlokStart = col; }
    col = 1 + vls.indexOf(fl_str("CommentExpertBlokEnd")); if (!col) { this.ok = false; } else { this.list1_col_CommentExpertBlokEnd = col; }


    if (this.list1_col_Актуальный_Сбор_КП) {
      // this.list1_col_Актуальный_Сбор_КП = this.list1_col_Актуальный_Сбор_КП
      //   + (this.list1_col_Актуальный_Сбор_КП - this.list1_col_FirstPrice) % 3 + 2;
      this.list1_col_Актуальный_Сбор_КП =
        this.list1_col_Актуальный_Сбор_КП
        + 3
        - (this.list1_col_Актуальный_Сбор_КП - this.list1_col_FirstPrice) % 3;
    }



  }


  toLog() {
    return `ClassColumnRow toLog()

this.list1_col_FirstPrice=${this.list1_col_FirstPrice}
this.list1_col_LastPrice=${this.list1_col_LastPrice}
this.list1_col_Product=${this.list1_col_Product}
this.list1_col_Group=${this.list1_col_Group}
this.list1_col_Coment=${this.list1_col_Coment}
this.list1_col_TopFormula=${this.list1_col_TopFormula}
this.list1_col_S=${this.list1_col_S}
this.list1_col_M=${this.list1_col_M}

this.voprosi_col_N=${this.voprosi_col_N}
this.voprosi_col_O=${this.voprosi_col_O}
this.voprosi_row_Header=${this.voprosi_row_Header}
this.voprosi_row_BodyFirst=${this.voprosi_row_BodyFirst}

this.list1_col_CommentBlokStart=${this.list1_col_CommentBlokStart}
this.list1_col_CommentBlokEnd=${this.list1_col_CommentBlokEnd}
this.list1_col_MAX_COL=${this.list1_col_MAX_COL}
this.list1_col_Актуальный_Сбор_КП=${this.list1_col_Актуальный_Сбор_КП}
    `
  }

}


// const classColRow = new ClassColumnRow();
/** @type {ClassColumnRow} */
let classColRow = undefined;

/** @returns {ClassColumnRow} */
function getClassColRow() {
  if (!classColRow) {
    classColRow = new ClassColumnRow();
  }
  return classColRow;
}



