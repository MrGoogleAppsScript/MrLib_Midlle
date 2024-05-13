

function onEditList_Проблемы(edit) {
  getContext().getSheetПроблемы().onEdit(edit);
}


class ClassSheet_Проблемы {
  constructor(url = getContext().getSpreadSheet().getUrl()) {
    // this.sheet = getContext().getSheetByName("Вопросы");
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Проблемы);

    this.col_G = nr("G");
    this.col_H = nr("H");
    this.row_Header = 1;
    this.row_BodyFirst = 2;


    this.url = url;
    this.rowConf = {
      head: { first: 1, last: 1, key: 1, },
      body: { first: 2, last: 2, },
    };

    this.names = {


      ДобавитьВопросПроблему: "Добавить Вопрос Проблему",
      ПереченьТоваров: "Перечень товаров по которым требуется уточнение",
      КоммЗвонка: "Комментарий по итогу звонка",
      ДобавитьВопрос: "Добавить вопрос по проекту",
      ДобавитьПроблему: "Добавить проблему по проекту",

      ДатаОзнакомления: "Дата ознакомления",
      ДатаОзнакомленияСОтветом: "Дата ознакомления с ответом",
      ДатаДобавленияПроблемы: "Дата добавления проблемы",
      ДатаДобавленияРешения: "Дата добавления решения",

      СКакойПочтыПисали: "С какой почты писали",
      НаКакуюПочтуПисали: "На какую почту писали",
      СутьВопроса: "Суть вопроса",
      КомментарийПоИтогуЗвонка: "Комментарий по итогу звонка",
      ПочтаОтвет: "Почта с которой получен ответ",

      ДатаПроблемы: `Дата проблемы`,
      ТоварПроблема: `Товар к которому относится проблема 
(не обязательно)`,
      ЕмаилПроблема: `Еmail к которомо относится проблема 
(не обязательно)`,
      СутьПроблемы: `Суть проблемы`,

      ПредлагаемоеРешение: "Предлагаемое решение",


      ДатаДобавленияПроблемы: "Дата добавления проблемы",
      ДатаДобавленияРешения: "Дата добавления решения",

      МенеджерСПроблемойОзнакомлен: "С проблемой ознакомлен Менеджер",
      МенеджерДатаСПроблемойОзнакомлен: "Дата ознакомленя Менеджера с проблемой",



      // 2-2 Проблемы по проекту
      // 4. Рядом с полем "С проблемой ознакомлен Менеджер" добавить аналогичное поле "С проблемой ознакомлен Эксперт"
      // 5. Добавить автоматическое поле "Дата ознакомленя Эксперта с проблемой"
      // 6. После поля "Предлагаемое решение" добавить столбцы "Комментарий Менеджера" и "Комментарий Эксперта" и "С решением ознакомлен Эксперт"
      // 7. Добавить автоматическое поле "Дата ознакомленя Эксперта с решением"
      // 8. Переименовать поле "С решением ознакомлен" в "С решением ознакомлен специалист по расчету"
      ЭкспертСПроблемойОзнакомлен: "С проблемой ознакомлен Эксперт",
      ЭкспертДатаОзнакомленяСПроблемой: "Дата ознакомленя Эксперта с проблемой",
      ЭкспертДатаОзнакомленяСРешением: "Дата ознакомленя Эксперта с решением",
      МенеджерКомментарий: "Комментарий Менеджера",
      ЭкспертКомментарий: "Комментарий Эксперта",
      ЭкспертСРешениемОзнакомлен: "С решением ознакомлен Эксперт",

      СпециалистПоРасчетуСРешениемОзнакомлен: "С решением ознакомлен специалист по расчету",
      СРешениемОзнакомленold: "С решением ознакомлен",
      СРешениемОзнакомлен: "С решением ознакомлен специалист по расчету",

    };


    this.ttst = [
     
    ];


    // this.requiredColNames();
  }


  requiredColNames() {
    try {
      let requiredColNamesArr = [
        // "Перечень товаров по которым требуется уточнение",
        // "Комментарий по итогу звонка",
        "key",
        this.names.ДатаДобавленияПроблемы,
        this.names.ДатаДобавленияРешения,
        this.names.МенеджерСПроблемойОзнакомлен,
        this.names.МенеджерДатаСПроблемойОзнакомлен,
        // 5. Добавить автоматическое поле "Дата ознакомленя Эксперта с проблемой"
        this.names.ЭкспертДатаОзнакомленяСПроблемой,
        // 7. Добавить автоматическое поле "Дата ознакомленя Эксперта с решением"
        this.names.ЭкспертДатаОзнакомленяСРешением,
      ];



      // 2-2 Проблемы по проекту

      // 8. Переименовать поле "С решением ознакомлен" в "С решением ознакомлен специалист по расчету"
      // if (isTest()) {
      if (true) {
        /** @type {SpreadsheetApp.Sheet} */
        let theSheet = this.getSheetModel_Проблемы().sheet;
        /** @type {[]} */
        let head_key = this.getSheetModel_Проблемы().head_key;

        // 4. Рядом с полем "С проблемой ознакомлен Менеджер" добавить аналогичное поле "С проблемой ознакомлен Эксперт"
        if (!head_key.includes(this.names.ЭкспертСПроблемойОзнакомлен)) {
          let col = head_key.indexOf(this.names.МенеджерСПроблемойОзнакомлен);
          theSheet.insertColumnAfter(col);
          theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.ЭкспертСПроблемойОзнакомлен);
          this.getSheetModel_Проблемы().init();
          this.getSheetModel_Проблемы().reset();
          head_key = this.getSheetModel_Проблемы().head_key;
        }


        // 6. После поля "Предлагаемое решение" добавить столбцы "Комментарий Менеджера" и "Комментарий Эксперта" и "С решением ознакомлен Эксперт"
        if (!head_key.includes(this.names.МенеджерКомментарий)) {
          let col = head_key.indexOf(this.names.ПредлагаемоеРешение);
          theSheet.insertColumnAfter(col);
          theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.ЭкспертКомментарий);
          theSheet.insertColumnAfter(col);
          theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.МенеджерКомментарий);
          this.getSheetModel_Проблемы().init();
          this.getSheetModel_Проблемы().reset();
          head_key = this.getSheetModel_Проблемы().head_key;
        }

        if (!head_key.includes(this.names.ЭкспертСРешениемОзнакомлен)) {
          let col = head_key.indexOf(this.names.СРешениемОзнакомленold);
          if (col > 0) {
            theSheet.insertColumnAfter(col);
            theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.ЭкспертСРешениемОзнакомлен);
            this.getSheetModel_Проблемы().init();
            this.getSheetModel_Проблемы().reset();
            head_key = this.getSheetModel_Проблемы().head_key;
          }
        }

        let col = head_key.indexOf(this.names.СРешениемОзнакомленold);
        if (col > 0) {
          this.getSheetModel_Проблемы().sheet.getRange(this.rowConf.head.key, col, 1, 1).setValue(this.names.СпециалистПоРасчетуСРешениемОзнакомлен);
        }


      }







      this.getSheetModel_Проблемы().requiredColNames(requiredColNamesArr);
      this.makeKeys(this.sheet, getContext().getSpreadSheet().getId());

      /** @type {[]} */
      let head_key = this.getSheetModel_Проблемы().head_key;
      // let col = head_key.indexOf(this.names.ДатаОзнакомления);
      // if (col > 0) {
      //   this.getSheetModel_Проблемы().sheet.getRange(this.rowConf.head.key, col, 1, 1).setValue(this.names.ДатаОзнакомленияСОтветом);
      // }



      let col = head_key.indexOf(this.names.МенеджерСПроблемойОзнакомлен);
      if (col > 0) {
        let rule = SpreadsheetApp.newDataValidation().requireValueInList(["Ознакомлен",]).build();
        this.getSheetModel_Проблемы()
          .sheet
          .getRange(this.rowConf.body.first, col, this.getSheetModel_Проблемы().sheet.getMaxRows() - this.rowConf.body.first + 1, 1)
          .setDataValidation(rule);
      }



      this.sheetModel_Проблемы = undefined;
    } catch (err) { mrErrToString(err); }
  }

  onEdit(edit) {
    // Получаем диапазон ячеек, в которых произошли изменения
    let range = edit.range;

    // Лист, на котором производились изменения
    let sheet = range.getSheet();

    // Проверяем, нужный ли это нам лист
    Logger.log("onEditList_Проблемы");
    if (sheet.getName() != getSettings().sheetName_Проблемы) {
      return false;
    }


    // let head_key = this.getSheetModel_Проблемы().head_key;
    let head_key = ["row"].concat(this.sheet.getRange(this.rowConf.head.key, 1, 1, this.sheet.getLastColumn()).getValues()[0]);
    this.col_G = head_key.indexOf(this.names.СРешениемОзнакомлен);
    if (this.col_G < 1) { return; }
    this.col_H = head_key.indexOf(this.names.ДатаОзнакомленияСОтветом);
    if (this.col_H < 1) { return; }

    let rowStart = range.rowStart;
    let rowEnd = range.rowEnd;

    if ((range.columnStart <= this.col_G) && (this.col_G <= range.columnEnd)) {
      // let d = new Date();
      let d = `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`;


      let vlsG = this.sheet.getRange(rowStart, this.col_G, rowEnd - rowStart + 1, 1).getValues();
      let vlsH = this.sheet.getRange(rowStart, this.col_H, rowEnd - rowStart + 1, 1).getValues();

      for (let i = 0, row = range.rowStart; i < vlsG.length; row++, i++) {
        if (row < this.row_BodyFirst) { continue; }
        if (vlsG[i][0] === true) { vlsH[i][0] = d; }
        if (fl_str(vlsG[i][0]) == "ДА") { vlsH[i][0] = d; }
      }

      this.sheet.getRange(rowStart, this.col_H, rowEnd - rowStart + 1, 1).setValues(vlsH);

    }


    return;
    // Дата добавления проблемы" и "Дата добавления решения"
    this.col_ДатаДобавленияПроблемы = head_key.indexOf(this.names.ДатаДобавленияПроблемы);
    this.col_ДатаДобавленияРешения = head_key.indexOf(this.names.ДатаДобавленияРешения);


    // Logger.log([
    //   this.col_G,
    //   this.col_H,
    //   nc(this.col_G),
    //   nc(this.col_H),
    //   this.col_ДатаДобавленияПроблемы,
    //   this.col_ДатаДобавленияРешения,
    //   nc(this.col_ДатаДобавленияПроблемы),
    //   nc(this.col_ДатаДобавленияРешения),
    //   head_key,
    // ]);

    if ((range.columnStart <= this.col_ДатаДобавленияПроблемы) && (this.col_ДатаДобавленияПроблемы <= range.columnEnd)) {
      let range = this.sheet.getRange(rowStart, this.col_ДатаДобавленияПроблемы, rowEnd - rowStart + 1, 1);
      let vls = this.sheet.getRange(rowStart, this.col_ДатаДобавленияПроблемы, rowEnd - rowStart + 1, 1).getValues();
      range.setValues(vls);
    }

    if ((range.columnStart <= this.col_ДатаДобавленияРешения) && (this.col_ДатаДобавленияРешения <= range.columnEnd)) {
      let range = this.sheet.getRange(rowStart, this.col_ДатаДобавленияРешения, rowEnd - rowStart + 1, 1);
      let vls = this.sheet.getRange(rowStart, this.col_ДатаДобавленияРешения, rowEnd - rowStart + 1, 1).getValues();
      range.setValues(vls);
    }

  }


  makeKeys(sheet, spreadSheetId) {
    Logger.log("makeKeys" + `${sheet.getName()} ${spreadSheetId}`);
    // if (!this.ttst.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { return; }

    // let нр = (() => { try { return `0000000${getContext().getНомерПроекта()}`.slice(-6) } catch (err) { } })();
    let нр = (() => { try { return `${getContext().getНомерПроекта()}` } catch (err) { } })();

    if (!нр) { Logger.log("Нет Номера Проекта"); return; }


    let headRow = sheet.getRange("1:1").getValues()[0]
    let col_key = headRow.indexOf("key") + 1;
    let col_last = headRow.indexOf("last") + 1;
    if (col_key == 0) return;
    /** @type {[[]]} */
    let vls = (() => { try { return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues(); } catch (err) { return [] } })();
    vls = vls.map((v, i) => [i + 2].concat(v));
    vls = vls.filter(v => {
      if (v[col_key] != '') {
        return false;
      }

      if (v.slice(1, col_key).join("") == '') {
        return false;
      }
      return true;
    });
    // Logger.log(vls.length);
    if (vls.length == 0) { return; }
    let t = "Проблема";
    vls.forEach(v => {
      let id = getContext().generateNextTimeId();
      let kk = [`pr:${нр}`, `id:${spreadSheetId}`, `${t}:${id}`].join(" ");
      sheet.getRange(v[0], col_key).setValue(kk);
    });

  }

  /** @param {Map} mapMalo */
  setMaloOtvetov(mapMalo) {
    Logger.log(" ClassSheet_Проблемы setMaloOtvetov S");


    let statisticsЗвонки = getContext().getSheetЗвонкиПроект().getStatistics();
    if (statisticsЗвонки.total == 0) { return; }
    if ((statisticsЗвонки.statusOn + statisticsЗвонки.statusNone) > 0) { return; }




    let m = undefined
    try {
      m = [...mapMalo.values()].filter(v => {
        if (!`${fl_str(v.malo)}`.includes(fl_str("Мало,"))) { return false; }
        return true;
      });
    } catch (err) {
      mrErrToString(err);
    }

    if (!Array.isArray(m)) { return; }

    let нр = (() => { try { return getContext().getНомерПроекта() } catch (err) { } })();

    if (!нр) { Logger.log("Нет Номера Проекта"); return; }
    нр = `${нр}`;
    let spreadSheetId = getContext().getSpreadSheet().getId();
    let t = "Проблема";
    let kk = [`pr:${нр}`, `id:${spreadSheetId}`, `${t}:Мало`].join(" ");
    let problem = {
      //№
      "№": "0",
      "Предлагаемое решение": "",
      "С решением ознакомлен": `${m.length == 0 ? "Да" : ""}`,

      "Предлагаемое решение": `${m.length == 0 ? "Проблема закрыта автоматически.\nПредложений достаточно у всех номенклатур" : ""}`,

      "key": kk,
      "Добавлено Автоматический": "Да",
      "Перечень товаров по которым требуется уточнение": m.length == 0 ? `Предложений достаточно у всех номенклатур` : `${m.map(v => v.name).join("\n")}`,
      "Суть проблемы": m.length == 0 ? `Проблема закрыта автоматически.\nПредложений достаточно у всех номенклатур` : `Проблема поставлена автоматически.\nМало предложений у ${m.length} ${m.length == 1 ? "номенклатуры" : "номенклатур"}`,
      "Товар к которому относится проблема \n(не обязательно)": m.length == 0 ? `Предложений достаточно у всех номенклатур` : `${m.map(v => v.name).join("\n")}`,
      "Дата проблемы": `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`,
    } // 
    problem[this.names.ДатаДобавленияПроблемы] = `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`;
    // Logger.log(JSON.stringify(problem, null, 2));
    this.requiredColNames();
    // if (!this.ttst.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { return; }

    let problem_has = (() => {
      try {
        // problem.err = "-";
        let ret = this.getSheetModel_Проблемы().getMap().get(problem.key);
        // Logger.log("problem_has ClassSheet_Проблемы setMaloOtvetov S")
        // Logger.log(JSON.stringify(ret, null, 2));
        return ret;
      }
      catch (err) { mrErrToString(err); return undefined; }
    })();

    if (problem_has) {
      problem[this.names.ДатаПроблемы] = problem_has[this.names.ДатаПроблемы];
      problem[this.names.ДатаДобавленияПроблемы] = problem_has[this.names.ДатаДобавленияПроблемы];
      problem[this.names.ДатаДобавленияРешения] = problem_has[this.names.ДатаДобавленияРешения];

      // problem.log = "Даты";
    }

    this.getSheetModel_Проблемы().updateItems([problem]);
    Logger.log(" ClassSheet_Проблемы setMaloOtvetov F");
  }

  service() {
    this.requiredColNames();
    // if (!testUrls.includes(this.url)) { return; }
    let sheetModelЗвонки = getContext().getSheetЗвонкиПроект().getSheetModel_Звонки();
    /** @type {Map} */
    let mapЗвонки = sheetModelЗвонки.getMap();
    // Logger.log(JSON.stringify([...mapЗвонки.values()], null, 2));
    let items = [...mapЗвонки.values()]
      .filter(item => {
        if (item[this.names.ДобавитьВопросПроблему] != this.names.ДобавитьПроблему) { return false; }
        if (this.getSheetModel_Проблемы().getMap().has(item["key"])) { return false; }
        return true;
      })
      .map(item => {
        let ret = {
          key: item["key"],
          "Добавлено Автоматический": "Да",
          "Перечень товаров по которым требуется уточнение": "Да",
          // "Комментарий по итогу звонка": item[this.names.КоммЗвонка],

          "Суть проблемы": `Проблема поставлена Координатором\n${item[this.names.СутьВопроса]}\n${item[this.names.КоммЗвонка]}`,
          "Товар к которому относится проблема \n(не обязательно)": `${item[this.names.ПереченьТоваров]}`,
          "Еmail к которомо относится проблема \n(не обязательно)": [item[this.names.СКакойПочтыПисали], item[this.names.НаКакуюПочтуПисали], item[this.names.ПочтаОтвет]].filter(e => e).join("\n"),
          "Дата проблемы": `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`,
        };
        ret[this.names.ДатаДобавленияПроблемы] = `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`;
        return ret;
      });

    // this.requiredColNames();
    this.getSheetModel_Проблемы().updateItems(items);
  }



  serviceДатаДобавления() {
    /** @type {Map} */

    let arrH = ["key", "last", "row",];
    let arrC = [
      this.names.СутьПроблемы,
      this.names.ПредлагаемоеРешение,
      this.names.СРешениемОзнакомлен,
      this.names.МенеджерСПроблемойОзнакомлен,
      this.names.ЭкспертСПроблемойОзнакомлен,
      this.names.ЭкспертСРешениемОзнакомлен,
    ];
    let couples = new Object();
    couples[this.names.СутьПроблемы] = this.names.ДатаДобавленияПроблемы;
    couples[this.names.ПредлагаемоеРешение] = this.names.ДатаДобавленияРешения;
    couples[this.names.СРешениемОзнакомлен] = this.names.ДатаОзнакомленияСОтветом;  // С решением ознакомлен	Дата ознакомления с ответом
    couples[this.names.МенеджерСПроблемойОзнакомлен] = this.names.МенеджерДатаСПроблемойОзнакомлен;
    couples[this.names.ЭкспертСПроблемойОзнакомлен] = this.names.ЭкспертДатаОзнакомленяСПроблемой;
    couples[this.names.ЭкспертСРешениемОзнакомлен] = this.names.ЭкспертДатаОзнакомленяСРешением;

    // this.head_key.length
    // changed_key
    let head_key = this.getSheetModel_Проблемы().head_key;
    this.getSheetModel_Проблемы().changed_key = head_key.map(h => {
      if (arrH.includes(h)) { return false; }
      return true;
    })
    let map = this.getSheetModel_Проблемы().getMap();


    let items = [...map.values()].map(item => {
      item.state = {
        current: (() => { try { return JSON.parse(JSON.stringify(this.getSheetModel_Проблемы().makeSnapshot(item))) } catch (err) { return new Object() } })(),
        past: (() => { try { return JSON.parse(item.last) } catch (err) { return new Object() } })(),
        changed: new Array(),
      }
      return item;
    }).filter(item => {
      let ret = (() => {
        for (let i = 0; i < arrC.length; i++) {
          if (!item.state.current[arrC[i]]) { item.state.current[arrC[i]] = "" }
          if (!item.state.past[arrC[i]]) { item.state.past[arrC[i]] = "" }

          if (`${item.state.current[arrC[i]]}` != `${item.state.past[arrC[i]]}`) {
            // Logger.log([typeof item.state.current[arrC[i]], typeof item.state.past[arrC[i]]]);
            // Logger.log([item.state.current[arrC[i]], item.state.past[arrC[i]]]);
            // Logger.log([`${item.state.current[arrC[i]]}`, `${item.state.past[arrC[i]]}`]);
            item.state.changed.push(arrC[i]);
          }
        }
        return item.state.changed.length > 0;

      })();
      return ret;
    }).map(item => {
      item.state.changed.forEach(ch => {
        // let colLast = head_key.indexOf("last");
        // let colFormula = head_key.indexOf(couples[ch]);
        // let sheet = this.getSheetModel_Проблемы().sheet;
        // let formula = sheet.getRange(item.row, colFormula).getFormula();
        // if (!formula) { return; }
        // let value = sheet.getRange(item.row, colFormula).getValue();
        // sheet.getRange(item.row, colFormula).setValue(value);
        // let snapshot = this.getSheetModel_Проблемы().makeSnapshot(item);
        // sheet.getRange(item.row, colLast).setValue(JSON.stringify(snapshot));
        // // Logger.log(JSON.stringify({ formula, value, snapshot, colLast, colFormula, }, null, 2));

        let date = item[couples[ch]];
        if (!date) {
          date = "Дата: " + Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")
          // sheet.getRange(item.row, colDate).setValue(datet);
          item[couples[ch]] = date;
          // return;
        }

        let snapshot = this.getSheetModel_Проблемы().makeSnapshot(item);
        item.last = JSON.stringify(snapshot);
        this.getSheetModel_Проблемы().updateItems([item]);

        Logger.log(JSON.stringify({ serviceДатаДобавления: "serviceДатаДобавления", snapshot, row: item.row }, null, 2));


      });

      return item;
    });


    // Logger.log(items.map(item => { return { row: item.row, key: item.key, changed: item.state.changed } }));

    // /** @param {Obj} obj */
    // isChanged(obj) {
    //   let snapshotObj = this.makeSnapshot(obj);
    //   if (JSON.stringify(snapshotObj, null, 2) === obj.last) { return false; }
    //   return true;
    // }

    // /** @param {Obj} obj */
    // makeSnapshot(obj) {
    //   let retObj = new Object();
    //   for (let i = this.head_key.length - 1; i >= 0; i--) {
    //     if (this.changed_key[i] !== true) { continue; }
    //     let key = this.head_key[i];
    //     if (!key) { continue; }
    //     retObj[key] = obj[key];
    //   }
    //   return retObj;
    // }


  }


  serviceСинзронизацияПроблем() {

    this.requiredColNames();

    // if (!this.ttst.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { return; }

    let url_buy = getContext().getUrls().buy;
    if (!url_buy) { return; }

    // MrLib_Midlle.getMrContextConstructor()
    let ssa_Внешнии = SpreadsheetApp.openByUrl(url_buy);

    let sheetNames = ssa_Внешнии.getSheets().map(sh => sh.getName());
    let sheetNameПроблемы_Внешние = (() => {
      if (sheetNames.includes(getSettings().sheetName_Проблемы)) return getSettings().sheetName_Проблемы;
      // if (sheetNames.includes(getSettings().sheetName_ВопросыО)) {
      //   ssa_Внешнии.getSheetByName(getSettings().sheetName_ВопросыО).setName(getSettings().sheetName_Вопросы);
      //   return getSettings().sheetName_Вопросы;
      // }
    })();

    if (!sheetNameПроблемы_Внешние) { return; }

    Logger.log("serviceСинзронизацияПроблем 2 " + `${sheetNameПроблемы_Внешние} ${url_buy}`);

    /** @type {Map} */
    let mapПроблемыП = this.getSheetModel_Проблемы().getMap();
    /** @type {Map} */
    let mapПроблемыВ = this.getSheetModel_Проблемы_Внешние(url_buy, sheetNameПроблемы_Внешние).getMap();

    this.makeKeys(this.getSheetModel_Проблемы_Внешние().sheet, this.getSheetModel_Проблемы_Внешние().context.getSpreadSheet().getId());


    Logger.log("serviceСинзронизацияПроблем 3 makeKeys");
    let добавитьП = [...mapПроблемыВ.keys()]
      .filter(k => {
        if (mapПроблемыП.has(k)) { return false; }
        return true;
      })
      .map(k => {
        // Logger.log(JSON.stringify(k, null, 2));
        let item = mapПроблемыВ.get(k);
        if (item["Дата проблемы"]) { item["Дата проблемы"] = (() => { try { return new Date(item["Дата проблемы"]) } catch { } })(); }
        return item;

      });
    if (добавитьП.length > 0) {
      // Logger.log(JSON.stringify(добавитьП, null, 2));
      this.getSheetModel_Проблемы().updateItemsObj(добавитьП);
    }

    //  return;
    Logger.log("serviceСинзронизацияПроблем 4 ");
    this.getSheetModel_Проблемы().reset();
    mapПроблемыП = this.getSheetModel_Проблемы().getMap();

    let обновитьВ = [...mapПроблемыВ.keys()]
      .filter(k => {
        if (!mapПроблемыП.has(k)) { return false; }
        return true;
      })
      .map(k => {
        let item = mapПроблемыП.get(k);
        // Суть проблемы	Предлагаемое решение
        // return { "key": k, "Ответ": item["Ответ"] };
        return { "key": k, "Предлагаемое решение": item["Предлагаемое решение"] };

      });
    if (обновитьВ.length > 0) {
      // Logger.log(JSON.stringify(обновитьВ, null, 2));
      this.getSheetModel_Проблемы_Внешние().updateItems(обновитьВ);
    }
    SpreadsheetApp.flush();


    this.getSheetModel_Проблемы_Внешние().reset();
    mapПроблемыВ = this.getSheetModel_Проблемы_Внешние().getMap();
    this.getSheetModel_Проблемы().updateItemsObj([...mapПроблемыВ.values()].map(item => {
      // Logger.log(item.key);
      // Logger.log(Number.isNaN (item["Дата проблемы"]));
      // Logger.log(new Date(item["Дата проблемы"]));
      // Logger.log(item["Дата проблемы"]);

      if (`${item["Дата проблемы"]}` != '') { item["Дата проблемы"] = (() => { try { return "Дата: " + Utilities.formatDate(new Date(item["Дата проблемы"]), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss") } catch { } })() }
      if (`${item[this.names.ДатаОзнакомленияСОтветом]}` != '') { item[this.names.ДатаОзнакомленияСОтветом] = (() => { try { return "Дата: " + Utilities.formatDate(new Date(item[this.names.ДатаОзнакомленияСОтветом]), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss") } catch { } })() }

      return item;
    }));
    Logger.log("serviceСинзронизацияПроблем 5 ");

  }


  getSheetModel_Проблемы_Внешние(url, sheetName) {
    if (!this.sheetModel_Проблемы_Внешние) {
      this.sheetModel_Проблемы_Внешние = MrLib_Midlle.makeSheetModelBy(url, sheetName, this.rowConf);
    }
    // Logger.log(this.sheetModel_Проблемы_Внешние.context)
    return this.sheetModel_Проблемы_Внешние;
  }



  getSheetModel_Проблемы() {
    if (!this.sheetModel_Проблемы) {
      this.sheetModel_Проблемы = MrLib_Midlle.makeSheetModelBy(this.url, this.sheet.getName(), this.rowConf);
    }
    return this.sheetModel_Проблемы;
  }














  exportToMidlle() {


    this.requiredColNames();

    let нр = (() => { try { return getContext().getНомерПроекта() } catch (err) { } })();
    if (!нр) { Logger.log("Нет Номера Проекта"); return; }
    нр = `${нр}`;

    let dataTaskSyncПроблемы = MrLib_Midlle.getDataTaskSyncПроблемы();
    dataTaskSyncПроблемы.param.urls.proj = this.url;
    dataTaskSyncПроблемы.param.projNumber = нр;

    // Logger.log(JSON.stringify(dataTaskSyncПроблемы, undefined, 2));

    MrLib_Midlle.ВыполненитьМасивЗадач("ClassSheet_Проблемы update ВыполненитьМасивЗадач", [dataTaskSyncПроблемы,], "ClassSheet_Проблемы");

  }



}


function testClassSheet_Проблемы() {
  getContext().getSheetПроблемы().requiredColNames();

}