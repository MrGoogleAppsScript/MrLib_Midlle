

function onEditList_Вопросы(edit) {
  getContext().getSheetVoprosi().onEdit(edit);
}


class ClassSheet_Voprosi {
  constructor(url = getContext().getSpreadSheet().getUrl()) {
    Logger.log("ClassSheet_Voprosi constructor S ");
    // this.sheet = getContext().getSheetByName("Вопросы");
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Вопросы);

    this.url = url;
    this.col_N = nr("N");
    this.col_O = nr("O");
    this.row_Header = 1;
    this.row_BodyFirst = 2;
    this.rowConf = {
      head: { first: 1, last: 1, key: 1, type: 1 },
      body: { first: 2, last: 2, },
    };

    this.names = {
      ОтветРазослан: "Ответ разослан",
      ДобавитьВопросПроблему: "Добавить Вопрос Проблему",
      ПереченьТоваров: "Перечень товаров по которым требуется уточнение",
      КоммЗвонка: "Комментарий по итогу звонка",
      ДобавитьВопрос: "Добавить вопрос по проекту",
      ДобавитьПроблему: "Добавить проблему по проекту",

      ДатаОтвета: "Дата Ответа",
      ДатаОтправкиОтвета: "Дата отправки Ответа",

      ДатаДобавленияВопроса: "Дата добавления вопроса",
      ДатаДобавленияОтвета: "Дата добавления ответа",

      СКакойПочтыПисали: "С какой почты писали",
      НаКакуюПочтуПисали: "На какую почту писали",
      СутьВопроса: "Суть вопроса",
      КомментарийПоИтогуЗвонка: "Комментарий по итогу звонка",
      ПочтаОтвет: "Почта с которой получен ответ",

      Номенклатура: "Номенклатура",
      ЕмаилПоставщика: "Еmail поставщика",
      Вопрос: "Вопрос",
      Ответ: "Ответ",

      МенеджерСВопросомОзнакомлен: "С вопросом ознакомлен Менеджер",
      МенеджерДатаСВопросомОзнакомлен: "Дата ознакомленя Менеджера с вопросом",




      // 2-1 Вопросы от Поставщиков
      // 1. Рядом с полем "С вопросом ознакомлен Менеджер" добавить аналогичное поле "С вопросом ознакомлен Эксперт"
      // 2. Добавить автоматическое поле "Дата ознакомленя Эксперта с вопросом"
      // 3. После поля "Ответ" добавить столбцы "Комментарий Менеджера" и "Комментарий Эксперта"


      ЭкспертСВопросомОзнакомлен: "С вопросом ознакомлен Эксперт",
      ЭкспертДатаСВопросомОзнакомленя: "Дата ознакомленя Эксперта с вопросом",


      МенеджерКомментарий: "Комментарий Менеджера",
      ЭкспертКомментарий: "Комментарий Эксперта",


      // "с вопросом ознакомлен (менеджер)",  "с вопросом ознакомлен (менеджер)- дата" 

    }




    // this.requiredColNames();
    Logger.log("ClassSheet_Voprosi constructor F ");
  }
  requiredColNames() {
    try {
      let requiredColNamesArr = [
        // "Перечень товаров по которым требуется уточнение",
        // "Комментарий по итогу звонка",
        "key",
        this.names.ДатаДобавленияВопроса,
        this.names.ДатаДобавленияОтвета,
        this.names.МенеджерСВопросомОзнакомлен,
        this.names.МенеджерДатаСВопросомОзнакомлен,
        this.names.ЭкспертДатаСВопросомОзнакомленя,

      ];


      // 2-1 Вопросы от Поставщиков
      // 1. Рядом с полем "С вопросом ознакомлен Менеджер" добавить аналогичное поле "С вопросом ознакомлен Эксперт"
      // 2. Добавить автоматическое поле "Дата ознакомленя Эксперта с вопросом"
      // 3. После поля "Ответ" добавить столбцы "Комментарий Менеджера" и "Комментарий Эксперта"

      // if (isTest()) {
      if (true) {
        /** @type {SpreadsheetApp.Sheet} */
        let theSheet = this.getSheetModel_Вопросы().sheet;
        /** @type {[]} */
        let head_key = this.getSheetModel_Вопросы().head_key;

        if (!head_key.includes(this.names.ЭкспертСВопросомОзнакомлен)) {
          let col = head_key.indexOf(this.names.МенеджерСВопросомОзнакомлен);
          theSheet.insertColumnAfter(col);
          theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.ЭкспертСВопросомОзнакомлен);
          this.getSheetModel_Вопросы().init();
          this.getSheetModel_Вопросы().reset();
          head_key = this.getSheetModel_Вопросы().head_key;
        }


        if (!head_key.includes(this.names.МенеджерКомментарий)) {
          let col = head_key.indexOf(this.names.Ответ);
          theSheet.insertColumnAfter(col);
          theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.ЭкспертКомментарий);
          theSheet.insertColumnAfter(col);
          theSheet.getRange(this.rowConf.head.key, col + 1, 1, 1).setValue(this.names.МенеджерКомментарий);
          this.getSheetModel_Вопросы().init();
          this.getSheetModel_Вопросы().reset();
          head_key = this.getSheetModel_Вопросы().head_key;
        }

        // let col = head_key.indexOf(this.names.ДатаОтвета);
        // if (col > 0) {
        //   this.getSheetModel_Вопросы().sheet.getRange(this.rowConf.head.key, col, 1, 1).setValue(this.names.ДатаОтправкиОтвета);
        // }


      }





      this.getSheetModel_Вопросы().requiredColNames(requiredColNamesArr);

      let head_key = this.getSheetModel_Вопросы().head_key;



      let col = head_key.indexOf(this.names.МенеджерСВопросомОзнакомлен);
      if (col > 0) {
        let rule = SpreadsheetApp.newDataValidation().requireValueInList(["Ознакомлен",]).build();
        this.getSheetModel_Вопросы()
          .sheet
          .getRange(this.rowConf.body.first, col, this.getSheetModel_Вопросы().sheet.getMaxRows() - this.rowConf.body.first + 1, 1)
          .setDataValidation(rule);
      }



      this.sheetModel_Вопросы = undefined;
      this.makeKeys(this.sheet, getContext().getSpreadSheet().getId());
    } catch (err) { mrErrToString(err); }
  }

  onEdit(edit) {
    // Получаем диапазон ячеек, в которых произошли изменения
    let range = edit.range;

    // Лист, на котором производились изменения
    let sheet = range.getSheet();

    // Проверяем, нужный ли это нам лист
    Logger.log("onEditList_Вопросы");
    // if (sheet.getName() != 'Вопросы') {
    if (sheet.getName() != getSettings().sheetName_Вопросы) {
      return false;
    }


    // let head_key = this.getSheetModel_Вопросы().head_key;
    let head_key = ["row"].concat(this.sheet.getRange(this.rowConf.head.key, 1, 1, this.sheet.getLastColumn()).getValues()[0]);
    this.col_N = head_key.indexOf(this.names.ОтветРазослан);

    this.col_O = head_key.indexOf(this.names.ДатаОтправкиОтвета);
    // Logger.log([this.col_O, this.col_N, nc(this.col_O), nc(this.col_N), head_key]);
    if (this.col_N < 1) { return; }
    if (this.col_O < 1) { return; }
    let rowStart = range.rowStart;
    let rowEnd = range.rowEnd;
    // Logger.log([this.col_O, this.col_N, nc(this.col_O), nc(this.col_N), head_key]);
    if ((range.columnStart <= this.col_N) && (this.col_N <= range.columnEnd)) {
      // let d = new Date();
      let d = `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`;


      let vlsN = this.sheet.getRange(rowStart, this.col_N, rowEnd - rowStart + 1, 1).getValues();
      let vlsO = this.sheet.getRange(rowStart, this.col_O, rowEnd - rowStart + 1, 1).getValues();

      for (let i = 0, row = range.rowStart; i < vlsN.length; row++, i++) {
        if (row < this.row_BodyFirst) { continue; }
        if (vlsN[i][0] === true) { vlsO[i][0] = d; }
        if (fl_str(vlsN[i][0]) == "ДА") { vlsO[i][0] = d; }
      }

      this.sheet.getRange(rowStart, this.col_O, rowEnd - rowStart + 1, 1).setValues(vlsO);

    }

    return;

    // Дата добавления вопроса" и "Дата добавления ответа"
    this.col_ДатаДобавленияВопроса = head_key.indexOf(this.names.ДатаДобавленияВопроса);
    this.col_ДатаДобавленияОтвета = head_key.indexOf(this.names.ДатаДобавленияОтвета);

    // Logger.log([nc(this.col_ДатаДобавленияВопроса), nc(this.col_ДатаДобавленияОтвета)]);

    if ((range.columnStart <= this.col_ДатаДобавленияВопроса) && (this.col_ДатаДобавленияВопроса <= range.columnEnd)) {
      let range = this.sheet.getRange(rowStart, this.col_ДатаДобавленияВопроса, rowEnd - rowStart + 1, 1);
      let vls = this.sheet.getRange(rowStart, this.col_ДатаДобавленияВопроса, rowEnd - rowStart + 1, 1).getValues();
      range.setValues(vls);
    }

    if ((range.columnStart <= this.col_ДатаДобавленияОтвета) && (this.col_ДатаДобавленияОтвета <= range.columnEnd)) {
      let range = this.sheet.getRange(rowStart, this.col_ДатаДобавленияОтвета, rowEnd - rowStart + 1, 1);
      let vls = this.sheet.getRange(rowStart, this.col_ДатаДобавленияОтвета, rowEnd - rowStart + 1, 1).getValues();
      range.setValues(vls);
    }


    // this.makeKeys();
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
    let t = "Вопрос";
    vls.forEach(v => {
      let id = getContext().generateNextTimeId();
      let kk = [`pr:${нр}`, `id:${spreadSheetId}`, `${t}:${id}`].join(" ");
      sheet.getRange(v[0], col_key).setValue(kk);
    });

  }


  service() {
    this.requiredColNames();
    // this.makeKeys();
    // if (!testUrls.includes(this.url)) { return; }
    let sheetModelЗвонки = getContext().getSheetЗвонкиПроект().getSheetModel_Звонки();
    /** @type {Map} */
    let mapЗвонки = sheetModelЗвонки.getMap();
    // Logger.log(JSON.stringify([...mapЗвонки.values()], null, 2));
    let items = [...mapЗвонки.values()]
      .filter(item => {
        if (item[this.names.ДобавитьВопросПроблему] != this.names.ДобавитьВопрос) { return false; }
        if (this.getSheetModel_Вопросы().getMap().has(item["key"])) { return false; }
        return true;
      })
      .map(item => {
        let ret = {
          key: item["key"],
          "Добавлено Автоматический": "Да",
          "Перечень товаров по которым требуется уточнение": "Да",
          // "Комментарий по итогу звонка": item[this.names.КоммЗвонка],

          "Вопрос": `Вопрос поставлен Координатором\n${item[this.names.СутьВопроса]}\n${item[this.names.КоммЗвонка]}`,
          "Номенклатура": `${item[this.names.ПереченьТоваров]}`,
          "Еmail поставщика": [item[this.names.СКакойПочтыПисали], item[this.names.НаКакуюПочтуПисали], item[this.names.ПочтаОтвет]].filter(e => e).join("\n"),
        };
        ret[this.names.ДатаДобавленияВопроса] = `Дата: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`;
        return ret;
      });
    // Logger.log(JSON.stringify(items, null, 2));
    // Logger.log(JSON.stringify(this.getSheetModel_Вопросы().sheet.getName(), null, 2));

    this.getSheetModel_Вопросы().updateItems(items);
  }

  serviceСинзронизацияВопросов() {
    this.requiredColNames();

    // if (!this.ttst.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { return; }

    // this.requiredColNames();
    let url_buy = getContext().getUrls().buy;
    if (!url_buy) { return; }

    // MrLib_Midlle.getMrContextConstructor()
    let ssa_Внешнии = SpreadsheetApp.openByUrl(url_buy);

    let sheetNames = ssa_Внешнии.getSheets().map(sh => sh.getName());
    let sheetNameВопросы_Внешние = (() => {
      if (sheetNames.includes(getSettings().sheetName_Вопросы)) return getSettings().sheetName_Вопросы;
      if (sheetNames.includes(getSettings().sheetName_ВопросыО)) {
        ssa_Внешнии.getSheetByName(getSettings().sheetName_ВопросыО).setName(getSettings().sheetName_Вопросы);
        return getSettings().sheetName_Вопросы;
      }
    })();

    if (!sheetNameВопросы_Внешние) { return; }

    Logger.log("serviceСинзронизацияВопросов 2 " + `${sheetNameВопросы_Внешние} ${url_buy}`);

    /** @type {Map} */
    let mapВопросыП = this.getSheetModel_Вопросы().getMap();
    /** @type {Map} */
    let mapВопросыВ = this.getSheetModel_Вопросы_Внешние(url_buy, sheetNameВопросы_Внешние).getMap();

    this.makeKeys(this.getSheetModel_Вопросы_Внешние().sheet, this.getSheetModel_Вопросы_Внешние().context.getSpreadSheet().getId());


    Logger.log("serviceСинзронизацияВопросов 3 makeKeys");
    let добавитьП = [...mapВопросыВ.keys()]
      .filter(k => {
        if (mapВопросыП.has(k)) { return false; }
        return true;
      })
      .map(k => {
        Logger.log(JSON.stringify(k, null, 2));
        let item = mapВопросыВ.get(k);
        return item;

      });
    if (добавитьП.length > 0) {
      // Logger.log(JSON.stringify(добавитьП, null, 2));
      this.getSheetModel_Вопросы().updateItems(добавитьП);
    }

    //  return;
    Logger.log("serviceСинзронизацияВопросов 4 ");
    this.getSheetModel_Вопросы().reset();
    mapВопросыП = this.getSheetModel_Вопросы().getMap();

    let обновитьВ = [...mapВопросыВ.keys()]
      .filter(k => {
        if (!mapВопросыП.has(k)) { return false; }
        return true;
      })
      .map(k => {
        let item = mapВопросыП.get(k);
        return { "key": k, "Ответ": item["Ответ"] };
      });
    if (обновитьВ.length > 0) {
      Logger.log(JSON.stringify(обновитьВ, null, 2));
      this.getSheetModel_Вопросы_Внешние().updateItems(обновитьВ);
    }
    SpreadsheetApp.flush();


    this.getSheetModel_Вопросы_Внешние().reset();
    mapВопросыВ = this.getSheetModel_Вопросы_Внешние().getMap();
    this.getSheetModel_Вопросы().updateItems([...mapВопросыВ.values()].map(item => {
      if (`${item[this.names.ДатаОтправкиОтвета]}` != '') { item[this.names.ДатаОтправкиОтвета] = (() => { try { return "Дата: " + Utilities.formatDate(new Date(item[this.names.ДатаОтправкиОтвета]), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss") } catch { } })() }

      return item;
    }));
    Logger.log("serviceСинзронизацияВопросов 5 ");

  }

  getSheetModel_Вопросы() {
    if (!this.sheetModel_Вопросы) {
      this.sheetModel_Вопросы = MrLib_Midlle.makeSheetModelBy(this.url, this.sheet.getName(), this.rowConf);
      // this.sheetModel_Вопросы_Внешние = new MrLib_Midlle.getMrClassSheetModel()(sheetName, getContext(), this.rowConf)
    }
    return this.sheetModel_Вопросы;
  }


  getSheetModel_Вопросы_Внешние(url, sheetName) {
    if (!this.sheetModel_Вопросы_Внешние) {
      this.sheetModel_Вопросы_Внешние = MrLib_Midlle.makeSheetModelBy(url, sheetName, this.rowConf);
      // this.sheetModel_Вопросы_Внешние = new MrLib_Midlle.getMrClassSheetModel()(sheetName, getContext(), this.rowConf)
      // this.sheetModel_Вопросы_Внешние = MrLib_Midlle.makeSheetModelBy(url, sheetName, this.rowConf);

    }
    return this.sheetModel_Вопросы_Внешние;
  }



  serviceДатаДобавления() {
    /** @type {Map} */

    let arrH = ["key", "last", "row",];
    let arrC = [
      this.names.Вопрос,
      this.names.Ответ,
      this.names.ОтветРазослан,
      this.names.МенеджерСВопросомОзнакомлен,
      this.names.ЭкспертСВопросомОзнакомлен,
    ];

    let couples = new Object()
    couples[this.names.Вопрос] = this.names.ДатаДобавленияВопроса;
    couples[this.names.Ответ] = this.names.ДатаДобавленияОтвета;
    couples[this.names.ОтветРазослан] = this.names.ДатаОтправкиОтвета; // Ответ разослан	Дата отправки Ответа
    couples[this.names.МенеджерСВопросомОзнакомлен] = this.names.МенеджерДатаСВопросомОзнакомлен;
    couples[this.names.ЭкспертСВопросомОзнакомлен] = this.names.ЭкспертДатаСВопросомОзнакомленя;

    // this.head_key.length
    // changed_key
    let head_key = this.getSheetModel_Вопросы().head_key;
    this.getSheetModel_Вопросы().changed_key = head_key.map(h => {
      if (arrH.includes(h)) { return false; }
      return true;
    })
    let map = this.getSheetModel_Вопросы().getMap();

    let items = [...map.values()].map(item => {
      item.state = {
        current: (() => { try { return JSON.parse(JSON.stringify(this.getSheetModel_Вопросы().makeSnapshot(item))) } catch (err) { return new Object() } })(),
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
            Logger.log(["есть изменения", item.state.current[arrC[i]], item.state.past[arrC[i]]]);
            // Logger.log([`${item.state.current[arrC[i]]}`, `${item.state.past[arrC[i]]}`]);
            item.state.changed.push(arrC[i]);
          }
        }
        return item.state.changed.length > 0;

      })();
      return ret;
    }).map(item => {
      item.state.changed.forEach(ch => {
        // return;
        // let colLast = head_key.indexOf("last");
        // let colDate = head_key.indexOf(couples[ch]);
        /** @type {SpreadsheetApp.Sheet} */
        // let sheet = this.getSheetModel_Вопросы().sheet;

        // let formula = sheet.getRange(item.row, colFormula).getFormula();
        // if (!formula) { return; }

        // let date = sheet.getRange(item.row, colDate).getValue();
        let date = item[couples[ch]];
        if (!date) {
          date = "Дата: " + Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")
          // sheet.getRange(item.row, colDate).setValue(datet);
          item[couples[ch]] = date;
          // return;
        }

        // let value = sheet.getRange(item.row, colDate).getValue();
        // sheet.getRange(item.row, colDate).setValue(value);
        let snapshot = this.getSheetModel_Вопросы().makeSnapshot(item);
        item.last = JSON.stringify(snapshot);
        // sheet.getRange(item.row, colLast).setValue(JSON.stringify(snapshot));
        item["Ед"] = { " Изм": item["Ед. Изм"] };
        this.getSheetModel_Вопросы().updateItems([item]);

        // Logger.log(JSON.stringify({ serviceДатаДобавления: "serviceДатаДобавления", snapshot, row: item.row }, null, 2));
      });

      return item;
    });




  }


  exportToMidlle() {

    this.requiredColNames();
    let нр = (() => { try { return getContext().getНомерПроекта() } catch (err) { } })();
    if (!нр) { Logger.log("Нет Номера Проекта"); return; }
    нр = `${нр}`;

    let dataTaskSyncВопросы = MrLib_Midlle.getDataTaskSyncВопросы();
    dataTaskSyncВопросы.param.urls.proj = this.url;
    dataTaskSyncВопросы.param.projNumber = нр;

    // Logger.log(JSON.stringify(dataTaskSyncВопросы, undefined, 2));

    MrLib_Midlle.ВыполненитьМасивЗадач("ClassSheet_Voprosi update ВыполненитьМасивЗадач", [dataTaskSyncВопросы,], "ClassSheet_Voprosi");
  }


  aa() {
    let a = `=IF(ISBLANK(R[0]C[-1]);; FILTER({
'1-1 Сбор КП'!R2C[2]:C[2] \\  
BYROW('1-1 Сбор КП'!R2C[2]:C[2];LAMBDA(e;VLOOKUP(R[0]C[-1];{'1-2 Запросы по товарным группам'!R[-3]C1:C1 \\ '1-2 Запросы по товарным группам'!R[-3]C3:C3};2;false))) \\ 
'1-1 Сбор КП'!R2C[4]:C[9] \\ 
'1-1 Сбор КП'!R2C[12]:C[12]};{'1-1 Сбор КП'!R2C[1]:C[1]=R[0]C[-1]}))`


  }

}

// function testClassSheet_Voprosi() {
//   getContext().getSheetVoprosi().requiredColNames();
//   // let shВопросы = new ClassSheet_Voprosi(urlBuy);
//   // shВопросы.requiredColNames();


// }