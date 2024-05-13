class ClassSheet_НастройкиПисем {
  constructor() { // class constructor
    this.sheet = getContext().getSheetByName(getSettings().sheetName_Настройки_писем);
    this.da = "Да";
    // this.strRangeFinishFlag = "K13";
    this.strRangeDaysBetweenLetterAndCall = "K12";

    // this.folderId = getSettings().folderId;
    // this.url_patternSpreadSheet = getSettings().url_таблица_закупки_шаблон;

    // this.rowHadsFirst = 1;
    // this.rowHadsLast = 2;


    // this.rowBodyFirst = 3;

    // this.rowBodyLast = this.findRowBodyLast();

    // this.makeCol();

    this.strRanges = {
      СтатусыЗвонков: "BB2:BB",

      auto_Finish: "C18",
      mailSent_Finish: "E18",

      auto_Continue: "C26",
      mailSent_Continue: "E26",

      daysBetweenLetterAndCall: "K12",
      summary: { row: 2, col: nr("AE") },
      triggersMem: "K11",
    }
  }

  serviceSummary() {
    let maloOtvetov = this.getMaloB8();
    let vls = getSummary(maloOtvetov);
    this.sheet.getRange(this.strRanges.summary.row, this.strRanges.summary.col, vls.length, vls[0].length).setValues(vls);
    let d = `Обновленио: ${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`;
    this.sheet.getRange(this.strRanges.summary.row, this.strRanges.summary.col).setNote(d);
  }


  getMaloB8() {

    let ret = getContext().getSheetList1().sheet.getRange("B8").getValue();
    if (!ret) { ret = 3; }
    if (isNaN(ret)) { ret = 3; }
    if (ret < 0) { ret = 3; }
    return ret;

  }


  getСтатусыЗвонков() {
    if (!this.СтатусыЗвонков) {
      let vls = this.sheet.getRange(this.strRanges.СтатусыЗвонков).getValues();
      let notes = this.sheet.getRange(this.strRanges.СтатусыЗвонков).getNotes();
      this.СтатусыЗвонков = new Map();
      vls.forEach((v, i) => {
        if (!v[0]) { return; }
        this.СтатусыЗвонков.set(v[0], notes[i][0] === 'true');
      });
    }
    return this.СтатусыЗвонков;
  }

  updateСтатусыЗвонков() {
    let dateNow = getContext().timeConstruct;
    if (dateNow.getHours() != 23) {
      Logger.log(`Час ${dateNow.getHours()} `
        + "Не подходящее Время: "
        + `${Utilities.formatDate(dateNow, "Europe/Moscow", "dd.MM.yyyy HH:mm:ss")}`);
      return;
    }

    let urlЗвонкиКоординатора = MrLib_Midlle.getSettings().urls.ЗвонкиКоординатора;
    Logger.log(urlЗвонкиКоординатора);
    let shModelTh = MrLib_Midlle.makeSheetModelBy(urlЗвонкиКоординатора, "Технический лист", {
      head: { first: 1, last: 1, key: 1, },
      body: { first: 2, last: 2, },
    });
    /** @type {Map} */
    let mapСтатусыЗвонков = shModelTh.getMap();

    // Logger.log(JSON.stringify([...mapСтатусыЗвонков.values()], null, 2));

    let rangeСтатусыЗвонков = this.sheet.getRange(this.strRanges.СтатусыЗвонков);
    rangeСтатусыЗвонков.clearContent();

    rangeСтатусыЗвонков.clearNote();

    let nts = rangeСтатусыЗвонков.getNotes();
    let vls = rangeСтатусыЗвонков.getValues();
    [...mapСтатусыЗвонков.values()].forEach((item, i) => {
      vls[i][0] = item["Статусы звонков"];
      nts[i][0] = item["Скрывать строку"];
    });
    // Logger.log(nts);
    // Logger.log(vls);
    rangeСтатусыЗвонков.setValues(vls);
    rangeСтатусыЗвонков.setNotes(nts);
    SpreadsheetApp.flush();

    // Logger.log(JSON.stringify([...this.getСтатусыЗвонков().values()], null, 2));


  }

  getDaysBetweenLetterAndCall() {
    let def = 1;
    let deys = this.sheet.getRange(this.strRanges.daysBetweenLetterAndCall).getValue();
    deys = Number.parseInt(deys);
    if (Number.isNaN(deys)) {
      deys = def;
    }
    if (!Number.isFinite(deys)) {
      deys = def;
    }
    if (deys < 0) { deys = 1 }
    return deys;
  }

  getMailDataFinish() {
    if (!this.mailDataFinish) {

      // this.mailDataFinish = {
      //   subject: this.sheet.getRange("C13").getValue(),//Тема письма
      //   htmlBody: this.sheet.getRange("C14").getValue(),//Текст Содержание
      //   attachmentsId: this.sheet.getRange("C16").getValue(),//Прикрепленный файл Вложение (id файла)
      //   attachments: undefined,
      //   replyTo: this.sheet.getRange("C17").getValue(),//Куда отвечать Адрес ответа
      //   project: (() => { try { return `<br><br>Проект № ${SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]}` } catch (err) { return "" } })()
      // }
      let tm = this.getTemplateForSendMails();
      if (!tm) {
        throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю. или нет доступа к файлу`;
        throw `Для ${getContext().getUserEmail()} нет шаблона`;
      }
      if (tm["Закрытие"] !== true) {
        throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`;
      }

      this.mailDataFinish = {
        // Закрытие Тема письма	
        // Закрытие Текст Содержание	
        // Закрытие Вложение	
        // Закрытие Адрес ответа
        subject: tm["Закрытие Тема письма"],//Тема письма
        htmlBody: tm["Закрытие Текст Содержание"],//Текст Содержание
        attachmentsId: tm["Закрытие Вложение"],//Прикрепленный файл Вложение (id файла)
        attachments: undefined,
        replyTo: tm["Закрытие Адрес ответа"],//Куда отвечать Адрес ответа
        project: (() => { try { return `<br><br>Проект № ${SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]}` } catch (err) { return "" } })()
      }



      this.mailDataFinish.attachments = (() => {
        if (this.mailDataFinish.attachmentsId.toString().length > 0) {
          var pdfFileBlob = DriveApp.getFileById(this.mailDataFinish.attachmentsId).getBlob();
          return [pdfFileBlob];
        } else {
          return undefined;
        }
      })();
      this.mailDataFinish.htmlBody = `${this.mailDataFinish.htmlBody} ${this.mailDataFinish.project}`;
    }
    return this.mailDataFinish;
  }

  getMailDataContinue() {
    if (!this.mailDataFinish) {
      // this.mailDataFinish = {
      //   subject: this.sheet.getRange("C21").getValue(),//Тема письма
      //   htmlBody: this.sheet.getRange("C22").getValue(),//Текст Содержание
      //   attachmentsId: this.sheet.getRange("C24").getValue(),//Прикрепленный файл Вложение (id файла)
      //   attachments: undefined,
      //   replyTo: this.sheet.getRange("C25").getValue(),//Куда отвечать Адрес ответа
      //   project: (() => { try { return `<br><br>Проект № ${SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]}` } catch (err) { return "" } })()
      // }

      let tm = this.getTemplateForSendMails();
      if (!tm) {
        throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю. или нет доступа к файлу`;
        throw `Для ${getContext().getUserEmail()} нет шаблона`;
      }
      if (tm["Положительный"] !== true) {
        throw `В настоящее время в базе отсутствуют реквизиты для отправки сообщений с этого емейла (${getContext().getUserEmail()}), обратитесь к руководителю.`;
      }
      // Logger.log(tm["Положительный"]);
      this.mailDataFinish = {
        // Положительный	
        // Положительный Тема письма	
        // Положительный Текст Содержание	
        // Положительный Вложение	
        // Положительный Адрес ответа

        subject: tm["Положительный Тема письма"],//Тема письма
        htmlBody: tm["Положительный Текст Содержание"],//Текст Содержание
        attachmentsId: tm["Положительный Вложение"],//Прикрепленный файл Вложение (id файла)
        attachments: undefined,
        replyTo: tm["Положительный Адрес ответа"],//Куда отвечать Адрес ответа
        project: (() => { try { return `<br><br>Проект № ${SpreadsheetApp.getActiveSpreadsheet().getName().trim().split(" ")[0]}` } catch (err) { return "" } })()
      }


      this.mailDataFinish.attachments = (() => {
        if (this.mailDataFinish.attachmentsId.toString().length > 0) {
          var pdfFileBlob = DriveApp.getFileById(this.mailDataFinish.attachmentsId).getBlob();
          return [pdfFileBlob];
        } else {
          return undefined;
        }
      })();
      this.mailDataFinish.htmlBody = `${this.mailDataFinish.htmlBody} ${this.mailDataFinish.project}`;

    }
    return this.mailDataFinish;
  }

  sendMail(mail) {
    Logger.log(JSON.stringify({ from_: getContext().getUserEmail() }, null, 2));
    Logger.log(JSON.stringify(mail, null, 2));
    if (testUrls.includes(getContext().getSpreadSheet().getUrl())) { // тест
      Logger.log("Тестовая таблица");
      return;
    }
    MailApp.sendEmail(mail);
    // Logger.log(JSON.stringify(mail, null, 2));
  }


  setFlagByRange(rangeMem, fl = this.da) {
    // let rangeMem = this.strRanges.finishFlag;
    this.sheet.getRange(rangeMem).setValue(fl);

    // Logger.log(`setFinishFlag Start `);

    let memNote = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).getNote();
    memNote = (() => { try { return JSON.parse(memNote); } catch { return new Object(); } })();
    memNote.mailMem = new DataClassFinishMem();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).setNote(JSON.stringify(memNote, undefined, 2));
    // Logger.log(`setFinishFlag Finish = ${JSON.stringify(memNote)}`);

  }



  getFlagMemNoteByRange(rangeMem, fl = this.da) {

    if (fl_str(this.sheet.getRange(rangeMem).getValue()) !== fl_str(fl)) { return false; }
    let memNote = this.sheet.getRange(rangeMem).getNote();
    memNote = (() => { try { return JSON.parse(memNote) } catch (err) { mrErrToString(err); return undefined } })();
    if (!memNote) { return false; }
    if (!memNote.finishMem) { return false; }
    let d = SpreadsheetApp.getActiveSpreadsheet().getId().slice(0, 10);
    if (memNote.finishMem.d != d) { return false; }
    return true;
  }

  getFlagByRange(rangeMem, fl = this.da) {
    if (fl_str(this.sheet.getRange(rangeMem).getValue()) !== fl_str(fl)) { return false; }
    return true;
  }





  onEditHelper(duration = 1 / 24 / 60 * 4.9) {
    return;
    if (!getContext().hasTime(duration)) { return; }



    // let vls_AJ1_AK1 = this.sheet.getRange("AJ1:AK1").getValues();  

    let vlsAK = this.sheet.getRange("AK1:AK").getValues();
    vlsAK = vlsAK.map(v => v[0]).slice(0, 52);

    let vlsAJ = this.sheet.getRange("AJ1:AJ").getValues();
    vlsAJ = vlsAJ.map(v => v[0]).slice(0, 52);

    // Logger.log("vls_AJ1_AK1=" + `${vlsAK}`);

    // if (vls_AJ1_AK1[0][0]!="JSON Заголовки"){ return;}
    // if (vls_AJ1_AK1[0][1]!="JSON Значения"){ return;}


    let url = "https://docs.google.com/spreadsheets/d/12KYFCKo/edit#gid=0";

    let ssa = SpreadsheetApp.openByUrl(url);
    // ssa.getSheetByName("Лист4").appendRow([SpreadsheetApp.getActive().getUrl(), SpreadsheetApp.getActive().getName(), ...vlsAJ]);
    // ssa.getSheetByName("Лист3").appendRow([SpreadsheetApp.getActive().getUrl(), SpreadsheetApp.getActive().getName(), ...vlsAK]);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ss.getOwner();
    let em = Session.getEffectiveUser().getEmail();

    // Logger.log(owner.getEmail());
    // Logger.log(Session.getEffectiveUser().getEmail());

    // if (SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() !== Session.getEffectiveUser().getEmail()) {
    //   Logger.log("Триггер может добавить только владелец");
    //   return;
    // }

    // ssa.getSheetByName("Лист7").appendRow([SpreadsheetApp.getActive().getUrl(), SpreadsheetApp.getActive().getName(), owner , em ]);


  }


  /** @param {Array} memArr */
  addToMemSentProductRequests(memArr, url) {

    // if (getContext().getSpreadSheet().getUrl() !== "https://docs.google.com/spreadsheets/d/1_k3PdkCo/edit") { // тест
    //   return;
    // }


    // if (!testUrls.includes(getContext().getSpreadSheet().getUrl())) { // тест
    //   return
    // }



    // memArr.forEach(item => { Logger.log(item) });
    // let url = getContext().getUrls().product;
    let shl = getContext().getSheetList1();
    let vlsPr = shl.sheet.getRange(shl.rowBodyFirst, shl.columnProduct, shl.rowBodyLast - shl.rowBodyFirst + 1, 1).getValues();
    vlsPr = vlsPr.flat(1);

    memArr = memArr.map(item => {
      item.НоменклатураТекст = item.Номенклатура;
      let row = (() => { return vlsPr.indexOf(item.НоменклатураТекст) + shl.rowBodyFirst })();
      item.Номенклатура = `'${getSettings().sheetName_Лист_1}'!R${row}C${shl.columnProduct}`;
      return item;
    });
    if (!this.sheetModel_Отправленные_Запросы) {
      let sheetName = getSettings().sheetName_Отпр_запросы; //"Отправленные запросы";
      this.sheetModel_Отправленные_Запросы = MrLib_Midlle.makeSheetModelBy(url, sheetName);
    }
    this.sheetModel_Отправленные_Запросы.updateItems(memArr);
  }

  getTemplateForSendMails() {
    try {
      // let email = getContext().getUserEmail()
      let url = "https://docs.google.com/spreadsheets/d/1U0C8/edit";
      let sheetName = "Шаблоны Писем";
      let sheetModel = MrLib_Midlle.makeSheetModelBy(url, sheetName);
      return sheetModel.getMap().get(getContext().getUserEmail());
    } catch (err) {
      mrErrToString(err);
      // throw mrErrToString(err);
    }
    return undefined;
  }





  /** @param {Map} memMeilsMap */
  addToMemSentMails(memMeilsMap, url) {
    try {
      // if (!testUrls.includes(getContext().getSpreadSheet().getUrl())) { // тест
      //   return
      // }

      if (memMeilsMap.size == 0) { return; }

      // Отправленные письма
      // memMeilsMap.forEach(item => { Logger.log(JSON.stringify(item, null, 2)) });
      let memProductRequests = new Array();

      memMeilsMap.forEach(m => {
        m.Номенклатуры.forEach(item => {
          if (!item) { return; }
          item.Письмо = m.key;
          memProductRequests.push(item);
        });
      });



      if (!this.sheetModel_Отправленные_Письма) {
        let sheetName = getSettings().sheetName_Отпр_письма;//"Отправленные письма";
        this.sheetModel_Отправленные_Письма = MrLib_Midlle.makeSheetModelBy(url, sheetName);
      }
      this.sheetModel_Отправленные_Письма.updateItems([...memMeilsMap.values()]);
      this.addToMemSentProductRequests(memProductRequests, url);
    } catch (err) { mrErrToString(err); }
  }


  getSheetModelОтпрПисьма(url) {
    let url_Письма = url;
    let sheetName_Отпр_письма = getSettings().sheetName_Отпр_письма;
    let sheetModel_Отпр_письма = MrLib_Midlle.makeSheetModelBy(url_Письма, sheetName_Отпр_письма);


    let requiredColNamesArr = [
      "key",
      "from",
      "to",
      "ДатаОтправки",
      "url",
      "mail",
      "Звонок",
      // "last",
    ];
    sheetModel_Отпр_письма.requiredColNames(requiredColNamesArr);
    sheetModel_Отпр_письма.sheet.hideSheet();
    return sheetModel_Отпр_письма;
  }


  getSheetModelОтпрЗапросы(url) {
    let url_Запросы = url;
    let sheetName_Отпр_запросы = getSettings().sheetName_Отпр_запросы;
    let sheetModel_Отпр_запросы = MrLib_Midlle.makeSheetModelBy(url_Запросы, sheetName_Отпр_запросы);


    let requiredColNamesArr = [

      // last
      "key",
      "НомерПроекта",
      "from",
      "to",
      "emailAddress",
      "url",
      "Номенклатура",
      "НоменклатураТекст",
      "ДатаОтправки",
      "Письмо",
      "Звонок",
    ];
    sheetModel_Отпр_запросы.requiredColNames(requiredColNamesArr);
    sheetModel_Отпр_запросы.sheet.hideSheet();
    return sheetModel_Отпр_запросы;
  }


  getSheetModelРасс_Закрытие(url) {
    let url_ = url;
    let sheetName = getSettings().sheetName_Расс_Зак;
    let sheetModel = MrLib_Midlle.makeSheetModelBy(url_, sheetName);
    let requiredColNamesArr = [
      "key",
      "НомерПроекта",
      "from",
      "to",
      "url",
      "ДатаОтправки",
    ];
    sheetModel.requiredColNames(requiredColNamesArr);
    sheetModel.sheet.hideSheet();
    return sheetModel;
  }


  getSheetModelРасс_Продолжение(url) {
    let url_ = url;
    let sheetName = getSettings().sheetName_Расс_Прод;
    let sheetModel = MrLib_Midlle.makeSheetModelBy(url_, sheetName);
    let requiredColNamesArr = [
      "key",
      "НомерПроекта",
      "from",
      "to",
      "url",
      "ДатаОтправки",
    ];
    sheetModel.requiredColNames(requiredColNamesArr);
    sheetModel.sheet.hideSheet();
    return sheetModel;
  }




  menuРассылкаЗакрытиеПроекта() {
    Logger.log(` menuРассылкаЗакрытиеПроекта  Start | From:${getContext().getUserEmail()}`);
    let snp = getContext().getSheetНастройки_писем();

    let mailData = snp.getMailDataFinish();
    if (!mailData) {
      throw `Для ${getContext().getUserEmail()} нет шаблона`;
    }

    let sl1 = getContext().getSheetList1();
    let emails1 = [...sl1.getEmailProductMap().keys()].map(k => {
      return `${k}`.trim().toLowerCase().split(" ").filter(e => {
        return isEmail(e);
      });
    }).flat();
    let emailsSet1 = new Set();  // Всем от кого есть ответы на листе 1 
    emails1.forEach(e => emailsSet1.add(e));


    let sl7 = getContext().getSheetList7();
    let emails7 = [].concat(sl7.getContragentEmail(), sl7.getContragentRezEmail())
    emails7 = emails7.map(k => {
      return `${k}`.trim().toLowerCase().split(" ").filter(e => {
        return isEmail(e);
      });
    }).flat();
    let emailsSet7 = new Set();  // Всем от кого есть ответы на листе 7 
    emails7.forEach(e => emailsSet7.add(e));
    // Logger.log(JSON.stringify([...emailsSet7.keys()]));


    let mapToFroms = new Map();

    let msОтпрПисьма = snp.getSheetModelОтпрПисьма(getContext().getSpreadSheet().getUrl())
    msОтпрПисьма.getMap().forEach(v => {
      // Logger.log([v.from != getContext().getUserEmail(), v.from, v.to, getContext().getUserEmail()]);
      if (v.from != getContext().getUserEmail()) { return; }

      mapToFroms.set(v.to, v.from);
    });



    let emailsSetЗакрытие = new Set();
    let msРасс_Закрытие = snp.getSheetModelРасс_Закрытие(getContext().getSpreadSheet().getUrl())
    msРасс_Закрытие.getMap().forEach((v, k) => { emailsSetЗакрытие.add(k); });

    let isDefUser = (() => {
      let memNote = this.sheet.getRange(this.strRanges.triggersMem).getNote();
      memNote = (() => { try { return JSON.parse(memNote) } catch (err) { mrErrToString(err); return false; } })();

      if (!memNote) { return false; }
      if (memNote.triggerMem.ТекущийПользователь != getContext().getUserEmail()) {
        return false;
      }
      Logger.log(`memNote.triggerMem.ТекущийПользователь ${memNote.triggerMem.ТекущийПользователь}`);
      return true;
    })();


    let emailSendMap = new Map();

    emailsSet1.forEach(em => {
      // Logger.log(em);

      if (emailsSet7.has(em)) { return; }
      if (emailsSetЗакрытие.has(em)) { return; }
      if (!mapToFroms.has(em)) {
        // Logger.log(`isDefUser= ${isDefUser}`);
        if (!isDefUser) { return; }
      }
      // Logger.log([em, "ok"]);
      emailSendMap.set(em, {
        "key": em,
        "НомерПроекта": getContext().getНомерПроекта(),
        "from": getContext().getUserEmail(),
        "to": em,
        "url": getContext().getSpreadSheet().getUrl(),
        "ДатаОтправки": getContext().timeConstruct,
      });
    });

    // Logger.log(JSON.stringify([...emailSendMap.keys()]));

    if (emailSendMap.size == 0) { return; }




    let mail = {
      to: undefined,
      replyTo: mailData.replyTo,
      subject: mailData.subject,
      htmlBody: mailData.htmlBody,
      attachments: mailData.attachments,
    };

    emailSendMap.forEach((v, k) => {
      mail.to = k;
      snp.sendMail(mail);
    });

    // snp.setFlagByRange(snp.strRanges.mailSent_Finish);
    msРасс_Закрытие.updateItems([...emailSendMap.values()]);

    msРасс_Закрытие.sheet.hideSheet();
  }


  menuРассылкаПродолжениеПроекта() {
    Logger.log(` menuРассылкаПродолжениеПроекта  Start | From:${getContext().getUserEmail()}`);
    let snp = getContext().getSheetНастройки_писем();


    let mailData = snp.getMailDataContinue();
    if (!mailData) {
      throw `Для ${getContext().getUserEmail()} нет шаблона`;
    }
    // let fl = snp.getFlagMemNoteByRange(snp.strRanges.mailSent_Continue);
    // // Logger.log(`fl=${fl}`);
    // if (fl) { Logger.log("Уже отправляли Email"); return; }
    // return;


    let sl7 = getContext().getSheetList7();
    let emails7 = [].concat(sl7.getContragentEmail(), sl7.getContragentRezEmail())
    emails7 = emails7.map(k => {
      return `${k}`.trim().toLowerCase().split(" ")
        .map(e => e.trim())
        .filter(e => { return isEmail(e); })
        .flat(2).map(e => {
          return `${e}`.split(",")
            .map(e => e.trim())
            .filter(e => { return isEmail(e); }).flat(2);
        });
    }).flat(4);




    let emailsSet7 = new Set();  // Всем от кого есть ответы на листе 1 
    emails7.forEach(e => {
      if (!isEmail(e)) { return; }
      emailsSet7.add(e);
    });

    let isDefUser = (() => {
      let memNote = this.sheet.getRange(this.strRanges.triggersMem).getNote();
      memNote = (() => { try { return JSON.parse(memNote) } catch (err) { mrErrToString(err); return false; } })();
      // Logger.log(`memNote.triggerMem.ТекущийПользователь ${memNote.triggerMem.ТекущийПользователь}`);
      if (!memNote) { return false; }
      if (memNote.triggerMem.ТекущийПользователь != getContext().getUserEmail()) {
        return false;
      }
      // Logger.log(`memNote.triggerMem.ТекущийПользователь ${memNote.triggerMem.ТекущийПользователь}`);
      return true;
    })();



    let mapToFroms = new Map();

    let msОтпрПисьма = snp.getSheetModelОтпрПисьма(getContext().getSpreadSheet().getUrl())
    msОтпрПисьма.getMap().forEach(v => {
      // Logger.log([v.from != getContext().getUserEmail(), v.from, v.to, getContext().getUserEmail()]);
      if (v.from != getContext().getUserEmail()) { return; }

      mapToFroms.set(v.to, v.from);
    });

    let emailsSetПрод = new Set();
    let smРасс_Прод = snp.getSheetModelРасс_Продолжение(getContext().getSpreadSheet().getUrl());
    smРасс_Прод.getMap().forEach((v, k) => { emailsSetПрод.add(k); });

    // Logger.log(JSON.stringify([...emailsSet7.keys()]));

    let emailSendMap = new Map();
    emailsSet7.forEach(em => {
      if (emailsSetПрод.has(em)) { return; }
      // if (!mapToFroms.has(em)) { return; }
      if (!mapToFroms.has(em)) {
        // return
        if (!isDefUser) { return; }
      }

      emailSendMap.set(em, {
        "key": em,
        "НомерПроекта": getContext().getНомерПроекта(),
        "from": getContext().getUserEmail(),
        "to": em,
        "url": getContext().getSpreadSheet().getUrl(),
        "ДатаОтправки": getContext().timeConstruct,
      });
    });

    Logger.log(JSON.stringify([...emailSendMap.keys()]));

    // return;
    if (emailSendMap.size == 0) { return; }


    let mail = {
      to: undefined,
      replyTo: mailData.replyTo,
      subject: mailData.subject,
      htmlBody: mailData.htmlBody,
      attachments: mailData.attachments,
    };

    emailSendMap.forEach((v, k) => {
      mail.to = k;
      snp.sendMail(mail);
    });

    // snp.setFlagByRange(snp.strRanges.mailSent_Continue);
    smРасс_Прод.updateItems([...emailSendMap.values()]);
    smРасс_Прод.sheet.hideSheet();
  }



}








// /**
//  * {
//   "triggerMem": {
//     "ПоследнееВыполнение": "2023-05-17T15:15:50.115Z",
//     "ТекущийПользователь": "zakupka@haiverk.ru",
//     "ВладелецТаблици": "zakupka@haiverk.ru",
//     "Тригеры": [
//       "triggerMrOnEditHelperExternal",
//       "triggerMrOnEditHelper",
//       "MrLib_top.triggerService",
//       "onEdit_trigger"
//     ],
//     "d": "1mXwNcxsjg"
//   }
// }
//  */



class DataClassFinishMem {
  constructor() {
    this.ПоследнееВыполнение = getContext().timeConstruct;
    this.ТекущийПользователь = (() => { return Session.getEffectiveUser().getEmail() })();
    this.ВладелецТаблици = (() => { return SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() })();
    this.d = SpreadsheetApp.getActiveSpreadsheet().getId().slice(0, 10);
  }
}


function service_setTriggerInfo_() {
  Logger.log(`service_setTriggerInfo_ Start `);
  let rangeMem = "K11";
  let memNote = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).getNote();
  memNote = (() => { try { return JSON.parse(memNote); } catch { return new Object(); } })();
  memNote.triggerMem = new DataClassTriggerMem();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).setNote(JSON.stringify(memNote, undefined, 2));
  Logger.log(`service_setTriggerInfo_ Finish = ${JSON.stringify(memNote)}`);

}


function service_updateСтатусыЗвонков() {
  getContext().getSheetНастройки_писем().updateСтатусыЗвонков();
  // tempupdateНастройки_писем();
}

function tempupdateНастройки_писем() {

  // if (!isTest()) { return; }


  // let sheet = getContext().getSheetНастройки_писем().sheet;
  // let тампусто = "&" + (() => { try { return sheet.getRange("AQ1:AT6").getValues().flat(2).join("") } catch (err) { return mrErrToString(err); } })();
  // if (тампусто == "&") { return; }
  // sheet.getRange("AQ1:AT6").clear();
  // sheet.getRange("AQ1:AT6").clearNote();
  // sheet.getRange("AQ1:AT6").clearDataValidations()

  // let vls = [
  //   ["Признак", "Показать", "Порядок"],
  //   ["♥️", true, 0],
  //   ["♠", true, 1],
  //   ["♣️", true, 2],
  //   ["♦️", true, 3],
  //   ["-", true, 4],
  // ]
  // sheet.getRange("AQ1:AS6").setValues(vls);
  // // sheet.getRange("AR2:AR6").  

  // let rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  // sheet.getRange("AR2:AR6").setDataValidation(rule);


}


/**
 * Задача 4.14 - Сделать авторассылку о закрытии проекта		
на листе настройка писем сделать шаблон (тема текст письма вложение на какой ящик ответ, аналогично текущей рассылке)

Всем от кого есть ответы на листе 1 
и кого нет на листе выбор поставщиков в столбце K и L отправить письмо по шаблону…

Учесть что емейлы могут быть двойные через пробел

Сделать настройку (галочку) на листе настройка писем рассылать ли автоответ. 

Сделать автоотправку, вопрос с какого ящика, 
когда все товары либо оплачены, либо частично оплачены, 
либо выбран “Текущий статус проекта” на листе 0,
либо выбран “статус расчета” на листе 0, 
либо по кнопке из меню

Сделать флаг, что автоответ отправлен, 
чтоб не было повторной отправки (например если два раза нажать в меню)

 */


function servicesАвтоРассылкаЗакрытияПродолженияПроекта() {
  Logger.log(` servicesАвтоРассылкаЗакрытияПроекта  Start`);

  let cSheet = new MrClassSheetПаспортПроекта();
  let map = new Map();
  cSheet.getMap().forEach((v, k) => { map.set(fl_str(k), v) });
  let fl = (() => {
    // Logger.log(map.get(fl_str("Текущий статус проекта")))
    if (map.get(fl_str("Текущий статус проекта"))) {
      if (map.get(fl_str("Текущий статус проекта"))["value"]) { return true; }
    }
    if (map.get(fl_str("статус расчета"))) {
      if (map.get(fl_str("статус расчета"))["value"]) { return true; }
    }
    // return false;

    let paymentsforAllProducts = true;
    // Logger.log(getContext().getSheetPay().getMapPercenOfPayment().size)
    getContext().getSheetPay().getMapPercenOfPayment().forEach((v, k) => {
      // Logger.log(v);
      // if (!paymentsforAllProducts) { return; }
      if (v) { return; }
      paymentsforAllProducts = false;
    });
    if (paymentsforAllProducts === true) { return true; }
    return false;
  })();

  Logger.log(` servicesАвтоРассылкаЗакрытияПроекта  fl=${fl}`);
  if (!fl) { return; }

  return;
  let snp = getContext().getSheetНастройки_писем();
  if (snp.getFlagByRange(snp.strRanges.auto_Finish)) { menuРассылкаЗакрытиеПроекта(); }
  if (snp.getFlagByRange(snp.strRanges.auto_Continue)) { menuРассылкаПродолжениеПроекта(); }

}

function menuРассылкаЗакрытиеПроекта() {
  Logger.log(` menuРассылкаЗакрытиеПроекта  Start`);

  let snp = getContext().getSheetНастройки_писем();
  snp.menuРассылкаЗакрытиеПроекта();
  return;
  let fl = snp.getFlagMemNoteByRange(snp.strRanges.mailSent_Finish);
  // Logger.log(`fl=${fl}`);
  if (fl) { Logger.log("Уже отправляли Email"); return; }


  let sl1 = getContext().getSheetList1();

  let emails1 = [...sl1.getEmailProductMap().keys()].map(k => {
    return `${k}`.trim().toLowerCase().split(" ").filter(e => {
      return isEmail(e);
    });
  }).flat();
  let emailsSet1 = new Set();  // Всем от кого есть ответы на листе 1 
  emails1.forEach(e => emailsSet1.add(e));


  let sl7 = getContext().getSheetList7();

  let emails7 = [].concat(sl7.getContragentEmail(), sl7.getContragentRezEmail())

  emails7 = emails7.map(k => {
    return `${k}`.trim().toLowerCase().split(" ").filter(e => {
      return isEmail(e);
    });
  }).flat();


  let emailsSet7 = new Set();  // Всем от кого есть ответы на листе 1 
  emails7.forEach(e => emailsSet7.add(e));

  // Logger.log(JSON.stringify([...emailsSet7.keys()]));

  let emailSendMap = new Map();

  emailsSet1.forEach(em => {
    if (emailsSet7.has(em)) { return; }
    emailSendMap.set(em, new Object());
  });

  Logger.log(JSON.stringify([...emailSendMap.keys()]));

  // return;
  if (emailSendMap.size == 0) { return; }


  let mailData = snp.getMailDataFinish();

  let mail = {
    to: undefined,
    replyTo: mailData.replyTo,
    subject: mailData.subject,
    htmlBody: mailData.htmlBody,
    attachments: mailData.attachments,
  };

  emailSendMap.forEach((v, k) => {
    mail.to = k;
    snp.sendMail(mail);
  });

  snp.setFlagByRange(snp.strRanges.mailSent_Finish);

}


function menuРассылкаПродолжениеПроекта() {
  Logger.log(` menuРассылкаПродолжениеПроекта  Start`);
  let snp = getContext().getSheetНастройки_писем();

  snp.menuРассылкаПродолжениеПроекта();
  return;

  let fl = snp.getFlagMemNoteByRange(snp.strRanges.mailSent_Continue);
  // Logger.log(`fl=${fl}`);
  if (fl) { Logger.log("Уже отправляли Email"); return; }

  // return;


  let sl7 = getContext().getSheetList7();

  let emails7 = [].concat(sl7.getContragentEmail(), sl7.getContragentRezEmail())

  emails7 = emails7.map(k => {
    return `${k}`.trim().toLowerCase().split(" ")
      .map(e => e.trim())
      .filter(e => { return isEmail(e); })
      .flat(2).map(e => {
        return `${e}`.split(",")
          .map(e => e.trim())
          .filter(e => { return isEmail(e); }).flat(2);
      });
  }).flat(4);




  let emailsSet7 = new Set();  // Всем от кого есть ответы на листе 1 
  emails7.forEach(e => {
    if (!isEmail(e)) { return; }
    emailsSet7.add(e);
  });

  // Logger.log(JSON.stringify([...emailsSet7.keys()]));

  let emailSendMap = new Map();
  emailsSet7.forEach(em => {
    emailSendMap.set(em, new Object());
  });

  Logger.log(JSON.stringify([...emailSendMap.keys()]));

  // return;
  if (emailSendMap.size == 0) { return; }

  let mailData = snp.getMailDataContinue();
  let mail = {
    to: undefined,
    replyTo: mailData.replyTo,
    subject: mailData.subject,
    htmlBody: mailData.htmlBody,
    attachments: mailData.attachments,
  };

  emailSendMap.forEach((v, k) => {
    mail.to = k;
    snp.sendMail(mail);
  });

  snp.setFlagByRange(snp.strRanges.mailSent_Continue);

}








/** @param {[Object]} toMemArrMailsInMilde*/
function addMailInMidle(toMemArrMailsInMilde) {
  return;

  let url = MrLib_Midlle.getSettings().urls.Письма;
  let sheetName = "Отправленные";
  let sheetModel_Отправленные = MrLib_Midlle.makeSheetModelBy(url, sheetName);
  sheetModel_Отправленные.updateItems(toMemArrMailsInMilde)

}







function getSummary(maloOtvetov = 3) {
  let ret = getContext().getSheetList1().getSummary(maloOtvetov, getContext().getAnalogColor());
  ret = ret.concat(getContext().getSheetList7().getSummary(maloOtvetov, getContext().getAnalogColor()));
  return ret;
}

function serviceSummaryНастройки_писем() {
  getContext().getSheetНастройки_писем().serviceSummary();
}





