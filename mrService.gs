
function triggerService() {
  try { tempupdateНастройки_писем(); } catch (err) { mrErrToString(err); }

  // this.ttst = [

  // ];

  // if (this.ttst.includes(SpreadsheetApp.getActiveSpreadsheet().getUrl())) { triggerService2(); return; }



  triggerService2(); return;

  // try { servicesСинхрШаблонаОплатыВнешней(); } catch (err) { }
  // paintUpdateBackgroundAllforSheetList7();
  try { serviceАрхив(); } catch (err) { mrErrToString(err); }
  try { serviceObserverDeleteTriggers(); } catch (err) { mrErrToString(err); }
  try { service_updateСтатусыЗвонков(); } catch (err) { mrErrToString(err); }
  try { serviceСинхрНомераПроекта_(); } catch (err) { mrErrToString(err); }



  // try { servicesПоискЗвонковПоПисьмамV2(); } catch (err) { mrErrToString(err); }
  try { servicesПоискЗвонковПоПисьмамV3(); } catch (err) { mrErrToString(err); }

  // try { servicesПоискЗвонковПоПисьмам(); } catch (err) { mrErrToString(err); }
  // try { servicesПоискЗвонков_(); } catch (err) { mrErrToString(err); }
  // try { servicesАктуализациЗвонков(); } catch (err) { mrErrToString(err); }
  try { servicesАктуализациЗвонковV2(); } catch (err) { mrErrToString(err); }
  try { triggerUpdateЗвонки(); } catch (err) { mrErrToString(err); }

  // MrLib_top.servicesПоискЗвонковПоПисьмам();
  // MrLib_top.servicesАктуализациЗвонков();
  try { serviceПодсчетНеразосланныхПисем_(); } catch (err) { mrErrToString(err); }
  try { triggerUpdateЭкспортВБитрикс_(); } catch (err) { mrErrToString(err); }
  // try { servicesАвтоРассылкаЗакрытияПроекта(); } catch (err) { mrErrToString(err); }

  service_look_triggers_();
  service_setTriggerInfo_();

  try { serviceCopyMemFromBuyExternalToProduct(); } catch (err) { mrErrToString(err); }
  try { getContext().getSheetList1().recalculate_report_malo(); } catch (err) { mrErrToString(err); }

  // try {
  //   [

  //   ].forEach(u => {
  //     SpreadsheetApp.openByUrl(u).getSheetByName("1-1 Сбор КП").getRange("A1").setValue(undefined);
  //   });

  // } catch (err) { }

  try { serviceДобавитьВопросПроблему(); } catch (err) { mrErrToString(err); }
  try { serviceСинзронизацияВопросов(); } catch (err) { mrErrToString(err); }

}











function triggerService2() {


  let cache = getContext().getCacheDocument();
  Logger.log(JSON.stringify(cache));
  if (!cache["IndexTriggerService"]) { cache["IndexTriggerService"] = 0; }
  if (!Array.isArray(cache["IndexTriggerServiceTry"])) { cache["IndexTriggerServiceTry"] = new Array(30); for (let i = 0; i < 30; i++) { cache["IndexTriggerServiceTry"][i] = 0; } }
  if (!cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]) { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]] = 0; }
  if (cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]] > 3) { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]] = 0; cache["IndexTriggerService"]++; }
  // cache["IndexTriggerService"]++;
  Logger.log(JSON.stringify(cache));

  while (true) {
    switch (cache["IndexTriggerService"]) {


      case 0:
        // try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n service_setTriggerInfo_"); service_setTriggerInfo_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("service_setTriggerInfo_ \n^^^^^^^^^^ ");

        // try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n service_look_triggers_"); service_look_triggers_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("service_look_triggers_ \n^^^^^^^^^^ ");

        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceАрхив"); serviceАрхив(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceАрхив \n^^^^^^^^^^ ");

      case 1:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceObserverDeleteTriggers"); serviceObserverDeleteTriggers(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceObserverDeleteTriggers \n^^^^^^^^^^ ");
      case 2:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n service_updateСтатусыЗвонков"); service_updateСтатусыЗвонков(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("service_updateСтатусыЗвонков \n^^^^^^^^^^ ");
      case 3:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceСинхрНомераПроекта_"); serviceСинхрНомераПроекта_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceСинхрНомераПроекта_ \n^^^^^^^^^^ ");
      case 4:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n servicesПоискЗвонковПоПисьмамV3"); servicesПоискЗвонковПоПисьмамV3(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("servicesПоискЗвонковПоПисьмамV3 \n^^^^^^^^^^ ");
      case 5:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n servicesАктуализациЗвонковV2"); servicesАктуализациЗвонковV2(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("servicesАктуализациЗвонковV2 \n^^^^^^^^^^ ");
      case 6:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n triggerUpdateЗвонки"); triggerUpdateЗвонки(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("triggerUpdateЗвонки \n^^^^^^^^^^ ");
      case 7:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceПодсчетНеразосланныхПисем_"); serviceПодсчетНеразосланныхПисем_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceПодсчетНеразосланныхПисем_ \n^^^^^^^^^^ ");
      case 8:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n triggerUpdateЭкспортВБитрикс_"); triggerUpdateЭкспортВБитрикс_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("triggerUpdateЭкспортВБитрикс_ \n^^^^^^^^^^ ");
      case 9:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n service_look_triggers_"); service_look_triggers_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("service_look_triggers_ \n^^^^^^^^^^ ");
      case 10:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n service_setTriggerInfo_"); service_setTriggerInfo_(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("service_setTriggerInfo_ \n^^^^^^^^^^ ");
      case 11:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceCopyMemFromBuyEx"); serviceCopyMemFromBuyExternalToProduct(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceCopyMemFromBuyEx \n^^^^^^^^^^ ");
      case 12:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n recalculate_report_malo"); getContext().getSheetList1().recalculate_report_malo(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("recalculate_report_malo \n^^^^^^^^^^ ");
      case 13:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceДобавитьВопросПроблему"); serviceДобавитьВопросПроблему(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceДобавитьВопросПроблему \n^^^^^^^^^^ ");
      case 14:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceСинзронизацияВопросов"); serviceСинзронизацияВопросов(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceСинзронизацияВопросов \n^^^^^^^^^^ ");
      case 15:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceСинзронизацияПроблем"); serviceСинзронизацияПроблем(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceСинзронизацияПроблем \n^^^^^^^^^^ ");
      case 16:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceРасчетРентабельностиКП"); serviceРасчетРентабельностиКП(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceРасчетРентабельностиКП \n^^^^^^^^^^ ");
      case 17:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceДатаДобавленияПроблемыВопроса"); serviceДатаДобавленияПроблемыВопроса(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceДатаДобавленияПроблемыВопроса \n^^^^^^^^^^ ");

      case 18:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceSummaryНастройки_писем"); serviceSummaryНастройки_писем(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceSummaryНастройки_писем \n^^^^^^^^^^ ");
      case 19:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n servisВертикаль"); servisВертикаль(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("servisВертикаль \n^^^^^^^^^^ ");

      case 20:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceSyncОплатыВнешнейИЛогистики"); serviceSyncОплатыВнешнейИЦентральной(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceSyncОплатыВнешнейИЛогистики \n^^^^^^^^^^ ");


      case 21:
        try { cache["IndexTriggerServiceTry"][cache["IndexTriggerService"]]++; Logger.log("VVVVVVV  \n serviceSyncПроектаИЦентральной"); serviceSyncПроектаИЦентральной(); } catch (err) { mrErrToString(err); } cache["IndexTriggerService"]++; getContext().saveCacheDocument(); Logger.log("serviceSyncПроектаИЦентральной \n^^^^^^^^^^ ");



// serviceSyncПроектаИЦентральной()
      default:
        cache["IndexTriggerServiceTry"] = undefined;
        cache["IndexTriggerService"] = 0;
        getContext().saveCacheDocument();
        return;
        break;
    }
  }

}









































class DataClassTriggerMem {
  constructor() {
    this.ПоследнееВыполнение = getContext().timeConstruct;
    this.ТекущийПользователь = (() => { return Session.getEffectiveUser().getEmail() })();
    this.ВладелецТаблици = (() => { return SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() })();
    this.Тригеры = ScriptApp.getProjectTriggers().map(trigger => trigger.getHandlerFunction());
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


function service_look_triggers_() {
  let triggers = ScriptApp.getProjectTriggers();
  let triggersName = triggers.map(trigger => trigger.getHandlerFunction()).sort();
  let scriptId = ScriptApp.getScriptId();
  let url = getContext().getSpreadSheet().getUrl();
  let email = Session.getEffectiveUser().getEmail();
  let owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  let name = SpreadsheetApp.getActiveSpreadsheet().getName();

  let НомерПроекта = MrLib_Midlle.getItemByKey("Проекты", url, "НомерПроекта", undefined);
  let url_pay;
  let url_buy;
  let sheetNames = getContext().getSpreadSheet().getSheets().map(s => s.getSheetName()).sort();
  let sheetNamesPay;
  let sheetNamesBuy;

  let externalSpreadSheetPay = getContext().getSheetPay().getExternalSpreadSheet();
  let externalSpreadSheetBuy = getContext().getSheetBuy().getExternalSpreadSheet();

  if (externalSpreadSheetBuy) {
    url_buy = externalSpreadSheetBuy.getUrl();
    sheetNamesBuy = externalSpreadSheetBuy.getSheets().map(s => s.getSheetName()).sort();
  }

  if (externalSpreadSheetPay) {
    url_pay = externalSpreadSheetPay.getUrl();
    sheetNamesPay = externalSpreadSheetPay.getSheets().map(s => s.getSheetName()).sort();
  }

  // let вопросы = undefined;
  // let fr = getSettings().sheetName_Вопросы;
  // let to = getSettings().sheetName_Вопросы_2_1;
  // try {
  //   getContext().getSheetByName(fr);
  //   вопросы = fr;
  // } catch (err) { mrErrToString(err); }

  // try {
  //   getContext().getSheetByName(to);
  //   вопросы = вопросы + " " + to;
  // } catch (err) { mrErrToString(err); }
  let sItem = {
    key: url,
    name,
    scriptId,
    url,
    email,
    owner,
    date: getContext().timeConstruct,
    triggersName,
    НомерПроекта,
    url_pay,
    url_buy,
    sheetNamesPay,
    sheetNamesBuy,
    sheetNames,
    // вопросы,
    ver: 1,
  };
  Logger.log(JSON.stringify(sItem, null, 2));

  MrLib_Midlle.updateItems("service", [sItem]);

  if (НомерПроекта) {
    let nnn = new Array();
    urls = getContext().getUrls();
    for (let k in urls) {
      if (!urls[k]) { continue; }
      if (k == "callCoord") { continue; }
      nnn.push({
        key: standartUrl(urls[k]),
        type: k,
        НомерПроекта,
        ver: 1,
        date: getContext().timeConstruct,
      });
    }
    MrLib_Midlle.updateItems("НомераПроектов", nnn);
  }
  // НомераПроектов



  // MrLib_Midlle.updateItems("Черновик", [{ key: НомерПроекта, url, url_pay, url_buy, sheetNamesPay, sheetNamesBuy, nameDoc, timeConstruct: getContext().timeConstruct }]);
  // MrLib_Midlle.updateItems("service", [{ key: НомерПроекта, url, url_pay, url_buy, sheetNamesPay, sheetNamesBuy, nameDoc, timeConstruct: getContext().timeConstruct }]);



}



function serviceСинхрНомераПроекта_() {
  try {
    let url = getContext().getSpreadSheet().getUrl();
    let nameDoc = getContext().getSpreadSheet().getName();

    let externalSpreadSheetPay = getContext().getSheetPay().getExternalSpreadSheet();
    // let externalSpreadSheetBuy = getContext().getSheetBuy().getExternalSpreadSheet();

    let url_pay;
    let НомерПроекта = getContext().getНомерПроекта();
    if (externalSpreadSheetPay) {
      url_pay = externalSpreadSheetPay.getUrl();

      let sheet_Исполнение_Внешние = externalSpreadSheetPay.getSheetByName(getContext().settings.sheetName_Исполнение_Внешние);
      if (sheet_Исполнение_Внешние) {

        let items = new Array();
        items.push({
          row: 3,
          col: "B",
          key: "Проект",
          Значение: НомерПроекта,
        });
        items.push(
          {
            row: 3,
            col: "G",
            key: "Время Последней Синхронизации",
            Значение: getContext().timeConstruct,
          });

        // items.push({

        // row: 3,
        // col: "C",
        //   key: "Плательщик",
        //   Значение: getContext().getSheetPay().getПлательщики().join("\n"),
        // });

        items.push({
          row: 3,
          col: "H",
          key: "Победил",
          Значение:
            // ((json_Реестр) => { try { return JSON.parse(json_Реестр)["Закупает"] } catch (err) { return undefined } })(MrLib_Midlle.getItemByKey("Сводная", НомерПроекта, Реестр)),
            ((json_Исполнение) => { try { return JSON.parse(json_Исполнение)["Победил"]; } catch (err) { return mrErrToString(err); } })(MrLib_Midlle.getItemByKey("Сводная", `${НомерПроекта}`, "Исполнение")),
        });

        items.push({
          row: 3,
          col: "E",
          key: "Таблица Основная",
          Значение: nameDoc,
        });

        items.push({
          row: 3,
          col: "F",
          key: "Ссылка",
          Значение: url,
        });


        items.push({
          row: 3,
          col: "I",
          key: "Вид подписания договора",
          Значение: "",
          // ((json_Исполнение) => {
          //   try {
          //     return JSON.parse(json_Исполнение)["Вид подписания договора"];
          //   } catch (err) {
          //     mrErrToString(err);
          //     return;
          //   }
          // })(MrLib_Midlle.getItemByKey("Сводная", `${НомерПроекта}`, "Исполнение")),
        });


        // items.push({
        //   row: 3,
        //   col: "K",
        //   key: "json_Исполнение",
        //   Значение: "",

        // });

        items.forEach(item => {
          sheet_Исполнение_Внешние.getRange(item.row, nr(item.col)).setValue(item.Значение);
        });

      }




    }




  } catch (err) {
    Logger.log(mrErrToString(err));
  }


}








/**
 * Задача 4.12 - Сделать автопостановку задачи на звонок		
 * 1. Если после автоматической рассылки писем прошло больше 1 рабочего дня (настраиваемый параметр),
 * а у товара меньше 3х цен то необходимо поставить задачу на обзвон тех поставщиков? 
 * от которых нет ответа, для этого копируем емейлы кому надо позвонить на лист 1-3 звонки. 
 * Вставляем название товаров по которому был отправлен запрос. 
 */
function servicesПоискЗвонков_() {
  let sheetЗвонки = new MrClassSheetЗвонки();
  sheetЗвонки.servicesПоискЗвонков();
}


function servicesПоискЗвонковПоПисьмам() {

  let url = getContext().getSpreadSheet().getUrl();

  if (url) {
    let prefix = "";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesПоискЗвонковПоПисьмам();
  } else { }

  url = (() => {
    try {
      return getContext().getUrls().buy;
    } catch (err) {
      mrErrToString(err);
      return undefined;
    }
  })();
  if (url) {
    let prefix = "В-";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesПоискЗвонковПоПисьмам();
  } else {
    Logger.log("Нет Внешней таблицы Закупок");
  }
}



function servicesАктуализациЗвонков() {

  let url = getContext().getSpreadSheet().getUrl();

  if (url) {
    let prefix = "";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesАктуализациЗвонков();
  } else { }
  // return;
  url = (() => {
    try {
      return getContext().getUrls().buy;
    } catch (err) {
      mrErrToString(err);
      return undefined;
    }
  })();
  if (url) {
    let prefix = "В-";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesАктуализациЗвонков();
  } else {
    Logger.log("Нет Внешней таблицы Закупок");
  }
}




// function servicesОтменаЗвонков() {
//   let sheetЗвонки = new MrClassSheetЗвонки();
//   sheetЗвонки.servicesОтменаЗвонков();
// }


// function renameSheet() {
//   let fr = getSettings().sheetName_Вопросы;
//   let to = getSettings().sheetName_Вопросы_2_1;
//   try {
//     getContext().getSheetByName(fr).setName(to);
//   } catch (err) { mrErrToString(err); }

//   try {
//     getContext().getSheetByName(to);
//   } catch (err) { mrErrToString(err); }


// }


function servicesПоискЗвонковПоПисьмамV2() {

  let url = getContext().getSpreadSheet().getUrl();

  if (url) {
    let prefix = "";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesПоискЗвонковПоПисьмамV2();
  } else { }
  // // return;
  url = (() => {
    try {
      return getContext().getUrls().buy;
    } catch (err) {
      mrErrToString(err);
      return undefined;
    }
  })();
  if (url) {
    let prefix = "В-";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesПоискЗвонковПоПисьмамV2();
  } else {
    Logger.log("Нет Внешней таблицы Закупок");
  }
}


function servicesАктуализациЗвонковV2() {

  let url = getContext().getSpreadSheet().getUrl();

  if (url) {
    let prefix = "";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesАктуализациЗвонковV2();
  } else { }
  // return;
  url = (() => {
    try {
      return getContext().getUrls().buy;
    } catch (err) {
      mrErrToString(err);
      return undefined;
    }
  })();
  if (url) {
    let prefix = "В-";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesАктуализациЗвонковV2();
  } else {
    Logger.log("Нет Внешней таблицы Закупок");
  }
}



function servicesПоискЗвонковПоПисьмамV3() {

  let url = getContext().getSpreadSheet().getUrl();

  if (url) {
    let prefix = "";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesПоискЗвонковПоПисьмамV3();
  } else { }
  // return;
  url = (() => {
    try {
      return getContext().getUrls().buy;
    } catch (err) {
      mrErrToString(err);
      return undefined;
    }
  })();
  if (url) {
    let prefix = "В-";
    let sheetЗвонки = new MrClassSheetЗвонки(url, prefix);
    sheetЗвонки.servicesПоискЗвонковПоПисьмамV2();
  } else {
    Logger.log("Нет Внешней таблицы Закупок");
  }
}


function serviceДобавитьВопросПроблему() {
  getContext().getSheetVoprosi().service();
  // getContext().getSheetVoprosi().serviceСинзронизацияВопросов();
  getContext().getSheetПроблемы().service();

  try {
    let urlBuy = getContext().getUrls().buy;
    if (urlBuy) {
      let shПоблемы = new ClassSheet_Проблемы(urlBuy);
      let shВопросы = new ClassSheet_Voprosi(urlBuy);
      shПоблемы.service();
      shВопросы.service();
    }
  } catch (err) {
    mrErrToString(err)
  }

}




function serviceСинзронизацияВопросов() {
  getContext().getSheetVoprosi().serviceСинзронизацияВопросов();
}



function serviceСинзронизацияПроблем() {
  getContext().getSheetПроблемы().serviceСинзронизацияПроблем();
}



function serviceДатаДобавленияПроблемыВопроса() {

  // getContext().getSheetVoprosi().ab();
  // SpreadsheetApp.flush();
  //   getContext().getSheetVoprosi().ab();
  Logger.log("serviceДатаДобавления S");



  getContext().getSheetПроблемы().serviceДатаДобавления();
  getContext().getSheetVoprosi().serviceДатаДобавления();

  try {
    let urlBuy = getContext().getUrls().buy;
    if (urlBuy) {
      let shПоблемы = new ClassSheet_Проблемы(urlBuy);
      let shВопросы = new ClassSheet_Voprosi(urlBuy);
      shПоблемы.serviceДатаДобавления();
      shВопросы.serviceДатаДобавления();
    }
  } catch (err) {
    mrErrToString(err)
  }

  Logger.log("serviceДатаДобавления F");

  try {
    Logger.log("exportToMidlle S");
    getContext().getSheetПроблемы().exportToMidlle();
    getContext().getSheetVoprosi().exportToMidlle();
    Logger.log("exportToMidlle F");
  } catch (err) { mrErrToString(err) }

}


function serviceSyncОплатыВнешнейИЦентральной() {

  let theUrlExternalPay = getContext().getUrls().pay;
  let theProjNumber = getContext().getНомерПроекта();

  Logger.log(JSON.stringify({ theProjNumber, theUrlExternalPay }));
  if (!theUrlExternalPay) {
    Logger.log(`Нет Таблици "Внешнеей оплаты"`);
    return
  }

  if (!theProjNumber) {
    Logger.log(`Нет номера проекта`);
    return;
  }

  // return
  MrLib_External.serviceSyncЛогистика(theUrlExternalPay, theProjNumber);
}





function serviceSyncПроектаИЦентральной() {

  let theUrlProj = getContext().getUrls().product;
  let theProjNumber = getContext().getНомерПроекта();

  Logger.log(JSON.stringify({ theProjNumber, theUrlProj }));
  if (!theUrlProj) {
    Logger.log(`Нет Таблици "Проекта"`);
    return
  }

  if (!theProjNumber) {
    Logger.log(`Нет номера проекта`);
    return;
  }

  theProjNumber = `${theProjNumber}`;
  // // key	Таблица Проекта	Проект
  let theTasks = [

    ((theTask) => {
      theTask.param.projNumber = theProjNumber;
      theTask.param.projHead = "Проект";
      theTask.param.urlHead = "Таблица Проекта";
      theTask.param.urls.proj = theUrlProj;
      theTask.param.sheetNames.proj = "ОтчетыПроектов";
      theTask.param.sheetNames.central = "ОтчетыПроектов";
      return theTask;
    })(MrLib_Midlle.getDataTaskSyncПроектаСЦентральной()),

  ];


  // return;
  Logger.log("+======================================================================================");
  MrLib_Midlle.ВыполненитьМасивЗадач("serviceSyncПроектаИЦентральной", theTasks, "serviceSyncПроектаИЦентральной");

}




function testCache() {

  let cache = getContext().getCacheDocument();
  Logger.log(JSON.stringify(cache));
  if (!cache["tt"]) { cache["tt"] = 0; }
  cache["tt"]++;
  Logger.log(JSON.stringify(cache));
  getContext().saveCacheDocument();

  switch (cache["tt"]) {
    case 0: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 1: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 2: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 3: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 4: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 5: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 6: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 7: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 8: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    case 9: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    // case 8: Logger.log(JSON.stringify(cache)); cache["tt"]++; getContext().saveCacheDocument();
    default:
      cache["tt"] = 0;
      break;
  }


}