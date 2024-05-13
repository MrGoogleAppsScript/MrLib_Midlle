


function triggrMakeCopyАрхив() {
  let destinationID = "1lXBM";
  menuMakeCopyАрхив(destinationID);
}








function menuMakeCopyАрхив(destinationID = "1QpZM") {

  // generates the timestamp and stores in variable formattedDate as year-month-date hour-minute-second
  var formattedDate = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd' 'HH:mm:ss");

  // gets the name of the original file and appends the word "copy" followed by the timestamp stored in formattedDate
  var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " | Архив " + formattedDate;

  // gets the destination folder by their ID. REPLACE xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx with your folder's ID that you can get by opening the folder in Google Drive and checking the URL in the browser's address bar
  var destination = DriveApp.getFolderById(destinationID);

  // gets the current Google Sheet file
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())

  // makes copy of "file" with "name" at the "destination"
  let copiedFile = file.makeCopy(name, destination);
  return copiedFile;
}


function serviceАрхив() {
  if (getContext().timeConstruct.getHours() != 20) { return; }
  if (getContext().timeConstruct.getMinutes() >= 31) { return; }
  let item = {
    date: Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "yyyy-MM-dd' 'HH:mm:ss"),
    // key: `${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "yyyy-MM-dd")}|${getContext().getSpreadSheet().getId()}|${getContext().getSpreadSheet().getName()}`,
    key: `${Utilities.formatDate(getContext().timeConstruct, "Europe/Moscow", "yyyy-MM-dd")} | ${getContext().getSpreadSheet().getId()}`,
    name: getContext().getSpreadSheet().getName(),
    from: getContext().getSpreadSheet().getUrl(),
    arhiv: undefined,
    ver: 1,
  };

  let hasKey = MrLib_Midlle.getItemByKey("АрхивПроектов", item.key, "key", undefined);
  if (hasKey) {
    item.hasKey = hasKey;
    Logger.log(JSON.stringify(item, null, 2))
    return;
  }
  // Logger.log(JSON.stringify(item, null, 2));

  let arhivFile = menuMakeCopyАрхив();
  item.arhiv = arhivFile.getUrl();
  MrLib_Midlle.updateItems("АрхивПроектов", [item]);
  Logger.log(JSON.stringify(item, null, 2));
}




function serviceObserverDeleteTriggers() {
  let mem = getContext().getMem("TriggersStart");
  if (!mem) {
    mem = {
      start: {
        first: getContext().timeConstruct.getTime(),
        date: getContext().timeConstruct,
        ver: 1,
      },
    }
  }
  mem.start.last = getContext().timeConstruct.getTime();
  getContext().setMem("TriggersStart", mem);
  let fl_delete = false;
  if (mem.start.first + (1000 * 60 * 60 * 24 * 45) < mem.start.last) { // 45 суток
    fl_delete = true;
  }



  if (fl_delete !== true) {
    Logger.log(" не Удаляем по дате"); return;
  }

  Logger.log(" По дате: Удалять");
  let СтатусПроекта = getContext().getСтатусПроекта();
  let Исполнение_date = getContext().getИсполнение_date();
  let НомерПроекта = getContext().getНомерПроекта();

  Logger.log({ fl_delete, НомерПроекта, СтатусПроекта, Исполнение_date, d: "Удаляем по дате" });

  if (getContext().getНомерПроекта()) {
    let статусыЗакрытияПроекта = getContext().getСтатусыЗакрытияПроекта().map(v => fl_str(v));
    if (fl_str(СтатусПроекта) !== fl_str("")) {
      Logger.log("Статус не Пуст.")
      fl_delete = false;
      if (статусыЗакрытияПроекта.includes(fl_str(СтатусПроекта))) {
        Logger.log("По Статусу: Удаляем.")
        fl_delete = true;
      } else { Logger.log("По Статусу: Не Удаляем.") }
    } else {
      Logger.log("Статус Пуст: Не Удаляем.");
      fl_delete = false;
    }
  }

  Logger.log({ fl_delete, НомерПроекта, СтатусПроекта, Исполнение_date, d: "udalaem" });

  if (fl_delete !== true) {
    Logger.log("Итого не Удаляем.");
    return;
  }


  Logger.log("Итого Удаляем.");


  let mesaje = {
    "Метка Триггеры Добавлены": (() => {
      let da = fl_str(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange("K11").getValue());
      if (fl_str("ДА") == da) {
        return "ДА";
      } else {
        return "Нет";
      }
    })(),

    "Текущий пользователь": (() => { return Session.getEffectiveUser().getEmail() })(),
    "Владелец таблици": (() => { return SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() })(),
    "Список Триггеров текущего пользователя": (() => {
      let triggers = ScriptApp.getProjectTriggers();
      let triggersName = triggers.map(trigger => trigger.getHandlerFunction());
      if (triggersName.length == 0) { return "Нет Триггеров"; }
      return `Кол-во: ${triggersName.length},  ${triggersName.join(", ")}`;
    })(),


    "Список Триггеров других пользователя": (() => {
      let url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      MrLib_Midlle.getItemByKey("service", url, "НомерПроекта", undefined);
      // ;"email"	"triggersName"	"triggersName.length"
      let email = (() => { return MrLib_Midlle.getItemByKey("service", url, "email", "нет"); })();
      let triggersName = (() => { return MrLib_Midlle.getItemByKey("service", url, "triggersName", undefined); })();
      let triggersCount = (() => { return MrLib_Midlle.getItemByKey("service", url, "triggersName.length", undefined); })();

      // if (triggersName) {
      //   triggersName = (() => { try { JSON.parse(triggersName).join(", ") } catch { return triggersName } })();
      // }
      return `Пользователь: ${email}, Кол-во: ${triggersCount}, Триггеры: ${triggersName}`;
    })(),
    // MrLib_Midlle.getItemByKey("Проекты", url, "НомерПроекта", undefined);
  }
  Logger.log(JSON.stringify(mesaje, null, 2));

  // return;
  if (fl_delete == true) {
    deleteTrigger();
    getContext().setMem("TriggersStart", undefined);
  }

}

function getСтатусыЗакрытияПроекта() {
  getContext().getСтатусыЗакрытияПроекта();
}