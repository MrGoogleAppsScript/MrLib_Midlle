function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('Email рассылка')
    // .addItem('Отправить письма', 'createEmail')
    .addItem('Отправить письма c файлами', 'MrLib_top.menuTriggerMenuCreateEmailРассылкаЗапросов')
    // .addItem('Отправить письма c файлами', 'MrLib_top.menuCreateEmailРассылкаЗапросов')
    .addItem('Отправить письма о закрытие проекта', 'MrLib_top.menuРассылкаЗакрытиеПроекта')
    .addItem('Отправить письма о продолжение проекта', 'MrLib_top.menuРассылкаПродолжениеПроекта')


    .addSeparator()
    // .addItem('Обновить заливку', 'updateBackground')
    .addItem('Обновить заливку', 'MrLib_top.updateBackgroundV2')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Удалить заливку')
      .addItem('Об отправке письма', 'delBackgroundMailBrown')
      .addItem('Об обработке цены', 'delBackgroundMailGreen'))
    /*.addSeparator()
    .addItem('Элемент 4', 'myFunction4')*/

    // Mistervova@mail.ru 24.12.2020 согласно тз от 22.12.2020 //Задача№4 Организовать Автоматическая корректировка формул
    // добавлин пункт в меню 
    .addSeparator()
    // .addItem('Удаление строк', 'correctFormulas')  // Удалить лишние строки
    .addItem('Удалить лишние строки', 'correctFormulas')  // Удалить лишние строки
    .addItem('Добавить столбцы на листе "1-1"', 'MrLib_top.menuAddEmtyColsBeforeАктуальный_Сбор_КП')  // Добавить столбцы на листе "1-1"
    .addItem('Выровнить столбцы на листе "1-1"', 'MrLib_top.menuFixSheetCompact') //

    // Выделение емейлов поставщиков
    // Mistervova@mail.ru 25.12.2020 согласно тз от 22.12.2020 // Задача№7 Выделение емейлов поставщиков цветом
    // добавлин пункт в меню // paint the email of contractors // paintCounteragentEmail
    .addSeparator()
    .addItem('Выделение одинаковых цен на "3-3"', 'paintBackgroundProductList7')
    .addItem('Обновить цвета поставщиков на "1-1" и "3-3"', 'paintBackgroundContragent')
    .addItem('Обновить цветовое выделение на "3-3"', 'paintUpdateBackgroundAllforSheetList7')
    // Найти E-mail в базе товаров  
    .addItem('Найти E-mail в базе товаров', 'findCounteragentEmail')
    .addItem('Создать таблицу закупки', 'menuMakeCopy')
    .addItem('Создать таблицу оплаты', 'MrLib_top.menuMakeCopy_оплаты')
    .addItem('Повторная рассылка', 'MrLib_top.copyReMailingPriductV2')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Триггеры')
      .addItem('Запустить триггеры', 'addNewTrigger')
      .addItem('Список запущеных триггеров', 'MrLib_top.investigatePresenceOfTriggers')
      // .addItem('Запустить триггеры на день', 'addNewTriggerOneDey')
    )

    .addItem('Пересчитать Цена Итого Предложений', 'MrLib_top.recalculate_report_malo')
    .addItem(`Экспорт в Битрикс`, `MrLib_top.menuЭкспортВБитрикс`)
    // .addItem('deleteTrigger', 'deleteTrigger')
    // deleteTrigger
    .addToUi();

  // addNewTrigger("onEdit_trigger");
}

function emailQuota() {
  var ss = SpreadsheetApp.getActive();
  // var shSet = ss.getSheetByName("Настройки писем");
  var shSet = ss.getSheetByName(getSettings().sheetName_Настройки_писем);
  //лимит рассылки
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  shSet.getRange("K1").setValue(emailQuotaRemaining);
}



function addNewTrigger() {
  let rangeMem = "K11";
  let da = fl_str("ДА");
  let hasTriggers = (() => {
    if (fl_str(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).getValue()) != da) { return false; }
    let memNote = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).getNote();

    memNote = (() => { try { return JSON.parse(memNote) } catch (err) { return undefined } })();

    if (!memNote) { return false; }
    if (!memNote.triggerMem) { return false; }
    if (!memNote.triggerMem.ПоследнееВыполнение) { return false; }
    let d = SpreadsheetApp.getActiveSpreadsheet().getId().slice(0, 10);
    if (memNote.triggerMem.d != d) { return false; }

    if ((new Date().getTime() - new Date(memNote.triggerMem.ПоследнееВыполнение).getTime() - 0.25 * DeyMilliseconds) > 0) { return false; }
    return true;
  })();


  // hasTriggers = false;
  if (hasTriggers) { Logger.log("триггеры были добавлены рание"); return; } else { Logger.log("триггеров нет и надо добавить"); }

  let ВладелецТаблици = (() => { return SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() })();
  let ТекущийПользователь = (() => { return Session.getEffectiveUser().getEmail() })();
  if (ВладелецТаблици != ТекущийПользователь) {
    Browser.msgBox(`Триггеры не установлены обратитесь к Владельцу таблицы ${ВладелецТаблици}`);
    return;
  } else { Logger.log("триггеров нет и надо добавить ТекущийПользователь ВладелецТаблици, что надо"); }


  // return;
  addNewTriggeronEditTrigger("onEdit_trigger");
  addNewTriggerMrOnEditHelper("triggerMrOnEditHelper");
  addNewTriggerMrOnEditHelper("triggerMrOnEditHelperExternal");
  addNewTriggerMrOnEditHelper("MrLib_top.triggerMrOnEditHelperExpertise");
  addNewTriggerService("MrLib_top.triggerService");

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange(rangeMem).setValue(da);
  service_setTriggerInfo_();
}

function investigatePresenceOfTriggers() {
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

  mesajeStr = "";
  for (let k in mesaje) {
    mesajeStr = `${mesajeStr} \n ${k}: ${mesaje[k]}.`;
  }

  Browser.msgBox(`${mesajeStr}`);
}



function addNewTriggeronEditTrigger(triggerName = "onEdit_trigger") {


  Logger.log(`Триггер проверяем на наличие триггера ${triggerName}`);

  // if (SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() !== Session.getEffectiveUser().getEmail()) {
  //   Logger.log("Триггер может добавить только владелец")
  //   return;
  // }


  if (typeof triggerName != "string") { return; }
  if (triggerName == "") { return; }

  // let triggers = ScriptApp.getScriptTriggers();

  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => Logger.log(trigger.getHandlerFunction()));
  let triggersName = triggers.map(trigger => trigger.getHandlerFunction());
  if (!triggersName.includes(triggerName)) {
    Logger.log(`Триггер ${triggerName} не найден. Добовляем.`);
    ScriptApp.newTrigger(triggerName).forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
  } else {
    Logger.log(`Триггер ${triggerName} найден.`);


  }
}

function addNewTriggerMrOnEditHelper(triggerName = "triggerMrOnEditHelper") {
  Logger.log(`Триггер проверяем на наличие триггера ${triggerName}`);

  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var owner = ss.getOwner();
  // Logger.log(owner.getEmail());
  // Logger.log(Session.getEffectiveUser().getEmail());

  // if (SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() !== Session.getEffectiveUser().getEmail()) {
  //   Logger.log("Триггер может добавить только владелец");
  //   return;
  // }


  if (typeof triggerName != "string") { return; }
  if (triggerName == "") { return; }

  // let triggers = ScriptApp.getScriptTriggers();
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => Logger.log(trigger.getHandlerFunction()));
  let triggersName = triggers.map(trigger => trigger.getHandlerFunction());

  if (!triggersName.includes(triggerName)) {
    Logger.log(`Триггер ${triggerName} не найден. Добовляем.`);
    // ScriptApp.newTrigger(triggerName).forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
    ScriptApp.newTrigger(triggerName).timeBased().everyMinutes(5).create();

  } else {
    Logger.log(`Триггер ${triggerName} найден.`);
  }


}

function addNewTriggerService(triggerName = "MrLib_top.triggerService") {
  Logger.log(`Триггер проверяем на наличие триггера ${triggerName}`);
  if (typeof triggerName != "string") { return; }
  if (triggerName == "") { return; }
  // let triggers = ScriptApp.getScriptTriggers();
  let triggers = ScriptApp.getProjectTriggers();
  // triggers.forEach(trigger => Logger.log(trigger.getHandlerFunction()));
  let triggersName = triggers.map(trigger => trigger.getHandlerFunction());

  if (!triggersName.includes(triggerName)) {
    Logger.log(`Триггер ${triggerName} не найден. Добовляем.`);
    // ScriptApp.newTrigger(triggerName).forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
    ScriptApp.newTrigger(triggerName).timeBased().everyMinutes(30).create();

  } else {
    Logger.log(`Триггер ${triggerName} найден.`);
  }
}


function deleteTrigger() {
  // return;

  try {
    // let Стоп = MrLib_Midlle.getItemByKey("Проекты", SpreadsheetApp.getActiveSpreadsheet().getUrl(), "Стоп", "juj");
    // Logger.log(`deleteTrigger item=${JSON.stringify(Стоп)}`);
    // MrLib_Midlle.updateItems("Проекты", [
    //   {
    //     key: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    //     // tr_dell: true,
    //     tr_dell: Стоп,
    //     tr_dell_date: new Date(),
    //     tr_dell_email: Session.getEffectiveUser().getEmail(),
    //   }]);
    // if (Стоп != true) { return; }
    let net = fl_str("Нет");
    deleteTriggerByName("onEdit_trigger");
    deleteTriggerByName("triggerMrOnEditHelper");
    deleteTriggerByName("triggerMrOnEditHelperExternal");
    deleteTriggerByName("MrLib_top.triggerMrOnEditHelperExpertise");
    deleteTriggerByName("MrLib_top.triggerService");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSettings().sheetName_Настройки_писем).getRange("K11").setValue(net);
  } catch (err) { }

}


function deleteTriggerByName(triggerName = "triggerПеренистиПисьма") {
  Logger.log(`Триггер проверяем на наличие триггера ${triggerName}`);
  if (typeof triggerName != "string") { return; }
  if (triggerName == "") { return; }
  // let triggers = ScriptApp.getScriptTriggers();
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => Logger.log(trigger.getHandlerFunction()));
  let triggersName = triggers.map(trigger => trigger.getHandlerFunction());
  if (!triggersName.includes(triggerName)) {
    Logger.log(`Триггер ${triggerName} не найден. Не Удаляем`);
  } else {
    // Logger.log(`Триггер ${triggerName} найден.`);
    Logger.log(`Триггер ${triggerName} найден. Удаляем`);
    let index = triggersName.indexOf(triggerName);
    ScriptApp.deleteTrigger(triggers[index]);
  }
}


