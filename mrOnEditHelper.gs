function triggerMrOnEditHelper(duration = 1 / 24 / 60 * 4.9) {


  // try {
  //   [
  //     "1QvKYo_0ORInyMx0kIl4Oakf4Y-IMZVCS39_bV6AGvEk",
  //   ].forEach(id => {
  //     DriveApp.getFileById(id).addEditor("script@ss-postavka.ru");
  //   });

  // } catch (err) {

  // }
  new MrOnEditHelper().triggerMrOnEditHelper(duration);
}




function triggerMrOnEditHelperExpertise(trirrgetInfo, duration = 1 / 24 / 60 * 4.9) {
  try { serviceCopyMemFromBuyExternalToProduct(); } catch (err) { mrErrToString(err); }
  new MrOnEditHelper().triggerMrOnEditHelperExpertise(duration);
}



function triggerMrOnEditHelperExternal(duration = 1 / 24 / 60 * 4.9) {
  // try { servicesСинхрШаблонаОплатыВнешней();    } catch (err) { }
  // try {  triggerService(); } catch (err) { }
  // try { service_setTriggerInfo_(); } catch (err) { mrErrToString(err); }

  //  renameSheet
  try {
    getContext().getSheetPay().triggerOnEditHelperPayExternal(duration);
  } catch (err) { mrErrToString(err); }



  try {
    getContext().getSheetBuy().triggerOnEditHelperBuyExternal(duration);
  } catch (err) { mrErrToString(err); }




}



class MrOnEditHelper {
  constructor() {


  }
  triggerMrOnEditHelper(duration = 1 / 24 / 60 * 4.9) {
    // return
    Logger.log(`MrOnEditHelper  triggerMrOnEditHelper  duration=${duration}`);
    // deleteTrigger();
    triggerExportValuesToMidlle();

    let tt = 0;
    while (getContext().hasTime(duration)) {
      // while (this.hasTask() && this.hasTime(duration)) {
      Logger.log(`MrOnEditHelper   triggerMrOnEditHelper Цыкл=${tt} ${new Date()} `);
      tt++;
      try {

        if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelper | мало времени break`); break; }
        Logger.log(` getContext().getSheetList1().fixSheet(duration);=${tt} ${new Date()} `);
        try { getContext().getSheetList1().fixSheet(duration); } catch (err) { mrErrToString(err); }

        if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelper | мало времени break`); break; }
        Logger.log(` getContext().getSheetList7().fixSheet(duration);=${tt} ${new Date()} `);
        try { getContext().getSheetList7().fixSheet(duration); } catch (err) { mrErrToString(err); }


        if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelper | мало времени break`); break; }
        Logger.log(` getContext().getSheetList7().onEditHelper(duration);=${tt} ${new Date()} `);
        try { getContext().getSheetList7().onEditHelper(duration); } catch (err) { mrErrToString(err); }


        // if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelper | мало времени break`); break; }
        // Logger.log(` getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).fixSheet(duration));=${tt} ${new Date()} `);
        // try { getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).fixSheet(duration)); } catch (err) { mrErrToString(err); }


        // if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelper | мало времени break`); break; }
        // Logger.log(` getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).onEditHelper(duration))=${tt} ${new Date()} `);
        // try { getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).onEditHelper(duration)); } catch (err) { mrErrToString(err); }


        if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelper | мало времени break`); break; }
        Logger.log(` getContext().getSheetBuy().onEditHelper(duration);=${tt} ${new Date()} `);
        try { getContext().getSheetBuy().onEditHelper(duration); } catch (err) { mrErrToString(err); }


        SpreadsheetApp.flush();
        mrContext = undefined;
        classColRow = undefined;
        // break
      } catch (err) {
        Logger.log(mrErrToString(err));
        // getContext().addError(err);

      } finally {

      }

      Utilities.sleep(getContext().tMin * 6);
      // break;
    }

    Logger.log(`MrOnEditHelper  triggerMrOnEditHelper  Finish`);
    return true;
  }



  triggerMrOnEditHelperExpertise(duration = 1 / 24 / 60 * 3.9) {
    // return
    // serviceСинзронизацияВопросов()
    Logger.log(`MrOnEditHelper  triggerMrOnEditHelperExpertise  duration=${duration}`);

    let tt = 0;
    while (getContext().hasTime(duration)) {
      // while (this.hasTask() && this.hasTime(duration)) {
      Logger.log(`MrOnEditHelper   triggerMrOnEditHelperExpertise Цыкл=${tt} ${new Date()} `);
      tt++;
      try {

        if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelperExpertise | мало времени break`); break; }
        Logger.log(` getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).onEditHelper(duration))=${tt} ${new Date()} `);
        try {
          getContext().getSheetNamesExpertise().forEach(sheetName =>
            getContext().getSheetExpertise(sheetName).onEditHelper(duration));
        } catch (err) { mrErrToString(err); }


        if (!getContext().hasTime(duration)) { Logger.log(`MrOnEditHelper triggerMrOnEditHelperExpertise | мало времени break`); break; }
        Logger.log(` getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).fixSheet(duration));=${tt} ${new Date()} `);
        try { getContext().getSheetNamesExpertise().forEach(sheetName => getContext().getSheetExpertise(sheetName).fixSheet(duration)); } catch (err) { mrErrToString(err); }




        // if (tt < 2) { try { serviceСинзронизацияВопросов(); } catch (err) { mrErrToString(err); } }

        SpreadsheetApp.flush();
        // mrContext = undefined;
        // break
      } catch (err) {
        Logger.log(mrErrToString(err));
        // getContext().addError(err);

      } finally {

      }

      Utilities.sleep(getContext().tMin * 6);
      // break;
    }

    Logger.log(`MrOnEditHelper  triggerMrOnEditHelper  Finish`);
    return true;
  }

}
