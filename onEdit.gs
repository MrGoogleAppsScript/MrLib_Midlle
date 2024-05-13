function onEdit(e) {

  try {
    var range = e.range;
    var row = range.getRow();
    var column = range.getColumn();
    var sheet = range.getSheet();
    var sheetName = range.getSheet().getSheetName();

    // Mistervova@mail.ru 23.12.2020 согласно тз от 22.12.2020 //Задача№1 Организовать поиск емейлов поставщиков по справочнику
    Logger.log("onEdit edit= " + JSON.stringify(e));
    Logger.log("onEdit sheetName=" + sheetName);


    // if (sheetName == 'Группы товаров') {
    if (sheetName == getSettings().sheetName_Группы_товаров) {
      onEditProductGroups(e);
    }
    // if (sheetName == 'Лист7') {
    if (sheetName == getSettings().sheetName_Лист_7) {
      onEditList_7(e);
    }

    // if (sheetName == 'Лист1') {
    if (sheetName == getSettings().sheetName_Лист_1) {
      onEditList_1(e);
    }

    // if (sheetName == 'Товары') {
    if (sheetName == getSettings().sheetName_Товары) {
      onEditList_Tovari(e);
    }

    // if (sheetName == 'Вопросы') {
    if (sheetName == getSettings().sheetName_Вопросы) {
      onEditList_Вопросы(e);
    }

    // if (sheetName == 'Вопросы') {
    if (sheetName == getSettings().sheetName_Проблемы) {
      onEditList_Проблемы(e);
    }




    if (getContext().isSheetExpertise(sheetName)) {
      onEditList_Expertise(e, sheetName);
    }


    // if (sheetName == "Группы товаров") {
    if (sheetName == getSettings().sheetName_Группы_товаров) {

      if (column > 8 && row > 3) {
        if (sheet.getRange(row, 1).getValue() != "")
          //SpreadsheetApp.getUi().alert(row + ";" + column + " " + sheetName); 
          unicEmailRow(row);
      }
      // if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
      //   unicEmailRow(row);
      // }

      try { serviceПодсчетНеразосланныхПисем_(); } catch (err) { mrErrToString(err); }
    }

    // if (sheetName == 'Лист1') {
    if (sheetName == getSettings().sheetName_Лист_1) {
      ProdBackground(e);
    }

  }
  catch (err) {

    SpreadsheetApp.getUi().alert(mrErrToString(err));
  }
  // try { onOpen() } catch (err) { mrErrToString(err); }

}
























function onEditList_Tovari(edit) {

  // Получаем диапазон ячеек, в которых произошли изменения
  let range = edit.range;

  // Лист, на котором производились изменения
  let sheet = range.getSheet();

  // Проверяем, нужный ли это нам лист
  Logger.log("onEditList_Tovari");
  // if (sheet.getName() != 'Товары') {
  if (sheet.getName() != getSettings().sheetName_Товары) {
    return false;
  }


  if (range.columnStart <= nr("C")) {
    getContext().getSheetProductGroups().updateAllRows();
  }


}


































// function unicEmailRow(unicRowIn) {
//   //  Logger.log("unicEmailRow unicRowIn= " + unicRowIn);

//   let firstRow = getContext().getProductGroup(unicRowIn).getFirstRow(); // может предыдущяя строка первая строка для этой группы 
//   if (firstRow != unicRowIn) { return; }


//   var unicRow = unicRowIn - 1;
//   var ss = SpreadsheetApp.getActive();
//   // var sh = ss.getSheetByName("Группы товаров");
//   var sh = ss.getSheetByName(getSettings().sheetName_Группы_товаров);
//   var lRow = sh.getLastRow();
//   var lCol = sh.getLastColumn();


//   // if (lCol<nr("J")) {lCol=nr("J");}
//   var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
//   Logger.log(`"unicEmailRow" ${unicRowIn} arrSh=${arrSh}`);


//   var sec = 10;//столбик откуда начинают писать email

//   // let sheetProductGroups = getContext().getSheetProductGroups(); //



//   if (arrSh[unicRow][0] != "" && arrSh[unicRow - 1][0] == "") {
//     var arrEmail = [];
//     var arrBG = [];
//     arrBG.push([]);
//     arrEmail.push([]);
//     for (let j = sec - 1; j < arrSh[unicRow].length; j++) {
//       if (arrSh[unicRow][j].toString().indexOf("@") > 0) {
//         if (arrEmail[0].indexOf(arrSh[unicRow][j]) == -1) {
//           arrEmail[0].push(arrSh[unicRow][j]);

//           // let email = fl_str(arrSh[unicRow][j]);
//           // let color = sheetProductGroups.getColorContragentForUnucRow(unicRowIn, email);

//           // // let color =   contragent.getColor();     
//           // arrBG[0].push(color);
//         }
//       }
//     }
//     sh.getRange(unicRow + 2, sec, 1, lCol).protect().remove();
//     sh.getRange(unicRow + 2, sec, 1, lCol).clearContent();   // добовляет лишнии пустые колонки 
//     // sh.getRange(unicRow + 2, sec, 1, lCol - sec + 1).clearContent();   // Не  добовляет лишнии пустые колонки 
//     if (arrEmail[0].length > 0) {
//       var r = sh.getRange(unicRow + 2, sec, arrEmail.length, arrEmail[0].length)

//       r.setValues(arrEmail);
//       // r.setBackgrounds(arrBG);
//       // Logger.log("unicEmailRow arrEmail= " + arrEmail);
//       sh.getRange(unicRow + 2, sec, 1, lCol).protect().setWarningOnly(true);
//     }
//   }
// }


// function unicEmail() {
//   Logger.log("начало");

//   ramoveAllProtection();

//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName("Группы товаров");
//   var lRow = sh.getLastRow();
//   var lCol = sh.getLastColumn();
//   var arrSh = sh.getRange(1, 1, lRow, lCol).getValues();
//   var flag = false;
//   var sec = 10;//столбик откуда начинают писать email
//   for (let i = 3; i < arrSh.length; i++) {
//     if (arrSh[i][0] != "") {
//       if (flag == false) {
//         flag = true;
//         var arrEmail = [];
//         arrEmail.push([]);
//         for (let j = sec - 1; j < arrSh[i].length; j++) {
//           if (arrSh[i][j].toString().indexOf("@") > 0) {
//             if (arrEmail[0].indexOf(arrSh[i][j]) == -1) {
//               arrEmail[0].push(arrSh[i][j]);
//             }
//           }
//         }
//         Logger.log(arrEmail);

//         sh.getRange(i + 2, sec, 1, lCol).clearContent();
//         if (arrEmail[0].length > 0) {
//           var r = sh.getRange(i + 2, sec, arrEmail.length, arrEmail[0].length)
//           r.setValues(arrEmail);
//           r.protect().setWarningOnly(true);
//           //var protection =r.protect();
//           //var me = Session.getEffectiveUser();
//           //protection.addEditor(me);
//           //.setWarningOnly(true); ;//.setDescription('Отфильтованные email запрещено редактировать.');
//           //r.protect().setWarningOnly(true);            
//         }
//       }//if (flag == false)
//     } else {
//       flag = false;
//     }
//   }
//   Logger.log("конец");
// }

// function ramoveAllProtection() {
//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName("Группы товаров");
//   var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
//   for (var i = 0; i < protections.length; i++) {
//     var protection = protections[i];
//     if (protection.canEdit()) {
//       protection.remove();
//     }
//   }
// }



