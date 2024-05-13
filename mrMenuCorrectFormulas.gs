// MisterVova@mail.ru 24.12.2020 добавил файл  согласно тз от 22.12.2020 // Задача№4 Организовать Автоматическая корректировка формул

// MisterVova@mail.ru 24.12.2020 добавил функцию   согласно тз от 22.12.2020 // Задача№4 Организовать Автоматическая корректировка формул
// вызывается из меню 
function correctFormulas() {
  let ss = SpreadsheetApp.getActive();

  // let sheetSetings = ss.getSheetByName("Настройки писем");
  let sheetSetings = ss.getSheetByName(getSettings().sheetName_Настройки_писем);
  let ui = SpreadsheetApp.getUi();

  let setings = sheetSetings.getRange(2, nr("X"), sheetSetings.getLastRow(), 4).getValues();
  let sheetNamesDelRow = new Array();
  let columnIndexY = new Array();
  let columnIndexZ = new Array();
  let columnIndexAA = new Array();
  let sheetNamesDelREF = new Array();

  // console.log("setings =" + setings );
  for (let i = setings.length - 1; i >= 0; i--) {
    // console.log( "i= "+ i+ " лист= " + setings [i][0] + " колонка =" + setings [i][1] );
    if (!setings[i][0]) {
      // setings.splice(i, 1);
      continue;
    }

    if (!setings[i][1]) {
      // setings.splice(i, 1);
      continue;
    }
    if (!setings[i][2]) {
      // setings.splice(i, 1);
      continue;
    }


    if (!setings[i][3]) {
      // setings.splice(i, 1);
      continue;
    }

    if (!parseInt(setings[i][3].toString(), 10)) {
      // setings.splice(i, 1);
      continue;
    }


    if (!ss.getSheetByName(setings[i][0].toString())) {
      ui.alert(
        'Лист "' + setings[i][0] + '"не найден ',
        // 'Нет @ на листе "'+ sheetNames[i] +'" проверьте колонку Nr=' + columnIndex[i], 
        `проверте лист "${getSettings().sheetName_Настройки_писем}" колонка "X" строка ` + (i + 2),

        ui.ButtonSet.OK
      );

      continue;

    }



    if (!checkSobachka(ss.getSheetByName(setings[i][0].toString()), nr(setings[i][2]))) {
      ui.alert(
        'Лист "' + setings[i][0] + '"будет пропущен',
        // 'Нет @ на листе "'+ sheetNames[i] +'" проверьте колонку Nr=' + columnIndex[i], 
        'В колонке ' + setings[i][2] + ' нет "@"',

        ui.ButtonSet.OK
      );
      continue;
    }

    if (!sheetNamesDelREF.includes(setings[i][0])) {
      sheetNamesDelREF.push(setings[i][0]);
    }

    sheetNamesDelRow.push(setings[i][0]);
    columnIndexY.push(nr(setings[i][1]));
    columnIndexZ.push(nr(setings[i][2]));

    let aa = parseInt(setings[i][3].toString(), 10);
    aa = (aa < 1 ? 1 : aa);
    columnIndexAA.push(aa);
  }
  Logger.log(JSON.stringify(
    sheetNamesDelRow,
    columnIndexY,
    columnIndexZ,
    columnIndexAA,
    sheetNamesDelREF,
  ), null, 2);

  // очистка строк 
  // for (let i = 0; i < sheetNamesDelRow.length; i++) {
  //   clearEmptiRowInSheet(ss.getSheetByName(sheetNamesDelRow[i]), columnIndexY[i], columnIndexZ[i], columnIndexAA[i]);

  // }

  // return;
  // удаление строк 
  for (let i = 0; i < sheetNamesDelRow.length; i++) {
    Logger.log(JSON.stringify(["Начало deleteEmptiRowInSheet", sheetNamesDelRow[i], columnIndexY[i], columnIndexZ[i], columnIndexAA[i]]));
    deleteEmptiRowInSheet(ss.getSheetByName(sheetNamesDelRow[i]), columnIndexY[i], columnIndexZ[i], columnIndexAA[i]);
    Logger.log(JSON.stringify(["Конец deleteEmptiRowInSheet", sheetNamesDelRow[i], columnIndexY[i], columnIndexZ[i], columnIndexAA[i]]));
  }

}



function checkSobachka(sheet, column) {

  let vl = sheet.getRange(1, column, sheet.getLastRow()).getValues();
  for (let row = vl.length - 1; row >= 0; row--) {
    if (vl[row].includes("@")) { return true; }
  }
  return false;
}

function clearEmptiRowInSheet(sheet, columnY, columnZ, columnAA = 1) {

  if (!checkSobachka(sheet, columnZ)) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(
      'Лист "' + sheet.getSheetName() + '"будет пропущен',
      'В колонке Nr=' + columnZ + ' нет "@"',
      ui.ButtonSet.OK
    );
    return;
  }
  //очистка строк пустые перед @ //начало  

  let rowSobachka = 0; // @ в строке с номером rowSobachka 
  let vcs = sheet.getRange(1, columnZ, sheet.getLastRow()).getValues();
  for (; rowSobachka < vcs.length; rowSobachka++) {
    if (vcs[rowSobachka].includes("@")) { break; }
  }
  rowSobachka++; // @ в строке с номером rowSobachka 
  // sheet.getRange(rowSobachka, columnY).setBackground("#00ff00");

  let rowEnd = rowSobachka;
  let rowStart = rowSobachka - 1;

  let sv = sheet.getRange(1, columnY, rowSobachka, 1).getValues();

  for (; rowStart > columnAA; rowStart--) {
    if (sv[rowStart - 1][0]) { break; } // нашли заполнееную строчку выше @  
  }
  rowStart++;
  // if ((rowEnd - rowStart) > 0) { sheet.deleteRows(rowStart, rowEnd - rowStart); }
  if ((rowEnd - rowStart) > 0) { sheet.getRange(rowStart, 1, rowEnd - rowStart, sheet.getLastColumn()).clearContent(); }
  if ((rowEnd - rowStart) > 0) { sheet.getRange(rowStart, 1, rowEnd - rowStart, sheet.getLastColumn()).setBackground("#999999"); }
  // очистка строк пустые перед @ //конец
}




function deleteEmptiRowInSheet(sheet, columnY, columnZ, columnAA = 1) {
  if (!checkSobachka(sheet, columnZ)) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(
      'Лист "' + sheet.getSheetName() + '"будет пропущен',
      'В колонке Nr=' + columnZ + ' нет "@"',
      ui.ButtonSet.OK
    );
    return;
  }


  // удаляет пустые перед @ //начало  
  let rowSobachka = 0; // @ в строке с номером rowSobachka 
  let vcs = sheet.getRange(1, columnZ, sheet.getLastRow()).getValues();
  for (; rowSobachka < vcs.length; rowSobachka++) {
    if (vcs[rowSobachka].includes("@")) { break; }
  }
  rowSobachka++; // @ в строке с номером rowSobachka 
  // sheet.getRange(rowSobachka, columnY).setBackground("#00ff00");

  let rowEnd = rowSobachka;
  let rowStart = rowSobachka - 1;

  let sv = sheet.getRange(1, columnY, rowSobachka, 1).getValues();

  for (; rowStart > columnAA; rowStart--) {
    if (sv[rowStart - 1][0]) { break; } // нашли заполнееную строчку выше @  
  }
  rowStart++;
  if ((rowEnd - rowStart) > 0) { sheet.deleteRows(rowStart, rowEnd - rowStart); }
  // удаляет пустые перед @ //конец
}

