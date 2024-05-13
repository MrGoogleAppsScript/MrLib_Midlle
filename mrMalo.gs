function getBackground(cellAdres) {
  var range = SpreadsheetApp.getActiveSheet().getRange(cellAdres),
    color = range.getBackground();
  return color;
}

function ЦВЕТФОНА(cellAdres) {
  return getBackground(cellAdres);
}


function countCellsWithBackgroundColor(maloB8 = 3, colorAnalog = "#ffff00", rHead, rBg, rPrice) {
  // let sheetName = "Лист1";
  let sheetName = getSettings().sheetName_Лист_1;
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);

  var vlPrice = sheet.getRange(rPrice).getValues();
  var vlBG = sheet.getRange(rBg).getBackgrounds();
  var vlHead = sheet.getRange(rHead).getValues();

  var priceAll = 0;
  var priceAnalog = 0;

  for (var i = 0; i < vlPrice.length; i++) {
    for (var j = 0; j < vlPrice[i].length; j++) {

      if (!vlHead[i][j]) { continue; }  // в шапке пусто  пропускаем
      if (!vlPrice[i][j]) { continue; } // цена пусто пропускаем // нет значения 
      priceAll++;
      if (vlBG[i][j] == colorAnalog) {
        priceAnalog++;
      }
    }
  }

  var retStr = `${priceAll}`;
  if (priceAll <= maloB8) { retStr = `Мало, всего ${retStr}`; }
  if (priceAnalog != 0) { retStr = `${retStr}, из них аналог ${priceAnalog}` }

  return retStr;
}




function malo(row, maloB8) {
  let colorAnalog = getContext().getAnalogColor();
  // let colFirst = "S";
  // let colLast = "HP";
  let colFirst = nc(getClassColRow().list1_col_FirstPrice);
  let colLast = nc(getClassColRow().list1_col_LastPrice);

  let rHead = `$${colFirst}$${1}:$${colLast}$${1}`;
  let rBg = `$${colFirst}$${1}:$${colLast}$${1}`;
  let rPrice = `$${colFirst}$${row}:$${colLast}$${row}`;

  return countCellsWithBackgroundColor(maloB8, colorAnalog, rHead, rBg, rPrice);
}

function мало(row, maloB8) {
  return malo(row, maloB8);
}

// function мало(row, maloB8){
//   return MrLib_top.malo(row, maloB8)
//  }



