// MisterVova@mail.ru 25.12.2020 добавил файл  согласно тз от 22.12.2020 // Задача№7 Выделение емейлов поставщиков цветом

// MisterVova@mail.ru 25.12.2020 добавил функцию   согласно тз от 22.12.2020 // Задача№7 Выделение емейлов поставщиков цветом
/*
Исходные данные 
  Лист Отзывы о поставщиках
  Настройка писем
  Лист 1
  Лист 7

Задача
  На листе 1 (первая строка) и на листе 7 (I:AR1000) проверяем все емейлы на листе “Отзывы о поставщиках” столбец B, 
  если есть такой емейл в списке то смотрим его статус (столбец С) далее переходим на лист “Настройки писем” столбцы V:W. 
  смотрим каким цветом залит соответствующий статус. 
  Таким же цветом выделяем и емейл Дополнительно проверяем каждый емейл на листе “Товары”, 
  если в столбце С указано “да” то емейл выделяем цветом согласно заливке ячейки К5 вкладки Настройка писем
  Сделать кнопку в меню для запуска этого скрипта
*/



function paintBackgroundProductList7() {

  let cList7 = getContext().getSheetList7();
  cList7.updateBackgroundProduct();

}

function paintBackgroundContragent() {
  let cList1 = getContext().getSheetList1();
  cList1.updateBackgroundAll();

  let cList7 = getContext().getSheetList7();
  cList7.updateBackgroundContragent();


}
function paintUpdateBackgroundAllforSheetList7() {
  let cList7 = getContext().getSheetList7();
  cList7.updateBackgroundAll();


}




function paintCounteragentAll() {

  getContext().upDate();

  paintCounteragentEmail_List_1();  // Лист1   // getSettings().sheetName_Лист_1
  paintCounteragentAll_List_7();  // Лист7  топ7 колонки Q-AR

}

function paintCounteragentEmail_Test() {
  let ss = SpreadsheetApp.getActive();
  let sheetList_t = ss.getSheetByName("Черновик");
  // let paintRange = sheetList_t.getRange(60,nr("N"),sheetList_t.getLastRow()-60, 1);
  let paintRange = sheetList_t.getRange(60, nr("B"), sheetList_t.getLastRow() - 60 + 1, nr("K") - nr("B") + 1);
  paintRange.setBackgrounds(paintCounteragentEmail_List_BG(paintRange));

}




function paintCounteragentEmail_List_1() {
  let cList1 = getContext().getSheetList1();
  cList1.updateBackgroundAll();


}



function paintCounteragentAll_List_7() {
  let cList7 = getContext().getSheetList7();
  cList7.updateBackgroundAll();
}




function paintCounteragentEmail_List_BG(paintRange) {

  let bg = paintRange.getBackgrounds();
  let vl = paintRange.getValues();

  for (let row = 0; row < bg.length; row++) {
    for (let col = 0; col < bg[row].length; col++) {
      let email = fl_str(vl[row][col]);

      if (!isEmail(email)) { continue }
      // Logger.log("row=" + row + " col=" + col);      
      let color = getContext().getCounteragentColor(fl_str(email));
      if (!color) continue;
      //  getContext().getCounteragent(fl_str(email)).logToConsole(); 
      // Logger.log("row=" + row + " col=" + col  + " email=" + email + " bg[row][col] old=" + bg[row][col]);
      bg[row][col] = color;
      // Logger.log("row=" + row + " col=" + col  + " email=" + email + " bg[row][col] new=" + bg[row][col]);              
      // Logger.log("_____________________" );
    }
  }

  return bg;
}







// mrFindCounteragentEmail















