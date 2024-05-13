
function nr(A1) {
  A1 = A1.replace(/\d/g, '')
  let i, l, chr,
    sum = 0,
    A = "A".charCodeAt(0),
    radix = "Z".charCodeAt(0) - A + 1;
  for (i = 0, l = A1.length; i < l; i++) {
    chr = A1.charCodeAt(i);
    sum = sum * radix + chr - A + 1
  }
  return sum;
}

/** @returns {String} */
function fl_str(str) {
  if (!str) { return ""; }
  return str.toString().replace(/ +/g, ' ').trim().toUpperCase();
}


function nc(column) {
  column = parseInt("" + column);
  if (isNaN(column)) { throw ('файл mrColumnToNr функция nrCol(): не найдено буквенное обозначение для колонки "' + column + '"'); }

  column = column - 1;
  switch (column) {
    case 0: { return "A"; }
    case 1: { return "B"; }
    case 2: { return "C"; }
    case 3: { return "D"; }
    case 4: { return "E"; }
    case 5: { return "F"; }
    case 6: { return "G"; }
    case 7: { return "H"; }
    case 8: { return "I"; }
    case 9: { return "J"; }
    case 10: { return "K"; }
    case 11: { return "L"; }
    case 12: { return "M"; }
    case 13: { return "N"; }
    case 14: { return "O"; }
    case 15: { return "P"; }
    case 16: { return "Q"; }
    case 17: { return "R"; }
    case 18: { return "S"; }
    case 19: { return "T"; }
    case 20: { return "U"; }
    case 21: { return "V"; }
    case 22: { return "W"; }
    case 23: { return "X"; }
    case 24: { return "Y"; }
    case 25: { return "Z"; }

    default: {
      if (column > 25) { return `${nc(column / 26)}${nc((column % 26) + 1)}`; }
    }
  }

  throw new Error('файл mrColumnToNr функция nc(): не найдено буквенное обозначение для колонки "' + column + '"');

}

function mrErrToString(err) {
  let ret = "ОШИБКА ВЫПОЛНЕНИЯ СКРИПТА \n " + "\nдата время:" + new Date() + "\nname: " + err.name + "\nmessage: " + err.message + "\nstack: " + err.stack;
  Logger.log(ret);
  return ret;
}

function isError(value) {
  if (!value) { return false; }
  return (value.toString()[0] == "#")
}


function standartUrl(url) {
  url = `${url}`.slice(0, 88);
  if (url.length < 88) {
    let url_patern = "https://docs.google.com/spreadsheets/d/ЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯЯ/edit";
    let ok_url = `${url}${url_patern.slice(url.length)}`;
    if (!ok_url.includes("Я")) { url = ok_url; }
  }
  return url;
}


function logMap(value, key, map) {
  try {
    console.log(`m[${key}] = ${value.logString()}`);
  }
  catch (err) {
    console.log(`m[${key}] = ${value}`);

  }
}



function CONTRAGENT(email) {
  if (!email) { throw "#не емайл" }
  if (!isEmail(email)) { throw "#не емайл" }


  return getContext().getCounteragent(fl_str(email)).logArray()

}




// function myFunctionss() {
//   let r = new RegExp("[♥️♠♣️♦️)(]", "gim" );
//     [
//         "aya (♥️)\naya (♥️)",
//         "aya (w)",
//         "aa (♠)",
//         "aza (♦️)",
//         "awa (♣️)",
//         "awza (♥️) (♥️♠♦️)",

//     ].map(e=>{
//       return  e
//       // .replace("(w)","")
//       .replace(r,"")
//       .trim();
//     }).forEach(e=>Logger.log(e+"|"));

// }

function myFunctionss() {
  let item = {
    a: 1,
    b: 3,
    c: "#ОШИБКА!",
    d: "#ERROR!",
  }
  for (let k in item) {
    if (item[k] == "#ОШИБКА!" || item[k] == "#ERROR!") {
        item[k] = undefined;
    }
  }
Logger.log(JSON.stringify(item));


}
