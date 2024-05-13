
let sheetNameFrom = getSettings().sheetName_Лист_1;
// ${sheetNameFrom}



function lookFormula(sheetName, row, col) {
  let sheet = getContext().getSheetByName(sheetName);
  return sheet.getRange(row, col).getFormula();
}

function lookFormulaR1C1(sheetName, row, col) {
  let sheet = getContext().getSheetByName(sheetName);
  let ret = sheet.getRange(row, col).getFormulaR1C1();
  Logger.log(`${ret}`);
  return ret;
}

function makeFormulaR1C1(sheetName, row, col) {
  // if (fl_str(sheetName) == fl_str("Лист7")) { return makeFormula_List7(row, col); }
  if (fl_str(sheetName) == fl_str(getSettings().sheetName_Лист_7)) { return makeFormula_List7(row, col); }
  // if (fl_str(sheetName) == fl_str("${sheetNameFrom}")) { return makeFormula_List1(row, col); }
  if (fl_str(sheetName) == fl_str(getSettings().sheetName_Лист_1)) { return makeFormula_List1(row, col); }
}





function makeFormula_List7(row, col) {

  if (col == nr("AZ")) { return `=ROW(INDIRECT(MIDB(FORMULATEXT(INDIRECT(ADDRESS(ROW();2;1;TRUE)));2;LENB(FORMULATEXT(INDIRECT(ADDRESS(ROW();2;1;TRUE)))))))`; }
  if (col == nr("AY")) { return `=INDIRECT(ADDRESS(R[0]C52;COLUMN(R1C233);1;TRUE;"${sheetNameFrom}"))`; }
  if (col == nr("AX")) { return `=INDIRECT(ADDRESS(R[0]C52;COLUMN(R1C4);1;TRUE;"${sheetNameFrom}"))`; }

}



function makeFormula_List1(row, col) {


  if (col == getClassColRow().list1_col_TopFormula)
  // { return f_List1_HY(); } // 03.09.2021  MrVova
  { return f_List1_ЛучшиеЦены(); } // 19.06.2023  MrVova


  if (col == getClassColRow().list1_col_КомментарийЭксперта)
  // { return f_List1_HY(); } // 03.09.2021  MrVova
  { return f_List1_КомментарийЭксперта(); } // 19.06.2023  MrVova



  if (col == getClassColRow().list1_col_ТопЭксперта)
  // { return f_List1_HY(); } // 03.09.2021  MrVova
  { return f_List1_ТопЭксперта(); } // 19.06.2023  MrVova

}


function f_List1_ТопЭксперта() {


  let cf = `C${getClassColRow().list1_col_FirstPrice}`;
  let cl = `C${getClassColRow().list1_col_LastPrice}`;

  let cc = `C${getClassColRow().list1_col_Coment}`;
  let ce = `C${getClassColRow().list1_col_КомментарийЭксперта}`;
  let sheetName = getSettings().sheetName_Лист_1

  // "Цена" \\ 
  // "Поставщик" \\ 
  // "Срок" \\ 
  // "Модель" \\ 
  // "Изготовитель" \\ 
  // "КомментарийЭксперта" \\ 
  // "Подходит" \\ 
  // "КомментарийМенеджера" \\ 


  let ret = `
LAMBDA(cc;тело;комент;комЭкс;BYROW(тело;LAMBDA(r;
TEXTJOIN("|||";FALSE; ARRAY_CONSTRAIN( Query( TRANSPOSE( 
{
{ ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))        };
{ ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))        };
{ ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))        };
{ ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r)+1 ); MOD(COLUMN(r)-cc;3)=1))        };
{ ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);комЭкс;2;0 );""))    };
{ ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);комЭкс;3;0 );""))    };
{ ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);комЭкс;4;0 );""))    };
{ ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
  { TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)})))) \\ TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
;2;0 );"")) };

{  ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))  }
}
);"Select * Where Col9 is not null Order by Col9";0);14;8)))
))
(
  
 COLUMN('${sheetName}'!R[0]${cf});
 '${sheetName}'!R[0]${cf}:R[0]${cl};
 '${sheetName}'!R[0]${cc};
 BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)))

)
`;


  // COLUMN('${sheetName}'!R[0]${cf});
  // '${sheetName}'!R[0]${cf}:R[0]${cl};
  // '${sheetName}'!R[0]${cc};
  // BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"|";FALSE;FALSE)))


  return ret;

}


function f_List1_КомментарийЭксперта() {


  let cf = `C${getClassColRow().list1_col_CommentExpertBlokStart}`;
  let cl = `C${getClassColRow().list1_col_CommentExpertBlokEnd}`;

  let sheetName = getSettings().sheetName_Лист_1

  let ret = `
    IFERROR(  
      TEXTJOIN("|||";true ; 
        BYROW(
        ARRAY_CONSTRAIN( SORT( ARRAYFORMULA(
          Split( TRANSPOSE (
            Split(TEXTJOIN("|||"; TRUE; '${sheetName}'!R[0]${cf}:R[0]${cl}) ;"|||";FALSE;FALSE)
          );"📎";FALSE;FALSE));5;FALSE );100;4)
        ;LAMBDA(r;TEXTJOIN("📎";FALSE ;r))
      )
    ;))`

  return ret;
}



function makeFormulaR1C1_V3D(r, c) {
  return `={IFERROR(   INDIRECT( CONCATENATE("'${sheetNameFrom}'!";   INDEX(SPLIT(R[0]C51;"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C17)+1)));"")\\IFERROR(INDEX(SPLIT(R[0]C51;"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C17)+2);"")}`;
}




function makeFormulaR1C1_V3(r, c) {
  return `=IFERROR( INDIRECT(  CONCATENATE("'${sheetNameFrom}'!";  INDEX(SPLIT(R[0]C51;"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C17)+1)));"")`;

}






function makeFormulaR1C1D(r, c) {
  return `=  {IFERROR(INDEX(SPLIT(R[0]C51;"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C17)+1);"")\\IFERROR(INDEX(SPLIT(R[0]C51;"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C17)+2);"")}`;
}

function makeFormulaR1C1S(r, c) {
  return `=IFERROR(INDEX(SPLIT(R[0]C51;"|||";FALSE;FALSE);1;COLUMN(R1C[0])-COLUMN(R1C17)+1);"")`;
}




function _getScriptId() {
  let ret = ScriptApp.getScriptId();
  return JSON.stringify(ret);
}

function _getUrl() {
  let ret = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return JSON.stringify(ret);
}



function f_List1_HY() {

  let cf = `C${getClassColRow().list1_col_FirstPrice}`;
  let cl = `C${getClassColRow().list1_col_LastPrice}`;

  let cc = `C${getClassColRow().list1_col_Coment}`;
  let cff = `C${getClassColRow().list1_col_FirstPrice + 3}`;

  return `=TEXTJOIN("|||";FALSE;

ARRAY_CONSTRAIN(FILTER(

SORTN(TRANSPOSE({

ARRAYFORMULA(FILTER(ARRAYFORMULA( ADDRESS(ROW('${sheetNameFrom}'!R[0]${cf}:R[0]${cl});COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl});1;TRUE)); IF( COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})>(COLUMN('${sheetNameFrom}'!R1${cff})-1) ; ISODD(COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})) ; TRUE ) ))
;

ARRAYFORMULA(FILTER(ARRAYFORMULA( ADDRESS(ROW('${sheetNameFrom}'!R1${cf}:R1${cl});COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl});1;TRUE)); IF( COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})>(COLUMN('${sheetNameFrom}'!R1${cff})-1) ; ISODD(COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})) ; TRUE ) ))
;

ARRAYFORMULA(IFERROR(VLOOKUP( FILTER(ARRAYFORMULA( ADDRESS(ROW('${sheetNameFrom}'!R1${cf}:R1${cl});COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl});1;TRUE)); IF( COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})>(COLUMN('${sheetNameFrom}'!R1${cff})-1) ; ISODD(COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})) ; TRUE ) );
{ TRANSPOSE (FILTER( {SPLIT('${sheetNameFrom}'!R[0]${cc};"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT('${sheetNameFrom}'!R[0]${cc};"|||";FALSE;FALSE) ); 1; 1)})))) \\TRANSPOSE (FILTER( {SPLIT('${sheetNameFrom}'!R[0]${cc};"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT('${sheetNameFrom}'!R[0]${cc};"|||";FALSE;FALSE) ); 1; 1)}))))};2;FALSE);""))
;

ARRAYFORMULA(FILTER(TRANSPOSE({"";"";"";TRANSPOSE (ARRAYFORMULA( ADDRESS(ROW('${sheetNameFrom}'!R[0]${cff}:R[0]${cl});COLUMN('${sheetNameFrom}'!R1${cff}:R1${cl});1;TRUE)))}); IF( COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})>(COLUMN('${sheetNameFrom}'!R1${cff})-1) ; ISEVEN(COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})) ; TRUE ) ))
;
FILTER( '${sheetNameFrom}'!R[0]${cf}:R[0]${cl};   IF( COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})>(COLUMN('${sheetNameFrom}'!R1${cff})-1) ; ISODD(COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})) ;  TRUE ) ) 


});10;0;5;TRUE)

;

SORTN(TRANSPOSE({FILTER( '${sheetNameFrom}'!R[0]${cf}:R[0]${cl};   IF( COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})>(COLUMN('${sheetNameFrom}'!R1${cff})-1) ; ISODD(COLUMN('${sheetNameFrom}'!R1${cf}:R1${cl})) ;  TRUE ) ) });10;0;1;TRUE)

)


;7;4)

)

`;

}






function f_List1_ЛучшиеЦены() {
  // if (isTest()) { return f_List1_ЛучшиеЦеныОктябрь2023_2(); }

  return f_List1_ЛучшиеЦеныОктябрь2023_2();

  let cf = `C${getClassColRow().list1_col_FirstPrice}`;
  let cl = `C${getClassColRow().list1_col_LastPrice}`;

  let cc = `C${getClassColRow().list1_col_Coment}`;
  let ce = `C${getClassColRow().list1_col_КомментарийЭксперта}`;
  let sheetName = getSettings().sheetName_Лист_1


  // return `LAMBDA(cc;тело;комент;комент)
  // (
  // COLUMN('1-1 Сбор КП'!R[0]C27);
  // '1-1 Сбор КП'!R[0]C27:R[0]C236;
  // '1-1 Сбор КП'!R[0]C244
  // )`



  return `LAMBDA(cc;тело;комент; комЭкс;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))  \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( Lambda(f;m;"Подходит: "& IFERROR( VLOOKUP( f; комЭкс; 4;0 );"-?")&", "&  m&" "&IFERROR(VLOOKUP( f; комЭкс; 2;0 );)&" "&IFERROR(VLOOKUP( f; комЭкс; 3;0 );)  )
(
 FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
 FILTER( r; MOD(COLUMN(r)-cc;3)=2)
) )   };
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))}
});"Select * Where Col6 is not null Order by Col6";0);7;5))
)
)
)
(


COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc};
BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)))

)`




  return `LAMBDA(cc;тело;комент;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))   \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=2))};
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))}
});"Select * Where Col6 is not null Order by Col6";0);7;5))
)
)
)
(

COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc}

)`

}






function f_List1_ЛучшиеЦеныОктябрь2023_1() {

  let cf = `C${getClassColRow().list1_col_FirstPrice}`;
  let cl = `C${getClassColRow().list1_col_LastPrice}`;

  let cc = `C${getClassColRow().list1_col_Coment}`;
  let ce = `C${getClassColRow().list1_col_КомментарийЭксперта}`;
  let sheetName = getSettings().sheetName_Лист_1;


  return `LAMBDA(cc;тело;комент; комЭкс; порядок;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))  \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( Lambda(f;m;"Подходит: "& IFERROR( VLOOKUP( f; комЭкс; 4;0 );"-?")&", "&  m&" "&IFERROR(VLOOKUP( f; комЭкс; 2;0 );)&" "&IFERROR(VLOOKUP( f; комЭкс; 3;0 );)  )
(
 FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
 FILTER( r; MOD(COLUMN(r)-cc;3)=2)
) )   };
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))};

BYCOL({ARRAYFORMULA( FILTER(  "'1-1 Сбор КП'!"&"R1C"&COLUMN(r) ; MOD(COLUMN(r)-cc;3)=0))};LAMBDA(aa; 
LAMBDA(a;ifs (
IFERROR(FIND("(♠)";a);0)<>0;IFERROR( VLOOKUP("♠";порядок;2;FALSE);9);
IFERROR(FIND("(♦️)";a);0)<>0;IFERROR( VLOOKUP("♦️";порядок;2;FALSE);9);
IFERROR(FIND("(♣️)";a);0)<>0;IFERROR( VLOOKUP("♣️";порядок;2;FALSE);9);
IFERROR(FIND("(♥️)";a);0)<>0;IFERROR( VLOOKUP("♥️";порядок;2;FALSE);9);
TRUE;IFERROR( VLOOKUP("-";порядок;2;FALSE);9)
))(INDIRECT( aa;FALSE)  )
 ))

});"Select * Where Col6 is not null Order by Col7, Col6";0);7;5))
)
)
)
(

COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc};
BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)));
{{"♥️"\\0};{"♠"\\1};{"♣️"\\2};{"♦️"\\3};{"-"\\4}}

)`


  return `LAMBDA(cc;тело;комент; комЭкс;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))  \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( Lambda(f;m;"Подходит: "& IFERROR( VLOOKUP( f; комЭкс; 4;0 );"-?")&", "&  m&" "&IFERROR(VLOOKUP( f; комЭкс; 2;0 );)&" "&IFERROR(VLOOKUP( f; комЭкс; 3;0 );)  )
(
 FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
 FILTER( r; MOD(COLUMN(r)-cc;3)=2)
) )   };
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))}
});"Select * Where Col6 is not null Order by Col6";0);7;5))
)
)
)
(


COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc};
BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)))

)`





}



function f_List1_ЛучшиеЦеныОктябрь2023_0() {

  let cf = `C${getClassColRow().list1_col_FirstPrice}`;
  let cl = `C${getClassColRow().list1_col_LastPrice}`;

  let cc = `C${getClassColRow().list1_col_Coment}`;
  let ce = `C${getClassColRow().list1_col_КомментарийЭксперта}`;
  let sheetName = getSettings().sheetName_Лист_1;



  // {'3-3 Выбор поставщиков'!R102C28:R106C28 \ '3-3 Выбор поставщиков'!R102C32:R106C32};
  // {'3-3 Выбор поставщиков'!R102C28:R106C28 \ '3-3 Выбор поставщиков'!R102C30:R106C30}






  return `LAMBDA(cc;тело;комент; комЭкс; порядок;показ;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))  \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( Lambda(f;m;"Подходит: "& IFERROR( VLOOKUP( f; комЭкс; 4;0 );"-?")&", "&  m&" "&IFERROR(VLOOKUP( f; комЭкс; 2;0 );)&" "&IFERROR(VLOOKUP( f; комЭкс; 3;0 );)  )
(
 FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
 FILTER( r; MOD(COLUMN(r)-cc;3)=2)
) )   };
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))};

BYCOL({ARRAYFORMULA( FILTER(  "'1-1 Сбор КП'!"&"R1C"&COLUMN(r) ; MOD(COLUMN(r)-cc;3)=0))};LAMBDA(aa; 
LAMBDA(a;ifs (
IFERROR(FIND("(♠)";a);0)<>0;IFERROR( VLOOKUP("♠";порядок;2;FALSE);9);
IFERROR(FIND("(♦️)";a);0)<>0;IFERROR( VLOOKUP("♦️";порядок;2;FALSE);9);
IFERROR(FIND("(♣️)";a);0)<>0;IFERROR( VLOOKUP("♣️";порядок;2;FALSE);9);
IFERROR(FIND("(♥️)";a);0)<>0;IFERROR( VLOOKUP("♥️";порядок;2;FALSE);9);
TRUE;IFERROR( VLOOKUP("-";порядок;2;FALSE);9)
))(INDIRECT( aa;FALSE)  )
 ));

BYCOL({ARRAYFORMULA( FILTER(  "'1-1 Сбор КП'!"&"R1C"&COLUMN(r) ; MOD(COLUMN(r)-cc;3)=0))};LAMBDA(aa; 
LAMBDA(a;ifs (
IFERROR(FIND("(♠)";a);0)<>0;IFERROR( VLOOKUP("♠";показ;2;FALSE);TRUE);
IFERROR(FIND("(♦️)";a);0)<>0;IFERROR( VLOOKUP("♦️";показ;2;FALSE);TRUE);
IFERROR(FIND("(♣️)";a);0)<>0;IFERROR( VLOOKUP("♣️";показ;2;FALSE);TRUE);
IFERROR(FIND("(♥️)";a);0)<>0;IFERROR( VLOOKUP("♥️";показ;2;FALSE);TRUE);
TRUE;IFERROR( VLOOKUP("-";показ;2;FALSE);TRUE)
))(INDIRECT( aa;FALSE)  )
 ))


});"Select * Where Col6 is not null and Col8=true Order by Col7, Col6";0);7;5))
)
)
)
(

COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc};
BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)));
{'Настройки писем'!R2C43:R6C43 \\ 'Настройки писем'!R2C45:R6C45};
{'Настройки писем'!R2C43:R6C43 \\ 'Настройки писем'!R2C44:R6C44}

)`


  return `LAMBDA(cc;тело;комент; комЭкс;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))  \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( Lambda(f;m;"Подходит: "& IFERROR( VLOOKUP( f; комЭкс; 4;0 );"-?")&", "&  m&" "&IFERROR(VLOOKUP( f; комЭкс; 2;0 );)&" "&IFERROR(VLOOKUP( f; комЭкс; 3;0 );)  )
(
 FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
 FILTER( r; MOD(COLUMN(r)-cc;3)=2)
) )   };
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))}
});"Select * Where Col6 is not null Order by Col6";0);7;5))
)
)
)
(


COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc};
BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)))

)`





}


function f_List1_ЛучшиеЦеныОктябрь2023_2() {

  let cf = `C${getClassColRow().list1_col_FirstPrice}`;
  let cl = `C${getClassColRow().list1_col_LastPrice}`;

  let cc = `C${getClassColRow().list1_col_Coment}`;
  let ce = `C${getClassColRow().list1_col_КомментарийЭксперта}`;
  let sheetName = getSettings().sheetName_Лист_1;
  let sheetName7 = getSettings().sheetName_Лист_7
  let rf = getContext().getRowSobachkaBySheetName(sheetName7) + 2;
  let rl = rf + 5 - 1;
  let ctf = getContext().getSheetList7().columnTopFirst + 1;
  let str_range_filter = `
  {'${sheetName7}'!R${rf}C${ctf + 0}:R${rl}C${ctf + 0} \\ '${sheetName7}'!R${rf}C${ctf + 4}:R${rl}C${ctf + 4}};
  {'${sheetName7}'!R${rf}C${ctf + 0}:R${rl}C${ctf + 0} \\ '${sheetName7}'!R${rf}C${ctf + 2}:R${rl}C${ctf + 2}}
  `





  return `LAMBDA(cc;тело;комент; комЭкс; порядок;показ;
BYROW(тело;
LAMBDA(r;
TEXTJOIN("|||";FALSE;ARRAY_CONSTRAIN( Query( TRANSPOSE( {
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0))};
{ARRAYFORMULA( IFERROR(VLOOKUP(  FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
{ 
  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA( ISODD({SEQUENCE(1; COLUMNS(SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))  \\  TRANSPOSE (FILTER( {SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE)}; ARRAYFORMULA(  ISEVEN({SEQUENCE(1; COLUMNS( SPLIT(INDEX(комент;1;0);"|||";FALSE;FALSE) ); 1; 1)}))))}
 ;2;0 );""))};
{ARRAYFORMULA( FILTER( ADDRESS(ROW(r); COLUMN(r) ); MOD(COLUMN(r)-cc;3)=1))};
{ARRAYFORMULA( Lambda(f;m;"Подходит: "& IFERROR( VLOOKUP( f; комЭкс; 4;0 );"-?")&", "&  m&" "&IFERROR(VLOOKUP( f; комЭкс; 2;0 );)&" "&IFERROR(VLOOKUP( f; комЭкс; 3;0 );)  )
(
 FILTER( ADDRESS(1; COLUMN(r) ); MOD(COLUMN(r)-cc;3)=0);
 FILTER( r; MOD(COLUMN(r)-cc;3)=2)
) )   };
{ARRAYFORMULA( FILTER( r; MOD(COLUMN(r)-cc;3)=0))};

BYCOL({ARRAYFORMULA( FILTER(  "'1-1 Сбор КП'!"&"R1C"&COLUMN(r) ; MOD(COLUMN(r)-cc;3)=0))};LAMBDA(aa; 
LAMBDA(a;ifs (
IFERROR(FIND("(♠)";a);0)<>0;IFERROR( VLOOKUP("♠";порядок;2;FALSE);5);
IFERROR(FIND("(♦️)";a);0)<>0;IFERROR( VLOOKUP("♦️";порядок;2;FALSE);5);
IFERROR(FIND("(♣️)";a);0)<>0;IFERROR( VLOOKUP("♣️";порядок;2;FALSE);5);
IFERROR(FIND("(♥️)";a);0)<>0;IFERROR( VLOOKUP("♥️";порядок;2;FALSE);5);
TRUE;IFERROR( VLOOKUP("-";порядок;2;FALSE);5)
))(INDIRECT( aa;FALSE)  )
 ));

BYCOL({ARRAYFORMULA( FILTER(  "'1-1 Сбор КП'!"&"R1C"&COLUMN(r) ; MOD(COLUMN(r)-cc;3)=0))};LAMBDA(aa; 
LAMBDA(a;ifs (
IFERROR(FIND("(♠)";a);0)<>0;IFERROR( VLOOKUP("♠";показ;2;FALSE);TRUE);
IFERROR(FIND("(♦️)";a);0)<>0;IFERROR( VLOOKUP("♦️";показ;2;FALSE);TRUE);
IFERROR(FIND("(♣️)";a);0)<>0;IFERROR( VLOOKUP("♣️";показ;2;FALSE);TRUE);
IFERROR(FIND("(♥️)";a);0)<>0;IFERROR( VLOOKUP("♥️";показ;2;FALSE);TRUE);
TRUE;IFERROR( VLOOKUP("-";показ;2;FALSE);TRUE)
))(INDIRECT( aa;FALSE)  )
 ))


});"Select * Where Col6 is not null and Col8=true Order by Col7, Col6";0);7;5))
)
)
)
(

COLUMN('${sheetName}'!R[0]${cf});
'${sheetName}'!R[0]${cf}:R[0]${cl};
'${sheetName}'!R[0]${cc};
BYROW( TRANSPOSE(Split('1-1 Сбор КП'!R[0]${ce};"|||";FALSE));LAMBDA(rc;Split(rc;"📎";FALSE;FALSE)));
${str_range_filter}
)`

}

