function fillCalendar() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let edts = sheet.getDataRange().getValues();
  let exceptDates = [];
  let hours = 0.0;
  let ids = [];

  let sheetLine = 0;
  edts.forEach ( function(edt) {
    sheetLine++;
    if ((edt[2] instanceof Date) && (edt[3] instanceof Number)) {
      addThisEdt ( edt, hours, ids );
    }
  })
}

function addThisEdt ( edt, hours, ids ) {
  let idsIn = edt[12].split(" \,\;");

  if (edt[9] instanceof Date) {
    let exceptDates = [edt[9]];
  } else {
    let exceptDates = edt[9].split(" \,\;");
    exceptDates.foreach ( function(element,indexED,ExcD) {
      ExcD[indexED] = Date(ExcD[indexED]);
    })
  }

  for ( countSemanas=0; countSemanas<edt[3]; countSemanas++ ) {
    //
  }
}
