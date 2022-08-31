function fillCalendar() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let edts = sheet.getDataRange().getValues();

  line = 0;
  edts.forEach ( function(edt) {
    line++;
    idsIn = edt[12].split(" \,\;");
    if (edt[9] instanceof Date) {
      exceptDates = [edt[9]];
    } else {
      exceptDates = edt[9].split(" \,\;");
      exceptDates.foreach ( function(element,indexED,ED) {
        ED[indexED] = Date(ED[indexED]);
      })
    }
    if ((edt[2] instanceof Date) && (edt[3] instanceof Number)) {
      //
      //
    }
  })
}
