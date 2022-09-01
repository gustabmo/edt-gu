function fillCalendar() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let edts = sheet.getDataRange().getValues();
  let exceptDates = [];
  let hours = 0.0;
  let ids = [];

  let sheetLine = 0;
  edts.forEach ( function(edt) {
    sheetLine++;
    Logger.log ( 'line '+sheetLine+' [0]:'+edt[0]+' [1]:'+edt[1]+' [2]:'+edt[2]);
    Logger.log ( '[0]:'+(edt[0] != ""));
    Logger.log ( '[1]:'+(edt[1] instanceof Date));
    Logger.log ( '[2]:'+(edt[2] instanceof Number));
    if ((edt[0] != "") && (edt[1] instanceof Date) && (edt[2] != "")) {
      addThisEdt ( edt, hours, ids );
    }
  })
}

function addThisEdt ( edt ) {
  Logger.log('@@ addThisEdt '+edt[0]);
  let calendarId = 'ei0npiqhdcsc3rc9k0oj5s4u2o';
  let idsIn = edt[11].split(" \,\;");
  let idsOut = [];
  let totalHours = 0;

  if (edt[9] instanceof Date) {
    let exceptDates = [edt[8]];
  } else {
    let exceptDates = edt[8].split(" \,\;");
    exceptDates.foreach ( function(element,indexED,excD) {
      excD[indexED] = Date(excD[indexED]);
    })
  }

  let date = edt[1];
  let numWeeks = Number(edt[2]);
  for ( countWeeks=0; countWeeks<numWeeks; countWeeks++ ) {
    for ( dow=0; dow<5; dow++ ) {
      let times = splitTime ( edt[3+dow] );
      if ((times != null) && (! exceptDates.includes ( Date,))) {
        idsOut.append (
          Calendar.Events.insert (
            {
              summary: edt[0],
              locaion: 'Ecole Rudolf Steiner de Genève',
              description: edt[0]+' '+edt[3+dow],
              start: {
                dateTime: (date+times[0]).toISOString()
              },
              end: {
                dateTime: (date+times[1]).toISOString()
              },
              colorId: Number(edt[9])
            },
            calendarId
          ).id
        );
        totalHours += times[1]-times[0];
        Logger.log('@@ added '+edt[0]);
      }
      date++;
    }
    date += 2; // jumps over the weekend
  }

  return { totalHours:totalHours, idsOut:idsOut };
}

function splitTimes ( st ) {
  sthm = st.split("-");
  if (len(thm) != 2) return null;
  st0 = thm[0].split(":");
  st1 = thm[1].split(":");
  if ((len(t0) != 2) || (len(t1) != 2)) return null;
  t0h = Number(st0[0]);
  t0m = Number(st0[1]);
  t1h = Number(st1[0]);
  t1m = Number(st1[1]);
  if (
    (typeof(t0h) != Number) || (typeof(t0m) != Number) || (typeof(t1h) != Number) || (typeof(t1m) != Number)
    ||
    (t0h < 0) || (t0h >= 24) || (t0m < 0) || (t0m >= 60)
    ||
    (t1h < 0) || (t1h >= 24) || (t1m < 0) || (t1m >= 60)
  ) return null;
  return [
    t0h/24 + t0m/24/60,
    t1h/24 + t1m/24/60
  ];
}


