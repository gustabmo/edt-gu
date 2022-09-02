function fillCalendar() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let edts = sheet.getDataRange().getValues();
  let ids = [];
  let hours = 0.0;

  let sheetLine = 0;
  edts.forEach ( function(edt) {
    sheetLine++;
    if ((sheetLine<5) // @@@@@@@@@@@@@@@@@@@@@@
    && (edt[0] != "") && (edt[1] instanceof Date) && (edt[2] != "")) {
      hours,ids = addThisEdt ( edt );
    }
  })
}

function addThisEdt ( edt ) {
  let calendarId = 'ei0npiqhdcsc3rc9k0oj5s4u2o@group.calendar.google.com';
  let idsIn = edt[11].split(" \,\;");
  let idsOut = [];
  let totalMs = 0;
  let exceptDates = [];

  Logger.log ( '@@ exceptdates 1:'+edt[8]);
  if (edt[8] instanceof Date) {
    exceptDates = [edt[8]];
  } else {
    exceptDates = edt[8].split(" ");
    Logger.log ( '@@ exceptdates 2:'+exceptDates);
    exceptDates.forEach ( function(element,indexED,excD) {
      Logger.log ( '@@ exceptdates 3.1:'+excD[indexED]);
      excD[indexED] = new Date(excD[indexED]+"T00:00:00Z");
      Logger.log ( '@@ exceptdates 3.2:'+excD[indexED]);
    })
  }
  Logger.log ( '@@ exceptdates 4:'+exceptDates);

  Logger.log ( '@@ datems 5.1:'+edt[1]);
  let datems = edt[1].getTime();
  Logger.log ( '@@ datems 5.2:'+datems);
  Logger.log ( '@@ datems 5.3:'+new Date(datems));
  let numWeeks = Number(edt[2]);
  numWeeks = 1; // @@@@@@@@@@@@@@@@@@@@@@@@@@@@
  for ( countWeeks=0; countWeeks<numWeeks; countWeeks++ ) {
    for ( dow=0; dow<5; dow++ ) {
      let times = splitTimes ( edt[3+dow] );
      if ((times != null) && (! exceptDates.includes ( new Date(datems) ))) {
        idsOut.push (
          Calendar.Events.insert (
            {
              summary: edt[0],
              location: 'Ecole Rudolf Steiner de Genève',
              description: edt[0]+' '+edt[3+dow],
              start: {
                dateTime: new Date(datems+times[0]).toISOString()
              },
              end: {
                dateTime: new Date(datems+times[1]).toISOString()
              },
              colorId: Number(edt[9])
            },
            calendarId
          ).id
        );
        totalMs += times[1]-times[0];
        Logger.log('@@ added '+edt[0]);
      }
      datems += 24*60*60*1000;
    }
    datems += 2*24*60*60*1000; // jumps over the weekend
  }

  return { totalHours:totalMs/1000/60/60, idsOut:idsOut };
}

function splitTimes ( st ) {
  sthm = st.split("-");
  if (sthm.length != 2) return null;
  st0 = sthm[0].split(":");
  st1 = sthm[1].split(":");
  if ((st0.length != 2) || (st1.length != 2)) return null;
  t0h = Number(st0[0]);
  t0m = Number(st0[1]);
  t1h = Number(st1[0]);
  t1m = Number(st1[1]);
  if (
    (isNaN(t0h) || isNaN(t0m) || isNaN(t1h) || isNaN(t1m))
    ||
    ((t0h < 0) || (t0h >= 24) || (t0m < 0) || (t0m >= 60))
    ||
    ((t1h < 0) || (t1h >= 24) || (t1m < 0) || (t1m >= 60))
  ) return null;
  return [
    (t0h*60 + t0m) * 60 * 1000,
    (t1h*60 + t1m) * 60 * 1000
  ];
}


