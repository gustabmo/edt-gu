function fillCalendar() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let edts = sheet.getDataRange().getValues();
  let exceptDates = [];
  let hours = 0.0;
  let ids = [];

  let sheetLine = 0;
  edts.forEach ( function(edt) {
    sheetLine++;
    //Logger.log ( 'line '+sheetLine+' [0]:'+edt[0]+' [1]:'+edt[1]+' [2]:'+edt[2]);
    if ((edt[0] != "") && (edt[1] instanceof Date) && (edt[2] != "")) {
      addThisEdt ( edt, hours, ids );
    }
  })
}

function addThisEdt ( edt ) {
  //Logger.log('@@ addThisEdt '+edt[0]);
  let calendarId = 'ei0npiqhdcsc3rc9k0oj5s4u2o';
  let idsIn = edt[11].split(" \,\;");
  let idsOut = [];
  let totalHours = 0;
  let exceptDates = [];

  if (edt[9] == '') {
  } else if (edt[9] instanceof Date) {
    exceptDates = [edt[8]];
  } else {
    exceptDates = edt[8].split(" \,\;");
    exceptDates.forEach ( function(element,indexED,excD) {
      excD[indexED] = Date(excD[indexED]);
    })
  }

  let datems = edt[1].getTime();
  let numWeeks = Number(edt[2]);
  Logger.log ( 'num weeks:'+numWeeks+ ' exceptDates:'+exceptDates);
  for ( countWeeks=0; countWeeks<numWeeks; countWeeks++ ) {
    for ( dow=0; dow<5; dow++ ) {
      let times = splitTimes ( edt[3+dow] );
      Logger.log('edt[]:'+edt[3+dow]+' times:'+times);
      if ((times != null) && (! exceptDates.includes ( Date ))) {
        idsOut.append (
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
        totalHours += times[1]-times[0];
        Logger.log('@@ added '+edt[0]);
      }
      date++;
    }
    date += 2; // jumps over the weekend
  }

  Logger.log ( 'end addthisedt ids:'+idsOut );
  return { totalHours:totalHours, idsOut:idsOut };
}

function splitTimes ( st ) {
  Logger.log('splitTimes 0 '+st);
  st01 = st.split("-");
  Logger.log('splitTimes 1 '+st01);
  if (st01.length != 2) return null;
  st0 = st01[0].split(":");
  st1 = st01[1].split(":");
  Logger.log('splitTimes 2 '+st0+' '+st1 );
  if ((st0.length != 2) || (st1.length != 2)) return null;
  t0h = Number(st0[0]);
  t0m = Number(st0[1]);
  t1h = Number(st1[0]);
  t1m = Number(st1[1]);
  Logger.log('splitTimes 3 '+t0h+' '+t0m+' '+t1h+' '+t1m );
  Logger.log('splitTimes 4 '+typeof(t0h)+' '+typeof(t0m)+' '+typeof(t1h)+' '+typeof(t1m) );
  Logger.log('splitTimes 4.1.1 '+(typeof(t0h) != Number));
  Logger.log('splitTimes 4.1.2 '+(typeof(t0m) != Number));
  Logger.log('splitTimes 4.1.3 '+(typeof(t1h) != Number));
  Logger.log('splitTimes 4.1.4 '+(typeof(t1m) != Number));
  Logger.log('splitTimes 4.2 '+(t0h < 0) || (t0h >= 24) || (t0m < 0) || (t0m >= 60) );
  Logger.log('splitTimes 4.3 '+(t0h < 0) || (t0h >= 24) || (t0m < 0) || (t0m >= 60) );
  if (
    isNaN(t0h) || isNaN(t0m) || isNaN(t1h) || isNaN(t1m)
    ||
    (t0h < 0) || (t0h >= 24) || (t0m < 0) || (t0m >= 60)
    ||
    (t0h < 0) || (t0h >= 24) || (t0m < 0) || (t0m >= 60)
  ) return null;
  Logger.log('splitTimes 5');
  return [
    (t0h*60 + t0m)*60*1000,
    (t1h*60 + t1m)*60*1000
  ];
}


