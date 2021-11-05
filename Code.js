// Tick Killer
/*
########## STEPS ##########
1. Establecer el Subscription ID, el API Token, el correo de usuario y la frecuencia de actualización usando un prompt
2. Extraer de Tick el id del usuario y almacenarlo en una variable para consultar las entradas
3. Extraer de Tick los Clientes y escribirlos en la hoja clients
4. Extraer de Tick los Proyectos y escribirlos en la hoja projects
5. Extraer de Tick las Tareas y escribirlas en la hoja tasks
6. Extraer de Tick las Entradas del usuario de los últimos 2 meses y escribirlas en la hoja entries para analizar
7. Extraer de Calendar los eventos semanales en la hoja calendar, ver: https://www.cloudbakers.com/blog/export-google-calendar-entries-to-a-google-spreadsheet
8. Clasificar las entradas del calendario por task y crear los entries con la duración de cada evento
9. Crear el loop para hacer cada POST de las entradas a Tick

########## TRIGGERS ##########
onInstall   ->  setApiCreds
menuItem    ->  updateTickData
SEMANAL     ->  importCalendar, postTick
*/

// Global variables
var subscriptionId = null
var apiToken = null
var email = null
var frequency = null
var rowCount = 2;

function onInstall(e) {
    onOpen(e);
}

function onOpen() {

    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Actualizar datos de Tick', 'updateTickData')
        .addItem("Configurar Add-On", "setApiCreds")
        .addToUi();
}

function setApiCreds() {
    var widget = HtmlService.createHtmlOutputFromFile("dialog.html");
    SpreadsheetApp.getUi().showModalDialog(widget, "Configuración Inicial");
}
function setUserCreds(form) {
    subscriptionId = form.subscriptionId;
    apiToken = form.apiToken;
    email = form.email;
    frequency = form.freq;

    console.log("subscriptionId: " + subscriptionId);
    console.log("apiToken:" + apiToken);
    console.log("email: " + email);
    console.log("frequency: " + frequency);
}


function importCalendar() {

    //
    // Export Google Calendar Events to a Google Spreadsheet
    //
    // This code retrieves events between 2 dates for the specified calendar.
    // It logs the results in the current spreadsheet starting at cell A2 listing the events,
    // dates/times, etc and even calculates event duration (via creating formulas in the spreadsheet) and formats the values.
    //
    // I do re-write the spreadsheet header in Row 1 with every run, as I found it faster to delete then entire sheet content,
    // change my parameters, and re-run my exports versus trying to save the header row manually...so be sure if you change
    // any code, you keep the header in agreement for readability!
    //
    // 1. Please modify the value for mycal to be YOUR calendar email address or one visible on your MY Calendars section of your Google Calendar
    // 2. Please modify the values for events to be the date/time range you want and any search parameters to find or omit calendar entires
    // Note: Events can be easily filtered out/deleted once exported from the calendar
    // 
    // Reference Websites:
    // https://developers.google.com/apps-script/reference/calendar/calendar
    // https://developers.google.com/apps-script/reference/calendar/calendar-event
    //

    var mycal = "jcorona@epa.digital";
    var cal = CalendarApp.getCalendarById(mycal);

    // Optional variations on getEvents
    // var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"));
    // var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
    // 
    // Explanation of how the search section works (as it is NOT quite like most things Google) as part of the getEvents function:
    //    {search: 'word1'}              Search for events with word1
    //    {search: '-word1'}             Search for events without word1
    //    {search: 'word1 word2'}        Search for events with word2 ONLY
    //    {search: 'word1-word2'}        Search for events with ????
    //    {search: 'word1 -word2'}       Search for events without word2
    //    {search: 'word1+word2'}        Search for events with word1 AND word2
    //    {search: 'word1+-word2'}       Search for events with word1 AND without word2
    //
    var events = cal.getEvents(new Date("January 12, 2021 00:00:00 CST"), new Date("January 15, 2021 23:59:59 CST"), { search: '' });


    var sheet = SpreadsheetApp.getActive().getSheetByName('calendar');
    // Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
    sheet.clearContents();

    // Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
    // of the getRange entry below
    var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event"]]
    var range = sheet.getRange(1, 1, 1, 14);
    range.setValues(header);


    // Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
    for (var i = 0; i < events.length; i++) {
        var row = i + 2;
        var myformula_placeholder = '';
        // Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
        // NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
        var details = [[mycal, events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()]];
        var range = sheet.getRange(row, 1, 1, 14);
        range.setValues(details);

        // Writing formulas from scripts requires that you write the formulas separate from non-formulas
        // Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
        var cell = sheet.getRange(row, 7);
        cell.setFormula('=(HOUR(F' + row + ')+(MINUTE(F' + row + ')/60))-(HOUR(E' + row + ')+(MINUTE(E' + row + ')/60))');
        cell.setNumberFormat('.00');

    }
}

function updateTickData() {
    getUserId();
    update();
}

function getUserId() {
    // get user id
    var headers = {
        "Authorization": "Token token=18f553771524ed323d95dd8444ea31c3",
        "Content-Type": "application/json",
        "User-Agent": "epa_test (jcorona@epa.digital)"
    };
    var options = {
        'method': 'GET',
        'headers': headers,
        'redirect': 'follow'
    };
    var request = UrlFetchApp.fetch('https://www.tickspot.com/99327/api/v2/users.json', options);
    var myObj = JSON.parse(request);
    userId = parseInt(myObj[0].id, 10);
    Logger.log("el user ID es: " + userId)
}

function getUserId() {
    // get user id
    var headers = {
        "Authorization": "Token token=18f553771524ed323d95dd8444ea31c3",
        "Content-Type": "application/json",
        "User-Agent": "epa_test (jcorona@epa.digital)"
    };
    var options = {
        'method': 'GET',
        'headers': headers,
        'redirect': 'follow'
    };
    var request = UrlFetchApp.fetch('https://www.tickspot.com/99327/api/v2/users.json', options);
    var myObj = JSON.parse(request);
    userId = parseInt(myObj[0].id, 10);
    Logger.log("el user ID es: " + userId)
}



function getClients() {
    // get user id
    var headers = {
        "Authorization": "Token token=18f553771524ed323d95dd8444ea31c3",
        "Content-Type": "application/json",
        "User-Agent": "epa_test (jcorona@epa.digital)"
    };
    var options = {
        'method': 'GET',
        'headers': headers,
        'redirect': 'follow'
    };
    var response = UrlFetchApp.fetch('https://www.tickspot.com/99327/api/v2/clients.json', options);
    var object = JSON.parse(response.getContentText());
    for (let i = 0; i < object.length; i++) {
      writeSheets(object[i])
    }
}

function writeSheets(object) {
    var values = Object.values(object);
    for (let i = 0; i < values.length; i++) {
      values[i] = String(values[i])
    }
    let sheetsValues = [];
    sheetsValues.push(values)
    var sheet = SpreadsheetApp.getActive().getSheetByName('getblue_data')
    sheet.getRange('A'+ rowCount + ':' + 'E' + rowCount).clearContent();
    sheet.getRange('A'+ rowCount + ':' + 'E' + rowCount).setValues(sheetsValues);
    rowCount++
}



function postTick() {
  // upload calendar entries
    var payload = JSON.stringify({
        "date": "2021-10-29",
        "hours": "8",
        "notes": "testing",
        "task_id": 15914019
    });

    var headers = {
        "Authorization": "Token token=18f553771524ed323d95dd8444ea31c3",
        "Content-Type": "application/json",
        "User-Agent": "epa_test (jcorona@epa.digital)"
    };

    var options = {
        'method': 'POST',
        'headers': headers,
        'body': 'raw',
        'payload': payload,
        'redirect': 'follow'
    };
    var send = UrlFetchApp.fetch('https://www.tickspot.com/99327/api/v2/entries.json', options);
    Logger.log(send);
}