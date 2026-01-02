
function logTrackingTrigger() {
  ScriptApp.newTrigger("logDailyTrackingToDatabase")
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(40)
    .create();
}


function calendarRescheduleTrigger() {
  ScriptApp.newTrigger("planSchedule")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .nearMinute(30)
    .create();
}


function deleteAllClockTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getEventType().toString() === "CLOCK") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function medTrackingTrigger() {
  ScriptApp.newTrigger("logMedTimeInputsToDatabase")
    .timeBased()
    .everyMinutes(30)
    .atHour(21)
    .nearMinute(58)
    .create();
}


function installUpcomingEventsTrigger() {
  // Avoid duplicates
  const handler = "logUpcomingCalendarEventsToDatabase";
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getEventType().toString() === "CLOCK" &&
    t.getHandlerFunction() === handler
  );
  if (exists) return;

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(15)
    .create();
}



function debugListTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    console.log({
      handler: t.getHandlerFunction(),
      type: t.getEventType(),
      source: t.getTriggerSource(),
      id: t.getUniqueId(),
    });
  });
}




function onOpen() {
  writeNext5UpcomingEventsToJ3N7();
  SpreadsheetApp.getUi()
    .createMenu('Funciones Manuales')
    .addItem("Importar Calendar Sheet", "importWeekFromToday")
    .addToUi();

  Utilities.sleep(500);
  updateEstudioWeekAndTotalSums();
}


function onEdit(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Pick the right dashboard config based on edited sheet
  const cfg = DASHBOARDS.find(d => d.sheetName === sheetName);
  if (!cfg) return;

  // Only react to the dropdown cell for that dashboard
  if (range.getA1Notation() !== cfg.dropdownCellA1) return;

  swapDashboardFromDatabase_(e, cfg);
}
