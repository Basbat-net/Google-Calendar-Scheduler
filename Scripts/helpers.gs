/**<--------------> FUNCIONES AUXILIARES <------------->**/

//REVISADO
function getNow_() {return new Date((new Date()).getTime() + NOW_OFFSET_HOURS * 60 * 60 * 1000);}
  
//REVISADO
function startOfDay_(d) { return new Date(new Date(d).setHours(0, 0, 0, 0));}

//REVISADO
function addDays_(d, days) {return new Date(new Date(d).getTime() + days * 86400000);}

//REVISADO
function isSameDay_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }

//REVISADO
function computeIntervalForPriority_(examDate, priority) {
  const bracket = priority === 1 ? { min: 22, max: 365 } : { min: (5 - priority) * 7, max: (5 - priority) * 7 + 7};
  return { desde: addDays_(examDate, -bracket.max), hasta: addDays_(examDate, -bracket.min), priority }; }
//REVISADO
function getFillRatio_(task) {
  if(!task.deadline || task.isOverdue || (task.deadline - now) <= 0 || !task.etaMinutesRemaining || task.etaMinutesRemaining <= 0) return 0;
  return (task.etaMinutesRemaining / ((task.deadline - now) / 60000)); }

//REVISADO
function fetchCalendar_(name) {
  const existing = CalendarApp.getCalendarsByName(name);
  if (existing && existing.length > 0) return existing[0];
  return CalendarApp.createCalendar(name); }


//REVISADO
//Limpia hasta la fecha que le digas los eventos que le digas de un calendario (a partir de despues del actual evento)
function clearAutomaticEvents_(cal, now, end) {
  let events;
  try {
    events = cal.getEvents(now, end);
  } catch (e) {
    Logger.log("getEvents falló en calendario '" + cal.getName() + "' (" + now + " → " + end + "): " + e);
    return; // si peta este calendario, lo ignoramos y seguimos
  }
  events.forEach(ev => {
    const s = ev.getStartTime();
    const e = ev.getEndTime();
    // Si el evento está activo justo ahora, lo respetamos
    if (s <= now && e > now) return;
    ev.deleteEvent();
  });
}

//REVISADO
function dateKeyFromDate_(d){return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);}

//REVISADO
function dateFromKey_(key) {
  const [y, m, d] = key.split('-').map(Number);
  return new Date(y, m - 1, d);
}

//REVISADO
function roundUpToNextHalfHour_(d) {
  const res = new Date(d);
  res.setSeconds(0, 0);
  const m = res.getMinutes();    
  res.setMinutes(m < 30 ? 30 : 0);
  if (res.getMinutes() >= 30) res.setHours(res.getHours() + 1);
  return res;
}

//REVISADO
function createMergedEventsForSegments_(cal, title, segments) {
  if (!cal || !segments || segments.length === 0) return;

  segments.sort((a, b) => a.start.getTime() - b.start.getTime());

  let current = {
    start: new Date(segments[0].start.getTime()),
    end: new Date(segments[0].end.getTime())
  };

  for (let i = 1; i < segments.length; i++) {
    const seg = segments[i];
    const segStart = seg.start;
    const segEnd = seg.end;

    if (segStart.getTime() <= current.end.getTime() + 10000) {
      if (segEnd.getTime() > current.end.getTime()) {
        current.end = new Date(segEnd.getTime());
      }
    } else {
      cal.createEvent(title, current.start, current.end);
      current = {
        start: new Date(segStart.getTime()),
        end: new Date(segEnd.getTime())
      };
    }
  }

  cal.createEvent(title, current.start, current.end);
}
