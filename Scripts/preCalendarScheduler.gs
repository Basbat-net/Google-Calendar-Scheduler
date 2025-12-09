/**<--------------> FUNCIONES DE ORGANIZACIÓN <------------->**/

//REVISADO
function getFixedEvents_(start, end) {
  const busy = [];
  const calendarNames = Object.keys(FIXED_CALENDARS); 
  // los ... es para que en el array, no te meta el subarray de FIXED_CALENDARS, que te meta cada elemento

  calendarNames.forEach(name => {
    const cal = fetchCalendar_(name);
    const events = cal.getEvents(start, end);

    events.forEach(ev => {
      let evStart = ev.getStartTime();
      let evEnd   = ev.getEndTime();

      const title = ev.getTitle() || ""; // si el titulo está vacio que sea "", el || es un operador or en la variable
                                         // si, es una criminalada, pero bueno he aprendido q existe asiq se usa
      
      let prepMinutes = 0;
      //analisis del churro
      // - / /i regex que ignora mayusculas por si h es minuscula o mayuscula
      // - \[ \] los caracteres de bracket pero les metes el escape porque si no se ralla
      // - lo que está entre los segundos () es lo que se guarda en la variable
      // - Ese bracket tiene un ? al final para que si no pongo decimal no pasa nada
      // - \d+ es para pillar uno o más digitos (para que pueda poner 12 por ejemplo)
      // - dentro del tercer (), hay un ?: que significa que te AGRUPE dentro del match de los 2os () en vez
      //   de que te lo meta dentro de otro elemento en el array, se ve que puedes hacer match de varias cosas a la vez
      // - El resto es \. punto para el decimal y \d+ para pillar todos los decimales
      // IMPORTANTE: la funcion match solo pilla lo que está dentro de los grupos capturantes () que no lleven un ?:,
      // asi que si la lías puede que match en vez de devolverte 
      const match = title.match(/\[(\d+(?:\.\d+)?)h\]/i);
      // IMPORTANTE 2: match es una basura y si te encuentra [12.5h], te va a meter en el indice 0 "[12.5h]", y 
      // en el indice 1 12.5, no se puede cambiar asi que hay que pillar siempre el 1 si le metes logica con regex
      if (match) {
        const prepHours = parseFloat(match[1]);  // ejemplo de lo que he puesto arriba
        if (!isNaN(prepHours) && prepHours > 0) prepMinutes = prepHours * 60;
        // pongo lo de isnan porque te puede pillar un [] o que el tiempo esté mal puesto        
      }

      // una vez  que tenemos el prepTime, se lo restamos al evento para artificialmente alargar el evento y que 
      // las cosas no se puedan poner en ese intervalo
      let effectiveStart = new Date(evStart.getTime() - prepMinutes * 60 * 1000);
      let effectiveEnd   = evEnd;

      // Hay una opción para que algunos calendarios tengan margenes hardcoded a todos los eventos, se añaden aqui 
      const clampHours = FIXED_CALENDARS[name] || 0;
      if (clampHours > 0) {
        const offsetMs = clampHours * 60 * 60 * 1000;
        effectiveStart = new Date(effectiveStart.getTime() - offsetMs);
        effectiveEnd   = new Date(effectiveEnd.getTime() + offsetMs);
      }

      // Solo metemos eventos en el intervalo q nos importa
      if (effectiveStart < end && effectiveEnd > start) {
        const clampedStart = effectiveStart < start ? start : effectiveStart;
        const clampedEnd   = effectiveEnd > end     ? end   : effectiveEnd;
        busy.push({ start: clampedStart, end: clampedEnd });
      }
    });
  });

  return busy;
}

//REVISADO
function buildFreeSlots_(start, end){
  const rangeStart = roundUpToNextHalfHour_(start);
  const rangeEnd   = roundUpToNextHalfHour_(end);
  const fixedBusyEvents = getFixedEvents_(rangeStart, rangeEnd); // aqui sacamos lo que no podemos tocar

  // Ahora iremos metiendo directamente los huecos ya fusionados
  const mergedFreeSlots = [];

  // Este bucle es el que mete los intervalos de media hora en el array
  for (let d = new Date(rangeStart); d < rangeEnd; d = new Date(d.getTime() + (24 * 60 * 60 * 1000))) {
    // itera cada dia
    const day = new Date(d.getFullYear(), d.getMonth(), d.getDate());

    // IMPORTANTE: Los intervalos son continuos, no se parten
    const dow = day.getDay(); // 0 = Sunday, 6 = Saturday
    const blocks = (dow === 0 || dow === 6)? INTERVALOS_USABLES_FINDES: INTERVALOS_USABLES_ENTRE_SEMANA;

    // Itera todo el dia en intervalos de media hora
    blocks.forEach(block => {
      let blockStart = new Date(day.getFullYear(), day.getMonth(), day.getDate(), block.startHour, block.startMinute);
      let blockEnd = new Date(day.getFullYear(), day.getMonth(), day.getDate(), block.endHour, block.endMinute);

      // para ignorar cosas fuera de los horarios de curro
      if (blockEnd <= rangeStart || blockStart >= rangeEnd) return;

      // filtra para que las cosas te caigan en el intervalo que toca
      if (blockStart < rangeStart && isSameDay_(blockStart, rangeStart)) blockStart = new Date(rangeStart); 
      if (blockEnd > rangeEnd && isSameDay_(blockEnd, rangeEnd)) blockEnd = new Date(rangeEnd); 
      let slotStart = new Date(blockStart);

      // analiza en intervaos de media hora el bloque
      while (slotStart < blockEnd && slotStart < rangeEnd) {
        const slotEnd = new Date(slotStart.getTime() + BLOCK_MINUTES * 60 * 1000);
        if (slotEnd > blockEnd || slotEnd > rangeEnd) break;
        if (slotEnd <=  getNow_() && (slotStart = slotEnd)) continue; // puedes asignar una variable como condicion lmaoo

        const busyFixed = fixedBusyEvents.some(ev => {
          // Mira a ver si alguno de los elementos de FixedBusyEvents se solapa con el intervalo de 30 min actual
          return ev.start < slotEnd && ev.end > slotStart;
        });

        if (!busyFixed) {
          // Si no se solapa, es un freeSlot y se intenta meter al array
          const last = mergedFreeSlots[mergedFreeSlots.length - 1]; // como es push siempre miramos el ultimo para comparar
          if (last && last.end.getTime() === slotStart.getTime()) last.end = new Date(slotEnd); 
            // si el anterior intervalo de mergedFreeSlots es consecutivo, en vez de añadirlo, se modifica el anterior
           else mergedFreeSlots.push({start: new Date(slotStart),end: new Date(slotEnd)});
        }
        slotStart = slotEnd;
      }
    });
  }
  Logger.log(JSON.stringify(mergedFreeSlots, null, 2));
  return mergedFreeSlots;
}

// REVISADO
function sortEventsByUrgency_(events) {
    const sortedEvents = [...events];

  // Se ordena directamente el array, ignorando las prioridades dado que luego solo sirven para cosas dentro de sortTowers
  sortedEvents.sort((a, b) => {
    const da = a.deadline ? a.deadline.getTime() : Infinity;
    const db = b.deadline ? b.deadline.getTime() : Infinity;
    if (da !== db) return da - db; // deadline más cercano primero

    // 1 - Prioridad overdue
    if (a.isOverdue || b.isOverdue) {
      if (a.isOverdue && !b.isOverdue) return -1;
      if (!a.isOverdue && b.isOverdue) return 1;

      // Si las dos están pasadas, va primero la que necesita más tiempo
      if (a.etaMinutesRemaining !== b.etaMinutesRemaining) {
        return b.etaMinutesRemaining - a.etaMinutesRemaining;
      }
    }

    // 1.5 - En prioridad 5, las Tareas van antes que Estudio
    if (a.prio === 5 && b.prio === 5 && a.type !== b.type) {
      if (a.type === "Tareas") return -1;
      if (b.type === "Tareas") return 1;
    }

    // 2 - Porcentaje de ocupación del tiempo restante hasta el deadline
    const fillA = getFillRatio_(a);
    const fillB = getFillRatio_(b);
    if (fillA !== fillB) return fillB - fillA;

    return 0;
  });

  Logger.log("sorted thingies");
  Logger.log(JSON.stringify(sortedEvents, null, 2));
  return sortedEvents;
}

// REVISADO
function getPendingTasks_(now) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TASKS_SHEET_NAME);// saca la sheet activa
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_TASK_ROW) return [];
  
  const numRows = lastRow - FIRST_TASK_ROW + 1;
  
  // te saca un array de arrays ordenado según le pongas ahi [fila][columna]
  const values = sheet.getRange(FIRST_TASK_ROW, COL_CONCEPTO, numRows, COL_PROGRESO).getValues();
  
  const tasks = [];
  for (let i = 0; i < values.length; i++) {
    //Sacamos todos los valores que nos interesan
    const rowIndex = FIRST_TASK_ROW + i;
    const row = values[i];
    const concepto = row[COL_CONCEPTO - 1];
    let deadline = row[COL_DEADLINE - 1];
    // ajustamos la deadline para que sea el dia de antes a las 23:59    
    if (deadline instanceof Date && !isNaN(deadline)) {
      const effDeadline = new Date(deadline.getTime());
      effDeadline.setDate(effDeadline.getDate() - 1);
      effDeadline.setHours(23, 59, 0, 0);
      deadline = effDeadline;
    }
    let prio = Number(row[COL_PRIO - 1]) || 0;
    const etaHours = Number(row[COL_ETA - 1]) || 0;
    let progreso = row[COL_PROGRESO - 1]/100;

    if (!concepto || !etaHours) continue;  // si no tiene concepto o eta callaooo
    const rawMinutes = etaHours * 60 * (1 - progreso);
    
    if (rawMinutes <= 0) continue; // para que ignore las que están completas
    let remainingMinutes = rawMinutes;


    let etaMinutesRemaining = Math.ceil(remainingMinutes / 60)*60; //lo redondeamos a la hora superior para que sea más facil colocarlo
    if (etaMinutesRemaining <= 0) continue;

    const isOverdue = deadline !== null && deadline.getTime() < now.getTime(); // booleano de si ya ha pasado la fecha
    if (isOverdue) prio = 5;
    // start = ahora
    const start = now;

    // allowedDays = todos los días desde ahora hasta deadline
    const allowedDays = [];
    let d = startOfDay_(start);
    const endD = startOfDay_(deadline);

    while (d <= endD) {
      allowedDays.push(dateKeyFromDate_(d));
      d = addDays_(d, 1);
    }

    tasks.push({
      type: "Tareas", rowIndex, concepto, start, deadline, prio, bracket: prio,                   // bracket = prio
      allowedDays, etaMinutesRemaining, totalMinutesOriginal: etaMinutesRemaining,
      blocksTotal: etaMinutesRemaining / 30, isOverdue });

  }

  return tasks;
}

// REVISADO
function getScaledStudyHours_(priority, start, end, now) {
  const baseHours = BASE_STUDY_HOURS[priority] || 0;
  if (!start || !end || baseHours <= 0) return 0;

  const weekStart = new Date(now), weekEnd = addDays_(weekStart, 7); // intervalo para la semana
  const overlapStart = Math.max(start.getTime(), weekStart.getTime()); // La fecha posterior de las dos es el inicio
  const overlapEnd = Math.min(end.getTime(), weekEnd.getTime()); // La fecha anterior de las dos es la final
  const overlapMs = Math.max(0, overlapEnd - overlapStart); // Restamos las dos para que salga un numero en ms
  const totalWeekMs = weekEnd - weekStart;

  // Redondea hacia arriba cada media hora pero te lo saca en horas (saca o .0 o .5)
  return Math.ceil((baseHours * 60 * (overlapMs / totalWeekMs)) / 30) * 30 / 60;
}

// REVISADO
function buildExamPriorityArray_() {

  // DATOS PARA OPERAR
  const cal = fetchCalendar_(EXAM_CALENDAR_NAME);
  const now = getNow_();              
  const today = startOfDay_(now);
  const windowEnd = addDays_(today, 28);
  const events = cal.getEvents(today, windowEnd);
  Logger.log("Found %s events in the next 4 weeks.", events.length);

  const result = [];

  // Intentamos colocar los intervalos de prioridad para todos los examenes en las proximas 4 semanas
  events.forEach(event => {
    const title = event.getTitle();
    const examStart = event.getStartTime();
    const examDate = startOfDay_(examStart);
    if (examDate < today || examDate > windowEnd) return;

    const diffDays = Math.floor((examDate - today) / (24 * 60 * 60 * 1000));
    if (diffDays < 0) return;

    // Para sacar los examenes, se tratan en intervalos de una semana hasta el examen, asi que siempre se van a partir
    // En dos intervalos uno de prioridad x y otro de prioridad x+5, esto saca esos numeros
    // Ejemplo: Examen Martes 12, Codigo ejecutado Domingo 3
    //          Intervalo prio 4: Domingo 3-> Lunes 4
    //          Intervalo prio 5: Martes 5 -> Lunes 11
    const currentPriority = Math.min(5, Math.max(1, 6 - Math.floor(diffDays / 7))); // saca la prioridad del 5 al 1 en función de cuanto de cerca está
                                                                                    // 5 < 7 dias, 4 < 14 dias, 3 < 21 dias, 2 < 28 dias, 1 > 28 dias
    const nextPriority = currentPriority + 1 + (currentPriority === 1) > 5 ? null : currentPriority + 1 + (currentPriority === 1);
    // Solo saca valor si la prioridad es inferior a 4
    
    // Intervalo 1 (el más lejano)
    const currentInterval = computeIntervalForPriority_(examDate, currentPriority); // devuelve inicio, final y prioridad en un dict
    if (currentInterval) {
      // Datos del intervalo
      let intervalStart = now;
      const deadline = new Date(currentInterval.hasta);
      deadline.setHours(examStart.getHours(), examStart.getMinutes());
      const hours = getScaledStudyHours_(currentInterval.priority, intervalStart, deadline, now);
      const minutes = Math.ceil((hours * 60) / 30) * 30;
      let allowedDays = [];
      let d1 = startOfDay_(intervalStart);
      let endD1 = startOfDay_(deadline);
      while (d1 <= endD1) {
        allowedDays.push(dateKeyFromDate_(d1));
        d1 = addDays_(d1, 1);
      }

      result.push({
        type:"Estudio", start:intervalStart, deadline:deadline,
        concepto:title, prio:currentInterval.priority, bracket:currentInterval.priority, isOverdue:false,
        allowedDays:allowedDays, etaMinutesRemaining:minutes, totalMinutesOriginal:minutes, blocksTotal:minutes/30
      });

    }

    // Intervalo 2 (el más cercano)
    if (nextPriority !== null) {
      const nextInterval = computeIntervalForPriority_(examDate, nextPriority);
        // Datos del intervalo
        const start = new Date(nextInterval.desde);
        start.setHours(examStart.getHours(), examStart.getMinutes());
        const deadlineNext = new Date(nextInterval.hasta);
        deadlineNext.setHours(examStart.getHours(), examStart.getMinutes());
        const hoursNext = getScaledStudyHours_(nextInterval.priority, start, deadlineNext, now);
        const minutesNext = Math.ceil((hoursNext * 60) / 30) * 30;
        let allowedDays2 = [];
        let d2 = startOfDay_(start);
        let endD2 = startOfDay_(deadlineNext);
        while (d2 <= endD2) {
          allowedDays2.push(dateKeyFromDate_(d2));
          d2 = addDays_(d2, 1);
        }

        result.push({
          type:"Estudio", start:start, deadline:deadlineNext,
          concepto:title, prio:nextInterval.priority, bracket:nextInterval.priority, isOverdue:false,
          allowedDays:allowedDays2, etaMinutesRemaining:minutesNext, totalMinutesOriginal:minutesNext, blocksTotal:minutesNext/30
        });


      }
  });

  // Una vez que hemos colocado todos los examenes, miramos a ver si alguno de las asignaturas está sin examen
  
  // Hacemos un set con todas las asignaturas
  const subjectsWithoutExams = new Set(SUBJECTS);
  // Si algun examen está en result, lo eliminamos del set
  result.forEach(item  => {
    const conceptoLower = String(item.concepto || "").toLowerCase(); // lowercase porsiaca
    SUBJECTS.forEach(subj => {
      const subjLower = subj.toLowerCase();
      if (conceptoLower.indexOf(subjLower) !== -1) subjectsWithoutExams.delete(subjLower); //los buscamos en el array de arriba
    });
  });

  // para las asignaturas que sobran les hacemos eventos de prioridad 1
  subjectsWithoutExams.forEach(subj => {
    const priority = 1;
    const start = today;
    const deadline = windowEnd;
    const hours = getScaledStudyHours_(priority, start, deadline, now);
    const minutes = Math.ceil((hours * 60) / 30) * 30;
    let allowedDays3 = [];
    let d3 = startOfDay_(start);
    let endD3 = startOfDay_(windowEnd);
    while (d3 <= endD3) {
      allowedDays3.push(dateKeyFromDate_(d3));
      d3 = addDays_(d3, 1);
    }

    result.push({
      type:"Estudio", start:start, deadline:windowEnd,
      concepto:subj, prio:priority, bracket:priority, isOverdue:false,
      allowedDays:allowedDays3, etaMinutesRemaining:minutes, totalMinutesOriginal:minutes, blocksTotal:minutes/30
    });

  });

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// REVISADO
// Esta funcion es la más complicada de todo el código, está separada por secciones y antes de cada seccion hay una explicación, además de los
// comentarios de cada cosa
// Nota: Aun queda por programar la parte que, para días con muchos eventos, purgue el día para que fisicamente quepan
function optimizeScheduleWithTowers_(freeSlots, eventsByBracket) {
  

  Logger.log("lo que entra a las torres");
  Logger.log(JSON.stringify(eventsByBracket, null, 2));
  // 1- PREPARACIÓN DE DATOS INICIALES
  // Extracción de los minutos por día
  const freeMinutesByDay = {};
  let globalRangeStart = null;
  let globalRangeEnd = null;

  freeSlots.forEach(slot => {
    const s = new Date(slot.start), e = new Date(slot.end);

    // Actualizamos los límites globales del rango
    if (!globalRangeStart || s < globalRangeStart) globalRangeStart = s;
    if (!globalRangeEnd   || e > globalRangeEnd)   globalRangeEnd   = e;

    // Vamos sumando minutos al dia en el diccionario
    const minutes = (e.getTime() - s.getTime()) / (60 * 1000);
    const key = dateKeyFromDate_(s);
    freeMinutesByDay[key] = (freeMinutesByDay[key] || 0) + minutes;
  });

  // Creacion de los arrays y dicts que nos servirán para luego
  const dayKeys = Object.keys(freeMinutesByDay).sort();
  const eventsMeta = [];
  const eventMetaById = {};
  let eventCounter = 0;
  
  eventsByBracket.forEach(ev => {
    const id = "E#" + (eventCounter++);
    const meta = { ...ev, id };
    eventsMeta.push(meta);
    eventMetaById[id] = meta;
  });

  // 2- ESTADO INICIAL DEL ALGORITMO
  // Primero, se va a intentar colocar todas las horas al final de su plazo, para crear la máxima disparidad posible
  const dayUsedBlocks = {};
  dayKeys.forEach(k => { dayUsedBlocks[k] = 0; });
  const eventDayBlocks = {};

  eventsMeta.forEach(meta => {
    // hace el bucle para cada evento
    const { id, type, blocksTotal, allowedDays } = meta;
    let remainingBlocks = blocksTotal;
    if (remainingBlocks <= 0) return;

    eventDayBlocks[id] = eventDayBlocks[id] || {concepto:meta["type"] + ": " + meta["concepto"]}; // Lo inicializa si está vacío

    // Mira en que dias puede poner los eventos, pero solo incluye los dias dentro del intervalo actual de freeSlots
    const allowedDesc = [...allowedDays]
      // Probamos con la lista de minutos libres por dia, y la usamos de referencia para que si un dia no está en ese rango, no la cuente,
      // que si no jode el algoritmo
      .filter(k => Object.prototype.hasOwnProperty.call(freeMinutesByDay, k))
      // Además, el día tiene que tener hueco útil antes de la hora de deadline de ese evento
      .filter(k => isDayUsableForEvent_(meta, k))
      .sort((a, b) =>
        dateFromKey_(b).getTime() - dateFromKey_(a).getTime()
      );

    // Si ninguno de los días permitidos está en el horizonte actual, no intentamos
    // colocarlo ahora; se colocará en futuros runs cuando el horizonte avance.
    if (allowedDesc.length === 0) return;

    // luego para cada evento, itera por dias
    for (let i = 0; i < allowedDesc.length && remainingBlocks > 0; i++) {
      const dayKey = allowedDesc[i];
      // Saca cuantos bloques de este evento tengo por ahora en este día
      const currentBlocksUsed = eventDayBlocks[id][dayKey] || 0;
      
      const MAX_BLOCKS_PER_DAY = 6; 
      // Si el bloque se puede colocar entero, el min será remainingBlocks, si no el máximo que se
      let capacityForThisEvent = Math.min(remainingBlocks, MAX_BLOCKS_PER_DAY - currentBlocksUsed);
      if (capacityForThisEvent <= 0) continue;

      let toAssign = capacityForThisEvent; // esto si es tasks se resuelve directamente aqui

      if (type === "Estudio") {
        // Solo vamos a poner bloques de más de una hora en Estudio, este bloque comprueba eso
        if (currentBlocksUsed === 0) {
          if (remainingBlocks === 1) {
            continue;
          }
          toAssign = Math.max(toAssign, 2); // Inicio minimo 2h
        }
        toAssign = Math.min(toAssign, capacityForThisEvent);

        const finalCount = currentBlocksUsed + toAssign;
        if (finalCount === 1) continue; //  si sobra uno, no se va a colocar en un día
      }

      if (toAssign <= 0) continue; // si es 0 se salta directamente

      eventDayBlocks[id][dayKey] = currentBlocksUsed + toAssign; // actualizamos cuanto se ha usado del evento
      dayUsedBlocks[dayKey] = (dayUsedBlocks[dayKey] || 0) + toAssign; // actualizamos cuanto se ha usado del dia
      remainingBlocks -= toAssign; // para el siguiente bucle
    }

    if (remainingBlocks > 0)  Logger.log("Evento '" + (meta.concepto || meta.id) + "' no ha podido colocar " + (remainingBlocks * 30) + " minutos en la fase base.");
  });


  // Resultados del 1er bucle
  let stats = computeOverflowStats_(dayUsedBlocks);
  Logger.log("Estado inicial (tras fase 1, antes de rebalancear torres):");
  Logger.log(JSON.stringify({
    freeMinutesByDay: freeMinutesByDay,
    dayUsedBlocks: dayUsedBlocks,
    stats: stats
  }, null, 2));

  // 3- ITERACION POR PAREJAS
  // Vamos a probar pareja por pareja (empezando por la más alta como source y pillando targets) a ver si, al mover un bloque de intervalo de sitio,
  // la diferencia entre el overflow de las parejas y la media baja, iterando primero desde arriba (iteraciones generales), luego una iteracion
  // por parejas que se reiniciará cada vez que un emparejamiento produzca una mejora, y luego dentro de cada emparejamiento, una iteración de todos los eventos
  // del source para ver cual es el mejor movimiento que se puede realizar
  let improvedGlobal = false;
  const MAX_GLOBAL_ITER = 5000;
  const MAX_PAIR_STEPS = 100;
  for (let iter = 0; iter < MAX_GLOBAL_ITER; iter++) {
    stats = computeOverflowStats_(dayUsedBlocks);
    const overflow = stats.overflow;

    // Inicializamos el dia con su iverflow
    const dayInfos = dayKeys.map(dayKey => ({
      dayKey: dayKey,
      ov: overflow[dayKey] || 0,
      usedBlocks: dayUsedBlocks[dayKey] || 0
    }));

    if (dayInfos.length < 2) break;

    // Ordenamos los dias por overflow
    dayInfos.sort((a, b) => a.ov - b.ov);

    const n = dayInfos.length;
    let anyImprovementThisRound = false;

    // Recorremos todas las posibles parejas (high, low)
    // si hay alguna mejora en alguna iteracion, anyImprovementThisRound va a salir del for loop y empezar otra vez desde 
    // el for loop de MAX_GLOBAL_ITER (el de justo encima) 
    for (let hiIdx = n - 1; hiIdx >= 0 && !anyImprovementThisRound; hiIdx--) {
      // Cogemos aquí la información del dia High
      const highInfo = dayInfos[hiIdx];
      const sourceDay = highInfo.dayKey;

      // si no hay bloques en el "high", no se puede mover nada desde él
      if ((dayUsedBlocks[sourceDay] || 0) <= 0) continue;

      for (let loIdx = 0; loIdx < n && !anyImprovementThisRound; loIdx++) {
        if (loIdx === hiIdx) continue; //para que no se empareje consigo mismo

        // Información de los dias que vamos a comprar
        const lowInfo = dayInfos[loIdx];
        const targetDay = lowInfo.dayKey;
        const ovHigh = highInfo.ov;
        const ovLow = lowInfo.ov;
        if (ovHigh <= ovLow) continue; // Si son iguales no vas a mejorar nada asiq lo dejas


        // COMPRACIÓN POR PAREJAS
        let improvedAny = false;
        for (let steps = 0; steps < MAX_PAIR_STEPS; steps++) {
          stats = computeOverflowStats_(dayUsedBlocks);
          const overflowPair = stats.overflow; // De donde sacas el oveflow

          // Miramos los overflows de los dos dias
          const ovSource = overflowPair[sourceDay] || 0; 
          const ovTarget = overflowPair[targetDay] || 0;

          // Si alguno de los target es más grande que el source, ya no es el más con diferencia y vamos a otro
          if (ovSource <= ovTarget) break;

          // Parametros de evaluacion
          const avg = (ovSource + ovTarget) / 2; // Constante 
          const currentMetric = Math.abs(ovSource - avg) + Math.abs(ovTarget - avg);

          let bestMove = null;
          let bestMetric = currentMetric;
          // aqui iteramos los eventos para encontrar el mejor movimiento posible
          for (const evId in eventDayBlocks) {
            // sacamos los datos del evento
            const meta = eventMetaById[evId];
            if (!meta) continue;

            const allowed = meta.allowedDays;
            if (allowed.indexOf(sourceDay) === -1 || allowed.indexOf(targetDay) === -1) continue; // si el evento no se puede en ninguno lo saltamos

            // El movimiento solo tiene sentido si el día es físicamente usable para ese evento (por hora de deadline)
            if (!isDayUsableForEvent_(meta, sourceDay) || !isDayUsableForEvent_(meta, targetDay)) continue;
            
            // Datos para probar el movimiento
            const type = meta.type;
            const blocksOnSource = eventDayBlocks[evId][sourceDay] || 0;
            const blocksOnTarget = eventDayBlocks[evId][targetDay] || 0;
            if (blocksOnSource <= 0) continue; // si en el source no hay nada de ese evento, no podemos mover cosas y lo saltamos

            // Para probar movimientos, como Estudio tiene una duración minima de 1h, puede que en target no haya de ese
            // evento y no se pueda mover solo media hora, asi que permitimos los dos movimientos
            const stepOptions = (type === "Estudio") ? [1, 2] : [1];

            // Probamos a encontrar un movimiento para ese evento que mejore la situación
            for (let s = 0; s < stepOptions.length; s++) {
              const stepBlocks = stepOptions[s]; // el movimiento que estamos probando
              if (blocksOnSource < stepBlocks) continue;

              // Evaluamos los nuevos conteos de ese evento en ambos sitios
              const newSourceCount = blocksOnSource - stepBlocks;
              const newTargetCount = blocksOnTarget + stepBlocks;


              if (type === "Estudio") {
                // esta arrow function dice si el movimiento es valido en base a si en ambos se queda un conteo valido de intervalos de media hora
                const validEstudio = c => c === 0 || (c >= 2 && c <= 6); 
                if (!validEstudio(newSourceCount) || !validEstudio(newTargetCount)) continue;
              } else {
                const validTarea = c => c >= 0 && c <= 6; // igual que arriba solo que para tasks
                if (!validTarea(newSourceCount) || !validTarea(newTargetCount)) continue;
              }

              const deltaMinutes = stepBlocks * 30;
              const newOvSource = ovSource - deltaMinutes;
              const newOvTarget = ovTarget + deltaMinutes;

              // Si el movimiento solo permuta los valores, es un movimiento ilegal y nos lo saltamos
              if (newOvSource === ovTarget && newOvTarget === ovSource) continue;

              // probamos si la diferencia respecto a la media de ambos ha cambiado, y si ha mejorado, lo metemos en bestMove
              const newMetric = Math.abs(newOvSource - avg) + Math.abs(newOvTarget - avg);
              if (newMetric + 1e-6 < bestMetric) {
                bestMetric = newMetric;
                bestMove = {evId: evId, sourceDay: sourceDay, targetDay: targetDay,
                            stepBlocks: stepBlocks, newOvSource: newOvSource, newOvTarget: newOvTarget};
                }
            }
          }

          if (!bestMove) break; // en caso de que no haya mejor movimiento, se pasa a la siguiente pareja

          // REALIZACIÓN DEL MOVIMIENTO DE EVENTOS ENTRE LAS PAREJAS
          // sacamos todos los valores de bestMove a variables individuales (destructuring)
          const {
            evId,
            sourceDay: sDay,
            targetDay: tDay,
            stepBlocks,
            newOvSource,
            newOvTarget
          } = bestMove;
          const metaMove = eventMetaById[evId]; //sacamos el evento al que se refiere el movimiento

          const prevSourceCount = eventDayBlocks[evId][sDay] || 0;
          const prevTargetCount = eventDayBlocks[evId][tDay] || 0;

          // quitamos ese numero de bloques del dia donde origina el movimiento (source o sDay)
          eventDayBlocks[evId][sDay] = prevSourceCount - stepBlocks;
          // si se queda vacío el día, eliminamos ese dia de los dias donde está ese evento
          if (eventDayBlocks[evId][sDay] <= 0) delete eventDayBlocks[evId][sDay];
          // metemos esos dias en el target (tday)
          eventDayBlocks[evId][tDay] = prevTargetCount + stepBlocks; 

          // modificamos los conteos de bloques usados de tDay y sDay
          dayUsedBlocks[sDay] = (dayUsedBlocks[sDay] || 0) - stepBlocks;
          dayUsedBlocks[tDay] = (dayUsedBlocks[tDay] || 0) + stepBlocks;

          Logger.log(
            "Movimiento parejas: " + (stepBlocks * 30) + " min del evento '" +
            (metaMove.concepto || evId) + "' de " + sDay + " a " + tDay +
            " (overflow " + ovSource + "," + ovTarget + " → " +
            newOvSource + "," + newOvTarget + ")"
          );

          improvedAny = true; // para ver que se ha mejorado algo
        }
        // en caso de que el emparejamiento haya mejorado algo, significa que aun hay cosas que mejorar y se sigue
        if (improvedAny) {
          anyImprovementThisRound = true; // si esto es true, se sale del for loop de parejas y se empieza de cero
          improvedGlobal = true;
        }
      }
    }
    // si al realizar una iteracion de todas las parejas no se ha encontrado ninguna mejora, significa que está en el estado optimo y se sale
    if (!anyImprovementThisRound) break;
  }

  // sacamos los resultados si algo ha mejorado (por lo general va a siempre mejorar)
  if (improvedGlobal) {
    stats = computeOverflowStats_(dayUsedBlocks);
    Logger.log("Estado tras rebalanceo por parejas (torres):");
    Logger.log(JSON.stringify({
      dayUsedBlocks: dayUsedBlocks,
      stats: stats
    }, null, 2));
  }

  
  // Recalcular stats tras meter el extra
  stats = computeOverflowStats_(dayUsedBlocks);


  const resultPre = {
    freeMinutesByDay: freeMinutesByDay,
    finalDayUsedMinutes: Object.fromEntries(
      dayKeys.map(k => [k, (dayUsedBlocks[k] || 0) * 30])
    ),
    overflowFinal: stats.overflow,
    meanOverflowFinal: stats.meanOverflow,
    FFinal: stats.F,
    eventDayBlocks: eventDayBlocks,
    eventsMeta: eventsMeta,
    dayKeys: dayKeys
  };

  Logger.log("Resultado pre purga optimización por torres (precalendario lógico):");
  Logger.log(JSON.stringify(resultPre, null, 2));


  // FASE DE EMPEQUEÑECIMIENTO DE DIAS LLENOS
  // Una vez que hemos encontrado la distribución optima, algunos dias puede que necesiten más horas que las que hay disponibles, asi que toca filtrar
  // La prioridad será siempre Tasks 5 > Estudio 5 >  Resto de osas

  // Copiamos el overflow para ver que hay cada cosa
  let overflowNow = stats.overflow || {};

  const extraTimes = {};// para los dias que necesito meter tiempo extra

  // iteramos cada dia por separado probando a ver cuales tienen overflow positivo
  dayKeys.forEach(dayKey => {
    let ovMinutes = overflowNow[dayKey] || 0; // los overflows del dia
    if (ovMinutes <= 0) return; 
    // Si pasa esto es que toca purgar
    Logger.log("estoy dentro de un dia de purgado");

    let blocksToRemove = Math.ceil(ovMinutes / 30); // El numero de bloques a quitar
    Logger.log("bloques a eliminar: " + blocksToRemove);
    Logger.log("eventDayBlocks: ");
    Logger.log(JSON.stringify(eventDayBlocks, null, 2));
    

    // Construimos la lista de candidatos de ese día que pueden perder bloques
    const candidates = [];

    for (const evId in eventDayBlocks) {
      const meta = eventMetaById[evId];
      if (!meta) continue;

      const blocks = eventDayBlocks[evId][dayKey] || 0;
      if (blocks <= 0) continue;

      const isStudy = (meta.type === "Estudio");
      const minBlocks = isStudy ? 4 : 2; // El estudio (a primeras) no puede bajar d 2h, el resto no puede bajar de 1h

      if (blocks > minBlocks) {
        candidates.push({
          concepto: eventDayBlocks[evId]["concepto"],
          evId: evId,
          blocks: blocks,
          isStudy: isStudy
        });
      }
    }

    Logger.log("candidatos");
    Logger.log(candidates);

    while (candidates.length > 0 && blocksToRemove > 0){
      candidates.sort((a, b) => b.blocks - a.blocks);
      for (let i = 0; i < candidates.length && candidates.length > 0; i++) {
        const evId  = candidates[i]["evId"];
        const meta = eventMetaById[evId];
        const isStudy = (meta.type === "Estudio");
        const minBlocks = isStudy ? 4 : 2;
        const currentBlocks = eventDayBlocks[evId][dayKey] || 0;
        eventDayBlocks[evId][dayKey] = currentBlocks - 1;
        dayUsedBlocks[dayKey] = (dayUsedBlocks[dayKey] || 0) - 1;
        Logger.log("bloque purgado: "+ eventDayBlocks[evId]["concepto"]);
        Logger.log("Bloques antes: " + blocksToRemove);
        blocksToRemove--;
        Logger.log("Bloques despues: " + blocksToRemove);
        if (blocksToRemove === 0) break;
        if ((currentBlocks - 1) <= minBlocks){
          Logger.log(candidates);
          Logger.log("Indice: " + i);
          Logger.log("destruimos: ");
          Logger.log(candidates[i]);
          candidates.splice(i, 1);
          Logger.log("post delete:");
          Logger.log(candidates);
          Logger.log("longitud candidatos: " + candidates.length);
          break;
        }   
      }
    }


    if (blocksToRemove > 0) {

      Logger.log("EventDayBlocks antes del tiempo extra");
      Logger.log(JSON.stringify(eventDayBlocks, null, 2));

      // Mapa auxiliar para poder ir de evId → meta (deadline, allowedDays, etc.)
      const metaById = {};
      eventsMeta.forEach(m => { metaById[m.id] = m; });

      // Distribuye numBlocks de tiempo extra de un evento concreto a lo largo de los días
      // permitidos hasta su deadline, intentando repartir en paquetes de 2 bloques por día
      // y balanceando para que no haya un día con mucha sobrecarga si otros aún tienen 0.
      function distributeExtraBlocks_(evId, originDayKey, numBlocks) {
        const meta = metaById[evId];
        // Si por lo que sea no tenemos meta, caemos al comportamiento antiguo (todo al día origen).
        if (!meta) {
          extraTimes[evId] = extraTimes[evId] || {
            concepto: eventDayBlocks[evId]["concepto"]
          };
          extraTimes[evId][originDayKey] =
            (extraTimes[evId][originDayKey] || 0) + numBlocks;
          return;
        }

        const deadlineKey = meta.deadline ? dateKeyFromDate_(meta.deadline) : null;
        const allowedRaw = (meta.allowedDays && meta.allowedDays.length)
          ? meta.allowedDays
          : dayKeys; // fallback: todas las dayKeys conocidas

        // Días candidatos: desde el día de origen hasta el deadline (incluido)
        const candidates = allowedRaw.filter(d =>
          d >= originDayKey && (!deadlineKey || d <= deadlineKey)
        );

        extraTimes[evId] = extraTimes[evId] || {
          concepto: eventDayBlocks[evId]["concepto"]
        };

        // Si no hay candidatos válidos, metemos todo en el propio día origen
        if (candidates.length === 0) {
          extraTimes[evId][originDayKey] =
            (extraTimes[evId][originDayKey] || 0) + numBlocks;
          return;
        }

        // Repartimos: mientras queden bloques, intentamos ponerlos en chunks de 2
        // y siempre en el día con menor carga extra actual.
        while (numBlocks > 0) {
          const chunk = (numBlocks >= 2) ? 2 : 1; // idealmente 2; si queda 1, lo añadimos encima

          // buscamos el día con menor carga extra actual para este evento
          let bestDay = candidates[0];
          let bestLoad = extraTimes[evId][bestDay] || 0;

          for (let i = 1; i < candidates.length; i++) {
            const d = candidates[i];
            const load = extraTimes[evId][d] || 0;
            if (load < bestLoad) {
              bestLoad = load;
              bestDay = d;
            }
          }

          extraTimes[evId][bestDay] =
            (extraTimes[evId][bestDay] || 0) + chunk;
          numBlocks -= chunk;
        }
      }

      // vamos sacando tiempo extra de ese día hasta 3h (6 bloques) o hasta agotar blocksToRemove
      let movedExtraBlocks = 0; // máximo 6 bloques (3h)

      while (blocksToRemove > 0 && movedExtraBlocks < 6) {
        // reconstruimos candidatos restantes del día
        const remaining = [];
        for (const evId in eventDayBlocks) {
          const blocks = eventDayBlocks[evId][dayKey] || 0;
          if (blocks > 0) {
            remaining.push({ evId, blocks });
          }
        }

        if (remaining.length === 0) break;

        // cogemos el evento con más bloques en ese día
        remaining.sort((a, b) => b.blocks - a.blocks);
        const { evId } = remaining[0];

        const available = eventDayBlocks[evId][dayKey] || 0;
        // no sacar más de lo disponible ni más de lo que falta por quitar ni más de 3h en total
        let toMove = Math.min(blocksToRemove, 6 - movedExtraBlocks, available);
        if (toMove <= 0) break;

        // opcional: evitar mover un solo bloque suelto si quieres mantener pares de 2
        // if (toMove === 1 && blocksToRemove > 1) toMove = 0;
        // if (toMove <= 0) break;

        // en vez de asignar todo al mismo dayKey, lo repartimos por días hasta el deadline
        distributeExtraBlocks_(evId, dayKey, toMove);

        eventDayBlocks[evId][dayKey] -= toMove;
        if (eventDayBlocks[evId][dayKey] <= 0) delete eventDayBlocks[evId][dayKey];

        dayUsedBlocks[dayKey] -= toMove;
        blocksToRemove -= toMove;
        movedExtraBlocks += toMove;
      }

      Logger.log("EventDayBlocks despues del tiempo extra");
      Logger.log(JSON.stringify(eventDayBlocks, null, 2));

      Logger.log("extraTimes:");
      Logger.log(JSON.stringify(extraTimes, null, 2));
    }


  });


  // Recalcular stats tras meter el extra
  stats = computeOverflowStats_(dayUsedBlocks);
  // Resultados finales
  const result = {
    freeMinutesByDay: freeMinutesByDay,
    finalDayUsedMinutes: Object.fromEntries(
      dayKeys.map(k => [k, (dayUsedBlocks[k] || 0) * 30])
    ),
    overflowFinal: stats.overflow,
    meanOverflowFinal: stats.meanOverflow,
    FFinal: stats.F,
    eventDayBlocks: eventDayBlocks,
    extraTimes: extraTimes,
    eventsMeta: eventsMeta,
    dayKeys: dayKeys
  };

  Logger.log("Resultado final optimización por torres (precalendario lógico):");
  Logger.log(JSON.stringify(result, null, 2));

  return result;
  
  // HELPER para sacar los overFlowStats
  // se tiene que meter aqui porque usa variables de dentro de la funcion y para no tener que meterlo en cada llamada
  function computeOverflowStats_(dayUsedBlocksSnapshot) {
    const dayMinutes = {};
    const overflow = {};
    let sumOverflow = 0;
    let sumAbsOverflow = 0;   
    let countDays = 0;

    dayKeys.forEach(dayKey => {
      const usedBlocks = dayUsedBlocksSnapshot[dayKey] || 0;
      const minutesAssigned = usedBlocks * 30;
      dayMinutes[dayKey] = minutesAssigned;
      const baseFree = freeMinutesByDay[dayKey] || 0;
      const ov = minutesAssigned - baseFree;
      overflow[dayKey] = ov;
      sumOverflow += ov;
      sumAbsOverflow += Math.abs(ov);  
      countDays++;
    });

    const meanOverflow = countDays > 0 ? (sumOverflow / countDays) : 0;

    let F = 0;
    dayKeys.forEach(dayKey => {
      const ov = overflow[dayKey];
      const diff = ov - meanOverflow;
      F += diff * diff;
    });

    return {
      dayMinutes: dayMinutes,
      overflow: overflow,
      meanOverflow: meanOverflow,
      F: F,
      sumAbsOverflow: sumAbsOverflow
    };
  }

  // helper para ver si ese dia es usable para ese día
  function isDayUsableForEvent_(meta, dayKey) {
  // Si no hay deadline (o no es una fecha válida), no restringimos
  if (!meta.deadline) return true;
  const deadline = new Date(meta.deadline);
  if (isNaN(deadline.getTime())) return true;

  // Solo nos importa el mismo día del deadline; otros días se tratan por fecha en allowedDays
  const deadlineKey = dateKeyFromDate_(deadline);
  if (deadlineKey !== dayKey) return true;

  const deadlineTime = deadline.getTime();

  // Buscamos si existe al menos un freeSlot en ese día que empiece antes del deadline
  for (let i = 0; i < freeSlots.length; i++) {
    const slot = freeSlots[i];
    const s = new Date(slot.start);
    const e = new Date(slot.end);
    if (dateKeyFromDate_(s) !== dayKey) continue;
    // Cualquier hueco que empiece antes del deadline se considera usable;
    // el ajuste fino de minutos ya lo hará scheduleTasksIntoFreeSlots_
    if (s.getTime() < deadlineTime && e.getTime() > s.getTime()) {
      return true;
    }
  }
  // No hay huecos reales antes del deadline en ese día → no usarlo en el precalendario
  return false;
}

}
