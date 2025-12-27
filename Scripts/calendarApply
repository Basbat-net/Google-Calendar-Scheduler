// REVISADO
function getSlotsIndicesOrderedByLunchProximity_(dayKey, workingSlots) {
  const dayDate = dateFromKey_(dayKey);
  const lunchMid = new Date(dayDate.getFullYear(),dayDate.getMonth(),dayDate.getDate(),14, 30, 0, 0);

  const items = [];

  for (let i = 0; i < workingSlots.length; i++) {
    const slot = workingSlots[i];
    if (!slot || !slot.start || !slot.end) continue;

    const slotStart = slot.start;
    const slotEnd = slot.end;
    if (!(slotStart instanceof Date) || !(slotEnd instanceof Date)) continue;

    const slotDayKey = dateKeyFromDate_(slotStart);
    if (slotDayKey !== dayKey) continue;

    // distancia mínima a la comida (si el slot está antes, miramos el final;
    // si está después, miramos el inicio; si cruzara, distancia 0)
    let distMs;
    if (slotEnd <= lunchMid) {
      distMs = lunchMid.getTime() - slotEnd.getTime();
    } else if (slotStart >= lunchMid) {
      distMs = slotStart.getTime() - lunchMid.getTime();
    } else {
      distMs = 0;
    }

    items.push({
      index: i,
      distMs: distMs,
      startTime: slotStart.getTime()
    });
  }

  items.sort((a, b) => {
    if (a.distMs !== b.distMs) {
      return a.distMs - b.distMs; // más cerca de la comida primero
    }
    // empate de distancia: preferimos el que empieza más tarde
    return b.startTime - a.startTime;
  });

  return items.map(it => it.index);
}

// REVISADO

function scheduleTasksIntoFreeSlots_(freeSlots, eventsByBracket) {
  Logger.log("lo que le entra de freeSlots");
  Logger.log(JSON.stringify(freeSlots, null, 0));

  // 1- esREPARAMOS EL PRECALENDARIO CON optimizeScheduleWithTowers_
  const towerResult = optimizeScheduleWithTowers_(freeSlots, eventsByBracket);
  if (!towerResult) {
    Logger.log("No se ha podido construir un precalendario; saliendo.");
    return;
  }
  const {eventDayBlocks, eventsMeta, extraTimes} = towerResult;

  // copiamos freeSlots para ir recortandolo
  const workingSlots = freeSlots.map(slot => ({
    start: new Date(slot.start),
    end: new Date(slot.end)
  }));

  // Preparamos los datos para la creación de los eventos
  const rawNow = getNow_();
  const now = roundUpToNextHalfHour_(rawNow);
  const segmentsByKey = {};
  // Ordenamos los eventos por urgencia
  const eventsMetaOrdered = [...eventsMeta].sort((a, b) => {
    const aIsTask = (a.type === "Tareas") ? 1 : 0;
    const bIsTask = (b.type === "Tareas") ? 1 : 0;
    if (aIsTask !== bIsTask) return bIsTask - aIsTask; // Tareas primero

    const pa = proximityOfLastScheduledMomentToDeadlineMs_(a, eventDayBlocks);
    const pb = proximityOfLastScheduledMomentToDeadlineMs_(b, eventDayBlocks);

    if (pa !== pb) return pa - pb;

    // Desempates razonables para que no “parezca que no ordena”
    const da = a.deadline instanceof Date ? a.deadline.getTime() : 0;
    const db = b.deadline instanceof Date ? b.deadline.getTime() : 0;
    if (da !== db) return da - db;

    // último desempate estable: id
    return String(a.id || "").localeCompare(String(b.id || ""));
  });

  function proximityOfLastScheduledMomentToDeadlineMs_(meta, eventDayBlocks) {
    if (!meta || !(meta.deadline instanceof Date)) return Number.POSITIVE_INFINITY;

    const id = meta.id;
    const blocksPerDay = eventDayBlocks && eventDayBlocks[id];
    if (!blocksPerDay) return Number.POSITIVE_INFINITY;

    // 1) último día con bloques
    let lastDayKey = null;
    Object.keys(blocksPerDay).forEach(dk => {
      if (dk === "concepto") return;
      const n = Number(blocksPerDay[dk] || 0);
      if (n > 0 && (!lastDayKey || dk > lastDayKey)) lastDayKey = dk;
    });
    if (!lastDayKey) return Number.POSITIVE_INFINITY;

    // 2) instante "último" de ese día (final de jornada usable)
    const lastMoment = getEndOfUsableDayForKey_(lastDayKey);

    // 3) proximidad: cuanto más cerca del deadline, más prioritario.
    //    - Si lastMoment cae después del deadline, lo tratamos como 0 (máxima urgencia),
    //      porque significa que ese evento “empuja” contra el deadline.
    const diff = meta.deadline.getTime() - lastMoment.getTime();
    return diff <= 0 ? 0 : diff;
  }

  // Usa exactamente tu lógica de intervalos (entre semana vs findes)
  function getEndOfUsableDayForKey_(dayKey) {
    const d = dateFromKey_(dayKey);
    const dow = d.getDay(); // 0 domingo, 6 sábado

    const intervals = (dow === 0 || dow === 6)
      ? INTERVALOS_USABLES_FINDES
      : INTERVALOS_USABLES_ENTRE_SEMANA;

    if (!intervals || intervals.length === 0) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
    }

    const last = intervals[intervals.length - 1];
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), last.endHour, last.endMinute, 0, 0);
  }

  Logger.log("EventsMetaOrdered");
  Logger.log(eventsMetaOrdered);

  // IMPORTANTE: Iteramos los eventos y en funcion de los freeSlots los colocamos, no miramos los slots para ver qué metemos
  eventsMetaOrdered.forEach(meta => {
    
    // Datos del evento
    const ev = meta.originalEvent;
    const id = meta.id;
    if (id == "E#0") Logger.log("lo puedo probar");
    const type = meta.type;
    const deadline = meta.deadline;
    const isTask = (type === "Tareas");
    const isShortTask = isTask && ev && ev.etaMinutesRemaining <= 60;
    const blocksPerDay = eventDayBlocks[id] || {};
    const totalBlocks = Object.values(blocksPerDay).reduce((acc, v) => acc + v, 0);
    if (totalBlocks <= 0) return;
    const targetCal = type ? (calendarByType[type] || (calendarByType[type] = fetchCalendar_(type))) : null;
    if (!targetCal) {Logger.log("No se encontró calendario para type='" + type + "' al programar '" + (meta.concepto || "") + "'"); return;}
    const title = meta.concepto || ("Tarea " + (ev.rowIndex != null ? ev.rowIndex : ""));
    const dayKeysForEvent = Object.keys(blocksPerDay)
      .filter(k => k !== "concepto")
      .sort();

    const segments = []; // aqui es donde se van a ir guardando en qué segmentos ponemos este evento en especifico

    // Iteramos todos los dias en donde el precalendario nos ha dicho que podemos colocarlos
    // Nota para el futuro: Este for loop es un poco una guarrada, y posiblemente se pueda hacer más 
    // limpio, pero ahora mismo prefiero hacer otras cosas antes que hacerle un refactor
    for (let d = 0; d < dayKeysForEvent.length; d++) {
      const dayKey = dayKeysForEvent[d];
      // el precalendario nos dice cuantos bloques por dia, sabiendo que esos bloques se pueden colocar
      // el donde no lo indica, asi que se decide aquí debajo
      let blocksForDay = blocksPerDay[dayKey]; 
      if (!blocksForDay || blocksForDay <= 0) continue;
      const dayDate = dateFromKey_(dayKey);
      const lunchMid = new Date(dayDate.getFullYear(),dayDate.getMonth(),dayDate.getDate(),14, 30, 0, 0);

      // las tareas y el estudio se hacen de maneras diferentes, asi que hay que tratarlos con diferente logica
      if (type === "Tareas") {
        if (id == "E#0") Logger.log("está entrando en tareas");
        // usamos los slots de ese día ordenados por cercanía a la comida
        const slotIndices = getSlotsIndicesOrderedByLunchProximity_(dayKey, workingSlots);
        if (id == "E#0") Logger.log("slot indices: " + slotIndices);
        for (let idx = 0; idx < slotIndices.length && blocksForDay > 0; idx++) {
          // Datos del slot
          const i = slotIndices[idx];
          const slot = workingSlots[i];
          if (!slot || !slot.start || !slot.end) continue;
          let slotStart = slot.start;
          let slotEnd = slot.end;
          if (!(slotStart instanceof Date) || !(slotEnd instanceof Date)) continue;
          if (slotEnd <= now) continue;
          const slotDayKey = dateKeyFromDate_(slotStart);
          if (slotDayKey !== dayKey) continue;

          if (id == "E#0") Logger.log("blocksForDay:" + blocksForDay);
          while (blocksForDay > 0) {
            if (slotEnd <= slotStart) break;
            const isMorning = slotEnd <= lunchMid; // si es por la mañana se trata diferente

            // Para tareas <= 1h, se pueden partir en cosas de media hora, si no, solo de hora en hora
            const minChunkMinutes = isShortTask ? 30 : 60;
            const blocksPerChunk = isShortTask ? 1 : 2;
            const chunkMillis = blocksPerChunk * (30 * 60 * 1000);

            if (id == "E#0") Logger.log("huuuuh");
            // Para la mañana, empieza a rellenar desde lo mas cerca a la comida
            if (isMorning) {
              if (id == "E#0") Logger.log("entra en mañana");
              //Datos del evento
              const availableMinutes = (slotEnd.getTime() - slotStart.getTime()) / (60 * 1000);
              if (id == "E#0") Logger.log("Pasa 1");
              if (availableMinutes < minChunkMinutes) break;
              const eventEnd = new Date(slotEnd.getTime());
              const eventStart = new Date(eventEnd.getTime() - chunkMillis);
              if (eventEnd <= now) break;
              if (id == "E#24") Logger.log("Pasa 2");
              if (eventStart < slotStart) break;
              if (id == "E#0") Logger.log("Pasa 3");
              if (id == "E#24") Logger.log(eventEnd);
              if (id == "E#24") Logger.log(deadline);              
              if (eventEnd > deadline)    break;
              if (id == "E#0") Logger.log("Pasa 4");
              // Colocamos una sección del evento en este segmento en especifico
              segments.push({ start: new Date(eventStart), end: new Date(eventEnd) });

              //Modificamos los datos del slot y del dia, dado que ya están usado, para que no se coloquen dos veces
              slotEnd = eventStart;
              slot.end = slotEnd;
              blocksForDay -= blocksPerChunk;

            } 

            // para la tarde, lo rellenamos en descendente desde la hora de la comida
            else {
              if (id == "E#0") Logger.log("entra en tarde");
              // Datos del evento
              const effectiveStart = new Date(Math.max(slotStart.getTime(), now.getTime()));
              const availableMinutes = (slotEnd.getTime() - effectiveStart.getTime()) / (60 * 1000);
              if (availableMinutes < minChunkMinutes) break;
              const eventStart = effectiveStart;
              const eventEnd = new Date(eventStart.getTime() + chunkMillis);
              if (id == "E#24") Logger.log(eventEnd);
              if (id == "E#24") Logger.log(deadline);    
              if (eventEnd > deadline) break;
      
              // Colocamos el evento
              segments.push({ start: new Date(eventStart), end: new Date(eventEnd) });

              // Modificamos los datos del slot y del dia
              slotStart = eventEnd;
              slot.start = slotStart;
              blocksForDay -= blocksPerChunk;
            }
          }
        }
        continue; // siguiente día o evento
      }

      // Parte para el estudio
      const slotIndices = getSlotsIndicesOrderedByLunchProximity_(dayKey, workingSlots); // Los slots, no los eventos

      for (let idx = 0; idx < slotIndices.length && blocksForDay >= 2; idx++) {
        if (id == "E#0") Logger.log("alocacion de la tarea como tal");
        // Datos del slot
        const i = slotIndices[idx];
        const slot = workingSlots[i];
        if (!slot || !slot.start || !slot.end) continue;
        let slotStart = slot.start;
        let slotEnd = slot.end;
        if (!(slotStart instanceof Date) || !(slotEnd instanceof Date)) continue;
        if (slotEnd <= now) continue;
        const slotDayKey = dateKeyFromDate_(slotStart);
        if (slotDayKey !== dayKey) continue;

        // Importante, los blocksForDay no pueden ser menores a 2 porque no podemos alocar media hora de estudio
        while (blocksForDay >= 2) {
          // Vemos cuanto espacio tenemos todavía
          if (id == "E#24") Logger.log("comprobaciones para id 24");
          if (id == "E#24") Logger.log(slotStart);
          if (id == "E#24") Logger.log(slotEnd);
          if (id == "E#24") Logger.log(deadline);    
            //if (eventEnd > deadline) break;
          if (slotEnd <= slotStart) break;
          let availableMinutes = (slotEnd.getTime() - slotStart.getTime()) / (60 * 1000);
          if (availableMinutes < 60) break; // menos de 1h → no abrimos cluster
          
          // Vemos cuantos slots podemos meter de una
          let capBlocks = Math.floor(availableMinutes / 30);
          let clusterBlocks = Math.min(capBlocks, blocksForDay);
          if (clusterBlocks < 2) break; // Nos hemos quedado sin espacio

          if (blocksForDay - clusterBlocks === 1) {
            if (clusterBlocks >= 3) {
              clusterBlocks -= 1; // dejamos 2 bloques para otro cluster
            } else {
              break;
            }
          }
          if (clusterBlocks < 2) break;
          const durationMs = clusterBlocks * (30 * 60 * 1000);
          const isMorning = slotEnd <= lunchMid;

          let clusterStart, clusterEnd;

          // Hacemos lo mismo que en tasks, si es mañana desde la comida si no desde la tarde
          if (isMorning) {
            // Datos para ver donde meter el chunk
            clusterEnd = new Date(Math.min(slotEnd.getTime(), deadline.getTime()));
            clusterStart = new Date(clusterEnd.getTime() - durationMs);
            if (clusterEnd <= now) break;
            if (clusterStart < slotStart) break;

            // Probamos a ver si lo podemos meter
            let maxBlocksByDeadline = Math.floor((deadline.getTime() - clusterStart.getTime()) / (30 * 60 * 1000));

            if (maxBlocksByDeadline < 2) break;
            if (clusterBlocks > maxBlocksByDeadline) {
              clusterBlocks = maxBlocksByDeadline;
              if (clusterBlocks < 2) break;
              const newDuration = clusterBlocks * (30 * 60 * 1000);
              clusterStart = new Date(clusterEnd.getTime() - newDuration);
            }
            // Si sale, se mete a segments
            segments.push({ start: new Date(clusterStart), end: new Date(clusterEnd) });

            // Actualizamos los datos
            blocksForDay -= clusterBlocks;
            slotEnd = clusterStart;
            slot.end = slotEnd;
          } 
          else {
            // Tarde
            // Datos para el bloque
            const effectiveStart = new Date(Math.max(slotStart.getTime(), now.getTime()));
            availableMinutes = (slotEnd.getTime() - effectiveStart.getTime()) / (60 * 1000);
            capBlocks = Math.floor(availableMinutes / 30);
            if (capBlocks < 2) break;
            clusterBlocks = Math.min(clusterBlocks, capBlocks);
            if (clusterBlocks < 2) break;
            clusterStart = effectiveStart;
            clusterEnd = new Date(clusterStart.getTime() + clusterBlocks * (30 * 60 * 1000));

            //Probamos si cabe
            let maxBlocksByDeadline = Math.floor((deadline.getTime() - clusterStart.getTime()) / (30 * 60 * 1000));
            if (id == "E#24") Logger.log("max blocks by deadline");
            if (id == "E#24") Logger.log(maxBlocksByDeadline);
            if (id == "E#24") Logger.log("event start");
            if (id == "E#24") Logger.log(clusterStart);

            if (id == "E#24") Logger.log("event end");
            if (id == "E#24") Logger.log(clusterEnd);

            if (maxBlocksByDeadline < 2) break;
            if (clusterBlocks > maxBlocksByDeadline) {
              clusterBlocks = maxBlocksByDeadline;
              if (clusterBlocks < 2) break;
              clusterEnd = new Date(clusterStart.getTime() + clusterBlocks * (30 * 60 * 1000));
            }

            // lo tiramos a segments
            segments.push({ start: new Date(clusterStart), end: new Date(clusterEnd) });

            // actualizamos cosas 
            blocksForDay -= clusterBlocks;
            slotStart = clusterEnd;
            slot.start = slotStart;
          }
        }
      }
    }

    if (segments.length > 0) {


      // Aquí es realmente donde se meten los segmentos dentro del segementsByKey que es la que luego pasará al calendar
      const key = type + '||' + title;
      if (!segmentsByKey[key])  segmentsByKey[key] = { cal: targetCal, title: title, segments: [] };
      segmentsByKey[key].segments.push(...segments);
    } 
    else {
      Logger.log("No se han podido crear segmentos reales en el calendario para '" + (meta.concepto || id) + "' con id: " + id + "y deadline:" + meta.deadline+" pese a tener bloques en el precalendario.");        
      Logger.log(segments.length);}
    
  });


  // FASE EXTRA DE HORAS DE ESTUDIO
  // Una vez que hemos colocado los estudios y las tareas, queremos llenar más horas con estudio 
  // en caso de que nos sobre sitio
  Logger.log("FASE EXTRA DE REAGRUPACIÓN DE ESTUDIO:");

  const horizon = getDynamicHorizonEnd_(now);

  // Sacamos los slots que nos sobran
  const freeAfterNormal = workingSlots
    .filter(s => s && s.start instanceof Date && s.end instanceof Date &&
      s.end > now && s.start < horizon && s.end > s.start)
    .map(s => ({
      start: new Date(Math.max(s.start.getTime(), now.getTime())),
      end:   new Date(Math.min(s.end.getTime(),   horizon.getTime()))
    }));

  // Lo convertimos a un diccionario
  const freeMinutesByDay = {};
  freeAfterNormal.forEach(slot => {
    const key = dateKeyFromDate_(slot.start);
    const minutes = (slot.end - slot.start) / 60000;
    freeMinutesByDay[key] = (freeMinutesByDay[key] || 0) + minutes;
  });

  // Filtramos solo los eventos de estudio importantes
  const priority5Events = eventsMeta.filter(e => e.type === "Estudio" && e.bracket === 5);

  if (priority5Events.length === 0) {
    Logger.log("FASE EXTRA: No hay eventos de estudio p5.");
  } 
  else {
    Logger.log("FASE EXTRA: Hay " + priority5Events.length + " eventos p5.");

    // sacamos los examenes que se pueden estudiar cada día
    const examsByDay = {};
    priority5Events.forEach(ev => {
      ev.allowedDays.forEach(dayKey => {
        if (!examsByDay[dayKey]) examsByDay[dayKey] = [];
        examsByDay[dayKey].push(ev);
      });
    });


    // iteramos cada día en el que tenemos evento de estudio prio 5
    // En los días donde hay prioridad 5, se va a intentar alocar el mismo numero total de horas extras por asignatura,
    // puede que un dia tenga más pero la semana se va a reparti de maenra
    Object.keys(examsByDay).forEach(dayKey => {
      // Lista de examenes para ese día
      const exList = examsByDay[dayKey];
      if (!freeMinutesByDay[dayKey] || freeMinutesByDay[dayKey] < 60) return;

      Logger.log("FASE EXTRA: Día " + dayKey + " tiene " + freeMinutesByDay[dayKey] + " min libres.");

      // Si solo hay un examen ese día, vamos a llenarlo todo con el evento de estudio
      if (exList.length === 1) {
        const ev = exList[0];
        while (placeOneHourExtra_(ev, dayKey)) {
          freeMinutesByDay[dayKey] -= 60;
        }
        return;
      }

      // Si llega aquí es que hay más de uno ese día

      const load = {};
      exList.forEach(ev => load[ev.id] = 0);

      while (freeMinutesByDay[dayKey] >= 60) {
        // Escogemos el examen con MENOS carga extra.
        // En empate, desempatamos por nombre para no favorecer siempre al primero.
        let best = null;
        let bestLoad = Infinity;

        exList.forEach(ev => {
          const l = load[ev.id] || 0; // probamos la load de cada examen extra
          if (l < bestLoad - 1e-6) {
            best = ev;
            bestLoad = l; // si es el de menos lo metemos aqui
          } else if (Math.abs(l - bestLoad) < 1e-6 && best) {
            // si son iguales desempata por nombre (para que el 1o no sea siempre el peor)
            const nameBest = String(best.concepto || "");
            const nameCur  = String(ev.concepto || "");
            if (nameCur.localeCompare(nameBest) < 0) {
              best = ev;
              bestLoad = l;
            }
          }
        });

        if (!best) break; // no se puede colocar ninguna

        const placed = placeOneHourExtra_(best, dayKey); // se mete el evento en el mejor punto
        if (!placed) break;

        // modificamos los datos para la siguiente iteracion
        freeMinutesByDay[dayKey] -= 60;
        load[best.id] = (load[best.id] || 0) + 60;
      }

    });
  }

  // FASE EXTRA: PLANIFICACIÓN DEL TIEMPO EXTRA AL FINAL DEL DÍA
  // Estas horas solo existen en días con overflow > 0, así que por definición
  // tienen que ir más allá de los límites normales de ese día.
  // En vez de encajarlas a la fuerza en el mismo día del overflow,
  // intentamos recolocarlas en cualquier día (hasta el deadline del evento)
  // donde se puedan FUSIONAR con un evento consecutivo del mismo tipo,
  // para evitar cortes raros entre asignaturas.

  if (extraTimes && Object.keys(extraTimes).length > 0) {
    const metaById = {};
    eventsMeta.forEach(m => { metaById[m.id] = m; });

    // Fin de jornada usable para un dayKey dado (usa los intervalos globales)
    function getEndOfUsableDayForKey_(dayKey) {
      const d = dateFromKey_(dayKey);   // 2025-12-09 → Date local ese día
      const dow = d.getDay();           // 0 = domingo, 6 = sábado

      const intervals = (dow === 0 || dow === 6)
        ? INTERVALOS_USABLES_FINDES
        : INTERVALOS_USABLES_ENTRE_SEMANA;

      // Si por lo que sea no hay intervalos, tomamos 23:59 como "final de día".
      if (!intervals || intervals.length === 0) {
        return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 0, 0);
      }

      const last = intervals[intervals.length - 1];
      return new Date(
        d.getFullYear(),
        d.getMonth(),
        d.getDate(),
        last.endHour,
        last.endMinute,
        0,
        0
      );
    }

    // Último segmento (y su key) ya colocado en ese día, si existe
    function getLastSegmentInfoForDay_(dayKey) {
      let lastSeg = null;
      let lastKey = null;

      Object.keys(segmentsByKey).forEach(k => {
        const entry = segmentsByKey[k];
        if (!entry || !entry.segments) return;
        entry.segments.forEach(seg => {
          if (!seg.start || !seg.end) return;
          if (dateKeyFromDate_(seg.start) !== dayKey) return;
          if (!lastSeg || seg.end.getTime() > lastSeg.end.getTime()) {
            lastSeg = seg;
            lastKey = k;
          }
        });
      });

      if (!lastSeg) return null;
      return {
        key: lastKey,
        end: new Date(lastSeg.end.getTime())
      };
    }

    // Elige el mejor día para colocar UN bloque extra de este evento:
    //  - Preferimos días donde el último segmento del día sea del mismo key (fusión perfecta).
    //  - Después, días que al menos tengan algún evento.
    //  - Si no hay nada, usamos un día de fallback (por ejemplo, uno de los días originales en extraTimes).
    function pickBestDayForExtraBlock_(keyExtra, metaExtra, fallbackDays) {
      const allowedDays = (metaExtra.allowedDays || []).slice();
      const deadline = metaExtra.deadline;
      const deadlineKey = deadline ? dateKeyFromDate_(deadline) : null;

      // Conjunto de candidatos: allowedDays ∩ (≤ deadline) ∪ fallbackDays
      const candidateSet = {};
      allowedDays.forEach(dk => {
        if (deadlineKey && dk > deadlineKey) return;
        candidateSet[dk] = true;
      });
      (fallbackDays || []).forEach(dk => {
        if (dk === "concepto") return;
        candidateSet[dk] = true;
      });

      const candidates = Object.keys(candidateSet).sort();
      if (candidates.length === 0) return null;

      let bestDay = null;
      let bestScore = -1;

      candidates.forEach(dk => {
        const info = getLastSegmentInfoForDay_(dk);
        let score = 0;
        if (info) {
          // 2 → el último evento del día ya es de este mismo key (fusión perfecta).
          // 1 → el día tiene eventos, pero de otra cosa.
          score = (info.key === keyExtra) ? 2 : 1;
        } else {
          // 0 → día vacío; solo lo usamos si no hay nada mejor.
          score = 0;
        }

        if (
          bestDay === null ||
          score > bestScore ||
          (score === bestScore && dk < bestDay) // desempate por día más cercano (alfabético YYYY-MM-DD)
        ) {
          bestDay = dk;
          bestScore = score;
        }
      });

      if (!bestDay) return null;

      const lastInfo = getLastSegmentInfoForDay_(bestDay);
      const baseStart = lastInfo
        ? new Date(lastInfo.end.getTime())
        : getEndOfUsableDayForKey_(bestDay);

      return { dayKey: bestDay, baseStart: baseStart };
    }

    Object.keys(extraTimes).forEach(evId => {
      const metaExtra = metaById[evId];
      if (!metaExtra) return;

      const titleExtra = metaExtra.concepto || ("Tarea " + (metaExtra.rowIndex != null ? metaExtra.rowIndex : ""));
      const typeExtra  = metaExtra.type;
      const calExtra   = calendarByType[typeExtra] || fetchCalendar_(typeExtra);
      if (!calExtra) return;

      const dayMap = extraTimes[evId];

      const keyExtra = typeExtra + "||" + titleExtra;
      if (!segmentsByKey[keyExtra]) {
        segmentsByKey[keyExtra] = { cal: calExtra, title: titleExtra, segments: [] };
      }

      // Total de bloques extra para este evento (sumando todos los días origen)
      let totalBlocksExtra = 0;
      const fallbackDays = [];
      Object.keys(dayMap).forEach(dayKey => {
        if (dayKey === "concepto") return;
        const blocks = dayMap[dayKey] || 0;
        if (blocks > 0) {
          totalBlocksExtra += blocks;
          fallbackDays.push(dayKey);
        }
      });
      if (totalBlocksExtra <= 0) return;

      // Vamos colocando cada bloque individualmente, eligiendo el mejor día cada vez
      for (let b = 0; b < totalBlocksExtra; b++) {
        const choice = pickBestDayForExtraBlock_(keyExtra, metaExtra, fallbackDays);
        if (!choice) break;

        const baseStart = choice.baseStart;

        // bloque de 30'
        const segStart = new Date(baseStart.getTime());
        const segEnd   = new Date(segStart.getTime() + 30 * 60 * 1000);

        segmentsByKey[keyExtra].segments.push({
          start: segStart,
          end:   segEnd
        });

        // Actualizamos internamente para que el siguiente bloque vea el nuevo "final de día"
        // (getLastSegmentInfoForDay_ leerá ya este nuevo segmento).
      }
    });
  }



  // REORDENADO Y AGRUPACIÓN DE EVENTOS DE ESTUDIO
  // Después de las horas extra, algunos eventos se te quedan huerfanos asi que hay que permutar y unificar algunos
  // para que no se te queden horas A B A B 
  // REORDENADO Y AGRUPACIÓN DE EVENTOS DE ESTUDIO
  // Después de las horas extra, algunos eventos se te quedan huerfanos asi que hay que permutar y unificar algunos
  // para que no se te queden horas A B A B 
  Logger.log("FASE FINAL: reordenar y agrupar segmentos por día.");

  // --- NUEVO: Mapa deadline por key REAL (mismo key que segmentsByKey usa) ---
  // IMPORTANTE: debe usar EXACTAMENTE el mismo "title" que usas al crear segmentsByKey.
  const deadlineByKey = {};
  (eventsMeta || []).forEach(meta => {
    const ev = meta ? meta.originalEvent : null;
    const type = meta ? meta.type : null;
    if (!type) return;

    const title = meta.concepto || ("Tarea " + (ev && ev.rowIndex != null ? ev.rowIndex : ""));
    const key = type + "||" + title;

    if (meta.deadline instanceof Date) {
      // si hay duplicados (mismo key), nos quedamos con el deadline más temprano por seguridad
      if (!deadlineByKey[key] || meta.deadline.getTime() < deadlineByKey[key].getTime()) {
        deadlineByKey[key] = new Date(meta.deadline.getTime());
      }
    }
  });

  // pasamos a segmentsByday para poder ordenarlos
  const segmentsByDay = {};
  Object.keys(segmentsByKey).forEach(key => {
    const entry = segmentsByKey[key];
    entry.segments.forEach(seg => {
      const dayKey = dateKeyFromDate_(seg.start);
      if (!segmentsByDay[dayKey]) segmentsByDay[dayKey] = [];
      segmentsByDay[dayKey].push({
        key: key,
        cal: entry.cal,
        title: entry.title,
        start: new Date(seg.start),
        end: new Date(seg.end)
      });
    });
  });

  // --- DEBUG (ANTES) SOLO para Estudio||tmm en 2026-01-05 ---
  (function debugBefore_() {
    const dk = "2026-01-05";
    const list = segmentsByDay[dk] || [];
    const tmm = list.filter(s => s.key === "Estudio||tmm");
    if (tmm.length > 0) {
      Logger.log("DEBUG BEFORE (day " + dk + ") Estudio||tmm:");
      tmm.forEach((s, i) => Logger.log(i + ") " + s.start + " -> " + s.end + " | dl=" + deadlineByKey[s.key]));
    }
  })();

  // 2) Para cada día:
  //    - Detectamos "brackets" de estudio consecutivos (sin huecos).
  //    - Dentro de cada bracket reordenamos los segmentos Estudio para agrupar
  //      por asignatura (A B A B → A A B B).
  //    - Después fusionamos segmentos consecutivos del mismo key.
  // iteramos cada día a ver si podemos simplificar algo
  Object.keys(segmentsByDay).forEach(dayKey => {
    const list = segmentsByDay[dayKey];
    if (!list || list.length === 0) return;

    // ordenamos los eventos por hora
    list.sort((a, b) => a.start.getTime() - b.start.getTime());

    // --- 2.1 Reordenar brackets consecutivos SOLO de Estudio ---
    let i = 0;
    while (i < list.length) {
      let j = i + 1;
      // bracket [i, j) donde cada evento empieza justo cuando termina el anterior
      while (j < list.length && list[j].start.getTime() === list[j - 1].end.getTime()) {
        j++;
      }

      // intenta buscar eventos consecutivos que pueden estar ordenados o no
      const size = j - i;
      if (size >= 2) {
        // si hay 2 o más es posible ordenar
        const bracketSegs = list.slice(i, j); // cortamos solo lo que nos interesa

        // chequeamos que todos los eventos sean de estudio (las tasks van a partir igual)
        const allStudy = bracketSegs.every(seg => seg.key.startsWith("Estudio||"));
        if (allStudy) {
          const bracketStart = new Date(bracketSegs[0].start);

          // --- FIX: SIMULACIÓN sin mutar los objetos reales ---
          // Creamos "plan items" con duración y referencia al original.
          const plan = bracketSegs.map(orig => ({
            orig: orig,
            key: orig.key,
            durMs: orig.end.getTime() - orig.start.getTime()
          }));

          // agrupamos las keys iguales para luego fusionarlos
          plan.sort((a, b) => {
            if (a.key < b.key) return -1;
            if (a.key > b.key) return 1;
            // fallback estable: por duración no hace falta, mantenemos orden original implícito
            return 0;
          });

          // Asignamos tiempos en un "schedule" temporal
          let cursor = new Date(bracketStart);
          const tempSchedule = plan.map(p => {
            const s = new Date(cursor);
            const e = new Date(cursor.getTime() + p.durMs);
            cursor = new Date(e);
            return { key: p.key, start: s, end: e, orig: p.orig };
          });

          // --- Validación de deadlines tras mover (ahora sí, con deadlines correctos) ---
          // Si algún segmento acaba después del deadline de su key, NO reordenamos el bracket.
          const violates = tempSchedule.some(seg => {
            const dl = deadlineByKey[seg.key];
            if (!(dl instanceof Date)) return false; // si no sabemos deadline, no bloqueamos
            return seg.end.getTime() > dl.getTime();
          });

          if (!violates) {
            // Aplicamos el schedule a los objetos reales SOLO si no viola
            tempSchedule.forEach(ts => {
              ts.orig.start = new Date(ts.start);
              ts.orig.end = new Date(ts.end);
            });
          } else {
            // Si viola, dejamos el bracket como estaba (no tocamos)
            // Logger.log("Bracket NO reordenado en " + dayKey + " por violación de deadline.");
          }
        }
      }

      i = j;
    }

    // una vez que los tenemos ya, los reordenamos por tiempo
    list.sort((a, b) => a.start.getTime() - b.start.getTime());

    // los segmentos están ordenados, pero aun no se han fusionadp
    const merged = [];
    for (let idx = 0; idx < list.length; idx++) {
      // informacion del segmento
      const seg = list[idx];
      if (merged.length === 0) {
        merged.push(seg);
        continue;
      }
      const last = merged[merged.length - 1];

      // si este y el anterior son iguales nos unimos
      if (last.key === seg.key && last.end.getTime() === seg.start.getTime()) {
        // incluso al fusionar, respetar deadline
        const dl = deadlineByKey[last.key];
        if (dl instanceof Date && seg.end.getTime() > dl.getTime()) {
          merged.push(seg); // no fusionamos si alargaría más allá del deadline
        } else {
          last.end = new Date(seg.end);
        }
      } else {
        merged.push(seg);
      }
    }

    segmentsByDay[dayKey] = merged;
  });

  // --- DEBUG (DESPUÉS) SOLO para Estudio||tmm en 2026-01-05 ---
  (function debugAfter_() {
    const dk = "2026-01-05";
    const list = segmentsByDay[dk] || [];
    const tmm = list.filter(s => s.key === "Estudio||tmm");
    if (tmm.length > 0) {
      Logger.log("DEBUG AFTER (day " + dk + ") Estudio||tmm:");
      tmm.forEach((s, i) => Logger.log(i + ") " + s.start + " -> " + s.end + " | dl=" + deadlineByKey[s.key]));
    }
  })();

  
  // Una vez que acaba el for loop, se usa createMergedEventsForSegments_ y se planta en el calendar
  Object.keys(segmentsByKey).forEach(key => {
    const entry = segmentsByKey[key];
    createMergedEventsForSegments_(entry.cal, entry.title, entry.segments);
  });
  
  Logger.log("scheduleTasksIntoFreeSlots_: planificación final completada.");

  // Helper para colocar 1 hora extra en un día determinado (agrupación)
  function placeOneHourExtra_(evMeta, dayKey) {
    // Datos del examen
    const examDeadline = evMeta.deadline;
    const examDayKey   = dateKeyFromDate_(examDeadline);
    const isExamDay    = (examDayKey === dayKey);

    // Datos del día
    const dayDate  = dateFromKey_(dayKey);
    const lunchMid = new Date(dayDate.getFullYear(), dayDate.getMonth(), dayDate.getDate(), 14, 30, 0, 0);

    // Slots que sobran ese día
    const slotIndices = getSlotsIndicesOrderedByLunchProximity_(dayKey, freeAfterNormal);
    if (!slotIndices || slotIndices.length === 0) return false;

    for (let idx = 0; idx < slotIndices.length; idx++) {
      // Datos del slot
      const i = slotIndices[idx];
      const slot = freeAfterNormal[i];
      if (!slot || !slot.start || !slot.end) continue;
      let slotStart = new Date(slot.start);
      let slotEnd   = new Date(slot.end);

      // Si es el mismo día del examen, no usar nada que empiece después del examen
      if (isExamDay) {
        if (slotStart >= examDeadline) continue;
        if (slotEnd > examDeadline) slotEnd = new Date(examDeadline);
      }

      // Por seguridad, ignorar slots degenerados
      if (slotEnd <= slotStart) continue;

      // Datos de lo que podemos usar del slot
      let availableMinutes = (slotEnd.getTime() - slotStart.getTime()) / 60000;
      if (availableMinutes < 60) continue;
      const isMorning = slotEnd <= lunchMid;

      let start, end;

      // Si es en morning, empezamos desde el final (desde la comida y restamos)
      if (isMorning) {
        end   = new Date(slotEnd.getTime());
        start = new Date(end.getTime() - 60 * 60000);

      } 
      // si es por la tarde, pillamos el punto y sumamos
      else {
        start = new Date(slotStart.getTime());
        end   = new Date(start.getTime() + 60 * 60000);
        if (end > slotEnd) continue;
      }

      // IMPORTANTE: Que no ponga fechas en el mismo día despues del examen
      if (end > examDeadline) {
        if (!isExamDay) continue;
        continue;
      }

      // actualizamos el restante de eventos que luego irán al calendar
      const title = evMeta.concepto;
      const cal = calendarByType["Estudio"] || fetchCalendar_("Estudio");
      const key = "Estudio||" + title;
      if (!segmentsByKey[key])   segmentsByKey[key] = { cal: cal, title: title, segments: [] };
      segmentsByKey[key].segments.push({ start: new Date(start), end: new Date(end) });

      // actualizamos el slot que hemos modificado
      if (isMorning) slot.end = start;
      else slot.start = end;
      return true;
    }
    return false;
    }
}
