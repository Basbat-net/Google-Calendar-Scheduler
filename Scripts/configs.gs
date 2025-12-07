/**<--------------> VARIABLES GENERALES <------------->**/

const TASKS_SHEET_NAME = "Tareas"; // El nombre de la sheet de excel        
const FIRST_TASK_ROW = 2;                 
// Las columnas de la hoja TASKS_SHEET_NAME en numeros
const COL_CONCEPTO = 1;   
const COL_DEADLINE = 2;   
const COL_PRIO = 3;       
const COL_PREP = 4;        
const COL_ETA = 5;        
const COL_PROGRESO = 6;   

const INTERVALOS_USABLES_ENTRE_SEMANA =         
        [ { startHour: 9,  startMinute: 30, endHour: 14, endMinute: 0 },
          { startHour: 15, startMinute: 0,  endHour: 22, endMinute: 0 }];
const INTERVALOS_USABLES_FINDES =         
        [ { startHour: 11, startMinute: 30, endHour: 14, endMinute: 0 },
          { startHour: 15, startMinute: 0,  endHour: 24, endMinute: 0 }];


//Calendarios estáticos para planificar entorno a ellos
// están con numero porque algunos calendarios les he querido meter padding (offset x delante y detrás)
const FIXED_CALENDARS = {"Clases":0,"General":0,"Laboratorio":0,"Laboratorios":0,"Examenes":1}
// IMPORTANTE: Los nombres de los examenes deben ser iguales a los de este array, si no se te va a rallar
const SUBJECTS = [
  "circuitos", "ciencia", "regu", "resis", "termo","analogica" ];


const BASE_STUDY_HOURS = [0,3,5,0,10,15]; // el indice n es el numero de horas para las cosas de prioridad n (prio 1->3h en el ejemplo)

//Calendarios dinámicos donde se hace el planning como tal
const TASK_CALENDAR_NAME = "Tareas";
const STUDY_CALENDAR_NAME = "Estudio";
const PROJECT_CALENDAR_NAME = "Proyectos";
const EXAM_CALENDAR_NAME = "Examenes";

const HORIZON_DAYS = 8;         // Con cuantos días de antelación opera el calendario
const BLOCK_MINUTES = 30;       // bloque minimo de trabajo para organizar

const NOW_OFFSET_HOURS = -0.5; // hay un bug ahora de que empieza en la media hora siguiente a la actual, lo dejo
                               // en -0.5 para corregirlo, pero en teoría debería ser 0

const calendarByType = {};
const now = getNow_(); 
