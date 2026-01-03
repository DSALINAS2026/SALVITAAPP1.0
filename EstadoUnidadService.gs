// ========= CONFIG =========
const SPREADSHEET_ID = "1Lyeb-ht-g41QMHJlgNhBg1dmxamlaieWOjCf6yZPsmg";

const SHEET_USERS = "Usuarios";
const SHEET_CHASIS = "ChasisBD";
const SHEET_MOV = "MovimientoUnidad";

// Columnas esperadas en ChasisBD (si faltan, se agregan)
const CHASIS_COLS_REQUIRED = [
  "IdChasis",
  "Sociedad",
  "Deposito",
  "Interno",
  "Dominio",
  "KmRecorridos",
  "Tipo",
  "CapacidadCarga",
  "Nro. Chasis",
  "Marca",
  "Modelo",
  "Motor",
  "Nro Motor",
  "Eje",
  "Mapa Cubierta",
  "Carroceria",
  "Año",
  "Estado",
  "Val.",
  "Usuario",
  "Fecha",
  "Ultima Modificacion"
];

// Columnas esperadas en MovimientoUnidad (si faltan, se agregan)
const MOV_COLS_REQUIRED = [
  "IdMov",
  "Unidad",
  "Tipo",            // ingreso/egreso
  "FechaMov",        // Date
  "HoraMov",         // hh:mm
  "Odómetro",
  "UltimoOdometro",
  "KmRecorridos",
  "Observacion",
  "Deposito",
  "Usuario",
  "Timestamp"
];

// ========= WEB =========
function doGet() {
  const t = HtmlService.createTemplateFromFile("Index");
  return t.evaluate()
    .setTitle("Mantenimiento - Web")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========= HELPERS =========
let __SS;
function _ss() {
  if (!__SS) __SS = SpreadsheetApp.openById(SPREADSHEET_ID);
  return __SS;
}

const __SHEETS = {};
function _sheet(name) {
  if (Array.isArray(name)) return _sheetAny(name);
  if (__SHEETS[name]) return __SHEETS[name];
  const sh = _ss().getSheetByName(name);
  if (!sh) throw new Error("No existe la pestaña: " + name);
  __SHEETS[name] = sh;
  return sh;
}

function _sheetAny(names){
  const list = (names||[]).filter(Boolean).map(String);
  for (const n of list){
    if (__SHEETS[n]) return __SHEETS[n];
    const sh = _ss().getSheetByName(n);
    if (sh){ __SHEETS[n]=sh; return sh; }
  }
  throw new Error("No existe ninguna pestaña: " + list.join(" / "));
}

function _getOrCreateSheet(name){
  if (Array.isArray(name)) name = name[0];
  let sh = _ss().getSheetByName(name);
  if (!sh) sh = _ss().insertSheet(name);
  __SHEETS[name] = sh;
  return sh;
}

function _nowMs(){ return Date.now(); }

