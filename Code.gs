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
function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _sheet(name) {
  const sh = _ss().getSheetByName(name);
  if (!sh) throw new Error("No existe la pestaña: " + name);
  return sh;
}
