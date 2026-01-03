// ===================== PREVENTIVOS (HOJAS PREVENTIVAS) =====================

// NUEVAS HOJAS
const PSV_SHEET_HP = "HojasPreventivas";
const PSV_SHEET_HP_TAREAS = "HojasPreventivasTareas";

// Columnas requeridas (✅ ahora la frecuencia es de la CABECERA)
const PSV_HP_COLS_REQUIRED = [
  "IdHP",   "NombreHP",   "Sector",   "Descripcion",   "CadaKm",   "CadaDias",   "Estado",   "CreadoPor",   "CreadoEl",   "AvisarAntesKm",   "AvisarAntesDias",   "ControlTipo",   "IntervaloKm",   "IntervaloDias",   "AvisoKm",   "AvisoDias",   "Activo"
];

// En tareas, dejamos CadaKm/CadaDias por compatibilidad (si ya existen),
// pero el sistema NUEVO no las usa.
const PSV_HP_T_COLS_REQUIRED = [
  "IdHP", "CodigoTarea", "NombreTarea", "Sistema", "Subsistema", "Sector",
  "CadaKm", "CadaDias",                 // (compat) se ignoran en el nuevo
  "Orden"
];

function ensurePreventivosSheets_(){
  const ss = _ss();

  // HojasPreventivas
  let sh = ss.getSheetByName(PSV_SHEET_HP);
  if (!sh) sh = ss.insertSheet(PSV_SHEET_HP);
  PSV_ensureCols_(sh, PSV_HP_COLS_REQUIRED);

  // HojasPreventivasTareas
  let st = ss.getSheetByName(PSV_SHEET_HP_TAREAS);
  if (!st) st = ss.insertSheet(PSV_SHEET_HP_TAREAS);
  PSV_ensureCols_(st, PSV_HP_T_COLS_REQUIRED);
}

function PSV_ensureCols_(sh, required){
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||"").trim());
  let changed = false;

  required.forEach(col => {
    if (!headers.includes(col)){
      headers.push(col);
      changed = true;
    }
  });

  if (changed){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
}

function PSV_uuid_(){
  return "HP-" + Utilities.getUuid().slice(0,8).toUpperCase();
}

function PSV_nowISO_(){
  return new Date().toISOString();
}

function PSV_getUserName_(){
  try {
    return Session.getActiveUser().getEmail() || "sistema";
  } catch(e){
    return "sistema";
  }
}

function PSV_readTable_(sheetName){
  ensurePreventivosSheets_();
  const sh = _sheet(sheetName);
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return { headers: sh.getRange(1,1,1,lc).getValues()[0], rows: [] };

  const headers = sh.getRange(1,1,1,lc).getValues()[0].map(h => String(h||"").trim());
  const values = sh.getRange(2,1,lr-1,lc).getValues();

  const rows = values.map((r,i) => {
    const obj = { _row: i+2 };
    headers.forEach((h,idx)=> obj[h] = r[idx]);
    return obj;
  });

  return { headers, rows };
}

// =====================================================
// OPTIONS / FILTROS
// =====================================================
function getHPFiltersOptions(token){
  ensurePreventivosSheets_();

  const SH_TAREAS = "MaestroTareas";
  let sectores = [];

  try{
    const t = PSV_readTable_(SH_TAREAS);
    const set = new Set();
    (t.rows||[]).forEach(r=>{
      const s = String(r["sector"] ?? r["Sector"] ?? "").trim();
      if (s) set.add(s);
    });
    sectores = Array.from(set).sort((a,b)=> a.localeCompare(b,"es"));
  }catch(e){
    sectores = [];
  }

  return {
    ok: true,
    sectores,
    estados: ["activo","inactivo"]
  };
}

function getHPCreateOptions(token){
  const SH_TAREAS = "MaestroTareas";
  let tareas = [];
  let sectores = [];

  try{
    const t = PSV_readTable_(SH_TAREAS);
    const set = new Set();

    tareas = (t.rows||[])
      .filter(r => !(r.borrado === true || String(r.borrado).toLowerCase() === "true"))
      .map(r => ({
        codigo: String(r.codigo||"").trim(),
        nombre: String(r.nombre||"").trim(),
        sistema: String(r.sistema||"").trim(),
        subsistema: String(r.subsistema||"").trim(),
        sector: String(r.sector||"").trim()
      }))
      .filter(x => x.codigo && x.nombre);

    tareas.forEach(x => { if (x.sector) set.add(x.sector); });
    sectores = Array.from(set).sort((a,b)=> a.localeCompare(b,"es"));
  }catch(e){
    tareas = [];
    sectores = [];
  }

  return { ok:true, tareas, sectores };
}

// =====================================================
// SEARCH (lista)
// - NO lista todo si no hay filtros/busqueda
// - Devuelve CadaKm/CadaDias de cabecera
// - Fallback: si cabecera vacía, toma la primer tarea (compat)
// =====================================================
function PSV_norm_(s){ return String(s||"").trim(); }

function PSV_getFallbackFrecuenciaFromTasks_(idHP){
  try{
    const tt = PSV_readTable_(PSV_SHEET_HP_TAREAS).rows || [];
    const first = tt.find(r => PSV_norm_(r.IdHP) === PSV_norm_(idHP));
    if (!first) return { CadaKm:"", CadaDias:"" };
    return {
      CadaKm: PSV_norm_(first.CadaKm),
      CadaDias: PSV_norm_(first.CadaDias)
    };
  }catch(e){
    return { CadaKm:"", CadaDias:"" };
  }
}

function searchHojasPreventivas(token, q){
  ensurePreventivosSheets_();

  const t = PSV_readTable_(PSV_SHEET_HP);
  const base = (t.rows || []).map(r => ({
    _row: r._row,
    IdHP: PSV_norm_(r.IdHP),
    NombreHP: PSV_norm_(r.NombreHP),
    Sector: PSV_norm_(r.Sector),
    Descripcion: PSV_norm_(r.Descripcion),
    CadaKm: PSV_norm_(r.CadaKm),
    CadaDias: PSV_norm_(r.CadaDias),
    Estado: PSV_norm_(r.Estado),
    CreadoPor: PSV_norm_(r.CreadoPor),
    CreadoEl: PSV_norm_(r.CreadoEl)
  }));

  const nombre = String(q?.nombre||"").toLowerCase().trim();
  const sector = PSV_norm_(q?.sector);
  const estado = PSV_norm_(q?.estado);

  const hasQuery = !!nombre || !!sector || !!estado;
  if (!hasQuery){
    return { ok:true, rows: [] };
  }

  const out = base
    .filter(r=>{
      const okN = !nombre || r.NombreHP.toLowerCase().includes(nombre);
      const okS = !sector || r.Sector === sector;
      const okE = !estado || r.Estado === estado;
      return okN && okS && okE;
    })
    .map(r=>{
      // ✅ compat: si cabecera no tiene frecuencia, la “toma” de tareas
      if (!r.CadaKm && !r.CadaDias){
        const fb = PSV_getFallbackFrecuenciaFromTasks_(r.IdHP);
        r.CadaKm = r.CadaKm || fb.CadaKm;
        r.CadaDias = r.CadaDias || fb.CadaDias;
      }
      return r;
    });

  return { ok:true, rows: out };
}

// =====================================================
// DETAILS (para ojito)
// =====================================================
function getHPDetails(token, idHP){
  ensurePreventivosSheets_();

  const hp = (PSV_readTable_(PSV_SHEET_HP).rows || []).find(r => PSV_norm_(r.IdHP) === PSV_norm_(idHP));
  if (!hp) throw new Error("No se encontró la hoja preventiva.");

  let head = {
    IdHP: PSV_norm_(hp.IdHP),
    NombreHP: PSV_norm_(hp.NombreHP),
    Sector: PSV_norm_(hp.Sector),
    Descripcion: PSV_norm_(hp.Descripcion),
    CadaKm: PSV_norm_(hp.CadaKm),
    CadaDias: PSV_norm_(hp.CadaDias),
    Estado: PSV_norm_(hp.Estado),
    CreadoPor: PSV_norm_(hp.CreadoPor),
    CreadoEl: PSV_norm_(hp.CreadoEl)
  };

  if (!head.CadaKm && !head.CadaDias){
    const fb = PSV_getFallbackFrecuenciaFromTasks_(head.IdHP);
    head.CadaKm = fb.CadaKm;
    head.CadaDias = fb.CadaDias;
  }

  const tasks = (PSV_readTable_(PSV_SHEET_HP_TAREAS).rows || [])
    .filter(r => PSV_norm_(r.IdHP) === PSV_norm_(idHP))
    .sort((a,b)=> Number(a.Orden||9999) - Number(b.Orden||9999))
    .map(r => ({
      CodigoTarea: PSV_norm_(r.CodigoTarea),
      NombreTarea: PSV_norm_(r.NombreTarea),
      Sistema: PSV_norm_(r.Sistema),
      Subsistema: PSV_norm_(r.Subsistema),
      Sector: PSV_norm_(r.Sector),
      Orden: Number(r.Orden||0)
    }));

  return { ok:true, head, tasks };
}

// =====================================================
// CREATE (✅ frecuencia en CABECERA)
// =====================================================
function addHojaPreventiva(token, payload){
  ensurePreventivosSheets_();

  const NombreHP = PSV_norm_(payload?.NombreHP);
  const Sector = PSV_norm_(payload?.Sector);
  const Descripcion = PSV_norm_(payload?.Descripcion);
  const Estado = PSV_norm_(payload?.Estado || "activo").toLowerCase();

  const CadaKm = PSV_norm_(payload?.CadaKm);
  const CadaDias = PSV_norm_(payload?.CadaDias);

// --- ALERTAS / TIPO DE CONTROL ---
// Regla:
// - Si solo hay Km => ControlTipo = "km" y AvisarAntesKm = 10% (si no viene)
// - Si solo hay Días => ControlTipo = "dia" y AvisarAntesDias = 7 (si no viene)
// - Si hay ambos, por defecto prioriza Km (podés cambiarlo desde UI en el futuro)
const tipo = (CadaKm ? "km" : "dia");

// Valores opcionales desde UI (si más adelante agregás campos)
const AvisarAntesKm_in = PSV_norm_(payload?.AvisarAntesKm);
const AvisarAntesDias_in = PSV_norm_(payload?.AvisarAntesDias);

const CadaKmNum = parseInt(String(CadaKm||"").replace(/[^\d]/g,""), 10);
const CadaDiasNum = parseInt(String(CadaDias||"").replace(/[^\d]/g,""), 10);

const defAvisarKm = (isFinite(CadaKmNum) && CadaKmNum>0) ? Math.max(1, Math.ceil(CadaKmNum * 0.10)) : "";
const defAvisarDias = (isFinite(CadaDiasNum) && CadaDiasNum>0) ? 7 : "";

const AvisarAntesKm = AvisarAntesKm_in ? parseInt(String(AvisarAntesKm_in).replace(/[^\d]/g,""),10) : defAvisarKm;
const AvisarAntesDias = AvisarAntesDias_in ? parseInt(String(AvisarAntesDias_in).replace(/[^\d]/g,""),10) : defAvisarDias;

  const tareas = Array.isArray(payload?.tareas) ? payload.tareas : [];

  if (!NombreHP) throw new Error("Nombre de hoja es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!CadaKm && !CadaDias) throw new Error("Debés completar CadaKm o CadaDias (al menos uno).");
  if (!tareas.length) throw new Error("Debés seleccionar al menos 1 tarea.");

  const IdHP = PSV_uuid_();
  const creadoPor = PSV_getUserName_();
  const creadoEl = PSV_nowISO_();

  // Guardar cabecera
  const sh = _sheet(PSV_SHEET_HP);
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  const row = {};

const _setCol = (variants, value) => {
  const arr = Array.isArray(variants) ? variants : [variants];
  const key = arr.find(k => h.includes(k));
  if (key) row[key] = value;
};
h.forEach(k => row[k] = "");
  row.IdHP = IdHP;
  row.NombreHP = NombreHP;
  row.Sector = Sector;
  row.Descripcion = Descripcion;
  
row.CadaKm = CadaKm;
row.CadaDias = CadaDias;

// Guardamos tipo + alertas (si las columnas existen en tu hoja)
_setCol(["ControlTipo","TipoControl","Tipo"], tipo);
if (tipo === "km") {
  _setCol(["AvisarAntesKm","AvisarAntesKms","AlertaKm","AvisoAntesKm"], AvisarAntesKm);
  _setCol(["AvisarAntesDias","AlertaDias","AvisoAntesDias"], "");
} else {
  _setCol(["AvisarAntesDias","AlertaDias","AvisoAntesDias"], AvisarAntesDias);
  _setCol(["AvisarAntesKm","AvisarAntesKms","AlertaKm","AvisoAntesKm"], "");
}

// Si tenés estas columnas, las dejamos completas para futuras vistas
_setCol(["IntervaloKm"], CadaKm);
_setCol(["IntervaloDias"], CadaDias);
_setCol(["Activo"], "SI");

row.Estado = Estado || "activo";
row.CreadoPor = creadoPor;
row.CreadoEl = creadoEl;

  sh.appendRow(h.map(k => row[k]));

  // Guardar tareas (sin frecuencia)
  const st = _sheet(PSV_SHEET_HP_TAREAS);
  const ht = st.getRange(1,1,1,st.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  tareas.forEach((t,idx)=>{
    const r = {};
    ht.forEach(k => r[k] = "");
    r.IdHP = IdHP;
    r.CodigoTarea = PSV_norm_(t.CodigoTarea);
    r.NombreTarea = PSV_norm_(t.NombreTarea);
    r.Sistema = PSV_norm_(t.Sistema);
    r.Subsistema = PSV_norm_(t.Subsistema);
    r.Sector = PSV_norm_(t.Sector);
    // compat (quedan vacías)
    r.CadaKm = "";
    r.CadaDias = "";
    r.Orden = idx+1;

    st.appendRow(ht.map(k => r[k]));
  });

  return { ok:true, IdHP, NombreHP };
}

// =====================================================
// UPDATE (EDITAR HOJA + REEMPLAZAR TAREAS)
// =====================================================
function updateHojaPreventiva(token, payload){
  ensurePreventivosSheets_();

  const IdHP = String(payload?.IdHP || "").trim();
  const NombreHP = PSV_norm_(payload?.NombreHP);
  const Sector = PSV_norm_(payload?.Sector);
  const Descripcion = PSV_norm_(payload?.Descripcion);
  const Estado = PSV_norm_(payload?.Estado || "activo").toLowerCase();
  const CadaKm = PSV_norm_(payload?.CadaKm);
  const CadaDias = PSV_norm_(payload?.CadaDias);
  const tareas = Array.isArray(payload?.tareas) ? payload.tareas : [];

  if (!IdHP) throw new Error("IdHP inválido.");
  if (!NombreHP) throw new Error("Nombre de hoja es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!CadaKm && !CadaDias) throw new Error("Debés completar CadaKm o CadaDias (al menos uno).");
  if (!tareas.length) throw new Error("Debés seleccionar al menos 1 tarea.");

  // --- 1) Actualizar cabecera ---
  const hpTable = PSV_readTable_(PSV_SHEET_HP);
  const hpRow = (hpTable.rows || []).find(r => PSV_norm_(r.IdHP) === PSV_norm_(IdHP));
  if (!hpRow) throw new Error("No se encontró la hoja preventiva para editar.");

  const sh = _sheet(PSV_SHEET_HP);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  const rowNumber = Number(hpRow._row); // fila real en la hoja
  const rowValues = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];

  const idx = {};
  headers.forEach((h,i)=> idx[h]=i);

  if (idx.NombreHP != null) rowValues[idx.NombreHP] = NombreHP;
  if (idx.Sector != null) rowValues[idx.Sector] = Sector;
  if (idx.Descripcion != null) rowValues[idx.Descripcion] = Descripcion;
  if (idx.CadaKm != null) rowValues[idx.CadaKm] = CadaKm;
  if (idx.CadaDias != null) rowValues[idx.CadaDias] = CadaDias;
  if (idx.Estado != null) rowValues[idx.Estado] = Estado || "activo";

  sh.getRange(rowNumber, 1, 1, headers.length).setValues([rowValues]);

  // --- 2) Reemplazar tareas (borra viejas y agrega nuevas) ---
  const st = _sheet(PSV_SHEET_HP_TAREAS);
  const tTable = PSV_readTable_(PSV_SHEET_HP_TAREAS);
  const tRows = (tTable.rows || []);

  // borrar descendente para no romper índices
  const toDelete = tRows
    .filter(r => PSV_norm_(r.IdHP) === PSV_norm_(IdHP))
    .map(r => Number(r._row))
    .sort((a,b)=> b-a);

  toDelete.forEach(rn => st.deleteRow(rn));

  // volver a leer headers de tareas
  const ht = st.getRange(1,1,1,st.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  tareas.forEach((t, i)=>{
    const r = {};
    ht.forEach(k => r[k] = "");
    r.IdHP = IdHP;
    r.CodigoTarea = PSV_norm_(t.CodigoTarea);
    r.NombreTarea = PSV_norm_(t.NombreTarea);
    r.Sistema = PSV_norm_(t.Sistema);
    r.Subsistema = PSV_norm_(t.Subsistema);
    r.Sector = PSV_norm_(t.Sector) || Sector;
    // compat (se ignoran)
    r.CadaKm = "";
    r.CadaDias = "";
    r.Orden = Number(t.Orden || (i+1));

    st.appendRow(ht.map(k => r[k]));
  });

  return { ok:true };
}


function getHPTareasById(token, idHP){
  ensurePreventivosSheets_();

  idHP = String(idHP || "").trim();
  if (!idHP) throw new Error("IdHP inválido.");

  const t = PSV_readTable_(PSV_SHEET_HP_TAREAS);
  const rows = (t.rows || [])
    .filter(r => String(r.IdHP || "").trim() === idHP)
    .map(r => ({
      CodigoTarea: String(r.CodigoTarea||"").trim(),
      NombreTarea: String(r.NombreTarea||"").trim(),
      Sistema: String(r.Sistema||"").trim(),
      Subsistema: String(r.Subsistema||"").trim(),
      Sector: String(r.Sector||"").trim(),
      Orden: Number(r.Orden || 0)
    }))
    .sort((a,b)=> (a.Orden||0) - (b.Orden||0));

  return { ok:true, rows };
}
