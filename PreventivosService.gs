// ===================== PREVENTIVOS (HOJAS PREVENTIVAS) =====================

// NUEVAS HOJAS
const SHEET_HP = "HojasPreventivas";
const SHEET_HP_TAREAS = "HojasPreventivasTareas";

// Columnas requeridas (✅ ahora la frecuencia es de la CABECERA)
const HP_COLS_REQUIRED = [
  "IdHP", "NombreHP", "Sector", "Descripcion",
  "CadaKm", "CadaDias",                 // ✅ NUEVO
  "Estado", "CreadoPor", "CreadoEl"
];

// En tareas, dejamos CadaKm/CadaDias por compatibilidad (si ya existen),
// pero el sistema NUEVO no las usa.
const HP_T_COLS_REQUIRED = [
  "IdHP", "CodigoTarea", "NombreTarea", "Sistema", "Subsistema", "Sector",
  "CadaKm", "CadaDias",                 // (compat) se ignoran en el nuevo
  "Orden"
];

function ensurePreventivosSheets_(){
  const ss = _ss();

  // HojasPreventivas
  let sh = ss.getSheetByName(SHEET_HP);
  if (!sh) sh = ss.insertSheet(SHEET_HP);
  ensureCols_(sh, HP_COLS_REQUIRED);

  // HojasPreventivasTareas
  let st = ss.getSheetByName(SHEET_HP_TAREAS);
  if (!st) st = ss.insertSheet(SHEET_HP_TAREAS);
  ensureCols_(st, HP_T_COLS_REQUIRED);
}

function ensureCols_(sh, required){
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

function _uuid_(){
  return "HP-" + Utilities.getUuid().slice(0,8).toUpperCase();
}

function _nowISO_(){
  return new Date().toISOString();
}

function _getUserName_(){
  try {
    return Session.getActiveUser().getEmail() || "sistema";
  } catch(e){
    return "sistema";
  }
}

function _readTable_(sheetName){
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
    const t = _readTable_(SH_TAREAS);
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
    const t = _readTable_(SH_TAREAS);
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
function _norm_(s){ return String(s||"").trim(); }

function _getFallbackFrecuenciaFromTasks_(idHP){
  try{
    const tt = _readTable_(SHEET_HP_TAREAS).rows || [];
    const first = tt.find(r => _norm_(r.IdHP) === _norm_(idHP));
    if (!first) return { CadaKm:"", CadaDias:"" };
    return {
      CadaKm: _norm_(first.CadaKm),
      CadaDias: _norm_(first.CadaDias)
    };
  }catch(e){
    return { CadaKm:"", CadaDias:"" };
  }
}

function searchHojasPreventivas(token, q){
  ensurePreventivosSheets_();

  const t = _readTable_(SHEET_HP);
  const base = (t.rows || []).map(r => ({
    _row: r._row,
    IdHP: _norm_(r.IdHP),
    NombreHP: _norm_(r.NombreHP),
    Sector: _norm_(r.Sector),
    Descripcion: _norm_(r.Descripcion),
    CadaKm: _norm_(r.CadaKm),
    CadaDias: _norm_(r.CadaDias),
    Estado: _norm_(r.Estado),
    CreadoPor: _norm_(r.CreadoPor),
    CreadoEl: _norm_(r.CreadoEl)
  }));

  const nombre = String(q?.nombre||"").toLowerCase().trim();
  const sector = _norm_(q?.sector);
  const estado = _norm_(q?.estado);

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
        const fb = _getFallbackFrecuenciaFromTasks_(r.IdHP);
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

  const hp = (_readTable_(SHEET_HP).rows || []).find(r => _norm_(r.IdHP) === _norm_(idHP));
  if (!hp) throw new Error("No se encontró la hoja preventiva.");

  let head = {
    IdHP: _norm_(hp.IdHP),
    NombreHP: _norm_(hp.NombreHP),
    Sector: _norm_(hp.Sector),
    Descripcion: _norm_(hp.Descripcion),
    CadaKm: _norm_(hp.CadaKm),
    CadaDias: _norm_(hp.CadaDias),
    Estado: _norm_(hp.Estado),
    CreadoPor: _norm_(hp.CreadoPor),
    CreadoEl: _norm_(hp.CreadoEl)
  };

  if (!head.CadaKm && !head.CadaDias){
    const fb = _getFallbackFrecuenciaFromTasks_(head.IdHP);
    head.CadaKm = fb.CadaKm;
    head.CadaDias = fb.CadaDias;
  }

  const tasks = (_readTable_(SHEET_HP_TAREAS).rows || [])
    .filter(r => _norm_(r.IdHP) === _norm_(idHP))
    .sort((a,b)=> Number(a.Orden||9999) - Number(b.Orden||9999))
    .map(r => ({
      CodigoTarea: _norm_(r.CodigoTarea),
      NombreTarea: _norm_(r.NombreTarea),
      Sistema: _norm_(r.Sistema),
      Subsistema: _norm_(r.Subsistema),
      Sector: _norm_(r.Sector),
      Orden: Number(r.Orden||0)
    }));

  return { ok:true, head, tasks };
}

// =====================================================
// CREATE (✅ frecuencia en CABECERA)
// =====================================================
function addHojaPreventiva(token, payload){
  ensurePreventivosSheets_();

  const NombreHP = _norm_(payload?.NombreHP);
  const Sector = _norm_(payload?.Sector);
  const Descripcion = _norm_(payload?.Descripcion);
  const Estado = _norm_(payload?.Estado || "activo").toLowerCase();

  const CadaKm = _norm_(payload?.CadaKm);
  const CadaDias = _norm_(payload?.CadaDias);

  const tareas = Array.isArray(payload?.tareas) ? payload.tareas : [];

  if (!NombreHP) throw new Error("Nombre de hoja es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!CadaKm && !CadaDias) throw new Error("Debés completar CadaKm o CadaDias (al menos uno).");
  if (!tareas.length) throw new Error("Debés seleccionar al menos 1 tarea.");

  const IdHP = _uuid_();
  const creadoPor = _getUserName_();
  const creadoEl = _nowISO_();

  // Guardar cabecera
  const sh = _sheet(SHEET_HP);
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  const row = {};
  h.forEach(k => row[k] = "");
  row.IdHP = IdHP;
  row.NombreHP = NombreHP;
  row.Sector = Sector;
  row.Descripcion = Descripcion;
  row.CadaKm = CadaKm;
  row.CadaDias = CadaDias;
  row.Estado = Estado || "activo";
  row.CreadoPor = creadoPor;
  row.CreadoEl = creadoEl;

  sh.appendRow(h.map(k => row[k]));

  // Guardar tareas (sin frecuencia)
  const st = _sheet(SHEET_HP_TAREAS);
  const ht = st.getRange(1,1,1,st.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  tareas.forEach((t,idx)=>{
    const r = {};
    ht.forEach(k => r[k] = "");
    r.IdHP = IdHP;
    r.CodigoTarea = _norm_(t.CodigoTarea);
    r.NombreTarea = _norm_(t.NombreTarea);
    r.Sistema = _norm_(t.Sistema);
    r.Subsistema = _norm_(t.Subsistema);
    r.Sector = _norm_(t.Sector);
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
  const NombreHP = _norm_(payload?.NombreHP);
  const Sector = _norm_(payload?.Sector);
  const Descripcion = _norm_(payload?.Descripcion);
  const Estado = _norm_(payload?.Estado || "activo").toLowerCase();
  const CadaKm = _norm_(payload?.CadaKm);
  const CadaDias = _norm_(payload?.CadaDias);
  const tareas = Array.isArray(payload?.tareas) ? payload.tareas : [];

  if (!IdHP) throw new Error("IdHP inválido.");
  if (!NombreHP) throw new Error("Nombre de hoja es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!CadaKm && !CadaDias) throw new Error("Debés completar CadaKm o CadaDias (al menos uno).");
  if (!tareas.length) throw new Error("Debés seleccionar al menos 1 tarea.");

  // --- 1) Actualizar cabecera ---
  const hpTable = _readTable_(SHEET_HP);
  const hpRow = (hpTable.rows || []).find(r => _norm_(r.IdHP) === _norm_(IdHP));
  if (!hpRow) throw new Error("No se encontró la hoja preventiva para editar.");

  const sh = _sheet(SHEET_HP);
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
  const st = _sheet(SHEET_HP_TAREAS);
  const tTable = _readTable_(SHEET_HP_TAREAS);
  const tRows = (tTable.rows || []);

  // borrar descendente para no romper índices
  const toDelete = tRows
    .filter(r => _norm_(r.IdHP) === _norm_(IdHP))
    .map(r => Number(r._row))
    .sort((a,b)=> b-a);

  toDelete.forEach(rn => st.deleteRow(rn));

  // volver a leer headers de tareas
  const ht = st.getRange(1,1,1,st.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  tareas.forEach((t, i)=>{
    const r = {};
    ht.forEach(k => r[k] = "");
    r.IdHP = IdHP;
    r.CodigoTarea = _norm_(t.CodigoTarea);
    r.NombreTarea = _norm_(t.NombreTarea);
    r.Sistema = _norm_(t.Sistema);
    r.Subsistema = _norm_(t.Subsistema);
    r.Sector = _norm_(t.Sector) || Sector;
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

  const t = _readTable_(SHEET_HP_TAREAS);
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
