// ===================== PREVENTIVOS (HOJAS PREVENTIVAS) =====================

// NUEVAS HOJAS
const SHEET_HP = "HojasPreventivas";
const SHEET_HP_TAREAS = "HojasPreventivasTareas";

// Columnas requeridas
const HP_COLS_REQUIRED = [
  "IdHP", "NombreHP", "Sector", "Descripcion", "Estado", "CreadoPor", "CreadoEl"
];

const HP_T_COLS_REQUIRED = [
  "IdHP", "CodigoTarea", "NombreTarea", "Sistema", "Subsistema", "Sector",
  "CadaKm", "CadaDias", "Orden"
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
  // ID simple y estable
  return "HP-" + Utilities.getUuid().slice(0,8).toUpperCase();
}

function _nowISO_(){
  return new Date().toISOString();
}

function _getUserName_(){
  // Usa sesión si existe; fallback email
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

function getHPFiltersOptions(token){
  // Si querés proteger por login, podés validar token acá (como en tus otros services)
  ensurePreventivosSheets_();

  // Sectores los sacamos desde MaestroTareas (ya lo tenés)
  // Reutilizamos tu fuente: MaestroTareasService suele tener getTareasSelectOptions
  // Pero como no estamos dentro del front, lo reconstruimos leyendo tabla de tareas.
  // Si tu tabla MaestroTareas no se llama así en sheet, avisame y lo ajusto.
  const SH_TAREAS = "MaestroTareas"; // coincide con tu service actual
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

function searchHojasPreventivas(token, q){
  ensurePreventivosSheets_();

  const t = _readTable_(SHEET_HP);
  const rows = (t.rows || []).map(r => ({
    _row: r._row,
    IdHP: String(r.IdHP||"").trim(),
    NombreHP: String(r.NombreHP||"").trim(),
    Sector: String(r.Sector||"").trim(),
    Descripcion: String(r.Descripcion||"").trim(),
    Estado: String(r.Estado||"").trim(),
    CreadoPor: String(r.CreadoPor||"").trim(),
    CreadoEl: String(r.CreadoEl||"").trim()
  }));

  const nombre = String(q?.nombre||"").toLowerCase().trim();
  const sector = String(q?.sector||"").trim();
  const estado = String(q?.estado||"").trim();

  // NO LISTAR TODO:
  // solo devolvemos si hay búsqueda o filtro aplicado
  const hasQuery = !!nombre || !!sector || !!estado;
  if (!hasQuery){
    return { ok:true, rows: [] };
  }

  const out = rows.filter(r=>{
    const okN = !nombre || r.NombreHP.toLowerCase().includes(nombre);
    const okS = !sector || r.Sector === sector;
    const okE = !estado || r.Estado === estado;
    return okN && okS && okE;
  });

  return { ok:true, rows: out };
}

function getHPCreateOptions(token){
  // Devuelve tareas del Maestro + sectores para selects
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

function addHojaPreventiva(token, payload){
  ensurePreventivosSheets_();

  const NombreHP = String(payload?.NombreHP||"").trim();
  const Sector = String(payload?.Sector||"").trim();
  const Descripcion = String(payload?.Descripcion||"").trim();
  const Estado = String(payload?.Estado||"activo").trim().toLowerCase();

  const tareas = Array.isArray(payload?.tareas) ? payload.tareas : [];

  if (!NombreHP) throw new Error("Nombre de hoja es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!tareas.length) throw new Error("Debés seleccionar al menos 1 tarea.");

  // Validar frecuencias
  tareas.forEach((t,i)=>{
    const km = String(t.CadaKm||"").trim();
    const di = String(t.CadaDias||"").trim();
    if (!km && !di) throw new Error(`Falta frecuencia en tarea #${i+1} (CadaKm o CadaDias).`);
  });

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
  row.Estado = Estado || "activo";
  row.CreadoPor = creadoPor;
  row.CreadoEl = creadoEl;

  sh.appendRow(h.map(k => row[k]));

  // Guardar tareas
  const st = _sheet(SHEET_HP_TAREAS);
  const ht = st.getRange(1,1,1,st.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());

  tareas.forEach((t,idx)=>{
    const r = {};
    ht.forEach(k => r[k] = "");
    r.IdHP = IdHP;
    r.CodigoTarea = String(t.CodigoTarea||"").trim();
    r.NombreTarea = String(t.NombreTarea||"").trim();
    r.Sistema = String(t.Sistema||"").trim();
    r.Subsistema = String(t.Subsistema||"").trim();
    r.Sector = String(t.Sector||"").trim();
    r.CadaKm = String(t.CadaKm||"").trim();
    r.CadaDias = String(t.CadaDias||"").trim();
    r.Orden = idx+1;

    st.appendRow(ht.map(k => r[k]));
  });

  return { ok:true, IdHP, NombreHP };
}
