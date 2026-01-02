// ===================== OT CREATE (CORRECTIVA) =====================
// IMPORTANTE: este archivo NO declara SHEET_OT / SHEET_OT_TAREAS para no chocar con otros módulos.
// Usa nombres locales seguros y, si ya existen constantes globales, las toma.

const SHEET_OTC_OT       = (typeof SHEET_OT !== "undefined") ? SHEET_OT : "OrdenesTrabajo";
const SHEET_OTC_DET      = (typeof SHEET_OT_TAREAS !== "undefined") ? SHEET_OT_TAREAS : "OT_Tareas";
const SHEET_OTC_CHASIS   = (typeof SHEET_CHASIS !== "undefined") ? SHEET_CHASIS : "ChasisBD";
const SHEET_OTC_TAREAS   = (typeof SHEET_TAREAS !== "undefined") ? SHEET_TAREAS : "MaestroTareas";

const OTC_COLS_REQUIRED = [
  "IdOT","NroOT","TipoOT","EstadoOT","Fecha","Interno","Dominio","Sociedad","Deposito","Sector",
  "Solicita","Descripcion","Usuario","Timestamp"
];

const OTC_DET_COLS_REQUIRED = [
  "IdDetalle","IdOT","CodigoTarea","NombreTarea","Sistema","Subsistema","Sector",
  "EstadoTarea","Check","Usuario","Timestamp"
];


// ---------- helpers ----------
function _otc_headers_(sh){
  const last = sh.getLastColumn();
  if (last === 0) return [];
  return sh.getRange(1,1,1,last).getValues()[0].map(h => (h||"").toString().trim());
}

function _otc_ensureSchema_(sheetName, required){
  const sh = _sheet(sheetName);
  const headers = _otc_headers_(sh);

  if (headers.length === 0){
    sh.getRange(1,1,1,required.length).setValues([required]);
    return;
  }

  const missing = required.filter(c => !headers.includes(c));
  if (missing.length){
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function _otc_ensureAll_(){
  _otc_ensureSchema_(SHEET_OTC_OT, OTC_COLS_REQUIRED);
  _otc_ensureSchema_(SHEET_OTC_DET, OTC_DET_COLS_REQUIRED);
}

function _otc_norm_(s){ return (s ?? "").toString().trim().toLowerCase(); }
function _otc_isTrue_(v){
  const s = (v ?? "").toString().trim().toLowerCase();
  return v === true || s === "true" || s === "1" || s === "si" || s === "sí";
}

function _otc_uuid8_(){
  return Utilities.getUuid().slice(0,8);
}

function _otc_nextNroOT_(){
  const sh = _sheet(SHEET_OTC_OT);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return 1;

  const h = v[0].map(x => (x||"").toString().trim());
  const iNro = h.indexOf("NroOT");
  if (iNro === -1) return 1;

  let max = 0;
  for (let r=1; r<v.length; r++){
    const raw = v[r][iNro];
    const n = Number(String(raw ?? "").replace(/[^\d]/g,""));
    if (isFinite(n) && n > max) max = n;
  }
  return max + 1;
}

// ===================== API: opciones para modal Crear OT =====================
// Devuelve unidades + tareas + listas para filtros
function getOTCreateOptions(token){
  _requireSession_(token);
  _otc_ensureAll_();

  // helper: índice de columna ignorando mayúsc/minúsc y espacios
  const idxCI = (headers, name) => {
    const n = (name||"").toString().trim().toLowerCase();
    return headers.findIndex(h => (h||"").toString().trim().toLowerCase() === n);
  };

  // ====== UNIDADES desde ChasisBD ======
  const ch = _sheet(SHEET_OTC_CHASIS);
  const cv = ch.getDataRange().getValues();

  let unidades = [];
  if (cv.length >= 2){
    const h = cv[0].map(x => (x||"").toString().trim());
    const iInt = idxCI(h,"Interno");
    const iDom = idxCI(h,"Dominio");
    const iSoc = idxCI(h,"Sociedad");
    const iDep = idxCI(h,"Deposito");
    const iEst = idxCI(h,"Estado");

    unidades = cv.slice(1).map(r=>{
      const Interno = (iInt===-1? "" : (r[iInt]??"").toString().trim());
      if (!Interno) return null;

      const Dominio = (iDom===-1? "" : (r[iDom]??"").toString().trim());
      const Sociedad = (iSoc===-1? "" : (r[iSoc]??"").toString().trim());
      const Deposito = (iDep===-1? "" : (r[iDep]??"").toString().trim());
      const Estado = (iEst===-1? "" : (r[iEst]??"").toString().trim());

      // si estado indica baja/inactivo, lo ocultamos (simple)
      const est = _otc_norm_(Estado);
      if (est.includes("baja") || est.includes("inact")) return null;

      return {
        value: Interno,
        label: `${Interno}${Dominio ? " — " + Dominio : ""}${Sociedad ? " — " + Sociedad : ""}`,
        Dominio, Sociedad, Deposito
      };
    }).filter(Boolean);

    unidades.sort((a,b)=> a.value.localeCompare(b.value,"es"));
  }

  // ====== TAREAS desde MaestroTareas (no borradas) ======
  const mt = _sheet(SHEET_OTC_TAREAS);
  const mv = mt.getDataRange().getValues();

  let tareas = [];
  let sectores = [];

  if (mv.length >= 2){
    const h = mv[0].map(x => (x||"").toString().trim());
    const iCod = idxCI(h,"codigo");
    const iNom = idxCI(h,"nombre");
    const iSis = idxCI(h,"sistema");
    const iSub = idxCI(h,"subsistema");
    const iSec = idxCI(h,"sector");
    const iBor = idxCI(h,"borrado");

    tareas = mv.slice(1).map(r=>{
      const borr = (iBor===-1? "" : (r[iBor]??""));
      if (_otc_isTrue_(borr)) return null;

      const codigo = (iCod===-1? "" : (r[iCod]??"").toString().trim());
      const nombre = (iNom===-1? "" : (r[iNom]??"").toString().trim());
      if (!codigo && !nombre) return null;

      const sistema = (iSis===-1? "" : (r[iSis]??"").toString().trim());
      const subsistema = (iSub===-1? "" : (r[iSub]??"").toString().trim());
      const sector = (iSec===-1? "" : (r[iSec]??"").toString().trim());

      if (sector) sectores.push(sector);

      return { codigo, nombre, sistema, subsistema, sector };
    }).filter(Boolean);

    const seen = new Set();
    sectores = sectores.filter(s => seen.has(s) ? false : (seen.add(s), true));
    sectores.sort((a,b)=>a.localeCompare(b,"es"));

    tareas.sort((a,b)=> (a.codigo||"").localeCompare((b.codigo||""),"es"));
  }

  return { ok:true, unidades, tareas, sectores };
}

// ===================== API: Crear OT Correctiva =====================
// payload: { Interno, Sector, Solicita, Descripcion, tareas:[{codigo,nombre,sistema,subsistema,sector}] }
function addOTCorrectiva(token, payload){
  const user = _requireSession_(token);
  _otc_ensureAll_();

  const Interno = (payload?.Interno ?? "").toString().trim();
  const Sector = (payload?.Sector ?? "").toString().trim();
  const Solicita = (payload?.Solicita ?? "").toString().trim();
  const Descripcion = (payload?.Descripcion ?? "").toString().trim();

  if (!Interno) throw new Error("Interno es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!Solicita) throw new Error("Quién solicita es obligatorio.");

  // buscar datos unidad (dominio/sociedad/deposito)
  const ch = _sheet(SHEET_OTC_CHASIS);
  const cv = ch.getDataRange().getValues();

  const idxCI = (headers, name) => {
    const n = (name||"").toString().trim().toLowerCase();
    return headers.findIndex(h => (h||"").toString().trim().toLowerCase() === n);
  };

  let Dominio="", Sociedad="", Deposito="";
  if (cv.length >= 2){
    const hh = cv[0].map(x => (x||"").toString().trim());
    const iInt = idxCI(hh,"Interno");
    const iDom = idxCI(hh,"Dominio");
    const iSoc = idxCI(hh,"Sociedad");
    const iDep = idxCI(hh,"Deposito");

    const row = cv.slice(1).find(r => (iInt===-1?"":(r[iInt]??"").toString().trim()) === Interno);
    if (row){
      Dominio  = (iDom===-1? "" : (row[iDom]??"").toString().trim());
      Sociedad = (iSoc===-1? "" : (row[iSoc]??"").toString().trim());
      Deposito = (iDep===-1? "" : (row[iDep]??"").toString().trim());
    }
  }

  // OT head
  const IdOT = _otc_uuid8_();
  const NroOT = _otc_nextNroOT_();
  const now = new Date();

  const sh = _sheet(SHEET_OTC_OT);
  const headers = _otc_headers_(sh);
  const row = new Array(headers.length).fill("");

  const setIf = (col, val) => {
    const i = headers.indexOf(col);
    if (i !== -1) row[i] = val;
  };

  setIf("IdOT", IdOT);
  setIf("NroOT", NroOT);
  setIf("TipoOT", "Correctiva");
  setIf("EstadoOT", "pendiente");
  setIf("Fecha", now);
  setIf("Interno", Interno);
  setIf("Dominio", Dominio);
  setIf("Sociedad", Sociedad);
  setIf("Deposito", Deposito);
  setIf("Sector", Sector);
  setIf("Solicita", Solicita);
  setIf("Descripcion", Descripcion);
  setIf("Usuario", user.u || "");
  setIf("Timestamp", now);

  sh.appendRow(row);

  // detalles
  const shD = _sheet(SHEET_OTC_DET);
  const dH = _otc_headers_(shD);

  const setD = (arr, col, val) => {
    const i = dH.indexOf(col);
    if (i !== -1) arr[i] = val;
  };

  const tareas = (payload?.tareas || []).map(t=>({
    codigo: (t.codigo ?? "").toString().trim(),
    nombre: (t.nombre ?? "").toString().trim(),
    sistema: (t.sistema ?? "").toString().trim(),
    subsistema: (t.subsistema ?? "").toString().trim(),
    sector: (t.sector ?? "").toString().trim()
  })).filter(t => t.codigo || t.nombre);

  if (!tareas.length) throw new Error("Tenés que seleccionar al menos 1 tarea.");

  tareas.forEach(t=>{
    const rr = new Array(dH.length).fill("");
    setD(rr, "IdDetalle", _otc_uuid8_());
    setD(rr, "IdOT", IdOT);
    setD(rr, "CodigoTarea", (t.codigo ?? "").toString().trim());
    setD(rr, "NombreTarea", (t.nombre ?? "").toString().trim());
    setD(rr, "Sistema", (t.sistema ?? "").toString().trim());
    setD(rr, "Subsistema", (t.subsistema ?? "").toString().trim());
    setD(rr, "Sector", (t.sector ?? "").toString().trim() || Sector);
    setD(rr, "EstadoTarea", "pendiente");
    setD(rr, "Check", true); // por defecto tildada
    setD(rr, "Usuario", user.u || "");
    setD(rr, "Timestamp", new Date());

    shD.appendRow(rr);
  });

  return { ok:true, IdOT, NroOT };
}
// ===================== OT CREATE (PREVENTIVA) =====================
// Crea una OT con tareas precargadas desde HojasPreventivas + HojasPreventivasTareas
const SHEET_OTC_HP        = (typeof SHEET_HP !== "undefined") ? SHEET_HP : "HojasPreventivas";
const SHEET_OTC_HP_TAREAS = (typeof SHEET_HP_TAREAS !== "undefined") ? SHEET_HP_TAREAS : "HojasPreventivasTareas";

// Opciones para modal Crear OT Preventiva (unidades + hojas activas)
function getOTPreventivaOptions(token){
  _requireSession_(token);
  _otc_ensureAll_();

  // si existe, asegura esquema preventivos
  try{ if (typeof ensurePreventivosSheets_ === "function") ensurePreventivosSheets_(); }catch(e){}

  const idxCI = (headers, name) => {
    const n = (name||"").toString().trim().toLowerCase();
    return headers.findIndex(h => (h||"").toString().trim().toLowerCase() === n);
  };

  // ====== UNIDADES desde ChasisBD ======
  const ch = _sheet(SHEET_OTC_CHASIS);
  const cv = ch.getDataRange().getValues();
  let unidades = [];

  if (cv.length >= 2){
    const h = cv[0].map(x => (x||"").toString().trim());
    const iInt = idxCI(h,"Interno");
    const iDom = idxCI(h,"Dominio");
    const iSoc = idxCI(h,"Sociedad");
    const iDep = idxCI(h,"Deposito");
    const iEst = idxCI(h,"Estado");

    unidades = cv.slice(1).map(r=>{
      const Interno = (iInt===-1? "" : (r[iInt]??"").toString().trim());
      if (!Interno) return null;

      const Dominio = (iDom===-1? "" : (r[iDom]??"").toString().trim());
      const Sociedad = (iSoc===-1? "" : (r[iSoc]??"").toString().trim());
      const Deposito = (iDep===-1? "" : (r[iDep]??"").toString().trim());
      const Estado = (iEst===-1? "" : (r[iEst]??"").toString().trim());

      // si estado indica baja/inactivo, lo ocultamos (simple)
      const est = _otc_norm_(Estado);
      if (est.includes("baja") || est.includes("inact")) return null;

      return {
        value: Interno,
        label: `${Interno}${Dominio ? " — " + Dominio : ""}${Sociedad ? " — " + Sociedad : ""}`,
        Dominio, Sociedad, Deposito
      };
    }).filter(Boolean);

    unidades.sort((a,b)=> a.value.localeCompare(b.value,"es"));
  }

  // ====== HOJAS PREVENTIVAS activas ======
  const sh = _sheet(SHEET_OTC_HP);
  const v = sh.getDataRange().getValues();
  let hojas = [];
  if (v.length >= 2){
    const h = v[0].map(x => (x||"").toString().trim());
    const iId = idxCI(h,"IdHP");
    const iNom = idxCI(h,"NombreHP");
    const iSec = idxCI(h,"Sector");
    const iEst = idxCI(h,"Estado");

    hojas = v.slice(1).map(r=>{
      const IdHP = (iId===-1? "" : (r[iId]??"").toString().trim());
      const NombreHP = (iNom===-1? "" : (r[iNom]??"").toString().trim());
      const Sector = (iSec===-1? "" : (r[iSec]??"").toString().trim());
      const Estado = (iEst===-1? "" : (r[iEst]??"").toString().trim());

      if (!IdHP && !NombreHP) return null;

      const est = _otc_norm_(Estado);
      if (est && (est.includes("baja") || est.includes("inact") || est.includes("anul"))) return null;

      return { IdHP: IdHP || NombreHP, NombreHP: NombreHP || IdHP, Sector };
    }).filter(Boolean);

    hojas.sort((a,b)=> (a.NombreHP||"").localeCompare((b.NombreHP||""),"es"));
  }

  return { ok:true, unidades, hojas };
}
// Crear OT Preventiva.
// payload: { Interno, IdHP, Sector, Solicita, Descripcion, tareas:[{codigo,nombre,sistema,subsistema,sector,check}] }
function addOTPreventiva(token, payload){
  const user = _requireSession_(token);
  _otc_ensureAll_();
  try{ if (typeof ensurePreventivosSheets_ === "function") ensurePreventivosSheets_(); }catch(e){}

  const Interno = (payload?.Interno ?? "").toString().trim();
  const IdHP = (payload?.IdHP ?? "").toString().trim();
  const Solicita = (payload?.Solicita ?? "").toString().trim();
  let Descripcion = (payload?.Descripcion ?? "").toString().trim();

  if (!Interno) throw new Error("Interno es obligatorio.");
  if (!IdHP) throw new Error("Hoja preventiva es obligatoria.");
  if (!Solicita) throw new Error("Quién solicita es obligatorio.");

  // buscar datos unidad (dominio/sociedad/deposito)
  const ch = _sheet(SHEET_OTC_CHASIS);
  const cv = ch.getDataRange().getValues();
  const idxCI = (headers, name) => {
    const n = (name||"").toString().trim().toLowerCase();
    return headers.findIndex(h => (h||"").toString().trim().toLowerCase() === n);
  };

  let Dominio="", Sociedad="", Deposito="";
  if (cv.length >= 2){
    const hh = cv[0].map(x => (x||"").toString().trim());
    const iInt = idxCI(hh,"Interno");
    const iDom = idxCI(hh,"Dominio");
    const iSoc = idxCI(hh,"Sociedad");
    const iDep = idxCI(hh,"Deposito");

    const row = cv.slice(1).find(r => (iInt===-1?"":(r[iInt]??"").toString().trim()) === Interno);
    if (row){
      Dominio  = (iDom===-1? "" : (row[iDom]??"").toString().trim());
      Sociedad = (iSoc===-1? "" : (row[iSoc]??"").toString().trim());
      Deposito = (iDep===-1? "" : (row[iDep]??"").toString().trim());
    }
  }

  // leer cabecera HP para sector + nombre
  const shHP = _sheet(SHEET_OTC_HP);
  const hv = shHP.getDataRange().getValues();
  if (hv.length < 2) throw new Error("No hay HojasPreventivas cargadas.");

  const hh = hv[0].map(x => (x||"").toString().trim());
  const iId = idxCI(hh,"IdHP");
  const iNom = idxCI(hh,"NombreHP");
  const iSec = idxCI(hh,"Sector");

  const hpRow = hv.slice(1).find(r=>{
    const id = (iId===-1? "" : (r[iId]??"").toString().trim());
    const nom = (iNom===-1? "" : (r[iNom]??"").toString().trim());
    return id === IdHP || nom === IdHP;
  });

  const HPNombre = hpRow ? (iNom===-1? "" : (hpRow[iNom]??"").toString().trim()) : "";
  const HPSector = hpRow ? (iSec===-1? "" : (hpRow[iSec]??"").toString().trim()) : "";

  const Sector = (payload?.Sector ?? "").toString().trim() || HPSector || "";

  if (!Descripcion && HPNombre) Descripcion = "Preventivo: " + HPNombre;
  if (HPNombre && Descripcion && !Descripcion.toLowerCase().includes(HPNombre.toLowerCase())){
    Descripcion = `${Descripcion} (HP: ${HPNombre})`;
  }

  // OT head
  const IdOT = _otc_uuid8_();
  const NroOT = _otc_nextNroOT_();
  const now = new Date();

  const shOT = _sheet(SHEET_OTC_OT);
  const otH = _otc_headers_(shOT);
  const row = new Array(otH.length).fill("");

  const setIf = (col, val) => {
    const i = otH.indexOf(col);
    if (i !== -1) row[i] = val;
  };

  setIf("IdOT", IdOT);
  setIf("NroOT", NroOT);
  setIf("TipoOT", "Preventiva");
  setIf("EstadoOT", "pendiente");
  setIf("Fecha", now);
  setIf("Interno", Interno);
  setIf("Dominio", Dominio);
  setIf("Sociedad", Sociedad);
  setIf("Deposito", Deposito);
  setIf("Sector", Sector);
  setIf("Solicita", Solicita);
  setIf("Descripcion", Descripcion);
  setIf("Usuario", user.u || "");
  setIf("Timestamp", now);

  shOT.appendRow(row);

  // detalles
  const tareas = (payload?.tareas || []).map(t=>({
    codigo: (t.codigo ?? "").toString().trim(),
    nombre: (t.nombre ?? "").toString().trim(),
    sistema: (t.sistema ?? "").toString().trim(),
    subsistema: (t.subsistema ?? "").toString().trim(),
    sector: (t.sector ?? "").toString().trim() || Sector,
    check: !!t.check
  })).filter(t => t.codigo || t.nombre);

  if (!tareas.length) throw new Error("La hoja preventiva no tiene tareas.");

  const shD = _sheet(SHEET_OTC_DET);
  const dH = _otc_headers_(shD);

  const setD = (arr, col, val) => {
    const i = dH.indexOf(col);
    if (i !== -1) arr[i] = val;
  };

  tareas.forEach(t=>{
    const rr = new Array(dH.length).fill("");
    setD(rr, "IdDetalle", _otc_uuid8_());
    setD(rr, "IdOT", IdOT);
    setD(rr, "CodigoTarea", t.codigo);
    setD(rr, "NombreTarea", t.nombre);
    setD(rr, "Sistema", t.sistema);
    setD(rr, "Subsistema", t.subsistema);
    setD(rr, "Sector", t.sector);
    setD(rr, "EstadoTarea", t.check ? "pendiente" : "anulada");
    setD(rr, "Check", t.check);
    setD(rr, "Usuario", user.u || "");
    setD(rr, "Timestamp", new Date());
    shD.appendRow(rr);
  });

  return { ok:true, IdOT, NroOT };
}
// (fin)

