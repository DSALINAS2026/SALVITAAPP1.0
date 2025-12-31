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
      const interno = (iInt===-1? "" : (r[iInt]??"").toString().trim());
      if (!interno) return null;

      const dominio = (iDom===-1? "" : (r[iDom]??"").toString().trim());
      const sociedad = (iSoc===-1? "" : (r[iSoc]??"").toString().trim());
      const deposito = (iDep===-1? "" : (r[iDep]??"").toString().trim());
      const estado = (iEst===-1? "" : (r[iEst]??"").toString().trim().toLowerCase());

      const label = [interno, dominio, sociedad].filter(Boolean).join(" — ");
      return { value: interno, label, dominio, sociedad, deposito, estado };
    }).filter(Boolean);

    const seen = new Set();
    unidades = unidades.filter(u => seen.has(u.value) ? false : (seen.add(u.value), true));
    unidades.sort((a,b)=>a.label.localeCompare(b.label,"es"));
  }

  // ====== TAREAS desde MaestroTareas (sin borradas) ======
  const mt = _sheet(SHEET_OTC_TAREAS);
  const mv = mt.getDataRange().getValues();

  let tareas = [];
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

      return { codigo, nombre, sistema, subsistema, sector };
    }).filter(Boolean);

    tareas.sort((a,b)=>{
      const aa = (a.codigo||"") + " " + (a.nombre||"");
      const bb = (b.codigo||"") + " " + (b.nombre||"");
      return aa.localeCompare(bb,"es");
    });
  }

  const sistemas = [...new Set(tareas.map(t=>t.sistema).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));
  const subsistemas = [...new Set(tareas.map(t=>t.subsistema).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));
  const sectores = [...new Set(tareas.map(t=>(t.sector||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));

  return { ok:true, unidades, tareas, sistemas, subsistemas, sectores };
}


// ===================== API: crear OT correctiva + detalles =====================
function addOTCorrectiva(token, payload){
  const user = _requireSession_(token);
  _otc_ensureAll_();

  const Interno = (payload?.Interno ?? "").toString().trim();
  const Dominio = (payload?.Dominio ?? "").toString().trim();
  const Sociedad = (payload?.Sociedad ?? "").toString().trim();
  const Deposito = (payload?.Deposito ?? "").toString().trim();
  const Sector = (payload?.Sector ?? "").toString().trim();
  const Solicita = (payload?.Solicita ?? "").toString().trim();
  const Descripcion = (payload?.Descripcion ?? "").toString().trim();

  const tareas = Array.isArray(payload?.tareas) ? payload.tareas : [];

  if (!Interno) throw new Error("Interno es obligatorio.");
  if (!Sector) throw new Error("Sector es obligatorio.");
  if (!Solicita) throw new Error("Quién solicita es obligatorio.");
  if (tareas.length < 1) throw new Error("Seleccioná al menos 1 tarea.");

  const IdOT = "OT-" + _otc_uuid8_();
  const NroOT = String(_otc_nextNroOT_()); // SOLO NÚMEROS
  const now = new Date();

  const sh = _sheet(SHEET_OTC_OT);
  const h = _otc_headers_(sh);
  const idx = (name)=>h.indexOf(name);

  const row = new Array(h.length).fill("");
  const setIf = (col, val) => { const c = idx(col); if (c !== -1) row[c] = val; };

  setIf("IdOT", IdOT);
  setIf("NroOT", NroOT);
  setIf("TipoOT", "correctiva");
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
  const hd = _otc_headers_(shD);
  const idxD = (name)=>hd.indexOf(name);

  tareas.forEach(t=>{
    const rr = new Array(hd.length).fill("");
    const setD = (col, val) => { const c = idxD(col); if (c !== -1) rr[c] = val; };

    setD("IdDetalle", "D-" + _otc_uuid8_());
    setD("IdOT", IdOT);
    setD("CodigoTarea", (t.codigo ?? "").toString().trim());
    setD("NombreTarea", (t.nombre ?? "").toString().trim());
    setD("Sistema", (t.sistema ?? "").toString().trim());
    setD("Subsistema", (t.subsistema ?? "").toString().trim());
    setD("Sector", (t.sector ?? "").toString().trim() || Sector);
    setD("EstadoTarea", "pendiente");
    setD("Check", true); // por defecto tildada
    setD("Usuario", user.u || "");
    setD("Timestamp", new Date());

    shD.appendRow(rr);
  });

  return { ok:true, IdOT, NroOT };
}
