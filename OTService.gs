// ===================== OT CONFIG (NO CHOCA CON OT EXISTENTE) =====================
const SHEET_OT_CREATE = "OrdenesTrabajo";
const SHEET_OT_TAREAS_CREATE = "OT_Tareas";
const SHEET_EMPL = "Empleados";

const OT_COLS_REQUIRED = [
  "IdOT","NroOT","TipoOT","EstadoOT","Fecha","Interno","Dominio","Sociedad","Deposito","Sector",
  "Solicita","Descripcion","Usuario","Timestamp"
];

const OT_TAREAS_COLS_REQUIRED = [
  "IdDetalle","IdOT","CodigoTarea","NombreTarea","Sistema","Subsistema","Sector",
  "EstadoTarea","Check","Usuario","Timestamp"
];

const EMPL_COLS_REQUIRED = ["Legajo","Nombre","Activo"];

// ===================== helpers schema =====================
function _getHeaders_(sh){
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1,1,1,lastCol).getValues()[0].map(h => (h||"").toString().trim());
}

function _ensureSchema_(sheetName, required){
  const sh = _sheet(sheetName);
  const headers = _getHeaders_(sh);
  if (headers.length === 0){
    sh.getRange(1,1,1,required.length).setValues([required]);
    return;
  }
  const missing = required.filter(c => !headers.includes(c));
  if (missing.length){
    sh.getRange(1, headers.length+1, 1, missing.length).setValues([missing]);
  }
}

function _ensureOT_(){
  _ensureSchema_(SHEET_OT_CREATE, OT_COLS_REQUIRED);
  _ensureSchema_(SHEET_OT_TAREAS_CREATE, OT_TAREAS_COLS_REQUIRED);
  _ensureSchema_(SHEET_EMPL, EMPL_COLS_REQUIRED);
}

function _norm_(s){ return (s ?? "").toString().trim().toLowerCase(); }
function _isTrue_(v){
  const s = (v ?? "").toString().trim().toLowerCase();
  return v === true || s === "true" || s === "1" || s === "si" || s === "sí";
}

// ===================== options =====================
function getOTFiltersOptions(token){
  _requireSession_(token);
  _ensureOT_();

  const sh = _sheet(SHEET_OT_CREATE);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return { ok:true, estados:[], sociedades:[], depositos:[], sectores:[] };

  const h = v[0].map(x => (x||"").toString().trim());
  const iEstado = h.indexOf("EstadoOT");
  const iSoc = h.indexOf("Sociedad");
  const iDep = h.indexOf("Deposito");
  const iSec = h.indexOf("Sector");

  const pick = (i) => i === -1 ? [] : v.slice(1).map(r => (r[i]??"").toString().trim()).filter(Boolean);

  const estados = [...new Set(pick(iEstado).map(_norm_))].filter(Boolean).sort();
  const sociedades = [...new Set(pick(iSoc))].sort((a,b)=>a.localeCompare(b,"es"));
  const depositos = [...new Set(pick(iDep))].sort((a,b)=>a.localeCompare(b,"es"));
  const sectores = [...new Set(pick(iSec))].sort((a,b)=>a.localeCompare(b,"es"));

  return { ok:true, estados, sociedades, depositos, sectores };
}

function getEmpleadosOptions(token){
  _requireSession_(token);
  _ensureOT_();

  const sh = _sheet(SHEET_EMPL);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return { ok:true, empleados:[] };

  const h = v[0].map(x => (x||"").toString().trim());
  const iLeg = h.indexOf("Legajo");
  const iNom = h.indexOf("Nombre");
  const iAct = h.indexOf("Activo");

  const empleados = v.slice(1).map(r => {
    const leg = (iLeg===-1?"":(r[iLeg]??"").toString().trim());
    const nom = (iNom===-1?"":(r[iNom]??"").toString().trim());
    const act = (iAct===-1? true : _isTrue_(r[iAct]));
    if (!leg && !nom) return null;
    if (!act) return null;
    return { value: leg, label: `${leg} — ${nom}`, nombre: nom };
  }).filter(Boolean);

  const seen = new Set();
  const uniq = empleados.filter(e => seen.has(e.value) ? false : (seen.add(e.value), true));
  uniq.sort((a,b)=>a.label.localeCompare(b.label,"es"));
  return { ok:true, empleados: uniq };
}

// ===================== SEARCH (NO CARGA TODO) =====================
function searchOT(token, q){
  _requireSession_(token);
  _ensureOT_();

  const interno = (q?.interno ?? "").toString().trim();
  const nroOT = (q?.nroOT ?? "").toString().trim();

  if (!interno && !nroOT) return { ok:true, rows: [], hint:"Ingresá Interno o N° OT y tocá Buscar." };

  const tipo = _norm_(q?.tipo || "");
  const estado = _norm_(q?.estado || "");
  const sociedad = (q?.sociedad || "").toString().trim();
  const deposito = (q?.deposito || "").toString().trim();
  const sector = (q?.sector || "").toString().trim();
  const estadoTarea = _norm_(q?.estadoTarea || "");

  const sh = _sheet(SHEET_OT_CREATE);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return { ok:true, rows: [] };

  const h = v[0].map(x => (x||"").toString().trim());
  const idx = (name)=>h.indexOf(name);
  const get = (row,i)=>(i===-1?"":row[i]);

  const iId = idx("IdOT");
  const iNro = idx("NroOT");
  const iTipo = idx("TipoOT");
  const iEst = idx("EstadoOT");
  const iFecha = idx("Fecha");
  const iInt = idx("Interno");
  const iDom = idx("Dominio");
  const iSoc = idx("Sociedad");
  const iDep = idx("Deposito");
  const iSec = idx("Sector");

  let allowedByTask = null;
  if (estadoTarea){
    allowedByTask = new Set();
    const shT = _sheet(SHEET_OT_TAREAS_CREATE);
    const vt = shT.getDataRange().getValues();
    if (vt.length >= 2){
      const ht = vt[0].map(x => (x||"").toString().trim());
      const itIdOT = ht.indexOf("IdOT");
      const itEstado = ht.indexOf("EstadoTarea");
      vt.slice(1).forEach(r=>{
        const id = (itIdOT===-1?"":(r[itIdOT]??"").toString().trim());
        const es = _norm_(itEstado===-1?"":(r[itEstado]??""));
        if (id && es === estadoTarea) allowedByTask.add(id);
      });
    }
  }

  const rows = v.slice(1).map((r,k)=>{
    const rowNum = k+2;

    const IdOT = (get(r,iId)||"").toString().trim();
    const NroOT = (get(r,iNro)||"").toString().trim();
    const TipoOT = _norm_(get(r,iTipo));
    const EstadoOT = _norm_(get(r,iEst));
    const Interno = (get(r,iInt)||"").toString().trim();
    const Dominio = (get(r,iDom)||"").toString().trim();
    const Sociedad = (get(r,iSoc)||"").toString().trim();
    const Deposito = (get(r,iDep)||"").toString().trim();
    const Sector = (get(r,iSec)||"").toString().trim();

    const fechaVal = get(r,iFecha);
    const FechaMs = (fechaVal instanceof Date) ? fechaVal.getTime() : (fechaVal ? new Date(fechaVal).getTime() : null);

    return { _row: rowNum, IdOT, NroOT, TipoOT, EstadoOT, FechaMs, Interno, Dominio, Sociedad, Deposito, Sector };
  }).filter(x=>{
    if (!x.IdOT) return false;

    const okInterno = !interno || _norm_(x.Interno) === _norm_(interno);
    const okNro = !nroOT || x.NroOT === nroOT;

    const okTipo = !tipo || x.TipoOT === tipo;
    const okEst = !estado || x.EstadoOT === estado;
    const okSoc = !sociedad || x.Sociedad === sociedad;
    const okDep = !deposito || x.Deposito === deposito;
    const okSec = !sector || x.Sector === sector;

    const okTask = !allowedByTask || allowedByTask.has(x.IdOT);

    return (okInterno && okNro && okTipo && okEst && okSoc && okDep && okSec && okTask);
  });

  rows.sort((a,b)=> (b.FechaMs??-1) - (a.FechaMs??-1));
  return { ok:true, rows };
}

// ===================== DETAILS =====================
function getOTDetails(token, idOT){
  _requireSession_(token);
  _ensureOT_();

  idOT = (idOT||"").toString().trim();
  if (!idOT) throw new Error("IdOT inválido.");

  const sh = _sheet(SHEET_OT_CREATE);
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => (x||"").toString().trim());
  const idx = (name)=>h.indexOf(name);
  const get = (row,i)=>(i===-1?"":row[i]);

  const iId = idx("IdOT");
  const iNro = idx("NroOT");
  const iTipo = idx("TipoOT");
  const iEst = idx("EstadoOT");
  const iFecha = idx("Fecha");
  const iInt = idx("Interno");
  const iDom = idx("Dominio");
  const iSoc = idx("Sociedad");
  const iDep = idx("Deposito");
  const iSec = idx("Sector");
  const iSol = idx("Solicita");
  const iDes = idx("Descripcion");

  let head = null;
  for (let r=1;r<v.length;r++){
    const id = (get(v[r],iId)||"").toString().trim();
    if (id === idOT){
      const fechaVal = get(v[r],iFecha);
      const FechaMs = (fechaVal instanceof Date) ? fechaVal.getTime() : (fechaVal ? new Date(fechaVal).getTime() : null);
      head = {
        IdOT: idOT,
        NroOT: (get(v[r],iNro)||"").toString().trim(),
        TipoOT: (get(v[r],iTipo)||"").toString().trim(),
        EstadoOT: (get(v[r],iEst)||"").toString().trim(),
        FechaMs,
        Interno: (get(v[r],iInt)||"").toString().trim(),
        Dominio: (get(v[r],iDom)||"").toString().trim(),
        Sociedad: (get(v[r],iSoc)||"").toString().trim(),
        Deposito: (get(v[r],iDep)||"").toString().trim(),
        Sector: (get(v[r],iSec)||"").toString().trim(),
        Solicita: (get(v[r],iSol)||"").toString().trim(),
        Descripcion: (get(v[r],iDes)||"").toString().trim(),
      };
      break;
    }
  }
  if (!head) throw new Error("No se encontró la OT.");

  const shT = _sheet(SHEET_OT_TAREAS_CREATE);
  const vt = shT.getDataRange().getValues();
  let tasks = [];
  if (vt.length >= 2){
    const ht = vt[0].map(x => (x||"").toString().trim());
    const it = (n)=>ht.indexOf(n);
    const itIdOT = it("IdOT");

    const itCod = it("CodigoTarea");
    const itNom = it("NombreTarea");
    const itSis = it("Sistema");
    const itSub = it("Subsistema");
    const itSec = it("Sector");
    const itEst = it("EstadoTarea");
    const itChk = it("Check");

    tasks = vt.slice(1).map((r,k)=>{
      const rowNum = k+2;
      const id = (itIdOT===-1?"":(r[itIdOT]??"").toString().trim());
      if (id !== idOT) return null;
      return {
        _row: rowNum,
        CodigoTarea: (itCod===-1?"":(r[itCod]??"").toString().trim()),
        NombreTarea: (itNom===-1?"":(r[itNom]??"").toString().trim()),
        Sistema: (itSis===-1?"":(r[itSis]??"").toString().trim()),
        Subsistema: (itSub===-1?"":(r[itSub]??"").toString().trim()),
        Sector: (itSec===-1?"":(r[itSec]??"").toString().trim()),
        EstadoTarea: (itEst===-1?"":(r[itEst]??"").toString().trim()),
        Check: _isTrue_(itChk===-1?false:r[itChk]),
      };
    }).filter(Boolean);
  }

  return { ok:true, head, tasks };
}

// ===================== ANULAR =====================
function anularOT(token, idOT, motivo){
  const user = _requireSession_(token);
  _ensureOT_();

  idOT = (idOT||"").toString().trim();
  if (!idOT) throw new Error("IdOT inválido.");

  const sh = _sheet(SHEET_OT_CREATE);
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => (x||"").toString().trim());
  const iId = h.indexOf("IdOT");
  const iEst = h.indexOf("EstadoOT");
  const iUsr = h.indexOf("Usuario");
  const iTS = h.indexOf("Timestamp");
  const iDes = h.indexOf("Descripcion");

  for (let r=1;r<v.length;r++){
    const id = (iId===-1?"":(v[r][iId]??"").toString().trim());
    if (id === idOT){
      if (iEst !== -1) sh.getRange(r+1,iEst+1).setValue("anulada");
      if (iUsr !== -1) sh.getRange(r+1,iUsr+1).setValue(user.u||"");
      if (iTS !== -1) sh.getRange(r+1,iTS+1).setValue(new Date());
      if (iDes !== -1 && motivo) {
        const prev = (v[r][iDes]??"").toString();
        sh.getRange(r+1,iDes+1).setValue((prev ? prev + "\n" : "") + "ANULADA: " + motivo);
      }
      return { ok:true };
    }
  }
  throw new Error("OT no encontrada.");
}

// ===================== CONFIRMAR (1 operario obligatorio + supervisor) =====================
function confirmarOT(token, payload){
  const user = _requireSession_(token);
  _ensureOT_();

  const idOT = (payload?.idOT||"").toString().trim();
  if (!idOT) throw new Error("IdOT inválido.");

  const operarios = (payload?.operarios || []).filter(Boolean);
  const supervisor = (payload?.supervisor || "").toString().trim();
  if (operarios.length < 1) throw new Error("Debés seleccionar al menos 1 operario.");
  if (!supervisor) throw new Error("Supervisor es obligatorio.");

  const shT = _sheet(SHEET_OT_TAREAS_CREATE);
  const vt = shT.getDataRange().getValues();
  if (vt.length < 2) throw new Error("OT no tiene tareas.");

  const ht = vt[0].map(x => (x||"").toString().trim());
  const iIdOT = ht.indexOf("IdOT");
  const iChk = ht.indexOf("Check");
  const iEstT = ht.indexOf("EstadoTarea");
  const iUsr = ht.indexOf("Usuario");
  const iTS = ht.indexOf("Timestamp");
  if (iIdOT === -1 || iChk === -1) throw new Error("Falta esquema OT_Tareas.");

  const checks = payload?.checks || [];
  const map = new Map(checks.map(x => [Number(x.row), !!x.checked]));

  let total = 0, ok = 0;

  for (let r=1;r<vt.length;r++){
    const rowNum = r+1;
    const id = (iIdOT===-1?"":(vt[r][iIdOT]??"").toString().trim());
    if (id !== idOT) continue;

    total++;
    const newVal = map.has(rowNum) ? map.get(rowNum) : _isTrue_(vt[r][iChk]);
    if (newVal) ok++;

    shT.getRange(rowNum, iChk+1).setValue(newVal);
    if (iEstT !== -1) shT.getRange(rowNum, iEstT+1).setValue(newVal ? "ok" : "pendiente");
    if (iUsr !== -1) shT.getRange(rowNum, iUsr+1).setValue(user.u||"");
    if (iTS !== -1) shT.getRange(rowNum, iTS+1).setValue(new Date());
  }

  if (total === 0) throw new Error("OT no tiene tareas.");

  const estadoOT = (ok === total) ? "confirmada" : "parcial";

  const sh = _sheet(SHEET_OT_CREATE);
  const v2 = sh.getDataRange().getValues();
  const h2 = v2[0].map(x => (x||"").toString().trim());
  const iId = h2.indexOf("IdOT");
  const iEst = h2.indexOf("EstadoOT");
  const iDes = h2.indexOf("Descripcion");
  const iUsr2 = h2.indexOf("Usuario");
  const iTS2 = h2.indexOf("Timestamp");

  for (let r=1;r<v2.length;r++){
    const id = (iId===-1?"":(v2[r][iId]??"").toString().trim());
    if (id === idOT){
      if (iEst !== -1) sh.getRange(r+1,iEst+1).setValue(estadoOT);
      if (iUsr2 !== -1) sh.getRange(r+1,iUsr2+1).setValue(user.u||"");
      if (iTS2 !== -1) sh.getRange(r+1,iTS2+1).setValue(new Date());

      if (iDes !== -1){
        const prev = (v2[r][iDes]??"").toString();
        const firma = `\nCONFIRMACION: Operarios(${operarios.join(", ")}) Supervisor(${supervisor}) - ${new Date().toLocaleString()}`;
        sh.getRange(r+1,iDes+1).setValue(prev + firma);
      }
      break;
    }
  }

  return { ok:true, estadoOT, ok, total };
}
