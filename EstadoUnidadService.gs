// ===================== ESTADO UNIDAD =====================
// Vista: estado de preventivos por unidad + reprogramación + historial
// Requiere: HojasPreventivas (SHEET_HP) con CadaKm/CadaDias y (nuevo) AvisoKm/AvisoDias
// Crea/usa: PreventivosUnidad, PreventivosUnidadHis

const SHEET_PU = "PreventivosUnidad";
const SHEET_PU_HIS = "PreventivosUnidadHis";

const PU_COLS_REQUIRED = [
  "IdPU",
  "Interno",
  "IdHP",
  "NombreHP",
  "Control",          // "km" | "dias"
  "CadaKm",
  "CadaDias",
  "AvisoKm",
  "AvisoDias",
  "UltimoKm",
  "UltimaFecha",      // Date
  "ProximoKm",
  "ProximaFecha",     // Date
  "EstadoCalc",       // normal | por_pasar | pasado
  "PendienteOT",      // TRUE/FALSE
  "IdOTPendiente",
  "UltimoIdOT",
  "UltimaAccion",     // confirmacion | reprogramacion | creacion
  "Usuario",
  "Timestamp"
];

const PU_HIS_COLS_REQUIRED = [
  "IdHis",
  "Interno",
  "IdHP",
  "NombreHP",
  "Accion",           // confirmacion | reprogramacion | creacion
  "Antes",            // json
  "Despues",          // json
  "Motivo",
  "Usuario",
  "Timestamp"
];

function ensureEstadoUnidadSheets_(){
  // Dependencias
  if (typeof ensurePreventivosSheets_ === "function") ensurePreventivosSheets_();

  const ss = _ss();

  let sh = ss.getSheetByName(SHEET_PU);
  if (!sh) sh = ss.insertSheet(SHEET_PU);
  ensureCols_(sh, PU_COLS_REQUIRED);

  let sh2 = ss.getSheetByName(SHEET_PU_HIS);
  if (!sh2) sh2 = ss.insertSheet(SHEET_PU_HIS);
  ensureCols_(sh2, PU_HIS_COLS_REQUIRED);
}

// Util: headers -> idx map
function _eu_headers_(sh){
  const last = sh.getLastColumn();
  if (last === 0) return [];
  return sh.getRange(1,1,1,last).getValues()[0].map(h => (h||"").toString().trim());
}

function _eu_norm_(v){
  return (v ?? "").toString().trim();
}

function _eu_num_(v){
  const n = parseInt((v ?? "").toString().replace(/[^\d]/g,""), 10);
  return isFinite(n) ? n : 0;
}

function _eu_isTrue_(v){
  const s = (v ?? "").toString().toLowerCase().trim();
  return (v === true) || s === "true" || s === "1" || s === "si" || s === "sí";
}

function _eu_date_(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isFinite(d.getTime()) ? d : null;
}

function _eu_isoDate_(d){
  if (!d) return "";
  const dd = new Date(d);
  if (!isFinite(dd.getTime())) return "";
  return dd.toISOString().slice(0,10);
}

function _eu_addDays_(d, days){
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + (days||0));
  return x;
}

function _eu_uuid8_(){
  return Math.random().toString(36).slice(2,10).toUpperCase();
}

function _eu_findHPById_(idHP){
  const sh = _sheet(SHEET_HP);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return null;
  const h = v[0].map(x => (x||"").toString().trim());
  const iId = h.indexOf("IdHP");
  const iNom = h.indexOf("NombreHP");
  const iSec = h.indexOf("Sector");
  const iCadaKm = h.indexOf("CadaKm");
  const iCadaDias = h.indexOf("CadaDias");
  const iAvisoKm = h.indexOf("AvisoKm");
  const iAvisoDias = h.indexOf("AvisoDias");

  for (let r=1;r<v.length;r++){
    const id = (iId===-1?"":(v[r][iId]??"").toString().trim());
    if (id === idHP){
      return {
        IdHP:id,
        NombreHP:(iNom===-1?"":(v[r][iNom]??"").toString().trim()),
        Sector:(iSec===-1?"":(v[r][iSec]??"").toString().trim()),
        CadaKm:_eu_num_(iCadaKm===-1? "" : v[r][iCadaKm]),
        CadaDias:_eu_num_(iCadaDias===-1? "" : v[r][iCadaDias]),
        AvisoKm:_eu_num_(iAvisoKm===-1? "" : v[r][iAvisoKm]),
        AvisoDias:_eu_num_(iAvisoDias===-1? "" : v[r][iAvisoDias]),
      };
    }
  }
  return null;
}

function _eu_findIdHPByName_(nombreHP){
  const name = _eu_norm_(nombreHP);
  if (!name) return "";
  const sh = _sheet(SHEET_HP);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return "";
  const h = v[0].map(x => (x||"").toString().trim());
  const iId = h.indexOf("IdHP");
  const iNom = h.indexOf("NombreHP");
  if (iId === -1 || iNom === -1) return "";
  for (let r=1;r<v.length;r++){
    const nom = (v[r][iNom]??"").toString().trim();
    if (nom && nom.toLowerCase() === name.toLowerCase()){
      return (v[r][iId]??"").toString().trim();
    }
  }
  return "";
}

function _eu_getUnidadByInterno_(interno){
  const sh = _sheet(SHEET_CHASIS);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return null;
  const h = v[0].map(x => (x||"").toString().trim());
  const iInt = h.indexOf("Interno");
  if (iInt === -1) return null;

  const get = (row, name) => {
    const i = h.indexOf(name);
    return i === -1 ? "" : (row[i] ?? "").toString().trim();
  };

  const target = _eu_norm_(interno);
  for (let r=1;r<v.length;r++){
    const cur = (v[r][iInt]??"").toString().trim();
    if (cur === target){
      return {
        Interno: target,
        Dominio: get(v[r],"Dominio"),
        Sociedad: get(v[r],"Sociedad"),
        Deposito: get(v[r],"Deposito"),
        Tipo: get(v[r],"Tipo"),
        KmRecorridos: _eu_num_(get(v[r],"KmRecorridos")),
        Marca: get(v[r],"Marca"),
        Modelo: get(v[r],"Modelo"),
        Motor: get(v[r],"Motor"),
        "Nro. Chasis": get(v[r],"Nro. Chasis"),
        "Nro Motor": get(v[r],"Nro Motor"),
        Estado: get(v[r],"Estado"),
      };
    }
  }
  return null;
}


function _eu_pendingOTMap_(interno){
  // Devuelve map {IdHP: {pendiente:true, idOT:""} } para la unidad
  try{
    const sh = _sheet("OrdenesTrabajo");
    const v = sh.getDataRange().getValues();
    if (v.length < 2) return {};
    const h = v[0].map(x => (x||"").toString().trim());
    const iTipo = h.indexOf("TipoOT");
    const iInt = h.indexOf("Interno");
    const iIdHP = h.indexOf("IdHP");
    const iEst = h.indexOf("EstadoOT");
    const iIdOT = h.indexOf("IdOT");
    if (iInt === -1 || iEst === -1) return {};

    const map = {};
    const target = (interno||"").toString().trim();

    for (let r=1;r<v.length;r++){
      const curInt = (iInt===-1?"":(v[r][iInt]??"").toString().trim());
      if (curInt !== target) continue;

      const tipo = (iTipo===-1?"":(v[r][iTipo]??"").toString().trim()).toLowerCase();
      if (tipo !== "preventiva") continue;

      const est = (iEst===-1?"":(v[r][iEst]??"").toString().trim()).toLowerCase();
      if (est === "confirmada" || est === "anulada") continue;

      const idHP = (iIdHP===-1?"":(v[r][iIdHP]??"").toString().trim());
      if (!idHP) continue;

      const idOT = (iIdOT===-1?"":(v[r][iIdOT]??"").toString().trim());

      // si ya hay uno, dejamos el primero encontrado
      if (!map[idHP]) map[idHP] = { pendiente:true, idOT };
    }
    return map;
  } catch(e){
    return {};
  }
}

function _eu_computeEstado_(control, unidadKm, hoy, rec){
  // Defaults
  const cadaKm   = _eu_num_(rec.CadaKm);
  const cadaDias = _eu_num_(rec.CadaDias);
  const avisoKm  = _eu_num_(rec.AvisoKm);
  const avisoDias= _eu_num_(rec.AvisoDias);

  let estado = "normal";
  let venceStr = "";
  let restan = 0;
  let pasado = 0;
  let actual = 0;

  if (control === "km"){
    const ultimoKm = _eu_num_(rec.UltimoKm);
    const proximoKm = _eu_num_(rec.ProximoKm);

    const cur = _eu_num_(unidadKm);
    actual = cur; // km actual
    const faltan = proximoKm - cur;
    restan = faltan;

    if (proximoKm <= 0){
      estado = "normal";
      venceStr = "-";
    } else if (cur > proximoKm){
      estado = "pasado";
      pasado = cur - proximoKm;
      venceStr = `${proximoKm}`;
    } else {
      // por pasar
      const umbral = (avisoKm > 0) ? avisoKm : Math.max(500, Math.round(cadaKm * 0.1)); // default
      if (faltan <= umbral) estado = "por_pasar";
      venceStr = `${proximoKm}`;
    }

    return { estado, venceStr, restan, pasado, actual, ultimo: ultimoKm, cada: cadaKm, proximo: proximoKm };

  } else {
    const ult = _eu_date_(rec.UltimaFecha);
    const prox = _eu_date_(rec.ProximaFecha);

    // "actual" = días desde última (si hay)
    if (ult) actual = Math.floor((hoy.getTime() - ult.getTime()) / 86400000);
    else actual = 0;

    if (!prox){
      estado = "normal";
      venceStr = "-";
    } else {
      const faltan = Math.ceil((prox.getTime() - hoy.getTime()) / 86400000);
      restan = faltan;

      if (hoy.getTime() > prox.getTime()){
        estado = "pasado";
        pasado = Math.ceil((hoy.getTime() - prox.getTime()) / 86400000);
        venceStr = _eu_isoDate_(prox);
      } else {
        const umbral = (avisoDias > 0) ? avisoDias : Math.max(7, Math.round(cadaDias * 0.1)); // default
        if (faltan <= umbral) estado = "por_pasar";
        venceStr = _eu_isoDate_(prox);
      }
    }

    return {
      estado, venceStr, restan, pasado, actual,
      ultimo: ult ? _eu_isoDate_(ult) : "",
      cada: cadaDias,
      proximo: prox ? _eu_isoDate_(prox) : ""
    };
  }
}

// ----------- PUBLIC: obtener estado por interno -----------
function getEstadoUnidad(token, interno){
  const user = _requireSession_(token);
  ensureEstadoUnidadSheets_();

  const intv = _eu_norm_(interno);
  if (!intv) throw new Error("Interno es obligatorio.");

  const unidad = _eu_getUnidadByInterno_(intv);
  if (!unidad) throw new Error("No se encontró la unidad (Interno).");

  const hoy = new Date();
  const pendingMap = _eu_pendingOTMap_(intv);

  // leer registros PU de esa unidad
  const sh = _sheet(SHEET_PU);
  const v = sh.getDataRange().getValues();
  const h = _eu_headers_(sh);

  const idx = (name)=>h.indexOf(name);
  const iInt = idx("Interno");
  if (v.length < 2 || iInt === -1){
    return { ok:true, unidad, rows:[], nextDue:null };
  }

  const rows = [];
  for (let r=1;r<v.length;r++){
    const curInt = (v[r][iInt]??"").toString().trim();
    if (curInt !== intv) continue;

    const rec = {};
    h.forEach((col, c)=> rec[col] = v[r][c]);

    // HP
    const idHP = _eu_norm_(rec.IdHP);
    const hp = idHP ? _eu_findHPById_(idHP) : null;

    const control = _eu_norm_(rec.Control) || (hp && hp.CadaKm ? "km" : "dias");

    // si hay cambios en HP, refrescamos frecuencias (sin tocar último)
    if (hp){
      rec.NombreHP = hp.NombreHP || rec.NombreHP;
      rec.CadaKm = hp.CadaKm;
      rec.CadaDias = hp.CadaDias;
      rec.AvisoKm = hp.AvisoKm;
      rec.AvisoDias = hp.AvisoDias;
    }

    const calc = _eu_computeEstado_(control, unidad.KmRecorridos, hoy, rec);

    rows.push({
      _row: r+1,
      IdPU: _eu_norm_(rec.IdPU),
      Interno: intv,
      IdHP: idHP,
      NombreHP: _eu_norm_(rec.NombreHP),
      Control: control,
      Ultimo: calc.ultimo,
      Cada: calc.cada,
      Proximo: calc.proximo,
      Actual: calc.actual,
      Restan: calc.restan,
      Pasado: calc.pasado,
      Vence: calc.venceStr,
      Estado: calc.estado,
      PendienteOT: !!pendingMap[idHP],
      IdOTPendiente: (pendingMap[idHP]?.idOT || _eu_norm_(rec.IdOTPendiente)),
    });
  }

  // orden: primero pasado, luego por_pasar, luego normal, y por “vence” más cercano
  const orderEstado = { pasado:0, por_pasar:1, normal:2 };
  rows.sort((a,b)=>{
    const ea = orderEstado[a.Estado] ?? 9;
    const eb = orderEstado[b.Estado] ?? 9;
    if (ea !== eb) return ea - eb;

    // km: menor Proximo; dias: menor fecha
    const aKey = (a.Control==="km") ? (_eu_num_(a.Proximo)||999999999) : (Date.parse(a.Proximo||"9999-12-31")||9999999999999);
    const bKey = (b.Control==="km") ? (_eu_num_(b.Proximo)||999999999) : (Date.parse(b.Proximo||"9999-12-31")||9999999999999);
    return aKey - bKey;
  });

  const nextDue = rows.length ? rows[0] : null;

  return { ok:true, unidad, rows, nextDue };
}

// ----------- REPROGRAMACIÓN MANUAL (con historial) -----------
function reprogramarPreventivoUnidad(token, payload){
  const user = _requireSession_(token);
  ensureEstadoUnidadSheets_();

  const Interno = _eu_norm_(payload?.Interno);
  const IdHP = _eu_norm_(payload?.IdHP);
  const Motivo = _eu_norm_(payload?.Motivo) || "reprogramación manual";

  if (!Interno) throw new Error("Interno es obligatorio.");
  if (!IdHP) throw new Error("IdHP es obligatorio.");

  const unidad = _eu_getUnidadByInterno_(Interno);
  if (!unidad) throw new Error("No se encontró la unidad (Interno).");

  const hp = _eu_findHPById_(IdHP);
  if (!hp) throw new Error("No se encontró la hoja preventiva (IdHP).");

  const control = (hp.CadaKm && hp.CadaKm > 0) ? "km" : "dias";
  const hoy = new Date();

  const sh = _sheet(SHEET_PU);
  const v = sh.getDataRange().getValues();
  const h = _eu_headers_(sh);
  const idx = (name)=>h.indexOf(name);

  const iInt = idx("Interno");
  const iIdHP = idx("IdHP");
  const iIdPU = idx("IdPU");

  if (iInt === -1 || iIdHP === -1) throw new Error("PreventivosUnidad sin columnas necesarias.");

  let rowNum = 0;
  for (let r=1;r<v.length;r++){
    const curInt = (v[r][iInt]??"").toString().trim();
    const curId = (v[r][iIdHP]??"").toString().trim();
    if (curInt === Interno && curId === IdHP){ rowNum = r+1; break; }
  }

  if (!rowNum){
    // si no existe, lo creamos (y lo reprogramamos)
    rowNum = sh.getLastRow() + 1;
    const rr = new Array(h.length).fill("");
    const set = (col,val)=>{ const c=idx(col); if (c!==-1) rr[c]=val; };
    set("IdPU", "PU-" + _eu_uuid8_());
    set("Interno", Interno);
    set("IdHP", IdHP);
    set("NombreHP", hp.NombreHP);
    set("Control", control);
    set("CadaKm", hp.CadaKm);
    set("CadaDias", hp.CadaDias);
    set("AvisoKm", hp.AvisoKm);
    set("AvisoDias", hp.AvisoDias);
    set("PendienteOT", "FALSE");
    set("IdOTPendiente", "");
    set("UltimaAccion", "creacion");
    set("Usuario", user.u||"");
    set("Timestamp", new Date());
    sh.appendRow(rr);
  }

  // leer antes
  const before = {};
  h.forEach((col,c)=> before[col] = sh.getRange(rowNum,c+1).getValue());

  // valores nuevos
  const setCell = (col, val)=>{
    const c = idx(col);
    if (c === -1) return;
    sh.getRange(rowNum, c+1).setValue(val);
  };

  setCell("NombreHP", hp.NombreHP);
  setCell("Control", control);
  setCell("CadaKm", hp.CadaKm);
  setCell("CadaDias", hp.CadaDias);
  setCell("AvisoKm", hp.AvisoKm);
  setCell("AvisoDias", hp.AvisoDias);

  if (control === "km"){
    const nuevoUltKm = _eu_num_(payload?.UltimoKm);
    if (nuevoUltKm <= 0) throw new Error("UltimoKm inválido.");
    const prox = nuevoUltKm + (hp.CadaKm||0);

    setCell("UltimoKm", nuevoUltKm);
    // Permite ajustar fecha realizada (opcional)
    const fechaISO = _eu_norm_(payload?.UltimaFechaISO);
    setCell("UltimaFecha", fechaISO ? fechaISO : hoy);
setCell("ProximoKm", prox);
    setCell("ProximaFecha", "");
  } else {
    const fechaISO = _eu_norm_(payload?.UltimaFechaISO);
    if (!fechaISO) throw new Error("UltimaFecha es obligatoria (YYYY-MM-DD).");
    const ult = new Date(fechaISO + "T00:00:00");
    if (!isFinite(ult.getTime())) throw new Error("UltimaFecha inválida.");
    const prox = _eu_addDays_(ult, (hp.CadaDias||0));

    setCell("UltimaFecha", ult);
    setCell("UltimoKm", unidad.KmRecorridos || 0);
    setCell("ProximaFecha", prox);
    setCell("ProximoKm", "");
  }

  setCell("PendienteOT", "FALSE");
  setCell("IdOTPendiente", "");
  setCell("UltimaAccion", "reprogramacion");
  setCell("Usuario", user.u||"");
  setCell("Timestamp", new Date());

  // after snapshot
  const after = {};
  h.forEach((col,c)=> after[col] = sh.getRange(rowNum,c+1).getValue());

  // historial
  _eu_appendHis_(Interno, IdHP, hp.NombreHP, "reprogramacion", before, after, Motivo, user);

  return { ok:true };
}

function _eu_appendHis_(Interno, IdHP, NombreHP, Accion, beforeObj, afterObj, Motivo, user){
  const shH = _sheet(SHEET_PU_HIS);
  const h = _eu_headers_(shH);
  const idx = (name)=>h.indexOf(name);

  const rr = new Array(h.length).fill("");
  const set = (col,val)=>{ const c=idx(col); if (c!==-1) rr[c]=val; };

  set("IdHis", "H-" + _eu_uuid8_());
  set("Interno", Interno);
  set("IdHP", IdHP);
  set("NombreHP", NombreHP || "");
  set("Accion", Accion);
  set("Antes", JSON.stringify(beforeObj || {}));
  set("Despues", JSON.stringify(afterObj || {}));
  set("Motivo", Motivo || "");
  set("Usuario", (user?.u || user?.Usuario || ""));
  set("Timestamp", new Date());

  shH.appendRow(rr);
}

// ----------- HOOK desde confirmar OT (preventiva) -----------
function EU_onConfirmPreventivoOT_(interno, idHP, idOT, user){
  try{
    ensureEstadoUnidadSheets_();

    const Interno = _eu_norm_(interno);
    if (!Interno) return;

    // Si no hay IdHP, intentamos resolver por nombre (desde OT)
    let IdHP = _eu_norm_(idHP);

    // unidad
    const unidad = _eu_getUnidadByInterno_(Interno);
    if (!unidad) return;

    const hoy = new Date();
    const curKm = unidad.KmRecorridos || 0;

    // si sigue vacío, no hacemos nada
    if (!IdHP) return;

    const hp = _eu_findHPById_(IdHP);
    if (!hp) return;

    const control = (hp.CadaKm && hp.CadaKm > 0) ? "km" : "dias";

    const sh = _sheet(SHEET_PU);
    const v = sh.getDataRange().getValues();
    const h = _eu_headers_(sh);
    const idx = (name)=>h.indexOf(name);

    const iInt = idx("Interno");
    const iIdHP = idx("IdHP");

    let rowNum = 0;
    for (let r=1;r<v.length;r++){
      const curInt = (v[r][iInt]??"").toString().trim();
      const curId = (v[r][iIdHP]??"").toString().trim();
      if (curInt === Interno && curId === IdHP){ rowNum = r+1; break; }
    }

    // si no existe, creamos
    if (!rowNum){
      const rr = new Array(h.length).fill("");
      const set = (col,val)=>{ const c=idx(col); if (c!==-1) rr[c]=val; };

      set("IdPU", "PU-" + _eu_uuid8_());
      set("Interno", Interno);
      set("IdHP", IdHP);
      set("NombreHP", hp.NombreHP);
      set("Control", control);
      set("CadaKm", hp.CadaKm);
      set("CadaDias", hp.CadaDias);
      set("AvisoKm", hp.AvisoKm);
      set("AvisoDias", hp.AvisoDias);

      // inicializamos como si fuera confirmación
      if (control === "km"){
        const prox = curKm + (hp.CadaKm||0);
        set("UltimoKm", curKm);
        set("UltimaFecha", hoy);
        set("ProximoKm", prox);
        set("ProximaFecha", "");
      } else {
        const prox = _eu_addDays_(hoy, (hp.CadaDias||0));
        set("UltimaFecha", hoy);
        set("UltimoKm", curKm);
        set("ProximaFecha", prox);
        set("ProximoKm", "");
      }

      set("PendienteOT", "FALSE");
      set("IdOTPendiente", "");
      set("UltimoIdOT", idOT || "");
      set("UltimaAccion", "confirmacion");
      set("Usuario", user?.u || "");
      set("Timestamp", new Date());

      sh.appendRow(rr);

      // historial creacion+confirmacion
      _eu_appendHis_(Interno, IdHP, hp.NombreHP, "creacion", {}, rr.reduce((o,_,i)=>{o[h[i]]=rr[i]; return o;},{}), "Creado automáticamente por confirmación de OT", user);
      return;
    }

    // snapshot before
    const before = {};
    h.forEach((col,c)=> before[col] = sh.getRange(rowNum,c+1).getValue());

    // actualizar fila existente a confirmación
    const setCell = (col, val)=>{
      const c = idx(col);
      if (c === -1) return;
      sh.getRange(rowNum, c+1).setValue(val);
    };

    setCell("NombreHP", hp.NombreHP);
    setCell("Control", control);
    setCell("CadaKm", hp.CadaKm);
    setCell("CadaDias", hp.CadaDias);
    setCell("AvisoKm", hp.AvisoKm);
    setCell("AvisoDias", hp.AvisoDias);

    if (control === "km"){
      const prox = curKm + (hp.CadaKm||0);
      setCell("UltimoKm", curKm);
      setCell("UltimaFecha", hoy);
      setCell("ProximoKm", prox);
      setCell("ProximaFecha", "");
    } else {
      const prox = _eu_addDays_(hoy, (hp.CadaDias||0));
      setCell("UltimaFecha", hoy);
      setCell("UltimoKm", curKm);
      setCell("ProximaFecha", prox);
      setCell("ProximoKm", "");
    }

    setCell("PendienteOT", "FALSE");
    setCell("IdOTPendiente", "");
    setCell("UltimoIdOT", idOT || "");
    setCell("UltimaAccion", "confirmacion");
    setCell("Usuario", user?.u || "");
    setCell("Timestamp", new Date());

    // after snapshot
    const after = {};
    h.forEach((col,c)=> after[col] = sh.getRange(rowNum,c+1).getValue());

    _eu_appendHis_(Interno, IdHP, hp.NombreHP, "confirmacion", before, after, "Confirmación de OT preventiva", user);

  } catch(e){
    // NO rompas el flujo de confirmar OT
    console.error("EU_onConfirmPreventivoOT_ error:", e);
  }
}
