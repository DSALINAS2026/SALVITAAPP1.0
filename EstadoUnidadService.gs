// ========= CONFIG =========
const SPREADSHEET_ID = "1Lyeb-ht-g41QMHJlgNhBg1dmxamlaieWOjCf6yZPsmg";

const SHEET_USERS = "Usuarios";
const SHEET_CHASIS = "ChasisBD";
const SHEET_MOV = "MovimientoUnidad";

// Columnas esperadas en ChasisBD (s// ===================== ESTADO UNIDAD SERVICE (v7 FIXED) =====================
// Objetivo: devolver el estado de preventivos de una unidad SIN romper nada.
// - Lee PreventivosUnidad (UltimoKm/UltimaFecha/ProximoKm/ProximaFecha/etc.)
// - Para KM: "Actual" = km recorridos desde el último (KmActual - UltimoKm)
// - Para DÍAS: "Actual" = días transcurridos desde ÚltimaFecha
// - Detecta OT preventiva pendiente por (Interno + IdHP) en ordenesTrabajo
// - Incluye campos "compatibles" (CadaStr/UltimoStr/...) por si tu UI vieja los usa
// - Incluye reprogramarPreventivoUnidadV7() con historial en PreventivosUnidadHis
// - Nunca tira throw hacia el cliente: siempre devuelve {ok:false,msg}

const EU7_SHEET_CHASIS      = "ChasisBD";
const EU7_SHEET_PREV_UNIDAD = "PreventivosUnidad";
const EU7_SHEET_OT          = "ordenesTrabajo";
const EU7_SHEET_HIS         = "PreventivosUnidadHis";

// ---------- helpers ----------
function EU7_normKey_(s){
  s = (s ?? "").toString().trim().toLowerCase();
  s = s.normalize("NFD").replace(/[\u0300-\u036f]/g,""); // sin acentos
  s = s.replace(/\s+/g," ");
  return s;
}
function EU7_buildIndex_(headers){
  const idx = {};
  headers.forEach((h,i)=>{ idx[EU7_normKey_(h)] = i; });
  return idx;
}
function EU7_get_(row, idx, ...names){
  for (const n of names){
    const k = EU7_normKey_(n);
    if (k in idx){
      const v = row[idx[k]];
      if (v !== "" && v !== null && typeof v !== "undefined") return v;
    }
  }
  return "";
}
function EU7_set_(row, idx, name, value){
  const k = EU7_normKey_(name);
  if (k in idx) row[idx[k]] = value;
}
function EU7_num_(v){
  const s = (v ?? "").toString().replace(/[^\d.-]/g,"");
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
function EU7_date_(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function EU7_fmtDate_(d){
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const yy = d.getFullYear();
  return `${dd}/${mm}/${yy}`;
}
function EU7_today_(){
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}
function EU7_addDays_(d, days){
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + Number(days||0));
  return x;
}
function EU7_daysBetween_(a,b){
  const ms = (a.getTime() - b.getTime());
  return Math.floor(ms / 86400000);
}
function EU7_makeId_(){
  return Utilities.getUuid().slice(0,8);
}

// ---------- main: estado ----------
function getEstadoUnidadV7(token, interno){
  try{
    const user = _requireSession_(token);

    interno = (interno || "").toString().trim();
    if (!interno) return { ok:false, msg:"Interno requerido." };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ----- chasis -----
    const shCh = ss.getSheetByName(EU7_SHEET_CHASIS);
    if (!shCh) return { ok:false, msg:`No existe la pestaña: ${EU7_SHEET_CHASIS}` };

    const chVals = shCh.getDataRange().getValues();
    if (chVals.length < 2) return { ok:false, msg:"ChasisBD sin datos." };

    const chH = chVals[0];
    const chIdx = EU7_buildIndex_(chH);

    let chRow = null;
    for (let r=1; r<chVals.length; r++){
      const val = (EU7_get_(chVals[r], chIdx, "Interno") || "").toString().trim();
      if (val === interno){ chRow = chVals[r]; break; }
    }
    if (!chRow) return { ok:false, msg:`No se encontró el interno ${interno} en ChasisBD.` };

    const chasis = {
      Interno: interno,
      Dominio: EU7_get_(chRow, chIdx, "Dominio"),
      Sociedad: EU7_get_(chRow, chIdx, "Sociedad"),
      Deposito: EU7_get_(chRow, chIdx, "Deposito", "Depósito"),
      Marca: EU7_get_(chRow, chIdx, "Marca"),
      "Nro. Chasis": EU7_get_(chRow, chIdx, "Nro. Chasis", "Nro Chasis", "Numero Chasis"),
      "Nro Motor": EU7_get_(chRow, chIdx, "Nro Motor", "Nro. Motor"),
      KmRecorridos: EU7_num_(EU7_get_(chRow, chIdx, "KmRecorridos", "Km Recorridos", "Km"))
    };

    // ----- preventivos unidad -----
    const shPU = ss.getSheetByName(EU7_SHEET_PREV_UNIDAD);
    if (!shPU) return { ok:false, msg:`No existe la pestaña: ${EU7_SHEET_PREV_UNIDAD}` };

    const puVals = shPU.getDataRange().getValues();
    if (puVals.length < 2){
      return { ok:true, chasis, items:[], me:{usuario:user?.Usuario||user?.u||"", rol:user?.Rol||user?.r||""} };
    }

    const puH = puVals[0];
    const puIdx = EU7_buildIndex_(puH);

    const today = EU7_today_();
    const kmActual = chasis.KmRecorridos;

    // ----- OTs pendientes para el interno (cargamos una vez) -----
    const shOT = ss.getSheetByName(EU7_SHEET_OT);
    const otPendByIdHP = {}; // {idHP: {IdOT,NroOT,EstadoOT}}
    if (shOT){
      const otVals = shOT.getDataRange().getValues();
      if (otVals.length >= 2){
        const otH = otVals[0];
        const otIdx = EU7_buildIndex_(otH);

        const closed = new Set(["confirmada","confirmado","ok","cerrada","cerrado","anulada","anulado"]);
        for (let r=1; r<otVals.length; r++){
          const row = otVals[r];
          const otInterno = (EU7_get_(row, otIdx, "Interno") || "").toString().trim();
          if (otInterno !== interno) continue;

          const tipo = (EU7_get_(row, otIdx, "TipoOT", "Tipo OT") || "").toString().trim().toLowerCase();
          if (tipo !== "preventiva") continue;

          const est = (EU7_get_(row, otIdx, "EstadoOT", "Estado OT") || "").toString().trim().toLowerCase();
          if (closed.has(est)) continue;

          const idHP = (EU7_get_(row, otIdx, "IdHP") || "").toString().trim();
          if (!idHP) continue;

          const ts = EU7_date_(EU7_get_(row, otIdx, "Timestamp"))?.getTime() || r;
          const prev = otPendByIdHP[idHP];
          if (!prev || ts > prev._ts){
            otPendByIdHP[idHP] = {
              IdOT: (EU7_get_(row, otIdx, "IdOT") || "").toString().trim(),
              NroOT: (EU7_get_(row, otIdx, "NroOT", "Nro OT") || "").toString().trim(),
              EstadoOT: est,
              _ts: ts
            };
          }
        }
      }
    }

    // ----- armar items -----
    const items = [];
    for (let r=1; r<puVals.length; r++){
      const row = puVals[r];
      const puInterno = (EU7_get_(row, puIdx, "Interno") || "").toString().trim();
      if (puInterno !== interno) continue;

      const idHP = (EU7_get_(row, puIdx, "IdHP") || "").toString().trim();
      const nombreHP = (EU7_get_(row, puIdx, "NombreHP", "Nombre Preventivo", "Preventivo") || "").toString().trim();

      const control = (EU7_get_(row, puIdx, "Control", "ControlTipo", "Control Tipo") || "").toString().trim().toLowerCase();
      const cadaKm = EU7_num_(EU7_get_(row, puIdx, "CadaKm", "IntervaloKm", "FrecuenciaKm"));
      const cadaDias = EU7_num_(EU7_get_(row, puIdx, "CadaDias", "IntervaloDias", "FrecuenciaDias"));

      const avisoKm = EU7_num_(EU7_get_(row, puIdx, "AvisoKm", "AvisarAntesKm"));
      const avisoDias = EU7_num_(EU7_get_(row, puIdx, "AvisoDias", "AvisarAntesDias"));

      const ultimoKm = EU7_num_(EU7_get_(row, puIdx, "UltimoKm", "Ultimo"));
      const ultimaFecha = EU7_date_(EU7_get_(row, puIdx, "UltimaFecha", "UltimoFecha", "Ultima Fecha"));

      const proximoKm = EU7_num_(EU7_get_(row, puIdx, "ProximoKm", "Proximo"));
      const proximaFecha = EU7_date_(EU7_get_(row, puIdx, "ProximaFecha", "ProximoFecha", "Proxima Fecha"));

      // decidir tipo (si Control viene vacío, inferimos)
      let tipo = control;
      if (!tipo){
        if (cadaKm && !cadaDias) tipo = "km";
        else if (!cadaKm && cadaDias) tipo = "dia";
        else if (cadaKm) tipo = "km";
        else tipo = "dia";
      }
      if (tipo === "dias") tipo = "dia";

      // defaults de aviso
      const avisoKmEff = avisoKm || (cadaKm ? Math.ceil(cadaKm * 0.10) : 0);
      const avisoDiasEff = avisoDias || 7;

      // calcular proximo si falta
      let proxKmEff = proximoKm;
      let proxFechaEff = proximaFecha;

      if (tipo === "km"){
        if (!proxKmEff && cadaKm){
          proxKmEff = (ultimoKm || 0) + cadaKm;
        }
      } else {
        if (!proxFechaEff && ultimaFecha && cadaDias){
          proxFechaEff = EU7_addDays_(ultimaFecha, cadaDias);
        }
      }

      // armar strings y estado
      let cadaStr = "";
      let ultimoStr = "";
      let actualStr = "";
      let proximoStr = "";
      let pasadoStr = "0";
      let clase = "ok";

      if (tipo === "km"){
        const actualDesdeUlt = Math.max(0, kmActual - (ultimoKm || 0));

        cadaStr = cadaKm ? `${cadaKm} Km` : "-";
        ultimoStr = (ultimoKm || ultimoKm === 0) ? `${ultimoKm} Km` : "-";
        actualStr = `${actualDesdeUlt} Km`;
        proximoStr = proxKmEff ? `${proxKmEff} Km` : "-";

        if (proxKmEff){
          const diff = kmActual - proxKmEff; // >0 pasado
          if (diff > 0){
            clase = "pasado";
            pasadoStr = `${diff} Km`;
          } else {
            pasadoStr = "0";
            if (avisoKmEff && kmActual >= (proxKmEff - avisoKmEff)){
              clase = "proximo";
            }
          }
        }
      } else {
        const diasDesdeUlt = ultimaFecha ? Math.max(0, EU7_daysBetween_(today, ultimaFecha)) : 0;

        cadaStr = cadaDias ? `${cadaDias} Días` : "-";
        ultimoStr = ultimaFecha ? EU7_fmtDate_(ultimaFecha) : "-";
        actualStr = `${diasDesdeUlt} Días`;
        proximoStr = proxFechaEff ? EU7_fmtDate_(proxFechaEff) : "-";

        if (proxFechaEff){
          const diffDays = EU7_daysBetween_(today, proxFechaEff); // >0 pasado
          if (diffDays > 0){
            clase = "pasado";
            pasadoStr = `${diffDays} Días`;
          } else {
            pasadoStr = "0";
            if (avisoDiasEff){
              const warnDate = EU7_addDays_(proxFechaEff, -avisoDiasEff);
              if (today.getTime() >= warnDate.getTime()){
                clase = "proximo";
              }
            }
          }
        }
      }

      const otPend = otPendByIdHP[idHP] || null;

      const item = {
        Interno: interno,
        IdHP: idHP,
        Codigo: idHP,
        NombreHP: nombreHP,
        Tipo: tipo,
        ControlTipo: tipo,

        Cada: cadaStr,
        Ultimo: ultimoStr,
        Actual: actualStr,
        Proximo: proximoStr,
        Pasado: pasadoStr,
        Clase: clase,

        // compat UI vieja
        CadaStr: cadaStr,
        UltimoStr: ultimoStr,
        ActualStr: actualStr,
        ProximoStr: proximoStr,
        PasadoStr: pasadoStr,

        NroOT: otPend ? (otPend.NroOT || "") : "",
        IdOT: otPend ? (otPend.IdOT || "") : "",
        OTPendiente: !!otPend
      };

      items.push(item);
    }

    return {
      ok:true,
      chasis,
      items,
      me:{ usuario:user?.Usuario||user?.u||"", rol:user?.Rol||user?.r||"" }
    };

  } catch(err){
    return { ok:false, msg: (err && err.message) ? err.message : String(err) };
  }
}

// ---------- reprogramar manual: setea NUEVO ÚLTIMO y recalcula PRÓXIMO ----------
function reprogramarPreventivoUnidadV7(token, interno, idHP, nuevo, motivo){
  try{
    const user = _requireSession_(token);

    interno = (interno || "").toString().trim();
    idHP = (idHP || "").toString().trim();
    motivo = (motivo || "").toString().trim();

    if (!interno) return { ok:false, msg:"Interno requerido." };
    if (!idHP) return { ok:false, msg:"IdHP requerido." };
    if (!motivo) return { ok:false, msg:"Motivo requerido." };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const shPU = ss.getSheetByName(EU7_SHEET_PREV_UNIDAD);
    if (!shPU) return { ok:false, msg:`No existe la pestaña: ${EU7_SHEET_PREV_UNIDAD}` };

    const vals = shPU.getDataRange().getValues();
    if (vals.length < 2) return { ok:false, msg:"PreventivosUnidad sin datos." };

    const headers = vals[0];
    const idx = EU7_buildIndex_(headers);

    let rowN = -1;
    for (let r=1; r<vals.length; r++){
      const rInterno = (EU7_get_(vals[r], idx, "Interno") || "").toString().trim();
      const rIdHP = (EU7_get_(vals[r], idx, "IdHP") || "").toString().trim();
      if (rInterno === interno && rIdHP === idHP){ rowN = r; break; }
    }
    if (rowN < 1) return { ok:false, msg:`No se encontró PreventivosUnidad para Interno=${interno} IdHP=${idHP}` };

    const row = vals[rowN];

    const control = (EU7_get_(row, idx, "Control", "ControlTipo", "Control Tipo") || "").toString().trim().toLowerCase();
    const cadaKm = EU7_num_(EU7_get_(row, idx, "CadaKm", "IntervaloKm"));
    const cadaDias = EU7_num_(EU7_get_(row, idx, "CadaDias", "IntervaloDias"));

    let tipo = control;
    if (!tipo){
      if (cadaKm && !cadaDias) tipo = "km";
      else if (!cadaKm && cadaDias) tipo = "dia";
      else if (cadaKm) tipo = "km";
      else tipo = "dia";
    }
    if (tipo === "dias") tipo = "dia";

    // snapshot antes
    const antes = {
      UltimoKm: EU7_get_(row, idx, "UltimoKm"),
      UltimaFecha: EU7_get_(row, idx, "UltimaFecha"),
      ProximoKm: EU7_get_(row, idx, "ProximoKm"),
      ProximaFecha: EU7_get_(row, idx, "ProximaFecha")
    };

    if (tipo === "km"){
      const n = parseInt(String(nuevo).replace(/[^\d]/g,""), 10);
      if (!isFinite(n) || n < 0) return { ok:false, msg:"Nuevo Km inválido." };

      EU7_set_(row, idx, "UltimoKm", n);
      if (cadaKm){
        EU7_set_(row, idx, "ProximoKm", n + cadaKm);
      }

    } else {
      // nuevo puede venir como "YYYY-MM-DD" o Date
      let d = null;
      if (nuevo instanceof Date) d = nuevo;
      else {
        const s = String(nuevo || "").trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return { ok:false, msg:"Nueva fecha inválida. Usá YYYY-MM-DD" };
        d = new Date(s + "T00:00:00");
      }
      if (!d || isNaN(d.getTime())) return { ok:false, msg:"Nueva fecha inválida." };

      EU7_set_(row, idx, "UltimaFecha", d);
      if (cadaDias){
        EU7_set_(row, idx, "ProximaFecha", EU7_addDays_(d, cadaDias));
      }
    }

    // auditoría
    EU7_set_(row, idx, "Usuario", user?.Usuario || user?.u || "");
    EU7_set_(row, idx, "Timestamp", new Date());
    EU7_set_(row, idx, "UltimaAccion", "REPROG");

    // guardar fila
    shPU.getRange(rowN+1, 1, 1, headers.length).setValues([row]);

    // historial en PreventivosUnidadHis si existe
    const shH = ss.getSheetByName(EU7_SHEET_HIS);
    if (shH){
      const hVals = shH.getDataRange().getValues();
      const hHeaders = hVals[0] || [];
      const hIdx = EU7_buildIndex_(hHeaders);

      const nueva = new Array(hHeaders.length).fill("");

      EU7_set_(nueva, hIdx, "IdHis", EU7_makeId_());
      EU7_set_(nueva, hIdx, "Interno", interno);
      EU7_set_(nueva, hIdx, "IdHP", idHP);
      EU7_set_(nueva, hIdx, "NombreHP", EU7_get_(row, idx, "NombreHP", "Nombre Preventivo", "Preventivo"));
      EU7_set_(nueva, hIdx, "Accion", "REPROG");
      EU7_set_(nueva, hIdx, "Antes", JSON.stringify(antes));
      EU7_set_(nueva, hIdx, "Despues", JSON.stringify({
        UltimoKm: EU7_get_(row, idx, "UltimoKm"),
        UltimaFecha: EU7_get_(row, idx, "UltimaFecha"),
        ProximoKm: EU7_get_(row, idx, "ProximoKm"),
        ProximaFecha: EU7_get_(row, idx, "ProximaFecha")
      }));
      EU7_set_(nueva, hIdx, "Motivo", motivo);
      EU7_set_(nueva, hIdx, "Usuario", user?.Usuario || user?.u || "");
      EU7_set_(nueva, hIdx, "Timestamp", new Date());

      shH.appendRow(nueva);
    }

    return { ok:true };

  } catch(err){
    return { ok:false, msg: (err && err.message) ? err.message : String(err) };
  }
}
i faltan, se agregan)
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

