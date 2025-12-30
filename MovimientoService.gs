function _getHeadersMov_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || "").toString().trim());
}

function _ensureMovSchema_() {
  const sh = _sheet(SHEET_MOV);
  const headers = _getHeadersMov_(sh);

  if (headers.length === 0) {
    sh.getRange(1,1,1,MOV_COLS_REQUIRED.length).setValues([MOV_COLS_REQUIRED]);
    return;
  }

  const missing = MOV_COLS_REQUIRED.filter(c => !headers.includes(c));
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}
function _fmtHoraHHMM_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");

  if (typeof v === "number" && isFinite(v)) {
    const totalMin = Math.round(v * 24 * 60);
    const hh = String(Math.floor(totalMin / 60) % 24).padStart(2, "0");
    const mm = String(totalMin % 60).padStart(2, "0");
    return `${hh}:${mm}`;
  }

  const s = v.toString().trim();
  const m = /^(\d{1,2}):(\d{2})$/.exec(s);
  if (m) return `${String(m[1]).padStart(2,"0")}:${m[2]}`;

  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  return s;
}

function _parseHoraToMinutes_(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec((hhmm || "").toString().trim());
  if (!m) return null;
  const hh = Number(m[1]), mm = Number(m[2]);
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  return hh * 60 + mm;
}

function _lastOdometerForUnidad_(unidad) {
  _ensureMovSchema_();
  const sh = _sheet(SHEET_MOV);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { last: 0, row: null, ts: null };

  const headers = values[0].map(h => (h || "").toString().trim());
  const idx = (name) => headers.indexOf(name);

  const iUnidad = idx("Unidad");
  const iOdo = idx("Odómetro");
  const iTS = idx("Timestamp");

  let best = { last: 0, row: null, ts: null };

  for (let r = 1; r < values.length; r++) {
    const u = (iUnidad === -1 ? "" : (values[r][iUnidad] || "").toString().trim());
    if (!u || u.toLowerCase() !== unidad.toLowerCase()) continue;

    const odo = Number(iOdo === -1 ? 0 : (values[r][iOdo] || 0));
    const tsVal = (iTS === -1 ? null : values[r][iTS]);

    let ts = null;
    if (tsVal instanceof Date) ts = tsVal.getTime();
    else if (tsVal) {
      const d = new Date(tsVal);
      if (!isNaN(d.getTime())) ts = d.getTime();
    }

    if (ts !== null) {
      if (best.ts === null || ts > best.ts) best = { last: odo, row: r + 1, ts };
    } else {
      // fallback
      best = { last: odo, row: r + 1, ts: best.ts };
    }
  }

  return best;
}

function _updateChasisKm_(unidad, newOdo, user) {
  _ensureChasisSchema_();
  const sh = _sheet(SHEET_CHASIS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => (h || "").toString().trim());
  const iInterno = headers.indexOf("Interno");
  const iKm = headers.indexOf("KmRecorridos");
  const iUM = headers.indexOf("Ultima Modificacion");
  const iU = headers.indexOf("Usuario");

  if (iInterno === -1 || iKm === -1) return;

  for (let r = 1; r < values.length; r++) {
    const interno = (values[r][iInterno] || "").toString().trim();
    if (interno && interno.toLowerCase() === unidad.toLowerCase()) {
      sh.getRange(r + 1, iKm + 1).setValue(Number(newOdo));
      if (iUM !== -1) sh.getRange(r + 1, iUM + 1).setValue(new Date());
      if (iU !== -1) sh.getRange(r + 1, iU + 1).setValue(user.u || "");
      return;
    }
  }
}

// ====== FIX robusto listado (no se rompe si falta columna) ======
function listMovimientos(token) {
  _requireSession_(token);
  _ensureMovSchema_();

  const sh = _sheet(SHEET_MOV);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok:true, rows: [] };

  const headers = values[0].map(h => (h || "").toString().trim());
  const idx = (name) => headers.indexOf(name);

  const iId = idx("IdMov");
  const iUnidad = idx("Unidad");
  const iTipo = idx("Tipo");
  const iFecha = idx("FechaMov");
  const iHora = idx("HoraMov");
  const iOdo = idx("Odómetro");
  const iUlt = idx("UltimoOdometro");
  const iKm = idx("KmRecorridos");
  const iObs = idx("Observacion");
  const iDep = idx("Deposito");
  const iUser = idx("Usuario");
  const iTS = idx("Timestamp");

  const get = (row, i) => (i === -1 ? "" : row[i]);

  const rows = values.slice(1).map((r, k) => {
    const rowNum = k + 2;

    const fechaVal = get(r, iFecha);
    const tsVal = get(r, iTS);

    const fechaMs = (fechaVal instanceof Date) ? fechaVal.getTime() : (fechaVal ? new Date(fechaVal).getTime() : null);
    const tsMs = (tsVal instanceof Date) ? tsVal.getTime() : (tsVal ? new Date(tsVal).getTime() : null);

    return {
      _row: rowNum,
      IdMov: (get(r, iId) || "").toString(),
      Unidad: (get(r, iUnidad) || "").toString(),
      Tipo: (get(r, iTipo) || "").toString().toLowerCase(),
      FechaMovMs: isNaN(fechaMs) ? null : fechaMs,
      HoraMov: _fmtHoraHHMM_(get(r, iHora)),
      Odometro: Number(get(r, iOdo) || 0),
      UltimoOdometro: Number(get(r, iUlt) || 0),
      KmRecorridos: Number(get(r, iKm) || 0),
      Observacion: (get(r, iObs) || "").toString(),
      Deposito: (get(r, iDep) || "").toString(),
      Usuario: (get(r, iUser) || "").toString(),
      TimestampMs: isNaN(tsMs) ? null : tsMs
    };
  });

  rows.sort((a,b) => {
    const ta = a.TimestampMs ?? -1;
    const tb = b.TimestampMs ?? -1;
    if (ta !== tb) return tb - ta;
    return b._row - a._row;
  });

  return { ok:true, rows };
}

// ====== endpoint PRO: opciones de Unidad/Deposito ======
function _distinctFromCol_(sheetName, colName) {
  const sh = _sheet(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => (h || "").toString().trim());
  const i = headers.indexOf(colName);
  if (i === -1) return [];

  const vals = values.slice(1).map(r => (r[i] ?? "").toString().trim()).filter(Boolean);
  return [...new Set(vals)].sort((a,b) => a.localeCompare(b,"es"));
}

function _distinctFromSheetColAIfExists_(sheetName) {
  const ss = _ss();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return null;
  const last = sh.getLastRow();
  if (last < 1) return [];
  const vals = sh.getRange(1,1,last,1).getValues().flat().map(v => (v??"").toString().trim()).filter(Boolean);
  return [...new Set(vals)].sort((a,b)=>a.localeCompare(b,"es"));
}

function getMovSelectOptions(token) {
  _requireSession_(token);
  _ensureChasisSchema_();
  _ensureMovSchema_();

  // Unidades: de ChasisBD
  const ch = _sheet(SHEET_CHASIS);
  const cvals = ch.getDataRange().getValues();

  let unidades = [];
  if (cvals.length >= 2) {
    const h = cvals[0].map(x => (x||"").toString().trim());
    const iInt = h.indexOf("Interno");
    const iDom = h.indexOf("Dominio");
    const iSoc = h.indexOf("Sociedad");
    const iDep = h.indexOf("Deposito");
    const iEst = h.indexOf("Estado");

    unidades = cvals.slice(1).map(r => {
      const interno = (iInt===-1? "" : (r[iInt]||"").toString().trim());
      if (!interno) return null;
      const dominio = (iDom===-1? "" : (r[iDom]||"").toString().trim());
      const sociedad = (iSoc===-1? "" : (r[iSoc]||"").toString().trim());
      const deposito = (iDep===-1? "" : (r[iDep]||"").toString().trim());
      const estado = (iEst===-1? "" : (r[iEst]||"").toString().trim().toLowerCase());
      // si querés, podés filtrar solo activos:
      // if (estado && estado !== "activo") return null;

      const label = [interno, dominio, sociedad].filter(Boolean).join(" — ");
      return { value: interno, label, deposito };
    }).filter(Boolean);

    const seen = new Set();
    unidades = unidades.filter(u => (seen.has(u.value) ? false : (seen.add(u.value), true)));
    unidades.sort((a,b) => a.label.localeCompare(b.label,"es"));
  }

  // Depósitos:
  let depositos = _distinctFromSheetColAIfExists_("Depositos");
  if (depositos === null) {
    const d1 = _distinctFromCol_(SHEET_CHASIS, "Deposito");
    const d2 = _distinctFromCol_(SHEET_MOV, "Deposito");
    depositos = [...new Set([...d1, ...d2])].sort((a,b)=>a.localeCompare(b,"es"));
  }

  return { ok:true, unidades, depositos };
}

// ====== endpoint: ultimo odometro al seleccionar unidad ======
function getUltimoOdometer(token, unidad) {
  _requireSession_(token);
  unidad = (unidad || "").toString().trim();
  if (!unidad) return { ok:true, last: 0 };
  const last = _lastOdometerForUnidad_(unidad).last || 0;
  return { ok:true, last: Number(last) };
}

// ====== agregar movimiento ======
function addMovimiento(token, payload) {
  const user = _requireSession_(token);
  _ensureMovSchema_();
  _ensureChasisSchema_();

  const Unidad = (payload.Unidad || "").toString().trim();
  const Tipo = (payload.Tipo || "").toString().trim().toLowerCase();
  const Deposito = (payload.Deposito || "").toString().trim();
  const Observacion = (payload.Observacion || "").toString().trim();

  const FechaISO = (payload.FechaISO || "").toString().trim(); // yyyy-mm-dd
  const HoraMov = (payload.HoraMov || "").toString().trim();   // hh:mm

  const Odometro = Number(payload.Odometro);

  if (!Unidad) throw new Error("Unidad (Interno) es obligatorio.");
  if (!["ingreso","egreso"].includes(Tipo)) throw new Error("Tipo debe ser ingreso o egreso.");
  if (!Deposito) throw new Error("Depósito es obligatorio.");
  if (!FechaISO) throw new Error("Fecha es obligatoria.");
  if (_parseHoraToMinutes_(HoraMov) === null) throw new Error("Hora inválida. Usá hh:mm.");
  if (!isFinite(Odometro) || Odometro <= 0) throw new Error("Odómetro inválido.");

  const d = new Date(FechaISO + "T00:00:00");
  if (isNaN(d.getTime())) throw new Error("Fecha inválida.");

  const lastInfo = _lastOdometerForUnidad_(Unidad);
  const UltimoOdometro = Number(lastInfo.last || 0);

  if (Odometro < UltimoOdometro) {
    throw new Error(`No se puede cargar un odómetro menor al último. Último: ${UltimoOdometro}`);
  }

  const KmRecorridos = Odometro - UltimoOdometro;
  const now = new Date();
  const IdMov = "MOV-" + Utilities.getUuid().slice(0,8);

  const sh = _sheet(SHEET_MOV);
  const headers = _getHeadersMov_(sh);
  const idx = (name) => headers.indexOf(name);

  const row = new Array(headers.length).fill("");
  const setIf = (col, val) => { const c = idx(col); if (c !== -1) row[c] = val; };

  setIf("IdMov", IdMov);
  setIf("Unidad", Unidad);
  setIf("Tipo", Tipo);
  setIf("FechaMov", d);
  setIf("HoraMov", HoraMov);
  setIf("Odómetro", Odometro);
  setIf("UltimoOdometro", UltimoOdometro);
  setIf("KmRecorridos", KmRecorridos);
  setIf("Observacion", Observacion);
  setIf("Deposito", Deposito);
  setIf("Usuario", user.u || "");
  setIf("Timestamp", now);

  sh.appendRow(row);

  // Actualiza ChasisBD.KmRecorridos con el nuevo odómetro
  _updateChasisKm_(Unidad, Odometro, user);

  return { ok:true, msg:"Movimiento guardado.", IdMov, UltimoOdometro, KmRecorridos };
}
