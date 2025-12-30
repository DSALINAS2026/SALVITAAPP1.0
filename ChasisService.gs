function _getHeaders_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || "").toString().trim());
}

function _ensureChasisSchema_() {
  const sh = _sheet(SHEET_CHASIS);
  const headers = _getHeaders_(sh);
  if (headers.length === 0) {
    sh.getRange(1,1,1,CHASIS_COLS_REQUIRED.length).setValues([CHASIS_COLS_REQUIRED]);
    return;
  }
  const missing = CHASIS_COLS_REQUIRED.filter(c => !headers.includes(c));
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function listChasis(token) {
  const user = _requireSession_(token);
  _ensureChasisSchema_();

  const sh = _sheet(SHEET_CHASIS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return {
      ok:true,
      isAdmin:_isAdmin_(user),
      summary:{ total:0, activo:0, inactivo:0, baja:0 },
      rows:[]
    };
  }

  const headers = values[0].map(h => (h || "").toString().trim());
  const idx = (name) => headers.indexOf(name);

  const iInterno = idx("Interno");
  const iDominio = idx("Dominio");
  const iSociedad = idx("Sociedad");
  const iDeposito = idx("Deposito");
  const iTipo = idx("Tipo");
  const iKm = idx("KmRecorridos");
  const iEstado = idx("Estado");
  const iId = idx("IdChasis");

  const rows = [];
  let a=0, i=0, b=0;

  values.slice(1).forEach((r, rowIndex0) => {
    const estado = ((r[iEstado] || "") + "").trim().toLowerCase() || "activo";
    if (estado === "activo") a++;
    else if (estado === "inactivo") i++;
    else if (estado === "baja") b++;

    rows.push({
      _row: rowIndex0 + 2,
      IdChasis: (r[iId] || "").toString(),
      Interno: (r[iInterno] || "").toString(),
      Dominio: (r[iDominio] || "").toString(),
      Sociedad: (r[iSociedad] || "").toString(),
      Deposito: (r[iDeposito] || "").toString(),
      Tipo: (r[iTipo] || "").toString(),
      KmRecorridos: (r[iKm] || "").toString(),
      Estado: estado,

      CapacidadCarga: (r[idx("CapacidadCarga")] || "").toString(),
      "Nro. Chasis": (r[idx("Nro. Chasis")] || "").toString(),
      Marca: (r[idx("Marca")] || "").toString(),
      Modelo: (r[idx("Modelo")] || "").toString(),
      Motor: (r[idx("Motor")] || "").toString(),
      "Nro Motor": (r[idx("Nro Motor")] || "").toString(),
      Eje: (r[idx("Eje")] || "").toString(),
      "Mapa Cubierta": (r[idx("Mapa Cubierta")] || "").toString(),
      Carroceria: (r[idx("Carroceria")] || "").toString(),
      "Año": (r[idx("Año")] || "").toString()
    });
  });

  return {
    ok:true,
    isAdmin:_isAdmin_(user),
    summary:{ total: rows.length, activo:a, inactivo:i, baja:b },
    rows
  };
}

function addChasis(token, payload) {
  const user = _requireSession_(token);
  if (!_isAdmin_(user)) throw new Error("Solo administradores pueden agregar chasis.");

  _ensureChasisSchema_();
  const sh = _sheet(SHEET_CHASIS);
  const headers = _getHeaders_(sh);
  const idx = (name) => headers.indexOf(name);

  const Interno = (payload.Interno || "").toString().trim();
  const Dominio = (payload.Dominio || "").toString().trim();
  const Sociedad = (payload.Sociedad || "").toString().trim();
  const Deposito = (payload.Deposito || "").toString().trim();
  const Tipo = (payload.Tipo || "").toString().trim();
  const KmRecorridos = (payload.KmRecorridos || "0").toString().trim();
  const Estado = ((payload.Estado || "activo") + "").toLowerCase();

  const CapacidadCarga = (payload.CapacidadCarga || "").toString().trim();
  const NroChasis = (payload["Nro. Chasis"] || "").toString().trim();
  const Marca = (payload.Marca || "").toString().trim();
  const Modelo = (payload.Modelo || "").toString().trim();
  const Motor = (payload.Motor || "").toString().trim();
  const NroMotor = (payload["Nro Motor"] || payload["NroMotor"] || "").toString().trim();
  const Eje = (payload.Eje || "").toString().trim();
  const MapaCubierta = (payload["Mapa Cubierta"] || "").toString().trim();
  const Carroceria = (payload.Carroceria || "").toString().trim();
  const Anio = (payload["Año"] || payload.Anio || "").toString().trim();

  if (!Interno) throw new Error("Interno es obligatorio.");
  if (!Dominio) throw new Error("Dominio/Patente es obligatorio.");

  const data = sh.getDataRange().getValues();
  const iInterno = idx("Interno");
  if (data.slice(1).some(r => (r[iInterno] || "").toString().trim().toLowerCase() === Interno.toLowerCase())) {
    throw new Error("Ya existe un chasis con ese Interno.");
  }

  const newRow = new Array(headers.length).fill("");
  const id = (payload.IdChasis || ("CH-" + Utilities.getUuid().slice(0,8))).toString();
  const now = new Date();

  function setIfExists(colName, value) {
    const c = idx(colName);
    if (c !== -1) newRow[c] = value;
  }

  setIfExists("IdChasis", id);
  setIfExists("Interno", Interno);
  setIfExists("Dominio", Dominio);
  setIfExists("Sociedad", Sociedad);
  setIfExists("Deposito", Deposito);
  setIfExists("Tipo", Tipo);
  setIfExists("KmRecorridos", KmRecorridos);
  setIfExists("Estado", ["activo","inactivo","baja"].includes(Estado) ? Estado : "activo");

  setIfExists("CapacidadCarga", CapacidadCarga);
  setIfExists("Nro. Chasis", NroChasis);
  setIfExists("Marca", Marca);
  setIfExists("Modelo", Modelo);
  setIfExists("Motor", Motor);
  setIfExists("Nro Motor", NroMotor);
  setIfExists("Eje", Eje);
  setIfExists("Mapa Cubierta", MapaCubierta);
  setIfExists("Carroceria", Carroceria);
  setIfExists("Año", Anio);

  // NO visibles
  setIfExists("Val.", (payload["Val."] ?? "").toString());
  setIfExists("Usuario", user.u || "");
  setIfExists("Fecha", now);
  setIfExists("Ultima Modificacion", now);

  sh.appendRow(newRow);
  return { ok:true, msg:"Chasis agregado.", id };
}

function setChasisEstado(token, rowNumber, estado) {
  const user = _requireSession_(token);
  if (!_isAdmin_(user)) throw new Error("Solo administradores pueden cambiar el estado.");

  _ensureChasisSchema_();
  const sh = _sheet(SHEET_CHASIS);
  const headers = _getHeaders_(sh);
  const iEstado = headers.indexOf("Estado");
  if (iEstado === -1) throw new Error("No existe la columna Estado.");

  estado = ((estado || "") + "").toLowerCase().trim();
  if (!["activo","inactivo","baja"].includes(estado)) throw new Error("Estado inválido.");

  sh.getRange(Number(rowNumber), iEstado + 1).setValue(estado);

  // Ultima modificación + usuario
  const colUM = headers.indexOf("Ultima Modificacion");
  if (colUM !== -1) sh.getRange(Number(rowNumber), colUM + 1).setValue(new Date());
  const colU = headers.indexOf("Usuario");
  if (colU !== -1) sh.getRange(Number(rowNumber), colU + 1).setValue(user.u || "");

  return { ok:true };
}

function updateChasisFields(token, rowNumber, fields) {
  const user = _requireSession_(token);
  if (!_isAdmin_(user)) throw new Error("Solo administradores pueden editar.");

  _ensureChasisSchema_();
  const sh = _sheet(SHEET_CHASIS);
  const headers = _getHeaders_(sh);

  const allowed = [
    "Sociedad","Deposito","Interno","Dominio","KmRecorridos","Tipo",
    "CapacidadCarga","Nro. Chasis","Marca","Modelo","Motor","Nro Motor",
    "Eje","Mapa Cubierta","Carroceria","Año",
    "Estado"
  ];

  const row = Number(rowNumber);

  allowed.forEach(k => {
    if (Object.prototype.hasOwnProperty.call(fields, k)) {
      const col = headers.indexOf(k);
      if (col !== -1) sh.getRange(row, col + 1).setValue((fields[k] ?? "").toString());
    }
  });

  // Ultima modificación + usuario
  const now = new Date();
  const colUM = headers.indexOf("Ultima Modificacion");
  if (colUM !== -1) sh.getRange(row, colUM + 1).setValue(now);

  const colU = headers.indexOf("Usuario");
  if (colU !== -1) sh.getRange(row, colU + 1).setValue(user.u || "");

  return { ok:true };
}

// ====== SELECT OPTIONS (Eje / Mapa Cubierta / Carroceria) ======

function _getColumnValuesDistinct_(sh, colName) {
  const headers = _getHeaders_(sh);
  const col = headers.indexOf(colName);
  if (col === -1) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, col + 1, lastRow - 1, 1).getValues()
    .flat()
    .map(v => (v ?? "").toString().trim())
    .filter(v => v);

  return [...new Set(vals)].sort((a,b) => a.localeCompare(b, "es"));
}

/**
 * Devuelve opciones para selects.
 * Si existe hoja "Eje" / "Mapa Cubierta" / "Carroceria" (col A), la usa.
 * Si no existe, toma valores únicos desde ChasisBD.
 */
function getSelectOptions(token) {
  _requireSession_(token);
  const ss = _ss();
  const ch = _sheet(SHEET_CHASIS);

  function fromSheetOrChasis(sheetName, colName) {
    const sh = ss.getSheetByName(sheetName);
    if (sh && sh.getLastRow() >= 1) {
      const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues()
        .flat()
        .map(v => (v ?? "").toString().trim())
        .filter(v => v);
      return [...new Set(vals)].sort((a,b) => a.localeCompare(b, "es"));
    }
    return _getColumnValuesDistinct_(ch, colName);
  }

  return {
    ok: true,
    ejes: fromSheetOrChasis("Eje", "Eje"),
    mapas: fromSheetOrChasis("Mapa Cubierta", "Mapa Cubierta"),
    carrocerias: fromSheetOrChasis("Carroceria", "Carroceria")
  };
}

