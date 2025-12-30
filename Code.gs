const SHEET_TAREAS = "MaestroTareas";
const TAREAS_COLS_REQUIRED = [
  "idTarea","codigo","nombre","usuario","borrado","sistema","subsistema","sector","fecha"
];

function _getHeadersTareas_(sh){
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1,1,1,lastCol).getValues()[0].map(h => (h||"").toString().trim());
}

function _ensureTareasSchema_(){
  const sh = _sheet(SHEET_TAREAS);
  const headers = _getHeadersTareas_(sh);

  if (headers.length === 0){
    sh.getRange(1,1,1,TAREAS_COLS_REQUIRED.length).setValues([TAREAS_COLS_REQUIRED]);
    return;
  }

  const missing = TAREAS_COLS_REQUIRED.filter(c => !headers.includes(c));
  if (missing.length){
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function _isBorradoVal_(v){
  const s = (v ?? "").toString().trim().toLowerCase();
  return (v === true || s === "true" || s === "1" || s === "si" || s === "sí");
}

function _distinctFromTareas_(colName){
  const sh = _sheet(SHEET_TAREAS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => (h||"").toString().trim());
  const i = headers.indexOf(colName);
  const iBor = headers.indexOf("borrado");
  if (i === -1) return [];

  const vals = values.slice(1)
    .filter(r => !_isBorradoVal_(iBor === -1 ? false : r[iBor]))
    .map(r => (r[i] ?? "").toString().trim())
    .filter(Boolean);

  return [...new Set(vals)].sort((a,b)=>a.localeCompare(b,"es"));
}

function listTareas(token){
  _requireSession_(token);
  _ensureTareasSchema_();

  const sh = _sheet(SHEET_TAREAS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok:true, rows: [] };

  const headers = values[0].map(h => (h||"").toString().trim());
  const idx = (name) => headers.indexOf(name);
  const get = (row,i) => (i === -1 ? "" : row[i]);

  const iId = idx("idTarea");
  const iCod = idx("codigo");
  const iNom = idx("nombre");
  const iSis = idx("sistema");
  const iSub = idx("subsistema");
  const iSec = idx("sector");
  const iUsr = idx("usuario");
  const iBor = idx("borrado");
  const iFec = idx("fecha");

  const rows = values.slice(1).map((r,k) => {
    const rowNum = k + 2;

    const borrado = _isBorradoVal_(get(r,iBor)) ? true : false;
    const fecVal = get(r,iFec);
    const fecMs = (fecVal instanceof Date) ? fecVal.getTime() : (fecVal ? new Date(fecVal).getTime() : null);

    return {
      _row: rowNum,
      idTarea: (get(r,iId)||"").toString(),
      codigo: (get(r,iCod)||"").toString(),
      nombre: (get(r,iNom)||"").toString(),
      sistema: (get(r,iSis)||"").toString(),
      subsistema: (get(r,iSub)||"").toString(),
      sector: (get(r,iSec)||"").toString(),
      usuario: (get(r,iUsr)||"").toString(),
      borrado,
      fechaMs: isNaN(fecMs) ? null : fecMs
    };
  }).filter(x => x.borrado === false);

  rows.sort((a,b) => (a.codigo||"").localeCompare(b.codigo||"","es") || (a.nombre||"").localeCompare(b.nombre||"","es"));

  return { ok:true, rows };
}

function listTareasIncluyendoEliminadas(token){
  _requireSession_(token);
  _ensureTareasSchema_();

  const sh = _sheet(SHEET_TAREAS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok:true, rows: [] };

  const headers = values[0].map(h => (h||"").toString().trim());
  const idx = (name) => headers.indexOf(name);
  const get = (row,i) => (i === -1 ? "" : row[i]);

  const iId = idx("idTarea");
  const iCod = idx("codigo");
  const iNom = idx("nombre");
  const iSis = idx("sistema");
  const iSub = idx("subsistema");
  const iSec = idx("sector");
  const iUsr = idx("usuario");
  const iBor = idx("borrado");
  const iFec = idx("fecha");

  const rows = values.slice(1).map((r,k) => {
    const rowNum = k + 2;

    const borrado = _isBorradoVal_(get(r,iBor)) ? true : false;
    const fecVal = get(r,iFec);
    const fecMs = (fecVal instanceof Date) ? fecVal.getTime() : (fecVal ? new Date(fecVal).getTime() : null);

    return {
      _row: rowNum,
      idTarea: (get(r,iId)||"").toString(),
      codigo: (get(r,iCod)||"").toString(),
      nombre: (get(r,iNom)||"").toString(),
      sistema: (get(r,iSis)||"").toString(),
      subsistema: (get(r,iSub)||"").toString(),
      sector: (get(r,iSec)||"").toString(),
      usuario: (get(r,iUsr)||"").toString(),
      borrado,
      fechaMs: isNaN(fecMs) ? null : fecMs
    };
  });

  rows.sort((a,b) => (a.codigo||"").localeCompare(b.codigo||"","es") || (a.nombre||"").localeCompare(b.nombre||"","es"));

  return { ok:true, rows };
}

function getTareasSelectOptions(token){
  _requireSession_(token);
  _ensureTareasSchema_();
  return {
    ok:true,
    sistemas: _distinctFromTareas_("sistema"),
    subsistemas: _distinctFromTareas_("subsistema"),
    sectores: _distinctFromTareas_("sector"),
  };
}

function addTarea(token, payload){
  const user = _requireSession_(token);
  _ensureTareasSchema_();

  const codigo = (payload.codigo || "").toString().trim();
  const nombre = (payload.nombre || "").toString().trim();
  const sistema = (payload.sistema || "").toString().trim();
  const subsistema = (payload.subsistema || "").toString().trim();
  const sector = (payload.sector || "").toString().trim();

  if (!codigo) throw new Error("Código es obligatorio.");
  if (!nombre) throw new Error("Nombre es obligatorio.");
  if (!sistema) throw new Error("Sistema es obligatorio.");
  if (!subsistema) throw new Error("Subsistema es obligatorio.");
  if (!sector) throw new Error("Sector es obligatorio.");

  const sh = _sheet(SHEET_TAREAS);
  const values = sh.getDataRange().getValues();

  if (values.length >= 2){
    const h = values[0].map(x => (x||"").toString().trim());
    const iCod = h.indexOf("codigo");
    const iBor = h.indexOf("borrado");
    if (iCod !== -1){
      const exists = values.slice(1).some(r => {
        const cod = (r[iCod]||"").toString().trim().toLowerCase();
        const bor = (iBor === -1 ? false : r[iBor]);
        return !_isBorradoVal_(bor) && cod === codigo.toLowerCase();
      });
      if (exists) throw new Error("Ya existe una tarea con ese código (no borrada).");
    }
  }

  const headers = _getHeadersTareas_(sh);
  const idx = (name) => headers.indexOf(name);
  const setIf = (row, col, val) => { const i = idx(col); if (i !== -1) row[i] = val; };

  const idTarea = "TAR-" + Utilities.getUuid().slice(0,8);
  const now = new Date();

  const row = new Array(headers.length).fill("");
  setIf(row, "idTarea", idTarea);
  setIf(row, "codigo", codigo);
  setIf(row, "nombre", nombre);
  setIf(row, "sistema", sistema);
  setIf(row, "subsistema", subsistema);
  setIf(row, "sector", sector);

  setIf(row, "usuario", user.u || "");
  setIf(row, "borrado", false);
  setIf(row, "fecha", now);

  sh.appendRow(row);

  return { ok:true, idTarea };
}

function borrarTarea(token, row){
  const user = _requireSession_(token);
  _ensureTareasSchema_();

  row = Number(row);
  if (!isFinite(row) || row < 2) throw new Error("Fila inválida.");

  const sh = _sheet(SHEET_TAREAS);
  const headers = _getHeadersTareas_(sh);
  const iBor = headers.indexOf("borrado");
  const iUsr = headers.indexOf("usuario");
  const iFec = headers.indexOf("fecha");

  if (iBor === -1) throw new Error("Falta columna borrado.");

  sh.getRange(row, iBor+1).setValue(true);
  if (iUsr !== -1) sh.getRange(row, iUsr+1).setValue(user.u || "");
  if (iFec !== -1) sh.getRange(row, iFec+1).setValue(new Date());

  return { ok:true };
}
