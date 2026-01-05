// ===================== ESTADO UNIDAD SERVICE (SAFE PATCH) =====================
// Este archivo NO declara SPREADSHEET_ID para evitar "already been declared".
// Usa nombres únicos EU7_* para no chocar con otros módulos.

// ---- Helpers ----
function EU7_ss_(){
  // 1) Si existe SPREADSHEET_ID global (definido en Code.gs), úsalo.
  try{
    if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  }catch(e){}
  // 2) Si existe propiedad guardada
  try{
    const pid = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (pid) return SpreadsheetApp.openById(pid);
  }catch(e){}
  // 3) Fallback
  return SpreadsheetApp.getActiveSpreadsheet();
}

function EU7_norm_(s){
  s = (s ?? '').toString().trim().toLowerCase();
  try{
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g,''); // sin acentos
  }catch(e){}
  s = s.replace(/\s+/g,' ');
  return s;
}
function EU7_idx_(headers){
  const idx = {};
  headers.forEach((h,i)=>{ idx[EU7_norm_(h)] = i; });
  return idx;
}
function EU7_get_(row, idx /*, names... */){
  for (let i=2;i<arguments.length;i++){
    const k = EU7_norm_(arguments[i]);
    if (k in idx){
      const v = row[idx[k]];
      if (v !== '' && v !== null && typeof v !== 'undefined') return v;
    }
  }
  return '';
}
function EU7_set_(row, idx, name, value){
  const k = EU7_norm_(name);
  if (k in idx) row[idx[k]] = value;
}
function EU7_sheet_(ss, names){
  for (const n of names){
    const sh = ss.getSheetByName(n);
    if (sh) return sh;
  }
  return null;
}
function EU7_toNum_(v){
  const n = Number(v);
  return isFinite(n) ? n : 0;
}
function EU7_toDate_(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  // intenta parse DD/MM/YYYY o YYYY-MM-DD
  const s = v.toString().trim();
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m){
    const dd = Number(m[1]), mm = Number(m[2]) - 1, yy = Number(m[3].length===2?('20'+m[3]):m[3]);
    const d = new Date(yy, mm, dd);
    if (!isNaN(d.getTime())) return d;
  }
  const d2 = new Date(s);
  if (!isNaN(d2.getTime())) return d2;
  return null;
}
function EU7_fmtDate_(d){
  if (!d) return '';
  if (!(d instanceof Date)) d = EU7_toDate_(d);
  if (!d) return '';
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yy = d.getFullYear();
  return `${dd}-${mm}-${yy}`;
}
function EU7_daysBetween_(a,b){
  const ms = 24*60*60*1000;
  const da = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  const db = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((db - da)/ms);
}

function EU7_requireAuth_(token){
  // Si existe validateToken o Auth similar, úsalo. Si no existe, deja pasar.
  try{
    if (typeof validateToken === 'function') {
      const r = validateToken(token);
      if (r && r.ok === false) return r;
    }
    if (typeof Auth_validateToken === 'function') {
      const r2 = Auth_validateToken(token);
      if (r2 && r2.ok === false) return r2;
    }
  }catch(e){
    return {ok:false, msg:'Token inválido', error:String(e)};
  }
  return {ok:true};
}

// ---- API: Estado Unidad ----
function getEstadoUnidadV7(token, interno){
  try{
    const auth = EU7_requireAuth_(token);
    if (!auth.ok) return auth;

    interno = (interno ?? '').toString().trim();
    if (!interno) return {ok:false, msg:'Interno requerido'};

    const ss = EU7_ss_();

    const shCh = EU7_sheet_(ss, ['ChasisBD']);
    const shPU = EU7_sheet_(ss, ['PreventivosUnidad']);
    const shOT = EU7_sheet_(ss, ['ordenesTrabajo','OrdenesTrabajo']);

    if (!shPU) return {ok:false, msg:'No existe hoja PreventivosUnidad'};

    // --- Chasis info ---
    let chasis = { Interno: interno };
    let kmActual = 0;
    if (shCh){
      const v = shCh.getDataRange().getValues();
      const h = v.shift();
      const ix = EU7_idx_(h);
      const row = v.find(r => (EU7_get_(r, ix, 'Interno')+'') === interno);
      if (row){
        chasis = {
          Interno: interno,
          Dominio: EU7_get_(row, ix, 'Dominio'),
          Sociedad: EU7_get_(row, ix, 'Sociedad'),
          Deposito: EU7_get_(row, ix, 'Deposito'),
          Marca: EU7_get_(row, ix, 'Marca'),
          'Nro Chasis': EU7_get_(row, ix, 'Nro. Chasis','Nro Chasis','NroChasis'),
          'Nro Motor': EU7_get_(row, ix, 'Nro Motor','Nro. Motor','Motor'),
          KmRecorridos: EU7_toNum_(EU7_get_(row, ix, 'KmRecorridos','Km Recorridos','Km'))
        };
        kmActual = EU7_toNum_(chasis.KmRecorridos);
      }
    }

    // --- OTs pendientes index (Interno+IdHP -> {nroOT,idOT,estado}) ---
    const otPend = {};
    if (shOT){
      const vv = shOT.getDataRange().getValues();
      const hh = vv.shift();
      const ox = EU7_idx_(hh);

      const iInterno = EU7_norm_('Interno');
      const iIdHP = EU7_norm_('IdHP');
      const iTipoOT = EU7_norm_('TipoOT');
      const iEstado = EU7_norm_('EstadoOT');
      const iNroOT = EU7_norm_('NroOT');
      const iIdOT  = EU7_norm_('IdOT');
      const iTS    = EU7_norm_('Timestamp');

      vv.forEach(r=>{
        const inx = (ox[iInterno]!=null)? r[ox[iInterno]] : '';
        if ((inx+'') !== interno) return;
        const tipo = (ox[iTipoOT]!=null)? (r[ox[iTipoOT]]+'') : '';
        if ((tipo+'').toLowerCase() !== 'preventiva') return;
        const est = (ox[iEstado]!=null)? (r[ox[iEstado]]+'').toLowerCase() : '';
        if (['anulada','confirmada','cerrada'].includes(est)) return;

        const idhp = (ox[iIdHP]!=null)? (r[ox[iIdHP]]+'') : '';
        if (!idhp) return;

        const key = `${interno}||${idhp}`;
        const ts = (ox[iTS]!=null)? r[ox[iTS]] : '';
        const tsv = (ts instanceof Date) ? ts.getTime() : (EU7_toDate_(ts)?.getTime() || 0);

        const cur = otPend[key];
        if (!cur || tsv >= cur._ts){
          otPend[key] = {
            nroOT: (ox[iNroOT]!=null)? (r[ox[iNroOT]]+'') : '',
            idOT:  (ox[iIdOT]!=null)? (r[ox[iIdOT]]+'') : '',
            estado: est,
            _ts: tsv
          };
        }
      });
    }

    // --- PreventivosUnidad rows ---
    const data = shPU.getDataRange().getValues();
    const head = data.shift();
    const px = EU7_idx_(head);

    const items = [];
    data.forEach(r=>{
      const inx = (EU7_get_(r, px, 'Interno')+'');
      if (inx !== interno) return;

      const idHP = (EU7_get_(r, px, 'IdHP')+'').trim();
      const nombre = EU7_get_(r, px, 'NombreHP','Nombre Preventivo','Nombre');
      const control = (EU7_get_(r, px, 'Control','ControlTipo','Tipo','Clase')+'').toLowerCase().trim();

      const cadaKm = EU7_toNum_(EU7_get_(r, px, 'CadaKm','IntervaloKm'));
      const cadaDias = EU7_toNum_(EU7_get_(r, px, 'CadaDias','IntervaloDias'));

      const ultKm = EU7_toNum_(EU7_get_(r, px, 'UltimoKm','Ultimo'));
      const ultF  = EU7_toDate_(EU7_get_(r, px, 'UltimaFecha','UltimoFecha'));

      const proxKm = EU7_toNum_(EU7_get_(r, px, 'ProximoKm','Proximo'));
      const proxF  = EU7_toDate_(EU7_get_(r, px, 'ProximaFecha','ProximoFecha'));

      // decidir tipo
      let tipo = control;
      if (tipo !== 'km' && tipo !== 'dia'){
        if (cadaKm > 0) tipo = 'km';
        else if (cadaDias > 0) tipo = 'dia';
        else tipo = 'km';
      }

      let Cada='', Ultimo='', Actual='', Proximo='', Pasado='0', clase='ok';

      if (tipo === 'km'){
        Cada = (cadaKm>0 ? `${cadaKm} Km` : '');
        Ultimo = (ultKm>0 ? `${ultKm} Km` : '0 Km');
        const act = Math.max(0, kmActual - ultKm);
        Actual = `${act} Km`;
        Proximo = (proxKm>0 ? `${proxKm} Km` : (cadaKm>0 ? `${ultKm + cadaKm} Km` : ''));
        const p = (proxKm>0 && kmActual > proxKm) ? (kmActual - proxKm) : 0;
        Pasado = `${p}`;
        if (p>0) clase='pasado';
        else{
          // próximo si dentro del aviso (10% por defecto)
          const aviso = EU7_toNum_(EU7_get_(r, px, 'AvisoKm','AvisarAntesKm','AlertaPctKm'));
          const umbral = (aviso>0 ? (proxKm - aviso) : (proxKm - Math.ceil((cadaKm||0)*0.10)));
          if (proxKm>0 && kmActual >= umbral) clase='proximo';
        }
      } else {
        Cada = (cadaDias>0 ? `${cadaDias} Días` : '');
        Ultimo = ultF ? EU7_fmtDate_(ultF) : '';
        const hoy = new Date();
        const dias = ultF ? Math.max(0, EU7_daysBetween_(ultF, hoy)) : 0;
        Actual = `${dias} Días`;
        Proximo = proxF ? EU7_fmtDate_(proxF) : '';
        const p = (proxF && hoy > proxF) ? EU7_daysBetween_(proxF, hoy) : 0;
        Pasado = `${p}`;
        if (p>0) clase='pasado';
        else{
          const avisoD = EU7_toNum_(EU7_get_(r, px, 'AvisoDias','AvisarAntesDias','AlertaDias'));
          const umbralD = avisoD>0 ? avisoD : 7;
          if (proxF){
            const faltan = EU7_daysBetween_(hoy, proxF);
            if (faltan <= umbralD) clase='proximo';
          }
        }
      }

      const key = `${interno}||${idHP}`;
      const pend = otPend[key] || null;

      items.push({
        Cod: idHP,
        Preventivo: nombre,
        Tipo: tipo,
        Cada, Ultimo, Actual, Proximo, Pasado,
        NroOT: pend ? (pend.nroOT || '') : '',
        hasOTPendiente: !!pend,
        idOT: pend ? (pend.idOT || '') : '',
        clase
      });
    });

    return {ok:true, chasis, items};

  }catch(e){
    return {ok:false, msg:'No se pudo obtener el estado.', error:String(e)};
  }
}

// ---- API: Reprogramar (modal lindo ya en UI) ----
function reprogramarPreventivoUnidadV7(token, interno, idHP, nuevoUltimo, motivo){
  try{
    const auth = EU7_requireAuth_(token);
    if (!auth.ok) return auth;

    interno = (interno ?? '').toString().trim();
    idHP = (idHP ?? '').toString().trim();
    motivo = (motivo ?? '').toString().trim();
    if (!interno || !idHP) return {ok:false, msg:'Datos incompletos'};
    if (!motivo) return {ok:false, msg:'Motivo obligatorio'};

    const ss = EU7_ss_();
    const shPU  = EU7_sheet_(ss, ['PreventivosUnidad']);
    const shHis = EU7_sheet_(ss, ['PreventivosUnidadHis','PreventivoUnidadHis','preveventivosunidadhis']);

    if (!shPU) return {ok:false, msg:'No existe hoja PreventivosUnidad'};

    const v = shPU.getDataRange().getValues();
    const h = v.shift();
    const ix = EU7_idx_(h);

    // Buscar fila
    const iInterno = EU7_norm_('Interno');
    const iIdHP    = EU7_norm_('IdHP');

    let rowIdx = -1;
    for (let i=0;i<v.length;i++){
      const r = v[i];
      if ((r[ix[iInterno]]+'')===interno && (r[ix[iIdHP]]+'')===idHP){
        rowIdx = i; break;
      }
    }
    if (rowIdx<0) return {ok:false, msg:'No existe preventivo para esa unidad'};

    const row = v[rowIdx];

    // Tipo/control
    let control = (EU7_get_(row, ix, 'Control','ControlTipo','Tipo','Clase')+'').toLowerCase().trim();
    const cadaKm   = EU7_toNum_(EU7_get_(row, ix, 'CadaKm','IntervaloKm'));
    const cadaDias = EU7_toNum_(EU7_get_(row, ix, 'CadaDias','IntervaloDias'));
    if (control !== 'km' && control !== 'dia'){
      control = (cadaDias>0 && cadaKm<=0) ? 'dia' : 'km';
    }

    // Snapshot antes (para historial)
    const antes = {
      UltimoKm: EU7_get_(row, ix, 'UltimoKm'),
      UltimaFecha: EU7_get_(row, ix, 'UltimaFecha'),
      ProximoKm: EU7_get_(row, ix, 'ProximoKm'),
      ProximaFecha: EU7_get_(row, ix, 'ProximaFecha')
    };

    // Aplicar nuevo ÚLTIMO + recalcular PRÓXIMO
    if (control === 'dia'){
      const d = EU7_toDate_(nuevoUltimo);
      if (!d) return {ok:false, msg:'Fecha inválida'};
      EU7_set_(row, ix, 'UltimaFecha', d);

      // Recalcular próximo si hay intervalo
      if (cadaDias > 0){
        const p = new Date(d.getTime());
        p.setDate(p.getDate() + cadaDias);
        EU7_set_(row, ix, 'ProximaFecha', p);
      }
    } else {
      const n = EU7_toNum_(nuevoUltimo);
      if (!n) return {ok:false, msg:'Km inválido'};
      EU7_set_(row, ix, 'UltimoKm', n);

      if (cadaKm > 0){
        EU7_set_(row, ix, 'ProximoKm', n + cadaKm);
      }
    }

    // Marcas
    EU7_set_(row, ix, 'UltimaAccion', 'REPROG');
    EU7_set_(row, ix, 'Usuario', (typeof getSessionUser === 'function') ? (getSessionUser(token)||'') : '');
    EU7_set_(row, ix, 'Timestamp', new Date());

    // Guardar
    shPU.getRange(rowIdx+2, 1, 1, h.length).setValues([row]);

    // Historial robusto (según encabezados)
    if (shHis){
      const hv = shHis.getRange(1,1,1,shHis.getLastColumn()).getValues()[0];
      const hx = EU7_idx_(hv);
      const out = new Array(hv.length).fill('');

      const setH = (k,val)=>{ const kk = EU7_norm_(k); if (kk in hx) out[hx[kk]] = val; };

      setH('IdHis', Utilities.getUuid());
      setH('Interno', interno);
      setH('IdHP', idHP);
      setH('NombreHP', EU7_get_(row, ix, 'NombreHP','Nombre Preventivo','Nombre'));
      setH('Accion', 'REPROG');
      setH('Antes', JSON.stringify(antes));
      setH('Despues', JSON.stringify({
        UltimoKm: EU7_get_(row, ix, 'UltimoKm'),
        UltimaFecha: EU7_get_(row, ix, 'UltimaFecha'),
        ProximoKm: EU7_get_(row, ix, 'ProximoKm'),
        ProximaFecha: EU7_get_(row, ix, 'ProximaFecha')
      }));
      setH('Motivo', motivo);
      setH('Usuario', (typeof getSessionUser === 'function') ? (getSessionUser(token)||'') : '');
      setH('Timestamp', new Date());

      shHis.appendRow(out);
      SpreadsheetApp.flush();
      // his row number is lastRow after append

    }

    // Leer fila luego de guardar para verificar
    SpreadsheetApp.flush();
    const afterRow = shPU.getRange(rowIdx+2, 1, 1, h.length).getValues()[0];
    const despues = {
      UltimoKm: EU7_get_(afterRow, ix, 'UltimoKm'),
      UltimaFecha: EU7_get_(afterRow, ix, 'UltimaFecha'),
      ProximoKm: EU7_get_(afterRow, ix, 'ProximoKm'),
      ProximaFecha: EU7_get_(afterRow, ix, 'ProximaFecha')
    };
    return {ok:true, row: rowIdx+2, antes, despues, hisSheet: shHis ? shHis.getName() : ''};

  }catch(e){
    return {ok:false, msg:'No se pudo reprogramar.', error:String(e)};
  }
}
function EU_reprog_FINAL_fixId(payload) {
  try {
    if (!payload) return { ok:false, msg:"Payload vacío" };

    const interno = payload.interno;
    const idHPraw = payload.idHP || payload.Cod || payload.codigo || "";
    const nombreHP = payload.nombre || payload.nombreHP || "";
    const motivo = payload.motivo;
    const nuevoUltimoKm = payload.nuevoUltimoKm;
    const nuevaUltimaFecha = payload.nuevaUltimaFecha;

    if (!interno || !motivo) return { ok:false, msg:"Datos incompletos (interno/motivo)" };

    const ss = SpreadsheetApp.getActive();
    const shPU = ss.getSheetByName("PreventivosUnidad");
    if (!shPU) return { ok:false, msg:"No existe hoja PreventivosUnidad" };

    const data = shPU.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const idx = {};
    headers.forEach((h,i)=> idx[String(h).trim()] = i);

    let rowIndex = -1;

    // 1) Buscar por PU-xxxx (IdPU)
    if (idHPraw && String(idHPraw).startsWith("PU-") && idx["IdPU"] != null) {
      rowIndex = rows.findIndex(r => String(r[idx["IdPU"]]) === String(idHPraw));
    }

    // 2) Buscar por HP-xxxx (IdHP)
    if (rowIndex === -1 && idHPraw && String(idHPraw).startsWith("HP-") && idx["IdHP"] != null) {
      rowIndex = rows.findIndex(r => String(r[idx["IdHP"]]) === String(idHPraw));
    }

    // 3) Buscar por Interno + NombreHP
    if (rowIndex === -1 && nombreHP && idx["Interno"] != null && idx["NombreHP"] != null) {
      rowIndex = rows.findIndex(r =>
        String(r[idx["Interno"]]) === String(interno) &&
        String(r[idx["NombreHP"]]).toLowerCase() === String(nombreHP).toLowerCase()
      );
    }

    if (rowIndex === -1) return { ok:false, msg:"No se encontró preventivo para reprogramar" };

    const sheetRow = rowIndex + 2;

    // helpers
    const has = (k) => idx[k] != null;
    const setCell = (k, v) => shPU.getRange(sheetRow, idx[k] + 1).setValue(v);

    const antes = {
      UltimoKm: has("UltimoKm") ? shPU.getRange(sheetRow, idx["UltimoKm"]+1).getValue() : "",
      UltimaFecha: has("UltimaFecha") ? shPU.getRange(sheetRow, idx["UltimaFecha"]+1).getValue() : "",
      ProximoKm: has("ProximoKm") ? shPU.getRange(sheetRow, idx["ProximoKm"]+1).getValue() : "",
      ProximaFecha: has("ProximaFecha") ? shPU.getRange(sheetRow, idx["ProximaFecha"]+1).getValue() : ""
    };

    // Actualiza ÚLTIMO
    if (nuevoUltimoKm != null && nuevoUltimoKm !== "" && has("UltimoKm")) {
      setCell("UltimoKm", Number(nuevoUltimoKm));
      if (has("UltimaFecha")) setCell("UltimaFecha", "");
    }
    if (nuevaUltimaFecha && has("UltimaFecha")) {
      setCell("UltimaFecha", new Date(nuevaUltimaFecha));
      if (has("UltimoKm")) setCell("UltimoKm", "");
    }

    // Recalcula PRÓXIMO
    if (nuevoUltimoKm != null && nuevoUltimoKm !== "" && has("CadaKm") && has("ProximoKm")) {
      const cadaKm = Number(shPU.getRange(sheetRow, idx["CadaKm"]+1).getValue()) || 0;
      setCell("ProximoKm", Number(nuevoUltimoKm) + cadaKm);
      if (has("ProximaFecha")) setCell("ProximaFecha", "");
    }
    if (nuevaUltimaFecha && has("CadaDias") && has("ProximaFecha")) {
      const cadaDias = Number(shPU.getRange(sheetRow, idx["CadaDias"]+1).getValue()) || 0;
      const base = new Date(nuevaUltimaFecha);
      base.setDate(base.getDate() + cadaDias);
      setCell("ProximaFecha", base);
      if (has("ProximoKm")) setCell("ProximoKm", "");
    }

    const despues = {
      UltimoKm: has("UltimoKm") ? shPU.getRange(sheetRow, idx["UltimoKm"]+1).getValue() : "",
      UltimaFecha: has("UltimaFecha") ? shPU.getRange(sheetRow, idx["UltimaFecha"]+1).getValue() : "",
      ProximoKm: has("ProximoKm") ? shPU.getRange(sheetRow, idx["ProximoKm"]+1).getValue() : "",
      ProximaFecha: has("ProximaFecha") ? shPU.getRange(sheetRow, idx["ProximaFecha"]+1).getValue() : ""
    };

    // Historial (busca preventiv/unidad/his en el nombre)
    const shHis = ss.getSheets().find(s => /preventiv.*unidad.*his/i.test(s.getName()));
    if (shHis) {
      shHis.appendRow([
        "HIS-"+Utilities.getUuid().slice(0,8),
        interno,
        idHPraw,
        nombreHP,
        "REPROG",
        JSON.stringify(antes),
        JSON.stringify(despues),
        motivo,
        (Session.getActiveUser().getEmail() || "ADMIN"),
        new Date()
      ]);
    }

    return { ok:true, puRow: sheetRow, antes, despues, hisSheet: shHis ? shHis.getName() : "" };

  } catch (e) {
    return { ok:false, msg:String(e) };
  }
}
