/**
 * Hook seguro: al confirmar una OT preventiva, actualiza PreventivosUnidad
 * para reiniciar el contador (KM o DÍAS) y recalcular el próximo.
 *
 * IMPORTANTE:
 * - NO declara SPREADSHEET_ID (usa el existente si está)
 * - NO rompe la confirmación si algo falla (OTService ya lo atrapa)
 *
 * Firma esperada por OTService.confirmarOT():
 *   EU_registerConfirmacionPreventivo_(interno, idHP, idOT, usuario, meta)
 */
function EU_registerConfirmacionPreventivo_(interno, idHP, idOT, usuario, meta) {
  interno = (interno ?? "").toString().trim();
  idHP    = (idHP ?? "").toString().trim();
  if (!interno || !idHP) return;

  const ss = (typeof SPREADSHEET_ID !== "undefined")
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  const shPU  = ss.getSheetByName("PreventivosUnidad");
  const shHP  = ss.getSheetByName("HojasPreventivas");
  const shCh  = ss.getSheetByName("ChasisBD");
  const shHis = ss.getSheetByName("PreventivosUnidadHis");

  if (!shPU) throw new Error("No existe la hoja PreventivosUnidad");

  const now = new Date(); // fecha real de confirmación

  // --- Helpers ---
  const norm = (x)=> (x ?? "").toString().trim();
  const toNum = (x)=>{
    const n = Number(String(x ?? "").toString().replace(/[^\d.-]/g,""));
    return isFinite(n) ? n : 0;
  };
  const getHeaderMap = (head)=>{
    const map = {};
    head.forEach((h,i)=>{ map[norm(h)] = i; });
    return map;
  };
  const idx = (map, ...names)=>{
    for (const n of names){
      const k = norm(n);
      if (k && Object.prototype.hasOwnProperty.call(map,k)) return map[k];
    }
    return -1;
  };

  // --- KM actual del chasis ---
  let kmActual = null;
  if (shCh){
    const v = shCh.getDataRange().getValues();
    if (v.length >= 2){
      const mapC = getHeaderMap(v[0]);
      const iInt = idx(mapC, "Interno");
      const iKm  = idx(mapC, "KmRecorridos","KM","Km");
      if (iInt !== -1 && iKm !== -1){
        for (let r=1;r<v.length;r++){
          if (norm(v[r][iInt]) === interno){
            kmActual = toNum(v[r][iKm]);
            break;
          }
        }
      }
    }
  }

  // --- Leer PreventivosUnidad ---
  let pu = shPU.getDataRange().getValues();
  const head = pu[0] || [];
  const mapPU = getHeaderMap(head);

  const iIdPU   = idx(mapPU, "IdPU","Id");
  const iIntPU  = idx(mapPU, "Interno");
  const iIdHP2  = idx(mapPU, "IdHP");
  const iNomHP  = idx(mapPU, "NombreHP","NombrePreventivo","Preventivo");
  const iCtrl   = idx(mapPU, "Control","ControlTipo","Tipo");
  const iCadaKm = idx(mapPU, "CadaKm","IntervaloKm");
  const iCadaD  = idx(mapPU, "CadaDias","IntervaloDias");
  const iAvisKm = idx(mapPU, "AvisoKm","AvisarAntesKm","AlertaKm","AlertaPctKm");
  const iAvisD  = idx(mapPU, "AvisoDias","AvisarAntesDias","AlertaDias");

  const iUltKm  = idx(mapPU, "UltimoKm");
  const iUltF   = idx(mapPU, "UltimaFecha","Ultimo","Ultima");
  const iProxKm = idx(mapPU, "ProximoKm");
  const iProxF  = idx(mapPU, "ProximaFecha");

  const iPend   = idx(mapPU, "PendienteOT");
  const iIdPend = idx(mapPU, "IdOTPendiente");
  const iUltId  = idx(mapPU, "UltimoIdOT");
  const iUltAcc = idx(mapPU, "UltimaAccion");
  const iUsr    = idx(mapPU, "Usuario");
  const iTS     = idx(mapPU, "Timestamp");

  if (iIntPU === -1 || iIdHP2 === -1) throw new Error("PreventivosUnidad: faltan columnas Interno/IdHP");

  let rowIdx = -1;
  for (let r=1;r<pu.length;r++){
    if (norm(pu[r][iIntPU]) === interno && norm(pu[r][iIdHP2]) === idHP){
      rowIdx = r;
      break;
    }
  }

  // --- Obtener info de HojasPreventivas (si hace falta) ---
  function getHPInfo(){
    const info = { NombreHP:"", Control:"", CadaKm:0, CadaDias:0, AvisoKm:0, AvisoDias:0 };
    if (!shHP) return info;
    const v = shHP.getDataRange().getValues();
    if (v.length < 2) return info;
    const mapH = getHeaderMap(v[0]);
    const iId = idx(mapH, "IdHP","Codigo","Id");
    const iNom= idx(mapH, "NombreHP","Nombre","Preventivo");
    const iCk = idx(mapH, "CadaKm","IntervaloKm");
    const iCd = idx(mapH, "CadaDias","IntervaloDias");
    const iCt = idx(mapH, "ControlTipo","Control","Tipo");
    const iAk = idx(mapH, "AvisarAntesKm","AvisoKm","AlertaPctKm","AlertaKm");
    const iAd = idx(mapH, "AvisarAntesDias","AvisoDias","AlertaDias");
    for (let r=1;r<v.length;r++){
      if (norm(v[r][iId]) === idHP){
        info.NombreHP = iNom!==-1 ? norm(v[r][iNom]) : "";
        info.CadaKm   = iCk!==-1 ? toNum(v[r][iCk]) : 0;
        info.CadaDias = iCd!==-1 ? toNum(v[r][iCd]) : 0;
        info.Control  = iCt!==-1 ? norm(v[r][iCt]).toLowerCase() : "";
        info.AvisoKm  = iAk!==-1 ? toNum(v[r][iAk]) : 0;
        info.AvisoDias= iAd!==-1 ? toNum(v[r][iAd]) : 0;
        break;
      }
    }
    if (!info.Control){
      if (info.CadaKm && !info.CadaDias) info.Control = "km";
      else if (info.CadaDias && !info.CadaKm) info.Control = "dia";
    }
    return info;
  }

  // --- Crear fila si no existe ---
  if (rowIdx === -1){
    const info = getHPInfo();
    const row = new Array(head.length).fill("");
    if (iIdPU !== -1) row[iIdPU] = "PU-" + Utilities.getUuid().slice(0,8);
    row[iIntPU] = interno;
    row[iIdHP2] = idHP;
    if (iNomHP !== -1) row[iNomHP] = info.NombreHP;
    if (iCtrl !== -1) row[iCtrl] = info.Control;
    if (iCadaKm !== -1) row[iCadaKm] = info.CadaKm || "";
    if (iCadaD  !== -1) row[iCadaD]  = info.CadaDias || "";
    if (iAvisKm !== -1 && info.AvisoKm) row[iAvisKm] = info.AvisoKm;
    if (iAvisD  !== -1 && info.AvisoDias) row[iAvisD] = info.AvisoDias;

    shPU.appendRow(row);

    // recargar para obtener índice real
    pu = shPU.getDataRange().getValues();
    rowIdx = pu.length - 1;
  }

  // --- Update existente ---
  const before = pu[rowIdx].slice();
  const row = pu[rowIdx];

  const control = (iCtrl !== -1 ? norm(row[iCtrl]).toLowerCase() : "");

  if (control === "km"){
    if (kmActual === null) throw new Error("No pude obtener KmRecorridos del chasis");
    if (iUltKm !== -1) row[iUltKm] = kmActual;
    const cadaKm = (iCadaKm !== -1) ? toNum(row[iCadaKm]) : 0;
    if (cadaKm > 0 && iProxKm !== -1) row[iProxKm] = kmActual + cadaKm;
    // Fecha también se actualiza por trazabilidad
    if (iUltF !== -1) row[iUltF] = now;
  } else if (control === "dia"){
    if (iUltF !== -1) row[iUltF] = now;
    const cadaDias = (iCadaD !== -1) ? toNum(row[iCadaD]) : 0;
    if (cadaDias > 0 && iProxF !== -1){
      const prox = new Date(now);
      prox.setDate(prox.getDate() + cadaDias);
      row[iProxF] = prox;
    }
    // UltimoKm no aplica
  } else {
    // Si no hay control, intentamos inferir
    const cadaKm = (iCadaKm !== -1) ? toNum(row[iCadaKm]) : 0;
    const cadaDias = (iCadaD !== -1) ? toNum(row[iCadaD]) : 0;
    if (cadaKm && !cadaDias){
      if (iCtrl !== -1) row[iCtrl] = "km";
      if (kmActual !== null && iUltKm !== -1) row[iUltKm] = kmActual;
      if (kmActual !== null && cadaKm > 0 && iProxKm !== -1) row[iProxKm] = kmActual + cadaKm;
      if (iUltF !== -1) row[iUltF] = now;
    } else if (cadaDias && !cadaKm){
      if (iCtrl !== -1) row[iCtrl] = "dia";
      if (iUltF !== -1) row[iUltF] = now;
      if (cadaDias > 0 && iProxF !== -1){
        const prox = new Date(now);
        prox.setDate(prox.getDate() + cadaDias);
        row[iProxF] = prox;
      }
    }
  }

  // Limpia OT pendiente y setea tracking
  if (iPend !== -1) row[iPend] = "";
  if (iIdPend !== -1) row[iIdPend] = "";
  if (iUltId !== -1) row[iUltId] = idOT || "";
  if (iUltAcc !== -1) row[iUltAcc] = "CONFIRM";
  if (iUsr !== -1) row[iUsr] = usuario || "";
  if (iTS !== -1) row[iTS] = now;

  // Guardar fila
  shPU.getRange(rowIdx+1, 1, 1, head.length).setValues([row]);

  // Historial
  if (shHis){
    const hv = shHis.getDataRange().getValues();
    const hh = hv[0] || [];
    // si el historial no tiene schema esperado, append simple
    shHis.appendRow([
      "HIS-" + Utilities.getUuid().slice(0,8),
      interno,
      idHP,
      "",
      "",
      "",
      "",
      usuario || "",
      now
    ]);
    // Intento de esquema extendido (si existen columnas)
    try{
      const mapH = getHeaderMap(hh);
      const cols = new Array(hh.length).fill("");
      const set = (name,val)=>{ const i = idx(mapH,name); if (i!==-1) cols[i]=val; };
      set("IdHis","HIS-" + Utilities.getUuid().slice(0,8));
      set("Interno",interno);
      set("IdHP",idHP);
      set("NombreHP", (iNomHP!==-1? norm(row[iNomHP]):""));
      set("Accion","CONFIRM");
      set("Antes", JSON.stringify(before));
      set("Despues", JSON.stringify(row));
      set("Motivo", "Confirmación OT " + (idOT||""));
      set("Usuario", usuario || "");
      set("Timestamp", now);
      if (hh.length && cols.some(x=>x!=="")){
        shHis.appendRow(cols);
      }
    }catch(e){}
  }
}
