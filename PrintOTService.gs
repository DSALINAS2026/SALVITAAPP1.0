/**
 * ==============================
 * PRINT OT - DATA SERVICE
 * Lee OT + tareas + datos unidad
 * ==============================
 */
function getOTPrintData(token, idOT){
  const u = _requireSession_(token);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shOT  = ss.getSheetByName("ordenesTrabajo");
  const shDet = ss.getSheetByName("OT_Tareas");
  const shCh  = ss.getSheetByName("ChasisBD");

  if(!shOT) throw new Error("No existe hoja ordenesTrabajo");
  if(!idOT) throw new Error("Falta idOT");

  // --- OT head
  const otVals = shOT.getDataRange().getValues();
  const otHead = otVals.shift();
  const idxOT = (n)=> otHead.indexOf(n);

  const r = otVals.find(row => String(row[idxOT("IdOT")]) === String(idOT));
  if(!r) throw new Error("No se encontró la OT: " + idOT);

  const head = {
    IdOT: r[idxOT("IdOT")],
    NroOT: r[idxOT("NroOT")],
    TipoOT: r[idxOT("TipoOT")],
    EstadoOT: r[idxOT("EstadoOT")],
    FechaCreacion: r[idxOT("Fecha")] || r[idxOT("Timestamp")] || r[idxOT("FechaMov")] || "",
    Interno: r[idxOT("Interno")],
    Dominio: r[idxOT("Dominio")],
    Sociedad: r[idxOT("Sociedad")],
    Deposito: r[idxOT("Deposito")],
    Sector: r[idxOT("Sector")],
    NombrePreventivo: r[idxOT("NombrePreventivo")] || r[idxOT("TipoPreventivo")] || "",
    Solicita: r[idxOT("Solicita")],
    Descripcion: r[idxOT("Descripcion")],
    Usuario: r[idxOT("Usuario")]
  };

  // --- Unidad extra (marca/chasis/motor/km)
  let unidad = {};
  if(shCh){
    const chVals = shCh.getDataRange().getValues();
    const chHead = chVals.shift();
    const ic = (n)=> chHead.indexOf(n);
    const chRow = chVals.find(row => String(row[ic("Interno")]) === String(head.Interno));
    if(chRow){
      unidad = {
        Marca: chRow[ic("Marca")] || chRow[ic("Marca ") ] || "",
        NroChasis: chRow[ic("Nro. Chasis")] || chRow[ic("Nro Chasis")] || chRow[ic("NroChasis")] || "",
        NroMotor: chRow[ic("Nro Motor")] || chRow[ic("Nro. Motor")] || chRow[ic("NroMotor")] || "",
        Km: chRow[ic("KmRecorridos")] || ""
      };
    }
  }

  // --- Tareas
  let tareas = [];
  if(shDet){
    const detVals = shDet.getDataRange().getValues();
    const detHead = detVals.shift();
    const idd = (n)=> detHead.indexOf(n);

    tareas = detVals
      .filter(row => String(row[idd("IdOT")]) === String(idOT))
      .map(row => ({
        CodigoTarea: row[idd("CodigoTarea")],
        NombreTarea: row[idd("NombreTarea")],
        Sector: row[idd("Sector")] || head.Sector || ""
      }));
  }

  return {
    ...head,
    ...unidad,
    Tareas: tareas
  };
}
