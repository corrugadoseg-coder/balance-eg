/*****************************************************************
 *  BALANCE EG — Backend v7: getAll() rápido
 *
 *  QUÉ CAMBIA Y POR QUÉ
 *  --------------------
 *  Antes, cada getter (IMP, SERV, CHEQUE, SUELDOS, G, V, GINF, GTCARGAS…)
 *  llamaba a ensureCol() por CADA columna en CADA getAll(). Cada ensureCol
 *  hacía 2 lecturas de encabezado (una en ensureCol + otra en colIndex).
 *  Eran ~175 lecturas a la planilla por getAll() → ~20 segundos.
 *
 *  Ahora los getters NO migran en cada lectura: solo crean la hoja (con
 *  TODOS los encabezados) si no existe. La migración de columnas se hace
 *  una sola vez con migrarEsquema() (ver abajo).
 *
 *  CÓMO APLICARLO
 *  --------------
 *  1) En el editor de Apps Script, REEMPLAZÁ las funciones
 *     IMP, CHEQUE, SERV, CHEMIT, SUELDOS, G, V, GINF, GTCARGAS
 *     por las versiones de este archivo.
 *  2) PEGÁ también migrarEsquema() y medirGetAll() (son nuevas).
 *  3) Ejecutá UNA vez migrarEsquema() desde el editor (asegura que las
 *     hojas existentes tengan todas las columnas). Es idempotente.
 *  4) Medí con medirGetAll() antes y después (mirá el Registro de ejecución).
 *  5) Redesplegá: Implementar → Administrar implementaciones → editar →
 *     Versión: NUEVA → Implementar. La URL no cambia.
 *
 *  ensureCol() y colIndex() SE CONSERVAN (los usa migrarEsquema()).
 *****************************************************************/

/* ---- getters: crean la hoja con TODOS los encabezados si falta; NO migran en cada lectura ---- */

function IMP() {
  var sh = SS.getSheetByName("Impuestos");
  if (!sh) {
    sh = SS.insertSheet("Impuestos");
    sh.appendRow(["Mes","IVAContador","RetBancIIBB","IIBBContador","AlicIIBB","CoefCABA","CoefBsAs","IVAExtracto","IVAVentas","IVACompras","IVACompNeto","PercIVACompras","PercIIBBCompras","RetIVAClientes","IVAExcluido","NetoCompras","NetoVentas","GastosNegro","MesPagoIVA","RetIIBBClientes","RetGanancias","SaldoFavorIIBB","RetSUSSClientes"]);
  }
  return sh;
}

function CHEQUE() {
  var sh = SS.getSheetByName("ImpCheque");
  if (!sh) {
    sh = SS.insertSheet("ImpCheque");
    sh.appendRow(["Mes","Total","Compensado","UsadoCCSS"]);
  }
  return sh;
}

function SERV() {
  var sh = SS.getSheetByName("Servicios");
  if (!sh) {
    sh = SS.insertSheet("Servicios");
    sh.appendRow(["ID","Mes","Servicio","Categoria","Rubro","Sub","FechaEmision","Vencimiento","FechaPago","TipoPago","Entidad","Monto","IVA","PercIVA","PercIIBB","Estado","Obs","EnBalance","GastoID","EnInforme","Debito","Pagado"]);
  }
  return sh;
}

function CHEMIT() {
  var sh = SS.getSheetByName("ChequesEmit");
  if (!sh) {
    sh = SS.insertSheet("ChequesEmit");
    sh.appendRow(["ID","Tipo","Nro","Beneficiario","CUIT","FechaPago","Monto","Estado","Obs"]);
  }
  return sh;
}

function SUELDOS() {
  var sh = SS.getSheetByName("Sueldos");
  if (!sh) {
    sh = SS.insertSheet("Sueldos");
    sh.appendRow(["ID","Mes","Tipo","Modalidad","Nombre","DepA","Q1A","Q2A","VacA","SacA","ManoB","Q1B","Q2B","VacB","SacB","Monto","Obs","GastoIdA","GastoIdB"]);
  }
  return sh;
}

function G() {
  var sh = SS.getSheetByName("Gastos");
  if (!sh) {
    sh = SS.insertSheet("Gastos");
    sh.appendRow(["ID","Fecha","Proveedor","Rubro","Subcategoria","Tipo","Monto","IVA","PercIVA","PercIIBB"]);
  }
  return sh;
}

function V() {
  var sh = SS.getSheetByName("Ventas");
  if (!sh) {
    sh = SS.insertSheet("Ventas");
    sh.appendRow(["Mes","Blanco","Negro","IVADebito","RetIVA","RetIIBB","RetGan"]);
  }
  return sh;
}

function GINF() {
  var sh = SS.getSheetByName("GastosInforme");
  if (!sh) {
    sh = SS.insertSheet("GastosInforme");
    sh.appendRow(["Mes","Rubro","Monto","Tipo","Proveedor"]);
  }
  return sh;
}

function GTCARGAS() {
  var sh = SS.getSheetByName("GTCargas");
  if (!sh) {
    sh = SS.insertSheet("GTCargas");
    sh.appendRow(["Comprobante","Data","FacturaMes","Total"]);
  }
  return sh;
}

/* ---- migración de esquema: se corre UNA sola vez (idempotente), no en cada getAll ---- */
function migrarEsquema() {
  var g = SS.getSheetByName("Gastos");
  if (g) {
    var gh = g.getRange(1,1,1,g.getLastColumn()).getValues()[0];
    if (gh.indexOf("Subcategoria") === -1) { g.insertColumnBefore(5); g.getRange(1,5).setValue("Subcategoria"); }
    ["IVA","PercIVA","PercIIBB"].forEach(function(c){ ensureCol(g,c); });
  }
  var v = SS.getSheetByName("Ventas");
  if (v) ["IVADebito","RetIVA","RetIIBB","RetGan"].forEach(function(c){ ensureCol(v,c); });
  var imp = SS.getSheetByName("Impuestos");
  if (imp) ["IVAContador","RetBancIIBB","IIBBContador","AlicIIBB","CoefCABA","CoefBsAs","IVAExtracto","IVAVentas","IVACompras","IVACompNeto","PercIVACompras","PercIIBBCompras","RetIVAClientes","IVAExcluido","NetoCompras","NetoVentas","GastosNegro","MesPagoIVA","RetIIBBClientes","RetGanancias","SaldoFavorIIBB","RetSUSSClientes"].forEach(function(c){ ensureCol(imp,c); });
  var ch = SS.getSheetByName("ImpCheque");
  if (ch) ["Total","Compensado","UsadoCCSS"].forEach(function(c){ ensureCol(ch,c); });
  var sv = SS.getSheetByName("Servicios");
  if (sv) ["Mes","Servicio","Categoria","Rubro","Sub","FechaEmision","Vencimiento","FechaPago","TipoPago","Entidad","Monto","IVA","PercIVA","PercIIBB","Estado","Obs","EnBalance","GastoID","EnInforme","Debito","Pagado"].forEach(function(c){ ensureCol(sv,c); });
  var ce = SS.getSheetByName("ChequesEmit");
  if (ce) ["Tipo","Nro","Beneficiario","CUIT","FechaPago","Monto","Estado","Obs"].forEach(function(c){ ensureCol(ce,c); });
  var su = SS.getSheetByName("Sueldos");
  if (su) ["Mes","Tipo","Modalidad","Nombre","DepA","Q1A","Q2A","VacA","SacA","ManoB","Q1B","Q2B","VacB","SacB","Monto","Obs","GastoIdA","GastoIdB"].forEach(function(c){ ensureCol(su,c); });
  var gi = SS.getSheetByName("GastosInforme");
  if (gi) ["Rubro","Monto","Tipo","Proveedor"].forEach(function(c){ ensureCol(gi,c); });
  var gt = SS.getSheetByName("GTCargas");
  if (gt) ["Comprobante","Data","FacturaMes","Total"].forEach(function(c){ ensureCol(gt,c); });
  return "Esquema verificado/migrado OK.";
}

/* ---- medición: corré esto antes y después; mirá el Registro de ejecución ---- */
function medirGetAll() {
  var t0 = new Date().getTime();
  var r = getAll();
  var ms = new Date().getTime() - t0;
  var info = "getAll(): " + ms + " ms  ·  gastos=" + r.data.gastos.length +
             "  ventas=" + Object.keys(r.data.ventas).length +
             "  impuestos=" + Object.keys(r.data.impuestos).length +
             "  gastosInforme=" + r.data.gastosInforme.length +
             "  gtCargas=" + Object.keys(r.data.gtCargas).length;
  Logger.log(info);
  return info;
}
