// ============================================================
// GOOGLE APPS SCRIPT — Liquidador Nokia 2026
// ============================================================
const SPREADSHEET_ID = '1BJbrJVuYADu41s2uBv42AIVCdFstNr7zvpsCyPHrTqI';

// ── GET ───────────────────────────────────────────────────────
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreate(ss, 'Data');
    var raw = sheet.getRange('A2').getValue();
    var data = raw ? JSON.parse(raw) : { sitios: [], gastos: [] };
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST ──────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data;
    // FormData
    if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    }
    // Plain text body
    else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    else { throw new Error('Sin datos'); }

    if (!data.sitios) throw new Error('Estructura inválida');

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Guardar JSON completo en Data
    var dataSheet = getOrCreate(ss, 'Data');
    if (dataSheet.getRange('A1').getValue() === '') {
      dataSheet.getRange('A1:B1').setValues([['json_data','last_updated']]);
    }
    dataSheet.getRange('A2').setValue(JSON.stringify(data));
    dataSheet.getRange('B2').setValue(new Date().toLocaleString('es-CO'));

    // Usar calcs pre-calculados si vienen del app
    var calcs = data.calcs || [];

    actualizarConsolidado(ss, data, calcs);
    actualizarGastos(ss, data);
    actualizarSitios(ss, data, calcs);

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        sitios: data.sitios.length,
        gastos: (data.gastos||[]).length
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── CONSOLIDADO ───────────────────────────────────────────────
function actualizarConsolidado(ss, data, calcs) {
  var sheet = getOrCreate(ss, 'Consolidado');
  sheet.clearContents();
  sheet.clearFormats();

  var headers = [
    'Sitio','Tipo','Fecha','LC','Empresa','Cat','CW',
    'TI','ADJ','CW Nokia','CR','Total Venta',
    'SubC TI','SubC ADJ','SubC CW','Mat TI','Mat CW','Logística','Adicionales','BackOffice','Total Costo',
    'Utilidad','% Margen'
  ];
  var rows = [headers];

  // Build lookup from calcs array
  var calcMap = {};
  calcs.forEach(function(cc) { calcMap[cc.id] = cc; });

  (data.sitios || []).forEach(function(s) {
    var cc = calcMap[s.id];
    if (!cc) {
      // Fallback if no calcs: use raw data
      var gastosS = (data.gastos||[]).filter(function(g){return g.sitio===s.id;});
      cc = {
        nombre:s.nombre, tipo:s.tipo, fecha:s.fecha,
        ciudad:(s.ciudad||'').replace('Ciudad_',''),
        lc:s.lc, empresa:'', cat:s.cat, tiene_cw:s.tiene_cw,
        nokiaTI:0,nokiaADJ:0,nokiaCW:s.cw_nokia||0,nokiaCR:0,
        totalVenta:s.cw_nokia||0,
        subcTI:0,subcADJ:0,subcCW:s.cw_costo||0,
        matTI:sumArr(gastosS,'Materiales TI'),
        matCW:sumArr(gastosS,'Materiales CW'),
        logist:sumArr(gastosS,'Logistica'),
        adicion:sumArr(gastosS,'Adicionales'),
        backoffice:(s.costos&&s.costos.backoffice)||0,
        totalCosto:0, utilidad:0, margen:0
      };
    }
    rows.push([
      cc.nombre||s.nombre, cc.tipo||s.tipo, cc.fecha||s.fecha,
      cc.lc||s.lc, cc.empresa||'', cc.cat||s.cat,
      s.tiene_cw ? 'Sí' : 'No',
      cc.nokiaTI||0, cc.nokiaADJ||0, cc.nokiaCW||0, cc.nokiaCR||0, cc.totalVenta||0,
      cc.subcTI||0, cc.subcADJ||0, cc.subcCW||0,
      cc.matTI||0, cc.matCW||0, cc.logist||0, cc.adicion||0, cc.backoffice||0,
      cc.totalCosto||0,
      cc.utilidad||0, (cc.margen||0)+'%'
    ]);
  });

  sheet.getRange(1,1,rows.length,headers.length).setValues(rows);

  // Header row: dark green
  formatRange(sheet.getRange(1,1,1,headers.length), '#144E4A','#CDFBF2', true);

  // Venta Nokia columns (8-12): light blue
  if (rows.length > 1) {
    sheet.getRange(2,8,rows.length-1,5).setBackground('#EFF6FF');
    sheet.getRange(2,12,rows.length-1,1).setBackground('#DBEAFE').setFontWeight('bold');
    // Costo SubC columns (13-21): light yellow
    sheet.getRange(2,13,rows.length-1,9).setBackground('#FFFBEB');
    sheet.getRange(2,21,rows.length-1,1).setBackground('#FEF3C7').setFontWeight('bold');
    // Utilidad/Margen
    sheet.getRange(2,22,rows.length-1,2).setFontWeight('bold');
  }

  // Total row
  var totalRow = rows.length + 1;
  var totals = ['TOTAL PROYECTO','','','','','',''];
  var sumCols = [8,9,10,11,12,13,14,15,16,17,18,19,20,21,22];
  for (var i = 1; i <= headers.length; i++) {
    if (sumCols.indexOf(i) !== -1) {
      sheet.getRange(totalRow, i).setFormula('=SUM('+colLetter(i)+'2:'+colLetter(i)+rows.length+')');
    } else if (i <= 7) {
      // skip
    }
  }
  sheet.getRange(totalRow,1).setValue('TOTAL PROYECTO');
  // % margen total
  sheet.getRange(totalRow,23).setFormula('=IF('+colLetter(12)+totalRow+'>0,('+colLetter(22)+totalRow+'/'+colLetter(12)+totalRow+')*100,0)&"%"');
  formatRange(sheet.getRange(totalRow,1,1,headers.length), '#1A4D1A','#CDFBF2', true);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

// ── GASTOS ────────────────────────────────────────────────────
function actualizarGastos(ss, data) {
  var sheet = getOrCreate(ss, 'Gastos');
  sheet.clearContents();
  sheet.clearFormats();

  var headers = ['Sitio','Tipo','Descripción','Valor'];
  var rows = [headers];
  (data.gastos||[]).forEach(function(g){
    rows.push([g.sitio||'', g.tipo||'', g.desc||'', g.valor||0]);
  });

  sheet.getRange(1,1,rows.length,4).setValues(rows);
  formatRange(sheet.getRange(1,1,1,4), '#FFD36C','#000000', true);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1,4);
}

// ── HOJA POR SITIO ────────────────────────────────────────────
function actualizarSitios(ss, data, calcs) {
  var calcMap = {};
  calcs.forEach(function(cc){ calcMap[cc.id]=cc; });

  (data.sitios||[]).forEach(function(s) {
    var sheetName = s.nombre.replace(/[:\\/\[\]*?']/g,'').substring(0,30);
    var sheet = getOrCreate(ss, sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    var cc = calcMap[s.id] || {};
    var gastosS = (data.gastos||[]).filter(function(g){return g.sitio===s.id;});

    var r = 1;

    // ── Info sitio ──
    var infoData = [
      ['SITIO:', s.nombre, '', 'Fecha:', s.fecha],
      ['Tipo:', s.tipo, '', 'Ciudad:', (s.ciudad||'').replace('Ciudad_','')],
      ['LC:', s.lc, '', 'Cat:', s.cat],
      ['CW:', s.tiene_cw?'Sí':'No', '', 'BackOffice:', cc.backoffice||0]
    ];
    sheet.getRange(r,1,infoData.length,5).setValues(infoData);
    formatRange(sheet.getRange(r,1,infoData.length,1),'#144E4A','#CDFBF2',true);
    r += infoData.length + 1;

    // ── Nokia Liquidación ──
    sheet.getRange(r,1,1,6).setValues([['NOKIA — LIQUIDACIÓN VENTA','','','','','']]);
    sheet.getRange(r,1,1,6).merge();
    formatRange(sheet.getRange(r,1,1,6),'#144E4A','#CDFBF2',true);
    r++;
    var nokiaHdr = [['Actividad','Sección','Unidad','Cantidad','P. Nokia','Total Nokia']];
    sheet.getRange(r,1,1,6).setValues(nokiaHdr);
    formatRange(sheet.getRange(r,1,1,6),'#1A4D1A','#CDFBF2',true);
    r++;

    (s.actividades||[]).forEach(function(act){
      sheet.getRange(r,1,1,6).setValues([[
        act.nombre||act.id||'', act.sec||'', act.unidad||'',
        act.cant||0, 0, 0
      ]]);
      r++;
    });
    if (s.tiene_cw && s.cw_nokia > 0) {
      sheet.getRange(r,1,1,6).setValues([['CW Obra Civil','CW','Sitio',1,s.cw_nokia,s.cw_nokia]]);
      r++;
    }
    // Total Nokia
    sheet.getRange(r,1,1,6).setValues([['TOTAL VENTA NOKIA','','','','',cc.totalVenta||0]]);
    formatRange(sheet.getRange(r,1,1,6),'#DBEAFE','#1E3A5F',true);
    r += 2;

    // ── SubC Liquidación ──
    sheet.getRange(r,1,1,6).setValues([['SUBCONTRATISTA — LIQUIDACIÓN PAGO','','','','','']]);
    sheet.getRange(r,1,1,6).merge();
    formatRange(sheet.getRange(r,1,1,6),'#3D2800','#FFF0CE',true);
    r++;
    sheet.getRange(r,1,1,6).setValues([['Actividad','Sección','Unidad','Cantidad','P. SubC','Total SubC']]);
    formatRange(sheet.getRange(r,1,1,6),'#5C3300','#FFD36C',true);
    r++;

    (s.actividades||[]).forEach(function(act){
      if(act.id==='PM') return;
      sheet.getRange(r,1,1,6).setValues([[
        act.nombre||act.id||'', act.sec||'', act.unidad||'',
        act.cant||0, 0, 0
      ]]);
      r++;
    });
    if (s.tiene_cw && s.cw_costo > 0) {
      sheet.getRange(r,1,1,6).setValues([['CW SubContratista','CW','Sitio',1,s.cw_costo,s.cw_costo]]);
      r++;
    }
    r++;

    // ── Costos Operativos ──
    sheet.getRange(r,1,1,2).setValues([['COSTOS OPERATIVOS','']]);
    formatRange(sheet.getRange(r,1,1,2),'#FFD36C','#000000',true);
    r++;
    var costRows = [
      ['SubC TI+ADJ+CR:', (cc.subcTI||0)+(cc.subcADJ||0)],
      ['SubC CW:', cc.subcCW||0],
      ['Materiales TI:', cc.matTI||0],
      ['Materiales CW:', cc.matCW||0],
      ['Logística:', cc.logist||0],
      ['Adicionales:', cc.adicion||0],
      ['BackOffice:', cc.backoffice||0]
    ];
    sheet.getRange(r,1,costRows.length,2).setValues(costRows);
    sheet.getRange(r,1,costRows.length,2).setBackground('#FFFAEE');
    r += costRows.length;
    sheet.getRange(r,1,1,2).setValues([['TOTAL COSTO:', cc.totalCosto||0]]);
    formatRange(sheet.getRange(r,1,1,2),'#FEF3C7','#78350F',true);
    r++;
    sheet.getRange(r,1,1,2).setValues([['UTILIDAD:', cc.utilidad||0]]);
    var utColor = (cc.margen||0)>=30?'#1A7A1A':(cc.margen||0)>=20?'#FFC000':'#C0392B';
    formatRange(sheet.getRange(r,1,1,2), utColor, '#FFFFFF', true);
    r++;
    sheet.getRange(r,1,1,2).setValues([['% MARGEN:', (cc.margen||0)+'%']]);
    formatRange(sheet.getRange(r,1,1,2), utColor, '#FFFFFF', true);
    r += 2;

    // ── Gastos ──
    if (gastosS.length > 0) {
      sheet.getRange(r,1,1,4).setValues([['GASTOS REGISTRADOS','','','']]);
      sheet.getRange(r,1,1,4).merge();
      formatRange(sheet.getRange(r,1,1,4),'#FFD36C','#000000',true);
      r++;
      sheet.getRange(r,1,1,4).setValues([['Tipo','Descripción','Valor','']]);
      formatRange(sheet.getRange(r,1,1,4),'#FFFAEE','#000000',true);
      r++;
      gastosS.forEach(function(g){
        sheet.getRange(r,1,1,3).setValues([[g.tipo||'', g.desc||'', g.valor||0]]);
        r++;
      });
    }

    sheet.autoResizeColumns(1,6);
  });
}

// ── Helpers ───────────────────────────────────────────────────
function getOrCreate(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function formatRange(range, bg, fg, bold) {
  range.setBackground(bg).setFontColor(fg).setFontWeight(bold?'bold':'normal');
}

function sumArr(arr, tipo) {
  return arr.filter(function(g){return g.tipo===tipo;})
    .reduce(function(a,g){return a+(g.valor||0);},0);
}

function colLetter(n) {
  var s='';
  while(n>0){s=String.fromCharCode(64+(n%26||26))+s;n=Math.floor((n-1)/26);}
  return s;
}
