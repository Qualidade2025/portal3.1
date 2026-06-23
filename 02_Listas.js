/** =========================================================
 * Listas (para Criar RNC)
 * ========================================================= */

function listarFornecedores() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Dados');
  if (!sh) throw new Error('Aba "Dados" não encontrada.');
  var data = sh.getDataRange().getValues();
  var hdrIdx=-1, cArea=-1, cUsuario=-1, cSenha=-1;
  for (var i=0;i<data.length;i++){
    var row = data[i].map(function(x){ return String(x||'').trim(); });
    cUsuario=row.indexOf('Usuário'); cSenha=row.indexOf('Senha');
    if (cUsuario>=0 && cSenha>=0){ cArea=row.indexOf('Área'); hdrIdx=i; break; }
  }
  if (hdrIdx===-1 || cArea===-1) return [];
  var set = {};
  for (var r=hdrIdx+1; r<data.length; r++){
    var v = String(data[r][cArea]||'').trim(); if (v) set[v]=true;
  }
  return Object.keys(set).sort();
}

function listarMotivos() { // globais (Dados!K2:K)
  var sh = SpreadsheetApp.getActive().getSheetByName('Dados');
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var vals = sh.getRange(2, 11, lastRow-1, 1).getValues(); // K
  var out = [];
  for (var i=0;i<vals.length;i++){ var v = String(vals[i][0]||'').trim(); if (v) out.push(v); }
  return out;
}

function listarMotivosPorFornecedor(fornecedor) {
  fornecedor = String(fornecedor || '').trim();
  var arr = [];

  if (fornecedor) {
    var sh = SpreadsheetApp.getActive().getSheetByName('Dados');
    if (sh) {
      var lastCol = sh.getLastColumn(), lastRow = Math.max(sh.getLastRow(), 2);
      if (lastCol >= MOTIVOS_ANCHOR_COL) {
        var header = sh.getRange(1, MOTIVOS_ANCHOR_COL, 1, lastCol - MOTIVOS_ANCHOR_COL + 1).getValues()[0];
        var colIndex = -1;
        for (var j=0; j<header.length; j++){
          if (String(header[j]||'').trim() === fornecedor){ colIndex = MOTIVOS_ANCHOR_COL + j; break; }
        }
        if (colIndex !== -1) {
          var colVals = sh.getRange(2, colIndex, lastRow-1, 1).getValues();
          for (var i=0;i<colVals.length;i++){
            var v = String(colVals[i][0]||'').trim();
            if (v) arr.push(v);
          }
        }
      }
    }
  }
  if (arr.length === 0) arr = listarMotivos(); // fallback global

  // dedup + ordena + "Outro"
  var set = {}, out = [];
  for (var i=0;i<arr.length;i++){
    var v = String(arr[i]||'').trim();
    if (!v || v.toLowerCase()==='outro') continue;
    if (!set[v]) { set[v]=true; out.push(v); }
  }
  out.sort();
  out.push('Outro');
  return out;
}

function listarMotivosResolucaoOutro(fornecedor) {
  fornecedor = String(fornecedor || '').trim();
  var sh = SpreadsheetApp.getActive().getSheetByName('Dados');
  if (!sh) throw new Error('Aba "Dados" não encontrada.');

  var lastRow = Math.max(sh.getLastRow(), 2);
  var lastCol = sh.getLastColumn();

  function normalizeKey(value) {
    return String(value || '').trim().toLowerCase();
  }

  function uniqueClean(values, skipSet) {
    var seen = {};
    var out = [];
    for (var i = 0; i < values.length; i++) {
      var v = String(values[i] || '').trim();
      var key = normalizeKey(v);
      if (!v || key === 'outro' || seen[key] || (skipSet && skipSet[key])) continue;
      seen[key] = true;
      out.push(v);
    }
    out.sort();
    return out;
  }

  var motivosFornecedor = [];
  if (fornecedor && lastCol >= MOTIVOS_ANCHOR_COL) {
    var endCol = Math.min(lastCol, 34); // AH
    if (endCol >= MOTIVOS_ANCHOR_COL) {
      var headers = sh.getRange(1, MOTIVOS_ANCHOR_COL, 1, endCol - MOTIVOS_ANCHOR_COL + 1).getValues()[0];
      var colIndex = -1;
      for (var h = 0; h < headers.length; h++) {
        if (String(headers[h] || '').trim() === fornecedor) {
          colIndex = MOTIVOS_ANCHOR_COL + h;
          break;
        }
      }
      if (colIndex !== -1) {
        motivosFornecedor = sh.getRange(2, colIndex, lastRow - 1, 1).getValues().map(function(row) { return row[0]; });
      }
    }
  }

  motivosFornecedor = uniqueClean(motivosFornecedor);
  var fornecedorSet = {};
  motivosFornecedor.forEach(function(v) { fornecedorSet[normalizeKey(v)] = true; });

  var motivosGerais = [];
  if (lastCol >= 12) {
    motivosGerais = sh.getRange(2, 12, lastRow - 1, 1).getValues().map(function(row) { return row[0]; });
  }
  motivosGerais = uniqueClean(motivosGerais, fornecedorSet);

  return {
    fornecedor: fornecedor,
    motivosFornecedor: motivosFornecedor,
    motivosGerais: motivosGerais
  };
}
