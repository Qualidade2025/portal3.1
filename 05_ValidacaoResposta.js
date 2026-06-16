/** =========================================================
 * Validação da Resposta (Qualidade)
 *  - validarResposta(token, rncId)
 *  - retornarResposta(token, rncId, motivo)
 * Dependências: _getSession_(token)
 * Planilhas: "Controle", "Log"
 * ========================================================= */

/**
* Qualidade: valida a resposta enviada pela área responsável.
 * Efeito: Controle!C ← "Resposta"
 */
function registrarDataHoraValidacao(sh, row, dataHora) {
  var dt = dataHora instanceof Date ? dataHora : new Date();
  var cols = _getControleColumnIndices_(sh);
  var range = sh.getRange(row, cols.dataValidacaoResposta.column);
  range.setValue(dt);
  range.setNumberFormat('dd/MM/yyyy HH:mm');
}