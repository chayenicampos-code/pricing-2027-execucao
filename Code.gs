/**
 * Backend do "Execução Individual — Pricing 2027".
 *
 * Recebe os registros que cada pessoa do time marca na página HTML e grava
 * numa aba por dia (formato AAAA-MM-DD) dentro desta planilha. A planilha
 * fica com o acesso que VOCÊ (dona do script) já tem — quem preenche a
 * página nunca recebe acesso a ela, só consegue mandar dados pra dentro.
 *
 * COMO IMPLANTAR (fazer uma vez só):
 * 1. Crie uma Google Sheet nova e vazia (ex: "Pricing 2027 — Execução
 *    Diária (privado)") — não compartilhe com ninguém, ela é só sua.
 * 2. Nessa planilha: Extensões → Apps Script.
 * 3. Apague o conteúdo padrão do Code.gs e cole o conteúdo deste arquivo.
 * 4. Em "Implantar" → "Nova implantação" → tipo "App da Web":
 *      Executar como: Eu (seu e-mail)
 *      Quem pode acessar: Qualquer pessoa
 * 5. Autorize quando pedir (é seu próprio script acessando sua própria
 *    planilha — normal aparecer o aviso "app não verificado", pode seguir).
 * 6. Copie a URL do app da Web gerada (termina em /exec) e cole na
 *    constante ENDPOINT_URL no topo do arquivo HTML.
 *
 * Cada pessoa que preenche a página NUNCA vê nem precisa de acesso a essa
 * planilha — só você, abrindo ela direto no Google Sheets.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tabName = data.date; // "AAAA-MM-DD"
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(["horário", "nome", "bloco", "subgrupo", "item", "ação"]);
      sheet.setFrozenRows(1);
    }
    var rows = (data.events || []).map(function (ev) {
      return [ev.horario, data.nome, ev.bloco, ev.subgrupo, ev.texto, ev.acao];
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, gravados: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, info: "Endpoint ativo. Use POST para registrar execuções." })
  ).setMimeType(ContentService.MimeType.JSON);
}
