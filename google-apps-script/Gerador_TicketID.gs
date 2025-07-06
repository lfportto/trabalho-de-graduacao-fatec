// ============================================================
// Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
// Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
// Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
// Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
// Descrição: Este script cria o Ticket ID de um pedido após o usuário
// enviar um formulário, isto é, o código numérico de seis algarismos
// utilizado para identificar e rastrear um pedido ao longo do processo.
// ============================================================

function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas ao formulário 1");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  // Se ainda não existe o cabeçalho "Ticket ID", cria na primeira linha
  if (sheet.getRange(1, lastColumn).getValue() !== "Ticket ID") {
    sheet.getRange(1, lastColumn + 1).setValue("Ticket ID");
  }

  // Gerar ID aleatório de 6 dígitos
  var ticketID = Math.floor(100000 + Math.random() * 900000);
  
  // Gravar o Ticket ID na nova coluna para a nova resposta
  sheet.getRange(lastRow, sheet.getLastColumn()).setValue(ticketID);

  // Substituir vírgulas por pontos na coluna B (nome do item)
  var colBValue = sheet.getRange(lastRow, 2).getValue(); // Coluna B = 2
  var newValue = colBValue.replace(/,/g, '.');

  if (newValue !== colBValue) {
    sheet.getRange(lastRow, 2).setValue(newValue);
  }
}
