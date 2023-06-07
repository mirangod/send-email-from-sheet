// Matheus C. @mirangod

/**
 * PT-BR: Função que envia um e-mail automático com base no preenchimento da
 * coluna B com um valor e a coluna C com um "CONCLUÍDO". As variáveis do
 * e-mail são definidas de acordo com a linhas das colunas J até P.
 * EN-US: Function that sends an automatic e-mail based on filling column B
 * with a value and column C with a "CONCLUÍDO". E-mail variables are
 * defined according to row from clumns J to P.
 */
function enviarCadastro() {
  // Váriaveis que envolvem o evento
  const sheetCadastro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO");
  const rangeStatus = sheetCadastro.getRange("D:D").getValues();

  for (var i in rangeStatus){
    if (rangeStatus[i] == "CONCLUÍDO" && sheetCadastro.getRange(parseInt(i)+1,3).getValue() != "" && sheetCadastro.getRange(parseInt(i)+1,17).getValue() == ""){
      sheetCadastro.getRange(parseInt(i)+1,17).setValue(Utilities.formatDate(new Date,"GMT-3","dd/MM/yyyy HH:mm:ss"));
      const horario = sheetCadastro.getRange(parseInt(i)+1,17).getValue();

      const t = HtmlService.createTemplateFromFile("e-mail");
      t.id = sheetCadastro.getRange(parseInt(i)+1,1).getValue();
      t.tipo = sheetCadastro.getRange(parseInt(i)+1,2).getValue();
      t.valor = sheetCadastro.getRange(parseInt(i)+1,3).getValue();
      t.horario = Utilities.formatDate(horario,"GMT-3","dd/MM/yyyy HH:mm:ss");
      t.resp1 = sheetCadastro.getRange(parseInt(i)+1,10).getValue();
      t.resp2 = sheetCadastro.getRange(parseInt(i)+1,11).getValue();
      t.resp3 = sheetCadastro.getRange(parseInt(i)+1,12).getValue();
      t.resp4 = sheetCadastro.getRange(parseInt(i)+1,13).getValue();
      t.resp5 = sheetCadastro.getRange(parseInt(i)+1,14).getValue();
      t.resp6 = sheetCadastro.getRange(parseInt(i)+1,15).getValue();
      t.resp7 = sheetCadastro.getRange(parseInt(i)+1,16).getValue();
      
      switch (sheetCadastro.getRange(parseInt(i)+1,2).getValue()) {
        case "FORNECEDOR":
          t.form1 = "NOME DO FORNECEDOR";
          t.form2 = "CNPJ";
          t.form3 = "INSCRIÇÃO ESTADUAL"
          t.form4 = "TELEFONE";
          t.form5 = "ENDEREÇO";
          t.form6 = "E-MAIL";
          t.form7 = "TIPO DE FORNECEDOR";

          const message1 = {
            name: "Cadastro de Fornecedor",
            to: sheetCadastro.getRange(parseInt(i)+1,8).getValue(),
            subject: "CADASTRO "+ sheetCadastro.getRange(parseInt(i)+1,1).getValue() + " - CONCLUÍDO",
            htmlBody: t.evaluate().getContent()
          }

          MailApp.sendEmail(message1);
          Logger.log("Email enviado...");
          break;
        case "ITEM":
          t.form1 = "NOME DO ITEM";
          t.form2 = "MARCA";
          t.form3 = "UNIDADE DE MEDIDA";
          t.form4 = "ESPECIFICAÇÕES";
          t.form5 = "FORNECEDOR DE INDICAÇÃO";
          t.form6 = "USO PRINCIPAL";
          t.form7 = "NCM";

          const message2 = {
            name: "Cadastro de Item",
            to: sheetCadastro.getRange(parseInt(i)+1,8).getValue(),
            subject: "CADASTRO "+ sheetCadastro.getRange(parseInt(i)+1,1).getValue() + " - CONCLUÍDO",
            htmlBody: t.evaluate().getContent()
          }

          MailApp.sendEmail(message2);
          Logger.log("Email enviado...");
          break;
      }
    }
  }
}