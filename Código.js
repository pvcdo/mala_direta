/**
 * doGet – recebe JSON, limpa preenchimentos antigos, salva novos dados
 *         e preenche o documento.
 */

function dataHojeFormatada() {
  var hoje = new Date();
  
  var dia = hoje.getDate();
  var mes = hoje.getMonth() + 1; // Janeiro é 0
  var ano = hoje.getFullYear();
  
  // Garante 2 dígitos para dia e mês
  var diaFormatado = (dia < 10 ? '0' : '') + dia;
  var mesFormatado = (mes < 10 ? '0' : '') + mes;
  
  var dataFormatada = diaFormatado + '-' + mesFormatado + '-' + ano;
  
  return dataFormatada;
}

function doGet(e) {

  const out = (msg, mime) =>
    ContentService.createTextOutput(msg)
                  .setMimeType(mime || ContentService.MimeType.TEXT);

  Logger.log('--- Início doGet ---');

  if (!e || !e.parameter || !e.parameter.dados) {
    Logger.log('Parâmetros ausentes.');
    return out('Parâmetros não informados.');
  }

  let dados;
  try {
    dados = JSON.parse(e.parameter.dados);
    Logger.log('JSON analisado: %s', JSON.stringify(dados));
  } catch (err) {
    Logger.log('Falha ao fazer JSON.parse: %s', err.message);
    return out('JSON inválido: ' + err.message);
  }

  const id_pasta_mala_direta = '11ReCNyq0RXgoUbHausmfzfCJWrM2x0s4'
  const pasta_mala = DriveApp.getFolderById(id_pasta_mala_direta)
  const arqs_mala = pasta_mala.getFilesByType(MimeType.GOOGLE_DOCS)

  const id_pasta_docs_gerados = '1xJryCCMbQ1pOt4zrNdfCY2UbMdn2NHfH'
  const pasta_docs_gerados = DriveApp.getFolderById(id_pasta_docs_gerados)

  const nome_documento = DocumentApp.getActiveDocument().getName()
  const nome_novo_doc = dados["Nº TICKET DE CONTRATAÇÃO"] + "_" + dataHojeFormatada() + "_" + dados["ATENDENTE"] + "_" + nome_documento

  while (arqs_mala.hasNext()) {
    const arq = arqs_mala.next()
    const nome = arq.getName()
    if(nome === nome_documento){
      const arq_copia = arq.makeCopy(nome_novo_doc, pasta_docs_gerados)
      const doc = DocumentApp.openById(arq_copia.getId())
      const body = doc.getBody()
      Object.entries(dados).forEach(([chave, valor]) => {
        try {
          body.replaceText('<<' + chave + '>>', String(valor));
        } catch (subErr) {
          Logger.log('Erro ao substituir "%s": %s', chave, subErr.message);
        }
      });
      const doc_pdf = "https://docs.google.com/document/d/"+arq_copia.getId()+"/export?format=pdf"

      return out(
        JSON.stringify({
          status : 'ok',
          mensagem : 'Documento atualizado com sucesso!',
          doc_pdf
        }),
        ContentService.MimeType.JSON
      );
    }
  }
}