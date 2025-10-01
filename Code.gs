// --- CONFIGURAÇÕES GLOBAIS ---
const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId();
const NOME_DA_ABA = 'Respostas ao formulário 1'; // Verifique se o nome está EXATO

const ID_PASTA_UPLOADS = 'ID_DA_SUA_PASTA_UPLOADS_DE_LAUDOS'; // <-- COLOQUE O ID AQUI
const ID_PASTA_PDFS = 'ID_DA_SUA_PASTA_LAUDOS_GERADOS_EM_PDF'; // <-- COLOQUE O ID AQUI
const ID_DOCUMENTO_MODELO = 'ID_DO_SEU_DOCUMENTO_MODELO'; // <-- COLOQUE O ID AQUI
// --- FIM DAS CONFIGURAÇÕES ---

// Função que cria o Web App
function doGet(e) {
  return HtmlService.createTemplateFromFile('Formulario').evaluate()
      .setTitle('Formulário de Perda de Garantia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// Inclui o conteúdo de outros arquivos .html no principal
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Processa os dados recebidos do formulário HTML
function processarFormulario(formObject) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
    const cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
    
    const pastaUploads = DriveApp.getFolderById(ID_PASTA_UPLOADS);
    
    // Faz o upload dos arquivos de imagem e pega suas URLs
    formObject['Foto do Produto 1'] = uploadFile(formObject.foto1, pastaUploads);
    formObject['Foto do Produto 2'] = uploadFile(formObject.foto2, pastaUploads);
    formObject['Foto Nº de Série'] = uploadFile(formObject.foto3, pastaUploads);

    // Organiza os dados na ordem correta das colunas da planilha
    const novaLinha = cabecalhos.map(cabecalho => formObject[cabecalho] || '');
    
    // Adiciona a nova linha de dados na planilha
    planilha.appendRow(novaLinha);
    
    const numeroDaNovaLinha = planilha.getLastRow();
    
    // Chama a função para gerar o laudo para a linha que acabamos de adicionar
    const resultadoLaudo = gerarLaudoPDF(numeroDaNovaLinha);

    return { success: true, message: resultadoLaudo.message };
  } catch (e) {
    return { success: false, message: 'Erro no servidor: ' + e.message };
  }
}

// Função auxiliar para fazer upload do arquivo
function uploadFile(fileData, folder) {
  if (fileData && fileData.mimeType && fileData.bytes) {
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    return file.getUrl(); // Retorna a URL do arquivo no Drive
  }
  return '';
}

// --- FUNÇÃO DE GERAÇÃO DE PDF (Adaptada da versão anterior) ---
function gerarLaudoPDF(numeroDaLinha) {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
  const cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
  const dados = planilha.getRange(numeroDaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];
  
  const dadosParaSubstituir = {};
  cabecalhos.forEach((cabecalho, index) => {
    dadosParaSubstituir[cabecalho] = dados[index];
  });
  
  const emailColaborador = dadosParaSubstituir['Endereço de e-mail']; // Usa a coluna correta
  if (!emailColaborador) {
    throw new Error('Coluna "Endereço de e-mail" não encontrada ou vazia.');
  }

  const pastaDestino = DriveApp.getFolderById(ID_PASTA_PDFS);
  const arquivoModelo = DriveApp.getFileById(ID_DOCUMENTO_MODELO);

  const nomeCliente = dadosParaSubstituir['Cliente'] || 'Cliente';
  const nomePedido = dadosParaSubstituir['Pedido'] || 'Pedido';
  const nomeNovoDocumento = `Laudo - Pedido ${nomePedido} - ${nomeCliente}`;

  const novoDocumento = arquivoModelo.makeCopy(nomeNovoDocumento, pastaDestino);
  const doc = DocumentApp.openById(novoDocumento.getId());
  const body = doc.getBody();

  // Substitui os textos
  for (const cabecalho in dadosParaSubstituir) {
    const valor = dadosParaSubstituir[cabecalho];
    if (typeof valor === 'string' && !valor.includes('drive.google.com')) {
       body.replaceText(`<<${cabecalho}>>`, valor);
    }
  }
  
  // Substitui as imagens
  ['Foto do Produto 1', 'Foto do Produto 2', 'Foto Nº de Série'].forEach(cabecalho => {
    const urlImagem = dadosParaSubstituir[cabecalho];
    if (urlImagem) {
      const placeholder = `<<${cabecalho}>>`;
      const rangeElement = body.findText(placeholder);
      if (rangeElement) {
        const element = rangeElement.getElement().getParent();
        if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
          element.clear();
          const imageId = urlImagem.match(/id=([a-zA-Z0-9_-]+)/)[1];
          const imagemBlob = DriveApp.getFileById(imageId).getBlob();
          const insertedImage = element.appendInlineImage(imagemBlob);
          
          const maxWidth = 450;
          const ratio = insertedImage.getWidth() / insertedImage.getHeight();
          insertedImage.setWidth(maxWidth);
          insertedImage.setHeight(maxWidth / ratio);
        }
      }
    }
  });

  doc.saveAndClose();
  const blobPdf = novoDocumento.getAs('application/pdf');
  const arquivoPdf = pastaDestino.createFile(blobPdf).setName(nomeNovoDocumento + '.pdf');
  DriveApp.getFileById(novoDocumento.getId()).setTrashed(true);

  const assuntoEmail = `Laudo de Perda de Garantia: Pedido ${nomePedido}`;
  const corpoEmail = `Olá,\n\nSegue em anexo o laudo de perda de garantia para o cliente ${nomeCliente}, referente ao pedido ${nomePedido}.\n\nAtenciosamente,\nSistema Automático de Laudos.`;
  
  GmailApp.sendEmail(emailColaborador, assuntoEmail, corpoEmail, { attachments: [arquivoPdf] });
  
  return { success: true, message: 'Laudo gerado e enviado com sucesso!' };
}
