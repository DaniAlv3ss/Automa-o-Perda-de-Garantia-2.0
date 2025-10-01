// --- CONFIGURAÇÕES GLOBAIS ---
const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId();
const NOME_DA_ABA = 'Formulario'; // Verifique se o nome está EXATO

const ID_PASTA_UPLOADS = '1TVsNXw7BHDFMj9_9VM8pBGJpULRyHBK5'; // <-- COLOQUE O ID AQUI
const ID_PASTA_PDFS = '1Hpk4-eKfrDCO3x6ELWS96mc2-pAgGQZd'; // <-- COLOQUE O ID AQUI
const ID_DOCUMENTO_MODELO = '1i_5C5DeXXtsT1qUMl29kg63I9lLHIfHJqyoRiy8GGNQ'; // <-- COLOQUE O ID AQUI
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

// --- FUNÇÃO OTIMIZADA PARA BUSCAR E ORDENAR OS LAUDOS (FORMATO ARRAY) ---
function obterLaudos() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'laudos_data_v4'; // Nova chave para o formato array
    const cachedData = cache.get(cacheKey);

    if (cachedData != null) {
      Logger.log("Laudos carregados do cache (formato array).");
      return JSON.parse(cachedData);
    }

    Logger.log("Laudos não encontrados no cache. Buscando na planilha.");
    const planilha = SpreadsheetApp.openById(ID_PLANILHA);
    const abaLaudos = planilha.getSheetByName('Laudos');
    if (!abaLaudos) {
      throw new Error('Aba "Laudos" não encontrada na planilha.');
    }
    const dados = abaLaudos.getRange(2, 1, abaLaudos.getLastRow() - 1, 2).getValues(); 
    
    // Filtra, ordena alfabeticamente pela categoria, e mapeia para um array de objetos
    const laudosArray = dados
      .filter(linha => linha[0] && linha[1]) // Garante que não adiciona linhas vazias
      .sort((a, b) => a[0].localeCompare(b[0])) // Ordena pelo nome da categoria
      .map(linha => ({ categoria: linha[0], laudo: linha[1] })); // Mapeia para o formato {categoria, laudo}

    // Adiciona a opção de laudo personalizado no início do array
    laudosArray.unshift({ categoria: 'Laudo Personalizado', laudo: '' });

    // Armazena o array no cache por 1 hora (3600 segundos)
    cache.put(cacheKey, JSON.stringify(laudosArray), 3600);
    Logger.log("Laudos ordenados (formato array) e armazenados no cache.");

    return laudosArray;
  } catch (e) {
    Logger.log(e.toString());
    return { error: e.message }; // Retorna um objeto de erro para o cliente
  }
}


// Processa os dados recebidos do formulário HTML
function processarFormulario(formObject) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
    const cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
    
    const pastaUploads = DriveApp.getFolderById(ID_PASTA_UPLOADS);
    
    formObject['Foto do Produto 1'] = uploadFile(formObject.foto1, pastaUploads);
    formObject['Foto do Produto 2'] = uploadFile(formObject.foto2, pastaUploads);
    formObject['Foto Nº de Série'] = uploadFile(formObject.foto3, pastaUploads);

    const novaLinha = cabecalhos.map(cabecalho => formObject[cabecalho] || '');
    
    planilha.appendRow(novaLinha);
    
    const numeroDaNovaLinha = planilha.getLastRow();
    
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

// --- FUNÇÃO DE GERAÇÃO DE PDF (Sem alterações) ---
function gerarLaudoPDF(numeroDaLinha) {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
  const cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
  const dados = planilha.getRange(numeroDaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];
  
  const dadosParaSubstituir = {};
  cabecalhos.forEach((cabecalho, index) => {
    dadosParaSubstituir[cabecalho] = dados[index];
  });
  
  const emailColaborador = dadosParaSubstituir['Endereço de e-mail'];
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

  for (const cabecalho in dadosParaSubstituir) {
    const valor = dadosParaSubstituir[cabecalho];
    if (typeof valor === 'string' && !valor.includes('drive.google.com')) {
        body.replaceText(`<<${cabecalho}>>`, valor);
    }
  }
  
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

