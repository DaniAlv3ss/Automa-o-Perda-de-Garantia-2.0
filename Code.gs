// --- CONFIGURAÇÕES GLOBAIS ---
const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId();
const NOME_DA_ABA = 'Formulario'; // Verifique se o nome está EXATO

const ID_PASTA_UPLOADS = '1TVsNXw7BHDFMj9_9VM8pBGJpULRyHBK5'; // <-- COLOQUE O ID AQUI
const ID_PASTA_PDFS = '1Hpk4-eKfrDCO3x6ELWS96mc2-pAgGQZd'; // <-- COLOQUE O ID AQUI
const ID_DOCUMENTO_MODELO = '1i_5C5DeXXtsT1qUMl29kg63I9lLHIfHJqyoRiy8GGNQ'; // <-- COLOQUE O ID AQUI
// --- FIM DAS CONFIGURAÇÕES ---

// Função que cria o Web App
function doGet(e) {
  Logger.log("--- INICIANDO doGet E TENTANDO OBTER E-MAIL DO USUÁRIO ---");
  let userEmail = '';
  let emailSource = 'Nenhum (falha)';
  try {
    userEmail = Session.getActiveUser().getEmail();
    if (userEmail) {
      emailSource = 'Usuário Ativo (getActiveUser)';
      Logger.log(`Sucesso ao obter e-mail do Usuário Ativo: ${userEmail}`);
    } else {
       Logger.log('getActiveUser() executado, mas retornou um e-mail vazio. Tentando próximo método.');
       throw new Error("getActiveUser retornou vazio.");
    }
  } catch (err) {
    Logger.log(`FALHA ao obter e-mail do Usuário Ativo. Erro: ${err.message}. Tentando método fallback.`);
    try {
      userEmail = Session.getEffectiveUser().getEmail();
      if (userEmail) {
        emailSource = 'Usuário Efetivo (getEffectiveUser)';
        Logger.log(`Sucesso ao obter e-mail do Usuário Efetivo: ${userEmail}`);
      } else {
        Logger.log('getEffectiveUser() executado, mas retornou um e-mail vazio.');
      }
    } catch (err2) {
      Logger.log(`FALHA CRÍTICA ao obter e-mail do Usuário Efetivo. O campo de e-mail será manual. Erro: ${err2.message}`);
    }
  }
  
  Logger.log(`Finalizando doGet. E-mail a ser enviado para o template: "${userEmail}" (Fonte: ${emailSource})`);
  const template = HtmlService.createTemplateFromFile('Formulario');
  template.userEmail = userEmail;
  return template.evaluate()
      .setTitle('Formulário de Perda de Garantia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// Inclui o conteúdo de outros arquivos .html no principal
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function obterLaudos() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'laudos_data_v4';
    const cachedData = cache.get(cacheKey);
    if (cachedData != null) { return JSON.parse(cachedData); }

    const planilha = SpreadsheetApp.openById(ID_PLANILHA);
    const abaLaudos = planilha.getSheetByName('Laudos');
    if (!abaLaudos) { throw new Error('Aba "Laudos" não encontrada na planilha.'); }
    const dados = abaLaudos.getRange(2, 1, abaLaudos.getLastRow() - 1, 2).getValues(); 
    
    const laudosArray = dados
      .filter(linha => linha[0] && linha[1])
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(linha => ({ categoria: linha[0], laudo: linha[1] }));

    laudosArray.unshift({ categoria: 'Laudo Personalizado', laudo: '' });
    cache.put(cacheKey, JSON.stringify(laudosArray), 3600);
    return laudosArray;
  } catch (e) {
    Logger.log(e.toString());
    return { error: e.message };
  }
}

// --- INÍCIO DA REESTRUTURAÇÃO PARA PROCESSAMENTO EM ETAPAS ---

// ETAPA 1: Salva os dados na planilha e faz o upload dos arquivos.
function etapa1_SalvarDadosEUploads(formObject) {
  try {
    Logger.log("--- ETAPA 1: INICIANDO ---");
    
    const fusoHorario = Session.getScriptTimeZone();
    const dataHoraAtual = Utilities.formatDate(new Date(), fusoHorario, 'dd/MM/yyyy HH:mm:ss');
    formObject['Data e Hora'] = dataHoraAtual;
    
    const pastaUploads = DriveApp.getFolderById(ID_PASTA_UPLOADS);
    formObject['Foto do Produto 1'] = uploadFile(formObject.foto1, pastaUploads);
    formObject['Foto do Produto 2'] = uploadFile(formObject.foto2, pastaUploads);
    formObject['Foto Nº de Série'] = uploadFile(formObject.foto3, pastaUploads);

    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
    const cabecalhosRaw = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
    const cabecalhos = cabecalhosRaw.map(h => h.trim()); 
    const novaLinha = cabecalhos.map(cabecalho => formObject[cabecalho] || '');
    planilha.appendRow(novaLinha);
    
    Logger.log("--- ETAPA 1: CONCLUÍDA ---");
    return { success: true, data: formObject };
  } catch (e) {
    Logger.log("!!! ERRO NA ETAPA 1: " + e.message + "\nStack: " + e.stack);
    return { success: false, message: 'Erro na Etapa 1 (Upload/Planilha): ' + e.message };
  }
}

// ETAPA 2: Cria o documento Google Docs e insere os dados e imagens.
function etapa2_CriarDocumento(dadosDaEtapa1) {
  try {
    Logger.log("--- ETAPA 2: INICIANDO ---");
    const dadosParaSubstituir = dadosDaEtapa1;
    
    const pastaDestino = DriveApp.getFolderById(ID_PASTA_PDFS);
    const arquivoModelo = DriveApp.getFileById(ID_DOCUMENTO_MODELO);
    const nomeCliente = dadosParaSubstituir['Cliente'] || 'Cliente';
    const nomePedido = dadosParaSubstituir['Pedido'] || 'Pedido';
    const nomeNovoDocumento = `Temp - Laudo - Pedido ${nomePedido} - ${nomeCliente}`;

    const novoDocumento = arquivoModelo.makeCopy(nomeNovoDocumento, pastaDestino);
    const docId = novoDocumento.getId();
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    for (const cabecalho in dadosParaSubstituir) {
      let valor = dadosParaSubstituir[cabecalho];
      if (valor instanceof Date) {
        valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      }
      const valorString = String(valor);
      if (!valorString.includes('drive.google.com')) {
          body.replaceText(`<<${cabecalho}>>`, valorString);
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
            let imageId = null;
            let match = urlImagem.match(/\/d\/([a-zA-Z0-9_-]{25,})/) || urlImagem.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
            if (match && match[1]) {
              imageId = match[1];
              try {
                const imagemBlob = DriveApp.getFileById(imageId).getBlob();
                const insertedImage = element.appendInlineImage(imagemBlob);
                const maxWidth = 450;
                const ratio = insertedImage.getWidth() / insertedImage.getHeight();
                insertedImage.setWidth(maxWidth);
                insertedImage.setHeight(maxWidth / ratio);
              } catch (imgError) {
                element.appendText(`[Erro ao carregar imagem]`);
              }
            } else {
              Logger.log(`URL de imagem inválida ou ID não encontrado: ${urlImagem}`);
              element.appendText('https://support.google.com/merchants/answer/12470638?hl=pt');
            }
          }
        }
      }
    });

    doc.saveAndClose();
    Logger.log("--- ETAPA 2: CONCLUÍDA ---");
    
    return { success: true, data: dadosDaEtapa1, tempDocId: docId, nomeDoArquivo: nomeNovoDocumento.replace('Temp - ','') };
  } catch (e) {
    Logger.log("!!! ERRO NA ETAPA 2: " + e.message + "\nStack: " + e.stack);
    return { success: false, message: 'Erro na Etapa 2 (Criação do Documento): ' + e.message };
  }
}

// ETAPA 3: Converte o documento para PDF, envia o e-mail e faz a limpeza.
function etapa3_FinalizarEEnviar(dadosDaEtapa2) {
  try {
    Logger.log("--- ETAPA 3: INICIANDO ---");
    const tempDocId = dadosDaEtapa2.tempDocId;
    const dadosParaEmail = dadosDaEtapa2.data;
    const nomeFinalDoArquivo = dadosDaEtapa2.nomeDoArquivo;
    
    const tempDocFile = DriveApp.getFileById(tempDocId);
    const pastaDestino = DriveApp.getFolderById(ID_PASTA_PDFS);

    const blobPdf = tempDocFile.getAs('application/pdf');
    const arquivoPdf = pastaDestino.createFile(blobPdf).setName(nomeFinalDoArquivo + '.pdf');
    
    tempDocFile.setTrashed(true);

    const emailColaborador = dadosParaEmail['Endereço de e-mail'];
    const nomeCliente = dadosParaEmail['Cliente'] || 'Cliente';
    const nomePedido = dadosParaEmail['Pedido'] || 'Pedido';
    const assuntoEmail = `Laudo de Perda de Garantia: Pedido ${nomePedido}`;
    const corpoEmail = `Olá,\n\nSegue em anexo o laudo de perda de garantia para o cliente ${nomeCliente}, referente ao pedido ${nomePedido}.\n\nAtenciosamente,\nSistema Automático de Laudos.`;
    GmailApp.sendEmail(emailColaborador, assuntoEmail, corpoEmail, { attachments: [arquivoPdf] });
    
    Logger.log("--- ETAPA 3: CONCLUÍDA ---");
    return { success: true, message: 'Laudo gerado e enviado com sucesso!' };
  } catch (e) {
    Logger.log("!!! ERRO NA ETAPA 3: " + e.message + "\nStack: " + e.stack);
    return { success: false, message: 'Erro na Etapa 3 (PDF/E-mail): ' + e.message };
  }
}

// Função auxiliar para fazer upload do arquivo
function uploadFile(fileData, folder) {
  if (fileData && fileData.mimeType && fileData.bytes) {
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    return file.getUrl();
  }
  return '';
}

