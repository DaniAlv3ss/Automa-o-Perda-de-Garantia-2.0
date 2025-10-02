// --- CONFIGURAÇÕES GLOBAIS ---
const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId();
const NOME_DA_ABA = 'Formulario'; // Verifique se o nome está EXATO

const ID_PASTA_UPLOADS = '1TVsNXw7BHDFMj9_9VM8pBGJpULRyHBK5'; // <-- COLOQUE O ID AQUI
const ID_PASTA_PDFS = '1Hpk4-eKfrDCO3x6ELWS96mc2-pAgGQZd'; // <-- COLOQUE O ID AQUI
const ID_DOCUMENTO_MODELO = '1i_5C5DeXXtsT1qUMl29kg63I9lLHIfHJqyoRiy8GGNQ'; // <-- COLOQUE O ID AQUI
// --- FIM DAS CONFIGURAÇÕES ---

// **INÍCIO DA CORREÇÃO REFORÇADA DE E-MAIL**
// Função que cria o Web App e passa o e-mail do usuário para o HTML
function doGet(e) {
  Logger.log("--- INICIANDO doGet E TENTANDO OBTER E-MAIL DO USUÁRIO ---");
  let userEmail = '';
  let emailSource = 'Nenhum (falha)'; // Para logging, para saber de onde o e-mail veio.

  try {
    // MÉTODO 1: Usuário Ativo. Ideal para Web Apps executados como "usuário que acessa o app".
    // Isso requer que o usuário autorize o acesso à sua identidade na primeira vez que usa o app.
    userEmail = Session.getActiveUser().getEmail();
    if (userEmail) {
      emailSource = 'Usuário Ativo (getActiveUser)';
      Logger.log(`Sucesso ao obter e-mail do Usuário Ativo: ${userEmail}`);
    } else {
       // Isso pode acontecer se o usuário não tiver um e-mail ou se a permissão foi negada.
       Logger.log('getActiveUser() executado, mas retornou um e-mail vazio. Tentando próximo método.');
       throw new Error("getActiveUser retornou vazio."); // Força a ida para o próximo catch
    }
  } catch (err) {
    Logger.log(`FALHA ao obter e-mail do Usuário Ativo. Erro: ${err.message}. Tentando método fallback.`);
    try {
      // MÉTODO 2: Usuário Efetivo. Geralmente, é o dono do script ou quem o autorizou.
      // Útil como fallback ou se o app for executado como "eu" (dono do script).
      userEmail = Session.getEffectiveUser().getEmail();
      if (userEmail) {
        emailSource = 'Usuário Efetivo (getEffectiveUser)';
        Logger.log(`Sucesso ao obter e-mail do Usuário Efetivo: ${userEmail}`);
      } else {
        Logger.log('getEffectiveUser() executado, mas retornou um e-mail vazio.');
      }
    } catch (err2) {
      Logger.log(`FALHA CRÍTICA ao obter e-mail do Usuário Efetivo. O campo de e-mail será manual. Erro: ${err2.message}`);
      // Se ambos os métodos falharem, userEmail permanecerá uma string vazia.
    }
  }
  
  Logger.log(`Finalizando doGet. E-mail a ser enviado para o template: "${userEmail}" (Fonte: ${emailSource})`);

  const template = HtmlService.createTemplateFromFile('Formulario');
  template.userEmail = userEmail; // Passa o e-mail (ou string vazia) para o HTML
  return template.evaluate()
      .setTitle('Formulário de Perda de Garantia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}
// **FIM DA CORREÇÃO REFORÇADA DE E-MAIL**

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

// Processa os dados recebidos do formulário HTML
function processarFormulario(formObject) {
  try {
    Logger.log("--- INICIANDO PROCESSAMENTO DO FORMULÁRIO ---");
    Logger.log("Dados recebidos do formulário: " + JSON.stringify(formObject, null, 2));

    const fusoHorario = Session.getScriptTimeZone();
    const dataHoraAtual = Utilities.formatDate(new Date(), fusoHorario, 'dd/MM/yyyy HH:mm:ss');
    formObject['Data e Hora'] = dataHoraAtual;
    
    const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
    const cabecalhosRaw = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
    const cabecalhos = cabecalhosRaw.map(h => h.trim()); 
    Logger.log("Cabeçalhos lidos e limpos da planilha: " + JSON.stringify(cabecalhos));
    
    const pastaUploads = DriveApp.getFolderById(ID_PASTA_UPLOADS);
    
    formObject['Foto do Produto 1'] = uploadFile(formObject.foto1, pastaUploads);
    formObject['Foto do Produto 2'] = uploadFile(formObject.foto2, pastaUploads);
    formObject['Foto Nº de Série'] = uploadFile(formObject.foto3, pastaUploads);

    const novaLinha = cabecalhos.map(cabecalho => formObject[cabecalho] || '');
    
    Logger.log("Nova linha a ser inserida na planilha: " + JSON.stringify(novaLinha));
    planilha.appendRow(novaLinha);
    Logger.log("Linha inserida com sucesso.");
    
    const numeroDaNovaLinha = planilha.getLastRow();
    const resultadoLaudo = gerarLaudoPDF(numeroDaNovaLinha);

    Logger.log("--- PROCESSAMENTO CONCLUÍDO COM SUCESSO ---");
    return { success: true, message: resultadoLaudo.message };
  } catch (e) {
    Logger.log("!!! ERRO NO SERVIDOR: " + e.message + "\nStack: " + e.stack);
    return { success: false, message: 'Erro no servidor: ' + e.message };
  }
}

function uploadFile(fileData, folder) {
  if (fileData && fileData.mimeType && fileData.bytes) {
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    return file.getUrl();
  }
  return '';
}

function gerarLaudoPDF(numeroDaLinha) {
  Logger.log(`--- INICIANDO GERAÇÃO DE PDF PARA A LINHA ${numeroDaLinha} ---`);
  const planilha = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(NOME_DA_ABA);
  const cabecalhosRaw = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];
  const cabecalhos = cabecalhosRaw.map(h => h.trim());
  const dados = planilha.getRange(numeroDaLinha, 1, 1, planilha.getLastColumn()).getValues()[0];
  
  const dadosParaSubstituir = {};
  cabecalhos.forEach((cabecalho, index) => {
    dadosParaSubstituir[cabecalho] = dados[index];
  });
  Logger.log("Dados mapeados para substituição no documento: " + JSON.stringify(dadosParaSubstituir, null, 2));
  
  const emailColaborador = dadosParaSubstituir['Endereço de e-mail'];
  if (!emailColaborador) { throw new Error('Coluna "Endereço de e-mail" não encontrada ou vazia.'); }

  const pastaDestino = DriveApp.getFolderById(ID_PASTA_PDFS);
  const arquivoModelo = DriveApp.getFileById(ID_DOCUMENTO_MODELO);

  const nomeCliente = dadosParaSubstituir['Cliente'] || 'Cliente';
  const nomePedido = dadosParaSubstituir['Pedido'] || 'Pedido';
  const nomeNovoDocumento = `Laudo - Pedido ${nomePedido} - ${nomeCliente}`;

  const novoDocumento = arquivoModelo.makeCopy(nomeNovoDocumento, pastaDestino);
  const doc = DocumentApp.openById(novoDocumento.getId());
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
            // **INÍCIO DA CORREÇÃO DE SINTAXE**
            element.appendText('https://support.google.com/merchants/answer/12470638?hl=en'); // O texto de erro deve ser uma string entre aspas/apóstrofos.
            // **FIM DA CORREÇÃO DE SINTAXE**
          }
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
  
  Logger.log("--- GERAÇÃO DE PDF CONCLUÍDA ---");
  return { success: true, message: 'Laudo gerado e enviado com sucesso!' };
}

