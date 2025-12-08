// Pasta onde ficam guardados os anexos que vêm do email [ficheiros por processar] (assegurado pelo N8N, que envia os anexos em pdf do email documentos@darkland.pt para a pasta)
const PASTA_GERAL_FICHEIROS = "1DKCSluenYGGNz05uLLzQwWwq-wYHCk54";
// Sub-pasta, "#0.1 - Comprovativos", onde são colocados os comprovativos depois de fazer os pagamentos
var PASTA_COMPROVATIVOS_ID = "1nBnKXyvtiUt7BMYdaQfYtaJbgBMe4xjR";
// Sub-pasta, "#0.2 - Extratos de conta", onde são colocados os extratos de conta
var PASTA_EXTRATOS_ID = "1X6qfsGSjve4IsD0HdZOtMdh7OK3oklLC";

// --> Departamentos Darkpurple --> [DP] 2. Financeiro e RH (FERH) --> [DP][FERH] RH --> [DP][FERH] Recibos de vencimento --> 0 - Por processar
const PASTA_GERAL_RECIBOS = "1AW34HuS07KvBrclaoHQIEmd0pyMCATHQ";

// --> Departamentos Darkland --> [DL] 2. Financeiro e RH (FERH) --> [DL][FERH] Contabilidade --> [DL][FERH] Faturas de compras
const PASTA_GERAL_FATURAS = "17Onz--A6H-AdeMon0AvK3wRCvkNQhsq2";

const PASTA_LIXO = "1BzcMPR93SXKjh7fM4IQZRlB-sefWBR6a";

const EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL = "financeiro@arrowplus.pt";

const CODIGO_EMPRESA = "DL";

// Spreadsheet de log de movimentos
const SHEET_MOVIMENTOS_ID = "1r6flE_SuZJ2VLV5xmUo2qIvIv7p4Z9JMk6pwz6m-nTI";
const SHEET_MOVIMENTOS_NOME = "Movimentos";
const DESTINO_PATH_DEPTH = 3; // Quantos pais escreve na spreadsheet do destino final de cada ficheiro

// Cache em memória das pastas de anos
var __cacheYearFolders = null;

/*
*
* COPIAR PARA OUTRAS EMPRESAS A PARTIR DAQUI:
*
*/

/**
 * ======================================================================================
 * FUNÇÃO PRINCIPAL DE DISTRIBUIÇÃO (ATIVADA POR ACCIONADOR)
 * ======================================================================================
 */
function distribuirFicheirosDoGeral() {
  var sourceFolder = DriveApp.getFolderById(PASTA_GERAL_FICHEIROS);
  var pastaGeralRecibos = DriveApp.getFolderById(PASTA_GERAL_RECIBOS);
  var pastaLixo = DriveApp.getFolderById(PASTA_LIXO);

  var recCount = 0; 
  var fileCount = 0; 
  var fileErrors = 0; 
  var nomesFicheirosMovidos = '';
  var errosFicheirosMovidos = '';

  var files = sourceFolder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    // === RESET DE VARIÁVEIS CRÍTICO (Evita contaminação entre ficheiros) ===
    let month = null;
    let year = null;
    let day = null;
    let dataDocumento = null;
    let valorATCUD = null;
    //let tipoDocumento = null;
    // =======================================================================

    let textoPDF = "";
    try {
      textoPDF = convertPDFToText(file.getId(), ['pt', 'en', null]) || "";
    } catch (e) {
      fileErrors++;
      errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ': Falha no OCR (' + e + ').\n';
      continue;
    }
    
    if (!textoPDF.trim()) { 
      fileErrors++;
      errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ': PDF sem texto após OCR.\n';
      continue;
    }

    // EXTRAÇÕES BASE
    //tipoDocumento = extractTipoDocumento(textoPDF);
    valorATCUD = extractATCUD(textoPDF);
    dataDocumento = extractDataDocumentoTaloes(textoPDF);
    
    // Normalização para comparações
    let t = (textoPDF || "").toLowerCase();

    // ------------------------------------------------------------------------
    // CASO ESPECIAL: MEO (Escreve na fatura "Este documento não serve de fatura"!)
    // ------------------------------------------------------------------------
    if(textoPDF.includes("504615947")){
       Logger.log("🛡️ MEO DETETADA: A neutralizar frase de recibo.");
       
       // Truque de Mestre: Removemos a frase "venenosa" da variável de texto 't'
       // Assim, a verificação 'ehReciboPT' lá em baixo já não vai disparar falsamente.
       t = t.replace(/este documento não serve de factura/gi, "xxx")
            .replace(/este documento não serve de fatura/gi, "xxx");
    }

    // ------------------------------------------------------------------------
    // CASO ESPECIAL: RADIUS ("Parte" a data da fatura da data de vencimento!)
    // ------------------------------------------------------------------------
    if(textoPDF.includes("509001319") || textoPDF.includes("Radius") || textoPDF.includes("radius")){
      Logger.log("🛡️ RADIUS DETETADA: A calcular a menor data (Data Fatura).");

      // Regex para capturar qualquer data DD-MM-AAAA ou DD/MM/AAAA
      var regexDatas = /(\d{2})[-./](\d{2})[-./](\d{4})/g;
      var todasDatas = [];
      var match;

      while ((match = regexDatas.exec(textoPDF)) !== null) {
         // match[1]=Dia, match[2]=Mes, match[3]=Ano
         var d = parseInt(match[1], 10);
         var m = parseInt(match[2], 10);
         var y = parseInt(match[3], 10);
         
         // Filtro de segurança: datas entre 2020 e 2030 (ignora lixo de OCR)
         if(y >= 2020 && y <= 2030 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
            var dateObj = new Date(y, m - 1, d); // Mes em JS é 0-11
            todasDatas.push({
               str: match[0].replace(/-/g, "/").replace(/\./g, "/"),
               obj: dateObj.getTime()
            });
         }
      }

      if (todasDatas.length > 0) {
         // Ordena cronologicamente: Data Mais Antiga -> Data Mais Recente
         todasDatas.sort(function(a, b) { return a.obj - b.obj; });

         // A primeira data da lista ordenada será a Data de Emissão (ex: 27/04)
         // As seguintes serão Vencimento (ex: 12/05)
         var menorData = todasDatas[0].str;
         
         if (dataDocumento !== menorData) {
            Logger.log("✅ Data RADIUS corrigida de " + dataDocumento + " para " + menorData);
            dataDocumento = menorData;
         }
      }
    }

    // ------------------------------------------------------------------------
    // CASO: CRÉDITO AGRÍCOLA (Comissões, etc.)
    // ------------------------------------------------------------------------
    if(valorATCUD && textoPDF.includes("www.creditoagricola.pt") && (textoPDF.includes("FACTURA") || textoPDF.includes("FATURA"))){
      Logger.log("[CA FATURA] " + fileName);

      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      ({ month, year } = _extrairMesAno(dataDocumento));

      const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);

      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
      }
      continue; 
    }

    // ------------------------------------------------------------------------
    // CASO ESPECIAL: DIGITAL OCEAN (DP)
    // ------------------------------------------------------------------------
    if(CODIGO_EMPRESA==="DP" && textoPDF.includes("DigitalOcean")){
      var dateRegex5 = /(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s+(\d{4})/i;
      var match5 = textoPDF.match(dateRegex5);

      if (match5) {
        var monthMapEnglish = {"january": "01", "february": "02", "march": "03", "april": "04", "may": "05", "june": "06", "july": "07", "august": "08", "september": "09", "october": "10", "november": "11", "december": "12"};
        month = monthMapEnglish[match5[1].toLowerCase()];
        year = match5[3];

        const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);
        if (resultado.sucesso) {
           fileCount++;
           nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
        } else {
           fileErrors++;
           errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
        }
        continue;
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": DOcean data não encontrada.\n";
        continue;
      }
    }

    // ------------------------------------------------------------------------
    // CASO ESPECIAL: AVISOS DE SEGUROS (DL) --> LIXO
    // ------------------------------------------------------------------------
    if(CODIGO_EMPRESA==="DL"){
      if (
        (textoPDF.includes("O seu seguro vai ser pago por débito") && textoPDF.includes("Aviso")) ||
        (textoPDF.includes("Condições Particulares da") && !textoPDF.includes("fatura")) ||
        (textoPDF.includes("Generali") && (textoPDF.includes("Autorização de Débito Direto") || textoPDF.includes("NOTA INFORMATIVA"))) ||
        (textoPDF.includes("EXTRACTO COBRANÇAS CLIENTE"))
      ) {
        file.moveTo(pastaLixo);
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": Não é para catalogar (Lixo/Aviso) ";
        continue;
      }
    }



    // ------------------------------------------------------------------------
    // CASO 1: RECIBO DE VENCIMENTO
    // ------------------------------------------------------------------------
    if (fileName.startsWith('REC_') && !textoPDF.includes("Fatura") && !textoPDF.includes("Fatura simplificada") && !textoPDF.includes("Fatura-recibo")) {
      Logger.log("CASO 1: Recibo Vencimento - " + fileName);
      copiarMoverELog_(file, pastaGeralRecibos, sourceFolder);
      recCount++;
      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + '\n';
      continue; 
    } 

    // ------------------------------------------------------------------------
    // CASO 2: EXTRATOS DE CONTA CORRENTE
    // ------------------------------------------------------------------------
    const ehExtrato = (
        t.includes("extracto de contas correntes") || 
        t.includes("documentos de clientes por liquidar") ||
        t.includes("extrato de conta de cliente") ||
        t.includes("listagem documentos divida") ||
        t.includes("conta-corrente") ||
        t.includes("extractos por conta")
    );

    if (ehExtrato) {
       Logger.log("CASO 2: Extrato de conta corrente - " + fileName);
       // Mover diretamente para a pasta de Extratos
       var pastaExtratos = DriveApp.getFolderById(PASTA_EXTRATOS_ID);
       copiarMoverELog_(file, pastaExtratos, sourceFolder);
       
       // Adiciona ao relatório apenas como informação (não conta como erro nem como fatura)
       errosFicheirosMovidos += '\n Info: ' + fileName + " arquivado em Extratos.";
       
       continue; // Salta para o próximo ficheiro
    }    

    // ------------------------------------------------------------------------
    // CASO 3: RECIBOS NACIONAIS
    // ------------------------------------------------------------------------
    const ehReciboPT = (
      (t.includes("recibo n.º") || t.includes("recibo nº") || t.includes("recibo nr") || 
       t.includes("recebemos a quantia de") || t.includes("recebemos a importância") ||
       t.includes("este documento não serve de factura") || t.includes("este documento não serve de fatura") ||
       t.includes("recibo cliente") || t.includes("total do recibo"))
      && !t.includes("fatura/recibo") && !t.includes("fatura-recibo")
    );

    if (ehReciboPT) {
      Logger.log("[RECIBO PT] " + fileName);

      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      
      ({ month, year } = _extrairMesAno(dataDocumento));

      // --- LÓGICA INTELIGENTE IN-LOOP ---
      // Tenta descobrir se o recibo pertence a uma fatura de meses anteriores
      var origemReal = _encontrarMesOrigemDaFatura(textoPDF, year, month);
      
      if (origemReal) {
         // Se encontrou a fatura original, muda o destino para lá!
         year = origemReal.year;
         month = origemReal.month;
      }
      // -----------------------------------

      // USAR HELPER DE MOVIMENTO
      const resultado = moverParaPastaFinal_(file, year, month, "#4 - Recibos", sourceFolder);
      
      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + '\n';
        Logger.log("Movidp para: " + resultado.pasta);
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
      }
      continue; 
    }

    
    // Lixo Geral
    /*
    if(textoPDF.includes("Abaixo se discriminam as faturas em divida, que solicitamos que sejam liquidadas.")) {
      file.moveTo(pastaLixo);
      fileErrors++;
      errosFicheirosMovidos += '\n Erro ' + fileName + ": Aviso pagamento (Lixo) ";
      continue;
    }
    */




    // ------------------------------------------------------------------------
    // CASO 4: FATURA NACIONAL (Tem ATCUD, não é CA)
    // ------------------------------------------------------------------------
    if(valorATCUD && !textoPDF.includes("www.creditoagricola.pt")){
      Logger.log("CASO 4: Fatura Nacional - " + fileName);

      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      ({ month, year } = _extrairMesAno(dataDocumento));

      const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);
      
      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
      }
      continue; 
    }


    // ------------------------------------------------------------------------
    // CASO 5.1: FATURA ESTRANGEIRA (INVOICE - EN)
    // ------------------------------------------------------------------------
    if((textoPDF.includes("invoice") || textoPDF.includes("Invoice") || textoPDF.includes("INVOICE")) &&
       (!textoPDF.includes("receipt") && !textoPDF.includes("Receipt") && !textoPDF.includes("RECEIPT")) &&
       !textoPDF.includes("www.creditoagricola.pt")){
      
      Logger.log("CASO 5.1: Invoice - " + fileName);
      
      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      ({ month, year } = _extrairMesAno(dataDocumento));

      const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);

      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
      }
      continue;
    }


    // ------------------------------------------------------------------------
    // CASO 5.2: FATURA ESTRANGEIRA S/ ATCUD (PT)
    // ------------------------------------------------------------------------
    if(!valorATCUD && 
      (textoPDF.includes("factura") || textoPDF.includes("Factura") || textoPDF.includes("FACTURA") ||
       textoPDF.includes("fatura") || textoPDF.includes("Fatura") || textoPDF.includes("FATURA"))) {

      Logger.log("CASO 5.2: Fatura PT sem ATCUD - " + fileName);

      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      ({ month, year } = _extrairMesAno(dataDocumento));

      const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);

      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
      }
      continue;
    }


    // ------------------------------------------------------------------------
    // CASO 6.2: RECIBO ESTRANGEIRO
    // ------------------------------------------------------------------------
    if(!valorATCUD && (textoPDF.includes("receipt") || textoPDF.includes("Receipt") || textoPDF.includes("RECEIPT"))){
      
      Logger.log("CASO 6.2: Receipt - " + fileName);

      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      ({ month, year } = _extrairMesAno(dataDocumento));

      // --- LÓGICA INTELIGENTE IN-LOOP ---
      // Tenta descobrir se o recibo pertence a uma fatura de meses anteriores
      var origemReal = _encontrarMesOrigemDaFatura(textoPDF, year, month);
      
      if (origemReal) {
         // Se encontrou a fatura original, muda o destino para lá!
         year = origemReal.year;
         month = origemReal.month;
      }
      // -----------------------------------

      const resultado = moverParaPastaFinal_(file, year, month, "#4 - Recibos", sourceFolder);

      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
      } else {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
      }
      continue;
    }


    // ------------------------------------------------------------------------
    // CASO 7: COMPROVATIVO PAGAMENTO (CA)
    // ------------------------------------------------------------------------
    var contaadebitar = "";
    if(CODIGO_EMPRESA==="DP") contaadebitar = "Conta a Debitar: 40294310603";
    if(CODIGO_EMPRESA==="DL") contaadebitar = "Conta a Debitar: 40334199557";
    if(CODIGO_EMPRESA==="DF") contaadebitar = "Conta a Debitar: 40399078640";

    if(textoPDF.includes(contaadebitar)){ 
      
      Logger.log("CASO 7: Comprovativo CA Identificado - " + fileName);
      
      // Apenas movemos para a pasta "Inbox" dos comprovativos. 
      // A função 'catalogarComprovativosArquivo()' que corre no fim do script fará o resto.
      var pastaComprovativos = DriveApp.getFolderById(PASTA_COMPROVATIVOS_ID);
      
      copiarMoverELog_(file, pastaComprovativos, sourceFolder);
      
      // Atualizar contadores (opcional, conta como ficheiro movido com sucesso)
      fileCount++;
      nomesFicheirosMovidos += '\n Comprovativo ' + fileName + ' enviado para processamento posterior.\n';
      
      continue;
    }

  } // Fim While Loop

  // ENVIAR EMAILS DE RESUMO
  enviarResumos_(recCount, fileCount, fileErrors, nomesFicheirosMovidos, errosFicheirosMovidos);

  // Processar comprovativos em arquivo (Arquivo Morto / Cópia)
  try {
    catalogarComprovativosArquivo();
  } catch (e) {
    Logger.log("ERRO no catalogador de comprovativos: " + e);
  }
}


/**
 * ======================================================================================
 * HELPERS DE MOVIMENTO E LÓGICA CORE (O NOVO CORAÇÃO DO SCRIPT)
 * ======================================================================================
 */

// Valida se data existe e tem formato correto
function _validarData(dataDoc, fileName) {
  if (!dataDoc || dataDoc.split("/").length !== 3) {
    Logger.log("Data inválida para " + fileName + ": " + dataDoc);
    return false;
  }
  return true;
}

// Extrai objeto {month, year, day} de string DD/MM/AAAA
function _extrairMesAno(dataStr) {
  const parts = dataStr.split("/");
  // Assumindo formato DD/MM/AAAA normalizado pelos extractors
  // Se length for 4 no último, é ano.
  if (parts[2].length === 4) {
    return { day: parts[0], month: parts[1], year: parts[2] };
  } else {
    // Formato AAAA/MM/DD
    return { day: parts[2], month: parts[1], year: parts[0] };
  }
}

// Encontra a pasta do ano (com cache simples se quiseres evoluir)
function getPastaAno_(year) {
  var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();
  while (pastasFaturas.hasNext()) {
    var p = pastasFaturas.next();
    if (p.getName() === year.toString()) return p;
  }
  return null;
}

// Encontra a pasta do mês dentro da pasta do ano
function getPastaMes_(pastaAno, month, year) {
  var pastasMes = pastaAno.getFolders();
  var nomeAlvo = "Faturas_" + CODIGO_EMPRESA + "_" + month + "/" + year;
  while (pastasMes.hasNext()) {
    var p = pastasMes.next();
    if (p.getName() === nomeAlvo) return p;
  }
  return null;
}

// Lógica de verificação do Excel (extraída do loop principal)
function verificarExcelNaPasta_(pastaMes, month, year, valorATCUD) {
  var nomeFicheiroExcel = "#0 - Faturas_" + CODIGO_EMPRESA + "_" + month + "/" + year;
  var files = pastaMes.getFilesByName(nomeFicheiroExcel);
  
  if (!files.hasNext()) {
    Logger.log("⚠️ Excel não encontrado: " + nomeFicheiroExcel);
    return false;
  }
  
  var ss = SpreadsheetApp.open(files.next());
  var tabs = ["Faturas e NCs normais", "Faturas e NCs com reembolso"];
  
  for (var i = 0; i < tabs.length; i++) {
    var sheet = ss.getSheetByName(tabs[i]);
    if (!sheet) continue;
    
    // Procura coluna ATCUD (helper global existente)
    // Nota: precisa do helper encontraColunaNoCabecalho
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    var colIndex = headers.indexOf("ATCUD / Nº Documento");
    
    if (colIndex === -1) continue;
    
    var dados = sheet.getRange(3, colIndex + 1, sheet.getLastRow() - 2).getValues().flat();
    if (dados.includes(valorATCUD)) {
      Logger.log("✅ ATCUD encontrado na tab " + tabs[i]);
      return true;
    }
  }
  Logger.log("❌ ATCUD não encontrado no Excel.");
  return false;
}

/**
 * HELPER DE MOVIMENTO UNIFICADO
 * Navega: Ano -> Mês -> SubPasta (ex: #1 ou #4) -> "PARA CATALOGAR" -> Move
 * Retorna: { sucesso: boolean, pasta: string, erro: string }
 */
function moverParaPastaFinal_(file, year, month, nomeSubPasta, sourceFolder) {
  const pastaAno = getPastaAno_(year);
  if (!pastaAno) return { sucesso: false, erro: "Pasta Ano " + year + " não encontrada." };

  const pastaMes = getPastaMes_(pastaAno, month, year);
  if (!pastaMes) return { sucesso: false, erro: "Pasta Mês " + month + "/" + year + " não encontrada." };

  // Dentro do mês, procurar a sub-pasta (#1 ou #4 ou #5)
  let pastaNivel1 = null;
  const it1 = pastaMes.getFolders();
  while(it1.hasNext()) {
    let p = it1.next();
    if (p.getName() === nomeSubPasta) { pastaNivel1 = p; break; }
  }

  if (!pastaNivel1) return { sucesso: false, erro: "Sub-pasta '" + nomeSubPasta + "' não encontrada em " + pastaMes.getName() };

  // Dentro da sub-pasta, procurar "PARA CATALOGAR"
  let pastaFinal = null;
  const it2 = pastaNivel1.getFolders();
  while(it2.hasNext()) {
    let p = it2.next();
    if (p.getName() === "PARA CATALOGAR") { pastaFinal = p; break; }
  }

  if (!pastaFinal) return { sucesso: false, erro: "Pasta 'PARA CATALOGAR' não encontrada." };

  // Mover
  Logger.log("✅ Movendo ficheiro para: " + pastaMes.getName() + " > " + nomeSubPasta);
  copiarMoverELog_(file, pastaFinal, sourceFolder);
  
  return { sucesso: true, pasta: pastaMes.getName() };
}

function enviarResumos_(recCount, fileCount, fileErrors, nomesFicheirosMovidos, errosFicheirosMovidos) {
  var summary;
  if(recCount>0){
    summary = 'Movidos ' + recCount + ' recibos.\n' + nomesFicheirosMovidos;
    MailApp.sendEmail(EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL, "Distribuição Recibos ("+CODIGO_EMPRESA+")", summary);
  }  
  if(fileCount>0){
    summary = 'Movidos ' + fileCount + ' documentos fiscais.\n' + nomesFicheirosMovidos;
    MailApp.sendEmail(EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL, "Distribuição Faturas ("+CODIGO_EMPRESA+")", summary);
  }  
  if(fileErrors>0){
    summary = 'Erros: ' + fileErrors + '\n' + errosFicheirosMovidos;
    MailApp.sendEmail(EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL, "ERROS Distribuição ("+CODIGO_EMPRESA+")", summary);
  } 
}


/**
 * ======================================================================================
 * HELPERS GLOBAIS (OCR, DATAS, ETC)
 * ======================================================================================
 */

function convertPDFToText(fileId, languages) {
  if (!fileId) throw new Error("convertPDFToText: fileId em falta.");
  if (!Array.isArray(languages)) languages = [languages || "pt"];

  const file = DriveApp.getFileById(fileId);
  const mime = file.getMimeType();
  
  if (mime === MimeType.GOOGLE_DOCS || mime === "application/vnd.google-apps.document") {
    return DocumentApp.openById(fileId).getBody().getText();
  }
  if (mime.indexOf("application/vnd.google-apps.") === 0) {
    return DocumentApp.openById(fileId).getBody().getText();
  }

  let lastError = null;
  for (let i = 0; i < languages.length; i++) {
    const lang = languages[i];
    const maxTentativas = 4;
    let esperaMs = 2000;

    for (let tentativa = 1; tentativa <= maxTentativas; tentativa++) {
      let docId = null;
      try {
        const blob = file.getBlob();
        const resource = { title: "OCR_TEMP_" + file.getName() };
        const options = { ocr: true, ocrLanguage: lang || undefined };

        let ocrResult;
        try {
           ocrResult = Drive.Files.insert(resource, blob, options);
        } catch (e) {
           resource.mimeType = "application/pdf";
           ocrResult = Drive.Files.insert(resource, blob, options);
        }

        if (!ocrResult || !ocrResult.id) throw new Error("ID nulo no OCR.");
        docId = ocrResult.id;
        Utilities.sleep(500);
        const doc = DocumentApp.openById(docId);
        const textContent = doc.getBody().getText();
        DriveApp.getFileById(docId).setTrashed(true);

        if (textContent && textContent.trim().length > 0) return textContent;
        // Se vazio, tenta próxima lingua
        break; 

      } catch (e) {
        if (docId) { try { DriveApp.getFileById(docId).setTrashed(true); } catch(err) {} }
        const msg = (e && e.message) ? e.message : e.toString();
        
        if (msg.includes("User rate limit exceeded") || msg.includes("403")) {
          if (tentativa < maxTentativas) {
            Utilities.sleep(esperaMs);
            esperaMs *= 2;
            continue; 
          } else {
            lastError = e;
          }
        } else {
          lastError = e;
          break; // erro fatal nesta lingua
        }
      }
    }
  }
  return ""; // Falhou tudo ou vazio
}

function getMovimentosSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_MOVIMENTOS_ID);
  let sh = ss.getSheetByName(SHEET_MOVIMENTOS_NOME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MOVIMENTOS_NOME);
    sh.appendRow(["Data", "Hora", "Nome do ficheiro", "Fonte (nome)", "Pais do destino (nome)", "Link atual"]);
  }
  return sh;
}

function setLinkCell_(sh, row, col, text, url) {
  const rich = SpreadsheetApp.newRichTextValue().setText(text || "").setLinkUrl(url || null).build();
  sh.getRange(row, col).setRichTextValue(rich);
}

function buildDestinoParentsTail_(folder, depth) {
  const MAX_HOPS = 25;
  const chain = [];
  let current = folder;
  for (let i = 0; i < MAX_HOPS && current; i++) {
    chain.push(current);
    const parents = current.getParents();
    if (!parents.hasNext()) break;
    current = parents.next();
  }
  chain.reverse();
  const rootNames = ["o meu disco","my drive","meu drive"];
  if (chain.length && rootNames.includes(chain[0].getName().toLowerCase())) chain.shift();
  if (chain.length) chain.pop(); 
  const parentsTail = chain.slice(-depth);
  return parentsTail.map(f => f.getName()).join(" / ");
}

function registarMovimento_(fileName, fonteFolder, destinoFolder, novoFicheiroUrl) {
  const tz = Session.getScriptTimeZone() || "Europe/Lisbon";
  const agora = new Date();
  const data = Utilities.formatDate(agora, tz, "yyyy-MM-dd");
  const hora = Utilities.formatDate(agora, tz, "HH:mm:ss");
  const sh = getMovimentosSheet_();
  sh.appendRow([data, hora, fileName, "", "", ""]);
  const row = sh.getLastRow();
  if (fonteFolder) setLinkCell_(sh, row, 4, fonteFolder.getName(), fonteFolder.getUrl());
  if (destinoFolder) {
    const caminhoCurto = buildDestinoParentsTail_(destinoFolder, DESTINO_PATH_DEPTH);
    setLinkCell_(sh, row, 5, caminhoCurto, destinoFolder.getUrl());
  }
  if (novoFicheiroUrl) setLinkCell_(sh, row, 6, "LINK", novoFicheiroUrl);
}

function copiarMoverELog_(file, destinoFolder, fonteFolder) {
  const novoFicheiro = file.makeCopy(destinoFolder);
  file.moveTo(DriveApp.getFolderById(PASTA_LIXO));
  registarMovimento_(file.getName(), fonteFolder, destinoFolder, novoFicheiro.getUrl());
  return novoFicheiro;
}

function extractTipoDocumento(text) {
  text = text.replace(/[/_-]/g, ' ').replace(/\s{2,}/g, ' '); 
  text = text.replace(/válido como recibo após/g, ''); 
  var tipos = ['fatura simplificada', 'factura simplificada', 'nota de crédito', 'fatura recibo', 'factura recibo', '2ª via', 'segunda via', 'fatura', 'factura', 'recibo de renda', 'recibo'];
  var tiposOutput = ['Fatura simplificada', 'Fatura simplificada', 'Nota de crédito', 'Fatura-recibo', 'Fatura-recibo', '2ª via fatura', '2ª via fatura', 'Fatura', 'Fatura', 'Recibo de renda', 'Recibo'];
  for (var i = 0; i < tipos.length; i++) {
    if (text.toLowerCase().includes(tipos[i])) return tiposOutput[i];
  }
  return null;
}

function extractATCUD(pdfText) {
  if (!pdfText) return null;
  var regex = /ATCUD:\s*([^\s]+)/;
  var match = pdfText.match(regex);
  if (match) return match[1];
  regex = /ATCUD\s*([^\s]+)/;
  match = pdfText.match(regex);
  if (match) return match[1];
  return null;
}

function extrairATCUDRecibosCA(texto) {
  var regex = /Informação Complementar:\s*([A-Z0-9]{2,}-[A-Z0-9]+)/i;
  var match = texto.match(regex);
  if (match) return match[1];
  return null;
}

function extractDataDocumentoTaloes(pdfText) {
  if (!pdfText) return null;
  const injectSpaces = s => String(s).replace(/\u00A0/g, ' ').replace(/([A-Za-z])(\d)/g, '$1 $2').replace(/(\d)([A-Za-z])/g, '$1 $2');
  const raw = injectSpaces(pdfText);
  const linesAll = raw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const badPOS = [/\b(pos|tpa|tp[ãa]g|terminal|redeunic|multibanco|mb\s*way|sibs)\b/i, /\bvisa\b/i, /\bmastercard\b/i, /\bmaestro\b/i, /\bam[ex|erican\s*express]\b/i, /\bautoriz[aç][aã]o\b/i, /\bauth(?:orization)?\b/i, /\baid\b/i, /\batc\b/i, /\btid\b/i, /\bnsu\b/i, /\bpan\b/i, /\barqc?\b/i, /\bcomprovativo\b/i, /\breceb[ií]do\b/i, /\bmerchant\s*copy\b/i, /\bclient\s*copy\b/i, /\blote\b/i, /\bref(?:\.|er[eê]ncia)?\b/i, /\btransa[cç][aã]o\b/i, /\bvenda\b/i, /\bpagamento\b/i];
  const timeRe = /\b\d{2}:\d{2}(?::\d{2})?\b/; 
  const lines = linesAll.filter(l => !isHardBad(l));
  const headerRe = /\b(?:fatura\/recibo|fatura|factura|nota\s+de\s+cr[eé]dito)\b/i;
  const dataFieldRe = /\bdata\s*:\s*(\d{2})[./-](\d{2})[./-](\d{4})\b/i;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (headerRe.test(l)) {
      const look = [l, lines[i+1], lines[i+2], lines[i+3]].filter(Boolean);
      for (const seg of look) {
        const m = seg.match(dataFieldRe);
        if (m) {
          const best = _safeDate_(m[1], m[2], m[3]);
          if (best) return best;
        }
      }
    }
  }

  const goodLabels = [
    /data\s*(?:de)?\s*emiss[aã]o/i, 
    /\bemiss[aã]o\b/i, 
    /\bemitida\b/i, 
    /\bemitida\s+em\b/i,
    /dt\.?\s*emiss[aã]o/i,
    /\bdata\s*doc(?:umento)?\b/i,
    /\bdata\s*:\b/i,
    /\bdata\s*da\s*fatura\b/i,
    /\binvoice\s+date\b/i, 
    /\bissue\s+date\b/i, 
    /\bfecha\s+de\s+emisi[oó]n\b/i
  ];
  const badLabels = [/\bvig[êe]ncia\b/i, /\bper[ií]odo\b/i, /\bvalidade\b/i, /\bcompet[êe]ncia\b/i, /\bref\.\s*(?:per[ií]odo|m[êe]s)\b/i, /\bintervalo\b/i, /\bvenc(?:imento)?\b/i, /\bprazo\b/i, /\bdue\b/i, /\bpayment\b/i];
  const hasRangeRe = /\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*(?:a|–|—|-)\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/i;

  const oneLine = raw.replace(/\s+/g, ' ');
  const isoNearLabel = new RegExp('(?:' + goodLabels.map(r=>r.source).join('|') + ')' + '[^0-9]{0,40}(20\\d{2})[./-](\\d{2})[./-](\\d{2})','i');
  let m = oneLine.match(isoNearLabel);
  if (m) {
    const fast = _safeDate_(m[3], m[2], m[1]);
    if (fast) return fast;
  }

  const rx = [
    { re:/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/g, norm:(y,m,d)=>({d,m,y, iso:true}) },
    { re:/\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/g,     norm:(d,m,y)=>({d,m,y}) },
    { re:/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\b/g,         norm:(d,m,y)=>({d,m,y, ambiguousYY:true}) },
    { re:/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})\b/gi, norm:(d,mon,y)=>({d, m:_monPT_(mon), y}) },
    { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})\b/gi, norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
    { re:/\b(\d{1,2})[.\- ](jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[.\- ](\d{4})\b/gi, norm:(d,mon,y)=>({d, m:_monPTabbrev_(mon), y}) }
  ];

  const candidates = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line) continue;
    if (hasRangeRe.test(line)) continue;
    if (badLabels.some(r => r.test(line))) continue;

    const looksPOS = badPOS.some(r => r.test(line)) || timeRe.test(line);
    const hasGood = goodLabels.some(r => r.test(line));

    for (const {re, norm} of rx) {
      re.lastIndex = 0;
      let mm;
      while ((mm = re.exec(line)) !== null) {
        const p = norm(...mm.slice(1));
        const dd2 = String(p.d).padStart(2,'0');
        const mm2 = String(p.m).padStart(2,'0');
        let yyyy = String(p.y);

        if ((p.ambiguousYY || yyyy.length === 2) && looksPOS && !hasGood) continue;

        if (p.ambiguousYY || yyyy.length === 2) {
          const reconciled = _reconcileYearWithNearbyISO(line, mm.index, dd2, mm2, 80);
          if (reconciled) {
            yyyy = reconciled;
          } else {
            if (!hasGood && looksPOS) continue;
            const n = +yyyy; yyyy = (n <= 79) ? (2000 + n) : (1900 + n);
          }
        }
        const safe = _safeDate_(dd2, mm2, yyyy);
        if (!safe) continue;

        const near = goodLabels.some(r => {
          r.lastIndex = 0;
          const lm = r.exec(line);
          return lm && lm.index >= 0 && (mm.index - lm.index) <= 30;
        });

        let score = 0;
        if (p.iso) score += 60;          
        if (hasGood) score += 120;
        if (near) score += 20;
        if (i < 12) score += 30;
        if (line.length <= 80) score += 5;
        if (looksPOS && !hasGood) score -= 150;
        if (score > 0) candidates.push({ safe, score, lineIndex: i, colIndex: mm.index });
      }
    }
  }
  if (candidates.length) {
    candidates.sort((a,b)=>{
      if (b.score !== a.score) return b.score - a.score;
      if (a.lineIndex !== b.lineIndex) return a.lineIndex - b.lineIndex;
      return a.colIndex - b.colIndex;
    });
    return candidates[0].safe;
  }
  return null;
}

// Helpers utilitários
function isHardBad(line) {
  const lo = line.toLowerCase();
  if (/\b2\s*(?:ª|a|\.ª)?\s*via\b/.test(lo)) return true;
  if (/\bsegunda\s+via\b/.test(lo)) return true;
  if (/\bduplicad[oa]\b/.test(lo)) return true;
  if (/\breimpress[ãa]o\b/.test(lo)) return true;
  if (/\bc[óo]pia\b/.test(lo)) return true;
  if (/\bvia\b.{0,20}\bgerad[ao]?\b/.test(lo)) return true;
  if (/\bgerad[ao]\s+em\b/.test(lo)) return true;
  if (/\ba\s+partir\s+d[eo]\b/.test(lo)) return true;
  return false;
}
function _reconcileYearWithNearbyISO(line, matchIndex, dd, mm, windowChars) {
  const around = windowChars || 60;
  const left  = Math.max(0, matchIndex - around);
  const right = Math.min(line.length, matchIndex + around);
  const ctx = line.slice(left, right);
  const iso = ctx.match(/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/);
  if (!iso) return null;
  const isoY = iso[1], isoM = String(iso[2]).padStart(2,'0'), isoD = String(iso[3]).padStart(2,'0');
  if ((dd === isoD && mm === isoM) || (dd === isoM && mm === isoD)) return isoY;
  return null;
}
function _safeDate_(dd, mm, yyyy) {
  const y = Number(yyyy), m = Number(mm), d = Number(dd);
  if (!y || !m || !d || y < 2000) return null;
  const dt = new Date(y, m-1, d);
  const today = new Date();
  if (isNaN(dt.getTime()) || dt > today) return null;
  if (dt.getFullYear()!==y || (dt.getMonth()+1)!==m || dt.getDate()!==d) return null;
  return `${String(d).padStart(2,'0')}/${String(m).padStart(2,'0')}/${String(y)}`;
}
function _monPT_(s){ const m={janeiro:1,fevereiro:2,'março':3,marco:3,abril:4,maio:5,junho:6,julho:7,agosto:8,setembro:9,outubro:10,novembro:11,dezembro:12}; return m[String(s).toLowerCase()]||s; }
function _monPTabbrev_(s){ const m={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12}; return m[String(s).toLowerCase().slice(0,3)]||s; }
function _monEN_(s){ const m={january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12}; return m[String(s).toLowerCase()]||s; }
function normYear(y){ y=String(y); if (y.length===2){ const n=+y; return (n<=79?2000+n:1900+n);} return y; }
function to2(n){ return String(n).padStart(2,'0'); }

/**
 * ======================================================================================
 * COPIAR COMPROVATIVOS PARA ARQUIVO
 * ======================================================================================
 */
function catalogarComprovativosArquivo() {
  var pastaComprovativos = DriveApp.getFolderById(PASTA_COMPROVATIVOS_ID);
  var pdfFiles = getPDFFilesInFolder(pastaComprovativos);
  Logger.log("Comprovativos encontrados: " + pdfFiles.length);

  for (var i = 0; i < pdfFiles.length; i++) {
    var file = pdfFiles[i];
    try {
      var texto = convertPDFToText(file.getId(), ['pt', 'en', null]);
      var atcud = extractATCUDFromText(texto);
      if (!atcud) { Logger.log("Sem ATCUD: " + file.getName()); continue; }

      var dataPagamentoStr = extractDateFromPayslip(texto);
      var anoPagamento = inferYearFromDateString(dataPagamentoStr) || new Date().getFullYear();

      var match = procurarFaturaPorATCUDNoArquivo(atcud, anoPagamento);

      if (match) {
        var novoNome = "COMP" + match.numeroDocumento + ".pdf";
        var ssFile = DriveApp.getFileById(match.spreadsheetId);
        var parents = ssFile.getParents();
        if (!parents.hasNext()) continue;

        var pastaMes = parents.next();
        var itComp = pastaMes.getFoldersByName("#5 - Comprovativos de pagamento");
        var pasta5 = itComp.hasNext() ? itComp.next() : pastaMes.createFolder("#5 - Comprovativos de pagamento");
        var itPara = pasta5.getFoldersByName("PARA CATALOGAR");
        var pastaParaCatalogar = itPara.hasNext() ? itPara.next() : pasta5.createFolder("PARA CATALOGAR");

        var copia = copiarMoverELog_(file, pastaParaCatalogar, pastaComprovativos);
        copia.setName(novoNome);
        Logger.log("Comprovativo arquivado: " + novoNome);
      }
    } catch (e) {
      Logger.log("Erro comprovativo " + file.getName() + ": " + e);
    }
  }
}

function getPDFFilesInFolder(folder) {
  var files = folder.getFiles(), pdfFiles = [];
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === "application/pdf") pdfFiles.push(file);
  }
  return pdfFiles;
}

function extractATCUDFromText(pdfText) {
  if (!pdfText) return null;
  var regex1 = /ATCUD:\s*([^\s]+)/i, match1 = pdfText.match(regex1);
  if (match1) return match1[1].trim();
  var regex2 = /ATCUD\s+([^\s]+)/i, match2 = pdfText.match(regex2);
  if (match2) return match2[1].trim();
  var text = pdfText.replace(/\s+/g, " ").toUpperCase();
  var regex3 = /\b([A-Z0-9]{8,}-\d{2,})\b/, match3 = text.match(regex3);
  if (match3) return match3[1].trim();
  return null;
}

function extractDateFromPayslip(content) {
  if (!content) return null;
  var words = content.split(/\s+/);
  for (var i = 0; i < words.length; i++) {
    var word = words[i];
    if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(word)) { var p=word.split(/[-/]/); return p[2]+"/"+p[1]+"/"+p[0]; }
    if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(word)) { var p=word.split(/[-/]/); return p[0]+"/"+p[1]+"/"+p[2]; }
  }
  var m = content.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) return ("0"+m[1]).slice(-2)+"/"+("0"+m[2]).slice(-2)+"/"+m[3];
  return null;
}

function inferYearFromDateString(dateStr) {
  if (!dateStr) return null;
  var parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) { var a=parseInt(parts[2],10); if (!isNaN(a) && a>1900 && a<2100) return a; }
  var m = dateStr.match(/(19\d{2}|20\d{2})/);
  if (m) return parseInt(m[1], 10);
  return null;
}

function procurarFaturaPorATCUDNoArquivo(atcud, anoPagamento) {
  if (!atcud) return null;
  var atcudNormalizado = String(atcud).replace(/\s/g, "");
  var yearWrappers = getYearFolders(); 
  if (!yearWrappers || !yearWrappers.length) return null;
  var candidatos = [];
  for (var i=0; i<yearWrappers.length; i++) { if (yearWrappers[i].year <= anoPagamento) candidatos.push(yearWrappers[i]); }
  if (!candidatos.length) candidatos = yearWrappers.slice();

  for (var a=0; a<candidatos.length; a++) {
    var wrapper = candidatos[a];
    var pastaAno = wrapper.folder;
    var itMeses = pastaAno.getFolders();
    while (itMeses.hasNext()) {
      var pastaMes = itMeses.next();
      var nm = pastaMes.getName();
      if (nm.indexOf("Faturas_") !== 0) continue; 
      var files = pastaMes.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
        try {
          var ss = SpreadsheetApp.openById(f.getId());
          var match = procurarATCUDNasAbasDeFaturas(ss, atcudNormalizado);
          if (match) {
            match.ano = wrapper.year;
            match.spreadsheetId = f.getId();
            match.spreadsheetName = f.getName();
            match.pastaMesNome = nm;
            return match;
          }
        } catch (e) { }
      }
    }
  }
  return null;
}

function getYearFolders() {
  if (__cacheYearFolders) return __cacheYearFolders;
  var root = DriveApp.getFolderById(PASTA_GERAL_FATURAS);
  var it = root.getFolders(), result = [];
  while (it.hasNext()) {
    var f = it.next(), name = f.getName();
    if (/^\d{4}$/.test(name)) result.push({ year: parseInt(name, 10), folder: f });
  }
  result.sort(function(a, b) { return b.year - a.year; });
  __cacheYearFolders = result;
  return result;
}

function procurarATCUDNasAbasDeFaturas(ss, atcudNormalizado) {
  var nomesAbas = ["Faturas e NCs normais", "Faturas e NCs com reembolso", "Outros documentos"];
  for (var i = 0; i < nomesAbas.length; i++) {
    var sheet = ss.getSheetByName(nomesAbas[i]);
    if (!sheet) continue;
    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha <= 2) continue;

    var colATCUD = encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", 2);
    var colComp = encontraColunaNoCabecalho(sheet, "Comprovativo de pagamento", 2);
    var colNumDoc = encontraColunaNoCabecalho(sheet, "Nº", 2);
    if (colNumDoc < 0) colNumDoc = encontraColunaNoCabecalho(sheet, "Número do documento", 2);

    if (colATCUD < 0 || colComp < 0 || colNumDoc < 0) continue;

    for (var row = 3; row <= ultimaLinha; row++) {
      var atcudLinha = sheet.getRange(row, colATCUD).getDisplayValue();
      if (!atcudLinha) continue;
      if (String(atcudLinha).replace(/\s/g, "") !== atcudNormalizado) continue;
      if (!sheet.getRange(row, colComp).isBlank()) continue;

      return { sheetName: nomesAbas[i], row: row, numeroDocumento: sheet.getRange(row, colNumDoc).getValue() };
    }
  }
  return null;
}

function encontraColunaNoCabecalho(sheet, columnName, linhaDoCabecalho) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return -1;
  var headerRowValues = sheet.getRange(linhaDoCabecalho, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headerRowValues.length; i++) {
    if (headerRowValues[i] === columnName) return i + 1;
  }
  return -1;
}

/**
 * Procura nos registos (Exceis) passados usando heurística avançada:
 * 1. Fornecedor (Nome) Match (Opcional/Reforço)
 * 2. Valor (Price) Match (Obrigatório)
 * 3. ID Parcial (Sequencial) Match (Obrigatório para desempatar)
 */
/**
 * Procura nos registos (Exceis) passados usando heurística avançada:
 * 1. Fornecedor (Nome) Match (Opcional/Reforço)
 * 2. Valor (Price) Match (Obrigatório)
 * 3. ID Parcial (Sequencial) Match (Obrigatório para desempatar)
 */
function _encontrarMesOrigemDaFatura(textoPdf, anoRecibo, mesRecibo) {
  Logger.log("🕵️ [SMART MATCH] A iniciar scan para Recibo de " + mesRecibo + "/" + anoRecibo);
  
  // Limpa o texto PDF e normaliza
  var textoNorm = textoPdf.toUpperCase().replace(/\s+/g, " ");
  
  var janelaMeses = 24; 
  var dataBase = new Date(anoRecibo, parseInt(mesRecibo)-1, 1);

  for (var i = 0; i < janelaMeses; i++) {
    var d = new Date(dataBase);
    d.setMonth(dataBase.getMonth() - i);
    var checkYear = d.getFullYear();
    var checkMonth = ("0" + (d.getMonth() + 1)).slice(-2);
    
     Logger.log("   > [" + (i+1) + "/24] A ver mês: " + checkMonth + "/" + checkYear);

    // 1. Obter Pastas
    var pAno = getPastaAno_(checkYear);
    if (!pAno) continue;
    var pMes = getPastaMes_(pAno, checkMonth, checkYear);
    if (!pMes) continue;

    var nomeFicheiroExcel = "#0 - Faturas_" + CODIGO_EMPRESA + "_" + checkMonth + "/" + checkYear;
    var files = pMes.getFilesByName(nomeFicheiroExcel);
    if (!files.hasNext()) {
        Logger.log("     ⚠️ Excel não encontrado: " + nomeFicheiroExcel);
       continue;
    }

    try {
      var fileExcel = files.next();
      var ss = SpreadsheetApp.open(fileExcel);
      var abas = ["Faturas e NCs normais", "Faturas e NCs com reembolso", "Outros documentos"];
      
      for (var k = 0; k < abas.length; k++) {
        var sheet = ss.getSheetByName(abas[k]);
        if (!sheet) continue;
        
        var lastRow = sheet.getLastRow();
        if (lastRow <= 2) continue;
        
        // --- MAPEAMENTO DE COLUNAS ---
        var colEntidade = encontraColunaNoCabecalho(sheet, "Fornecedor", 2);
        
        var colValor = encontraColunaNoCabecalho(sheet, "Valor total", 2);

        var colNum = encontraColunaNoCabecalho(sheet, "Nº", 2);
        if (colNum < 0) colNum = encontraColunaNoCabecalho(sheet, "Número do documento", 2);
        
        var colATCUD = encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", 2);
        
        // Debug das colunas encontradas (descomenta se achares que ele não está a ler as colunas)
         Logger.log("     Aba '" + abas[k] + "' Cols: Ent=" + colEntidade + " Val=" + colValor + " Num=" + colNum + " ATCUD=" + colATCUD);

        if (colValor < 0 && colNum < 0 && colATCUD < 0) continue;
        
        // Ler dados em memória
        var dados = sheet.getRange(3, 1, lastRow-2, sheet.getLastColumn()).getValues();
        
        for (var r = 0; r < dados.length; r++) {
          var rowData = dados[r];
          
          // === PASSO 1: VERIFICAR VALOR (O filtro mais forte) ===
          // === PASSO 1: VERIFICAR VALOR (Robusto com Espaços e Milhares) ===
          var valorMatch = false;
          var valorExcel = rowData[colValor-1];
            
          // Garante que tratamos como número
          var valorFloat = null;
          if (typeof valorExcel === 'number') {
             valorFloat = valorExcel;
          } else if (typeof valorExcel === 'string') {
             // Tenta limpar "1.000,00 €" para 1000.00
             var limpo = valorExcel.replace(/[^0-9.,-]/g, "").replace(",", "."); 
             // Se tiver múltiplos pontos, é chato, mas o parseFloat costuma lidar com o formato padrão JS
             valorFloat = parseFloat(limpo);
          }

          if (valorFloat !== null && !isNaN(valorFloat)) {
             // 1. Formato Base: "1000.00" e "1000,00"
             var vRaw = valorFloat.toFixed(2); // "1000.00"
             var vVirgula = vRaw.replace(".", ","); // "1000,00"

             // 2. Formato com Separador de Milhares (Espaço): "1 000,00"
             // Regex insere espaço a cada 3 digitos na parte inteira
             var partes = vVirgula.split(",");
             var inteiroComEspacos = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
             var vComEspaco = inteiroComEspacos + "," + partes[1]; // "1 000,00"

             // 3. Formato com Separador de Milhares (Ponto): "1.000,00"
             var inteiroComPontos = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
             var vComPonto = inteiroComPontos + "," + partes[1]; // "1.000,00"
             
             // DEBUG (opcional, para veres o que ele está a testar)
             //Logger.log("Testando valores: " + vRaw + " | " + vVirgula + " | " + vComEspaco);

             if (textoPdf.includes(vRaw) || 
                 textoPdf.includes(vVirgula) || 
                 textoPdf.includes(vComEspaco) || 
                 textoPdf.includes(vComPonto)) {
                 valorMatch = true;
             }
          }

          // === PASSO 2: VERIFICAR FORNECEDOR (Opcional mas ajuda no log) ===
          var entidadeMatch = false;
          if (colEntidade > 0) {
            var entidadeExcel = String(rowData[colEntidade-1]).toUpperCase().trim();
            if (entidadeExcel.length > 2) {
               var primeiraPalavra = entidadeExcel.split(" ")[0];
               if (primeiraPalavra.length > 2 && textoNorm.includes(primeiraPalavra)) {
                 entidadeMatch = true;
               }
            }
          }

          // === PASSO 3: VERIFICAR ID PARCIAL ===
          var idMatch = false;
          var idEncontrado = "";

          // Tenta pelo Nº Fatura
          if (colNum > 0) {
             var numDoc = String(rowData[colNum-1]).toUpperCase().trim();
             var sequencial = extractSequencial_(numDoc);
             if (sequencial && sequencial.length >= 2) {
               if (textoNorm.includes(sequencial)) {
                 idMatch = true;
                 idEncontrado = "Nº " + numDoc + " (Seq: " + sequencial + ")";
               }
             }
          }

          // Tenta pelo ATCUD
          if (!idMatch && colATCUD > 0) {
             var atcud = String(rowData[colATCUD-1]).toUpperCase().trim();
             var seqAtcud = extractSequencial_(atcud);
             if (seqAtcud && seqAtcud.length >= 2) {
               if (textoNorm.includes(seqAtcud)) {
                 idMatch = true;
                 idEncontrado = "ATCUD " + atcud + " (Seq: " + seqAtcud + ")";
               }
             }
          }

          // === DECISÃO FINAL ===
          // Exige: Valor E (ID ou Entidade Forte)
          // Mas na tua lógica pediste ID explicitamente.
          if (valorMatch && idMatch) {
            Logger.log("✅ SMART MATCH CONFIRMADO!");
            Logger.log("   -> Ficheiro: " + checkMonth + "/" + checkYear);
            Logger.log("   -> Valor: " + rowData[colValor-1]);
            Logger.log("   -> ID Validado: " + idEncontrado);
            return { year: String(checkYear), month: String(checkMonth) };
          } else if (valorMatch && !idMatch) {
            Logger.log("      ❌ Valor bateu, mas ID não encontrado no PDF. (Excel ID: " + (colNum>0?rowData[colNum-1]:"N/A") + ")");
          }
        }
      }
    } catch(e) {
      Logger.log("❌ Erro ao ler excel " + checkMonth + "/" + checkYear + ": " + e.message);
    }
  }
  
  Logger.log("🏁 Sem match heurístico nos últimos " + janelaMeses + " meses.");
  return null;
}

// Helper para extrair a parte final (sequencial) de uma fatura ou ATCUD
function extractSequencial_(str) {
  if (!str) return null;
  // Divide por /, -, ou espaço
  var parts = str.split(/[\/\-\s]/);
  // Pega na última parte que seja numérica
  for (var i = parts.length - 1; i >= 0; i--) {
    var p = parts[i].replace(/[^0-9]/g, ""); // limpa letras
    if (p.length > 0) return p;
  }
  return null;
}