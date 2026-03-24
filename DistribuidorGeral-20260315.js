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
var __cachePastaAno = {};
var __cachePastaMes = {};

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
    valorATCUD = extractATCUD(textoPDF);
    let t = (textoPDF || "").toLowerCase();

    // =====================================================================
    // PASSO 1: CONSENSO DE DATA (5 fontes → data mais votada)
    // =====================================================================
    let consenso = _consensoData(fileName, textoPDF);
    dataDocumento = consenso.data;

    // Correção Radius: se Radius, forçar menor data (override consenso)
    let dataRadius = _corrigirDataRadius(textoPDF, t);
    if (dataRadius && dataRadius !== dataDocumento) {
      Logger.log("Data RADIUS corrigida de " + dataDocumento + " para " + dataRadius);
      dataDocumento = dataRadius;
    }

    // =====================================================================
    // PASSO 2: CLASSIFICAÇÃO DE TIPO VIA IA
    // =====================================================================
    let iaTipo = null;
    try {
      let iaResult = classificarDocumentoViaIA(textoPDF.substring(0, 4000));
      if (iaResult && iaResult.tipo) {
        iaTipo = iaResult.tipo.toLowerCase().trim();
        Logger.log("IA tipo: " + iaTipo + " | " + fileName);

        // --- Lixo ---
        if (iaTipo === "lixo") {
          file.moveTo(pastaLixo);
          fileErrors++; errosFicheirosMovidos += '\n ' + fileName + ": Lixo (IA).\n";
          continue;
        }

        // --- Recibo de Vencimento ---
        if (iaTipo === "recibo_vencimento") {
          copiarMoverELog_(file, pastaGeralRecibos, sourceFolder);
          recCount++; nomesFicheirosMovidos += '\n Ficheiro ' + fileName + '\n';
          Logger.log("✅ IA → Recibo Vencimento: " + fileName);
          continue;
        }

        // --- Extrato ---
        if (iaTipo === "extrato") {
          copiarMoverELog_(file, DriveApp.getFolderById(PASTA_EXTRATOS_ID), sourceFolder);
          errosFicheirosMovidos += '\n Info: ' + fileName + " arquivado em Extratos (IA).\n";
          Logger.log("✅ IA → Extrato: " + fileName);
          continue;
        }

        // --- Comprovativo ---
        if (iaTipo === "comprovativo") {
          copiarMoverELog_(file, DriveApp.getFolderById(PASTA_COMPROVATIVOS_ID), sourceFolder);
          fileCount++; nomesFicheirosMovidos += '\n Comprovativo ' + fileName + ' (IA)\n';
          Logger.log("✅ IA → Comprovativo: " + fileName);
          continue;
        }

        // --- Fatura / Nota de Crédito (com data do consenso) ---
        if ((iaTipo === "fatura" || iaTipo === "nota_credito") && _validarData(dataDocumento, fileName)) {
          ({ month, year } = _extrairMesAno(dataDocumento));
          const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);
          if (resultado.sucesso) {
            fileCount++; nomesFicheirosMovidos += '\n ' + fileName + ' → ' + resultado.pasta + ' (IA: ' + iaTipo + ', ' + consenso.votos + ' votos)\n';
            Logger.log("✅ IA → " + iaTipo + ": " + fileName);
          } else {
            fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
          }
          continue;
        }

        // --- Recibo (com data do consenso + Smart Match) ---
        if (iaTipo === "recibo" && _validarData(dataDocumento, fileName)) {
          ({ month, year } = _extrairMesAno(dataDocumento));
          var origemReal = _encontrarMesOrigemDaFatura(textoPDF, year, month);
          if (origemReal) { year = origemReal.year; month = origemReal.month; }
          const resultado = moverParaPastaFinal_(file, year, month, "#4 - Recibos", sourceFolder);
          if (resultado.sucesso) {
            fileCount++; nomesFicheirosMovidos += '\n ' + fileName + ' → ' + resultado.pasta + ' (IA: recibo, ' + consenso.votos + ' votos)\n';
            Logger.log("✅ IA → Recibo: " + fileName);
          } else {
            fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": " + resultado.erro + "\n";
          }
          continue;
        }

        Logger.log("IA tipo=" + iaTipo + " mas sem data válida → fallback regex: " + fileName);
      }
    } catch (eClassif) {
      Logger.log("IA classificação falhou → fallback regex: " + String(eClassif).substring(0, 100));
    }

    // ------------------------------------------------------------------------
    // CASO ESPECIAL: MEO (Escreve na fatura "Este documento não serve de fatura"!)
    // ------------------------------------------------------------------------
    
    // Truque de Mestre:
    const ehMEO = textoPDF.includes("504 615 947");
    if(ehMEO) Logger.log("🛡️ MEO DETETADA: A neutralizar hipótese de recibo.");

    // NOTA: Correção Radius já feita no Passo 1 (via _corrigirDataRadius)

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

    // NOTA: Digital Ocean — data já determinada pelo consenso (IAs lidam com datas em inglês)

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
      Logger.log("✅ CASO 1: Recibo Vencimento - " + fileName);
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
       Logger.log("✅ CASO 2: Extrato de conta corrente - " + fileName);
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
    
    var ehReciboPT = (
      (
        t.includes("recibo n.º") || t.includes("recibo nº") || t.includes("recibo nr") || 
        t.includes("recebemos a quantia de") || t.includes("recebemos a importância") ||
        t.includes("recibo cliente") || t.includes("total do recibo") ||
        
        // Mantemos a frase da MEO aqui, o filtro vem depois
        t.includes("este documento não serve de factura") || t.includes("este documento não serve de fatura")
      )
      
      // Filtros de Segurança
      && !t.includes("fatura/recibo") 
      && !t.includes("fatura-recibo")
    );
    
    // ------------------------------------------------------------------------
    // [Passo 2] EXCLUSÃO GLOBAL (A MÁGICA FINAL)
    // ------------------------------------------------------------------------
    // Se ehMEO for TRUE, o resultado da linha é FALSE, anulando ehReciboPT.
    // Se ehMEO for FALSE, o resultado é ehReciboPT (mantém o que deu no Passo 1).
    ehReciboPT = ehReciboPT && !ehMEO;

    if(ehReciboPT) {
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
        Logger.log("✅ CASO 3: Recibo PT movido para: " + resultado.pasta);
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
      Logger.log("✅ CASO 4: Fatura Nacional - " + fileName);

      if(!_validarData(dataDocumento, fileName)) {
        fileErrors++; errosFicheirosMovidos += '\n Erro ' + fileName + ": Data inválida.\n"; continue;
      }
      ({ month, year } = _extrairMesAno(dataDocumento));

      Logger.log("Diagnóstico: Tentando mover para pasta " + month + "/" + year);

      const resultado = moverParaPastaFinal_(file, year, month, "#1 - Faturas e NCs normais", sourceFolder);
      
      if (resultado.sucesso) {
        fileCount++;
        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + resultado.pasta + '\n';
      } else {
        Logger.log("ERRO AO MOVER: " + resultado.erro); // <-- GARANTA QUE ESTA LINHA TEM LOG.
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
      
      Logger.log("✅ CASO 5.1: Invoice - " + fileName);
      
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

      Logger.log("✅ CASO 5.2: Fatura PT sem ATCUD - " + fileName);

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
      
      Logger.log("✅ CASO 6.2: Receipt - " + fileName);

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
      
      Logger.log("✅ CASO 7: Comprovativo CA - " + fileName);
      
      // Apenas movemos para a pasta "Inbox" dos comprovativos. 
      // A função 'catalogarComprovativosArquivo()' que corre no fim do script fará o resto.
      var pastaComprovativos = DriveApp.getFolderById(PASTA_COMPROVATIVOS_ID);
      
      copiarMoverELog_(file, pastaComprovativos, sourceFolder);
      
      // Atualizar contadores (opcional, conta como ficheiro movido com sucesso)
      fileCount++;
      nomesFicheirosMovidos += '\n Comprovativo ' + fileName + ' enviado para processamento posterior.\n';
      
      continue;
    }

    // ------------------------------------------------------------------------
    // CASO DEFAULT: Ficheiro não classificado por nenhum método
    // ------------------------------------------------------------------------
    fileErrors++;
    errosFicheirosMovidos += '\n ' + fileName + ": Não classificado (sem caso aplicável, IA e regex falharam).\n";
    Logger.log("SEM CASO: " + fileName);

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
 * HELPERS DE IA
 * ======================================================================================
 */
function chamarGroq(prompt) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty("GROQ_API_KEY");
  const url = "https://api.groq.com/openai/v1/chat/completions";

  const payload = {
    model: "meta-llama/llama-4-scout-17b-16e-instruct",
    messages: [{ role: "user", content: prompt }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(json.error.message || JSON.stringify(json.error));
  }
  if (!json.choices || !json.choices.length) {
    throw new Error("Resposta inesperada: " + JSON.stringify(json));
  }

  return json.choices[0].message.content;
}

function chamarMistral(prompt) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty("MISTRAL_API_KEY");
  const url = "https://api.mistral.ai/v1/chat/completions";

  const payload = {
    model: "mistral-small-latest",
    messages: [{ role: "user", content: prompt }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error("Mistral: " + (json.error.message || JSON.stringify(json.error)));
  }
  if (!json.choices || !json.choices.length) {
    throw new Error("Mistral resposta inesperada: " + JSON.stringify(json));
  }

  return json.choices[0].message.content;
}

function chamarGemini(prompt, modelo) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!API_KEY) throw new Error("GEMINI_API_KEY não configurada");
  var modelId = modelo || "gemini-2.0-flash";
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelId + ":generateContent?key=" + API_KEY;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0, maxOutputTokens: 100 }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error("Gemini: " + (json.error.message || JSON.stringify(json.error)));
  }

  var text = "";
  try { text = json.candidates[0].content.parts[0].text; } catch (e) {
    throw new Error("Gemini resposta inesperada: " + JSON.stringify(json).substring(0, 200));
  }
  return text;
}

function extrairDataViaIA(texto) {
  const prompt = "Extraia apenas a data de emissão do seguinte texto de um documento.\n" +
    "Retorne SOMENTE a data no formato DD/MM/AAAA, sem mais nada.\n" +
    'Se não encontrar, retorne "Não encontrada".\n\nTexto:\n' + texto;

  // Tenta Mistral primeiro, Groq como fallback
  try {
    var resultado = chamarMistral(prompt);
    Logger.log("IA utilizada: Mistral (Small)");
    return resultado;
  } catch (eMistral) {
    Logger.log("Mistral falhou: " + String(eMistral).substring(0, 100) + " → a tentar Groq...");
    var resultado = chamarGroq(prompt);
    Logger.log("IA utilizada: Groq (Llama 4 Scout)");
    return resultado;
  }
}

/**
 * ======================================================================================
 * ALGORITMO DE CONSENSO DE DATAS (5 fontes → data mais votada)
 * ======================================================================================
 *
 * Fontes (6):
 * 1. Regex OCR (extractDataDocumentoTaloes)
 * 2. IA Mistral (Small, grátis)
 * 3. IA Groq (Llama 4 Scout, grátis)
 * 4. IA Gemini 2.0 Flash (grátis)
 * 5. IA Gemini 3.1 Flash Lite Preview (grátis)
 * 6. Nome do ficheiro (parse FONTE_YYYY-MM_ID.pdf → só MM/YYYY, voto parcial)
 *
 * O consenso escolhe a data com mais votos (maioria simples).
 * Para o nome do ficheiro (só MM/YYYY), conta como voto se mês e ano coincidem.
 */
function _extrairDataDoNomeFicheiro(fileName) {
  // Formato: FONTE_YYYY-MM_ID.pdf → extraímos YYYY e MM
  var m = fileName.match(/[_](\d{4})-(\d{2})[_]/);
  if (m) return { month: m[2], year: m[1] }; // Sem dia — voto parcial
  return null;
}

function _normalizarDataIA(resultado) {
  if (!resultado) return null;
  var limpo = resultado.trim().replace(/```/g, "").replace(/["""]/g, "").trim();
  var m = limpo.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) {
    var safe = _safeDate_(m[1], m[2], m[3]);
    return safe; // DD/MM/YYYY ou null
  }
  return null;
}

function _extrairDataViaModelo(texto, nomeFuncao, funcaoChamar) {
  var prompt = "Extraia apenas a data de emissão do seguinte texto de um documento.\n" +
    "Retorne SOMENTE a data no formato DD/MM/AAAA, sem mais nada.\n" +
    'Se não encontrar, retorne "Não encontrada".\n\nTexto:\n' + texto;
  try {
    var resultado = funcaoChamar(prompt);
    var data = _normalizarDataIA(resultado);
    Logger.log("  [CONSENSO] " + nomeFuncao + ": " + (data || "sem data"));
    return data;
  } catch (e) {
    Logger.log("  [CONSENSO] " + nomeFuncao + " ERRO: " + String(e).substring(0, 80));
    return null;
  }
}

function _consensoData(fileName, textoPDF) {
  Logger.log("🗳️ [CONSENSO] A iniciar para: " + fileName);
  var textoParaIA = textoPDF.substring(0, 4000);
  var votos = {}; // { "DD/MM/YYYY": count }
  var fontes = {}; // { "DD/MM/YYYY": ["fonte1", "fonte2"] }

  function registarVoto(data, fonte) {
    if (!data) return;
    if (!votos[data]) { votos[data] = 0; fontes[data] = []; }
    votos[data]++;
    fontes[data].push(fonte);
  }

  // --- FONTE 1: Regex OCR (rápido, sem API) ---
  var dataRegex = extractDataDocumentoTaloes(textoPDF);
  registarVoto(dataRegex, "Regex");

  // --- FONTE 2: IA Mistral ---
  var dataMistral = _extrairDataViaModelo(textoParaIA, "Mistral", chamarMistral);
  registarVoto(dataMistral, "Mistral");

  // --- FONTE 3: IA Groq ---
  var dataGroq = _extrairDataViaModelo(textoParaIA, "Groq", chamarGroq);
  registarVoto(dataGroq, "Groq");

  // --- FONTE 4: IA Gemini 2.0 Flash ---
  var dataGeminiPro = _extrairDataViaModelo(textoParaIA, "Gemini 2.0 Flash", function(p) { return chamarGemini(p, "gemini-2.0-flash"); });
  registarVoto(dataGeminiPro, "Gemini 2.0 Flash");

  // --- FONTE 5: IA Gemini 3.1 Flash Lite ---
  var dataGeminiLite = _extrairDataViaModelo(textoParaIA, "Gemini 3.1 Flash Lite", function(p) { return chamarGemini(p, "gemini-3.1-flash-lite-preview"); });
  registarVoto(dataGeminiLite, "Gemini Lite");

  // --- FONTE 6: Nome do ficheiro (voto parcial — só MM/YYYY) ---
  var nomeInfo = _extrairDataDoNomeFicheiro(fileName);
  if (nomeInfo) {
    // Procura entre as datas já votadas qual tem o mês/ano que bate
    var votouPorNome = false;
    for (var d in votos) {
      var parts = d.split("/");
      if (parts[1] === nomeInfo.month && parts[2] === nomeInfo.year) {
        registarVoto(d, "Nome");
        votouPorNome = true;
        break;
      }
    }
    if (!votouPorNome) {
      Logger.log("  [CONSENSO] Nome: " + nomeInfo.month + "/" + nomeInfo.year + " (sem match de dia)");
    }
  } else {
    Logger.log("  [CONSENSO] Nome: formato não reconhecido");
  }

  // --- APURAMENTO: data com mais votos ---
  var melhorData = null;
  var melhorVotos = 0;
  for (var d in votos) {
    if (votos[d] > melhorVotos) {
      melhorVotos = votos[d];
      melhorData = d;
    }
  }

  if (melhorData) {
    Logger.log("🗳️ [CONSENSO] RESULTADO: " + melhorData + " (" + melhorVotos + "/6 votos: " + fontes[melhorData].join(", ") + ")");
  } else {
    Logger.log("🗳️ [CONSENSO] RESULTADO: SEM DATA (nenhuma fonte devolveu data válida)");
  }

  return {
    data: melhorData,
    votos: melhorVotos,
    fontes: melhorData ? fontes[melhorData] : [],
    todas: votos
  };
}

/**
 * Classificação completa de documento via IA (tipo + data + fornecedor + NIF + valor).
 * Uma única chamada substitui toda a cadeia de if/else.
 * Retorna objeto { tipo, data, fornecedor, nif, valor } ou null se falhar.
 */
function classificarDocumentoViaIA(texto) {
  const prompt = `Classifica este documento PDF. Responde SÓ com JSON válido (sem markdown, sem \`\`\`):
{"tipo":"...","data":"DD/MM/AAAA","fornecedor":"...","nif":"...","valor":"..."}

Tipos possíveis (escolhe UM):
- fatura = fatura, factura, invoice, fatura-recibo, factura/recibo (FR), fatura simplificada, comissões bancárias, débitos de serviços. ATENÇÃO: "Válido como recibo após boa cobrança" NÃO torna o documento num recibo — continua a ser fatura.
- nota_credito = nota de crédito, credit note
- recibo = recibo de pagamento, receipt, quitação. Documento que CONFIRMA o pagamento de uma fatura já emitida. NÃO é recibo de vencimento/salário.
- recibo_vencimento = recibo de salário, payslip, vencimento
- extrato = extrato bancário, conta-corrente, extracto
- comprovativo = comprovativo de transferência/pagamento bancário
- lixo = avisos, notificações, publicidade, documentos sem valor fiscal
- desconhecido = impossível determinar

"data" = data de EMISSÃO (não vencimento). Formato DD/MM/AAAA.
Campos desconhecidos = null.

Texto:
${texto}`;

  var resposta, iaUsada;
  try {
    resposta = chamarMistral(prompt);
    iaUsada = "Mistral (Small)";
  } catch (eMistral) {
    Logger.log("Mistral classificação falhou: " + String(eMistral).substring(0, 80) + " → Groq...");
    resposta = chamarGroq(prompt);
    iaUsada = "Groq (Llama 4 Scout)";
  }
  Logger.log("IA classificação utilizada: " + iaUsada);

  // Limpar possível wrapping markdown
  resposta = resposta.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

  var obj = JSON.parse(resposta);
  if (!obj || !obj.tipo) return null;
  return obj;
}

/**
 * Extrai a menor data do texto (usada para corrigir faturas Radius).
 * Retorna string DD/MM/AAAA ou null.
 */
function _corrigirDataRadius(textoPDF, t) {
  if (!t.includes("509001319") && !t.includes("radius")) return null;

  Logger.log("RADIUS DETETADA: A calcular a menor data (Data Fatura).");
  var regexDatas = /(\d{2})[-./](\d{2})[-./](\d{4})/g;
  var todasDatas = [];
  var match;

  while ((match = regexDatas.exec(textoPDF)) !== null) {
    var d = parseInt(match[1], 10);
    var m = parseInt(match[2], 10);
    var y = parseInt(match[3], 10);
    if (y >= 2020 && y <= 2030 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      var dateObj = new Date(y, m - 1, d);
      todasDatas.push({
        str: match[0].replace(/-/g, "/").replace(/\./g, "/"),
        obj: dateObj.getTime()
      });
    }
  }

  if (todasDatas.length > 0) {
    todasDatas.sort(function(a, b) { return a.obj - b.obj; });
    return todasDatas[0].str;
  }
  return null;
}

/**
 * Calcula MD5 hash de um ficheiro do Drive.
 */
function _computeFileHash(file) {
  var bytes = file.getBlob().getBytes();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes);
  return digest.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

/**
 * Verifica se um ficheiro já existe (duplicado) nas pastas indicadas.
 * Lista ficheiros via Drive API v2 (paginado), filtra por tamanho em código,
 * e só calcula hash MD5 se o tamanho coincidir.
 */
function _verificarDuplicados(file, pastasParaVerificar) {
  var tamanhoNovo = file.getSize();
  var hashNovo = null; // Lazy: só calcula se encontrar ficheiro com mesmo tamanho

  for (var p = 0; p < pastasParaVerificar.length; p++) {
    var pasta = pastasParaVerificar[p];
    if (!pasta) continue;
    var folderId = pasta.getId();

    var q = "'" + folderId + "' in parents and trashed=false";
    var pageToken = null;

    do {
      var resp = Drive.Files.list({
        q: q,
        pageToken: pageToken,
        maxResults: 200,
        fields: "nextPageToken,items(id,title,fileSize)"
      });

      var items = resp.items || [];
      for (var i = 0; i < items.length; i++) {
        // Filtrar por tamanho em código (Drive API v2 não suporta fileSize na query)
        if (Number(items[i].fileSize) !== tamanhoNovo) continue;
        // Mesmo tamanho → calcular hash para confirmar
        if (!hashNovo) hashNovo = _computeFileHash(file);
        var existente = DriveApp.getFileById(items[i].id);
        if (_computeFileHash(existente) === hashNovo) {
          Logger.log("DUPLICADO: " + file.getName() + " = " + items[i].title + " (em " + pasta.getName() + ")");
          return true;
        }
      }

      pageToken = resp.nextPageToken || null;
    } while (pageToken);
  }
  return false;
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

// Encontra a pasta do ano (com cache em memória)
function getPastaAno_(year) {
  var key = year.toString();
  if (key in __cachePastaAno) return __cachePastaAno[key];
  var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();
  while (pastasFaturas.hasNext()) {
    var p = pastasFaturas.next();
    if (p.getName() === key) {
      __cachePastaAno[key] = p;
      return p;
    }
  }
  __cachePastaAno[key] = null;
  return null;
}

// Encontra a pasta do mês dentro da pasta do ano (com cache em memória)
function getPastaMes_(pastaAno, month, year) {
  var nomeAlvo = "Faturas_" + CODIGO_EMPRESA + "_" + month + "/" + year;
  if (nomeAlvo in __cachePastaMes) return __cachePastaMes[nomeAlvo];
  var pastasMes = pastaAno.getFolders();
  while (pastasMes.hasNext()) {
    var p = pastasMes.next();
    if (p.getName() === nomeAlvo) {
      __cachePastaMes[nomeAlvo] = p;
      return p;
    }
  }
  __cachePastaMes[nomeAlvo] = null;
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

  // Verificar duplicados: sub-pasta destino + irmã (#1 ↔ #2) e respetivos PARA CATALOGAR
  var pastasParaVerificar = [pastaFinal, pastaNivel1];
  var irma = (nomeSubPasta === "#1 - Faturas e NCs normais") ? "#2 - Faturas e NCs com reembolso"
           : (nomeSubPasta === "#2 - Faturas e NCs com reembolso") ? "#1 - Faturas e NCs normais"
           : null;
  if (irma) {
    var itIrma = pastaMes.getFolders();
    while (itIrma.hasNext()) {
      var pIrma = itIrma.next();
      if (pIrma.getName() === irma) {
        pastasParaVerificar.push(pIrma);
        var itPC = pIrma.getFolders();
        while (itPC.hasNext()) {
          var pPC = itPC.next();
          if (pPC.getName() === "PARA CATALOGAR") { pastasParaVerificar.push(pPC); break; }
        }
        break;
      }
    }
  }
  if (_verificarDuplicados(file, pastasParaVerificar)) {
    file.moveTo(DriveApp.getFolderById(PASTA_LIXO));
    return { sucesso: false, erro: "Duplicado (ficheiro idêntico já existe em " + pastaMes.getName() + ")." };
  }

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
  
  // Funções Auxiliares: InjectSpaces e limpeza de linhas, permanecem as mesmas
  const injectSpaces = s => String(s).replace(/\u00A0/g, ' ').replace(/([A-Za-z])(\d)/g, '$1 $2').replace(/(\d)([A-Za-z])/g, '$1 $2');
  const raw = injectSpaces(pdfText);
  const linesAll = raw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  
  // (Variáveis de filtragem permanecem as mesmas)
  const badPOS = [/\b(pos|tpa|tp[ãa]g|terminal|redeunic|multibanco|mb\s*way|sibs)\b/i, /\bvisa\b/i, /\bmastercard\b/i, /\bmaestro\b/i, /\bam[ex|erican\s*express]\b/i, /\bautoriz[aç][aã]o\b/i, /\bauth(?:orization)?\b/i, /\baid\b/i, /\batc\b/i, /\btid\b/i, /\bnsu\b/i, /\bpan\b/i, /\barqc?\b/i, /\bcomprovativo\b/i, /\breceb[ií]do\b/i, /\bmerchant\s*copy\b/i, /\bclient\s*copy\b/i, /\blote\b/i, /\bref(?:\.|er[eê]ncia)?\b/i, /\btransa[cç][aã]o\b/i, /\bvenda\b/i, /\bpagamento\b/i];
  const timeRe = /\b\d{2}:\d{2}(?::\d{2})?\b/; 
  const lines = linesAll.filter(l => !isHardBad(l));
  
  // RegEx de cabeçalho
  const headerRe = /\b(?:fatura\/recibo|fatura|factura|nota\s+de\s+cr[eé]dito)\b/i;
  
  // RegEx ORIGINAL (DD/MM/AAAA)
  const dataFieldRe = /\bdata\s*:\s*(\d{2})[./-](\d{2})[./-](\d{4})\b/i;
  
  // NOVO REGEX: Captura o formato ISO AAAA-MM-DD perto do label "Data:" (CORREÇÃO CHAVE)
  const dataISOFieldRe = /\bdata\s*:\s*(20\d{2})[./-](\d{2})[./-](\d{2})\b/i; 

  // --- 1. FAST EXIT (PERTO DO CABEÇALHO) ---
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (headerRe.test(l)) {
      const look = [l, lines[i+1], lines[i+2], lines[i+3]].filter(Boolean);
      for (const seg of look) {
        
        // Tenta DD/MM/AAAA primeiro
        let m = seg.match(dataFieldRe);
        if (m) {
          const best = _safeDate_(m[1], m[2], m[3]);
          if (best) return best;
        }

        // Tenta AAAA-MM-DD (CORREÇÃO)
        m = seg.match(dataISOFieldRe);
        if (m) {
          // m[1]=Ano, m[2]=Mês, m[3]=Dia -> Inverter para (Dia, Mês, Ano)
          const best = _safeDate_(m[3], m[2], m[1]); 
          if (best) return best;
        }
      }
    }
  }

  // --- 2. HEURÍSTICA DE PONTUAÇÃO (Restante da Lógica) ---
  const goodLabels = [
    /data\s*(?:de)?\s*emiss[aã]o/i, 
    /\bemiss[aã]o\b/i, 
    /\bemitida\b/i, 
    /\bemitida\s+em\b/i,
    /dt\.?\s*emiss[aã]o/i,
    /\bdata\s*doc(?:umento)?\b/i,
    /\bdata\s*:\b/i, // <-- Este label é capturado aqui para o ISO Near Label
    /\bdata\s*da\s*fatura\b/i,
    /\binvoice\s+date\b/i, 
    /\bissue\s+date\b/i, 
    /\bfecha\s+de\s+emisi[oó]n\b/i
  ];
  const badLabels = [/\bvig[êe]ncia\b/i, /\bper[ií]odo\b/i, /\bvalidade\b/i, /\bcompet[êe]ncia\b/i, /\bref\.\s*(?:per[ií]odo|m[êe]s)\b/i, /\bintervalo\b/i, /\bvenc(?:imento)?\b/i, /\bprazo\b/i, /\bdue\b/i, /\bpayment\b/i];
  const hasRangeRe = /\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*(?:a|–|—|-)\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/i;

  const oneLine = raw.replace(/\s+/g, ' ');
  
  // ISO NEAR LABEL (Continua a ser útil para ISO sem o 'Data:' explícito no Fast Exit)
  const isoNearLabel = new RegExp('(?:' + goodLabels.map(r=>r.source).join('|') + ')' + '[^0-9]{0,40}(20\\d{2})[./-](\\d{2})[./-](\\d{2})','i');
  let m = oneLine.match(isoNearLabel);
  if (m) {
    // m[1]=Y, m[2]=M, m[3]=D
    const fast = _safeDate_(m[3], m[2], m[1]);
    if (fast) return fast;
  }

  // Regras de Expressões Regulares de data (rx)
  const rx = [
    // [0] AAAA/MM/DD ou AAAA-MM-DD. O norm faz a inversão (d:m[3], m:m[2], y:m[1])
    { re:/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/g, norm:(y,m,d)=>({d,m,y, iso:true}) },
    // [1] DD/MM/AAAA
    { re:/\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/g,     norm:(d,m,y)=>({d,m,y}) },
    // [2] DD/MM/AA
    { re:/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\b/g,         norm:(d,m,y)=>({d,m,y, ambiguousYY:true}) },
    // (As restantes regras de extenso e abreviaturas permanecem iguais)
    { re:/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})\b/gi, norm:(d,mon,y)=>({d, m:_monPT_(mon), y}) },
    { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})\b/gi, norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
    { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2})\s*,?\s*(\d{4})\b/gi, norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
    { re:/\b(\d{1,2})[.\- ](jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[.\- ](\d{4})\b/gi, norm:(d,mon,y)=>({d, m:_monPTabbrev_(mon), y}) }
  ];

  const candidates = [];
  // (A lógica de iteração e pontuação dos candidatos permanece a mesma, usando as regras 'rx' acima)
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
        // ... (resto da lógica de normalização, verificação de ano, e pontuação) ...
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
  
  // (Lógica de desempate final)
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

      var dataPagamentoStr = extractDateFromPayslip(texto);
      var anoPagamento = inferYearFromDateString(dataPagamentoStr) || new Date().getFullYear();

      // Tentar match por ATCUD primeiro, fallback para valor+entidade
      var match = null;
      if (atcud) {
        match = procurarFaturaPorATCUDNoArquivo(atcud, anoPagamento);
        if (match) Logger.log("✅ Comprovativo match por ATCUD: " + atcud);
      }
      if (!match) {
        Logger.log("⚡ Sem ATCUD ou sem match ATCUD para " + file.getName() + " → fallback valor+entidade");
        Logger.log("   Texto (500 chars): " + (texto || "").substring(0, 500).replace(/\n/g, " "));
        // Extrair todos os valores monetários do texto para debug
        var valoresNoTexto = (texto || "").match(/\d[\d\s.,]*\d/g) || [];
        Logger.log("   Valores no texto: " + valoresNoTexto.slice(0, 15).join(" | "));
        match = procurarFaturaPorValorEntidade_(texto, anoPagamento);
        if (match) Logger.log("✅ Comprovativo match por VALOR+ENTIDADE: " + match.numeroDocumento);
      }

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

/**
 * Fallback para comprovativos sem ATCUD: procura fatura por valor + entidade.
 * Percorre Sheets dos últimos 6 meses.
 */
function procurarFaturaPorValorEntidade_(textoPdf, anoPagamento) {
  if (!textoPdf) return null;
  var textoLower = textoPdf.toLowerCase();

  var yearWrappers = getYearFolders();
  if (!yearWrappers || !yearWrappers.length) return null;

  // Percorrer últimos 6 meses
  var agora = new Date();
  var mesAtual = agora.getMonth() + 1;
  var anoAtual = anoPagamento || agora.getFullYear();

  // Extrair o descritivo/complementar do comprovativo (frequentemente contém nº fatura)
  var descritivoMatch = textoPdf.match(/(?:Descritivo|Complementar|Informação\s*Complementar)[:\s]*([^\n]{2,40})/i);
  var descritivo = descritivoMatch ? descritivoMatch[1].trim() : "";
  if (descritivo) Logger.log("     Descritivo/Complementar: " + descritivo);

  for (var offset = 0; offset < 12; offset++) {
    var d = new Date(anoAtual, mesAtual - 1 - offset, 1);
    var checkYear = d.getFullYear();
    var checkMonth = ("0" + (d.getMonth() + 1)).slice(-2);

    var pAno = getPastaAno_(checkYear);
    if (!pAno) continue;
    var pMes = getPastaMes_(pAno, checkMonth, checkYear);
    if (!pMes) continue;

    var nomeFicheiroExcel = "#0 - Faturas_" + CODIGO_EMPRESA + "_" + checkMonth + "/" + checkYear;
    var files = pMes.getFilesByName(nomeFicheiroExcel);
    if (!files.hasNext()) continue;

    try {
      var fileExcel = files.next();
      var ss = SpreadsheetApp.open(fileExcel);
      var abas = ["Faturas e NCs normais", "Faturas e NCs com reembolso", "Outros documentos"];

      for (var k = 0; k < abas.length; k++) {
        var sheet = ss.getSheetByName(abas[k]);
        if (!sheet) continue;
        var lastRow = sheet.getLastRow();
        if (lastRow <= 2) continue;

        var colEntidade = encontraColunaNoCabecalho(sheet, "Fornecedor", 2);
        var colValor = encontraColunaNoCabecalho(sheet, "Valor total", 2);
        var colNumDoc = encontraColunaNoCabecalho(sheet, "Nº", 2);
        if (colNumDoc < 0) colNumDoc = encontraColunaNoCabecalho(sheet, "Número do documento", 2);
        var colComp = encontraColunaNoCabecalho(sheet, "Comprovativo de pagamento", 2);

        if (colValor < 0 || colEntidade < 0) continue;

        var dados = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).getValues();

        for (var r = 0; r < dados.length; r++) {
          var rowData = dados[r];

          // Verificar valor (formato flexível)
          var valorExcel = rowData[colValor - 1];
          var valorFloat = null;
          if (typeof valorExcel === 'number') {
            valorFloat = valorExcel;
          } else if (typeof valorExcel === 'string') {
            var limpo = valorExcel.replace(/[^0-9.,-]/g, "").replace(",", ".");
            valorFloat = parseFloat(limpo);
          }

          var valorMatch = false;
          if (valorFloat !== null && !isNaN(valorFloat) && valorFloat > 0) {
            var vRaw = valorFloat.toFixed(2);
            var vVirgula = vRaw.replace(".", ",");
            var partes = vVirgula.split(",");
            var vComEspaco = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ") + "," + partes[1];
            var vComPonto = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".") + "," + partes[1];

            if (textoPdf.includes(vRaw) || textoPdf.includes(vVirgula) ||
                textoPdf.includes(vComEspaco) || textoPdf.includes(vComPonto)) {
              valorMatch = true;
            }
          }

          if (!valorMatch) continue;

          // Verificar entidade (tudo em minúsculas)
          var entidadeExcel = String(rowData[colEntidade - 1] || "").toLowerCase().trim();
          if (entidadeExcel.length <= 2) continue;

          var primeiraPalavra = entidadeExcel.split(/\s+/)[0];
          if (primeiraPalavra.length <= 2) {
            // Se primeira palavra é muito curta (ex: "A"), usar as primeiras duas
            var palavras = entidadeExcel.split(/\s+/);
            primeiraPalavra = palavras.slice(0, 2).join(" ");
          }

          // Comparar sem acentos
          var textoSemAcentos = textoLower.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
          var entidadeSemAcentos = primeiraPalavra.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

          var entidadeEncontrada = textoLower.includes(primeiraPalavra) || textoSemAcentos.includes(entidadeSemAcentos);

          // Se entidade não bate, tentar match por descritivo + nº documento
          var descritivoMatch2 = false;
          if (!entidadeEncontrada && descritivo && colNumDoc > 0) {
            var numDocSheet = String(rowData[colNumDoc - 1]).trim();
            if (numDocSheet.length >= 2 && descritivo.includes(numDocSheet)) {
              descritivoMatch2 = true;
              Logger.log("     📎 Descritivo '" + descritivo + "' contém nº doc '" + numDocSheet + "'");
            }
          }

          if (!entidadeEncontrada && !descritivoMatch2) {
            Logger.log("     ⚠️ Valor " + valorFloat + " bateu, mas entidade '" + primeiraPalavra + "' não encontrada no texto");
            continue;
          }

          // Match confirmado!
          var numDoc = colNumDoc > 0 ? String(rowData[colNumDoc - 1]) : "?";
          Logger.log("   -> Match: " + checkMonth + "/" + checkYear + " | Entidade: " + entidadeExcel + " | Valor: " + valorFloat + " | Doc: " + numDoc);

          return {
            ano: checkYear,
            spreadsheetId: fileExcel.getId(),
            spreadsheetName: fileExcel.getName(),
            pastaMesNome: pMes.getName(),
            numeroDocumento: numDoc,
            colComp: colComp,
            row: r + 3, // 1-indexed, data starts at row 3
            aba: abas[k],
          };
        }
      }
    } catch (e) {
      Logger.log("❌ Erro ao ler excel " + checkMonth + "/" + checkYear + " (valor+entidade): " + e.message);
    }
  }

  Logger.log("🏁 Sem match valor+entidade nos últimos 6 meses.");
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

        // === NÍVEL 0: Extrair nºs de fatura do texto do recibo e procurar na coluna ATCUD ===
        // Recibos frequentemente listam as faturas que pagam (ex: "Fatura 1 2600/000077")
        // Extrair os sequenciais dessas referências e procurar na coluna ATCUD
        if (colATCUD > 0) {
          var seqsRecibo = new Set();

          // Padrão 1: "XXXX/NNNNNN" (série/número) — extrair só a parte após a barra
          // Exclui anos isolados (2024/, 2025/, 2026/) que não são nºs de fatura
          var refsBarras = textoPdf.match(/\b\d{1,4}\/0*(\d{2,})\b/g) || [];
          for (var q = 0; q < refsBarras.length; q++) {
            var partesBarra = refsBarras[q].split("/");
            var serie = partesBarra[0];
            var seqRef = partesBarra[1].replace(/^0+/, "") || "0";
            // Excluir se a série é um ano (2020-2030) e o sequencial parece data (mês/dia)
            var serieNum = parseInt(serie);
            if (serieNum >= 2020 && serieNum <= 2030) continue;
            if (seqRef.length >= 2) seqsRecibo.add(seqRef);
          }

          // Padrão 2: Tabela de documentos no recibo — "Fatura" ou "Nota de Crédito" seguido de nº
          // Captura: "Fatura1 2600/000077" → "000077" → "77"
          var refsTabela = textoPdf.match(/(?:Fatura|Factura|Nota de [Cc]rédito)\s*\d*\s+\d{1,4}\/0*(\d{2,})/g) || [];
          for (var q2 = 0; q2 < refsTabela.length; q2++) {
            var numTabela = refsTabela[q2].match(/\/0*(\d{2,})/);
            if (numTabela) seqsRecibo.add(numTabela[1].replace(/^0+/, "") || "0");
          }

          // Padrão 3: "Nº Documento" seguido de nº com barra (contexto de tabela de recibo)
          var refsNDoc = textoPdf.match(/N[º°]\s*Documento[^\n]{0,30}?\d{1,4}\/0*(\d{2,})/gi) || [];
          for (var q3 = 0; q3 < refsNDoc.length; q3++) {
            var numNDoc = refsNDoc[q3].match(/\/0*(\d{2,})/);
            if (numNDoc) seqsRecibo.add(numNDoc[1].replace(/^0+/, "") || "0");
          }

          if (seqsRecibo.size > 0) {
            Logger.log("     [NÍVEL 0] Referências a faturas encontradas no recibo: " + Array.from(seqsRecibo).join(", "));

            for (var r0 = 0; r0 < dados.length; r0++) {
              var atcudExcel = String(dados[r0][colATCUD-1]).toUpperCase().trim();
              if (!atcudExcel || atcudExcel.length < 3) continue;

              // Extrair sequencial do ATCUD (parte após último "-" ou últimos dígitos)
              var atcudSeq = atcudExcel.replace(/.*[-]0*/, "");

              var encontrouRef = false;
              var seqArr = Array.from(seqsRecibo);
              for (var sq = 0; sq < seqArr.length; sq++) {
                if (atcudSeq === seqArr[sq] || atcudExcel.includes(seqArr[sq])) {
                  encontrouRef = true;
                  break;
                }
              }

              if (encontrouRef) {
                // Confirmar com valor
                var valExcel0 = dados[r0][colValor-1];
                var valFloat0 = typeof valExcel0 === 'number' ? valExcel0 : parseFloat(String(valExcel0).replace(/[^0-9.,-]/g, "").replace(",", "."));

                if (valFloat0 && !isNaN(valFloat0)) {
                  var v0 = valFloat0.toFixed(2);
                  var v0v = v0.replace(".", ",");
                  if (textoPdf.includes(v0) || textoPdf.includes(v0v)) {
                    Logger.log("✅ SMART MATCH NÍVEL 0 (Ref. fatura no recibo + ATCUD + Valor)");
                    Logger.log("   -> Destino: " + checkMonth + "/" + checkYear);
                    Logger.log("   -> ATCUD: " + atcudExcel + " | Valor: " + valExcel0);
                    return { year: String(checkYear), month: String(checkMonth) };
                  }
                }
              }
            }
          }
        }

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

          // === PASSO 3: VERIFICAR ID (3 níveis de confiança) ===
          var idEncontrado = "";
          var nivelMatch = 0; // 0=nenhum, 1=nº completo, 2=ATCUD completo, 3=sequencial longo+valor+entidade

          // NÍVEL 1: Número COMPLETO da fatura no texto (ex: "FT 2026/13", "FT A/13")
          if (colNum > 0) {
             var numDoc = String(rowData[colNum-1]).toUpperCase().trim();
             if (numDoc.length >= 3) {
               // Normalizar espaços para match mais flexível
               var numDocNorm = numDoc.replace(/\s+/g, " ");
               var numDocSemEspacos = numDoc.replace(/\s+/g, "");
               if (textoNorm.includes(numDocNorm) || textoNorm.includes(numDocSemEspacos)) {
                 nivelMatch = 1;
                 idEncontrado = "Nº COMPLETO " + numDoc;
               }
             }
          }

          // NÍVEL 2: ATCUD COMPLETO no texto (sempre único)
          if (nivelMatch === 0 && colATCUD > 0) {
             var atcud = String(rowData[colATCUD-1]).toUpperCase().trim();
             if (atcud.length >= 5) {
               // ATCUD tem formato tipo "ABCD1234-5678" — procurar completo
               var atcudNorm = atcud.replace(/\s+/g, "");
               if (textoNorm.includes(atcud) || textoNorm.replace(/\s+/g, "").includes(atcudNorm)) {
                 nivelMatch = 2;
                 idEncontrado = "ATCUD COMPLETO " + atcud;
               }
             }
          }

          // NÍVEL 3 (fallback): Sequencial LONGO (6+ dígitos) + valor + entidade (os 3 juntos)
          if (nivelMatch === 0) {
             var seqMatch = false;
             var seqEncontrado = "";

             if (colNum > 0) {
               var numDoc3 = String(rowData[colNum-1]).toUpperCase().trim();
               var sequencial = extractSequencial_(numDoc3);
               if (sequencial && sequencial.length >= 6 && textoNorm.includes(sequencial)) {
                 seqMatch = true;
                 seqEncontrado = "Nº " + numDoc3 + " (Seq: " + sequencial + ")";
               }
             }
             if (!seqMatch && colATCUD > 0) {
               var atcud3 = String(rowData[colATCUD-1]).toUpperCase().trim();
               var seqAtcud = extractSequencial_(atcud3);
               if (seqAtcud && seqAtcud.length >= 6 && textoNorm.includes(seqAtcud)) {
                 seqMatch = true;
                 seqEncontrado = "ATCUD " + atcud3 + " (Seq: " + seqAtcud + ")";
               }
             }

             // Nível 3 exige os 3: sequencial + valor + entidade
             if (seqMatch && valorMatch && entidadeMatch) {
               nivelMatch = 3;
               idEncontrado = seqEncontrado + " + Valor + Entidade";
             }
          }

          // === DECISÃO FINAL ===
          // Nível 1 ou 2: match seguro (número ou ATCUD completo) — valor confirma
          // Nível 3: fallback (sequencial longo + valor + entidade juntos)
          if (nivelMatch === 1 && valorMatch) {
            Logger.log("✅ SMART MATCH NÍVEL 1 (Nº completo + Valor)");
            Logger.log("   -> Destino: " + checkMonth + "/" + checkYear);
            Logger.log("   -> Valor: " + rowData[colValor-1] + " | ID: " + idEncontrado);
            return { year: String(checkYear), month: String(checkMonth) };
          } else if (nivelMatch === 2 && valorMatch) {
            Logger.log("✅ SMART MATCH NÍVEL 2 (ATCUD completo + Valor)");
            Logger.log("   -> Destino: " + checkMonth + "/" + checkYear);
            Logger.log("   -> Valor: " + rowData[colValor-1] + " | ID: " + idEncontrado);
            return { year: String(checkYear), month: String(checkMonth) };
          } else if (nivelMatch === 3) {
            Logger.log("✅ SMART MATCH NÍVEL 3 (Seq longo + Valor + Entidade)");
            Logger.log("   -> Destino: " + checkMonth + "/" + checkYear);
            Logger.log("   -> " + idEncontrado);
            return { year: String(checkYear), month: String(checkMonth) };
          } else if (valorMatch) {
            Logger.log("      ⚠️ Valor bateu, mas ID insuficiente. (Excel: " + (colNum>0?rowData[colNum-1]:"N/A") + ", Entidade: " + (entidadeMatch?"SIM":"NÃO") + ")");
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