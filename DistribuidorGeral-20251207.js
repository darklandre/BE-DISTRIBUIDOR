// Pasta onde ficam guardados os anexos que vêm do email [ficheiros por processar] (assegurado pelo N8N, que envia os anexos em pdf do email documentos@darkland.pt para a pasta)
const PASTA_GERAL_FICHEIROS = "1DKCSluenYGGNz05uLLzQwWwq-wYHCk54";
// Sub-pasta, "Comprovativos", onde são colocados os comprovativos depois de fazer os pagamentos
var PASTA_COMPROVATIVOS_ID = "1nBnKXyvtiUt7BMYdaQfYtaJbgBMe4xjR";

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









// COPIAR PARA OUTRAS EMPRESAS A PARTIR DAQUI:

// Cache em memória das pastas de anos
var __cacheYearFolders = null;

function distribuirFicheirosDoGeral() {

  var sourceFolder = DriveApp.getFolderById(PASTA_GERAL_FICHEIROS);
  var pastaGeralRecibos = DriveApp.getFolderById(PASTA_GERAL_RECIBOS);
  var pastaLixo = DriveApp.getFolderById(PASTA_LIXO);

  var recCount = 0; // Número de recibos de vencimento movidos
  var fileCount = 0; // Número de ficheiros movidos
  var fileErrors = 0; // Número de tentativas de ficheiros
  var nomesFicheirosMovidos = '';
  var errosFicheirosMovidos = '';

  var files = sourceFolder.getFiles();
  
  while (files.hasNext()) {

    var file = files.next();
    var fileName = file.getName();

    let textoPDF = "";
    try {
      textoPDF = convertPDFToText(file.getId(), 'pt') || "";
    } catch (e) {
      fileErrors++;
      errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ': Falha no OCR (' + e + ').\n';
      continue;
    }
    if (!textoPDF.trim()) { // nada reconhecido
      fileErrors++;
      errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ': PDF sem texto após OCR.\n';
      continue;
    }

    // EXTRAÇÕES BASE (precisamos já disto para todos os ramos)
    var tipoDocumento = extractTipoDocumento(textoPDF);
    var valorATCUD = extractATCUD(textoPDF);
    //var dataDocumento = extractDataDocumento(textoPDF);
    var dataDocumento = extractDataDocumentoTaloes(textoPDF);
    var existePastaMesAnoFaturas = 0;

    // Para todas as empresas:
  
    const t = (textoPDF || "").toLowerCase();

    // CASO DE EXTRATOS DE CONTA CORRENTE (ECC)
     
    const ehExtrato =
      (
        t.includes("extracto de contas correntes") || t.includes("documentos de clientes por liquidar")
      )

    if (ehExtrato) {

    }    

    // CASO DE RECIBOS PT -- NÃO FATURAS
    
    const ehReciboPT = (
      (
        t.includes("recibo n.º") ||
        t.includes("recibo nº")  ||
        t.includes("recibo nr")  ||
        t.includes("recebemos a quantia de") ||
        t.includes("recebemos a importância") ||
        t.includes("este documento não serve de factura") ||
        t.includes("este documento não serve de fatura") ||
        t.includes("recibo cliente") ||
        t.includes("total do recibo")
      )
      &&
      !t.includes("fatura/recibo") &&
      !t.includes("fatura-recibo")
    );

    if (ehReciboPT) {
    
      Logger.log("CASO RECIBO");

      Logger.log("FILE: " + fileName);
      Logger.log("TEXTO: " + textoPDF);
      Logger.log("TIPO: " + tipoDocumento);
      Logger.log("ATCUD: " + valorATCUD);
      Logger.log("DATA: " + dataDocumento);

      if(!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Data inválida ou não encontrada.\n";
        continue;
      }

      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      Logger.log("MES: " + month);
      Logger.log("ANO: " + year);

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;
              Logger.log("ENTRY 1");

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#4 - Recibos") {

                  Logger.log("ENTRY 2");

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      var novo = copiarMoverELog_(file, pastaNivelDois, sourceFolder, null, null);
                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          } 

        }
      }
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;
    }

    // Faturas do Crédito Agricola (comissões, cartão refeição, etc)
    if(valorATCUD && textoPDF.includes("www.creditoagricola.pt") && (textoPDF.includes("FACTURA") || textoPDF.includes("FATURA"))){

      if (!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Caso válido mas data inválida ou não encontrada.\n";
        continue;
      }

      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      Logger.log("MES: " + month);
      Logger.log("ANO: " + year);

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;
              Logger.log("ENTRY 1");

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              //Encontra pasta "#1 - Faturas e NCs normais"
              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#1 - Faturas e NCs normais") {

                  Logger.log("ENTRY 2");

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          } 
        }
      }
      if(existePastaMesAnoFaturas)
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;
    }

    // 0 - Casos especiais de faturas (Digital Ocean) 
    //     REGRA: Não existe
    //     ACÇÃO: Extrai data e envia para a pasta das faturas daquele mês (onde existe um catalogador)

    if(CODIGO_EMPRESA==="DP"){
      if(textoPDF.includes("DigitalOcean")){

        var dateRegex5 = /(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s+(\d{4})/i;
        var match5 = textoPDF.match(dateRegex5);

        if (!match5) {
          fileErrors++;
          errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Ficheiro do DigitalOcean com erro na data (em EN) não encontrada.\n";
          continue;
        }

        if (match5) {
          var monthName = match5[1].toLowerCase();
          var day = match5[2];
          var year = match5[3];

          var monthMapEnglish = {
            "january": "01", "february": "02", "march": "03", "april": "04", "may": "05", "june": "06",
            "july": "07", "august": "08", "september": "09", "october": "10", "november": "11", "december": "12"
          };

          var month = monthMapEnglish[monthName];

        }

        var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

        while (pastasFaturas.hasNext()) {
          var pastaFaturas = pastasFaturas.next();
          if (pastaFaturas.getName() === year) {

            var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

            while (pastaFaturasDentroDoAnoCerto.hasNext()) {
              var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

              if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

                existePastaMesAnoFaturas = 1;
                Logger.log("ENTRY 1");

                var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

                //Encontra pasta "#1 - Faturas e NCs normais"
                while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                  var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                  if (pastaNivelUm.getName() === "#1 - Faturas e NCs normais") {

                    Logger.log("ENTRY 2");

                    var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                    //Encontra pasta "PARA CATALOGAR"
                    while (iteradorPastasCatalogacao.hasNext()) {

                      var pastaNivelDois = iteradorPastasCatalogacao.next();
                      if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                        Logger.log("ENTRY 3");

                        Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                        copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                        fileCount++;
                        nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                      }

                    }
                  }
                }
              }
            }
            
            if(!existePastaMesAnoFaturas){
              Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
              fileErrors++;
              errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
            } 

          }
        }
        continue; // Ao ser distribuido, não irá verificar mais nenhum caso;

      }
    }
    if(CODIGO_EMPRESA==="DL"){
      // Avisos Tranquilidade/Generali
      if(textoPDF.includes("O seu seguro vai ser pago por débito") && textoPDF.includes("(Este documento não serve de fatura)") && textoPDF.includes("Aviso") ){
        file.moveTo(pastaLixo);
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não é para catalogar (aviso seguro) ";
        continue; // Não irá verificar mais nenhum caso;
      }
      // Avisos alterações seguros Tranquilidade/Generali
      if(textoPDF.includes("Condições Particulares da") && !textoPDF.includes("fatura")){
        file.moveTo(pastaLixo);
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não é para catalogar (condições seguro) ";
        continue; // Não irá verificar mais nenhum caso;
      }        
      // Avisos alterações seguros Tranquilidade/Generali
      if(textoPDF.includes("Generali") && textoPDF.includes("Autorização de Débito Direto") ){
        file.moveTo(pastaLixo);
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não é para catalogar (autorização débito direto seguro) ";
        continue; // Não irá verificar mais nenhum caso;
      }
      // Avisos alterações seguros Tranquilidade/Generali
      if(textoPDF.includes("Generali") && textoPDF.includes("NOTA INFORMATIVA") ){
        file.moveTo(pastaLixo);
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não é para catalogar (nota seguro) ";
        continue; // Não irá verificar mais nenhum caso;
      }
      // Extratos conta corrente BigMat
      if(textoPDF.includes("EXTRACTO COBRANÇAS CLIENTE") ){
        file.moveTo(pastaLixo);
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não é para catalogar (extrato bigmat) ";
        continue; // Não irá verificar mais nenhum caso;
      }
    
    }
    // Para todas as empresas:
    // Extratos (Calculo Fiscal Dr. Barroso)
    if(textoPDF.includes("Abaixo se discriminam as faturas em divida, que solicitamos que sejam liquidadas.") ){
      file.moveTo(pastaLixo);
      fileErrors++;
      errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não é para catalogar (aviso seguro) ";
      continue; // Não irá verificar mais nenhum caso;
    }

    // 1 - Se for um recibo de vencimento
    //     REGRA: O nome do ficheiro começa com "REC"
    //     ACÇÃO: Envia para a pasta dos recibos (onde existe um catalogador para verificar onde pertence)

    if (fileName.startsWith('REC_') && 
        !textoPDF.includes("Fatura") &&
        !textoPDF.includes("Fatura simplificada") &&
        !textoPDF.includes("Fatura-recibo")
        ) {

      Logger.log("CASO 1!");

      copiarMoverELog_(file, pastaGeralRecibos, sourceFolder);
      recCount++;
      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + '\n';
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;
    } 

    // 2 - Se for uma fatura nacional
    //     REGRA: Tem ATCUD e não contém 'www.creditoagricola.pt' (porque os comprovativos têm ATCUDs e esse texto)
    //     ACÇÃO: Extrai data e envia para a pasta das faturas daquele mês (onde existe um catalogador)

    existePastaMesAnoFaturas = 0;

    if(valorATCUD && !textoPDF.includes("www.creditoagricola.pt")){

      Logger.log("CASO 2!");

      Logger.log("FILE: " + fileName);
      Logger.log("TEXTO: " + textoPDF);
      Logger.log("TIPO: " + tipoDocumento);
      Logger.log("ATCUD: " + valorATCUD);
      Logger.log("DATA: " + dataDocumento);

      if(!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Caso válido mas data inválida ou não encontrada.\n";
        continue;
      }
      
      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      Logger.log("MES: " + month);
      Logger.log("ANO: " + year);

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;
              Logger.log("ENTRY 1");

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              //Encontra pasta "#1 - Faturas e NCs normais"
              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#1 - Faturas e NCs normais") {

                  Logger.log("ENTRY 2");

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!"); 
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          }

        }

      }
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;  

    }

    // 3.1 - Se for uma fatura estrangeira em Inglês
    //     REGRA: Tem "invoice" ou MAIS...
    //     ACÇÃO: Extrai data e envia para a pasta das faturas daquele mês (onde existe um catalogador)

    existePastaMesAnoFaturas = 0;

    if(textoPDF.includes("invoice") || textoPDF.includes("Invoice") || textoPDF.includes("INVOICE"))
    if(!textoPDF.includes("receipt") && !textoPDF.includes("Receipt") && !textoPDF.includes("RECEIPT"))
    if(!textoPDF.includes("www.creditoagricola.pt")){

      Logger.log("CASO 3.1!");

      Logger.log("FILE: " + fileName);
      Logger.log("TEXTO: " + textoPDF);
      Logger.log("TIPO: " + tipoDocumento);
      Logger.log("ATCUD: " + valorATCUD);
      Logger.log("DATA: " + dataDocumento);

      if(!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Caso válido mas data inválida ou não encontrada.\n";
        continue;
      }

      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      Logger.log("MES: " + month);
      Logger.log("ANO: " + year);

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;
              Logger.log("ENTRY 1");

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              //Encontra pasta "#1 - Faturas e NCs normais"
              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#1 - Faturas e NCs normais") {

                  Logger.log("ENTRY 2");

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          } 
        }
      }
      if(existePastaMesAnoFaturas)
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;
    }

    // 3.2 - Se for uma fatura estrangeira sem atcud mas em português (e.g. Amazon, Adobe, etc)
    //     REGRA: Tem "fatura" ou "factura"
    //     ACÇÃO: Extrai data e envia para a pasta das faturas daquele mês (onde existe um catalogador)

    existePastaMesAnoFaturas = 0;

    if(!valorATCUD)
    if(textoPDF.includes("factura") || textoPDF.includes("Factura") || textoPDF.includes("FACTURA")
    || textoPDF.includes("fatura") || textoPDF.includes("Fatura") || textoPDF.includes("FATURA")){

      Logger.log("CASO 3.2");

      Logger.log("FILE: " + fileName);
      Logger.log("TEXTO: " + textoPDF);
      Logger.log("TIPO: " + tipoDocumento);
      Logger.log("ATCUD: " + valorATCUD);
      Logger.log("DATA: " + dataDocumento);

      if(!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Caso válido mas data inválida ou não encontrada.\n";
        continue;
      }

      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      Logger.log("MES: " + month);
      Logger.log("ANO: " + year);

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;
              Logger.log("ENTRY 1");

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              //Encontra pasta "#1 - Faturas e NCs normais"
              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#1 - Faturas e NCs normais") {

                  Logger.log("ENTRY 2");

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          } 

        }
      }
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;
    }

    // 4 - Se for um recibo
    //     REGRA: 
    //     ACÇÃO: Só move se a fatura já está catalogada (problema - pode ser de outro mês)
    // 4.1 (nacional)

    existePastaMesAnoFaturas = 0;

    // 4.2 (estrangeiro)
    if(!valorATCUD)
    if(textoPDF.includes("receipt") || textoPDF.includes("Receipt") || textoPDF.includes("RECEIPT")){

      Logger.log("CASO 4.2");

      Logger.log("FILE: " + fileName);
      Logger.log("TEXTO: " + textoPDF);
      Logger.log("TIPO: " + tipoDocumento);
      Logger.log("ATCUD: " + valorATCUD);
      Logger.log("DATA: " + dataDocumento);

      if (!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Caso válido mas data inválida ou não encontrada.\n";
        continue;
      }

      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      Logger.log("MES: " + month);
      Logger.log("ANO: " + year);

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;
              Logger.log("ENTRY 1");

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              //Encontra pasta "#1 - Faturas e NCs normais"
              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#4 - Recibos") {

                  Logger.log("ENTRY 2");

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          } 

        }
      }
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;
    }

    existePastaMesAnoFaturas = 0;

    // 5 - Se for um comprovativo de pagamento do CA
    //     REGRA: Contém 'www.creditoagricola.pt' e 'Conta a Debitar: 40294310603'
    //     ACÇÃO: Vai procurar em todas as faturas daquele mês se lá está o ATCUD (caso não encontre naquele mês, anda para trás no mês até chegar a Janeiro daquele ano)

    var contaadebitar = "";

    if(CODIGO_EMPRESA==="DP") contaadebitar = "Conta a Debitar: 40294310603";
    if(CODIGO_EMPRESA==="DL") contaadebitar = "Conta a Debitar: 40334199557";
    if(CODIGO_EMPRESA==="DF") contaadebitar = "Conta a Debitar: 40399078640";
    if(textoPDF.includes(contaadebitar)){ 
      Logger.log("CASO 5");  

      Logger.log("FILE: " + fileName);
      Logger.log("TEXTO: " + textoPDF);
      valorATCUD = extrairATCUDRecibosCA(textoPDF);
      Logger.log("ATCUD: " + valorATCUD);
      Logger.log("DATA: " + dataDocumento);

      if (!dataDocumento || dataDocumento.split("/").length !== 3) {
        fileErrors++;
        errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Caso válido mas data inválida ou não encontrada.\n";
        continue;
      }

      var month = dataDocumento.split("/")[1].toString();
          
      if( dataDocumento.split("/")[2].toString().length == 4){  // SE VEM NO FORMATO DD/MM/AAAA
        var year = dataDocumento.split("/")[2].toString();
        var day = dataDocumento.split("/")[0].toString();
      }
      else{                                                   // SE VEM NO FORMATO AAAA/MM/DD
        var year = dataDocumento.split("/")[0].toString();
        var day = dataDocumento.split("/")[2].toString();
      }

      var pastasFaturas = DriveApp.getFolderById(PASTA_GERAL_FATURAS).getFolders();

      while (pastasFaturas.hasNext()) {
        var pastaFaturas = pastasFaturas.next();
        if (pastaFaturas.getName() === year) {

          var pastaFaturasDentroDoAnoCerto = pastaFaturas.getFolders();

          while (pastaFaturasDentroDoAnoCerto.hasNext()) {
            var pastaFaturasMes = pastaFaturasDentroDoAnoCerto.next();

            // Tenta ir ao mês do comprovativo primeiro:
            if (pastaFaturasMes.getName() === "Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year) {

              existePastaMesAnoFaturas = 1;

              // Verifica se o ATCUD existe nos ATCUDs do ficheiro excel que se chama #0 - Faturas_+CODIGO_EMPRESA+"_"+month+"/"+year:

              var nomeFicheiroExcel = "#0 - Faturas_" + CODIGO_EMPRESA + "_" + month + "/" + year;
              var ficheirosDentroDaPasta = pastaFaturasMes.getFilesByName(nomeFicheiroExcel);

              if (!ficheirosDentroDaPasta.hasNext()) {
                Logger.log("⚠️ Ficheiro Excel não encontrado: " + nomeFicheiroExcel);
                fileErrors++;
                errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Ficheiro Excel não encontrado (" + nomeFicheiroExcel + "). \n";
                continue;
              }
              else 
                Logger.log("⚠️ Excel encontrado: " + nomeFicheiroExcel);

              var ficheiroExcel = ficheirosDentroDaPasta.next();
              var spreadsheet = SpreadsheetApp.open(ficheiroExcel);

              var tabsParaVerificar = ["Faturas e NCs normais", "Faturas e NCs com reembolso"];
              var atcudEncontrado = false;

              for (var i = 0; i < tabsParaVerificar.length; i++) {
                var sheet = spreadsheet.getSheetByName(tabsParaVerificar[i]);
                if (!sheet) continue;

                var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
                var colIndex = headers.indexOf("ATCUD / Nº Documento");
                if (colIndex === -1) continue;

                var dadosATCUD = sheet.getRange(3, colIndex + 1, sheet.getLastRow() - 2).getValues().flat();

                var linhaIndex = dadosATCUD.findIndex(function(valor) {
                  return valor === valorATCUD;
                });

                if (linhaIndex !== -1) {
                  var linhaEncontrada = linhaIndex + 3; // linha real na folha
                  Logger.log("✅ ATCUD encontrado no Excel, tab " + tabsParaVerificar[i] + ", linha " + linhaEncontrada-2); // O -dois é referente ao header, que tem que ser subtraído
                  atcudEncontrado = true;
                  break;
                }

              }

              if (!atcudEncontrado) {
                Logger.log("❌ ATCUD não encontrado no Excel: " + valorATCUD);
                fileErrors++;
                errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": ATCUD não encontrado no Excel (" + valorATCUD + "). \n";
                continue;
              }

              var iteradorPastasDentroDoMesEAnoCerto = pastaFaturasMes.getFolders();

              //Encontra pasta "#5 - Comprovativos de pagamento"
              while (iteradorPastasDentroDoMesEAnoCerto.hasNext()) {

                var pastaNivelUm = iteradorPastasDentroDoMesEAnoCerto.next();
                if (pastaNivelUm.getName() === "#5 - Comprovativos de pagamento") {

                  var iteradorPastasCatalogacao = pastaNivelUm.getFolders();

                  //Encontra pasta "PARA CATALOGAR"
                  while (iteradorPastasCatalogacao.hasNext()) {

                    var pastaNivelDois = iteradorPastasCatalogacao.next();
                    if (pastaNivelDois.getName() === "PARA CATALOGAR") {

                      Logger.log("ENTRY 3");

                      Logger.log("FOUND and copy made to "+pastaFaturasMes.getName());
                      copiarMoverELog_(file, pastaNivelDois, sourceFolder);

                      fileCount++;
                      nomesFicheirosMovidos += '\n Ficheiro ' + fileName + ' para pasta ' + pastaFaturasMes.getName() + '\n';
                    }

                  }
                }
              }
            }
          }
          
          if(!existePastaMesAnoFaturas){
            Logger.log("NÃO ENCONTRO A PASTA DAS FATURAS DESSE ANO E MÊS!");  
            fileErrors++;
            errosFicheirosMovidos += '\n Erro com ficheiro ' + fileName + ": Não existe pasta das faturas desse ano e mês " + "(Faturas_"+CODIGO_EMPRESA+"_"+month+"/"+year+"). \n";
          } 

        }
      }
      continue; // Ao ser distribuido, não irá verificar mais nenhum caso;

    }

  }

  var summary;

  // Caso em que catalogaram recibos de vencimento:

  if(recCount>0){
    summary = 'Foi EXECUTADA distribuição de ficheiros da pasta GERAL.\n\n' +
                'Movidos(s) ' + recCount + ' ficheiro(s) da pasta GERAL para pasta mãe dos RECIBOS (aguarda(m) catalogação).\n\n' +
                'Ficheiros em questão:\n' + nomesFicheirosMovidos;
    MailApp.sendEmail(EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL, "Distribuição de recibos de vencimento EXECUTADA ("+CODIGO_EMPRESA+")", summary);
  }  

  // Caso em que catalogaram faturas/documentos fiscais:

  if(fileCount>0){
    summary = 'Foi EXECUTADA distribuição de ficheiros da pasta GERAL.\n\n' +
                'Movidos(s) ' + fileCount + ' ficheiro(s) da pasta GERAL para pasta mãe das FATURAS (aguarda(m) catalogação).\n\n' +
                'Ficheiros em questão:\n' + nomesFicheirosMovidos;
    MailApp.sendEmail(EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL, "Distribuição de ficheiros EXECUTADA ("+CODIGO_EMPRESA+")", summary);

  }  

  // Caso em que *tentaram* catalogar faturas/documentos fiscais:

  if(fileErrors>0){
    summary = 'Foi TENTADA distribuição de ficheiros da pasta GERAL.\n\n' +
                'Encontrado(s) ' + fileErrors + ' erro(s) ao tentar mover ficheiro(s) da pasta GERAL para pasta mãe das FATURAS.\n\n' +
                'Erros:\n' + errosFicheirosMovidos;
    MailApp.sendEmail(EMAIL_NOTIFICACAO_DISTRIBUIDOR_GERAL, "Distribuição de ficheiros TENTADA ("+CODIGO_EMPRESA+")", summary);

  }  
  
  
}

// Função para obter o nome do mês
function obterNomeMes(mes) {
  var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  return meses[mes - 1];
}

// Função para obter a data e hora atual no formato AAAAMMDD HH:MM
function obterDataHoraAtual() {
  var now = new Date();
  var dia = ('0' + now.getDate()).slice(-2);
  var mes = ('0' + (now.getMonth() + 1)).slice(-2);
  var ano = now.getFullYear();
  var hora = ('0' + now.getHours()).slice(-2);
  var minuto = ('0' + now.getMinutes()).slice(-2);
  
  return ano + mes + dia + ' ' + hora + ':' + minuto;
}

function convertPDFToText(fileId, language) {
  if (!fileId) {
    throw new Error("convertPDFToText: fileId em falta.");
  }

  language = language || "pt"; // PT por omissão

  var file = DriveApp.getFileById(fileId);
  var mime = file.getMimeType();

  Logger.log("convertPDFToText: ficheiro=" + file.getName() + " mime=" + mime);

  // 1) Se já for um Google Docs, lemos texto direto (sem OCR)
  if (mime === MimeType.GOOGLE_DOCS || mime === "application/vnd.google-apps.document") {
    Logger.log("convertPDFToText: ficheiro já é Google Docs, a ler texto diretamente.");
    return DocumentApp.openById(fileId).getBody().getText();
  }

  // 2) Se for outro tipo Google Apps (Sheets, Slides, etc.), não faz sentido OCR
  if (mime.indexOf("application/vnd.google-apps.") === 0) {
    Logger.log("convertPDFToText: ficheiro é tipo Google Apps (" + mime + "), a tentar ler como Doc.");
    return DocumentApp.openById(fileId).getBody().getText();
  }

  // 3) Para PDFs/imagens, usamos OCR via Drive.Files.insert
  var maxTentativas = 4;
  var esperaMs = 5000; // 5s, depois 10s

  for (var tentativa = 1; tentativa <= maxTentativas; tentativa++) {
    try {
      var blob = file.getBlob();
      Logger.log("convertPDFToText: a fazer OCR, blob contentType=" + blob.getContentType());

      // NOTA: não definimos mimeType no resource -> deixamos o Drive decidir
      var ocrResult = Drive.Files.insert(
        {
          title: "OCR temp - " + file.getName()
          // sem mimeType aqui!
        },
        blob,
        {
          ocr: true,
          ocrLanguage: language
        }
      );

      if (!ocrResult || !ocrResult.id) {
        throw new Error("Falha a criar documento OCR temporário.");
      }

      var docId = ocrResult.id;
      var textContent = DocumentApp.openById(docId).getBody().getText();

      // apaga o Doc temporário
      DriveApp.getFileById(docId).setTrashed(true);
      return textContent;

    } catch (e) {
      var msg = (e && e.message) ? e.message : e.toString();
      var isOcrRateLimit =
        msg.indexOf("User rate limit exceeded for OCR") !== -1 ||
        msg.indexOf("User rate limit exceeded") !== -1;

      if (!isOcrRateLimit) {
        throw e; // erro real
      }

      Logger.log(
        "Limite de OCR atingido na tentativa " +
        tentativa +
        " para o ficheiro " + fileId +
        ". A aguardar " + (esperaMs / 1000) + "s. Mensagem: " + msg
      );

      if (tentativa === maxTentativas) {
        throw new Error(
          "Falhou o OCR para o ficheiro " +
          fileId +
          " após " + maxTentativas + " tentativas. Último erro: " + msg
        );
      }

      Utilities.sleep(esperaMs);
      esperaMs *= 2;
    }
  }

  throw new Error("convertPDFToText: chegou ao fim sem devolver texto.");
}

function getMovimentosSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_MOVIMENTOS_ID);
  let sh = ss.getSheetByName(SHEET_MOVIMENTOS_NOME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MOVIMENTOS_NOME);
    sh.appendRow([
      "Data",
      "Hora",
      "Nome do ficheiro",
      "Fonte (nome)",
      "Pais do destino (nome)",
      "Link atual"
    ]);
  }
  return sh;
}

function setLinkCell_(sh, row, col, text, url) {
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(text || "")
    .setLinkUrl(url || null)
    .build();
  sh.getRange(row, col).setRichTextValue(rich);
}

// devolve apenas os últimos N pais (sem “O meu disco” e sem a própria pasta)
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
  if (chain.length) chain.pop(); // remove a própria pasta destino
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

  if (fonteFolder) {
    setLinkCell_(sh, row, 4, fonteFolder.getName(), fonteFolder.getUrl());
  }
  if (destinoFolder) {
    const caminhoCurto = buildDestinoParentsTail_(destinoFolder, DESTINO_PATH_DEPTH);
    setLinkCell_(sh, row, 5, caminhoCurto, destinoFolder.getUrl());
  }
  if (novoFicheiroUrl) {
    setLinkCell_(sh, row, 6, "LINK", novoFicheiroUrl);
  }
}

/**
 * Move “standardizado”: copia para o destino, manda o original para o LIXO
 * e regista o movimento na sheet.
 */
function copiarMoverELog_(file, destinoFolder, fonteFolder) {
  const novoFicheiro = file.makeCopy(destinoFolder);
  file.moveTo(DriveApp.getFolderById(PASTA_LIXO));
  registarMovimento_(file.getName(), fonteFolder, destinoFolder, novoFicheiro.getUrl());
  return novoFicheiro;
}

function extractTipoDocumento(text) {

  text = text.replace(/[/_-]/g, ' '); // retira _, -, etc
  text = text.replace(/\s{2,}/g, ' '); // retira todos os espaços menos 1 se houverem vários, por exemplo "fatura     recibo" -> "fatura recibo"

  //Logger.log("TEXTO: " + text);

  // Esta parte é para não confundir o 
  var pattern = /válido como recibo após/g;
  var match = text.match(pattern);
  if (match) {
    text = text.replace(pattern, ''); // Replace the matched expression with an empty string
  }

  // List of possible tipo values
  // Nunca mudar esta ordem!
  var tipos = [
    'fatura simplificada',
    'factura simplificada',
    'nota de crédito',
    'fatura recibo',
    'factura recibo',
    '2ª via',
    'segunda via',
    'fatura',
    'factura',
    'recibo de renda',
    'recibo'
  ];

  // Corresponding output tipo values with consistent case
  var tiposOutput = [
    'Fatura simplificada',
    'Fatura simplificada',
    'Nota de crédito',
    'Fatura-recibo',
    'Fatura-recibo',
    '2ª via fatura',
    '2ª via fatura',
    'Fatura',
    'Fatura',
    'Recibo de renda',
    'Recibo'
  ];

  // Iterate over the list of possible tipos
  for (var i = 0; i < tipos.length; i++) {
    // Check if the tipo is present in the text (case insensitive)
    if (text.toLowerCase().includes(tipos[i])) {
      return tiposOutput[i]; // Return the matched tipo with consistent case
    }
  }

  // If no tipo is found, return null
  return null;
}

function extractATCUD(pdfText) {

  //Logger.log("TEXTO: " + pdfText);
  
  if (pdfText) {
    // Match the pattern "ATCUD:" followed by any characters until a space or end of line
    var regex = /ATCUD:\s*([^\s]+)/;
    var match = pdfText.match(regex);
    if (match) {
      return match[1]; // Extract the code after "ATCUD:"
    }
    else{
      var regex = /ATCUD\s*([^\s]+)/;
      var match = pdfText.match(regex);
      if (match) {
        return match[1]; // Extract the code after "ATCUD:"
      }
    }
  }
  return null; // Return null if ATCUD is not found
}

function extrairATCUDRecibosCA(texto) {

  var regex = /Informação Complementar:\s*([A-Z0-9]{2,}-[A-Z0-9]+)/i;
  var match = texto.match(regex);
  if (match) {
    return match[1]; // O ATCUD
  }
  return null;
}

function extractDataDocumento(pdfText) {
  const today = new Date();
  const results = [];

  // Lista de regex + lógica de extração
  const patterns = [
    {
      regex: /(\d{2}\/\d{2}\/\d{4})|(\d{2}-\d{2}-\d{4})|(\d{4}-\d{2}-\d{2})/g,
      parser: (m) => m.replace(/-/g, "/")
    },
    {
      regex: /(\d{1,2})\s+de\s+(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+(\d{4})/gi,
      parser: (m) => {
        const match = m.match(/(\d{1,2})\s+de\s+([a-zç]+)\s+(\d{4})/i);
        const meses = {
          janeiro: "01", fevereiro: "02", março: "03", abril: "04", maio: "05", junho: "06",
          julho: "07", agosto: "08", setembro: "09", outubro: "10", novembro: "11", dezembro: "12"
        };
        return `${match[1]}/${meses[match[2].toLowerCase()]}/${match[3]}`;
      }
    },
    {
      regex: /(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(\d{4})/gi,
      parser: (m) => {
        const match = m.match(/(\d{1,2})\s+([a-z]+)\s+(\d{4})/i);
        const meses = {
          jan: "01", fev: "02", mar: "03", abr: "04", mai: "05", jun: "06",
          jul: "07", ago: "08", set: "09", out: "10", nov: "11", dez: "12"
        };
        return `${match[1]}/${meses[match[2].toLowerCase()]}/${match[3]}`;
      }
    },
    {
      regex: /(\d{1,2})-(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)-(\d{4})/gi,
      parser: (m) => {
        const match = m.match(/(\d{1,2})-([A-Z]+)-(\d{4})/i);
        const meses = {
          jan: "01", fev: "02", mar: "03", abr: "04", mai: "05", jun: "06",
          jul: "07", ago: "08", set: "09", out: "10", nov: "11", dez: "12"
        };
        return `${match[1]}/${meses[match[2].toLowerCase()]}/${match[3]}`;
      }
    },
    {
      regex: /(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s+(\d{4})/gi,
      parser: (m) => {
        const match = m.match(/([a-z]+)\s+(\d{1,2}),\s+(\d{4})/i);
        const meses = {
          january: "01", february: "02", march: "03", april: "04", may: "05", june: "06",
          july: "07", august: "08", september: "09", october: "10", november: "11", december: "12"
        };
        return `${match[2]}/${meses[match[1].toLowerCase()]}/${match[3]}`;
      }
    },
    {
      regex: /(\d{4})\/(\d{2})\/(\d{2})/g,
      parser: (m) => {
        const parts = m.split("/");
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
      }
    },
    {
      regex: /(\d{2}\/\d{2}\/\d{4})|(\d{2}-\d{2}-\d{4})|(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{2})/g,
      parser: (m) => {
        const clean = m.replace(/-/g, "/");
        const parts = clean.split("/");
        let day, month, year;

        if (parts[0].length === 4) {
          year = parts[0]; month = parts[1]; day = parts[2];
        } else if (parts[2].length === 2) {
          day = parts[0]; month = parts[1]; year = "20" + parts[2];
        } else {
          day = parts[0]; month = parts[1]; year = parts[2];
        }

        return `${day}/${month}/${year}`;
      }
    },
    {
      regex: /(\d{2})\.(\d{2})\.(\d{4})/g,
      parser: (m) => m.replace(/\./g, "/")
    },
    {
      regex: /(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})/gi,
      parser: (m) => {
        const match = m.match(/(\d{1,2})\s+([a-z]+)\s+(\d{4})/i);
        const meses = {
          jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
          jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
        };
        return `${match[1]}/${meses[match[2].toLowerCase()]}/${match[3]}`;
      }
    },
    {
      regex: /(\d{2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*-(\d{4})/gi,
      parser: (m) => {
        const match = m.match(/(\d{2})-([a-z]+)-(\d{4})/i);
        const meses = {
          jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
          jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
        };
        return `${match[1]}/${meses[match[2].toLowerCase()]}/${match[3]}`;
      }
    }
  ];

  // Processar todos os padrões
  for (let { regex, parser } of patterns) {
    const matches = [...pdfText.matchAll(regex)];

    for (let match of matches) {
      try {
        const formatted = parser(match[0]);
        const [d, m, y] = formatted.split("/").map(n => parseInt(n));
        const date = new Date(y, m - 1, d);

        if (date <= today) {
          return formatted;
        }
      } catch (e) {
        // Ignora erros de parsing
        continue;
      }
    }
  }

  return null;
}




















































































// Esta função testa todas as faturas de 2025 ou uma pasta específica
// Como tem tempo limitado, guarda a última pasta que viu e recomeça aí na execução seguinte
function bancadaDeTestes2025() {

  // === MODO DE EXECUÇÃO ===
  // Se estiver vazio => percorre toda a árvore do ano (Faturas_DL_MM/AAAA -> #1 e #2).
  // Se tiver um ID => processa APENAS essa pasta.
  const ONLY_FOLDER_ID = "102LMZvC1914IhB90_zJRvobOo5t-mv9G"; // ex.: "1AbCdefGhIJkLmNoP"

  // (opcionais) override do esperado quando usas ONLY_FOLDER_ID
  const ONLY_EXPECTED_MM   = "";   // "01".."12" ou "" para NA
  const ONLY_EXPECTED_YEAR = 2025; // usado só se ONLY_EXPECTED_MM tiver valor

  function _isOCRRateLimitError_(err) {
    return String(err).indexOf('User rate limit exceeded for OCR') >= 0;
  }


  // === CONTROLO DE RETOMA =====================

  const RESET_PROGRESS = false; // <-- muda para true para recomeçar do zero, caso contrário continua onde parou na última execução

  if (RESET_PROGRESS) {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty("BENCH_IDX");
    // apaga tokens de paginação e posição
    const all = props.getKeys();
    all.forEach(k => {
      if (k.startsWith("BENCH_TOKEN_")) props.deleteProperty(k);
      if (k.startsWith("BENCH_POS_")) props.deleteProperty(k);
    });
    Logger.log("🔄 Progresso apagado → recomeça do zero na próxima execução.");
    // não sai — deixa continuar, agora vazio
  }





  // ======= CONFIGURAÇÃO =======
  const YEAR = 2025;
  const PARENT_FOLDER_ID = "1z3HZIF1EoaCbQyAPSU2XNVnouFc62x3M";

  // ======= TEMPO / RETOMA =======
  const TIME_BUDGET_MS = 5 * 60 * 1000;
  const SLEEP_MS = 120;
  const start = Date.now();

  // ======= PREPARA SHEET =======
  const { ss, sheet } = _bench_createNewSheet_();
  _bench_writeHeaderIfEmpty_(sheet);
  _bench_ensureFormatting_(sheet);

  // ======= ALVO DINÂMICO + RETOMA =======
  const props = PropertiesService.getScriptProperties();
  const IDX_KEY = "BENCH_IDX";
  const folders = ONLY_FOLDER_ID
  ? [[ONLY_EXPECTED_MM || "", ONLY_FOLDER_ID]]           // [[tag, folderId]]
  : _bench_getTargetsFromParent_(PARENT_FOLDER_ID, YEAR);

  let idx = Number(props.getProperty(IDX_KEY) || 0);
  if (idx < 0 || idx >= folders.length) idx = 0;

  // Helper para flush + guardar estado e sair
  function _flushAndCheckpointAndReturn_(sheet, rows, {TOKEN_KEY, pageToken, POS_KEY, pos, idx, tag, reason}) {
    if (rows && rows.length) {
      const r = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length);
      r.setValues(rows);
      rows.length = 0;
    }
    if (TOKEN_KEY) props.setProperty(TOKEN_KEY, pageToken || "");
    if (POS_KEY != null) props.setProperty(POS_KEY, String(pos || 0));
    props.setProperty(IDX_KEY, String(idx));
    Logger.log("[%s] ⚠️ %s — retoma guardada.", tag, reason || "Checkpoint");
    return true; // sinal para sair do chamador
  }

  // ======= LOOP POR PASTAS =======
  for (; idx < folders.length; idx++) {

    const [mm, folderId] = folders[idx];
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log("[%s] ERRO: pasta inválida (%s) | %s", mm, folderId, String(e));
      continue;
    }
    const folderName = folder.getName();
    const tag = mm; // tag = mês
    Logger.log("[%s] ===== Pasta: %s (%s) =====", tag, folderName, folderId);

    // Paginação Drive v2
    const TOKEN_KEY = "BENCH_TOKEN_" + folderId;
    const POS_KEY   = "BENCH_POS_"   + folderId;

    let pageToken = props.getProperty(TOKEN_KEY) || null;

    const q = `'${folderId}' in parents and trashed=false`;
    const fields = "nextPageToken,items(id,title,mimeType,fileSize)";
    const MAX_RESULTS = 50;

    while (true) {
      if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) {
        if (_flushAndCheckpointAndReturn_(sheet, [], {TOKEN_KEY, pageToken, POS_KEY, pos:0, idx, tag, reason:"Time budget (antes da lista)"})) return;
      }

      const resp = Drive.Files.list({ q, pageToken, maxResults: MAX_RESULTS, fields, orderBy: "title" });
      const items = resp.items || [];
      if (!items.length) {
        props.deleteProperty(TOKEN_KEY);
        props.deleteProperty(POS_KEY);
        break;
      }

      // posição dentro da página (retoma)
      let pos = Number(props.getProperty(POS_KEY) || 0);
      if (pos < 0 || pos > items.length) pos = 0;

      const rows = [];

      // usa loop indexado para respeitar 'pos'
      for (let i = pos; i < items.length; i++) {
        const it = items[i];

        // time budget check dentro do batch
        if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) {
          if (_flushAndCheckpointAndReturn_(sheet, rows, {TOKEN_KEY, pageToken, POS_KEY, pos: i, idx, tag, reason:"Time budget (no batch)"})) return;
        }

        const mime = (it.mimeType || "").toLowerCase();
        const title = it.title || "";
        if (!mime.includes("pdf") && !/\.pdf$/i.test(title)) {
          // mesmo assim avança a posição para não reprocessar
          props.setProperty(POS_KEY, String(i + 1));
          continue;
        }

        let data = null, origem = "", nota = "";
        const MAX_ATTEMPTS = 6;

        for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
          try {
            const texto = ocrPDF_(it.id, "pt");

            data = extractDataDocumentoTaloes(texto);
            origem = data ? "SIMPLES" : "";

            if (!data) {
              data = extractDataDocumento_Simples_PERMISSIVA(texto);
              origem = data ? "PERMISSIVA" : "";
            }
            break; // sucesso

          } catch (e) {
            if (_isOCRRateLimitError_(e)) {
              if (attempt < MAX_ATTEMPTS) {
                const delay = 500 * Math.pow(2, attempt - 1); // 0.5s,1s,2s,4s,8s
                Utilities.sleep(delay);
                continue;
              } else {
                // flush + checkpoint + sair
                _flushAndCheckpointAndReturn_(sheet, rows, {TOKEN_KEY, pageToken, POS_KEY, pos: i, idx, tag, reason:`OCR rate limit após ${MAX_ATTEMPTS} tentativas`});
                return;
              }
            } else {
              origem = "ERRO";
              nota = String(e);
              break;
            }
          }
        }

        if (!data && !origem) {
          origem = "ERRO";
          nota = nota || "Sem data extraída";
        }

        const expected = (ONLY_FOLDER_ID && ONLY_EXPECTED_MM)
          ? `${ONLY_EXPECTED_MM}/${ONLY_EXPECTED_YEAR}`
          : (tag ? `${tag}/${YEAR}` : "");
          
        const verdict = _bench_compare_(data, expected);

        rows.push([
          new Date(),            // Timestamp (escrita)
          tag,                   // Tag (MM)
          folderName,            // Pasta
          title,                 // Ficheiro
          data || "",            // Data Detectada (DD/MM/AAAA)
          expected || "",        // Esperado (MM/AAAA)
          verdict,               // Veredicto
          origem || "",          // Origem
          `=HYPERLINK("https://drive.google.com/file/d/${it.id}/view","${it.id}")`,
          nota
        ]);

        // avançar posição e persistir retoma
        props.setProperty(POS_KEY, String(i + 1));

        Utilities.sleep(SLEEP_MS);
      } // for i

      // flush final da página
      if (rows.length) {
        const shRange = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length);
        shRange.setValues(rows);
      }

      // próxima página
      pageToken = resp.nextPageToken || null;
      if (!pageToken) {
        props.deleteProperty(TOKEN_KEY);
        props.deleteProperty(POS_KEY);
        break;
      }
    } // while páginas

    // terminou a pasta; limpa token/pos e avança idx
    props.deleteProperty("BENCH_TOKEN_" + folderId);
    props.deleteProperty("BENCH_POS_" + folderId);
    props.setProperty(IDX_KEY, String(idx + 1));
    Logger.log("[%s] -------- FIM PASTA --------", tag);
  } // for

  // terminou tudo
  PropertiesService.getScriptProperties().deleteProperty("BENCH_IDX");
  Logger.log("✅ Bancada concluída.");
}

function _bench_ensureFormatting_(sheet) {
  // Regras existentes...
  const rules = sheet.getConditionalFormatRules();
  if (!rules || !rules.length) {
    const colVerd = 7, lastRow = Math.max(sheet.getLastRow(), 1000);
    const okRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OK").setBackground("#2e7d32").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, colVerd, lastRow)]).build();
    const failRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("FAIL").setBackground("#c62828").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, colVerd, lastRow)]).build();
    sheet.setConditionalFormatRules([okRule, failRule]);
  }

  // **Aqui está o que interessa:** mostrar data+hora no Timestamp
  sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
}


/** Reset à retoma (recomeça do início) */
function benchDatasPDF_Reset() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getKeys();
  all.forEach(k => { if (k.indexOf("BENCH_") === 0) props.deleteProperty(k); });
  Logger.log("Retoma limpa.");
}

function _bench_createNewSheet_() {
  const tz = "Europe/Lisbon";
  const name = "Bench Datas PDF - " + Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");
  const ss = SpreadsheetApp.create(name);
  ss.setSpreadsheetTimeZone(tz);

  // cria a tab "Resultados"
  let sheet = ss.getSheetByName("Resultados");
  if (!sheet) sheet = ss.insertSheet("Resultados"); else sheet.setName("Resultados");

  // apaga todas as outras tabs
  ss.getSheets()
    .filter(sh => sh.getName() !== "Resultados")
    .forEach(sh => ss.deleteSheet(sh));

  return { ss, sheet };
}

function _bench_writeHeaderIfEmpty_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1,1,1,10).setValues([[
    "Timestamp (DD/MM/AAAA HH:MM:SS)",
    "Tag","Pasta","Ficheiro",
    "Data Detectada (DD/MM/AAAA)", // isto é a da fatura, mantém-se
    "Esperado (MM/AAAA)","Veredicto","Origem","File ID","Nota"
  ]]);
  sheet.setFrozenRows(1);
}

/** Para tags 01..12 e R01..R12 devolve "MM/YYYY"; para tags sem mês, "" */
function _bench_expectedFromTag_(tag, year) {
  const m = String(tag || "");
  if (/^R?\d{2}$/.test(m)) {
    const mm = m.replace(/^R/,'');
    return `${mm}/${year}`;
  }
  return "";
}

/** Compara: "DD/MM/YYYY[ HH:MM]" vs "MM/YYYY" → OK/FAIL/NA */
function _bench_compare_(detected, expected) {
  if (!expected) return "NA";
  if (!detected) return "FAIL";

  // apanha só a parte da data antes do espaço (se houver hora)
  const datePart = String(detected).trim().split(/\s+/)[0]; // "DD/MM/YYYY"
  const dm = datePart.split("/");
  if (dm.length !== 3) return "FAIL";

  const mm = dm[1], yyyy = dm[2];
  const [expMM, expYY] = expected.split("/");
  return (mm === expMM && yyyy === expYY) ? "OK" : "FAIL";
}



























function listarDatasEmVariasPastas() {
  // ======= CONFIGURAÇÃO =======
  const YEAR = 2025;
  const PARENT_FOLDER_ID = "1z3HZIF1EoaCbQyAPSU2XNVnouFc62x3M"; // pasta-mãe do ano

  // ======= TEMPO / RETOMA =======
  const TIME_BUDGET_MS = 5 * 60 * 1000;
  const SLEEP_MS = 120;
  const start = Date.now();

  // ======= PREPARA SHEET =======
  const { ss, sheet } = _bench_createNewSheet_();
  Logger.log("Sheet pronta em: %s", ss.getUrl());
  _bench_writeHeaderIfEmpty_(sheet);
  _bench_ensureFormatting_(sheet);

  // ======= ALVO DINÂMICO + RETOMA =======
  const props = PropertiesService.getScriptProperties();
  const IDX_KEY = "BENCH_IDX";
  const folders = _bench_getTargetsFromParent_(PARENT_FOLDER_ID, YEAR); // array de [mm, folderId]

  let idx = Number(props.getProperty(IDX_KEY) || 0);
  if (idx < 0 || idx >= folders.length) idx = 0;

  // ======= LOOP POR PASTAS =======
  for (; idx < folders.length; idx++) {
    if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) break;

    const [mm, folderId] = folders[idx];
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log("[%s] ERRO: pasta inválida (%s) | %s", mm, folderId, String(e));
      continue;
    }
    const folderName = folder.getName();
    const tag = mm; // tag = mês
    Logger.log("[%s] ===== Pasta: %s (%s) =====", tag, folderName, folderId);

    // Paginação Drive v2
    const TOKEN_KEY = "BENCH_TOKEN_" + folderId;
    let pageToken = props.getProperty(TOKEN_KEY) || null;

    const q = `'${folderId}' in parents and trashed=false`;
    const fields = "nextPageToken,items(id,title,mimeType,fileSize)";
    const MAX_RESULTS = 50;

    while (true) {
      if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) {
        props.setProperty(TOKEN_KEY, pageToken || "");
        props.setProperty(IDX_KEY, String(idx));
        Logger.log("[%s] ⚠️ Time budget, guardado token para retoma.", tag);
        return;
      }

      const resp = Drive.Files.list({
        q, pageToken, maxResults: MAX_RESULTS, fields, orderBy: "title"
      });

      const items = resp.items || [];
      if (!items.length) {
        props.deleteProperty(TOKEN_KEY);
        break;
      }

      const rows = [];
      for (let it of items) {
        const mime = (it.mimeType || "").toLowerCase();
        const title = it.title || "";
        if (!mime.includes("pdf") && !/\.pdf$/i.test(title)) continue;

        let data = null, origem = "", nota = "";

        try {
          const texto = ocrPDF_(it.id, "pt");
          data = extractDataDocumento_Simplesv1(texto);
          origem = data ? "SIMPLES" : "";

          if (!data) {
            data = extractDataDocumento_Simples_PERMISSIVA(texto);
            origem = data ? "PERMISSIVA" : "";
          }

        } catch (e) {
          data = null;
          origem = "ERRO";
          nota = String(e);
        }

        // Esperado = MM/YYYY (tag é o mês)
        const expected = `${tag}/${YEAR}`;
        const verdict = _bench_compare_(data, expected);

        rows.push([
          new Date(),            // Timestamp (escrita)
          tag,                   // Tag (MM)
          folderName,            // Pasta
          title,                 // Ficheiro
          data || "",            // Data Detectada (DD/MM/AAAA)
          expected || "",        // Esperado (MM/AAAA)
          verdict,               // Veredicto
          origem || "",          // Origem
          `=HYPERLINK("https://drive.google.com/file/d/${it.id}/view","${it.id}")`,
          nota
        ]);

        Utilities.sleep(SLEEP_MS);
        if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) break;
      }

      if (rows.length) {
        const shRange = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length);
        shRange.setValues(rows);
      }

      pageToken = resp.nextPageToken || null;
      if (!pageToken) {
        props.deleteProperty(TOKEN_KEY);
        break;
      }
    } // while páginas

    // terminou a pasta; limpa token e avança idx
    props.deleteProperty("BENCH_TOKEN_" + folderId);
    props.setProperty(IDX_KEY, String(idx + 1));
    Logger.log("[%s] -------- FIM PASTA --------", tag);
  } // for

  // terminou tudo
  PropertiesService.getScriptProperties().deleteProperty("BENCH_IDX");
  Logger.log("✅ Bancada concluída.");
}

/** Constrói (ou lê do cache) a lista de pastas a processar para o ano */
function _bench_getTargetsFromParent_(parentFolderId, year) {
  const props = PropertiesService.getScriptProperties();
  const KEY = "BENCH_TARGETS_CACHE";
  const cached = props.getProperty(KEY);
  if (cached) {
    try {
      const arr = JSON.parse(cached);
      if (Array.isArray(arr) && arr.length) return arr;
    } catch (_) {}
  }

  // Sem cache -> varrer Drive
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('Ative o Advanced Drive Service (Drive v2).');
  }

  // 1) Encontrar as "Faturas_DL_MM/2025" diretamente na pasta-mãe
  const parentQ = `'${parentFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const fields = "nextPageToken,items(id,title)";
  const targets = []; // array de [tag, folderId]

  let pageToken = null;
  const monthRe = new RegExp(`^Faturas_DL_(\\d{2})\\/${year}$`, 'i');

  do {
    const resp = Drive.Files.list({
      q: parentQ,
      pageToken,
      maxResults: 1000,
      fields,
      orderBy: "title"
    });
    const items = resp.items || [];
    for (const it of items) {
      const title = it.title || "";
      const m = title.match(monthRe);
      if (!m) continue;
      const mm = m[1];

      // 2) Dentro de cada mês, procurar #1 e #2
      const subQ = `'${it.id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
      let subToken = null;
      do {
        const sub = Drive.Files.list({
          q: subQ,
          pageToken: subToken,
          maxResults: 1000,
          fields,
          orderBy: "title"
        });
        const subs = sub.items || [];
        for (const sf of subs) {
          const nm = (sf.title||"").toLowerCase();
          if (nm === "#1 - faturas e ncs normais".toLowerCase() ||
              nm === "#2 - faturas e ncs com reembolso".toLowerCase()) {
            // Tag = MM (mesmo esperado para #1 e #2)
            targets.push([mm, sf.id]);
          }
        }
        subToken = sub.nextPageToken || null;
      } while (subToken);
    }
    pageToken = resp.nextPageToken || null;
  } while (pageToken);

  // Ordena por mês crescente e por nome para estabilidade
  targets.sort((a,b)=> (a[0]===b[0]) ? 0 : (a[0] < b[0] ? -1 : 1));

  if (!targets.length) {
    throw new Error("Não encontrei nenhuma pasta 'Faturas_DL_MM/" + year + "'. Verifica a pasta-mãe.");
  }

  props.setProperty(KEY, JSON.stringify(targets));
  return targets;
}

/** Limpa o cache de targets (se a estrutura de pastas mudar) */
function benchTargets_ResetCache() {
  PropertiesService.getScriptProperties().deleteProperty("BENCH_TARGETS_CACHE");
  Logger.log("Cache de targets limpa.");
}

/** OCR via Drive v2 -> Google Doc -> texto (apaga o doc temporário).
 *  Evita "Invalid Value" usando só 1 idioma por tentativa.
 */
function ocrPDF_(fileId) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('Ative o Advanced Drive Service (Drive v2).');
  }
  const pdf  = DriveApp.getFileById(fileId);
  const blob = pdf.getBlob();
  const langs = ['pt', 'en', null]; // PT -> EN -> sem idioma
  let lastErr;

  for (var i = 0; i < langs.length; i++) {
    const lang = langs[i];
    let created = null;
    try {
      // Tenta sem forçar mimetype; se falhar, força application/pdf
      try {
        created = Drive.Files.insert(
          { title: pdf.getName().replace(/\.pdf$/i, '') },
          blob,
          { ocr: true, ocrLanguage: lang || undefined, fields: 'id' }
        );
      } catch (_) {
        created = Drive.Files.insert(
          { title: pdf.getName().replace(/\.pdf$/i, ''), mimeType: 'application/pdf' },
          blob,
          { ocr: true, ocrLanguage: lang || undefined, fields: 'id' }
        );
      }

      // Pequeno polling: o Doc pode demorar ~200–800ms a “ficar pronto”
      var text = '';
      var t0 = Date.now();
      while (!text && (Date.now() - t0) < 4000) {
        try {
          text = DocumentApp.openById(created.id).getBody().getText() || '';
          if (text.trim()) break;
        } catch (e) {
          // ainda a preparar — espera e tenta de novo
        }
        Utilities.sleep(200);
      }

      if (text && text.trim()) return text;
      // se vazio, tenta próximo idioma
    } catch (err) {
      lastErr = err;
      // se for Invalid Value com idioma, avança para o próximo
      if (String(err).indexOf('Invalid Value') >= 0) continue;
    } finally {
      try { if (created && created.id) DriveApp.getFileById(created.id).setTrashed(true); } catch (_) {}
    }
  }
  throw lastErr || new Error('OCR falhou em todas as tentativas');
}






















function extractDataDocumentoTaloes(pdfText) {
  if (!pdfText) return null;

  // ===== Helpers =====
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
  const injectSpaces = s => String(s)
    .replace(/\u00A0/g, ' ')
    .replace(/([A-Za-z])(\d)/g, '$1 $2')
    .replace(/(\d)([A-Za-z])/g, '$1 $2');

  // ===== Normalização =====
  const raw = injectSpaces(pdfText);
  const linesAll = raw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

  // ===== Filtros “hard” (2ª via / gerada em …) =====
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

  // ===== Filtros POS/Comprovativo =====
  const badPOS = [
    /\b(pos|tpa|tp[ãa]g|terminal|redeunic|multibanco|mb\s*way|sibs)\b/i,
    /\bvisa\b/i, /\bmastercard\b/i, /\bmaestro\b/i, /\bam[ex|erican\s*express]\b/i,
    /\bautoriz[aç][aã]o\b/i, /\bauth(?:orization)?\b/i,
    /\baid\b/i, /\batc\b/i, /\btid\b/i, /\bnsu\b/i, /\bpan\b/i, /\barqc?\b/i,
    /\bcomprovativo\b/i, /\breceb[ií]do\b/i, /\bmerchant\s*copy\b/i, /\bclient\s*copy\b/i,
    /\blote\b/i, /\bref(?:\.|er[eê]ncia)?\b/i, /\btransa[cç][aã]o\b/i, /\bvenda\b/i, /\bpagamento\b/i,
  ];
  const timeRe = /\b\d{2}:\d{2}(?::\d{2})?\b/; // HH:MM(:SS)

  const lines = linesAll.filter(l => !isHardBad(l));

  // ===== Fast path com cabeçalho/documento =====
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

  // ===== Labels “boas” e “más” =====
  const goodLabels = [
    /data\s*(?:de)?\s*emiss[aã]o/i, /\bemiss[aã]o\b/i, /\bemitida\b/i, /\bemitida\s+em\b/i,
    /dt\.?\s*emiss[aã]o/i, /\bdata\s*doc(?:umento)?\b/i, /\bdata\s*:\b/i,
    /\binvoice\s+date\b/i, /\bissue\s+date\b/i, /\bfecha\s+de\s+emisi[oó]n\b/i
  ];
  const badLabels = [
    /\bvig[êe]ncia\b/i, /\bper[ií]odo\b/i, /\bvalidade\b/i, /\bcompet[êe]ncia\b/i,
    /\bref\.\s*(?:per[ií]odo|m[êe]s)\b/i, /\bintervalo\b/i,
    /\bvenc(?:imento)?\b/i, /\bprazo\b/i, /\bdue\b/i, /\bpayment\b/i
  ];
  const hasRangeRe = /\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*(?:a|–|—|-)\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/i;

  // ISO junto de labels boas
  const oneLine = raw.replace(/\s+/g, ' ');
  const isoNearLabel = new RegExp('(?:' + goodLabels.map(r=>r.source).join('|') + ')' +
    '[^0-9]{0,40}(20\\d{2})[./-](\\d{2})[./-](\\d{2})','i');
  let m = oneLine.match(isoNearLabel);
  if (m) {
    const fast = _safeDate_(m[3], m[2], m[1]);
    if (fast) return fast;
  }

  // ===== Scanner =====
  const rx = [
    { re:/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/g, norm:(y,m,d)=>({d,m,y, iso:true}) },
    { re:/\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/g,     norm:(d,m,y)=>({d,m,y}) },
    { re:/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\b/g,         norm:(d,m,y)=>({d,m,y, ambiguousYY:true}) },
    { re:/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})\b/gi,
      norm:(d,mon,y)=>({d, m:_monPT_(mon), y}) },
    { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})\b/gi,
      norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
    { re:/\b(\d{1,2})[.\- ](jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[.\- ](\d{4})\b/gi,
      norm:(d,mon,y)=>({d, m:_monPTabbrev_(mon), y}) }
  ];

  const candidates = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line) continue;
    if (hasRangeRe.test(line)) continue;
    if (badLabels.some(r => r.test(line))) continue;

    // descarta linhas típicas de talão
    const looksPOS = badPOS.some(r => r.test(line)) || timeRe.test(line);
    // (não descartamos já: ainda podemos apanhar uma data boa com label forte; avaliamos mais abaixo)

    const hasGood = goodLabels.some(r => r.test(line));

    for (const {re, norm} of rx) {
      re.lastIndex = 0;
      let mm;
      while ((mm = re.exec(line)) !== null) {
        const p = norm(...mm.slice(1));
        const dd2 = String(p.d).padStart(2,'0');
        const mm2 = String(p.m).padStart(2,'0');
        let yyyy = String(p.y);

        // Regras POS + ambíguo:
        // - Se é ambíguo (YY) e a linha parece POS e não há label boa → IGNORA este match.
        if ((p.ambiguousYY || yyyy.length === 2) && looksPOS && !hasGood) {
          continue;
        }

        // Reconciliar YY com ISO na mesma linha (se existir)
        if (p.ambiguousYY || yyyy.length === 2) {
          const reconciled = _reconcileYearWithNearbyISO(line, mm.index, dd2, mm2, 80);
          if (reconciled) {
            yyyy = reconciled;
          } else {
            // Sem ISO e sem label boa: se ainda é ambíguo e a linha parece POS → descarta.
            if (!hasGood && looksPOS) continue;
            // fallback YY->YYYY (00–79 => 2000+, 80–99 => 1900+)
            const n = +yyyy; yyyy = (n <= 79) ? (2000 + n) : (1900 + n);
          }
        }

        const safe = _safeDate_(dd2, mm2, yyyy);
        if (!safe) continue;

        // proximidade a labels boas
        const near = goodLabels.some(r => {
          r.lastIndex = 0;
          const lm = r.exec(line);
          return lm && lm.index >= 0 && (mm.index - lm.index) <= 30;
        });

        // scoring
        let score = 0;
        if (p.iso) score += 60;            // mais bónus para ISO
        if (hasGood) score += 120;
        if (near) score += 20;
        if (i < 12) score += 30;
        if (line.length <= 80) score += 5;

        // Penalização forte se parece POS e não há label boa
        if (looksPOS && !hasGood) score -= 150;

        if (score > 0) {
          candidates.push({ safe, score, lineIndex: i, colIndex: mm.index });
        }
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

  // ===== Fallback rápido (sem intervalos e sem linhas POS) =====
  const cleaned = lines
    .filter(l => !hasRangeRe.test(l) && !badPOS.some(r => r.test(l)))
    .join('\n');
  const quick = [
    { r:/\b(20\d{2})[./-](\d{2})[./-](\d{2})\b/g, p:s=>{const a=s.split(/[./-]/); return _safeDate_(a[2],a[1],a[0]);} },
    { r:/\b(\d{2})[./-](\d{2})[./-](\d{4})\b/g,   p:s=>{const a=s.split(/[./-]/); return _safeDate_(a[0],a[1],a[2]);} }
  ];
  for (const {r,p} of quick) {
    let q; r.lastIndex=0;
    while ((q=r.exec(cleaned))!==null) {
      const out = p(q[0]);
      if (out) return out;
    }
  }
  return null;

  // ===== helpers meses =====
  function _monPT_(s){ const m={janeiro:1,fevereiro:2,'março':3,marco:3,abril:4,maio:5,junho:6,julho:7,agosto:8,setembro:9,outubro:10,novembro:11,dezembro:12}; return m[String(s).toLowerCase()]||s; }
  function _monPTabbrev_(s){ const m={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12}; return m[String(s).toLowerCase().slice(0,3)]||s; }
  function _monEN_(s){ const m={january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12}; return m[String(s).toLowerCase()]||s; }
}



function extractDataDocumento_Simplesv1(pdfText) {
  if (!pdfText) return null;

  //HELPER
  function _reconcileYearWithNearbyISO(line, matchIndex, dd, mm, windowChars) {
    const around = windowChars || 60;
    const left  = Math.max(0, matchIndex - around);
    const right = Math.min(line.length, matchIndex + around);
    const ctx = line.slice(left, right);

    const iso = ctx.match(/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/);
    if (!iso) return null;

    const isoY = iso[1];
    const isoM = String(iso[2]).padStart(2,'0');
    const isoD = String(iso[3]).padStart(2,'0');

    // se coincide (ou OCR trocou D↔M), herda o ano do ISO
    if ((dd === isoD && mm === isoM) || (dd === isoM && mm === isoD)) {
      return isoY;
    }
    return null;
  }

  // 0) Normalização leve
  const injectSpaces = s => String(s)
    .replace(/\u00A0/g, ' ')
    .replace(/([A-Za-z])(\d)/g, '$1 $2')
    .replace(/(\d)([A-Za-z])/g, '$1 $2');
  const raw = injectSpaces(pdfText);

  // 1) Filtrar “linhas venenosas” (2ª via / cópias / reimpressão / gerada em …)
  function isHardBad(line) {
    const lo = line.toLowerCase();
    if (/\b2\s*(?:ª|a|\.ª)?\s*via\b/.test(lo)) return true;
    if (/\bsegunda\s+via\b/.test(lo)) return true;
    if (/\bduplicad[oa]\b/.test(lo)) return true;
    if (/\breimpress[ãa]o\b/.test(lo)) return true;
    if (/\bc[óo]pia\b/.test(lo)) return true;
    if (/\bvia\b.{0,20}\bgerad[ao]?\b/.test(lo)) return true; // "via gerada"
    if (/\bgerad[ao]\s+em\b/.test(lo)) return true;           // "gerada em 12-04-2025"
    if (/\ba\s+partir\s+d[eo]\b/.test(lo)) return true;
    return false;
  }

  const linesAll = raw.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const lines = linesAll.filter(l => !isHardBad(l));

  // 2) FAST PATH — cabeçalho documental: "FATURA/RECIBO|FATURA|NOTA DE CRÉDITO ... DATA: dd/mm/aaaa"
  //    (Protege contra a “2ª via gerada em …” que removemos acima)
  const headerRe = /\b(?:fatura\/recibo|fatura|factura|nota\s+de\s+cr[eé]dito)\b/i;
  const dataFieldRe = /\bdata\s*:\s*(\d{2})[./-](\d{2})[./-](\d{4})\b/i;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (headerRe.test(l)) {
      // tenta apanhar "DATA: ..." na própria linha OU nas 3 próximas linhas curtas
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

  // 3) Outra via rápida — ISO AAAA-MM-DD junto a labels bons (evita “vencimento”/intervalos)
  const goodLabels = [
    /data\s*(?:de)?\s*emiss[aã]o/i, // Data de Emissão
    /\bemiss[aã]o\b/i,              // emissão
    /\bemitida\b/i,                 // emitida
    /\bemitida\s+em\b/i,            // emitida em
    /dt\.?\s*emiss[aã]o/i,          // DT EMISSÃO
    /\bdata\s*doc(?:umento)?\b/i,   // Data Doc
    /\bdata\s*:\b/i,                // Data:
    /\binvoice\s+date\b/i,
    /\bissue\s+date\b/i,
    /\bfecha\s+de\s+emisi[oó]n\b/i
  ];
  const oneLine = raw.replace(/\s+/g, ' ');
  const isoNearLabel = new RegExp(
    '(?:' + goodLabels.map(r=>r.source).join('|') + ')' +
    '[^0-9]{0,40}(20\\d{2})[./-](\\d{2})[./-](\\d{2})','i'
  );
  let m = oneLine.match(isoNearLabel);
  if (m) {
    const fast = _safeDate_(m[3], m[2], m[1]);
    if (fast) return fast;
  }

  // 4) Scanner geral com exclusões
  const badLabels = [
    /\bvig[êe]ncia\b/i,
    /\bper[ií]odo\b/i,
    /\bvalidade\b/i,
    /\bcompet[êe]ncia\b/i,
    /\bref\.\s*(?:per[ií]odo|m[êe]s)\b/i,
    /\bintervalo\b/i,
    /\bvenc(?:imento)?\b/i,   // novo
    /\bprazo\b/i,             // novo
    /\bdue\b/i,               // novo
    /\bpayment\b/i            // novo
  ];
  const hasRangeRe = /\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*(?:a|–|—|-)\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/i;

  /* VERSÃO FUNCIONAL E MUITO BOA (só não apanha a data dos talões):
    const rx = [
      { re:/\b(\d{1,2})[\/](\d{1,2})[\/](\d{2,4})\b/g,                  norm:(d,m,y)=>({d,m,y}) },
      { re:/\b(\d{1,2})[-](\d{1,2})[-](\d{2,4})\b/g,                    norm:(d,m,y)=>({d,m,y}) },
      { re:/\b(\d{1,2})[.](\d{1,2})[.](\d{2,4})\b/g,                    norm:(d,m,y)=>({d,m,y}) },
      { re:/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/g,          norm:(y,m,d)=>({d,m,y}) },
      { re:/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})\b/gi,
        norm:(d,mon,y)=>({d, m:_monPT_(mon), y}) },
      { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})\b/gi,
        norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
      { re:/\b(\d{1,2})[.\- ](jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[.\- ](\d{4})\b/gi,
        norm:(d,mon,y)=>({d, m:_monPTabbrev_(mon), y}) }
    ];
  */

  const rx = [
    // ISO com ano de 4 dígitos (prioridade)
    { re:/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/g,
      norm:(y,m,d)=>({d,m,y, iso:true}) },

    // DD/MM/YYYY | DD-MM-YYYY | DD.MM.YYYY
    { re:/\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/g,
      norm:(d,m,y)=>({d,m,y}) },

    // DD/MM/YY | DD-MM-YY (AMBÍGUO → reconciliar)
    { re:/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\b/g,
      norm:(d,m,y)=>({d,m,y, ambiguousYY:true}) },

    // Meses por extenso
    { re:/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})\b/gi,
      norm:(d,mon,y)=>({d, m:_monPT_(mon), y}) },
    { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})\b/gi,
      norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
    { re:/\b(\d{1,2})[.\- ](jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[.\- ](\d{4})\b/gi,
      norm:(d,mon,y)=>({d, m:_monPTabbrev_(mon), y}) }
  ];


  const candidates = [];
  for (let i=0;i<lines.length;i++) {
    const line = lines[i];
    if (!line) continue;
    if (hasRangeRe.test(line)) continue;
    if (badLabels.some(r=>r.test(line))) continue;

    const hasGood = goodLabels.some(r=>r.test(line));
    for (const {re,norm} of rx) {
      re.lastIndex = 0;
      let mm;
      while ((mm = re.exec(line)) !== null) {

        /* VERSÃO FUNCIONAL E MUITO BOA (só não apanha a data dos talões):
        const p = norm(...mm.slice(1));
        const dd = to2(p.d), mth = to2(p.m), yyyy = normYear(p.y);
        const safe = _safeDate_(dd, mth, yyyy);
        if (!safe) continue;

        // proximidade label-bom (<=30 chars antes)
        const near = goodLabels.some(r=>{
          r.lastIndex=0;
          const lm = r.exec(line);
          return lm && lm.index>=0 && (mm.index - lm.index) <= 30;
        });

        // rejeitar datas coladas a termos de pagamento/vencimento
        const around = 30; // janela de proximidade
        const left  = Math.max(0, mm.index - around);
        const right = Math.min(line.length, mm.index + mm[0].length + around);
        const ctx = line.slice(left, right).toLowerCase();

        let score = 0;
        if (hasGood) score += 120;
        if (near)    score += 15;
        if (i < 12)  score += 30;
        if (line.length <= 80) score += 5;
        candidates.push({ safe, score, lineIndex:i, colIndex:mm.index });
        */

        const p = norm(...mm.slice(1));

        // normaliza dia/mes para 2 dígitos
        const dd2 = String(p.d).padStart(2,'0');
        const mm2 = String(p.m).padStart(2,'0');
        let yyyy = String(p.y);

        // se ano tem 2 dígitos → tenta reconciliar com ISO próximo na MESMA LINHA
        if (p.ambiguousYY || yyyy.length === 2) {
          const reconciled = _reconcileYearWithNearbyISO(line, mm.index, dd2, mm2, 60);
          if (reconciled) {
            yyyy = reconciled;
          } else {
            // fallback YY → YYYY (regra 00–79 => 2000+, 80–99 => 1900+)
            const n = +yyyy;
            yyyy = (n <= 79) ? (2000 + n) : (1900 + n);
          }
        }

        const safe = _safeDate_(dd2, mm2, yyyy);
        if (!safe) continue;

        // proximidade a labels boas
        const near = goodLabels.some(r=>{
          r.lastIndex=0;
          const lm = r.exec(line);
          return lm && lm.index>=0 && (mm.index - lm.index) <= 30;
        });

        let score = 0;
        if (p.iso) score += 40;      // bónus extra para ISO yyyy-mm-dd
        const hasGood = goodLabels.some(r=>r.test(line));
        if (hasGood) score += 120;
        if (near)    score += 15;
        if (i < 12)  score += 30;
        if (line.length <= 80) score += 5;

        candidates.push({ safe, score, lineIndex:i, colIndex:mm.index });

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

  // 5) fallback rápido no texto limpo de intervalos
  const cleaned = lines.filter(l=>!hasRangeRe.test(l)).join('\n');
  const quick = [
    { r:/\b(20\d{2})[./-](\d{2})[./-](\d{2})\b/g, p:s=>{const a=s.split(/[./-]/); return _safeDate_(a[2],a[1],a[0]);} },
    { r:/\b(\d{2})[./-](\d{2})[./-](\d{4})\b/g,   p:s=>{const a=s.split(/[./-]/); return _safeDate_(a[0],a[1],a[2]);} }
  ];
  for (const {r,p} of quick) {
    let q; r.lastIndex=0;
    while ((q=r.exec(cleaned))!==null) {
      const out = p(q[0]);
      if (out) return out;
    }
  }
  return null;

  // helpers
  function normYear(y){ y=String(y); if (y.length===2){ const n=+y; return (n<=79?2000+n:1900+n);} return y; }
  function to2(n){ return String(n).padStart(2,'0'); }
  function _monPT_(s){ const m={janeiro:1,fevereiro:2,'março':3,marco:3,abril:4,maio:5,junho:6,julho:7,agosto:8,setembro:9,outubro:10,novembro:11,dezembro:12}; return m[String(s).toLowerCase()]||s; }
  function _monPTabbrev_(s){ const m={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12}; return m[String(s).toLowerCase().slice(0,3)]||s; }
  function _monEN_(s){ const m={january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12}; return m[String(s).toLowerCase()]||s; }
}

//BACKUP: também é boa, a outra parece melhor:
function extractDataDocumento_Simplesv2(pdfText) {
  if (!pdfText) return null;

  // 0) Normalização leve
  const injectSpaces = s => String(s)
    .replace(/\u00A0/g, ' ')
    .replace(/([A-Za-z])(\d)/g, '$1 $2')
    .replace(/(\d)([A-Za-z])/g, '$1 $2');
  const raw = injectSpaces(pdfText);

  // 1) Filtrar “linhas venenosas” (2ª via / cópias / reimpressão / gerada em …)
  function isHardBad(line) {
    const lo = line.toLowerCase();
    if (/\b2\s*(?:ª|a|\.ª)?\s*via\b/.test(lo)) return true;
    if (/\bsegunda\s+via\b/.test(lo)) return true;
    if (/\bduplicad[oa]\b/.test(lo)) return true;
    if (/\breimpress[ãa]o\b/.test(lo)) return true;
    if (/\bc[óo]pia\b/.test(lo)) return true;
    if (/\bvia\b.{0,20}\bgerad[ao]?\b/.test(lo)) return true; // "via gerada"
    if (/\bgerad[ao]\s+em\b/.test(lo)) return true;           // "gerada em 12-04-2025"
    return false;
  }

  const linesAll = raw.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const lines = linesAll.filter(l => !isHardBad(l));

  // 2) FAST PATH — cabeçalho documental: "FATURA/RECIBO|FATURA|NOTA DE CRÉDITO ... DATA: dd/mm/aaaa"
  //    (Protege contra a “2ª via gerada em …” que removemos acima)
  const headerRe = /\b(?:fatura\/recibo|fatura|factura|nota\s+de\s+cr[eé]dito)\b/i;
  const dataFieldRe = /\bdata\s*:\s*(\d{2})[./-](\d{2})[./-](\d{4})\b/i;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (headerRe.test(l)) {
      // tenta apanhar "DATA: ..." na própria linha OU nas 3 próximas linhas curtas
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

  // 3) Outra via rápida — ISO AAAA-MM-DD junto a labels bons (evita “vencimento”/intervalos)
  const goodLabels = [
    /data\s*(?:de)?\s*emiss[aã]o/i, // Data de Emissão
    /\bemiss[aã]o\b/i,              // emissão
    /\bemitida\b/i,                 // emitida
    /\bemitida\s+em\b/i,            // emitida em
    /dt\.?\s*emiss[aã]o/i,          // DT EMISSÃO
    /\bdata\s*doc(?:umento)?\b/i,   // Data Doc
    /\bdata\s*:\b/i,                // Data:
    /\binvoice\s+date\b/i,
    /\bissue\s+date\b/i,
    /\bfecha\s+de\s+emisi[oó]n\b/i
  ];
  const oneLine = raw.replace(/\s+/g, ' ');
  const isoNearLabel = new RegExp(
    '(?:' + goodLabels.map(r=>r.source).join('|') + ')' +
    '[^0-9]{0,40}(20\\d{2})[./-](\\d{2})[./-](\\d{2})','i'
  );
  let m = oneLine.match(isoNearLabel);
  if (m) {
    const fast = _safeDate_(m[3], m[2], m[1]);
    if (fast) return fast;
  }

  // 4) Scanner geral com exclusões
  const badLabels = [
    /\bvig[êe]ncia\b/i,
    /\bper[ií]odo\b/i,
    /\bvalidade\b/i,
    /\bcompet[êe]ncia\b/i,
    /\bref\.\s*(?:per[ií]odo|m[êe]s)\b/i,
    /\bintervalo\b/i,
    /\bvenc(?:imento)?\b/i,
    /\bprazo\b/i,
    /\bdue\b/i,
    /\bpayment\b/i,
    /\bpago\b/i,
    /\brecebido\b/i
  ];
  const hasRangeRe = /\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*(?:a|–|—|-)\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/i;

  const rx = [
    { re:/\b(\d{1,2})[\/](\d{1,2})[\/](\d{2,4})\b/g,                  norm:(d,m,y)=>({d,m,y}) },
    { re:/\b(\d{1,2})[-](\d{1,2})[-](\d{2,4})\b/g,                    norm:(d,m,y)=>({d,m,y}) },
    { re:/\b(\d{1,2})[.](\d{1,2})[.](\d{2,4})\b/g,                    norm:(d,m,y)=>({d,m,y}) },
    { re:/\b(20\d{2})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\b/g,          norm:(y,m,d)=>({d,m,y}) },
    { re:/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})\b/gi,
      norm:(d,mon,y)=>({d, m:_monPT_(mon), y}) },
    { re:/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})\b/gi,
      norm:(mon,d,y)=>({d, m:_monEN_(mon), y}) },
    { re:/\b(\d{1,2})[.\- ](jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[.\- ](\d{4})\b/gi,
      norm:(d,mon,y)=>({d, m:_monPTabbrev_(mon), y}) }
  ];

  const candidates = [];
  for (let i=0;i<lines.length;i++) {
    const line = lines[i];
    if (!line) continue;
    if (hasRangeRe.test(line)) continue;
    if (badLabels.some(r=>r.test(line))) continue;

    const hasGood = goodLabels.some(r=>r.test(line));
    for (const {re,norm} of rx) {
      re.lastIndex = 0;
      let mm;
      while ((mm = re.exec(line)) !== null) {
        const p = norm(...mm.slice(1));
        const dd = to2(p.d), mth = to2(p.m), yyyy = normYear(p.y);
        const safe = _safeDate_(dd, mth, yyyy);
        if (!safe) continue;

        // proximidade label-bom (<=30 chars antes)
        const near = goodLabels.some(r=>{
          r.lastIndex=0;
          const lm = r.exec(line);
          return lm && lm.index>=0 && (mm.index - lm.index) <= 30;
        });

        let score = 0;
        if (hasGood) score += 120;
        if (near)    score += 15;
        if (i < 12)  score += 30;
        if (line.length <= 80) score += 5;
        candidates.push({ safe, score, lineIndex:i, colIndex:mm.index });
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

  // 5) fallback rápido no texto limpo de intervalos
  const cleaned = lines.filter(l=>!hasRangeRe.test(l)).join('\n');
  const quick = [
    { r:/\b(20\d{2})[./-](\d{2})[./-](\d{2})\b/g, p:s=>{const a=s.split(/[./-]/); return _safeDate_(a[2],a[1],a[0]);} },
    { r:/\b(\d{2})[./-](\d{2})[./-](\d{4})\b/g,   p:s=>{const a=s.split(/[./-]/); return _safeDate_(a[0],a[1],a[2]);} }
  ];
  for (const {r,p} of quick) {
    let q; r.lastIndex=0;
    while ((q=r.exec(cleaned))!==null) {
      const out = p(q[0]);
      if (out) return out;
    }
  }
  return null;

  // helpers
  function normYear(y){ y=String(y); if (y.length===2){ const n=+y; return (n<=79?2000+n:1900+n);} return y; }
  function to2(n){ return String(n).padStart(2,'0'); }
  function _monPT_(s){ const m={janeiro:1,fevereiro:2,'março':3,marco:3,abril:4,maio:5,junho:6,julho:7,agosto:8,setembro:9,outubro:10,novembro:11,dezembro:12}; return m[String(s).toLowerCase()]||s; }
  function _monPTabbrev_(s){ const m={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12}; return m[String(s).toLowerCase().slice(0,3)]||s; }
  function _monEN_(s){ const m={january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12}; return m[String(s).toLowerCase()]||s; }
}





function _safeDate_(dd, mm, yyyy) {
  const y = Number(yyyy), m = Number(mm), d = Number(dd);
  if (!y || !m || !d) return null;
  if (y < 2000) return null;
  const dt = new Date(y, m-1, d);
  const today = new Date();
  if (isNaN(dt.getTime()) || dt > today) return null;
  if (dt.getFullYear()!==y || (dt.getMonth()+1)!==m || dt.getDate()!==d) return null;
  return `${String(d).padStart(2,'0')}/${String(m).padStart(2,'0')}/${String(y)}`;
}

























/** Extrai data (prioriza labels), normaliza DD/MM/AAAA e rejeita datas futuras. */
function extractDataDocumento_Simples_PERMISSIVA(pdfText) {
  if (!pdfText) return null;

  // Helpers locais
  const norm = (s) => (s || "").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  const injectSpaces = (s) => (s || "").replace(/(\d)([A-Za-z])/g,'$1 $2').replace(/([A-Za-z])(\d)/g,'$1 $2');

  // Normalizações para apanhar casos como "2025-01-13Data:"
  const textSpaced = injectSpaces(pdfText);
  const one = textSpaced.replace(/\s+/g, ' ');

  // 1) Labels fortes (igual ao teu, mas com o texto já “spaced”)
  let m = one.match(/\b(data(?:\s+de\s+emiss[aã]o)?|data\s+doc(?:umento)?|invoice\s+date|issue\s+date)\s*[:\-]?\s*(\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4})/i);
  if (m) {
    const p = m[2].replace(/[.\-]/g,'/').split('/');
    return _safeDatePERMISSIVA_(p[0].padStart(2,'0'), p[1].padStart(2,'0'), p[2].length===2 ? '20'+p[2] : p[2]);
  }
  m = one.match(/\b(invoice\s+date|issue\s+date)\s*[:\-]?\s*(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2}),\s*(\d{4})/i);
  if (m) {
    const mapEN = {january:"01",february:"02",march:"03",april:"04",may:"05",june:"06",july:"07",august:"08",september:"09",october:"10",november:"11",december:"12"};
    return _safeDatePERMISSIVA_(String(m[3]).padStart(2,'0'), mapEN[m[2].toLowerCase()], m[4]);
  }
  m = one.match(/\b(data(?:\s+de\s+emiss[aã]o)?|data\s+doc(?:umento)?)\s*[:\-]?\s*(\d{1,2})\s+de\s+(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+(\d{4})/i);
  if (m) {
    const mapPT = {"janeiro":"01","fevereiro":"02","março":"03","abril":"04","maio":"05","junho":"06","julho":"07","agosto":"08","setembro":"09","outubro":"10","novembro":"11","dezembro":"12"};
    return _safeDatePERMISSIVA_(String(m[2]).padStart(2,'0'), mapPT[m[3].toLowerCase()], m[4]);
  }

  // 2) SCAN POR LINHAS PREFERENCIAIS (ANTES DO CLEAN/FALLBACK)
  //    - inclui "data:" sem "venc" próximo
  //    - ignora vencimento/validade/prazo/due/etc.
  const preferTerms = ['emissao','emitid','data doc','data documento','invoice date','issue date','documento:','fatura:','factura:','data:'];
  const excludeTerms = ['venc','vencimento','prazo','validade','due','payment','limite'];

  const dateFinders = [
    // DD/MM/AAAA ou DD-MM-AAAA
    (s) => {
      const x = s.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/);
      if (!x) return null;
      const yy = x[3].length===2 ? ('20'+x[3]) : x[3];
      return _safeDatePERMISSIVA_(String(x[1]).padStart(2,'0'), String(x[2]).padStart(2,'0'), yy);
    },
    // AAAA/MM/DD ou AAAA-MM-DD
    (s) => {
      const x = s.match(/\b(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\b/);
      if (!x) return null;
      return _safeDatePERMISSIVA_(String(x[3]).padStart(2,'0'), String(x[2]).padStart(2,'0'), x[1]);
    },
    // "13 de janeiro de 2025"
    (s) => {
      const x = s.match(/\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de?\s+(\d{4})/i);
      if (!x) return null;
      const mapPT = {"janeiro":"01","fevereiro":"02","março":"03","abril":"04","maio":"05","junho":"06","julho":"07","agosto":"08","setembro":"09","outubro":"10","novembro":"11","dezembro":"12"};
      return _safeDatePERMISSIVA_(String(x[1]).padStart(2,'0'), mapPT[x[2].toLowerCase()], x[3]);
    }
  ];

  const lines = textSpaced.split(/\r?\n/);
  for (let raw of lines) {
    const ln = raw.trim();
    if (!ln) continue;
    const lnNorm = norm(ln);

    // descartar linhas "más"
    if (excludeTerms.some(t => lnNorm.includes(t))) continue;

    // aceitar linhas "boas"
    if (preferTerms.some(t => lnNorm.includes(t))) {
      // também proteger "data:" quando houver "venc" colado algures
      if (lnNorm.includes('data:') && lnNorm.includes('venc')) continue;

      for (const finder of dateFinders) {
        const got = finder(ln);
        if (got) return got;
      }
    }
  }

  // 3) Limpa linhas de período/vigência para evitar falsos positivos no fallback
  const cleaned = textSpaced
    .split(/\r?\n/)
    .filter(l => {
      const lo = norm(l);
      if (excludeTerms.some(t => lo.includes(t))) return false; // ignora venc/validade/prazo/due
      if (/\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}\s*(a|\-|–|—)\s*\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}/i.test(lo)) return false; // intervalos
      return true;
    })
    .join('\n');

  // 4) Fallbacks gerais (os teus), já em texto limpo
  const pats = [
    { r: /(\d{2}[\/\-]\d{2}[\/\-]\d{4})/g, p: s => s.replace(/-/g,'/') },
    { r: /(\d{4}[\/\-]\d{2}[\/\-]\d{2})/g, p: s => { const a=s.replace(/-/g,'/').split('/'); return `${a[2]}/${a[1]}/${a[0]}`; } },
    { r: /(\d{1,2})\s+de\s+(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+(\d{4})/gi,
      p: s => { const x=s.match(/(\d{1,2})\s+de\s+([a-zç]+)\s+(\d{4})/i);
                const map={"janeiro":"01","fevereiro":"02","março":"03","abril":"04","maio":"05","junho":"06","julho":"07","agosto":"08","setembro":"09","outubro":"10","novembro":"11","dezembro":"12"};
                return `${String(x[1]).padStart(2,'0')}/${map[x[2].toLowerCase()]}/${x[3]}`; } }
  ];
  for (const {r,p} of pats) {
    const all = [...cleaned.matchAll(r)];
    for (const mm of all) {
      const [dd,mm_,yy] = p(mm[0]).split('/');
      const safe = _safeDatePERMISSIVA_(dd, mm_, yy);
      if (safe) return safe;
    }
  }

  return null;
}




function _safeDatePERMISSIVA_(dd, mm, yyyy) {
  const d = new Date(Number(yyyy), Number(mm)-1, Number(dd));
  const today = new Date();
  if (isNaN(d.getTime()) || d > today) return null;
  return `${String(dd).padStart(2,'0')}/${String(mm).padStart(2,'0')}/${String(yyyy)}`;
}



















/**
 * 
 *  CÓDIDO PARA CATALOGAR COMPROVATIVOS
 * 
 */

/**
 * Função principal
 * - percorre todos os PDFs na pasta de comprovativos
 * - extrai ATCUD e data de pagamento
 * - procura, por anos para trás, a fatura que tem esse ATCUD
 * - se encontrar, APENAS renomeia o comprovativo e faz log do que iria fazer na folha
 */
/**
 * COPIAR COMPROVATIVOS PARA O ARQUIVO (SEM CATALOGAR NAS SHEETS)
 *
 * - percorre todos os PDFs na pasta de comprovativos (PASTA_COMPROVATIVOS_ID)
 * - extrai ATCUD e data de pagamento
 * - encontra o mês/ano onde está a fatura
 * - cria uma CÓPIA renomeada para COMP<N>.pdf dentro de:
 *        #5 - Comprovativos de pagamento / PARA CATALOGAR
 * - move o ficheiro original para PASTA_LIXO
 */
function simularCatalogacaoComprovativosArquivo() {
  var pastaComprovativos = DriveApp.getFolderById(PASTA_COMPROVATIVOS_ID);
  var pdfFiles = getPDFFilesInFolder(pastaComprovativos);
  Logger.log("Encontrados " + pdfFiles.length + " comprovativos na pasta de entrada.");

  for (var i = 0; i < pdfFiles.length; i++) {
    var file = pdfFiles[i];
    Logger.log("=======================================");
    Logger.log("A processar comprovativo: " + file.getName());

    try {
      var texto = convertPDFToText(file.getId(), "pt");

      var atcud = extractATCUDFromText(texto);
      if (!atcud) {
        Logger.log("  -> ATCUD não encontrado no texto. Comprovativo ignorado.");
        continue;
      }
      Logger.log("  -> ATCUD extraído: " + atcud);

      var dataPagamentoStr = extractDateFromPayslip(texto);
      var anoPagamento = inferYearFromDateString(dataPagamentoStr);
      if (!anoPagamento) anoPagamento = new Date().getFullYear();

      var match = procurarFaturaPorATCUDNoArquivo(atcud, anoPagamento);

      if (match) {
        var novoNome = "COMP" + match.numeroDocumento + ".pdf";

        Logger.log(
          "  -> MATCH ENCONTRADO!\n" +
          "     Ano.: " + match.ano + "\n" +
          "     Pasta Faturas.: " + match.pastaMesNome + "\n" +
          "     Ficheiro.: " + match.spreadsheetName + "\n" +
          "     Aba.: " + match.sheetName + "\n" +
          "     Linha.: " + match.row + "\n" +
          "     Nº Documento.: " + match.numeroDocumento + "\n" +
          "     Novo nome da CÓPIA: " + novoNome
        );

        // 1) Descobrir a pasta do mês onde está o ficheiro de faturas
        var ssFile = DriveApp.getFileById(match.spreadsheetId);
        var parents = ssFile.getParents();
        if (!parents.hasNext()) {
          Logger.log("  -> AVISO: Ficheiro sem pasta-pai. Não foi possível localizar mês.");
          continue;
        }

        var pastaMes = parents.next(); // "Faturas_DL_MM/AAAA"

        // 2) Dentro dela, garantir #5 - Comprovativos de pagamento
        var itComp = pastaMes.getFoldersByName("#5 - Comprovativos de pagamento");
        var pasta5 = itComp.hasNext()
          ? itComp.next()
          : pastaMes.createFolder("#5 - Comprovativos de pagamento");

        // 3) Dentro de #5, garantir a subpasta "PARA CATALOGAR"
        var itPara = pasta5.getFoldersByName("PARA CATALOGAR");
        var pastaParaCatalogar = itPara.hasNext()
          ? itPara.next()
          : pasta5.createFolder("PARA CATALOGAR");

        // 4) Criar a cópia
        var copia = file.makeCopy(novoNome, pastaParaCatalogar);

        Logger.log(
          "  -> Cópia criada: '" + copia.getName() +
          "' em '#5 - Comprovativos de pagamento/ PARA CATALOGAR' de '" +
          pastaMes.getName() + "'."
        );

        // 5) Mover o original para PASTA_LIXO
        try {
          var pastaLixo = DriveApp.getFolderById(PASTA_LIXO);
          file.moveTo(pastaLixo);
          Logger.log("  -> Ficheiro original movido para PASTA_LIXO.");
        } catch (eLixo) {
          Logger.log("  -> ERRO ao mover o original para PASTA_LIXO: " + eLixo);
        }

      } else {
        Logger.log("  -> NENHUM MATCH encontrado para ATCUD '" + atcud + "'.");
      }

    } catch (e) {
      Logger.log("  -> ERRO ao processar '" + file.getName() + "': " + e);
    }
  }

  Logger.log("Fim da execução: cópia de comprovativos (sem catalogar).");
}



/**
 * Interpreta uma string de data (idealmente "DD/MM/AAAA") e devolve o ano (number)
 */
function inferYearFromDateString(dateStr) {
  if (!dateStr) return null;
  var parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    var ano = parseInt(parts[2], 10);
    if (!isNaN(ano) && ano > 1900 && ano < 2100) {
      return ano;
    }
  }
  // fallback: tenta apanhar um grupo de 4 dígitos
  var m = dateStr.match(/(19\d{2}|20\d{2})/);
  if (m) {
    var ano2 = parseInt(m[1], 10);
    if (!isNaN(ano2)) return ano2;
  }
  return null;
}

/**
 * Procura uma fatura com determinado ATCUD nas várias folhas de faturas
 * - começa no ano do pagamento
 * - se não encontrar, vai para anos anteriores existentes
 * Retorna:
 *   {
 *     ano: number,
 *     spreadsheetId: string,
 *     spreadsheetName: string,
 *     sheetName: string,
 *     row: number,
 *     numeroDocumento: string|number
 *   }
 * ou null se não encontrar.
 */
function procurarFaturaPorATCUDNoArquivo(atcud, anoPagamento) {
  if (!atcud) return null;

  var atcudNormalizado = String(atcud).replace(/\s/g, "");
  var yearFolders = getYearFolders(); // devolve array de pastas de ano (nomes "2023", "2024", ...)

  if (!yearFolders || !yearFolders.length) return null;

  // Construir lista de anos candidatos: pastas cujo ano <= anoPagamento
  var candidatos = [];
  for (var i = 0; i < yearFolders.length; i++) {
    var pastaAno = yearFolders[i];
    var nomeAno = pastaAno.getName().trim();
    var anoNum = parseInt(nomeAno, 10);

    if (!isNaN(anoNum) && anoNum <= anoPagamento) {
      candidatos.push(pastaAno);
    }
  }

  // Se não houver nenhuma <= anoPagamento, tentamos todos os anos disponíveis
  if (!candidatos.length) {
    candidatos = yearFolders.slice();
  }

  // Percorre anos candidatos (já vêm ordenados por getYearFolders)
  for (var a = 0; a < candidatos.length; a++) {
    var pastaAno = candidatos[a];
    var anoStr = pastaAno.getName().trim();
    var ano = parseInt(anoStr, 10);

    if (isNaN(ano)) continue;

    Logger.log("  -> A procurar ATCUD " + atcudNormalizado + " no ano " + ano + "...");

    var itMeses = pastaAno.getFolders();
    while (itMeses.hasNext()) {
      var pastaMes = itMeses.next();
      var nomePastaMes = pastaMes.getName();

      // Consideramos apenas pastas do tipo "Faturas_DL_mês/ano"
      if (nomePastaMes.indexOf("Faturas_DL_") !== 0) {
        continue;
      }

      Logger.log("     > Pasta mês: " + nomePastaMes);

      var files = pastaMes.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

        var ssId = f.getId();
        var ssName = f.getName();

        Logger.log("       - A abrir ficheiro de faturas: " + ssName);

        var ss = SpreadsheetApp.openById(ssId);
        var match = procurarATCUDNasAbasDeFaturas(ss, atcudNormalizado);

        if (match) {
          // Completar info de contexto e devolver logo
          match.ano = ano;
          match.spreadsheetId = ssId;
          match.spreadsheetName = ssName;
          match.pastaMesNome = nomePastaMes;
          return match;
        }
      }
    }
  }

  // Se chegou aqui, não encontrou nada
  return null;
}


/**
 * Procura o ATCUD normalizado nas 3 abas de faturas de um ficheiro:
 *  - "Faturas e NCs normais"
 *  - "Faturas e NCs com reembolso"
 *  - "Outros documentos"
 * Usa:
 *  - Coluna "ATCUD / Nº Documento"
 *  - Coluna "Comprovativo de pagamento" (tem de estar vazia)
 *  - Coluna "Nº" ou "Número do documento" para obter o número do documento
 */
function procurarATCUDNasAbasDeFaturas(ss, atcudNormalizado) {
  var nomesAbas = [
    "Faturas e NCs normais",
    "Faturas e NCs com reembolso",
    "Outros documentos"
  ];
  var linhaCabecalho = 2;

  for (var i = 0; i < nomesAbas.length; i++) {
    var nomeAba = nomesAbas[i];
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) continue;

    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha <= linhaCabecalho) continue;

    var colATCUD = encontraColunaNoCabecalho(sheet, "ATCUD / Nº Documento", linhaCabecalho);
    var colComp = encontraColunaNoCabecalho(sheet, "Comprovativo de pagamento", linhaCabecalho);

    // Números de documento podem ter cabeçalho "Nº" ou "Número do documento"
    var colNumDoc = encontraColunaNoCabecalho(sheet, "Nº", linhaCabecalho);
    if (colNumDoc < 0) {
      colNumDoc = encontraColunaNoCabecalho(sheet, "Número do documento", linhaCabecalho);
    }

    if (colATCUD < 0 || colComp < 0 || colNumDoc < 0) {
      continue;
    }

    for (var row = linhaCabecalho + 1; row <= ultimaLinha; row++) {
      var atcudLinha = sheet.getRange(row, colATCUD).getDisplayValue();
      if (!atcudLinha) continue;

      var atcudLinhaNorm = String(atcudLinha).replace(/\s/g, "");
      if (atcudLinhaNorm !== atcudNormalizado) continue;

      var rangeComp = sheet.getRange(row, colComp);
      if (!rangeComp.isBlank()) {
        // já tem comprovativo, ignoramos
        continue;
      }

      var numeroDocumento = sheet.getRange(row, colNumDoc).getValue();

      return {
        sheetName: nomeAba,
        row: row,
        numeroDocumento: numeroDocumento
      };
    }
  }

  return null;
}

// ==============================
// Helpers reutilizados da Bernardete
// ==============================

function getPDFFilesInFolder(folder) {
  var files = folder.getFiles();
  var pdfFiles = [];
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === "application/pdf") {
      pdfFiles.push(file);
    }
  }
  return pdfFiles;
}


/**
 * Extrai ATCUD do texto de um PDF.
 * Versão simples reutilizando o que já estava a funcionar.
 */
function extractATCUDFromText(pdfText) {
  if (!pdfText) return null;

  // Primeiro tenta padrão "ATCUD: XXXXX"
  var regex1 = /ATCUD:\s*([^\s]+)/i;
  var match1 = pdfText.match(regex1);
  if (match1 && match1[1]) {
    return match1[1].trim();
  }

  // Depois tenta "ATCUD XXXXX"
  var regex2 = /ATCUD\s+([^\s]+)/i;
  var match2 = pdfText.match(regex2);
  if (match2 && match2[1]) {
    return match2[1].trim();
  }

  // Se nada, tenta apanhar tokens tipo ABCD23EF-12345
  var text = pdfText.replace(/\s+/g, " ").toUpperCase();
  var regex3 = /\b([A-Z0-9]{8,}-\d{2,})\b/;
  var match3 = text.match(regex3);
  if (match3 && match3[1]) {
    return match3[1].trim();
  }

  return null;
}

/**
 * Versão simplificada da extractDateFromPayslip:
 * tenta encontrar a primeira data plausível no texto e devolve em "DD/MM/AAAA"
 */
function extractDateFromPayslip(content) {
  if (!content) return null;
  var words = content.split(/\s+/);
  for (var i = 0; i < words.length; i++) {
    var word = words[i];

    // AAAA-MM-DD ou AAAA/MM/DD
    if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(word)) {
      var parts1 = word.split(/[-/]/);
      return parts1[2] + "/" + parts1[1] + "/" + parts1[0];
    }

    // DD-MM-AAAA ou DD/MM/AAAA
    if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(word)) {
      var parts2 = word.split(/[-/]/);
      return parts2[0] + "/" + parts2[1] + "/" + parts2[2];
    }
  }

  // fallback: tenta padrão dd/mm/aaaa algures no texto completo
  var m = content.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) {
    var d = ("0" + m[1]).slice(-2);
    var mth = ("0" + m[2]).slice(-2);
    var y = m[3];
    return d + "/" + mth + "/" + y;
  }

  return null;
}

// ==============================
// Helpers de navegação de pastas / anos
// ==============================

/**
 * Devolve array de objetos {year, folder} para cada pasta de ano em PASTA_COMPROVATIVOS_ID
 * Ex.: nome da pasta "2024", "2025", etc.
 */
function getYearFolders() {
  if (__cacheYearFolders) {
    return __cacheYearFolders;
  }

  var root = DriveApp.getFolderById(PASTA_GERAL_FATURAS);
  var it = root.getFolders();
  var result = [];

  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (/^\d{4}$/.test(name)) {
      result.push({
        year: parseInt(name, 10),
        folder: f
      });
    }
  }

  // anos por ordem decrescente (mais recente primeiro)
  result.sort(function(a, b) {
    return b.year - a.year;
  });

  __cacheYearFolders = result;
  return result;
}


/**
 * Procura uma coluna pelo nome exato na linha de cabeçalho.
 * Devolve o número da coluna (1-based) ou -1 se não encontrar.
 */
function encontraColunaNoCabecalho(sheet, columnName, linhaDoCabecalho) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return -1;

  var headerRowValues = sheet
    .getRange(linhaDoCabecalho, 1, 1, lastColumn)
    .getValues()[0];

  for (var i = 0; i < headerRowValues.length; i++) {
    if (headerRowValues[i] === columnName) {
      return i + 1;
    }
  }

  return -1;
}