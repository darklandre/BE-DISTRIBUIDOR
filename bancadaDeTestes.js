/*
          ATENÇÃO!!!

    5. O Veredicto sobre a Bancada de Testes
Fizeste bem em manter. No entanto, a bancada usa funções de extração de data (extractDataDocumentoTaloes, extractDataDocumento_Simples_PERMISSIVA) que não são usadas na função principal distribuirFicheirosDoGeral (que usa extractDataDocumento ou extractDataDocumentoTaloes dependendo do sítio).

Regra de Ouro: A Bancada de Testes deve testar as mesmas funções que o Distribuidor usa. Caso contrário, estás a testar uma realidade paralela.

*/

// =================================================================
// FUNÇÕES DE BANCADA DE TESTES OTIMIZADAS
// =================================================================

/**
 * **************** BANCADA DE TESTES: OTIMIZADA ****************
 * * Testa faturas, reduzindo I/O e chamadas Utilities.sleep()
 * para máxima performance dentro do limite de 5 minutos do Apps Script.
 * 
 **/

function bancadaDeTestes2025() {

  // === MODO DE EXECUÇÃO ===
  // Se estiver vazio => percorre toda a árvore do ano (Faturas_DL_MM/AAAA -> #1 e #2).
  // Se tiver um ID => processa APENAS essa pasta.
  const ONLY_FOLDER_ID = "";//"102LMZvC1914IhB90_zJRvobOo5t-mv9G"; 

  // (opcionais) override do esperado quando usas ONLY_FOLDER_ID
  const ONLY_EXPECTED_MM   = "01";   // "01".."12" ou "" para NA
  const ONLY_EXPECTED_YEAR = "2025"; //2025; // usado só se ONLY_EXPECTED_MM tiver valor

  function _isOCRRateLimitError_(err) {
    return String(err).indexOf('User rate limit exceeded for OCR') >= 0;
  }


  // === CONTROLO DE RETOMA =====================

  const RESET_PROGRESS = false; // <-- muda para true para recomeçar do zero

  if (RESET_PROGRESS) {
    benchDatasPDF_Reset(); // Usa a função de reset unificada
    Logger.log("🔄 Progresso apagado → recomeça do zero na próxima execução.");
    // Não sai — deixa continuar, agora vazio
  }


  // ======= CONFIGURAÇÃO =======
  const YEAR = 2025;
  const PARENT_FOLDER_ID = "1z3HZIF1EoaCbQyAPSU2XNVnouFc62x3M";

  // ======= TEMPO / RETOMA =======
  const TIME_BUDGET_MS = 5 * 60 * 1000;
  // const SLEEP_MS = 120; // REMOVIDO PARA MÁXIMA VELOCIDADE
  const start = Date.now();

  // ======= PREPARA SHEET =======
  const { ss, sheet } = _bench_createNewSheet_();
  _bench_writeHeaderIfEmpty_(sheet);
  _bench_ensureFormatting_(sheet);

  // ======= ALVO DINÂMICO + RETOMA =======
  const props = PropertiesService.getScriptProperties();
  const IDX_KEY = "BENCH_IDX";
  const folders = ONLY_FOLDER_ID
  ? [[ONLY_EXPECTED_MM || "", ONLY_FOLDER_ID]]           // [[tag, folderId]]
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
    // Grava propriedades APENAS no checkpoint
    if (TOKEN_KEY) props.setProperty(TOKEN_KEY, pageToken || "");
    if (POS_KEY != null) props.setProperty(POS_KEY, String(pos || 0));
    props.setProperty(IDX_KEY, String(idx));
    Logger.log("[%s] ⚠️ %s — retoma guardada. Continua em ficheiro %s.", tag, reason || "Checkpoint", String(pos || 0));
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
    const POS_KEY   = "BENCH_POS_"   + folderId;

    let pageToken = props.getProperty(TOKEN_KEY) || null;

    const q = `'${folderId}' in parents and trashed=false`;
    const fields = "nextPageToken,items(id,title,mimeType,fileSize)";
    const MAX_RESULTS = 50;

    while (true) {
      if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) {
        // Usa pos=0 aqui (o pos correto será determinado mais tarde)
        if (_flushAndCheckpointAndReturn_(sheet, [], {TOKEN_KEY, pageToken, POS_KEY:null, pos:0, idx, tag, reason:"Time budget (antes da lista)"})) return;
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
      let current_pos = pos; // Rastreador de posição local para a página

      const rows = [];

      // usa loop indexado para respeitar 'pos'
      for (let i = pos; i < items.length; i++) {
        const it = items[i];
        current_pos = i; // Guarda a posição que vai ser processada/deixada

        // time budget check dentro do batch
        if ((Date.now() - start) > (TIME_BUDGET_MS - 8000)) {
          // Checkpoint: grava as rows recolhidas e a posição atual (i)
          if (_flushAndCheckpointAndReturn_(sheet, rows, {TOKEN_KEY, pageToken, POS_KEY, pos: i, idx, tag, reason:"Time budget (no batch)"})) return;
        }

        const mime = (it.mimeType || "").toLowerCase();
        const title = it.title || "";
        if (!mime.includes("pdf") && !/\.pdf$/i.test(title)) {
          // Arquivo não-PDF, apenas avança o rastreador
          current_pos = i + 1;
          continue;
        }

        let data = null, origem = "", nota = "", dataIA = "";
        const MAX_ATTEMPTS = 1;

        for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
          try {
            const texto = convertPDFToText(it.id); 

            // Tenta 1: Processamento CASO ESPECIAL RADIUS (tal como acontece na realidade)
            if(texto.includes("509001319") || texto.includes("Radius") || texto.includes("radius")){
              
              Logger.log("🛡️ RADIUS DETETADA: A calcular a menor data (Data Fatura).");

              // Regex para capturar qualquer data DD-MM-AAAA ou DD/MM/AAAA
              // NOTA: Usa (\d{1,2}) para dia/mês, mais robusto.
              var regexDatas = /(\d{1,2})[-./](\d{1,2})[-./](\d{4})/g; 
              var todasDatas = [];
              var match;

              while ((match = regexDatas.exec(texto)) !== null) {
                var d = parseInt(match[1], 10);
                var m = parseInt(match[2], 10);
                var y = parseInt(match[3], 10);
                
                // Filtro de segurança: datas válidas e dentro de um período razoável
                if(y >= 2020 && y <= 2030 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
                  var dateObj = new Date(y, m - 1, d);
                  // Validação extra para datas inválidas (e.g. 30/02)
                  if(dateObj.getFullYear() === y && dateObj.getMonth() === m - 1 && dateObj.getDate() === d) {
                    todasDatas.push({
                      str: match[0].replace(/-/g, "/").replace(/\./g, "/"),
                      obj: dateObj.getTime()
                    });
                  }
                }
              }

              if (todasDatas.length > 0) {
                // Ordena cronologicamente: Data Mais Antiga -> Data Mais Recente
                todasDatas.sort(function(a, b) { return a.obj - b.obj; });

                var menorData = todasDatas[0].str;
                
                // ** CORREÇÃO CHAVE: Atribuir a data e definir a origem **
                data = menorData;
                origem = "RADIUS"; 
                // FIM DA CORREÇÃO
                
                Logger.log("✅ Data RADIUS detetada: " + menorData);

              }
            }
            // Tenta 2: Lógica Normal/Permissiva 
            else {
              data = extractDataDocumentoTaloes(texto);
              origem = data ? "SIMPLES" : "";

              if (!data) {
                data = extractDataDocumento_Simples_PERMISSIVA(texto);
                origem = data ? "PERMISSIVA" : "";
              }
            }
            // ------------------------------------------------------------------------

            // Chama a IA independentemente do método anterior
            try {
              dataIA = extrairDataViaIA(texto.substring(0, 4000));
              Utilities.sleep(1500); // ← aumenta para 1.5s entre chamadas
            } catch(eIA) {
              dataIA = "ERRO_IA: " + String(eIA).substring(0, 80);
            }

            break; // sucesso

          } catch (e) {
            if (_isOCRRateLimitError_(e)) {
              if (attempt < MAX_ATTEMPTS) {
                const delay = 500 * Math.pow(2, attempt - 1); // Backoff exponencial
                Utilities.sleep(delay); // **Sleep necessário para o Backoff do Rate Limit**
                continue;
              } else {
                // Flush + checkpoint + sair
                _flushAndCheckpointAndReturn_(sheet, rows, {TOKEN_KEY, pageToken, POS_KEY, pos: i, idx, tag, reason:`OCR rate limit após ${MAX_ATTEMPTS} tentativas`});
                return;
              }
            } else {
              origem = "ERRO";
              nota = String(e);
              break;
            }
          }
        } // for attempt

        if (!data && !origem) {
          origem = "ERRO";
          nota = nota || "Sem data extraída";
        }

        const expected = (ONLY_FOLDER_ID && ONLY_EXPECTED_MM)
          ? `${ONLY_EXPECTED_MM}/${ONLY_EXPECTED_YEAR}`
          : (tag ? `${tag}/${YEAR}` : "");
          
        const verdict = _bench_compare_(data, expected);

        rows.push([
          new Date(),            // Timestamp (escrita)
          tag,                   // Tag (MM)
          folderName,            // Pasta
          title,                 // Ficheiro
          data || "",            // Data Detectada (DD/MM/AAAA)
          expected || "",        // Esperado (MM/AAAA)
          verdict,               // Veredicto
          origem || "",          // Origem
          `=HYPERLINK("https://drive.google.com/file/d/${it.id}/view","${it.id}")`,
          nota,
          dataIA  // ← coluna nova
        ]);

        current_pos = i + 1; // Avança o rastreador local
        
        // ** Utilities.sleep(SLEEP_MS); ** REMOVIDO PARA MÁXIMA VELOCIDADE
      } // for i

      // flush final da página (Escrita no Sheet)
      if (rows.length) {
        const shRange = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length);
        shRange.setValues(rows);
        // Persiste a POSIÇÃO no Drive.PropertiesService, caso o timeout ocorra
        // antes de avançarmos para a próxima página ou pasta.
        if (current_pos > pos) {
          props.setProperty(POS_KEY, String(current_pos));
        }
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

// =================================================================
// FUNÇÕES AUXILIARES DE SUPORTE
// =================================================================

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
  sheet.getRange(1,1,1,11).setValues([[
    "Timestamp (DD/MM/AAAA HH:MM:SS)",
    "Tag","Pasta","Ficheiro",
    "Data Detectada (DD/MM/AAAA)", // isto é a da fatura, mantém-se
    "Esperado (MM/AAAA)","Veredicto","Origem","File ID","Nota","Data IA (Groq)"
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

        let data = null, origem = "", nota = "", dataIA = "";

        try {
          const texto = convertPDFToText(it.id, "pt");
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
          nota,
          dataIA  // ← coluna nova
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

/**
 * ========================================================================
 * BANCADA DE TESTES DO CONSENSO
 * ========================================================================
 * Percorre uma pasta (ou a #0), corre o consenso em cada PDF,
 * e escreve os resultados numa spreadsheet nova para análise.
 *
 * Colunas: Ficheiro | Regex | Mistral | Groq | Gemini 2.0 Flash | Gemini Lite | Nome | CONSENSO | Votos | Link
 *
 * Configuração:
 *   FOLDER_ID  → pasta a testar (vazio = usa #0)
 *   MAX_FILES  → limite de ficheiros (0 = todos)
 */
function testarConsensoEmMassa() {
  // === CONFIGURAÇÃO ===
  var FOLDER_ID = ""; // Vazio = usa PASTA_GERAL_FICHEIROS (#0)
  var MAX_FILES = 0;  // 0 = sem limite

  var folderId = FOLDER_ID || PASTA_GERAL_FICHEIROS;
  var folder = DriveApp.getFolderById(folderId);
  Logger.log("📂 Pasta: " + folder.getName() + " (" + folderId + ")");

  // Criar spreadsheet de resultados
  var tz = "Europe/Lisbon";
  var ssName = "Consenso Teste - " + Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");
  var ss = SpreadsheetApp.create(ssName);
  ss.setSpreadsheetTimeZone(tz);
  var sheet = ss.getSheets()[0];
  sheet.setName("Resultados");

  // Cabeçalho
  sheet.getRange(1, 1, 1, 11).setValues([[
    "Ficheiro", "Regex", "Mistral", "Groq", "Gemini 2.0 Flash", "Gemini Lite", "Nome (MM/YYYY)",
    "CONSENSO", "Votos", "Fontes", "Link"
  ]]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#4a148c").setFontColor("#ffffff");

  Logger.log("📊 Sheet: " + ss.getUrl());

  // Iterar PDFs
  var files = folder.getFiles();
  var count = 0;
  var TIME_BUDGET = 5 * 60 * 1000 - 15000; // 5 min - margem
  var start = Date.now();

  while (files.hasNext()) {
    if (MAX_FILES > 0 && count >= MAX_FILES) break;
    if ((Date.now() - start) > TIME_BUDGET) {
      Logger.log("⏱️ Time budget atingido após " + count + " ficheiros.");
      break;
    }

    var file = files.next();
    if (file.getMimeType() !== "application/pdf") continue;

    var fileName = file.getName();
    count++;
    Logger.log("[" + count + "] " + fileName);

    var textoPDF = "";
    try {
      textoPDF = convertPDFToText(file.getId(), ['pt', 'en', null]) || "";
    } catch (e) {
      sheet.appendRow([fileName, "ERRO OCR", "", "", "", "", "", "", "", String(e).substring(0, 80), ""]);
      continue;
    }

    if (!textoPDF.trim()) {
      sheet.appendRow([fileName, "PDF vazio", "", "", "", "", "", "", "", "", ""]);
      continue;
    }

    // Correr consenso detalhado (com resultados individuais)
    var textoParaIA = textoPDF.substring(0, 4000);
    var prompt = "Extraia apenas a data de emissão do seguinte texto de um documento.\n" +
      "Retorne SOMENTE a data no formato DD/MM/AAAA, sem mais nada.\n" +
      'Se não encontrar, retorne "Não encontrada".\n\nTexto:\n' + textoParaIA;

    var dataRegex = extractDataDocumentoTaloes(textoPDF) || "";
    var dataMistral = "", dataGroq = "", dataGeminiPro = "", dataGeminiLite = "";

    try { dataMistral = _normalizarDataIA(chamarMistral(prompt)) || ""; } catch (e) { dataMistral = "ERRO"; }
    try { dataGroq = _normalizarDataIA(chamarGroq(prompt)) || ""; } catch (e) { dataGroq = "ERRO"; }
    try { dataGeminiPro = _normalizarDataIA(chamarGemini(prompt, "gemini-2.0-flash")) || ""; } catch (e) { dataGeminiPro = "ERRO"; }
    try { dataGeminiLite = _normalizarDataIA(chamarGemini(prompt, "gemini-3.1-flash-lite-preview")) || ""; } catch (e) { dataGeminiLite = "ERRO"; }

    var nomeInfo = _extrairDataDoNomeFicheiro(fileName);
    var nomeStr = nomeInfo ? (nomeInfo.month + "/" + nomeInfo.year) : "";

    // Consenso completo (reutiliza o resultado)
    var consenso = _consensoData(fileName, textoPDF);

    var link = '=HYPERLINK("https://drive.google.com/file/d/' + file.getId() + '/view","abrir")';

    sheet.appendRow([
      fileName,
      dataRegex,
      dataMistral,
      dataGroq,
      dataGeminiPro,
      dataGeminiLite,
      nomeStr,
      consenso.data || "SEM DATA",
      consenso.votos + "/6",
      consenso.fontes.join(", "),
      link
    ]);
  }

  // Formatação condicional: colorir consenso
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var colConsenso = 8;
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("SEM DATA").setBackground("#c62828").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, colConsenso, lastRow)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("/").setBackground("#2e7d32").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, colConsenso, lastRow)]).build()
  ]);

  // Auto-resize
  for (var c = 1; c <= 11; c++) sheet.autoResizeColumn(c);

  Logger.log("✅ Teste concluído: " + count + " ficheiros processados.");
  Logger.log("📊 Resultados: " + ss.getUrl());
}

/**
 * ========================================================================
 * BANCADA DE CONSENSO 2025 — Percorre TODOS os PDFs do ano
 * ========================================================================
 * Para cada PDF em #1 e #2 de cada mês, corre o consenso de 6 fontes
 * e compara o resultado com o mês esperado (pasta onde está).
 *
 * Suporta retoma automática (ScriptProperties) — se exceder o time budget
 * de 5 minutos, grava o progresso e continua na próxima execução.
 *
 * Para recomeçar do zero: corre consensoReset() primeiro.
 *
 * Resultados escritos numa spreadsheet com veredicto OK/FAIL.
 */
function testarConsenso2025() {
  // === CONFIGURAÇÃO ===
  var YEAR = 2025;
  var PARENT_FOLDER_ID = "1z3HZIF1EoaCbQyAPSU2XNVnouFc62x3M";
  var TIME_BUDGET_MS = 5 * 60 * 1000 - 15000; // 5 min - 15s margem
  var start = Date.now();

  var props = PropertiesService.getScriptProperties();
  var IDX_KEY = "CONS_IDX";
  var TOKEN_KEY_PREFIX = "CONS_TOKEN_";
  var POS_KEY_PREFIX = "CONS_POS_";
  var SS_KEY = "CONS_SHEET_ID";

  // === SHEET DE RESULTADOS (reutiliza entre execuções) ===
  var ssId = props.getProperty(SS_KEY);
  var ss, sheet;
  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
      sheet = ss.getSheetByName("Consenso");
    } catch (e) { ssId = null; }
  }
  if (!ssId) {
    var tz = "Europe/Lisbon";
    var ssName = "Consenso 2025 - " + Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");
    ss = SpreadsheetApp.create(ssName);
    ss.setSpreadsheetTimeZone(tz);
    sheet = ss.getSheets()[0];
    sheet.setName("Consenso");
    sheet.getRange(1, 1, 1, 13).setValues([[
      "Mês", "Pasta", "Ficheiro",
      "Regex", "Mistral", "Groq", "Gemini 2.0 Flash", "Gemini Lite", "Nome",
      "CONSENSO", "Esperado", "Veredicto", "Link"
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#4a148c").setFontColor("#ffffff");
    props.setProperty(SS_KEY, ss.getId());
  }
  Logger.log("📊 Sheet: " + ss.getUrl());

  // === ALVOS (reutiliza _bench_getTargetsFromParent_ da bancada existente) ===
  var folders = _bench_getTargetsFromParent_(PARENT_FOLDER_ID, YEAR);

  var idx = Number(props.getProperty(IDX_KEY) || 0);
  if (idx < 0 || idx >= folders.length) idx = 0;

  // === LOOP POR PASTAS ===
  for (; idx < folders.length; idx++) {
    if ((Date.now() - start) > TIME_BUDGET_MS) {
      props.setProperty(IDX_KEY, String(idx));
      Logger.log("⏱️ Time budget — retoma em idx=" + idx);
      return;
    }

    var mm = folders[idx][0];
    var folderId = folders[idx][1];
    var folder;
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { continue; }
    var folderName = folder.getName();
    var esperado = mm + "/" + YEAR;
    Logger.log("[" + mm + "] " + folderName);

    // Paginação Drive
    var TOKEN_KEY = TOKEN_KEY_PREFIX + folderId;
    var POS_KEY = POS_KEY_PREFIX + folderId;
    var pageToken = props.getProperty(TOKEN_KEY) || null;

    var q = "'" + folderId + "' in parents and trashed=false";
    var fields = "nextPageToken,items(id,title,mimeType)";

    while (true) {
      if ((Date.now() - start) > TIME_BUDGET_MS) {
        props.setProperty(TOKEN_KEY, pageToken || "");
        props.setProperty(IDX_KEY, String(idx));
        Logger.log("⏱️ Time budget (página) — retoma em idx=" + idx);
        return;
      }

      var resp = Drive.Files.list({ q: q, pageToken: pageToken, maxResults: 50, fields: fields, orderBy: "title" });
      var items = resp.items || [];
      if (!items.length) { props.deleteProperty(TOKEN_KEY); props.deleteProperty(POS_KEY); break; }

      var pos = Number(props.getProperty(POS_KEY) || 0);
      if (pos < 0 || pos > items.length) pos = 0;
      var rows = [];

      for (var i = pos; i < items.length; i++) {
        if ((Date.now() - start) > TIME_BUDGET_MS) {
          // Flush + checkpoint
          if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
          props.setProperty(POS_KEY, String(i));
          props.setProperty(TOKEN_KEY, pageToken || "");
          props.setProperty(IDX_KEY, String(idx));
          Logger.log("⏱️ Time budget (ficheiro " + i + ") — retoma guardada.");
          return;
        }

        var it = items[i];
        var mime = (it.mimeType || "").toLowerCase();
        var title = it.title || "";
        if (!mime.includes("pdf") && !/\.pdf$/i.test(title)) continue;

        // OCR
        var textoPDF = "";
        try {
          textoPDF = convertPDFToText(it.id, ['pt', 'en', null]) || "";
        } catch (e) {
          rows.push([mm, folderName, title, "ERRO OCR", "", "", "", "", "", "", esperado, "ERRO", ""]);
          continue;
        }
        if (!textoPDF.trim()) {
          rows.push([mm, folderName, title, "PDF vazio", "", "", "", "", "", "", esperado, "ERRO", ""]);
          continue;
        }

        // Resultados individuais
        var textoParaIA = textoPDF.substring(0, 4000);
        var prompt = "Extraia apenas a data de emissão do seguinte texto de um documento.\n" +
          "Retorne SOMENTE a data no formato DD/MM/AAAA, sem mais nada.\n" +
          'Se não encontrar, retorne "Não encontrada".\n\nTexto:\n' + textoParaIA;

        var dataRegex = extractDataDocumentoTaloes(textoPDF) || "";
        var dataMistral = "", dataGroq = "", dataGeminiPro = "", dataGeminiLite = "";

        try { dataMistral = _normalizarDataIA(chamarMistral(prompt)) || ""; } catch (e) { dataMistral = "ERRO"; }
        try { dataGroq = _normalizarDataIA(chamarGroq(prompt)) || ""; } catch (e) { dataGroq = "ERRO"; }
        try { dataGeminiPro = _normalizarDataIA(chamarGemini(prompt, "gemini-2.0-flash")) || ""; } catch (e) { dataGeminiPro = "ERRO"; }
        try { dataGeminiLite = _normalizarDataIA(chamarGemini(prompt, "gemini-3.1-flash-lite-preview")) || ""; } catch (e) { dataGeminiLite = "ERRO"; }

        var nomeInfo = _extrairDataDoNomeFicheiro(title);
        var nomeStr = nomeInfo ? (nomeInfo.month + "/" + nomeInfo.year) : "";

        // Consenso
        var consenso = _consensoData(title, textoPDF);
        var consensoData = consenso.data || "";

        // Veredicto: comparar MM/YYYY do consenso com esperado
        var veredicto = "FAIL";
        if (consensoData) {
          var cParts = consensoData.split("/");
          if (cParts.length === 3 && (cParts[1] + "/" + cParts[2]) === esperado) {
            veredicto = "OK";
          }
        }

        var link = '=HYPERLINK("https://drive.google.com/file/d/' + it.id + '/view","abrir")';

        rows.push([
          mm, folderName, title,
          dataRegex, dataMistral, dataGroq, dataGeminiPro, dataGeminiLite, nomeStr,
          consensoData || "SEM DATA", esperado, veredicto, link
        ]);
      } // for items

      // Flush
      if (rows.length) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }

      // Próxima página
      pageToken = resp.nextPageToken || null;
      props.deleteProperty(POS_KEY);
      if (!pageToken) { props.deleteProperty(TOKEN_KEY); break; }
    } // while páginas

    props.deleteProperty(TOKEN_KEY_PREFIX + folderId);
    props.deleteProperty(POS_KEY_PREFIX + folderId);
    props.setProperty(IDX_KEY, String(idx + 1));
    Logger.log("[" + mm + "] FIM PASTA");
  } // for folders

  // Formatação final
  var lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OK").setBackground("#2e7d32").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, 12, lastRow)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("FAIL").setBackground("#c62828").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, 12, lastRow)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("ERRO").setBackground("#e65100").setFontColor("#ffffff")
      .setRanges([sheet.getRange(2, 12, lastRow)]).build()
  ]);
  for (var c = 1; c <= 13; c++) sheet.autoResizeColumn(c);

  // Limpar estado de retoma
  props.deleteProperty(IDX_KEY);
  props.deleteProperty(SS_KEY);
  Logger.log("✅ Bancada Consenso 2025 concluída.");
  Logger.log("📊 Resultados: " + ss.getUrl());
}

/** Limpa o estado de retoma do teste de consenso (recomeça do zero) */
function consensoReset() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getKeys();
  all.forEach(function(k) { if (k.indexOf("CONS_") === 0) props.deleteProperty(k); });
  // Limpar também cache de targets
  props.deleteProperty("BENCH_TARGETS_CACHE");
  Logger.log("🔄 Retoma do consenso limpa — recomeça do zero na próxima execução.");
}

function _safeDatePERMISSIVA_(dd, mm, yyyy) {
  const d = new Date(Number(yyyy), Number(mm)-1, Number(dd));
  const today = new Date();
  if (isNaN(d.getTime()) || d > today) return null;
  return `${String(dd).padStart(2,'0')}/${String(mm).padStart(2,'0')}/${String(yyyy)}`;
}