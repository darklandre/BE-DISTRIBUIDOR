/*
          ATENÇÃO!!!

    5. O Veredicto sobre a Bancada de Testes
Fizeste bem em manter. No entanto, a bancada usa funções de extração de data (extractDataDocumentoTaloes, extractDataDocumento_Simples_PERMISSIVA) que não são usadas na função principal distribuirFicheirosDoGeral (que usa extractDataDocumento ou extractDataDocumentoTaloes dependendo do sítio).

Regra de Ouro: A Bancada de Testes deve testar as mesmas funções que o Distribuidor usa. Caso contrário, estás a testar uma realidade paralela.

*/

/**
 *
 **************** BANCADA DE TESTES: ****************
 * 
 * Esta função testa todas as faturas de 2025 ou uma pasta específica
 * Como tem tempo limitado, guarda a última pasta que viu e recomeça aí na execução seguinte
 * 
 */

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
            const texto = convertPDFToText(it.id, ['pt', 'en', null]);

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

// OCR via Drive v2 -> Google Doc -> texto (apaga o doc temporário). Evita "Invalid Value" usando só 1 idioma por tentativa.
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