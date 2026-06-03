const SHEET_OBRAS = "DiarioObra";
const SHEET_USERS = "Usuarios";
const COLS_OBRAS = ['ID','NOME','OBJETO','FISCAL','FORNECEDOR','VALOR','INSTRUMENTO','ID_GED',
  'INICIO','PRAZO_DIAS','PRAZO','STATUS','STATUS_ORIGINAL','PROGRESSO','RESPONSAVEL',
  'AVISOS','PENDENCIAS','ATUALIZACOES','COMENTARIOS_QUINZENAIS',
  'ADITIVO_DIAS','ADITIVO_VALOR','LINK_PASTA','VERSAO','VISTORIA_REALIZADA','VISTORIA_DATA','DATA_FINALIZACAO','ID_TERMO_RECEBIMENTO','EVENTOS'];
const COLS_USERS = ['NOME','SENHA_HASH','PERFIL'];

// Nível 1: Chave secreta
const APP_KEY = "3abeb74aceffdafb4ff157b193149c58cabc2a19173f5d7b4dfc94cec45ea39a";

// Nível 3: Domínio permitido
const DOMINIO_PERMITIDO = "gelado00.github.io";

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    // Verificar chave secreta
    if (!e.parameter.appKey || e.parameter.appKey !== APP_KEY) {
      const errJson = JSON.stringify({ error: "Acesso não autorizado" });
      const cb = e.parameter.callback;
      if (cb) return ContentService.createTextOutput(`${cb}(${errJson})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
      return ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = getSpreadsheet();
    const action = e.parameter.action;
    const callback = e.parameter.callback;
    let result;

    if      (action === "login")           result = login(ss, e.parameter.nome, e.parameter.senha);
    else if (action === "load")            result = loadObras(ss.getSheetByName(SHEET_OBRAS));
    else if (action === "saveOne")         result = saveOneObra(ss.getSheetByName(SHEET_OBRAS), JSON.parse(e.parameter.obra));
    else if (action === "deleteObra")      result = deleteObra(ss.getSheetByName(SHEET_OBRAS), e.parameter.id);
    else if (action === "loadAvisos")      result = loadAvisos(e.parameter.destinatario);
    else if (action === "marcarLido")      result = marcarLido(e.parameter.id);
    else if (action === "marcarTodosLidos") result = marcarTodosLidos(e.parameter.destinatario);
    else if (action === "criarAviso")      result = criarAviso(e.parameter);
    else result = { error: "Ação inválida" };

    const json = JSON.stringify(result);
    if (callback) return ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errJson = JSON.stringify({ error: err.toString() });
    const callback = e.parameter && e.parameter.callback;
    if (callback) return ContentService.createTextOutput(`${callback}(${errJson})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet() {
  const files = DriveApp.getFilesByName("DiarioDeObra");
  let ss = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.create("DiarioDeObra");
  if (!ss.getSheetByName(SHEET_OBRAS)) ss.insertSheet(SHEET_OBRAS).appendRow(COLS_OBRAS);
  if (!ss.getSheetByName(SHEET_USERS)) {
    const s = ss.insertSheet(SHEET_USERS);
    s.appendRow(COLS_USERS);
    s.appendRow(['Fernando','2d403645fd44261178636caf25a777e19e71c6dea77b621399225252c675b765','ADM']);
    s.appendRow(['Carine','a665a45920422f9d417e4867efdc4fb8a04a1f3fff1fa07e998e86f7f7a27ae3','TECNICO']);
    s.appendRow(['Davi','a665a45920422f9d417e4867efdc4fb8a04a1f3fff1fa07e998e86f7f7a27ae3','TECNICO']);
  }
  return ss;
}

function login(ss, nome, senhaHash) {
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, error: "Aba de usuários não encontrada" };
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "Nenhum usuário cadastrado" };
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (const row of data) {
    if (String(row[0]).trim().toLowerCase() === String(nome).trim().toLowerCase() &&
        String(row[1]).trim() === String(senhaHash).trim()) {
      return { success: true, nome: row[0], perfil: row[2] };
    }
  }
  return { success: false, error: "Usuário ou senha incorretos" };
}

function loadObras(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, obras: [] };
  const lastCol = Math.max(sheet.getLastColumn(), COLS_OBRAS.length);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const obras = data.filter(r => r[0] !== '').map(r => ({
    id: r[0], nome: r[1], objeto: r[2], fiscal: r[3], fornecedor: r[4],
    valor: r[5], instrumento: r[6], idGed: r[7], inicio: r[8], prazoDias: r[9],
    prazo: r[10], status: r[11], statusOriginal: r[12], progresso: Number(r[13])||0,
    responsavel: r[14],
    avisos: safeJson(r[15]),
    pendencias: safeJson(r[16]),
    atualizacoes: safeJson(r[17]),
    comentariosQuinzenais: safeJson(r[18]),
    aditivoDias: r[19]||'', aditivoValor: r[20]||'', linkPasta: r[21]||'',
    versao: Number(r[22])||0,
    vistoriaRealizada: r[23]===true||r[23]==='TRUE'||r[23]==='true',
    vistoriaData: r[24]||'',
    dataFinalizacao: r[25]||'',
    idTermoRecebimento: r[26]||'',
    eventos: safeJson(r[27]),
  }));
  return { success: true, obras };
}

function saveOneObra(sheet, obra) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = ids.findIndex(id => String(id) === String(obra.id));
    if (idx >= 0) {
      // Lê a linha existente antes de sobrescrever (para avisos automáticos)
      const existingRow = sheet.getRange(idx + 2, 1, 1, COLS_OBRAS.length).getValues()[0];
      const versaoAtual = Number(existingRow[22]) || 0;
      const versaoObra = Number(obra.versao) || 0;
      if (versaoAtual > versaoObra) {
        return { success: false, conflito: true, versaoAtual };
      }
      const obraAntiga = { responsavel: existingRow[14], pendencias: safeJson(existingRow[16]) };
      obra.versao = versaoObra + 1;
      sheet.getRange(idx + 2, 1, 1, COLS_OBRAS.length).setValues([obraToRow(obra)]);
      gerarAvisosAutomaticos(obraAntiga, obra);
      return { success: true, versao: obra.versao };
    }
  }
  obra.versao = 1;
  sheet.appendRow(obraToRow(obra));
  gerarAvisosAutomaticos(null, obra);
  return { success: true, versao: 1 };
}

function deleteObra(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false };
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(i => String(i) === String(id));
  if (idx >= 0) { sheet.deleteRow(idx + 2); return { success: true }; }
  return { success: false, error: "Obra não encontrada" };
}

function obraToRow(o) {
  return [
    o.id, o.nome||'', o.objeto||'', o.fiscal||'', o.fornecedor||'',
    o.valor||'', o.instrumento||'', o.idGed||'', o.inicio||'', o.prazoDias||'',
    o.prazo||'', o.status||'', o.statusOriginal||'', o.progresso||0, o.responsavel||'',
    JSON.stringify(o.avisos||[]),
    JSON.stringify(o.pendencias||[]),
    JSON.stringify(o.atualizacoes||[]),
    JSON.stringify(o.comentariosQuinzenais||[]),
    o.aditivoDias||'', o.aditivoValor||'', o.linkPasta||'',
    o.versao||0,
    o.vistoriaRealizada||false,
    o.vistoriaData||'',
    o.dataFinalizacao||'',
    o.idTermoRecebimento||'',
    JSON.stringify(o.eventos||[]),
  ];
}

function safeJson(val) {
  try { return val ? JSON.parse(val) : []; } catch(e) { return []; }
}

// ─────────────────────────────────────────────────────────────────────────────
// CENTRAL DE AVISOS
// ─────────────────────────────────────────────────────────────────────────────

function getAvisosSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet
    ? SpreadsheetApp.getActiveSpreadsheet()
    : DriveApp.getFilesByName("DiarioDeObra").hasNext()
      ? SpreadsheetApp.open(DriveApp.getFilesByName("DiarioDeObra").next())
      : null;
  if (!ss) throw new Error("Planilha não encontrada");
  let sheet = ss.getSheetByName('Avisos');
  if (!sheet) {
    sheet = ss.insertSheet('Avisos');
    sheet.appendRow(['ID','DESTINATARIO','TIPO','MENSAGEM','OBRA_ID','OBRA_NOME','DATA','LIDO','REMETENTE']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToAviso(row) {
  return {
    id:           String(row[0] || ''),
    destinatario: String(row[1] || ''),
    tipo:         String(row[2] || ''),
    mensagem:     String(row[3] || ''),
    obraId:       String(row[4] || ''),
    obraNome:     String(row[5] || ''),
    data:         String(row[6] || ''),
    lido:         row[7] === true || String(row[7]).toUpperCase() === 'TRUE',
    remetente:    String(row[8] || ''),
  };
}

function loadAvisos(destinatario) {
  const sheet = getAvisosSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { avisos: [] };

  const dest = String(destinatario || '').trim();
  const avisos = [];

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const rowDest = String(row[1] || '').trim();
    if (dest === '' || rowDest === dest) {
      avisos.push(rowToAviso(row));
    }
  }

  // Não lidos primeiro, depois por data decrescente
  avisos.sort((a, b) => {
    if (a.lido !== b.lido) return a.lido ? 1 : -1;
    return b.data.localeCompare(a.data);
  });

  return { avisos };
}

function marcarLido(id) {
  const sheet = getAvisosSheet();
  const data  = sheet.getDataRange().getValues();
  const strId = String(id || '');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === strId) {
      sheet.getRange(i + 1, 8).setValue(true); // coluna H = LIDO
      return { success: true };
    }
  }
  return { success: false, erro: 'Aviso não encontrado' };
}

function marcarTodosLidos(destinatario) {
  const sheet = getAvisosSheet();
  const data  = sheet.getDataRange().getValues();
  const dest  = String(destinatario || '').trim();

  for (let i = 1; i < data.length; i++) {
    const rowDest = String(data[i][1] || '').trim();
    const jaLido  = data[i][7] === true || String(data[i][7]).toUpperCase() === 'TRUE';
    if (rowDest === dest && !jaLido) {
      sheet.getRange(i + 1, 8).setValue(true);
    }
  }
  return { success: true };
}

function criarAviso(params) {
  const sheet = getAvisosSheet();
  const id    = String(Date.now()) + String(Math.floor(Math.random() * 9999));
  const data  = new Date().toISOString();

  sheet.appendRow([
    id,
    String(params.destinatario || ''),
    String(params.tipo         || ''),
    String(params.mensagem     || ''),
    String(params.obraId       || ''),
    String(params.obraNome     || ''),
    data,
    false,
    String(params.remetente    || 'Sistema'),
  ]);

  return { success: true, id };
}

function gerarAvisosAutomaticos(obraAntiga, novaObra) {
  if (!novaObra || !novaObra.responsavel) return;

  const novoResp = String(novaObra.responsavel).trim();
  const nome     = String(novaObra.nome || '').trim();
  const obraId   = String(novaObra.id   || '');

  // 1. Obra nova atribuída a um técnico
  if (!obraAntiga && novoResp) {
    criarAviso({
      destinatario: novoResp,
      tipo:         'nova_obra',
      mensagem:     `Você foi atribuído à obra ${nome}`,
      obraId,
      obraNome:     nome,
      remetente:    'Sistema',
    });
    return;
  }

  // 2. Responsável trocado
  if (obraAntiga && String(obraAntiga.responsavel || '').trim() !== novoResp) {
    criarAviso({
      destinatario: novoResp,
      tipo:         'nova_obra',
      mensagem:     `Você foi atribuído à obra ${nome}`,
      obraId,
      obraNome:     nome,
      remetente:    'Sistema',
    });
  }

  // 3. Novas pendências adicionadas
  if (obraAntiga) {
    const idsAntigos = new Set(
      (obraAntiga.pendencias || []).map(p => String(p.id))
    );
    const novasPends = (novaObra.pendencias || []).filter(
      p => !idsAntigos.has(String(p.id))
    );
    novasPends.forEach(p => {
      criarAviso({
        destinatario: novoResp,
        tipo:         'nova_pendencia',
        mensagem:     `Nova pendência em ${nome}: ${p.texto}`,
        obraId,
        obraNome:     nome,
        remetente:    'Sistema',
      });
    });
  }
}
