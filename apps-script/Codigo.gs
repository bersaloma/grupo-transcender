/**
 * Quadro de Servidores, Diretrizes e Atas · Grupo Transcender de NA
 * Liga o site à planilha (lê o quadro e salva edições com histórico),
 * ao Google Docs do manual de procedimentos (só leitura)
 * e gera as atas administrativas em Google Docs, na pasta "Administrativas" do Drive do grupo.
 *
 * Instalação (uma vez, logado na conta do grupo):
 * 1. Na planilha: Extensões → Apps Script. Apague o que estiver lá e cole este arquivo.
 * 2. Engrenagem "Configurações do projeto" → Propriedades do script → Adicionar:
 *      PIN = (a senha de serviço combinada com a mesa)
 * 3. Implantar → Nova implantação → tipo "App da Web"
 *      Executar como: Eu (grupotranscenderna@gmail.com)
 *      Quem pode acessar: Qualquer pessoa
 * 4. Autorize, copie a URL que termina em /exec e mande para atualizar o site.
 *
 * Para atualizar este código depois: cole a versão nova, salve e vá em
 * Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar.
 * Assim a URL continua a mesma.
 */

const ABA_QUADRO = 0;            // primeira aba = quadro atual
const ABA_HISTORICO = 'Histórico';
const DOC_DIRETRIZES = '1dKjOKiC5X1EVZK23AyOegQxUxxAb_GGPBne9fxK8Uds'; // manual de procedimentos
const PASTA_ATAS_ID = '1LULlyLiGUQzCCLoxgQIsFxgNWa7ylxLh'; // pasta Administrativas do Drive do grupo
const ABA_ATAS = 'Atas';
const ULTIMA_ATA_PADRAO = 17;    // última ata feita fora do site

/* Rode esta função uma vez pelo editor (Executar) para autorizar o acesso ao Drive. */
function autorizar() {
  DriveApp.getFolderById(PASTA_ATAS_ID).getName();
  DocumentApp.openById(DOC_DIRETRIZES).getName();
  Logger.log('Autorizado. Próxima ata: ' + proximaAta());
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ata === 'proxima') {
    return json({ ok: true, numero: proximaAta() });
  }
  if (e && e.parameter && e.parameter.doc === 'diretrizes') {
    try { return json({ ok: true, blocos: lerDiretrizes() }); }
    catch (err) { return json({ ok: false, erro: String(err) }); }
  }
  return json({ ok: true, linhas: lerQuadro() });
}

/* Lê o Google Docs e devolve os parágrafos em ordem: títulos, texto e itens de lista. */
function lerDiretrizes() {
  const corpo = DocumentApp.openById(DOC_DIRETRIZES).getBody();
  const H = DocumentApp.ParagraphHeading, T = DocumentApp.ElementType;
  const blocos = [];
  for (let i = 0; i < corpo.getNumChildren(); i++) {
    const el = corpo.getChild(i);
    const tipoEl = el.getType();
    if (tipoEl !== T.PARAGRAPH && tipoEl !== T.LIST_ITEM) continue;
    const texto = el.asText().getText().trim();
    if (!texto || /^[-_—=*]{3,}$/.test(texto)) continue;
    let tipo = 'p';
    if (tipoEl === T.LIST_ITEM) tipo = 'li';
    else {
      const h = el.asParagraph().getHeading();
      if (h === H.TITLE || h === H.HEADING1) tipo = 'h1';
      else if (h === H.SUBTITLE || h === H.HEADING2) tipo = 'h2';
      else if (h === H.HEADING3 || h === H.HEADING4) tipo = 'h3';
    }
    blocos.push({ tipo: tipo, html: htmlDoTexto(el.asText()) });
  }
  return blocos;
}

/* Converte o texto em HTML simples, mantendo só o negrito. */
function htmlDoTexto(t) {
  const txt = t.getText();
  const cortes = (t.getTextAttributeIndices() || [0]).concat([txt.length]);
  let out = '';
  for (let k = 0; k < cortes.length - 1; k++) {
    const a = cortes[k], b = cortes[k + 1];
    if (b <= a) continue;
    const pedaco = txt.substring(a, b).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    out += t.isBold(a) && pedaco.trim() ? '<strong>' + pedaco + '</strong>' : pedaco;
  }
  return out.trim();
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const pin = PropertiesService.getScriptProperties().getProperty('PIN');
    if (!pin || String(d.pin || '') !== String(pin)) return json({ ok: false, erro: 'Senha de serviço incorreta.' });
    if (d.acao === 'ata') return json(criarAta(d));
    if (!d.autor || !d.grupo || !d.encargo) return json({ ok: false, erro: 'Dados incompletos.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      salvarEncargo(SpreadsheetApp.getActiveSpreadsheet(), d.grupo, d.encargo, d.nome, d.fim, d.autor, true);
    } finally {
      lock.releaseLock();
    }
    return json({ ok: true, linhas: lerQuadro() });
  } catch (err) {
    return json({ ok: false, erro: String(err) });
  }
}

/* Grava um encargo no quadro e registra no Histórico. Devolve true se algo mudou. */
function salvarEncargo(ss, grupo, encargo, nome, fim, autor, sempre) {
  nome = String(nome || '').trim(); fim = String(fim || '').trim();
  const aba = ss.getSheets()[ABA_QUADRO];
  const valores = aba.getDataRange().getDisplayValues();
  let linha = -1;
  for (let i = 1; i < valores.length; i++) {
    if (norm(valores[i][0]) === norm(grupo) && norm(valores[i][1]) === norm(encargo)) { linha = i; break; }
  }
  let antesNome = '', antesFim = '';
  if (linha !== -1) { antesNome = valores[linha][2]; antesFim = valores[linha][3]; }
  if (!sempre && norm(antesNome) === norm(nome) && norm(antesFim) === norm(fim)) return false;
  if (linha === -1) aba.appendRow([grupo, encargo, nome, fim]);
  else aba.getRange(linha + 1, 3, 1, 2).setValues([[nome, fim]]);
  aba.getRange(1, 3, aba.getMaxRows(), 2).setNumberFormat('@'); // mantém datas como texto dd/mm/aaaa

  let hist = ss.getSheetByName(ABA_HISTORICO);
  if (!hist) {
    hist = ss.insertSheet(ABA_HISTORICO);
    hist.appendRow(['Data e hora', 'Quem editou', 'Grupo', 'Encargo', 'Nome antes', 'Fim antes', 'Nome depois', 'Fim depois']);
    hist.setFrozenRows(1);
  }
  hist.appendRow([new Date(), autor, grupo, encargo, antesNome, antesFim, nome, fim]);
  return true;
}

/* ===== Atas administrativas ===== */
function proximaAta() {
  const p = PropertiesService.getScriptProperties();
  return Number(p.getProperty('ULTIMA_ATA') || ULTIMA_ATA_PADRAO) + 1;
}

const MESES = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
function dataBR(iso) { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + '/' + m[2] + '/' + m[1] : String(iso || ''); }
function dataExtenso(iso) { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + ' de ' + MESES[+m[2] - 1] + ' de ' + m[1] : String(iso || ''); }
function pessoaTxt(p) { return p && p.nome ? p.nome + (p.pres ? ' (' + p.pres + ')' : '') : 'Vago'; }
function termoTxt(p) { return p && p.nome ? p.nome + (p.fim ? ' (até ' + p.fim + ')' : '') : 'Vago'; }

/* Acrescenta parágrafo ou item de lista com partes em negrito: partes = [[texto, negrito], ...] */
function rico(corpo, partes, lista) {
  const txt = partes.map(function (x) { return x[0]; }).join('');
  const el = lista ? corpo.appendListItem(txt || ' ').setGlyphType(DocumentApp.GlyphType.BULLET) : corpo.appendParagraph(txt || ' ');
  if (!lista) el.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  const t = el.editAsText(); t.setBold(false);
  let pos = 0;
  partes.forEach(function (x) { if (x[1] && x[0]) t.setBold(pos, pos + x[0].length - 1, true); pos += x[0].length; });
  return el;
}
function titulo(corpo, txt, nivel) {
  return corpo.appendParagraph(txt).setHeading(nivel || DocumentApp.ParagraphHeading.HEADING2);
}

function criarAta(d) {
  const a = d.ata || {};
  const n = Number(a.numero) || proximaAta();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const doc = DocumentApp.create('ATA N.' + n + ' (' + dataBR(a.data) + ') - Reunião Administrativa');
    const c = doc.getBody();
    const H1 = DocumentApp.ParagraphHeading.HEADING1;
    c.getParagraphs()[0].setText('ATA N.' + n + ' · REUNIÃO ADMINISTRATIVA ORDINÁRIA · GRUPO TRANSCENDER DE NARCÓTICOS ANÔNIMOS').setHeading(H1);
    rico(c, [['Data: ', true], [dataExtenso(a.data) + '.', false]]);
    rico(c, [['Horário: ', true], [(a.inicio || '19h00') + '.', false]]);

    titulo(c, '1. Abertura da Reunião');
    rico(c, [['A reunião iniciou-se às ' + (a.inicio || '19h00') + ' com:', false]]);
    ['Oração da Serenidade', 'Oração do Serviço', 'Leitura das 12 Tradições de Narcóticos Anônimos', 'Leitura dos 12 Conceitos de Serviço de Narcóticos Anônimos']
      .forEach(function (x) { rico(c, [[x, false]], true); });
    rico(c, [['Foi reforçado o compromisso do grupo com o serviço e a inclusão de todas as pessoas trans, travestis, dissidentes de gênero e não bináries em recuperação.', false]]);

    titulo(c, '2. Lista de chamada');
    rico(c, [['Mesa de servidores:', true]]);
    (a.mesa || []).forEach(function (m) { rico(c, [[m.enc + ': ', true], [pessoaTxt(m), false]], true); });
    rico(c, [['Moderação e apoio:', true]]);
    (a.dias || []).forEach(function (x) {
      rico(c, [[x.dia + ' 19h: ', true], ['Moderação ' + pessoaTxt(x.mod) + ' / Apoio ' + pessoaTxt(x.apoio), false]], true);
    });

    titulo(c, '3. Relatórios e breves relatos');
    rico(c, [['Mesa de servidores:', true]]);
    (a.mesa || []).forEach(function (m) {
      rico(c, [[m.enc + (m.nome ? ' (' + m.nome + ')' : '') + ': ', true], [m.nome ? (m.relato || 'sem relato.') : 'Vago', false]], true);
    });
    rico(c, [['Breve relato sobre as reuniões:', true]]);
    (a.dias || []).forEach(function (x) {
      const vago = !(x.mod && x.mod.nome) && !(x.apoio && x.apoio.nome);
      rico(c, [[x.dia + ' 19h: ', true], [vago ? 'Vago' : (x.relato || 'sem ocorrências.'), false]], true);
    });

    let sec = 4;
    if (String(a.novos || '').trim()) { titulo(c, (sec++) + '. Assuntos novos'); rico(c, [[a.novos.trim(), false]]); }
    if (String(a.velhos || '').trim()) { titulo(c, (sec++) + '. Assuntos velhos'); rico(c, [[a.velhos.trim(), false]]); }

    titulo(c, (sec++) + '. Moções');
    const mocoes = (a.mocoes || []).filter(function (m) { return (m.titulo || m.descricao || '').trim(); });
    if (!mocoes.length) rico(c, [['Não houve moções nesta reunião.', false]]);
    mocoes.forEach(function (m, i) {
      rico(c, [['Moção ' + (i + 1) + ': ', true], [m.titulo || '', false]]);
      if (m.endosso) rico(c, [['Endosso: ', true], [m.endosso, false]], true);
      if (m.descricao) rico(c, [['Descrição: ', true], [m.descricao, false]], true);
      if (m.emenda) rico(c, [['Emenda: ', true], [m.emenda, false]], true);
      const res = [m.resultado, m.votos].filter(Boolean).join(': ');
      rico(c, [['Resultado: ', true], [res || 'não informado', false]], true);
    });

    titulo(c, (sec++) + '. Eletivas');
    rico(c, [['Mesa de servidores:', true]]);
    (a.eleMesa || []).forEach(function (m) { rico(c, [[m.enc + ': ', true], [termoTxt(m), false]], true); });
    rico(c, [['Moderação e apoio:', true]]);
    (a.eleDias || []).forEach(function (x) {
      rico(c, [[x.dia + ': ', true], ['Moderação ' + termoTxt(x.mod) + ' / Apoio ' + termoTxt(x.apoio), false]], true);
    });

    titulo(c, (sec++) + '. Encerramento');
    rico(c, [['A reunião foi encerrada' + (a.termino ? ' às ' + a.termino : '') + ' com a Oração da Serenidade, após a conclusão de todas as pautas.', false]]);
    rico(c, [['Ata redigida por: ', true], [(a.autor || d.autor || '') + (a.encargoAutor ? ', ' + a.encargoAutor : '') + ', Grupo Transcender de NA.', false]]);
    rico(c, [['Registrada em: ', true], [Utilities.formatDate(new Date(), 'America/Sao_Paulo', "dd/MM/yyyy 'às' HH:mm") + '.', false]]);
    doc.saveAndClose();

    // pasta do grupo
    DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(PASTA_ATAS_ID));

    // eletivas no quadro
    let alteracoes = 0;
    if (a.atualizarQuadro) {
      const autor = 'Ata N.' + n + ' (' + (a.autor || d.autor || '') + ')';
      (a.eleMesa || []).forEach(function (m) { if (salvarEncargo(ss, 'Mesa', m.enc, m.nome, m.fim, autor, false)) alteracoes++; });
      (a.eleDias || []).forEach(function (x) {
        if (salvarEncargo(ss, 'Moderação', x.dia, x.mod.nome, x.mod.fim, autor, false)) alteracoes++;
        if (salvarEncargo(ss, 'Apoio', x.dia, x.apoio.nome, x.apoio.fim, autor, false)) alteracoes++;
      });
    }

    // registro na aba Atas
    let reg = ss.getSheetByName(ABA_ATAS);
    if (!reg) {
      reg = ss.insertSheet(ABA_ATAS);
      reg.appendRow(['Nº', 'Data da reunião', 'Redigida por', 'Moções', 'Encargos atualizados no quadro', 'Documento', 'Registrada em']);
      reg.setFrozenRows(1);
    }
    const resumo = mocoes.map(function (m) { return (m.titulo || 'Moção') + ' (' + ([m.resultado, m.votos].filter(Boolean).join(' ') || 'sem resultado') + ')'; }).join('; ');
    reg.appendRow([n, dataBR(a.data), (a.autor || d.autor || '') + (a.encargoAutor ? ', ' + a.encargoAutor : ''), resumo || 'Sem moções', alteracoes, doc.getUrl(), new Date()]);

    const p = PropertiesService.getScriptProperties();
    if (n > Number(p.getProperty('ULTIMA_ATA') || ULTIMA_ATA_PADRAO)) p.setProperty('ULTIMA_ATA', String(n));

    return { ok: true, numero: n, url: doc.getUrl(), alteracoes: alteracoes, linhas: lerQuadro() };
  } finally {
    lock.releaseLock();
  }
}

function lerQuadro() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheets()[ABA_QUADRO];
  return aba.getDataRange().getDisplayValues();
}

function norm(s) { return String(s || '').trim().toLowerCase(); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
