/**
 * Quadro de Servidores e Diretrizes · Grupo Transcender de NA
 * Liga o site à planilha (lê o quadro e salva edições com histórico)
 * e ao Google Docs do manual de procedimentos (só leitura).
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

function doGet(e) {
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
    if (!d.autor || !d.grupo || !d.encargo) return json({ ok: false, erro: 'Dados incompletos.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const aba = ss.getSheets()[ABA_QUADRO];
      const valores = aba.getDataRange().getDisplayValues();
      let linha = -1;
      for (let i = 1; i < valores.length; i++) {
        if (norm(valores[i][0]) === norm(d.grupo) && norm(valores[i][1]) === norm(d.encargo)) { linha = i; break; }
      }
      let antesNome = '', antesFim = '';
      if (linha === -1) {
        aba.appendRow([d.grupo, d.encargo, d.nome || '', d.fim || '']);
      } else {
        antesNome = valores[linha][2]; antesFim = valores[linha][3];
        aba.getRange(linha + 1, 3, 1, 2).setValues([[d.nome || '', d.fim || '']]);
      }
      aba.getRange(1, 3, aba.getMaxRows(), 2).setNumberFormat('@'); // mantém datas como texto dd/mm/aaaa

      let hist = ss.getSheetByName(ABA_HISTORICO);
      if (!hist) {
        hist = ss.insertSheet(ABA_HISTORICO);
        hist.appendRow(['Data e hora', 'Quem editou', 'Grupo', 'Encargo', 'Nome antes', 'Fim antes', 'Nome depois', 'Fim depois']);
        hist.setFrozenRows(1);
      }
      hist.appendRow([new Date(), d.autor, d.grupo, d.encargo, antesNome, antesFim, d.nome || '', d.fim || '']);
    } finally {
      lock.releaseLock();
    }
    return json({ ok: true, linhas: lerQuadro() });
  } catch (err) {
    return json({ ok: false, erro: String(err) });
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
