// ================================================================
//  ACERVO DO TERREIRO — Apps Script Web App v4
//  Projeto: "Acervo do Terreiro"
//  Descrição da implantação: "Acervo do Terreiro — Web App"
//
//  1. Renomeie o projeto: clique em "Projeto sem título" → digite
//     "Acervo do Terreiro" → OK
//  2. Selecione "setup" → ▶ Executar
//  3. Implantar → Gerenciar implantações → lápis
//     Versão: Nova versão
//     Descrição: Acervo do Terreiro — Web App
//     → Implantar
// ================================================================

const ABA_ACERVO  = 'Acervo';
const ABA_ESTOQUE = 'Estoque';

// ── PALETA ───────────────────────────────────────────────
// Cabeçalho Acervo  : fundo vinho escuro, texto dourado claro
// Cabeçalho Estoque : fundo verde musgo escuro, texto verde claro
// Status OK         : verde suave  #e6f4ea / #1e6b3a
// Status Repor      : âmbar suave  #fef9e7 / #7d5c00
// Status Crítico    : vermelho suave #fdecea / #8b0000
// Status Em uso     : azul suave   #e8f0fe / #1a56a0
// Status Danificado : laranja suave #fef3e2 / #7a3800

// ── SETUP ────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _configurarAcervo(ss);
  _configurarEstoque(ss);
  _excluirPaginasPadrao(ss);
  SpreadsheetApp.getUi().alert(
    '✅ Acervo do Terreiro configurado!\n\n' +
    '• Abas: Acervo + Estoque\n' +
    '• Página1 removida\n' +
    '• Filtros, cores e formatação aplicados\n\n' +
    'Lembre de renomear o projeto:\n' +
    'Clique em "Projeto sem título" → "Acervo do Terreiro"');
}

// ── ABA ACERVO ───────────────────────────────────────────
function _configurarAcervo(ss) {
  const aba = ss.getSheetByName(ABA_ACERVO) || ss.insertSheet(ABA_ACERVO);
  const COLS = ['ID','Nome','Categoria','Subcategoria','Orixá / Entidade',
                'Status','Localização / Armário','Qtd','Foto (URL)',
                'Observações','Data cadastro'];
  const N = COLS.length;

  // Cabeçalho
  const hr = aba.getRange(1, 1, 1, N);
  hr.setValues([COLS])
    .setBackground('#2c1a10')      // vinho escuro
    .setFontColor('#f0d090')       // dourado claro
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(false);
  aba.setRowHeight(1, 34);

  // Larguras — ocupa ~1500px (tela cheia no Sheets, sem scroll horizontal)
  // ID | Nome | Categoria | Subcategoria | Orixá/Entidade | Status | Localização | Qtd | Foto | Observações | Data
  [65, 230, 160, 145, 145, 125, 160, 55, 85, 210, 110]
    .forEach((w,i) => aba.setColumnWidth(i+1, w));

  // Remove colunas extras
  const mc = aba.getMaxColumns();
  if (mc > N) aba.deleteColumns(N+1, mc-N);

  // Linhas: 28px para dados, max 1000
  const mr = aba.getMaxRows();
  if (mr > 1000) aba.deleteRows(1001, mr-1000);
  if (mr > 1) aba.setRowHeightsForced(2, aba.getMaxRows()-1, 26);

  // Dados: fundo branco/cinza zebra é difícil de manter via script;
  // usamos fundo branco limpo + bordas leves
  aba.getRange(2, 1, aba.getMaxRows()-1, N)
    .setBackground('#fffdf8')
    .setFontColor('#2a1a0e')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  // Centraliza ID, Qtd, Data
  aba.getRange(2,1,aba.getMaxRows()-1,1).setHorizontalAlignment('center'); // ID
  aba.getRange(2,8,aba.getMaxRows()-1,1).setHorizontalAlignment('center'); // Qtd
  aba.getRange(2,11,aba.getMaxRows()-1,1).setHorizontalAlignment('center'); // Data

  // Congela
  aba.setFrozenRows(1);
  aba.setFrozenColumns(0);

  // Validação Categoria
  aba.getRange('C2:C1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Roupas e indumentárias','Ferramentas e objetos rituais',
                           'Ervas e plantas','Alimentos e oferendas'], true)
      .setAllowInvalid(false).build());

  // Validação Status
  aba.getRange('F2:F1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Disponível','Em uso','Danificado','Necessita reposição'], true)
      .setAllowInvalid(false).build());

  // Formatação condicional Status (coluna F) — cores suaves e legíveis
  aba.setConditionalFormatRules([
    { v:'Disponível',          bg:'#e6f4ea', f:'#1e6b3a' },
    { v:'Em uso',              bg:'#e8f0fe', f:'#1a56a0' },
    { v:'Danificado',          bg:'#fef3e2', f:'#7a3800' },
    { v:'Necessita reposição', bg:'#fdecea', f:'#8b0000' },
  ].map(({v,bg,f}) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(v).setBackground(bg).setFontColor(f)
      .setRanges([aba.getRange('F2:F1000')]).build()));

  // Filtro
  try { aba.getFilter().remove(); } catch(e) {}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();

  // Protege cabeçalho
  try { aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true); } catch(e) {}
}

// ── ABA ESTOQUE ──────────────────────────────────────────
function _configurarEstoque(ss) {
  const aba = ss.getSheetByName(ABA_ESTOQUE) || ss.insertSheet(ABA_ESTOQUE);
  const COLS = ['ID','Nome','Categoria','Unidade',
                'Qtd Atual','Qtd Mínima','% Estoque','Status','Atualizado em'];
  const N = COLS.length;

  // Cabeçalho
  const hr = aba.getRange(1,1,1,N);
  hr.setValues([COLS])
    .setBackground('#1a3020')      // verde musgo escuro
    .setFontColor('#b8f0c8')       // verde claro
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(false);
  aba.setRowHeight(1, 34);

  // Larguras — ocupa ~1500px (tela cheia, sem scroll horizontal)
  // ID | Nome | Categoria | Unidade | Qtd Atual | Qtd Mínima | % Estoque | Status | Data
  [65, 310, 210, 120, 110, 110, 100, 145, 130]
    .forEach((w,i) => aba.setColumnWidth(i+1, w));

  // Remove colunas extras
  const mc = aba.getMaxColumns();
  if (mc > N) aba.deleteColumns(N+1, mc-N);

  // Linhas
  const mr = aba.getMaxRows();
  if (mr > 1000) aba.deleteRows(1001, mr-1000);
  if (mr > 1) aba.setRowHeightsForced(2, aba.getMaxRows()-1, 26);

  // Dados
  aba.getRange(2,1,aba.getMaxRows()-1,N)
    .setBackground('#f8fff8')
    .setFontColor('#0a1a0a')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  // Centraliza colunas numéricas e data
  [1,5,6,7,9].forEach(col =>
    aba.getRange(2,col,aba.getMaxRows()-1,1).setHorizontalAlignment('center'));

  // Congela
  aba.setFrozenRows(1);
  aba.setFrozenColumns(0);

  // Formatação condicional Status (H) — cores suaves
  aba.setConditionalFormatRules([
    { v:'✅ Em dia', bg:'#e6f4ea', f:'#1e6b3a' },
    { v:'⚠️ Repor',  bg:'#fef9e7', f:'#7d5c00' },
    { v:'🔴 Crítico',bg:'#fdecea', f:'#8b0000' },
  ].map(({v,bg,f}) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(v).setBackground(bg).setFontColor(f)
      .setRanges([aba.getRange('H2:H1000')]).build()));

  // Filtro
  try { aba.getFilter().remove(); } catch(e) {}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();

  // Protege cabeçalho
  try { aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true); } catch(e) {}
}

// ── REMOVE ABAS PADRÃO ───────────────────────────────────
function _excluirPaginasPadrao(ss) {
  ['Página1','Planilha1','Plan1','Sheet1'].forEach(nome => {
    const s = ss.getSheetByName(nome);
    if (s) try { ss.deleteSheet(s); } catch(e) {}
  });
}

// ── HELPERS ──────────────────────────────────────────────
function _dataFormatada() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

// ── POST ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.acao === 'inserir')         return _inserirAcervo(d);
    if (d.acao === 'estoque-inserir') return _inserirEstoque(d);
    return saida({ ok: false, erro: 'Ação desconhecida: ' + d.acao });
  } catch(err) {
    return saida({ ok: false, erro: err.message });
  }
}

function _inserirAcervo(d) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if (!aba) return saida({ ok: false, erro: 'Aba Acervo não encontrada. Execute setup().' });
  const id  = 'ACE-' + Utilities.getUuid().substring(0,6).toUpperCase();
  const lin = aba.getLastRow() + 1;
  aba.appendRow([id, d.nome, d.categoria, d.subcategoria||'', d.orixa||'',
    d.status||'Disponível', d.local||'', Number(d.quantidade)||1,
    d.foto||'', d.observacoes||'', _dataFormatada()]);
  const r = aba.getRange(lin,1,1,11);
  r.setBackground('#fffdf8').setFontColor('#2a1a0e')
   .setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  aba.getRange(lin,1).setHorizontalAlignment('center');
  aba.getRange(lin,8).setHorizontalAlignment('center');
  aba.getRange(lin,11).setHorizontalAlignment('center');
  aba.setRowHeight(lin, 26);
  return saida({ ok: true, id });
}

function _inserirEstoque(d) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!aba) return saida({ ok: false, erro: 'Aba Estoque não encontrada. Execute setup().' });
  const atual  = Number(d.atual)  || 0;
  const minimo = Number(d.minimo) || 1;
  const pct    = Math.round((atual / minimo) * 100);
  const status = pct >= 100 ? '✅ Em dia' : pct >= 50 ? '⚠️ Repor' : '🔴 Crítico';
  const data   = _dataFormatada();

  // Atualiza se já existe pelo nome
  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][1]||'').toLowerCase().trim() === (d.nome||'').toLowerCase().trim()) {
      const lin = i + 1;
      aba.getRange(lin,4).setValue(d.unidade||'unidade(s)');
      aba.getRange(lin,5).setValue(atual);
      aba.getRange(lin,6).setValue(minimo);
      aba.getRange(lin,7).setValue(pct+'%');
      aba.getRange(lin,8).setValue(status);
      aba.getRange(lin,9).setValue(data);
      return saida({ ok: true, acao: 'atualizado' });
    }
  }

  // Novo
  const id  = 'EST-' + Utilities.getUuid().substring(0,6).toUpperCase();
  const lin = aba.getLastRow() + 1;
  aba.appendRow([id, d.nome, d.categoria||'', d.unidade||'unidade(s)',
    atual, minimo, pct+'%', status, data]);
  const r = aba.getRange(lin,1,1,9);
  r.setBackground('#f8fff8').setFontColor('#0a1a0a')
   .setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  [1,5,6,7,9].forEach(col => aba.getRange(lin,col).setHorizontalAlignment('center'));
  aba.setRowHeight(lin, 26);
  return saida({ ok: true, id, acao: 'inserido' });
}

// ── GET ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const acao = (e.parameter||{}).acao;
    if (acao === 'listar')         return _listarAcervo();
    if (acao === 'estoque-listar') return _listarEstoque();
    return saida({ ok: true, msg: 'Acervo do Terreiro v4 — Web App ativo.' });
  } catch(err) {
    return saida({ ok: false, erro: err.message });
  }
}

function _listarAcervo() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if (!aba) return saida({ ok: false, erro: 'Aba Acervo não encontrada.' });
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],subcategoria:l[3],orixa:l[4],
    status:l[5],local:l[6],quantidade:l[7],foto:l[8],observacoes:l[9],dataCadastro:l[10]
  }))});
}

function _listarEstoque() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!aba) return saida({ ok: false, erro: 'Aba Estoque não encontrada.' });
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],unidade:l[3],
    atual:l[4],minimo:l[5],pct:l[6],status:l[7],dataCadastro:l[8]
  }))});
}

function saida(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
