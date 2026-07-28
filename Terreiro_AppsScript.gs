// ================================================================
//  ACERVO DO TERREIRO — Apps Script Web App
//  1. Cole este código em Extensões → Apps Script
//  2. Selecione "setup" no dropdown → clique ▶ Executar
//  3. Implantar → Nova implantação → App da Web
//     • Executar como: Eu mesmo
//     • Quem tem acesso: Qualquer pessoa
//  4. Copie a URL /exec e cole no app
// ================================================================

const ABA_ACERVO  = 'Acervo';
const ABA_ESTOQUE = 'Estoque';

// ── SETUP ────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _criarAcervo(ss);
  _criarEstoque(ss);
  ['Plan1','Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0) try { ss.deleteSheet(s); } catch(e) {}
  });
  SpreadsheetApp.getUi().alert('✅ Acervo do Terreiro configurado!\nAbas criadas: Acervo + Estoque');
}

function _criarAcervo(ss) {
  const aba = ss.getSheetByName(ABA_ACERVO) || ss.insertSheet(ABA_ACERVO);
  if (aba.getLastRow() > 0) return; // já configurada

  const cols = ['ID','Nome','Categoria','Subcategoria','Orixá / Entidade',
                'Status','Localização / Armário','Quantidade','Foto (URL)','Observações','Data de cadastro'];
  const hr = aba.getRange(1,1,1,cols.length);
  hr.setValues([cols])
    .setBackground('#3d2b1f').setFontColor('#f5e6c8')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  [80,220,160,140,140,110,160,80,180,220,120].forEach((w,i) => aba.setColumnWidth(i+1,w));
  aba.setFrozenRows(1);

  // Remove colunas e linhas extras
  const totalCols = aba.getMaxColumns();
  if (totalCols > cols.length) aba.deleteColumns(cols.length+1, totalCols-cols.length);

  // Validações
  aba.getRange('C2:C2000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Roupas e indumentárias','Ferramentas e objetos rituais','Ervas e plantas','Alimentos e oferendas'],true)
      .setAllowInvalid(false).build());
  aba.getRange('F2:F2000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Disponível','Em uso','Danificado','Necessita reposição'],true)
      .setAllowInvalid(false).build());

  // Formatação condicional de status
  const regras = [];
  [{v:'Disponível',bg:'#d4edda',f:'#155724'},{v:'Em uso',bg:'#fff3cd',f:'#856404'},
   {v:'Danificado',bg:'#f8d7da',f:'#721c24'},{v:'Necessita reposição',bg:'#ffeeba',f:'#856404'}]
  .forEach(({v,bg,f}) => regras.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
      .setBackground(bg).setFontColor(f).setRanges([aba.getRange('F2:F2000')]).build()));
  aba.setConditionalFormatRules(regras);
  aba.getRange(1,1,1,cols.length).protect().setDescription('Cabeçalho Acervo').setWarningOnly(true);
}

function _criarEstoque(ss) {
  const aba = ss.getSheetByName(ABA_ESTOQUE) || ss.insertSheet(ABA_ESTOQUE);
  if (aba.getLastRow() > 0) return; // já configurada

  const cols = ['ID','Nome','Categoria','Unidade','Qtd Atual','Qtd Mínima','% Estoque','Status','Data de atualização'];
  const hr = aba.getRange(1,1,1,cols.length);
  hr.setValues([cols])
    .setBackground('#1a3a2a').setFontColor('#c8f5e0')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  [80,220,160,110,90,90,90,120,130].forEach((w,i) => aba.setColumnWidth(i+1,w));
  aba.setFrozenRows(1);

  // Remove colunas extras
  const totalCols = aba.getMaxColumns();
  if (totalCols > cols.length) aba.deleteColumns(cols.length+1, totalCols-cols.length);

  // Formatação condicional de status
  const regras = [];
  [{v:'🔴 Crítico',bg:'#f8d7da',f:'#721c24'},{v:'⚠️ Repor',bg:'#fff3cd',f:'#856404'},{v:'✅ Em dia',bg:'#d4edda',f:'#155724'}]
  .forEach(({v,bg,f}) => regras.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
      .setBackground(bg).setFontColor(f).setRanges([aba.getRange('H2:H2000')]).build()));
  aba.setConditionalFormatRules(regras);
  aba.getRange(1,1,1,cols.length).protect().setDescription('Cabeçalho Estoque').setWarningOnly(true);
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
  const id = 'ACE-' + Utilities.getUuid().substring(0,6).toUpperCase();
  aba.appendRow([id, d.nome, d.categoria, d.subcategoria||'', d.orixa||'',
    d.status||'Disponível', d.local||'', Number(d.quantidade)||1,
    d.foto||'', d.observacoes||'',
    d.dataCadastro||new Date().toLocaleDateString('pt-BR')]);
  const ul = aba.getLastRow();
  aba.getRange(ul,8).setHorizontalAlignment('center');
  aba.getRange(ul,11).setHorizontalAlignment('center');
  return saida({ ok: true, id });
}

function _inserirEstoque(d) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!aba) return saida({ ok: false, erro: 'Aba Estoque não encontrada. Execute setup().' });

  const atual  = Number(d.atual)  || 0;
  const minimo = Number(d.minimo) || 1;
  const pct    = Math.round((atual / minimo) * 100);
  const status = pct >= 100 ? '✅ Em dia' : pct >= 50 ? '⚠️ Repor' : '🔴 Crítico';
  const data   = d.dataCadastro || new Date().toLocaleDateString('pt-BR');

  // Se item já existe → atualiza
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if ((dados[i][1]||'').toLowerCase().trim() === (d.nome||'').toLowerCase().trim()) {
      const lin = i + 1;
      aba.getRange(lin,4).setValue(d.unidade||'unidade(s)');
      aba.getRange(lin,5).setValue(atual);
      aba.getRange(lin,6).setValue(minimo);
      aba.getRange(lin,7).setValue(pct + '%');
      aba.getRange(lin,8).setValue(status);
      aba.getRange(lin,9).setValue(data);
      return saida({ ok: true, acao: 'atualizado' });
    }
  }

  // Novo item
  const id = 'EST-' + Utilities.getUuid().substring(0,6).toUpperCase();
  aba.appendRow([id, d.nome, d.categoria||'', d.unidade||'unidade(s)', atual, minimo, pct+'%', status, data]);
  return saida({ ok: true, id, acao: 'inserido' });
}

// ── GET ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const acao = (e.parameter||{}).acao;
    if (acao === 'listar')         return _listarAcervo();
    if (acao === 'estoque-listar') return _listarEstoque();
    return saida({ ok: true, msg: 'Web App ativo.' });
  } catch(err) {
    return saida({ ok: false, erro: err.message });
  }
}

function _listarAcervo() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if (!aba) return saida({ ok: false, erro: 'Aba Acervo não encontrada.' });
  const [,...linhas] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: linhas.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],subcategoria:l[3],orixa:l[4],
    status:l[5],local:l[6],quantidade:l[7],foto:l[8],observacoes:l[9],dataCadastro:l[10]
  }))});
}

function _listarEstoque() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!aba) return saida({ ok: false, erro: 'Aba Estoque não encontrada.' });
  const [,...linhas] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: linhas.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],unidade:l[3],
    atual:l[4],minimo:l[5],pct:l[6],status:l[7],dataCadastro:l[8]
  }))});
}

function saida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
