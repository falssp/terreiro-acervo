// ================================================================
//  ACERVO DO TERREIRO — Apps Script v5
//  Projeto: "Acervo do Terreiro"
//  1. Cole em Extensões → Apps Script
//  2. setup() → ▶ Executar
//  3. Implantar → Nova versão → Implantar
// ================================================================

const ABA_ACERVO      = 'Acervo';
const ABA_CONSUMIVEIS = 'Consumíveis';

// ── NÍVEIS DE ALERTA (% do mínimo) ──────────────────────
// ✅ OK       ≥ 75%
// ⚠️ Repor   25–74%
// 🔴 Alerta   5–24%
// 🚨 Urgente  < 5%  (inclui zero)

function _nivelEstoque(atual, minimo) {
  if (!minimo || minimo === 0) return { nivel: 'sem-minimo', label: '— Sem mínimo', pct: 0 };
  const pct = Math.round((atual / minimo) * 100);
  if (pct >= 75) return { nivel: 'ok',      label: '✅ OK',             pct };
  if (pct >= 25) return { nivel: 'repor',   label: '⚠️ Repor',          pct };
  if (pct >= 5)  return { nivel: 'alerta',  label: '🔴 Alerta',         pct };
  return            { nivel: 'urgente', label: '🚨 Urgente / Zero', pct };
}

// ── SETUP ────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _configurarAcervo(ss);
  _configurarConsumiveis(ss);
  _excluirPaginasPadrao(ss);
  SpreadsheetApp.getUi().alert(
    '✅ Acervo do Terreiro v5 configurado!\n\n' +
    '• Aba Acervo: cadastro geral de itens\n' +
    '• Aba Consumíveis: estoque visual por categoria\n\n' +
    'Níveis de alerta:\n' +
    '  ✅ OK — ≥ 75% do mínimo\n' +
    '  ⚠️ Repor — 25 a 74%\n' +
    '  🔴 Alerta — 5 a 24%\n' +
    '  🚨 Urgente/Zero — < 5%');
}

// ── ABA ACERVO ───────────────────────────────────────────
function _configurarAcervo(ss) {
  const aba = ss.getSheetByName(ABA_ACERVO) || ss.insertSheet(ABA_ACERVO);
  const COLS = ['ID','Nome','Categoria','Subcategoria','Orixá / Entidade',
                'Status','Localização / Armário','Qtd','Foto (URL)',
                'Observações','Data cadastro'];
  const N = COLS.length;
  const hr = aba.getRange(1,1,1,N);
  hr.setValues([COLS])
    .setBackground('#2c1a10').setFontColor('#f0d090')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
  aba.setRowHeight(1, 34);
  [90,240,165,150,150,130,165,55,90,200,115].forEach((w,i) => aba.setColumnWidth(i+1,w));
  const mc = aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  const mr = aba.getMaxRows(); if(mr>1000) aba.deleteRows(1001,mr-1000);
  if(mr>1) aba.setRowHeightsForced(2,aba.getMaxRows()-1,26);
  aba.getRange(2,1,aba.getMaxRows()-1,N).setBackground('#fffdf8').setFontColor('#2a1a0e').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  aba.getRange(2,1,aba.getMaxRows()-1,1).setHorizontalAlignment('center');
  aba.getRange(2,8,aba.getMaxRows()-1,1).setHorizontalAlignment('center');
  aba.getRange(2,11,aba.getMaxRows()-1,1).setHorizontalAlignment('center');
  aba.setFrozenRows(1); aba.setFrozenColumns(0);
  aba.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Roupas e indumentárias','Ferramentas e objetos rituais','Ervas e plantas','Alimentos e oferendas'],true).setAllowInvalid(false).build());
  aba.getRange('F2:F1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Disponível','Em uso','Danificado','Necessita reposição'],true).setAllowInvalid(false).build());
  aba.setConditionalFormatRules([
    {v:'Disponível',bg:'#e6f4ea',f:'#1e6b3a'},{v:'Em uso',bg:'#e8f0fe',f:'#1a56a0'},
    {v:'Danificado',bg:'#fef3e2',f:'#7a3800'},{v:'Necessita reposição',bg:'#fdecea',f:'#8b0000'},
  ].map(({v,bg,f})=>SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v).setBackground(bg).setFontColor(f).setRanges([aba.getRange('F2:F1000')]).build()));
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

// ── ABA CONSUMÍVEIS ──────────────────────────────────────
function _configurarConsumiveis(ss) {
  const aba = ss.getSheetByName(ABA_CONSUMIVEIS) || ss.insertSheet(ABA_CONSUMIVEIS);

  const COLS = ['ID','Categoria','Item','Unidade','Qtd Atual','Qtd Mínima',
                '% Estoque','Nível','Fornecedor Sugerido','Preço Unit. (R$)',
                'Qtd/Pacote','Preço Pacote (R$)','Link Compra','Atualizado em'];
  const N = COLS.length;

  const hr = aba.getRange(1,1,1,N);
  hr.setValues([COLS])
    .setBackground('#1a1a3a').setFontColor('#c8d8f0')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  aba.setRowHeight(1, 40);

  // Larguras pensadas para não ter scroll
  [70,120,190,90,80,85,80,120,170,100,80,110,190,110].forEach((w,i) => aba.setColumnWidth(i+1,w));

  const mc = aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  const mr = aba.getMaxRows(); if(mr>500) aba.deleteRows(501,mr-500);
  if(mr>1) aba.setRowHeightsForced(2,aba.getMaxRows()-1,26);

  aba.setFrozenRows(1); aba.setFrozenColumns(0);

  // Validação Categoria
  const cats = ['Velas','Bebidas','Fumo','Ervas e Defumação','Alimentos','Flores e Naturais'];
  aba.getRange('B2:B500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cats,true).setAllowInvalid(false).build());

  // Validação Nível (coluna H) — calculado pelo script mas pode ser visto
  aba.getRange('H2:H500').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ OK','⚠️ Repor','🔴 Alerta','🚨 Urgente / Zero','— Sem mínimo'],true)
    .setAllowInvalid(true).build());

  // Formatação condicional Nível (coluna H)
  aba.setConditionalFormatRules([
    {v:'✅ OK',            bg:'#e6f4ea',f:'#1e6b3a'},
    {v:'⚠️ Repor',        bg:'#fef9e7',f:'#7d5c00'},
    {v:'🔴 Alerta',       bg:'#fdecea',f:'#8b0000'},
    {v:'🚨 Urgente / Zero',bg:'#f5c6cb',f:'#5c0000'},
  ].map(({v,bg,f})=>SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v).setBackground(bg).setFontColor(f).setRanges([aba.getRange('H2:H500')]).build()));

  // Preço unitário automático: se tiver pacote, calcula; senão usa o direto
  // Fórmula em coluna J (Preço Unit.) para linhas com dados
  // Feito via script ao inserir — não fórmula para evitar referências circulares

  // Dados iniciais: itens padrão do terreiro
  const itensBase = [
    // [Categoria, Item, Unidade, QtdAtual, QtdMin, Fornecedor, PrecoUnit, QtdPacote, PrecoPacote]
    ['Velas','Vela 7 dias branca','unidade',0,10,'','',1,''],
    ['Velas','Vela 7 dias vermelha','unidade',0,6,'','',1,''],
    ['Velas','Vela 7 dias amarela','unidade',0,6,'','',1,''],
    ['Velas','Vela 7 dias azul','unidade',0,6,'','',1,''],
    ['Velas','Vela 7 dias verde','unidade',0,6,'','',1,''],
    ['Velas','Vela 7 dias preta','unidade',0,4,'','',1,''],
    ['Velas','Vela votiva branca','pacote',0,2,'','',12,''],
    ['Velas','Vela taça branca','unidade',0,6,'','',1,''],
    ['Bebidas','Cachaça','garrafa',0,3,'','',1,''],
    ['Bebidas','Cerveja preta','unidade',0,6,'','',1,''],
    ['Bebidas','Vinho tinto seco','garrafa',0,2,'','',1,''],
    ['Bebidas','Mel','pote',0,2,'','',1,''],
    ['Bebidas','Azeite de dendê','garrafa',0,1,'','',1,''],
    ['Bebidas','Água mineral','garrafa',0,6,'','',1,''],
    ['Fumo','Cigarro (maço)','maço',0,3,'','',1,''],
    ['Fumo','Charuto','unidade',0,4,'','',1,''],
    ['Fumo','Cigarrilha','unidade',0,6,'','',1,''],
    ['Fumo','Cachimbo (tabaco)','pacote',0,1,'','',1,''],
    ['Ervas e Defumação','Incenso (caixinha)','caixa',0,3,'','',1,''],
    ['Ervas e Defumação','Pemba branca','unidade',0,4,'','',1,''],
    ['Ervas e Defumação','Pemba colorida','unidade',0,4,'','',1,''],
    ['Ervas e Defumação','Ervas para banho (mix)','maço',0,3,'','',1,''],
    ['Ervas e Defumação','Erva para defumação','maço',0,2,'','',1,''],
    ['Alimentos','Farofa','kg',0,1,'','',1,''],
    ['Alimentos','Pipoca','pacote',0,2,'','',1,''],
    ['Alimentos','Milho de pipoca','kg',0,1,'','',1,''],
    ['Alimentos','Azeite de oliva','garrafa',0,1,'','',1,''],
    ['Alimentos','Amendoim','pacote',0,2,'','',1,''],
    ['Alimentos','Inhame','unidade',0,4,'','',1,''],
    ['Flores e Naturais','Rosas brancas','dúzia',0,1,'','',1,''],
    ['Flores e Naturais','Girassol','unidade',0,6,'','',1,''],
    ['Flores e Naturais','Flores mistas','buquê',0,1,'','',1,''],
    ['Flores e Naturais','Folhas de bananeira','folha',0,4,'','',1,''],
  ];

  // Só insere se aba estiver vazia (sem dados na linha 2)
  if (aba.getRange(2,1).getValue() === '') {
    const data = _dataFormatada();
    itensBase.forEach((item, idx) => {
      const lin = idx + 2;
      const [cat,nome,unid,atual,min,forn,precoUnit,qtdPac,precoPac] = item;
      const { label, pct } = _nivelEstoque(atual, min);
      aba.getRange(lin, 1).setValue('CSM-' + String(idx+1).padStart(3,'0'));
      aba.getRange(lin, 2).setValue(cat);
      aba.getRange(lin, 3).setValue(nome);
      aba.getRange(lin, 4).setValue(unid);
      aba.getRange(lin, 5).setValue(atual);
      aba.getRange(lin, 6).setValue(min);
      aba.getRange(lin, 7).setValue(pct + '%');
      aba.getRange(lin, 8).setValue(label);
      aba.getRange(lin, 9).setValue(forn);
      aba.getRange(lin,10).setValue(precoUnit);
      aba.getRange(lin,11).setValue(qtdPac);
      aba.getRange(lin,12).setValue(precoPac);
      aba.getRange(lin,13).setValue('');
      aba.getRange(lin,14).setValue(data);
    });
    // Formata bloco de dados
    const bloco = aba.getRange(2, 1, itensBase.length, N);
    bloco.setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
    [1,5,6,7,10,11,12].forEach(c => aba.getRange(2,c,itensBase.length,1).setHorizontalAlignment('center'));
    aba.getRange(2,1,itensBase.length,1).setFontColor('#888888');
  }

  // Filtro
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

// ── REMOVE ABAS PADRÃO ───────────────────────────────────
function _excluirPaginasPadrao(ss) {
  ['Página1','Planilha1','Plan1','Sheet1','Estoque'].forEach(nome => {
    const s = ss.getSheetByName(nome);
    if(s) try{ss.deleteSheet(s);}catch(e){}
  });
}

function _dataFormatada() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

// ── POST ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.acao === 'inserir')              return _inserirAcervo(d);
    if (d.acao === 'consumivel-atualizar') return _atualizarConsumivel(d);
    if (d.acao === 'consumivel-inserir')   return _inserirConsumivel(d);
    return saida({ ok: false, erro: 'Ação desconhecida: ' + d.acao });
  } catch(err) {
    return saida({ ok: false, erro: err.message });
  }
}

function _inserirAcervo(d) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if(!aba) return saida({ok:false,erro:'Aba Acervo não encontrada.'});
  const id  = 'ACE-' + Utilities.getUuid().substring(0,6).toUpperCase();
  const lin = aba.getLastRow() + 1;
  aba.appendRow([id,d.nome,d.categoria,d.subcategoria||'',d.orixa||'',
    d.status||'Disponível',d.local||'',Number(d.quantidade)||1,
    d.foto||'',d.observacoes||'',_dataFormatada()]);
  aba.getRange(lin,1,1,11).setBackground('#fffdf8').setFontColor('#2a1a0e').setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  aba.getRange(lin,1).setHorizontalAlignment('center');
  aba.getRange(lin,8).setHorizontalAlignment('center');
  aba.getRange(lin,11).setHorizontalAlignment('center');
  aba.setRowHeight(lin,26);
  return saida({ok:true,id});
}

function _atualizarConsumivel(d) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba) return saida({ok:false,erro:'Aba Consumíveis não encontrada.'});
  const rows = aba.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    // Busca por ID ou por nome exato
    const matchId   = d.id   && rows[i][0] === d.id;
    const matchNome = d.nome && (rows[i][2]||'').toLowerCase().trim() === d.nome.toLowerCase().trim();
    if(matchId || matchNome){
      const lin    = i + 1;
      const atual  = Number(d.atual) || 0;
      const minimo = Number(d.minimo || rows[i][5]) || 1;
      const { label, pct } = _nivelEstoque(atual, minimo);
      const precoUnit = d.precoUnit || rows[i][9] || '';
      const qtdPac    = d.qtdPacote || rows[i][10] || 1;
      const precoPac  = d.precoPacote || rows[i][11] || '';
      // Calcula preço unitário se vier de pacote
      let unitCalc = precoUnit;
      if(!precoUnit && precoPac && qtdPac > 1) unitCalc = (Number(precoPac)/Number(qtdPac)).toFixed(2);
      aba.getRange(lin,5).setValue(atual);
      aba.getRange(lin,6).setValue(minimo);
      aba.getRange(lin,7).setValue(pct+'%');
      aba.getRange(lin,8).setValue(label);
      if(d.fornecedor) aba.getRange(lin,9).setValue(d.fornecedor);
      if(unitCalc)     aba.getRange(lin,10).setValue(unitCalc);
      if(d.qtdPacote)  aba.getRange(lin,11).setValue(qtdPac);
      if(d.precoPacote)aba.getRange(lin,12).setValue(precoPac);
      if(d.link)       aba.getRange(lin,13).setValue(d.link);
      aba.getRange(lin,14).setValue(_dataFormatada());
      return saida({ok:true,acao:'atualizado',nivel:label,pct});
    }
  }
  return saida({ok:false,erro:'Item não encontrado. Use consumivel-inserir para criar.'});
}

function _inserirConsumivel(d) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba) return saida({ok:false,erro:'Aba Consumíveis não encontrada.'});
  const atual  = Number(d.atual) || 0;
  const minimo = Number(d.minimo) || 0;
  const { label, pct } = _nivelEstoque(atual, minimo);
  const qtdPac  = Number(d.qtdPacote) || 1;
  const precoPac= d.precoPacote || '';
  let unitCalc  = d.precoUnit || '';
  if(!unitCalc && precoPac && qtdPac > 1) unitCalc = (Number(precoPac)/qtdPac).toFixed(2);
  const idx = aba.getLastRow() - 1; // linhas de dados
  const id  = 'CSM-' + String(idx+1).padStart(3,'0');
  const lin = aba.getLastRow() + 1;
  aba.appendRow([id,d.categoria||'',d.nome,d.unidade||'unidade',
    atual,minimo,pct+'%',label,
    d.fornecedor||'',unitCalc,qtdPac,precoPac,d.link||'',_dataFormatada()]);
  aba.getRange(lin,1,1,14).setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  [1,5,6,7,10,11,12].forEach(c=>aba.getRange(lin,c).setHorizontalAlignment('center'));
  aba.setRowHeight(lin,26);
  return saida({ok:true,id,acao:'inserido',nivel:label,pct});
}

// ── GET ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const acao = (e.parameter||{}).acao;
    if(acao==='listar')             return _listarAcervo();
    if(acao==='consumiveis-listar') return _listarConsumiveis();
    if(acao==='ml-buscar')          return _buscarML(e.parameter.q, e.parameter.categoria);
    return saida({ok:true,msg:'Acervo do Terreiro v5 — Web App ativo.'});
  } catch(err) {
    return saida({ok:false,erro:err.message});
  }
}

function _listarAcervo() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if(!aba) return saida({ok:false,erro:'Aba Acervo não encontrada.'});
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ok:true,itens:rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],subcategoria:l[3],orixa:l[4],
    status:l[5],local:l[6],quantidade:l[7],foto:l[8],observacoes:l[9],dataCadastro:l[10]
  }))});
}

function _listarConsumiveis() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba) return saida({ok:false,erro:'Aba Consumíveis não encontrada.'});
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ok:true,itens:rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],categoria:l[1],nome:l[2],unidade:l[3],
    atual:l[4],minimo:l[5],pct:l[6],nivel:l[7],
    fornecedor:l[8],precoUnit:l[9],qtdPacote:l[10],precoPacote:l[11],
    link:l[12],atualizado:l[13]
  }))});
}

// ── MERCADO LIVRE (API pública, gratuita) ─────────────────
function _buscarML(query, categoria) {
  try {
    const q = encodeURIComponent(query || '');
    const url = `https://api.mercadolibre.com/sites/MLB/search?q=${q}&limit=5`;
    const resp = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
    if(resp.getResponseCode() !== 200) return saida({ok:false,erro:'ML indisponível'});
    const data = JSON.parse(resp.getContentText());
    const results = (data.results||[]).slice(0,5).map(r=>({
      titulo:    r.title,
      preco:     r.price,
      link:      r.permalink,
      thumbnail: r.thumbnail,
      vendedor:  (r.seller||{}).nickname || '',
      qtdPacote: _detectarQuantidade(r.title),
    }));
    return saida({ok:true,resultados:results});
  } catch(err) {
    return saida({ok:false,erro:err.message});
  }
}

function _detectarQuantidade(titulo) {
  // Tenta extrair "pacote com X", "caixa com X", "X unidades" do título
  const m = titulo.match(/(\d+)\s*(un|unid|pç|pc|peças|velas|charuto|cigarros)/i);
  return m ? parseInt(m[1]) : 1;
}

function saida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
