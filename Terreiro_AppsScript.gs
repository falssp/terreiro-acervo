// ================================================================
//  ACERVO DO TERREIRO — Apps Script v6
//  Abas: Acervo | Consumíveis | Filhos de Santo | Entidades/Orixás | Financeiro
//  1. Cole em Extensões → Apps Script
//  2. setup() → ▶ Executar
//  3. Implantar → Nova versão → Implantar
// ================================================================

const ABA_ACERVO      = 'Acervo';
const ABA_CONSUMIVEIS = 'Consumíveis';
const ABA_FILHOS      = 'Filhos de Santo';
const ABA_ENTIDADES   = 'Entidades e Orixás';
const ABA_FINANCEIRO  = 'Financeiro';

// ── NÍVEIS DE ESTOQUE ────────────────────────────────────
function _nivelEstoque(atual, minimo) {
  if (!minimo || minimo === 0) return { nivel:'sem-minimo', label:'— Sem mínimo', pct:0 };
  const pct = Math.round((atual / minimo) * 100);
  if (pct >= 75) return { nivel:'ok',      label:'✅ OK',              pct };
  if (pct >= 25) return { nivel:'repor',   label:'⚠️ Repor',           pct };
  if (pct >= 5)  return { nivel:'alerta',  label:'🔴 Alerta',          pct };
  return             { nivel:'urgente', label:'🚨 Urgente / Zero',  pct };
}

function _dataFormatada() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

// ── MENU CUSTOMIZADO ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏛️ Terreiro')
    .addItem('⚙️ Configurar planilha (setup)', 'setup')
    .addSeparator()
    .addItem('🔄 Recalcular níveis de estoque', 'recalcularEstoque')
    .addItem('🛒 Gerar lista de compras', 'gerarListaCompras')
    .addSeparator()
    .addItem('📅 Ver datas do mês atual', 'verDatasMes')
    .addItem('⚠️ Ver obrigações próximas (90 dias)', 'verObrigacoes')
    .addSeparator()
    .addItem('🧹 Limpar filtros — todas as abas', 'limparFiltros')
    .addToUi();
}

// ── SETUP PRINCIPAL ──────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _configurarAcervo(ss);
  _configurarConsumiveis(ss);
  _configurarFilhos(ss);
  _configurarEntidades(ss);
  _configurarFinanceiro(ss);
  _excluirPaginasPadrao(ss);
  _autoResizeConsumiveis(ss);
  // Reordena abas
  const ordem = [ABA_ACERVO, ABA_CONSUMIVEIS, ABA_FILHOS, ABA_ENTIDADES, ABA_FINANCEIRO];
  ordem.forEach((nome, i) => {
    const aba = ss.getSheetByName(nome);
    if (aba) ss.setActiveSheet(aba), ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(ABA_ACERVO));
  SpreadsheetApp.getUi().alert(
    '✅ Acervo do Terreiro v6 configurado!\n\n' +
    '• Acervo · Consumíveis · Filhos de Santo\n' +
    '• Entidades e Orixás · Financeiro\n\n' +
    'Use o menu 🏛️ Terreiro para ações rápidas.');
}

// ── ABA ACERVO ───────────────────────────────────────────
function _configurarAcervo(ss) {
  const aba = ss.getSheetByName(ABA_ACERVO) || ss.insertSheet(ABA_ACERVO);
  const COLS = ['ID','Nome','Categoria','Subcategoria','Orixá / Entidade',
                'Status','Localização / Armário','Qtd','Foto (URL)','Observações','Data cadastro'];
  const N = COLS.length;
  const hr = aba.getRange(1,1,1,N);
  hr.setValues([COLS]).setBackground('#2c1a10').setFontColor('#f0d090')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(false);
  aba.setRowHeight(1, 34);
  [90,240,165,150,150,130,165,55,90,200,115].forEach((w,i) => aba.setColumnWidth(i+1,w));
  const mc=aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  const mr=aba.getMaxRows(); if(mr>1000) aba.deleteRows(1001,mr-1000);
  if(mr>1) aba.setRowHeightsForced(2,aba.getMaxRows()-1,26);
  aba.getRange(1,1,1,N).setHorizontalAlignment('center');
  aba.getRange(2,1,aba.getMaxRows()-1,N).setBackground('#fffdf8').setFontColor('#2a1a0e')
    .setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  [1,8,11].forEach(c=>aba.getRange(2,c,aba.getMaxRows()-1,1).setHorizontalAlignment('center'));
  aba.setFrozenRows(1);
  aba.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Roupas e indumentárias','Ferramentas e objetos rituais','Ervas e plantas','Alimentos e oferendas'],true)
    .setAllowInvalid(false).build());
  aba.getRange('F2:F1000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Disponível','Em uso','Danificado','Necessita reposição'],true)
    .setAllowInvalid(false).build());
  aba.setConditionalFormatRules([
    {v:'Disponível',bg:'#e6f4ea',f:'#1e6b3a'},{v:'Em uso',bg:'#e8f0fe',f:'#1a56a0'},
    {v:'Danificado',bg:'#fef3e2',f:'#7a3800'},{v:'Necessita reposição',bg:'#fdecea',f:'#8b0000'},
  ].map(({v,bg,f})=>SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
    .setBackground(bg).setFontColor(f).setRanges([aba.getRange('F2:F1000')]).build()));
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

// ── ABA CONSUMÍVEIS ──────────────────────────────────────
function _configurarConsumiveis(ss) {
  const aba = ss.getSheetByName(ABA_CONSUMIVEIS) || ss.insertSheet(ABA_CONSUMIVEIS);
  const COLS = ['ID','Categoria','Item','Unidade','Qtd Atual','Qtd Mínima',
                '% Estoque','Nível','Fornecedor','Preço Unit.','Qtd/Pac',
                'Preço Pacote','Link Compra','Atualizado em'];
  const N = COLS.length;
  aba.getRange(1,1,1,N).setValues([COLS]).setBackground('#1a1a3a').setFontColor('#c8d8f0')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(false);
  aba.setRowHeight(1,34);
  const mc=aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  const mr=aba.getMaxRows(); if(mr>500) aba.deleteRows(501,mr-500);
  if(mr>1) aba.setRowHeightsForced(2,aba.getMaxRows()-1,26);
  aba.setFrozenRows(1);
  aba.getRange('B2:B500').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Velas','Bebidas','Fumo','Ervas e Defumação','Alimentos','Flores e Naturais'],true)
    .setAllowInvalid(false).build());
  aba.setConditionalFormatRules([
    {v:'✅ OK',bg:'#e6f4ea',f:'#1e6b3a'},{v:'⚠️ Repor',bg:'#fef9e7',f:'#7d5c00'},
    {v:'🔴 Alerta',bg:'#fdecea',f:'#8b0000'},{v:'🚨 Urgente / Zero',bg:'#f5c6cb',f:'#5c0000'},
  ].map(({v,bg,f})=>SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
    .setBackground(bg).setFontColor(f).setRanges([aba.getRange('H2:H500')]).build()));
  if(aba.getRange(2,1).getValue()===''){
    const data=_dataFormatada();
    const itens=[
      ['Velas','Vela 7 dias branca','unidade',0,10],['Velas','Vela 7 dias vermelha','unidade',0,6],
      ['Velas','Vela 7 dias amarela','unidade',0,6],['Velas','Vela 7 dias azul','unidade',0,6],
      ['Velas','Vela 7 dias verde','unidade',0,6],['Velas','Vela 7 dias preta','unidade',0,4],
      ['Velas','Vela votiva branca','pacote',0,2],['Velas','Vela taça branca','unidade',0,6],
      ['Bebidas','Cachaça','garrafa',0,3],['Bebidas','Cerveja preta','unidade',0,6],
      ['Bebidas','Vinho tinto seco','garrafa',0,2],['Bebidas','Mel','pote',0,2],
      ['Bebidas','Azeite de dendê','garrafa',0,1],['Bebidas','Água mineral','garrafa',0,6],
      ['Fumo','Cigarro (maço)','maço',0,3],['Fumo','Charuto','unidade',0,4],
      ['Fumo','Cigarrilha','unidade',0,6],['Fumo','Cachimbo (tabaco)','pacote',0,1],
      ['Ervas e Defumação','Incenso (caixinha)','caixa',0,3],['Ervas e Defumação','Pemba branca','unidade',0,4],
      ['Ervas e Defumação','Pemba colorida','unidade',0,4],['Ervas e Defumação','Ervas para banho (mix)','maço',0,3],
      ['Ervas e Defumação','Erva para defumação','maço',0,2],['Alimentos','Farofa','kg',0,1],
      ['Alimentos','Pipoca','pacote',0,2],['Alimentos','Milho de pipoca','kg',0,1],
      ['Alimentos','Azeite de oliva','garrafa',0,1],['Alimentos','Amendoim','pacote',0,2],
      ['Alimentos','Inhame','unidade',0,4],['Flores e Naturais','Rosas brancas','dúzia',0,1],
      ['Flores e Naturais','Girassol','unidade',0,6],['Flores e Naturais','Flores mistas','buquê',0,1],
      ['Flores e Naturais','Folhas de bananeira','folha',0,4],
    ];
    itens.forEach(([cat,nome,unid,atual,min],idx)=>{
      const lin=idx+2,{label,pct}=_nivelEstoque(atual,min);
      aba.getRange(lin,1).setValue('CSM-'+String(idx+1).padStart(3,'0'));
      aba.getRange(lin,2).setValue(cat); aba.getRange(lin,3).setValue(nome);
      aba.getRange(lin,4).setValue(unid); aba.getRange(lin,5).setValue(atual);
      aba.getRange(lin,6).setValue(min); aba.getRange(lin,7).setValue(pct+'%');
      aba.getRange(lin,8).setValue(label); aba.getRange(lin,9).setValue('');
      aba.getRange(lin,10).setValue(''); aba.getRange(lin,11).setValue(1);
      aba.getRange(lin,12).setValue(''); aba.getRange(lin,13).setValue('');
      aba.getRange(lin,14).setValue(data);
    });
  }
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

function _autoResizeConsumiveis(ss) {
  const aba=ss.getSheetByName(ABA_CONSUMIVEIS); if(!aba)return;
  const N=14;
  aba.getRange(1,1,1,N).setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  for(let c=1;c<=N;c++) aba.autoResizeColumn(c);
  const mins=[60,90,165,80,75,80,75,120,150,85,65,95,155,105];
  mins.forEach((w,i)=>{if(aba.getColumnWidth(i+1)<w)aba.setColumnWidth(i+1,w);});
  aba.setRowHeight(1,34);
  aba.getRange(1,1,1,N).setHorizontalAlignment('center');
  [1,4,5,6,7,8,10,11,12,14].forEach(c=>aba.getRange(2,c,498,1).setHorizontalAlignment('center'));
  [2,3,9,13].forEach(c=>aba.getRange(2,c,498,1).setHorizontalAlignment('left'));
}

// ── ABA FILHOS DE SANTO ──────────────────────────────────
function _configurarFilhos(ss) {
  const aba = ss.getSheetByName(ABA_FILHOS) || ss.insertSheet(ABA_FILHOS);
  if(aba.getLastRow()>0 && aba.getRange(1,1).getValue()==='ID') {
    _formatarFilhos(aba); return; // já existe — só reformata
  }
  const COLS = [
    'ID','Nome de Candomblé','Nome Social','Contato',
    'Data de Aniversário','Data de Feitura','Orixá de Cabeça','Adjuntó / Juntó',
    'Nação','Próxima Obrigação','Obrigações Realizadas','Status no Terreiro',
    'Observações','Cadastrado em'
  ];
  const N=COLS.length;
  aba.getRange(1,1,1,N).setValues([COLS]).setBackground('#2a0a2a').setFontColor('#e8c8f0')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(false);
  aba.setRowHeight(1,34);
  [70,180,160,130,110,110,130,130,90,120,170,120,200,110]
    .forEach((w,i)=>aba.setColumnWidth(i+1,w));
  const mc=aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  const mr=aba.getMaxRows(); if(mr>300) aba.deleteRows(301,mr-300);
  if(mr>1) aba.setRowHeightsForced(2,aba.getMaxRows()-1,26);
  _formatarFilhos(aba);
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

function _formatarFilhos(aba) {
  const N=14;
  const orixa=['Exu','Ogum','Oxossi','Xangô','Oxum','Iemanjá','Oxalá','Iansã','Obá',
               'Nanã','Obaluaê / Omolu','Oxumaré','Preto-Velho','Caboclo','Outro'];
  const status=['Ativo','Inativo','Suspenso','Abiã','Iaô','Ebomi','Ogã / Ekedi'];
  const nacoes=['Ketu','Angola','Jeje','Efon','Ijexá','Nagô','Omolokô','Umbanda','Outra'];
  aba.getRange('G2:G300').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(orixa,true).setAllowInvalid(true).build());
  aba.getRange('H2:H300').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(orixa,true).setAllowInvalid(true).build());
  aba.getRange('I2:I300').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(nacoes,true).setAllowInvalid(true).build());
  aba.getRange('L2:L300').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(status,true).setAllowInvalid(true).build());
  // Datas: formata colunas E, F, J
  ['E2:E300','F2:F300','J2:J300'].forEach(r=>
    aba.getRange(r).setNumberFormat('dd/MM/yyyy'));
  // Formatação condicional por status
  aba.setConditionalFormatRules([
    {v:'Ativo',bg:'#e6f4ea',f:'#1e6b3a'},{v:'Ebomi',bg:'#e8f0fe',f:'#1a56a0'},
    {v:'Suspenso',bg:'#fdecea',f:'#8b0000'},{v:'Inativo',bg:'#f5f5f5',f:'#666666'},
    {v:'Abiã',bg:'#fff9e6',f:'#7d5c00'},{v:'Iaô',bg:'#f0e6ff',f:'#4a0080'},
  ].map(({v,bg,f})=>SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
    .setBackground(bg).setFontColor(f).setRanges([aba.getRange('L2:L300')]).build()));
  // Alerta próxima obrigação nos próximos 90 dias (coluna J — laranja)
  aba.getRange(2,1,aba.getMaxRows()-1,N).setFontSize(10).setVerticalAlignment('middle');
  aba.setFrozenRows(1);
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
}

// ── ABA ENTIDADES / ORIXÁS ───────────────────────────────
function _configurarEntidades(ss) {
  const aba = ss.getSheetByName(ABA_ENTIDADES) || ss.insertSheet(ABA_ENTIDADES);
  if(aba.getLastRow()>0 && aba.getRange(1,1).getValue()==='Orixá / Entidade') return;
  const COLS = [
    'Orixá / Entidade','Nação / Qualidade','Data da Festa','Dia da Semana',
    'Cores','Oferendas Preferidas','Bebidas / Alimentos','Itens do Acervo',
    'Saudação','Observações Rituais'
  ];
  const N=COLS.length;
  aba.getRange(1,1,1,N).setValues([COLS]).setBackground('#1a2a1a').setFontColor('#c8f0c8')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(false);
  aba.setRowHeight(1,34);
  [160,140,100,100,120,220,180,180,120,220].forEach((w,i)=>aba.setColumnWidth(i+1,w));
  const mc=aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  aba.setFrozenRows(1);
  aba.getRange('C2:C50').setNumberFormat('dd/MM');
  // Dados base dos orixás mais comuns
  const orixas=[
    ['Exu','Todas','Segunda-feira','Segunda','Preto e vermelho','Pimenta, dendê, farofa, cachaça','Cachaça, vinho tinto','Ogó, tridentes, sete chaves','Laroyê Exu!','Guardião das encruzilhadas'],
    ['Ogum','Ketu / Angola','23/04','Terça','Verde e preto','Feijão preto, carne, dendê','Cerveja preta, vinho tinto','Espada, ferramentas de ferro','Ogum Yê!','Orixá do ferro e da guerra'],
    ['Oxossi','Ketu','—','Quinta','Azul e verde','Milho branco, inhame, mel','Mel, água de coco','Arco e flecha','Okê Arô!','Caçador, orixá da fartura'],
    ['Xangô','Ketu','04/12','Quarta','Vermelho e branco','Acarajé, vatapá, azeite','Vinho tinto, cerveja','Machado duplo (oxé)','Kaô Kabiesilê!','Senhor da justiça'],
    ['Oxum','Ketu','08/12','Sábado','Amarelo e dourado','Mel, acarajé, milho amarelo','Mel, champanhe, laranja','Abebê, espelho, leque','Ora Iê Iê Ô!','Orixá do amor e das águas doces'],
    ['Iemanjá','Ketu','02/02','Sábado','Azul e branco','Melão, uva branca, arroz','Champanhe, água, leite','Abebê de prata','Odoyá!','Rainha do mar'],
    ['Oxalá','Ketu','—','Sexta','Branco','Inhame, arroz, canjica branca','Água, leite','Opaxorô, pano da costa branco','Êpa Babá!','Pai criador, orixá da paz'],
    ['Iansã','Ketu','04/12','Terça','Vermelho e marrom','Acarajé, abará','Vinho tinto, cerveja','Espada, eruexim','Eparrêi Iansã!','Orixá dos ventos e dos raios'],
    ['Nanã','Ketu','26/07','Segunda','Roxo e branco','Inhame, milho, canjica','Água, leite','Ibiri (vassoura de palha)','Salúbà Nanã!','Mais antiga dos orixás, orixá da lama'],
    ['Obaluaê / Omolu','Ketu','—','Segunda','Preto, branco e vermelho','Pipoca, milho, coco','Vinho tinto, dendê','Xaxará','Atotô Obaluaê!','Orixá da saúde e das doenças'],
    ['Preto-Velho','Umbanda','13/05','Segunda','Branco e preto','Fumo, cachaça, mel','Cachaça, mel','Cachimbo, bengala','Ave, meu filho!','Espírito de sabedoria e cura'],
    ['Caboclo','Umbanda','—','Terça','Verde e amarelo','Mel, frutas, charuto','Cerveja, cachaça','Arco e flecha, cocar','Okê Caboclo!','Espírito da natureza'],
    ['Exu / Pombagira','Umbanda','—','Segunda e sexta','Vermelho e preto','Rosa vermelha, champanhe','Champanhe, vinho tinto','Rosas, tridentes','Laroyê!','Guardião(ã) das encruzilhadas'],
  ];
  if(aba.getRange(2,1).getValue()===''){
    aba.getRange(2,1,orixas.length,N).setValues(orixas);
    aba.getRange(2,1,orixas.length,N).setFontSize(10).setVerticalAlignment('middle').setWrap(true);
    aba.setRowHeightsForced(2,orixas.length,42);
  }
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

// ── ABA FINANCEIRO ───────────────────────────────────────
function _configurarFinanceiro(ss) {
  const aba = ss.getSheetByName(ABA_FINANCEIRO) || ss.insertSheet(ABA_FINANCEIRO);
  if(aba.getLastRow()>0 && aba.getRange(1,1).getValue()==='ID') return;
  const COLS=['ID','Data','Tipo','Categoria','Descrição','Valor (R$)','Responsável','Comprovante','Observações'];
  const N=COLS.length;
  aba.getRange(1,1,1,N).setValues([COLS]).setBackground('#1a2a0a').setFontColor('#d0f0b0')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(false);
  aba.setRowHeight(1,34);
  [70,95,90,140,220,100,130,130,200].forEach((w,i)=>aba.setColumnWidth(i+1,w));
  const mc=aba.getMaxColumns(); if(mc>N) aba.deleteColumns(N+1,mc-N);
  const mr=aba.getMaxRows(); if(mr>2000) aba.deleteRows(2001,mr-2000);
  if(mr>1) aba.setRowHeightsForced(2,aba.getMaxRows()-1,26);
  aba.setFrozenRows(1);
  // Validações
  aba.getRange('C2:C2000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Entrada','Saída'],true).setAllowInvalid(false).build());
  const cats=['Dízimo / Doação','Evento / Gira','Venda','Outros (entrada)',
              'Consumíveis','Manutenção','Indumentária / Acervo','Ritual / Oferenda','Outros (saída)'];
  aba.getRange('D2:D2000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(cats,true).setAllowInvalid(true).build());
  aba.getRange('B2:B2000').setNumberFormat('dd/MM/yyyy');
  aba.getRange('F2:F2000').setNumberFormat('R$ #,##0.00');
  // Formatação condicional Entrada/Saída
  aba.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Entrada')
      .setBackground('#e6f4ea').setFontColor('#1e6b3a').setRanges([aba.getRange('C2:C2000')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Saída')
      .setBackground('#fdecea').setFontColor('#8b0000').setRanges([aba.getRange('C2:C2000')]).build(),
  ]);
  aba.getRange(2,1,aba.getMaxRows()-1,N).setFontSize(10).setVerticalAlignment('middle');
  [1,2,3,5,6].forEach(c=>aba.getRange(2,c,aba.getMaxRows()-1,1).setHorizontalAlignment('center'));
  try{aba.getFilter().remove();}catch(e){}
  aba.getRange(1,1,Math.max(2,aba.getLastRow()),N).createFilter();
  try{aba.getRange(1,1,1,N).protect().setDescription('Cabeçalho').setWarningOnly(true);}catch(e){}
}

// ── AÇÕES DO MENU ────────────────────────────────────────
function recalcularEstoque() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const aba=ss.getSheetByName(ABA_CONSUMIVEIS);
  if(!aba){SpreadsheetApp.getUi().alert('Aba Consumíveis não encontrada.');return;}
  const rows=aba.getDataRange().getValues();
  let atualizados=0;
  for(let i=1;i<rows.length;i++){
    const atual=Number(rows[i][4])||0, min=Number(rows[i][5])||0;
    if(min>0){
      const {label,pct}=_nivelEstoque(atual,min);
      aba.getRange(i+1,7).setValue(pct+'%');
      aba.getRange(i+1,8).setValue(label);
      atualizados++;
    }
  }
  SpreadsheetApp.getUi().alert(`✅ ${atualizados} itens recalculados.`);
}

function gerarListaCompras() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const aba=ss.getSheetByName(ABA_CONSUMIVEIS);
  if(!aba){SpreadsheetApp.getUi().alert('Aba Consumíveis não encontrada.');return;}
  const rows=aba.getDataRange().getValues();
  const lista=rows.slice(1).filter(r=>r[0]!==''&&!String(r[7]).includes('OK'));
  if(!lista.length){SpreadsheetApp.getUi().alert('✅ Nenhum item precisa ser reposto!');return;}
  const urgente=lista.filter(r=>String(r[7]).includes('Urgente'));
  const alerta=lista.filter(r=>String(r[7]).includes('Alerta'));
  const repor=lista.filter(r=>String(r[7]).includes('Repor'));
  let msg='🛒 LISTA DE COMPRAS\n'+'─'.repeat(30)+'\n';
  if(urgente.length){msg+='🚨 URGENTE:\n';urgente.forEach(r=>msg+=`  • ${r[2]} — faltam ${Math.max(0,r[5]-r[4])} ${r[3]}\n`);}
  if(alerta.length){msg+='\n🔴 ALERTA:\n';alerta.forEach(r=>msg+=`  • ${r[2]} — faltam ${Math.max(0,r[5]-r[4])} ${r[3]}\n`);}
  if(repor.length){msg+='\n⚠️ REPOR:\n';repor.forEach(r=>msg+=`  • ${r[2]} — faltam ${Math.max(0,r[5]-r[4])} ${r[3]}\n`);}
  SpreadsheetApp.getUi().alert(msg);
}

function verDatasMes() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hoje=new Date(), mes=hoje.getMonth(), ano=hoje.getFullYear();
  const nomeMes=Utilities.formatDate(hoje,Session.getScriptTimeZone(),'MMMM/yyyy');
  let msg=`📅 DATAS DE ${nomeMes.toUpperCase()}\n`+'─'.repeat(30)+'\n';
  // Aniversariantes
  const abaF=ss.getSheetByName(ABA_FILHOS);
  if(abaF&&abaF.getLastRow()>1){
    const rows=abaF.getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
    const aniv=rows.filter(r=>{
      if(!r[4])return false;
      const d=new Date(r[4]); return d.getMonth()===mes;
    });
    if(aniv.length){
      msg+='🎂 Aniversariantes:\n';
      aniv.sort((a,b)=>new Date(a[4]).getDate()-new Date(b[4]).getDate());
      aniv.forEach(r=>msg+=`  • ${_dataFormatada_(r[4])} — ${r[1]||r[2]}\n`);
    }else{msg+='🎂 Sem aniversariantes este mês.\n';}
  }
  // Festas de orixás
  const abaE=ss.getSheetByName(ABA_ENTIDADES);
  if(abaE&&abaE.getLastRow()>1){
    const rows=abaE.getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
    const festas=rows.filter(r=>{
      if(!r[2]||r[2]==='—')return false;
      const parts=String(r[2]).split('/');
      return parts.length>=2&&(parseInt(parts[1])-1)===mes;
    });
    if(festas.length){
      msg+='\n🥁 Festas do mês:\n';
      festas.forEach(r=>msg+=`  • ${r[2]} — ${r[0]}\n`);
    }
  }
  SpreadsheetApp.getUi().alert(msg);
}

function _dataFormatada_(d) {
  if(!d)return'—';
  try{return Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd/MM');}catch(e){return'—';}
}

function verObrigacoes() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const aba=ss.getSheetByName(ABA_FILHOS);
  if(!aba){SpreadsheetApp.getUi().alert('Aba Filhos de Santo não encontrada.');return;}
  const hoje=new Date(), limite=new Date(hoje);
  limite.setDate(limite.getDate()+90);
  const rows=aba.getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
  const prox=rows.filter(r=>{
    if(!r[9])return false;
    const d=new Date(r[9]); return d>=hoje&&d<=limite;
  });
  if(!prox.length){SpreadsheetApp.getUi().alert('✅ Nenhuma obrigação nos próximos 90 dias.');return;}
  prox.sort((a,b)=>new Date(a[9])-new Date(b[9]));
  let msg='⚠️ OBRIGAÇÕES — PRÓXIMOS 90 DIAS\n'+'─'.repeat(30)+'\n';
  prox.forEach(r=>{
    const dias=Math.round((new Date(r[9])-hoje)/(1000*60*60*24));
    msg+=`  • ${r[1]||r[2]} — ${_dataFormatada_(r[9])} (em ${dias} dias)\n`;
  });
  SpreadsheetApp.getUi().alert(msg);
}

function limparFiltros() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  [ABA_ACERVO,ABA_CONSUMIVEIS,ABA_FILHOS,ABA_ENTIDADES,ABA_FINANCEIRO].forEach(nome=>{
    const aba=ss.getSheetByName(nome);
    if(aba){try{const f=aba.getFilter();if(f)f.sort(1,true);}catch(e){}}
  });
  SpreadsheetApp.getUi().alert('✅ Filtros limpos e ordenação restaurada.');
}

// ── REMOVE ABAS PADRÃO ───────────────────────────────────
function _excluirPaginasPadrao(ss) {
  ['Página1','Planilha1','Plan1','Sheet1','Estoque'].forEach(nome=>{
    const s=ss.getSheetByName(nome);
    if(s)try{ss.deleteSheet(s);}catch(e){}
  });
}

// ── POST ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const d=JSON.parse(e.postData.contents);
    if(d.acao==='inserir')              return _inserirAcervo(d);
    if(d.acao==='consumivel-atualizar') return _atualizarConsumivel(d);
    if(d.acao==='consumivel-inserir')   return _inserirConsumivel(d);
    if(d.acao==='filho-inserir')        return _inserirFilho(d);
    if(d.acao==='financeiro-inserir')   return _inserirFinanceiro(d);
    return saida({ok:false,erro:'Ação desconhecida: '+d.acao});
  }catch(err){return saida({ok:false,erro:err.message});}
}

function _inserirAcervo(d) {
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if(!aba)return saida({ok:false,erro:'Aba Acervo não encontrada.'});
  const id='ACE-'+Utilities.getUuid().substring(0,6).toUpperCase();
  const lin=aba.getLastRow()+1;
  aba.appendRow([id,d.nome,d.categoria,d.subcategoria||'',d.orixa||'',
    d.status||'Disponível',d.local||'',Number(d.quantidade)||1,
    d.foto||'',d.observacoes||'',_dataFormatada()]);
  aba.getRange(lin,1,1,11).setBackground('#fffdf8').setFontColor('#2a1a0e')
    .setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  [1,8,11].forEach(c=>aba.getRange(lin,c).setHorizontalAlignment('center'));
  aba.setRowHeight(lin,26);
  return saida({ok:true,id});
}

function _atualizarConsumivel(d) {
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba)return saida({ok:false,erro:'Aba Consumíveis não encontrada.'});
  const rows=aba.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    const matchId=d.id&&rows[i][0]===d.id;
    const matchNome=d.nome&&(rows[i][2]||'').toLowerCase().trim()===d.nome.toLowerCase().trim();
    if(matchId||matchNome){
      const lin=i+1,atual=Number(d.atual)||0,minimo=Number(d.minimo||rows[i][5])||1;
      const {label,pct}=_nivelEstoque(atual,minimo);
      let unitCalc=d.precoUnit||rows[i][9]||'';
      if(!unitCalc&&d.precoPacote&&d.qtdPacote>1)unitCalc=(Number(d.precoPacote)/Number(d.qtdPacote)).toFixed(2);
      aba.getRange(lin,5).setValue(atual); aba.getRange(lin,6).setValue(minimo);
      aba.getRange(lin,7).setValue(pct+'%'); aba.getRange(lin,8).setValue(label);
      if(d.fornecedor)aba.getRange(lin,9).setValue(d.fornecedor);
      if(unitCalc)aba.getRange(lin,10).setValue(unitCalc);
      if(d.qtdPacote)aba.getRange(lin,11).setValue(d.qtdPacote);
      if(d.precoPacote)aba.getRange(lin,12).setValue(d.precoPacote);
      if(d.link)aba.getRange(lin,13).setValue(d.link);
      aba.getRange(lin,14).setValue(_dataFormatada());
      return saida({ok:true,acao:'atualizado',nivel:label,pct});
    }
  }
  return saida({ok:false,erro:'Item não encontrado.'});
}

function _inserirConsumivel(d) {
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba)return saida({ok:false,erro:'Aba Consumíveis não encontrada.'});
  const atual=Number(d.atual)||0,minimo=Number(d.minimo)||0;
  const {label,pct}=_nivelEstoque(atual,minimo);
  const qtdPac=Number(d.qtdPacote)||1,precoPac=d.precoPacote||'';
  let unitCalc=d.precoUnit||'';
  if(!unitCalc&&precoPac&&qtdPac>1)unitCalc=(Number(precoPac)/qtdPac).toFixed(2);
  const idx=aba.getLastRow()-1,id='CSM-'+String(idx+1).padStart(3,'0');
  const lin=aba.getLastRow()+1;
  aba.appendRow([id,d.categoria||'',d.nome,d.unidade||'unidade',
    atual,minimo,pct+'%',label,d.fornecedor||'',unitCalc,qtdPac,precoPac,d.link||'',_dataFormatada()]);
  aba.getRange(lin,1,1,14).setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('left');
  [1,4,5,6,7,10,11,12].forEach(c=>aba.getRange(lin,c).setHorizontalAlignment('center'));
  aba.setRowHeight(lin,26);
  return saida({ok:true,id,acao:'inserido',nivel:label,pct});
}

function _inserirFilho(d) {
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FILHOS);
  if(!aba)return saida({ok:false,erro:'Aba Filhos de Santo não encontrada.'});
  const id='FLH-'+Utilities.getUuid().substring(0,6).toUpperCase();
  const lin=aba.getLastRow()+1;
  aba.appendRow([id,d.nomeSanto||'',d.nomeSocial||'',d.contato||'',
    d.aniversario||'',d.feitura||'',d.orixa||'',d.adjunto||'',
    d.nacao||'',d.proximaObrigacao||'',d.obrigacoesRealizadas||'',
    d.statusTerreiro||'Ativo',d.observacoes||'',_dataFormatada()]);
  aba.getRange(lin,1,1,14).setFontSize(10).setVerticalAlignment('middle');
  aba.setRowHeight(lin,26);
  return saida({ok:true,id});
}

function _inserirFinanceiro(d) {
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FINANCEIRO);
  if(!aba)return saida({ok:false,erro:'Aba Financeiro não encontrada.'});
  const id=(d.tipo==='Entrada'?'ENT':'SAI')+'-'+Utilities.getUuid().substring(0,6).toUpperCase();
  const lin=aba.getLastRow()+1;
  aba.appendRow([id,d.data||_dataFormatada(),d.tipo,d.categoria,
    d.descricao||'',Number(d.valor)||0,d.responsavel||'',d.comprovante||'',d.observacoes||'']);
  aba.getRange(lin,1,1,9).setFontSize(10).setVerticalAlignment('middle');
  aba.getRange(lin,6).setNumberFormat('R$ #,##0.00');
  [1,2,3,5,6].forEach(c=>aba.getRange(lin,c).setHorizontalAlignment('center'));
  aba.setRowHeight(lin,26);
  return saida({ok:true,id});
}

// ── GET ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const acao=(e.parameter||{}).acao;
    if(acao==='listar')             return _listarAcervo();
    if(acao==='consumiveis-listar') return _listarConsumiveis();
    if(acao==='filhos-listar')      return _listarFilhos();
    if(acao==='financeiro-resumo')  return _resumoFinanceiro();
    if(acao==='datas-mes')          return _datasDoMes();
    if(acao==='ml-buscar')          return _buscarML(e.parameter.q);
    return saida({ok:true,msg:'Acervo do Terreiro v6 — Web App ativo.'});
  }catch(err){return saida({ok:false,erro:err.message});}
}

function _listarAcervo(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if(!aba)return saida({ok:false,erro:'Aba Acervo não encontrada.'});
  const [,...rows]=aba.getDataRange().getValues();
  return saida({ok:true,itens:rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],subcategoria:l[3],orixa:l[4],
    status:l[5],local:l[6],quantidade:l[7],foto:l[8],observacoes:l[9],dataCadastro:l[10]
  }))});
}

function _listarConsumiveis(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba)return saida({ok:false,erro:'Aba Consumíveis não encontrada.'});
  const [,...rows]=aba.getDataRange().getValues();
  return saida({ok:true,itens:rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],categoria:l[1],nome:l[2],unidade:l[3],atual:l[4],minimo:l[5],
    pct:l[6],nivel:l[7],fornecedor:l[8],precoUnit:l[9],qtdPacote:l[10],
    precoPacote:l[11],link:l[12],atualizado:l[13]
  }))});
}

function _listarFilhos(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FILHOS);
  if(!aba)return saida({ok:false,erro:'Aba Filhos de Santo não encontrada.'});
  const [,...rows]=aba.getDataRange().getValues();
  // Segredo de santo: exibe só o orixá, sem nenhum nome identificador
  return saida({ok:true,itens:rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],
    orixa:l[6],
    nacao:l[8],
    feitura:_dataFormatada_(l[5]),
    proximaObrigacao:_dataFormatada_(l[9]),
    statusTerreiro:l[11]
  }))});
}

function _resumoFinanceiro(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FINANCEIRO);
  if(!aba)return saida({ok:false,erro:'Aba Financeiro não encontrada.'});
  const [,...rows]=aba.getDataRange().getValues().slice(1);
  const hoje=new Date(),mes=hoje.getMonth(),ano=hoje.getFullYear();
  let totalEntradas=0,totalSaidas=0,entradasMes=0,saidasMes=0;
  const porCategoria={};
  rows.filter(r=>r[0]!=='').forEach(r=>{
    const val=Number(r[5])||0,tipo=r[2];
    const d=new Date(r[1]),eMes=d.getMonth()===mes&&d.getFullYear()===ano;
    if(tipo==='Entrada'){totalEntradas+=val;if(eMes)entradasMes+=val;}
    else{totalSaidas+=val;if(eMes)saidasMes+=val;}
    const cat=r[3]||'Outros';
    porCategoria[cat]=(porCategoria[cat]||0)+(tipo==='Saída'?val:0);
  });
  return saida({ok:true,
    saldo:totalEntradas-totalSaidas,
    totalEntradas,totalSaidas,
    entradasMes,saidasMes,saldoMes:entradasMes-saidasMes,
    porCategoria
  });
}

function _datasDoMes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hoje=new Date(),mes=hoje.getMonth(),ano=hoje.getFullYear();
  const result={aniversariantes:[],festas:[],obrigacoes:[]};
  const abaF=ss.getSheetByName(ABA_FILHOS);
  if(abaF&&abaF.getLastRow()>1){
    const rows=abaF.getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
    rows.forEach(r=>{
      if(r[4]){const d=new Date(r[4]);if(d.getMonth()===mes)result.aniversariantes.push({nome:r[1]||r[2],data:_dataFormatada_(r[4]),orixa:r[6]});}
      if(r[9]){const d=new Date(r[9]);if(d.getMonth()===mes&&d.getFullYear()===ano)result.obrigacoes.push({nome:r[1]||r[2],data:_dataFormatada_(r[9])});}
    });
  }
  const abaE=ss.getSheetByName(ABA_ENTIDADES);
  if(abaE&&abaE.getLastRow()>1){
    const rows=abaE.getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
    rows.forEach(r=>{
      if(r[2]&&r[2]!=='—'){
        const parts=String(r[2]).split('/');
        if(parts.length>=2&&(parseInt(parts[1])-1)===mes)
          result.festas.push({entidade:r[0],data:r[2],saudacao:r[8]||''});
      }
    });
  }
  return saida({ok:true,...result});
}

function _buscarML(query){
  try{
    const q=encodeURIComponent(query||'');
    const resp=UrlFetchApp.fetch(`https://api.mercadolibre.com/sites/MLB/search?q=${q}&limit=5`,{muteHttpExceptions:true});
    if(resp.getResponseCode()!==200)return saida({ok:false,erro:'ML indisponível'});
    const data=JSON.parse(resp.getContentText());
    const results=(data.results||[]).slice(0,5).map(r=>({
      titulo:r.title,preco:r.price,link:r.permalink,thumbnail:r.thumbnail,
      vendedor:(r.seller||{}).nickname||'',qtdPacote:_detectarQuantidade(r.title),
    }));
    return saida({ok:true,resultados:results});
  }catch(err){return saida({ok:false,erro:err.message});}
}

function _detectarQuantidade(titulo){
  const m=titulo.match(/(\d+)\s*(un|unid|pç|pc|peças|velas|charuto|cigarros)/i);
  return m?parseInt(m[1]):1;
}

function saida(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
