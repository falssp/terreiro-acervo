// ================================================================
//  GESTÃO DO TERREIRO — Ile Ase Vodun Ogum Ayres — Apps Script v12
//  Corrigido: filtros até linha 201, colunas extras removidas,
//  dropdown Nível em Consumíveis, Admins sem filtro
// ================================================================

const ABA_ACERVO      = 'Acervo';
const ABA_CONSUMIVEIS = 'Consumíveis';
const ABA_FILHOS      = 'Filhos de Santo';
const ABA_ENTIDADES   = 'Entidades e Orixás';
const ABA_FINANCEIRO  = 'Financeiro';
const ABA_ADMINS      = 'Admins';

const SALT    = 'ile_ase_salt_v9_2024';
const DEV_KEY = 'ile_ase_dev_2024_falsp';

const PERMISSOES_TODAS = ['estoque','datas','obrigacoes','financeiro','mailing','filhos','configuracoes'];

const ORIXA_LIST         = ['Iansã','Iemanjá','Logun Edé','Nanã','Obá','Obaluaê / Omolu','Ogum','Oxaguiã / Oxalufã','Oxalá','Oxossi','Oxum','Oxumaré','Xangô','Outro'];
const NACAO_LIST         = ['Angola','Efon','Ijexá','Jeje','Ketu','Nagô','Omolokô','Umbanda','Outra'];
const STATUS_FILHO_LIST  = ['Abiã','Ativo','Ebomi','Iaô','Inativo','Ogã / Ekedi','Suspenso'];
const CAT_CONSUMIVEL_LIST= ['Alimentos','Bebidas','Ervas e Defumação','Flores e Naturais','Fumo','Velas','Outro'];
const NIVEL_LIST         = ['— Sem mínimo','✅ OK','⚠️ Repor','🔴 Alerta','🚨 Urgente / Zero'];
const CAT_ACERVO_LIST   = ['Alimentos e Oferendas','Ervas e Plantas','Ferramentas e Objetos Rituais','Roupas e Indumentárias','Outro'];
const STATUS_ACERVO_LIST = ['Danificado','Disponível','Em Uso','Necessita Reposição','N/A'];
const ENTIDADE_LIST      = ['Baiano / Baiana','Boiadeiro / Boiadeira','Caboclo / Cabocla','Cigano / Cigana','Criança / Erê','Exu','Marinheiro / Marinheira','Ogum Beira-Mar','Pomba-Gira','Preto-Velho / Preta-Velha','Zé Pilintra','Outro'];
const CAT_FINANCEIRO_LIST= ['Consumíveis','Dízimo / Doação','Evento / Gira','Indumentária / Acervo','Manutenção','Ritual / Oferenda','Venda','Outros'];

// ── MENU ─────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏛️ Gestão do Terreiro')
    .addItem('⚙️ Configurar planilha', 'setup')
    .addSeparator()
    .addItem('📅 Datas do mês atual', 'verDatasMes')
    .addItem('🛒 Lista de compras', 'gerarListaCompras')
    .addItem('⚠️ Obrigações próximas (180 dias)', 'verObrigacoes')
    .addItem('🔄 Recalcular níveis de estoque', 'recalcularEstoque')
    .addToUi();
}

// ── SETUP ────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const criadas = [];
  if (!ss.getSheetByName(ABA_ACERVO))      { _criarAcervo(ss);      criadas.push(ABA_ACERVO); }
  if (!ss.getSheetByName(ABA_CONSUMIVEIS)) { _criarConsumiveis(ss); criadas.push(ABA_CONSUMIVEIS); }
  if (!ss.getSheetByName(ABA_FILHOS))      { _criarFilhos(ss);      criadas.push(ABA_FILHOS); }
  if (!ss.getSheetByName(ABA_ENTIDADES))   { _criarEntidades(ss);   criadas.push(ABA_ENTIDADES); }
  if (!ss.getSheetByName(ABA_FINANCEIRO))  { _criarFinanceiro(ss);  criadas.push(ABA_FINANCEIRO); }
  if (!ss.getSheetByName(ABA_ADMINS))      { _criarAdmins(ss);      criadas.push(ABA_ADMINS); }
  _excluirPadrao(ss);
  SpreadsheetApp.getUi().alert(
    criadas.length === 0
      ? '✅ Todas as abas já existem.\nNenhuma alteração foi feita na estrutura.'
      : '✅ Criadas: ' + criadas.join(', ') +
        '\n\nAcesse admin.html para criar o primeiro usuário (Pai de Santo).'
  );
}

// ── HELPERS DE ESTRUTURA ─────────────────────────────────
function _cabecalho(aba, cols, bg, fg) {
  aba.getRange(1, 1, 1, cols.length)
    .setValues([cols])
    .setBackground(bg).setFontColor(fg)
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(false);
  aba.setRowHeight(1, 34);
  aba.setFrozenRows(1);
}

function _validacao(aba, cel, lista) {
  aba.getRange(cel).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(lista, true)
      .setAllowInvalid(false)
      .build()
  );
}

function _formatCond(aba, range, regras) {
  aba.setConditionalFormatRules(
    regras.map(({ v, bg, f }) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(v)
        .setBackground(bg).setFontColor(f)
        .setRanges([aba.getRange(range)])
        .build()
    )
  );
}

// Remove colunas além do necessário e aplica filtro na range correta
function _finalizarAba(aba, numCols, numLinhasDados) {
  const totalCols = aba.getMaxColumns();
  if (totalCols > numCols) {
    aba.deleteColumns(numCols + 1, totalCols - numCols);
  }
  // Filtro cobre cabeçalho + área de dados
  aba.getRange(1, 1, numLinhasDados + 1, numCols).createFilter();
}

function _novaAba(ss, nome) { return ss.insertSheet(nome); }

function _excluirPadrao(ss) {
  const ABAS_VALIDAS = [ABA_ACERVO, ABA_CONSUMIVEIS, ABA_FILHOS, ABA_ENTIDADES, ABA_FINANCEIRO, ABA_ADMINS];
  ss.getSheets().forEach(s => {
    if (!ABAS_VALIDAS.includes(s.getName())) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });
}

// ── ABA ADMINS ───────────────────────────────────────────
function _criarAdmins(ss) {
  const aba = ss.insertSheet(ABA_ADMINS);
  _cabecalho(aba,
    ['E-mail','Senha','Nome','Nível','Permissões','Tentativas','Bloqueado até','Último acesso'],
    '#1a0a2a','#e0c8ff');
  aba.setColumnWidth(1, 220);
  aba.setColumnWidth(2, 280);
  aba.setColumnWidth(5, 300);
  // Remove colunas extras (8 colunas usadas)
  const totalCols = aba.getMaxColumns();
  if (totalCols > 8) aba.deleteColumns(9, totalCols - 8);
}

// ── AUTENTICAÇÃO ─────────────────────────────────────────
function _hash(texto) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    texto + SALT,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function _adminsVazio() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  if (!aba) return true;
  return aba.getLastRow() <= 1;
}

function _autenticar(email, senha) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_ADMINS);
  if (!aba) return { ok: false, erro: 'Sistema não configurado. Execute o setup.' };
  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0]||'').toLowerCase().trim() !== email.toLowerCase().trim()) continue;
    const bloqAte = rows[i][6];
    if (bloqAte) {
      try { if (new Date(bloqAte) > new Date()) return { ok: false, erro: 'Acesso bloqueado temporariamente. Aguarde 30 minutos ou solicite desbloqueio ao Pai de Santo.' }; } catch(e) {}
    }
    if (_hash(senha) !== rows[i][1]) {
      const tentativas = (Number(rows[i][5]) || 0) + 1;
      aba.getRange(i+1, 6).setValue(tentativas);
      if (tentativas >= 3) {
        const bloq = new Date(); bloq.setMinutes(bloq.getMinutes() + 30);
        aba.getRange(i+1, 7).setValue(bloq.toISOString());
        aba.getRange(i+1, 6).setValue(0);
        return { ok: false, erro: 'Acesso bloqueado por 30 minutos após 3 tentativas incorretas.' };
      }
      return { ok: false, erro: 'E-mail ou senha incorretos. (' + (3 - tentativas) + ' tentativa(s) restante(s))' };
    }
    aba.getRange(i+1, 6).setValue(0);
    aba.getRange(i+1, 7).setValue('');
    aba.getRange(i+1, 8).setValue(_dataFormatada());
    const permissoes = (rows[i][4] || '').split(',').map(p => p.trim()).filter(Boolean);
    return { ok: true, token: _gerarToken(email, permissoes), nome: rows[i][2], nivel: rows[i][3], permissoes };
  }
  return { ok: false, erro: 'E-mail ou senha incorretos.' };
}

function _criarPrimeirAdmin(email, senha, nome) {
  if (!_adminsVazio()) return { ok: false, erro: 'Já existe um administrador cadastrado.' };
  if (!email || !email.includes('@')) return { ok: false, erro: 'E-mail inválido.' };
  if (!senha || senha.length < 6) return { ok: false, erro: 'A senha precisa ter pelo menos 6 caracteres.' };
  if (!nome || nome.trim().length < 2) return { ok: false, erro: 'Informe um nome.' };
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  aba.appendRow([email.toLowerCase().trim(), _hash(senha), nome.trim(), 'Pai de Santo', PERMISSOES_TODAS.join(','), 0, '', _dataFormatada()]);
  return _autenticar(email, senha);
}

function _gerarToken(email, permissoes) {
  const token = Utilities.getUuid();
  const expira = new Date(); expira.setHours(expira.getHours() + 8);
  PropertiesService.getScriptProperties().setProperty('tok_' + token, JSON.stringify({ email, permissoes, expira: expira.toISOString() }));
  return token;
}

function _validarToken(token) {
  if (!token) return null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('tok_' + token);
    if (!raw) return null;
    const d = JSON.parse(raw);
    if (new Date(d.expira) < new Date()) { PropertiesService.getScriptProperties().deleteProperty('tok_' + token); return null; }
    return d;
  } catch(e) { return null; }
}

function _revogarToken(token) { if (token) PropertiesService.getScriptProperties().deleteProperty('tok_' + token); }
function _isDev(devkey) { return devkey === DEV_KEY; }

// ── POST ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.acao === 'inserir')              return _inserirAcervo(d);
    if (d.acao === 'consumivel-atualizar') return _atualizarConsumivel(d);
    if (d.acao === 'consumivel-inserir')   return _inserirConsumivel(d);
    if (d.acao === 'filho-inserir')        return _inserirFilho(d);
    if (d.acao === 'financeiro-inserir')   return _inserirFinanceiro(d);
    if (d.acao === 'login')                return saida(_autenticar(d.email, d.senha));
    if (d.acao === 'logout')               { _revogarToken(d.token); return saida({ ok: true }); }
    if (d.acao === 'primeiro-admin')       return saida(_criarPrimeirAdmin(d.email, d.senha, d.nome));
    if (d.acao === 'trocar-senha')         return _trocarSenha(d);
    const sessao = _validarToken(d.token);
    if (!sessao) return saida({ ok: false, erro: 'Sessão expirada.', code: 401 });
    const tem = p => sessao.permissoes.includes(p);
    if (d.acao === 'admin-criar')       { if (!tem('configuracoes')) return saida({ ok: false, erro: 'Sem permissão.', code: 403 }); return _criarAdmin(d); }
    if (d.acao === 'admin-desbloquear') { if (!tem('configuracoes')) return saida({ ok: false, erro: 'Sem permissão.', code: 403 }); return _desbloquearAdmin(d.email); }
    return saida({ ok: false, erro: 'Ação desconhecida: ' + d.acao });
  } catch(err) { return saida({ ok: false, erro: err.message }); }
}

// ── GET ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter || {};
    const acao = p.acao;
    if (_isDev(p.devkey)) {
      if (acao === 'admins-raw') return _listarAdminsRaw();
      return saida({ ok: true, msg: 'Dev access OK', adminsVazio: _adminsVazio() });
    }
    if (acao === 'listar')             return _listarAcervo();
    if (acao === 'consumiveis-listar') return _listarConsumiveis();
    if (acao === 'filhos-listar')      return _listarFilhos();
    if (acao === 'datas-mes')          return _datasDoMes();
    if (acao === 'ml-buscar')          return _buscarML(p.q);
    if (acao === 'admins-vazio')       return saida({ ok: true, vazio: _adminsVazio() });
    const sessao = _validarToken(p.token);
    if (!sessao) return saida({ ok: false, erro: 'Não autorizado.', code: 401 });
    const tem = perm => sessao.permissoes.includes(perm);
    if (acao === 'financeiro-resumo') { if (!tem('financeiro')) return saida({ ok: false, erro: 'Sem permissão.', code: 403 }); return _resumoFinanceiro(); }
    if (acao === 'admins-listar')     { if (!tem('configuracoes')) return saida({ ok: false, erro: 'Sem permissão.', code: 403 }); return _listarAdmins(); }
    return saida({ ok: true, msg: 'Gestão do Terreiro — Ile Ase Vodun Ogum Ayres — v12' });
  } catch(err) { return saida({ ok: false, erro: err.message }); }
}

// ── GESTÃO DE ADMINS ─────────────────────────────────────
function _criarAdmin(d) {
  if (!d.email || !d.email.includes('@')) return saida({ ok: false, erro: 'E-mail inválido.' });
  if (!d.senha || d.senha.length < 6)     return saida({ ok: false, erro: 'Senha muito curta (mín. 6 caracteres).' });
  if (!d.nome)                             return saida({ ok: false, erro: 'Informe o nome.' });
  const aba  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0]||'').toLowerCase() === d.email.toLowerCase()) return saida({ ok: false, erro: 'E-mail já cadastrado.' });
  }
  aba.appendRow([d.email.toLowerCase().trim(), _hash(d.senha), d.nome.trim(), d.nivel || 'Admin', (d.permissoes || []).join(','), 0, '', _dataFormatada()]);
  return saida({ ok: true });
}

function _desbloquearAdmin(email) {
  const aba  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0]||'').toLowerCase() === (email||'').toLowerCase()) {
      aba.getRange(i+1, 6).setValue(0); aba.getRange(i+1, 7).setValue(''); return saida({ ok: true });
    }
  }
  return saida({ ok: false, erro: 'Admin não encontrado.' });
}

function _trocarSenha(d) {
  const sessao = _validarToken(d.token);
  if (!sessao) return saida({ ok: false, erro: 'Sessão expirada.' });
  if (!d.senhaAtual || !d.senhaNova) return saida({ ok: false, erro: 'Preencha todos os campos.' });
  if (d.senhaNova.length < 6) return saida({ ok: false, erro: 'A nova senha precisa ter pelo menos 6 caracteres.' });
  const aba  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  const rows = aba.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0]||'').toLowerCase() === sessao.email.toLowerCase()) {
      if (_hash(d.senhaAtual) !== rows[i][1]) return saida({ ok: false, erro: 'Senha atual incorreta.' });
      aba.getRange(i+1, 2).setValue(_hash(d.senhaNova)); return saida({ ok: true });
    }
  }
  return saida({ ok: false, erro: 'Usuário não encontrado.' });
}

function _listarAdmins() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  if (!aba) return saida({ ok: false, erro: 'Aba Admins não encontrada.' });
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ ok: true, admins: rows.filter(l => l[0] !== '').map(l => ({
    email: l[0], nome: l[2], nivel: l[3],
    permissoes: (l[4]||'').split(',').map(p=>p.trim()).filter(Boolean),
    tentativas: l[5], bloqueado: !!l[6] && new Date(l[6]) > new Date(),
    bloqAte: l[6] ? _dataFormatada_(l[6]) : '', ultimoAcesso: l[7]
  }))});
}

function _listarAdminsRaw() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ADMINS);
  if (!aba) return saida({ ok: false, erro: 'Aba não encontrada.' });
  return saida({ ok: true, total: aba.getLastRow() - 1 });
}

// ── LISTAGENS ────────────────────────────────────────────
function _listarAcervo() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if (!aba) return saida({ ok: false, erro: 'Aba Acervo não encontrada.' });
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],nome:l[1],categoria:l[2],subcategoria:l[3],orixa:l[4],entidade:l[5],
    status:l[6],local:l[7],quantidade:l[8],foto:l[9],observacoes:l[10],dataCadastro:l[11]
  }))});
}

function _listarConsumiveis() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if (!aba) return saida({ ok: false, erro: 'Aba Consumíveis não encontrada.' });
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0],categoria:l[1],nome:l[2],unidade:l[3],atual:l[4],minimo:l[5],
    pct:l[6],nivel:l[7],fornecedor:l[8],precoUnit:l[9],qtdPacote:l[10],
    precoPacote:l[11],link:l[12],atualizado:l[13]
  }))});
}

function _listarFilhos() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FILHOS);
  if (!aba) return saida({ ok: false, erro: 'Aba Filhos não encontrada.' });
  const [,...rows] = aba.getDataRange().getValues();
  return saida({ ok: true, itens: rows.filter(l=>l[0]!=='').map(l=>({
    id:l[0], orixa:l[6], nacao:l[8],
    feitura:_dataFormatada_(l[5]),
    proximaObrigacao:_dataFormatada_(l[9]),
    statusTerreiro:l[11]
  }))});
}

function _resumoFinanceiro() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FINANCEIRO);
  if (!aba) return saida({ ok: false, erro: 'Aba Financeiro não encontrada.' });
  const hoje=new Date(),mes=hoje.getMonth(),ano=hoje.getFullYear();
  let tE=0,tS=0,eM=0,sM=0; const pc={};
  aba.getDataRange().getValues().slice(1).filter(r=>r[0]!=='').forEach(r=>{
    const v=Number(r[5])||0,t=r[2]; let em=false;
    try{const d=new Date(r[1]);em=d.getMonth()===mes&&d.getFullYear()===ano;}catch(e){}
    if(t==='Entrada'){tE+=v;if(em)eM+=v;}else{tS+=v;if(em)sM+=v;}
    pc[r[3]||'Outros']=(pc[r[3]||'Outros']||0)+(t==='Saída'?v:0);
  });
  return saida({ok:true,saldo:tE-tS,totalEntradas:tE,totalSaidas:tS,entradasMes:eM,saidasMes:sM,saldoMes:eM-sM,porCategoria:pc});
}

function _datasDoMes() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),hoje=new Date(),mes=hoje.getMonth(),ano=hoje.getFullYear();
  const r={aniversariantes:[],festas:[],obrigacoes:[]};
  const aF=ss.getSheetByName(ABA_FILHOS);
  if(aF&&aF.getLastRow()>1){
    aF.getDataRange().getValues().slice(1).filter(l=>l[0]!=='').forEach(l=>{
      try{if(l[4]&&new Date(l[4]).getMonth()===mes)r.aniversariantes.push({nome:l[1]||l[2],data:_dataFormatada_(l[4]),orixa:l[6]});}catch(e){}
      try{const d=new Date(l[9]);if(l[9]&&d.getMonth()===mes&&d.getFullYear()===ano)r.obrigacoes.push({nome:l[1]||l[2],data:_dataFormatada_(l[9])});}catch(e){}
    });
  }
  const aE=ss.getSheetByName(ABA_ENTIDADES);
  if(aE&&aE.getLastRow()>1){
    aE.getDataRange().getValues().slice(1).filter(l=>l[0]!=='').forEach(l=>{
      if(l[2]&&l[2]!==''){const p=String(l[2]).split('/');if(p.length>=2&&(parseInt(p[1])-1)===mes)r.festas.push({entidade:l[0],data:l[2],saudacao:l[8]||''});}
    });
  }
  return saida({ok:true,...r});
}

function _buscarML(q) {
  try{
    const resp=UrlFetchApp.fetch('https://api.mercadolibre.com/sites/MLB/search?q='+encodeURIComponent(q||'')+'&limit=5',{muteHttpExceptions:true});
    if(resp.getResponseCode()!==200)return saida({ok:false,erro:'ML indisponível'});
    const data=JSON.parse(resp.getContentText());
    return saida({ok:true,resultados:(data.results||[]).slice(0,5).map(r=>({
      titulo:r.title,preco:r.price,link:r.permalink,thumbnail:r.thumbnail,
      vendedor:(r.seller||{}).nickname||'',
      qtdPacote:parseInt((r.title.match(/(\d+)\s*(un|unid|pç|pc|peças|velas)/i)||[,1])[1])||1
    }))});
  }catch(e){return saida({ok:false,erro:e.message});}
}

// ── INSERÇÕES ────────────────────────────────────────────
function _inserirAcervo(d){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACERVO);
  if(!aba)return saida({ok:false,erro:'Aba Acervo não encontrada.'});
  const id='ACE-'+Utilities.getUuid().substring(0,6).toUpperCase();
  aba.appendRow([id,d.nome,d.categoria,d.subcategoria||'',d.orixa||'',d.entidade||'',d.status||'Disponível',d.local||'',Number(d.quantidade)||1,d.foto||'',d.observacoes||'',_dataFormatada()]);
  return saida({ok:true,id});
}

function _atualizarConsumivel(d){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba)return saida({ok:false,erro:'Aba não encontrada.'});
  const rows=aba.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    const mId=d.id&&rows[i][0]===d.id,mNome=d.nome&&(rows[i][2]||'').toLowerCase().trim()===d.nome.toLowerCase().trim();
    if(mId||mNome){
      const lin=i+1,a=Number(d.atual)||0,m=Number(d.minimo||rows[i][5])||1;
      const{label,pct}=_nivelEstoque(a,m);
      let u=d.precoUnit||rows[i][9]||'';
      if(!u&&d.precoPacote&&Number(d.qtdPacote)>1)u=(Number(d.precoPacote)/Number(d.qtdPacote)).toFixed(2);
      aba.getRange(lin,5).setValue(a);aba.getRange(lin,6).setValue(m);
      aba.getRange(lin,7).setValue(pct+'%');aba.getRange(lin,8).setValue(label);
      if(d.fornecedor)aba.getRange(lin,9).setValue(d.fornecedor);
      if(u)aba.getRange(lin,10).setValue(u);
      if(d.qtdPacote)aba.getRange(lin,11).setValue(d.qtdPacote);
      if(d.precoPacote)aba.getRange(lin,12).setValue(d.precoPacote);
      if(d.link)aba.getRange(lin,13).setValue(d.link);
      aba.getRange(lin,14).setValue(_dataFormatada());
      return saida({ok:true,acao:'atualizado',nivel:label,pct});
    }
  }
  return saida({ok:false,erro:'Item não encontrado.'});
}

function _inserirConsumivel(d){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba)return saida({ok:false,erro:'Aba não encontrada.'});
  const a=Number(d.atual)||0,m=Number(d.minimo)||0;
  const{label,pct}=_nivelEstoque(a,m);
  const qP=Number(d.qtdPacote)||1,pP=d.precoPacote||'';
  let u=d.precoUnit||'';if(!u&&pP&&qP>1)u=(Number(pP)/qP).toFixed(2);
  aba.appendRow(['CSM-'+String(aba.getLastRow()).padStart(3,'0'),d.categoria||'',d.nome,d.unidade||'unidade',a,m,pct+'%',label,d.fornecedor||'',u,qP,pP,d.link||'',_dataFormatada()]);
  return saida({ok:true,acao:'inserido',nivel:label,pct});
}

function _inserirFilho(d){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FILHOS);
  if(!aba)return saida({ok:false,erro:'Aba não encontrada.'});
  const id='FLH-'+Utilities.getUuid().substring(0,6).toUpperCase();
  aba.appendRow([id,d.nomeSanto||'',d.nomeSocial||'',d.contato||'',d.aniversario||'',d.feitura||'',d.orixa||'',d.adjunto||'',d.nacao||'',d.proximaObrigacao||'',d.obrigacoesRealizadas||'',d.statusTerreiro||'Ativo',d.observacoes||'',_dataFormatada()]);
  return saida({ok:true,id});
}

function _inserirFinanceiro(d){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FINANCEIRO);
  if(!aba)return saida({ok:false,erro:'Aba não encontrada.'});
  const id=(d.tipo==='Entrada'?'ENT':'SAI')+'-'+Utilities.getUuid().substring(0,6).toUpperCase();
  aba.appendRow([id,d.data||_dataFormatada(),d.tipo,d.categoria,d.descricao||'',Number(d.valor)||0,d.responsavel||'',d.comprovante||'',d.observacoes||'']);
  return saida({ok:true,id});
}

// ── CRIAÇÃO DE ABAS ──────────────────────────────────────
function _criarAcervo(ss){
  const aba=_novaAba(ss,ABA_ACERVO);
  const COLS = ['ID','Nome','Categoria','Subcategoria','Orixá de Cabeça','Entidade / Linha','Status','Localização / Armário','Qtd','Foto (URL)','Observações','Data cadastro'];
  _cabecalho(aba, COLS, '#2c1a10','#f0d090');
  _validacao(aba,'C2:C200', CAT_ACERVO_LIST);
  _validacao(aba,'E2:E200', ORIXA_LIST);
  _validacao(aba,'F2:F200', ENTIDADE_LIST);
  _validacao(aba,'G2:G200', STATUS_ACERVO_LIST);
  _formatCond(aba,'G2:G200',[
    {v:'Disponível',      bg:'#e6f4ea',f:'#1e6b3a'},
    {v:'Em Uso',          bg:'#e8f0fe',f:'#1a56a0'},
    {v:'Danificado',      bg:'#fef3e2',f:'#7a3800'},
    {v:'Necessita Reposição',bg:'#fdecea',f:'#8b0000'}
  ]);
  _finalizarAba(aba, COLS.length, 200);
}

function _criarConsumiveis(ss){
  const aba=_novaAba(ss,ABA_CONSUMIVEIS);
  const COLS = ['ID','Categoria','Item','Unidade','Qtd Atual','Qtd Mínima','% Estoque','Nível','Fornecedor','Preço Unit.','Qtd/Pac','Preço Pacote','Link Compra','Atualizado em'];
  _cabecalho(aba, COLS, '#1a1a3a','#c8d8f0');
  _validacao(aba,'B2:B200', CAT_CONSUMIVEL_LIST);
  _validacao(aba,'H2:H200', NIVEL_LIST);
  _formatCond(aba,'H2:H200',[
    {v:'✅ OK',             bg:'#e6f4ea',f:'#1e6b3a'},
    {v:'⚠️ Repor',         bg:'#fef9e7',f:'#7d5c00'},
    {v:'🔴 Alerta',        bg:'#fdecea',f:'#8b0000'},
    {v:'🚨 Urgente / Zero',bg:'#f5c6cb',f:'#5c0000'}
  ]);
  const dt=_dataFormatada();
  [
    ['Velas','Vela 7 dias branca','unidade',0,10],
    ['Velas','Vela 7 dias vermelha','unidade',0,6],
    ['Velas','Vela 7 dias amarela','unidade',0,6],
    ['Velas','Vela 7 dias azul','unidade',0,6],
    ['Velas','Vela 7 dias verde','unidade',0,6],
    ['Velas','Vela 7 dias preta','unidade',0,4],
    ['Velas','Vela votiva branca','pacote',0,2],
    ['Velas','Vela taça branca','unidade',0,6],
    ['Bebidas','Cachaça','garrafa',0,3],
    ['Bebidas','Cerveja preta','unidade',0,6],
    ['Bebidas','Vinho tinto seco','garrafa',0,2],
    ['Bebidas','Mel','pote',0,2],
    ['Bebidas','Azeite de dendê','garrafa',0,1],
    ['Bebidas','Água mineral','garrafa',0,6],
    ['Fumo','Cigarro (maço)','maço',0,3],
    ['Fumo','Charuto','unidade',0,4],
    ['Fumo','Cigarrilha','unidade',0,6],
    ['Fumo','Cachimbo (tabaco)','pacote',0,1],
    ['Ervas e Defumação','Incenso (caixinha)','caixa',0,3],
    ['Ervas e Defumação','Pemba branca','unidade',0,4],
    ['Ervas e Defumação','Pemba colorida','unidade',0,4],
    ['Ervas e Defumação','Ervas para banho (mix)','maço',0,3],
    ['Ervas e Defumação','Erva para defumação','maço',0,2],
    ['Alimentos','Farofa','kg',0,1],
    ['Alimentos','Pipoca','pacote',0,2],
    ['Alimentos','Milho de pipoca','kg',0,1],
    ['Alimentos','Azeite de oliva','garrafa',0,1],
    ['Alimentos','Amendoim','pacote',0,2],
    ['Alimentos','Inhame','unidade',0,4],
    ['Flores e Naturais','Rosas brancas','dúzia',0,1],
    ['Flores e Naturais','Girassol','unidade',0,6],
    ['Flores e Naturais','Flores mistas','buquê',0,1],
    ['Flores e Naturais','Folhas de bananeira','folha',0,4]
  ].forEach(([cat,nome,unid,atual,min],idx)=>{
    const{label,pct}=_nivelEstoque(atual,min);
    aba.getRange(idx+2,1,1,14).setValues([['CSM-'+String(idx+1).padStart(3,'0'),cat,nome,unid,atual,min,pct+'%',label,'','',1,'','',dt]]);
  });
  _finalizarAba(aba, COLS.length, 200);
}

function _criarFilhos(ss){
  const aba=_novaAba(ss,ABA_FILHOS);
  const COLS = ['ID','Nome de Candomblé','Nome Social','Contato','Data de Aniversário','Data de Feitura','Orixá de Cabeça','Adjuntó','Nação','Próxima Obrigação','Obrigações Realizadas','Status no Terreiro','Observações','Cadastrado em'];
  _cabecalho(aba, COLS, '#2a0a2a','#e8c8f0');
  _validacao(aba,'G2:G200', ORIXA_LIST);
  _validacao(aba,'H2:H200', ORIXA_LIST);
  _validacao(aba,'I2:I200', NACAO_LIST);
  _validacao(aba,'L2:L200', STATUS_FILHO_LIST);
  aba.getRange('E2:F200').setNumberFormat('dd/MM/yyyy');
  aba.getRange('J2:J200').setNumberFormat('dd/MM/yyyy');
  _formatCond(aba,'L2:L200',[
    {v:'Ativo',      bg:'#e6f4ea',f:'#1e6b3a'},
    {v:'Ebomi',      bg:'#e8f0fe',f:'#1a56a0'},
    {v:'Suspenso',   bg:'#fdecea',f:'#8b0000'},
    {v:'Inativo',    bg:'#f5f5f5',f:'#666'},
    {v:'Abiã',       bg:'#fff9e6',f:'#7d5c00'},
    {v:'Iaô',        bg:'#f0e6ff',f:'#4a0080'}
  ]);
  _finalizarAba(aba, COLS.length, 200);
}

function _criarEntidades(ss){
  const aba=_novaAba(ss,ABA_ENTIDADES);
  const COLS = ['Orixá / Entidade','Nação / Qualidade','Data da Festa','Dia da Semana','Cores','Oferendas Preferidas','Bebidas / Alimentos','Itens do Acervo','Saudação','Observações Rituais'];
  _cabecalho(aba, COLS, '#1a2a1a','#c8f0c8');
  aba.getRange('C2:C50').setNumberFormat('dd/MM');
  const dados=[
    ['Exu','Todas','','Segunda','Preto e vermelho','Pimenta, dendê, farofa, cachaça','Cachaça, vinho tinto','Ogó, tridentes, sete chaves','Laroyê Exu!','Guardião das encruzilhadas'],
    ['Ogum','Ketu / Angola','23/04','Terça','Verde e preto','Feijão preto, carne, dendê','Cerveja preta, vinho tinto','Espada, ferramentas de ferro','Ogum Yê!','Orixá do ferro e da guerra'],
    ['Oxossi','Ketu','','Quinta','Azul e verde','Milho branco, inhame, mel','Mel, água de coco','Arco e flecha','Okê Arô!','Caçador, orixá da fartura'],
    ['Xangô','Ketu','04/12','Quarta','Vermelho e branco','Acarajé, vatapá, azeite','Vinho tinto, cerveja','Machado duplo (oxé)','Kaô Kabiesilê!','Senhor da justiça'],
    ['Oxum','Ketu','08/12','Sábado','Amarelo e dourado','Mel, acarajé, milho amarelo','Mel, champanhe, laranja','Abebê, espelho, leque','Ora Iê Iê Ô!','Orixá do amor e das águas doces'],
    ['Iemanjá','Ketu','02/02','Sábado','Azul e branco','Melão, uva branca, arroz','Champanhe, água, leite','Abebê de prata','Odoyá!','Rainha do mar'],
    ['Oxalá','Ketu','','Sexta','Branco','Inhame, arroz, canjica branca','Água, leite','Opaxorô, pano da costa branco','Êpa Babá!','Pai criador, orixá da paz'],
    ['Iansã','Ketu','04/12','Terça','Vermelho e marrom','Acarajé, abará','Vinho tinto, cerveja','Espada, eruexim','Eparrêi Iansã!','Orixá dos ventos e dos raios'],
    ['Nanã','Ketu','26/07','Segunda','Roxo e branco','Inhame, milho, canjica','Água, leite','Ibiri (vassoura de palha)','Salúbà Nanã!','Mais antiga dos orixás, orixá da lama'],
    ['Obaluaê / Omolu','Ketu','','Segunda','Preto, branco e vermelho','Pipoca, milho, coco','Vinho tinto, dendê','Xaxará','Atotô Obaluaê!','Orixá da saúde e das doenças'],
    ['Preto-Velho','Umbanda','13/05','Segunda','Branco e preto','Fumo, cachaça, mel','Cachaça, mel','Cachimbo, bengala','Ave, meu filho!','Espírito de sabedoria e cura'],
    ['Caboclo','Umbanda','','Terça','Verde e amarelo','Mel, frutas, charuto','Cerveja, cachaça','Arco e flecha, cocar','Okê Caboclo!','Espírito da natureza'],
    ['Exu / Pombagira','Umbanda','','Segunda e sexta','Vermelho e preto','Rosa vermelha, champanhe','Champanhe, vinho tinto','Rosas, tridentes','Laroyê!','Guardião(ã) das encruzilhadas']
  ];
  aba.getRange(2,1,dados.length,10).setValues(dados);
  _finalizarAba(aba, COLS.length, dados.length);
}

function _criarFinanceiro(ss){
  const aba=_novaAba(ss,ABA_FINANCEIRO);
  const COLS = ['ID','Data','Tipo','Categoria','Descrição','Valor (R$)','Responsável','Comprovante','Observações'];
  _cabecalho(aba, COLS, '#1a2a0a','#d0f0b0');
  aba.getRange('B2:B200').setNumberFormat('dd/MM/yyyy');
  aba.getRange('F2:F200').setNumberFormat('R$ #,##0.00');
  _validacao(aba,'C2:C200',['Entrada','Saída']);
  _validacao(aba,'D2:D200', CAT_FINANCEIRO_LIST);
  _formatCond(aba,'C2:C200',[
    {v:'Entrada',bg:'#e6f4ea',f:'#1e6b3a'},
    {v:'Saída',  bg:'#fdecea',f:'#8b0000'}
  ]);
  _finalizarAba(aba, COLS.length, 200);
}

// ── MENU ─────────────────────────────────────────────────
function recalcularEstoque(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba){SpreadsheetApp.getUi().alert('Aba não encontrada.');return;}
  const rows=aba.getDataRange().getValues();let n=0;
  for(let i=1;i<rows.length;i++){
    const a=Number(rows[i][4])||0,m=Number(rows[i][5])||0;
    if(m>0){const{label,pct}=_nivelEstoque(a,m);aba.getRange(i+1,7).setValue(pct+'%');aba.getRange(i+1,8).setValue(label);n++;}
  }
  SpreadsheetApp.getUi().alert('✅ '+n+' itens recalculados.');
}

function gerarListaCompras(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONSUMIVEIS);
  if(!aba){SpreadsheetApp.getUi().alert('Aba não encontrada.');return;}
  const nok=aba.getDataRange().getValues().slice(1).filter(r=>r[0]!==''&&!String(r[7]).includes('OK'));
  if(!nok.length){SpreadsheetApp.getUi().alert('✅ Nenhum item precisa ser reposto!');return;}
  let msg='🛒 LISTA DE COMPRAS\n'+'─'.repeat(28)+'\n';
  ['Urgente','Alerta','Repor'].forEach(t=>{
    const g=nok.filter(r=>String(r[7]).includes(t));
    if(g.length){msg+='\n'+(t==='Urgente'?'🚨':t==='Alerta'?'🔴':'⚠️')+' '+t+':\n';g.forEach(r=>msg+='  • '+r[2]+' — faltam '+Math.max(0,r[5]-r[4])+' '+r[3]+'\n');}
  });
  SpreadsheetApp.getUi().alert(msg);
}

function verDatasMes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),hoje=new Date(),mes=hoje.getMonth();
  let msg='📅 '+Utilities.formatDate(hoje,Session.getScriptTimeZone(),'MMMM/yyyy').toUpperCase()+'\n'+'─'.repeat(28)+'\n';
  const aF=ss.getSheetByName(ABA_FILHOS);
  if(aF&&aF.getLastRow()>1){
    const aniv=aF.getDataRange().getValues().slice(1).filter(r=>{try{return r[4]&&new Date(r[4]).getMonth()===mes;}catch(e){return false;}});
    if(aniv.length){msg+='\n🎂 Aniversariantes:\n';aniv.forEach(r=>msg+='  • '+_dataFormatada_(r[4])+' — '+(r[1]||r[2])+'\n');}
  }
  const aE=ss.getSheetByName(ABA_ENTIDADES);
  if(aE&&aE.getLastRow()>1){
    const f=aE.getDataRange().getValues().slice(1).filter(r=>{if(!r[2]||r[2]==='')return false;const p=String(r[2]).split('/');return p.length>=2&&(parseInt(p[1])-1)===mes;});
    if(f.length){msg+='\n🥁 Festas:\n';f.forEach(r=>msg+='  • '+r[2]+' — '+r[0]+'\n');}
  }
  SpreadsheetApp.getUi().alert(msg);
}

function verObrigacoes(){
  const aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FILHOS);
  if(!aba){SpreadsheetApp.getUi().alert('Aba não encontrada.');return;}
  const hoje=new Date(),lim=new Date(hoje);lim.setDate(lim.getDate()+180);
  const prox=aba.getDataRange().getValues().slice(1).filter(r=>{
    if(!r[9])return false;
    try{const d=new Date(r[9]);return d>=hoje&&d<=lim;}catch(e){return false;}
  });
  if(!prox.length){SpreadsheetApp.getUi().alert('✅ Nenhuma obrigação nos próximos 180 dias.');return;}
  let msg='⚠️ OBRIGAÇÕES — 180 DIAS\n'+'─'.repeat(28)+'\n';
  prox.sort((a,b)=>new Date(a[9])-new Date(b[9]));
  prox.forEach(r=>{const dias=Math.round((new Date(r[9])-hoje)/(864e5));msg+='  • '+(r[1]||r[2])+' — '+_dataFormatada_(r[9])+' (em '+dias+' dias)\n';});
  SpreadsheetApp.getUi().alert(msg);
}

// ── HELPERS ──────────────────────────────────────────────
function _nivelEstoque(a,m){
  if(!m)return{nivel:'sem-minimo',label:'— Sem mínimo',pct:0};
  const pct=Math.round((a/m)*100);
  if(pct>=75)return{nivel:'ok',label:'✅ OK',pct};
  if(pct>=25)return{nivel:'repor',label:'⚠️ Repor',pct};
  if(pct>=5)return{nivel:'alerta',label:'🔴 Alerta',pct};
  return{nivel:'urgente',label:'🚨 Urgente / Zero',pct};
}
function _dataFormatada(){return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy');}
function _dataFormatada_(d){if(!d)return'—';try{return Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd/MM/yyyy');}catch(e){return'—';}}
function saida(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
