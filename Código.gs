/** ============== CONFIG & SETUP ============== **/
const SS = SpreadsheetApp.getActive();
const SHEET_SETTINGS   = SS.getSheetByName('Settings');
const SHEET_CUSTOMERS  = SS.getSheetByName('Customers');
const SHEET_TX         = SS.getSheetByName('Transactions');
const SHEET_USERS      = SS.getSheetByName('Users');

// Sessão: 6h (21600s)
const SESSION_TTL_SEC = 21600;

/** ============== WEB APP ============== **/
function doGet(e){
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cashback da Casa')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** ============== UTIL: HASH / SESSÕES / CPF ============== **/
function hashSenha(plain){
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain);
  return raw.map(b => ('0' + (b + 256).toString(16)).slice(-2)).join('');
}

function gerarToken(){ return Utilities.getUuid(); }

function saveSession(token, username){
  CacheService.getScriptCache().put(token, username, SESSION_TTL_SEC);
}

function getSessionUser(token){
  return CacheService.getScriptCache().get(token); // username ou null
}

function invalidateSession(token){
  CacheService.getScriptCache().remove(token);
}

// >>> NORMALIZADOR DE CPF (resolve zeros à esquerda)
function _normCPF(v){
  return String(v || '').replace(/\D/g,'').padStart(11,'0');
}

/** ============== USERS ============== **/
function _findUserRow(username){
  const vals = SHEET_USERS.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++){
    if (String(vals[i][0]) === String(username)) return i + 1;
  }
  return -1;
}

function createUser(username, plainPassword, role){
  const hash = hashSenha(plainPassword);
  SHEET_USERS.appendRow([username, hash, role || 'admin', true, true, new Date(), '', '']);
  return { ok: true };
}

function verifyCredentials(username, plainPassword){
  const row = _findUserRow(username);
  if (row === -1) return { ok:false, msg:'Usuário não encontrado' };
  const r = SHEET_USERS.getRange(row, 1, 1, 8).getValues()[0];
  const ativo = !!r[3];
  if (!ativo) return { ok:false, msg:'Usuário inativo' };
  const storedHash = r[1];
  const attemptHash = hashSenha(plainPassword);
  if (storedHash === attemptHash) return { ok:true, must_change: !!r[4] };
  return { ok:false, msg:'Senha inválida' };
}

function setNewPassword(username, newPlain){
  const row = _findUserRow(username);
  if (row === -1) return { ok:false, msg:'Usuário não encontrado' };
  SHEET_USERS.getRange(row, 2).setValue(hashSenha(newPlain)); // password_hash
  SHEET_USERS.getRange(row, 5).setValue(false);               // must_change = false
  SHEET_USERS.getRange(row, 7).setValue(new Date());          // last_login
  return { ok:true };
}

/** ============== SETTINGS & CLIENTES ============== **/
function _getSettings(){
  const o = {};
  const rows = SHEET_SETTINGS.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++){
    o[String(rows[i][0]).trim()] = rows[i][1];
  }
  o.cashback_percent   = Number(o.cashback_percent || 5);
  o.validade_dias      = Number(o.validade_dias || 90);
  o.ticket_min         = Number(o.ticket_min || 30);
  o.teto_por_transacao = Number(o.teto_por_transacao || 20);
  o.teto_por_cpf_mes   = Number(o.teto_por_cpf_mes || 50);
  // limite diário opcional; se não existir, muito alto
  o.teto_por_cpf_dia   = Number(o.teto_por_cpf_dia || 999999);
  return o;
}

function _findCustomerRowByCPF(cpf){
  const alvo = _normCPF(cpf);
  const vals = SHEET_CUSTOMERS.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++){
    const v = _normCPF(vals[i][0]);
    if (v === alvo) return i + 1;
  }
  return -1;
}

// >>> AJUSTADO: cria linha com setValues garantindo formato texto e sem pular linha
function _getCustomer(cpf){
  const norm = _normCPF(cpf);
  const row = _findCustomerRowByCPF(norm);
  if (row === -1){
    const newRow = SHEET_CUSTOMERS.getLastRow() + 1;
    SHEET_CUSTOMERS.getRange(newRow, 1).setNumberFormat('@');
    SHEET_CUSTOMERS.getRange(newRow, 1, 1, 6).setValues([[norm, '', '', 0, '', new Date()]]);
    return { cpf: norm, nome: '', telefone: '', saldo_centavos: 0, ultimo_uso: '' };
  }
  const r = SHEET_CUSTOMERS.getRange(row, 1, 1, 6).getValues()[0];
  return {
    cpf: _normCPF(r[0]),
    nome: r[1] || '',
    telefone: r[2] || '',
    saldo_centavos: Number(r[3]) || 0,
    ultimo_uso: r[4] || ''
  };
}

// >>> AJUSTADO: sincroniza espelho e garante formato texto
function _setCustomerBalance(cpf, saldo_centavos){
  const norm = _normCPF(cpf);
  const row = _findCustomerRowByCPF(norm);
  const agora = new Date();
  if (row === -1){
    const newRow = SHEET_CUSTOMERS.getLastRow() + 1;
    SHEET_CUSTOMERS.getRange(newRow, 1).setNumberFormat('@');
    SHEET_CUSTOMERS.getRange(newRow, 1, 1, 6).setValues([[norm, '', '', saldo_centavos, agora, agora]]);
  } else {
    SHEET_CUSTOMERS.getRange(row, 1).setNumberFormat('@');
    SHEET_CUSTOMERS.getRange(row, 1).setValue(norm);
    SHEET_CUSTOMERS.getRange(row, 4).setValue(saldo_centavos);
    SHEET_CUSTOMERS.getRange(row, 5).setValue(agora);
  }
}

/** ============== AUTH MIDDLEWARE ============== **/
function requireAuth(token){
  const user = getSessionUser(token);
  if (!user) throw new Error('Sessão inválida ou expirada');
  return user; // username
}

/** ============== HELPERS DE CÁLCULO (dinâmico) ============== **/
// Saldo válido agora (considera validade para CREDITO; RESGATE/AJUSTE sempre contam)
function _saldoAtualElegivel_(cpf){
  const alvo = _normCPF(cpf);
  const st = _getSettings();
  const hoje = new Date();
  const rows = SHEET_TX.getDataRange().getValues().slice(1);
  let saldo = 0;

  rows.forEach(r=>{
    if(!r[0]) return;
    const ts   = new Date(r[0]);
    const tipo = String(r[1]);
    const c    = _normCPF(r[2]);
    const v    = Number(r[3])||0; // CREDITO (+), RESGATE (-), AJUSTE (+/-)
    if (c !== alvo) return;

    if (tipo === 'CREDITO'){
      const expira = new Date(ts);
      expira.setDate(expira.getDate() + st.validade_dias);
      if (hoje <= expira) saldo += v; // só créditos não vencidos
    } else if (tipo === 'RESGATE' || tipo === 'AJUSTE'){
      saldo += v;
    }
  });

  return Math.max(0, saldo);
}

// Somatório de CREDITO já emitido no mês e no dia (para apoio)
function _totaisCreditoMesDia_(cpf){
  const alvo = _normCPF(cpf);
  const tz = Session.getScriptTimeZone();
  const rows = SHEET_TX.getDataRange().getValues().slice(1);
  const mesRef = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const diaRef = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  let mesCent = 0, diaCent = 0;

  rows.forEach(r=>{
    if(!r[0]) return;
    const ts   = new Date(r[0]);
    const tipo = String(r[1]);
    const c    = _normCPF(r[2]);
    const v    = Number(r[3])||0;
    if (tipo !== 'CREDITO' || c !== alvo) return;

    const ym  = Utilities.formatDate(ts, tz, 'yyyy-MM');
    const ymd = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    if (ym === mesRef)  mesCent += v;
    if (ymd === diaRef) diaCent += v;
  });

  return { mesCent, diaCent };
}

// Folgas que governam o crédito permitido agora.
// Se "saldoElegOverride" for informado, evita recalcular o saldo elegível.
function _limitesRestantesMesDia_(cpf, saldoElegOverride){
  const st = _getSettings();
  const saldoEleg = (typeof saldoElegOverride === 'number') ? saldoElegOverride : _saldoAtualElegivel_(cpf);

  const { mesCent, diaCent } = _totaisCreditoMesDia_(cpf);

  const tetoMesCent = Math.round(st.teto_por_cpf_mes * 100);
  const tetoDiaCent = Math.round(st.teto_por_cpf_dia * 100);

  const restantePorSaldo = Math.max(0, tetoMesCent - saldoEleg);
  const restantePorHistorico = Math.max(0, tetoMesCent - mesCent);
  const restCreditoMes = Math.min(restantePorSaldo, restantePorHistorico);
  const restDia = Math.max(0, tetoDiaCent - diaCent);

  return { restCreditoMes, restDia };
}

/** ============== AUTH API ============== **/
function apiLogin(payload){
  const u = String(payload.username || '').trim();
  const p = String(payload.password || '');
  const res = verifyCredentials(u, p);
  if (!res.ok) return { ok:false, msg: res.msg };
  const token = gerarToken();
  saveSession(token, u);
  return { ok:true, token, must_change: !!res.must_change, username: u };
}

function apiChangePassword(payload){
  const token = String(payload.token || '');
  const sessUser = getSessionUser(token);
  const username = String(payload.username || '');
  const newPass  = String(payload.newPassword || '');
  if (!sessUser || sessUser !== username) return { ok:false, msg:'Sessão inválida' };
  if (newPass.length < 8) return { ok:false, msg:'Senha deve ter pelo menos 8 caracteres.' };
  return setNewPassword(username, newPass);
}

function apiLogout(payload){
  const token = String(payload.token || '');
  invalidateSession(token);
  return { ok:true };
}

/** ============== BUSINESS API (dinâmicas) ============== **/
// Consulta saldo (recalcula validade) e devolve folgas de teto
function apiGetSaldo(payload){
  const token = String(payload.token || '');
  requireAuth(token);
  const cpf = _normCPF(payload.cpf);

  _getCustomer(cpf); // garante cadastro
  const saldoEleg = _saldoAtualElegivel_(cpf);

  const { restCreditoMes, restDia } = _limitesRestantesMesDia_(cpf, saldoEleg);
  _setCustomerBalance(cpf, saldoEleg); // espelho

  return {
    ok:true,
    cpf,
    saldo: saldoEleg / 100,
    limiteRestanteMes: restCreditoMes / 100,
    limiteRestanteDia: restDia / 100
  };
}

// Lançar compra (respeita limites mensal/diário)
function apiLancarCompra(payload){
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(20000);
  if (!locked) {
    return { ok:false, msg:'Sistema temporariamente ocupado. Tente novamente.' };
  }
  try {
    const token = String(payload.token || '');
    const operador = requireAuth(token);
    const cpf = _normCPF(payload.cpf);
    const valorCompra = Number(payload.valorCompra || 0);
    const notaRef = String(payload.notaRef || '');

    const st = _getSettings();
    if (valorCompra < st.ticket_min) {
      return { ok:false, msg: `Ticket mínimo R$ ${st.ticket_min}` };
    }

    let creditoCent = Math.round(valorCompra * (st.cashback_percent / 100) * 100);
    const tetoTransCent = Math.round(st.teto_por_transacao * 100);
    if (creditoCent > tetoTransCent) creditoCent = tetoTransCent;

    const { restCreditoMes, restDia } = _limitesRestantesMesDia_(cpf);
    let permitido = Math.min(restCreditoMes, restDia);

    if (permitido <= 0){
      const msg = restCreditoMes <= 0
        ? `Limite mensal de R$ ${st.teto_por_cpf_mes.toFixed(2)} já atingido para este CPF.`
        : `Limite diário de R$ ${st.teto_por_cpf_dia.toFixed(2)} já atingido para este CPF.`;
      return { ok:false, msg };
    }
    if (creditoCent > permitido) creditoCent = permitido;

    // grava linha com CPF normalizado e coluna como texto
    const newRow = [ new Date(), 'CREDITO', cpf, creditoCent,
                     Math.round(valorCompra * 100), operador, notaRef, '' ];
    const lr = SHEET_TX.getLastRow() + 1;
    SHEET_TX.getRange(lr,3).setNumberFormat('@'); // coluna CPF
    SHEET_TX.appendRow(newRow);

    const saldoEleg = _saldoAtualElegivel_(cpf);
    _setCustomerBalance(cpf, saldoEleg);

    const limited = (creditoCent < Math.round(valorCompra * (st.cashback_percent / 100) * 100));
    return { ok:true, cpf, credito: (creditoCent / 100), saldoAtual: (saldoEleg / 100), limited };
  } catch (e) {
    return { ok:false, msg: e.message };
  } finally {
    try { if (locked) lock.releaseLock(); } catch(_) {}
  }
}

// Resgatar usa saldo elegível (já considera vencimentos)
function apiResgatar(payload){
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(20000);
  if (!locked) {
    return { ok:false, msg:'Sistema temporariamente ocupado. Tente novamente.' };
  }
  try {
    const token = String(payload.token || '');
    const operador = requireAuth(token);
    const cpf = _normCPF(payload.cpf);
    const valorResgate = Number(payload.valorResgate || 0);

    const saldoEleg = _saldoAtualElegivel_(cpf);
    const resgateCent = Math.round(valorResgate * 100);
    if (resgateCent <= 0) return { ok:false, msg:'Valor inválido' };
    if (resgateCent > saldoEleg) return { ok:false, msg:'Saldo insuficiente' };

    const lr = SHEET_TX.getLastRow() + 1;
    SHEET_TX.getRange(lr,3).setNumberFormat('@');
    SHEET_TX.appendRow([ new Date(), 'RESGATE', cpf, -resgateCent, '', operador, '', '' ]);

    const novoSaldo = _saldoAtualElegivel_(cpf);
    _setCustomerBalance(cpf, novoSaldo);

    return { ok:true, cpf, resgatado: (resgateCent / 100), saldoAtual: (novoSaldo / 100) };
  } catch (e) {
    return { ok:false, msg: e.message };
  } finally {
    try { if (locked) lock.releaseLock(); } catch(_) {}
  }
}

/** ============== NOVAS APIs - MELHORIAS ============== **/

// Dashboard: métricas do dia e mês atual
function apiGetDashboard(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);
    
    Logger.log('=== Dashboard: Iniciando ===');
    
    const tz = Session.getScriptTimeZone();
    const hoje = new Date();
    const mesAtual = Utilities.formatDate(hoje, tz, 'yyyy-MM');
    const diaAtual = Utilities.formatDate(hoje, tz, 'yyyy-MM-dd');
    
    const txData = SHEET_TX.getDataRange().getValues();
    
    let creditoDia = 0, creditoMes = 0;
    let resgateDia = 0, resgateMes = 0;
    let txDia = 0, txMes = 0;
    const cpfsHoje = {};
    const cpfsMes = {};
    
    // Pula header (linha 0)
    for (let i = 1; i < txData.length; i++) {
      const row = txData[i];
      if (!row[0]) continue;
      
      const ts = new Date(row[0]);
      const tipo = String(row[1] || '');
      const cpf = _normCPF(row[2]);
      const valor = Number(row[3]) || 0; // JÁ está em centavos
      
      const ym = Utilities.formatDate(ts, tz, 'yyyy-MM');
      const ymd = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
      
      if (ymd === diaAtual) {
        txDia++;
        if (cpf) cpfsHoje[cpf] = true;
        if (tipo === 'CREDITO') creditoDia += valor;
        if (tipo === 'RESGATE') resgateDia += Math.abs(valor);
      }
      
      if (ym === mesAtual) {
        txMes++;
        if (cpf) cpfsMes[cpf] = true;
        if (tipo === 'CREDITO') creditoMes += valor;
        if (tipo === 'RESGATE') resgateMes += Math.abs(valor);
      }
    }
    
    const totalClientes = Math.max(0, SHEET_CUSTOMERS.getLastRow() - 1);
    
    Logger.log('Dashboard: Hoje=' + txDia + ' tx, Mês=' + txMes + ' tx');
    
    return {
      ok: true,
      hoje: {
        transacoes: txDia,
        clientesUnicos: Object.keys(cpfsHoje).length,
        creditoGerado: creditoDia / 100,
        resgateFeito: resgateDia / 100
      },
      mes: {
        transacoes: txMes,
        clientesUnicos: Object.keys(cpfsMes).length,
        creditoGerado: creditoMes / 100,
        resgateFeito: resgateMes / 100,
        saldoPendente: (creditoMes - resgateMes) / 100
      },
      totalClientes
    };
  } catch (e) {
    Logger.log('ERRO apiGetDashboard: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

// Histórico completo de um CPF
function apiGetHistorico(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);
    const cpfDigitado = String(payload.cpf || '');
    const cpf = _normCPF(cpfDigitado);
    
    Logger.log('Buscando histórico para CPF: ' + cpf);
    
    const txData = SHEET_TX.getDataRange().getValues();
    const historico = [];
    
    // Pula o header (linha 0)
    for (let i = 1; i < txData.length; i++){
      const row = txData[i];
      if (!row[0]) continue; // pula linha vazia
      
      const cpfLinha = _normCPF(row[2]);
      if (cpfLinha !== cpf) continue;
      
      historico.push({
        data: row[0], // já é Date do Sheets
        tipo: String(row[1] || ''),
        valor: (Number(row[3]) || 0) / 100,
        valorCompra: row[4] ? (Number(row[4]) / 100) : null,
        operador: String(row[5] || ''),
        nota: String(row[6] || '')
      });
    }
    
    Logger.log('Encontradas ' + historico.length + ' transações');
    
    // Ordena do mais recente para o mais antigo
    historico.sort((a, b) => {
      const dateA = new Date(a.data);
      const dateB = new Date(b.data);
      return dateB - dateA;
    });
    
    const cliente = _getCustomer(cpf);
    const saldo = _saldoAtualElegivel_(cpf);
    _setCustomerBalance(cpf, saldo); // mantém espelho coerente
    
    return {
      ok: true,
      cpf,
      nome: cliente.nome,
      saldoAtual: saldo / 100,
      historico
    };
  } catch (e) {
    Logger.log('ERRO apiGetHistorico: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

// >>> AJUSTADO: Busca de clientes com saldo DINÂMICO
function apiBuscarClientes(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);
    const termo = String(payload.termo || '').toLowerCase().trim();
    
    Logger.log('Buscando clientes com termo: ' + termo);
    
    if (termo.length < 3) {
      return { ok: true, clientes: [] };
    }
    
    const custData = SHEET_CUSTOMERS.getDataRange().getValues();
    const resultados = [];
    const termoSemPontos = termo.replace(/\D/g, '');
    
    // Pula o header (linha 0)
    for (let i = 1; i < custData.length; i++){
      const row = custData[i];
      const cpfRaw = row[0];
      if (!cpfRaw) continue;
      
      const cpf = _normCPF(cpfRaw);
      const nomeLower = String(row[1] || '').toLowerCase().trim();
      const telefone = String(row[2] || '').trim();

      const cpfSemPontos = cpf.replace(/\D/g, '');
      const matchCPF = cpfSemPontos.includes(termoSemPontos);
      const matchNome = nomeLower && nomeLower.includes(termo);
      
      if (matchCPF || matchNome) {
        const saldoEleg = _saldoAtualElegivel_(cpf); // dinâmico
        _setCustomerBalance(cpf, saldoEleg);         // sincroniza espelho

        if (saldoEleg > 0) {
          resultados.push({
            cpf,
            cpf_formatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4'),
            nome: row[1] || 'Cliente ' + cpf.substring(0, 3) + '...',
            telefone: telefone || '-',
            saldo: saldoEleg / 100
          });
        }
      }
      
      if (resultados.length >= 10) break;
    }
    
    Logger.log('Encontrados ' + resultados.length + ' clientes');
    
    return { ok: true, clientes: resultados };
  } catch (e) {
    Logger.log('ERRO apiBuscarClientes: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

function _coletarClientesComSaldo_(opcoes){
  const incluirZeros = !!(opcoes && opcoes.incluirZeros);

  const settings = _getSettings();
  const validadeDias = Number(settings.validade_dias || 0);
  const validadeSegura = Number.isFinite(validadeDias) ? validadeDias : 0;
  const hoje = new Date();

  const txRows = SHEET_TX.getDataRange().getValues();
  const saldoMap = {};
  const ultimoUsoMap = {};

  for (let i = 1; i < txRows.length; i++){
    const row = txRows[i];
    if (!row[0] || !row[2]) continue;

    const tipo = String(row[1] || '').trim();
    const cpf = _normCPF(row[2]);
    if (!cpf) continue;

    const valorCent = Number(row[3]);
    const tsRaw = row[0];
    const ts = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
    if (!Number.isFinite(valorCent) || !(ts instanceof Date) || isNaN(ts.getTime())) continue;

    if (!saldoMap[cpf]) saldoMap[cpf] = 0;

    if (!ultimoUsoMap[cpf] || ts > ultimoUsoMap[cpf]) {
      ultimoUsoMap[cpf] = ts;
    }

    if (tipo === 'CREDITO'){
      const expira = new Date(ts);
      expira.setDate(expira.getDate() + validadeSegura);
      if (!isNaN(expira.getTime()) && hoje <= expira) {
        saldoMap[cpf] += valorCent;
      }
    } else if (tipo === 'RESGATE' || tipo === 'AJUSTE'){
      saldoMap[cpf] += valorCent;
    }
  }

  const custData = SHEET_CUSTOMERS.getDataRange().getValues();
  const clientes = [];
  const saldoUpdates = [];
  const ultimoUsoUpdates = [];

  for (let i = 1; i < custData.length; i++){
    const row = custData[i];
    const cpfBruto = row[0];

    if (!cpfBruto){
      saldoUpdates.push([0]);
      ultimoUsoUpdates.push(['']);
      continue;
    }

    const cpf = _normCPF(cpfBruto);

    const nome = String(row[1] || '').trim();
    const telefone = String(row[2] || '').trim();

    const temSaldoCalculado = Object.prototype.hasOwnProperty.call(saldoMap, cpf);
    const saldoPlanilha = Number(row[3]) || 0;
    const saldoEleg = Math.max(0, temSaldoCalculado ? Number(saldoMap[cpf]) || 0 : saldoPlanilha);

    let ultimoUso = ultimoUsoMap[cpf];
    if (!ultimoUso){
      const rawUso = row[4];
      if (rawUso instanceof Date && !isNaN(rawUso.getTime())){
        ultimoUso = rawUso;
      }
    }

    saldoUpdates.push([saldoEleg]);
    ultimoUsoUpdates.push([ultimoUso || '']);

    if (incluirZeros || saldoEleg > 0) {
      clientes.push({
        cpf,
        cpf_formatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4'),
        nome: nome || 'Cliente ' + cpf.substring(0, 3) + '...',
        telefone: telefone || '-',
        saldo: saldoEleg / 100,
        ultimo_uso: ultimoUso || ''
      });
    }
  }

  if (saldoUpdates.length || ultimoUsoUpdates.length){
    try {
      if (saldoUpdates.length){
        SHEET_CUSTOMERS.getRange(2, 4, saldoUpdates.length, 1).setValues(saldoUpdates);
      }
      if (ultimoUsoUpdates.length){
        SHEET_CUSTOMERS.getRange(2, 5, ultimoUsoUpdates.length, 1).setValues(ultimoUsoUpdates);
      }
    } catch (syncErr) {
      Logger.log('Aviso: falha ao sincronizar espelho de clientes: ' + syncErr.message);
    }
  }

  clientes.sort((a, b) => b.saldo - a.saldo);
  return clientes;
}

// >>> AJUSTADO: TOP clientes calculando saldo DINÂMICO e atualizando espelho
function apiGetTopClientes(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);

    Logger.log('=== TOP Clientes: Iniciando ===');

    const clientes = _coletarClientesComSaldo_({ incluirZeros: false });
    Logger.log('Total clientes considerados: ' + clientes.length);

    const top20 = clientes.slice(0, 20);

    Logger.log('Retornando TOP ' + top20.length);

    return { ok: true, clientes: top20 };
  } catch (e) {
    Logger.log('ERRO apiGetTopClientes: ' + e.message);
    Logger.log('Stack: ' + e.stack);

    try {
      const fallback = _listarClientesDaSheet_();
      Logger.log('Retornando fallback com ' + fallback.length + ' clientes.');
      return { ok: true, clientes: fallback };
    } catch (fallbackErr) {
      Logger.log('ERRO fallback TOP clientes: ' + fallbackErr.message);
      return { ok: false, msg: e.message };
    }
  }
}

function apiGetClientesRelatorio(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);

    Logger.log('Gerando relatório completo de clientes');

    const clientes = _coletarClientesComSaldo_({ incluirZeros: true });
    Logger.log('Total clientes retornados: ' + clientes.length);

    return { ok: true, clientes };
  } catch (e) {
    Logger.log('ERRO apiGetClientesRelatorio: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

function _listarClientesDaSheet_(){
  const custData = SHEET_CUSTOMERS.getDataRange().getValues();
  const clientes = [];

  for (let i = 1; i < custData.length; i++){
    const row = custData[i];
    const cpfBruto = row[0];
    if (!cpfBruto) continue;

    const cpf = _normCPF(cpfBruto);
    const saldoCent = Number(row[3]) || 0;
    if (saldoCent <= 0) continue;

    clientes.push({
      cpf,
      cpf_formatado: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4'),
      nome: String(row[1] || '').trim() || 'Cliente ' + cpf.substring(0, 3) + '...',
      telefone: String(row[2] || '').trim() || '-',
      saldo: saldoCent / 100,
      ultimo_uso: (row[4] instanceof Date && !isNaN(row[4].getTime())) ? row[4] : ''
    });
  }

  clientes.sort((a, b) => b.saldo - a.saldo);
  return clientes.slice(0, 20);
}

/** ============== JOB: EXPIRAR CRÉDITOS (opcional/manutenção) ============== **/
// Mantém a aba Customers espelhada; não é obrigatório (saldo já é dinâmico).
function jobExpirarCreditos(){
  const st = _getSettings();
  const rows = SHEET_TX.getDataRange().getValues().slice(1);
  const hoje = new Date();
  const map = {}; // cpf -> saldo válido

  rows.forEach(r => {
    if (!r[0]) return;
    const ts = new Date(r[0]);
    const tipo = String(r[1]);
    const cpf  = _normCPF(r[2]);
    const v    = Number(r[3]) || 0;

    if (!map[cpf]) map[cpf] = 0;

    if (tipo === 'CREDITO'){
      const expira = new Date(ts);
      expira.setDate(expira.getDate() + st.validade_dias);
      if (hoje <= expira) map[cpf] += v;
    } else if (tipo === 'RESGATE' || tipo === 'AJUSTE'){
      map[cpf] += v;
    }
  });

  const vals = SHEET_CUSTOMERS.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++){
    const cpf = _normCPF(vals[i][0]);
    const saldoNovo = Math.max(0, map[cpf] || 0);
    SHEET_CUSTOMERS.getRange(i + 1, 1).setNumberFormat('@');
    SHEET_CUSTOMERS.getRange(i + 1, 1).setValue(cpf);
    SHEET_CUSTOMERS.getRange(i + 1, 4).setValue(saldoNovo);
  }
}

/** ============== UTIL DIVERSOS / EXTRAS ============== **/
function addLojista() {
  // troque aqui o usuário/senha inicial
  createUser("lojista1", "Senha123", "admin");
}

/** ============== FUNÇÃO DE DIAGNÓSTICO ============== **/
function testarLeituraDados(){
  Logger.log('=== TESTE DE LEITURA DE DADOS ===');
  
  try {
    // Testa conexão com planilhas
    Logger.log('1. Testando conexão com Customers...');
    const custRows = SHEET_CUSTOMERS.getLastRow();
    Logger.log('   - Linhas em Customers: ' + custRows);
    
    Logger.log('2. Testando conexão com Transactions...');
    const txRows = SHEET_TX.getLastRow();
    Logger.log('   - Linhas em Transactions: ' + txRows);
    
    // Testa leitura de 3 primeiras linhas de Customers
    Logger.log('3. Lendo primeiras 3 linhas de Customers...');
    if (custRows >= 2) {
      const custData = SHEET_CUSTOMERS.getRange(2, 1, Math.min(3, custRows-1), 6).getValues();
      custData.forEach((row, idx) => {
        Logger.log('   Linha ' + (idx+2) + ': CPF=' + row[0] + ', Nome=' + row[1] + ', Saldo=' + row[3]);
      });
    }
    
    // Testa leitura de 3 primeiras transações
    Logger.log('4. Lendo primeiras 3 linhas de Transactions...');
    if (txRows >= 2) {
      const txData = SHEET_TX.getRange(2, 1, Math.min(3, txRows-1), 8).getValues();
      txData.forEach((row, idx) => {
        Logger.log('   Linha ' + (idx+2) + ': Data=' + row[0] + ', Tipo=' + row[1] + ', CPF=' + row[2] + ', Valor=' + row[3]);
      });
    }
    
    // Testa função de normalização de CPF
    Logger.log('5. Testando normalização de CPF...');
    const teste1 = _normCPF('123.456.789-01');
    Logger.log('   "123.456.789-01" -> "' + teste1 + '"');
    const teste2 = _normCPF('12345678901');
    Logger.log('   "12345678901" -> "' + teste2 + '"');
    const teste3 = _normCPF('11111111111');
    Logger.log('   "11111111111" -> "' + teste3 + '"');
    
    Logger.log('=== TESTE CONCLUÍDO ===');
    return 'Verifique os logs acima';
    
  } catch (e) {
    Logger.log('ERRO NO TESTE: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return 'ERRO: ' + e.message;
  }
}

// Testa buscar histórico do CPF 11111111111
function testarHistoricoCPF(){
  Logger.log('=== TESTE HISTÓRICO CPF 11111111111 ===');
  
  try {
    const cpfTeste = '11111111111';
    Logger.log('1. CPF para buscar: ' + cpfTeste);
    Logger.log('2. CPF normalizado: ' + _normCPF(cpfTeste));
    
    const txData = SHEET_TX.getDataRange().getValues();
    Logger.log('3. Total de linhas em Transactions: ' + txData.length);
    
    let encontradas = 0;
    for (let i = 1; i < txData.length; i++){
      const row = txData[i];
      if (!row[0]) {
        Logger.log('   Linha ' + (i+1) + ': VAZIA - pulando');
        continue;
      }
      
      const cpfLinha = String(row[2] || '');
      const cpfNorm = _normCPF(cpfLinha);
      
      Logger.log('   Linha ' + (i+1) + ': CPF bruto="' + cpfLinha + '" -> normalizado="' + cpfNorm + '"');
      
      if (cpfNorm === _normCPF(cpfTeste)) {
        encontradas++;
        Logger.log('      ✓ MATCH! Tipo=' + row[1] + ', Valor=' + row[3]);
      }
    }
    
    Logger.log('4. Total de transações encontradas: ' + encontradas);
    Logger.log('=== FIM DO TESTE ===');
    return 'Encontradas ' + encontradas + ' transações';
    
  } catch (e) {
    Logger.log('ERRO: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return 'ERRO: ' + e.message;
  }
}

// Testa buscar TOP clientes
function testarTopClientes(){
  Logger.log('=== TESTE TOP CLIENTES ===');
  
  try {
    const custData = SHEET_CUSTOMERS.getDataRange().getValues();
    Logger.log('1. Total de linhas em Customers: ' + custData.length);
    
    const clientes = [];
    
    for (let i = 1; i < custData.length; i++){
      const row = custData[i];
      
      Logger.log('2. Linha ' + (i+1) + ':');
      Logger.log('   - CPF bruto: "' + row[0] + '"');
      Logger.log('   - Nome: "' + row[1] + '"');
      Logger.log('   - Saldo: ' + row[3]);
      
      if (!row[0]) {
        Logger.log('   - STATUS: CPF vazio, pulando');
        continue;
      }
      
      const cpf = _normCPF(row[0]);
      const saldoCentavos = Number(row[3]) || 0;
      
      Logger.log('   - CPF normalizado: "' + cpf + '"');
      Logger.log('   - Saldo em centavos: ' + saldoCentavos);
      
      if (saldoCentavos > 0) {
        clientes.push({
          cpf: cpf,
          saldo: saldoCentavos / 100
        });
        Logger.log('   - STATUS: ✓ ADICIONADO');
      } else {
        Logger.log('   - STATUS: Saldo zero, ignorando');
      }
    }
    
    Logger.log('3. Total de clientes com saldo: ' + clientes.length);
    clientes.forEach((c, idx) => {
      Logger.log('   Cliente ' + (idx+1) + ': CPF=' + c.cpf + ', Saldo=R$ ' + c.saldo.toFixed(2));
    });
    
    Logger.log('=== FIM DO TESTE ===');
    return 'Encontrados ' + clientes.length + ' clientes';
    
  } catch (e) {
    Logger.log('ERRO: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return 'ERRO: ' + e.message;
  }
}

// API para testar do front-end
function apiTestarDados(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);
    
    const resultado = testarLeituraDados();
    
    return {
      ok: true,
      msg: 'Teste executado. Verifique os logs em Apps Script > Execuções',
      detalhes: {
        customersRows: SHEET_CUSTOMERS.getLastRow(),
        transactionsRows: SHEET_TX.getLastRow()
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// API para testar histórico
function apiTestarHistorico(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);
    
    const resultado = testarHistoricoCPF();
    
    return {
      ok: true,
      msg: resultado + ' - Verifique os logs detalhados'
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// API para testar TOP clientes
function apiTestarTop(payload){
  try {
    const token = String(payload.token || '');
    requireAuth(token);
    
    const resultado = testarTopClientes();
    
    return {
      ok: true,
      msg: resultado + ' - Verifique os logs detalhados'
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// Resumo mensal (opcional)
function apiResumoMensal(payload){
  try{
    const token = String(payload.token || '');
    requireAuth(token);

    const tz = Session.getScriptTimeZone();
    let ym = String(payload.ym || '');
    if (!ym){ ym = Utilities.formatDate(new Date(), tz, 'yyyy-MM'); }

    const rows = SHEET_TX.getDataRange().getValues().slice(1);
    let creditCent = 0, resgateCent = 0;
    const cpfsSet = new Set();

    rows.forEach(r=>{
      const ts = r[0] ? new Date(r[0]) : null;
      if (!ts) return;
      const tsYM = Utilities.formatDate(ts, tz, 'yyyy-MM');
      if (tsYM !== ym) return;

      const tipo = r[1];
      const cpf  = _normCPF(r[2] || '');
      const v    = Number(r[3]) || 0;
      if (cpf) cpfsSet.add(cpf);
      if (tipo === 'CREDITO') creditCent += v;
      else if (tipo === 'RESGATE') resgateCent += (-v);
    });

    const clientesUnicos = Array.from(cpfsSet).filter(Boolean).length;
    return {
      ok: true,
      ym,
      clientesUnicos,
      credit: (creditCent/100),
      resgatado: (resgateCent/100),
      saldoPendente: ((creditCent - resgateCent)/100)
    };
  } catch(e){
    return { ok:false, msg:e.message };
  }
}

/** ============== MIGRAÇÃO (rode uma vez) ============== **/
// >>> AJUSTADO: garante coluna toda como texto antes de regravar
function migrarCpfParaTexto(){
  // Customers: Coluna A (CPF)
  const lc = SHEET_CUSTOMERS.getLastRow();
  if (lc >= 2){
    const colA = SHEET_CUSTOMERS.getRange(2,1,lc-1,1);
    colA.setNumberFormat('@'); // texto
    const valsC = colA.getValues().map(r => [_normCPF(r[0])]);
    colA.setValues(valsC);
  }

  // Transactions: Coluna C (CPF)
  const lt = SHEET_TX.getLastRow();
  if (lt >= 2){
    const colC = SHEET_TX.getRange(2,3,lt-1,1);
    colC.setNumberFormat('@'); // texto
    const valsT = colC.getValues().map(r => [_normCPF(r[0])]);
    colC.setValues(valsT);
  }
  
  SpreadsheetApp.flush();
  return { ok:true, msg:'CPF normalizado e colunas definidas como Texto.' };
}
