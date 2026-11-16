/**
 * API do Painel "Apontador de Produção" (TV) — planilha Marcador ULISSES / aba BD
 * -------------------------------------------------------------------------------
 * Faz os filtros e as contas no servidor e devolve um JSON pequeno para o painel.
 * Usa ContentService (não HtmlService), então NÃO aparece a faixa
 * "Este aplicativo foi criado por um usuário do Google Apps Script".
 *
 * Parâmetros (todos opcionais, menos o token):
 *   token=...                 senha (igual à do index.html)
 *   de=dd/MM/aaaa             data inicial  (padrão: hoje, ou o último dia com dados)
 *   ate=dd/MM/aaaa            data final    (padrão: igual a "de")
 *   operador=Nome             filtra um operador
 *   processo=Nome             filtra um processo
 *   limite=400                máximo de lançamentos devolvidos na tabela
 *   callback=fn               JSONP (o painel usa isso; funciona em TV antiga)
 *   nocache=1                 ignora o cache
 *
 * Regras (iguais ao Looker Studio):
 *   Meta / Realizado  = soma das colunas
 *   Saldo             = Realizado − Meta
 *   Eficiência        = média das eficiências dos lançamentos (coluna Eficiência)
 *   Prontas           = soma do Realizado onde "Processo Final?" = TRUE
 */

// ===================== CONFIGURAÇÃO =====================
var SPREADSHEET_ID = '1Ve9JRGD43cLo3HlWegFtoeN8FixPgXdwCaUtqQa3HOM';
var ABA = 'BD';
var TOKEN = 'ulisses-e6b4d1b0b15c';   // mesma senha do index.html ('' desativa)
var CACHE_SEGUNDOS = 45;
var LIMITE_PADRAO = 400;
var MAX_DIAS = 62;                 // trava para não pedirem um período gigante
// ========================================================

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (TOKEN && p.token !== TOKEN) {
      return responder_({ ok: false, erro: 'nao autorizado' }, p.callback);
    }
    var chave = 'p_' + [p.de, p.ate, p.operador, p.processo, p.limite].join('|');
    chave = chave.length > 200 ? chave.substr(0, 200) : chave;
    var cache = CacheService.getScriptCache();

    if (CACHE_SEGUNDOS > 0 && !p.nocache) {
      var emCache = cache.get(chave);
      if (emCache) return responderTexto_(emCache, p.callback);
    }

    var corpo = JSON.stringify(montarPainel_(p));

    if (CACHE_SEGUNDOS > 0 && corpo.length < 95000) {
      try { cache.put(chave, corpo, CACHE_SEGUNDOS); } catch (err) {}
    }
    return responderTexto_(corpo, p.callback);
  } catch (err) {
    return responder_({ ok: false, erro: String(err) }, p.callback);
  }
}

function montarPainel_(p) {
  var tz = Session.getScriptTimeZone();
  var agora = new Date();
  var hoje = Utilities.formatDate(agora, tz, 'dd/MM/yyyy');

  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABA);
  if (!sh) throw new Error('Aba "' + ABA + '" não encontrada');

  var ultimaLinha = sh.getLastRow();
  var ultimaCol = sh.getLastColumn();
  var cab = sh.getRange(1, 1, 1, ultimaCol).getDisplayValues()[0];

  var C = {
    data: col_(cab, ['Data']),
    hora: col_(cab, ['Hora']),
    produto: col_(cab, ['PRODUTO', 'Produto']),
    funcao: col_(cab, ['Função', 'Funcao']),
    meta: col_(cab, ['Meta']),
    realizado: col_(cab, ['Realizado']),
    eficiencia: col_(cab, ['Eficiência', 'Eficiencia']),
    final: col_(cab, ['Processo Final?', 'Processo Final']),
    inicio: col_(cab, ['HORA INÍCIO', 'HORA INICIO', 'Hora Inicio']),
    fim: col_(cab, ['HORA FIM', 'Hora Fim']),
    operador: col_(cab, ['Operador.', 'Operador']),
    processo: col_(cab, ['Processo.', 'Processo'])
  };
  if (C.data < 0 || C.meta < 0 || C.realizado < 0) {
    throw new Error('Cabeçalho da aba BD mudou: preciso das colunas Data, Meta e Realizado');
  }

  var txt = [], val = [];
  if (ultimaLinha >= 2) {
    var rng = sh.getRange(2, 1, ultimaLinha - 1, ultimaCol);
    txt = rng.getDisplayValues();
    val = rng.getValues();
  }

  // Datas existentes
  var datasMap = {};
  for (var i = 0; i < txt.length; i++) {
    var d0 = String(txt[i][C.data]).trim();
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(d0)) datasMap[d0] = true;
  }
  var datas = Object.keys(datasMap).sort(function (a, b) { return kd_(b) - kd_(a); });

  // Período
  var de = validaData_(p.de) || (datasMap[hoje] ? hoje : (datas[0] || hoje));
  var ate = validaData_(p.ate) || de;
  if (kd_(ate) < kd_(de)) { var tmp = de; de = ate; ate = tmp; }
  var kDe = kd_(de), kAte = kd_(ate);
  if (diasEntre_(de, ate) > MAX_DIAS) throw new Error('Período maior que ' + MAX_DIAS + ' dias');

  var fOp = (p.operador || '').trim();
  var fProc = (p.processo || '').trim().toLowerCase();
  var limite = Math.min(Number(p.limite) || LIMITE_PADRAO, 2000);

  var opsMap = {}, procsMap = {};
  var T = { meta: 0, realizado: 0, prontas: 0, somaEf: 0, nEf: 0, n: 0 };
  var O = {};
  var linhas = [];

  for (var r = 0; r < txt.length; r++) {
    var data = String(txt[r][C.data]).trim();
    if (!datasMap[data]) continue;
    var k = kd_(data);
    if (k < kDe || k > kAte) continue;

    var funcao = C.funcao >= 0 ? String(txt[r][C.funcao]).trim() : '';
    var operador = C.operador >= 0 ? String(txt[r][C.operador]).trim() : '';
    var processo = C.processo >= 0 ? String(txt[r][C.processo]).trim() : '';
    if ((!operador || operador === '#N/A') && funcao) {
      var partes = funcao.split('-');
      operador = partes.shift().trim();
      if (!processo || processo === '#N/A') processo = partes.join('-').trim();
    }
    if (!operador) operador = '(sem nome)';

    // listas para os filtros (dentro do período, antes de filtrar operador/processo)
    opsMap[operador] = true;
    if (processo) procsMap[processo] = true;

    if (fOp && operador !== fOp) continue;
    if (fProc && processo.toLowerCase() !== fProc) continue;

    var meta = num_(val[r][C.meta]);
    var real = num_(val[r][C.realizado]);
    var ef = C.eficiencia >= 0 ? efic_(val[r][C.eficiencia]) : null;
    if (ef === null && meta > 0) ef = real / meta;
    var fin = C.final >= 0 ? verdadeiro_(val[r][C.final]) : false;

    T.n++;
    T.meta += meta;
    T.realizado += real;
    if (fin) T.prontas += real;
    if (ef !== null) { T.somaEf += ef; T.nEf++; }

    var o = O[operador] || (O[operador] = { meta: 0, realizado: 0, somaEf: 0, nEf: 0 });
    o.meta += meta; o.realizado += real;
    if (ef !== null) { o.somaEf += ef; o.nEf++; }

    linhas.push({
      data: data,
      hora: C.hora >= 0 ? String(txt[r][C.hora]).trim() : '',
      produto: C.produto >= 0 ? String(txt[r][C.produto]).trim() : '',
      operador: operador,
      processo: processo,
      inicio: limpa_(C.inicio >= 0 ? txt[r][C.inicio] : ''),
      fim: limpa_(C.fim >= 0 ? txt[r][C.fim] : ''),
      meta: meta > 0 ? meta : null,
      realizado: real,
      eficiencia: ef === null ? null : arred_(ef),
      _k: k
    });
  }

  // Mais recentes primeiro (data desc, hora desc)
  linhas.sort(function (a, b) {
    if (a._k !== b._k) return b._k - a._k;
    return a.hora < b.hora ? 1 : (a.hora > b.hora ? -1 : 0);
  });
  var total = linhas.length;
  linhas = linhas.slice(0, limite);
  for (var x = 0; x < linhas.length; x++) delete linhas[x]._k;

  var porOperador = Object.keys(O).map(function (nome) {
    var g = O[nome];
    return {
      operador: nome,
      meta: g.meta,
      realizado: g.realizado,
      eficiencia: g.nEf ? arred_(g.somaEf / g.nEf) : null
    };
  }).sort(function (a, b) {
    var ea = a.eficiencia === null ? -1 : a.eficiencia;
    var eb = b.eficiencia === null ? -1 : b.eficiencia;
    return eb - ea;
  });

  return {
    ok: true,
    de: de,
    ate: ate,
    hoje: hoje,
    operador: fOp,
    processo: p.processo || '',
    atualizadoEm: Utilities.formatDate(agora, tz, 'HH:mm:ss'),
    datas: datas.slice(0, 90),
    operadores: Object.keys(opsMap).sort(ordPt_),
    processos: Object.keys(procsMap).sort(ordPt_),
    totais: {
      meta: T.meta,
      realizado: T.realizado,
      prontas: T.prontas,
      saldo: T.realizado - T.meta,
      eficiencia: T.nEf ? arred_(T.somaEf / T.nEf) : null,
      lancamentos: T.n
    },
    porOperador: porOperador,
    linhas: linhas,
    totalLinhas: total
  };
}

// ---------- auxiliares ----------
function col_(cab, nomes) {
  for (var n = 0; n < nomes.length; n++)
    for (var c = 0; c < cab.length; c++)
      if (String(cab[c]).trim() === nomes[n]) return c;
  return -1;
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === '' || typeof v === 'undefined') return 0;
  var n = parseFloat(String(v).replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

/** Eficiência vem como 0,4569 (número) ou "45,69%" (texto); erro (#DIV/0!) vira null. */
function efic_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v || '').trim();
  if (!s || s.charAt(0) === '#') return null;
  var n = parseFloat(s.replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, ''));
  if (!isFinite(n)) return null;
  return s.indexOf('%') >= 0 ? n / 100 : n;
}

function verdadeiro_(v) {
  if (v === true) return true;
  var s = String(v || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'VERDADEIRO' || s === 'SIM' || s === 'Y' || s === '1';
}

function limpa_(v) {
  var s = String(v || '').trim();
  return (!s || s.charAt(0) === '#') ? '' : s;
}

function validaData_(s) {
  s = String(s || '').trim();
  return /^\d{2}\/\d{2}\/\d{4}$/.test(s) ? s : '';
}

function kd_(d) { var p = d.split('/'); return Number(p[2] + p[1] + p[0]); }

function diasEntre_(a, b) {
  var pa = a.split('/'), pb = b.split('/');
  var da = new Date(pa[2], pa[1] - 1, pa[0]), db = new Date(pb[2], pb[1] - 1, pb[0]);
  return Math.round((db - da) / 86400000);
}

function ordPt_(a, b) { return a.localeCompare(b, 'pt-BR'); }

function arred_(x) { return Math.round(x * 10000) / 10000; }

function responder_(obj, callback) { return responderTexto_(JSON.stringify(obj), callback); }

function responderTexto_(texto, callback) {
  if (callback && /^[A-Za-z_$][\w$.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + texto + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

/** Rode no editor para testar sem implantar (veja o resultado em "Registro de execução"). */
function testarPainel() {
  var r = montarPainel_({});
  r.linhas = r.linhas.slice(0, 5);
  Logger.log(JSON.stringify(r, null, 2));
}
