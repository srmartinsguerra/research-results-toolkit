/**
 * RESEARCH RESULTS TOOLKIT — RESULTADOS E MASCARAMENTO
 * Arquivo único para Google Apps Script (runtime V8).
 * Contém os cinco módulos completos, incluindo o embaralhamento da TABELA 5.
 *
 * Instalar no projeto vinculado à planilha MESTRA.
 * Substitui os cinco arquivos anteriores; não manter cópias das mesmas funções.
 * Preservar os demais arquivos da MESTRA e seu appsscript.json.
 * Configure RRT_AUTHORIZED_EMAIL e RRT_SPREADSHEET_ID nas propriedades privadas.
 * Primeira execução após configurar: RRT_INSTALAR_E_LIGAR.
 *
 * A geração principal da MESTRA encontra tabelasAutomatico() neste arquivo.
 * LIGAR mascara antes da escrita; DESLIGAR regenera os originais da base atual.
 * Validação local com serviços simulados não equivale a execução no Google.
 */


/* ======================================================================
 * MÓDULO 1 DE 5 — RESULT_MASKING.gs
 * ====================================================================== */

/**
 * Research Results Toolkit - integração de geração e mascaramento, v4 com TABELA 5.
 * Instalar NO PROJETO DA MESTRA, junto dos quatro geradores corrigidos.
 * Não executar a versão antiga depois da migração.
 * DESLIGAR regenera os resultados originais a partir da base atual.
 * Backups antigos não são aplicados sobre linhas novas.
 */
var CM_CFG = Object.freeze({
  // Configure these privately in Apps Script > Project Settings > Script Properties.
  email: (PropertiesService.getScriptProperties().getProperty('RRT_AUTHORIZED_EMAIL') || '').trim().toLowerCase(),
  planilha: (PropertiesService.getScriptProperties().getProperty('RRT_SPREADSHEET_ID') || '').trim(),
  chave: 'RRT_MASCARAMENTO_V4',
  revisao: 'v4-generic-labels-1',
  abas: ['CONSORT', 'TABELA 3', 'TABELA 4', 'TABELA 5'],
  limite: 1500
});
var CM_LOCK = { n: 0, lock: null };
var CM_FORCE = false;

// Trava reentrante na MESMA execução; execuções distintas disputam a trava real.
function cmTrava_() {
  var adquiriu = false;
  return {
    tryLock: function(ms) {
      if (adquiriu) return true;
      if (!CM_LOCK.n) {
        CM_LOCK.lock = LockService.getDocumentLock();
        if (!CM_LOCK.lock) throw new Error('Instale no projeto vinculado à planilha mestra.');
        if (!CM_LOCK.lock.tryLock(ms)) return false;
      }
      CM_LOCK.n++;
      adquiriu = true;
      return true;
    },
    releaseLock: function() {
      if (!adquiriu) return;
      adquiriu = false;
      if (--CM_LOCK.n === 0) CM_LOCK.lock.releaseLock();
    }
  };
}
function cmProps_() { return PropertiesService.getScriptProperties(); }
function cmEstado_() {
  var json = cmProps_().getProperty(CM_CFG.chave);
  if (!json) throw new Error('Integração ainda não instalada. Execute RRT_INSTALAR_E_LIGAR.');
  var e = JSON.parse(json);
  if (e.versao !== 4 || e.planilha !== CM_CFG.planilha || typeof e.ligado !== 'boolean') {
    throw new Error('Estado de mascaramento incompatível. Nenhum resultado foi publicado.');
  }
  return e;
}
function cmVerificarConfiguracao_() {
  if (!CM_CFG.email || !CM_CFG.planilha) {
    throw new Error('Configure RRT_AUTHORIZED_EMAIL e RRT_SPREADSHEET_ID nas propriedades do script antes de executar.');
  }
}
function cmAutorizar_(manual) {
  cmVerificarConfiguracao_();
  var efetivo = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  var ativo = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (efetivo !== CM_CFG.email || (manual && ativo !== CM_CFG.email)) {
    throw new Error('Esta operação deve ser executada pela conta ' + CM_CFG.email + '.');
  }
}
function cmSS_() {
  cmVerificarConfiguracao_();
  var ss = SpreadsheetApp.getActive();
  if (!ss || ss.getId() !== CM_CFG.planilha) throw new Error('Projeto fora da planilha mestra configurada.');
  return ss;
}
function cmChaveAba_(nome) { return CM_CFG.chave + ':' + nome; }
function cmDiario_(nome) {
  return JSON.parse(cmProps_().getProperty(cmChaveAba_(nome)) || 'null');
}
function cmPrecisaEscrever_(nome) {
  var e = cmEstado_(), d = cmDiario_(nome);
  return CM_FORCE || !d || d.pendente || d.ciclo !== e.ciclo || d.revisao !== CM_CFG.revisao;
}
function cmConcluir_(nome) {
  if (CM_CFG.abas.indexOf(nome) < 0) return;
  var d = cmDiario_(nome);
  if (!d) throw new Error('Escrita sem registro de preparação: ' + nome);
  d.pendente = false;
  d.concluido = new Date().toISOString();
  cmProps_().setProperty(cmChaveAba_(nome), JSON.stringify(d));
}
function cmNormal_(v) {
  return String(v == null ? '' : v).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[‐‑‒–—-]/g, ' ').replace(/\s+/g, ' ').trim();
}
function cmInteiro_(v) { return typeof v === 'number' && Number.isSafeInteger(v) && v >= 0; }
function cmP_(v) {
  if (typeof v === 'number') return Number.isFinite(v) && v >= 0 && v <= 1;
  if (typeof v !== 'string' || !/^[<≤>≥=]?\s*(?:\d+(?:[.,]\d+)?|[.,]\d+)(?:e[-+]?\d+)?$/i.test(v.trim())) return false;
  var n = Number(v.trim().replace(/^[<≤>≥=]\s*/, '').replace(',', '.'));
  return Number.isFinite(n) && n >= 0 && n <= 1;
}
function cmSelecionar_(nome, r) {
  var a = [], vistos = {}, m = r.matriz;
  var preservadas = {};
  (r.preservar || []).forEach(function(p) { preservadas[p.l + ':' + p.c] = true; });
  function add(l, c, tipo, trechos) {
    var key = (l + 1) + ':' + (c + 1);
    if (preservadas[key]) throw new Error('Tentativa de mascarar célula do painel: ' + nome + ' ' + key);
    if (vistos[key]) throw new Error('Seleção duplicada: ' + nome + ' ' + key);
    vistos[key] = true;
    a.push({ l: l + 1, c: c + 1, tipo: tipo, valor: m[l][c], trechos: trechos || null });
    if (a.length > CM_CFG.limite) throw new Error('Limite de células excedido.');
  }
  if (nome === 'TABELA 3') {
    var headers = [];
    m.forEach(function(row, i) {
      var x = row.map(cmNormal_);
      if (x.indexOf('n') >= 0 && x.indexOf('%') >= 0 && x.indexOf('qui quadrado') >= 0 && x.indexOf('exato de fisher') >= 0) headers.push(i);
    });
    if (!headers.length) throw new Error('Cabeçalhos da TABELA 3 não reconhecidos; publicação bloqueada.');
    headers.forEach(function(ini, ix) {
      var fim = ix + 1 < headers.length ? headers[ix + 1] : m.length;
      m[ini].map(cmNormal_).forEach(function(h, c) {
        for (var l = ini + 1; l < fim; l++) {
          var v = m[l][c];
          if (h === 'n' && cmInteiro_(v)) add(l, c, 'n');
          if ((h === 'qui quadrado' || h === 'exato de fisher') && cmP_(v)) add(l, c, 'p');
        }
      });
    });
  } else if (nome === 'TABELA 4') {
    var reconhecidas = 0;
    m.forEach(function(row, l) {
      var label = cmNormal_(row[0]);
      var count = label === 'restauracoes com tempo registrado, n';
      var failure = label === 'falhas, n (%)';
      if (!count && !failure) return;
      reconhecidas++;
      for (var c = 2; c < row.length; c++) {
        if (count && cmInteiro_(row[c])) add(l, c, 'n');
        if (failure && typeof row[c] === 'string') {
          var found = /^(\s*)(\d+)(\s*\(\s*\d+(?:[.,]\d+)?\s*%?\s*\)\s*)$/.exec(row[c]);
          if (found) add(l, c, 'n', [{ inicio: found[1].length, fim: found[1].length + found[2].length, numero: Number(found[2]) }]);
        }
      }
    });
    if (!reconhecidas) throw new Error('Linhas de contagem da TABELA 4 não reconhecidas; publicação bloqueada.');
  } else if (nome === 'TABELA 5') {
    if (!Array.isArray(r.cmCamposT5)) {
      throw new Error('Atualize também TABELA3.gs: faltam as coordenadas dos campos da TABELA 5. Nenhum dado foi escrito.');
    }
    r.cmCamposT5.forEach(function(p) {
      if (!Number.isSafeInteger(p.l) || !Number.isSafeInteger(p.c) || p.l < 1 || p.c < 1 ||
          !m[p.l - 1] || p.c > m[p.l - 1].length || typeof p.lote !== 'string' || !p.lote) {
        throw new Error('Coordenada inválida no mascaramento da TABELA 5.');
      }
      add(p.l - 1, p.c - 1, 'campo');
      a[a.length - 1].lote = p.lote;
    });
  } else if (nome === 'CONSORT') {
    m.forEach(function(row, l) { row.forEach(function(v, c) {
      if (typeof v !== 'string' || preservadas[(l+1) + ':' + (c+1)]) return;
      var re = /\bn\s*=\s*(\d+)(?![\d.,])/gi, found, ts = [];
      while ((found = re.exec(v)) !== null) {
        var ini = found.index + found[0].length - found[1].length;
        ts.push({ inicio: ini, fim: ini + found[1].length, numero: Number(found[1]) });
      }
      if (ts.length) add(l, c, 'n', ts);
    }); });
    if (!a.length) throw new Error('Contagens n = não localizadas no CONSORT; publicação bloqueada.');
  }
  // Cabeçalhos de grupos também trazem N como texto nas tabelas.
  if (nome === 'TABELA 3' || nome === 'TABELA 4' || nome === 'TABELA 5') {
    m.forEach(function(row, l) { row.forEach(function(v, c) {
      if (typeof v !== 'string' || vistos[(l+1) + ':' + (c+1)]) return;
      var re = /\bn\s*=\s*(\d+)(?![\d.,])/gi, found, ts = [];
      while ((found = re.exec(v)) !== null) {
        var ini = found.index + found[0].length - found[1].length;
        ts.push({ inicio: ini, fim: ini + found[1].length, numero: Number(found[1]) });
      }
      if (ts.length) add(l, c, 'n', ts);
    }); });
  }
  return a;
}
function cmSubstituir_(texto, ts) {
  var s = String(texto);
  ts.slice().reverse().forEach(function(t) { s = s.slice(0, t.inicio) + String(t.novo) + s.slice(t.fim); });
  return s;
}
function cmOffset_(offset, ts) {
  var delta = 0;
  ts.forEach(function(t) {
    if (offset >= t.fim) delta += String(t.novo).length - (t.fim - t.inicio);
    else if (offset > t.inicio) delta += Math.min(offset - t.inicio, String(t.novo).length) - (offset - t.inicio);
  });
  return offset + delta;
}
function cmMascarar_(nome, original) {
  var r = Object.assign({}, original);
  r.matriz = original.matriz.map(function(row) { return row.slice(); });
  r.formatos = JSON.parse(JSON.stringify(original.formatos || []));
  var registros = cmSelecionar_(nome, r), numeros = [];
  registros.forEach(function(x) {
    if (x.tipo !== 'n') return;
    if (x.trechos) x.trechos.forEach(function(t) { numeros.push({ alvo: t, valor: t.numero }); });
    else numeros.push({ alvo: x, valor: x.valor });
  });
  var mistura = numeros.map(function(n) { return n.valor; });
  for (var i = mistura.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * i), tmp = mistura[i];
    mistura[i] = mistura[j]; mistura[j] = tmp;
  }
  numeros.forEach(function(n, i) { n.alvo.novo = mistura[i]; });
  // TABELA 5: cada campo é misturado apenas com o mesmo campo no mesmo estudo.
  // Não se movem linhas completas, pois isso preservaria as associações originais.
  var lotes = {};
  registros.filter(function(x) { return x.tipo === 'campo'; }).forEach(function(x) {
    (lotes[x.lote] || (lotes[x.lote] = [])).push(x);
  });
  Object.keys(lotes).forEach(function(chave) {
    var celulas = lotes[chave], valores = celulas.map(function(x) { return x.valor; });
    for (var i = valores.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * i), tmp = valores[i];
      valores[i] = valores[j]; valores[j] = tmp;
    }
    celulas.forEach(function(x, i) { x.novo = valores[i]; });
  });
  var usados = {};
  registros.forEach(function(x) {
    if (x.tipo === 'p') {
      var cod = 10000 + Math.floor(Math.random() * 90000);
      // Sondagem limitada, mesmo se o gerador aleatório repetir valores.
      while (usados[cod]) cod = cod === 99999 ? 10000 : cod + 1;
      usados[cod] = true; x.novo = String(cod);
      (r.cmTextos || (r.cmTextos = [])).push({ l: x.l, c: x.c });
    } else if (x.trechos) x.novo = cmSubstituir_(x.valor, x.trechos);
    r.matriz[x.l - 1][x.c - 1] = x.novo;
    r.formatos.forEach(function(f) {
      if (f.tipo !== 'rico' || f.l !== x.l || f.c !== x.c) return;
      if (String(f.valor.texto) !== String(x.valor)) throw new Error('Texto rico divergente em ' + nome + '.');
      f.valor.texto = String(x.novo);
      f.valor.negritos = (f.valor.negritos || []).map(function(par) {
        return x.trechos ? [cmOffset_(par[0], x.trechos), cmOffset_(par[1], x.trechos)] : [0, String(x.novo).length];
      });
    });
  });
  return { render: r, registros: registros };
}
function cmBackup_(nome, registros, ciclo) {
  cmAutorizar_(false);
  var json = JSON.stringify({ versao: 4, planilha: CM_CFG.planilha, aba: nome,
    data: new Date().toISOString(), ciclo: ciclo, registros: registros });
  var arquivo = DriveApp.createFile(Utilities.newBlob(json, 'application/json',
    'RRT_V4_' + nome.replace(/ /g, '_') + '_' + Utilities.getUuid() + '.json'));
  var dono = arquivo.getOwner();
  if (!dono || dono.getEmail().toLowerCase() !== CM_CFG.email ||
      arquivo.getSharingAccess() !== DriveApp.Access.PRIVATE ||
      arquivo.getEditors().length || arquivo.getViewers().length ||
      arquivo.getBlob().getDataAsString('UTF-8') !== json) {
    throw new Error('Backup não validado. Escrita da aba cancelada.');
  }
  return arquivo.getId();
}
// Chamada antes de clear/setValues, com a trava dos geradores já adquirida.
function cmPreparar_(ss, nome, r) {
  if (CM_CFG.abas.indexOf(nome) < 0) return r;
  if (ss.getId() !== CM_CFG.planilha) throw new Error('Planilha de destino incorreta.');
  var e = cmEstado_(), anterior = cmDiario_(nome), pacote = null;
  var backup = anterior ? anterior.backup : null;
  if (e.ligado) {
    pacote = cmMascarar_(nome, r);
    backup = cmBackup_(nome, pacote.registros, e.ciclo);
  }
  cmProps_().setProperty(cmChaveAba_(nome), JSON.stringify({ ciclo: e.ciclo, ligado: e.ligado,
    revisao: CM_CFG.revisao, pendente: true, backup: backup, anterior: anterior ? anterior.backup : null,
    iniciado: new Date().toISOString() }));
  return pacote ? pacote.render : r;
}
function cmExecutores_() {
  return [
    { nome: 'CONSORT', executar: csExecutar_ },
    { nome: 'TABELA 3', executar: t1Executar_ },
    { nome: 'TABELA 4', executar: t2Executar_ },
    { nome: 'TABELA 5', executar: t3Executar_ }
  ];
}
function cmReconstruir_(forcar) {
  var anteriores = CM_FORCE, erros = [], mensagens = [];
  CM_FORCE = !!forcar;
  try {
    cmExecutores_().forEach(function(item) {
      try {
        var msg = item.executar(false);
        if (cmPrecisaEscreverSemForca_(item.nome)) throw new Error(msg || 'Geração não concluída.');
        mensagens.push(msg);
      } catch (err) { erros.push(item.nome + ': ' + err.message); }
    });
  } finally { CM_FORCE = anteriores; }
  if (erros.length) throw new Error('PAROU / atualização parcial. Execute RRT_REPROCESSAR para retomar. ' + erros.join(' | '));
  return mensagens.join('\n');
}
function cmPrecisaEscreverSemForca_(nome) {
  var e = cmEstado_(), d = cmDiario_(nome);
  return !d || d.pendente || d.ciclo !== e.ciclo || d.revisao !== CM_CFG.revisao;
}
function cmMudar_(ligado) {
  cmAutorizar_(true); cmSS_();
  var lock = cmTrava_();
  if (!lock.tryLock(30000)) throw new Error('Outra geração em andamento; tente novamente.');
  try {
    var raw = cmProps_().getProperty(CM_CFG.chave);
    var antigo = raw ? cmEstado_() : null;
    if (!antigo || antigo.ligado !== ligado) {
      cmProps_().setProperty(CM_CFG.chave, JSON.stringify({ versao: 4, planilha: CM_CFG.planilha,
        ligado: ligado, ciclo: Utilities.getUuid(), data: new Date().toISOString() }));
    }
    // Também recupera instalação/execução parcial sem reutilizar backups V3.
    cmReconstruir_(false);
    var msg = ligado ? 'LIGADO. Novas gerações manterão o mascaramento.' :
      'DESLIGADO. Resultados originais recalculados a partir da base e do corte atuais.';
    console.log(msg); return msg;
  } finally { lock.releaseLock(); }
}
function RRT_INSTALAR_E_LIGAR() { return cmMudar_(true); }
function RRT_LIGAR() { return cmMudar_(true); }
function RRT_DESLIGAR() { return cmMudar_(false); }
function RRT_REPROCESSAR() {
  cmAutorizar_(true); cmSS_();
  var lock = cmTrava_();
  if (!lock.tryLock(30000)) throw new Error('Outra geração em andamento.');
  try { var msg = cmReconstruir_(true); console.log(msg); return msg; }
  finally { lock.releaseLock(); }
}
function RRT_DIAGNOSTICO() {
  cmAutorizar_(true); cmSS_();
  var lock = cmTrava_();
  if (!lock.tryLock(30000)) throw new Error('Outra geração em andamento.');
  try {
    var e = cmEstado_();
    var linhas = ['Modo desejado: ' + (e.ligado ? 'LIGADO' : 'DESLIGADO')];
    CM_CFG.abas.forEach(function(nome) {
      var d = cmDiario_(nome);
      linhas.push(nome + ': ' + (cmPrecisaEscreverSemForca_(nome) ? 'PENDENTE - execute RRT_REPROCESSAR' :
        'última escrita concluída') + (d && d.backup ? '; backup privado registrado' : ''));
    });
    if (cmProps_().getProperty('RRT_MASCARAMENTO_V3_BACKUP') || cmProps_().getProperty('RRT_NUMEROS_COFRE_ATIVO')) {
      linhas.push('Referência antiga preservada. A V4 não usa nem restaura o backup antigo.');
    }
    linhas.push('Diagnóstico de estado; não verifica alterações manuais nas células.');
    var msg = linhas.join('\n'); console.log(msg); return msg;
  } finally { lock.releaseLock(); }
}
/** Ponto já chamado por rodarRotinas e rotinaDiariaBanco na MESTRA. */
function tabelasAutomatico() {
  cmSS_();
  var lock = cmTrava_();
  if (!lock.tryLock(30000)) throw new Error('Outra geração em andamento.');
  try { return cmReconstruir_(false); }
  finally { lock.releaseLock(); }
}

// Executado DEPOIS dos formatos originais, antes de setValues.
function cmFormatos_(sh, r) {
  (r.cmTextos || []).forEach(function(p) { sh.getRange(p.l, p.c).setNumberFormat('@'); });
}


/* ======================================================================
 * MÓDULO 2 DE 5 — CONSORT.gs
 * ====================================================================== */

var CS_ABA_SAIDA = 'CONSORT';
var CS_VERSAO = 'public-v1';
var CS_REAGIR_PROPRIA_ABA = true;
var CS_ROTULO_FIGURA = 'Figura 1.';
var CS_TITULO = 'Fluxograma dos participantes do estudo {estudo}, segundo o CONSORT 2010{corte}.';
var CS_TITULO_PAINEL = 'Estudo {estudo}: preencha as células amarelas (corte e elegibilidade); o restante vem do banco.';
var CS_MESES_RECHAMADA = 6;
var CS_VAZIO_FIGURA = '____';
var CS_COR_ETAPA = '#dbe5f1';
var CS_COR_ENTRADA = '#fff2cc';
var CS_COR_STATUS = '#595959';
var CS_COR_AVISO = '#cc0000';
var CS_FORMATO_DATA = 'dd/mm/yyyy';
var CS_TITULOS_GRUPO = {
  '0': 'Grupo 1',
  '1': 'Grupo 2',
  '2': 'Grupo 3',
  '3': 'Grupo 4'
};
var CS_MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
var CS_EXTENSO = ['Nenhum', 'Um', 'Dois', 'Três', 'Quatro', 'Cinco', 'Seis', 'Sete', 'Oito', 'Nove',
'Dez'];
var CS_CAMPOS_PAINEL = [
  { chave: 'corte', rotulo: 'Corte do estudo (CONSORT e Tabelas 1 a 3)', tipo: 'modo',
    ajuda: 'auto: todas as restaurações e todo o acompanhamento até hoje. por período: usa a data inicial e a data final abaixo.' },
  { chave: 'inicio', rotulo: 'Data inicial (Data 0 a partir de)', tipo: 'data',
    ajuda: 'Entram as restaurações cimentadas (Data 0) a partir desta data. Dois cliques abrem o calendário.' },
  { chave: 'fim', rotulo: 'Data final (Data 0 e acompanhamento até)', tipo: 'data',
    ajuda: 'Entram as restaurações cimentadas até esta data; rechamadas, falhas e tempos depois dela são ignorados. Dois cliques abrem o calendário.' },
  { chave: 'coderest', rotulo: 'Coderest até (última restauração incluída)', tipo: 'numero',
    ajuda: 'Entram só as restaurações com coderest até este número. Vale junto com as datas. Vazio = sem limite.' },
  { chave: 'avaliados', rotulo: 'Avaliados para elegibilidade (pacientes)' },
  { chave: 'naoAtenderam', rotulo: 'Não atenderam aos critérios de inclusão' },
  { chave: 'recusaram', rotulo: 'Recusaram-se a participar' },
  { chave: 'outros', rotulo: 'Outros motivos' }
];
var CS_LARGURA_FAIXA = 34;
var CS_LARGURA_RECUO = 10;
var CS_LARGURAS_FIGURA = [290, 24, 290, 30, 250, 40];
var CS_CAMPOS = {
  codpac: ['codepac', 'codpac'],
  coderest: ['coderest'],
  estudo: ['estudo'],
  grupo: ['grupo'],
  dente: ['n⁰ dente', 'Nº dente', 'N° dente'],
  data0: ['Data 0'],
  dataSucesso: ['data sucesso'],
  tempoSucesso: ['tempo sucesso'],
  sucesso: ['sucesso'],
  dataSobrevida: ['data sobrevida'],
  tempoSobrevida: ['tempo sobrevida'],
  sobrevida: ['sobrevida']
};
var CS_ABAS_FONTE = ['BANCO DE ANÁLISE', 'BANCO DE ANALISE', 'Folha1', 'Planilha1',
'Sheet1'];
var CS_LINHA_CABECALHO = 1;
var CS_LINHA_DADOS = 2;
var CS_HORA_ROTINA = null;
var CS_MINUTOS_VIGIA = 30;
var CS_ATRASO_SEGUNDOS = 90;
var CS_JANELA_IGNORAR_MS = 45000;
var CS_LINHAS_ENTRE_TABELAS = 3;
var CS_MARCADORES_VAZIOS = ['.', '..', 'na', 'n/a', 'n.a.', 'nd', 'n/d', '-', '--', 's/i', '?',
  'sem informacao', 'sem informação', 'null', 'undefined'];
var CS_ROTULO_SEM_INFO = 'Sem informação';
var CS_ROTULO_TOTAL = 'Total';
var CS_ROTULO_SEM_GRUPO = 'grupo não informado\nsem sorteio';
var CS_ROTULO_LINHA_ESTUDO = 'estudo';
var CS_ROTULO_LINHA_GRUPO = 'grupo';
var CS_GRUPOS_POR_ESTUDO = [
  { estudo: '0', grupos: ['0', '1'] },
  { estudo: '1', grupos: ['2', '3'] }
];
var CS_ESTUDO_POR_TEXTO = { 'estudoa': '0', 'estudob': '1', 'studya': '0', 'studyb': '1', 'a': '0', 'b': '1' };
var CS_ROTULOS = {
  estudo: { '0': 'Estudo A', '1': 'Estudo B' },
  grupo: {
    '0': 'Grupo 1',
    '1': 'Grupo 2',
    '2': 'Grupo 3',
    '3': 'Grupo 4'
  }
};
var CS_NOME_ESTUDO = { '0': 'A', '1': 'B' };
var CS_COR_SECAO = '#efefef';
var CS_COR_VALOR_P = '#fff2cc';
var CS_CORES_DESTAQUE = [
  { chave: '0', nome: 'A', linha: '#cfe2f3', dados: '#6fa8dc' },
  { chave: '1', nome: 'B', linha: '#fff2cc', dados: '#ffe599' },
  { chave: 'sem', nome: 'sem sorteio', linha: '#d9d2e9', dados: '#b4a7d6' }
];
var CS_ACENTOS = {
  'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
  'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
  'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
  'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
  'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
  'ç': 'c', 'ñ': 'n', 'ý': 'y', 'ÿ': 'y'
};
var CS_CRITERIOS_FDI = {
  fraturaretencao: 'Fratura/perda de retenção',
  adaptacaomarginal: 'Desadaptação marginal',
  contatoproximal: 'Ausência de contato proximal',
  formacontorno: 'Forma/contorno inadequado',
  oclusaodesgaste: 'Oclusão/desgaste',
  cariemargem: 'Cárie na margem',
  defeitotecduromargem: 'Defeito de tecido duro na margem',
  hipersensibposop: 'Hipersensibilidade pós-operatória',
  brilhotextura: 'Brilho/textura',
  coloracaomarginal: 'Coloração marginal',
  estabcor: 'Estabilidade de cor',
  satisfacao: 'Insatisfação do paciente',
  avalradiografica: 'Alteração radiográfica'
};
var CS_ESCORE_FDI_FALHA = 4;
var CS_PAINEL_ABA = 'CONSORT';
var CS_PAINEL_POS = 'CS_ENTRADAS_POS';
var CS_PAINEL_VAL = 'CS_ENTRADAS_VAL';
var CS_MODO_AUTO = 'auto';
var CS_MODO_PERIODO = 'por período';
var CS_ESPERA_TRAVA_PAINEL_MS = 120000;
var CS_VANTAGEM_PAINEL_MS = 2000;
var CS_CODIGOS_FALHA_SUCESSO = [1, 2];
var CS_CODIGOS_FALHA_SOBREVIDA = [2];
var CS_DIAS_ANO = 365.25;
function GERAR_CONSORT() {
  return csExecutar_(true);
}
function consortRotina() {
  return csExecutar_(false);
}
function consortReconstruir() {
  csLimparAgendamento_();
  return csExecutar_(false);
}
function consortVigia() {
  var atual = csImpressao_();
  var visto = csProps_().getProperty('CS_IMPRESSAO') || '';
  if (atual && atual === visto && !cmPrecisaEscrever_(CS_ABA_SAIDA)) return;
  csExecutar_(false);
}
function consortOnChange(e) {
  if (csPainelMudou_()) return csAtualizarPeloPainel_();
  if (csOrigemMudanca_(e) !== 'fonte') return;
  if (csGeracaoRecente_()) return;
  csAgendar_();
}
function configurarConsort(silencioso) {
  var ss = csSS_();
  desligarConsort(true);
  csProps_().deleteProperty('CS_DESLIGADO');
  var hora = csHoraRotina_();
  var falhas = [];
  var pedidos = [
    { rotulo: 'reconstrução diária', criar: function() {
      ScriptApp.newTrigger('consortRotina').timeBased().everyDays(1).atHour(hora).create();
    } },
    { rotulo: 'reação a edições', criar: function() {
      ScriptApp.newTrigger('consortOnChange').forSpreadsheet(ss).onChange().create();
    } },
    { rotulo: 'vigia', criar: function() {
      ScriptApp.newTrigger('consortVigia').timeBased().everyMinutes(CS_MINUTOS_VIGIA).create();
    } }
  ];
  pedidos.forEach(function(p) {
    try {
      p.criar();
    } catch (err) {
      falhas.push(p.rotulo + ' (' + err.message + ')');
    }
  });
  var msg = 'Gatilhos da aba ' + CS_ABA_SAIDA + ' instalados: vigia a cada ' +
  CS_MINUTOS_VIGIA +
    ' min, reação a edições no banco e no painel do CONSORT e reconstrução diária às ' + hora
  + 'h.';
  if (falhas.length) {
    msg = 'ATENÇÃO: não foi possível instalar ' + falhas.join('; ') + '. ' +
      'O projeto aceita no máximo 20 gatilhos por conta; confira em Acionadores. ' + msg;
  }
  if (silencioso !== true) csAvisar_(msg);
  return msg;
}
function desligarConsort(interno) {
  var alvos = ['consortRotina', 'consortOnChange', 'consortReconstruir', 'consortVigia'];
  var removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (alvos.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  csProps_().deleteProperty('CS_AGENDADO');
  if (interno !== true) csProps_().setProperty('CS_DESLIGADO', '1');
  return removidos + ' gatilho(s) da aba ' + CS_ABA_SAIDA + ' removido(s).';
}
function csGruposEstudos_(sh, blocos) {
  if (!sh || !blocos || !blocos.estudos) return;
  var nC = sh.getMaxColumns();
  var quer = [0];
  for (var c = 1; c <= nC; c++) quer.push(0);
  blocos.estudos.forEach(function(b) {
    for (var k = b.c1; k <= Math.min(b.c2, nC); k++) quer[k] = 1;
  });
  var atual = [0];
  var igual = true;
  for (var c2 = 1; c2 <= nC; c2++) {
    atual.push(sh.getColumnGroupDepth(c2));
    if (atual[c2] !== quer[c2]) igual = false;
  }
  if (igual) return;
  var maximo = Math.max.apply(null, atual);
  for (var d = maximo; d >= 1; d--) {
    csFaixas_(nC, function(k) { return atual[k] >= d; }).forEach(function(f) {
      sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(-1);
    });
  }
  sh.showColumns(1, nC);
  csFaixas_(nC, function(k) { return quer[k] === 1; }).forEach(function(f) {
    sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(1);
  });
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
}
function csFaixas_(n, teste) {
  var faixas = [];
  var ini = 0;
  for (var c = 1; c <= n + 1; c++) {
    var ok = c <= n && teste(c);
    if (ok && !ini) ini = c;
    if (!ok && ini) {
      faixas.push([ini, c - 1]);
      ini = 0;
    }
  }
  return faixas;
}
function csDestaque_(registros) {
  var mapa = {};
  registros.forEach(function(r) {
    if (!r.pacote) return;
    var chave = r.grupo ? r.estudo : 'sem';
    if (!csCorDestaque_(chave)) return;
    if (!mapa[chave]) mapa[chave] = {};
    if (!mapa[chave][r.pacote.linha]) mapa[chave][r.pacote.linha] = {};
    var celulas = mapa[chave][r.pacote.linha];
    for (var c in r.pacote.usadas) {
      if (r.pacote.usadas.hasOwnProperty(c)) celulas[c] = true;
    }
  });
  return mapa;
}
function csCorDestaque_(chave) {
  for (var i = 0; i < CS_CORES_DESTAQUE.length; i++) {
    if (CS_CORES_DESTAQUE[i].chave === chave) return CS_CORES_DESTAQUE[i];
  }
  return null;
}
function csRegrasDestaque_(mapa, nC) {
  var dados = [];
  var linhas = [];
  CS_CORES_DESTAQUE.forEach(function(cor) {
    var porLinha = mapa[cor.chave];
    var numeros = porLinha ? Object.keys(porLinha).map(Number) : [];
    if (!numeros.length) return;
    var ultima = Math.max.apply(null, numeros);
    var porTrecho = {};
    numeros.forEach(function(l) {
      csFaixas_(nC, function(c) { return !!porLinha[l][c - 1]; }).forEach(function(t) {
        var k = t[0] + ':' + t[1];
        if (!porTrecho[k]) porTrecho[k] = {};
        porTrecho[k][l] = true;
      });
    });
    var a1Dados = [];
    Object.keys(porTrecho).forEach(function(k) {
      var c = k.split(':').map(Number);
      csFaixas_(ultima, function(l) { return !!porTrecho[k][l]; }).forEach(function(t) {
        a1Dados.push(csA1_(t[0], c[0], t[1] - t[0] + 1, c[1] - c[0] + 1));
      });
    });
    var a1Linhas = csFaixas_(ultima, function(l) { return !!porLinha[l]; }).map(function(t) {
      return csA1_(t[0], 1, t[1] - t[0] + 1, nC);
    });
    if (a1Dados.length) dados.push({ formula: csFormulaDestaque_(cor.nome, 'dados lidos'), cor:
      cor.dados, a1: a1Dados.sort() });
    linhas.push({ formula: csFormulaDestaque_(cor.nome, 'linha analisada'), cor: cor.linha, a1:
      a1Linhas.sort() });
  });
  return { dados: dados, linhas: linhas };
}
function csFormulaDestaque_(categoria, tipo) {
  return '=N("RRT · ' + CS_ABA_SAIDA + ' · ' + categoria + ' · ' + tipo + '")=0';
}
function csMarcaDestaque_(regra) {
  try {
    var cond = regra.getBooleanCondition();
    if (!cond) return null;
    var m = String((cond.getCriteriaValues() || [])[0] || '').replace(/\s+/g, ' ')
      .match(/^\s*=\s*N\(\s*"RRT\s*·\s*(.*?)\s*·\s*(.*?)\s*·\s*(dados lidos|linha analisada)"\s*\)\s*=\s*0\s*$/i);
    return m ? { dono: m[1].trim().toUpperCase(), dados: m[3].toLowerCase() === 'dados lidos' } :
    null;
  } catch (err) {
    return null;
  }
}
function csAplicarDestaque_(sh, mapa) {
  if (!sh || !mapa) return '';
  try {
    var desejo = csRegrasDestaque_(mapa, sh.getLastColumn());
    var outras = [];
    var dados = [];
    var linhas = [];
    var minhas = [];
    var fase = 0;
    var emOrdem = true;
    sh.getConditionalFormatRules().forEach(function(regra) {
      var marca = csMarcaDestaque_(regra);
      var f = !marca ? 0 : (marca.dados ? 1 : 2);
      if (f < fase) emOrdem = false;
      fase = Math.max(fase, f);
      if (!marca) outras.push(regra);
      else if (marca.dono === CS_ABA_SAIDA.toUpperCase()) minhas.push(regra);
      else if (marca.dados) dados.push(regra);
      else linhas.push(regra);
    });
    if (emOrdem && csMesmasRegras_(minhas, desejo.dados.concat(desejo.linhas))) return '';
    var nova = function(d) {
      return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(d.formula).setBackground
      (d.cor)
        .setRanges(sh.getRangeList(d.a1).getRanges()).build();
    };
    sh.setConditionalFormatRules(outras.concat(dados, desejo.dados.map(nova), linhas,
    desejo.linhas.map(nova)));
    return '';
  } catch (err) {
    return ' ATENÇÃO: não foi possível colorir as linhas analisadas na aba ' + sh.getName() + ' ('
    + err.message + ').';
  }
}
function csMesmasRegras_(atuais, desejadas) {
  if (atuais.length !== desejadas.length) return false;
  var chave = function(formula, cor, a1) {
    return JSON.stringify([String(formula || '').replace(/\s+/g, '').toLowerCase(), String(cor ||
    '').toLowerCase(),
      a1.map(function(x) { return String(x).toUpperCase(); }).sort()]);
  };
  var a = atuais.map(function(regra) {
    var cond = regra.getBooleanCondition();
    return chave((cond.getCriteriaValues() || [])[0], cond.getBackground(),
      regra.getRanges().map(function(r) { return r.getA1Notation(); }));
  }).sort();
  var d = desejadas.map(function(x) { return chave(x.formula, x.cor, x.a1); }).sort();
  return JSON.stringify(a) === JSON.stringify(d);
}
function csSS_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Nenhuma planilha ativa.');
  return ss;
}
function csProps_() {
  return PropertiesService.getDocumentProperties();
}
function csHoraRotina_() {
  if (CS_HORA_ROTINA !== null && CS_HORA_ROTINA !== undefined) return Number(CS_HORA_ROTINA);
  if (typeof HORA_ROTINA_BANCO !== 'undefined' && HORA_ROTINA_BANCO !== null) {
    return (Number(HORA_ROTINA_BANCO) + 1) % 24;
  }
  return 7;
}
function csAgendar_() {
  var props = csProps_();
  var pendente = Number(props.getProperty('CS_AGENDADO') || 0);
  var agora = Date.now();
  if (pendente && agora - pendente < CS_ATRASO_SEGUNDOS * 1000 * 3) return;
  props.setProperty('CS_AGENDADO', String(agora));
  ScriptApp.newTrigger('consortReconstruir').timeBased().after(CS_ATRASO_SEGUNDOS *
  1000).create();
}
function csLimparAgendamento_() {
  csProps_().deleteProperty('CS_AGENDADO');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'consortReconstruir') ScriptApp.deleteTrigger(t);
  });
}
function csGeracaoRecente_() {
  var ultima = Number(csProps_().getProperty('CS_ULTIMA_GERACAO') || 0);
  return ultima && Date.now() - ultima < CS_JANELA_IGNORAR_MS;
}
function csOrigemMudanca_(e) {
  var ignorados = ['FORMAT', 'OTHER'];
  if (e && e.changeType && ignorados.indexOf(String(e.changeType)) >= 0) return '';
  var ativa = null;
  try {
    ativa = (e && e.source ? e.source : csSS_()).getActiveSheet();
  } catch (err) {
    return 'fonte';
  }
  if (!ativa) return 'fonte';
  var nome = ativa.getName();
  if (nome === CS_PAINEL_ABA) return 'painel';
  if (nome === CS_ABA_SAIDA) return CS_REAGIR_PROPRIA_ABA ? 'painel' : '';
  return CS_ABAS_FONTE.indexOf(nome) >= 0 ? 'fonte' : '';
}
function csImpressao_() {
  var ss = csSS_();
  try {
    return 'drv:' + DriveApp.getFileById(ss.getId()).getLastUpdated().getTime();
  } catch (err) {}
  try {
    var sh = csAbaFonte_(ss);
    return 'dim:' + sh.getLastRow() + 'x' + sh.getLastColumn();
  } catch (err2) {}
  return '';
}
function csExecutar_(interativo, doPainel) {
  var trava = cmTrava_();
  if (!trava.tryLock(doPainel ? CS_ESPERA_TRAVA_PAINEL_MS : 30000)) {
    if (doPainel) {
      csAgendar_();
      return 'Outra execução em andamento; reconstrução da aba ' + CS_ABA_SAIDA + ' agendada.';
    }
    var ocupado = 'Outra execução da aba ' + CS_ABA_SAIDA + ' está em andamento.';
    if (interativo) csAvisar_(ocupado);
    return ocupado;
  }
  try {
    var ss = csSS_();
    var fonte = csAbaFonte_(ss);
    var entrada = csLerFonte_(fonte);
    var render = csMontar_(entrada, ss);
    var assinatura = csAssinatura_(render);
    var props = csProps_();
    var existente = ss.getSheetByName(CS_ABA_SAIDA);
    if (existente && !props.getProperty('CS_ASSINATURA') && existente.getLastRow() > 0) {
      var alheia = 'A aba "' + CS_ABA_SAIDA + '" já existe com conteúdo que não foi gerado por este script; nada foi alterado. ' +
        'Renomeie essa aba ou troque CS_ABA_SAIDA no início do arquivo e rode de novo.';
      if (interativo) throw new Error(alheia);
      return alheia;
    }
    if (!cmPrecisaEscrever_(CS_ABA_SAIDA) && !interativo && existente && assinatura === props.getProperty('CS_ASSINATURA')) {
      props.setProperty('CS_IMPRESSAO', csImpressao_());
      if (render.chaveCorte) props.setProperty('CS_CORTE_USADO', render.chaveCorte);
      if (render.chavePainel !== undefined) props.setProperty('CS_PAINEL_LIDO',
      render.chavePainel);
      csGruposEstudos_(existente, render.blocos);
      return 'Aba ' + CS_ABA_SAIDA + ' já está em dia.' + csAplicarDestaque_(fonte,
      render.destaque);
    }
    csEscrever_(ss, render);
    if (render.depoisDeEscrever) render.depoisDeEscrever(ss);
    csGruposEstudos_(ss.getSheetByName(CS_ABA_SAIDA), render.blocos);
    var avisoCores = csAplicarDestaque_(fonte, render.destaque);
    props.setProperty('CS_ASSINATURA', assinatura);
    props.setProperty('CS_ULTIMA_GERACAO', String(Date.now()));
    props.setProperty('CS_IMPRESSAO', csImpressao_());
    if (render.chaveCorte) props.setProperty('CS_CORTE_USADO', render.chaveCorte);
    if (render.chavePainel !== undefined) props.setProperty('CS_PAINEL_LIDO',
    render.chavePainel);
    cmConcluir_(CS_ABA_SAIDA);
    var msg = 'Aba ' + CS_ABA_SAIDA + ' atualizada. ' + (render.resumo || '') + ' Fonte: ' +
    fonte.getName() + '.';
    if (render.faltando && render.faltando.length) {
      msg += ' ATENÇÃO: colunas não localizadas na origem: ' + render.faltando.join(', ') + '.';
    }
    msg += avisoCores;
    if (interativo) {
      msg += csGarantirGatilhos_();
      csAvisar_(msg);
    }
    return msg;
  } finally {
    trava.releaseLock();
  }
}
function csGarantirGatilhos_() {
  try {
    if (csProps_().getProperty('CS_DESLIGADO')) {
      return ' Atualização automática desta aba desligada por desligarConsort; para religar, rode configurarConsort.';
    }
    var instalado = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'consortOnChange';
    });
    return instalado ? '' : ' ' + configurarConsort(true);
  } catch (err) {
    return ' Não foi possível instalar a atualização automática (' + err.message + '); rode configurarConsort.';
  }
}
function csAvisar_(texto) {
  try {
    SpreadsheetApp.getActive().toast(texto, CS_ABA_SAIDA, 8);
  } catch (err) {}
}
function csAbaFonte_(ss) {
  for (var i = 0; i < CS_ABAS_FONTE.length; i++) {
    var sh = ss.getSheetByName(CS_ABAS_FONTE[i]);
    if (sh) return sh;
  }
  throw new Error('Nenhuma aba de origem encontrada. Procurei por: ' +
  CS_ABAS_FONTE.join(', ') + '.');
}
function csLerFonte_(sh) {
  var nCols = sh.getLastColumn();
  var nLinhas = sh.getLastRow() - CS_LINHA_DADOS + 1;
  if (nCols < 2) throw new Error('A aba "' + sh.getName() + '" não tem colunas suficientes.');
  var cabecalho = sh.getRange(CS_LINHA_CABECALHO, 1, 1, nCols).getDisplayValues()[0];
  var dados = nLinhas > 0 ? sh.getRange(CS_LINHA_DADOS, 1, nLinhas, nCols).getValues() :
  [];
  return { cabecalho: cabecalho, dados: dados };
}
function csAssinatura_(render) {
  var texto = JSON.stringify([CS_VERSAO, render.matriz, render.formatos, render.larguras,
  render.alturas]);
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, texto,
    Utilities.Charset.UTF_8);
    var hex = '';
    for (var i = 0; i < bytes.length; i++) {
      var b = (bytes[i] + 256) % 256;
      hex += (b < 16 ? '0' : '') + b.toString(16);
    }
    return hex;
  } catch (err) {
    return 'len' + texto.length;
  }
}
function csChave_(v) {
  var s = String(v === null || v === undefined ? '' : v).toLowerCase();
  var saida = '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    saida += CS_ACENTOS[c] || c;
  }
  return saida.replace(/[^a-z0-9]+/g, '');
}
function csVazio_(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return isNaN(v.getTime());
  var s = String(v).trim();
  if (!s) return true;
  return CS_MARCADORES_VAZIOS.indexOf(s.toLowerCase()) >= 0;
}
function csTexto_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + ('0' + v.getDate()).slice(-2);
  }
  return String(v).trim();
}
function csNumero_(v) {
  if (csVazio_(v)) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  if (v instanceof Date) return null;
  var s = String(v).trim().replace(/\s/g, '').replace(/^r\$/i, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function csData_(v) {
  if (csVazio_(v)) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var iso = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var br = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
  return null;
}
function csIndices_(cabecalho, campos) {
  var porChave = {};
  for (var c = 0; c < cabecalho.length; c++) {
    var k = csChave_(cabecalho[c]);
    if (!k) continue;
    if (!porChave[k]) porChave[k] = [];
    porChave[k].push(c);
  }
  var mapa = {};
  for (var campo in campos) {
    if (!campos.hasOwnProperty(campo)) continue;
    var achados = [];
    campos[campo].forEach(function(alias) {
      (porChave[csChave_(alias)] || []).forEach(function(idx) {
        if (achados.indexOf(idx) < 0) achados.push(idx);
      });
    });
    mapa[campo] = achados;
  }
  return mapa;
}
function csFaltando_(idx, campos, exigidos) {
  var faltando = [];
  exigidos.forEach(function(alternativas) {
    var achou = alternativas.some(function(c) { return idx[c] && idx[c].length; });
    if (!achou) faltando.push(campos[alternativas[0]][0]);
  });
  return faltando;
}
function csValorBruto_(linha, indices) {
  if (!indices || !indices.length) return '';
  for (var i = 0; i < indices.length; i++) {
    var v = linha[indices[i]];
    if (!csVazio_(v)) {
      if (linha.usadas) linha.usadas[indices[i]] = true;
      return v;
    }
  }
  return '';
}
function csUsar_(linha, colunas) {
  if (!linha.usadas) return;
  colunas.forEach(function(c) {
    if (typeof c === 'number' && c >= 0 && !csVazio_(linha[c])) linha.usadas[c] = true;
  });
}
function csCodigoGrupo_(v) {
  if (csVazio_(v)) return '';
  var n = csNumero_(v);
  return n !== null ? String(n) : csTexto_(v);
}
function csEstudoEfetivo_(estudo, grupo) {
  if (!csVazio_(estudo)) {
    var n = csNumero_(estudo);
    if (n !== null) return String(n);
    var k = csChave_(estudo);
    if (CS_ESTUDO_POR_TEXTO.hasOwnProperty(k)) return CS_ESTUDO_POR_TEXTO[k];
    return csTexto_(estudo);
  }
  for (var i = 0; i < CS_GRUPOS_POR_ESTUDO.length; i++) {
    if (CS_GRUPOS_POR_ESTUDO[i].grupos.indexOf(grupo) >= 0) return
    CS_GRUPOS_POR_ESTUDO[i].estudo;
  }
  return '';
}
function csChavePaciente_(codpac, numeroLinha) {
  if (csVazio_(codpac)) return '#linha' + numeroLinha;
  var n = csNumero_(codpac);
  if (n !== null) return 'p' + n;
  return 'p' + (csChave_(codpac) || csTexto_(codpac));
}
function csEstudos_(registros) {
  var lista = [];
  CS_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var grupos = cfg.grupos.slice();
    var extras = [];
    var semGrupo = 0;
    registros.forEach(function(r) {
      if (r.estudo !== cfg.estudo) return;
      if (!r.grupo) { semGrupo++; return; }
      if (grupos.indexOf(r.grupo) < 0 && extras.indexOf(r.grupo) < 0) extras.push(r.grupo);
    });
    extras.sort(csCompararChaves_);
    lista.push({ estudo: cfg.estudo, grupos: grupos.concat(extras), temSemGrupo: semGrupo > 0
});
  });
  return lista;
}
function csRotuloGrupo_(g) {
  if (CS_ROTULOS.grupo.hasOwnProperty(g)) return CS_ROTULOS.grupo[g];
  return 'Grupo ' + g;
}
function csRotuloEstudo_(e) {
  return CS_ROTULOS.estudo.hasOwnProperty(e) ? CS_ROTULOS.estudo[e] : 'estudo ' + e;
}
function csNomeEstudo_(e) {
  return CS_NOME_ESTUDO.hasOwnProperty(e) ? CS_NOME_ESTUDO[e] : String(e);
}
function csArredondarPar_(v, casas) {
  var f = Math.pow(10, casas);
  var x = v * f;
  var piso = Math.floor(x);
  var frac = x - piso;
  var r;
  if (Math.abs(frac - 0.5) < 1e-9) r = (piso % 2 === 0) ? piso : piso + 1;
  else r = Math.round(x);
  return r / f;
}
function csFmt_(v, casas) {
  if (v === null || v === undefined || isNaN(v)) return '';
  return csArredondarPar_(v, casas).toFixed(casas).replace('.', ',');
}
function csNumeroP_(p) {
  if (p === null || p === undefined || isNaN(p)) return '';
  if (p < 0.001) return '< 0,001';
  return csArredondarPar_(p, 3);
}
function csPlural_(n, singular, plural) {
  return n === 1 ? singular : plural;
}
function csContar_(obj) {
  var n = 0;
  for (var k in obj) {
    if (obj.hasOwnProperty(k)) n++;
  }
  return n;
}
function csLinhaVazia_(n) {
  var linha = [];
  for (var i = 0; i < n; i++) linha.push('');
  return linha;
}
function csCompararChaves_(a, b) {
  var na = csNumero_(a);
  var nb = csNumero_(b);
  if (na !== null && nb !== null) return na - nb;
  if (na !== null) return -1;
  if (nb !== null) return 1;
  return a < b ? -1 : (a > b ? 1 : 0);
}
function csLetra_(c) {
  var s = '';
  while (c > 0) {
    var m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}
function csA1_(l, c, nl, nc) {
  var a = csLetra_(c) + l;
  nl = nl || 1;
  nc = nc || 1;
  if (nl === 1 && nc === 1) return a;
  return a + ':' + csLetra_(c + nc - 1) + (l + nl - 1);
}
function csNovoRender_() {
  return { matriz: [], nColunas: 1, formatos: [], larguras: {}, alturas: {}, faltando: [], resumo: '' };
}
function csPor_(render, l, c, valor) {
  while (render.matriz.length < l) render.matriz.push([]);
  var linha = render.matriz[l - 1];
  while (linha.length < c) linha.push('');
  linha[c - 1] = valor;
}
function csFechar_(render) {
  var nC = render.nColunas;
  render.matriz.forEach(function(linha) {
    if (linha.length > nC) nC = linha.length;
  });
  render.nColunas = nC;
  render.matriz.forEach(function(linha) {
    while (linha.length < nC) linha.push('');
  });
  return render;
}
function csMarcar_(render, tipo, l, c, nl, nc, valor) {
  render.formatos.push({ tipo: tipo, l: l, c: c, nl: nl || 1, nc: nc || 1, valor: valor });
}
function csLargura_(render, col, px) {
  if (!render.larguras[col] || render.larguras[col] < px) render.larguras[col] = px;
}
function csAltura_(render, linha, px) {
  if (!render.alturas[linha] || render.alturas[linha] < px) render.alturas[linha] = px;
}
function csContarLinhas_(texto, larguraPx, pxPorCaractere) {
  var porLinha = Math.max(1, Math.floor((larguraPx - 8) / (pxPorCaractere || 6.6)));
  var total = 0;
  String(texto).split('\n').forEach(function(paragrafo) {
    var palavras = paragrafo.split(/\s+/).filter(function(p) { return p; });
    if (!palavras.length) { total++; return; }
    var atual = 0;
    var linhas = 1;
    palavras.forEach(function(p) {
      var tam = p.length;
      if (atual === 0) {
        atual = tam;
      } else if (atual + 1 + tam <= porLinha) {
        atual += 1 + tam;
      } else {
        linhas++;
        atual = tam;
      }
      while (atual > porLinha) {
        linhas++;
        atual -= porLinha;
      }
    });
    total += linhas;
  });
  return total;
}
function csAlturaPara_(nLinhas) {
  return nLinhas <= 1 ? 21 : 21 + 15 * (nLinhas - 1);
}
function csEscrever_(ss, r) {
  var nomeAba = r.aba || CS_ABA_SAIDA;
  r = cmPreparar_(ss, nomeAba, r);
  var sh = ss.getSheetByName(nomeAba);
  if (!sh) sh = ss.insertSheet(nomeAba);
  var preservar = r.preservar || [];
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  if (preservar.length) sh.clearFormats();
  else sh.clear();
  sh.clearNotes();
  if (r.validacoes) sh.getRange(1, 1, sh.getMaxRows(),
  sh.getMaxColumns()).clearDataValidations();
  if (!r.matriz.length) r.matriz.push(csLinhaVazia_(Math.max(1, r.nColunas)));
  csFechar_(r);
  var nL = r.matriz.length;
  var nC = r.nColunas;
  if (sh.getMaxColumns() < nC) sh.insertColumnsAfter(sh.getMaxColumns(), nC -
  sh.getMaxColumns());
  if (sh.getMaxRows() < nL) sh.insertRowsAfter(sh.getMaxRows(), nL - sh.getMaxRows());
  var tudo = sh.getRange(1, 1, nL, nC);
  csAplicarGrupos_(sh, r.formatos, function(f) { return f.tipo === 'formato'; });
  cmFormatos_(sh, r);
  if (preservar.length) {
    csRetangulos_(nL, nC, preservar).forEach(function(q) {
      sh.getRange(q.l, q.c, q.nl, q.nc).setValues(r.matriz.slice(q.l - 1, q.l - 1 +
      q.nl).map(function(linha) {
        return linha.slice(q.c - 1, q.c - 1 + q.nc);
      }));
    });
  } else {
    tudo.setValues(r.matriz);
  }
  tudo.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(false);
  r.formatos.forEach(function(f) {
    if (f.tipo === 'mesclar') sh.getRange(f.l, f.c, f.nl, f.nc).merge();
  });
  csAplicarGrupos_(sh, r.formatos, function(f) {
    return f.tipo !== 'formato' && f.tipo !== 'mesclar' && f.tipo !== 'rico';
  });
  var negrito = null;
  r.formatos.forEach(function(f) {
    if (f.tipo !== 'rico') return;
    if (!negrito) negrito = SpreadsheetApp.newTextStyle().setBold(true).build();
    var construtor = SpreadsheetApp.newRichTextValue().setText(f.valor.texto);
    f.valor.negritos.forEach(function(par) {
      if (par[1] > par[0]) construtor.setTextStyle(par[0], par[1], negrito);
    });
    sh.getRange(f.l, f.c).setRichTextValue(construtor.build());
  });
  (r.validacoes || []).forEach(function(v) {
    sh.getRange(v.l, v.c).setDataValidation(v.regra);
  });
  csMedidas_(nC, r.larguras, 62, function(ini, qtd, px) { sh.setColumnWidths(ini, qtd, px); });
  csMedidas_(nL, r.alturas, 21, function(ini, qtd, px) { sh.setRowHeights(ini, qtd, px); });
  if (r.congelarColunas) sh.setFrozenColumns(r.congelarColunas);
  if (r.congelarLinhas) sh.setFrozenRows(r.congelarLinhas);
  if (sh.getMaxRows() > nL) sh.deleteRows(nL + 1, sh.getMaxRows() - nL);
  if (sh.getMaxColumns() > nC) sh.deleteColumns(nC + 1, sh.getMaxColumns() - nC);
  SpreadsheetApp.flush();
}
function csRetangulos_(nL, nC, preservar) {
  var furos = {};
  preservar.forEach(function(p) { furos[p.l + ',' + p.c] = true; });
  var segmentos = [];
  for (var c = 1; c <= nC; c++) {
    var lista = [];
    var ini = 0;
    for (var l = 1; l <= nL + 1; l++) {
      var livre = l <= nL && !furos[l + ',' + c];
      if (livre && !ini) ini = l;
      if (!livre && ini) {
        lista.push([ini, l - 1]);
        ini = 0;
      }
    }
    segmentos.push(JSON.stringify(lista));
  }
  var retangulos = [];
  var c0 = 1;
  for (var k = 2; k <= nC + 1; k++) {
    if (k <= nC && segmentos[k - 1] === segmentos[c0 - 1]) continue;
    JSON.parse(segmentos[c0 - 1]).forEach(function(seg) {
      retangulos.push({ l: seg[0], c: c0, nl: seg[1] - seg[0] + 1, nc: k - c0 });
    });
    c0 = k;
  }
  return retangulos;
}
function csMedidas_(n, mapa, padrao, aplicar) {
  var ini = 1;
  var atual = mapa[1] || padrao;
  for (var i = 2; i <= n + 1; i++) {
    var v = i <= n ? (mapa[i] || padrao) : null;
    if (v !== atual) {
      aplicar(ini, i - ini, atual);
      ini = i;
      atual = v;
    }
  }
}
function csAplicarGrupos_(sh, formatos, filtro) {
  var grupos = {};
  var ordem = [];
  formatos.forEach(function(f) {
    if (!filtro(f)) return;
    var chave = f.tipo + '|' + JSON.stringify(f.valor);
    if (!grupos[chave]) {
      grupos[chave] = { f: f, a1: [] };
      ordem.push(chave);
    }
    grupos[chave].a1.push(csA1_(f.l, f.c, f.nl, f.nc));
  });
  ordem.forEach(function(k) {
    var g = grupos[k];
    for (var i = 0; i < g.a1.length; i += 300) {
      csAplicarLista_(sh.getRangeList(g.a1.slice(i, i + 300)), g.f);
    }
  });
}
function csAplicarLista_(lista, f) {
  var v = f.valor;
  if (f.tipo === 'negrito') lista.setFontWeight(v ? 'bold' : 'normal');
  else if (f.tipo === 'italico') lista.setFontStyle(v ? 'italic' : 'normal');
  else if (f.tipo === 'fundo') lista.setBackground(v);
  else if (f.tipo === 'formato') lista.setNumberFormat(v);
  else if (f.tipo === 'alinhaH') lista.setHorizontalAlignment(v);
  else if (f.tipo === 'alinhaV') lista.setVerticalAlignment(v);
  else if (f.tipo === 'quebra') lista.setWrap(v);
  else if (f.tipo === 'tamanho') lista.setFontSize(v);
  else if (f.tipo === 'rotacao') lista.setTextRotation(v);
  else if (f.tipo === 'corTexto') lista.setFontColor(v);
  else if (f.tipo === 'borda') {
    if (v.estilo === 'tracejada') {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    } else {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null);
    }
  }
}
function csCaixa_(render, l, c, nl, nc, estilo) {
  csMarcar_(render, 'borda', l, c, nl, nc, { t: true, l: true, b: true, r: true, estilo: estilo || 'solida' });
}
function csLinhaBaixo_(render, l, c, nc) {
  csMarcar_(render, 'borda', l, c, 1, nc, { t: null, l: null, b: true, r: null, estilo: 'solida' });
}
function csLaterais_(render, l, c, nl) {
  csMarcar_(render, 'borda', l, c, nl, 1, { t: null, l: true, b: null, r: true, estilo: 'solida' });
}
function csBlocos_(cabecalho) {
  var inicios = [];
  for (var c = 0; c < cabecalho.length; c++) {
    var m = csChave_(cabecalho[c]).match(/^data(\d+)$/);
    if (m) inicios.push({ indice: Number(m[1]), col: c });
  }
  var blocos = [];
  inicios.forEach(function(b, i) {
    var fim = i + 1 < inicios.length ? inicios[i + 1].col : cabecalho.length;
    var bloco = { indice: b.indice, data: b.col, avaliacao: -1, causa: -1, criterios: [] };
    for (var c = b.col + 1; c < fim; c++) {
      var raiz = csChave_(cabecalho[c]).replace(/\d+$/, '');
      if (raiz === 'avaliacaofinal') {
        if (bloco.avaliacao < 0) bloco.avaliacao = c;
      } else if (raiz === 'causafalha') {
        if (bloco.causa < 0) bloco.causa = c;
      } else if (CS_CRITERIOS_FDI.hasOwnProperty(raiz)) {
        bloco.criterios.push({ col: c, rotulo: CS_CRITERIOS_FDI[raiz] });
      }
    }
    blocos.push(bloco);
  });
  return blocos;
}
function csLerBlocos_(linha, blocos) {
  return blocos.map(function(b) {
    var aval = b.avaliacao >= 0 ? csNumero_(linha[b.avaliacao]) : null;
    var causa = b.causa >= 0 ? csNumero_(linha[b.causa]) : null;
    var criticos = [];
    var colunasCriticas = [];
    b.criterios.forEach(function(cr) {
      var escore = csNumero_(linha[cr.col]);
      if (escore === null || escore < CS_ESCORE_FDI_FALHA) return;
      colunasCriticas.push(cr.col);
      if (criticos.indexOf(cr.rotulo) < 0) criticos.push(cr.rotulo);
    });
    return { data: csData_(linha[b.data]), avaliacao: aval, causa: causa, criticos: criticos,
      colunasCriticas: colunasCriticas };
  });
}
function csDiasEntre_(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}
function csLerPainel_(ss) {
  var valores = {};
  try {
    valores = JSON.parse(csProps_().getProperty(CS_PAINEL_VAL) || '{}') || {};
  } catch (err) {
    valores = {};
  }
  var lidos = csLerCelulasPainel_(ss);
  if (lidos) {
    for (var est in lidos) {
      if (!lidos.hasOwnProperty(est)) continue;
      if (!valores[est]) valores[est] = {};
      for (var campo in lidos[est]) {
        if (lidos[est].hasOwnProperty(campo)) valores[est][campo] = lidos[est][campo];
      }
    }
  }
  return { valores: valores, lidos: lidos };
}
function csLerCelulasPainel_(ss) {
  var pos = null;
  try {
    pos = JSON.parse(csProps_().getProperty(CS_PAINEL_POS) || 'null');
  } catch (err) {
    pos = null;
  }
  var sh = ss ? ss.getSheetByName(CS_PAINEL_ABA) : null;
  if (!sh || !pos) return null;
  var celulas = [];
  for (var est in pos) {
    if (!pos.hasOwnProperty(est)) continue;
    for (var campo in pos[est]) {
      if (!pos[est].hasOwnProperty(campo)) continue;
      var m = String(pos[est][campo]).match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      var c = 0;
      for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
      celulas.push({ est: est, campo: campo, l: Number(m[2]), c: c });
    }
  }
  if (!celulas.length) return null;
  var l1 = Infinity, c1 = Infinity, l2 = 0, c2 = 0;
  celulas.forEach(function(x) {
    l1 = Math.min(l1, x.l);
    c1 = Math.min(c1, x.c);
    l2 = Math.max(l2, x.l);
    c2 = Math.max(c2, x.c);
  });
  if (l2 > sh.getMaxRows() || c2 > sh.getMaxColumns()) return null;
  var grade = sh.getRange(l1, c1, l2 - l1 + 1, c2 - c1 + 1).getDisplayValues();
  var lidos = {};
  celulas.forEach(function(x) {
    var v = grade[x.l - l1][x.c - c1];
    if (!lidos[x.est]) lidos[x.est] = {};
    lidos[x.est][x.campo] = String(v === null || v === undefined ? '' : v).trim();
  });
  return lidos;
}
function csPainelMudou_() {
  try {
    var ss = csSS_();
    var props = csProps_();
    if (CS_ABA_SAIDA === CS_PAINEL_ABA) {
      var lidos = csLerCelulasPainel_(ss);
      return !!lidos && JSON.stringify(lidos) !== (props.getProperty('CS_PAINEL_LIDO') || '');
    }
    return csChaveCortes_(csCortes_(ss, new Date())) !==
    (props.getProperty('CS_CORTE_USADO') || '');
  } catch (err) {
    return false;
  }
}
function csAtualizarPeloPainel_() {
  if (CS_ABA_SAIDA !== CS_PAINEL_ABA) Utilities.sleep(CS_VANTAGEM_PAINEL_MS);
  return csExecutar_(false, true);
}
function csChaveCortes_(cortes) {
  return CS_GRUPOS_POR_ESTUDO.map(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (!c.periodo) return cfg.estudo + ':auto';
    return cfg.estudo + ':' + (c.coderest !== null ? c.coderest : '') + '#' + (c.inicio ?
    csDataBr_(c.inicio) : '') +
      '-' + (c.fim ? csDataBr_(c.fim) : '');
  }).join('|');
}
function csFaltaCoderest_(idx, cortes) {
  if (idx.coderest && idx.coderest.length) return false;
  return CS_GRUPOS_POR_ESTUDO.some(function(cfg) { return cortes.estudos[cfg.estudo].coderest !== null; });
}
function csModoCorte_(v) {
  var k = csChave_(v);
  return !k || k.indexOf('auto') === 0 ? CS_MODO_AUTO : CS_MODO_PERIODO;
}
function csDataEstrita_(v) {
  if (csVazio_(v)) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(),
  v.getDate());
  var s = String(v).trim();
  var br = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(\D|$)/);
  var iso = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(\D|$)/);
  var a, m, d;
  if (br) {
    a = Number(br[3]);
    m = Number(br[2]);
    d = Number(br[1]);
  } else if (iso) {
    a = Number(iso[1]);
    m = Number(iso[2]);
    d = Number(iso[3]);
  } else {
    return null;
  }
  var data = new Date(a, m - 1, d);
  if (data.getFullYear() !== a || data.getMonth() !== m - 1 || data.getDate() !== d) return null;
  return data;
}
function csDia_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}
function csCorte_(v, hoje) {
  var corte = { periodo: false, inicio: null, fim: null, coderest: null, aviso: '',
    referencia: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()) };
  if (!v || csModoCorte_(v.corte) !== CS_MODO_PERIODO) return corte;
  var inicio = csDataEstrita_(v.inicio);
  var fim = csDataEstrita_(v.fim);
  var coderest = csVazio_(v.coderest) ? null : csNumero_(v.coderest);
  var invalidos = [];
  if (!csVazio_(v.inicio) && !inicio) invalidos.push('a data inicial');
  if (!csVazio_(v.fim) && !fim) invalidos.push('a data final');
  if (!csVazio_(v.coderest) && (coderest === null || coderest < 1 || Math.round(coderest) !==
  coderest)) {
    invalidos.push('o coderest');
  }
  if (invalidos.length) {
    corte.aviso = 'não entendi ' + invalidos.join(' nem ') + '; corte ignorado, valendo auto.';
    return corte;
  }
  if (!inicio && !fim && coderest === null) {
    corte.aviso = 'preencha as datas e/ou o coderest; sem nenhum deles, vale auto.';
    return corte;
  }
  if (inicio && fim && inicio.getTime() > fim.getTime()) {
    corte.aviso = 'a data inicial é depois da data final; corte ignorado, valendo auto.';
    return corte;
  }
  corte.periodo = true;
  corte.inicio = inicio;
  corte.fim = fim;
  corte.coderest = coderest;
  if (fim && fim.getTime() < corte.referencia.getTime()) corte.referencia = fim;
  return corte;
}
function csCortes_(ss, hoje) {
  var painel = csLerPainel_(ss);
  var cortes = { estudos: {}, padrao: csCorte_(null, hoje), valores: painel.valores, lidos:
    painel.lidos };
  CS_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    cortes.estudos[cfg.estudo] = csCorte_(painel.valores[cfg.estudo], hoje);
  });
  return cortes;
}
function csCorteDe_(cortes, estudo) {
  return cortes.estudos.hasOwnProperty(estudo) ? cortes.estudos[estudo] : cortes.padrao;
}
function csDentroDoCorte_(corte, data0, coderest) {
  if (!corte.periodo) return true;
  if (corte.coderest !== null && (coderest === null || coderest > corte.coderest)) return false;
  if (!data0) return true;
  var t = csDia_(data0);
  if (corte.inicio && t < corte.inicio.getTime()) return false;
  if (corte.fim && corte.coderest === null && t > corte.fim.getTime()) return false;
  return true;
}
function csData0NoCorte_(corte, data0) {
  if (!data0 || !corte.periodo || !corte.fim) return data0;
  return csDia_(data0) > corte.fim.getTime() ? null : data0;
}
function csDataBr_(d) {
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}
function csTextoCorte_(corte) {
  if (!corte.periodo) return '';
  var partes = [];
  if (corte.coderest !== null) partes.push('coderest até ' + corte.coderest);
  if (corte.inicio && corte.fim) partes.push(csDataBr_(corte.inicio) + ' a ' + csDataBr_(corte.fim));
  else if (corte.inicio) partes.push('a partir de ' + csDataBr_(corte.inicio));
  else if (corte.fim) partes.push('até ' + csDataBr_(corte.fim));
  return partes.join('; ');
}
function csRotuloComCorte_(rotulo, corte) {
  var texto = csTextoCorte_(corte);
  return texto ? rotulo + ' (corte: ' + texto + ')' : rotulo;
}
function csAvisosCorte_(cortes) {
  var avisos = [];
  CS_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (c && c.aviso) avisos.push('corte do estudo ' + csNomeEstudo_(cfg.estudo) + ': ' + c.aviso);
  });
  return avisos;
}
function csDesfecho_(linha, idx, blocos, data0, corte, lidos) {
  if (data0 && !csData0NoCorte_(corte, data0)) {
    return { sucesso: csDesfechoVazio_(), sobrevida: csDesfechoVazio_() };
  }
  var usadas = linha.usadas;
  var doBanco = {};
  var d = {};
  [['sucesso', idx.dataSucesso, idx.tempoSucesso, idx.sucesso, idx.causaSucesso,
  CS_CODIGOS_FALHA_SUCESSO],
    ['sobrevida', idx.dataSobrevida, idx.tempoSobrevida, idx.sobrevida, idx.causaSobrevida,
  CS_CODIGOS_FALHA_SOBREVIDA]
  ].forEach(function(p) {
    linha.usadas = doBanco[p[0]] = {};
    d[p[0]] = csDesfechoBanco_(linha, p[1], p[2], p[3], p[4], p[5], data0);
  });
  linha.usadas = usadas;
  if (corte.periodo && corte.fim && data0) {
    var fim = corte.fim.getTime();
    var refeito = null;
    ['sucesso', 'sobrevida'].forEach(function(k) {
      if (!d[k].data || csDia_(d[k].data) <= fim) return;
      if (!refeito) refeito = csDesfechoBlocos_(lidos || csLerBlocos_(linha, blocos), blocos, data0,
      fim);
      d[k] = refeito[k];
      doBanco[k] = null;
      refeito.usados[k].forEach(function(i) { csUsar_(linha, [blocos[i].data, blocos[i].avaliacao]); });
      if (refeito.travas[k] >= 0) csUsar_(linha, [blocos[refeito.travas[k]].causa]);
    });
  }
  ['sucesso', 'sobrevida'].forEach(function(k) {
    if (!usadas || !doBanco[k]) return;
    for (var c in doBanco[k]) {
      if (doBanco[k].hasOwnProperty(c)) usadas[c] = true;
    }
  });
  return d;
}
function csDesfechoVazio_() {
  return { evento: false, codigo: null, causa: null, data: null, dias: null, anos: null };
}
function csDesfechoBanco_(linha, iData, iTempo, iCodigo, iCausa, falhas, data0) {
  var codigo = csNumero_(csValorBruto_(linha, iCodigo));
  var data = csData_(csValorBruto_(linha, iData));
  var dias = null;
  var anos = null;
  if (data0 && data) {
    var n = csDiasEntre_(data0, data);
    if (n >= 0) {
      dias = n;
      anos = n / CS_DIAS_ANO;
    }
  } else {
    var t = csNumero_(csValorBruto_(linha, iTempo));
    if (t !== null && t >= 0) {
      anos = t;
      if (data0) data = new Date(data0.getFullYear(), data0.getMonth(), data0.getDate() +
      Math.round(t * CS_DIAS_ANO));
    }
  }
  return { evento: codigo !== null && falhas.indexOf(codigo) >= 0, codigo: codigo,
    causa: csNumero_(csValorBruto_(linha, iCausa)), data: data, dias: dias, anos: anos };
}
function csDesfechoBlocos_(lidos, blocos, data0, fim) {
  var ordem = [];
  for (var k = 0; k < blocos.length; k++) ordem.push(k);
  ordem.sort(function(a, b) { return blocos[a].indice - blocos[b].indice || a - b; });
  var reg = { sucesso: null, sobrevida: null };
  var travado = { sucesso: false, sobrevida: false };
  var usados = { sucesso: [], sobrevida: [] };
  var travas = { sucesso: -1, sobrevida: -1 };
  var falhas = { sucesso: CS_CODIGOS_FALHA_SUCESSO, sobrevida:
    CS_CODIGOS_FALHA_SOBREVIDA };
  ordem.forEach(function(i) {
    var bl = lidos[i];
    if (!bl.data || csDia_(bl.data) > fim) return;
    if (blocos[i].indice !== 0 && bl.avaliacao === null) return;
    var r = { data: bl.data, codigo: bl.avaliacao, causa: bl.causa };
    ['sucesso', 'sobrevida'].forEach(function(k) {
      if (travado[k]) return;
      reg[k] = r;
      usados[k].push(i);
      if (r.codigo !== null && falhas[k].indexOf(r.codigo) >= 0) {
        travado[k] = true;
        travas[k] = i;
      }
    });
  });
  return {
    sucesso: csDesfechoRegistro_(reg.sucesso, travado.sucesso, data0),
    sobrevida: csDesfechoRegistro_(reg.sobrevida, travado.sobrevida, data0),
    usados: usados,
    travas: travas
  };
}
function csDesfechoRegistro_(r, travado, data0) {
  var data = r ? r.data : data0;
  var n = csDiasEntre_(data0, data);
  return { evento: travado, codigo: travado ? r.codigo : 0, causa: travado ? r.causa : 0, data: data,
    dias: n >= 0 ? n : null, anos: n >= 0 ? n / CS_DIAS_ANO : null };
}
function csMontar_(entrada, ss) {
  var idx = csIndices_(entrada.cabecalho, CS_CAMPOS);
  var blocos = csBlocos_(entrada.cabecalho);
  var faltando = csFaltando_(idx, CS_CAMPOS, [['grupo'], ['data0'], ['sucesso'], ['sobrevida']]);
  if (blocos.length < 2) faltando.push('Data 1 (blocos de rechamada)');
  var cortes = csCortes_(ss, new Date());
  if (csFaltaCoderest_(idx, cortes)) faltando.push('coderest (usado no corte)');
  var registros = csRegistros_(entrada.dados, idx, blocos, cortes);
  var estudos = csEstudos_(registros);
  var render = csNovoRender_();
  render.validacoes = [];
  var mapa = {};
  var guardados = {};
  var celulas = [];
  var partes = [];
  var r0 = 1 + CS_CAMPOS_PAINEL.length + 3;
  var linhas = null;
  estudos.forEach(function(est, k) {
    var cb = csBase_(k);
    var entradas = cortes.valores[est.estudo] || {};
    var corte = csCorteDe_(cortes, est.estudo);
    var painel = csPainel_(render, est, entradas, corte, cb);
    mapa[est.estudo] = painel.mapa;
    guardados[est.estudo] = painel.valores;
    celulas = celulas.concat(painel.celulas);
    var res = csFigura_(render, est, registros, entradas, r0, corte, cb);
    linhas = res.linhas;
    partes.push(csRotuloComCorte_('estudo ' + csNomeEstudo_(est.estudo) + ': ' +
    res.randomizadas + ' randomizadas, ' +
      res.semGrupo + ' ainda não randomizadas', corte));
  });
  if (linhas) {
    csEtapa_(render, linhas.inclusao, linhas.incluidos, 'Inclusão');
    csEtapa_(render, linhas.alocacao, linhas.alocacao, 'Alocação');
    csEtapa_(render, linhas.acompanhamento, linhas.acompanhamento, 'Acompanhamento');
    csEtapa_(render, linhas.analise, linhas.analise, 'Análise');
  }
  var nC = csBase_(estudos.length) - 2;
  for (var c = 1; c <= nC; c++) csLargura_(render, c, csLarguraCol_(c));
  if (render.nColunas < nC) render.nColunas = nC;
  render.blocos = { rotulos: [1, 2], estudos: estudos.map(function(est, k) {
    return { chave: est.estudo, nome: csNomeEstudo_(est.estudo), c1: csBase_(k), c2:
        Math.min(csBase_(k) + 4, nC) };
  }) };
  if (csPainelNoLugar_(ss, mapa)) render.preservar = celulas;
  render.chavePainel = cortes.lidos ? JSON.stringify(cortes.lidos) : '';
  render.faltando = faltando;
  render.resumo = partes.concat(csAvisosCorte_(cortes)).join('; ') + '.';
  render.congelarColunas = 1;
  render.depoisDeEscrever = function() {
    var props = csProps_();
    props.setProperty(CS_PAINEL_POS, JSON.stringify(mapa));
    props.setProperty(CS_PAINEL_VAL, JSON.stringify(guardados));
  };
  render.destaque = csDestaque_(registros.filter(function(r) {
    return estudos.some(function(est) { return est.estudo === r.estudo; });
  }));
  return csFechar_(render);
}
function csRegistros_(dados, idx, blocos, cortes) {
  var registros = [];
  var base = 0;
  for (var b = 0; b < blocos.length; b++) {
    if (blocos[b].indice === 0) {
      base = b;
      break;
    }
  }
  dados.forEach(function(linha, i) {
    linha.usadas = {};
    var codpac = csValorBruto_(linha, idx.codpac);
    var grupoBruto = csValorBruto_(linha, idx.grupo);
    var estudoBruto = csValorBruto_(linha, idx.estudo);
    var dente = csValorBruto_(linha, idx.dente);
    if (csVazio_(codpac) && csVazio_(grupoBruto) && csVazio_(dente) &&
    csVazio_(estudoBruto)) return;
    var grupo = csCodigoGrupo_(grupoBruto);
    var estudo = csEstudoEfetivo_(estudoBruto, grupo);
    if (!estudo) return;
    var data0Banco = csData_(csValorBruto_(linha, idx.data0));
    var corte = csCorteDe_(cortes, estudo);
    if (!csDentroDoCorte_(corte, data0Banco, csNumero_(csValorBruto_(linha, idx.coderest))))
    return;
    var data0 = csData0NoCorte_(corte, data0Banco);
    if (grupo && corte.fim && !data0) return;
    var identificacao = {};
    for (var c in linha.usadas) {
      if (linha.usadas.hasOwnProperty(c)) identificacao[c] = true;
    }
    var lidos = csLerBlocos_(linha, blocos);
    var rechamadas = [];
    lidos.forEach(function(bl, k) {
      if (k === base || !bl.data || !data0) return;
      if (corte.fim && csDia_(bl.data) > corte.fim.getTime()) return;
      rechamadas.push(bl.data);
      csUsar_(linha, [blocos[k].data]);
    });
    var avalBase = data0 && lidos.length ? lidos[base].avaliacao : null;
    if (!rechamadas.length && avalBase !== null) csUsar_(linha, [blocos[base].avaliacao]);
    var d = csDesfecho_(linha, idx, blocos, data0Banco, corte, lidos);
    registros.push({
      paciente: csChavePaciente_(codpac, i + CS_LINHA_DADOS),
      estudo: estudo,
      grupo: grupo,
      data0: data0,
      rechamadas: rechamadas,
      falhaNoDia: !rechamadas.length && ((avalBase !== null &&
      CS_CODIGOS_FALHA_SUCESSO.indexOf(avalBase) >= 0) ||
        (d.sucesso.evento && d.sucesso.anos === 0)),
      eSucesso: d.sucesso.evento,
      eSobrevida: d.sobrevida.evento,
      tSucesso: d.sucesso.anos,
      tSobrevida: d.sobrevida.anos,
      pacote: { linha: i + CS_LINHA_DADOS, usadas: grupo ? linha.usadas : identificacao }
    });
  });
  return registros;
}
function csSomarMeses_(d, meses) {
  var r = new Date(d.getFullYear(), d.getMonth() + meses, 1);
  var ultimoDia = new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate();
  r.setDate(Math.min(d.getDate(), ultimoDia));
  return r;
}
function csMesesCompletos_(a, b) {
  var m = (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
  if (b.getDate() < a.getDate()) m--;
  return m;
}
function csClasse_(r, hoje) {
  if (!r.data0) return 'semDatas';
  if (r.rechamadas.length) return 'comRechamada';
  if (r.falhaNoDia) return 'falhaNoDia';
  if (csSomarMeses_(r.data0, CS_MESES_RECHAMADA).getTime() <= hoje.getTime()) return 'atrasada';
  return 'aguardando';
}
function csResumoGrupo_(regs, hoje) {
  var r = { n: regs.length, pacientes: {}, semDatas: 0, comRechamada: 0, seis: 0, doze: 0,
    aguardando: [], atrasadas: [], falhaNoDia: 0, falhasSucesso: 0, falhasSobrevida: 0,
    comTempo: 0, tempoPositivo: 0 };
  regs.forEach(function(x) {
    r.pacientes[x.paciente] = true;
    if (x.eSucesso) r.falhasSucesso++;
    if (x.eSobrevida) r.falhasSobrevida++;
    if (x.tSucesso !== null) {
      r.comTempo++;
      if (x.tSucesso > 0 || x.eSucesso || (x.tSobrevida !== null && x.tSobrevida > 0))
      r.tempoPositivo++;
    }
    var classe = csClasse_(x, hoje);
    if (classe === 'semDatas') {
      r.semDatas++;
    } else if (classe === 'comRechamada') {
      r.comRechamada++;
      var ultima = x.rechamadas[0];
      x.rechamadas.forEach(function(d) { if (d.getTime() > ultima.getTime()) ultima = d; });
      var meses = csMesesCompletos_(x.data0, ultima);
      if (meses >= 6) r.seis++;
      if (meses >= 12) r.doze++;
    } else if (classe === 'falhaNoDia') {
      r.falhaNoDia++;
    } else if (classe === 'atrasada') {
      r.atrasadas.push(x.data0);
    } else {
      r.aguardando.push(x.data0);
    }
  });
  r.nPacientes = csContar_(r.pacientes);
  return r;
}
function csChaveMes_(d) {
  return d.getFullYear() * 12 + d.getMonth();
}
function csTextoMes_(k, comAno) {
  var texto = CS_MESES[k % 12];
  return comAno ? texto + '/' + Math.floor(k / 12) : texto;
}
function csFaixaMeses_(datas) {
  var chaves = datas.map(csChaveMes_);
  var min = Math.min.apply(null, chaves);
  var max = Math.max.apply(null, chaves);
  if (min === max) return csTextoMes_(min, true);
  if (Math.floor(min / 12) === Math.floor(max / 12)) {
    return csTextoMes_(min, false) + '–' + csTextoMes_(max, true);
  }
  return csTextoMes_(min, true) + '–' + csTextoMes_(max, true);
}
function csListaMeses_(datas) {
  var chaves = [];
  datas.forEach(function(d) {
    var k = csChaveMes_(d);
    if (chaves.indexOf(k) < 0) chaves.push(k);
  });
  chaves.sort(function(a, b) { return a - b; });
  if (chaves.length > 3) return csFaixaMeses_(datas);
  return chaves.map(function(k) { return csTextoMes_(k, true); }).join(', ');
}
function csValorEntrada_(v) {
  var t = String(v === null || v === undefined ? '' : v).trim();
  return t ? t : CS_VAZIO_FIGURA;
}
function csSomaEntradas_(entradas, chaves) {
  var soma = 0;
  var algum = false;
  chaves.forEach(function(k) {
    var n = csNumero_(entradas[k]);
    if (n !== null) {
      soma += n;
      algum = true;
    }
  });
  return algum ? String(soma) : CS_VAZIO_FIGURA;
}
function csPainelNoLugar_(ss, mapa) {
  if (!ss || !ss.getSheetByName(CS_ABA_SAIDA)) return false;
  var anterior = null;
  try {
    anterior = JSON.parse(csProps_().getProperty(CS_PAINEL_POS) || 'null');
  } catch (err) {
    anterior = null;
  }
  return !!anterior && JSON.stringify(anterior) === JSON.stringify(mapa);
}
function csPainel_(render, est, entradas, corte, cb) {
  var mapa = {};
  var valores = {};
  var celulas = [];
  var nCorte = 0;
  csPor_(render, 1, cb, CS_TITULO_PAINEL.replace('{estudo}', csNomeEstudo_(est.estudo)));
  csMarcar_(render, 'negrito', 1, cb, 1, 1, true);
  CS_CAMPOS_PAINEL.forEach(function(en, k) {
    var l = 2 + k;
    var c = cb + 2;
    var bruto = entradas[en.chave];
    var texto = bruto === null || bruto === undefined ? '' : String(bruto).trim();
    var celula = texto;
    if (en.tipo === 'modo') {
      texto = csModoCorte_(texto);
      celula = texto;
    } else if (en.tipo === 'data') {
      var d = csDataEstrita_(texto);
      if (d) {
        texto = csDataBr_(d);
        celula = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' +
        d.getDate()).slice(-2);
      }
    } else if (en.tipo === 'numero') {
      var n = csNumero_(texto);
      if (n !== null) {
        celula = n;
        texto = String(n);
      }
    }
    if (en.tipo) nCorte++;
    csPor_(render, l, cb, en.rotulo);
    csPor_(render, l, c, celula);
    csMarcar_(render, 'formato', l, c, 1, 1, en.tipo === 'data' ? CS_FORMATO_DATA : (en.tipo
    === 'numero' ? '0' : '@'));
    csMarcar_(render, 'fundo', l, c, 1, 1, CS_COR_ENTRADA);
    csMarcar_(render, 'alinhaH', l, c, 1, 1, 'center');
    csCaixa_(render, l, c, 1, 1);
    if (en.tipo) render.validacoes.push({ l: l, c: c, regra: csRegra_(en) });
    mapa[en.chave] = csA1_(l, c);
    celulas.push({ l: l, c: c });
    valores[en.chave] = texto;
  });
  if (nCorte) {
    var status = csStatusCorte_(corte);
    var cs = cb + 4;
    csPor_(render, 2, cs, status.texto);
    if (nCorte > 1) csMarcar_(render, 'mesclar', 2, cs, nCorte, 1, true);
    csMarcar_(render, 'quebra', 2, cs, nCorte, 1, true);
    csMarcar_(render, 'alinhaV', 2, cs, nCorte, 1, 'top');
    csMarcar_(render, 'tamanho', 2, cs, nCorte, 1, 9);
    csMarcar_(render, status.aviso ? 'negrito' : 'italico', 2, cs, nCorte, 1, true);
    csMarcar_(render, 'corTexto', 2, cs, nCorte, 1, status.aviso ? CS_COR_AVISO :
    CS_COR_STATUS);
  }
  return { mapa: mapa, valores: valores, celulas: celulas };
}
function csRegra_(en) {
  var b = SpreadsheetApp.newDataValidation();
  if (en.tipo === 'modo') b = b.requireValueInList([CS_MODO_AUTO, CS_MODO_PERIODO],
  true);
  else if (en.tipo === 'numero') b = b.requireNumberGreaterThanOrEqualTo(1);
  else b = b.requireDate();
  return b.setAllowInvalid(false).setHelpText(en.ajuda).build();
}
function csStatusCorte_(corte) {
  if (corte.aviso) return { texto: 'ATENÇÃO: ' + corte.aviso, aviso: true };
  if (!corte.periodo) {
    return { texto: 'Auto: entram todas as restaurações e todo o acompanhamento até hoje; atualiza sozinho.', aviso: false };
  }
  var partes = [];
  if (corte.coderest !== null) partes.push('coderest até ' + corte.coderest);
  if (corte.inicio && corte.fim) {
    partes.push('cimentadas de ' + csDataBr_(corte.inicio) + ' a ' + csDataBr_(corte.fim));
  } else if (corte.inicio) {
    partes.push('cimentadas a partir de ' + csDataBr_(corte.inicio));
  } else if (corte.fim) {
    partes.push('cimentadas até ' + csDataBr_(corte.fim));
  }
  partes.push('acompanhamento até ' + (corte.fim ? csDataBr_(corte.fim) : 'hoje'));
  return { texto: 'Em vigor: ' + partes.join('; ') + '.', aviso: false };
}
function csBase_(k) {
  return 3 + CS_LARGURAS_FIGURA.length * k;
}
function csLarguraCol_(c) {
  if (c === 1) return CS_LARGURA_FAIXA;
  if (c === 2) return CS_LARGURA_RECUO;
  return CS_LARGURAS_FIGURA[(c - 3) % CS_LARGURAS_FIGURA.length];
}
function csLarguraSoma_(c, nc) {
  var total = 0;
  for (var k = c; k < c + nc; k++) total += csLarguraCol_(k);
  return total;
}
function csCaixaTexto_(render, l, c, nc, linhas, op) {
  var texto = linhas.join('\n');
  csPor_(render, l, c, texto);
  if (nc > 1) csMarcar_(render, 'mesclar', l, c, 1, nc, true);
  csMarcar_(render, 'quebra', l, c, 1, nc, true);
  csMarcar_(render, 'alinhaH', l, c, 1, nc, op.centro ? 'center' : 'left');
  if (op.topo) csMarcar_(render, 'alinhaV', l, c, 1, nc, 'top');
  csCaixa_(render, l, c, 1, nc, op.tracejada ? 'tracejada' : 'solida');
  if (op.negritoPrimeira) {
    csMarcar_(render, 'rico', l, c, 1, 1, { texto: texto, negritos: [[0, linhas[0].length]] });
  }
  var nLinhas = csContarLinhas_(texto, csLarguraSoma_(c, nc), 6.2);
  csAltura_(render, l, 12 + 16 * nLinhas);
}
function csSeta_(render, l, c, nc, simbolo) {
  csPor_(render, l, c, simbolo);
  if (nc > 1) csMarcar_(render, 'mesclar', l, c, 1, nc, true);
  csMarcar_(render, 'alinhaH', l, c, 1, nc, 'center');
  csMarcar_(render, 'tamanho', l, c, 1, nc, 12);
}
function csEtapa_(render, l1, l2, texto) {
  var nl = l2 - l1 + 1;
  csPor_(render, l1, 1, texto);
  if (nl > 1) csMarcar_(render, 'mesclar', l1, 1, nl, 1, true);
  csMarcar_(render, 'fundo', l1, 1, nl, 1, CS_COR_ETAPA);
  csMarcar_(render, 'rotacao', l1, 1, nl, 1, 90);
  csMarcar_(render, 'negrito', l1, 1, nl, 1, true);
  csMarcar_(render, 'alinhaH', l1, 1, nl, 1, 'center');
}
function csFigura_(render, est, registros, entradas, r0, corte, cb) {
  var hoje = corte.referencia;
  var regs = registros.filter(function(r) { return r.estudo === est.estudo; });
  var comGrupo = regs.filter(function(r) { return !!r.grupo; });
  var semGrupo = regs.length - comGrupo.length;
  var pacientes = {};
  comGrupo.forEach(function(r) { pacientes[r.paciente] = true; });
  var nPac = csContar_(pacientes);
  var textoCorte = csTextoCorte_(corte);
  var titulo = CS_ROTULO_FIGURA + ' ' + CS_TITULO.replace('{estudo}',
  csNomeEstudo_(est.estudo))
    .replace('{corte}', textoCorte ? ' (corte: ' + textoCorte + ')' : '');
  csPor_(render, r0, cb, titulo);
  csMarcar_(render, 'tamanho', r0, cb, 1, 1, 11);
  csMarcar_(render, 'rico', r0, cb, 1, 1, { texto: titulo, negritos: [[0, CS_ROTULO_FIGURA.length]]
});
  csAltura_(render, r0 + 1, 8);
  var lE = r0 + 2;
  csCaixaTexto_(render, lE, cb, 3, ['Avaliados para elegibilidade',
    '(n = ' + csValorEntrada_(entradas.avaliados) + ' pacientes)'],
    { tracejada: true, centro: true, negritoPrimeira: true });
  csSeta_(render, lE, cb + 3, 1, '►');
  csCaixaTexto_(render, lE, cb + 4, 1, [
    'Excluídos (n = ' + csSomaEntradas_(entradas, ['naoAtenderam', 'recusaram', 'outros']) + ')',
    '• Não atenderam aos critérios de inclusão (n = ' + csValorEntrada_(entradas.naoAtenderam)
  + ')',
    '• Recusaram-se a participar (n = ' + csValorEntrada_(entradas.recusaram) + ')',
    '• Outros motivos (n = ' + csValorEntrada_(entradas.outros) + ')'],
    { tracejada: true, negritoPrimeira: true });
  csSeta_(render, lE + 1, cb, 3, '▼');
  csAltura_(render, lE + 1, 22);
  var lI = lE + 2;
  csCaixaTexto_(render, lI, cb, 3, ['Incluídos e randomizados',
    comGrupo.length + ' ' + csPlural_(comGrupo.length, 'restauração', 'restaurações') + ' em ' +
  nPac + ' ' +
    csPlural_(nPac, 'paciente', 'pacientes'),
    '(unidade de randomização: restauração)'], { centro: true, negritoPrimeira: true });
  csSeta_(render, lI, cb + 3, 1, '►');
  csCaixaTexto_(render, lI, cb + 4, 1, ['Ainda não randomizados',
    'n = ' + semGrupo + ' ' + csPlural_(semGrupo, 'restauração, excluída da análise',
      'restaurações, excluídas da análise')], { negritoPrimeira: true });
  var lA = lI + 2;
  var lF = lA + 2;
  var lN = lF + 2;
  [lI + 1, lA + 1, lF + 1].forEach(function(l) { csAltura_(render, l, 22); });
  est.grupos.forEach(function(g, i) {
    var col = cb + 2 * i;
    var rg = comGrupo.filter(function(r) { return r.grupo === g; });
    var s = csResumoGrupo_(rg, hoje);
    var naoCimentadas = csPlural_(s.semDatas, 'não cimentada', 'não cimentadas');
    [lI + 1, lA + 1, lF + 1].forEach(function(l) { csSeta_(render, l, col, 1, '▼'); });
    csCaixaTexto_(render, lA, col, 1, [
      CS_TITULOS_GRUPO.hasOwnProperty(g) ? CS_TITULOS_GRUPO[g] :
    csRotuloGrupo_(g),
      'Alocadas: n = ' + s.n + ' ' + csPlural_(s.n, 'restauração', 'restaurações') + ' (' + s.nPacientes
    + ' ' +
        csPlural_(s.nPacientes, 'paciente', 'pacientes') + ')',
      '• Receberam a intervenção alocada (n = ' + (s.n - s.semDatas) + ')',
      '• Não receberam a intervenção alocada (n = ' + s.semDatas + ')' + (s.semDatas ? ': sem cimentação registrada' : '')],
      { negritoPrimeira: true, topo: true });
    var acomp = ['Com ≥ 1 rechamada (n = ' + s.comRechamada + ')',
      '• ≥ 6 meses (n = ' + s.seis + ')',
      '• ≥ 12 meses (n = ' + s.doze + ')'];
    if (s.aguardando.length) {
      acomp.push('Aguardando a 1ª rechamada, ' + csPlural_(s.aguardando.length, 'cimentada',
      'cimentadas') +
        ' em ' + csFaixaMeses_(s.aguardando) + ' (n = ' + s.aguardando.length + ')');
    }
    if (s.falhaNoDia) acomp.push('Falha no dia da cimentação (n = ' + s.falhaNoDia + ')');
    if (s.atrasadas.length) {
      acomp.push('Sem rechamada desde a cimentação, ' + csListaMeses_(s.atrasadas) +
        ' (n = ' + s.atrasadas.length + ')');
    }
    if (s.semDatas) {
      acomp.push(csPlural_(s.semDatas, 'Não cimentada', 'Não cimentadas') + ', sem acompanhamento (n = ' + s.semDatas + ')');
    }
    if (!s.atrasadas.length && !s.semDatas) acomp.push('Perdas de seguimento (n = 0)');
    acomp.push('Falhas: sucesso n = ' + s.falhasSucesso + '; sobrevida n = ' +
    s.falhasSobrevida);
    csCaixaTexto_(render, lF, col, 1, acomp, { topo: true });
    csCaixaTexto_(render, lN, col, 1, [
      'Analisadas (n = ' + (s.n - s.semDatas) + ')',
      '• Características (Tabela 1): n = ' + (s.n - s.semDatas),
      '• Sucesso e sobrevida (Kaplan-Meier): n = ' + s.comTempo + ' com tempo registrado; ' +
        s.tempoPositivo + ' com tempo > 0',
      '• ' + csPlural_(s.semDatas, 'Excluída', 'Excluídas') + ' da análise (n = ' + s.semDatas + ')' +
        (s.semDatas ? ': ' + naoCimentadas : '')],
      { negritoPrimeira: true, topo: true });
  });
  var fim = lN;
  var nota = csNotaPacientes_(est, comGrupo.filter(function(r) { return !!r.data0; }));
  if (nota) {
    fim = lN + 2;
    csAltura_(render, lN + 1, 8);
    csPor_(render, fim, cb, nota);
    csMarcar_(render, 'mesclar', fim, cb, 1, 5, true);
    csMarcar_(render, 'quebra', fim, cb, 1, 5, true);
    csMarcar_(render, 'tamanho', fim, cb, 1, 5, 9);
    csAltura_(render, fim, 8 + 13 * csContarLinhas_(nota, csLarguraSoma_(cb, 5), 5.6));
  }
  return { fim: fim, randomizadas: comGrupo.length, semGrupo: semGrupo,
    linhas: { inclusao: lE, incluidos: lI, alocacao: lA, acompanhamento: lF, analise: lN } };
}
function csNotaPacientes_(est, comGrupo) {
  var porPaciente = {};
  comGrupo.forEach(function(r) {
    if (!porPaciente[r.paciente]) porPaciente[r.paciente] = {};
    porPaciente[r.paciente][r.grupo] = (porPaciente[r.paciente][r.grupo] || 0) + 1;
  });
  var multiplos = 0;
  var todosUm = true;
  for (var p in porPaciente) {
    if (!porPaciente.hasOwnProperty(p) || p.indexOf('#linha') === 0) continue;
    var grupos = [];
    for (var g in porPaciente[p]) {
      if (porPaciente[p].hasOwnProperty(g)) grupos.push(g);
    }
    if (grupos.length < 2) continue;
    multiplos++;
    grupos.forEach(function(gg) { if (porPaciente[p][gg] !== 1) todosUm = false; });
  }
  if (!multiplos) return '';
  var somas = est.grupos.map(function(g) {
    var vistos = {};
    comGrupo.forEach(function(r) { if (r.grupo === g) vistos[r.paciente] = true; });
    return csContar_(vistos);
  });
  var quem = multiplos < CS_EXTENSO.length ? CS_EXTENSO[multiplos] : String(multiplos);
  return quem + ' ' + csPlural_(multiplos, 'paciente recebeu', 'pacientes receberam') + ' ' +
    (todosUm ? 'uma restauração em cada grupo' : 'restaurações em mais de um grupo') +
    '; por isso, a soma de pacientes por grupo (' + somas.join(' + ') +
    ') é maior que o total de pacientes analisados (' + csContar_(porPaciente) + ').';
}


/* ======================================================================
 * MÓDULO 3 DE 5 — TABELA1.gs
 * ====================================================================== */

var T1_ABA_SAIDA = 'TABELA 3';
var T1_VERSAO = 'public-v1';
var T1_REAGIR_PROPRIA_ABA = false;
var T1_ROTULO_TABELA = 'Tabela 3 –';
var T1_TITULO = 'Distribuição da amostra por características do paciente e da restauração, segundo o estudo e o grupo.';
var T1_CALCULAR_P = true;
var T1_ROTULO_VALOR_P = 'Valor p';
var T1_ROTULO_QUI = 'Qui-quadrado';
var T1_ROTULO_FISHER = 'Exato de Fisher';
var T1_ROTULO_VARIAVEL = 'Variável';
var T1_ROTULOS_RISCO = null;
var T1_FISHER_MAX_TABELAS = 2000000;
var T1_SIMULACOES = 20000;
var T1_LARGURA_ROTULO = 280;
var T1_LARGURA_N = 86;
var T1_LARGURA_PCT = 52;
var T1_LARGURA_P = 110;
var T1_LARGURA_ESPACO = 24;
var T1_SALARIOS_MINIMOS = [
  { desde: '2010-01-01', valor: 510 },
  { desde: '2011-01-01', valor: 540 },
  { desde: '2011-03-01', valor: 545 },
  { desde: '2012-01-01', valor: 622 },
  { desde: '2013-01-01', valor: 678 },
  { desde: '2014-01-01', valor: 724 },
  { desde: '2015-01-01', valor: 788 },
  { desde: '2016-01-01', valor: 880 },
  { desde: '2017-01-01', valor: 937 },
  { desde: '2018-01-01', valor: 954 },
  { desde: '2019-01-01', valor: 998 },
  { desde: '2020-01-01', valor: 1039 },
  { desde: '2020-02-01', valor: 1045 },
  { desde: '2021-01-01', valor: 1100 },
  { desde: '2022-01-01', valor: 1212 },
  { desde: '2023-01-01', valor: 1302 },
  { desde: '2023-05-01', valor: 1320 },
  { desde: '2024-01-01', valor: 1412 },
  { desde: '2025-01-01', valor: 1518 },
  { desde: '2026-01-01', valor: 1621 }
];
var T1_FAIXAS_IDADE = [
  { rotulo: '< 18', max: 17 },
  { rotulo: '18 – 30', min: 18, max: 30 },
  { rotulo: '31 – 45', min: 31, max: 45 },
  { rotulo: '46 – 64', min: 46, max: 64 },
  { rotulo: '≥ 65', min: 65 }
];
var T1_FAIXAS_RENDA = [
  { rotulo: '≤ 1 salário mínimo', max: 1 },
  { rotulo: '> 1 a 3 salários mínimos', acimaDe: 1, max: 3 },
  { rotulo: '> 3 salários mínimos', acimaDe: 3 }
];
var T1_FAIXAS_DENTES = [
  { rotulo: '1', min: 1, max: 1 },
  { rotulo: '2', min: 2, max: 2 },
  { rotulo: '≥ 3', min: 3 }
];
var T1_ESCOLARIDADE = [
  { rotulo: 'Até médio incompleto', codigos: ['0', '1', '3'],
    textos: ['Fundamental incompleto', 'Médio incompleto', 'Fundamental completo'] },
  { rotulo: 'Médio completo / superior incompleto', codigos: ['2', '4'],
    textos: ['Superior incompleto', 'Médio completo'] },
  { rotulo: 'Superior completo', codigos: ['5'], textos: ['Superior completo'] }
];
var T1_CAMPOS = {
  codpac: ['codepac', 'codpac'],
  coderest: ['coderest'],
  estudo: ['estudo'],
  grupo: ['grupo'],
  dente: ['n⁰ dente', 'Nº dente', 'N° dente'],
  idade: ['idade'],
  nascimento: ['Data nasc.', 'Data de nascimento'],
  consulta: ['Data 1ª consulta'],
  data0: ['Data 0'],
  sexo: ['Sexo'],
  escolaridade: ['Grau escolaridade', 'Grau de escolaridade', 'Escolaridade'],
  renda: ['Renda familiar', 'Renda familiar mensal'],
  conjunto: ['Conjuto', 'Conjunto'],
  arcada: ['Arcada'],
  faces: ['Nº faces restantes', 'Nº de faces restantes'],
  vitalidade: ['Vitalidade'],
  riscoOclusal: ['Nível estresse oclusal', 'Risco oclusal'],
  riscoCarie: ['riscocárie', 'Risco de cárie']
};
var T1_SECOES = [
  { chave: 'paciente', rotulo: 'Variáveis relacionadas ao paciente' },
  { chave: 'restauracao', rotulo: 'Variáveis relacionadas à restauração' }
];
var T1_VARIAVEIS = [
  { rotulo: 'Idade (em anos)', campo: 'idade', secao: 'paciente', faixas: T1_FAIXAS_IDADE },
  { rotulo: 'Sexo', campo: 'sexo', secao: 'paciente', fixas: ['Feminino', 'Masculino'] },
  { rotulo: 'Escolaridade', campo: 'escolaridade', secao: 'paciente' },
  { rotulo: 'Renda familiar mensal', campo: 'renda', secao: 'paciente', faixas: T1_FAIXAS_RENDA
},
  { rotulo: 'Nº de dentes, por paciente', campo: 'dentesPaciente', secao: 'paciente', faixas:
  T1_FAIXAS_DENTES },
  { rotulo: 'Tipo de dente', campo: 'conjunto', secao: 'restauracao', fixas: ['Molar', 'Pré-molar'] },
  { rotulo: 'Arcada', campo: 'arcada', secao: 'restauracao', fixas: ['Superior', 'Inferior'] },
  { rotulo: 'Nº de faces restantes', campo: 'faces', secao: 'restauracao', fixas: ['0', '1', '2', '3', '4'] },
  { rotulo: 'Vitalidade', campo: 'vitalidade', secao: 'restauracao', fixas: ['Vital', 'Não vital'] },
  { rotulo: 'Ano de colocação', campo: 'ano', secao: 'restauracao' },
  { rotulo: 'Risco oclusal', campo: 'riscoOclusal', secao: 'restauracao' },
  { rotulo: 'Risco de cárie', campo: 'riscoCarie', secao: 'restauracao' }
];
var T1_ABA_CONFERENCIA = 'CONFERÊNCIA TABELA 1';
var T1_COLUNAS_CONFERENCIA = [
  { rotulo: 'Estudo', largura: 90 },
  { rotulo: 'Grupo', largura: 150 },
  { rotulo: 'codepac', largura: 70 },
  { rotulo: 'coderest', largura: 70, formato: '0' },
  { rotulo: 'Nº dente', largura: 70, formato: '@' },
  { rotulo: 'Linha no banco', largura: 70, formato: '0' },
  { rotulo: 'Data 0', largura: 90, formato: 'dd/mm/yyyy' },
  { rotulo: 'Data do salário mínimo', largura: 110, formato: 'dd/mm/yyyy' },
  { rotulo: 'Renda no banco', largura: 120, formato: 'R$ #,##0.00' },
  { rotulo: 'Salário mínimo usado', largura: 110, formato: 'R$ #,##0.00' },
  { rotulo: 'Renda em salários mínimos', largura: 110, formato: '0.000' },
  { rotulo: 'Faixa de renda (Tabela 1)', largura: 190 },
  { rotulo: 'Nº de dentes do paciente no estudo', largura: 120, formato: '0' },
  { rotulo: 'Faixa de nº de dentes (Tabela 1)', largura: 120 },
  { rotulo: 'Observação', largura: 360 }
];
var T1_ABAS_FONTE = ['BANCO DE ANÁLISE', 'BANCO DE ANALISE', 'Folha1', 'Planilha1',
'Sheet1'];
var T1_LINHA_CABECALHO = 1;
var T1_LINHA_DADOS = 2;
var T1_HORA_ROTINA = null;
var T1_MINUTOS_VIGIA = 30;
var T1_ATRASO_SEGUNDOS = 90;
var T1_JANELA_IGNORAR_MS = 45000;
var T1_LINHAS_ENTRE_TABELAS = 3;
var T1_MARCADORES_VAZIOS = ['.', '..', 'na', 'n/a', 'n.a.', 'nd', 'n/d', '-', '--', 's/i', '?',
  'sem informacao', 'sem informação', 'null', 'undefined'];
var T1_ROTULO_SEM_INFO = 'Sem informação';
var T1_ROTULO_TOTAL = 'Total';
var T1_ROTULO_SEM_GRUPO = 'grupo não informado\nsem sorteio';
var T1_ROTULO_LINHA_ESTUDO = 'estudo';
var T1_ROTULO_LINHA_GRUPO = 'grupo';
var T1_GRUPOS_POR_ESTUDO = [
  { estudo: '0', grupos: ['0', '1'] },
  { estudo: '1', grupos: ['2', '3'] }
];
var T1_ESTUDO_POR_TEXTO = { 'estudoa': '0', 'estudob': '1', 'studya': '0', 'studyb': '1', 'a': '0', 'b': '1' };
var T1_ROTULOS = {
  estudo: { '0': 'Estudo A', '1': 'Estudo B' },
  grupo: {
    '0': 'Grupo 1',
    '1': 'Grupo 2',
    '2': 'Grupo 3',
    '3': 'Grupo 4'
  }
};
var T1_NOME_ESTUDO = { '0': 'A', '1': 'B' };
var T1_COR_SECAO = '#efefef';
var T1_COR_VALOR_P = '#fff2cc';
var T1_CORES_DESTAQUE = [
  { chave: '0', nome: 'A', linha: '#cfe2f3', dados: '#6fa8dc' },
  { chave: '1', nome: 'B', linha: '#fff2cc', dados: '#ffe599' },
  { chave: 'sem', nome: 'sem sorteio', linha: '#d9d2e9', dados: '#b4a7d6' }
];
var T1_ACENTOS = {
  'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
  'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
  'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
  'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
  'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
  'ç': 'c', 'ñ': 'n', 'ý': 'y', 'ÿ': 'y'
};
var T1_PAINEL_ABA = 'CONSORT';
var T1_PAINEL_POS = 'CS_ENTRADAS_POS';
var T1_PAINEL_VAL = 'CS_ENTRADAS_VAL';
var T1_MODO_AUTO = 'auto';
var T1_MODO_PERIODO = 'por período';
var T1_ESPERA_TRAVA_PAINEL_MS = 120000;
var T1_VANTAGEM_PAINEL_MS = 2000;
function GERAR_TABELA1() {
  return t1Executar_(true);
}
function tabela1Rotina() {
  return t1Executar_(false);
}
function tabela1Reconstruir() {
  t1LimparAgendamento_();
  return t1Executar_(false);
}
function tabela1Vigia() {
  var atual = t1Impressao_();
  var visto = t1Props_().getProperty('T1_IMPRESSAO') || '';
  if (atual && atual === visto && !cmPrecisaEscrever_(T1_ABA_SAIDA)) return;
  t1Executar_(false);
}
function tabela1OnChange(e) {
  if (t1PainelMudou_()) return t1AtualizarPeloPainel_();
  if (t1OrigemMudanca_(e) !== 'fonte') return;
  if (t1GeracaoRecente_()) return;
  t1Agendar_();
}
function configurarTabela1(silencioso) {
  var ss = t1SS_();
  desligarTabela1(true);
  t1Props_().deleteProperty('T1_DESLIGADO');
  var hora = t1HoraRotina_();
  var falhas = [];
  var pedidos = [
    { rotulo: 'reconstrução diária', criar: function() {
      ScriptApp.newTrigger('tabela1Rotina').timeBased().everyDays(1).atHour(hora).create();
    } },
    { rotulo: 'reação a edições', criar: function() {
      ScriptApp.newTrigger('tabela1OnChange').forSpreadsheet(ss).onChange().create();
    } },
    { rotulo: 'vigia', criar: function() {
      ScriptApp.newTrigger('tabela1Vigia').timeBased().everyMinutes(T1_MINUTOS_VIGIA).create();
    } }
  ];
  pedidos.forEach(function(p) {
    try {
      p.criar();
    } catch (err) {
      falhas.push(p.rotulo + ' (' + err.message + ')');
    }
  });
  var msg = 'Gatilhos da aba ' + T1_ABA_SAIDA + ' instalados: vigia a cada ' +
  T1_MINUTOS_VIGIA +
    ' min, reação a edições no banco e no painel do CONSORT e reconstrução diária às ' + hora
  + 'h.';
  if (falhas.length) {
    msg = 'ATENÇÃO: não foi possível instalar ' + falhas.join('; ') + '. ' +
      'O projeto aceita no máximo 20 gatilhos por conta; confira em Acionadores. ' + msg;
  }
  if (silencioso !== true) t1Avisar_(msg);
  return msg;
}
function desligarTabela1(interno) {
  var alvos = ['tabela1Rotina', 'tabela1OnChange', 'tabela1Reconstruir', 'tabela1Vigia'];
  var removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (alvos.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  t1Props_().deleteProperty('T1_AGENDADO');
  if (interno !== true) t1Props_().setProperty('T1_DESLIGADO', '1');
  return removidos + ' gatilho(s) da aba ' + T1_ABA_SAIDA + ' removido(s).';
}
function t1GruposEstudos_(sh, blocos) {
  if (!sh || !blocos || !blocos.estudos) return;
  var nC = sh.getMaxColumns();
  var quer = [0];
  for (var c = 1; c <= nC; c++) quer.push(0);
  blocos.estudos.forEach(function(b) {
    for (var k = b.c1; k <= Math.min(b.c2, nC); k++) quer[k] = 1;
  });
  var atual = [0];
  var igual = true;
  for (var c2 = 1; c2 <= nC; c2++) {
    atual.push(sh.getColumnGroupDepth(c2));
    if (atual[c2] !== quer[c2]) igual = false;
  }
  if (igual) return;
  var maximo = Math.max.apply(null, atual);
  for (var d = maximo; d >= 1; d--) {
    t1Faixas_(nC, function(k) { return atual[k] >= d; }).forEach(function(f) {
      sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(-1);
    });
  }
  sh.showColumns(1, nC);
  t1Faixas_(nC, function(k) { return quer[k] === 1; }).forEach(function(f) {
    sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(1);
  });
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
}
function t1Faixas_(n, teste) {
  var faixas = [];
  var ini = 0;
  for (var c = 1; c <= n + 1; c++) {
    var ok = c <= n && teste(c);
    if (ok && !ini) ini = c;
    if (!ok && ini) {
      faixas.push([ini, c - 1]);
      ini = 0;
    }
  }
  return faixas;
}
function t1Destaque_(registros) {
  var mapa = {};
  registros.forEach(function(r) {
    if (!r.pacote) return;
    var chave = r.grupo ? r.estudo : 'sem';
    if (!t1CorDestaque_(chave)) return;
    if (!mapa[chave]) mapa[chave] = {};
    if (!mapa[chave][r.pacote.linha]) mapa[chave][r.pacote.linha] = {};
    var celulas = mapa[chave][r.pacote.linha];
    for (var c in r.pacote.usadas) {
      if (r.pacote.usadas.hasOwnProperty(c)) celulas[c] = true;
    }
  });
  return mapa;
}
function t1CorDestaque_(chave) {
  for (var i = 0; i < T1_CORES_DESTAQUE.length; i++) {
    if (T1_CORES_DESTAQUE[i].chave === chave) return T1_CORES_DESTAQUE[i];
  }
  return null;
}
function t1RegrasDestaque_(mapa, nC) {
  var dados = [];
  var linhas = [];
  T1_CORES_DESTAQUE.forEach(function(cor) {
    var porLinha = mapa[cor.chave];
    var numeros = porLinha ? Object.keys(porLinha).map(Number) : [];
    if (!numeros.length) return;
    var ultima = Math.max.apply(null, numeros);
    var porTrecho = {};
    numeros.forEach(function(l) {
      t1Faixas_(nC, function(c) { return !!porLinha[l][c - 1]; }).forEach(function(t) {
        var k = t[0] + ':' + t[1];
        if (!porTrecho[k]) porTrecho[k] = {};
        porTrecho[k][l] = true;
      });
    });
    var a1Dados = [];
    Object.keys(porTrecho).forEach(function(k) {
      var c = k.split(':').map(Number);
      t1Faixas_(ultima, function(l) { return !!porTrecho[k][l]; }).forEach(function(t) {
        a1Dados.push(t1A1_(t[0], c[0], t[1] - t[0] + 1, c[1] - c[0] + 1));
      });
    });
    var a1Linhas = t1Faixas_(ultima, function(l) { return !!porLinha[l]; }).map(function(t) {
      return t1A1_(t[0], 1, t[1] - t[0] + 1, nC);
    });
    if (a1Dados.length) dados.push({ formula: t1FormulaDestaque_(cor.nome, 'dados lidos'), cor:
      cor.dados, a1: a1Dados.sort() });
    linhas.push({ formula: t1FormulaDestaque_(cor.nome, 'linha analisada'), cor: cor.linha, a1:
      a1Linhas.sort() });
  });
  return { dados: dados, linhas: linhas };
}
function t1FormulaDestaque_(categoria, tipo) {
  return '=N("RRT · ' + T1_ABA_SAIDA + ' · ' + categoria + ' · ' + tipo + '")=0';
}
function t1MarcaDestaque_(regra) {
  try {
    var cond = regra.getBooleanCondition();
    if (!cond) return null;
    var m = String((cond.getCriteriaValues() || [])[0] || '').replace(/\s+/g, ' ')
      .match(/^\s*=\s*N\(\s*"RRT\s*·\s*(.*?)\s*·\s*(.*?)\s*·\s*(dados lidos|linha analisada)"\s*\)\s*=\s*0\s*$/i);
    return m ? { dono: m[1].trim().toUpperCase(), dados: m[3].toLowerCase() === 'dados lidos' } :
    null;
  } catch (err) {
    return null;
  }
}
function t1AplicarDestaque_(sh, mapa) {
  if (!sh || !mapa) return '';
  try {
    var desejo = t1RegrasDestaque_(mapa, sh.getLastColumn());
    var outras = [];
    var dados = [];
    var linhas = [];
    var minhas = [];
    var fase = 0;
    var emOrdem = true;
    sh.getConditionalFormatRules().forEach(function(regra) {
      var marca = t1MarcaDestaque_(regra);
      var f = !marca ? 0 : (marca.dados ? 1 : 2);
      if (f < fase) emOrdem = false;
      fase = Math.max(fase, f);
      if (!marca) outras.push(regra);
      else if (marca.dono === T1_ABA_SAIDA.toUpperCase()) minhas.push(regra);
      else if (marca.dados) dados.push(regra);
      else linhas.push(regra);
    });
    if (emOrdem && t1MesmasRegras_(minhas, desejo.dados.concat(desejo.linhas))) return '';
    var nova = function(d) {
      return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(d.formula).setBackground
      (d.cor)
        .setRanges(sh.getRangeList(d.a1).getRanges()).build();
    };
    sh.setConditionalFormatRules(outras.concat(dados, desejo.dados.map(nova), linhas,
    desejo.linhas.map(nova)));
    return '';
  } catch (err) {
    return ' ATENÇÃO: não foi possível colorir as linhas analisadas na aba ' + sh.getName() + ' ('
    + err.message + ').';
  }
}
function t1MesmasRegras_(atuais, desejadas) {
  if (atuais.length !== desejadas.length) return false;
  var chave = function(formula, cor, a1) {
    return JSON.stringify([String(formula || '').replace(/\s+/g, '').toLowerCase(), String(cor ||
    '').toLowerCase(),
      a1.map(function(x) { return String(x).toUpperCase(); }).sort()]);
  };
  var a = atuais.map(function(regra) {
    var cond = regra.getBooleanCondition();
    return chave((cond.getCriteriaValues() || [])[0], cond.getBackground(),
      regra.getRanges().map(function(r) { return r.getA1Notation(); }));
  }).sort();
  var d = desejadas.map(function(x) { return chave(x.formula, x.cor, x.a1); }).sort();
  return JSON.stringify(a) === JSON.stringify(d);
}
function t1SS_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Nenhuma planilha ativa.');
  return ss;
}
function t1Props_() {
  return PropertiesService.getDocumentProperties();
}
function t1HoraRotina_() {
  if (T1_HORA_ROTINA !== null && T1_HORA_ROTINA !== undefined) return Number(T1_HORA_ROTINA);
  if (typeof HORA_ROTINA_BANCO !== 'undefined' && HORA_ROTINA_BANCO !== null) {
    return (Number(HORA_ROTINA_BANCO) + 1) % 24;
  }
  return 7;
}
function t1Agendar_() {
  var props = t1Props_();
  var pendente = Number(props.getProperty('T1_AGENDADO') || 0);
  var agora = Date.now();
  if (pendente && agora - pendente < T1_ATRASO_SEGUNDOS * 1000 * 3) return;
  props.setProperty('T1_AGENDADO', String(agora));
  ScriptApp.newTrigger('tabela1Reconstruir').timeBased().after(T1_ATRASO_SEGUNDOS *
  1000).create();
}
function t1LimparAgendamento_() {
  t1Props_().deleteProperty('T1_AGENDADO');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tabela1Reconstruir') ScriptApp.deleteTrigger(t);
  });
}
function t1GeracaoRecente_() {
  var ultima = Number(t1Props_().getProperty('T1_ULTIMA_GERACAO') || 0);
  return ultima && Date.now() - ultima < T1_JANELA_IGNORAR_MS;
}
function t1OrigemMudanca_(e) {
  var ignorados = ['FORMAT', 'OTHER'];
  if (e && e.changeType && ignorados.indexOf(String(e.changeType)) >= 0) return '';
  var ativa = null;
  try {
    ativa = (e && e.source ? e.source : t1SS_()).getActiveSheet();
  } catch (err) {
    return 'fonte';
  }
  if (!ativa) return 'fonte';
  var nome = ativa.getName();
  if (nome === T1_PAINEL_ABA) return 'painel';
  if (nome === T1_ABA_SAIDA) return T1_REAGIR_PROPRIA_ABA ? 'painel' : '';
  return T1_ABAS_FONTE.indexOf(nome) >= 0 ? 'fonte' : '';
}
function t1Impressao_() {
  var ss = t1SS_();
  try {
    return 'drv:' + DriveApp.getFileById(ss.getId()).getLastUpdated().getTime();
  } catch (err) {}
  try {
    var sh = t1AbaFonte_(ss);
    return 'dim:' + sh.getLastRow() + 'x' + sh.getLastColumn();
  } catch (err2) {}
  return '';
}
function t1Executar_(interativo, doPainel) {
  var trava = cmTrava_();
  if (!trava.tryLock(doPainel ? T1_ESPERA_TRAVA_PAINEL_MS : 30000)) {
    if (doPainel) {
      t1Agendar_();
      return 'Outra execução em andamento; reconstrução da aba ' + T1_ABA_SAIDA + ' agendada.';
    }
    var ocupado = 'Outra execução da aba ' + T1_ABA_SAIDA + ' está em andamento.';
    if (interativo) t1Avisar_(ocupado);
    return ocupado;
  }
  try {
    var ss = t1SS_();
    var fonte = t1AbaFonte_(ss);
    var entrada = t1LerFonte_(fonte);
    var render = t1Montar_(entrada, ss);
    var assinatura = t1Assinatura_(render);
    var props = t1Props_();
    var existente = ss.getSheetByName(T1_ABA_SAIDA);
    if (existente && !props.getProperty('T1_ASSINATURA') && existente.getLastRow() > 0) {
      var alheia = 'A aba "' + T1_ABA_SAIDA + '" já existe com conteúdo que não foi gerado por este script; nada foi alterado. ' +
        'Renomeie essa aba ou troque T1_ABA_SAIDA no início do arquivo e rode de novo.';
      if (interativo) throw new Error(alheia);
      return alheia;
    }
    if (!cmPrecisaEscrever_(T1_ABA_SAIDA) && !interativo && existente && assinatura === props.getProperty('T1_ASSINATURA')) {
      props.setProperty('T1_IMPRESSAO', t1Impressao_());
      if (render.chaveCorte) props.setProperty('T1_CORTE_USADO', render.chaveCorte);
      if (render.chavePainel !== undefined) props.setProperty('T1_PAINEL_LIDO',
      render.chavePainel);
      t1GruposEstudos_(existente, render.blocos);
      return 'Aba ' + T1_ABA_SAIDA + ' já está em dia.' + t1AplicarDestaque_(fonte,
      render.destaque);
    }
    t1Escrever_(ss, render);
    if (render.depoisDeEscrever) render.depoisDeEscrever(ss);
    t1GruposEstudos_(ss.getSheetByName(T1_ABA_SAIDA), render.blocos);
    var avisoCores = t1AplicarDestaque_(fonte, render.destaque);
    props.setProperty('T1_ASSINATURA', assinatura);
    props.setProperty('T1_ULTIMA_GERACAO', String(Date.now()));
    props.setProperty('T1_IMPRESSAO', t1Impressao_());
    if (render.chaveCorte) props.setProperty('T1_CORTE_USADO', render.chaveCorte);
    if (render.chavePainel !== undefined) props.setProperty('T1_PAINEL_LIDO',
    render.chavePainel);
    cmConcluir_(T1_ABA_SAIDA);
    var msg = 'Aba ' + T1_ABA_SAIDA + ' atualizada. ' + (render.resumo || '') + ' Fonte: ' +
    fonte.getName() + '.';
    if (render.faltando && render.faltando.length) {
      msg += ' ATENÇÃO: colunas não localizadas na origem: ' + render.faltando.join(', ') + '.';
    }
    msg += avisoCores;
    if (interativo) {
      msg += t1GarantirGatilhos_();
      t1Avisar_(msg);
    }
    return msg;
  } finally {
    trava.releaseLock();
  }
}
function t1GarantirGatilhos_() {
  try {
    if (t1Props_().getProperty('T1_DESLIGADO')) {
      return ' Atualização automática desta aba desligada por desligarTabela1; para religar, rode configurarTabela1.';
    }
    var instalado = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'tabela1OnChange';
    });
    return instalado ? '' : ' ' + configurarTabela1(true);
  } catch (err) {
    return ' Não foi possível instalar a atualização automática (' + err.message + '); rode configurarTabela1.';
  }
}
function t1Avisar_(texto) {
  try {
    SpreadsheetApp.getActive().toast(texto, T1_ABA_SAIDA, 8);
  } catch (err) {}
}
function t1AbaFonte_(ss) {
  for (var i = 0; i < T1_ABAS_FONTE.length; i++) {
    var sh = ss.getSheetByName(T1_ABAS_FONTE[i]);
    if (sh) return sh;
  }
  throw new Error('Nenhuma aba de origem encontrada. Procurei por: ' +
  T1_ABAS_FONTE.join(', ') + '.');
}
function t1LerFonte_(sh) {
  var nCols = sh.getLastColumn();
  var nLinhas = sh.getLastRow() - T1_LINHA_DADOS + 1;
  if (nCols < 2) throw new Error('A aba "' + sh.getName() + '" não tem colunas suficientes.');
  var cabecalho = sh.getRange(T1_LINHA_CABECALHO, 1, 1, nCols).getDisplayValues()[0];
  var dados = nLinhas > 0 ? sh.getRange(T1_LINHA_DADOS, 1, nLinhas, nCols).getValues() :
  [];
  return { cabecalho: cabecalho, dados: dados };
}
function t1Assinatura_(render) {
  var texto = JSON.stringify([T1_VERSAO, render.matriz, render.formatos, render.larguras,
  render.alturas]);
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, texto,
    Utilities.Charset.UTF_8);
    var hex = '';
    for (var i = 0; i < bytes.length; i++) {
      var b = (bytes[i] + 256) % 256;
      hex += (b < 16 ? '0' : '') + b.toString(16);
    }
    return hex;
  } catch (err) {
    return 'len' + texto.length;
  }
}
function t1Chave_(v) {
  var s = String(v === null || v === undefined ? '' : v).toLowerCase();
  var saida = '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    saida += T1_ACENTOS[c] || c;
  }
  return saida.replace(/[^a-z0-9]+/g, '');
}
function t1Vazio_(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return isNaN(v.getTime());
  var s = String(v).trim();
  if (!s) return true;
  return T1_MARCADORES_VAZIOS.indexOf(s.toLowerCase()) >= 0;
}
function t1Texto_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + ('0' + v.getDate()).slice(-2);
  }
  return String(v).trim();
}
function t1Numero_(v) {
  if (t1Vazio_(v)) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  if (v instanceof Date) return null;
  var s = String(v).trim().replace(/\s/g, '').replace(/^r\$/i, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function t1Data_(v) {
  if (t1Vazio_(v)) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var iso = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  var br = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
  return null;
}
function t1Indices_(cabecalho, campos) {
  var porChave = {};
  for (var c = 0; c < cabecalho.length; c++) {
    var k = t1Chave_(cabecalho[c]);
    if (!k) continue;
    if (!porChave[k]) porChave[k] = [];
    porChave[k].push(c);
  }
  var mapa = {};
  for (var campo in campos) {
    if (!campos.hasOwnProperty(campo)) continue;
    var achados = [];
    campos[campo].forEach(function(alias) {
      (porChave[t1Chave_(alias)] || []).forEach(function(idx) {
        if (achados.indexOf(idx) < 0) achados.push(idx);
      });
    });
    mapa[campo] = achados;
  }
  return mapa;
}
function t1Faltando_(idx, campos, exigidos) {
  var faltando = [];
  exigidos.forEach(function(alternativas) {
    var achou = alternativas.some(function(c) { return idx[c] && idx[c].length; });
    if (!achou) faltando.push(campos[alternativas[0]][0]);
  });
  return faltando;
}
function t1ValorBruto_(linha, indices) {
  if (!indices || !indices.length) return '';
  for (var i = 0; i < indices.length; i++) {
    var v = linha[indices[i]];
    if (!t1Vazio_(v)) {
      if (linha.usadas) linha.usadas[indices[i]] = true;
      return v;
    }
  }
  return '';
}
function t1Usar_(linha, colunas) {
  if (!linha.usadas) return;
  colunas.forEach(function(c) {
    if (typeof c === 'number' && c >= 0 && !t1Vazio_(linha[c])) linha.usadas[c] = true;
  });
}
function t1CodigoGrupo_(v) {
  if (t1Vazio_(v)) return '';
  var n = t1Numero_(v);
  return n !== null ? String(n) : t1Texto_(v);
}
function t1EstudoEfetivo_(estudo, grupo) {
  if (!t1Vazio_(estudo)) {
    var n = t1Numero_(estudo);
    if (n !== null) return String(n);
    var k = t1Chave_(estudo);
    if (T1_ESTUDO_POR_TEXTO.hasOwnProperty(k)) return T1_ESTUDO_POR_TEXTO[k];
    return t1Texto_(estudo);
  }
  for (var i = 0; i < T1_GRUPOS_POR_ESTUDO.length; i++) {
    if (T1_GRUPOS_POR_ESTUDO[i].grupos.indexOf(grupo) >= 0) return
    T1_GRUPOS_POR_ESTUDO[i].estudo;
  }
  return '';
}
function t1ChavePaciente_(codpac, numeroLinha) {
  if (t1Vazio_(codpac)) return '#linha' + numeroLinha;
  var n = t1Numero_(codpac);
  if (n !== null) return 'p' + n;
  return 'p' + (t1Chave_(codpac) || t1Texto_(codpac));
}
function t1Estudos_(registros) {
  var lista = [];
  T1_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var grupos = cfg.grupos.slice();
    var extras = [];
    var semGrupo = 0;
    registros.forEach(function(r) {
      if (r.estudo !== cfg.estudo) return;
      if (!r.grupo) { semGrupo++; return; }
      if (grupos.indexOf(r.grupo) < 0 && extras.indexOf(r.grupo) < 0) extras.push(r.grupo);
    });
    extras.sort(t1CompararChaves_);
    lista.push({ estudo: cfg.estudo, grupos: grupos.concat(extras), temSemGrupo: semGrupo > 0
});
  });
  return lista;
}
function t1RotuloGrupo_(g) {
  if (T1_ROTULOS.grupo.hasOwnProperty(g)) return T1_ROTULOS.grupo[g];
  return 'Grupo ' + g;
}
function t1RotuloEstudo_(e) {
  return T1_ROTULOS.estudo.hasOwnProperty(e) ? T1_ROTULOS.estudo[e] : 'estudo ' + e;
}
function t1NomeEstudo_(e) {
  return T1_NOME_ESTUDO.hasOwnProperty(e) ? T1_NOME_ESTUDO[e] : String(e);
}
function t1ArredondarPar_(v, casas) {
  var f = Math.pow(10, casas);
  var x = v * f;
  var piso = Math.floor(x);
  var frac = x - piso;
  var r;
  if (Math.abs(frac - 0.5) < 1e-9) r = (piso % 2 === 0) ? piso : piso + 1;
  else r = Math.round(x);
  return r / f;
}
function t1Fmt_(v, casas) {
  if (v === null || v === undefined || isNaN(v)) return '';
  return t1ArredondarPar_(v, casas).toFixed(casas).replace('.', ',');
}
function t1NumeroP_(p) {
  if (p === null || p === undefined || isNaN(p)) return '';
  if (p < 0.001) return '< 0,001';
  return t1ArredondarPar_(p, 3);
}
function t1Plural_(n, singular, plural) {
  return n === 1 ? singular : plural;
}
function t1Contar_(obj) {
  var n = 0;
  for (var k in obj) {
    if (obj.hasOwnProperty(k)) n++;
  }
  return n;
}
function t1LinhaVazia_(n) {
  var linha = [];
  for (var i = 0; i < n; i++) linha.push('');
  return linha;
}
function t1CompararChaves_(a, b) {
  var na = t1Numero_(a);
  var nb = t1Numero_(b);
  if (na !== null && nb !== null) return na - nb;
  if (na !== null) return -1;
  if (nb !== null) return 1;
  return a < b ? -1 : (a > b ? 1 : 0);
}
function t1Letra_(c) {
  var s = '';
  while (c > 0) {
    var m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}
function t1A1_(l, c, nl, nc) {
  var a = t1Letra_(c) + l;
  nl = nl || 1;
  nc = nc || 1;
  if (nl === 1 && nc === 1) return a;
  return a + ':' + t1Letra_(c + nc - 1) + (l + nl - 1);
}
function t1NovoRender_() {
  return { matriz: [], nColunas: 1, formatos: [], larguras: {}, alturas: {}, faltando: [], resumo: '' };
}
function t1Por_(render, l, c, valor) {
  while (render.matriz.length < l) render.matriz.push([]);
  var linha = render.matriz[l - 1];
  while (linha.length < c) linha.push('');
  linha[c - 1] = valor;
}
function t1Fechar_(render) {
  var nC = render.nColunas;
  render.matriz.forEach(function(linha) {
    if (linha.length > nC) nC = linha.length;
  });
  render.nColunas = nC;
  render.matriz.forEach(function(linha) {
    while (linha.length < nC) linha.push('');
  });
  return render;
}
function t1Marcar_(render, tipo, l, c, nl, nc, valor) {
  render.formatos.push({ tipo: tipo, l: l, c: c, nl: nl || 1, nc: nc || 1, valor: valor });
}
function t1Largura_(render, col, px) {
  if (!render.larguras[col] || render.larguras[col] < px) render.larguras[col] = px;
}
function t1Altura_(render, linha, px) {
  if (!render.alturas[linha] || render.alturas[linha] < px) render.alturas[linha] = px;
}
function t1ContarLinhas_(texto, larguraPx, pxPorCaractere) {
  var porLinha = Math.max(1, Math.floor((larguraPx - 8) / (pxPorCaractere || 6.6)));
  var total = 0;
  String(texto).split('\n').forEach(function(paragrafo) {
    var palavras = paragrafo.split(/\s+/).filter(function(p) { return p; });
    if (!palavras.length) { total++; return; }
    var atual = 0;
    var linhas = 1;
    palavras.forEach(function(p) {
      var tam = p.length;
      if (atual === 0) {
        atual = tam;
      } else if (atual + 1 + tam <= porLinha) {
        atual += 1 + tam;
      } else {
        linhas++;
        atual = tam;
      }
      while (atual > porLinha) {
        linhas++;
        atual -= porLinha;
      }
    });
    total += linhas;
  });
  return total;
}
function t1AlturaPara_(nLinhas) {
  return nLinhas <= 1 ? 21 : 21 + 15 * (nLinhas - 1);
}
function t1Escrever_(ss, r) {
  var nomeAba = r.aba || T1_ABA_SAIDA;
  r = cmPreparar_(ss, nomeAba, r);
  var sh = ss.getSheetByName(nomeAba);
  if (!sh) sh = ss.insertSheet(nomeAba);
  var preservar = r.preservar || [];
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  if (preservar.length) sh.clearFormats();
  else sh.clear();
  sh.clearNotes();
  if (r.validacoes) sh.getRange(1, 1, sh.getMaxRows(),
  sh.getMaxColumns()).clearDataValidations();
  if (!r.matriz.length) r.matriz.push(t1LinhaVazia_(Math.max(1, r.nColunas)));
  t1Fechar_(r);
  var nL = r.matriz.length;
  var nC = r.nColunas;
  if (sh.getMaxColumns() < nC) sh.insertColumnsAfter(sh.getMaxColumns(), nC -
  sh.getMaxColumns());
  if (sh.getMaxRows() < nL) sh.insertRowsAfter(sh.getMaxRows(), nL - sh.getMaxRows());
  var tudo = sh.getRange(1, 1, nL, nC);
  t1AplicarGrupos_(sh, r.formatos, function(f) { return f.tipo === 'formato'; });
  cmFormatos_(sh, r);
  if (preservar.length) {
    t1Retangulos_(nL, nC, preservar).forEach(function(q) {
      sh.getRange(q.l, q.c, q.nl, q.nc).setValues(r.matriz.slice(q.l - 1, q.l - 1 +
      q.nl).map(function(linha) {
        return linha.slice(q.c - 1, q.c - 1 + q.nc);
      }));
    });
  } else {
    tudo.setValues(r.matriz);
  }
  tudo.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(false);
  r.formatos.forEach(function(f) {
    if (f.tipo === 'mesclar') sh.getRange(f.l, f.c, f.nl, f.nc).merge();
  });
  t1AplicarGrupos_(sh, r.formatos, function(f) {
    return f.tipo !== 'formato' && f.tipo !== 'mesclar' && f.tipo !== 'rico';
  });
  var negrito = null;
  r.formatos.forEach(function(f) {
    if (f.tipo !== 'rico') return;
    if (!negrito) negrito = SpreadsheetApp.newTextStyle().setBold(true).build();
    var construtor = SpreadsheetApp.newRichTextValue().setText(f.valor.texto);
    f.valor.negritos.forEach(function(par) {
      if (par[1] > par[0]) construtor.setTextStyle(par[0], par[1], negrito);
    });
    sh.getRange(f.l, f.c).setRichTextValue(construtor.build());
  });
  (r.validacoes || []).forEach(function(v) {
    sh.getRange(v.l, v.c).setDataValidation(v.regra);
  });
  t1Medidas_(nC, r.larguras, 62, function(ini, qtd, px) { sh.setColumnWidths(ini, qtd, px); });
  t1Medidas_(nL, r.alturas, 21, function(ini, qtd, px) { sh.setRowHeights(ini, qtd, px); });
  if (r.congelarColunas) sh.setFrozenColumns(r.congelarColunas);
  if (r.congelarLinhas) sh.setFrozenRows(r.congelarLinhas);
  if (sh.getMaxRows() > nL) sh.deleteRows(nL + 1, sh.getMaxRows() - nL);
  if (sh.getMaxColumns() > nC) sh.deleteColumns(nC + 1, sh.getMaxColumns() - nC);
  SpreadsheetApp.flush();
}
function t1Retangulos_(nL, nC, preservar) {
  var furos = {};
  preservar.forEach(function(p) { furos[p.l + ',' + p.c] = true; });
  var segmentos = [];
  for (var c = 1; c <= nC; c++) {
    var lista = [];
    var ini = 0;
    for (var l = 1; l <= nL + 1; l++) {
      var livre = l <= nL && !furos[l + ',' + c];
      if (livre && !ini) ini = l;
      if (!livre && ini) {
        lista.push([ini, l - 1]);
        ini = 0;
      }
    }
    segmentos.push(JSON.stringify(lista));
  }
  var retangulos = [];
  var c0 = 1;
  for (var k = 2; k <= nC + 1; k++) {
    if (k <= nC && segmentos[k - 1] === segmentos[c0 - 1]) continue;
    JSON.parse(segmentos[c0 - 1]).forEach(function(seg) {
      retangulos.push({ l: seg[0], c: c0, nl: seg[1] - seg[0] + 1, nc: k - c0 });
    });
    c0 = k;
  }
  return retangulos;
}
function t1Medidas_(n, mapa, padrao, aplicar) {
  var ini = 1;
  var atual = mapa[1] || padrao;
  for (var i = 2; i <= n + 1; i++) {
    var v = i <= n ? (mapa[i] || padrao) : null;
    if (v !== atual) {
      aplicar(ini, i - ini, atual);
      ini = i;
      atual = v;
    }
  }
}
function t1AplicarGrupos_(sh, formatos, filtro) {
  var grupos = {};
  var ordem = [];
  formatos.forEach(function(f) {
    if (!filtro(f)) return;
    var chave = f.tipo + '|' + JSON.stringify(f.valor);
    if (!grupos[chave]) {
      grupos[chave] = { f: f, a1: [] };
      ordem.push(chave);
    }
    grupos[chave].a1.push(t1A1_(f.l, f.c, f.nl, f.nc));
  });
  ordem.forEach(function(k) {
    var g = grupos[k];
    for (var i = 0; i < g.a1.length; i += 300) {
      t1AplicarLista_(sh.getRangeList(g.a1.slice(i, i + 300)), g.f);
    }
  });
}
function t1AplicarLista_(lista, f) {
  var v = f.valor;
  if (f.tipo === 'negrito') lista.setFontWeight(v ? 'bold' : 'normal');
  else if (f.tipo === 'italico') lista.setFontStyle(v ? 'italic' : 'normal');
  else if (f.tipo === 'fundo') lista.setBackground(v);
  else if (f.tipo === 'formato') lista.setNumberFormat(v);
  else if (f.tipo === 'alinhaH') lista.setHorizontalAlignment(v);
  else if (f.tipo === 'alinhaV') lista.setVerticalAlignment(v);
  else if (f.tipo === 'quebra') lista.setWrap(v);
  else if (f.tipo === 'tamanho') lista.setFontSize(v);
  else if (f.tipo === 'rotacao') lista.setTextRotation(v);
  else if (f.tipo === 'corTexto') lista.setFontColor(v);
  else if (f.tipo === 'borda') {
    if (v.estilo === 'tracejada') {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    } else {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null);
    }
  }
}
function t1Caixa_(render, l, c, nl, nc, estilo) {
  t1Marcar_(render, 'borda', l, c, nl, nc, { t: true, l: true, b: true, r: true, estilo: estilo || 'solida' });
}
function t1LinhaBaixo_(render, l, c, nc) {
  t1Marcar_(render, 'borda', l, c, 1, nc, { t: null, l: null, b: true, r: null, estilo: 'solida' });
}
function t1Laterais_(render, l, c, nl) {
  t1Marcar_(render, 'borda', l, c, nl, 1, { t: null, l: true, b: null, r: true, estilo: 'solida' });
}
function t1Quadro_(render, q) {
  t1Laterais_(render, q.r1, 1, q.rFim - q.r1 + 1);
  [q.r1, q.rGrupo, q.rFimCabecalho, q.rFim].forEach(function(l) {
    t1LinhaBaixo_(render, l, 1, 1);
  });
  q.caixas.forEach(function(cx) {
    var nc = cx.c2 - cx.c1 + 1;
    t1Caixa_(render, q.r1 + 1, cx.c1, q.rFim - q.r1, nc);
    t1LinhaBaixo_(render, q.rGrupo, cx.c1, nc);
    t1LinhaBaixo_(render, q.rFimCabecalho, cx.c1, nc);
  });
}
function t1CabecalhoGrupo_(render, l, c, texto, larguraPx) {
  t1Por_(render, l, c, texto);
  t1Marcar_(render, 'quebra', l, c, 1, 1, true);
  t1Altura_(render, l, t1AlturaPara_(t1ContarLinhas_(texto, larguraPx, 7)));
}
function t1LerPainel_(ss) {
  var valores = {};
  try {
    valores = JSON.parse(t1Props_().getProperty(T1_PAINEL_VAL) || '{}') || {};
  } catch (err) {
    valores = {};
  }
  var lidos = t1LerCelulasPainel_(ss);
  if (lidos) {
    for (var est in lidos) {
      if (!lidos.hasOwnProperty(est)) continue;
      if (!valores[est]) valores[est] = {};
      for (var campo in lidos[est]) {
        if (lidos[est].hasOwnProperty(campo)) valores[est][campo] = lidos[est][campo];
      }
    }
  }
  return { valores: valores, lidos: lidos };
}
function t1LerCelulasPainel_(ss) {
  var pos = null;
  try {
    pos = JSON.parse(t1Props_().getProperty(T1_PAINEL_POS) || 'null');
  } catch (err) {
    pos = null;
  }
  var sh = ss ? ss.getSheetByName(T1_PAINEL_ABA) : null;
  if (!sh || !pos) return null;
  var celulas = [];
  for (var est in pos) {
    if (!pos.hasOwnProperty(est)) continue;
    for (var campo in pos[est]) {
      if (!pos[est].hasOwnProperty(campo)) continue;
      var m = String(pos[est][campo]).match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      var c = 0;
      for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
      celulas.push({ est: est, campo: campo, l: Number(m[2]), c: c });
    }
  }
  if (!celulas.length) return null;
  var l1 = Infinity, c1 = Infinity, l2 = 0, c2 = 0;
  celulas.forEach(function(x) {
    l1 = Math.min(l1, x.l);
    c1 = Math.min(c1, x.c);
    l2 = Math.max(l2, x.l);
    c2 = Math.max(c2, x.c);
  });
  if (l2 > sh.getMaxRows() || c2 > sh.getMaxColumns()) return null;
  var grade = sh.getRange(l1, c1, l2 - l1 + 1, c2 - c1 + 1).getDisplayValues();
  var lidos = {};
  celulas.forEach(function(x) {
    var v = grade[x.l - l1][x.c - c1];
    if (!lidos[x.est]) lidos[x.est] = {};
    lidos[x.est][x.campo] = String(v === null || v === undefined ? '' : v).trim();
  });
  return lidos;
}
function t1PainelMudou_() {
  try {
    var ss = t1SS_();
    var props = t1Props_();
    if (T1_ABA_SAIDA === T1_PAINEL_ABA) {
      var lidos = t1LerCelulasPainel_(ss);
      return !!lidos && JSON.stringify(lidos) !== (props.getProperty('T1_PAINEL_LIDO') || '');
    }
    return t1ChaveCortes_(t1Cortes_(ss, new Date())) !==
    (props.getProperty('T1_CORTE_USADO') || '');
  } catch (err) {
    return false;
  }
}
function t1AtualizarPeloPainel_() {
  if (T1_ABA_SAIDA !== T1_PAINEL_ABA) Utilities.sleep(T1_VANTAGEM_PAINEL_MS);
  return t1Executar_(false, true);
}
function t1ChaveCortes_(cortes) {
  return T1_GRUPOS_POR_ESTUDO.map(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (!c.periodo) return cfg.estudo + ':auto';
    return cfg.estudo + ':' + (c.coderest !== null ? c.coderest : '') + '#' + (c.inicio ?
    t1DataBr_(c.inicio) : '') +
      '-' + (c.fim ? t1DataBr_(c.fim) : '');
  }).join('|');
}
function t1FaltaCoderest_(idx, cortes) {
  if (idx.coderest && idx.coderest.length) return false;
  return T1_GRUPOS_POR_ESTUDO.some(function(cfg) { return cortes.estudos[cfg.estudo].coderest !== null; });
}
function t1ModoCorte_(v) {
  var k = t1Chave_(v);
  return !k || k.indexOf('auto') === 0 ? T1_MODO_AUTO : T1_MODO_PERIODO;
}
function t1DataEstrita_(v) {
  if (t1Vazio_(v)) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(),
  v.getDate());
  var s = String(v).trim();
  var br = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(\D|$)/);
  var iso = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(\D|$)/);
  var a, m, d;
  if (br) {
    a = Number(br[3]);
    m = Number(br[2]);
    d = Number(br[1]);
  } else if (iso) {
    a = Number(iso[1]);
    m = Number(iso[2]);
    d = Number(iso[3]);
  } else {
    return null;
  }
  var data = new Date(a, m - 1, d);
  if (data.getFullYear() !== a || data.getMonth() !== m - 1 || data.getDate() !== d) return null;
  return data;
}
function t1Dia_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}
function t1Corte_(v, hoje) {
  var corte = { periodo: false, inicio: null, fim: null, coderest: null, aviso: '',
    referencia: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()) };
  if (!v || t1ModoCorte_(v.corte) !== T1_MODO_PERIODO) return corte;
  var inicio = t1DataEstrita_(v.inicio);
  var fim = t1DataEstrita_(v.fim);
  var coderest = t1Vazio_(v.coderest) ? null : t1Numero_(v.coderest);
  var invalidos = [];
  if (!t1Vazio_(v.inicio) && !inicio) invalidos.push('a data inicial');
  if (!t1Vazio_(v.fim) && !fim) invalidos.push('a data final');
  if (!t1Vazio_(v.coderest) && (coderest === null || coderest < 1 || Math.round(coderest) !==
  coderest)) {
    invalidos.push('o coderest');
  }
  if (invalidos.length) {
    corte.aviso = 'não entendi ' + invalidos.join(' nem ') + '; corte ignorado, valendo auto.';
    return corte;
  }
  if (!inicio && !fim && coderest === null) {
    corte.aviso = 'preencha as datas e/ou o coderest; sem nenhum deles, vale auto.';
    return corte;
  }
  if (inicio && fim && inicio.getTime() > fim.getTime()) {
    corte.aviso = 'a data inicial é depois da data final; corte ignorado, valendo auto.';
    return corte;
  }
  corte.periodo = true;
  corte.inicio = inicio;
  corte.fim = fim;
  corte.coderest = coderest;
  if (fim && fim.getTime() < corte.referencia.getTime()) corte.referencia = fim;
  return corte;
}
function t1Cortes_(ss, hoje) {
  var painel = t1LerPainel_(ss);
  var cortes = { estudos: {}, padrao: t1Corte_(null, hoje), valores: painel.valores, lidos:
    painel.lidos };
  T1_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    cortes.estudos[cfg.estudo] = t1Corte_(painel.valores[cfg.estudo], hoje);
  });
  return cortes;
}
function t1CorteDe_(cortes, estudo) {
  return cortes.estudos.hasOwnProperty(estudo) ? cortes.estudos[estudo] : cortes.padrao;
}
function t1DentroDoCorte_(corte, data0, coderest) {
  if (!corte.periodo) return true;
  if (corte.coderest !== null && (coderest === null || coderest > corte.coderest)) return false;
  if (!data0) return true;
  var t = t1Dia_(data0);
  if (corte.inicio && t < corte.inicio.getTime()) return false;
  if (corte.fim && corte.coderest === null && t > corte.fim.getTime()) return false;
  return true;
}
function t1Data0NoCorte_(corte, data0) {
  if (!data0 || !corte.periodo || !corte.fim) return data0;
  return t1Dia_(data0) > corte.fim.getTime() ? null : data0;
}
function t1DataBr_(d) {
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}
function t1TextoCorte_(corte) {
  if (!corte.periodo) return '';
  var partes = [];
  if (corte.coderest !== null) partes.push('coderest até ' + corte.coderest);
  if (corte.inicio && corte.fim) partes.push(t1DataBr_(corte.inicio) + ' a ' + t1DataBr_(corte.fim));
  else if (corte.inicio) partes.push('a partir de ' + t1DataBr_(corte.inicio));
  else if (corte.fim) partes.push('até ' + t1DataBr_(corte.fim));
  return partes.join('; ');
}
function t1RotuloComCorte_(rotulo, corte) {
  var texto = t1TextoCorte_(corte);
  return texto ? rotulo + ' (corte: ' + texto + ')' : rotulo;
}
function t1AvisosCorte_(cortes) {
  var avisos = [];
  T1_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (c && c.aviso) avisos.push('corte do estudo ' + t1NomeEstudo_(cfg.estudo) + ': ' + c.aviso);
  });
  return avisos;
}
function t1Montar_(entrada, ss) {
  var idx = t1Indices_(entrada.cabecalho, T1_CAMPOS);
  var faltando = t1Faltando_(idx, T1_CAMPOS, [['grupo'], ['sexo'], ['idade', 'nascimento'],
    ['escolaridade'], ['renda'], ['conjunto', 'dente'], ['arcada', 'dente'], ['faces'], ['vitalidade'],
    ['riscoOclusal'], ['riscoCarie']]);
  var cortes = t1Cortes_(ss, new Date());
  if (t1FaltaCoderest_(idx, cortes)) faltando.push('coderest (usado no corte)');
  var registros = t1Registros_(entrada.dados, idx, cortes);
  var render = t1NovoRender_();
  var res = t1Tabela_(render, t1Estudos_(registros), registros, 1, cortes);
  render.faltando = faltando;
  render.resumo = res.partes.concat(t1AvisosCorte_(cortes)).join('; ') + '.';
  render.chaveCorte = t1ChaveCortes_(cortes);
  render.congelarColunas = 1;
  render.congelarLinhas = res.linhasCabecalho;
  render.blocos = { rotulos: [1, 2], estudos: res.blocos };
  render.destaque = t1Destaque_(registros);
  return t1Fechar_(render);
}
function CONFERIR_TABELA1() {
  var trava = cmTrava_();
  if (!trava.tryLock(30000)) {
    var ocupado = 'Outra execução está em andamento; rode CONFERIR_TABELA1 de novo em instantes.';
    t1Avisar_(ocupado);
    return ocupado;
  }
  try {
    var ss = t1SS_();
    var alheia = t1ConferenciaAlheia_(ss);
    if (alheia) throw new Error(alheia);
    var entrada = t1LerFonte_(t1AbaFonte_(ss));
    var idx = t1Indices_(entrada.cabecalho, T1_CAMPOS);
    var cortes = t1Cortes_(ss, new Date());
    var registros = t1Registros_(entrada.dados, idx, cortes);
    t1Escrever_(ss, t1Conferencia_(registros, cortes, t1OutrosDoPaciente_(entrada.dados, idx,
    registros)));
    t1Props_().setProperty('T1_CONFERENCIA_ID',
    String(ss.getSheetByName(T1_ABA_CONFERENCIA).getSheetId()));
    var msg = 'Aba ' + T1_ABA_CONFERENCIA + ' atualizada: ' + registros.length + ' ' +
      t1Plural_(registros.length, 'restauração', 'restaurações') + ', a mesma amostra da Tabela 1.';
    t1Avisar_(msg);
    return msg;
  } finally {
    trava.releaseLock();
  }
}
function t1ConferenciaAlheia_(ss) {
  var sh = ss.getSheetByName(T1_ABA_CONFERENCIA);
  if (!sh || sh.getLastRow() === 0) return '';
  if (String(sh.getSheetId()) === t1Props_().getProperty('T1_CONFERENCIA_ID')) return '';
  return 'A aba "' + T1_ABA_CONFERENCIA + '" já existe com conteúdo que não foi gerado por este script; nada foi alterado. ' +
    'Renomeie essa aba e rode CONFERIR_TABELA1 de novo.';
}
function t1Conferencia_(registros, cortes, outros) {
  var render = t1NovoRender_();
  render.aba = T1_ABA_CONFERENCIA;
  var cols = T1_COLUNAS_CONFERENCIA;
  var nC = cols.length;
  var agora = new Date();
  var textos = T1_GRUPOS_POR_ESTUDO.map(function(cfg) {
    var t = t1TextoCorte_(t1CorteDe_(cortes, cfg.estudo));
    return 'estudo ' + t1NomeEstudo_(cfg.estudo) + ': ' + (t ? 'corte ' + t : 'auto');
  });
  t1Por_(render, 1, 1, 'Conferência da Tabela 1: renda e nº de dentes de cada restauração da amostra da Tabela 1');
  t1Por_(render, 2, 1, textos.join(' · ') + '. Gerada em ' + t1DataBr_(agora) + ' às ' +
    ('0' + agora.getHours()).slice(-2) + ':' + ('0' + agora.getMinutes()).slice(-2) + '.');
  t1Marcar_(render, 'negrito', 1, 1, 1, 1, true);
  t1Marcar_(render, 'tamanho', 1, 1, 1, 1, 11);
  t1Marcar_(render, 'italico', 2, 1, 1, 1, true);
  var rCab = 4;
  var linhasCab = 1;
  cols.forEach(function(col, i) {
    t1Por_(render, rCab, i + 1, col.rotulo);
    t1Largura_(render, i + 1, col.largura);
    linhasCab = Math.max(linhasCab, t1ContarLinhas_(col.rotulo, col.largura, 7));
  });
  t1Altura_(render, rCab, t1AlturaPara_(linhasCab));
  t1Marcar_(render, 'negrito', rCab, 1, 1, nC, true);
  t1Marcar_(render, 'fundo', rCab, 1, 1, nC, T1_COR_SECAO);
  t1Marcar_(render, 'quebra', rCab, 1, 1, nC, true);
  var ordenados = t1OrdemConferencia_(registros);
  ordenados.forEach(function(r, k) {
    var c = r.conferencia;
    [
      t1NomeEstudo_(r.estudo),
      r.grupo ? t1RotuloGrupo_(r.grupo) : 'sem sorteio',
      t1Vazio_(c.codpac) ? '' : c.codpac,
      c.coderest === null ? '' : c.coderest,
      c.dentes.join(' / '),
      c.linha,
      c.data0 || '',
      c.renda.data || '',
      c.renda.bruto,
      c.renda.salario === null ? '' : c.renda.salario,
      c.renda.sm === null ? '' : c.renda.sm,
      r.renda || T1_ROTULO_SEM_INFO,
      r.nDentes,
      r.dentesPaciente || T1_ROTULO_SEM_INFO,
      t1Observacao_(r, (outros || {})[c.linha])
    ].forEach(function(v, i) { t1Por_(render, rCab + 1 + k, i + 1, v); });
  });
  if (ordenados.length) {
    cols.forEach(function(col, i) {
      if (col.formato) t1Marcar_(render, 'formato', rCab + 1, i + 1, ordenados.length, 1,
      col.formato);
    });
  }
  render.nColunas = nC;
  render.congelarLinhas = rCab;
  return t1Fechar_(render);
}
function t1OrdemConferencia_(registros) {
  var estudos = T1_GRUPOS_POR_ESTUDO.map(function(cfg) { return cfg.estudo; });
  function posEstudo(e) {
    var i = estudos.indexOf(e);
    return i < 0 ? estudos.length : i;
  }
  function posGrupo(r) {
    if (!r.grupo) return 2;
    for (var i = 0; i < T1_GRUPOS_POR_ESTUDO.length; i++) {
      if (T1_GRUPOS_POR_ESTUDO[i].grupos.indexOf(r.grupo) >= 0) return 0;
    }
    return 1;
  }
  function coderest(r) {
    return r.conferencia.coderest === null ? Number.MAX_VALUE : r.conferencia.coderest;
  }
  return registros.slice().sort(function(a, b) {
    return posEstudo(a.estudo) - posEstudo(b.estudo) || posGrupo(a) - posGrupo(b) ||
      t1CompararChaves_(a.grupo, b.grupo) || coderest(a) - coderest(b);
  });
}
function t1Observacao_(r, outros) {
  var notas = [];
  if (r.conferencia.dentes.length > 1) {
    notas.push('nº de dente diferente entre as colunas do banco: ' + r.conferencia.dentes.join(' / '));
  }
  if (r.mesmoDente.length) {
    notas.push('outra restauração do paciente com o mesmo nº de dente: ' +
    r.mesmoDente.map(function(o) {
      return t1Identificacao_(o.conferencia.coderest, o.conferencia.linha);
    }).join(', '));
  }
  if (outros && outros.length) notas.push('o paciente também tem ' + outros.join(' e '));
  return notas.join('; ');
}
function t1Identificacao_(coderest, linha) {
  return coderest === null ? 'linha ' + linha : 'coderest ' + coderest;
}
function t1OutrosDoPaciente_(dados, idx, registros) {
  var naAmostra = {};
  registros.forEach(function(r) { naAmostra[r.conferencia.linha] = r; });
  var porPaciente = {};
  dados.forEach(function(linha, i) {
    var codpac = t1ValorBruto_(linha, idx.codpac);
    if (t1Vazio_(codpac)) return;
    var grupo = t1CodigoGrupo_(t1ValorBruto_(linha, idx.grupo));
    var k = t1EstudoEfetivo_(t1ValorBruto_(linha, idx.estudo), grupo) + '|' +
    t1ChavePaciente_(codpac, 0);
    if (!porPaciente[k]) porPaciente[k] = [];
    porPaciente[k].push({ linha: i + T1_LINHA_DADOS, grupo: grupo,
      coderest: t1Numero_(t1ValorBruto_(linha, idx.coderest)) });
  });
  var saida = {};
  registros.forEach(function(r) {
    var proprio = r.conferencia.linha;
    saida[proprio] = (porPaciente[r.estudo + '|' + r.paciente] || []).filter(function(o) {
      var outro = naAmostra[o.linha];
      return o.linha !== proprio && !(outro && t1ChaveContagem_(outro) ===
      t1ChaveContagem_(r));
    }).map(function(o) {
      var motivo = !naAmostra[o.linha] ? 'fora da amostra' : (o.grupo ? 'sorteada' : 'sem sorteio');
      return t1Identificacao_(o.coderest, o.linha) + ' (' + motivo + ')';
    });
  });
  return saida;
}
function t1DentesDaLinha_(linha, indices) {
  var vistos = [];
  (indices || []).forEach(function(c) {
    var v = linha[c];
    if (t1Vazio_(v)) return;
    var n = t1Numero_(v);
    var t = n !== null ? String(n) : t1Texto_(v);
    if (vistos.indexOf(t) < 0) vistos.push(t);
  });
  return vistos;
}
function t1Registros_(dados, idx, cortes) {
  var registros = [];
  dados.forEach(function(linha, i) {
    linha.usadas = {};
    var codpac = t1ValorBruto_(linha, idx.codpac);
    var grupoBruto = t1ValorBruto_(linha, idx.grupo);
    var estudoBruto = t1ValorBruto_(linha, idx.estudo);
    var dente = t1ValorBruto_(linha, idx.dente);
    if (t1Vazio_(codpac) && t1Vazio_(grupoBruto) && t1Vazio_(dente) && t1Vazio_(estudoBruto))
    return;
    var grupo = t1CodigoGrupo_(grupoBruto);
    var estudo = t1EstudoEfetivo_(estudoBruto, grupo);
    var corte = t1CorteDe_(cortes, estudo);
    var data0 = t1Data_(t1ValorBruto_(linha, idx.data0));
    if (!t1DentroDoCorte_(corte, data0, t1Numero_(t1ValorBruto_(linha, idx.coderest)))) return;
    if (!t1Data0NoCorte_(corte, data0)) return;
    var nDente = t1Numero_(dente);
    var renda = t1RendaDetalhe_(linha, idx);
    registros.push({
      paciente: t1ChavePaciente_(codpac, i + T1_LINHA_DADOS),
      dente: t1Vazio_(dente) ? '' : (nDente !== null ? String(nDente) : t1Chave_(dente)),
      estudo: estudo,
      grupo: grupo,
      idade: t1Faixa_(T1_FAIXAS_IDADE, t1Idade_(linha, idx)),
      sexo: t1Sexo_(t1ValorBruto_(linha, idx.sexo)),
      escolaridade: t1Escolaridade_(t1ValorBruto_(linha, idx.escolaridade)),
      renda: t1Faixa_(T1_FAIXAS_RENDA, renda.sm),
      conjunto: t1Conjunto_(t1ValorBruto_(linha, idx.conjunto), dente),
      arcada: t1Arcada_(t1ValorBruto_(linha, idx.arcada), dente),
      faces: t1Inteiro_(t1ValorBruto_(linha, idx.faces)),
      vitalidade: t1Vitalidade_(t1ValorBruto_(linha, idx.vitalidade)),
      ano: String(data0.getFullYear()),
      riscoOclusal: t1Risco_(t1ValorBruto_(linha, idx.riscoOclusal)),
      riscoCarie: t1Risco_(t1ValorBruto_(linha, idx.riscoCarie)),
      pacote: { linha: i + T1_LINHA_DADOS, usadas: linha.usadas },
      conferencia: { codpac: codpac, coderest: t1Numero_(t1ValorBruto_(linha, idx.coderest)),
        data0: data0, renda: renda,
        dentes: t1DentesDaLinha_(linha, idx.dente), linha: i + T1_LINHA_DADOS }
    });
  });
  t1DentesPorPaciente_(registros);
  return registros;
}
function t1DentesPorPaciente_(registros) {
  var porPaciente = {};
  registros.forEach(function(r) {
    var k = t1ChaveContagem_(r);
    if (!porPaciente[k]) porPaciente[k] = [];
    porPaciente[k].push(r);
  });
  registros.forEach(function(r) {
    var juntos = porPaciente[t1ChaveContagem_(r)];
    r.nDentes = juntos.length;
    r.dentesPaciente = t1Faixa_(T1_FAIXAS_DENTES, r.nDentes);
    r.mesmoDente = juntos.filter(function(o) { return o !== r && !!r.dente && o.dente === r.dente;
});
  });
}
function t1ChaveContagem_(r) {
  return r.estudo + '|' + (r.grupo ? 'g' : 's') + '|' + r.paciente;
}
function t1Faixa_(faixas, n) {
  if (n === null || n === undefined || isNaN(n)) return null;
  for (var i = 0; i < faixas.length; i++) {
    var f = faixas[i];
    if (f.min !== undefined && f.min !== null && n < f.min - 1e-9) continue;
    if (f.acimaDe !== undefined && f.acimaDe !== null && n <= f.acimaDe + 1e-9) continue;
    if (f.max !== undefined && f.max !== null && n > f.max + 1e-9) continue;
    return f.rotulo;
  }
  return null;
}
function t1Idade_(linha, idx) {
  var nasc = t1Data_(t1ValorBruto_(linha, idx.nascimento));
  var base = t1Data_(t1ValorBruto_(linha, idx.data0));
  if (nasc && base) {
    var anos = base.getFullYear() - nasc.getFullYear();
    if (base.getMonth() < nasc.getMonth() ||
        (base.getMonth() === nasc.getMonth() && base.getDate() < nasc.getDate())) anos--;
    if (anos >= 0) return anos;
  }
  var direta = t1Numero_(t1ValorBruto_(linha, idx.idade));
  return direta !== null && direta >= 0 ? Math.floor(direta) : null;
}
function t1SemAcento_(s) {
  var t = String(s).toLowerCase();
  var saida = '';
  for (var i = 0; i < t.length; i++) saida += T1_ACENTOS[t.charAt(i)] || t.charAt(i);
  return saida;
}
function t1RendaDetalhe_(linha, idx) {
  var bruto = t1ValorBruto_(linha, idx.renda);
  var r = { bruto: t1Vazio_(bruto) ? '' : bruto, data: null, salario: null, sm: null };
  if (t1Vazio_(bruto) || bruto instanceof Date) return r;
  if (typeof bruto === 'string') {
    var t = t1SemAcento_(bruto);
    if (/salario|\bsm\b/.test(t)) {
      var m = t.match(/(\d+(?:[.,]\d+)?)\s*(?:\/\s*(\d+))?\s*(?:salarios?|sm)/);
      if (m) {
        var q = parseFloat(m[1].replace(',', '.'));
        if (m[2]) q = q / parseFloat(m[2]);
        r.sm = isNaN(q) ? null : q;
        return r;
      }
      var palavras = { um: 1, uma: 1, meio: 0.5, meia: 0.5, dois: 2, duas: 2, tres: 3, quatro: 4,
        cinco: 5, seis: 6, sete: 7, oito: 8, nove: 9, dez: 10 };
      for (var p in palavras) {
        if (palavras.hasOwnProperty(p) && new RegExp('(^|[^a-z])' + p +
        '[^a-z]*(salarios?|sm)').test(t)) {
          r.sm = palavras[p];
          return r;
        }
      }
      r.sm = 1;
      return r;
    }
  }
  var reais = t1Numero_(bruto);
  if (reais === null || reais < 0) return r;
  r.data = t1Data_(t1ValorBruto_(linha, idx.consulta)) ||
    t1Data_(t1ValorBruto_(linha, idx.data0)) || new Date();
  r.salario = t1SalarioVigente_(r.data);
  r.sm = reais / r.salario;
  return r;
}
function t1SalarioVigente_(data) {
  var alvo = data.getFullYear() + '-' + ('0' + (data.getMonth() + 1)).slice(-2) + '-' + ('0' +
  data.getDate()).slice(-2);
  var achado = T1_SALARIOS_MINIMOS[0].valor;
  for (var i = 0; i < T1_SALARIOS_MINIMOS.length; i++) {
    if (T1_SALARIOS_MINIMOS[i].desde <= alvo) achado = T1_SALARIOS_MINIMOS[i].valor;
  }
  return achado;
}
function t1Escolaridade_(bruto) {
  if (t1Vazio_(bruto)) return null;
  var n = t1Numero_(bruto);
  var k = t1Chave_(bruto);
  for (var i = 0; i < T1_ESCOLARIDADE.length; i++) {
    var e = T1_ESCOLARIDADE[i];
    if (n !== null) {
      if (e.codigos.indexOf(String(n)) >= 0) return e.rotulo;
    } else if (e.textos.some(function(t) { return t1Chave_(t) === k; })) {
      return e.rotulo;
    }
  }
  return null;
}
function t1Sexo_(bruto) {
  if (t1Vazio_(bruto)) return null;
  var n = t1Numero_(bruto);
  if (n === 0) return 'Masculino';
  if (n === 1) return 'Feminino';
  var k = t1Chave_(bruto);
  if (k === 'f' || k.indexOf('fem') === 0 || k.indexOf('mulher') === 0) return 'Feminino';
  if (k === 'm' || k.indexOf('masc') === 0 || k.indexOf('homem') === 0) return 'Masculino';
  return t1Texto_(bruto);
}
function t1Conjunto_(bruto, dente) {
  if (!t1Vazio_(bruto)) {
    var k = t1Chave_(bruto);
    if (k.indexOf('premolar') === 0) return 'Pré-molar';
    if (k.indexOf('molar') === 0) return 'Molar';
    if (k.indexOf('incisivo') === 0) return 'Incisivo';
    if (k.indexOf('canino') === 0) return 'Canino';
    return t1Texto_(bruto);
  }
  var n = t1Numero_(dente);
  if (n === null) return null;
  var u = Math.abs(Math.round(n)) % 10;
  if (u === 1 || u === 2) return 'Incisivo';
  if (u === 3) return 'Canino';
  if (u === 4 || u === 5) return 'Pré-molar';
  if (u >= 6 && u <= 8) return 'Molar';
  return null;
}
function t1Arcada_(bruto, dente) {
  if (!t1Vazio_(bruto)) {
    var k = t1Chave_(bruto);
    if (k.indexOf('sup') === 0 || k === 'maxila') return 'Superior';
    if (k.indexOf('inf') === 0 || k === 'mandibula') return 'Inferior';
    return t1Texto_(bruto);
  }
  var n = t1Numero_(dente);
  if (n === null) return null;
  var q = Math.floor(Math.abs(Math.round(n)) / 10);
  if (q === 1 || q === 2 || q === 5 || q === 6) return 'Superior';
  if (q === 3 || q === 4 || q === 7 || q === 8) return 'Inferior';
  return null;
}
function t1Inteiro_(bruto) {
  var n = t1Numero_(bruto);
  if (n === null) return null;
  return String(Math.round(n));
}
function t1Vitalidade_(bruto) {
  if (t1Vazio_(bruto)) return null;
  var n = t1Numero_(bruto);
  if (n === 1) return 'Vital';
  if (n === 0) return 'Não vital';
  var k = t1Chave_(bruto);
  if (k === 'sim' || k === 's' || k === 'vital') return 'Vital';
  if (k === 'nao' || k === 'n' || k === 'naovital') return 'Não vital';
  return t1Texto_(bruto);
}
function t1Risco_(bruto) {
  if (t1Vazio_(bruto)) return null;
  var n = t1Numero_(bruto);
  var codigo = null;
  if (n === 0 || n === 1) {
    codigo = String(n);
  } else {
    var k = t1Chave_(bruto);
    if (k.indexOf('baixo') === 0) codigo = '0';
    else if (k.indexOf('alto') === 0) codigo = '1';
  }
  if (codigo === null) return t1Texto_(bruto);
  return T1_ROTULOS_RISCO ? T1_ROTULOS_RISCO[codigo] : codigo;
}
function t1Fixas_(v) {
  if (v.faixas) return v.faixas.map(function(f) { return f.rotulo; });
  if (v.campo === 'riscoOclusal' || v.campo === 'riscoCarie') {
    return T1_ROTULOS_RISCO ? [T1_ROTULOS_RISCO['0'], T1_ROTULOS_RISCO['1']] : ['0',
    '1'];
  }
  if (v.campo === 'escolaridade') return T1_ESCOLARIDADE.map(function(e) { return e.rotulo;
});
  return (v.fixas || []).slice();
}
function t1RotuloSecao_(chave) {
  for (var i = 0; i < T1_SECOES.length; i++) {
    if (T1_SECOES[i].chave === chave) return T1_SECOES[i].rotulo;
  }
  return '';
}
function t1Tabela_(render, estudos, registros, r1, cortes) {
  var blocos = [];
  var colunas = [];
  var caixas = [];
  var partes = [];
  var faixas = [];
  var c = 3;
  estudos.forEach(function(est, i) {
    if (i > 0) {
      t1Largura_(render, c, T1_LARGURA_ESPACO);
      c++;
    }
    var comGrupo = registros.filter(function(r) { return r.estudo === est.estudo && !!r.grupo; });
    var corte = t1CorteDe_(cortes, est.estudo);
    var bloco = { estudo: est.estudo, c1: c, grupos: [], corte: corte };
    est.grupos.forEach(function(g) {
      var col = { chave: est.estudo + '|' + g, rotulo: t1RotuloGrupo_(g), col: c,
        regs: comGrupo.filter(function(r) { return r.grupo === g; }) };
      bloco.grupos.push(col);
      colunas.push(col);
      c += 2;
    });
    colunas.push({ chave: est.estudo + '|__total', rotulo: T1_ROTULO_TOTAL, col: c, regs:
      comGrupo });
    c += 2;
    bloco.colQui = c;
    bloco.colFisher = c + 1;
    c += 2;
    caixas.push({ c1: bloco.c1, c2: c - 1 });
    faixas.push({ chave: est.estudo, nome: t1NomeEstudo_(est.estudo), c1: bloco.c1, c2: c - 1 });
    blocos.push(bloco);
    partes.push(t1RotuloComCorte_('estudo ' + t1NomeEstudo_(est.estudo) + ': n = ' +
    comGrupo.length, corte));
  });
  var semGrupo = registros.filter(function(r) { return !r.grupo; });
  if (semGrupo.length) {
    t1Largura_(render, c, T1_LARGURA_ESPACO);
    c++;
    colunas.push({ chave: '__sem', rotulo: T1_ROTULO_SEM_GRUPO, col: c, regs: semGrupo
});
    caixas.push({ c1: c, c2: c + 1 });
    faixas.push({ chave: 'sem', nome: 'Sem sorteio', c1: c, c2: c + 1 });
    c += 2;
    partes.push('sem sorteio: ' + semGrupo.length);
  }
  var nC = c - 1;
  if (render.nColunas < nC) render.nColunas = nC;
  t1Largura_(render, 1, T1_LARGURA_ROTULO);
  t1Largura_(render, 2, T1_LARGURA_ESPACO);
  colunas.forEach(function(col) {
    t1Largura_(render, col.col, T1_LARGURA_N);
    t1Largura_(render, col.col + 1, T1_LARGURA_PCT);
  });
  blocos.forEach(function(b) {
    t1Largura_(render, b.colQui, T1_LARGURA_P);
    t1Largura_(render, b.colFisher, T1_LARGURA_P);
  });
  t1Por_(render, r1, 1, T1_ROTULO_TABELA);
  t1Por_(render, r1, 2, T1_TITULO);
  t1Marcar_(render, 'negrito', r1, 1, 1, nC, true);
  t1Marcar_(render, 'tamanho', r1, 1, 1, nC, 11);
  var rEstudo = r1 + 1;
  var rGrupo = r1 + 2;
  var rN = r1 + 3;
  var rSub = r1 + 4;
  t1Por_(render, rEstudo, 1, T1_ROTULO_LINHA_ESTUDO);
  t1Por_(render, rGrupo, 1, T1_ROTULO_LINHA_GRUPO);
  t1Por_(render, rSub, 1, T1_ROTULO_VARIAVEL);
  blocos.forEach(function(b) {
    t1Por_(render, rEstudo, b.c1, t1RotuloComCorte_(t1RotuloEstudo_(b.estudo), b.corte));
    t1Por_(render, rGrupo, b.colQui, T1_ROTULO_VALOR_P);
    t1Por_(render, rSub, b.colQui, T1_ROTULO_QUI);
    t1Por_(render, rSub, b.colFisher, T1_ROTULO_FISHER);
  });
  colunas.forEach(function(col) {
    t1CabecalhoGrupo_(render, rGrupo, col.col, col.rotulo, T1_LARGURA_N);
    t1Por_(render, rN, col.col, 'n = ' + col.regs.length);
    t1Por_(render, rSub, col.col, 'n');
    t1Por_(render, rSub, col.col + 1, '%');
  });
  t1Marcar_(render, 'negrito', rEstudo, 1, 4, nC, true);
  var l = rSub + 1;
  var secaoAtual = null;
  T1_VARIAVEIS.forEach(function(v) {
    if (v.secao !== secaoAtual) {
      secaoAtual = v.secao;
      t1Por_(render, l, 1, t1RotuloSecao_(v.secao));
      t1Marcar_(render, 'negrito', l, 1, 1, 1, true);
      t1Marcar_(render, 'fundo', l, 1, 1, nC, T1_COR_SECAO);
      l++;
    }
    var contagem = t1Bloco_(v, colunas);
    t1Por_(render, l, 1, v.rotulo);
    t1Marcar_(render, 'negrito', l, 1, 1, 1, true);
    blocos.forEach(function(b) {
      t1Marcar_(render, 'fundo', l, b.colQui, 1, 2, T1_COR_VALOR_P);
      t1Marcar_(render, 'formato', l, b.colQui, 1, 2, '0.000');
      if (!T1_CALCULAR_P) return;
      var p = t1ValoresP_(contagem, b.grupos);
      t1Por_(render, l, b.colQui, t1NumeroP_(p.qui));
      t1Por_(render, l, b.colFisher, t1NumeroP_(p.fisher));
    });
    l++;
    contagem.categorias.forEach(function(cat) {
      t1Por_(render, l, 1, '    ' + cat.rotulo);
      if (cat.rotulo === T1_ROTULO_SEM_INFO) t1Marcar_(render, 'italico', l, 1, 1, 1, true);
      colunas.forEach(function(col) {
        var n = cat.contagem[col.chave] || 0;
        var base = cat.rotulo === T1_ROTULO_SEM_INFO ? col.regs.length :
        contagem.validos[col.chave];
        t1Por_(render, l, col.col, n);
        t1Por_(render, l, col.col + 1, base ? t1ArredondarPar_(n * 100 / base, 1) : '');
      });
      l++;
    });
  });
  var rFim = l - 1;
  colunas.forEach(function(col) {
    t1Marcar_(render, 'formato', rSub + 1, col.col, rFim - rSub, 1, '0');
    t1Marcar_(render, 'formato', rSub + 1, col.col + 1, rFim - rSub, 1, '0.0');
  });
  t1Marcar_(render, 'formato', r1, 1, rFim - r1 + 1, 1, '@');
  t1Marcar_(render, 'alinhaH', r1, 1, rFim - r1 + 1, 1, 'left');
  t1Marcar_(render, 'alinhaH', rGrupo, 2, rFim - rGrupo + 1, nC - 1, 'center');
  blocos.forEach(function(b) {
    t1Marcar_(render, 'alinhaH', rEstudo, b.c1, 1, 1, b.corte.periodo ? 'left' : 'center');
  });
  t1Quadro_(render, { r1: r1, rGrupo: rGrupo, rFimCabecalho: rSub, rFim: rFim, caixas: caixas });
  return { fim: rFim, partes: partes, linhasCabecalho: rSub - r1 + 1, blocos: faixas };
}
function t1Bloco_(v, colunas) {
  var ordem = t1Fixas_(v);
  var contagem = {};
  var extras = [];
  var validos = {};
  var temSemInfo = false;
  ordem.forEach(function(k) { contagem[k] = {}; });
  colunas.forEach(function(col) {
    validos[col.chave] = 0;
    col.regs.forEach(function(r) {
      var cat = r[v.campo];
      var k = (cat === null || cat === undefined || cat === '') ? T1_ROTULO_SEM_INFO :
      String(cat);
      if (k === T1_ROTULO_SEM_INFO) temSemInfo = true;
      else validos[col.chave]++;
      if (!contagem[k]) {
        contagem[k] = {};
        if (k !== T1_ROTULO_SEM_INFO) extras.push(k);
      }
      contagem[k][col.chave] = (contagem[k][col.chave] || 0) + 1;
    });
  });
  extras.sort(t1CompararChaves_);
  ordem = ordem.concat(extras);
  if (temSemInfo) ordem.push(T1_ROTULO_SEM_INFO);
  return {
    validos: validos,
    categorias: ordem.map(function(k) { return { rotulo: k, contagem: contagem[k] || {} }; })
  };
}
function t1ValoresP_(bloco, grupos) {
  var tabela = [];
  bloco.categorias.forEach(function(cat) {
    if (cat.rotulo === T1_ROTULO_SEM_INFO) return;
    var linha = grupos.map(function(col) { return cat.contagem[col.chave] || 0; });
    if (linha.some(function(x) { return x > 0; })) tabela.push(linha);
  });
  if (tabela.length < 2) return { qui: null, fisher: null };
  var manter = [];
  for (var j = 0; j < grupos.length; j++) {
    if (tabela.some(function(l) { return l[j] > 0; })) manter.push(j);
  }
  if (manter.length < 2) return { qui: null, fisher: null };
  tabela = tabela.map(function(l) { return manter.map(function(j) { return l[j]; }); });
  return { qui: t1QuiQuadrado_(tabela), fisher: t1Fisher_(tabela) };
}
function t1Margens_(t) {
  var somaL = [];
  var somaC = [];
  var n = 0;
  var i, j;
  for (j = 0; j < t[0].length; j++) somaC.push(0);
  for (i = 0; i < t.length; i++) {
    var s = 0;
    for (j = 0; j < t[i].length; j++) {
      s += t[i][j];
      somaC[j] += t[i][j];
    }
    somaL.push(s);
    n += s;
  }
  return { somaL: somaL, somaC: somaC, n: n };
}
function t1QuiQuadrado_(t) {
  var m = t1Margens_(t);
  if (!m.n) return null;
  var x2 = 0;
  for (var i = 0; i < t.length; i++) {
    for (var j = 0; j < t[i].length; j++) {
      var e = m.somaL[i] * m.somaC[j] / m.n;
      if (e <= 0) return null;
      var d = t[i][j] - e;
      x2 += d * d / e;
    }
  }
  return t1PQui_(x2, (t.length - 1) * (t[0].length - 1));
}
function t1LogFatoriais_(n) {
  var lf = [0];
  for (var i = 1; i <= n; i++) lf.push(lf[i - 1] + Math.log(i));
  return lf;
}
function t1Fisher_(t) {
  if (t[0].length === 2) {
    var exato = t1FisherRx2_(t);
    if (exato !== null) return exato;
  }
  var n = 0;
  t.forEach(function(l) { l.forEach(function(x) { n += x; }); });
  return t1FisherSimulado_(t, Math.max(2000, Math.min(T1_SIMULACOES, Math.floor(2e7 /
  Math.max(1, n)))));
}
function t1FisherRx2_(t) {
  var nLin = t.length;
  var rs = [];
  var c1 = 0;
  var n = 0;
  t.forEach(function(l) {
    rs.push(l[0] + l[1]);
    c1 += l[0];
    n += l[0] + l[1];
  });
  var lf = t1LogFatoriais_(n);
  function lcomb(a, b) {
    return lf[a] - lf[b] - lf[a - b];
  }
  var lbin = rs.map(function(r) {
    var a = [];
    for (var x = 0; x <= r; x++) a.push(lcomb(r, x));
    return a;
  });
  var lTotal = lcomb(n, c1);
  var obs = 0;
  t.forEach(function(l, i) { obs += lbin[i][l[0]]; });
  var limite = obs + 1e-7;
  var resto = [];
  resto[nLin] = 0;
  for (var i = nLin - 1; i >= 0; i--) resto[i] = resto[i + 1] + rs[i];
  var maximo = [];
  var minimo = [];
  maximo[nLin] = [0];
  minimo[nLin] = [0];
  for (i = nLin - 1; i >= 0; i--) {
    var teto = Math.min(c1, resto[i]);
    maximo[i] = [];
    minimo[i] = [];
    for (var rem = 0; rem <= teto; rem++) {
      var mx = -Infinity;
      var mn = Infinity;
      var de = Math.max(0, rem - resto[i + 1]);
      var ate = Math.min(rs[i], rem);
      for (var x = de; x <= ate; x++) {
        var a = lbin[i][x] + maximo[i + 1][rem - x];
        var b = lbin[i][x] + minimo[i + 1][rem - x];
        if (a > mx) mx = a;
        if (b < mn) mn = b;
      }
      maximo[i].push(mx);
      minimo[i].push(mn);
    }
  }
  var soma = 0;
  var visitados = 0;
  var estourou = false;
  function percorrer(k, restante, lp) {
    if (estourou) return;
    if (++visitados > T1_FISHER_MAX_TABELAS) {
      estourou = true;
      return;
    }
    if (lp + maximo[k][restante] <= limite) {
      soma += Math.exp(lp + lcomb(resto[k], restante) - lTotal);
      return;
    }
    if (lp + minimo[k][restante] > limite) return;
    var de = Math.max(0, restante - resto[k + 1]);
    var ate = Math.min(rs[k], restante);
    for (var x = de; x <= ate; x++) percorrer(k + 1, restante - x, lp + lbin[k][x]);
  }
  percorrer(0, c1, 0);
  if (estourou) return null;
  return Math.min(1, soma);
}
function t1FisherSimulado_(t, repeticoes) {
  var m = t1Margens_(t);
  var lf = t1LogFatoriais_(m.n);
  var nL = m.somaL.length;
  var nC = m.somaC.length;
  var observado = 0;
  var i, j;
  for (i = 0; i < nL; i++) {
    for (j = 0; j < nC; j++) observado += lf[t[i][j]];
  }
  var rotulos = [];
  for (i = 0; i < nL; i++) {
    for (var k = 0; k < m.somaL[i]; k++) rotulos.push(i);
  }
  var colunaDe = [];
  for (j = 0; j < nC; j++) {
    for (var q = 0; q < m.somaC[j]; q++) colunaDe.push(j);
  }
  var sim = [];
  for (i = 0; i < nL * nC; i++) sim.push(0);
  var sorteio = t1Aleatorio_(t1Semente_(t));
  var total = rotulos.length;
  var extremos = 0;
  var piso = Math.floor;
  for (var r = 0; r < repeticoes; r++) {
    for (i = total - 1; i > 0; i--) {
      var alvo = piso(sorteio() * (i + 1));
      var tmp = rotulos[i];
      rotulos[i] = rotulos[alvo];
      rotulos[alvo] = tmp;
    }
    for (i = 0; i < sim.length; i++) sim[i] = 0;
    for (i = 0; i < total; i++) sim[rotulos[i] * nC + colunaDe[i]]++;
    var estatistica = 0;
    for (i = 0; i < sim.length; i++) estatistica += lf[sim[i]];
    if (estatistica >= observado - 1e-7) extremos++;
  }
  return (1 + extremos) / (repeticoes + 1);
}
function t1Semente_(t) {
  var h = 2166136261;
  var texto = JSON.stringify(t);
  for (var i = 0; i < texto.length; i++) {
    h ^= texto.charCodeAt(i);
    h = Math.imul(h, 16777619) >>> 0;
  }
  return h >>> 0;
}
function t1Aleatorio_(semente) {
  var a = semente >>> 0;
  var imul = Math.imul;
  return function() {
    a = (a + 0x6D2B79F5) >>> 0;
    var t = a;
    t = imul(t ^ (t >>> 15), t | 1) >>> 0;
    t = (t ^ (t + imul(t ^ (t >>> 7), t | 61))) >>> 0;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function t1LogGama_(x) {
  var c = [76.18009172947146, -86.50532032941677, 24.01409824083091,
    -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  var y = x;
  var tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  var ser = 1.000000000190015;
  for (var j = 0; j < 6; j++) ser += c[j] / ++y;
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}
function t1GamaQ_(a, x) {
  if (x <= 0) return 1;
  var lg = t1LogGama_(a);
  if (x < a + 1) {
    var ap = a;
    var soma = 1 / a;
    var termo = soma;
    for (var n = 1; n <= 1000; n++) {
      ap++;
      termo *= x / ap;
      soma += termo;
      if (Math.abs(termo) < Math.abs(soma) * 1e-15) break;
    }
    return 1 - soma * Math.exp(-x + a * Math.log(x) - lg);
  }
  var minimo = 1e-300;
  var b = x + 1 - a;
  var c = 1 / minimo;
  var d = 1 / b;
  var h = d;
  for (var i = 1; i <= 1000; i++) {
    var an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < minimo) d = minimo;
    c = b + an / c;
    if (Math.abs(c) < minimo) c = minimo;
    d = 1 / d;
    var del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-15) break;
  }
  return Math.exp(-x + a * Math.log(x) - lg) * h;
}
function t1PQui_(x2, gl) {
  if (!(gl > 0) || !isFinite(x2) || x2 < 0) return null;
  return Math.max(0, Math.min(1, t1GamaQ_(gl / 2, x2 / 2)));
}


/* ======================================================================
 * MÓDULO 4 DE 5 — TABELA2.gs
 * ====================================================================== */

var T2_ABA_SAIDA = 'TABELA 4';
var T2_VERSAO = 'public-v1';
var T2_REAGIR_PROPRIA_ABA = false;
var T2_ROTULO_TABELA = 'Tabela 4 –';
var T2_TITULO = 'Taxas de sucesso e de sobrevida das restaurações, segundo o estudo e o grupo, {horizontes}.';
var T2_ROTULO_VALOR_P = 'Valor p';
var T2_HORIZONTES_FIXOS = [0.5, 1];
var T2_MAX_ANOS = 10;
var T2_Z = 1.959963984540054;
var T2_LARGURA_ROTULO = 330;
var T2_LARGURA_VALOR = 130;
var T2_LARGURA_P = 72;
var T2_LARGURA_ESPACO = 24;
var T2_CAMPOS = {
  codpac: ['codepac', 'codpac'],
  coderest: ['coderest'],
  estudo: ['estudo'],
  grupo: ['grupo'],
  dente: ['n⁰ dente', 'Nº dente', 'N° dente'],
  data0: ['Data 0'],
  dataSucesso: ['data sucesso'],
  tempoSucesso: ['tempo sucesso'],
  sucesso: ['sucesso'],
  dataSobrevida: ['data sobrevida'],
  tempoSobrevida: ['tempo sobrevida'],
  sobrevida: ['sobrevida']
};
var T2_DESFECHOS = [
  { chave: 'sucesso', secao: 'Sucesso', taxa: 'Taxa de sucesso (Kaplan-Meier), % (IC 95%)' },
  { chave: 'sobrevida', secao: 'Sobrevida', taxa: 'Taxa de sobrevida (Kaplan-Meier), % (IC 95%)' }
];
var T2_ABAS_FONTE = ['BANCO DE ANÁLISE', 'BANCO DE ANALISE', 'Folha1',
'Planilha1', 'Sheet1'];
var T2_LINHA_CABECALHO = 1;
var T2_LINHA_DADOS = 2;
var T2_HORA_ROTINA = null;
var T2_MINUTOS_VIGIA = 30;
var T2_ATRASO_SEGUNDOS = 90;
var T2_JANELA_IGNORAR_MS = 45000;
var T2_LINHAS_ENTRE_TABELAS = 3;
var T2_MARCADORES_VAZIOS = ['.', '..', 'na', 'n/a', 'n.a.', 'nd', 'n/d', '-',
'--', 's/i', '?',
  'sem informacao', 'sem informação', 'null', 'undefined'];
var T2_ROTULO_SEM_INFO = 'Sem informação';
var T2_ROTULO_TOTAL = 'Total';
var T2_ROTULO_SEM_GRUPO = 'grupo não informado\nsem sorteio';
var T2_ROTULO_LINHA_ESTUDO = 'estudo';
var T2_ROTULO_LINHA_GRUPO = 'grupo';
var T2_GRUPOS_POR_ESTUDO = [
  { estudo: '0', grupos: ['0', '1'] },
  { estudo: '1', grupos: ['2', '3'] }
];
var T2_ESTUDO_POR_TEXTO = { 'estudoa': '0', 'estudob': '1', 'studya': '0', 'studyb': '1', 'a': '0', 'b': '1' };
var T2_ROTULOS = {
  estudo: { '0': 'Estudo A', '1': 'Estudo B' },
  grupo: {
    '0': 'Grupo 1',
    '1': 'Grupo 2',
    '2': 'Grupo 3',
    '3': 'Grupo 4'
  }
};
var T2_NOME_ESTUDO = { '0': 'A', '1': 'B' };
var T2_COR_SECAO = '#efefef';
var T2_COR_VALOR_P = '#fff2cc';
var T2_CORES_DESTAQUE = [
  { chave: '0', nome: 'A', linha: '#cfe2f3', dados: '#6fa8dc' },
  { chave: '1', nome: 'B', linha: '#fff2cc', dados: '#ffe599' },
  { chave: 'sem', nome: 'sem sorteio', linha: '#d9d2e9', dados: '#b4a7d6' }
];
var T2_ACENTOS = {
  'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
  'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
  'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
  'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
  'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
  'ç': 'c', 'ñ': 'n', 'ý': 'y', 'ÿ': 'y'
};
var T2_CRITERIOS_FDI = {
  fraturaretencao: 'Fratura/perda de retenção',
  adaptacaomarginal: 'Desadaptação marginal',
  contatoproximal: 'Ausência de contato proximal',
  formacontorno: 'Forma/contorno inadequado',
  oclusaodesgaste: 'Oclusão/desgaste',
  cariemargem: 'Cárie na margem',
  defeitotecduromargem: 'Defeito de tecido duro na margem',
  hipersensibposop: 'Hipersensibilidade pós-operatória',
  brilhotextura: 'Brilho/textura',
  coloracaomarginal: 'Coloração marginal',
  estabcor: 'Estabilidade de cor',
  satisfacao: 'Insatisfação do paciente',
  avalradiografica: 'Alteração radiográfica'
};
var T2_ESCORE_FDI_FALHA = 4;
var T2_PAINEL_ABA = 'CONSORT';
var T2_PAINEL_POS = 'CS_ENTRADAS_POS';
var T2_PAINEL_VAL = 'CS_ENTRADAS_VAL';
var T2_MODO_AUTO = 'auto';
var T2_MODO_PERIODO = 'por período';
var T2_ESPERA_TRAVA_PAINEL_MS = 120000;
var T2_VANTAGEM_PAINEL_MS = 2000;
var T2_CODIGOS_FALHA_SUCESSO = [1, 2];
var T2_CODIGOS_FALHA_SOBREVIDA = [2];
var T2_DIAS_ANO = 365.25;
function GERAR_TABELA2() {
  return t2Executar_(true);
}
function tabela2Rotina() {
  return t2Executar_(false);
}
function tabela2Reconstruir() {
  t2LimparAgendamento_();
  return t2Executar_(false);
}
function tabela2Vigia() {
  var atual = t2Impressao_();
  var visto = t2Props_().getProperty('T2_IMPRESSAO') || '';
  if (atual && atual === visto && !cmPrecisaEscrever_(T2_ABA_SAIDA)) return;
  t2Executar_(false);
}
function tabela2OnChange(e) {
  if (t2PainelMudou_()) return t2AtualizarPeloPainel_();
  if (t2OrigemMudanca_(e) !== 'fonte') return;
  if (t2GeracaoRecente_()) return;
  t2Agendar_();
}
function configurarTabela2(silencioso) {
  var ss = t2SS_();
  desligarTabela2(true);
  t2Props_().deleteProperty('T2_DESLIGADO');
  var hora = t2HoraRotina_();
  var falhas = [];
  var pedidos = [
    { rotulo: 'reconstrução diária', criar: function() {
      ScriptApp.newTrigger('tabela2Rotina').timeBased().everyDays(1).atHour(hora).create();
    } },
    { rotulo: 'reação a edições', criar: function() {
      ScriptApp.newTrigger('tabela2OnChange').forSpreadsheet(ss).onChange().create()
      ;
    } },
    { rotulo: 'vigia', criar: function() {
      ScriptApp.newTrigger('tabela2Vigia').timeBased().everyMinutes(T2_MINUTOS_VIGIA
      ).create();
    } }
  ];
  pedidos.forEach(function(p) {
    try {
      p.criar();
    } catch (err) {
      falhas.push(p.rotulo + ' (' + err.message + ')');
    }
  });
  var msg = 'Gatilhos da aba ' + T2_ABA_SAIDA + ' instalados: vigia a cada ' +
  T2_MINUTOS_VIGIA +
    ' min, reação a edições no banco e no painel do CONSORT e reconstrução diária às ' + hora + 'h.';
  if (falhas.length) {
    msg = 'ATENÇÃO: não foi possível instalar ' + falhas.join('; ') + '. ' +
      'O projeto aceita no máximo 20 gatilhos por conta; confira em Acionadores. ' + msg;
  }
  if (silencioso !== true) t2Avisar_(msg);
  return msg;
}
function desligarTabela2(interno) {
  var alvos = ['tabela2Rotina', 'tabela2OnChange', 'tabela2Reconstruir',
  'tabela2Vigia'];
  var removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (alvos.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  t2Props_().deleteProperty('T2_AGENDADO');
  if (interno !== true) t2Props_().setProperty('T2_DESLIGADO', '1');
  return removidos + ' gatilho(s) da aba ' + T2_ABA_SAIDA + ' removido(s).';
}
function t2GruposEstudos_(sh, blocos) {
  if (!sh || !blocos || !blocos.estudos) return;
  var nC = sh.getMaxColumns();
  var quer = [0];
  for (var c = 1; c <= nC; c++) quer.push(0);
  blocos.estudos.forEach(function(b) {
    for (var k = b.c1; k <= Math.min(b.c2, nC); k++) quer[k] = 1;
  });
  var atual = [0];
  var igual = true;
  for (var c2 = 1; c2 <= nC; c2++) {
    atual.push(sh.getColumnGroupDepth(c2));
    if (atual[c2] !== quer[c2]) igual = false;
  }
  if (igual) return;
  var maximo = Math.max.apply(null, atual);
  for (var d = maximo; d >= 1; d--) {
    t2Faixas_(nC, function(k) { return atual[k] >= d; }).forEach(function(f) {
      sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(-1);
    });
  }
  sh.showColumns(1, nC);
  t2Faixas_(nC, function(k) { return quer[k] === 1; }).forEach(function(f) {
    sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(1);
  });
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
}
function t2Faixas_(n, teste) {
  var faixas = [];
  var ini = 0;
  for (var c = 1; c <= n + 1; c++) {
    var ok = c <= n && teste(c);
    if (ok && !ini) ini = c;
    if (!ok && ini) {
      faixas.push([ini, c - 1]);
      ini = 0;
    }
  }
  return faixas;
}
function t2Destaque_(registros) {
  var mapa = {};
  registros.forEach(function(r) {
    if (!r.pacote) return;
    var chave = r.grupo ? r.estudo : 'sem';
    if (!t2CorDestaque_(chave)) return;
    if (!mapa[chave]) mapa[chave] = {};
    if (!mapa[chave][r.pacote.linha]) mapa[chave][r.pacote.linha] = {};
    var celulas = mapa[chave][r.pacote.linha];
    for (var c in r.pacote.usadas) {
      if (r.pacote.usadas.hasOwnProperty(c)) celulas[c] = true;
    }
  });
  return mapa;
}
function t2CorDestaque_(chave) {
  for (var i = 0; i < T2_CORES_DESTAQUE.length; i++) {
    if (T2_CORES_DESTAQUE[i].chave === chave) return T2_CORES_DESTAQUE[i];
  }
  return null;
}
function t2RegrasDestaque_(mapa, nC) {
  var dados = [];
  var linhas = [];
  T2_CORES_DESTAQUE.forEach(function(cor) {
    var porLinha = mapa[cor.chave];
    var numeros = porLinha ? Object.keys(porLinha).map(Number) : [];
    if (!numeros.length) return;
    var ultima = Math.max.apply(null, numeros);
    var porTrecho = {};
    numeros.forEach(function(l) {
      t2Faixas_(nC, function(c) { return !!porLinha[l][c - 1];
}).forEach(function(t) {
        var k = t[0] + ':' + t[1];
        if (!porTrecho[k]) porTrecho[k] = {};
        porTrecho[k][l] = true;
      });
    });
    var a1Dados = [];
    Object.keys(porTrecho).forEach(function(k) {
      var c = k.split(':').map(Number);
      t2Faixas_(ultima, function(l) { return !!porTrecho[k][l];
}).forEach(function(t) {
        a1Dados.push(t2A1_(t[0], c[0], t[1] - t[0] + 1, c[1] - c[0] + 1));
      });
    });
    var a1Linhas = t2Faixas_(ultima, function(l) { return !!porLinha[l];
}).map(function(t) {
      return t2A1_(t[0], 1, t[1] - t[0] + 1, nC);
    });
    if (a1Dados.length) dados.push({ formula: t2FormulaDestaque_(cor.nome,
      'dados lidos'), cor: cor.dados, a1: a1Dados.sort() });
    linhas.push({ formula: t2FormulaDestaque_(cor.nome, 'linha analisada'),
      cor: cor.linha, a1: a1Linhas.sort() });
  });
  return { dados: dados, linhas: linhas };
}
function t2FormulaDestaque_(categoria, tipo) {
  return '=N("RRT · ' + T2_ABA_SAIDA + ' · ' + categoria + ' · ' + tipo +
  '")=0';
}
function t2MarcaDestaque_(regra) {
  try {
    var cond = regra.getBooleanCondition();
    if (!cond) return null;
    var m = String((cond.getCriteriaValues() || [])[0] || '').replace(/\s+/g,
    ' ')
      .match(/^\s*=\s*N\(\s*"RRT\s*·\s*(.*?)\s*·\s*(.*?)\s*·\s*(dados lidos|linha analisada)"\s*\)\s*=\s*0\s*$/i);
    return m ? { dono: m[1].trim().toUpperCase(), dados: m[3].toLowerCase()
      === 'dados lidos' } : null;
  } catch (err) {
    return null;
  }
}
function t2AplicarDestaque_(sh, mapa) {
  if (!sh || !mapa) return '';
  try {
    var desejo = t2RegrasDestaque_(mapa, sh.getLastColumn());
    var outras = [];
    var dados = [];
    var linhas = [];
    var minhas = [];
    var fase = 0;
    var emOrdem = true;
    sh.getConditionalFormatRules().forEach(function(regra) {
      var marca = t2MarcaDestaque_(regra);
      var f = !marca ? 0 : (marca.dados ? 1 : 2);
      if (f < fase) emOrdem = false;
      fase = Math.max(fase, f);
      if (!marca) outras.push(regra);
      else if (marca.dono === T2_ABA_SAIDA.toUpperCase()) minhas.push(regra);
      else if (marca.dados) dados.push(regra);
      else linhas.push(regra);
    });
    if (emOrdem && t2MesmasRegras_(minhas,
    desejo.dados.concat(desejo.linhas))) return '';
    var nova = function(d) {
      return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(d.formula).setBackground(d.cor)
        .setRanges(sh.getRangeList(d.a1).getRanges()).build();
    };
    sh.setConditionalFormatRules(outras.concat(dados, desejo.dados.map(nova),
    linhas, desejo.linhas.map(nova)));
    return '';
  } catch (err) {
    return ' ATENÇÃO: não foi possível colorir as linhas analisadas na aba ' +
    sh.getName() + ' (' + err.message + ').';
  }
}
function t2MesmasRegras_(atuais, desejadas) {
  if (atuais.length !== desejadas.length) return false;
  var chave = function(formula, cor, a1) {
    return JSON.stringify([String(formula || '').replace(/\s+/g,
    '').toLowerCase(), String(cor || '').toLowerCase(),
      a1.map(function(x) { return String(x).toUpperCase(); }).sort()]);
  };
  var a = atuais.map(function(regra) {
    var cond = regra.getBooleanCondition();
    return chave((cond.getCriteriaValues() || [])[0], cond.getBackground(),
      regra.getRanges().map(function(r) { return r.getA1Notation(); }));
  }).sort();
  var d = desejadas.map(function(x) { return chave(x.formula, x.cor, x.a1);
}).sort();
  return JSON.stringify(a) === JSON.stringify(d);
}
function t2SS_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Nenhuma planilha ativa.');
  return ss;
}
function t2Props_() {
  return PropertiesService.getDocumentProperties();
}
function t2HoraRotina_() {
  if (T2_HORA_ROTINA !== null && T2_HORA_ROTINA !== undefined) return Number(T2_HORA_ROTINA);
  if (typeof HORA_ROTINA_BANCO !== 'undefined' && HORA_ROTINA_BANCO !== null)
  {
    return (Number(HORA_ROTINA_BANCO) + 1) % 24;
  }
  return 7;
}
function t2Agendar_() {
  var props = t2Props_();
  var pendente = Number(props.getProperty('T2_AGENDADO') || 0);
  var agora = Date.now();
  if (pendente && agora - pendente < T2_ATRASO_SEGUNDOS * 1000 * 3) return;
  props.setProperty('T2_AGENDADO', String(agora));
  ScriptApp.newTrigger('tabela2Reconstruir').timeBased().after(T2_ATRASO_SEGUNDOS * 1000).create();
}
function t2LimparAgendamento_() {
  t2Props_().deleteProperty('T2_AGENDADO');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tabela2Reconstruir')
    ScriptApp.deleteTrigger(t);
  });
}
function t2GeracaoRecente_() {
  var ultima = Number(t2Props_().getProperty('T2_ULTIMA_GERACAO') || 0);
  return ultima && Date.now() - ultima < T2_JANELA_IGNORAR_MS;
}
function t2OrigemMudanca_(e) {
  var ignorados = ['FORMAT', 'OTHER'];
  if (e && e.changeType && ignorados.indexOf(String(e.changeType)) >= 0)
  return '';
  var ativa = null;
  try {
    ativa = (e && e.source ? e.source : t2SS_()).getActiveSheet();
  } catch (err) {
    return 'fonte';
  }
  if (!ativa) return 'fonte';
  var nome = ativa.getName();
  if (nome === T2_PAINEL_ABA) return 'painel';
  if (nome === T2_ABA_SAIDA) return T2_REAGIR_PROPRIA_ABA ? 'painel' : '';
  return T2_ABAS_FONTE.indexOf(nome) >= 0 ? 'fonte' : '';
}
function t2Impressao_() {
  var ss = t2SS_();
  try {
    return 'drv:' +
    DriveApp.getFileById(ss.getId()).getLastUpdated().getTime();
  } catch (err) {}
  try {
    var sh = t2AbaFonte_(ss);
    return 'dim:' + sh.getLastRow() + 'x' + sh.getLastColumn();
  } catch (err2) {}
  return '';
}
function t2Executar_(interativo, doPainel) {
  var trava = cmTrava_();
  if (!trava.tryLock(doPainel ? T2_ESPERA_TRAVA_PAINEL_MS : 30000)) {
    if (doPainel) {
      t2Agendar_();
      return 'Outra execução em andamento; reconstrução da aba ' +
      T2_ABA_SAIDA + ' agendada.';
    }
    var ocupado = 'Outra execução da aba ' + T2_ABA_SAIDA + ' está em andamento.';
    if (interativo) t2Avisar_(ocupado);
    return ocupado;
  }
  try {
    var ss = t2SS_();
    var fonte = t2AbaFonte_(ss);
    var entrada = t2LerFonte_(fonte);
    var render = t2Montar_(entrada, ss);
    var assinatura = t2Assinatura_(render);
    var props = t2Props_();
    var existente = ss.getSheetByName(T2_ABA_SAIDA);
    if (existente && !props.getProperty('T2_ASSINATURA') &&
    existente.getLastRow() > 0) {
      var alheia = 'A aba "' + T2_ABA_SAIDA + '" já existe com conteúdo que não foi gerado por este script; nada foi alterado. ' +
        'Renomeie essa aba ou troque T2_ABA_SAIDA no início do arquivo e rode de novo.';
      if (interativo) throw new Error(alheia);
      return alheia;
    }
    if (!cmPrecisaEscrever_(T2_ABA_SAIDA) && !interativo && existente && assinatura ===
    props.getProperty('T2_ASSINATURA')) {
      props.setProperty('T2_IMPRESSAO', t2Impressao_());
      if (render.chaveCorte) props.setProperty('T2_CORTE_USADO',
      render.chaveCorte);
      if (render.chavePainel !== undefined)
      props.setProperty('T2_PAINEL_LIDO', render.chavePainel);
      t2GruposEstudos_(existente, render.blocos);
      return 'Aba ' + T2_ABA_SAIDA + ' já está em dia.' +
      t2AplicarDestaque_(fonte, render.destaque);
    }
    t2Escrever_(ss, render);
    if (render.depoisDeEscrever) render.depoisDeEscrever(ss);
    t2GruposEstudos_(ss.getSheetByName(T2_ABA_SAIDA), render.blocos);
    var avisoCores = t2AplicarDestaque_(fonte, render.destaque);
    props.setProperty('T2_ASSINATURA', assinatura);
    props.setProperty('T2_ULTIMA_GERACAO', String(Date.now()));
    props.setProperty('T2_IMPRESSAO', t2Impressao_());
    if (render.chaveCorte) props.setProperty('T2_CORTE_USADO',
    render.chaveCorte);
    if (render.chavePainel !== undefined) props.setProperty('T2_PAINEL_LIDO',
    render.chavePainel);
    cmConcluir_(T2_ABA_SAIDA);
    var msg = 'Aba ' + T2_ABA_SAIDA + ' atualizada. ' + (render.resumo || '')
    + ' Fonte: ' + fonte.getName() + '.';
    if (render.faltando && render.faltando.length) {
      msg += ' ATENÇÃO: colunas não localizadas na origem: ' +
      render.faltando.join(', ') + '.';
    }
    msg += avisoCores;
    if (interativo) {
      msg += t2GarantirGatilhos_();
      t2Avisar_(msg);
    }
    return msg;
  } finally {
    trava.releaseLock();
  }
}
function t2GarantirGatilhos_() {
  try {
    if (t2Props_().getProperty('T2_DESLIGADO')) {
      return ' Atualização automática desta aba desligada por desligarTabela2; para religar, rode configurarTabela2.';
    }
    var instalado = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'tabela2OnChange';
    });
    return instalado ? '' : ' ' + configurarTabela2(true);
  } catch (err) {
    return ' Não foi possível instalar a atualização automática (' +
    err.message + '); rode configurarTabela2.';
  }
}
function t2Avisar_(texto) {
  try {
    SpreadsheetApp.getActive().toast(texto, T2_ABA_SAIDA, 8);
  } catch (err) {}
}
function t2AbaFonte_(ss) {
  for (var i = 0; i < T2_ABAS_FONTE.length; i++) {
    var sh = ss.getSheetByName(T2_ABAS_FONTE[i]);
    if (sh) return sh;
  }
  throw new Error('Nenhuma aba de origem encontrada. Procurei por: ' +
  T2_ABAS_FONTE.join(', ') + '.');
}
function t2LerFonte_(sh) {
  var nCols = sh.getLastColumn();
  var nLinhas = sh.getLastRow() - T2_LINHA_DADOS + 1;
  if (nCols < 2) throw new Error('A aba "' + sh.getName() + '" não tem colunas suficientes.');
  var cabecalho = sh.getRange(T2_LINHA_CABECALHO, 1, 1,
  nCols).getDisplayValues()[0];
  var dados = nLinhas > 0 ? sh.getRange(T2_LINHA_DADOS, 1, nLinhas,
  nCols).getValues() : [];
  return { cabecalho: cabecalho, dados: dados };
}
function t2Assinatura_(render) {
  var texto = JSON.stringify([T2_VERSAO, render.matriz, render.formatos,
  render.larguras, render.alturas]);
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, texto,
    Utilities.Charset.UTF_8);
    var hex = '';
    for (var i = 0; i < bytes.length; i++) {
      var b = (bytes[i] + 256) % 256;
      hex += (b < 16 ? '0' : '') + b.toString(16);
    }
    return hex;
  } catch (err) {
    return 'len' + texto.length;
  }
}
function t2Chave_(v) {
  var s = String(v === null || v === undefined ? '' : v).toLowerCase();
  var saida = '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    saida += T2_ACENTOS[c] || c;
  }
  return saida.replace(/[^a-z0-9]+/g, '');
}
function t2Vazio_(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return isNaN(v.getTime());
  var s = String(v).trim();
  if (!s) return true;
  return T2_MARCADORES_VAZIOS.indexOf(s.toLowerCase()) >= 0;
}
function t2Texto_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/'
    + ('0' + v.getDate()).slice(-2);
  }
  return String(v).trim();
}
function t2Numero_(v) {
  if (t2Vazio_(v)) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  if (v instanceof Date) return null;
  var s = String(v).trim().replace(/\s/g, '').replace(/^r\$/i, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function t2Data_(v) {
  if (t2Vazio_(v)) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var iso = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1,
  Number(iso[3]));
  var br = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
  return null;
}
function t2Indices_(cabecalho, campos) {
  var porChave = {};
  for (var c = 0; c < cabecalho.length; c++) {
    var k = t2Chave_(cabecalho[c]);
    if (!k) continue;
    if (!porChave[k]) porChave[k] = [];
    porChave[k].push(c);
  }
  var mapa = {};
  for (var campo in campos) {
    if (!campos.hasOwnProperty(campo)) continue;
    var achados = [];
    campos[campo].forEach(function(alias) {
      (porChave[t2Chave_(alias)] || []).forEach(function(idx) {
        if (achados.indexOf(idx) < 0) achados.push(idx);
      });
    });
    mapa[campo] = achados;
  }
  return mapa;
}
function t2Faltando_(idx, campos, exigidos) {
  var faltando = [];
  exigidos.forEach(function(alternativas) {
    var achou = alternativas.some(function(c) { return idx[c] &&
      idx[c].length; });
    if (!achou) faltando.push(campos[alternativas[0]][0]);
  });
  return faltando;
}
function t2ValorBruto_(linha, indices) {
  if (!indices || !indices.length) return '';
  for (var i = 0; i < indices.length; i++) {
    var v = linha[indices[i]];
    if (!t2Vazio_(v)) {
      if (linha.usadas) linha.usadas[indices[i]] = true;
      return v;
    }
  }
  return '';
}
function t2Usar_(linha, colunas) {
  if (!linha.usadas) return;
  colunas.forEach(function(c) {
    if (typeof c === 'number' && c >= 0 && !t2Vazio_(linha[c]))
    linha.usadas[c] = true;
  });
}
function t2CodigoGrupo_(v) {
  if (t2Vazio_(v)) return '';
  var n = t2Numero_(v);
  return n !== null ? String(n) : t2Texto_(v);
}
function t2EstudoEfetivo_(estudo, grupo) {
  if (!t2Vazio_(estudo)) {
    var n = t2Numero_(estudo);
    if (n !== null) return String(n);
    var k = t2Chave_(estudo);
    if (T2_ESTUDO_POR_TEXTO.hasOwnProperty(k)) return T2_ESTUDO_POR_TEXTO[k];
    return t2Texto_(estudo);
  }
  for (var i = 0; i < T2_GRUPOS_POR_ESTUDO.length; i++) {
    if (T2_GRUPOS_POR_ESTUDO[i].grupos.indexOf(grupo) >= 0) return
    T2_GRUPOS_POR_ESTUDO[i].estudo;
  }
  return '';
}
function t2ChavePaciente_(codpac, numeroLinha) {
  if (t2Vazio_(codpac)) return '#linha' + numeroLinha;
  var n = t2Numero_(codpac);
  if (n !== null) return 'p' + n;
  return 'p' + (t2Chave_(codpac) || t2Texto_(codpac));
}
function t2Estudos_(registros) {
  var lista = [];
  T2_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var grupos = cfg.grupos.slice();
    var extras = [];
    var semGrupo = 0;
    registros.forEach(function(r) {
      if (r.estudo !== cfg.estudo) return;
      if (!r.grupo) { semGrupo++; return; }
      if (grupos.indexOf(r.grupo) < 0 && extras.indexOf(r.grupo) < 0)
      extras.push(r.grupo);
    });
    extras.sort(t2CompararChaves_);
    lista.push({ estudo: cfg.estudo, grupos: grupos.concat(extras),
      temSemGrupo: semGrupo > 0 });
  });
  return lista;
}
function t2RotuloGrupo_(g) {
  if (T2_ROTULOS.grupo.hasOwnProperty(g)) return T2_ROTULOS.grupo[g];
  return 'Grupo ' + g;
}
function t2RotuloEstudo_(e) {
  return T2_ROTULOS.estudo.hasOwnProperty(e) ? T2_ROTULOS.estudo[e] : 'estudo ' + e;
}
function t2NomeEstudo_(e) {
  return T2_NOME_ESTUDO.hasOwnProperty(e) ? T2_NOME_ESTUDO[e] : String(e);
}
function t2ArredondarPar_(v, casas) {
  var f = Math.pow(10, casas);
  var x = v * f;
  var piso = Math.floor(x);
  var frac = x - piso;
  var r;
  if (Math.abs(frac - 0.5) < 1e-9) r = (piso % 2 === 0) ? piso : piso + 1;
  else r = Math.round(x);
  return r / f;
}
function t2Fmt_(v, casas) {
  if (v === null || v === undefined || isNaN(v)) return '';
  return t2ArredondarPar_(v, casas).toFixed(casas).replace('.', ',');
}
function t2NumeroP_(p) {
  if (p === null || p === undefined || isNaN(p)) return '';
  if (p < 0.001) return '< 0,001';
  return t2ArredondarPar_(p, 3);
}
function t2Plural_(n, singular, plural) {
  return n === 1 ? singular : plural;
}
function t2Contar_(obj) {
  var n = 0;
  for (var k in obj) {
    if (obj.hasOwnProperty(k)) n++;
  }
  return n;
}
function t2LinhaVazia_(n) {
  var linha = [];
  for (var i = 0; i < n; i++) linha.push('');
  return linha;
}
function t2CompararChaves_(a, b) {
  var na = t2Numero_(a);
  var nb = t2Numero_(b);
  if (na !== null && nb !== null) return na - nb;
  if (na !== null) return -1;
  if (nb !== null) return 1;
  return a < b ? -1 : (a > b ? 1 : 0);
}
function t2Letra_(c) {
  var s = '';
  while (c > 0) {
    var m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}
function t2A1_(l, c, nl, nc) {
  var a = t2Letra_(c) + l;
  nl = nl || 1;
  nc = nc || 1;
  if (nl === 1 && nc === 1) return a;
  return a + ':' + t2Letra_(c + nc - 1) + (l + nl - 1);
}
function t2NovoRender_() {
  return { matriz: [], nColunas: 1, formatos: [], larguras: {}, alturas: {},
    faltando: [], resumo: '' };
}
function t2Por_(render, l, c, valor) {
  while (render.matriz.length < l) render.matriz.push([]);
  var linha = render.matriz[l - 1];
  while (linha.length < c) linha.push('');
  linha[c - 1] = valor;
}
function t2Fechar_(render) {
  var nC = render.nColunas;
  render.matriz.forEach(function(linha) {
    if (linha.length > nC) nC = linha.length;
  });
  render.nColunas = nC;
  render.matriz.forEach(function(linha) {
    while (linha.length < nC) linha.push('');
  });
  return render;
}
function t2Marcar_(render, tipo, l, c, nl, nc, valor) {
  render.formatos.push({ tipo: tipo, l: l, c: c, nl: nl || 1, nc: nc || 1,
    valor: valor });
}
function t2Largura_(render, col, px) {
  if (!render.larguras[col] || render.larguras[col] < px) render.larguras[col]
  = px;
}
function t2Altura_(render, linha, px) {
  if (!render.alturas[linha] || render.alturas[linha] < px)
  render.alturas[linha] = px;
}
function t2ContarLinhas_(texto, larguraPx, pxPorCaractere) {
  var porLinha = Math.max(1, Math.floor((larguraPx - 8) / (pxPorCaractere ||
  6.6)));
  var total = 0;
  String(texto).split('\n').forEach(function(paragrafo) {
    var palavras = paragrafo.split(/\s+/).filter(function(p) { return p; });
    if (!palavras.length) { total++; return; }
    var atual = 0;
    var linhas = 1;
    palavras.forEach(function(p) {
      var tam = p.length;
      if (atual === 0) {
        atual = tam;
      } else if (atual + 1 + tam <= porLinha) {
        atual += 1 + tam;
      } else {
        linhas++;
        atual = tam;
      }
      while (atual > porLinha) {
        linhas++;
        atual -= porLinha;
      }
    });
    total += linhas;
  });
  return total;
}
function t2AlturaPara_(nLinhas) {
  return nLinhas <= 1 ? 21 : 21 + 15 * (nLinhas - 1);
}
function t2Escrever_(ss, r) {
  var nomeAba = r.aba || T2_ABA_SAIDA;
  r = cmPreparar_(ss, nomeAba, r);
  var sh = ss.getSheetByName(nomeAba);
  if (!sh) sh = ss.insertSheet(nomeAba);
  var preservar = r.preservar || [];
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  if (preservar.length) sh.clearFormats();
  else sh.clear();
  sh.clearNotes();
  if (r.validacoes) sh.getRange(1, 1, sh.getMaxRows(),
  sh.getMaxColumns()).clearDataValidations();
  if (!r.matriz.length) r.matriz.push(t2LinhaVazia_(Math.max(1, r.nColunas)));
  t2Fechar_(r);
  var nL = r.matriz.length;
  var nC = r.nColunas;
  if (sh.getMaxColumns() < nC) sh.insertColumnsAfter(sh.getMaxColumns(), nC -
  sh.getMaxColumns());
  if (sh.getMaxRows() < nL) sh.insertRowsAfter(sh.getMaxRows(), nL -
  sh.getMaxRows());
  var tudo = sh.getRange(1, 1, nL, nC);
  t2AplicarGrupos_(sh, r.formatos, function(f) { return f.tipo === 'formato';
});
  cmFormatos_(sh, r);
  if (preservar.length) {
    t2Retangulos_(nL, nC, preservar).forEach(function(q) {
      sh.getRange(q.l, q.c, q.nl, q.nc).setValues(r.matriz.slice(q.l - 1, q.l
      - 1 + q.nl).map(function(linha) {
        return linha.slice(q.c - 1, q.c - 1 + q.nc);
      }));
    });
  } else {
    tudo.setValues(r.matriz);
  }
  tudo.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(false);
  r.formatos.forEach(function(f) {
    if (f.tipo === 'mesclar') sh.getRange(f.l, f.c, f.nl, f.nc).merge();
  });
  t2AplicarGrupos_(sh, r.formatos, function(f) {
    return f.tipo !== 'formato' && f.tipo !== 'mesclar' && f.tipo !== 'rico';
  });
  var negrito = null;
  r.formatos.forEach(function(f) {
    if (f.tipo !== 'rico') return;
    if (!negrito) negrito =
    SpreadsheetApp.newTextStyle().setBold(true).build();
    var construtor = SpreadsheetApp.newRichTextValue().setText(f.valor.texto);
    f.valor.negritos.forEach(function(par) {
      if (par[1] > par[0]) construtor.setTextStyle(par[0], par[1], negrito);
    });
    sh.getRange(f.l, f.c).setRichTextValue(construtor.build());
  });
  (r.validacoes || []).forEach(function(v) {
    sh.getRange(v.l, v.c).setDataValidation(v.regra);
  });
  t2Medidas_(nC, r.larguras, 62, function(ini, qtd, px) {
    sh.setColumnWidths(ini, qtd, px); });
  t2Medidas_(nL, r.alturas, 21, function(ini, qtd, px) { sh.setRowHeights(ini,
    qtd, px); });
  if (r.congelarColunas) sh.setFrozenColumns(r.congelarColunas);
  if (r.congelarLinhas) sh.setFrozenRows(r.congelarLinhas);
  if (sh.getMaxRows() > nL) sh.deleteRows(nL + 1, sh.getMaxRows() - nL);
  if (sh.getMaxColumns() > nC) sh.deleteColumns(nC + 1, sh.getMaxColumns() -
  nC);
  SpreadsheetApp.flush();
}
function t2Retangulos_(nL, nC, preservar) {
  var furos = {};
  preservar.forEach(function(p) { furos[p.l + ',' + p.c] = true; });
  var segmentos = [];
  for (var c = 1; c <= nC; c++) {
    var lista = [];
    var ini = 0;
    for (var l = 1; l <= nL + 1; l++) {
      var livre = l <= nL && !furos[l + ',' + c];
      if (livre && !ini) ini = l;
      if (!livre && ini) {
        lista.push([ini, l - 1]);
        ini = 0;
      }
    }
    segmentos.push(JSON.stringify(lista));
  }
  var retangulos = [];
  var c0 = 1;
  for (var k = 2; k <= nC + 1; k++) {
    if (k <= nC && segmentos[k - 1] === segmentos[c0 - 1]) continue;
    JSON.parse(segmentos[c0 - 1]).forEach(function(seg) {
      retangulos.push({ l: seg[0], c: c0, nl: seg[1] - seg[0] + 1, nc: k - c0
});
    });
    c0 = k;
  }
  return retangulos;
}
function t2Medidas_(n, mapa, padrao, aplicar) {
  var ini = 1;
  var atual = mapa[1] || padrao;
  for (var i = 2; i <= n + 1; i++) {
    var v = i <= n ? (mapa[i] || padrao) : null;
    if (v !== atual) {
      aplicar(ini, i - ini, atual);
      ini = i;
      atual = v;
    }
  }
}
function t2AplicarGrupos_(sh, formatos, filtro) {
  var grupos = {};
  var ordem = [];
  formatos.forEach(function(f) {
    if (!filtro(f)) return;
    var chave = f.tipo + '|' + JSON.stringify(f.valor);
    if (!grupos[chave]) {
      grupos[chave] = { f: f, a1: [] };
      ordem.push(chave);
    }
    grupos[chave].a1.push(t2A1_(f.l, f.c, f.nl, f.nc));
  });
  ordem.forEach(function(k) {
    var g = grupos[k];
    for (var i = 0; i < g.a1.length; i += 300) {
      t2AplicarLista_(sh.getRangeList(g.a1.slice(i, i + 300)), g.f);
    }
  });
}
function t2AplicarLista_(lista, f) {
  var v = f.valor;
  if (f.tipo === 'negrito') lista.setFontWeight(v ? 'bold' : 'normal');
  else if (f.tipo === 'italico') lista.setFontStyle(v ? 'italic' : 'normal');
  else if (f.tipo === 'fundo') lista.setBackground(v);
  else if (f.tipo === 'formato') lista.setNumberFormat(v);
  else if (f.tipo === 'alinhaH') lista.setHorizontalAlignment(v);
  else if (f.tipo === 'alinhaV') lista.setVerticalAlignment(v);
  else if (f.tipo === 'quebra') lista.setWrap(v);
  else if (f.tipo === 'tamanho') lista.setFontSize(v);
  else if (f.tipo === 'rotacao') lista.setTextRotation(v);
  else if (f.tipo === 'corTexto') lista.setFontColor(v);
  else if (f.tipo === 'borda') {
    if (v.estilo === 'tracejada') {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null, '#000000',
      SpreadsheetApp.BorderStyle.DASHED);
    } else {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null);
    }
  }
}
function t2Caixa_(render, l, c, nl, nc, estilo) {
  t2Marcar_(render, 'borda', l, c, nl, nc, { t: true, l: true, b: true, r:
    true, estilo: estilo || 'solida' });
}
function t2LinhaBaixo_(render, l, c, nc) {
  t2Marcar_(render, 'borda', l, c, 1, nc, { t: null, l: null, b: true, r:
    null, estilo: 'solida' });
}
function t2Laterais_(render, l, c, nl) {
  t2Marcar_(render, 'borda', l, c, nl, 1, { t: null, l: true, b: null, r:
    true, estilo: 'solida' });
}
function t2Quadro_(render, q) {
  t2Laterais_(render, q.r1, 1, q.rFim - q.r1 + 1);
  [q.r1, q.rGrupo, q.rFimCabecalho, q.rFim].forEach(function(l) {
    t2LinhaBaixo_(render, l, 1, 1);
  });
  q.caixas.forEach(function(cx) {
    var nc = cx.c2 - cx.c1 + 1;
    t2Caixa_(render, q.r1 + 1, cx.c1, q.rFim - q.r1, nc);
    t2LinhaBaixo_(render, q.rGrupo, cx.c1, nc);
    t2LinhaBaixo_(render, q.rFimCabecalho, cx.c1, nc);
  });
}
function t2CabecalhoGrupo_(render, l, c, texto, larguraPx) {
  t2Por_(render, l, c, texto);
  t2Marcar_(render, 'quebra', l, c, 1, 1, true);
  t2Altura_(render, l, t2AlturaPara_(t2ContarLinhas_(texto, larguraPx, 7)));
}
function t2Blocos_(cabecalho) {
  var inicios = [];
  for (var c = 0; c < cabecalho.length; c++) {
    var m = t2Chave_(cabecalho[c]).match(/^data(\d+)$/);
    if (m) inicios.push({ indice: Number(m[1]), col: c });
  }
  var blocos = [];
  inicios.forEach(function(b, i) {
    var fim = i + 1 < inicios.length ? inicios[i + 1].col : cabecalho.length;
    var bloco = { indice: b.indice, data: b.col, avaliacao: -1, causa: -1,
      criterios: [] };
    for (var c = b.col + 1; c < fim; c++) {
      var raiz = t2Chave_(cabecalho[c]).replace(/\d+$/, '');
      if (raiz === 'avaliacaofinal') {
        if (bloco.avaliacao < 0) bloco.avaliacao = c;
      } else if (raiz === 'causafalha') {
        if (bloco.causa < 0) bloco.causa = c;
      } else if (T2_CRITERIOS_FDI.hasOwnProperty(raiz)) {
        bloco.criterios.push({ col: c, rotulo: T2_CRITERIOS_FDI[raiz] });
      }
    }
    blocos.push(bloco);
  });
  return blocos;
}
function t2LerBlocos_(linha, blocos) {
  return blocos.map(function(b) {
    var aval = b.avaliacao >= 0 ? t2Numero_(linha[b.avaliacao]) : null;
    var causa = b.causa >= 0 ? t2Numero_(linha[b.causa]) : null;
    var criticos = [];
    var colunasCriticas = [];
    b.criterios.forEach(function(cr) {
      var escore = t2Numero_(linha[cr.col]);
      if (escore === null || escore < T2_ESCORE_FDI_FALHA) return;
      colunasCriticas.push(cr.col);
      if (criticos.indexOf(cr.rotulo) < 0) criticos.push(cr.rotulo);
    });
    return { data: t2Data_(linha[b.data]), avaliacao: aval, causa: causa,
      criticos: criticos, colunasCriticas: colunasCriticas };
  });
}
function t2DiasEntre_(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}
function t2LerPainel_(ss) {
  var valores = {};
  try {
    valores = JSON.parse(t2Props_().getProperty(T2_PAINEL_VAL) || '{}') || {};
  } catch (err) {
    valores = {};
  }
  var lidos = t2LerCelulasPainel_(ss);
  if (lidos) {
    for (var est in lidos) {
      if (!lidos.hasOwnProperty(est)) continue;
      if (!valores[est]) valores[est] = {};
      for (var campo in lidos[est]) {
        if (lidos[est].hasOwnProperty(campo)) valores[est][campo] =
        lidos[est][campo];
      }
    }
  }
  return { valores: valores, lidos: lidos };
}
function t2LerCelulasPainel_(ss) {
  var pos = null;
  try {
    pos = JSON.parse(t2Props_().getProperty(T2_PAINEL_POS) || 'null');
  } catch (err) {
    pos = null;
  }
  var sh = ss ? ss.getSheetByName(T2_PAINEL_ABA) : null;
  if (!sh || !pos) return null;
  var celulas = [];
  for (var est in pos) {
    if (!pos.hasOwnProperty(est)) continue;
    for (var campo in pos[est]) {
      if (!pos[est].hasOwnProperty(campo)) continue;
      var m = String(pos[est][campo]).match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      var c = 0;
      for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) -
      64);
      celulas.push({ est: est, campo: campo, l: Number(m[2]), c: c });
    }
  }
  if (!celulas.length) return null;
  var l1 = Infinity, c1 = Infinity, l2 = 0, c2 = 0;
  celulas.forEach(function(x) {
    l1 = Math.min(l1, x.l);
    c1 = Math.min(c1, x.c);
    l2 = Math.max(l2, x.l);
    c2 = Math.max(c2, x.c);
  });
  if (l2 > sh.getMaxRows() || c2 > sh.getMaxColumns()) return null;
  var grade = sh.getRange(l1, c1, l2 - l1 + 1, c2 - c1 +
  1).getDisplayValues();
  var lidos = {};
  celulas.forEach(function(x) {
    var v = grade[x.l - l1][x.c - c1];
    if (!lidos[x.est]) lidos[x.est] = {};
    lidos[x.est][x.campo] = String(v === null || v === undefined ? '' :
    v).trim();
  });
  return lidos;
}
function t2PainelMudou_() {
  try {
    var ss = t2SS_();
    var props = t2Props_();
    if (T2_ABA_SAIDA === T2_PAINEL_ABA) {
      var lidos = t2LerCelulasPainel_(ss);
      return !!lidos && JSON.stringify(lidos) !==
      (props.getProperty('T2_PAINEL_LIDO') || '');
    }
    return t2ChaveCortes_(t2Cortes_(ss, new Date())) !==
    (props.getProperty('T2_CORTE_USADO') || '');
  } catch (err) {
    return false;
  }
}
function t2AtualizarPeloPainel_() {
  if (T2_ABA_SAIDA !== T2_PAINEL_ABA) Utilities.sleep(T2_VANTAGEM_PAINEL_MS);
  return t2Executar_(false, true);
}
function t2ChaveCortes_(cortes) {
  return T2_GRUPOS_POR_ESTUDO.map(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (!c.periodo) return cfg.estudo + ':auto';
    return cfg.estudo + ':' + (c.coderest !== null ? c.coderest : '') + '#' +
    (c.inicio ? t2DataBr_(c.inicio) : '') +
      '-' + (c.fim ? t2DataBr_(c.fim) : '');
  }).join('|');
}
function t2FaltaCoderest_(idx, cortes) {
  if (idx.coderest && idx.coderest.length) return false;
  return T2_GRUPOS_POR_ESTUDO.some(function(cfg) { return cortes.estudos[cfg.estudo].coderest !== null; });
}
function t2ModoCorte_(v) {
  var k = t2Chave_(v);
  return !k || k.indexOf('auto') === 0 ? T2_MODO_AUTO : T2_MODO_PERIODO;
}
function t2DataEstrita_(v) {
  if (t2Vazio_(v)) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v).trim();
  var br = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(\D|$)/);
  var iso = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(\D|$)/);
  var a, m, d;
  if (br) {
    a = Number(br[3]);
    m = Number(br[2]);
    d = Number(br[1]);
  } else if (iso) {
    a = Number(iso[1]);
    m = Number(iso[2]);
    d = Number(iso[3]);
  } else {
    return null;
  }
  var data = new Date(a, m - 1, d);
  if (data.getFullYear() !== a || data.getMonth() !== m - 1 || data.getDate()
  !== d) return null;
  return data;
}
function t2Dia_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}
function t2Corte_(v, hoje) {
  var corte = { periodo: false, inicio: null, fim: null, coderest: null,
    aviso: '',
    referencia: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate())
};
  if (!v || t2ModoCorte_(v.corte) !== T2_MODO_PERIODO) return corte;
  var inicio = t2DataEstrita_(v.inicio);
  var fim = t2DataEstrita_(v.fim);
  var coderest = t2Vazio_(v.coderest) ? null : t2Numero_(v.coderest);
  var invalidos = [];
  if (!t2Vazio_(v.inicio) && !inicio) invalidos.push('a data inicial');
  if (!t2Vazio_(v.fim) && !fim) invalidos.push('a data final');
  if (!t2Vazio_(v.coderest) && (coderest === null || coderest < 1 ||
  Math.round(coderest) !== coderest)) {
    invalidos.push('o coderest');
  }
  if (invalidos.length) {
    corte.aviso = 'não entendi ' + invalidos.join(' nem ') + '; corte ignorado, valendo auto.';
    return corte;
  }
  if (!inicio && !fim && coderest === null) {
    corte.aviso = 'preencha as datas e/ou o coderest; sem nenhum deles, vale auto.';
    return corte;
  }
  if (inicio && fim && inicio.getTime() > fim.getTime()) {
    corte.aviso = 'a data inicial é depois da data final; corte ignorado, valendo auto.';
    return corte;
  }
  corte.periodo = true;
  corte.inicio = inicio;
  corte.fim = fim;
  corte.coderest = coderest;
  if (fim && fim.getTime() < corte.referencia.getTime()) corte.referencia =
  fim;
  return corte;
}
function t2Cortes_(ss, hoje) {
  var painel = t2LerPainel_(ss);
  var cortes = { estudos: {}, padrao: t2Corte_(null, hoje), valores:
    painel.valores, lidos: painel.lidos };
  T2_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    cortes.estudos[cfg.estudo] = t2Corte_(painel.valores[cfg.estudo], hoje);
  });
  return cortes;
}
function t2CorteDe_(cortes, estudo) {
  return cortes.estudos.hasOwnProperty(estudo) ? cortes.estudos[estudo] :
  cortes.padrao;
}
function t2DentroDoCorte_(corte, data0, coderest) {
  if (!corte.periodo) return true;
  if (corte.coderest !== null && (coderest === null || coderest >
  corte.coderest)) return false;
  if (!data0) return true;
  var t = t2Dia_(data0);
  if (corte.inicio && t < corte.inicio.getTime()) return false;
  if (corte.fim && corte.coderest === null && t > corte.fim.getTime()) return false;
  return true;
}
function t2Data0NoCorte_(corte, data0) {
  if (!data0 || !corte.periodo || !corte.fim) return data0;
  return t2Dia_(data0) > corte.fim.getTime() ? null : data0;
}
function t2DataBr_(d) {
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() +
  1)).slice(-2) + '/' + d.getFullYear();
}
function t2TextoCorte_(corte) {
  if (!corte.periodo) return '';
  var partes = [];
  if (corte.coderest !== null) partes.push('coderest até ' + corte.coderest);
  if (corte.inicio && corte.fim) partes.push(t2DataBr_(corte.inicio) + ' a ' +
  t2DataBr_(corte.fim));
  else if (corte.inicio) partes.push('a partir de ' +
  t2DataBr_(corte.inicio));
  else if (corte.fim) partes.push('até ' + t2DataBr_(corte.fim));
  return partes.join('; ');
}
function t2RotuloComCorte_(rotulo, corte) {
  var texto = t2TextoCorte_(corte);
  return texto ? rotulo + ' (corte: ' + texto + ')' : rotulo;
}
function t2AvisosCorte_(cortes) {
  var avisos = [];
  T2_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (c && c.aviso) avisos.push('corte do estudo ' +
    t2NomeEstudo_(cfg.estudo) + ': ' + c.aviso);
  });
  return avisos;
}
function t2Desfecho_(linha, idx, blocos, data0, corte, lidos) {
  if (data0 && !t2Data0NoCorte_(corte, data0)) {
    return { sucesso: t2DesfechoVazio_(), sobrevida: t2DesfechoVazio_() };
  }
  var usadas = linha.usadas;
  var doBanco = {};
  var d = {};
  [['sucesso', idx.dataSucesso, idx.tempoSucesso, idx.sucesso,
  idx.causaSucesso, T2_CODIGOS_FALHA_SUCESSO],
    ['sobrevida', idx.dataSobrevida, idx.tempoSobrevida, idx.sobrevida,
  idx.causaSobrevida, T2_CODIGOS_FALHA_SOBREVIDA]
  ].forEach(function(p) {
    linha.usadas = doBanco[p[0]] = {};
    d[p[0]] = t2DesfechoBanco_(linha, p[1], p[2], p[3], p[4], p[5], data0);
  });
  linha.usadas = usadas;
  if (corte.periodo && corte.fim && data0) {
    var fim = corte.fim.getTime();
    var refeito = null;
    ['sucesso', 'sobrevida'].forEach(function(k) {
      if (!d[k].data || t2Dia_(d[k].data) <= fim) return;
      if (!refeito) refeito = t2DesfechoBlocos_(lidos || t2LerBlocos_(linha,
      blocos), blocos, data0, fim);
      d[k] = refeito[k];
      doBanco[k] = null;
      refeito.usados[k].forEach(function(i) { t2Usar_(linha, [blocos[i].data,
        blocos[i].avaliacao]); });
      if (refeito.travas[k] >= 0) t2Usar_(linha,
      [blocos[refeito.travas[k]].causa]);
    });
  }
  ['sucesso', 'sobrevida'].forEach(function(k) {
    if (!usadas || !doBanco[k]) return;
    for (var c in doBanco[k]) {
      if (doBanco[k].hasOwnProperty(c)) usadas[c] = true;
    }
  });
  return d;
}
function t2DesfechoVazio_() {
  return { evento: false, codigo: null, causa: null, data: null, dias: null,
    anos: null };
}
function t2DesfechoBanco_(linha, iData, iTempo, iCodigo, iCausa, falhas,
data0) {
  var codigo = t2Numero_(t2ValorBruto_(linha, iCodigo));
  var data = t2Data_(t2ValorBruto_(linha, iData));
  var dias = null;
  var anos = null;
  if (data0 && data) {
    var n = t2DiasEntre_(data0, data);
    if (n >= 0) {
      dias = n;
      anos = n / T2_DIAS_ANO;
    }
  } else {
    var t = t2Numero_(t2ValorBruto_(linha, iTempo));
    if (t !== null && t >= 0) {
      anos = t;
      if (data0) data = new Date(data0.getFullYear(), data0.getMonth(),
      data0.getDate() + Math.round(t * T2_DIAS_ANO));
    }
  }
  return { evento: codigo !== null && falhas.indexOf(codigo) >= 0, codigo:
    codigo,
    causa: t2Numero_(t2ValorBruto_(linha, iCausa)), data: data, dias: dias,
    anos: anos };
}
function t2DesfechoBlocos_(lidos, blocos, data0, fim) {
  var ordem = [];
  for (var k = 0; k < blocos.length; k++) ordem.push(k);
  ordem.sort(function(a, b) { return blocos[a].indice - blocos[b].indice || a
    - b; });
  var reg = { sucesso: null, sobrevida: null };
  var travado = { sucesso: false, sobrevida: false };
  var usados = { sucesso: [], sobrevida: [] };
  var travas = { sucesso: -1, sobrevida: -1 };
  var falhas = { sucesso: T2_CODIGOS_FALHA_SUCESSO, sobrevida:
    T2_CODIGOS_FALHA_SOBREVIDA };
  ordem.forEach(function(i) {
    var bl = lidos[i];
    if (!bl.data || t2Dia_(bl.data) > fim) return;
    if (blocos[i].indice !== 0 && bl.avaliacao === null) return;
    var r = { data: bl.data, codigo: bl.avaliacao, causa: bl.causa };
    ['sucesso', 'sobrevida'].forEach(function(k) {
      if (travado[k]) return;
      reg[k] = r;
      usados[k].push(i);
      if (r.codigo !== null && falhas[k].indexOf(r.codigo) >= 0) {
        travado[k] = true;
        travas[k] = i;
      }
    });
  });
  return {
    sucesso: t2DesfechoRegistro_(reg.sucesso, travado.sucesso, data0),
    sobrevida: t2DesfechoRegistro_(reg.sobrevida, travado.sobrevida, data0),
    usados: usados,
    travas: travas
  };
}
function t2DesfechoRegistro_(r, travado, data0) {
  var data = r ? r.data : data0;
  var n = t2DiasEntre_(data0, data);
  return { evento: travado, codigo: travado ? r.codigo : 0, causa: travado ?
    r.causa : 0, data: data,
    dias: n >= 0 ? n : null, anos: n >= 0 ? n / T2_DIAS_ANO : null };
}
function t2Montar_(entrada, ss) {
  var idx = t2Indices_(entrada.cabecalho, T2_CAMPOS);
  var blocos = t2Blocos_(entrada.cabecalho);
  var faltando = t2Faltando_(idx, T2_CAMPOS, [['grupo'], ['data0'],
  ['dataSucesso', 'tempoSucesso'],
    ['sucesso'], ['dataSobrevida', 'tempoSobrevida'], ['sobrevida']]);
  var cortes = t2Cortes_(ss, new Date());
  if (t2FaltaCoderest_(idx, cortes)) faltando.push('coderest (usado no corte)');
  var algumCorte = T2_GRUPOS_POR_ESTUDO.some(function(cfg) { return
    !!cortes.estudos[cfg.estudo].fim; });
  if (algumCorte && blocos.length < 2) faltando.push('Data 1 (blocos de rechamada, usados no corte)');
  var registros = t2Registros_(entrada.dados, idx, blocos, cortes);
  var render = t2NovoRender_();
  var res = t2Tabela_(render, t2Estudos_(registros), registros, 1, cortes);
  render.faltando = faltando;
  render.resumo = res.partes.concat(t2AvisosCorte_(cortes)).join('; ') + '.';
  render.chaveCorte = t2ChaveCortes_(cortes);
  render.congelarColunas = 1;
  render.congelarLinhas = res.linhasCabecalho;
  render.blocos = { rotulos: [1, 2], estudos: res.blocos };
  render.destaque = t2Destaque_(registros);
  return t2Fechar_(render);
}
function t2Registros_(dados, idx, blocos, cortes) {
  var registros = [];
  dados.forEach(function(linha, i) {
    linha.usadas = {};
    var codpac = t2ValorBruto_(linha, idx.codpac);
    var grupoBruto = t2ValorBruto_(linha, idx.grupo);
    var estudoBruto = t2ValorBruto_(linha, idx.estudo);
    var dente = t2ValorBruto_(linha, idx.dente);
    if (t2Vazio_(codpac) && t2Vazio_(grupoBruto) && t2Vazio_(dente) &&
    t2Vazio_(estudoBruto)) return;
    var grupo = t2CodigoGrupo_(grupoBruto);
    var estudo = t2EstudoEfetivo_(estudoBruto, grupo);
    var data0 = t2Data_(t2ValorBruto_(linha, idx.data0));
    var corte = t2CorteDe_(cortes, estudo);
    if (!t2DentroDoCorte_(corte, data0, t2Numero_(t2ValorBruto_(linha,
    idx.coderest)))) return;
    if (!t2Data0NoCorte_(corte, data0)) return;
    var d = t2Desfecho_(linha, idx, blocos, data0, corte);
    registros.push({
      estudo: estudo,
      grupo: grupo,
      tsucesso: d.sucesso.anos,
      esucesso: d.sucesso.evento ? 1 : 0,
      tsobrevida: d.sobrevida.anos,
      esobrevida: d.sobrevida.evento ? 1 : 0,
      pacote: { linha: i + T2_LINHA_DADOS, usadas: linha.usadas }
    });
  });
  return registros;
}
function t2Observacoes_(regs, chave) {
  var obs = [];
  regs.forEach(function(r) {
    var t = r['t' + chave];
    if (t === null || t === undefined) return;
    obs.push({ t: t, e: r['e' + chave] });
  });
  return obs;
}
function t2Horizontes_(regs) {
  var maximo = 0;
  regs.forEach(function(r) {
    if (r.tsucesso !== null && r.tsucesso > maximo) maximo = r.tsucesso;
    if (r.tsobrevida !== null && r.tsobrevida > maximo) maximo = r.tsobrevida;
  });
  var lista = T2_HORIZONTES_FIXOS.slice();
  for (var a = 2; a <= Math.min(T2_MAX_ANOS, Math.floor(maximo + 1e-9)); a++)
  lista.push(a);
  return lista;
}
function t2RotuloHorizonte_(h) {
  if (h < 1) return Math.round(h * 12) + ' meses';
  return h + (h === 1 ? ' ano' : ' anos');
}
function t2TextoHorizontes_(lista) {
  var itens = [];
  lista.forEach(function(h) {
    if (h < 1) itens.push(Math.round(h * 12) + ' meses');
  });
  var anos = lista.filter(function(h) { return h >= 1; });
  anos.forEach(function(a, i) {
    if (i < anos.length - 1) itens.push(String(a));
    else itens.push(a + (anos.length === 1 && a === 1 ? ' ano' : ' anos'));
  });
  if (!itens.length) return '';
  if (itens.length === 1) return 'aos ' + itens[0];
  return 'aos ' + itens.slice(0, -1).join(', ') + ' e ' + itens[itens.length -
  1];
}
function t2Tabela_(render, estudos, registros, r1, cortes) {
  var blocos = [];
  var colunas = [];
  var caixas = [];
  var analisadas = [];
  var faixas = [];
  var c = 3;
  estudos.forEach(function(est, i) {
    if (i > 0) {
      t2Largura_(render, c, T2_LARGURA_ESPACO);
      c++;
    }
    var comGrupo = registros.filter(function(r) { return r.estudo ===
      est.estudo && !!r.grupo; });
    analisadas = analisadas.concat(comGrupo);
    var bloco = { estudo: est.estudo, c1: c, grupos: [], total: est.estudo +
      '|__total',
      corte: t2CorteDe_(cortes, est.estudo) };
    est.grupos.forEach(function(g) {
      var col = { chave: est.estudo + '|' + g, rotulo: t2RotuloGrupo_(g), col:
        c,
        regs: comGrupo.filter(function(r) { return r.grupo === g; }) };
      bloco.grupos.push(col);
      colunas.push(col);
      c++;
    });
    colunas.push({ chave: bloco.total, rotulo: T2_ROTULO_TOTAL, col: c, regs:
      comGrupo });
    c++;
    bloco.colP = c;
    c++;
    caixas.push({ c1: bloco.c1, c2: bloco.colP });
    faixas.push({ chave: est.estudo, nome: t2NomeEstudo_(est.estudo), c1:
      bloco.c1, c2: bloco.colP });
    blocos.push(bloco);
  });
  var semGrupo = registros.filter(function(r) { return !r.grupo; });
  if (semGrupo.length) {
    t2Largura_(render, c, T2_LARGURA_ESPACO);
    c++;
    colunas.push({ chave: '__sem', rotulo: T2_ROTULO_SEM_GRUPO, col: c, regs:
      semGrupo });
    caixas.push({ c1: c, c2: c });
    faixas.push({ chave: 'sem', nome: 'Sem sorteio', c1: c, c2: c });
    c++;
  }
  var nC = c - 1;
  if (render.nColunas < nC) render.nColunas = nC;
  t2Largura_(render, 1, T2_LARGURA_ROTULO);
  t2Largura_(render, 2, T2_LARGURA_ESPACO);
  colunas.forEach(function(col) { t2Largura_(render, col.col,
    T2_LARGURA_VALOR); });
  blocos.forEach(function(b) { t2Largura_(render, b.colP, T2_LARGURA_P); });
  var horizontes = t2Horizontes_(analisadas);
  t2Por_(render, r1, 1, T2_ROTULO_TABELA);
  t2Por_(render, r1, 2, T2_TITULO.replace('{horizontes}',
  t2TextoHorizontes_(horizontes)));
  t2Marcar_(render, 'negrito', r1, 1, 1, nC, true);
  t2Marcar_(render, 'tamanho', r1, 1, 1, nC, 11);
  var rEstudo = r1 + 1;
  var rGrupo = r1 + 2;
  var rN = r1 + 3;
  t2Por_(render, rEstudo, 1, T2_ROTULO_LINHA_ESTUDO);
  t2Por_(render, rGrupo, 1, T2_ROTULO_LINHA_GRUPO);
  blocos.forEach(function(b) {
    t2Por_(render, rEstudo, b.c1, t2RotuloComCorte_(t2RotuloEstudo_(b.estudo),
    b.corte));
    t2Por_(render, rGrupo, b.colP, T2_ROTULO_VALOR_P);
  });
  colunas.forEach(function(col) {
    t2CabecalhoGrupo_(render, rGrupo, col.col, col.rotulo, T2_LARGURA_VALOR);
    t2Por_(render, rN, col.col, 'n = ' + col.regs.length);
  });
  t2Marcar_(render, 'negrito', rEstudo, 1, 3, nC, true);
  var l = rN + 1;
  var resumo = {};
  var linhasNumero = [];
  function rotuloNegrito(texto) {
    t2Por_(render, l, 1, texto);
    t2Marcar_(render, 'negrito', l, 1, 1, 1, true);
  }
  T2_DESFECHOS.forEach(function(d) {
    t2Por_(render, l, 1, d.secao);
    t2Marcar_(render, 'negrito', l, 1, 1, 1, true);
    t2Marcar_(render, 'fundo', l, 1, 1, nC, T2_COR_SECAO);
    l++;
    var porColuna = {};
    colunas.forEach(function(col) {
      var obs = t2Observacoes_(col.regs, d.chave);
      porColuna[col.chave] = { obs: obs, desc: t2Descritiva_(obs), km:
        t2Km_(obs) };
    });
    if (d.chave === 'sucesso') {
      blocos.forEach(function(b) {
        var tot = porColuna[b.total].obs;
        var nf = 0;
        tot.forEach(function(o) { nf += o.e; });
        resumo[b.estudo] = { comTempo: tot.length, falhas: nf };
      });
    }
    rotuloNegrito('Restaurações com tempo registrado, n');
    colunas.forEach(function(col) { t2Por_(render, l, col.col,
      porColuna[col.chave].obs.length); });
    linhasNumero.push(l);
    l++;
    rotuloNegrito('Tempo de acompanhamento em anos');
    l++;
    t2Por_(render, l, 1, '    Média (DP)');
    colunas.forEach(function(col) {
      var ds = porColuna[col.chave].desc;
      if (!ds) return;
      t2Por_(render, l, col.col, t2Fmt_(ds.media, 2) + ' (' + (ds.dp === null
      ? '–' : t2Fmt_(ds.dp, 2)) + ')');
    });
    l++;
    t2Por_(render, l, 1, '    Mediana (mín–máx)');
    colunas.forEach(function(col) {
      var ds = porColuna[col.chave].desc;
      if (!ds) return;
      t2Por_(render, l, col.col, t2Fmt_(ds.mediana, 2) + ' (' + t2Fmt_(ds.min,
      2) + '–' + t2Fmt_(ds.max, 2) + ')');
    });
    l++;
    rotuloNegrito('Falhas, n (%)');
    colunas.forEach(function(col) {
      var obs = porColuna[col.chave].obs;
      if (!obs.length) return;
      var nf = 0;
      obs.forEach(function(o) { nf += o.e; });
      t2Por_(render, l, col.col, nf + ' (' + t2Fmt_(nf * 100 / obs.length, 1)
      + ')');
    });
    l++;
    rotuloNegrito(d.taxa);
    blocos.forEach(function(b) {
      t2Marcar_(render, 'fundo', l, b.colP, 1, 1, T2_COR_VALOR_P);
      t2Marcar_(render, 'formato', l, b.colP, 1, 1, '0.000');
      t2Por_(render, l, b.colP,
      t2NumeroP_(t2LogRank_(b.grupos.map(function(col) {
        return porColuna[col.chave].obs;
      }))));
    });
    l++;
    horizontes.forEach(function(h) {
      t2Por_(render, l, 1, '    ' + t2RotuloHorizonte_(h));
      colunas.forEach(function(col) {
        var km = porColuna[col.chave].km;
        if (!km.n) return;
        t2Por_(render, l, col.col, t2TextoKm_(t2KmEm_(km, h)));
      });
      l++;
    });
  });
  var rFim = l - 1;
  colunas.forEach(function(col) {
    t2Marcar_(render, 'formato', rN + 1, col.col, rFim - rN, 1, '@');
  });
  linhasNumero.forEach(function(ln) {
    colunas.forEach(function(col) {
      t2Marcar_(render, 'formato', ln, col.col, 1, 1, '0');
    });
  });
  t2Marcar_(render, 'formato', r1, 1, rFim - r1 + 1, 1, '@');
  t2Marcar_(render, 'alinhaH', r1, 1, rFim - r1 + 1, 1, 'left');
  t2Marcar_(render, 'alinhaH', rGrupo, 2, rFim - rGrupo + 1, nC - 1,
  'center');
  blocos.forEach(function(b) {
    t2Marcar_(render, 'alinhaH', rEstudo, b.c1, 1, 1, b.corte.periodo ? 'left'
    : 'center');
  });
  t2Quadro_(render, { r1: r1, rGrupo: rGrupo, rFimCabecalho: rN, rFim: rFim,
    caixas: caixas });
  var partes = blocos.map(function(b) {
    var x = resumo[b.estudo] || { comTempo: 0, falhas: 0 };
    return t2RotuloComCorte_('estudo ' + t2NomeEstudo_(b.estudo) + ': ' +
    x.comTempo + ' com tempo registrado, ' +
      x.falhas + ' falha(s) de sucesso', b.corte);
  });
  if (semGrupo.length) partes.push('sem sorteio: ' + semGrupo.length);
  return { fim: rFim, partes: partes, linhasCabecalho: rN - r1 + 1, blocos:
    faixas };
}
function t2Descritiva_(obs) {
  var n = obs.length;
  if (!n) return null;
  var v = obs.map(function(o) { return o.t; }).sort(function(a, b) { return a
    - b; });
  var soma = 0;
  v.forEach(function(x) { soma += x; });
  var media = soma / n;
  var dp = null;
  if (n > 1) {
    var q = 0;
    v.forEach(function(x) { q += (x - media) * (x - media); });
    dp = Math.sqrt(q / (n - 1));
  }
  var mediana = n % 2 ? v[(n - 1) / 2] : (v[n / 2 - 1] + v[n / 2]) / 2;
  return { n: n, media: media, dp: dp, mediana: mediana, min: v[0], max: v[n -
    1] };
}
function t2TempoKm_(o) {
  return (o.e && o.t <= 0) ? 1 / T2_DIAS_ANO : o.t;
}
function t2Km_(obs) {
  var dados = obs.map(function(o) { return { t: t2TempoKm_(o), e: o.e }; });
  dados.sort(function(a, b) { return a.t - b.t || b.e - a.e; });
  var passos = [];
  var s = 1;
  var soma = 0;
  var i = 0;
  var n = dados.length;
  while (i < n) {
    var t = dados[i].t;
    var risco = n - i;
    var d = 0;
    var j = i;
    while (j < n && dados[j].t === t) {
      d += dados[j].e;
      j++;
    }
    if (d > 0) {
      s *= (1 - d / risco);
      soma = risco > d ? soma + d / (risco * (risco - d)) : Infinity;
      passos.push({ t: t, s: s, soma: soma });
    }
    i = j;
  }
  return { n: n, passos: passos, maxT: n ? dados[n - 1].t : null, sFinal: s };
}
function t2KmEm_(km, h) {
  if (!km.n) return null;
  if (km.maxT < h - 1e-12 && km.sFinal > 0) return null;
  var s = 1;
  var soma = 0;
  km.passos.forEach(function(p) {
    if (p.t <= h + 1e-12) {
      s = p.s;
      soma = p.soma;
    }
  });
  var ic = null;
  if (s > 0 && s < 1 && isFinite(soma)) {
    var w = T2_Z * Math.sqrt(soma) / Math.abs(Math.log(s));
    ic = [Math.pow(s, Math.exp(w)), Math.pow(s, Math.exp(-w))];
  }
  return { s: s, ic: ic };
}
function t2TextoKm_(r) {
  if (!r) return '–';
  var texto = t2Fmt_(r.s * 100, 1);
  if (r.ic) texto += ' (' + t2Fmt_(r.ic[0] * 100, 1) + '–' + t2Fmt_(r.ic[1] *
  100, 1) + ')';
  return texto;
}
function t2LogRank_(gruposObs) {
  var grupos = gruposObs.filter(function(obs) { return obs.length > 0; });
  var k = grupos.length;
  if (k < 2) return null;
  var todos = [];
  grupos.forEach(function(obs, g) {
    obs.forEach(function(o) { todos.push({ t: t2TempoKm_(o), e: o.e, g: g });
});
  });
  var tempos = [];
  todos.forEach(function(o) {
    if (o.e && tempos.indexOf(o.t) < 0) tempos.push(o.t);
  });
  if (!tempos.length) return null;
  tempos.sort(function(a, b) { return a - b; });
  var dim = k - 1;
  var O = [];
  var E = [];
  var V = [];
  for (var a = 0; a < k; a++) {
    O.push(0);
    E.push(0);
  }
  for (a = 0; a < dim; a++) {
    var linhaV = [];
    for (var b = 0; b < dim; b++) linhaV.push(0);
    V.push(linhaV);
  }
  tempos.forEach(function(t) {
    var nG = [];
    var dG = [];
    for (var g = 0; g < k; g++) {
      nG.push(0);
      dG.push(0);
    }
    todos.forEach(function(o) {
      if (o.t >= t) nG[o.g]++;
      if (o.t === t && o.e) dG[o.g]++;
    });
    var n = 0;
    var d = 0;
    for (g = 0; g < k; g++) {
      n += nG[g];
      d += dG[g];
    }
    if (!n || !d) return;
    for (g = 0; g < k; g++) {
      O[g] += dG[g];
      E[g] += d * nG[g] / n;
    }
    if (n < 2) return;
    var fator = d * (n - d) / (n - 1);
    for (var x = 0; x < dim; x++) {
      for (var y = 0; y < dim; y++) {
        V[x][y] += fator * (nG[x] / n) * ((x === y ? 1 : 0) - nG[y] / n);
      }
    }
  });
  var u = [];
  for (a = 0; a < dim; a++) u.push(O[a] - E[a]);
  var sol = t2Resolver_(V, u);
  if (!sol) return null;
  var x2 = 0;
  for (a = 0; a < dim; a++) x2 += u[a] * sol[a];
  return t2PQui_(x2, dim);
}
function t2Resolver_(A, b) {
  var n = b.length;
  var m = A.map(function(linha, i) { return linha.slice().concat([b[i]]); });
  for (var col = 0; col < n; col++) {
    var piv = col;
    for (var r = col + 1; r < n; r++) {
      if (Math.abs(m[r][col]) > Math.abs(m[piv][col])) piv = r;
    }
    if (Math.abs(m[piv][col]) < 1e-12) return null;
    var tmp = m[col];
    m[col] = m[piv];
    m[piv] = tmp;
    for (r = 0; r < n; r++) {
      if (r === col) continue;
      var f = m[r][col] / m[col][col];
      for (var c = col; c <= n; c++) m[r][c] -= f * m[col][c];
    }
  }
  return m.map(function(linha, i) { return linha[n] / linha[i]; });
}
function t2LogGama_(x) {
  var c = [76.18009172947146, -86.50532032941677, 24.01409824083091,
    -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  var y = x;
  var tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  var ser = 1.000000000190015;
  for (var j = 0; j < 6; j++) ser += c[j] / ++y;
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}
function t2GamaQ_(a, x) {
  if (x <= 0) return 1;
  var lg = t2LogGama_(a);
  if (x < a + 1) {
    var ap = a;
    var soma = 1 / a;
    var termo = soma;
    for (var n = 1; n <= 1000; n++) {
      ap++;
      termo *= x / ap;
      soma += termo;
      if (Math.abs(termo) < Math.abs(soma) * 1e-15) break;
    }
    return 1 - soma * Math.exp(-x + a * Math.log(x) - lg);
  }
  var minimo = 1e-300;
  var b = x + 1 - a;
  var c = 1 / minimo;
  var d = 1 / b;
  var h = d;
  for (var i = 1; i <= 1000; i++) {
    var an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < minimo) d = minimo;
    c = b + an / c;
    if (Math.abs(c) < minimo) c = minimo;
    d = 1 / d;
    var del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-15) break;
  }
  return Math.exp(-x + a * Math.log(x) - lg) * h;
}
function t2PQui_(x2, gl) {
  if (!(gl > 0) || !isFinite(x2) || x2 < 0) return null;
  return Math.max(0, Math.min(1, t2GamaQ_(gl / 2, x2 / 2)));
}


/* ======================================================================
 * MÓDULO 5 DE 5 — TABELA3.gs
 * ====================================================================== */

var T3_ABA_SAIDA = 'TABELA 5';
var T3_VERSAO = 'public-v1';
var T3_REAGIR_PROPRIA_ABA = false;
var T3_ROTULO_TABELA = 'Tabela 5 –';
var T3_TITULO = 'Descrição das falhas observadas e suas características, segundo o estudo.';
var T3_ROTULO_SEM_SORTEIO = 'grupo não informado / sem sorteio';
var T3_ROTULO_ESTUDO_COLUNA = 'Estudo';
var T3_ROTULO_ESTUDO_DESCONHECIDO = 'Não informado';
var T3_LARGURA_ESPACO = 24;
var T3_DIAS_MES = 365.25 / 12;
var T3_MOTIVOS = {
  '1': 'Descimentação',
  '2': 'Fratura da restauração',
  '3': 'Endodontia',
  '4': 'Periodontia',
  '5': 'Fratura radicular',
  '6': 'Fratura do pino',
  '7': 'Cárie',
  '8': 'Estética'
};
var T3_CODIGO_DESCIMENTACAO = 1;
var T3_ROTULO_SEM_MOTIVO = 'Não informado';
var T3_ROTULO_NENHUMA = 'Nenhuma falha registrada até o momento.';
var T3_COLUNAS_T3 = [
  { chave: 'dente', rotulo: 'Dente', largura: 52, alinha: 'left', formato: '@'
},
  { chave: 'sexo', rotulo: 'Sexo', largura: 46, alinha: 'left', formato: '@'
},
  { chave: 'grupo', rotulo: 'Grupo', largura: 140, alinha: 'left', formato:
  '@', quebra: true },
  { chave: 'faces', rotulo: 'Nº de faces restantes', largura: 78, alinha:
  'center', formato: '0' },
  { chave: 'vitalidade', rotulo: 'Vitalidade', largura: 84, alinha: 'center',
  formato: '@' },
  { chave: 'msobrevida', rotulo: 'Tempo de sobrevida (meses)', largura: 92,
  alinha: 'center', formato: '0.0' },
  { chave: 'msucesso', rotulo: 'Tempo de sucesso (meses)', largura: 92,
  alinha: 'center', formato: '0.0' },
  { chave: 'motivo', rotulo: 'Motivo da falha', largura: 200, alinha: 'left',
  formato: '@', quebra: true },
  { chave: 'reparo', rotulo: 'Reparo', largura: 150, alinha: 'left', formato:
  '@', quebra: true },
  { chave: 'situacao', rotulo: 'Situação atual', largura: 100, alinha: 'left',
  formato: '@' }
];
var T3_CAMPOS = {
  codpac: ['codepac', 'codpac'],
  coderest: ['coderest'],
  estudo: ['estudo'],
  grupo: ['grupo'],
  dente: ['n⁰ dente', 'Nº dente', 'N° dente'],
  sexo: ['Sexo'],
  faces: ['Nº faces restantes', 'Nº de faces restantes'],
  vitalidade: ['Vitalidade'],
  data0: ['Data 0'],
  dataSucesso: ['data sucesso'],
  tempoSucesso: ['tempo sucesso'],
  sucesso: ['sucesso'],
  causaSucesso: ['causa sucesso'],
  dataSobrevida: ['data sobrevida'],
  tempoSobrevida: ['tempo sobrevida'],
  sobrevida: ['sobrevida']
};
var T3_ABAS_FONTE = ['BANCO DE ANÁLISE', 'BANCO DE ANALISE', 'Folha1',
'Planilha1', 'Sheet1'];
var T3_LINHA_CABECALHO = 1;
var T3_LINHA_DADOS = 2;
var T3_HORA_ROTINA = null;
var T3_MINUTOS_VIGIA = 30;
var T3_ATRASO_SEGUNDOS = 90;
var T3_JANELA_IGNORAR_MS = 45000;
var T3_LINHAS_ENTRE_TABELAS = 3;
var T3_MARCADORES_VAZIOS = ['.', '..', 'na', 'n/a', 'n.a.', 'nd', 'n/d', '-',
'--', 's/i', '?',
  'sem informacao', 'sem informação', 'null', 'undefined'];
var T3_ROTULO_SEM_INFO = 'Sem informação';
var T3_ROTULO_TOTAL = 'Total';
var T3_ROTULO_SEM_GRUPO = 'grupo não informado\nsem sorteio';
var T3_ROTULO_LINHA_ESTUDO = 'estudo';
var T3_ROTULO_LINHA_GRUPO = 'grupo';
var T3_GRUPOS_POR_ESTUDO = [
  { estudo: '0', grupos: ['0', '1'] },
  { estudo: '1', grupos: ['2', '3'] }
];
var T3_ESTUDO_POR_TEXTO = { 'estudoa': '0', 'estudob': '1', 'studya': '0', 'studyb': '1', 'a': '0', 'b': '1' };
var T3_ROTULOS = {
  estudo: { '0': 'Estudo A', '1': 'Estudo B' },
  grupo: {
    '0': 'Grupo 1',
    '1': 'Grupo 2',
    '2': 'Grupo 3',
    '3': 'Grupo 4'
  }
};
var T3_NOME_ESTUDO = { '0': 'A', '1': 'B' };
var T3_COR_SECAO = '#efefef';
var T3_COR_VALOR_P = '#fff2cc';
var T3_CORES_DESTAQUE = [
  { chave: '0', nome: 'A', linha: '#cfe2f3', dados: '#6fa8dc' },
  { chave: '1', nome: 'B', linha: '#fff2cc', dados: '#ffe599' },
  { chave: 'sem', nome: 'sem sorteio', linha: '#d9d2e9', dados: '#b4a7d6' }
];
var T3_ACENTOS = {
  'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
  'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
  'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
  'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
  'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
  'ç': 'c', 'ñ': 'n', 'ý': 'y', 'ÿ': 'y'
};
var T3_CRITERIOS_FDI = {
  fraturaretencao: 'Fratura/perda de retenção',
  adaptacaomarginal: 'Desadaptação marginal',
  contatoproximal: 'Ausência de contato proximal',
  formacontorno: 'Forma/contorno inadequado',
  oclusaodesgaste: 'Oclusão/desgaste',
  cariemargem: 'Cárie na margem',
  defeitotecduromargem: 'Defeito de tecido duro na margem',
  hipersensibposop: 'Hipersensibilidade pós-operatória',
  brilhotextura: 'Brilho/textura',
  coloracaomarginal: 'Coloração marginal',
  estabcor: 'Estabilidade de cor',
  satisfacao: 'Insatisfação do paciente',
  avalradiografica: 'Alteração radiográfica'
};
var T3_ESCORE_FDI_FALHA = 4;
var T3_PAINEL_ABA = 'CONSORT';
var T3_PAINEL_POS = 'CS_ENTRADAS_POS';
var T3_PAINEL_VAL = 'CS_ENTRADAS_VAL';
var T3_MODO_AUTO = 'auto';
var T3_MODO_PERIODO = 'por período';
var T3_ESPERA_TRAVA_PAINEL_MS = 120000;
var T3_VANTAGEM_PAINEL_MS = 2000;
var T3_CODIGOS_FALHA_SUCESSO = [1, 2];
var T3_CODIGOS_FALHA_SOBREVIDA = [2];
var T3_DIAS_ANO = 365.25;
function GERAR_TABELA3() {
  return t3Executar_(true);
}
function tabela3Rotina() {
  return t3Executar_(false);
}
function tabela3Reconstruir() {
  t3LimparAgendamento_();
  return t3Executar_(false);
}
function tabela3Vigia() {
  var atual = t3Impressao_();
  var visto = t3Props_().getProperty('T3_IMPRESSAO') || '';
  if (atual && atual === visto && !cmPrecisaEscrever_(T3_ABA_SAIDA)) return;
  t3Executar_(false);
}
function tabela3OnChange(e) {
  if (t3PainelMudou_()) return t3AtualizarPeloPainel_();
  if (t3OrigemMudanca_(e) !== 'fonte') return;
  if (t3GeracaoRecente_()) return;
  t3Agendar_();
}
function configurarTabela3(silencioso) {
  var ss = t3SS_();
  desligarTabela3(true);
  t3Props_().deleteProperty('T3_DESLIGADO');
  var hora = t3HoraRotina_();
  var falhas = [];
  var pedidos = [
    { rotulo: 'reconstrução diária', criar: function() {
      ScriptApp.newTrigger('tabela3Rotina').timeBased().everyDays(1).atHour(hora).create();
    } },
    { rotulo: 'reação a edições', criar: function() {
      ScriptApp.newTrigger('tabela3OnChange').forSpreadsheet(ss).onChange().create()
      ;
    } },
    { rotulo: 'vigia', criar: function() {
      ScriptApp.newTrigger('tabela3Vigia').timeBased().everyMinutes(T3_MINUTOS_VIGIA
      ).create();
    } }
  ];
  pedidos.forEach(function(p) {
    try {
      p.criar();
    } catch (err) {
      falhas.push(p.rotulo + ' (' + err.message + ')');
    }
  });
  var msg = 'Gatilhos da aba ' + T3_ABA_SAIDA + ' instalados: vigia a cada ' +
  T3_MINUTOS_VIGIA +
    ' min, reação a edições no banco e no painel do CONSORT e reconstrução diária às ' + hora + 'h.';
  if (falhas.length) {
    msg = 'ATENÇÃO: não foi possível instalar ' + falhas.join('; ') + '. ' +
      'O projeto aceita no máximo 20 gatilhos por conta; confira em Acionadores. ' + msg;
  }
  if (silencioso !== true) t3Avisar_(msg);
  return msg;
}
function desligarTabela3(interno) {
  var alvos = ['tabela3Rotina', 'tabela3OnChange', 'tabela3Reconstruir',
  'tabela3Vigia'];
  var removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (alvos.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  t3Props_().deleteProperty('T3_AGENDADO');
  if (interno !== true) t3Props_().setProperty('T3_DESLIGADO', '1');
  return removidos + ' gatilho(s) da aba ' + T3_ABA_SAIDA + ' removido(s).';
}
function t3GruposEstudos_(sh, blocos) {
  if (!sh || !blocos || !blocos.estudos) return;
  var nC = sh.getMaxColumns();
  var quer = [0];
  for (var c = 1; c <= nC; c++) quer.push(0);
  blocos.estudos.forEach(function(b) {
    for (var k = b.c1; k <= Math.min(b.c2, nC); k++) quer[k] = 1;
  });
  var atual = [0];
  var igual = true;
  for (var c2 = 1; c2 <= nC; c2++) {
    atual.push(sh.getColumnGroupDepth(c2));
    if (atual[c2] !== quer[c2]) igual = false;
  }
  if (igual) return;
  var maximo = Math.max.apply(null, atual);
  for (var d = maximo; d >= 1; d--) {
    t3Faixas_(nC, function(k) { return atual[k] >= d; }).forEach(function(f) {
      sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(-1);
    });
  }
  sh.showColumns(1, nC);
  t3Faixas_(nC, function(k) { return quer[k] === 1; }).forEach(function(f) {
    sh.getRange(1, f[0], 1, f[1] - f[0] + 1).shiftColumnGroupDepth(1);
  });
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
}
function t3Faixas_(n, teste) {
  var faixas = [];
  var ini = 0;
  for (var c = 1; c <= n + 1; c++) {
    var ok = c <= n && teste(c);
    if (ok && !ini) ini = c;
    if (!ok && ini) {
      faixas.push([ini, c - 1]);
      ini = 0;
    }
  }
  return faixas;
}
function t3Destaque_(registros) {
  var mapa = {};
  registros.forEach(function(r) {
    if (!r.pacote) return;
    var chave = r.grupo ? r.estudo : 'sem';
    if (!t3CorDestaque_(chave)) return;
    if (!mapa[chave]) mapa[chave] = {};
    if (!mapa[chave][r.pacote.linha]) mapa[chave][r.pacote.linha] = {};
    var celulas = mapa[chave][r.pacote.linha];
    for (var c in r.pacote.usadas) {
      if (r.pacote.usadas.hasOwnProperty(c)) celulas[c] = true;
    }
  });
  return mapa;
}
function t3CorDestaque_(chave) {
  for (var i = 0; i < T3_CORES_DESTAQUE.length; i++) {
    if (T3_CORES_DESTAQUE[i].chave === chave) return T3_CORES_DESTAQUE[i];
  }
  return null;
}
function t3RegrasDestaque_(mapa, nC) {
  var dados = [];
  var linhas = [];
  T3_CORES_DESTAQUE.forEach(function(cor) {
    var porLinha = mapa[cor.chave];
    var numeros = porLinha ? Object.keys(porLinha).map(Number) : [];
    if (!numeros.length) return;
    var ultima = Math.max.apply(null, numeros);
    var porTrecho = {};
    numeros.forEach(function(l) {
      t3Faixas_(nC, function(c) { return !!porLinha[l][c - 1];
}).forEach(function(t) {
        var k = t[0] + ':' + t[1];
        if (!porTrecho[k]) porTrecho[k] = {};
        porTrecho[k][l] = true;
      });
    });
    var a1Dados = [];
    Object.keys(porTrecho).forEach(function(k) {
      var c = k.split(':').map(Number);
      t3Faixas_(ultima, function(l) { return !!porTrecho[k][l];
}).forEach(function(t) {
        a1Dados.push(t3A1_(t[0], c[0], t[1] - t[0] + 1, c[1] - c[0] + 1));
      });
    });
    var a1Linhas = t3Faixas_(ultima, function(l) { return !!porLinha[l];
}).map(function(t) {
      return t3A1_(t[0], 1, t[1] - t[0] + 1, nC);
    });
    if (a1Dados.length) dados.push({ formula: t3FormulaDestaque_(cor.nome,
      'dados lidos'), cor: cor.dados, a1: a1Dados.sort() });
    linhas.push({ formula: t3FormulaDestaque_(cor.nome, 'linha analisada'),
      cor: cor.linha, a1: a1Linhas.sort() });
  });
  return { dados: dados, linhas: linhas };
}
function t3FormulaDestaque_(categoria, tipo) {
  return '=N("RRT · ' + T3_ABA_SAIDA + ' · ' + categoria + ' · ' + tipo +
  '")=0';
}
function t3MarcaDestaque_(regra) {
  try {
    var cond = regra.getBooleanCondition();
    if (!cond) return null;
    var m = String((cond.getCriteriaValues() || [])[0] || '').replace(/\s+/g,
    ' ')
      .match(/^\s*=\s*N\(\s*"RRT\s*·\s*(.*?)\s*·\s*(.*?)\s*·\s*(dados lidos|linha analisada)"\s*\)\s*=\s*0\s*$/i);
    return m ? { dono: m[1].trim().toUpperCase(), dados: m[3].toLowerCase()
      === 'dados lidos' } : null;
  } catch (err) {
    return null;
  }
}
function t3AplicarDestaque_(sh, mapa) {
  if (!sh || !mapa) return '';
  try {
    var desejo = t3RegrasDestaque_(mapa, sh.getLastColumn());
    var outras = [];
    var dados = [];
    var linhas = [];
    var minhas = [];
    var fase = 0;
    var emOrdem = true;
    sh.getConditionalFormatRules().forEach(function(regra) {
      var marca = t3MarcaDestaque_(regra);
      var f = !marca ? 0 : (marca.dados ? 1 : 2);
      if (f < fase) emOrdem = false;
      fase = Math.max(fase, f);
      if (!marca) outras.push(regra);
      else if (marca.dono === T3_ABA_SAIDA.toUpperCase()) minhas.push(regra);
      else if (marca.dados) dados.push(regra);
      else linhas.push(regra);
    });
    if (emOrdem && t3MesmasRegras_(minhas,
    desejo.dados.concat(desejo.linhas))) return '';
    var nova = function(d) {
      return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(d.formula).setBackground(d.cor)
        .setRanges(sh.getRangeList(d.a1).getRanges()).build();
    };
    sh.setConditionalFormatRules(outras.concat(dados, desejo.dados.map(nova),
    linhas, desejo.linhas.map(nova)));
    return '';
  } catch (err) {
    return ' ATENÇÃO: não foi possível colorir as linhas analisadas na aba ' +
    sh.getName() + ' (' + err.message + ').';
  }
}
function t3MesmasRegras_(atuais, desejadas) {
  if (atuais.length !== desejadas.length) return false;
  var chave = function(formula, cor, a1) {
    return JSON.stringify([String(formula || '').replace(/\s+/g,
    '').toLowerCase(), String(cor || '').toLowerCase(),
      a1.map(function(x) { return String(x).toUpperCase(); }).sort()]);
  };
  var a = atuais.map(function(regra) {
    var cond = regra.getBooleanCondition();
    return chave((cond.getCriteriaValues() || [])[0], cond.getBackground(),
      regra.getRanges().map(function(r) { return r.getA1Notation(); }));
  }).sort();
  var d = desejadas.map(function(x) { return chave(x.formula, x.cor, x.a1);
}).sort();
  return JSON.stringify(a) === JSON.stringify(d);
}
function t3SS_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Nenhuma planilha ativa.');
  return ss;
}
function t3Props_() {
  return PropertiesService.getDocumentProperties();
}
function t3HoraRotina_() {
  if (T3_HORA_ROTINA !== null && T3_HORA_ROTINA !== undefined) return Number(T3_HORA_ROTINA);
  if (typeof HORA_ROTINA_BANCO !== 'undefined' && HORA_ROTINA_BANCO !== null)
  {
    return (Number(HORA_ROTINA_BANCO) + 1) % 24;
  }
  return 7;
}
function t3Agendar_() {
  var props = t3Props_();
  var pendente = Number(props.getProperty('T3_AGENDADO') || 0);
  var agora = Date.now();
  if (pendente && agora - pendente < T3_ATRASO_SEGUNDOS * 1000 * 3) return;
  props.setProperty('T3_AGENDADO', String(agora));
  ScriptApp.newTrigger('tabela3Reconstruir').timeBased().after(T3_ATRASO_SEGUNDOS * 1000).create();
}
function t3LimparAgendamento_() {
  t3Props_().deleteProperty('T3_AGENDADO');
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tabela3Reconstruir')
    ScriptApp.deleteTrigger(t);
  });
}
function t3GeracaoRecente_() {
  var ultima = Number(t3Props_().getProperty('T3_ULTIMA_GERACAO') || 0);
  return ultima && Date.now() - ultima < T3_JANELA_IGNORAR_MS;
}
function t3OrigemMudanca_(e) {
  var ignorados = ['FORMAT', 'OTHER'];
  if (e && e.changeType && ignorados.indexOf(String(e.changeType)) >= 0)
  return '';
  var ativa = null;
  try {
    ativa = (e && e.source ? e.source : t3SS_()).getActiveSheet();
  } catch (err) {
    return 'fonte';
  }
  if (!ativa) return 'fonte';
  var nome = ativa.getName();
  if (nome === T3_PAINEL_ABA) return 'painel';
  if (nome === T3_ABA_SAIDA) return T3_REAGIR_PROPRIA_ABA ? 'painel' : '';
  return T3_ABAS_FONTE.indexOf(nome) >= 0 ? 'fonte' : '';
}
function t3Impressao_() {
  var ss = t3SS_();
  try {
    return 'drv:' +
    DriveApp.getFileById(ss.getId()).getLastUpdated().getTime();
  } catch (err) {}
  try {
    var sh = t3AbaFonte_(ss);
    return 'dim:' + sh.getLastRow() + 'x' + sh.getLastColumn();
  } catch (err2) {}
  return '';
}
function t3Executar_(interativo, doPainel) {
  var trava = cmTrava_();
  if (!trava.tryLock(doPainel ? T3_ESPERA_TRAVA_PAINEL_MS : 30000)) {
    if (doPainel) {
      t3Agendar_();
      return 'Outra execução em andamento; reconstrução da aba ' +
      T3_ABA_SAIDA + ' agendada.';
    }
    var ocupado = 'Outra execução da aba ' + T3_ABA_SAIDA + ' está em andamento.';
    if (interativo) t3Avisar_(ocupado);
    return ocupado;
  }
  try {
    var ss = t3SS_();
    var fonte = t3AbaFonte_(ss);
    var entrada = t3LerFonte_(fonte);
    var render = t3Montar_(entrada, ss);
    var assinatura = t3Assinatura_(render);
    var props = t3Props_();
    var existente = ss.getSheetByName(T3_ABA_SAIDA);
    if (existente && !props.getProperty('T3_ASSINATURA') &&
    existente.getLastRow() > 0) {
      var alheia = 'A aba "' + T3_ABA_SAIDA + '" já existe com conteúdo que não foi gerado por este script; nada foi alterado. ' +
        'Renomeie essa aba ou troque T3_ABA_SAIDA no início do arquivo e rode de novo.';
      if (interativo) throw new Error(alheia);
      return alheia;
    }
    if (!cmPrecisaEscrever_(T3_ABA_SAIDA) && !interativo && existente && assinatura ===
    props.getProperty('T3_ASSINATURA')) {
      props.setProperty('T3_IMPRESSAO', t3Impressao_());
      if (render.chaveCorte) props.setProperty('T3_CORTE_USADO',
      render.chaveCorte);
      if (render.chavePainel !== undefined)
      props.setProperty('T3_PAINEL_LIDO', render.chavePainel);
      t3GruposEstudos_(existente, render.blocos);
      return 'Aba ' + T3_ABA_SAIDA + ' já está em dia.' +
      t3AplicarDestaque_(fonte, render.destaque);
    }
    t3Escrever_(ss, render);
    if (render.depoisDeEscrever) render.depoisDeEscrever(ss);
    t3GruposEstudos_(ss.getSheetByName(T3_ABA_SAIDA), render.blocos);
    var avisoCores = t3AplicarDestaque_(fonte, render.destaque);
    props.setProperty('T3_ASSINATURA', assinatura);
    props.setProperty('T3_ULTIMA_GERACAO', String(Date.now()));
    props.setProperty('T3_IMPRESSAO', t3Impressao_());
    if (render.chaveCorte) props.setProperty('T3_CORTE_USADO',
    render.chaveCorte);
    if (render.chavePainel !== undefined) props.setProperty('T3_PAINEL_LIDO',
    render.chavePainel);
    cmConcluir_(T3_ABA_SAIDA);
    var msg = 'Aba ' + T3_ABA_SAIDA + ' atualizada. ' + (render.resumo || '')
    + ' Fonte: ' + fonte.getName() + '.';
    if (render.faltando && render.faltando.length) {
      msg += ' ATENÇÃO: colunas não localizadas na origem: ' +
      render.faltando.join(', ') + '.';
    }
    msg += avisoCores;
    if (interativo) {
      msg += t3GarantirGatilhos_();
      t3Avisar_(msg);
    }
    return msg;
  } finally {
    trava.releaseLock();
  }
}
function t3GarantirGatilhos_() {
  try {
    if (t3Props_().getProperty('T3_DESLIGADO')) {
      return ' Atualização automática desta aba desligada por desligarTabela3; para religar, rode configurarTabela3.';
    }
    var instalado = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === 'tabela3OnChange';
    });
    return instalado ? '' : ' ' + configurarTabela3(true);
  } catch (err) {
    return ' Não foi possível instalar a atualização automática (' +
    err.message + '); rode configurarTabela3.';
  }
}
function t3Avisar_(texto) {
  try {
    SpreadsheetApp.getActive().toast(texto, T3_ABA_SAIDA, 8);
  } catch (err) {}
}
function t3AbaFonte_(ss) {
  for (var i = 0; i < T3_ABAS_FONTE.length; i++) {
    var sh = ss.getSheetByName(T3_ABAS_FONTE[i]);
    if (sh) return sh;
  }
  throw new Error('Nenhuma aba de origem encontrada. Procurei por: ' +
  T3_ABAS_FONTE.join(', ') + '.');
}
function t3LerFonte_(sh) {
  var nCols = sh.getLastColumn();
  var nLinhas = sh.getLastRow() - T3_LINHA_DADOS + 1;
  if (nCols < 2) throw new Error('A aba "' + sh.getName() + '" não tem colunas suficientes.');
  var cabecalho = sh.getRange(T3_LINHA_CABECALHO, 1, 1,
  nCols).getDisplayValues()[0];
  var dados = nLinhas > 0 ? sh.getRange(T3_LINHA_DADOS, 1, nLinhas,
  nCols).getValues() : [];
  return { cabecalho: cabecalho, dados: dados };
}
function t3Assinatura_(render) {
  var texto = JSON.stringify([T3_VERSAO, render.matriz, render.formatos,
  render.larguras, render.alturas]);
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, texto,
    Utilities.Charset.UTF_8);
    var hex = '';
    for (var i = 0; i < bytes.length; i++) {
      var b = (bytes[i] + 256) % 256;
      hex += (b < 16 ? '0' : '') + b.toString(16);
    }
    return hex;
  } catch (err) {
    return 'len' + texto.length;
  }
}
function t3Chave_(v) {
  var s = String(v === null || v === undefined ? '' : v).toLowerCase();
  var saida = '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    saida += T3_ACENTOS[c] || c;
  }
  return saida.replace(/[^a-z0-9]+/g, '');
}
function t3Vazio_(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return isNaN(v.getTime());
  var s = String(v).trim();
  if (!s) return true;
  return T3_MARCADORES_VAZIOS.indexOf(s.toLowerCase()) >= 0;
}
function t3Texto_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/'
    + ('0' + v.getDate()).slice(-2);
  }
  return String(v).trim();
}
function t3Numero_(v) {
  if (t3Vazio_(v)) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  if (v instanceof Date) return null;
  var s = String(v).trim().replace(/\s/g, '').replace(/^r\$/i, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function t3Data_(v) {
  if (t3Vazio_(v)) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var iso = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1,
  Number(iso[3]));
  var br = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
  return null;
}
function t3Indices_(cabecalho, campos) {
  var porChave = {};
  for (var c = 0; c < cabecalho.length; c++) {
    var k = t3Chave_(cabecalho[c]);
    if (!k) continue;
    if (!porChave[k]) porChave[k] = [];
    porChave[k].push(c);
  }
  var mapa = {};
  for (var campo in campos) {
    if (!campos.hasOwnProperty(campo)) continue;
    var achados = [];
    campos[campo].forEach(function(alias) {
      (porChave[t3Chave_(alias)] || []).forEach(function(idx) {
        if (achados.indexOf(idx) < 0) achados.push(idx);
      });
    });
    mapa[campo] = achados;
  }
  return mapa;
}
function t3Faltando_(idx, campos, exigidos) {
  var faltando = [];
  exigidos.forEach(function(alternativas) {
    var achou = alternativas.some(function(c) { return idx[c] &&
      idx[c].length; });
    if (!achou) faltando.push(campos[alternativas[0]][0]);
  });
  return faltando;
}
function t3ValorBruto_(linha, indices) {
  if (!indices || !indices.length) return '';
  for (var i = 0; i < indices.length; i++) {
    var v = linha[indices[i]];
    if (!t3Vazio_(v)) {
      if (linha.usadas) linha.usadas[indices[i]] = true;
      return v;
    }
  }
  return '';
}
function t3Usar_(linha, colunas) {
  if (!linha.usadas) return;
  colunas.forEach(function(c) {
    if (typeof c === 'number' && c >= 0 && !t3Vazio_(linha[c]))
    linha.usadas[c] = true;
  });
}
function t3CodigoGrupo_(v) {
  if (t3Vazio_(v)) return '';
  var n = t3Numero_(v);
  return n !== null ? String(n) : t3Texto_(v);
}
function t3EstudoEfetivo_(estudo, grupo) {
  if (!t3Vazio_(estudo)) {
    var n = t3Numero_(estudo);
    if (n !== null) return String(n);
    var k = t3Chave_(estudo);
    if (T3_ESTUDO_POR_TEXTO.hasOwnProperty(k)) return T3_ESTUDO_POR_TEXTO[k];
    return t3Texto_(estudo);
  }
  for (var i = 0; i < T3_GRUPOS_POR_ESTUDO.length; i++) {
    if (T3_GRUPOS_POR_ESTUDO[i].grupos.indexOf(grupo) >= 0) return
    T3_GRUPOS_POR_ESTUDO[i].estudo;
  }
  return '';
}
function t3ChavePaciente_(codpac, numeroLinha) {
  if (t3Vazio_(codpac)) return '#linha' + numeroLinha;
  var n = t3Numero_(codpac);
  if (n !== null) return 'p' + n;
  return 'p' + (t3Chave_(codpac) || t3Texto_(codpac));
}
function t3Estudos_(registros) {
  var lista = [];
  T3_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var grupos = cfg.grupos.slice();
    var extras = [];
    var semGrupo = 0;
    registros.forEach(function(r) {
      if (r.estudo !== cfg.estudo) return;
      if (!r.grupo) { semGrupo++; return; }
      if (grupos.indexOf(r.grupo) < 0 && extras.indexOf(r.grupo) < 0)
      extras.push(r.grupo);
    });
    extras.sort(t3CompararChaves_);
    lista.push({ estudo: cfg.estudo, grupos: grupos.concat(extras),
      temSemGrupo: semGrupo > 0 });
  });
  return lista;
}
function t3RotuloGrupo_(g) {
  if (T3_ROTULOS.grupo.hasOwnProperty(g)) return T3_ROTULOS.grupo[g];
  return 'Grupo ' + g;
}
function t3RotuloEstudo_(e) {
  return T3_ROTULOS.estudo.hasOwnProperty(e) ? T3_ROTULOS.estudo[e] : 'estudo ' + e;
}
function t3NomeEstudo_(e) {
  return T3_NOME_ESTUDO.hasOwnProperty(e) ? T3_NOME_ESTUDO[e] : String(e);
}
function t3ArredondarPar_(v, casas) {
  var f = Math.pow(10, casas);
  var x = v * f;
  var piso = Math.floor(x);
  var frac = x - piso;
  var r;
  if (Math.abs(frac - 0.5) < 1e-9) r = (piso % 2 === 0) ? piso : piso + 1;
  else r = Math.round(x);
  return r / f;
}
function t3Fmt_(v, casas) {
  if (v === null || v === undefined || isNaN(v)) return '';
  return t3ArredondarPar_(v, casas).toFixed(casas).replace('.', ',');
}
function t3NumeroP_(p) {
  if (p === null || p === undefined || isNaN(p)) return '';
  if (p < 0.001) return '< 0,001';
  return t3ArredondarPar_(p, 3);
}
function t3Plural_(n, singular, plural) {
  return n === 1 ? singular : plural;
}
function t3Contar_(obj) {
  var n = 0;
  for (var k in obj) {
    if (obj.hasOwnProperty(k)) n++;
  }
  return n;
}
function t3LinhaVazia_(n) {
  var linha = [];
  for (var i = 0; i < n; i++) linha.push('');
  return linha;
}
function t3CompararChaves_(a, b) {
  var na = t3Numero_(a);
  var nb = t3Numero_(b);
  if (na !== null && nb !== null) return na - nb;
  if (na !== null) return -1;
  if (nb !== null) return 1;
  return a < b ? -1 : (a > b ? 1 : 0);
}
function t3Letra_(c) {
  var s = '';
  while (c > 0) {
    var m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}
function t3A1_(l, c, nl, nc) {
  var a = t3Letra_(c) + l;
  nl = nl || 1;
  nc = nc || 1;
  if (nl === 1 && nc === 1) return a;
  return a + ':' + t3Letra_(c + nc - 1) + (l + nl - 1);
}
function t3NovoRender_() {
  return { matriz: [], nColunas: 1, formatos: [], larguras: {}, alturas: {},
    faltando: [], resumo: '' };
}
function t3Por_(render, l, c, valor) {
  while (render.matriz.length < l) render.matriz.push([]);
  var linha = render.matriz[l - 1];
  while (linha.length < c) linha.push('');
  linha[c - 1] = valor;
}
function t3Fechar_(render) {
  var nC = render.nColunas;
  render.matriz.forEach(function(linha) {
    if (linha.length > nC) nC = linha.length;
  });
  render.nColunas = nC;
  render.matriz.forEach(function(linha) {
    while (linha.length < nC) linha.push('');
  });
  return render;
}
function t3Marcar_(render, tipo, l, c, nl, nc, valor) {
  render.formatos.push({ tipo: tipo, l: l, c: c, nl: nl || 1, nc: nc || 1,
    valor: valor });
}
function t3Largura_(render, col, px) {
  if (!render.larguras[col] || render.larguras[col] < px) render.larguras[col]
  = px;
}
function t3Altura_(render, linha, px) {
  if (!render.alturas[linha] || render.alturas[linha] < px)
  render.alturas[linha] = px;
}
function t3ContarLinhas_(texto, larguraPx, pxPorCaractere) {
  var porLinha = Math.max(1, Math.floor((larguraPx - 8) / (pxPorCaractere ||
  6.6)));
  var total = 0;
  String(texto).split('\n').forEach(function(paragrafo) {
    var palavras = paragrafo.split(/\s+/).filter(function(p) { return p; });
    if (!palavras.length) { total++; return; }
    var atual = 0;
    var linhas = 1;
    palavras.forEach(function(p) {
      var tam = p.length;
      if (atual === 0) {
        atual = tam;
      } else if (atual + 1 + tam <= porLinha) {
        atual += 1 + tam;
      } else {
        linhas++;
        atual = tam;
      }
      while (atual > porLinha) {
        linhas++;
        atual -= porLinha;
      }
    });
    total += linhas;
  });
  return total;
}
function t3AlturaPara_(nLinhas) {
  return nLinhas <= 1 ? 21 : 21 + 15 * (nLinhas - 1);
}
function t3Escrever_(ss, r) {
  var nomeAba = r.aba || T3_ABA_SAIDA;
  r = cmPreparar_(ss, nomeAba, r);
  var sh = ss.getSheetByName(nomeAba);
  if (!sh) sh = ss.insertSheet(nomeAba);
  var preservar = r.preservar || [];
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  if (preservar.length) sh.clearFormats();
  else sh.clear();
  sh.clearNotes();
  if (r.validacoes) sh.getRange(1, 1, sh.getMaxRows(),
  sh.getMaxColumns()).clearDataValidations();
  if (!r.matriz.length) r.matriz.push(t3LinhaVazia_(Math.max(1, r.nColunas)));
  t3Fechar_(r);
  var nL = r.matriz.length;
  var nC = r.nColunas;
  if (sh.getMaxColumns() < nC) sh.insertColumnsAfter(sh.getMaxColumns(), nC -
  sh.getMaxColumns());
  if (sh.getMaxRows() < nL) sh.insertRowsAfter(sh.getMaxRows(), nL -
  sh.getMaxRows());
  var tudo = sh.getRange(1, 1, nL, nC);
  t3AplicarGrupos_(sh, r.formatos, function(f) { return f.tipo === 'formato';
});
  cmFormatos_(sh, r);
  if (preservar.length) {
    t3Retangulos_(nL, nC, preservar).forEach(function(q) {
      sh.getRange(q.l, q.c, q.nl, q.nc).setValues(r.matriz.slice(q.l - 1, q.l
      - 1 + q.nl).map(function(linha) {
        return linha.slice(q.c - 1, q.c - 1 + q.nc);
      }));
    });
  } else {
    tudo.setValues(r.matriz);
  }
  tudo.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle').setWrap(false);
  r.formatos.forEach(function(f) {
    if (f.tipo === 'mesclar') sh.getRange(f.l, f.c, f.nl, f.nc).merge();
  });
  t3AplicarGrupos_(sh, r.formatos, function(f) {
    return f.tipo !== 'formato' && f.tipo !== 'mesclar' && f.tipo !== 'rico';
  });
  var negrito = null;
  r.formatos.forEach(function(f) {
    if (f.tipo !== 'rico') return;
    if (!negrito) negrito =
    SpreadsheetApp.newTextStyle().setBold(true).build();
    var construtor = SpreadsheetApp.newRichTextValue().setText(f.valor.texto);
    f.valor.negritos.forEach(function(par) {
      if (par[1] > par[0]) construtor.setTextStyle(par[0], par[1], negrito);
    });
    sh.getRange(f.l, f.c).setRichTextValue(construtor.build());
  });
  (r.validacoes || []).forEach(function(v) {
    sh.getRange(v.l, v.c).setDataValidation(v.regra);
  });
  t3Medidas_(nC, r.larguras, 62, function(ini, qtd, px) {
    sh.setColumnWidths(ini, qtd, px); });
  t3Medidas_(nL, r.alturas, 21, function(ini, qtd, px) { sh.setRowHeights(ini,
    qtd, px); });
  if (r.congelarColunas) sh.setFrozenColumns(r.congelarColunas);
  if (r.congelarLinhas) sh.setFrozenRows(r.congelarLinhas);
  if (sh.getMaxRows() > nL) sh.deleteRows(nL + 1, sh.getMaxRows() - nL);
  if (sh.getMaxColumns() > nC) sh.deleteColumns(nC + 1, sh.getMaxColumns() -
  nC);
  SpreadsheetApp.flush();
}
function t3Retangulos_(nL, nC, preservar) {
  var furos = {};
  preservar.forEach(function(p) { furos[p.l + ',' + p.c] = true; });
  var segmentos = [];
  for (var c = 1; c <= nC; c++) {
    var lista = [];
    var ini = 0;
    for (var l = 1; l <= nL + 1; l++) {
      var livre = l <= nL && !furos[l + ',' + c];
      if (livre && !ini) ini = l;
      if (!livre && ini) {
        lista.push([ini, l - 1]);
        ini = 0;
      }
    }
    segmentos.push(JSON.stringify(lista));
  }
  var retangulos = [];
  var c0 = 1;
  for (var k = 2; k <= nC + 1; k++) {
    if (k <= nC && segmentos[k - 1] === segmentos[c0 - 1]) continue;
    JSON.parse(segmentos[c0 - 1]).forEach(function(seg) {
      retangulos.push({ l: seg[0], c: c0, nl: seg[1] - seg[0] + 1, nc: k - c0
});
    });
    c0 = k;
  }
  return retangulos;
}
function t3Medidas_(n, mapa, padrao, aplicar) {
  var ini = 1;
  var atual = mapa[1] || padrao;
  for (var i = 2; i <= n + 1; i++) {
    var v = i <= n ? (mapa[i] || padrao) : null;
    if (v !== atual) {
      aplicar(ini, i - ini, atual);
      ini = i;
      atual = v;
    }
  }
}
function t3AplicarGrupos_(sh, formatos, filtro) {
  var grupos = {};
  var ordem = [];
  formatos.forEach(function(f) {
    if (!filtro(f)) return;
    var chave = f.tipo + '|' + JSON.stringify(f.valor);
    if (!grupos[chave]) {
      grupos[chave] = { f: f, a1: [] };
      ordem.push(chave);
    }
    grupos[chave].a1.push(t3A1_(f.l, f.c, f.nl, f.nc));
  });
  ordem.forEach(function(k) {
    var g = grupos[k];
    for (var i = 0; i < g.a1.length; i += 300) {
      t3AplicarLista_(sh.getRangeList(g.a1.slice(i, i + 300)), g.f);
    }
  });
}
function t3AplicarLista_(lista, f) {
  var v = f.valor;
  if (f.tipo === 'negrito') lista.setFontWeight(v ? 'bold' : 'normal');
  else if (f.tipo === 'italico') lista.setFontStyle(v ? 'italic' : 'normal');
  else if (f.tipo === 'fundo') lista.setBackground(v);
  else if (f.tipo === 'formato') lista.setNumberFormat(v);
  else if (f.tipo === 'alinhaH') lista.setHorizontalAlignment(v);
  else if (f.tipo === 'alinhaV') lista.setVerticalAlignment(v);
  else if (f.tipo === 'quebra') lista.setWrap(v);
  else if (f.tipo === 'tamanho') lista.setFontSize(v);
  else if (f.tipo === 'rotacao') lista.setTextRotation(v);
  else if (f.tipo === 'corTexto') lista.setFontColor(v);
  else if (f.tipo === 'borda') {
    if (v.estilo === 'tracejada') {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null, '#000000',
      SpreadsheetApp.BorderStyle.DASHED);
    } else {
      lista.setBorder(v.t, v.l, v.b, v.r, null, null);
    }
  }
}
function t3Caixa_(render, l, c, nl, nc, estilo) {
  t3Marcar_(render, 'borda', l, c, nl, nc, { t: true, l: true, b: true, r:
    true, estilo: estilo || 'solida' });
}
function t3LinhaBaixo_(render, l, c, nc) {
  t3Marcar_(render, 'borda', l, c, 1, nc, { t: null, l: null, b: true, r:
    null, estilo: 'solida' });
}
function t3Laterais_(render, l, c, nl) {
  t3Marcar_(render, 'borda', l, c, nl, 1, { t: null, l: true, b: null, r:
    true, estilo: 'solida' });
}
function t3Blocos_(cabecalho) {
  var inicios = [];
  for (var c = 0; c < cabecalho.length; c++) {
    var m = t3Chave_(cabecalho[c]).match(/^data(\d+)$/);
    if (m) inicios.push({ indice: Number(m[1]), col: c });
  }
  var blocos = [];
  inicios.forEach(function(b, i) {
    var fim = i + 1 < inicios.length ? inicios[i + 1].col : cabecalho.length;
    var bloco = { indice: b.indice, data: b.col, avaliacao: -1, causa: -1,
      criterios: [] };
    for (var c = b.col + 1; c < fim; c++) {
      var raiz = t3Chave_(cabecalho[c]).replace(/\d+$/, '');
      if (raiz === 'avaliacaofinal') {
        if (bloco.avaliacao < 0) bloco.avaliacao = c;
      } else if (raiz === 'causafalha') {
        if (bloco.causa < 0) bloco.causa = c;
      } else if (T3_CRITERIOS_FDI.hasOwnProperty(raiz)) {
        bloco.criterios.push({ col: c, rotulo: T3_CRITERIOS_FDI[raiz] });
      }
    }
    blocos.push(bloco);
  });
  return blocos;
}
function t3LerBlocos_(linha, blocos) {
  return blocos.map(function(b) {
    var aval = b.avaliacao >= 0 ? t3Numero_(linha[b.avaliacao]) : null;
    var causa = b.causa >= 0 ? t3Numero_(linha[b.causa]) : null;
    var criticos = [];
    var colunasCriticas = [];
    b.criterios.forEach(function(cr) {
      var escore = t3Numero_(linha[cr.col]);
      if (escore === null || escore < T3_ESCORE_FDI_FALHA) return;
      colunasCriticas.push(cr.col);
      if (criticos.indexOf(cr.rotulo) < 0) criticos.push(cr.rotulo);
    });
    return { data: t3Data_(linha[b.data]), avaliacao: aval, causa: causa,
      criticos: criticos, colunasCriticas: colunasCriticas };
  });
}
function t3DiasEntre_(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}
function t3LerPainel_(ss) {
  var valores = {};
  try {
    valores = JSON.parse(t3Props_().getProperty(T3_PAINEL_VAL) || '{}') || {};
  } catch (err) {
    valores = {};
  }
  var lidos = t3LerCelulasPainel_(ss);
  if (lidos) {
    for (var est in lidos) {
      if (!lidos.hasOwnProperty(est)) continue;
      if (!valores[est]) valores[est] = {};
      for (var campo in lidos[est]) {
        if (lidos[est].hasOwnProperty(campo)) valores[est][campo] =
        lidos[est][campo];
      }
    }
  }
  return { valores: valores, lidos: lidos };
}
function t3LerCelulasPainel_(ss) {
  var pos = null;
  try {
    pos = JSON.parse(t3Props_().getProperty(T3_PAINEL_POS) || 'null');
  } catch (err) {
    pos = null;
  }
  var sh = ss ? ss.getSheetByName(T3_PAINEL_ABA) : null;
  if (!sh || !pos) return null;
  var celulas = [];
  for (var est in pos) {
    if (!pos.hasOwnProperty(est)) continue;
    for (var campo in pos[est]) {
      if (!pos[est].hasOwnProperty(campo)) continue;
      var m = String(pos[est][campo]).match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      var c = 0;
      for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) -
      64);
      celulas.push({ est: est, campo: campo, l: Number(m[2]), c: c });
    }
  }
  if (!celulas.length) return null;
  var l1 = Infinity, c1 = Infinity, l2 = 0, c2 = 0;
  celulas.forEach(function(x) {
    l1 = Math.min(l1, x.l);
    c1 = Math.min(c1, x.c);
    l2 = Math.max(l2, x.l);
    c2 = Math.max(c2, x.c);
  });
  if (l2 > sh.getMaxRows() || c2 > sh.getMaxColumns()) return null;
  var grade = sh.getRange(l1, c1, l2 - l1 + 1, c2 - c1 +
  1).getDisplayValues();
  var lidos = {};
  celulas.forEach(function(x) {
    var v = grade[x.l - l1][x.c - c1];
    if (!lidos[x.est]) lidos[x.est] = {};
    lidos[x.est][x.campo] = String(v === null || v === undefined ? '' :
    v).trim();
  });
  return lidos;
}
function t3PainelMudou_() {
  try {
    var ss = t3SS_();
    var props = t3Props_();
    if (T3_ABA_SAIDA === T3_PAINEL_ABA) {
      var lidos = t3LerCelulasPainel_(ss);
      return !!lidos && JSON.stringify(lidos) !==
      (props.getProperty('T3_PAINEL_LIDO') || '');
    }
    return t3ChaveCortes_(t3Cortes_(ss, new Date())) !==
    (props.getProperty('T3_CORTE_USADO') || '');
  } catch (err) {
    return false;
  }
}
function t3AtualizarPeloPainel_() {
  if (T3_ABA_SAIDA !== T3_PAINEL_ABA) Utilities.sleep(T3_VANTAGEM_PAINEL_MS);
  return t3Executar_(false, true);
}
function t3ChaveCortes_(cortes) {
  return T3_GRUPOS_POR_ESTUDO.map(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (!c.periodo) return cfg.estudo + ':auto';
    return cfg.estudo + ':' + (c.coderest !== null ? c.coderest : '') + '#' +
    (c.inicio ? t3DataBr_(c.inicio) : '') +
      '-' + (c.fim ? t3DataBr_(c.fim) : '');
  }).join('|');
}
function t3FaltaCoderest_(idx, cortes) {
  if (idx.coderest && idx.coderest.length) return false;
  return T3_GRUPOS_POR_ESTUDO.some(function(cfg) { return cortes.estudos[cfg.estudo].coderest !== null; });
}
function t3ModoCorte_(v) {
  var k = t3Chave_(v);
  return !k || k.indexOf('auto') === 0 ? T3_MODO_AUTO : T3_MODO_PERIODO;
}
function t3DataEstrita_(v) {
  if (t3Vazio_(v)) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v).trim();
  var br = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(\D|$)/);
  var iso = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(\D|$)/);
  var a, m, d;
  if (br) {
    a = Number(br[3]);
    m = Number(br[2]);
    d = Number(br[1]);
  } else if (iso) {
    a = Number(iso[1]);
    m = Number(iso[2]);
    d = Number(iso[3]);
  } else {
    return null;
  }
  var data = new Date(a, m - 1, d);
  if (data.getFullYear() !== a || data.getMonth() !== m - 1 || data.getDate()
  !== d) return null;
  return data;
}
function t3Dia_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}
function t3Corte_(v, hoje) {
  var corte = { periodo: false, inicio: null, fim: null, coderest: null,
    aviso: '',
    referencia: new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate())
};
  if (!v || t3ModoCorte_(v.corte) !== T3_MODO_PERIODO) return corte;
  var inicio = t3DataEstrita_(v.inicio);
  var fim = t3DataEstrita_(v.fim);
  var coderest = t3Vazio_(v.coderest) ? null : t3Numero_(v.coderest);
  var invalidos = [];
  if (!t3Vazio_(v.inicio) && !inicio) invalidos.push('a data inicial');
  if (!t3Vazio_(v.fim) && !fim) invalidos.push('a data final');
  if (!t3Vazio_(v.coderest) && (coderest === null || coderest < 1 ||
  Math.round(coderest) !== coderest)) {
    invalidos.push('o coderest');
  }
  if (invalidos.length) {
    corte.aviso = 'não entendi ' + invalidos.join(' nem ') + '; corte ignorado, valendo auto.';
    return corte;
  }
  if (!inicio && !fim && coderest === null) {
    corte.aviso = 'preencha as datas e/ou o coderest; sem nenhum deles, vale auto.';
    return corte;
  }
  if (inicio && fim && inicio.getTime() > fim.getTime()) {
    corte.aviso = 'a data inicial é depois da data final; corte ignorado, valendo auto.';
    return corte;
  }
  corte.periodo = true;
  corte.inicio = inicio;
  corte.fim = fim;
  corte.coderest = coderest;
  if (fim && fim.getTime() < corte.referencia.getTime()) corte.referencia =
  fim;
  return corte;
}
function t3Cortes_(ss, hoje) {
  var painel = t3LerPainel_(ss);
  var cortes = { estudos: {}, padrao: t3Corte_(null, hoje), valores:
    painel.valores, lidos: painel.lidos };
  T3_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    cortes.estudos[cfg.estudo] = t3Corte_(painel.valores[cfg.estudo], hoje);
  });
  return cortes;
}
function t3CorteDe_(cortes, estudo) {
  return cortes.estudos.hasOwnProperty(estudo) ? cortes.estudos[estudo] :
  cortes.padrao;
}
function t3DentroDoCorte_(corte, data0, coderest) {
  if (!corte.periodo) return true;
  if (corte.coderest !== null && (coderest === null || coderest >
  corte.coderest)) return false;
  if (!data0) return true;
  var t = t3Dia_(data0);
  if (corte.inicio && t < corte.inicio.getTime()) return false;
  if (corte.fim && corte.coderest === null && t > corte.fim.getTime()) return false;
  return true;
}
function t3Data0NoCorte_(corte, data0) {
  if (!data0 || !corte.periodo || !corte.fim) return data0;
  return t3Dia_(data0) > corte.fim.getTime() ? null : data0;
}
function t3DataBr_(d) {
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() +
  1)).slice(-2) + '/' + d.getFullYear();
}
function t3TextoCorte_(corte) {
  if (!corte.periodo) return '';
  var partes = [];
  if (corte.coderest !== null) partes.push('coderest até ' + corte.coderest);
  if (corte.inicio && corte.fim) partes.push(t3DataBr_(corte.inicio) + ' a ' +
  t3DataBr_(corte.fim));
  else if (corte.inicio) partes.push('a partir de ' +
  t3DataBr_(corte.inicio));
  else if (corte.fim) partes.push('até ' + t3DataBr_(corte.fim));
  return partes.join('; ');
}
function t3RotuloComCorte_(rotulo, corte) {
  var texto = t3TextoCorte_(corte);
  return texto ? rotulo + ' (corte: ' + texto + ')' : rotulo;
}
function t3AvisosCorte_(cortes) {
  var avisos = [];
  T3_GRUPOS_POR_ESTUDO.forEach(function(cfg) {
    var c = cortes.estudos[cfg.estudo];
    if (c && c.aviso) avisos.push('corte do estudo ' +
    t3NomeEstudo_(cfg.estudo) + ': ' + c.aviso);
  });
  return avisos;
}
function t3Desfecho_(linha, idx, blocos, data0, corte, lidos) {
  if (data0 && !t3Data0NoCorte_(corte, data0)) {
    return { sucesso: t3DesfechoVazio_(), sobrevida: t3DesfechoVazio_() };
  }
  var usadas = linha.usadas;
  var doBanco = {};
  var d = {};
  [['sucesso', idx.dataSucesso, idx.tempoSucesso, idx.sucesso,
  idx.causaSucesso, T3_CODIGOS_FALHA_SUCESSO],
    ['sobrevida', idx.dataSobrevida, idx.tempoSobrevida, idx.sobrevida,
  idx.causaSobrevida, T3_CODIGOS_FALHA_SOBREVIDA]
  ].forEach(function(p) {
    linha.usadas = doBanco[p[0]] = {};
    d[p[0]] = t3DesfechoBanco_(linha, p[1], p[2], p[3], p[4], p[5], data0);
  });
  linha.usadas = usadas;
  if (corte.periodo && corte.fim && data0) {
    var fim = corte.fim.getTime();
    var refeito = null;
    ['sucesso', 'sobrevida'].forEach(function(k) {
      if (!d[k].data || t3Dia_(d[k].data) <= fim) return;
      if (!refeito) refeito = t3DesfechoBlocos_(lidos || t3LerBlocos_(linha,
      blocos), blocos, data0, fim);
      d[k] = refeito[k];
      doBanco[k] = null;
      refeito.usados[k].forEach(function(i) { t3Usar_(linha, [blocos[i].data,
        blocos[i].avaliacao]); });
      if (refeito.travas[k] >= 0) t3Usar_(linha,
      [blocos[refeito.travas[k]].causa]);
    });
  }
  ['sucesso', 'sobrevida'].forEach(function(k) {
    if (!usadas || !doBanco[k]) return;
    for (var c in doBanco[k]) {
      if (doBanco[k].hasOwnProperty(c)) usadas[c] = true;
    }
  });
  return d;
}
function t3DesfechoVazio_() {
  return { evento: false, codigo: null, causa: null, data: null, dias: null,
    anos: null };
}
function t3DesfechoBanco_(linha, iData, iTempo, iCodigo, iCausa, falhas,
data0) {
  var codigo = t3Numero_(t3ValorBruto_(linha, iCodigo));
  var data = t3Data_(t3ValorBruto_(linha, iData));
  var dias = null;
  var anos = null;
  if (data0 && data) {
    var n = t3DiasEntre_(data0, data);
    if (n >= 0) {
      dias = n;
      anos = n / T3_DIAS_ANO;
    }
  } else {
    var t = t3Numero_(t3ValorBruto_(linha, iTempo));
    if (t !== null && t >= 0) {
      anos = t;
      if (data0) data = new Date(data0.getFullYear(), data0.getMonth(),
      data0.getDate() + Math.round(t * T3_DIAS_ANO));
    }
  }
  return { evento: codigo !== null && falhas.indexOf(codigo) >= 0, codigo:
    codigo,
    causa: t3Numero_(t3ValorBruto_(linha, iCausa)), data: data, dias: dias,
    anos: anos };
}
function t3DesfechoBlocos_(lidos, blocos, data0, fim) {
  var ordem = [];
  for (var k = 0; k < blocos.length; k++) ordem.push(k);
  ordem.sort(function(a, b) { return blocos[a].indice - blocos[b].indice || a
    - b; });
  var reg = { sucesso: null, sobrevida: null };
  var travado = { sucesso: false, sobrevida: false };
  var usados = { sucesso: [], sobrevida: [] };
  var travas = { sucesso: -1, sobrevida: -1 };
  var falhas = { sucesso: T3_CODIGOS_FALHA_SUCESSO, sobrevida:
    T3_CODIGOS_FALHA_SOBREVIDA };
  ordem.forEach(function(i) {
    var bl = lidos[i];
    if (!bl.data || t3Dia_(bl.data) > fim) return;
    if (blocos[i].indice !== 0 && bl.avaliacao === null) return;
    var r = { data: bl.data, codigo: bl.avaliacao, causa: bl.causa };
    ['sucesso', 'sobrevida'].forEach(function(k) {
      if (travado[k]) return;
      reg[k] = r;
      usados[k].push(i);
      if (r.codigo !== null && falhas[k].indexOf(r.codigo) >= 0) {
        travado[k] = true;
        travas[k] = i;
      }
    });
  });
  return {
    sucesso: t3DesfechoRegistro_(reg.sucesso, travado.sucesso, data0),
    sobrevida: t3DesfechoRegistro_(reg.sobrevida, travado.sobrevida, data0),
    usados: usados,
    travas: travas
  };
}
function t3DesfechoRegistro_(r, travado, data0) {
  var data = r ? r.data : data0;
  var n = t3DiasEntre_(data0, data);
  return { evento: travado, codigo: travado ? r.codigo : 0, causa: travado ?
    r.causa : 0, data: data,
    dias: n >= 0 ? n : null, anos: n >= 0 ? n / T3_DIAS_ANO : null };
}
function t3Montar_(entrada, ss) {
  var idx = t3Indices_(entrada.cabecalho, T3_CAMPOS);
  var blocos = t3Blocos_(entrada.cabecalho);
  var faltando = t3Faltando_(idx, T3_CAMPOS, [['grupo'], ['dente'], ['sexo'],
  ['faces'], ['vitalidade'],
    ['data0'], ['sucesso'], ['causaSucesso'], ['sobrevida']]);
  var cortes = t3Cortes_(ss, new Date());
  if (t3FaltaCoderest_(idx, cortes)) faltando.push('coderest (usado no corte)');
  var registros = t3Registros_(entrada.dados, idx, blocos, cortes);
  var render = t3NovoRender_();
  var res = t3Tabela_(render, t3Estudos_(registros), registros, 1, cortes);
  render.faltando = faltando;
  render.resumo = res.partes.concat(t3AvisosCorte_(cortes)).join('; ') + '.';
  render.chaveCorte = t3ChaveCortes_(cortes);
  render.congelarLinhas = res.linhasCabecalho;
  render.blocos = { rotulos: [1], estudos: res.blocos };
  render.destaque = t3Destaque_(registros);
  return t3Fechar_(render);
}
function t3Registros_(dados, idx, blocos, cortes) {
  var registros = [];
  dados.forEach(function(linha, i) {
    linha.usadas = {};
    var codpac = t3ValorBruto_(linha, idx.codpac);
    var grupoBruto = t3ValorBruto_(linha, idx.grupo);
    var estudoBruto = t3ValorBruto_(linha, idx.estudo);
    var dente = t3ValorBruto_(linha, idx.dente);
    if (t3Vazio_(codpac) && t3Vazio_(grupoBruto) && t3Vazio_(dente) &&
    t3Vazio_(estudoBruto)) return;
    var grupo = t3CodigoGrupo_(grupoBruto);
    var estudo = t3EstudoEfetivo_(estudoBruto, grupo);
    var data0 = t3Data_(t3ValorBruto_(linha, idx.data0));
    var corte = t3CorteDe_(cortes, estudo);
    if (!t3DentroDoCorte_(corte, data0, t3Numero_(t3ValorBruto_(linha,
    idx.coderest)))) return;
    if (!t3Data0NoCorte_(corte, data0)) return;
    var registro = { estudo: estudo, grupo: grupo, falhou: false, pacote: {
        linha: i + T3_LINHA_DADOS, usadas: linha.usadas } };
    registros.push(registro);
    var lidos = t3LerBlocos_(linha, blocos);
    var d = t3Desfecho_(linha, idx, blocos, data0, corte, lidos);
    var falhaSucesso = d.sucesso.evento;
    var perdida = d.sobrevida.evento;
    if (!falhaSucesso && !perdida) return;
    var blocoFalha = null;
    var colunasFalha = null;
    for (var k = 0; k < lidos.length; k++) {
      if (corte.fim && (!lidos[k].data || t3Dia_(lidos[k].data) >
      corte.fim.getTime())) continue;
      if (lidos[k].avaliacao !== null &&
      T3_CODIGOS_FALHA_SUCESSO.indexOf(lidos[k].avaliacao) >= 0) {
        blocoFalha = lidos[k];
        colunasFalha = blocos[k];
        break;
      }
    }
    var codigo = falhaSucesso ? d.sucesso.codigo : (blocoFalha ?
    blocoFalha.avaliacao : d.sobrevida.codigo);
    if (!falhaSucesso && blocoFalha) t3Usar_(linha, [colunasFalha.avaliacao]);
    var causa = d.sucesso.causa;
    if (!causa && blocoFalha && blocoFalha.causa) {
      causa = blocoFalha.causa;
      t3Usar_(linha, [colunasFalha.causa]);
    }
    var motivo;
    if (causa && T3_MOTIVOS.hasOwnProperty(String(causa))) motivo =
    T3_MOTIVOS[String(causa)];
    else if (blocoFalha && blocoFalha.criticos.length) {
      motivo = blocoFalha.criticos.join('; ');
      t3Usar_(linha, blocoFalha.colunasCriticas);
    } else if (causa) motivo = 'Código ' + causa;
    else motivo = T3_ROTULO_SEM_MOTIVO;
    var nDente = t3Numero_(dente);
    var faces = t3Numero_(t3ValorBruto_(linha, idx.faces));
    registro.falhou = true;
    registro.ordemDente = nDente === null ? Infinity : nDente;
    registro.coderest = t3Numero_(t3ValorBruto_(linha, idx.coderest));
    registro.valores = {
      dente: t3Vazio_(dente) ? '' : (nDente !== null ? String(nDente) :
      t3Texto_(dente)),
      sexo: t3Sexo_(t3ValorBruto_(linha, idx.sexo)),
      grupo: grupo ? t3RotuloGrupo_(grupo) : (estudo ? t3NomeEstudo_(estudo) :
      T3_ROTULO_ESTUDO_DESCONHECIDO),
      faces: faces === null ? '' : Math.round(faces),
      vitalidade: t3Vitalidade_(t3ValorBruto_(linha, idx.vitalidade)),
      msobrevida: t3Meses_(d.sobrevida),
      msucesso: t3Meses_(d.sucesso),
      motivo: motivo,
      reparo: codigo === 1 ? 'Reparável' + (causa === T3_CODIGO_DESCIMENTACAO
      ? ' (recimentação)' : '') :
        (codigo === 2 ? 'Não reparável' : ''),
      situacao: perdida ? 'Perdida' : 'Em função'
    };
  });
  return registros;
}
function t3Meses_(desfecho) {
  if (desfecho.dias !== null) return t3ArredondarPar_(desfecho.dias /
  T3_DIAS_MES, 1);
  return desfecho.anos !== null ? t3ArredondarPar_(desfecho.anos * 12, 1) :
  '';
}
function t3Sexo_(bruto) {
  if (t3Vazio_(bruto)) return '';
  var n = t3Numero_(bruto);
  if (n === 0) return 'M';
  if (n === 1) return 'F';
  var k = t3Chave_(bruto);
  if (k === 'f' || k.indexOf('fem') === 0) return 'F';
  if (k === 'm' || k.indexOf('masc') === 0) return 'M';
  return t3Texto_(bruto);
}
function t3Vitalidade_(bruto) {
  if (t3Vazio_(bruto)) return '';
  var n = t3Numero_(bruto);
  if (n === 1) return 'Vital';
  if (n === 0) return 'Não vital';
  var k = t3Chave_(bruto);
  if (k === 'sim' || k === 's' || k === 'vital') return 'Vital';
  if (k === 'nao' || k === 'n' || k === 'naovital') return 'Não vital';
  return t3Texto_(bruto);
}
function t3Tabela_(render, estudos, registros, r1, cortes) {
  var cols = T3_COLUNAS_T3;
  // Coordenadas dos campos individuais; usadas pelo mascaramento integrado.
  render.cmCamposT5 = [];
  var blocos = [];
  estudos.forEach(function(est) {
    blocos.push({
      chave: est.estudo,
      rotulo: t3RotuloEstudo_(est.estudo),
      textoCorte: t3TextoCorte_(t3CorteDe_(cortes, est.estudo)),
      nome: 'estudo ' + t3NomeEstudo_(est.estudo),
      ordem: est.grupos,
      chaveOrdem: 'grupo',
      semSorteio: false,
      regs: registros.filter(function(r) { return r.estudo === est.estudo &&
        !!r.grupo && r.falhou; })
    });
  });
  if (registros.some(function(r) { return !r.grupo; })) {
    blocos.push({
      chave: 'sem',
      rotulo: T3_ROTULO_SEM_SORTEIO,
      textoCorte: '',
      nome: 'sem sorteio',
      ordem: estudos.map(function(e) { return e.estudo; }),
      chaveOrdem: 'estudo',
      semSorteio: true,
      regs: registros.filter(function(r) { return !r.grupo && r.falhou; })
    });
  }
  var rRotulo = r1 + 1;
  var rCab = r1 + 2;
  var linhasCab = 1;
  cols.forEach(function(col) {
    linhasCab = Math.max(linhasCab, t3ContarLinhas_(col.rotulo, col.largura,
    7));
  });
  t3Altura_(render, rCab, t3AlturaPara_(linhasCab));
  t3Largura_(render, 1, T3_LARGURA_ESPACO);
  var c = 2;
  var fim = rCab + 1;
  var partes = [];
  var faixas = [];
  blocos.forEach(function(b, k) {
    if (k > 0) {
      t3Largura_(render, c, T3_LARGURA_ESPACO);
      c++;
    }
    var c1 = c;
    b.regs.sort(function(x, y) {
      if (x.ordemDente !== y.ordemDente) return x.ordemDente < y.ordemDente ?
      -1 : 1;
      var gx = b.ordem.indexOf(x[b.chaveOrdem]);
      var gy = b.ordem.indexOf(y[b.chaveOrdem]);
      if (gx !== gy) return gx - gy;
      var tx = x.valores.msucesso === '' ? Infinity : x.valores.msucesso;
      var ty = y.valores.msucesso === '' ? Infinity : y.valores.msucesso;
      if (tx !== ty) return tx < ty ? -1 : 1;
      var cx = x.coderest === null ? Infinity : x.coderest;
      var cy = y.coderest === null ? Infinity : y.coderest;
      return cx === cy ? 0 : (cx < cy ? -1 : 1);
    });
    var n = b.regs.length;
    t3Por_(render, rRotulo, c1, b.rotulo + ' (n = ' + n + ' ' +
      t3Plural_(n, 'restauração com falha', 'restaurações com falha') +
      (b.textoCorte ? '; corte: ' + b.textoCorte : '') + ')');
    t3Marcar_(render, 'formato', rRotulo, c1, 1, 1, '@');
    cols.forEach(function(col, i) {
      t3Largura_(render, c1 + i, col.largura);
      var rotulo = b.semSorteio && col.chave === 'grupo' ?
      T3_ROTULO_ESTUDO_COLUNA : col.rotulo;
      t3Por_(render, rCab, c1 + i, rotulo);
    });
    t3Marcar_(render, 'formato', rCab, c1, 1, cols.length, '@');
    t3Marcar_(render, 'quebra', rCab, c1, 1, cols.length, true);
    var l = rCab + 1;
    if (!n) {
      t3Por_(render, l, c1, T3_ROTULO_NENHUMA);
      t3Marcar_(render, 'italico', l, c1, 1, 1, true);
      l++;
    }
    b.regs.forEach(function(r) {
      var linhas = 1;
      cols.forEach(function(col, i) {
        var v = r.valores[col.chave];
        t3Por_(render, l, c1 + i, v === null || v === undefined ? '' : v);
        if (v !== null && v !== undefined && v !== '') {
          render.cmCamposT5.push({
            l: l, c: c1 + i, lote: String(b.chave) + ':' + col.chave
          });
        }
        if (col.quebra) linhas = Math.max(linhas, t3ContarLinhas_(String(v),
        col.largura, 6.2));
      });
      if (linhas > 1) t3Altura_(render, l, t3AlturaPara_(linhas));
      l++;
    });
    var bFim = l - 1;
    cols.forEach(function(col, i) {
      t3Marcar_(render, 'alinhaH', rCab, c1 + i, bFim - rCab + 1, 1,
      col.alinha);
      if (n) {
        t3Marcar_(render, 'formato', rCab + 1, c1 + i, bFim - rCab, 1,
        col.formato);
        if (col.quebra) t3Marcar_(render, 'quebra', rCab + 1, c1 + i, bFim -
        rCab, 1, true);
      }
    });
    t3Caixa_(render, rRotulo, c1, bFim - rRotulo + 1, cols.length);
    t3LinhaBaixo_(render, rRotulo, c1, cols.length);
    t3LinhaBaixo_(render, rCab, c1, cols.length);
    faixas.push({ chave: b.chave, nome: b.semSorteio ? 'Sem sorteio' :
      t3NomeEstudo_(b.chave), c1: c1, c2: c1 + cols.length - 1 });
    c += cols.length;
    if (bFim > fim) fim = bFim;
    partes.push(b.nome + ': ' + n + ' falha(s)' + (b.textoCorte ? ' (corte: '
    + b.textoCorte + ')' : ''));
  });
  var nC = c - 1;
  if (render.nColunas < nC) render.nColunas = nC;
  t3Por_(render, r1, 1, T3_ROTULO_TABELA + ' ' + T3_TITULO);
  t3Marcar_(render, 'formato', r1, 1, 1, 1, '@');
  t3Marcar_(render, 'negrito', r1, 1, 1, nC, true);
  t3Marcar_(render, 'tamanho', r1, 1, 1, nC, 11);
  t3Marcar_(render, 'negrito', rRotulo, 1, 2, nC, true);
  return { fim: fim, partes: partes, linhasCabecalho: rCab - r1 + 1, blocos:
    faixas };
}
