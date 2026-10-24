// Gera as bases fictícias do painel "KPI Diário - Armazém".
// Simula as duas origens do painel original:
//   dados/oracle_tos/  -> extrações CSV do sistema operacional do terminal (TOS, banco Oracle)
//   dados/sharepoint/  -> planilhas de controle mantidas pelas equipes no SharePoint
// Todos os nomes, contêineres, placas e números são aleatórios. A semente é fixa,
// então cada execução gera exatamente as mesmas bases.

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const RAIZ = path.join(__dirname, '..', 'dados');
const DIR_TOS = path.join(RAIZ, 'oracle_tos');
const DIR_SP = path.join(RAIZ, 'sharepoint');
const INICIO = Date.UTC(2025, 0, 2);
const FIM = Date.UTC(2026, 9, 23, 23, 59); // última extração: 23/10/2026
const HORA = 3600e3;
const DIA = 24 * HORA;
const MIN = 60e3;

// ---------- aleatoriedade reprodutível ----------
let semente = 20261024;
function rnd() {
  semente |= 0; semente = (semente + 0x6d2b79f5) | 0;
  let t = Math.imul(semente ^ (semente >>> 15), 1 | semente);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
}
const unif = (a, b) => a + (b - a) * rnd();
const inteiro = (a, b) => Math.floor(unif(a, b + 1));
const escolher = (lista) => lista[Math.floor(rnd() * lista.length)];
function ponderado(pares) {
  const total = pares.reduce((s, [, p]) => s + p, 0);
  let r = rnd() * total;
  for (const [v, p] of pares) { if ((r -= p) <= 0) return v; }
  return pares[pares.length - 1][0];
}
function normal(media, desvio) {
  const u = 1 - rnd(), v = rnd();
  return media + desvio * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
}

// ---------- calendário (feriados nacionais e do município do Rio de Janeiro) ----------
const FERIADOS = [
  ['2025-01-01', 'Confraternização Universal'], ['2025-01-20', 'São Sebastião (municipal)'],
  ['2025-03-03', 'Carnaval'], ['2025-03-04', 'Carnaval'], ['2025-04-18', 'Paixão de Cristo'],
  ['2025-04-21', 'Tiradentes'], ['2025-04-23', 'São Jorge (estadual)'], ['2025-05-01', 'Dia do Trabalho'],
  ['2025-06-19', 'Corpus Christi'], ['2025-09-07', 'Independência do Brasil'], ['2025-10-12', 'Nossa Senhora Aparecida'],
  ['2025-11-02', 'Finados'], ['2025-11-15', 'Proclamação da República'], ['2025-11-20', 'Dia da Consciência Negra'],
  ['2025-12-25', 'Natal'],
  ['2026-01-01', 'Confraternização Universal'], ['2026-01-20', 'São Sebastião (municipal)'],
  ['2026-02-16', 'Carnaval'], ['2026-02-17', 'Carnaval'], ['2026-04-03', 'Paixão de Cristo'],
  ['2026-04-21', 'Tiradentes'], ['2026-04-23', 'São Jorge (estadual)'], ['2026-05-01', 'Dia do Trabalho'],
  ['2026-06-04', 'Corpus Christi'], ['2026-09-07', 'Independência do Brasil'], ['2026-10-12', 'Nossa Senhora Aparecida'],
  ['2026-11-02', 'Finados'], ['2026-11-15', 'Proclamação da República'], ['2026-11-20', 'Dia da Consciência Negra'],
  ['2026-12-25', 'Natal'],
];
const SET_FERIADOS = new Set(FERIADOS.map(([d]) => d));
const isoDia = (t) => new Date(t).toISOString().slice(0, 10);
const inicioDia = (t) => Math.floor(t / DIA) * DIA;
const diaSemana = (t) => new Date(t).getUTCDay(); // 0 = domingo
const ehUtil = (t) => diaSemana(t) !== 0 && diaSemana(t) !== 6 && !SET_FERIADOS.has(isoDia(t));

// Mesma regra usada no Power Query (fnPrazoDiaUtil): se o início passa da hora de corte,
// a contagem começa no dia seguinte; soma N dias úteis e fixa o horário-limite.
function prazoUtil(inicio, diasUteis, horaLimite, horaCorte) {
  let base = inicioDia(inicio);
  const horaInicio = (inicio - base) / HORA;
  if (horaCorte != null && horaInicio > horaCorte) base += DIA;
  let d = base, n = 0;
  if (diasUteis === 0) { while (!ehUtil(d)) d += DIA; }
  while (n < diasUteis) { d += DIA; if (ehUtil(d)) n++; }
  return d + horaLimite * HORA;
}

// Índice de desempenho por mês: oscila e melhora ao longo de 2026 (dá vida aos gráficos de tendência).
function taxaNoPrazo(t, base) {
  const d = new Date(t);
  const meses = (d.getUTCFullYear() - 2025) * 12 + d.getUTCMonth();
  const sazonal = 0.04 * Math.sin(meses / 1.7);
  const melhoria = meses >= 12 ? 0.05 : 0;
  return Math.min(0.97, base + sazonal + melhoria);
}

const diasOperacao = [];
for (let d = INICIO; d <= inicioDia(FIM); d += DIA) diasOperacao.push(d);

// ---------- formatação ----------
const pad = (n) => String(n).padStart(2, '0');
function dataHoraIso(t) {
  if (t == null) return '';
  const d = new Date(t);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}
const paraExcel = (t) => (t == null ? null : new Date(Math.round(t / MIN) * MIN));
function gravarCsv(arquivo, colunas, linhas) {
  const esc = (v) => {
    if (v == null) return '';
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const corpo = [colunas.join(','), ...linhas.map((l) => colunas.map((c) => esc(l[c])).join(','))].join('\n');
  fs.writeFileSync(path.join(DIR_TOS, arquivo), '﻿' + corpo + '\n', 'utf8');
  console.log(`  oracle_tos/${arquivo}: ${linhas.length} linhas`);
}

// Escreve uma aba no formato das planilhas de controle: título, linhas de cabeçalho
// decorativas e só então a tabela (o Power Query precisa pular essas linhas).
function abaControle(wb, nome, titulo, linhasTitulo, colunas, linhas) {
  const ws = wb.addWorksheet(nome);
  ws.addRow([titulo]).font = { bold: true, size: 14 };
  ws.addRow(['Planilha de controle - dados fictícios para portfólio']);
  for (let i = 2; i < linhasTitulo; i++) ws.addRow([]);
  const cab = ws.addRow(colunas);
  cab.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  cab.eachCell((c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } }; });
  for (const l of linhas) ws.addRow(colunas.map((c) => l[c] ?? null));
  colunas.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    col.width = Math.max(12, Math.min(38, c.length + 4));
    if (/DATA|TIME IN/.test(c)) col.numFmt = 'dd/mm/yyyy hh:mm';
  });
  return linhas.length;
}

// ---------- cadastros ----------
const CLIENTES = [
  ['Atlântica Comércio Exterior Ltda', 'Distribuição', 'RJ'], ['Boreal Pneus e Borrachas S.A.', 'Automotivo', 'SP'],
  ['Cerrado Mineração e Equipamentos', 'Mineração', 'MG'], ['Delta Química Industrial', 'Químico', 'RJ'],
  ['Estrela do Mar Têxtil', 'Têxtil', 'SC'], ['Farol Eletrônicos Importação', 'Eletroeletrônicos', 'SP'],
  ['Guanabara Alimentos do Brasil', 'Alimentos', 'RJ'], ['Horizonte Autopeças Ltda', 'Automotivo', 'MG'],
  ['Ipê Farmacêutica S.A.', 'Farmacêutico', 'RJ'], ['Jangada Bebidas e Insumos', 'Bebidas', 'ES'],
  ['Kuara Cosméticos', 'Cosméticos', 'SP'], ['Litoral Máquinas Agrícolas', 'Agronegócio', 'PR'],
  ['Maré Alta Siderurgia', 'Siderurgia', 'RJ'], ['Norte Papel e Embalagens', 'Papel e Celulose', 'BA'],
  ['Orla Plásticos Técnicos', 'Químico', 'RJ'], ['Pampa Agroinsumos', 'Agronegócio', 'RS'],
  ['Restinga Vidros Especiais', 'Construção', 'RJ'], ['Sertão Energia Solar', 'Energia', 'PE'],
  ['Tropical Frutas Desidratadas', 'Alimentos', 'BA'], ['Umbu Calçados e Couros', 'Calçados', 'RS'],
  ['Vento Sul Equipamentos Navais', 'Naval', 'RJ'], ['Xingu Madeiras Nobres', 'Construção', 'PA'],
  ['Zênite Tecnologia Médica', 'Saúde', 'SP'], ['Aurora Utilidades Domésticas', 'Varejo', 'RJ'],
  ['Brisa Cabos e Conectores', 'Eletroeletrônicos', 'SP'],
].map(([nome, segmento, uf], i) => ({ ID_CLIENTE: 1001 + i, RAZAO_SOCIAL: nome.toUpperCase(), SEGMENTO: segmento, UF: uf, peso: 1 + (i % 5 === 0 ? 3 : i % 3) }));

const TRANSPORTADORAS = [
  'Rota Certa Transportes', 'Via Litoral Logística', 'Expresso Serra Azul', 'TransGuanabara Cargas',
  'Caminho Seguro Ltda', 'Norte-Sul Rodoviário', 'Ágil Express Logística', 'Ponte Nova Transportes',
  'Carga Leve Distribuição', 'Atalaia Fretes', 'Bandeirante Cargas Pesadas', 'Mar e Terra Logística',
  'Planalto Transportes', 'Rodomar Logística', 'Vale Verde Fretamentos',
].map((nome, i) => ({ ID_TRANSPORTADORA: 501 + i, NOME_TRANSPORTADORA: nome.toUpperCase(), UF: escolher(['RJ', 'RJ', 'SP', 'MG', 'ES']) }));

const ARMADORES = ['OCEANIC LINES', 'ATLÂNTICO SHIPPING', 'PACÍFICO CONTAINER LINE', 'AUSTRAL MARITIME', 'NORDIC BLUE CARRIERS'];
const NAVIOS = ['ATLAS MARINER', 'BLUE HORIZON', 'COSTA VERDE STAR', 'GOLDEN WAVE', 'OCEAN PIONEER', 'SOUTHERN CROSS', 'TROPIC EXPRESS', 'VEGA TRADER', 'CORAL BAY', 'SILVER TIDE'];
const PREFIXOS_CTNR = ['TPXU', 'GLDU', 'ARMU', 'BRZU', 'OCNU', 'ATLU', 'SULU', 'MRAU'];
const LOCAIS_PATIO = ['PÁTIO A1', 'PÁTIO A2', 'PÁTIO A3', 'PÁTIO B1', 'PÁTIO B2', 'PÁTIO C1', 'PÁTIO C2', 'PÁTIO REEFER'];
const DOCAS = ['ARMAZÉM 1 - DOCA 01', 'ARMAZÉM 1 - DOCA 02', 'ARMAZÉM 1 - DOCA 03', 'ARMAZÉM 2 - DOCA 04', 'ARMAZÉM 2 - DOCA 05', 'ÁREA DE VISTORIA'];

const usados = new Set();
function novoContainer() {
  let c;
  do { c = escolher(PREFIXOS_CTNR) + String(inteiro(0, 9999999)).padStart(7, '0'); } while (usados.has(c));
  usados.add(c);
  return c;
}
function novaPlaca() {
  const L = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return `${L[inteiro(0, 25)]}${L[inteiro(0, 25)]}${L[inteiro(0, 25)]}${inteiro(0, 9)}${L[inteiro(0, 25)]}${inteiro(0, 9)}${inteiro(0, 9)}`;
}
const clienteAleatorio = () => ponderado(CLIENTES.map((c) => [c, c.peso]));
const horaUtil = (dia, de, ate) => dia + unif(de, ate) * HORA;
// Gera um instante "dentro" ou "fora" de um prazo, respeitando um limite mínimo.
function contraPrazo(prazo, minimo, noPrazo, folgaMaxH, atrasoMaxH) {
  if (noPrazo) return Math.max(minimo + unif(0.3, 1.5) * HORA, prazo - unif(0.2, folgaMaxH) * HORA);
  return Math.max(minimo + HORA, prazo + unif(0.3, atrasoMaxH) * HORA);
}

// ---------- metas (planilha de parâmetros) ----------
const METAS = [
  { PROCESSO: 'CONFERÊNCIA', ETAPA: 'POSICIONAMENTO', SERVICO_KPI: 'TODOS', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 1, HORA_LIMITE: 7, HORA_CORTE: 19, MINUTOS: null, DESCRICAO: 'Pedido até 19h: posicionar até D+1 útil às 07h' },
  { PROCESSO: 'CONFERÊNCIA', ETAPA: 'ENVIO DO RELATÓRIO', SERVICO_KPI: 'DESUNITIZAÇÃO CONFERÊNCIA TOTAL', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 3, HORA_LIMITE: 16, HORA_CORTE: 16, MINUTOS: null, DESCRICAO: 'Relatório à RFB até D+3 útil às 16h' },
  { PROCESSO: 'CONFERÊNCIA', ETAPA: 'ENVIO DO RELATÓRIO', SERVICO_KPI: 'DESUNITIZAÇÃO CONFERÊNCIA PARCIAL', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 2, HORA_LIMITE: 19, HORA_CORTE: 16, MINUTOS: null, DESCRICAO: 'Relatório à RFB até D+2 útil às 19h' },
  { PROCESSO: 'CONFERÊNCIA', ETAPA: 'ENVIO DO RELATÓRIO', SERVICO_KPI: 'RFB - SERVIÇOS SEM DESOVA', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 2, HORA_LIMITE: 16, HORA_CORTE: 16, MINUTOS: null, DESCRICAO: 'Relatório à RFB até D+2 útil às 16h' },
  { PROCESSO: 'CONFERÊNCIA', ETAPA: 'ENVIO DO RELATÓRIO', SERVICO_KPI: 'MAPA/ANVISA - SERVIÇOS SEM DESOVA', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 2, HORA_LIMITE: 11, HORA_CORTE: 16, MINUTOS: null, DESCRICAO: 'Liberação MAPA/ANVISA até D+2 útil às 11h' },
  { PROCESSO: 'ENTREGA', ETAPA: 'CICLO DO CAMINHÃO', SERVICO_KPI: 'NORMAL', TIPO_META: 'MINUTOS', DIAS_UTEIS: null, HORA_LIMITE: null, HORA_CORTE: null, MINUTOS: 120, DESCRICAO: 'Gate-in a gate-out em até 2h' },
  { PROCESSO: 'ENTREGA', ETAPA: 'CICLO DO CAMINHÃO', SERVICO_KPI: 'ESPECIAL', TIPO_META: 'MINUTOS', DIAS_UTEIS: null, HORA_LIMITE: null, HORA_CORTE: null, MINUTOS: 180, DESCRICAO: 'Carga especial (IMO, excesso, frágil): até 3h' },
  { PROCESSO: 'ENTREGA', ETAPA: 'CICLO DO CAMINHÃO', SERVICO_KPI: 'DTA', TIPO_META: 'MINUTOS', DIAS_UTEIS: null, HORA_LIMITE: null, HORA_CORTE: null, MINUTOS: 180, DESCRICAO: 'Trânsito aduaneiro (DTA): até 3h' },
  { PROCESSO: 'RECEBIMENTO', ETAPA: 'POSICIONAMENTO', SERVICO_KPI: 'TODOS', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 1, HORA_LIMITE: 7, HORA_CORTE: 19, MINUTOS: null, DESCRICAO: 'Solicitação até 19h: posicionar até D+1 útil às 07h' },
  { PROCESSO: 'RECEBIMENTO', ETAPA: 'DESUNITIZAÇÃO', SERVICO_KPI: 'TODOS', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 1, HORA_LIMITE: 19, HORA_CORTE: 16, MINUTOS: null, DESCRICAO: 'Desova concluída até D+1 útil às 19h após o posicionamento' },
  { PROCESSO: 'RECEBIMENTO', ETAPA: 'PRESENÇA DE CARGA', SERVICO_KPI: 'TODOS', TIPO_META: 'MINUTOS', DIAS_UTEIS: null, HORA_LIMITE: null, HORA_CORTE: null, MINUTOS: 360, DESCRICAO: 'Carga armazenada e informada em até 6h após a desova' },
  { PROCESSO: 'RECEBIMENTO', ETAPA: 'ENVIO DO RELATÓRIO', SERVICO_KPI: 'TODOS', TIPO_META: 'DIA ÚTIL', DIAS_UTEIS: 1, HORA_LIMITE: 12, HORA_CORTE: null, MINUTOS: null, DESCRICAO: 'Relatório de desova até D+1 útil às 12h' },
];
const meta = (proc, etapa, serv = 'TODOS') => METAS.find((m) => m.PROCESSO === proc && m.ETAPA === etapa && m.SERVICO_KPI === serv);
const aplicarMeta = (m, inicio) => (m.TIPO_META === 'MINUTOS' ? inicio + m.MINUTOS * MIN : prazoUtil(inicio, m.DIAS_UTEIS, m.HORA_LIMITE, m.HORA_CORTE));

// =====================================================================
// 1. CONFERÊNCIA ADUANEIRA
// =====================================================================
const SERVICOS_CONF = [
  // [serviço no TOS, grupo fiscal, serviço do KPI, peso]
  ['DESUNITIZAÇÃO CONFERÊNCIA TOTAL', 'RFB', 'DESUNITIZAÇÃO CONFERÊNCIA TOTAL', 22],
  ['DESUNITIZAÇÃO CONFERÊNCIA PARCIAL', 'RFB', 'DESUNITIZAÇÃO CONFERÊNCIA PARCIAL', 18],
  ['POS. CONFERÊNCIA FÍSICA RFB', 'RFB', 'RFB - SERVIÇOS SEM DESOVA', 30],
  ['POS. VISTORIA MAPA', 'MAPA/ANVISA', 'MAPA/ANVISA - SERVIÇOS SEM DESOVA', 14],
  ['POS. VISTORIA ANVISA', 'MAPA/ANVISA', 'MAPA/ANVISA - SERVIÇOS SEM DESOVA', 8],
  ['POS. FOTOGRAFIA CONTÊINER', 'RFB', null, 4], // não entra no KPI (filtrado no Power Query)
];
const ISOS = [['22G1', 30], ['42G1', 25], ['45G1', 35], ['22R1', 4], ['45R1', 6]];

const tosConf = [], spConf = [], justConf = [], solicitacoes = [];
let idPosic = 70001, nSolic = 1;
for (const dia of diasOperacao) {
  const util = ehUtil(dia);
  if (diaSemana(dia) === 0) continue;
  const qtd = util ? inteiro(6, 13) : (diaSemana(dia) === 6 ? inteiro(0, 3) : 0);
  for (let i = 0; i < qtd; i++) {
    const [servico, grupo, servKpi] = ponderado(SERVICOS_CONF.map((s) => [s, s[3]]));
    const cli = clienteAleatorio();
    const ctnr = novoContainer();
    const pedido = horaUtil(dia, 7, 21.5);
    const entrada = rnd() < 0.1 ? pedido + unif(2, 30) * HORA : pedido - unif(0.5, 12) * DIA;
    const programado = Math.max(pedido, entrada) + unif(0.3, 5) * HORA;
    const taxa = taxaNoPrazo(dia, 0.84);
    const prazoPosic = aplicarMeta(meta('CONFERÊNCIA', 'POSICIONAMENTO'), Math.max(pedido, entrada));
    const posicionado = contraPrazo(prazoPosic, programado, rnd() < taxa + 0.05, 10, 22);
    // RFB: o prazo conta do pedido do fiscal no sistema de agendamento (iConfere);
    // MAPA/ANVISA: conta da criação do pedido no TOS (ou da entrada do contêiner, se posterior).
    const pedidoIconfere = grupo === 'RFB' ? pedido + unif(0, 2.5) * HORA : null;
    const inicioServ = grupo === 'RFB' ? pedidoIconfere : Math.max(pedido, entrada);
    let fechamento = null, liberacao = null;
    if (servKpi) {
      const prazoServ = aplicarMeta(meta('CONFERÊNCIA', 'ENVIO DO RELATÓRIO', servKpi), inicioServ);
      fechamento = contraPrazo(prazoServ, posicionado + 2 * HORA, rnd() < taxa, 30, 48);
      if (servico.startsWith('DESUNITIZAÇÃO')) liberacao = posicionado + unif(1, Math.max(1.5, (fechamento - posicionado) / HORA - 1)) * HORA;
    } else {
      fechamento = posicionado + unif(0.5, 3) * HORA;
    }
    let status = 'FINALIZADO';
    if (rnd() < 0.02) { status = 'CANCELADO'; fechamento = null; liberacao = null; }
    else if (fechamento > FIM) { status = 'EM ANDAMENTO'; fechamento = null; if (liberacao > FIM) liberacao = null; }
    const id = idPosic++;
    tosConf.push({
      ID_POSIC_MERC: id, CONTAINER: ctnr, ISO: ponderado(ISOS), ID_CLIENTE: cli.ID_CLIENTE, SERVICO: servico,
      GRUPO_FISCAL: grupo, ENTRADA_CTNR: dataHoraIso(entrada), DATA_PEDIDO: dataHoraIso(pedido),
      DATA_PROGRAMADO: dataHoraIso(programado), DATA_POSICIONADO: posicionado > FIM ? '' : dataHoraIso(posicionado),
      DATA_FECHAMENTO: dataHoraIso(fechamento), STATUS: status,
    });
    if (grupo === 'RFB' && servKpi) {
      // planilha manual: contêiner às vezes digitado em minúsculas ou com espaços
      const digitado = rnd() < 0.06 ? ` ${ctnr.toLowerCase()}` : rnd() < 0.04 ? `${ctnr.slice(0, 4)} ${ctnr.slice(4)}` : ctnr;
      spConf.push({
        'CONTAINER': digitado, 'ID POSIC': id, 'DATA CRIAÇÃO PEDIDO (ICONFERE)': paraExcel(pedidoIconfere),
        'DATA:HORA LIBERAÇÃO': paraExcel(liberacao), 'SETOR': servico.startsWith('DESUNITIZAÇÃO') ? 'DESOVA' : 'CONFERÊNCIA',
        'CANAL': ponderado([['AMARELO', 55], ['VERMELHO', 45]]), 'OBSERVAÇÃO': rnd() < 0.05 ? 'Fiscal remarcou a conferência' : null,
      });
      if (rnd() < 0.015) spConf.push({ ...spConf[spConf.length - 1] }); // linha duplicada por engano
    }
    if (fechamento && servKpi) {
      const prazoServ = aplicarMeta(meta('CONFERÊNCIA', 'ENVIO DO RELATÓRIO', servKpi), inicioServ);
      if (fechamento > prazoServ && rnd() < 0.75) {
        justConf.push({ 'ID POSIC': id, 'CONTAINER': ctnr, 'JUSTIFICATIVA': escolher(['Fiscal não compareceu no horário agendado', 'Carga com avaria, aguardando vistoria', 'Falta de equipe no turno', 'Equipamento (empilhadeira) indisponível', 'Divergência documental informada pelo cliente', 'Chuva forte interrompeu a operação']), 'ÁREA RESPONSÁVEL': escolher(['OPERAÇÃO', 'FISCALIZAÇÃO', 'CLIENTE', 'MANUTENÇÃO']) });
      }
    }
    solicitacoes.push({
      'Nº': nSolic++, 'DATA:HORA SOLICITAÇÃO': paraExcel(pedido), 'CONTAINER': ctnr, 'TAM': ['22G1', '22R1'].includes(tosConf[tosConf.length - 1].ISO) ? 20 : 40,
      'TS': grupo === 'RFB' ? (servico.startsWith('DESUNITIZAÇÃO') ? 'CONFERÊNCIA COM DESOVA' : 'CONFERÊNCIA RFB') : (servico.includes('MAPA') ? 'VISTORIA MAPA' : 'VISTORIA ANVISA'),
      'LOCAL ATUAL': escolher(LOCAIS_PATIO), 'LOCAL PLANEJADO': servico.startsWith('DESUNITIZAÇÃO') ? escolher(DOCAS.slice(0, 5)) : 'ÁREA DE VISTORIA',
      'SETOR': grupo === 'RFB' ? 'CONFERÊNCIA' : 'VISTORIA', 'SOLICITANTE': 'FISCALIZAÇÃO',
      'POSICIONADO?': posicionado <= FIM && status !== 'CANCELADO' ? 'SIM' : 'NÃO', 'STATUS': status === 'CANCELADO' ? 'CANCELADO' : (posicionado <= FIM ? 'CONCLUÍDO' : 'PENDENTE'),
    });
  }
}

// =====================================================================
// 2. ENTREGA DE CARGA SOLTA (ciclo do caminhão)
// =====================================================================
const tosEntrega = [], justEntrega = [];
let ticket = 900001, lote = 250001;
for (const dia of diasOperacao) {
  if (diaSemana(dia) === 0 || SET_FERIADOS.has(isoDia(dia))) continue;
  const qtd = diaSemana(dia) === 6 ? inteiro(6, 14) : inteiro(26, 44);
  const taxa = taxaNoPrazo(dia, 0.8);
  for (let i = 0; i < qtd; i++) {
    const cli = clienteAleatorio();
    const transp = escolher(TRANSPORTADORAS);
    const janela = dia + inteiro(diaSemana(dia) === 6 ? 8 : 7, diaSemana(dia) === 6 ? 12 : 17) * HORA;
    const tipoDoc = ponderado([['DI', 45], ['DUIMP', 30], ['DTA', 15], ['DSI', 10]]);
    const especial = tipoDoc !== 'DTA' && rnd() < 0.12 ? 'S' : 'N';
    const tipo = tipoDoc === 'DTA' ? 'DTA' : especial === 'S' ? 'ESPECIAL' : 'NORMAL';
    const volumes = inteiro(2, tipo === 'ESPECIAL' ? 30 : 70);
    const status = ponderado([['FINALIZADO', 93], ['NO SHOW', 4], ['CANCELADO', 3]]);
    const lin = {
      TICKET: ticket++, STATUS: status, DATA_JANELA: dataHoraIso(janela), PLACA: novaPlaca(), ID_TRANSPORTADORA: transp.ID_TRANSPORTADORA,
      LOTE: `${new Date(dia).getUTCFullYear() % 100}${String(lote++).slice(-6)}`, ID_CLIENTE: cli.ID_CLIENTE, TIPO_DOCUMENTO: tipoDoc,
      ATENDIMENTO_ESPECIAL: especial, QTD_VOLUMES: volumes, PESO_KG: Math.round(volumes * unif(80, 900)),
      DATA_PRE_GATE: '', DATA_GATE_IN: '', INICIO_CARREGAMENTO: '', FIM_CARREGAMENTO: '', DATA_GATE_OUT: '',
    };
    if (status === 'FINALIZADO') {
      const metaMin = meta('ENTREGA', 'CICLO DO CAMINHÃO', tipo).MINUTOS;
      const gateIn = janela + Math.max(-30, Math.min(120, normal(10, 20))) * MIN;
      const preGate = Math.min(janela - unif(5, 90) * MIN, gateIn - unif(5, 25) * MIN);
      const noPrazo = rnd() < taxa;
      const cicloAlvo = noPrazo ? unif(0.45, 0.97) * metaMin : metaMin * unif(1.05, 1.9);
      const espera = cicloAlvo * unif(0.15, 0.35);
      const deslocSaida = unif(8, 20);
      const inicio = gateIn + espera * MIN;
      const fim = gateIn + (cicloAlvo - deslocSaida) * MIN;
      const gateOut = gateIn + cicloAlvo * MIN;
      Object.assign(lin, {
        DATA_PRE_GATE: dataHoraIso(preGate), DATA_GATE_IN: dataHoraIso(gateIn), INICIO_CARREGAMENTO: dataHoraIso(inicio),
        FIM_CARREGAMENTO: dataHoraIso(fim), DATA_GATE_OUT: dataHoraIso(gateOut),
      });
      if (!noPrazo && rnd() < 0.7) {
        justEntrega.push({
          'DATA JANELA': paraExcel(janela), 'PLACA': lin.PLACA,
          'JUSTIFICATIVA': escolher(['Inspeção fiscal durante o carregamento', 'Avaria na carga, aguardando registro', 'Equipamento indisponível', 'Carga em local de difícil acesso', 'Documentação do motorista incompleta', 'Fila de caminhões acima da capacidade', 'Excesso de volumes para uma única janela']),
          'ÁREA RESPONSÁVEL': escolher(['OPERAÇÃO', 'TRANSPORTADOR', 'FISCALIZAÇÃO', 'MANUTENÇÃO']),
        });
      }
    }
    tosEntrega.push(lin);
  }
}

// =====================================================================
// 3. RECEBIMENTO E DESUNITIZAÇÃO
// =====================================================================
const spReceb = [], tosMov = [];
let idMov = 3000001;
for (const dia of diasOperacao) {
  if (!ehUtil(dia)) continue;
  const qtd = inteiro(4, 10);
  const taxa = taxaNoPrazo(dia, 0.83);
  for (let i = 0; i < qtd; i++) {
    const cli = clienteAleatorio();
    const ctnr = novoContainer();
    const iso = ponderado(ISOS);
    const modal = ponderado([['FCL', 65], ['LCL', 35]]);
    const terminal = ponderado([['TERMINAL', 75], ['OUTRO TERMINAL (DTA)', 25]]);
    const fimNavio = dia - unif(0.5, 4) * DIA;
    const timeIn = terminal === 'TERMINAL' ? fimNavio + unif(1, 10) * HORA : fimNavio + unif(1, 3) * DIA;
    const solicProcesso = horaUtil(dia, 7, 20);
    const solicPosic = solicProcesso + unif(0.1, 3) * HORA;
    const pPosic = aplicarMeta(meta('RECEBIMENTO', 'POSICIONAMENTO'), solicPosic);
    const posicionado = contraPrazo(pPosic, solicPosic, rnd() < taxa + 0.05, 10, 20);
    const programado = posicionado + unif(0.3, 3) * HORA;
    const pDes = aplicarMeta(meta('RECEBIMENTO', 'DESUNITIZAÇÃO'), posicionado);
    const fimDes = contraPrazo(pDes, programado + (modal === 'LCL' ? 5 : 3) * HORA, rnd() < taxa, 14, 30);
    const armazenagem = rnd() < taxa + 0.05 ? fimDes + unif(0.3, 5.5) * HORA : fimDes + unif(6.5, 22) * HORA;
    const pRel = aplicarMeta(meta('RECEBIMENTO', 'ENVIO DO RELATÓRIO'), fimDes);
    const relatorio = contraPrazo(pRel, fimDes, rnd() < taxa + 0.03, 16, 30);
    const cancelada = rnd() < 0.03;
    const status = cancelada ? 'CANCELADO' : relatorio <= FIM ? 'FINALIZADO' : posicionado <= FIM ? 'EM ANDAMENTO' : 'AGUARDANDO POSICIONAMENTO';
    const limite = (t) => (cancelada || t > FIM ? null : paraExcel(t));
    const atrasado = fimDes > pDes || armazenagem - fimDes > 6 * HORA || relatorio > pRel;
    spReceb.push({
      'STATUS': status,
      'CLIENTE': rnd() < 0.08 ? `${cli.RAZAO_SOCIAL.toLowerCase()} ` : cli.RAZAO_SOCIAL, // digitação manual despadronizada
      'BL': `GLD${inteiro(10000000, 99999999)}`, 'CONTAINER': ctnr, 'TAM': ['22G1', '22R1'].includes(iso) ? 20 : 40,
      'MODAL': modal, 'ARMADOR': escolher(ARMADORES), 'NAVIO': escolher(NAVIOS),
      'SERVIÇO': modal === 'FCL' ? ponderado([['DESOVA TOTAL', 80], ['DESOVA PARCIAL', 20]]) : 'DESOVA TOTAL',
      'TERMINAL DE DESCARGA': terminal, 'FREE TIME (DIAS)': escolher([7, 7, 10, 14, 21]),
      'FIM OPERAÇÃO DO NAVIO': paraExcel(fimNavio), 'TIME IN': paraExcel(timeIn), 'SOLICITAÇÃO DO PROCESSO': paraExcel(solicProcesso),
      'DATA / HORA PROGRAMADO': limite(programado), 'DATA / HORA FIM DESUNITIZAÇÃO': limite(fimDes),
      'DATA / HORA ARMAZENAGEM': limite(armazenagem), 'DATA / HORA ENVIO DO RELATÓRIO': limite(relatorio),
      'DESOVA CANCELADA?': cancelada ? 'SIM' : 'NÃO',
      'JUSTIFICATIVA': atrasado && !cancelada && rnd() < 0.7 ? escolher(['Carga mal estufada pelo exportador', 'Aguardando liberação do cliente', 'Falta de espaço na área de armazenagem', 'Equipe direcionada para conferência fiscal', 'Contêiner com avaria estrutural', 'Sistema TOS indisponível']) : null,
    });
    if (rnd() < 0.01) spReceb.push({ ...spReceb[spReceb.length - 1] });
    solicitacoes.push({
      'Nº': nSolic++, 'DATA:HORA SOLICITAÇÃO': paraExcel(solicPosic), 'CONTAINER': ctnr, 'TAM': ['22G1', '22R1'].includes(iso) ? 20 : 40,
      'TS': 'DESOVA', 'LOCAL ATUAL': escolher(LOCAIS_PATIO), 'LOCAL PLANEJADO': escolher(DOCAS.slice(0, 5)), 'SETOR': 'RECEBIMENTO',
      'SOLICITANTE': 'CLIENTE', 'POSICIONADO?': !cancelada && posicionado <= FIM ? 'SIM' : 'NÃO',
      'STATUS': cancelada ? 'CANCELADO' : posicionado <= FIM ? 'CONCLUÍDO' : 'PENDENTE',
    });
    if (!cancelada && posicionado <= FIM) {
      tosMov.push({ ID_MOV: idMov++, CONTAINER: ctnr, TIPO_MOVIMENTO: 'POSICIONAMENTO ARMAZÉM', EXECUTADO: dataHoraIso(posicionado), LOCAL_ORIGEM: escolher(LOCAIS_PATIO), LOCAL_DESTINO: escolher(DOCAS.slice(0, 5)) });
      // um posicionamento que precisou ser refeito: o KPI deve usar o primeiro
      if (rnd() < 0.04) tosMov.push({ ID_MOV: idMov++, CONTAINER: ctnr, TIPO_MOVIMENTO: 'POSICIONAMENTO ARMAZÉM', EXECUTADO: dataHoraIso(Math.min(FIM, posicionado + unif(1, 6) * HORA)), LOCAL_ORIGEM: escolher(LOCAIS_PATIO), LOCAL_DESTINO: escolher(DOCAS.slice(0, 5)) });
      if (fimDes <= FIM) tosMov.push({ ID_MOV: idMov++, CONTAINER: ctnr, TIPO_MOVIMENTO: 'RETORNO PÁTIO (VAZIO)', EXECUTADO: dataHoraIso(Math.min(FIM, fimDes + unif(0.5, 4) * HORA)), LOCAL_ORIGEM: 'ARMAZÉM', LOCAL_DESTINO: escolher(LOCAIS_PATIO) });
    }
  }
}

// Solicitações que não são de armazém e devem ser filtradas no Power Query (como no painel original)
for (let i = 0; i < 900; i++) {
  const dia = escolher(diasOperacao.filter(ehUtil));
  solicitacoes.push({
    'Nº': nSolic++, 'DATA:HORA SOLICITAÇÃO': paraExcel(horaUtil(dia, 7, 20)), 'CONTAINER': novoContainer(), 'TAM': escolher([20, 40]),
    'TS': escolher(['LIMPEZA', 'LIMPEZA / REPESAR', 'REMOÇÃO PARA OUTRO PÁTIO']), 'LOCAL ATUAL': escolher(LOCAIS_PATIO),
    'LOCAL PLANEJADO': escolher(LOCAIS_PATIO), 'SETOR': 'PÁTIO', 'SOLICITANTE': 'INTERNO', 'POSICIONADO?': 'SIM', 'STATUS': 'CONCLUÍDO',
  });
}
solicitacoes.sort((a, b) => a['DATA:HORA SOLICITAÇÃO'] - b['DATA:HORA SOLICITAÇÃO']).forEach((s, i) => { s['Nº'] = i + 1; });

// ---------- gravação ----------
async function main() {
  fs.mkdirSync(DIR_TOS, { recursive: true });
  fs.mkdirSync(DIR_SP, { recursive: true });
  console.log('Gerando bases fictícias...');

  gravarCsv('tos_clientes.csv', ['ID_CLIENTE', 'RAZAO_SOCIAL', 'SEGMENTO', 'UF'], CLIENTES);
  gravarCsv('tos_transportadoras.csv', ['ID_TRANSPORTADORA', 'NOME_TRANSPORTADORA', 'UF'], TRANSPORTADORAS);
  gravarCsv('tos_posicionamento_conferencia.csv', Object.keys(tosConf[0]), tosConf);
  gravarCsv('tos_agendamento_entrega.csv', Object.keys(tosEntrega[0]), tosEntrega);
  gravarCsv('tos_movimentacao_conteiner.csv', Object.keys(tosMov[0]), tosMov);

  const salvar = async (arquivo, montar) => {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Gerador de dados fictícios';
    const n = montar(wb);
    await wb.xlsx.writeFile(path.join(DIR_SP, arquivo));
    console.log(`  sharepoint/${arquivo}: ${n}`);
  };
  await salvar('Controle de Conferencia Aduaneira.xlsx', (wb) => abaControle(wb, 'Controle', 'CONTROLE DE CONFERÊNCIA ADUANEIRA - ARMAZÉM', 3, Object.keys(spConf[0]), spConf) + ' linhas');
  await salvar('Solicitacao de Posicionamento.xlsx', (wb) => abaControle(wb, 'POSICIONAMENTOS', 'SOLICITAÇÕES DE POSICIONAMENTO - PÁTIO x ARMAZÉM', 4, Object.keys(solicitacoes[0]), solicitacoes) + ' linhas');
  await salvar('Controle do Recebimento - Desova.xlsx', (wb) => abaControle(wb, 'Acomp_Desova', 'ACOMPANHAMENTO DE RECEBIMENTO E DESUNITIZAÇÃO', 5, Object.keys(spReceb[0]), spReceb) + ' linhas');
  await salvar('Justificativas.xlsx', (wb) => {
    const a = abaControle(wb, 'Entrega', 'JUSTIFICATIVAS - ENTREGA DE CARGA SOLTA', 2, Object.keys(justEntrega[0]), justEntrega);
    const b = abaControle(wb, 'Conferencia', 'JUSTIFICATIVAS - CONFERÊNCIA ADUANEIRA', 2, Object.keys(justConf[0]), justConf);
    return `${a} (entrega) + ${b} (conferência) linhas`;
  });
  await salvar('Parametros e Metas.xlsx', (wb) => {
    const ws = wb.addWorksheet('Metas');
    ws.columns = Object.keys(METAS[0]).map((k) => ({ header: k, key: k, width: k === 'DESCRICAO' ? 60 : 20 }));
    METAS.forEach((m) => ws.addRow(m));
    const wf = wb.addWorksheet('Feriados');
    wf.columns = [{ header: 'DATA', key: 'DATA', width: 14, style: { numFmt: 'dd/mm/yyyy' } }, { header: 'DESCRICAO', key: 'DESCRICAO', width: 34 }];
    FERIADOS.forEach(([d, n]) => wf.addRow({ DATA: new Date(`${d}T00:00:00Z`), DESCRICAO: n }));
    return `${METAS.length} metas + ${FERIADOS.length} feriados`;
  });
  console.log('Concluído.');
}
main().catch((e) => { console.error(e); process.exit(1); });
