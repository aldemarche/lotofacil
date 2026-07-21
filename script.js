'use strict';

/* =====================================================================
   FECHAMENTO ESTATÍSTICO LOTOFÁCIL
   ---------------------------------------------------------------------
   Aplicação 100% client-side. Lê o histórico completo de concursos de
   uma planilha Excel e gera 4 jogos usando:
     - Frequência Histórica
     - Frequência Recente (últimos 100 concursos)
     - Cadeia de Markov de Primeira Ordem (repetição / entrada)
   Todo o cálculo é recalculado do zero a cada novo arquivo carregado.
   ===================================================================== */

/* ---------------------------------------------------------------------
   CONFIGURAÇÃO CENTRAL DE PESOS (nada de valores fixos espalhados)
   --------------------------------------------------------------------- */
const SCORE_WEIGHTS = {
  historicalFrequency: 0.30,
  recentFrequency: 0.30,
  transitionProbability: 0.40,
};

const TOTAL_NUMBERS = 25;     // dezenas possíveis: 1 a 25
const DRAWN_PER_CONTEST = 15; // dezenas sorteadas por concurso
const RECENT_WINDOW = 100;    // janela da frequência recente

/* Número de linhas usadas como amostra para detectar automaticamente
   em qual coluna começam as 15 dezenas (ver detectNumberColumnStart). */
const COLUMN_DETECTION_SAMPLE_SIZE = 60;

/* =====================================================================
   PASSO 1 — LEITURA DA PLANILHA
   ===================================================================== */

/**
 * Lê o arquivo Excel selecionado pelo usuário e retorna a matriz bruta
 * de linhas (array de arrays), exatamente como está na planilha.
 * @param {File} file
 * @returns {Promise<Array<Array<any>>>}
 */
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        // header:1 => retorna arrays "crus", sem tentar inferir objetos,
        // o que nos permite acessar as colunas por índice fixo (A..Q).
        const rawRows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: null,
          blankrows: false,
        });

        resolve(rawRows);
      } catch (err) {
        reject(new Error('Não foi possível ler o arquivo. Verifique se é um Excel válido.'));
      }
    };

    reader.onerror = () => reject(new Error('Falha ao carregar o arquivo selecionado.'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Converte o valor de uma célula em número, ou null se não for numérico.
 * Trata strings com espaços, vírgulas decimais e valores vazios.
 * @param {any} cell
 * @returns {number|null}
 */
function cellToNumber(cell) {
  if (cell === null || cell === undefined || cell === '') return null;
  const n = typeof cell === 'string' ? Number(cell.trim().replace(',', '.')) : Number(cell);
  return Number.isFinite(n) ? n : null;
}

/**
 * Verifica se uma sequência de 15 valores representa um concurso válido:
 * 15 dezenas distintas, cada uma entre 1 e 25.
 * @param {Array<number|null>} numbers
 * @returns {boolean}
 */
function isValidDrawSequence(numbers) {
  return (
    numbers.length === DRAWN_PER_CONTEST &&
    numbers.every((n) => n !== null && n >= 1 && n <= TOTAL_NUMBERS) &&
    new Set(numbers).size === DRAWN_PER_CONTEST
  );
}

/**
 * Detecta automaticamente em qual coluna (índice, começando em 0) se
 * inicia o bloco de 15 dezenas sorteadas. Em vez de assumir uma posição
 * fixa (ex.: sempre C:Q), varremos uma amostra de linhas e procuramos,
 * em cada uma, uma janela de 15 colunas consecutivas cujos valores
 * formem um concurso válido (15 dezenas distintas entre 1 e 25).
 * A coluna inicial mais frequente entre as linhas da amostra é adotada
 * para a planilha inteira — o que também lida automaticamente com
 * linhas de cabeçalho, título ou metadados antes dos dados.
 * @param {Array<Array<any>>} rawRows
 * @returns {number|null} índice da coluna inicial, ou null se não detectada
 */
function detectNumberColumnStart(rawRows) {
  const votes = new Map();
  const sampleSize = Math.min(rawRows.length, COLUMN_DETECTION_SAMPLE_SIZE);

  for (let i = 0; i < sampleSize; i++) {
    const row = rawRows[i];
    if (!row || row.length < DRAWN_PER_CONTEST) continue;

    const maxStart = row.length - DRAWN_PER_CONTEST;
    for (let start = 0; start <= maxStart; start++) {
      const window = row.slice(start, start + DRAWN_PER_CONTEST).map(cellToNumber);
      if (isValidDrawSequence(window)) {
        votes.set(start, (votes.get(start) || 0) + 1);
      }
    }
  }

  if (votes.size === 0) return null;

  let bestStart = null;
  let bestCount = -1;
  for (const [start, count] of votes) {
    if (count > bestCount) {
      bestStart = start;
      bestCount = count;
    }
  }
  return bestStart;
}

/**
 * Converte as linhas brutas da planilha em uma lista de concursos,
 * cada um contendo exatamente 15 dezenas ordenadas de forma crescente.
 * A coluna inicial das dezenas é detectada automaticamente (ver
 * detectNumberColumnStart), então cabeçalhos, títulos e colunas extras
 * não atrapalham a leitura. Linhas vazias ou inválidas são ignoradas.
 * A última posição do array retornado é sempre o concurso mais recente,
 * pois a ordem original das linhas da planilha é preservada.
 * @param {Array<Array<any>>} rawRows
 * @returns {Array<Array<number>>}
 */
function parseContests(rawRows) {
  const startColumn = detectNumberColumnStart(rawRows);
  if (startColumn === null) return [];

  const contests = [];

  for (const row of rawRows) {
    if (!row || row.length < startColumn + DRAWN_PER_CONTEST) continue;

    const numbers = row.slice(startColumn, startColumn + DRAWN_PER_CONTEST).map(cellToNumber);

    if (!isValidDrawSequence(numbers)) continue; // ignora cabeçalho / linha inválida

    contests.push(numbers.slice().sort((a, b) => a - b));
  }

  return contests;
}

/* =====================================================================
   PASSO 2 — ÚLTIMO CONCURSO: SORTEADAS x AUSENTES
   ===================================================================== */

/**
 * @param {Array<Array<number>>} contests
 * @returns {{drawn:number[], absent:number[]}}
 */
function splitLastContest(contests) {
  const lastContest = contests[contests.length - 1];
  const drawnSet = new Set(lastContest);

  const drawn = [...lastContest].sort((a, b) => a - b);
  const absent = [];

  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    if (!drawnSet.has(n)) absent.push(n);
  }

  return { drawn, absent };
}

/* =====================================================================
   PASSO 3 — FREQUÊNCIA HISTÓRICA
   ===================================================================== */

/**
 * Retorna um mapa {dezena: frequenciaHistorica} para 1..25.
 * frequenciaHistorica = ocorrencias / totalConcursos
 * @param {Array<Array<number>>} contests
 * @returns {Map<number, number>}
 */
function calculateHistoricalFrequency(contests) {
  const occurrences = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) occurrences.set(n, 0);

  for (const contest of contests) {
    for (const number of contest) {
      occurrences.set(number, occurrences.get(number) + 1);
    }
  }

  const totalContests = contests.length;
  const frequency = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    frequency.set(n, totalContests > 0 ? occurrences.get(n) / totalContests : 0);
  }

  return frequency;
}

/* =====================================================================
   PASSO 4 — FREQUÊNCIA RECENTE (ÚLTIMOS 100 CONCURSOS)
   ===================================================================== */

/**
 * @param {Array<Array<number>>} contests
 * @returns {Map<number, number>}
 */
function calculateRecentFrequency(contests) {
  const windowSize = Math.min(RECENT_WINDOW, contests.length);
  const recentContests = contests.slice(contests.length - windowSize);

  const occurrences = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) occurrences.set(n, 0);

  for (const contest of recentContests) {
    for (const number of contest) {
      occurrences.set(number, occurrences.get(number) + 1);
    }
  }

  const frequency = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    frequency.set(n, windowSize > 0 ? occurrences.get(n) / windowSize : 0);
  }

  return frequency;
}

/* =====================================================================
   PASSO 5 e 6 — CADEIA DE MARKOV DE PRIMEIRA ORDEM
   (probabilidade de repetição e probabilidade de entrada)
   ===================================================================== */

/**
 * Percorre todos os pares consecutivos de concursos (i, i+1) e calcula,
 * para cada dezena, a probabilidade de repetir no concurso seguinte
 * dado que apareceu no concurso atual.
 * @param {Array<Array<number>>} contests
 * @returns {Map<number, number>}
 */
function calculateMarkovRepetition(contests) {
  const totalApareceu = new Map();
  const totalRepetiu = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    totalApareceu.set(n, 0);
    totalRepetiu.set(n, 0);
  }

  for (let i = 0; i < contests.length - 1; i++) {
    const current = new Set(contests[i]);
    const next = new Set(contests[i + 1]);

    for (let n = 1; n <= TOTAL_NUMBERS; n++) {
      if (current.has(n)) {
        totalApareceu.set(n, totalApareceu.get(n) + 1);
        if (next.has(n)) {
          totalRepetiu.set(n, totalRepetiu.get(n) + 1);
        }
      }
    }
  }

  const probability = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    const apareceu = totalApareceu.get(n);
    probability.set(n, apareceu > 0 ? totalRepetiu.get(n) / apareceu : 0);
  }

  return probability;
}

/**
 * Percorre todos os pares consecutivos de concursos (i, i+1) e calcula,
 * para cada dezena, a probabilidade de entrar no concurso seguinte
 * dado que estava ausente no concurso atual.
 * @param {Array<Array<number>>} contests
 * @returns {Map<number, number>}
 */
function calculateMarkovEntry(contests) {
  const totalAusente = new Map();
  const totalEntrou = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    totalAusente.set(n, 0);
    totalEntrou.set(n, 0);
  }

  for (let i = 0; i < contests.length - 1; i++) {
    const current = new Set(contests[i]);
    const next = new Set(contests[i + 1]);

    for (let n = 1; n <= TOTAL_NUMBERS; n++) {
      if (!current.has(n)) {
        totalAusente.set(n, totalAusente.get(n) + 1);
        if (next.has(n)) {
          totalEntrou.set(n, totalEntrou.get(n) + 1);
        }
      }
    }
  }

  const probability = new Map();
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    const ausente = totalAusente.get(n);
    probability.set(n, ausente > 0 ? totalEntrou.get(n) / ausente : 0);
  }

  return probability;
}

/* =====================================================================
   PASSO 7 — NORMALIZAÇÃO MIN-MAX
   ===================================================================== */

/**
 * Normaliza uma lista de valores numéricos para o intervalo [0, 1].
 * Caso todos os valores sejam iguais, retorna 1 para todos (regra da spec).
 * @param {number[]} values
 * @returns {number[]}
 */
function normalizeMetrics(values) {
  const minimo = Math.min(...values);
  const maximo = Math.max(...values);

  if (maximo === minimo) {
    return values.map(() => 1);
  }

  return values.map((valor) => (valor - minimo) / (maximo - minimo));
}

/* =====================================================================
   PASSO 8 — CÁLCULO DO SCORE
   ===================================================================== */

/**
 * Calcula o score de cada dezena de uma lista, usando os pesos definidos
 * em SCORE_WEIGHTS. A normalização Min-Max é feita dentro do próprio
 * subconjunto (sorteadas ou ausentes), já que cada grupo usa uma métrica
 * de transição diferente (repetição vs. entrada).
 * @param {number[]} numbers - dezenas a pontuar (15 sorteadas ou 10 ausentes)
 * @param {Map<number, number>} historicalFrequency
 * @param {Map<number, number>} recentFrequency
 * @param {Map<number, number>} transitionProbability - repetição ou entrada
 * @returns {Array<{numero:number, frequenciaHistorica:number, frequenciaRecente:number, probabilidadeTransicao:number, score:number}>}
 */
function calculateScores(numbers, historicalFrequency, recentFrequency, transitionProbability) {
  const historicalValues = numbers.map((n) => historicalFrequency.get(n));
  const recentValues = numbers.map((n) => recentFrequency.get(n));
  const transitionValues = numbers.map((n) => transitionProbability.get(n));

  const normalizedHistorical = normalizeMetrics(historicalValues);
  const normalizedRecent = normalizeMetrics(recentValues);
  const normalizedTransition = normalizeMetrics(transitionValues);

  return numbers.map((numero, index) => {
    const score =
      normalizedHistorical[index] * SCORE_WEIGHTS.historicalFrequency +
      normalizedRecent[index] * SCORE_WEIGHTS.recentFrequency +
      normalizedTransition[index] * SCORE_WEIGHTS.transitionProbability;

    return {
      numero,
      frequenciaHistorica: historicalValues[index],
      frequenciaRecente: recentValues[index],
      probabilidadeTransicao: transitionValues[index],
      score,
    };
  });
}

/* =====================================================================
   PASSO 9 e 10 — RANKING E DEFINIÇÃO DOS GRUPOS
   ===================================================================== */

/**
 * Ordena por score decrescente e separa em FIXAS / GRUPO A / GRUPO B
 * (para as sorteadas) ou FIXAS AUSENTES / GRUPO R / GRUPO S (ausentes).
 * @param {Array} rankedScores - já ordenado por score decrescente
 * @param {number} fixedCount
 * @param {number} groupSize
 * @returns {{fixed:number[], groupA:number[], groupB:number[]}}
 */
function splitIntoGroups(rankedScores, fixedCount, groupSize) {
  const sortedNumbers = rankedScores.map((item) => item.numero);

  return {
    fixed: sortedNumbers.slice(0, fixedCount),
    groupA: sortedNumbers.slice(fixedCount, fixedCount + groupSize),
    groupB: sortedNumbers.slice(fixedCount + groupSize, fixedCount + groupSize * 2),
  };
}

/**
 * Gera os seis grupos finais a partir dos scores já calculados.
 * @param {Array} drawnScores
 * @param {Array} absentScores
 * @returns {object} grupos nomeados: fixas, grupoA, grupoB, fixasAusentes, grupoR, grupoS
 */
function generateGroups(drawnScores, absentScores) {
  const rankedDrawn = [...drawnScores].sort((a, b) => b.score - a.score);
  const rankedAbsent = [...absentScores].sort((a, b) => b.score - a.score);

  const drawnGroups = splitIntoGroups(rankedDrawn, 3, 6);   // 3 fixas + 6 + 6 = 15
  const absentGroups = splitIntoGroups(rankedAbsent, 2, 4); // 2 fixas + 4 + 4 = 10

  return {
    fixas: drawnGroups.fixed,
    grupoA: drawnGroups.groupA,
    grupoB: drawnGroups.groupB,
    fixasAusentes: absentGroups.fixed,
    grupoR: absentGroups.groupA,
    grupoS: absentGroups.groupB,
    rankedDrawn,
    rankedAbsent,
  };
}

/* =====================================================================
   PASSO 11 — GERAÇÃO DOS 4 JOGOS
   ===================================================================== */

/**
 * Monta os 4 jogos combinando os grupos, cada um com exatamente 15 dezenas,
 * ordenados de forma crescente.
 * @param {object} groups
 * @returns {Array<number[]>} 4 jogos
 */
function generateGames(groups) {
  const buildGame = (groupA, groupR) =>
    [...groups.fixas, ...groupA, ...groups.fixasAusentes, ...groupR].sort((a, b) => a - b);

  return [
    buildGame(groups.grupoA, groups.grupoR), // Jogo 1
    buildGame(groups.grupoA, groups.grupoS), // Jogo 2
    buildGame(groups.grupoB, groups.grupoR), // Jogo 3
    buildGame(groups.grupoB, groups.grupoS), // Jogo 4
  ];
}

/* =====================================================================
   RENDERIZAÇÃO — TABELAS
   ===================================================================== */

/**
 * Formata um número entre 0 e 1 como percentual com 1 casa decimal.
 * @param {number} value
 * @returns {string}
 */
function formatPercent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

/**
 * Formata a dezena sempre com dois dígitos (ex: 3 -> "03").
 * @param {number} n
 * @returns {string}
 */
function formatNumber(n) {
  return String(n).padStart(2, '0');
}

/**
 * Renderiza as duas tabelas (sorteadas e ausentes) já ordenadas por score.
 * @param {Array} rankedDrawn
 * @param {Array} rankedAbsent
 * @param {number[]} fixedDrawn
 * @param {number[]} fixedAbsent
 */
function renderTables(rankedDrawn, rankedAbsent, fixedDrawn, fixedAbsent) {
  const maxScore = Math.max(...rankedDrawn.map((r) => r.score), ...rankedAbsent.map((r) => r.score), 0.0001);

  const buildRows = (items, fixedList) =>
    items
      .map((item) => {
        const isFixed = fixedList.includes(item.numero);
        const barWidth = Math.max(4, Math.round((item.score / maxScore) * 60));
        return `
          <tr class="${isFixed ? 'is-fixed' : ''}">
            <td>${formatNumber(item.numero)}</td>
            <td>${formatPercent(item.frequenciaHistorica)}</td>
            <td>${formatPercent(item.frequenciaRecente)}</td>
            <td>${formatPercent(item.probabilidadeTransicao)}</td>
            <td>
              <div class="score-cell">
                <span>${item.score.toFixed(3)}</span>
                <span class="score-bar" style="width:${barWidth}px"></span>
              </div>
            </td>
          </tr>`;
      })
      .join('');

  document.querySelector('#drawnTable tbody').innerHTML = buildRows(rankedDrawn, fixedDrawn);
  document.querySelector('#absentTable tbody').innerHTML = buildRows(rankedAbsent, fixedAbsent);
}

/* =====================================================================
   RENDERIZAÇÃO — MAPA DE DEZENAS (1 a 25)
   ===================================================================== */

/**
 * @param {number[]} drawn
 * @param {number[]} absent
 */
function renderBallBoard(drawn, absent) {
  const drawnSet = new Set(drawn);
  const grid = document.getElementById('ballGrid');

  let html = '';
  for (let n = 1; n <= TOTAL_NUMBERS; n++) {
    const stateClass = drawnSet.has(n) ? 'drawn' : 'absent';
    html += `<div class="ball ${stateClass}">${formatNumber(n)}</div>`;
  }
  grid.innerHTML = html;
}

/* =====================================================================
   RENDERIZAÇÃO — GRUPOS
   ===================================================================== */

/**
 * @param {object} groups
 */
function renderGroups(groups) {
  const groupDefinitions = [
    { title: 'Fixas', tone: 'tone-fixed', numbers: groups.fixas },
    { title: 'Grupo A', tone: 'tone-a', numbers: groups.grupoA },
    { title: 'Grupo B', tone: 'tone-b', numbers: groups.grupoB },
    { title: 'Fixas ausentes', tone: 'tone-fixed-abs', numbers: groups.fixasAusentes },
    { title: 'Grupo R', tone: 'tone-r', numbers: groups.grupoR },
    { title: 'Grupo S', tone: 'tone-s', numbers: groups.grupoS },
  ];

  const html = groupDefinitions
    .map(
      (group) => `
      <div class="group-card ${group.tone}">
        <h4>${group.title}</h4>
        <div class="group-numbers">
          ${group.numbers.map((n) => `<span class="chip">${formatNumber(n)}</span>`).join('')}
        </div>
      </div>`
    )
    .join('');

  document.getElementById('groupsGrid').innerHTML = html;
}

/* =====================================================================
   RENDERIZAÇÃO — JOGOS
   ===================================================================== */

/**
 * @param {Array<number[]>} games
 */
function renderGames(games) {
  const html = games
    .map(
      (game, index) => `
      <div class="game-card">
        <h4>Jogo ${index + 1}</h4>
        <div class="game-numbers">
          ${game.map((n) => `<div class="ball">${formatNumber(n)}</div>`).join('')}
        </div>
      </div>`
    )
    .join('');

  document.getElementById('gamesGrid').innerHTML = html;
}

/* =====================================================================
   RENDERIZAÇÃO — RESUMO E EXPLICAÇÃO AUTOMÁTICA
   ===================================================================== */

/**
 * @param {number} totalContests
 * @param {number} lastContestNumber
 */
function renderSummary(totalContests, lastContestNumber) {
  document.getElementById('totalContests').textContent = totalContests.toLocaleString('pt-BR');
  document.getElementById('lastContestNumber').textContent = lastContestNumber ?? '—';
}

/**
 * Monta o texto explicativo exigido pela especificação.
 * @param {number} totalContests
 */
function renderExplanation(totalContests) {
  const text = `
    Foram analisados ${totalContests.toLocaleString('pt-BR')} concursos presentes na planilha carregada.
    O modelo combina três indicadores para cada dezena: a Frequência Histórica (quantas vezes ela apareceu
    em toda a base), a Frequência Recente (comportamento nos últimos 100 concursos) e uma Cadeia de Markov
    de primeira ordem, que mede a probabilidade de repetição (para dezenas sorteadas) ou de entrada
    (para dezenas ausentes) no concurso seguinte. As três dezenas fixas do jogo são as que somaram o maior
    Score entre as dezenas sorteadas no último concurso, enquanto as duas dezenas fixas ausentes são as
    que apresentaram o maior Score de entrada entre as dezenas que não saíram. Os demais grupos (A, B, R e S)
    foram definidos automaticamente a partir do mesmo ranking estatístico, e os 4 jogos combinam esses grupos
    para formar diferentes fechamentos de 15 dezenas cada.
  `.replace(/\s+/g, ' ').trim();

  document.getElementById('explanationText').textContent = text;
}

/* =====================================================================
   ORQUESTRAÇÃO PRINCIPAL
   ===================================================================== */

/**
 * Executa o pipeline completo de análise a partir da lista de concursos
 * já convertida (parseContests) e atualiza toda a interface.
 * @param {Array<Array<number>>} contests
 */
function runAnalysis(contests) {
  if (contests.length < 2) {
    showStatus('A planilha precisa conter pelo menos 2 concursos válidos para calcular a Cadeia de Markov.', 'error');
    return;
  }

  const { drawn, absent } = splitLastContest(contests);

  const historicalFrequency = calculateHistoricalFrequency(contests);
  const recentFrequency = calculateRecentFrequency(contests);
  const repetitionProbability = calculateMarkovRepetition(contests);
  const entryProbability = calculateMarkovEntry(contests);

  const drawnScores = calculateScores(drawn, historicalFrequency, recentFrequency, repetitionProbability);
  const absentScores = calculateScores(absent, historicalFrequency, recentFrequency, entryProbability);

  const groups = generateGroups(drawnScores, absentScores);
  const games = generateGames(groups);

  const lastContestRow = contests.length; // posição sequencial (linha) do último concurso lido

  renderSummary(contests.length, lastContestRow);
  renderBallBoard(drawn, absent);
  renderTables(groups.rankedDrawn, groups.rankedAbsent, groups.fixas, groups.fixasAusentes);
  renderGroups(groups);
  renderGames(games);
  renderExplanation(contests.length);

  document.getElementById('resultsWrapper').classList.remove('hidden');
  showStatus(`Planilha processada com sucesso: ${contests.length} concursos analisados.`, 'success');
}

/* =====================================================================
   INTERFACE — UPLOAD E MENSAGENS DE STATUS
   ===================================================================== */

/**
 * Exibe uma mensagem de status para o usuário (erro, sucesso ou carregando).
 * @param {string} message
 * @param {'error'|'success'|'loading'} type
 */
function showStatus(message, type) {
  const el = document.getElementById('statusMessage');
  el.textContent = message;
  el.className = `status-message ${type}`;
}

/**
 * Processa o arquivo selecionado: lê, converte e roda a análise completa.
 * Trata erros de arquivo inválido ou dados insuficientes.
 * @param {File} file
 */
async function handleFileSelected(file) {
  if (!file) return;

  const validExtension = /\.(xlsx|xls)$/i.test(file.name);
  if (!validExtension) {
    showStatus('Arquivo inválido. Selecione uma planilha Excel (.xlsx ou .xls).', 'error');
    return;
  }

  showStatus('Lendo planilha...', 'loading');

  try {
    const rawRows = await readExcel(file);
    const contests = parseContests(rawRows);

    if (contests.length === 0) {
      showStatus(
        `Não foi possível identificar as 15 dezenas em nenhuma linha (${rawRows.length} linhas lidas). ` +
        'Confirme que a planilha tem uma linha por concurso, com 15 dezenas (1 a 25) em colunas consecutivas.',
        'error'
      );
      document.getElementById('resultsWrapper').classList.add('hidden');
      return;
    }

    runAnalysis(contests);
  } catch (err) {
    showStatus(err.message || 'Ocorreu um erro inesperado ao processar o arquivo.', 'error');
    document.getElementById('resultsWrapper').classList.add('hidden');
  }
}

/* =====================================================================
   EVENTOS DE INTERFACE
   ===================================================================== */

function initUploadEvents() {
  const fileInput = document.getElementById('fileInput');
  const dropzone = document.getElementById('dropzone');

  fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    handleFileSelected(file);
  });

  // Suporte a arrastar-e-soltar, mantendo a mesma validação e pipeline.
  ['dragover', 'dragenter'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (event) => {
    const file = event.dataTransfer.files[0];
    handleFileSelected(file);
  });
}

document.addEventListener('DOMContentLoaded', initUploadEvents);
