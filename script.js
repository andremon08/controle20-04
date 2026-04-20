// ─── CONFIGURAÇÃO ───────────────────────────────────────────────────────────
const BASE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSJq1BdeNlo6gvM1vBhtgD88MRevuRrODf2NmVESwH5CMQ6VBkuZMUaNEr8xCoHeJlmnlsJaDV_Cj9L/pub';

const URL_VERBAS     = BASE_URL + '?gid=1303157015&single=true&output=csv';
const URL_SERVIDORES = BASE_URL + '?gid=1533392322&single=true&output=csv';
const URL_CARGOS     = BASE_URL + '?gid=1823673227&single=true&output=csv';
const URL_ESTRUTURAS = BASE_URL + '?gid=46958645&single=true&output=csv';

// ─── REGRAS DE NEGÓCIO ──────────────────────────────────────────────────────
const TETO_VEREADOR  = 92998.45;
const TOLERANCIA     = 0.13;
const MAX_SERVIDORES = 9;

// Cargo especial: não consome vaga, verba nem aparece na estrutura
const CARGO_CEDIDO = 'CEDIDOS DE OUTRAS ENTIDADES SEM ÔNUS';
const isCedido = cargo => cargo.trim().toUpperCase().includes('CEDIDOS DE OUTRAS ENTIDADES');

// Lista exata de lotações especiais com estrutura fixa definida por lei.
// TUDO que não estiver aqui e não começar com "Bloco" ou "Liderança" = Gabinete de Vereador.
const LOTACOES_ESPECIAIS = {
    'GABINETE DA PRESIDÊNCIA':         ['CC-1', 'CC-5', 'CC-6', 'CC-7'],
    'GABINETE DA 1ª VICE-PRESIDÊNCIA': ['CC-4', 'CC-6'],
    'GABINETE DA 2ª VICE-PRESIDÊNCIA': ['CC-4', 'CC-7'],
    'GABINETE DA 1ª SECRETARIA':       ['CC-3', 'CC-6', 'CC-7'],
    'GABINETE DA 2ª SECRETARIA':       ['CC-4', 'CC-6', 'CC-7'],
    'GABINETE DA 3ª SECRETARIA':       ['CC-5', 'CC-7'],
    'GABINETE DA 4ª SECRETARIA':       ['CC-5', 'CC-7'],
};

// ─── ESTADO GLOBAL ──────────────────────────────────────────────────────────
let dadosVerbas     = [];
let dadosServidores = [];
let tabelaCargos    = {};  // { 'CC-1': 12000.00, ... }
let dadosEstruturas = {};  // { 'Gabinete X': ['CC-1','CC-3','CC-6'], ... }
let _todasSugestoes = []; // sugestões geradas para o filtro em tempo real
let saldo_atual     = 0;  // saldo do gabinete atual, usado no filtro

// Estado atual para exportação
let _exportEstado = {
    mes: '', gab: '', tipo: '',
    servidores: [], estrutura: [], responsavel: ''
};

// ─── INICIALIZAÇÃO ──────────────────────────────────────────────────────────
function iniciar() {
    setStatus('', 'Carregando...');

    Promise.all([
        carregarCSV(URL_VERBAS,     'verbas'),
        carregarCSV(URL_SERVIDORES, 'servidores'),
        carregarCSV(URL_CARGOS,     'cargos'),
        carregarCSV(URL_ESTRUTURAS, 'estruturas'),
    ])
    .then(([verbas, servidores, cargos, estruturas]) => {
        try {
            dadosVerbas     = verbas;
            dadosServidores = servidores;
            tabelaCargos    = construirTabelaCargos(cargos);
            dadosEstruturas = construirEstruturas(estruturas);
            preencherFiltros();
            setStatus('ok', 'Dados carregados');
        } catch (err) {
            console.error('[Erro ao processar dados]', err);
            setStatus('erro', 'Erro ao processar dados');
        }
    })
    .catch(err => {
        console.error('[Erro ao carregar]', err);
        setStatus('erro', typeof err === 'string' ? err : 'Erro ao carregar dados');
    });
}

function carregarCSV(url, nome) {
    return new Promise((resolve, reject) => {
        // Timeout de 15s para não ficar pendurado
        const timer = setTimeout(() => {
            reject(`Timeout ao carregar "${nome}". Verifique se a planilha está publicada.`);
        }, 15000);

        Papa.parse(url, {
            download: true,
            header: true,
            skipEmptyLines: true,
            complete: r => {
                clearTimeout(timer);
                if (!r.data || r.data.length === 0) {
                    console.warn(`[${nome}] Aba vazia ou sem dados.`);
                }
                resolve(r.data || []);
            },
            error: e => {
                clearTimeout(timer);
                console.error(`[${nome}] Erro PapaParse:`, e);
                // Resolve com array vazio em vez de rejeitar, para não travar tudo
                // caso uma aba ainda não tenha dados
                resolve([]);
            },
        });
    });
}

function construirTabelaCargos(linhas) {
    const tabela = {};
    linhas.forEach(l => {
        const cargo   = (l['Cargo'] || '').trim();
        const salario = parseMoeda(l['Salário'] || l['Salario'] || '0');
        if (cargo) tabela[cargo] = salario;
    });
    return tabela;
}

// Aba estruturas: colunas Gabinete | Cargo
// Uma linha por cargo da estrutura (com repetições para mesmo CC)
function construirEstruturas(linhas) {
    const estruturas = {};
    linhas.forEach(l => {
        const gab   = (l['Gabinete'] || '').trim();
        const cargo = (l['Cargo'] || '').trim();
        if (!gab || !cargo) return;
        // Extrai código CC do nome completo (ex: "CHEFE DE GABINETE - CC-1" → "CC-1")
        const match = cargo.toUpperCase().match(/CC-\d+/);
        const cc = match ? match[0] : cargo.toUpperCase();
        if (!estruturas[gab]) estruturas[gab] = [];
        estruturas[gab].push(cc);
    });
    return estruturas;
}

// ─── CLASSIFICAÇÃO ──────────────────────────────────────────────────────────
// Retorna 'mesa_diretora', 'bloco' ou 'vereador'
function classificarTipo(gabinete) {
    const g = gabinete.trim().toUpperCase();
    // Verifica lista exata de lotações especiais (case-insensitive)
    for (const nome of Object.keys(LOTACOES_ESPECIAIS)) {
        if (g === nome.toUpperCase()) return 'mesa_diretora';
    }
    // Blocos e lideranças: nome contém essas palavras
    if (/\bbloco\b/i.test(g) || /\blideran/i.test(g)) return 'bloco';
    // Todo o resto = gabinete de vereador
    return 'vereador';
}

// ─── FILTROS ────────────────────────────────────────────────────────────────
function preencherFiltros() {
    const mesSelect = document.getElementById('mesSelect');
    const gabSelect = document.getElementById('gabineteSelect');

    const mesesSet = new Set();
    const gabSet   = new Set();

    dadosVerbas.forEach(l => {
        if (l['Mês'])      mesesSet.add(l['Mês'].trim());
        if (l['Gabinete']) gabSet.add(l['Gabinete'].trim());
    });

    const mesesOrdenados = [...mesesSet].sort((a, b) => {
        const [ma, aa] = a.split('/').map(Number);
        const [mb, ab] = b.split('/').map(Number);
        return aa !== ab ? aa - ab : ma - mb;
    });
    const gabsOrdenados = [...gabSet].sort((a, b) => a.localeCompare(b, 'pt-BR'));

    mesSelect.innerHTML = '<option value="">Selecione o mês...</option>';
    gabSelect.innerHTML = '<option value="">Selecione a lotação...</option>';

    mesesOrdenados.forEach(m => mesSelect.innerHTML += `<option value="${m}">${m}</option>`);
    gabsOrdenados.forEach(g  => gabSelect.innerHTML += `<option value="${g}">${g}</option>`);

    mesSelect.addEventListener('change', atualizarPainel);
    gabSelect.addEventListener('change', atualizarPainel);
}


// ─── ATUALIZAÇÃO PRINCIPAL ──────────────────────────────────────────────────
function atualizarPainel() {
    const mes = document.getElementById('mesSelect').value.trim();
    const gab = document.getElementById('gabineteSelect').value.trim();

    ocultarTudo();
    if (!mes || !gab) return;

    const { inicio, fim } = intervaloMes(mes);

    const verbaMes = dadosVerbas.find(
        l => l['Mês']?.trim() === mes && l['Gabinete']?.trim() === gab
    ) || {};

    // Filtra servidores ativos no mês via datas
    const servidoresMes = dadosServidores
        .filter(l => (l['Gabinete'] || '').trim() === gab)
        .map(l => {
            const admissao       = parseData(l['Admissão']   || l['Admissao']   || '');
            const exoneracao     = parseData(l['Exoneração'] || l['Exoneracao'] || '');
            const ativo          = estaAtivo(admissao, exoneracao, inicio, fim);
            const exoneradoNoMes = !!(exoneracao && exoneracao >= inicio && exoneracao <= fim);
            return { ...l, admissao, exoneracao, ativo, exoneradoNoMes };
        })
        .filter(l => l.ativo);

    const tipo = classificarTipo(gab);
    atualizarTopbarBadges(mes, tipo);

    // Salva estado para exportação
    _exportEstado.mes  = mes;
    _exportEstado.gab  = gab;
    _exportEstado.tipo = tipo;

    if (tipo === 'vereador') {
        renderizarVereador(gab, verbaMes, servidoresMes);
    } else {
        renderizarEspecial(gab, tipo, verbaMes, servidoresMes);
    }
}

// ─── TOPBAR BADGES ──────────────────────────────────────────────────────────
function atualizarTopbarBadges(mes, tipo) {
    let tipoBadge = '';
    if (tipo === 'vereador')       tipoBadge = `<span class="badge badge-tipo-vereador">Gabinete de Vereador</span>`;
    else if (tipo === 'mesa_diretora') tipoBadge = `<span class="badge badge-tipo-especial">Mesa Diretora</span>`;
    else                           tipoBadge = `<span class="badge badge-tipo-especial">Bloco / Liderança</span>`;
    document.getElementById('topbarBadges').innerHTML =
        `<span class="badge badge-mes">${mes}</span>${tipoBadge}`;
}

// ─── PAINEL VEREADOR ────────────────────────────────────────────────────────
function renderizarVereador(gab, verbaMes, servidores) {
    document.getElementById('painelVereador').classList.remove('escondido');

    const responsavel = (verbaMes['Responsável'] || verbaMes['Responsavel'] || '').trim();

    // Separa servidores ativos (não exonerados) dos exonerados no mês
    // Cedidos de outras entidades sem ônus são excluídos de verba, contagem e estrutura
    const servsAtivos     = servidores.filter(s => !s.exoneradoNoMes && !isCedido(s['Cargo'] || ''));
    const servsExonerados = servidores.filter(s => s.exoneradoNoMes);

    // Verba utilizada considera apenas ativos (exonerado = vaga liberada)
    let verbaUtil = 0;
    servsAtivos.forEach(s => {
        const cargo = (s['Cargo'] || '').trim();
        verbaUtil += tabelaCargos[cargo] || 0;
    });

    // ── Estrutura do gabinete ──
    const estrutura = dadosEstruturas[gab] || [];
    const totalVagasEstrutura = estrutura.length;

    const saldo  = TETO_VEREADOR - verbaUtil;
    const pct    = TETO_VEREADOR > 0 ? Math.min((verbaUtil / TETO_VEREADOR) * 100, 100) : 0;
    const nServs = servsAtivos.filter(s => (s['Nome do Servidor'] || '').trim()).length;

    // Vagas baseadas na estrutura; exonerados no mês liberam vagas
    const totalRef = totalVagasEstrutura > 0 ? totalVagasEstrutura : MAX_SERVIDORES;
    const vagasLiv = totalRef - nServs;

    document.getElementById('vVerbaTotal').textContent    = moeda(TETO_VEREADOR);
    document.getElementById('vVerbaUtil').textContent     = moeda(verbaUtil);
    document.getElementById('vSaldo').textContent         = moeda(Math.abs(saldo));
    document.getElementById('vServidores').textContent    = `${nServs} / ${totalRef}`;
    document.getElementById('vServidoresSub').textContent = vagasLiv === 1 ? '1 vaga disponível' : vagasLiv > 1 ? `${vagasLiv} vagas disponíveis` : 'sem vagas disponíveis';
    document.getElementById('vVerbaUtilSub').textContent  = responsavel ? `Resp.: ${responsavel}` : 'soma dos salários';

    const cardSaldo = document.getElementById('vSaldo').closest('.card');
    cardSaldo.classList.remove('card-saldo-negativo');
    if (saldo < 0) {
        cardSaldo.classList.add('card-saldo-negativo');
        document.getElementById('vSaldoSub').textContent = 'acima do teto';
    } else {
        document.getElementById('vSaldoSub').textContent = 'disponível no teto';
    }

    // Barra de progresso
    const fill  = document.getElementById('vProgressoFill');
    const pctEl = document.getElementById('vProgressoPct');
    fill.style.width  = pct.toFixed(1) + '%';
    pctEl.textContent = pct.toFixed(1) + '%';
    fill.classList.remove('aviso', 'perigo');
    if (pct >= 100)     fill.classList.add('perigo');
    else if (pct >= 85) fill.classList.add('aviso');

    // Alertas
    const temCC1 = servsAtivos.some(s => {
        const c = (s['Cargo'] || '').trim().toUpperCase();
        return (c.endsWith('CC-1') || c.endsWith('- CC-1') || c === 'CC-1');
    });
    const alertas = [];
    if (!temCC1) alertas.push({ tipo: 'erro', msg: 'Chefe de Gabinete (CC-1) não está lotado. Este cargo é obrigatório.' });
    if (nServs > MAX_SERVIDORES) alertas.push({ tipo: 'erro', msg: `O gabinete possui ${nServs} servidores, acima do limite legal de ${MAX_SERVIDORES}.` });
    if (verbaUtil > TETO_VEREADOR + TOLERANCIA) {
        alertas.push({ tipo: 'erro', msg: `Verba utilizada (${moeda(verbaUtil)}) excede o teto legal de ${moeda(TETO_VEREADOR)}.` });
    } else if (verbaUtil > TETO_VEREADOR) {
        alertas.push({ tipo: 'aviso', msg: 'Verba dentro da margem de tolerância de R$ 0,13.' });
    }
    if (servsExonerados.length > 0) {
        const nomes = servsExonerados.map(s => (s['Nome do Servidor'] || '').trim()).join(', ');
        alertas.push({ tipo: 'aviso', msg: `Exonerado(s) neste mês: ${nomes}.` });
    }
    if (alertas.length === 0) alertas.push({ tipo: 'ok', msg: 'Gabinete em conformidade com as regras legais.' });
    renderizarAlertas('alertasVereador', alertas);

    // ── Estrutura com cards agrupados ──
    const resumoEl = document.getElementById('vEstruturaResumo');
    const grade    = document.getElementById('gradeEstrutura');

    if (estrutura.length === 0) {
        grade.innerHTML = `<p style="font-size:14px;color:var(--muted);font-style:italic;padding:4px 0">Estrutura não cadastrada para este gabinete.</p>`;
        resumoEl.textContent = '';
    } else {
        // Constrói lista de CCs ativos (cópia fresca, não consumida)
        const cargosAtivosCC = servsAtivos.map(s => extrairCC(s['Cargo'] || ''));

        // Constrói estrutura efetiva: começa com os slots formais e adiciona
        // slots extras se houver mais servidores ativos do que slots na estrutura
        const estruturaEfetiva = [...estrutura];
        const consumivelCheck = [...cargosAtivosCC];
        // Marca quais slots da estrutura formal estão ocupados
        estruturaEfetiva.forEach((cc, i) => {
            const idx = consumivelCheck.indexOf(cc);
            if (idx !== -1) consumivelCheck.splice(idx, 1);
        });
        // Sobram em consumivelCheck os servidores que não têm slot na estrutura — adiciona como extra
        consumivelCheck.forEach(cc => estruturaEfetiva.push(cc));

        // Conta vagos (slots não preenchidos)
        const consumivelVagos = [...cargosAtivosCC];
        const vagosCount = estruturaEfetiva.filter(cc => {
            const idx = consumivelVagos.indexOf(cc);
            if (idx !== -1) { consumivelVagos.splice(idx, 1); return false; }
            return true;
        }).length;

        const totalSlots = estruturaEfetiva.length;
        resumoEl.textContent = vagosCount === 1
            ? `1 vaga livre de ${totalSlots}`
            : vagosCount > 1
                ? `${vagosCount} vagas livres de ${totalSlots}`
                : `${totalSlots} de ${totalSlots} preenchidos`;

        grade.innerHTML = renderizarEstruturaCCs(estruturaEfetiva, [...cargosAtivosCC]);
    }

    // ── Tabela servidores ──
    const tbody = document.getElementById('corpoTabelaVereador');
    const tfoot = document.getElementById('rodapeTabelaVereador');
    tbody.innerHTML = '';

    // Verba total tabela inclui todos (ativos + exonerados no mês para histórico)
    let verbaTotalTabela = 0;
    servidores.forEach(s => {
        const nome      = (s['Nome do Servidor'] || '').trim();
        const cargo     = (s['Cargo'] || '').trim();
        const matricula = (s['Matrícula'] || s['Matricula'] || '').trim();
        const sal       = tabelaCargos[cargo] || 0;
        const admStr    = s.admissao   ? formatarData(s.admissao)   : '—';
        const exoStr    = s.exoneracao ? formatarData(s.exoneracao) : '—';
        if (!nome) return;
        if (!s.exoneradoNoMes) verbaTotalTabela += sal;
        const exoTag = s.exoneradoNoMes ? `<span class="tag-exonerado">Exonerado</span>` : '';
        tbody.innerHTML += `
            <tr class="${s.exoneradoNoMes ? 'tr-exonerado' : ''}">
                <td class="col-matricula">${matricula || '—'}</td>
                <td>${nome}${exoTag}</td>
                <td>${cargo || '—'}</td>
                <td class="col-salario">${sal > 0 ? moeda(sal) : '—'}</td>
                <td class="col-data">${admStr}</td>
                <td class="col-data">${exoStr}</td>
            </tr>`;
    });

    tfoot.innerHTML = `
        <tr>
            <td colspan="3">Total</td>
            <td class="col-salario col-salario-total">${moeda(verbaUtil)}</td>
            <td colspan="2"></td>
        </tr>`;

    // ── Sugestão de Composição ──
    renderizarSugestao(saldo, nServs);

    // Salva estado para exportação
    _exportEstado.servidores  = servsAtivos;
    _exportEstado.estrutura   = dadosEstruturas[gab] || [];
    _exportEstado.responsavel = responsavel;
}

// ─── PAINEL ESPECIAL ────────────────────────────────────────────────────────
function renderizarEspecial(gab, tipo, verbaMes, servidores) {
    document.getElementById('painelEspecial').classList.remove('escondido');

    const responsavel = (verbaMes['Responsável'] || verbaMes['Responsavel'] || '').trim();
    const elResp       = document.getElementById('eResponsavel');
    const elDestaque   = document.getElementById('eResponsavelDestaque');
    elResp.textContent = responsavel || 'Não informado';
    if (responsavel) {
        elDestaque.style.display = 'flex';
    } else {
        elDestaque.style.display = 'none';
    }

    // Estrutura esperada
    let estrutura = [];
    if (tipo === 'mesa_diretora') {
        const chave = Object.keys(LOTACOES_ESPECIAIS).find(
            k => k.toUpperCase() === gab.trim().toUpperCase()
        );
        estrutura = chave ? LOTACOES_ESPECIAIS[chave] : [];
    } else {
        estrutura = ['CC-8'];
    }

    // Alertas
    const servsAtivosEspecial = servidores.filter(s => !s.exoneradoNoMes);
    const ccsAtivos = servsAtivosEspecial.map(s => extrairCC(s['Cargo'] || ''));
    const vagasLivres = estrutura.filter(cc => {
        const idx = ccsAtivos.indexOf(cc);
        if (idx !== -1) { ccsAtivos.splice(idx, 1); return false; }
        return true;
    });
    const alertas = [];
    if (tipo === 'bloco') {
        const blocoAtivo = servsAtivosEspecial.some(s => extrairCC(s['Cargo'] || '') === 'CC-8');
        if (!blocoAtivo) alertas.push({ tipo: 'aviso', msg: 'O cargo CC-8 desta lotação está vago.' });
        else             alertas.push({ tipo: 'ok',   msg: 'Lotação regularmente ocupada.' });
    } else {
        if (vagasLivres.length === estrutura.length)
            alertas.push({ tipo: 'aviso', msg: 'Nenhum cargo desta lotação está ocupado.' });
        else if (vagasLivres.length > 0)
            alertas.push({ tipo: 'aviso', msg: `${vagasLivres.length} cargo(s) com vaga disponível: ${vagasLivres.join(', ')}.` });
        else
            alertas.push({ tipo: 'ok', msg: 'Todos os cargos da estrutura estão ocupados.' });
    }
    const exoneradosNoMes = servidores.filter(s => s.exoneradoNoMes);
    if (exoneradosNoMes.length > 0) {
        const nomes = exoneradosNoMes.map(s => (s['Nome do Servidor'] || '').trim()).join(', ');
        alertas.push({ tipo: 'aviso', msg: `Exonerado(s) neste mês: ${nomes}.` });
    }
    renderizarAlertas('alertasEspecial', alertas);

    // ── Tabela primeiro ──
    const tbody = document.getElementById('corpoTabelaEspecial');
    tbody.innerHTML = '';
    servidores.forEach(s => {
        const nome      = (s['Nome do Servidor'] || '').trim();
        const cargo     = (s['Cargo'] || '').trim();
        const matricula = (s['Matrícula'] || s['Matricula'] || '').trim();
        const sal       = tabelaCargos[cargo] || 0;
        const admStr    = s.admissao   ? formatarData(s.admissao)   : '—';
        const exoStr    = s.exoneracao ? formatarData(s.exoneracao) : '—';
        if (!nome) return;
        const exoTag = s.exoneradoNoMes ? `<span class="tag-exonerado">Exonerado</span>` : '';
        tbody.innerHTML += `
            <tr class="${s.exoneradoNoMes ? 'tr-exonerado' : ''}">
                <td class="col-matricula">${matricula || '—'}</td>
                <td>${nome}${exoTag}</td>
                <td>${cargo || '—'}</td>
                <td class="col-salario">${sal > 0 ? moeda(sal) : '—'}</td>
                <td class="col-data">${admStr}</td>
                <td class="col-data">${exoStr}</td>
            </tr>`;
    });
    if (!tbody.innerHTML) {
        tbody.innerHTML = `<tr><td colspan="6" style="color:var(--muted);font-style:italic;text-align:center;padding:16px">Nenhum servidor lotado neste período</td></tr>`;
    }

    // ── Estrutura da Lotação com cards ──
    const resumoEl = document.getElementById('eEstruturaResumo');
    const grade    = document.getElementById('gradeEspecial');

    if (estrutura.length === 0) {
        grade.innerHTML = `<p style="font-size:14px;color:var(--muted);font-style:italic;padding:4px 0">Estrutura não definida para esta lotação.</p>`;
        resumoEl.textContent = '';
    } else {
        // Reconstrói cargos ativos para o render (ccsAtivos foi consumido acima)
        const cargosAtivosRender = servsAtivosEspecial.map(s => extrairCC(s['Cargo'] || ''));

        // Conta vagos para o resumo
        const tempVagos = [...cargosAtivosRender];
        const vagosCount = estrutura.filter(cc => {
            const idx = tempVagos.indexOf(cc);
            if (idx !== -1) { tempVagos.splice(idx, 1); return false; }
            return true;
        }).length;

        resumoEl.textContent = vagosCount === 1
            ? `1 vaga livre de ${estrutura.length}`
            : vagosCount > 1
                ? `${vagosCount} vagas livres de ${estrutura.length}`
                : `${estrutura.length} de ${estrutura.length} preenchidos`;

        grade.innerHTML = renderizarEstruturaCCs(estrutura, cargosAtivosRender);
    }

    // Salva estado para exportação
    _exportEstado.servidores  = servsAtivosEspecial;
    _exportEstado.estrutura   = estrutura;
    _exportEstado.responsavel = responsavel;
}

// ─── SUGESTÃO DE COMPOSIÇÃO ─────────────────────────────────────────────────
function renderizarSugestao(saldo, nServsAtivos) {
    saldo_atual = saldo; // salva para uso no filtro em tempo real
    const secao      = document.getElementById('secaoSugestao');
    const introEl    = document.getElementById('sugestaoIntro');
    const gridEl     = document.getElementById('sugestaoGrid');
    const resumo     = document.getElementById('sugestaoResumo');
    const filtroWrap = document.getElementById('sugestaoFiltroWrap');
    const filtroInput = document.getElementById('sugestaoFiltro');

    const vagasRestantes = MAX_SERVIDORES - nServsAtivos;

    if (vagasRestantes <= 0 || saldo <= 0) {
        secao.classList.add('escondido');
        return;
    }
    secao.classList.remove('escondido');

    const vagasStr = vagasRestantes === 1 ? '1 vaga disponível' : `${vagasRestantes} vagas disponíveis`;
    resumo.textContent = `${vagasStr} — saldo ${moeda(saldo)}`;

    // CCs disponíveis excluindo CC-1
    const ccMap = {};
    Object.entries(tabelaCargos).forEach(([nome, sal]) => {
        const cc = extrairCC(nome);
        if (!cc.startsWith('CC-') || cc === 'CC-1') return;
        if (!ccMap[cc] || sal < ccMap[cc]) ccMap[cc] = sal;
    });
    const ccs = Object.entries(ccMap)
        .map(([cc, sal]) => ({ cc, sal }))
        .sort((a, b) => a.sal - b.sal);

    if (ccs.length === 0) {
        introEl.innerHTML = `<p class="sugestao-intro">Tabela de cargos não carregada.</p>`;
        filtroWrap.style.display = 'none';
        return;
    }

    // Gera todas as combinações com repetição de tamanho 1..vagasRestantes
    const sugestoes = [];
    const vistos = new Set();

    function combinar(inicio, atual, custoAtual) {
        if (atual.length > 0) {
            const chave = [...atual].sort().join('+');
            if (!vistos.has(chave)) {
                vistos.add(chave);
                sugestoes.push({ ccs: [...atual], custo: custoAtual });
            }
        }
        if (atual.length === vagasRestantes) return;
        for (let i = inicio; i < ccs.length; i++) {
            const novoCusto = custoAtual + ccs[i].sal;
            if (novoCusto > saldo + TOLERANCIA) break;
            combinar(i, [...atual, ccs[i].cc], novoCusto);
        }
    }
    combinar(0, [], 0);

    if (sugestoes.length === 0) {
        introEl.innerHTML = `<p class="sugestao-intro">Não há cargos que caibam no saldo disponível de ${moeda(saldo)}.</p>`;
        filtroWrap.style.display = 'none';
        gridEl.innerHTML = ''; // limpa cards de gabinete anterior
        _todasSugestoes = [];
        return;
    }

    // Ordena: mais cargos primeiro, depois por custo decrescente
    sugestoes.sort((a, b) =>
        b.ccs.length !== a.ccs.length ? b.ccs.length - a.ccs.length : b.custo - a.custo
    );
    _todasSugestoes = sugestoes;

    const vagasIntro = vagasRestantes === 1 ? '1 vaga disponível' : `${vagasRestantes} vagas disponíveis`;
    introEl.innerHTML = `<p class="sugestao-intro">
        Com <strong>${vagasIntro}</strong> e saldo de <strong>${moeda(saldo)}</strong>,
        abaixo estão todas as combinações possíveis dentro do teto legal de ${moeda(TETO_VEREADOR)}:
    </p>`;

    // Mostra filtro e reseta
    filtroWrap.style.display = 'block';
    filtroInput.value = '';

    renderizarCardsSugestao('');

    // Filtro em tempo real — remove e recria o listener para evitar duplicatas
    const novoInput = filtroInput.cloneNode(true);
    filtroInput.parentNode.replaceChild(novoInput, filtroInput);
    novoInput.addEventListener('input', () => {
        renderizarCardsSugestao(novoInput.value.trim());
    });
}

function renderizarCardsSugestao(filtro) {
    const gridEl = document.getElementById('sugestaoGrid');

    // Suporte a múltiplos termos separados por vírgula
    // Ex: "CC2, CC3" → filtra cards que contenham CC-2 E CC-3
    const termos = filtro
        .split(',')
        .map(t => t.trim().toUpperCase().replace(/\s/g, ''))
        .filter(t => t.length > 0);

    const normalizar = cc => cc.replace('-', '');

    const filtradas = termos.length === 0
        ? _todasSugestoes
        : _todasSugestoes.filter(s => {
            // Cada termo deve estar presente pelo menos uma vez na combinação
            const ccsNorm = s.ccs.map(normalizar);
            return termos.every(termo => {
                const termoNorm = termo.replace('-', '');
                return ccsNorm.some(cc => cc.includes(termoNorm));
            });
        });

    if (filtradas.length === 0) {
        const termosLabel = termos.join(', ');
        gridEl.innerHTML = `<p class="sugestao-sem-resultado">Nenhuma combinação encontrada para "${termosLabel}".</p>`;
        return;
    }

    gridEl.innerHTML = filtradas.map(s => {
        const n = s.ccs.length;
        const saldoPos = saldo_atual - s.custo;
        const contagem = {};
        s.ccs.forEach(cc => { contagem[cc] = (contagem[cc] || 0) + 1; });
        const ccsLabel = Object.entries(contagem)
            .map(([cc, qtd]) => qtd > 1 ? `${qtd}× ${cc}` : cc)
            .join(' + ');
        const cargoStr = n === 1 ? '1 cargo' : `${n} cargos`;
        return `
            <div class="sugestao-card">
                <div class="sugestao-card-titulo">${cargoStr}</div>
                <div class="sugestao-card-ccs">${ccsLabel}</div>
                <div class="sugestao-card-total">+ ${moeda(s.custo)}</div>
                <div class="sugestao-card-saldo">Saldo restante: ${moeda(saldoPos)}</div>
            </div>`;
    }).join('');
}



// Extrai o código CC do nome completo do cargo
// "CHEFE DE GABINETE PARLAMENTAR - CC-1" → "CC-1"
function extrairCC(nomeCargo) {
    const match = nomeCargo.trim().toUpperCase().match(/CC-\d+/);
    return match ? match[0] : nomeCargo.trim().toUpperCase();
}

// Renderiza a estrutura do gabinete agrupada por CC com disposição triangular
function renderizarEstruturaCCs(estrutura, cargosAtivosCC) {
    const consumivel = [...cargosAtivosCC];

    const slots = estrutura.map(cc => {
        const idx = consumivel.indexOf(cc);
        if (idx !== -1) { consumivel.splice(idx, 1); return { cc, estado: 'ocupado' }; }
        return { cc, estado: 'vago' };
    });

    // Agrupa por CC preservando ordem de primeiro aparecimento
    const ordem = [];
    const grupos = {};
    slots.forEach(s => {
        if (!grupos[s.cc]) { grupos[s.cc] = []; ordem.push(s.cc); }
        grupos[s.cc].push(s.estado);
    });

    // Número de linhas = máximo de repetições de qualquer CC
    const maxSlots = Math.max(...ordem.map(cc => grupos[cc].length));

    // Grid: 1 coluna por tipo de CC, gap uniforme entre todos
    // Cada CC ocupa sempre a mesma coluna, repetições vão para linhas abaixo
    const numCols = ordem.length;
    const templateCols = Array(numCols).fill('102px').join(' ');

    let cellsHTML = '';
    ordem.forEach((cc, colIdx) => {
        const n = grupos[cc].length;
        // Preenche os slots ocupados/vagos
        for (let row = 0; row < n; row++) {
            const estado = grupos[cc][row];
            cellsHTML += `<div class="cc-card ${estado}" style="grid-column:${colIdx+1};grid-row:${row+1}">
                <span class="cc-card-label">${cc}</span>
                <span class="cc-card-status">${estado === 'ocupado' ? '● ocupado' : '○ vago'}</span>
            </div>`;
        }
        // Linhas acima do máximo ficam vazias (não precisam de placeholder — grid-template-rows cuida)
    });

    return `<div class="grade-cc-grid" style="grid-template-columns:${templateCols};grid-template-rows:repeat(${maxSlots},82px)">${cellsHTML}</div>`;
}


function ocultarTudo() {
    document.getElementById('estadoInicial').style.display = 'none';
    document.getElementById('painelVereador').classList.add('escondido');
    document.getElementById('painelEspecial').classList.add('escondido');
    document.getElementById('secaoSugestao').classList.add('escondido');
    document.getElementById('topbarBadges').innerHTML = '';
    const mes = document.getElementById('mesSelect').value;
    const gab = document.getElementById('gabineteSelect').value;
    if (!mes || !gab) document.getElementById('estadoInicial').style.display = 'flex';
}

function renderizarAlertas(idEl, alertas) {
    const el = document.getElementById(idEl);
    el.innerHTML = '';
    const icons = {
        erro:  `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`,
        aviso: `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
        ok:    `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`,
    };
    alertas.forEach(a => {
        el.innerHTML += `<div class="alerta alerta-${a.tipo}">${icons[a.tipo]}<span>${a.msg}</span></div>`;
    });
}

// "DD/MM/AAAA" → Date
function parseData(str) {
    if (!str || !str.trim()) return null;
    const partes = str.trim().split('/');
    if (partes.length !== 3) return null;
    const [d, m, a] = partes.map(Number);
    if (!d || !m || !a) return null;
    return new Date(a, m - 1, d);
}

// "MM/AAAA" → { inicio, fim }
function intervaloMes(mesAno) {
    const [m, a] = mesAno.split('/').map(Number);
    return { inicio: new Date(a, m - 1, 1), fim: new Date(a, m, 0) };
}

function estaAtivo(admissao, exoneracao, inicio, fim) {
    if (!admissao) return false;
    if (admissao > fim) return false;
    if (exoneracao && exoneracao < inicio) return false;
    return true;
}

function moeda(valor) {
    return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function parseMoeda(str) {
    // Aceita "12.000,00" e "12000.00"
    const s = str.toString().trim();
    if (s.includes(',')) return parseFloat(s.replace(/\./g, '').replace(',', '.')) || 0;
    return parseFloat(s) || 0;
}

function formatarData(date) {
    if (!date) return '—';
    return date.toLocaleDateString('pt-BR');
}

function setStatus(tipo, msg) {
    const dot = document.querySelector('.status-dot');
    const txt = document.getElementById('statusConexao');
    dot.className = 'status-dot' + (tipo ? ' ' + tipo : '');
    txt.textContent = msg;
}

iniciar();

// ─── EXPORTAÇÃO ─────────────────────────────────────────────────────────────

function dadosParaExport() {
    const e       = _exportEstado;
    const isVer   = e.tipo === 'vereador';
    const titulo  = isVer ? 'Relatório do Gabinete' : 'Relatório de Lotação';
    const secComp = isVer ? 'Composição do Gabinete' : 'Composição da Lotação';

    const servsLimpos = e.servidores.filter(s =>
        !s.exoneradoNoMes && !isCedido(s['Cargo'] || '')
    );

    const linhasServs = servsLimpos.map(s => ({
        'Matrícula': (s['Matrícula'] || s['Matricula'] || '').trim() || '—',
        'Nome':      (s['Nome do Servidor'] || '').trim(),
        'Cargo':     (s['Cargo'] || '').trim(),
        'Salário':   moeda(tabelaCargos[(s['Cargo'] || '').trim()] || 0),
        'Admissão':  s.admissao ? formatarData(s.admissao) : '—',
    }));

    const consumivel = servsLimpos.map(s => extrairCC(s['Cargo'] || ''));
    const linhasComp = e.estrutura.map(cc => {
        const idx = consumivel.indexOf(cc);
        const ocupado = idx !== -1;
        if (ocupado) consumivel.splice(idx, 1);
        return { 'Cargo': cc, 'Status': ocupado ? 'Ocupado' : 'Vago' };
    });

    return { titulo, secComp, linhasServs, linhasComp, e };
}

function exportarCSV() {
    const { linhasServs, linhasComp, e } = dadosParaExport();
    const nome = e.gab.replace(/[/\\?%*:|"<>]/g, '-');

    // Arquivo 1: Servidores
    const blob1 = new Blob([Papa.unparse(linhasServs)], { type: 'text/csv;charset=utf-8;' });
    const a1 = document.createElement('a');
    a1.href = URL.createObjectURL(blob1);
    a1.download = `${nome} - Servidores.csv`;
    a1.click();

    // Arquivo 2: Composição (pequeno delay para o browser não bloquear)
    setTimeout(() => {
        const blob2 = new Blob([Papa.unparse(linhasComp)], { type: 'text/csv;charset=utf-8;' });
        const a2 = document.createElement('a');
        a2.href = URL.createObjectURL(blob2);
        a2.download = `${nome} - Composição.csv`;
        a2.click();
    }, 400);
}

function exportarXLSX() {
    const { linhasServs, linhasComp, e } = dadosParaExport();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(linhasServs), 'Servidores');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(linhasComp),  'Composição');
    XLSX.writeFile(wb, `${e.gab} - ${e.mes}.xlsx`.replace(/[/\\?%*:|"<>]/g, '-'));
}

function exportarPDF() {
    const { jsPDF } = window.jspdf;
    const { titulo, secComp, linhasServs, linhasComp, e } = dadosParaExport();
    const isVer = e.tipo === 'vereador';
    const doc   = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const PW = 210, PH = 297, ML = 18, MR = 18;

    // Logo via elemento HTML
    const LOGO = document.getElementById('logoBase64')?.value || '';
    // Dimensões da logo no PDF: proporcional ao original 1920x1080 → cabe em ~80x45mm
    const LW = 72, LH = 40;

    // ── CABEÇALHO ──
    // Fundo branco puro no topo
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, PW, 48, 'F');

    // Logo no canto superior direito
    doc.addImage(LOGO, 'PNG', PW - MR - LW, 4, LW, LH);

    // Linha institucional à esquerda
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.setTextColor(40, 40, 40);
    doc.text('CONTROLE ORÇAMENTÁRIO LEGISLATIVO', ML, 14);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 100);
    doc.text(isVer ? 'Relatório do Gabinete' : 'Relatório de Lotação', ML, 21);

    // Linha divisória
    doc.setDrawColor(30, 30, 30);
    doc.setLineWidth(0.6);
    doc.line(ML, 46, PW - MR, 46);
    doc.setLineWidth(0.15);
    doc.setDrawColor(180, 180, 180);
    doc.line(ML, 47.5, PW - MR, 47.5);

    // ── BLOCO IDENTIFICAÇÃO ──
    let y = 56;

    // Nome da lotação — maior destaque
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(15, 15, 15);
    doc.text(e.gab, ML, y);
    y += 7;

    // Responsável (só para especiais)
    if (!isVer && e.responsavel) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.setTextColor(80, 80, 80);
        doc.text('Vereador Responsável: ', ML, y);
        const labelW = doc.getTextWidth('Vereador Responsável: ');
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(30, 30, 30);
        doc.text(e.responsavel, ML + labelW, y);
        y += 5;
    }

    // Linha separadora leve
    doc.setDrawColor(210, 210, 210);
    doc.setLineWidth(0.2);
    doc.line(ML, y + 1, PW - MR, y + 1);
    y += 7;

    // ── TABELA SERVIDORES ──
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(30, 30, 30);
    doc.text('SERVIDORES LOTADOS', ML, y);
    y += 3;

    doc.autoTable({
        startY: y,
        margin: { left: ML, right: MR },
        head: [['Matrícula', 'Nome do Servidor', 'Cargo', 'Salário', 'Admissão']],
        body: linhasServs.map(r => [r['Matrícula'], r['Nome'], r['Cargo'], r['Salário'], r['Admissão']]),
        styles: {
            font: 'helvetica', fontSize: 8, cellPadding: 3,
            textColor: [25, 25, 25], lineColor: [200, 200, 200], lineWidth: 0.15,
        },
        headStyles: {
            fillColor: [35, 35, 35], textColor: [255, 255, 255],
            fontStyle: 'bold', fontSize: 7.5,
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        columnStyles: {
            0: { cellWidth: 22 },
            1: { cellWidth: 'auto' },
            2: { cellWidth: 30 },
            3: { cellWidth: 26, halign: 'right' },
            4: { cellWidth: 22 },
        },
    });

    y = doc.lastAutoTable.finalY + 10;

    // ── COMPOSIÇÃO ──
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(30, 30, 30);
    doc.text(secComp.toUpperCase(), ML, y);
    y += 3;

    doc.autoTable({
        startY: y,
        margin: { left: ML, right: MR },
        tableWidth: 80,
        head: [['Cargo (CC)', 'Status']],
        body: linhasComp.map(r => [r['Cargo'], r['Status']]),
        styles: {
            font: 'helvetica', fontSize: 8, cellPadding: 3,
            textColor: [25, 25, 25], lineColor: [200, 200, 200], lineWidth: 0.15,
        },
        headStyles: {
            fillColor: [35, 35, 35], textColor: [255, 255, 255],
            fontStyle: 'bold', fontSize: 7.5,
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        columnStyles: { 0: { cellWidth: 40 }, 1: { cellWidth: 38 } },
        didParseCell(data) {
            if (data.column.index === 1 && data.section === 'body') {
                data.cell.styles.textColor = data.cell.raw === 'Ocupado' ? [30, 110, 60] : [200, 60, 50];
                data.cell.styles.fontStyle = 'bold';
            }
        },
    });

    // ── RODAPÉ ──
    const today  = new Date();
    const dataStr = today.toLocaleDateString('pt-BR');
    const total  = doc.internal.getNumberOfPages();
    const FLW = 44, FLH = 25; // logo menor no rodapé

    for (let i = 1; i <= total; i++) {
        doc.setPage(i);

        // Linha dupla antes do rodapé
        doc.setDrawColor(30, 30, 30);
        doc.setLineWidth(0.6);
        doc.line(ML, PH - 22, PW - MR, PH - 22);
        doc.setLineWidth(0.15);
        doc.setDrawColor(180, 180, 180);
        doc.line(ML, PH - 21, PW - MR, PH - 21);

        // Logo pequena à direita do rodapé
        doc.addImage(LOGO, 'PNG', PW - MR - FLW, PH - 24, FLW, FLH);

        // Texto do rodapé à esquerda
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7);
        doc.setTextColor(80, 80, 80);
        doc.text('Câmara Municipal de Curitiba', ML, PH - 14);
        doc.text('Controle Orçamentário Legislativo', ML, PH - 10);
        doc.setTextColor(130, 130, 130);
        doc.text(`Gerado em ${dataStr}   ·   Página ${i} de ${total}`, ML, PH - 6);
    }

    // Nome do arquivo: lotação + data AAAA-MM-DD
    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    doc.save(`${e.gab} - ${ymd}.pdf`.replace(/[/\?%*:|"<>]/g, '-'));
}

// ─── NAVEGAÇÃO ───────────────────────────────────────────────────────────────
function navegarPara(pagina) {
    const paginaPainel    = document.querySelector('.conteudo-principal:not(#paginaRelatorios)');
    const paginaRelat     = document.getElementById('paginaRelatorios');
    const navPainel       = document.getElementById('navPainel');
    const navRelatorios   = document.getElementById('navRelatorios');

    if (pagina === 'relatorios') {
        paginaPainel.classList.add('escondido');
        paginaRelat.classList.remove('escondido');
        navPainel.classList.remove('ativo');
        navRelatorios.classList.add('ativo');
        preencherListaRelat();
    } else {
        paginaPainel.classList.remove('escondido');
        paginaRelat.classList.add('escondido');
        navPainel.classList.add('ativo');
        navRelatorios.classList.remove('ativo');
    }
}

// ─── PÁGINA RELATÓRIOS ───────────────────────────────────────────────────────
let _relatSelecionadas = new Set();
let _relatMesAtual = '';

function preencherListaRelat() {
    // Usa o mês selecionado no painel, ou o mais recente disponível
    _relatMesAtual = document.getElementById('mesSelect').value ||
        (_opsMes && _opsMes.length ? _opsMes[_opsMes.length - 1] : '');

    const lista = document.getElementById('relatLista');
    lista.innerHTML = '';

    // Coleta todas as lotações da aba verbas
    const gabs = [...new Set(dadosVerbas.map(l => l['Gabinete']?.trim()).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'pt-BR'));

    gabs.forEach(gab => {
        const tipo = classificarTipo(gab);
        const tipoLabel = tipo === 'vereador' ? 'Gabinete' : tipo === 'mesa_diretora' ? 'Mesa Diretora' : 'Bloco/Liderança';
        const checked = _relatSelecionadas.has(gab) ? 'checked' : '';
        const selClass = _relatSelecionadas.has(gab) ? ' selecionado' : '';

        const div = document.createElement('div');
        div.className = `relat-item${selClass}`;
        div.dataset.gab = gab;
        div.innerHTML = `
            <input type="checkbox" id="relat_${gab}" ${checked} onchange="toggleRelat('${gab.replace(/'/g, "\\'")}', this)">
            <label class="relat-item-nome" for="relat_${gab}">${gab}</label>
            <span class="relat-item-tipo">${tipoLabel}</span>`;
        lista.appendChild(div);
    });

    atualizarContadorRelat();

    // Busca em tempo real
    const busca = document.getElementById('relatBusca');
    busca.addEventListener('input', filtrarListaRelat);
}

function filtrarListaRelat() {
    const q = document.getElementById('relatBusca').value.toLowerCase();
    document.querySelectorAll('.relat-item').forEach(item => {
        const nome = (item.dataset.gab || '').toLowerCase();
        item.classList.toggle('escondido', q.length > 0 && !nome.includes(q));
    });
}

function limparBuscaRelat() {
    document.getElementById('relatBusca').value = '';
    filtrarListaRelat();
    document.getElementById('relatBusca').focus();
}

function toggleRelat(gab, cb) {
    const item = cb.closest('.relat-item');
    if (cb.checked) {
        _relatSelecionadas.add(gab);
        item.classList.add('selecionado');
    } else {
        _relatSelecionadas.delete(gab);
        item.classList.remove('selecionado');
    }
    atualizarContadorRelat();
}

function atualizarContadorRelat() {
    const n = _relatSelecionadas.size;
    document.getElementById('relatContador').textContent =
        n === 0 ? 'Nenhuma selecionada' : n === 1 ? '1 selecionada' : `${n} selecionadas`;
}

function selecionarTodasRelat() {
    document.querySelectorAll('.relat-item:not(.escondido)').forEach(item => {
        _relatSelecionadas.add(item.dataset.gab);
        item.classList.add('selecionado');
        item.querySelector('input[type="checkbox"]').checked = true;
    });
    atualizarContadorRelat();
}

function limparSelecaoRelat() {
    _relatSelecionadas.clear();
    document.querySelectorAll('.relat-item').forEach(item => {
        item.classList.remove('selecionado');
        item.querySelector('input[type="checkbox"]').checked = false;
    });
    atualizarContadorRelat();
}

// Prepara dados de todas as lotações selecionadas para o mês atual
function dadosRelatTodasLotacoes() {
    const mes = _relatMesAtual;
    if (!mes || _relatSelecionadas.size === 0) {
        alert('Selecione pelo menos uma lotação e verifique se um mês está disponível.');
        return null;
    }

    const { inicio, fim } = intervaloMes(mes);

    return [..._relatSelecionadas].map(gab => {
        const tipo     = classificarTipo(gab);
        const verbaMes = dadosVerbas.find(l => l['Mês']?.trim() === mes && l['Gabinete']?.trim() === gab) || {};
        const responsavel = (verbaMes['Responsável'] || verbaMes['Responsavel'] || '').trim();
        const isVer    = tipo === 'vereador';

        const servsAtivos = dadosServidores
            .filter(l => (l['Gabinete'] || '').trim() === gab)
            .map(l => {
                const admissao   = parseData(l['Admissão']   || l['Admissao']   || '');
                const exoneracao = parseData(l['Exoneração'] || l['Exoneracao'] || '');
                return { ...l, admissao, exoneracao,
                    ativo: estaAtivo(admissao, exoneracao, inicio, fim),
                    exoneradoNoMes: !!(exoneracao && exoneracao >= inicio && exoneracao <= fim) };
            })
            .filter(l => l.ativo && !l.exoneradoNoMes && !isCedido(l['Cargo'] || ''));

        let estrutura = [];
        if (tipo === 'mesa_diretora') {
            const chave = Object.keys(LOTACOES_ESPECIAIS).find(k => k.toUpperCase() === gab.trim().toUpperCase());
            estrutura = chave ? LOTACOES_ESPECIAIS[chave] : [];
        } else if (tipo === 'bloco') {
            estrutura = ['CC-8'];
        } else {
            estrutura = dadosEstruturas[gab] || [];
        }

        const linhasServs = servsAtivos.map(s => ({
            'Matrícula': (s['Matrícula'] || s['Matricula'] || '').trim() || '—',
            'Nome':      (s['Nome do Servidor'] || '').trim(),
            'Cargo':     (s['Cargo'] || '').trim(),
            'Salário':   moeda(tabelaCargos[(s['Cargo'] || '').trim()] || 0),
            'Admissão':  s.admissao ? formatarData(s.admissao) : '—',
        }));

        const consumivel = servsAtivos.map(s => extrairCC(s['Cargo'] || ''));
        const linhasComp = estrutura.map(cc => {
            const idx = consumivel.indexOf(cc);
            const ocupado = idx !== -1;
            if (ocupado) consumivel.splice(idx, 1);
            return { 'Lotação': gab, 'Cargo': cc, 'Status': ocupado ? 'Ocupado' : 'Vago' };
        });

        return { gab, tipo, isVer, responsavel, linhasServs, linhasComp };
    });
}

// ── Exportar PDF (uma página por lotação) ──
function relatExportarPDF() {
    const lotes = dadosRelatTodasLotacoes();
    if (!lotes) return;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const PW = 210, PH = 297, ML = 18, MR = 18;
    const LW = 72, LH = 40, FLW = 44, FLH = 25;

    const LOGO = document.querySelector('#logoBase64')?.value || '';

    lotes.forEach((lotacao, idx) => {
        if (idx > 0) doc.addPage();

        const { gab, isVer, responsavel, linhasServs, linhasComp } = lotacao;
        const titulo = isVer ? 'Relatório do Gabinete' : 'Relatório de Lotação';
        const secComp = isVer ? 'Composição do Gabinete' : 'Composição da Lotação';

        // Cabeçalho
        doc.setFillColor(255, 255, 255);
        doc.rect(0, 0, PW, 48, 'F');
        if (LOGO) doc.addImage(LOGO, 'PNG', PW - MR - LW, 4, LW, LH);

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(13);
        doc.setTextColor(40, 40, 40);
        doc.text('CONTROLE ORÇAMENTÁRIO LEGISLATIVO', ML, 14);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.setTextColor(100, 100, 100);
        doc.text(titulo, ML, 21);

        doc.setDrawColor(30, 30, 30);
        doc.setLineWidth(0.6);
        doc.line(ML, 46, PW - MR, 46);
        doc.setLineWidth(0.15);
        doc.setDrawColor(180, 180, 180);
        doc.line(ML, 47.5, PW - MR, 47.5);

        let y = 56;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(16);
        doc.setTextColor(15, 15, 15);
        doc.text(gab, ML, y);
        y += 7;

        if (!isVer && responsavel) {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(80, 80, 80);
            const lw = doc.getTextWidth('Vereador Responsável: ');
            doc.text('Vereador Responsável: ', ML, y);
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(30, 30, 30);
            doc.text(responsavel, ML + lw, y);
            y += 5;
        }

        doc.setDrawColor(210, 210, 210);
        doc.setLineWidth(0.2);
        doc.line(ML, y + 1, PW - MR, y + 1);
        y += 7;

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.setTextColor(30, 30, 30);
        doc.text('SERVIDORES LOTADOS', ML, y);
        y += 3;

        doc.autoTable({
            startY: y, margin: { left: ML, right: MR },
            head: [['Matrícula', 'Nome do Servidor', 'Cargo', 'Salário', 'Admissão']],
            body: linhasServs.map(r => [r['Matrícula'], r['Nome'], r['Cargo'], r['Salário'], r['Admissão']]),
            styles: { font: 'helvetica', fontSize: 8, cellPadding: 3, textColor: [25,25,25], lineColor: [200,200,200], lineWidth: 0.15 },
            headStyles: { fillColor: [35,35,35], textColor: [255,255,255], fontStyle: 'bold', fontSize: 7.5 },
            alternateRowStyles: { fillColor: [245,245,245] },
            columnStyles: { 0:{cellWidth:22}, 1:{cellWidth:'auto'}, 2:{cellWidth:30}, 3:{cellWidth:26,halign:'right'}, 4:{cellWidth:22} },
        });

        y = doc.lastAutoTable.finalY + 10;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.setTextColor(30, 30, 30);
        doc.text(secComp.toUpperCase(), ML, y);
        y += 3;

        doc.autoTable({
            startY: y, margin: { left: ML, right: MR }, tableWidth: 80,
            head: [['Cargo (CC)', 'Status']],
            body: linhasComp.map(r => [r['Cargo'], r['Status']]),
            styles: { font: 'helvetica', fontSize: 8, cellPadding: 3, textColor: [25,25,25], lineColor: [200,200,200], lineWidth: 0.15 },
            headStyles: { fillColor: [35,35,35], textColor: [255,255,255], fontStyle: 'bold', fontSize: 7.5 },
            alternateRowStyles: { fillColor: [245,245,245] },
            columnStyles: { 0:{cellWidth:40}, 1:{cellWidth:38} },
            didParseCell(data) {
                if (data.column.index === 1 && data.section === 'body') {
                    data.cell.styles.textColor = data.cell.raw === 'Ocupado' ? [30,110,60] : [200,60,50];
                    data.cell.styles.fontStyle = 'bold';
                }
            },
        });
    });

    // Rodapé em todas as páginas
    const today  = new Date();
    const dataStr = today.toLocaleDateString('pt-BR');
    const total  = doc.internal.getNumberOfPages();
    for (let i = 1; i <= total; i++) {
        doc.setPage(i);
        doc.setDrawColor(30,30,30); doc.setLineWidth(0.6);
        doc.line(ML, PH-22, PW-MR, PH-22);
        doc.setLineWidth(0.15); doc.setDrawColor(180,180,180);
        doc.line(ML, PH-21, PW-MR, PH-21);
        if (LOGO) doc.addImage(LOGO, 'PNG', PW-MR-FLW, PH-24, FLW, FLH);
        doc.setFont('helvetica','normal'); doc.setFontSize(7); doc.setTextColor(80,80,80);
        doc.text('Câmara Municipal de Curitiba', ML, PH-14);
        doc.text('Controle Orçamentário Legislativo', ML, PH-10);
        doc.setTextColor(130,130,130);
        doc.text(`Gerado em ${dataStr}   ·   Página ${i} de ${total}`, ML, PH-6);
    }

    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    doc.save(`Relatorio-Consolidado-${ymd}.pdf`);
}

// ── Exportar XLSX ──
function relatExportarXLSX() {
    const lotes = dadosRelatTodasLotacoes();
    if (!lotes) return;

    const wb = XLSX.utils.book_new();
    const todasServs = lotes.flatMap(l => l.linhasServs.map(r => ({ 'Lotação': l.gab, ...r })));
    const todasComp  = lotes.flatMap(l => l.linhasComp);

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(todasServs), 'Servidores');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(todasComp),  'Composição');

    const ymd = new Date();
    const nome = ymd.getFullYear() + '-' + String(ymd.getMonth()+1).padStart(2,'0') + '-' + String(ymd.getDate()).padStart(2,'0');
    XLSX.writeFile(wb, `Relatorio-Consolidado-${nome}.xlsx`);
}

// ── Exportar CSV ──
function relatExportarCSV() {
    const lotes = dadosRelatTodasLotacoes();
    if (!lotes) return;

    const linhas = lotes.flatMap(l =>
        l.linhasServs.map(r => ({
            'Lotação':   l.gab,
            'Matrícula': r['Matrícula'],
            'Nome':      r['Nome'],
            'Cargo':     r['Cargo'],
            'Salário':   r['Salário'],
            'Admissão':  r['Admissão'],
        }))
    );

    const blob = new Blob([Papa.unparse(linhas)], { type: 'text/csv;charset=utf-8;' });
    const today = new Date();
    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Relatorio-Consolidado-${ymd}.csv`;
    a.click();
}
