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
    const blob = new Blob(
        [`Lotação: ${e.gab}\n\nServidores\n${Papa.unparse(linhasServs)}\n\nComposição\n${Papa.unparse(linhasComp)}`],
        { type: 'text/csv;charset=utf-8;' }
    );
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${nome}.csv`;
    a.click();
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
    const paginaPainel    = document.querySelector('.conteudo-principal:not(#paginaRelatorios):not(#paginaCalculadora)');
    const paginaRelat     = document.getElementById('paginaRelatorios');
    const paginaCalc      = document.getElementById('paginaCalculadora');
    const navPainel       = document.getElementById('navPainel');
    const navRelatorios   = document.getElementById('navRelatorios');
    const navCalculadora  = document.getElementById('navCalculadora');

    // Esconde tudo
    paginaPainel.classList.add('escondido');
    paginaRelat.classList.add('escondido');
    paginaCalc.classList.add('escondido');
    [navPainel, navRelatorios, navCalculadora].forEach(n => n && n.classList.remove('ativo'));

    if (pagina === 'relatorios') {
        paginaRelat.classList.remove('escondido');
        navRelatorios.classList.add('ativo');
        preencherListaRelat();
        atualizarBotoesSecao();
    } else if (pagina === 'calculadora') {
        paginaCalc.classList.remove('escondido');
        navCalculadora.classList.add('ativo');
        iniciarCalculadora();
    } else {
        paginaPainel.classList.remove('escondido');
        navPainel.classList.add('ativo');
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
    if (_relatSecaoServs) {
        const todasServs = lotes.flatMap(l => l.linhasServs.map(r => ({ 'Lotação': l.gab, ...r })));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(todasServs), 'Servidores');
    }
    if (_relatSecaoComp) {
        const todasComp = lotes.flatMap(l => l.linhasComp);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(todasComp), 'Composição');
    }

    const today = new Date();
    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    XLSX.writeFile(wb, `Relatorio-Consolidado-${ymd}.xlsx`);
}

// ── Exportar CSV ──
function relatExportarCSV() {
    // Só disponível quando apenas uma seção está ativa
    if (_relatSecaoServs && _relatSecaoComp) return;

    const lotes = dadosRelatTodasLotacoes();
    if (!lotes) return;

    let linhas;
    let sufixo;

    if (_relatSecaoServs) {
        // Apenas relação de servidores
        linhas = lotes.flatMap(l =>
            l.linhasServs.map(r => ({
                'Lotação':   l.gab,
                'Matrícula': r['Matrícula'],
                'Nome':      r['Nome'],
                'Cargo':     r['Cargo'],
                'Salário':   r['Salário'],
                'Admissão':  r['Admissão'],
            }))
        );
        sufixo = 'Servidores';
    } else {
        // Apenas composição
        linhas = lotes.flatMap(l => l.linhasComp);
        sufixo = 'Composição';
    }

    const blob = new Blob([Papa.unparse(linhas)], { type: 'text/csv;charset=utf-8;' });
    const today = new Date();
    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Relatorio-Consolidado-${sufixo}-${ymd}.csv`;
    a.click();
}

// ─── TOGGLE SEÇÕES RELATÓRIO ─────────────────────────────────────────────────
let _relatSecaoServs = true;
let _relatSecaoComp  = true;

function toggleSecao(sec) {
    if (sec === 'servs') {
        // Não permite desmarcar se a outra já está desmarcada
        if (!_relatSecaoServs && !_relatSecaoComp) return;
        if (_relatSecaoServs && !_relatSecaoComp) return; // única ativa
        _relatSecaoServs = !_relatSecaoServs;
    } else {
        if (_relatSecaoComp && !_relatSecaoServs) return; // única ativa
        _relatSecaoComp = !_relatSecaoComp;
    }
    atualizarBotoesSecao();
}

function atualizarBotoesSecao() {
    const btnS = document.getElementById('btnSecaoServs');
    const btnC = document.getElementById('btnSecaoComp');
    if (btnS) btnS.classList.toggle('ativo', _relatSecaoServs);
    if (btnC) btnC.classList.toggle('ativo', _relatSecaoComp);

    // CSV desativado quando ambas ativas
    const ambas = _relatSecaoServs && _relatSecaoComp;
    document.querySelectorAll('#paginaRelatorios .btn-csv').forEach(b => {
        b.classList.toggle('desativado', ambas);
    });
}

// ─── CALCULADORA DE COMPOSIÇÃO ───────────────────────────────────────────────

// Estado da calculadora: { 'CC-1': 1, 'CC-2': 0, ... }
let _calcQtd = {};

// Lista de CCs na ordem do formulário
const CALC_CCS_ORDEM = ['CC-1','CC-2','CC-3','CC-4','CC-5','CC-6','CC-7','CC-8','CC-9'];

function calcLimparTudo() {
    CALC_CCS_ORDEM.forEach(cc => { _calcQtd[cc] = cc === 'CC-1' ? 1 : 0; });
    // Reset qtd displays
    CALC_CCS_ORDEM.forEach(cc => {
        if (cc === 'CC-1') return;
        const el = document.getElementById(`calcQtd_${cc}`);
        if (el) el.textContent = '0';
    });
    atualizarCalcIndicadores();
}

function iniciarCalculadora() {
    // Só monta a grade uma vez
    if (document.getElementById('calcGrade').children.length > 0) {
        atualizarCalcIndicadores();
        return;
    }

    // Inicializa quantidades
    CALC_CCS_ORDEM.forEach(cc => { _calcQtd[cc] = cc === 'CC-1' ? 1 : 0; });

    const grade = document.getElementById('calcGrade');
    grade.innerHTML = '';

    CALC_CCS_ORDEM.forEach(cc => {
        const sal    = tabelaCargos[cc] || obterSalarioCC(cc);
        const isCC1  = cc === 'CC-1';
        const cargo  = isCC1 ? 'Chefe de Gabinete Parlamentar' : 'Assessor de Gabinete Parlamentar';

        const div = document.createElement('div');
        div.className = `calc-cc-card${isCC1 ? ' ativo bloqueado' : ''}`;
        div.id = `calcCard_${cc}`;
        div.innerHTML = `
            <div class="calc-cc-nome">${cc}</div>
            <div class="calc-cc-cargo">${cargo}</div>
            <div class="calc-cc-salario">${moeda(sal)}</div>
            <div class="calc-cc-controle">
                ${isCC1
                    ? `<span class="calc-cc-fixo">Obrigatório</span>`
                    : `<button class="calc-cc-btn" id="calcMenos_${cc}" onclick="calcAjustar('${cc}',-1)" disabled>−</button>
                       <span class="calc-cc-qtd" id="calcQtd_${cc}">0</span>
                       <button class="calc-cc-btn" id="calcMais_${cc}" onclick="calcAjustar('${cc}',+1)">+</button>`
                }
            </div>`;
        grade.appendChild(div);
    });

    atualizarCalcIndicadores();
}

// Obtém salário do CC pela tabela ou por valor fixo do formulário
function obterSalarioCC(cc) {
    // Busca na tabelaCargos por qualquer cargo que contenha o CC
    for (const [nome, sal] of Object.entries(tabelaCargos)) {
        if (extrairCC(nome) === cc) return sal;
    }
    // Fallback: valores do formulário
    const fallback = {
        'CC-1': 18599.66, 'CC-2': 16533.01, 'CC-3': 14466.48,
        'CC-4': 12399.85, 'CC-5': 10333.14, 'CC-6': 8266.53,
        'CC-7': 6199.89,  'CC-8': 4133.19,  'CC-9': 3099.98
    };
    return fallback[cc] || 0;
}

function calcAjustar(cc, delta) {
    const totalVagas = CALC_CCS_ORDEM.reduce((s, c) => s + _calcQtd[c], 0);
    const sal = tabelaCargos[cc] || obterSalarioCC(cc);
    const totalVerba = calcVerbaAtual();

    if (delta > 0) {
        if (totalVagas >= 9) return;                          // limite de vagas
        if (totalVerba + sal > TETO_VEREADOR + TOLERANCIA) return; // limite financeiro
        _calcQtd[cc]++;
    } else {
        if (_calcQtd[cc] <= 0) return;
        _calcQtd[cc]--;
    }

    atualizarCalcIndicadores();
}

function calcVerbaAtual() {
    return CALC_CCS_ORDEM.reduce((s, cc) => {
        const sal = tabelaCargos[cc] || obterSalarioCC(cc);
        return s + sal * _calcQtd[cc];
    }, 0);
}

function atualizarCalcIndicadores() {
    const totalVagas = CALC_CCS_ORDEM.reduce((s, cc) => s + _calcQtd[cc], 0);
    const totalVerba = calcVerbaAtual();
    const saldo      = TETO_VEREADOR - totalVerba;
    const vagasLiv   = 9 - totalVagas;
    const pctVagas   = (totalVagas / 9) * 100;
    const pctVerba   = Math.min((totalVerba / TETO_VEREADOR) * 100, 100);

    // Vagas
    document.getElementById('calcVagasUsadas').textContent = `${totalVagas} / 9`;
    document.getElementById('calcVagasSub').textContent    =
        vagasLiv === 1 ? '1 vaga disponível' : vagasLiv > 1 ? `${vagasLiv} vagas disponíveis` : 'sem vagas disponíveis';
    const barVagas = document.getElementById('calcBarVagas');
    barVagas.style.width = pctVagas.toFixed(1) + '%';
    barVagas.className = 'calc-ind-bar' + (totalVagas >= 9 ? ' perigo' : totalVagas >= 7 ? ' aviso' : '');

    // Verba
    document.getElementById('calcVerbaUsada').textContent = moeda(totalVerba);
    const barVerba = document.getElementById('calcBarVerba');
    barVerba.style.width = pctVerba.toFixed(1) + '%';
    barVerba.className = 'calc-ind-bar' + (pctVerba >= 100 ? ' perigo' : pctVerba >= 85 ? ' aviso' : '');

    // Saldo
    const cardSaldo = document.getElementById('calcCardSaldo');
    cardSaldo.classList.toggle('calc-ind-saldo-neg', saldo < 0);
    document.getElementById('calcSaldo').textContent = moeda(Math.abs(saldo));
    document.getElementById('calcSaldoSub').textContent = saldo < 0 ? 'acima do teto legal' : 'dentro do teto legal';

    // Atualiza visual dos cards e botões
    CALC_CCS_ORDEM.forEach(cc => {
        if (cc === 'CC-1') return;
        const qtd   = _calcQtd[cc];
        const card  = document.getElementById(`calcCard_${cc}`);
        const qtdEl = document.getElementById(`calcQtd_${cc}`);
        const menos = document.getElementById(`calcMenos_${cc}`);
        const mais  = document.getElementById(`calcMais_${cc}`);

        if (!card) return;
        qtdEl.textContent = qtd;
        card.classList.toggle('ativo', qtd > 0);

        // Menos: desabilita se qtd = 0
        menos.disabled = qtd <= 0;

        // Mais: desabilita e marca esgotado se vagas cheias ou verba esgotada
        const sal = tabelaCargos[cc] || obterSalarioCC(cc);
        const vagasCheias = totalVagas >= 9;
        const semSaldo    = totalVerba + sal > TETO_VEREADOR + TOLERANCIA;
        const naoPodeAdicionar = vagasCheias || semSaldo;
        mais.disabled = naoPodeAdicionar;
        // Esgotado: não tem qtd e não pode adicionar
        card.classList.toggle('esgotado', naoPodeAdicionar && qtd === 0);
    });
}

// ── Gerar formulário PDF ──
function gerarFormularioPDF() {
    const { jsPDF } = window.jspdf;
    const nomeVer = document.getElementById('calcVereador').value.trim();
    if (!nomeVer) {
        alert('Por favor, informe o nome do(a) vereador(a) antes de gerar o formulário.');
        document.getElementById('calcVereador').focus();
        return;
    }

    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const PW = 210, PH = 297, ML = 18, MR = 18, MT = 15;
    const CW = PW - ML - MR;

    // Logo e cabeçalho
    const LOGO = document.getElementById('logoBase64')?.value || '';
    if (LOGO) doc.addImage(LOGO, 'PNG', ML, MT, 36, 20);

    // Título à direita da logo, preto
    const titleX = ML + 42; // começa após a logo (36mm + gap)
    const titleW = CW - 42;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.setTextColor(0, 0, 0);
    doc.text('FORMULÁRIO', titleX + titleW / 2, MT + 6, { align: 'center' });
    doc.setFontSize(10.5);
    const tituloSplit = doc.splitTextToSize('DE ESCOLHA DE COMPOSIÇÃO DE GABINETE PARLAMENTAR', titleW);
    doc.text(tituloSplit, titleX + titleW / 2, MT + 12, { align: 'center' });

    // Borda ao redor do cabeçalho
    doc.setDrawColor(80, 80, 80);
    doc.setLineWidth(0.5);
    doc.rect(ML, MT, CW, 20);

    let y = MT + 28;

    // Texto introdutório
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(0);
    doc.text('EXCELENTÍSSIMO SENHOR PRESIDENTE DA COMISSÃO EXECUTIVA,', ML, y, { align: 'left' });
    y += 9;

    // Linha do vereador — nome sem caixa, com vírgula
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.5);
    doc.text('O (A) Vereador(a) ', ML, y);
    const lwLabel = doc.getTextWidth('O (A) Vereador(a) ');
    doc.setFont('helvetica', 'bold');
    doc.text(nomeVer + ',', ML + lwLabel, y);
    const lwNome = doc.getTextWidth(nomeVer + ',');

    y += 6;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.5);
    const linha2 = 'na escolha da organização de seu Gabinete Parlamentar, nos termos do Art. 4º da Lei Municipal 16.546/2025, opta pela seguinte composição:';
    const splitLinha2 = doc.splitTextToSize(linha2, CW);
    doc.text(splitLinha2, ML, y);
    y += splitLinha2.length * 5 + 4;

    // Tabela de composição
    // Monta lista de cargos escolhidos
    const cargosEscolhidos = [];
    CALC_CCS_ORDEM.forEach(cc => {
        for (let i = 0; i < _calcQtd[cc]; i++) {
            cargosEscolhidos.push(cc);
        }
    });
    // Preenche até 9 linhas
    while (cargosEscolhidos.length < 9) cargosEscolhidos.push('');

    const tabelaX   = ML + 30;
    const tabelaW   = 100;
    const colNomeW  = 78;
    const colCCW    = 22;
    const rowH      = 8;

    doc.setDrawColor(0);
    doc.setLineWidth(0.3);

    // Linha CC-1 (obrigatória, preenchida)
    doc.setFillColor(240, 240, 240);
    doc.rect(tabelaX, y, tabelaW, rowH, 'FD');
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.setTextColor(0);
    doc.text('1 Chefe de Gabinete Parlamentar', tabelaX + 2, y + 5.5);
    doc.setFont('helvetica', 'bold');
    doc.text('CC1', tabelaX + colNomeW + 2, y + 5.5);
    doc.line(tabelaX + colNomeW, y, tabelaX + colNomeW, y + rowH);
    y += rowH;

    // Linhas de assessores (slots 2..9)
    for (let i = 1; i < 9; i++) {
        const cc = cargosEscolhidos[i] || '';
        doc.setFillColor(255, 255, 255);
        doc.rect(tabelaX, y, tabelaW, rowH, 'FD');
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        doc.setTextColor(0);
        doc.text('1 Assessor de Gabinete Parlamentar', tabelaX + 2, y + 5.5);
        // Caixa do CC
        doc.setDrawColor(0);
        doc.rect(tabelaX + colNomeW, y, colCCW, rowH);
        if (cc) {
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(8.5);
            doc.text(cc, tabelaX + colNomeW + colCCW / 2, y + 5.5, { align: 'center' });
        }
        doc.line(tabelaX + colNomeW, y, tabelaX + colNomeW, y + rowH);
        y += rowH;
    }

    y += 8;

    // Obs.
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(0);
    const obs = 'Obs.: Nos termos do Art. 4º, §1º do respectivo diploma legal, o limite máximo de despesa com pessoal do gabinete com os cargos previstos deverá respeitar o somatório dos vencimentos dos cargos do CC-1 ao CC-6 mais dois cargos CC-7 (total de R$ 92.998,45).';
    const splitObs = doc.splitTextToSize(obs, CW);
    doc.text(splitObs, ML, y);
    y += splitObs.length * 4.2 + 10;

    // Data e assinatura
    const today = new Date();
    const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
    const diaStr = String(today.getDate()).padStart(2, '0');
    const mesStr = meses[today.getMonth()];
    const anoStr = today.getFullYear();

    // Linha da data com caixas
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text('Curitiba,', ML, y);
    let xCur = ML + doc.getTextWidth('Curitiba,') + 2;

    // Caixa dia
    doc.rect(xCur, y - 5, 12, 7);
    doc.setFont('helvetica', 'bold');
    doc.text(diaStr, xCur + 6, y - 0.5, { align: 'center' });
    xCur += 14;

    doc.setFont('helvetica', 'normal');
    doc.text('de', xCur, y);
    xCur += doc.getTextWidth('de') + 2;

    // Caixa mês
    doc.rect(xCur, y - 5, 28, 7);
    doc.setFont('helvetica', 'bold');
    doc.text(mesStr, xCur + 14, y - 0.5, { align: 'center' });
    xCur += 30;

    doc.setFont('helvetica', 'normal');
    doc.text('de', xCur, y);
    xCur += doc.getTextWidth('de') + 2;

    // Caixa ano
    doc.rect(xCur, y - 5, 18, 7);
    doc.setFont('helvetica', 'bold');
    doc.text(String(anoStr), xCur + 9, y - 0.5, { align: 'center' });

    // Linha de assinatura (em branco)
    const assinX = PW - MR - 60;
    doc.setDrawColor(0);
    doc.setLineWidth(0.3);
    doc.line(assinX, y, PW - MR, y);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(0);
    doc.text('Assinatura do(a) Vereador(a)', assinX + 30, y + 4, { align: 'center' });

    y += 16;

    // Tabela de vencimentos
    doc.setFillColor(220, 220, 220);
    doc.rect(ML, y, CW, 6, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(0);
    doc.text('VENCIMENTOS VIGENTES CONFORME SÍMBOLO', PW / 2, y + 4.2, { align: 'center' });
    y += 6;

    const venc = [
        ['CHEFE DE GABINETE PARLAMENTAR CC-1',     'R$ 18.599,66', 'ASSESSOR DE GABINETE PARLAMENTAR CC-6', 'R$ 8.266,53'],
        ['ASSESSOR DE GABINETE PARLAMENTAR CC-2',  'R$ 16.533,01', 'ASSESSOR DE GABINETE PARLAMENTAR CC-7', 'R$ 6.199,89'],
        ['ASSESSOR DE GABINETE PARLAMENTAR CC-3',  'R$ 14.466,48', 'ASSESSOR DE GABINETE PARLAMENTAR CC-8', 'R$ 4.133,19'],
        ['ASSESSOR DE GABINETE PARLAMENTAR CC-4',  'R$ 12.399,85', 'ASSESSOR DE GABINETE PARLAMENTAR CC-9', 'R$ 3.099,98'],
        ['ASSESSOR DE GABINETE PARLAMENTAR CC-5',  'R$ 10.333,14', '', ''],
    ];

    const hVenc = 5.5;
    // Total width = CW = 174mm. Each half = 87mm: cargo=65, valor=22
    const colW  = [65, 22, 65, 22];
    const colX  = [ML, ML+65, ML+87, ML+152];

    // Cabeçalho da tabela
    ['CARGO','VALOR','CARGO','VALOR'].forEach((h, ci) => {
        doc.setFillColor(200, 200, 200);
        doc.rect(colX[ci], y, colW[ci], hVenc, 'FD');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(7);
        doc.text(h, colX[ci] + colW[ci] / 2, y + 3.8, { align: 'center' });
    });
    y += hVenc;

    venc.forEach(row => {
        row.forEach((cell, ci) => {
            doc.setFillColor(255, 255, 255);
            doc.rect(colX[ci], y, colW[ci], hVenc, 'FD');
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(6.5);
            doc.setTextColor(0);
            const align = ci % 2 === 1 ? 'right' : 'left';
            const xText = align === 'right' ? colX[ci] + colW[ci] - 1 : colX[ci] + 1;
            doc.text(cell, xText, y + 3.8, { align });
        });
        y += hVenc;
    });

    // Borda geral ao redor de todo o formulário
    doc.setDrawColor(100, 100, 100);
    doc.setLineWidth(0.8);
    doc.rect(ML - 2, MT - 2, CW + 4, y - MT + 4);

    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    const nomeArq = nomeVer.replace(/[/\\?%*:|"<>]/g, '-');
    doc.save(`Formulario-Composicao-${nomeArq}-${ymd}.pdf`);
}
