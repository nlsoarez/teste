// ============================================
// CONFIGURAÇÃO DO SHAREPOINT DA CLARO
// ============================================

// URL de download direto do arquivo Excel no SharePoint
const SHAREPOINT_EXCEL_URL = 'https://corpclarobr-my.sharepoint.com/personal/nelson_soares_claro_com_br/_layouts/15/download.aspx?UniqueId=0EDD0D96-4E15-4704-826F-8E505DA1AAFD';

// Alternativa se a URL acima não funcionar (formato alternativo)
// const SHAREPOINT_EXCEL_URL = 'https://corpclarobr-my.sharepoint.com/:x:/r/personal/nelson_soares_claro_com_br/Documents/Analista%20Certificado.xlsx?download=1';

// ============================================
// CONFIGURAÇÃO DO SISTEMA
// ============================================

// Cache de elementos DOM
const matriculaInput = document.getElementById("matricula");
const resultadoDiv = document.getElementById("resultado");
const consultarBtn = document.getElementById("consultar-btn");
const statusDiv = document.getElementById("status");
const dataAtualizacaoSpan = document.getElementById("data-atualizacao");
const refreshBtn = document.getElementById("refresh-btn");

// Variáveis globais
let employees = [];
let employeeLookup = {};
let lastUpdate = null;
let isInitialLoad = true;

// ============================================
// CONFIGURAÇÃO DE METAS (atualizadas)
// ============================================

const METAS = {
    "ETIT": {
        "MÓVEL": 80,
        "RESIDENCIAL": 90,
        "EMPRESARIAL": 90
    },
    "Assertividade": {
        "MÓVEL": 85,
        "RESIDENCIAL": 70,
        "EMPRESARIAL": null
    },
    "DPA": {
        "CERTIFICACAO": 85,
        "INDIVIDUAL": 90
    }
};

// ============================================
// FUNÇÕES PRINCIPAIS PARA SHAREPOINT
// ============================================

/**
 * Carrega dados do arquivo Excel no SharePoint
 */
async function carregarDadosSharePoint() {
    try {
        mostrarStatus("🔄 Conectando ao SharePoint da Claro...", "loading");
        
        // Adiciona timestamp para evitar cache
        const urlComTimestamp = `${SHAREPOINT_EXCEL_URL}&t=${Date.now()}`;
        
        console.log("Tentando acessar:", urlComTimestamp);
        
        // Configurações para requisição ao SharePoint
        const requestOptions = {
            method: 'GET',
            mode: 'cors',
            credentials: 'omit', // Não enviar credenciais (SharePoint público)
            headers: {
                'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Cache-Control': 'no-cache'
            }
        };
        
        // Faz a requisição
        const response = await fetch(urlComTimestamp, requestOptions);
        
        if (!response.ok) {
            throw new Error(`Erro HTTP: ${response.status} ${response.statusText}`);
        }
        
        // Verifica se é um arquivo Excel
        const contentType = response.headers.get('content-type');
        if (!contentType || !contentType.includes('spreadsheetml')) {
            throw new Error('O arquivo não parece ser um Excel válido');
        }
        
        // Converte para ArrayBuffer
        const arrayBuffer = await response.arrayBuffer();
        
        // Processa o Excel
        processarExcel(arrayBuffer);
        
        // Atualiza data/hora
        atualizarDataHora();
        
        // Mostra sucesso
        const msg = `✅ Dados atualizados: ${employees.length} funcionários`;
        mostrarStatus(msg, "success", 3000);
        
        console.log("Dados carregados com sucesso:", employees);
        
    } catch (error) {
        console.error("Erro ao carregar dados do SharePoint:", error);
        tratarErroCarregamento(error);
    }
}

/**
 * Processa o arquivo Excel baixado
 */
function processarExcel(arrayBuffer) {
    try {
        // Lê o arquivo Excel usando SheetJS
        const workbook = XLSX.read(arrayBuffer, { 
            type: 'array',
            cellDates: true,
            cellStyles: true
        });
        
        // Pega a primeira planilha (ou específica)
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Converte para JSON
        const rawData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: "",
            blankrows: false
        });
        
        // Encontra os cabeçalhos
        let headerRowIndex = 0;
        for (let i = 0; i < Math.min(5, rawData.length); i++) {
            if (rawData[i].some(cell => 
                cell && typeof cell === 'string' && 
                (cell.toLowerCase().includes('matricula') || cell.toLowerCase().includes('nome'))
            )) {
                headerRowIndex = i;
                break;
            }
        }
        
        const headers = rawData[headerRowIndex].map(h => 
            String(h || "").trim().toLowerCase()
        );
        
        // Mapeia índices das colunas
        const colIndices = {
            matricula: headers.findIndex(h => h.includes('matricula') || h.includes('matrícula')),
            nome: headers.findIndex(h => h.includes('nome')),
            setor: headers.findIndex(h => h.includes('setor')),
            etit: headers.findIndex(h => h.includes('etit')),
            dpa: headers.findIndex(h => h.includes('dpa')),
            assertividade: headers.findIndex(h => h.includes('assertividade') || h.includes('acerto'))
        };
        
        // Processa os dados
        employees = [];
        
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
            const row = rawData[i];
            
            // Verifica se a linha tem matrícula
            const matricula = row[colIndices.matricula];
            if (!matricula && matricula !== 0) continue;
            
            // Cria objeto do funcionário
            const emp = {
                Matricula: String(matricula).trim().toUpperCase(),
                Nome: String(row[colIndices.nome] || "").trim(),
                Setor: formatarSetor(String(row[colIndices.setor] || "")),
                ETIT: formatarPorcentagem(row[colIndices.etit]),
                DPA: formatarPorcentagem(row[colIndices.dpa]),
                Assertividade: formatarPorcentagem(row[colIndices.assertividade])
            };
            
            // Valida dados básicos
            if (emp.Matricula && emp.Nome) {
                employees.push(emp);
            }
        }
        
        // Atualiza lookup
        atualizarLookup();
        
    } catch (error) {
        console.error("Erro ao processar Excel:", error);
        throw new Error("Falha ao processar planilha Excel");
    }
}

/**
 * Formata valores de porcentagem
 */
function formatarPorcentagem(valor) {
    if (valor === null || valor === undefined || valor === "") return "-";
    
    const strValor = String(valor).trim();
    
    // Se já for porcentagem formatada
    if (strValor.includes('%')) return strValor;
    
    // Se for número, converte para porcentagem
    const num = parseFloat(strValor.replace(',', '.'));
    if (!isNaN(num)) {
        if (num <= 1) {
            // Assume que é decimal (0.85 → 85%)
            return Math.round(num * 100) + "%";
        } else if (num <= 100) {
            // Já está em porcentagem (85 → 85%)
            return Math.round(num) + "%";
        } else {
            // Valor acima de 100 (pode ser DPA acima da meta)
            return num + "%";
        }
    }
    
    return strValor || "-";
}

/**
 * Formata o setor
 */
function formatarSetor(setor) {
    const setorUpper = setor.toUpperCase().trim();
    
    if (setorUpper.includes('EMPRESARIAL')) return 'EMPRESARIAL';
    if (setorUpper.includes('RESIDENCIAL')) return 'RESIDENCIAL';
    if (setorUpper.includes('MÓVEL') || setorUpper.includes('MOVEL')) return 'MÓVEL';
    
    return setorUpper || 'NÃO INFORMADO';
}

/**
 * Trata erros de carregamento
 */
function tratarErroCarregamento(error) {
    // Tenta usar cache local primeiro
    const cache = localStorage.getItem('claro_indicadores_cache');
    const cacheTime = localStorage.getItem('claro_indicadores_cache_time');
    
    if (cache && cacheTime) {
        const horaCache = parseInt(cacheTime);
        const agora = Date.now();
        const horasPassadas = (agora - horaCache) / (1000 * 60 * 60);
        
        if (horasPassadas < 24) { // Cache com menos de 24 horas
            employees = JSON.parse(cache);
            atualizarLookup();
            
            const horaFormatada = new Date(horaCache).toLocaleTimeString('pt-BR');
            mostrarStatus(`⚠ Usando dados em cache (${horaFormatada})`, "warning", 5000);
            return;
        }
    }
    
    // Se não tem cache ou está muito antigo, usa dados de fallback
    usarDadosFallback();
    
    if (isInitialLoad) {
        mostrarStatus("❌ SharePoint offline - usando dados locais", "error", 5000);
    } else {
        mostrarStatus("❌ Falha na atualização - mantendo dados atuais", "error", 3000);
    }
}

/**
 * Dados de fallback (atualize com dados recentes)
 */
function usarDadosFallback() {
    employees = [
        // Cole aqui seus dados atuais
        { "Matricula": "N6088107", "Nome": "LEANDRO GONÇALVES DE CARVALHO", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "64%", "Assertividade": "-" },
        { "Matricula": "N5619600", "Nome": "BRUNO COSTA BUCARD", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "60%", "Assertividade": "-" },
        { "Matricula": "N0238475", "Nome": "MARLEY MARQUES RIBEIRO", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "-", "Assertividade": "-" },
        { "Matricula": "N0189105", "Nome": "IGOR MARCELINO DE MARINS", "Setor": "EMPRESARIAL", "ETIT": "100%", "DPA": "77%", "Assertividade": "-" },
        { "Matricula": "N5737414", "Nome": "SANDRO DA SILVA CARVALHO", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "101%", "Assertividade": "-" },
        { "Matricula": "N5713690", "Nome": "GABRIELA TAVARES DA SILVA", "Setor": "EMPRESARIAL", "ETIT": "90%", "DPA": "74%", "Assertividade": "-" },
        { "Matricula": "N5802257", "Nome": "MAGNO FERRAREZ DE MORAIS", "Setor": "EMPRESARIAL", "ETIT": "96%", "DPA": "85%", "Assertividade": "-" },
        { "Matricula": "F201714", "Nome": "FERNANDA MESQUITA DE FREITAS", "Setor": "EMPRESARIAL", "ETIT": "90%", "DPA": "76%", "Assertividade": "-" },
        { "Matricula": "N6173055", "Nome": "JEFFERSON LUIS GONÇALVES COITINHO", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "72%", "Assertividade": "-" },
        { "Matricula": "N0125317", "Nome": "ROBERTO SILVA DO NASCIMENTO", "Setor": "EMPRESARIAL", "ETIT": "100%", "DPA": "91%", "Assertividade": "-" },
        { "Matricula": "N5819183", "Nome": "RODRIGO PIRES BERNARDINO", "Setor": "EMPRESARIAL", "ETIT": "94%", "DPA": "67%", "Assertividade": "-" },
        { "Matricula": "N5926003", "Nome": "SUELLEN HERNANDEZ DA SILVA", "Setor": "EMPRESARIAL", "ETIT": "88%", "DPA": "-", "Assertividade": "-" },
        { "Matricula": "N5932064", "Nome": "MONICA DA SILVA RODRIGUES", "Setor": "EMPRESARIAL", "ETIT": "100%", "DPA": "91%", "Assertividade": "-" },
        { "Matricula": "N5923221", "Nome": "KELLY PINHEIRO LIRA", "Setor": "RESIDENCIAL", "ETIT": "-", "DPA": "-", "Assertividade": "-" },
        { "Matricula": "N5772086", "Nome": "THIAGO PEREIRA DA SILVA", "Setor": "RESIDENCIAL", "ETIT": "100%", "DPA": "109%", "Assertividade": "80%" },
        { "Matricula": "N0239871", "Nome": "LEONARDO FERREIRA LIMA DE ALMEIDA", "Setor": "RESIDENCIAL", "ETIT": "100%", "DPA": "82%", "Assertividade": "100%" },
        { "Matricula": "N5577565", "Nome": "MARISTELLA MARCIA DOS SANTOS", "Setor": "RESIDENCIAL", "ETIT": "100%", "DPA": "85%", "Assertividade": "82%" },
        { "Matricula": "N5972428", "Nome": "CRISTIANE HERMOGENES DA SILVA", "Setor": "RESIDENCIAL", "ETIT": "100%", "DPA": "80%", "Assertividade": "88%" },
        { "Matricula": "N4014011", "Nome": "ALAN MARINHO DIAS", "Setor": "RESIDENCIAL", "ETIT": "100%", "DPA": "63%", "Assertividade": "100%" },
        { "Matricula": "F106664", "Nome": "RAISSA LIMA DE OLIVEIRA", "Setor": "RESIDENCIAL", "ETIT": "100%", "DPA": "98%", "Assertividade": "89%" }
    ];
    
    atualizarLookup();
}

/**
 * Atualiza o objeto de busca rápida e salva cache
 */
function atualizarLookup() {
    employeeLookup = {};
    employees.forEach(emp => {
        if (emp.Matricula) {
            employeeLookup[emp.Matricula.toUpperCase()] = emp;
        }
    });
    
    // Salva em cache local
    localStorage.setItem('claro_indicadores_cache', JSON.stringify(employees));
    localStorage.setItem('claro_indicadores_cache_time', Date.now().toString());
    
    console.log("Lookup atualizado:", Object.keys(employeeLookup).length, "funcionários");
}

/**
 * Mostra mensagem de status
 */
function mostrarStatus(mensagem, tipo = "info", timeout = null) {
    if (!statusDiv) return;
    
    statusDiv.textContent = mensagem;
    statusDiv.className = `status-info ${tipo}`;
    statusDiv.style.display = "block";
    
    if (timeout) {
        setTimeout(() => {
            statusDiv.style.display = "none";
        }, timeout);
    }
}

/**
 * Atualiza data e hora da última atualização
 */
function atualizarDataHora() {
    const agora = new Date();
    
    dataAtualizacaoSpan.textContent = 
        `${agora.getDate().toString().padStart(2, '0')}/` +
        `${(agora.getMonth() + 1).toString().padStart(2, '0')} ` +
        `${agora.getHours().toString().padStart(2, '0')}:` +
        `${agora.getMinutes().toString().padStart(2, '0')}`;
    
    lastUpdate = agora;
    isInitialLoad = false;
}

// ============================================
// FUNÇÕES DE NEGÓCIO
// ============================================

function definirMeta(setor, tipo) {
    if (tipo === "DPA") {
        return {
            certificacao: METAS.DPA.CERTIFICACAO,
            individual: METAS.DPA.INDIVIDUAL
        };
    }
    const setorNormalizado = setor.toUpperCase();
    if (tipo === "ETIT") return METAS.ETIT[setorNormalizado] || 0;
    return METAS.Assertividade[setorNormalizado] || 0;
}

function parseIndicatorValue(valor) {
    if (!valor || valor === "-" || valor === "–" || valor === "_" || 
        valor === "Não informado" || valor === "N/A") return null;
    
    // Remove % e converte vírgula para ponto
    const str = String(valor).replace('%', '').replace(',', '.').trim();
    const num = parseFloat(str);
    
    return isNaN(num) ? null : num;
}

function considerarDentroMeta(valor, setor, tipo, metaType = "individual") {
    const setorNormalizado = setor.toUpperCase();
    
    // Assertividade não se aplica ao setor EMPRESARIAL
    if (tipo === "Assertividade" && setorNormalizado === "EMPRESARIAL") {
        return true;
    }
    
    const valorNumerico = parseIndicatorValue(valor);
    if (valorNumerico === null) return true;
    
    const meta = tipo === "DPA" 
        ? definirMeta(setor, tipo)[metaType]
        : definirMeta(setor, tipo);
    
    if (meta === null) return true; // Se meta não se aplica
    
    return valorNumerico >= meta;
}

function formatarValor(valor) {
    if (!valor || valor === "-" || valor === "–" || valor === "_" || 
        valor === "Não informado") return "-";
    return valor;
}

function handleKeyPress(event) {
    if (event.key === "Enter") {
        consultar();
    }
}

function consultar() {
    const matricula = matriculaInput.value.trim().toUpperCase();
    resultadoDiv.innerHTML = "";
    
    if (!matricula) {
        resultadoDiv.innerHTML = "<p class='error'>Por favor, digite uma matrícula.</p>";
        return;
    }

    const empregado = employeeLookup[matricula];
    
    if (!empregado) {
        resultadoDiv.innerHTML = "<p class='error'>Matrícula não encontrada na base atual.</p>";
        return;
    }

    const setor = empregado.Setor.toUpperCase();
    
    // Verificar indicadores
    const etitOk = considerarDentroMeta(empregado.ETIT, setor, "ETIT");
    const assertividadeOk = setor === "EMPRESARIAL" ? null : considerarDentroMeta(empregado.Assertividade, setor, "Assertividade");
    const dpaCertificando = considerarDentroMeta(empregado.DPA, setor, "DPA", "certificacao");
    const dpaMetaIndividual = considerarDentroMeta(empregado.DPA, setor, "DPA", "individual");
    
    // Verificar certificação
    const certificando = etitOk && 
                       (setor === "EMPRESARIAL" || assertividadeOk) && 
                       dpaCertificando;
    
    const mensagemDPA = !dpaMetaIndividual && dpaCertificando ? 
        '<div class="meta-warning">Certificando, mas abaixo da meta individual (90%)</div>' : 
        '';

    // Formatar display da Assertividade para EMPRESARIAL
    const assertividadeDisplay = setor === "EMPRESARIAL" ?
        `<div class="indicator-row">
            <span class="indicator-name">Assertividade:</span>
            <span class="indicator-value not-applicable">N/A</span>
            <span class="meta-value">(Não se aplica)</span>
        </div>` :
        `<div class="indicator-row">
            <span class="indicator-name">Assertividade:</span>
            <span class="indicator-value ${assertividadeOk ? '' : 'warning'}">${formatarValor(empregado.Assertividade)}</span>
            <span class="meta-value">(Meta: ${definirMeta(setor, "Assertividade")}%)</span>
        </div>`;

    // Mostrar resultados
    resultadoDiv.innerHTML = 
        `<div class="employee-info">
            <h2>${empregado.Nome}</h2>
            <p><strong>Setor:</strong> ${empregado.Setor}</p>
            <p><small>Matrícula: ${empregado.Matricula}</small></p>
        </div>
        
        <div class="indicator-row">
            <span class="indicator-name">ETIT:</span>
            <span class="indicator-value ${etitOk ? '' : 'warning'}">${formatarValor(empregado.ETIT)}</span>
            <span class="meta-value">(Meta: ${definirMeta(setor, "ETIT")}%)</span>
        </div>
        
        ${assertividadeDisplay}
        
        <div class="indicator-row dpa-info">
            <span class="indicator-name">DPA:</span>
            <span class="indicator-value ${dpaMetaIndividual ? '' : 'warning'}">${formatarValor(empregado.DPA)}</span>
            <span class="meta-value">(Individual: ${METAS.DPA.INDIVIDUAL}%, Certificação: ${METAS.DPA.CERTIFICACAO}%)</span>
        </div>
        ${mensagemDPA}
        
        <div class="certification ${certificando ? 'success' : 'warning'}">
            ${certificando ? '✅ CERTIFICANDO' : '❌ NÃO CERTIFICANDO'}
        </div>
        
        <div class="info-rodape">
            <p><small>Dados atualizados em: ${dataAtualizacaoSpan.textContent}</small></p>
            <p><small>Fonte: SharePoint Claro</small></p>
        </div>`;
}

// ============================================
// INICIALIZAÇÃO
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    console.log("Sistema de Indicadores Claro - Inicializando...");
    
    // Tenta carregar cache primeiro para resposta rápida
    const cache = localStorage.getItem('claro_indicadores_cache');
    if (cache) {
        employees = JSON.parse(cache);
        atualizarLookup();
        mostrarStatus("📂 Dados em cache carregados", "info", 2000);
    }
    
    // Carrega dados do SharePoint
    carregarDadosSharePoint();
    
    // Event listeners
    matriculaInput.addEventListener('keypress', handleKeyPress);
    consultarBtn.addEventListener('click', consultar);
    
    // Botão de atualização
    refreshBtn.addEventListener('click', () => {
        refreshBtn.classList.add('loading');
        carregarDadosSharePoint().finally(() => {
            setTimeout(() => {
                refreshBtn.classList.remove('loading');
            }, 1000);
        });
    });
    
    // Atualizar automaticamente a cada 10 minutos
    setInterval(() => {
        console.log("Atualização automática do SharePoint");
        carregarDadosSharePoint();
    }, 10 * 60 * 1000);
    
    // Focar no campo de matrícula
    matriculaInput.focus();
    
    // Sugestão de matrícula
    matriculaInput.addEventListener('focus', () => {
        if (!matriculaInput.value) {
            matriculaInput.placeholder = "Ex: N6088107, N5619600...";
        }
    });
});
