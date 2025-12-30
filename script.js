// ============================
// CONFIGURAÇÃO DO ONEDRIVE
// ============================
// Substitua este link pelo link de download direto do seu arquivo Excel
const ONEDRIVE_EXCEL_URL = 'https://corpclarobr-my.sharepoint.com/personal/fernando_pereiracunha_claro_com_br/_layouts/15/download.aspx?share=SUA_CHAVE_AQUI';

// Ou use este formato alternativo se o acima não funcionar
// const ONEDRIVE_EXCEL_URL = 'https://corpclarobr.sharepoint.com/sites/SEU_SITE/_layouts/15/download.aspx?UniqueId=SEU_ID&Translate=false&tempauth=SEU_TOKEN';

// ============================
// CACHE DE ELEMENTOS DOM
// ============================
const matriculaInput = document.getElementById("matricula");
const resultadoDiv = document.getElementById("resultado");
const consultarBtn = document.getElementById("consultar-btn");
const statusDiv = document.getElementById("status");
const dataAtualizacaoSpan = document.getElementById("data-atualizacao");
const refreshBtn = document.getElementById("refresh-btn");

// ============================
// VARIÁVEIS GLOBAIS
// ============================
let employees = [];
let employeeLookup = {};
let lastUpdate = null;

// ============================
// CONFIGURAÇÃO DE METAS
// ============================
const METAS = {
    "ETIT": {
        "MÓVEL": 80,
        "RESIDENCIAL": 90,
        "EMPRESARIAL": 90
    },
    "Assertividade": {
        "MÓVEL": 85,
        "RESIDENCIAL": 70,
        "EMPRESARIAL": null // Não se aplica
    },
    "DPA": {
        "CERTIFICACAO": 85,
        "INDIVIDUAL": 90
    }
};

// ============================
// FUNÇÕES PRINCIPAIS
// ============================

/**
 * Carrega dados do OneDrive Excel
 */
async function carregarDadosOneDrive() {
    try {
        mostrarStatus("Conectando ao OneDrive...", "loading");
        
        // Baixar o arquivo Excel
        const response = await fetch(ONEDRIVE_EXCEL_URL);
        
        if (!response.ok) {
            throw new Error(`Erro HTTP: ${response.status} ${response.statusText}`);
        }
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Ler o Excel usando SheetJS
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Pegar a primeira planilha
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Converter para JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: "-"
        });
        
        // Encontrar os índices das colunas
        const headerRow = jsonData[0];
        const colIndices = {
            matricula: headerRow.findIndex(col => 
                col && (col.toLowerCase().includes('matricula') || col.toLowerCase().includes('matrícula'))
            ),
            nome: headerRow.findIndex(col => 
                col && col.toLowerCase().includes('nome')
            ),
            setor: headerRow.findIndex(col => 
                col && col.toLowerCase().includes('setor')
            ),
            etit: headerRow.findIndex(col => 
                col && col.toLowerCase().includes('etit')
            ),
            dpa: headerRow.findIndex(col => 
                col && col.toLowerCase().includes('dpa')
            ),
            assertividade: headerRow.findIndex(col => 
                col && (col.toLowerCase().includes('assertividade') || col.toLowerCase().includes('acerto'))
            )
        };
        
        // Processar os dados
        employees = [];
        
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            
            // Verificar se é uma linha válida
            if (!row || !row[colIndices.matricula]) continue;
            
            const emp = {
                Matricula: String(row[colIndices.matricula] || "").trim().toUpperCase(),
                Nome: String(row[colIndices.nome] || "").trim(),
                Setor: String(row[colIndices.setor] || "").trim(),
                ETIT: formatarValorExcel(row[colIndices.etit]),
                DPA: formatarValorExcel(row[colIndices.dpa]),
                Assertividade: formatarValorExcel(row[colIndices.assertividade])
            };
            
            employees.push(emp);
        }
        
        // Atualizar lookup
        atualizarLookup();
        
        // Atualizar data/hora
        atualizarDataHora();
        
        // Mostrar sucesso
        mostrarStatus(`✓ Dados atualizados: ${employees.length} registros`, "success", 3000);
        
        console.log("Dados carregados com sucesso:", employees);
        
    } catch (error) {
        console.error("Erro ao carregar dados do OneDrive:", error);
        
        // Tentar usar cache local ou dados de fallback
        const cached = localStorage.getItem('claro_indicator_cache');
        if (cached) {
            const { data, timestamp } = JSON.parse(cached);
            
            // Verificar se o cache tem menos de 1 hora
            const umaHora = 60 * 60 * 1000;
            if (Date.now() - timestamp < umaHora) {
                employees = data;
                atualizarLookup();
                mostrarStatus("Usando dados em cache (última atualização: " + 
                    new Date(timestamp).toLocaleTimeString('pt-BR') + ")", "warning", 3000);
                return;
            }
        }
        
        // Usar dados de fallback
        usarDadosFallback();
        mostrarStatus("⚠ Usando dados locais (erro no OneDrive)", "error", 5000);
    }
}

/**
 * Formata valores do Excel (remove formatação de porcentagem)
 */
function formatarValorExcel(valor) {
    if (!valor && valor !== 0) return "-";
    
    const strValor = String(valor);
    
    // Se já tiver %, manter
    if (strValor.includes('%')) return strValor;
    
    // Se for número, converter para porcentagem
    if (!isNaN(valor) && valor !== null) {
        return Math.round(valor * 100) + "%";
    }
    
    return strValor;
}

/**
 * Atualiza o objeto de busca rápida
 */
function atualizarLookup() {
    employeeLookup = {};
    employees.forEach(emp => {
        if (emp.Matricula) {
            employeeLookup[emp.Matricula.toUpperCase()] = emp;
        }
    });
    
    // Salvar em cache
    salvarCache();
}

/**
 * Salva dados em cache local
 */
function salvarCache() {
    const cacheData = {
        data: employees,
        timestamp: Date.now()
    };
    localStorage.setItem('claro_indicator_cache', JSON.stringify(cacheData));
}

/**
 * Dados de fallback (caso o OneDrive falhe)
 */
function usarDadosFallback() {
    employees = [
        // ... seus dados atuais aqui ...
        { "Matricula": "N6088107", "Nome": "LEANDRO GONÇALVES DE CARVALHO", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "64%", "Assertividade": "-" },
        { "Matricula": "N5619600", "Nome": "BRUNO COSTA BUCARD", "Setor": "EMPRESARIAL", "ETIT": "-", "DPA": "60%", "Assertividade": "-" },
        // ... resto dos dados ...
    ];
    atualizarLookup();
}

/**
 * Mostra mensagem de status
 */
function mostrarStatus(mensagem, tipo = "info", timeout = null) {
    statusDiv.textContent = mensagem;
    statusDiv.className = "status-info " + tipo;
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
    const options = { 
        day: '2-digit', 
        month: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    };
    
    dataAtualizacaoSpan.textContent = 
        `${agora.getDate().toString().padStart(2, '0')}/` +
        `${(agora.getMonth() + 1).toString().padStart(2, '0')} ` +
        `${agora.getHours().toString().padStart(2, '0')}:` +
        `${agora.getMinutes().toString().padStart(2, '0')}`;
    
    lastUpdate = agora;
}

// ============================
// FUNÇÕES DE NEGÓCIO (mantidas)
// ============================

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
    if (valor === "-" || valor === "–" || valor === "_" || valor === "Não informado" || !valor) return null;
    return parseFloat(valor.toString().replace("%", "").replace(",", "."));
}

function considerarDentroMeta(valor, setor, tipo, metaType = "individual") {
    const setorNormalizado = setor.toUpperCase();
    
    if (tipo === "Assertividade" && setorNormalizado === "EMPRESARIAL") {
        return true;
    }
    
    const valorNumerico = parseIndicatorValue(valor);
    if (valorNumerico === null) return true;
    
    const meta = tipo === "DPA" 
        ? definirMeta(setor, tipo)[metaType]
        : definirMeta(setor, tipo);
        
    return valorNumerico >= meta;
}

function formatarValor(valor) {
    if (!valor || valor === "-" || valor === "–" || valor === "_" || valor === "Não informado") return "-";
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
        resultadoDiv.innerHTML = "<p class='error'>Matrícula não encontrada.</p>";
        return;
    }

    const setor = empregado.Setor.toUpperCase();
    
    // Verificar indicadores
    const etitOk = considerarDentroMeta(empregado.ETIT, setor, "ETIT");
    const assertividadeOk = setor === "EMPRESARIAL" ? null : considerarDentroMeta(empregado.Assertividade, setor, "Assertividade");
    const dpaCertificando = considerarDentroMeta(empregado.DPA, setor, "DPA", "certificacao");
    const dpaMetaIndividual = considerarDentroMeta(empregado.DPA, setor, "DPA", "individual");
    
    // Para certificação, Assertividade não conta para EMPRESARIAL
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
            <span class="meta-value">(Meta Individual: ${METAS.DPA.INDIVIDUAL}%, Certificação: ${METAS.DPA.CERTIFICACAO}%)</span>
        </div>
        ${mensagemDPA}
        
        <div class="certification ${certificando ? 'success' : 'warning'}">
            ${certificando ? '✅ Certificando' : '❌ Não certificando'}
        </div>
        
        <div class="atualizacao-info">
            <small>Dados atualizados: ${dataAtualizacaoSpan.textContent}</small>
        </div>`;
}

// ============================
// EVENT LISTENERS
// ============================

document.addEventListener('DOMContentLoaded', () => {
    // Carregar dados ao iniciar
    carregarDadosOneDrive();
    
    // Event listeners
    matriculaInput.addEventListener('keypress', handleKeyPress);
    consultarBtn.addEventListener('click', consultar);
    
    // Botão de atualização
    refreshBtn.addEventListener('click', () => {
        refreshBtn.classList.add('loading');
        carregarDadosOneDrive().finally(() => {
            setTimeout(() => {
                refreshBtn.classList.remove('loading');
            }, 1000);
        });
    });
    
    // Atualizar a cada 15 minutos (900000 ms)
    setInterval(() => {
        carregarDadosOneDrive();
    }, 15 * 60 * 1000);
    
    // Focar no campo de matrícula
    matriculaInput.focus();
});
