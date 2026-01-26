/**
 * ===================================================================================
 * MODELO_ETP.JS - GERADOR DE MODELO EM BRANCO (MODO ADMINISTRATIVO)
 * ===================================================================================
 * 
 * Este script adiciona uma funcionalidade oculta para gerar um documento DOCX em branco
 * contendo todos os campos do ETP, formatado de acordo com o padrão oficial.
 * 
 * COMO ACESSAR:
 * Adicione "?modo=admin" ao final da URL do aplicativo (ex: index.html?modo=admin).
 * 
 */

document.addEventListener('DOMContentLoaded', () => {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('modo') === 'admin') {
        initAdminOverlay();
    }
});

/**
 * Cria a interface administrativa sobreposta ao app.
 */
function initAdminOverlay() {
    const elementsToHide = ['.tab-container', '.etp-toolbar', 'header', '#beta-alert-banner'];
    elementsToHide.forEach(selector => {
        const el = document.querySelector(selector);
        if (el) el.style.display = 'none';
    });

    const overlay = document.createElement('div');
    Object.assign(overlay.style, {
        position: 'fixed', top: '0', left: '0', width: '100%', height: '100%',
        backgroundColor: '#f4f6f9', display: 'flex', justifyContent: 'center', alignItems: 'center',
        zIndex: '9999', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    overlay.innerHTML = `
        <div style="background: white; padding: 50px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); text-align: center; width: 100%; max-width: 500px; border: 1px solid #dee2e6;">
            <div style="font-size: 52px; color: #004080; margin-bottom: 25px;"><i class="fas fa-file-word"></i></div>
            <h1 style="color: #333; font-size: 24px; margin-bottom: 15px; font-weight: 600;">Gerador de Modelos de ETP</h1>
            <p style="color: #666; margin-bottom: 30px; line-height: 1.6;">
                Selecione qual estrutura de modelo em branco você deseja baixar.
            </p>
            
            <div style="display: flex; flex-direction: column; gap: 15px;">
                <button id="btnGenCompleto" style="padding: 14px; font-size: 16px; background-color: #0056b3; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 500; transition: 0.2s;">
                    <i class="fas fa-download" style="margin-right: 8px;"></i> Modelo ETP Completo
                </button>
                
                <button id="btnGenSimplificado" style="padding: 14px; font-size: 16px; background-color: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 500; transition: 0.2s;">
                    <i class="fas fa-download" style="margin-right: 8px;"></i> Modelo ETP Simplificado
                </button>
            </div>

            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px;">
                <a href="${window.location.pathname}" style="color: #6c757d; text-decoration: none; font-size: 14px; display: inline-flex; align-items: center; gap: 5px;">
                    <i class="fas fa-arrow-left"></i> Voltar ao Aplicativo
                </a>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);

    // Listeners para os botões
    document.getElementById('btnGenCompleto').addEventListener('click', () => generateBlankTemplate(false));
    document.getElementById('btnGenSimplificado').addEventListener('click', () => generateBlankTemplate(true));
}

// --- DICIONÁRIO DE NOMES AMIGÁVEIS ---
// Mapeia o ID do grupo condicional (do config.js) para o texto que aparecerá na nota do DOCX.
const friendlyConditionNames = {
    // Identificação
    "etp_tipo": "tipo de ETP",
    "etp_auth": "autorização de ETP simplificado",
    // Cap 1
    "alinhamento_estrategico": "alinhamento estratégico",
    // Cap 3
    "natureza_continua": "natureza contínua",
    "beneficios_vigencia": "benefícios da vigência",
    "garantia": "garantia",
    "justificativa_garantia": "justificativa de garantia fora do padrão",
    "assistencia": "assistência técnica",
    "justificativa_assistencia": "justificativa de assistência fora do padrão",
    "transferencia": "transferência de conhecimento",
    "capacitacao": "capacitação",
    "contratacao_adicional": "contratação adicional",
    "ajuste_contratacoes": "ajuste em contratações",
    "acessibilidade": "acessibilidade",
    // Cap 7
    "parcelamento": "forma de parcelamento", // Nome genérico para o select pai
    "parcelamento_grupo_unico": "agrupamento único",
    "parcelamento_grupos_separados": "grupos separados",
    // Cap 8
    "modalidade_contratacao": "modalidade da contratação",
    "contratacao_dispensa": "dispensa de licitação",
    "contratacao_inexigibilidade": "inexigibilidade de licitação",
    "dispensa_outra_hipotese": "outra hipótese de dispensa",
    "qualificacao_tecnica": "qualificação técnica",
    "amostras": "amostras ou prova de conceito",
    "vistoria": "vistoria",
    "confidencialidade": "confidencialidade",
    "subcontratacao": "subcontratação",
    "consorcio": "formação de consórcio", // Select pai
    "consorcio_proibicao": "proibição de consórcio",
    "consorcio_permissao": "permissão de consórcio",
    "limite_consorcio": "limite de consorciadas",
    "cooperativas": "participação de cooperativas",
    "cooperativas_proibicao": "proibição de cooperativas",
    "estrangeiras": "empresas estrangeiras",
    "estrangeiras_proibicao": "proibição de estrangeiras",
    "estrangeiras_permissao": "permissão de estrangeiras",
    "margem_preferencia": "margem de preferência",
    "pessoa_fisica": "participação de pessoa física",
    "pessoa_fisica_proibicao": "proibição de pessoa física",
    // Cap 9
    "quantitativo_inferior": "quantitativo inferior",
    "precos_diferentes": "preços diferentes",
    "mais_de_um_fornecedor": "mais de um fornecedor",
    "adesao_futura": "adesão futura",
    "prorrogacao_ata": "prorrogação da ata",
    "renovacao_quantidades_ata": "renovação de quantidades"
};

/**
 * Retorna o nome amigável da condição que ativa um determinado campo alvo.
 * @param {string} targetFieldId - O ID do div conditional que está sendo exibido.
 */
function getFriendlyTriggerText(targetFieldId) {
    if (typeof conditionalFieldIds === 'undefined') return "opção anterior";

    // Procura qual chave em conditionalFieldIds contém o targetFieldId
    for (const [triggerKey, targets] of Object.entries(conditionalFieldIds)) {
        if (targets.includes(targetFieldId)) {
            // Se encontrou, retorna o nome amigável ou a própria chave se não houver tradução
            return friendlyConditionNames[triggerKey] || triggerKey;
        }
    }
    return "opção anterior";
}

/**
 * Função Principal de Geração do DOCX
 */
/**
 * Função Principal de Geração do DOCX
 */
function generateBlankTemplate(isSimplificado) {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, BorderStyle, AlignmentType } = window.docx;

    const FONT_FAMILY = "Calibri";
    const COLOR_BLACK = "000000";
    const docTitleText = isSimplificado ? "ESTUDO TÉCNICO PRELIMINAR SIMPLIFICADO" : "ESTUDO TÉCNICO PRELIMINAR COMPLETO";

    // Definição de Estilos com Espaçamentos Ajustados
    const styles = {
        docTitle: { 
            size: 32, 
            bold: true, 
            font: FONT_FAMILY, 
            color: COLOR_BLACK, 
            allCaps: true 
        },
        chapterTitle: { 
            size: 28, 
            bold: true, 
            font: FONT_FAMILY, 
            color: COLOR_BLACK, 
            // AUMENTADO: before: 1200 para duplicar o espaço antes do capítulo
            spacing: { before: 1400, after: 600 } 
        },
        dynamicSectionTitle: { 
            size: 24, 
            bold: true, 
            font: FONT_FAMILY, 
            color: "333333", 
            italics: true, 
            spacing: { before: 800, after: 400 } 
        },
        fieldLabel: { 
            size: 22, 
            bold: true, 
            font: FONT_FAMILY, 
            color: COLOR_BLACK, 
            // AUMENTADO: before: 600 cria uma linha em branco antes de cada item/subitem
            spacing: { before: 600, after: 300 } 
        },
        instruction: { 
            size: 20, 
            italics: true, 
            color: "666666", 
            font: FONT_FAMILY, 
            spacing: { after: 100 } 
        },
        placeholder: { 
            size: 22, 
            color: "555555", 
            italics: true, 
            font: FONT_FAMILY 
        },
        optionText: { 
            size: 22, 
            color: COLOR_BLACK, 
            font: FONT_FAMILY 
        }
    };

    const docChildren = [];

    // Título Principal do Documento
    docChildren.push(new Paragraph({
        children: [new TextRun({ text: docTitleText, ...styles.docTitle })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 }
    }));

    const tabOrder = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];

    tabOrder.forEach(tabId => {
        const tabEl = document.getElementById(tabId);
        if (!tabEl) return;

        const isChapterHiddenInSimplificado = tabEl.classList.contains('simplificado-hide');

        // 1. Título do Capítulo (O espaçamento 'before' duplicado está no estilo chapterTitle)
        const h2 = tabEl.querySelector('h2');
        if (h2) {
            let titleText = h2.innerText.trim().replace(/^\W+/, ''); // Remove ícones/espaços iniciais
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: titleText, ...styles.chapterTitle })],
                heading: HeadingLevel.HEADING_1,
                border: { bottom: { color: "BFBFBF", space: 4, value: "single", size: 6 } }
            }));
        }

        // 2. Se o capítulo inteiro não se aplica ao simplificado
        if (isSimplificado && isChapterHiddenInSimplificado) {
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: "Nota: Este capítulo não se aplica ao ETP Simplificado.", ...styles.instruction })],
                spacing: { after: 400 }
            }));
            return; // Pula para o próximo capítulo
        }

        // 3. Blocos Dinâmicos (Soluções, Riscos, etc.)
        const dynamicContainers = {
            'cap2': { id: 'solucoes_mercado_container', config: dynamicItemConfigs.solucao, label: "Solução de Mercado" },
            'cap5': { id: 'contratacoes_anteriores_container', config: dynamicItemConfigs.contratacao, label: "Contratação Anterior" },
            'cap6': { id: 'riscos_container', config: dynamicItemConfigs.risco, label: "Risco Identificado" },
            'anexos': { id: 'anexos_container', config: dynamicItemConfigs.anexo, label: "Anexo" }
        };

        if (dynamicContainers[tabId]) {
            const dynData = dynamicContainers[tabId];
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: `Exemplo de Bloco Dinâmico: ${dynData.label}`, ...styles.dynamicSectionTitle })],
                spacing: { before: 200 }
            }));
            
            // Renderiza um template "em branco" do item dinâmico para pegar os labels
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = dynData.config.template(1, false); 
            
            // Itera sobre os SUBITENS do bloco dinâmico
            tempDiv.querySelectorAll('.form-group').forEach(group => {
                const lbl = group.querySelector('label');
                if(lbl) {
                    const labelText = lbl.textContent.trim().replace(/\d+\.1\./, 'X.');
                    // O estilo 'fieldLabel' aqui garante o espaçamento antes de cada subitem
                    docChildren.push(new Paragraph({ 
                        children: [new TextRun({ text: labelText, ...styles.fieldLabel })],
                        keepNext: true // Mantém título junto com a resposta
                    }));
                    
                    // Verifica se dentro do template há selects
                    const select = group.querySelector('select');
                    if (select) {
                        Array.from(select.options).forEach(opt => {
                            if (opt.value) { 
                                docChildren.push(new Paragraph({
                                    children: [
                                        new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }),
                                        new TextRun({ text: opt.text, ...styles.optionText })
                                    ],
                                    indent: { left: 400 }
                                }));
                            }
                        });
                    } else {
                        // Caso contrário (textarea/input), placeholder padrão
                        docChildren.push(new Paragraph({ children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })], spacing: { after: 200 } }));
                    }
                }
            });
        }

        // 4. Campos Estáticos do Formulário
        tabEl.querySelectorAll('.form-group').forEach(group => {
            // Ignora campos que estão dentro de templates dinâmicos
            if (group.closest('.solucao-item, .risco-item, .contratacao-item, .anexo-item')) return;
            
            // Filtro de itens desativados no Simplificado
            if (isSimplificado && group.classList.contains('simplificado-hide')) return;

            // Tratamento especial para o item 3.1
            if (tabId === 'cap3') {
                const isCompletoWrapper = group.id === 'c3_1_completo_wrapper';
                const isSimplificadoWrapper = group.id === 'c3_1_simplificado_wrapper';
                if (isSimplificado && isCompletoWrapper) return;
                if (!isSimplificado && isSimplificadoWrapper) return;
            }

            const label = group.querySelector('label') || group.querySelector('.label-with-help label');
            if (!label) return;

            // Imprime o Título do Item
            // O estilo 'fieldLabel' aqui garante o espaçamento antes de cada item principal
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: label.textContent.trim(), ...styles.fieldLabel })],
                keepNext: true
            }));

            // 5. Nota Condicional (com nomes amigáveis)
            if (group.classList.contains('conditional-field')) {
                const friendlyName = getFriendlyTriggerText(group.id);
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: `Nota: Preencher somente se necessário conforme a opção "${friendlyName}".`, ...styles.instruction })]
                }));
            }

            // 6. Renderização das Opções de Resposta
            const checkOptions = group.querySelectorAll('.checkbox-group label');
            const selectElement = group.querySelector('select');

            if (checkOptions.length > 0) {
                // Renderiza Radio/Checkboxes
                checkOptions.forEach(opt => {
                    docChildren.push(new Paragraph({
                        children: [
                            new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }),
                            new TextRun({ text: opt.textContent.trim(), ...styles.optionText })
                        ],
                        indent: { left: 400 }
                    }));
                });
            } else if (selectElement) {
                // Renderiza Select como lista de opções
                const options = Array.from(selectElement.options);
                let hasValidOptions = false;
                
                options.forEach(opt => {
                    if (opt.value && opt.value.trim() !== "") {
                        hasValidOptions = true;
                        docChildren.push(new Paragraph({
                            children: [
                                new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }),
                                new TextRun({ text: opt.text.trim(), ...styles.optionText })
                            ],
                            indent: { left: 400 }
                        }));
                    }
                });

                if (!hasValidOptions) {
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })],
                        spacing: { after: 200 }
                    }));
                }

            } else {
                // Caso C: Inputs de Texto / Textarea
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })],
                    spacing: { after: 200 }
                }));
            }
        });
    });

    const filename = isSimplificado ? "Modelo_ETP_Simplificado.docx" : "Modelo_ETP_Completo.docx";

    // Gera e salva o arquivo
    Packer.toBlob(new Document({
        sections: [{
            properties: { page: { margin: { top: 1440, bottom: 1440, left: 1134, right: 1134 } } },
            children: docChildren,
        }],
    })).then(blob => saveAs(blob, filename));
}