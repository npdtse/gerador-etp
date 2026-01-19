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
    // Esconde o conteúdo principal visualmente para focar na ferramenta admin
    const elementsToHide = ['.tab-container', '.etp-toolbar', 'header', '#beta-alert-banner'];
    elementsToHide.forEach(selector => {
        const el = document.querySelector(selector);
        if (el) el.style.display = 'none';
    });

    // Cria overlay Admin
    const overlay = document.createElement('div');
    Object.assign(overlay.style, {
        position: 'fixed', top: '0', left: '0', width: '100%', height: '100%',
        backgroundColor: '#f4f6f9', display: 'flex', justifyContent: 'center', alignItems: 'center',
        zIndex: '9999', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    overlay.innerHTML = `
        <div style="background: white; padding: 50px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); text-align: center; max-width: 600px; border: 1px solid #dee2e6;">
            <div style="font-size: 52px; color: #004080; margin-bottom: 25px;"><i class="fas fa-file-word"></i></div>
            <h1 style="color: #333; font-size: 26px; margin-bottom: 15px; font-weight: 600;">Gerador de Modelo de ETP</h1>
            <p style="color: #666; margin-bottom: 35px; line-height: 1.6; font-size: 1.05em;">
                Esta ferramenta gera um arquivo <strong>.docx</strong> contendo a estrutura completa do ETP, 
                incluindo campos dinâmicos e regras condicionais.<br><br>
                O documento gerado serve como <em>template</em> para preenchimento manual no SEI.
            </p>
            <button id="btnGenTemplate" style="padding: 14px 28px; font-size: 16px; background-color: #0056b3; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 500; transition: background 0.2s;">
                <i class="fas fa-download" style="margin-right: 8px;"></i> Baixar Modelo em Branco
            </button>
            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px;">
                <a href="${window.location.pathname}" style="color: #6c757d; text-decoration: none; font-size: 14px; display: inline-flex; align-items: center; gap: 5px;">
                    <i class="fas fa-arrow-left"></i> Voltar ao Aplicativo
                </a>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);

    const btn = document.getElementById('btnGenTemplate');
    btn.addEventListener('click', generateBlankTemplate);
    btn.addEventListener('mouseover', () => btn.style.backgroundColor = '#004080');
    btn.addEventListener('mouseout', () => btn.style.backgroundColor = '#0056b3');
}

/**
 * Mapeia os IDs dos campos condicionais de volta para o nome do campo que os ativa.
 */
function buildReverseConditionalMap() {
    const map = {};
    if (typeof conditionalFieldIds !== 'undefined') {
        for (const [triggerId, targets] of Object.entries(conditionalFieldIds)) {
            targets.forEach(targetId => {
                let triggerLabelText = triggerId;
                
                // Tenta achar inputs com esse name (radio) ou id (select)
                const triggerEl = document.getElementById(triggerId) || document.querySelector(`input[name="${triggerId}"]`);
                if (triggerEl) {
                    // Busca o label mais próximo
                    const labelEl = triggerEl.closest('.form-group')?.querySelector('label');
                    if (labelEl) {
                        // Limpa o texto do label (remove numeração se houver para ficar mais limpo)
                        triggerLabelText = labelEl.textContent.trim().replace(/^\d+(\.\d+)*\s*/, '');
                    }
                }
                map[targetId] = triggerLabelText;
            });
        }
    }
    return map;
}

/**
 * Função Principal de Geração do DOCX
 */
function generateBlankTemplate() {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, BorderStyle, AlignmentType } = window.docx;

    // --- Definições de Estilo (Baseado no script.js e identidade visual oficial) ---
    const FONT_FAMILY = "Calibri"; // Padrão oficial
    const COLOR_BLACK = "000000";
    const COLOR_DARK_GRAY = "333333";
    const COLOR_LIGHT_GRAY = "666666";

    const styles = {
        docTitle: { 
            size: 32, // 16pt
            bold: true, 
            font: FONT_FAMILY, 
            color: COLOR_BLACK,
            allCaps: true
        },
        chapterTitle: { 
            size: 28, // 14pt
            bold: true, 
            font: FONT_FAMILY, 
            color: COLOR_BLACK, // Mantendo preto para sobriedade
            spacing: { before: 600, after: 200 } // Espaço maior antes de novos capítulos
        },
        dynamicSectionTitle: {
            size: 24, // 12pt
            bold: true,
            font: FONT_FAMILY,
            color: COLOR_DARK_GRAY,
            italics: true,
            spacing: { before: 300, after: 100 }
        },
        fieldLabel: { 
            size: 22, // 11pt
            bold: true, 
            font: FONT_FAMILY, 
            color: COLOR_BLACK, 
            spacing: { before: 300, after: 100 } 
        },
        instruction: { 
            size: 20, // 10pt
            italics: true, 
            color: COLOR_LIGHT_GRAY, 
            font: FONT_FAMILY,
            spacing: { after: 100 }
        },
        placeholder: { 
            size: 22, // 11pt
            color: "555555", 
            italics: true,
            font: FONT_FAMILY 
        },
        optionText: {
            size: 22, // 11pt
            font: FONT_FAMILY,
            color: COLOR_BLACK
        }
    };

    const docChildren = [];
    const conditionalMap = buildReverseConditionalMap();

    // 1. Cabeçalho do Documento
    docChildren.push(new Paragraph({
        children: [new TextRun({ text: "ESTUDO TÉCNICO PRELIMINAR", ...styles.docTitle })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 }
    }));
    
    docChildren.push(new Paragraph({
        children: [new TextRun({ text: "(Modelo para Preenchimento)", size: 24, font: FONT_FAMILY, italics: true })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 }
    }));

    // Ordem das abas para varredura
    const tabOrder = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];

    tabOrder.forEach(tabId => {
        const tabEl = document.getElementById(tabId);
        if (!tabEl) return;

        // --- Título do Capítulo ---
        const h2 = tabEl.querySelector('h2');
        if (h2) {
            let titleText = h2.innerText.trim();
            // Remove ícones do FontAwesome se tiverem sido capturados no innerText
            titleText = titleText.replace(/^\W+/, ''); 

            docChildren.push(new Paragraph({
                children: [new TextRun({ text: titleText, ...styles.chapterTitle })],
                heading: HeadingLevel.HEADING_1,
                border: { bottom: { color: "BFBFBF", space: 4, value: "single", size: 6 } } // Linha separadora
            }));
        }

        // --- Configuração para Itens Dinâmicos (Soluções, Riscos, etc) ---
        const dynamicContainers = {
            'cap2': { id: 'solucoes_mercado_container', config: dynamicItemConfigs.solucao, label: "Solução de Mercado" },
            'cap5': { id: 'contratacoes_anteriores_container', config: dynamicItemConfigs.contratacao, label: "Contratação Anterior" },
            'cap6': { id: 'riscos_container', config: dynamicItemConfigs.risco, label: "Risco Identificado" },
            'anexos': { id: 'anexos_container', config: dynamicItemConfigs.anexo, label: "Anexo" }
        };

        if (dynamicContainers[tabId]) {
            const dynData = dynamicContainers[tabId];
            
            // Bloco de instrução para item dinâmico
            docChildren.push(new Paragraph({
                children: [
                    new TextRun({ text: `Exemplo de Preenchimento: ${dynData.label}`, ...styles.dynamicSectionTitle })
                ],
                border: { left: { color: "E0E0E0", space: 10, value: "single", size: 20 } }, // Barra lateral para destacar
                spacing: { before: 300, after: 100 }
            }));
            
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: `(Nota: Copie e cole este bloco abaixo para cada ${dynData.label.toLowerCase()} adicional necessária)`, ...styles.instruction })],
                spacing: { after: 300 }
            }));

            // Gera HTML temporário para extrair labels do template do config.js
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = dynData.config.template(1, false);
            
            // Extrai labels do template
            const dynLabels = tempDiv.querySelectorAll('label');
            dynLabels.forEach(lbl => {
                let labelText = lbl.textContent.trim();
                
                // Remove a numeração fixa "1." para ficar genérico (ex: "2.1.1" vira "Título...")
                // Ou mantém para referência. Vamos manter a referência "X.1" para indicar hierarquia.
                labelText = labelText.replace(/\d+\.1\./, 'X.'); 

                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: labelText, ...styles.fieldLabel })]
                }));
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })],
                    spacing: { after: 200 }
                }));
            });

            // Adiciona um separador visual de fim de bloco dinâmico
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: "--- Fim do Bloco Dinâmico ---", size: 18, color: "999999" })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 300, after: 400 }
            }));
        }

        // --- Campos Estáticos do Capítulo ---
        const formGroups = Array.from(tabEl.querySelectorAll('.form-group'));
        
        formGroups.forEach(group => {
            // Ignora se estiver dentro de um item dinâmico (já tratados acima)
            if (group.closest('.solucao-item, .risco-item, .contratacao-item, .anexo-item')) return;

            const label = group.querySelector('label');
            let fieldTitle = "";
            let isConditional = false;
            let conditionText = "";

            // Verifica Condicional
            if (group.id && group.classList.contains('conditional-field')) {
                isConditional = true;
                const triggerName = conditionalMap[group.id];
                conditionText = triggerName 
                    ? `Preencher somente se a resposta anterior referente a "${triggerName}" indicar necessidade.` 
                    : "Campo de preenchimento condicional.";
            }

            // Extração do Título do Campo
            if (label) {
                const inputId = label.getAttribute('for');
                const input = document.getElementById(inputId);
                
                // Lógica para Radios/Checkboxes: Pegar o título do grupo, não da opção
                if (input && (input.type === 'radio' || input.type === 'checkbox')) {
                    if (!group.classList.contains('checkbox-group')) {
                         // É um label de opção isolada (ex: "Sim"), tentamos pegar o pai
                         const parentLabel = group.closest('.form-group')?.querySelector('.label-with-help label');
                         if(parentLabel) fieldTitle = parentLabel.textContent.trim();
                         else fieldTitle = label.textContent.trim(); // Fallback
                    }
                } else {
                    fieldTitle = label.textContent.trim();
                }
            } else {
                // Tenta pegar texto de label-with-help (comum no seu HTML)
                const labelWithHelp = group.querySelector('.label-with-help label');
                if (labelWithHelp) fieldTitle = labelWithHelp.textContent.trim();
            }

            if (fieldTitle) {
                // 1. Label do Campo
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: fieldTitle, ...styles.fieldLabel })],
                    keepNext: true // Tenta manter o label junto com a resposta/opções
                }));

                // 2. Instrução Condicional (se houver)
                if (isConditional) {
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: `Nota: ${conditionText}`, ...styles.instruction })]
                    }));
                }

                // 3. Área de Resposta (Opções ou Texto)
                const radioGroup = group.querySelector('.checkbox-group');
                
                if (radioGroup) {
                    // Se for múltipla escolha, lista as opções visualmente
                    const options = radioGroup.querySelectorAll('label');
                    options.forEach(opt => {
                        docChildren.push(new Paragraph({
                            children: [
                                new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }), // Caixa vazia simulada
                                new TextRun({ text: opt.textContent.trim(), ...styles.optionText })
                            ],
                            indent: { left: 400 }, // Recuo para parecer lista
                            spacing: { after: 100 }
                        }));
                    });
                } else {
                    // Campo de Texto Livre
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })],
                        spacing: { after: 200 }
                    }));
                }
            }
        });
    });

    // --- Geração e Download ---
    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: 1440, // 1 polegada (twips)
                        bottom: 1440,
                        left: 1134, // ~2 cm
                        right: 1134
                    }
                }
            },
            children: docChildren,
        }],
    });

    Packer.toBlob(doc).then(blob => {
        saveAs(blob, "Modelo_ETP_Em_Branco.docx");
    });
}