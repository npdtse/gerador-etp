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
function generateBlankTemplate(isSimplificado) {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, BorderStyle, AlignmentType } = window.docx;

    const FONT_FAMILY = "Calibri";
    const COLOR_BLACK = "000000";
    const docTitleText = isSimplificado ? "ESTUDO TÉCNICO PRELIMINAR SIMPLIFICADO" : "ESTUDO TÉCNICO PRELIMINAR";

    const styles = {
        docTitle: { size: 32, bold: true, font: FONT_FAMILY, color: COLOR_BLACK, allCaps: true },
        chapterTitle: { size: 28, bold: true, font: FONT_FAMILY, color: COLOR_BLACK, spacing: { before: 600, after: 200 } },
        dynamicSectionTitle: { size: 24, bold: true, font: FONT_FAMILY, color: "333333", italics: true, spacing: { before: 300, after: 100 } },
        fieldLabel: { size: 22, bold: true, font: FONT_FAMILY, color: COLOR_BLACK, spacing: { before: 300, after: 100 } },
        instruction: { size: 20, italics: true, color: "666666", font: FONT_FAMILY, spacing: { after: 100 } },
        placeholder: { size: 22, color: "555555", italics: true, font: FONT_FAMILY }
    };

    const docChildren = [];
    const conditionalMap = buildReverseConditionalMap();

    // Título Principal
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

        // Título do Capítulo
        const h2 = tabEl.querySelector('h2');
        if (h2) {
            let titleText = h2.innerText.trim().replace(/^\W+/, '');
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: titleText, ...styles.chapterTitle })],
                heading: HeadingLevel.HEADING_1,
                border: { bottom: { color: "BFBFBF", space: 4, value: "single", size: 6 } }
            }));
        }

        // Se o capítulo inteiro não se aplica ao simplificado
        if (isSimplificado && isChapterHiddenInSimplificado) {
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: "Nota: Este capítulo não se aplica ao ETP Simplificado.", ...styles.instruction })],
                spacing: { after: 400 }
            }));
            return; // Pula para o próximo capítulo
        }

        // Blocos Dinâmicos (apenas se o capítulo estiver ativo)
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
                spacing: { before: 300 }
            }));
            
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = dynData.config.template(1, false);
            tempDiv.querySelectorAll('label').forEach(lbl => {
                docChildren.push(new Paragraph({ children: [new TextRun({ text: lbl.textContent.trim().replace(/\d+\.1\./, 'X.'), ...styles.fieldLabel })] }));
                docChildren.push(new Paragraph({ children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })], spacing: { after: 200 } }));
            });
        }

        // Campos Estáticos
        tabEl.querySelectorAll('.form-group').forEach(group => {
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

            docChildren.push(new Paragraph({
                children: [new TextRun({ text: label.textContent.trim(), ...styles.fieldLabel })],
                keepNext: true
            }));

            // Nota condicional
            if (group.classList.contains('conditional-field')) {
                const triggerName = conditionalMap[group.id] || "resposta anterior";
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: `Nota: Preencher somente se necessário conforme a opção "${triggerName}".`, ...styles.instruction })]
                }));
            }

            // Opções (Radio/Checkbox) ou Placeholder
            const options = group.querySelectorAll('.checkbox-group label');
            if (options.length > 0) {
                options.forEach(opt => {
                    docChildren.push(new Paragraph({
                        children: [
                            new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }),
                            new TextRun({ text: opt.textContent.trim(), size: 22, font: FONT_FAMILY })
                        ],
                        indent: { left: 400 }
                    }));
                });
            } else {
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })],
                    spacing: { after: 200 }
                }));
            }
        });
    });

    const filename = isSimplificado ? "Modelo_ETP_Simplificado.docx" : "Modelo_ETP_Completo.docx";

    Packer.toBlob(new Document({
        sections: [{
            properties: { page: { margin: { top: 1440, bottom: 1440, left: 1134, right: 1134 } } },
            children: docChildren,
        }],
    })).then(blob => saveAs(blob, filename));
}