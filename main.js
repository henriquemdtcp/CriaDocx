import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
         AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel } from 'docx';

// Adicionar log inicial
console.log('🟢 Script main.js carregado com sucesso!');

const CORES = {
  TITULO: "203864",
  DESTAQUE: "E7E6E6",
  ESTRUTURA: "4472C4",
  CONTEUDO: "70AD47",
  PROCEDIMENTO: "FFC000",
  PRAZO: "C0504D",
  COMPETENCIA: "9B59B6",
  EXEMPLO: "E67E22"
};

console.log('🟢 Constantes de cores definidas');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

console.log('🟢 Bordas configuradas');

function cellTitulo(texto, cor) {
  return new TableCell({
    borders,
    width: { size: 9360, type: WidthType.DXA },
    shading: { fill: cor, type: ShadingType.CLEAR },
    margins: { top: 120, bottom: 120, left: 180, right: 180 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: texto, bold: true, size: 28, color: "FFFFFF" })]
      })
    ]
  });
}

function cellConteudo(conteudo) {
  return new TableCell({
    borders,
    width: { size: 9360, type: WidthType.DXA },
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    children: Array.isArray(conteudo) ? conteudo : [conteudo]
  });
}

function cellDupla(label, valor, corLabel) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 3120, type: WidthType.DXA },
        shading: { fill: corLabel, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: label, bold: true, size: 22, color: "FFFFFF" })]
          })
        ]
      }),
      new TableCell({
        borders,
        width: { size: 6240, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: Array.isArray(valor) ? valor : [valor]
      })
    ]
  });
}

function subtitulo(texto, icone = "▸") {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [
      new TextRun({ text: icone + " ", size: 26, bold: true, color: CORES.TITULO }),
      new TextRun({ text: texto, size: 26, bold: true, color: CORES.TITULO })
    ]
  });
}

function itemLista(texto, cor) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [
      new TextRun({ text: "● ", size: 24, bold: true, color: cor }),
      new TextRun({ text: texto, size: 22 })
    ]
  });
}

function destaque(label, valor, cor) {
  return new Paragraph({
    spacing: { before: 100, after: 100 },
    children: [
      new TextRun({ text: label + ": ", size: 22, bold: true, color: cor }),
      new TextRun({ text: valor, size: 22 })
    ]
  });
}

function boxAtencao(texto, icone = "⚠️", cor = "FFF4E6", corTexto = CORES.PRAZO) {
  return new Paragraph({
    spacing: { before: 200, after: 200, line: 340 },
    shading: { fill: cor, type: ShadingType.CLEAR },
    margins: { top: 150, bottom: 150, left: 200, right: 200 },
    children: [
      new TextRun({ text: icone + " ", size: 22, bold: true, color: corTexto }),
      new TextRun({ text: texto, size: 22, bold: true, color: corTexto })
    ]
  });
}

function boxDica(texto) {
  return new Paragraph({
    spacing: { before: 200, after: 200, line: 340 },
    shading: { fill: "E8F8F5", type: ShadingType.CLEAR },
    margins: { top: 150, bottom: 150, left: 200, right: 200 },
    children: [
      new TextRun({ text: "✅ ", size: 22, bold: true, color: CORES.CONTEUDO }),
      new TextRun({ text: texto, size: 22 })
    ]
  });
}

function boxErro(texto) {
  return new Paragraph({
    spacing: { before: 200, after: 200, line: 340 },
    shading: { fill: "FDEDEC", type: ShadingType.CLEAR },
    margins: { top: 150, bottom: 150, left: 200, right: 200 },
    children: [
      new TextRun({ text: "❌ ", size: 22, bold: true, color: CORES.PRAZO }),
      new TextRun({ text: texto, size: 22 })
    ]
  });
}

function linhaSeparacao(cor = CORES.TITULO) {
  return new Paragraph({
    spacing: { before: 300, after: 300 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: cor } }
  });
}

function espaco(tamanho = 200) {
  return new Paragraph({ text: "", spacing: { before: tamanho } });
}

function paragrafo(texto) {
  return new Paragraph({
    spacing: { before: 80, after: 200, line: 340 },
    children: [new TextRun({ text: texto, size: 22 })]
  });
}

console.log('🟢 Todas as funções auxiliares definidas');

// CRIAR O DOCUMENTO
console.log('🔄 Iniciando criação do documento...');

const doc = new Document({
  styles: {
    default: { 
      document: { 
        run: { font: "Arial", size: 24 },
        paragraph: { spacing: { line: 360 } }
      } 
    },
    paragraphStyles: [
      {
        id: "Heading1",
        run: { size: 36, bold: true, font: "Arial", color: CORES.TITULO },
        paragraph: { 
          spacing: { before: 360, after: 240 }, 
          outlineLevel: 0,
          alignment: AlignmentType.CENTER
        }
      },
      {
        id: "Heading2",
        run: { size: 30, bold: true, font: "Arial", color: CORES.TITULO },
        paragraph: { 
          spacing: { before: 300, after: 200 }, 
          outlineLevel: 1,
          shading: { fill: CORES.DESTAQUE, type: ShadingType.CLEAR }
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      new Paragraph({ text: "PROVA DISCURSIVA - CÂMARA DOS DEPUTADOS", heading: HeadingLevel.HEADING_1 }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 200 },
        children: [new TextRun({ text: "Guia Completo: Parecer Administrativo e Questões Discursivas", size: 22, italics: true, color: "666666" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
        children: [new TextRun({ text: "Analista Legislativo - Área: Técnica", size: 18, color: "999999" })]
      }),

      new Paragraph({ text: "1. PANORAMA GERAL DA PROVA DISCURSIVA", heading: HeadingLevel.HEADING_2 }),
      espaco(200),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        rows: [
          new TableRow({ children: [cellTitulo("ESTRUTURA E PONTUAÇÃO", CORES.ESTRUTURA)] }),
          cellDupla("Peso Total", new Paragraph({ children: [new TextRun({ text: "60 pontos (25% da nota final)", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.ESTRUTURA),
          cellDupla("Duração", new Paragraph({ children: [new TextRun({ text: "3 horas (turno da tarde)", size: 22, bold: true })] }), CORES.ESTRUTURA),
          cellDupla("Nota Mínima", new Paragraph({ children: [new TextRun({ text: "30 pontos no conjunto", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.ESTRUTURA)
        ]
      }),

      espaco(200),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        rows: [
          new TableRow({ children: [cellTitulo("COMPOSIÇÃO DA PROVA", CORES.CONTEUDO)] }),
          cellDupla("Peça Técnica", new Paragraph({ children: [new TextRun({ text: "Até 50 linhas → ", size: 22 }), new TextRun({ text: "30 pontos", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.CONTEUDO),
          cellDupla("Questão 1", new Paragraph({ children: [new TextRun({ text: "Até 20 linhas → ", size: 22 }), new TextRun({ text: "15 pontos", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.CONTEUDO),
          cellDupla("Questão 2", new Paragraph({ children: [new TextRun({ text: "Até 20 linhas → ", size: 22 }), new TextRun({ text: "15 pontos", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.CONTEUDO)
        ]
      }),

      espaco(200),
      subtitulo("Gestão Estratégica do Tempo", "📐"),
      espaco(120),
      itemLista("Peça técnica: 1h40 a 1h50 (concentra 50% da nota discursiva)", CORES.PROCEDIMENTO),
      itemLista("Questão 1 (20 linhas): 30 a 35 minutos", CORES.PROCEDIMENTO),
      itemLista("Questão 2 (20 linhas): 30 a 35 minutos", CORES.PROCEDIMENTO),
      itemLista("Revisão final: 5 a 10 minutos", CORES.PROCEDIMENTO),
      espaco(200),
      boxAtencao("A peça técnica vale metade da nota discursiva. Priorize tempo e atenção nela!", "⚠️", "FFF4E6", CORES.PRAZO),
      
      linhaSeparacao(CORES.TITULO),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200 },
        children: [new TextRun({ text: "Material elaborado para concurso Câmara dos Deputados", size: 18, color: "999999", italics: true })]
      })
    ]
  }]
});

console.log('🟢 Documento criado com sucesso!');

// ADICIONAR CONSOLE DE DEBUG NA TELA
function addDebugLog(message, type = 'info') {
  const debugContainer = document.getElementById('debugLogs');
  if (!debugContainer) {
    const container = document.createElement('div');
    container.id = 'debugLogs';
    container.style.cssText = `
      position: fixed;
      bottom: 20px;
      right: 20px;
      background: rgba(0,0,0,0.9);
      color: #00ff00;
      padding: 15px;
      border-radius: 10px;
      max-width: 400px;
      max-height: 300px;
      overflow-y: auto;
      font-family: monospace;
      font-size: 12px;
      z-index: 9999;
      box-shadow: 0 4px 20px rgba(0,0,0,0.5);
    `;
    document.body.appendChild(container);
  }
  
  const log = document.createElement('div');
  const timestamp = new Date().toLocaleTimeString();
  
  const colors = {
    'info': '#00ff00',
    'success': '#00ffff',
    'error': '#ff0000',
    'warning': '#ffff00'
  };
  
  log.style.color = colors[type] || colors.info;
  log.style.marginBottom = '5px';
  log.innerHTML = `[${timestamp}] ${message}`;
  
  document.getElementById('debugLogs').appendChild(log);
  document.getElementById('debugLogs').scrollTop = document.getElementById('debugLogs').scrollHeight;
  
  console.log(`[${type.toUpperCase()}] ${message}`);
}

// FUNÇÃO PARA GERAR O DOCUMENTO COM DEBUG (VERSÃO CORRIGIDA PARA BROWSER)
window.gerarDocumento = async function() {
    addDebugLog('🔵 Função gerarDocumento() chamada', 'info');
   
    const btn = document.getElementById('btnGerar');
    const status = document.getElementById('status');
   
    if (!btn) {
      addDebugLog('❌ ERRO: Botão não encontrado!', 'error');
      return;
    }
   
    if (!status) {
      addDebugLog('❌ ERRO: Elemento status não encontrado!', 'error');
      return;
    }
   
    addDebugLog('✅ Elementos HTML encontrados', 'success');
   
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span> Gerando documento...';
    status.style.display = 'flex';
    status.className = 'status processing';
    status.innerHTML = '<span class="spinner"></span> Processando... Isso pode levar alguns segundos';
   
    addDebugLog('🔄 Interface atualizada - processamento iniciado', 'info');
   
    try {
        addDebugLog('📦 Verificando objeto Document...', 'info');
        
        // Definição expandida do documento conforme solicitado
        const doc = new Document({
            styles: {
                default: {
                    document: {
                        run: { font: "Arial", size: 24 },
                        paragraph: { spacing: { line: 360 } }
                    }
                },
                paragraphStyles: [
                    {
                        id: "Heading1",
                        run: { size: 36, bold: true, font: "Arial", color: CORES.TITULO },
                        paragraph: {
                            spacing: { before: 360, after: 240 },
                            outlineLevel: 0,
                            alignment: AlignmentType.CENTER
                        }
                    },
                    {
                        id: "Heading2",
                        run: { size: 30, bold: true, font: "Arial", color: CORES.TITULO },
                        paragraph: {
                            spacing: { before: 300, after: 200 },
                            outlineLevel: 1,
                            shading: { fill: CORES.DESTAQUE, type: ShadingType.CLEAR }
                        }
                    }
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size: { width: 12240, height: 15840 },
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
                    }
                },
                children: [
                    new Paragraph({ text: "PROVA DISCURSIVA - CÂMARA DOS DEPUTADOS", heading: HeadingLevel.HEADING_1 }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 120, after: 200 },
                        children: [new TextRun({ text: "Guia Completo: Parecer Administrativo e Questões Discursivas", size: 22, italics: true, color: "666666" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 300 },
                        children: [new TextRun({ text: "Analista Legislativo - Área: Técnica", size: 18, color: "999999" })]
                    }),
                    new Paragraph({ text: "1. PANORAMA GERAL DA PROVA DISCURSIVA", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("ESTRUTURA E PONTUAÇÃO", CORES.ESTRUTURA)] }),
                            cellDupla("Peso Total", new Paragraph({ children: [new TextRun({ text: "60 pontos (25% da nota final)", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.ESTRUTURA),
                            cellDupla("Duração", new Paragraph({ children: [new TextRun({ text: "3 horas (turno da tarde)", size: 22, bold: true })] }), CORES.ESTRUTURA),
                            cellDupla("Nota Mínima", new Paragraph({ children: [new TextRun({ text: "30 pontos no conjunto das provas discursivas", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.ESTRUTURA)
                        ]
                    }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("COMPOSIÇÃO DA PROVA", CORES.CONTEUDO)] }),
                            cellDupla("Peça Técnica", new Paragraph({ children: [new TextRun({ text: "Até 50 linhas → ", size: 22 }), new TextRun({ text: "30 pontos", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.CONTEUDO),
                            cellDupla("Questão 1", new Paragraph({ children: [new TextRun({ text: "Até 20 linhas → ", size: 22 }), new TextRun({ text: "15 pontos", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.CONTEUDO),
                            cellDupla("Questão 2", new Paragraph({ children: [new TextRun({ text: "Até 20 linhas → ", size: 22 }), new TextRun({ text: "15 pontos", size: 22, bold: true, color: CORES.PRAZO })] }), CORES.CONTEUDO)
                        ]
                    }),
                    espaco(200),
                    subtitulo("Gestão Estratégica do Tempo", "📐"),
                    espaco(120),
                    itemLista("Peça técnica: 1h40 a 1h50 (concentra 50% da nota discursiva)", CORES.PROCEDIMENTO),
                    itemLista("Questão 1 (20 linhas): 30 a 35 minutos", CORES.PROCEDIMENTO),
                    itemLista("Questão 2 (20 linhas): 30 a 35 minutos", CORES.PROCEDIMENTO),
                    itemLista("Revisão final: 5 a 10 minutos", CORES.PROCEDIMENTO),
                    espaco(200),
                    boxAtencao("A peça técnica vale metade da nota discursiva. Priorize tempo e atenção nela!", "⚠️", "FFF4E6", CORES.PRAZO),
                    espaco(300),
                    new Paragraph({ text: "2. NATUREZA DA PEÇA TÉCNICA", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("TIPOS DE PEÇAS POSSÍVEIS", CORES.EXEMPLO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    new Paragraph({ spacing: { before: 120, after: 100 }, children: [new TextRun({ text: "Probabilidade de cobrança:", size: 22, bold: true, color: CORES.EXEMPLO })] }),
                                    itemLista("Parecer Administrativo/Técnico: ~90%", CORES.EXEMPLO),
                                    itemLista("Nota Técnica: ~8%", CORES.EXEMPLO),
                                    itemLista("Informação Técnica: ~2%", CORES.EXEMPLO),
                                    itemLista("Despacho Técnico: raro", CORES.EXEMPLO)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    boxDica("Prepare-se com foco no Parecer Administrativo. Ele é aceito como Nota Técnica sem penalização e demonstra domínio institucional completo."),
                    espaco(300),
                    new Paragraph({ text: "3. ESTRUTURA DO PARECER ADMINISTRATIVO", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    subtitulo("3.1 Cabeçalho / Identificação", "📝"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("CABEÇALHO", CORES.ESTRUTURA)] }),
                            cellDupla("Formato", new Paragraph({ children: [new TextRun({ text: "Parecer nº X/2026 – [Sigla da Unidade Técnica]", size: 22, bold: true })] }), CORES.ESTRUTURA),
                            cellDupla("Exemplo", new Paragraph({ children: [new TextRun({ text: "Parecer nº X/2026 – CONLE", size: 22 })] }), CORES.ESTRUTURA)
                        ]
                    }),
                    espaco(200),
                    boxDica("Nunca invente número. Use sempre 'X' quando o enunciado não fornecer. A sigla da unidade pode ser fictícia padrão (ex: CONLE, DAL, etc)."),
                    espaco(300),
                    subtitulo("3.2 Processo", "📝"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("IDENTIFICAÇÃO DO PROCESSO", CORES.ESTRUTURA)] }),
                            cellDupla("Formato", new Paragraph({ children: [new TextRun({ text: "Processo nº X", size: 22, bold: true })] }), CORES.ESTRUTURA)
                        ]
                    }),
                    espaco(200),
                    boxErro("Nunca invente número de processo. O 'X' é 100% aceitável e recomendado quando não fornecido no enunciado."),
                    espaco(300),
                    subtitulo("3.3 Ementa", "📝"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("EMENTA - REGRAS ESSENCIAIS", CORES.PRAZO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    destaque("Formato", "CAIXA ALTA", CORES.PRAZO),
                                    destaque("Estrutura", "Frases nominais (sem verbos no início)", CORES.PRAZO),
                                    destaque("Conteúdo", "Palavras-chave do tema em ordem lógica", CORES.PRAZO),
                                    destaque("Pontuação", "Pontos finais separando tópicos", CORES.PRAZO)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("EXEMPLO CORRETO", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120 },
                                        children: [new TextRun({ text: "EMENTA: PROCESSO LEGISLATIVO ORÇAMENTÁRIO. CRÉDITO ADICIONAL ESPECIAL. CALAMIDADE PÚBLICA. REGIME DE TRAMITAÇÃO. URGÊNCIA E PRIORIDADE. COMPETÊNCIA DA MESA DIRETORA E DA PRESIDÊNCIA DA CÂMARA DOS DEPUTADOS.", size: 20, bold: true })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(200),
                    boxErro("ERRO GRAVÍSSIMO: 'EMENTA: Trata-se de...' → Ementa não admite verbos no início!"),
                    espaco(120),
                    boxDica("Liste os temas na ordem que serão abordados no parecer. Isso ajuda o corretor a mapear sua resposta."),
                    espaco(300),
                    subtitulo("3.4 Relatório (ou Histórico)", "📝"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("RELATÓRIO", CORES.ESTRUTURA)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    paragrafo("Seção onde você descreve objetivamente o caso apresentado no enunciado, sem emitir opinião ou análise."),
                                    destaque("Extensão máxima", "6 a 8 linhas", CORES.PRAZO),
                                    destaque("Fechamento padrão", "É o relatório. / É o relatório. Passo a opinar.", CORES.ESTRUTURA)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("EXEMPLO DE RELATÓRIO", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120, line: 340 },
                                        children: [new TextRun({ text: "Trata-se de Projeto de Lei do Congresso Nacional que visa à abertura de crédito adicional especial destinado ao atendimento de despesas urgentes decorrentes de situação de calamidade pública reconhecida pelo Congresso Nacional. No curso da tramitação, foi apresentado requerimento para adoção do regime de urgência, com o objetivo de acelerar a deliberação da matéria, suscitando questionamentos quanto à adequação do regime proposto e às competências institucionais envolvidas na condução do processo legislativo.\n\nÉ o relatório. Passo a opinar.", size: 21 })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(200),
                    boxAtencao("A banca NÃO pontua narrativa. Seja objetivo! O Relatório serve apenas para contextualizar.", "⚠️", "FFF4E6", CORES.PRAZO),
                    espaco(300),
                    subtitulo("3.5 Parecer / Fundamentação (O CORAÇÃO DA PEÇA)", "⚖️"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("FUNDAMENTAÇÃO - REGRAS ESTRATÉGICAS", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    paragrafo("Esta é a seção que concentra a pontuação. Aqui você responde TODOS os quesitos do comando."),
                                    new Paragraph({ spacing: { before: 120, after: 80 }, children: [new TextRun({ text: "Estrutura recomendada por quesito:", size: 22, bold: true, color: CORES.CONTEUDO })] }),
                                    itemLista("Use conectivos que espelhem o comando: 'Quanto à...', 'No que se refere a...', 'Sob o aspecto de...'", CORES.CONTEUDO),
                                    itemLista("Um parágrafo por quesito (facilita o mapeamento pelo corretor)", CORES.CONTEUDO),
                                    itemLista("Linguagem técnica, objetiva e formal", CORES.CONTEUDO),
                                    itemLista("Fundamente com normas (CF, leis, regimento interno)", CORES.CONTEUDO)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("CONECTIVOS ESTRATÉGICOS POR QUESITO", CORES.PROCEDIMENTO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    itemLista("Quanto à competência...", CORES.PROCEDIMENTO),
                                    itemLista("No que se refere ao procedimento...", CORES.PROCEDIMENTO),
                                    itemLista("Sob o aspecto da legalidade...", CORES.PROCEDIMENTO),
                                    itemLista("No âmbito da gestão administrativa...", CORES.PROCEDIMENTO),
                                    itemLista("Quanto à natureza jurídica...", CORES.PROCEDIMENTO),
                                    itemLista("No tocante à tramitação...", CORES.PROCEDIMENTO)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    boxDica("Esses conectivos ajudam o corretor a identificar exatamente onde você respondeu cada quesito. Use-os estrategicamente!"),
                    espaco(300),
                    subtitulo("3.6 Conclusão / Encaminhamento", "✅"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("CONCLUSÃO", CORES.ESTRUTURA)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    paragrafo("Seção final onde você apresenta sua opinião técnica fundamentada e encaminha à autoridade superior."),
                                    new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun({ text: "Fórmula padrão recomendada:", size: 22, bold: true, color: CORES.ESTRUTURA })] }),
                                    new Paragraph({
                                        spacing: { before: 100, after: 100, line: 340 },
                                        shading: { fill: "E8F8F5", type: ShadingType.CLEAR },
                                        margins: { top: 120, bottom: 120, left: 150, right: 150 },
                                        children: [new TextRun({ text: "Ante o exposto, opina-se [favoravelmente/contrariamente] à medida, nos termos acima delineados.\n\nEncaminha-se à consideração superior.", size: 21, italics: true })]
                                    })
                                ])]
                            })
                        ]
                    }),
                    espaco(300),
                    subtitulo("3.7 Local, Data e Assinatura", "📝"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("FECHAMENTO", CORES.ESTRUTURA)] }),
                            cellDupla("Formato", new Paragraph({ children: [new TextRun({ text: "Brasília, [dia] de [mês] de 2026.", size: 22, bold: true }), new TextRun({ text: "\n\nAnalista Legislativo", size: 22, bold: true })] }), CORES.ESTRUTURA),
                            cellDupla("Exemplo", new Paragraph({ children: [new TextRun({ text: "Brasília, 08 de março de 2026.", size: 22 }), new TextRun({ text: "\n\nAnalista Legislativo", size: 22 })] }), CORES.ESTRUTURA)
                        ]
                    }),
                    espaco(200),
                    boxErro("NUNCA coloque: nome real, assinatura criativa ou matrícula fictícia. Use apenas 'Analista Legislativo'."),
                    espaco(300),
                    new Paragraph({ text: "4. QUESTÕES DISCURSIVAS (20 LINHAS)", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("ESTRATÉGIA PARA QUESTÕES DE 20 LINHAS", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    boxErro("❌ NÃO faça: introdução, conclusão ou 'enrolação'"),
                                    espaco(120),
                                    boxDica("✅ FAÇA: vá direto ao ponto, um parágrafo por tópico, linguagem técnica direta"),
                                    espaco(160),
                                    new Paragraph({ spacing: { before: 120, after: 80 }, children: [new TextRun({ text: "Estrutura ideal:", size: 22, bold: true, color: CORES.CONTEUDO })] }),
                                    itemLista("Identifique quantos quesitos há na questão", CORES.CONTEUDO),
                                    itemLista("Um parágrafo para cada quesito", CORES.CONTEUDO),
                                    itemLista("Use definição + consequência no mesmo parágrafo", CORES.CONTEUDO),
                                    itemLista("Fundamente com normas quando possível", CORES.CONTEUDO),
                                    itemLista("Escreva entre 15-20 linhas (aproveite o espaço!)", CORES.CONTEUDO)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("EXEMPLO DE BOA RESPOSTA", CORES.EXEMPLO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120, line: 340 },
                                        children: [new TextRun({ text: "Questão hipotética: Diferencie regime de urgência e regime de prioridade.\n\nResposta modelo:\n\nO regime de urgência consiste em procedimento especial de tramitação que implica redução de prazos regimentais e preferência absoluta na pauta de deliberações, sendo cabível apenas nas hipóteses expressamente previstas na Constituição Federal e no Regimento Interno da Câmara dos Deputados. Tal regime tem como consequência a supressão de determinadas etapas procedimentais e a inclusão automática da matéria na Ordem do Dia, exigindo deliberação em prazo determinado.\n\nJá o regime de prioridade confere precedência na apreciação da proposição, sem, contudo, alterar os prazos regimentais ou suprimir etapas do processo legislativo. A matéria em regime de prioridade será apreciada antes das demais que não gozem de regime especial, mas preserva-se a integralidade do rito procedimental aplicável. Assim, enquanto a urgência excepcionalmente altera prazos e procedimentos, a prioridade apenas reordena a sequência de apreciação das matérias.", size: 20, italics: true })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(200),
                    boxDica("Note como a resposta: (1) define cada conceito, (2) aponta consequências práticas, (3) diferencia claramente os institutos, (4) usa linguagem técnica precisa."),
                    espaco(300),
                    new Paragraph({ text: "5. CRITÉRIOS DE AVALIAÇÃO DO CEBRASPE", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("SISTEMA DE PONTUAÇÃO", CORES.PRAZO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    paragrafo("O Cebraspe avalia duas dimensões separadamente:"),
                                    espaco(120),
                                    destaque("NC (Nota de Conteúdo)", "Domínio do tema, correção técnica, completude da resposta", CORES.CONTEUDO),
                                    destaque("NL (Nota de Linguagem)", "Correção gramatical, clareza, coesão, adequação ao registro formal", CORES.ESTRUTURA),
                                    espaco(160),
                                    boxAtencao("REGRA CRÍTICA: Quanto mais linhas você escrever, menor o peso de cada erro. Por isso, use o espaço disponível!", "📐", "FFF4E6", CORES.PRAZO)
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("APROVEITAMENTO ESTRATÉGICO DE LINHAS", CORES.PROCEDIMENTO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    itemLista("Peça técnica: escrever 40-48 linhas das 50 disponíveis", CORES.PROCEDIMENTO),
                                    itemLista("Questões de 20 linhas: usar 15-20 linhas, se houver conteúdo", CORES.PROCEDIMENTO),
                                    itemLista("Nunca deixe questão em branco ou com menos de 10 linhas", CORES.PROCEDIMENTO),
                                    espaco(160),
                                    boxDica("Mais linhas = mais diluição de erros gramaticais. Use todo o espaço com conteúdo relevante!")
                                ])]
                            })
                        ]
                    }),
                    espaco(300),
                    new Paragraph({ text: "6. ESTRATÉGIAS DE EXECUÇÃO NA PROVA", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    subtitulo("6.1 Ordem Recomendada de Execução", "📋"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("SEQUÊNCIA ESTRATÉGICA", CORES.PROCEDIMENTO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    new Paragraph({ spacing: { before: 120, after: 100 }, children: [new TextRun({ text: "1º PASSO: ", size: 22, bold: true, color: CORES.PROCEDIMENTO }), new TextRun({ text: "Ler todos os enunciados (5 min)", size: 22 })] }),
                                    new Paragraph({ spacing: { before: 80, after: 100 }, children: [new TextRun({ text: "2º PASSO: ", size: 22, bold: true, color: CORES.PROCEDIMENTO }), new TextRun({ text: "Fazer rascunho da PEÇA TÉCNICA por palavras-chave (10-15 min)", size: 22 })] }),
                                    new Paragraph({ spacing: { before: 80, after: 100 }, children: [new TextRun({ text: "3º PASSO: ", size: 22, bold: true, color: CORES.PROCEDIMENTO }), new TextRun({ text: "Passar a PEÇA TÉCNICA a limpo (1h20-1h30)", size: 22 })] }),
                                    new Paragraph({ spacing: { before: 80, after: 100 }, children: [new TextRun({ text: "4º PASSO: ", size: 22, bold: true, color: CORES.PROCEDIMENTO }), new TextRun({ text: "Responder QUESTÃO 1 direto na folha definitiva (30-35 min)", size: 22 })] }),
                                    new Paragraph({ spacing: { before: 80, after: 100 }, children: [new TextRun({ text: "5º PASSO: ", size: 22, bold: true, color: CORES.PROCEDIMENTO }), new TextRun({ text: "Responder QUESTÃO 2 direto na folha definitiva (30-35 min)", size: 22 })] }),
                                    new Paragraph({ spacing: { before: 80, after: 100 }, children: [new TextRun({ text: "6º PASSO: ", size: 22, bold: true, color: CORES.PROCEDIMENTO }), new TextRun({ text: "Revisão final pontual (5-10 min)", size: 22 })] })
                                ])]
                            })
                        ]
                    }),
                    espaco(200),
                    boxAtencao("NÃO faça rascunho completo das questões de 20 linhas. Vá direto para a folha definitiva com um mental map dos tópicos.", "⚠️", "FFF4E6", CORES.PRAZO),
                    espaco(300),
                    subtitulo("6.2 Técnica do Rascunho por Palavras-Chave", "💡"),
                    espaco(160),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("MÉTODO DE RASCUNHO EFICIENTE", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    paragrafo("Para a PEÇA TÉCNICA, faça rascunho apenas com palavras-chave de cada quesito:"),
                                    espaco(120),
                                    new Paragraph({
                                        spacing: { before: 100, after: 100, line: 340 },
                                        shading: { fill: "FEF5E7", type: ShadingType.CLEAR },
                                        margins: { top: 120, bottom: 120, left: 150, right: 150 },
                                        children: [new TextRun({ text: "Exemplo de rascunho:\n\na) Natureza crédito especial → sem dotação específica → autorização legislativa + indicação recursos → calamidade = tratamento diferenciado\n\nb) Procedimento → iniciativa Executivo → análise comissões orçamentárias → Congresso Nacional\n\nc) Urgência vs Prioridade → urgência = redução prazos + preferência absoluta → prioridade = precedência sem supressão etapas\n\nd) Competências → Mesa = delibera aspectos formais + atribuições administrativas → Presidência = dirige trabalhos + define pauta + zela regimento\n\ne) Conclusão → tramitação regular + avaliar adequação urgência + preferir prioridade se mais compatível", size: 20, italics: true })]
                                    }),
                                    espaco(160),
                                    boxDica("Com esse rascunho de palavras-chave, você economiza tempo e já organiza mentalmente a estrutura da resposta.")
                                ])]
                            })
                        ]
                    }),
                    espaco(300),
                    new Paragraph({ text: "7. TEMAS QUENTES PARA A PROVA", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        columnWidths: [3120, 3120, 3120],
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: CORES.COMPETENCIA, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "TEMA", bold: true, size: 22, color: "FFFFFF" })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: CORES.COMPETENCIA, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PROBABILIDADE", bold: true, size: 22, color: "FFFFFF" })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, shading: { fill: CORES.COMPETENCIA, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PONTOS-CHAVE", bold: true, size: 22, color: "FFFFFF" })] })] })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Processo Legislativo Orçamentário", size: 21, bold: true })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "MUITO ALTA", size: 21, bold: true, color: CORES.PRAZO })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Créditos adicionais, iniciativa, análise, aprovação", size: 20 })] })] })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Código de Ética e Decoro", size: 21, bold: true })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "MUITO ALTA", size: 21, bold: true, color: CORES.PRAZO })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Tramitação de representação, competências, procedimento", size: 20 })] })] })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Regimes de Tramitação", size: 21, bold: true })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ALTA", size: 21, bold: true, color: CORES.EXEMPLO })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Urgência × Prioridade, requisitos, efeitos", size: 20 })] })] })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Governança e Gestão de Riscos", size: 21, bold: true })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "MÉDIA-ALTA", size: 21, bold: true, color: CORES.PROCEDIMENTO })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "SWOT, BSC, aplicação institucional", size: 20 })] })] })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Competências Mesa/Presidência", size: 21, bold: true })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ALTA", size: 21, bold: true, color: CORES.EXEMPLO })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Atribuições, limites, pauta, tramitação", size: 20 })] })] })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Comissões (permanentes/temporárias)", size: 21, bold: true })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "MÉDIA-ALTA", size: 21, bold: true, color: CORES.PROCEDIMENTO })] })] }),
                                    new TableCell({ borders, width: { size: 3120, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Criação, competência, tramitação", size: 20 })] })] })
                                ]
                            })
                        ]
                    }),
                    espaco(300),
                    new Paragraph({ text: "8. EXEMPLO COMPLETO 1 - PARECER ADMINISTRATIVO", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("ENUNCIADO SIMULADO 1", CORES.EXEMPLO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120, line: 340 },
                                        children: [new TextRun({ text: "No âmbito da Câmara dos Deputados, foi apresentado Projeto de Lei do Congresso Nacional (PLN) visando à abertura de crédito adicional especial, destinado ao atendimento de despesas urgentes decorrentes de calamidade pública reconhecida pelo Congresso Nacional.\n\nDurante a tramitação, parlamentares requereram a adoção do regime de urgência, com o objetivo de acelerar a deliberação da matéria. Questiona-se, contudo, a adequação desse regime ao caso concreto, bem como as competências institucionais envolvidas na condução do processo.\n\nNa condição de Analista Legislativo, elabore peça de natureza técnica, na forma de Parecer Administrativo, abordando, necessariamente, os seguintes aspectos:\n\na) a natureza jurídica dos créditos adicionais, com destaque para o crédito especial, bem como os requisitos constitucionais e legais para sua abertura;\n\nb) o procedimento legislativo aplicável aos projetos que tratam de matéria orçamentária, inclusive quanto à iniciativa e à deliberação pelo Congresso Nacional;\n\nc) a distinção entre regime de urgência e regime de prioridade, avaliando a adequação do pedido formulado;\n\nd) as competências da Mesa Diretora e da Presidência da Câmara dos Deputados na definição da pauta e na condução da tramitação da proposição;\n\ne) a conclusão técnica quanto à regularidade do procedimento adotado, com o devido encaminhamento à autoridade competente.", size: 20 })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(300),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("RESPOSTA MODELO", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120, line: 340 },
                                        children: [new TextRun({ text: "PARECER Nº X/2026 – CONLE\n\nProcesso nº X\n\nEMENTA: PROCESSO LEGISLATIVO ORÇAMENTÁRIO. CRÉDITO ADICIONAL ESPECIAL. CALAMIDADE PÚBLICA. REGIME DE TRAMITAÇÃO. URGÊNCIA E PRIORIDADE. COMPETÊNCIA DA MESA DIRETORA E DA PRESIDÊNCIA DA CÂMARA DOS DEPUTADOS.\n\nI – RELATÓRIO\n\nTrata-se de Projeto de Lei do Congresso Nacional que visa à abertura de crédito adicional especial destinado ao atendimento de despesas urgentes decorrentes de situação de calamidade pública reconhecida pelo Congresso Nacional. No curso da tramitação, foi apresentado requerimento para adoção do regime de urgência, com o objetivo de acelerar a deliberação da matéria, suscitando questionamentos quanto à adequação do regime proposto e às competências institucionais envolvidas na condução do processo legislativo.\n\nÉ o relatório. Passo a opinar.\n\nII – PARECER\n\nQuanto à natureza jurídica dos créditos adicionais, cumpre destacar que o crédito especial destina-se à realização de despesas para as quais não haja dotação orçamentária específica, dependendo, para sua abertura, de autorização legislativa prévia e da indicação dos recursos correspondentes, nos termos da Constituição Federal e da legislação orçamentária vigente. Em situações de calamidade pública reconhecida pelo Congresso Nacional, admite-se tratamento diferenciado quanto a determinados requisitos fiscais, sem afastar, contudo, a necessidade de observância do devido processo legislativo.\n\nNo que se refere ao procedimento legislativo aplicável, os projetos que tratam de créditos adicionais são de iniciativa do Poder Executivo e submetem-se à apreciação do Congresso Nacional, com análise pelas comissões competentes, em especial as de natureza orçamentária, observadas as normas regimentais e constitucionais pertinentes.\n\nSob o aspecto da tramitação, impõe-se distinguir o regime de urgência do regime de prioridade. O regime de urgência implica redução de prazos e preferência absoluta na pauta, sendo cabível apenas nas hipóteses expressamente previstas no ordenamento jurídico e no Regimento Interno. Já o regime de prioridade confere precedência na apreciação da matéria, sem a supressão integral das etapas procedimentais. Assim, a adoção do regime de urgência deve ser avaliada à luz da excepcionalidade do caso concreto e da compatibilidade com as normas regimentais, podendo o regime de prioridade revelar-se medida mais adequada.\n\nQuanto às competências institucionais, compete à Mesa Diretora deliberar sobre aspectos formais da tramitação das proposições, bem como exercer atribuições administrativas e regimentais. À Presidência da Câmara dos Deputados incumbe dirigir os trabalhos legislativos, definir a pauta de deliberações e zelar pela observância do Regimento Interno, inclusive quanto à admissibilidade e ao processamento dos regimes de tramitação requeridos.\n\nNo âmbito da legalidade e da regularidade procedimental, verifica-se que a tramitação do projeto deve observar rigorosamente as normas constitucionais e regimentais, cabendo à Presidência e à Mesa Diretora assegurar que eventual adoção de regime especial esteja devidamente fundamentada e em consonância com o ordenamento jurídico.\n\nIII – CONCLUSÃO\n\nAnte o exposto, opina-se pela regular tramitação do Projeto de Lei do Congresso Nacional destinado à abertura de crédito adicional especial, recomendando-se a avaliação criteriosa da adequação do regime de urgência, à luz das normas constitucionais e regimentais aplicáveis, sem prejuízo da adoção do regime de prioridade, se mais compatível com o caso concreto.\n\nEncaminha-se à consideração superior.\n\nBrasília, 08 de março de 2026.\n\nAnalista Legislativo", size: 19 })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(300),
                    new Paragraph({ text: "9. EXEMPLO COMPLETO 2 - PARECER ADMINISTRATIVO", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("ENUNCIADO SIMULADO 2", CORES.EXEMPLO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120, line: 340 },
                                        children: [new TextRun({ text: "Chegou à Mesa da Câmara dos Deputados representação formulada por partido político contra Deputado Federal, imputando-lhe suposta prática de ato incompatível com o decoro parlamentar, nos termos do Código de Ética e Decoro Parlamentar da Casa.\n\nParalelamente, no contexto do fortalecimento da governança institucional, a Administração da Câmara avalia a aplicação de instrumentos de gestão estratégica e de riscos, como a Matriz SWOT e o Balanced Scorecard (BSC), para aprimorar a atuação das comissões parlamentares, especialmente no tratamento de processos sensíveis e de elevado impacto institucional.\n\nDiante desse cenário, elabore peça de natureza técnica, na forma de Nota Técnica ou Parecer Administrativo, abordando, obrigatoriamente, os seguintes pontos:\n\na) a tramitação da representação por quebra de decoro parlamentar, indicando a competência dos órgãos envolvidos e as fases do procedimento;\n\nb) o papel das comissões permanentes e temporárias, com destaque para sua criação, competências e limites de atuação no caso concreto;\n\nc) a competência da Mesa Diretora quanto ao recebimento e ao encaminhamento da representação;\n\nd) a aplicabilidade de instrumentos de governança e gestão de riscos, como a Matriz SWOT e o BSC, no aprimoramento da atuação institucional das comissões;\n\ne) a conclusão técnica com recomendações administrativas voltadas ao fortalecimento da governança e da segurança decisória no âmbito da Câmara dos Deputados.", size: 20 })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(300),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("RESPOSTA MODELO", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo(
                                    new Paragraph({
                                        spacing: { before: 120, after: 120, line: 340 },
                                        children: [new TextRun({ text: "PARECER Nº X/2026 – DAL\n\nProcesso nº X\n\nEMENTA: PROCESSO LEGISLATIVO. CÓDIGO DE ÉTICA. QUEBRA DE DECORO PARLAMENTAR. TRAMITAÇÃO DE REPRESENTAÇÃO. COMISSÕES PARLAMENTARES. GOVERNANÇA. GESTÃO DE RISCOS. PROVIDÊNCIAS.\n\nI – RELATÓRIO\n\nTrata-se de representação apresentada contra Deputado Federal, imputando-lhe suposta prática de ato incompatível com o decoro parlamentar, nos termos do Código de Ética e Decoro Parlamentar da Câmara dos Deputados. A demanda foi encaminhada à Mesa Diretora, suscitando análise quanto ao procedimento aplicável, às competências institucionais envolvidas e às providências administrativas cabíveis.\n\nÉ o relatório. Passo a opinar.\n\nII – PARECER\n\nQuanto à competência, o recebimento inicial da representação cabe à Mesa Diretora, a quem incumbe o exame formal e o encaminhamento ao Conselho de Ética e Decoro Parlamentar, órgão responsável pela instrução e apreciação do mérito, observadas as normas regimentais pertinentes.\n\nNo que se refere ao procedimento, a representação deve observar as fases de admissibilidade, instrução, contraditório e ampla defesa, culminando com parecer conclusivo do órgão competente, a ser submetido ao Plenário.\n\nSob o aspecto da legalidade, as comissões permanentes e temporárias atuam nos limites de suas atribuições, sendo vedada a extrapolação de competência ou a supressão de etapas essenciais do processo.\n\nNo âmbito da gestão administrativa, a adoção de instrumentos de governança e gestão de riscos, como a Matriz SWOT e o Balanced Scorecard, contribui para o aprimoramento do controle institucional, da previsibilidade decisória e da mitigação de riscos reputacionais e operacionais.\n\nIII – CONCLUSÃO\n\nAnte o exposto, opina-se favoravelmente à regular tramitação da representação, com observância do procedimento legal e das boas práticas de governança.\n\nEncaminha-se à consideração superior.\n\nBrasília, 08 de março de 2026.\n\nAnalista Legislativo", size: 19 })]
                                    })
                                )]
                            })
                        ]
                    }),
                    espaco(300),
                    new Paragraph({ text: "10. CHECKLIST PRÉ-PROVA", heading: HeadingLevel.HEADING_2 }),
                    espaco(200),
                    new Table({
                        width: { size: 9360, type: WidthType.DXA },
                        rows: [
                            new TableRow({ children: [cellTitulo("CHECKLIST DE PREPARAÇÃO", CORES.CONTEUDO)] }),
                            new TableRow({
                                children: [cellConteudo([
                                    new Paragraph({ spacing: { before: 120, after: 100 }, children: [new TextRun({ text: "✓ Estrutura do Parecer", size: 22, bold: true, color: CORES.ESTRUTURA })] }),
                                    itemLista("Sei montar: Cabeçalho, Processo, Ementa, Relatório, Parecer, Conclusão, Fecho", CORES.ESTRUTURA),
                                    itemLista("Domino a fórmula da Ementa (CAIXA ALTA, frases nominais, sem verbos)", CORES.ESTRUTURA),
                                    itemLista("Sei usar conectivos estratégicos para espelhar quesitos", CORES.ESTRUTURA),
                                    espaco(160),
                                    new Paragraph({ spacing: { before: 120, after: 100 }, children: [new TextRun({ text: "✓ Questões de 20 Linhas", size: 22, bold: true, color: CORES.CONTEUDO })] }),
                                    itemLista("Sei que não preciso de introdução nem conclusão", CORES.CONTEUDO),
                                    itemLista("Vou direto ao ponto com um parágrafo por quesito", CORES.CONTEUDO),
                                    itemLista("Uso definição + consequência no mesmo parágrafo", CORES.CONTEUDO),
                                    espaco(160),
                                    new Paragraph({ spacing: { before: 120, after: 100 }, children: [new TextRun({ text: "✓ Gestão do Tempo", size: 22, bold: true, color: CORES.PROCEDIMENTO })] }),
                                    itemLista("Sei alocar 1h40-1h50 para a peça técnica", CORES.PROCEDIMENTO),
                                    itemLista("Faço rascunho apenas da peça, por palavras-chave", CORES.PROCEDIMENTO),
                                    itemLista("Vou direto na folha definitiva nas questões de 20 linhas", CORES.PROCEDIMENTO),
                                    espaco(160),
                                    new Paragraph({ spacing: { before: 120, after: 100 }, children: [new TextRun({ text: "✓ Conteúdo dos Temas Quentes", size: 22, bold: true, color: CORES.COMPETENCIA })] }),
                                    itemLista("Processo Legislativo Orçamentário (créditos adicionais)", CORES.COMPETENCIA),
                                    itemLista("Código de Ética e Decoro (tramitação de representação)", CORES.COMPETENCIA),
                                    itemLista("Regimes de Tramitação (urgência × prioridade)", CORES.COMPETENCIA),
                                    itemLista("Competências da Mesa Diretora e Presidência", CORES.COMPETENCIA),
                                    itemLista("Comissões (criação, competências, limites)", CORES.COMPETENCIA),
                                    itemLista("Governança e Gestão de Riscos (SWOT, BSC)", CORES.COMPETENCIA)
                                ])]
                            })
                        ]
                    }),
                    linhaSeparacao(CORES.TITULO),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200 },
                        children: [new TextRun({ text: "Material elaborado com base em análise de editais Cebraspe e padrões de correção", size: 18, color: "999999", italics: true })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 80 },
                        children: [new TextRun({ text: "Atualizado em: Fevereiro de 2026", size: 18, color: "999999", italics: true })]
                    })
                ]
            }]
        });

        addDebugLog('✅ Objeto Document criado com sucesso', 'success');
       
        // ====== MUDANÇA AQUI: Usar toBlob() ao invés de toBuffer() ======
        addDebugLog('🔄 Chamando Packer.toBlob()...', 'info');
        const blob = await Packer.toBlob(doc);
        addDebugLog(`✅ Blob gerado com sucesso! Tamanho: ${blob.size} bytes`, 'success');
       
        addDebugLog('🔄 Criando URL para download...', 'info');
        const url = window.URL.createObjectURL(blob);
        addDebugLog('✅ URL criada: ' + url.substring(0, 50) + '...', 'success');
       
        addDebugLog('🔄 Criando elemento <a> para download...', 'info');
        const link = document.createElement('a');
        link.href = url;
        link.download = 'Camara_Deputados_Prova_Discursiva_Guia_Completo.docx';
       
        addDebugLog('🔄 Adicionando link ao DOM...', 'info');
        document.body.appendChild(link);
       
        addDebugLog('🔄 Disparando click() no link...', 'info');
        link.click();
       
        addDebugLog('🔄 Removendo link do DOM...', 'info');
        document.body.removeChild(link);
       
        addDebugLog('🔄 Liberando URL...', 'info');
        window.URL.revokeObjectURL(url);
       
        status.className = 'status success';
        status.textContent = '✅ Documento gerado com sucesso! O download deve iniciar automaticamente.';
        btn.textContent = 'Gerar Novamente';
       
        addDebugLog('🎉 PROCESSO CONCLUÍDO COM SUCESSO!', 'success');
       
    } catch (error) {
        addDebugLog('❌ ERRO CAPTURADO: ' + error.message, 'error');
        addDebugLog('📋 Stack trace: ' + error.stack, 'error');
       
        status.className = 'status error';
        status.textContent = '❌ Erro ao gerar documento: ' + error.message;
        btn.textContent = 'Tentar Novamente';
       
        console.error('Erro completo:', error);
    } finally {
        btn.disabled = false;
        addDebugLog('🔵 Finally: botão reativado', 'info');
    }
}

console.log('🟢 Função gerarDocumento() atribuída ao window');

window.addEventListener('DOMContentLoaded', () => {
  console.log('🟢 DOM carregado');
  addDebugLog('✅ Página carregada completamente', 'success');
  addDebugLog('✅ Script inicializado com sucesso', 'success');
  addDebugLog('ℹ️ Clique no botão para gerar o documento', 'info');
});


