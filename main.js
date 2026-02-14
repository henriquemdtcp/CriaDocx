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
        if (!doc) {
          throw new Error('Documento não foi criado corretamente');
        }
        addDebugLog('✅ Objeto Document válido', 'success');
        
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

