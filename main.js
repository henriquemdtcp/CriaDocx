import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
         AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel } from 'docx';

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

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

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

// CRIAR O DOCUMENTO (versão resumida - você pode adicionar mais seções depois)
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
      boxAtencao("A peça técnica vale metade da nota discursiva. Priorize tempo e atenção nela!", "⚠️", "FFF4E6", CORES.PRAZO),
      
      // ADICIONE MAIS SEÇÕES AQUI conforme o código completo anterior...
      
      linhaSeparacao(CORES.TITULO),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200 },
        children: [new TextRun({ text: "Material elaborado para concurso Câmara dos Deputados", size: 18, color: "999999", italics: true })]
      })
    ]
  }]
});

// FUNÇÃO PARA GERAR O DOCUMENTO
window.gerarDocumento = async function() {
    const btn = document.getElementById('btnGerar');
    const status = document.getElementById('status');
    
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span> Gerando documento...';
    status.style.display = 'flex';
    status.className = 'status processing';
    status.innerHTML = '<span class="spinner"></span> Processando... Isso pode levar alguns segundos';
    
    try {
        const buffer = await Packer.toBuffer(doc);
        
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
        });
        
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'Camara_Deputados_Prova_Discursiva_Guia_Completo.docx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        status.className = 'status success';
        status.textContent = '✅ Documento gerado com sucesso! O download deve iniciar automaticamente.';
        btn.textContent = 'Gerar Novamente';
    } catch (error) {
        status.className = 'status error';
        status.textContent = '❌ Erro ao gerar documento: ' + error.message;
        btn.textContent = 'Tentar Novamente';
        console.error('Erro completo:', error);
    } finally {
        btn.disabled = false;
    }
}