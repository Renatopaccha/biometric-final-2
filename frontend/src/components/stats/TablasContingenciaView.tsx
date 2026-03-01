import { ArrowLeft, ChevronDown, Loader2, Sparkles } from 'lucide-react';
import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { ActionToolbar } from './ActionToolbar';
import { useDataContext } from '../../context/DataContext';
import { getCrosstabStats } from '../../api/stats';
import { sendChatMessage } from '../../api/ai';
import type { ContingencyTableResponse, ContingencyCellData } from '../../types/stats';

// Estilos para Excel WYSIWYG
const excelStyles = {
  mainHeader: {
    fill: { fgColor: { rgb: "D1E4FC" } },
    font: { bold: true, color: { rgb: "1F2937" }, sz: 11 },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center", wrapText: true }
  },
  categoryHeader: {
    fill: { fgColor: { rgb: "E3F2FD" } },
    font: { bold: true, sz: 10, color: { rgb: "1F2937" } },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  metricLabel: {
    fill: { fgColor: { rgb: "F9FAFB" } },
    font: { bold: false, sz: 9, color: { rgb: "4B5563" } },
    border: {
      top: { style: "thin", color: { rgb: "E5E7EB" } },
      bottom: { style: "thin", color: { rgb: "E5E7EB" } },
      left: { style: "thin", color: { rgb: "E5E7EB" } },
      right: { style: "thin", color: { rgb: "E5E7EB" } }
    },
    alignment: { horizontal: "left", vertical: "center", indent: 1 }
  },
  categoryCell: {
    font: { bold: true, sz: 10, color: { rgb: "1F2937" } },
    border: {
      top: { style: "thin", color: { rgb: "E5E7EB" } },
      bottom: { style: "thin", color: { rgb: "E5E7EB" } },
      left: { style: "thin", color: { rgb: "E5E7EB" } },
      right: { style: "thin", color: { rgb: "E5E7EB" } }
    },
    alignment: { horizontal: "left", vertical: "center" }
  },
  cellNumber: {
    font: { sz: 10 },
    border: {
      top: { style: "thin", color: { rgb: "E5E7EB" } },
      bottom: { style: "thin", color: { rgb: "E5E7EB" } },
      left: { style: "thin", color: { rgb: "E5E7EB" } },
      right: { style: "thin", color: { rgb: "E5E7EB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  totalRow: {
    fill: { fgColor: { rgb: "F1F5F9" } },
    font: { bold: true, sz: 10 },
    border: {
      top: { style: "medium", color: { rgb: "9CA3AF" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  }
};

interface TablasContingenciaViewProps {
  onBack: () => void;
  onNavigateToChat?: (chatId?: string) => void;
}

export function TablasContingenciaView({ onBack, onNavigateToChat }: TablasContingenciaViewProps) {
  const { sessionId, columns } = useDataContext();

  // State for variable selection
  const [rowVar, setRowVar] = useState<string>('');
  const [colVar, setColVar] = useState<string>('');
  const [segmentBy, setSegmentBy] = useState<string>('');
  const [selectedMetrics, setSelectedMetrics] = useState<string[]>(['frecuencia', 'pct_fila', 'pct_columna', 'pct_total']);

  // State for active segment (horizontal tabs)
  const [activeSegment, setActiveSegment] = useState<string>('General');

  // State for data
  const [tableData, setTableData] = useState<ContingencyTableResponse | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // AI Interpretation State
  const [analysisResult, setAnalysisResult] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [activeChatId, setActiveChatId] = useState<string | null>(null);
  const [isNavigating, setIsNavigating] = useState(false);

  // Dropdown states
  const [showRowDropdown, setShowRowDropdown] = useState(false);
  const [showColDropdown, setShowColDropdown] = useState(false);
  const [showSegmentDropdown, setShowSegmentDropdown] = useState(false);
  const [showMetricsDropdown, setShowMetricsDropdown] = useState(false);

  const rowRef = useRef<HTMLDivElement>(null);
  const colRef = useRef<HTMLDivElement>(null);
  const segmentRef = useRef<HTMLDivElement>(null);
  const metricsRef = useRef<HTMLDivElement>(null);

  // Click outside handlers
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (rowRef.current && !rowRef.current.contains(event.target as Node)) {
        setShowRowDropdown(false);
      }
      if (colRef.current && !colRef.current.contains(event.target as Node)) {
        setShowColDropdown(false);
      }
      if (segmentRef.current && !segmentRef.current.contains(event.target as Node)) {
        setShowSegmentDropdown(false);
      }
      if (metricsRef.current && !metricsRef.current.contains(event.target as Node)) {
        setShowMetricsDropdown(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Normalized table data structure
  const normalizedTableData = tableData ? {
    segments: tableData.segments || ['General'],
    tables: tableData.tables || { 'General': tableData as any },
    segment_by: tableData.segment_by || null
  } : null;

  const currentTable = normalizedTableData?.tables[activeSegment];

  // Fetch data when variables change
  useEffect(() => {
    const fetchData = async () => {
      if (!sessionId || !rowVar || !colVar) {
        setTableData(null);
        return;
      }

      if (rowVar === colVar) {
        setError('Las variables de fila y columna deben ser diferentes');
        return;
      }

      setLoading(true);
      setError(null);

      try {
        const response = await getCrosstabStats(sessionId, rowVar, colVar, segmentBy || undefined);
        setTableData(response);

        // Set active segment dynamically based on response
        if (response.segments && response.segments.length > 0) {
          setActiveSegment(response.segments[0]); // Select first available segment
        } else {
          setActiveSegment('General'); // Fallback to 'General'
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Error al calcular tabla de contingencia');
        setTableData(null);
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [sessionId, rowVar, colVar, segmentBy]);

  // Reset AI state when variables change
  useEffect(() => {
    setAnalysisResult(null);
    setActiveChatId(null);
  }, [rowVar, colVar, activeSegment]);

  const toggleMetric = (metric: string) => {
    setSelectedMetrics(prev =>
      prev.includes(metric) ? prev.filter(m => m !== metric) : [...prev, metric]
    );
  };

  // Helper: Get metric label
  const getMetricLabel = (metricKey: string): string => {
    const labels: Record<string, string> = {
      'frecuencia': 'N',
      'pct_fila': '% Fila',
      'pct_columna': '% Columna',
      'pct_total': '% Total'
    };
    return labels[metricKey] || metricKey;
  };

  // Helper: Get value from cell for metric
  const getCellValue = (data: ContingencyCellData, metricKey: string): number => {
    switch (metricKey) {
      case 'frecuencia': return data.count;
      case 'pct_fila': return data.row_percent;
      case 'pct_columna': return data.col_percent;
      case 'pct_total': return data.total_percent;
      default: return 0;
    }
  };

  // ========================================================================
  // EXPORT FUNCTIONS (CRÍTICO: AoA Implementation)
  // ========================================================================
  const handleExportExcel = () => {
    try {
      console.log('Iniciando exportación Contingencia (AoA)...');

      if (!normalizedTableData) {
        alert('No hay datos para exportar.');
        return;
      }

      if (!XLSX || !XLSX.utils) {
        throw new Error('Librería XLSX no cargada');
      }

      const wb = XLSX.utils.book_new();
      let hasData = false;

      // Iterar sobre todos los segmentos para crear hojas
      normalizedTableData.segments.forEach(segmentName => {
        const table = normalizedTableData.tables[segmentName];
        if (!table) return;

        // Construir Array of Arrays (AoA)
        const aoaData: any[][] = [];

        // 1. Encabezados Principales
        // Fila 0: Títulos generales
        const row0 = [
          `Tabla de Contingencia: ${table.row_variable} vs ${table.col_variable}`,
          '',
          ...Array(table.col_categories.length).fill('')
        ];
        aoaData.push(row0);

        // Fila 1: Segmento
        const row1 = [`Segmento: ${segmentName}`, '', ...Array(table.col_categories.length).fill('')];
        aoaData.push(row1);
        aoaData.push([]); // Espacio

        // 2. Encabezados de Tabla
        // Fila Headers: [Var Fila \ Var Col] | [Col Cat 1] | [Col Cat 2] ... | Total
        const headerRow = [
          `${table.row_variable} \\ ${table.col_variable}`,
          ...table.col_categories,
          'Total'
        ];
        aoaData.push(headerRow);

        // 3. Filas de Datos
        // Iteramos filas (categorías de variable fila)
        table.row_categories.forEach(rowCat => {
          // Para contingencia, a veces queremos mostrar múltiples métricas por celda.
          // Para Excel plano, una opción es mostrar N (frecuencia) principal, o múltiples filas por categoría.
          // Siguiendo el requerimiento de "Matriz", haremos una fila por categoría mostrando FRECUENCIA (N).
          // Si el usuario quiere ver porcentajes, técnicamente necesitaríamos más filas o celdas complejas.
          // Por simplicidad y robustez del "AoA" solicitado, exportaremos FRECUENCIA principal.
          // Opcionalmente, podemos añadir filas adicionales para %, pero el usuario pidió "Matriz de Arrays".

          // Vamos a exportar un bloque por métrica seleccionada para ser completos pero ordenados.

          selectedMetrics.forEach(metric => {
            const metricLabel = getMetricLabel(metric);
            const rowData: any[] = [`${rowCat} (${metricLabel})`];

            // Valores para cada columna
            table.col_categories.forEach(colCat => {
              const cell = table.cells[rowCat][colCat];
              const val = getCellValue(cell, metric);
              // Formato: si es pct, dividir por 100 para formato Excel %
              rowData.push(metric === 'frecuencia' ? val : val / 100);
            });

            // Total Fila
            const totalCell = table.row_totals[rowCat];
            const totalVal = getCellValue(totalCell, metric);
            rowData.push(metric === 'frecuencia' ? totalVal : totalVal / 100);

            aoaData.push(rowData);
          });
        });

        // 4. Fila de Totales de Columna
        selectedMetrics.forEach(metric => {
          const metricLabel = getMetricLabel(metric);
          const totalRow = [`Total General (${metricLabel})`];

          table.col_categories.forEach(colCat => {
            const colTotal = table.col_totals[colCat];
            const val = getCellValue(colTotal, metric);
            totalRow.push(metric === 'frecuencia' ? val : val / 100);
          });

          // Grand Total
          const grandTotal = metric === 'frecuencia' ? table.grand_total : 100.0;
          totalRow.push(metric === 'frecuencia' ? grandTotal : grandTotal / 100);

          aoaData.push(totalRow);
        });

        // Crear hoja desde AoA
        const ws = XLSX.utils.aoa_to_sheet(aoaData);

        // Aplicar estilos
        // Rango de datos empieza en fila 4 (índice 3)
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

        // Estilo Header
        for (let C = 0; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: 3, c: C });
          if (ws[addr]) ws[addr].s = excelStyles.categoryHeader;
        }

        // Estilos de celdas de datos (porcentajes)
        // Recorremos desde fila 4 hasta el final
        for (let R = 4; R <= range.e.r; ++R) {
          for (let C = 1; C <= range.e.c; ++C) { // Saltamos primera columna (etiquetas)
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[addr]) continue;

            // Detectar si es fila de porcentaje (basado en etiqueta en col 0, o lógica de bloque)
            // Simplificación: Mirar valor. Si es <= 1 y no es entero, probablemente porcentaje.
            // Mejor: Usar el índice de fila para saber qué métrica es.
            // Como es complejo, aplicamos formato numérico genérico o % si es pequeño.

            const val = ws[addr].v;
            if (typeof val === 'number' && val <= 1 && val !== 0 && !Number.isInteger(val)) {
              ws[addr].t = 'n';
              ws[addr].s = { numFmt: '0.00%' };
            } else {
              ws[addr].t = 'n';
              ws[addr].s = excelStyles.cellNumber;
            }
          }
        }

        // Ajustar anchos
        ws['!cols'] = [{ wch: 30 }, ...Array(table.col_categories.length + 1).fill({ wch: 12 })];

        // Añadir hoja al libro
        XLSX.utils.book_append_sheet(wb, ws, segmentName.substring(0, 31));
        hasData = true;
      });

      if (!hasData) {
        alert('No se generaron datos para exportar.');
        return;
      }

      const fileName = `Contingencia_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, fileName);
      console.log('Exportación Contingencia completada.');

    } catch (error) {
      console.error('Error exporting Contingency Excel:', error);
      alert(`Error al exportar: ${error instanceof Error ? error.message : String(error)}`);
    }
  };

  // Exportar a PDF
  const handleExportPDF = async () => {
    if (!normalizedTableData || !currentTable) {
      alert('No hay datos para exportar');
      return;
    }

    try {
      const jsPDFModule = await import('jspdf');
      const autoTableModule = await import('jspdf-autotable');
      const jsPDF = jsPDFModule.default || jsPDFModule.jsPDF;
      const autoTable = autoTableModule.default;

      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();
      let yPos = 20;

      // Headers PDF
      doc.setFontSize(16);
      doc.setFont('helvetica', 'bold');
      doc.text('Tabla de Contingencia', pageWidth / 2, yPos, { align: 'center' });
      yPos += 10;
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.text(`Generado: ${new Date().toLocaleString()}`, pageWidth / 2, yPos, { align: 'center' });
      yPos += 10;

      // Iterar segmentos
      normalizedTableData.segments.forEach((segmentName, segIdx) => {
        const table = normalizedTableData.tables[segmentName];
        if (!table) return;

        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text(`Segmento: ${segmentName}`, 14, yPos);
        yPos += 6;

        const head = [['', ...table.col_categories, 'Total']];

        // Construir body para autotable
        const body: any[] = [];

        table.row_categories.forEach(rowCat => {
          selectedMetrics.forEach((metric, idx) => {
            const label = idx === 0 ? rowCat : ''; // Agrupar visualmente
            const metricLabel = getMetricLabel(metric);
            const row = [`${label} ${idx === 0 ? '' : `(${metricLabel})`}`];

            // Cols
            table.col_categories.forEach(colCat => {
              const val = getCellValue(table.cells[rowCat][colCat], metric);
              row.push(metric === 'frecuencia' ? val.toString() : val.toFixed(1) + '%');
            });
            // Total
            const tot = getCellValue(table.row_totals[rowCat], metric);
            row.push(metric === 'frecuencia' ? tot.toString() : tot.toFixed(1) + '%');

            body.push(row);
          });
        });

        autoTable(doc, {
          startY: yPos,
          head: head,
          body: body,
          theme: 'grid',
          styles: { fontSize: 8 },
          headStyles: { fillColor: [220, 220, 220], textColor: 20 },
          margin: { top: 20 }
        });

        yPos = (doc as any).lastAutoTable.finalY + 15;
        if (yPos > 250) { doc.addPage(); yPos = 20; }
      });

      doc.save(`Contingencia_${rowVar}_vs_${colVar}.pdf`);

    } catch (error) {
      console.error('Error al exportar PDF:', error);
      alert('Error al exportar PDF.');
    }
  };

  // ========================================================================
  // AI INTERPRETATION HANDLERS
  // ========================================================================
  const generateContingencyContext = (): string => {
    if (!rowVar || !colVar || !currentTable) return "No hay datos.";
    return `Tabla de contingencia ${rowVar} vs ${colVar}. Total: ${currentTable.grand_total}.`;
  };

  const handleAIInterpretation = async () => {
    if (analysisResult || isAnalyzing) return;
    if (!rowVar || !colVar) { alert('Selecciona variables primero'); return; }

    setIsAnalyzing(true);
    try {
      const prompt = `Analiza tabla contingencia: ${rowVar} vs ${colVar}. Contexto: ${generateContingencyContext()}.`;
      const response = await sendChatMessage({
        session_id: sessionId,
        message: prompt,
        history: []
      });
      if (response.success) setAnalysisResult(response.response);
      else setError("Error IA");
    } catch (e) { setError("Error conexión IA"); }
    finally { setIsAnalyzing(false); }
  };

  const handleContinueToChat = async () => {
    if (!onNavigateToChat || !sessionId) return;
    if (onNavigateToChat) onNavigateToChat(activeChatId || undefined);
  };


  // Render
  return (
    <div className="h-full flex flex-col bg-slate-50">
      {/* Header */}
      <div className="bg-white border-b border-slate-200 px-6 py-6 shadow-sm">
        <div className="flex items-center gap-4 mb-4">
          <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-lg">
            <ArrowLeft className="w-5 h-5 text-slate-600" />
          </button>
          <h2 className="text-2xl font-bold text-slate-900">Tablas de Contingencia</h2>
        </div>

        {/* Controls */}
        <div className="flex items-center gap-3 ml-4">
          {/* Row Var */}
          <div className="relative" ref={rowRef}>
            <button onClick={() => setShowRowDropdown(!showRowDropdown)} className="px-4 py-2 border rounded-lg bg-white">
              {rowVar ? `Fila: ${rowVar}` : 'Variable Fila'} <ChevronDown className="inline w-4 h-4" />
            </button>
            {showRowDropdown && (
              <div className="absolute top-full bg-white border shadow-lg z-10 max-h-60 overflow-auto w-64">
                {columns.map(c => (
                  <div key={c} onClick={() => { setRowVar(c); setShowRowDropdown(false); }} className="px-4 py-2 hover:bg-gray-50 cursor-pointer">
                    {c}
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Col Var */}
          <div className="relative" ref={colRef}>
            <button onClick={() => setShowColDropdown(!showColDropdown)} className="px-4 py-2 border rounded-lg bg-white">
              {colVar ? `Columna: ${colVar}` : 'Variable Columna'} <ChevronDown className="inline w-4 h-4" />
            </button>
            {showColDropdown && (
              <div className="absolute top-full bg-white border shadow-lg z-10 max-h-60 overflow-auto w-64">
                {columns.map(c => (
                  <div key={c} onClick={() => { setColVar(c); setShowColDropdown(false); }} className="px-4 py-2 hover:bg-gray-50 cursor-pointer">
                    {c}
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Segment By */}
          <div className="relative" ref={segmentRef}>
            <button onClick={() => setShowSegmentDropdown(!showSegmentDropdown)} className="px-4 py-2 border rounded-lg bg-white">
              {segmentBy ? `Segmento: ${segmentBy}` : 'Segmentar'} <ChevronDown className="inline w-4 h-4" />
            </button>
            {showSegmentDropdown && (
              <div className="absolute top-full bg-white border shadow-lg z-10 max-h-60 overflow-auto w-64">
                <div onClick={() => { setSegmentBy(''); setShowSegmentDropdown(false); }} className="px-4 py-2 hover:bg-gray-50 cursor-pointer italic">Sin segmentación</div>
                {columns.map(c => (
                  <div key={c} onClick={() => { setSegmentBy(c); setShowSegmentDropdown(false); }} className="px-4 py-2 hover:bg-gray-50 cursor-pointer">{c}</div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 overflow-auto p-8">
        {loading && <div className="flex justify-center"><Loader2 className="animate-spin" /></div>}
        {error && <div className="text-red-500 bg-red-50 p-4 rounded">{error}</div>}

        {!loading && !error && currentTable && (
          <div className="bg-white rounded-lg shadow border p-6 overflow-auto">
            <h3 className="text-lg font-bold mb-4">{rowVar} vs {colVar} {segmentBy ? `(${activeSegment})` : ''}</h3>

            <table className="w-full border-collapse">
              <thead>
                <tr>
                  <th className="border p-2 bg-gray-100">{rowVar} \ {colVar}</th>
                  {currentTable.col_categories.map(c => <th key={c} className="border p-2 bg-gray-50">{c}</th>)}
                  <th className="border p-2 bg-gray-100">Total</th>
                </tr>
              </thead>
              <tbody>
                {currentTable.row_categories.map(rowCat => (
                  <tr key={rowCat}>
                    <td className="border p-2 font-medium">{rowCat}</td>
                    {currentTable.col_categories.map(colCat => (
                      <td key={colCat} className="border p-2 text-center">
                        <div className="flex flex-col text-xs">
                          <span className="font-bold text-sm">{currentTable.cells[rowCat][colCat].count}</span>
                          <span className="text-gray-500">{currentTable.cells[rowCat][colCat].row_percent.toFixed(1)}%</span>
                        </div>
                      </td>
                    ))}
                    <td className="border p-2 text-center font-bold">
                      {currentTable.row_totals[rowCat].count}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {!loading && !error && !currentTable && (
          <div className="text-center text-gray-500 mt-20">Selecciona variables para ver la tabla</div>
        )}
      </div>

      {/* Toolbar */}
      {currentTable && (
        <ActionToolbar
          onExportExcel={handleExportExcel}
          onExportPDF={handleExportPDF}
          onAIInterpretation={handleAIInterpretation}
          onContinueToChat={handleContinueToChat}
          isAnalyzing={isAnalyzing}
          isNavigating={isNavigating}
        />
      )}
    </div>
  );
}