import { ArrowLeft, ChevronDown, X, Loader2, Pencil, Save, Sparkles } from 'lucide-react';
import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { useDataContext } from '../../context/DataContext';
import { getFrequencyStats } from '../../api/stats';
import { sendChatMessage } from '../../api/ai';
import { ActionToolbar } from './ActionToolbar';
import type { FrequencyTableResult, FrequencyRow } from '../../types/stats';

// Estilos para Excel WYSIWYG
const excelStyles = {
  header: {
    fill: { fgColor: { rgb: "F3F4F6" } },
    font: { bold: true, color: { rgb: "1F2937" }, sz: 11 },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "medium", color: { rgb: "9CA3AF" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  cell: {
    font: { sz: 10 },
    border: {
      top: { style: "thin", color: { rgb: "E5E7EB" } },
      bottom: { style: "thin", color: { rgb: "E5E7EB" } },
      left: { style: "thin", color: { rgb: "E5E7EB" } },
      right: { style: "thin", color: { rgb: "E5E7EB" } }
    },
    alignment: { vertical: "center" }
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
  total: {
    fill: { fgColor: { rgb: "E5E7EB" } },
    font: { bold: true, sz: 10 },
    border: {
      top: { style: "medium", color: { rgb: "9CA3AF" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  totalLabel: {
    fill: { fgColor: { rgb: "E5E7EB" } },
    font: { bold: true, sz: 10 },
    border: {
      top: { style: "medium", color: { rgb: "9CA3AF" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { vertical: "center" }
  }
};

interface TablasFrecuenciaViewProps {
  onBack: () => void;
  onNavigateToChat?: (chatId?: string) => void;
}

export function TablasFrecuenciaView({ onBack, onNavigateToChat }: TablasFrecuenciaViewProps) {
  const { sessionId, columns } = useDataContext();
  const [selectedVars, setSelectedVars] = useState<string[]>([]);
  const [segmentBy, setSegmentBy] = useState<string>('');
  const [showVarsDropdown, setShowVarsDropdown] = useState(false);
  const [showSegmentDropdown, setShowSegmentDropdown] = useState(false);
  const [activeSegment, setActiveSegment] = useState<string>('General');

  // Nueva estructura de datos segmentados
  const [segments, setSegments] = useState<string[]>([]);
  const [tablesData, setTablesData] = useState<Record<string, FrequencyTableResult[]>>({});
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Estado para modo edición
  const [isEditing, setIsEditing] = useState(false);
  const [editedData, setEditedData] = useState<Record<string, FrequencyTableResult[]>>({});

  // AI Interpretation State
  const [analysisResult, setAnalysisResult] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [activeChatId, setActiveChatId] = useState<string | null>(null);
  const [isNavigating, setIsNavigating] = useState(false);

  const varsRef = useRef<HTMLDivElement>(null);
  const segmentRef = useRef<HTMLDivElement>(null);

  const categoricalVariables = columns;

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (varsRef.current && !varsRef.current.contains(event.target as Node)) {
        setShowVarsDropdown(false);
      }
      if (segmentRef.current && !segmentRef.current.contains(event.target as Node)) {
        setShowSegmentDropdown(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Cargar datos de frecuencia cuando cambian las variables o segmentación
  useEffect(() => {
    async function fetchFrequencyData() {
      if (!sessionId || selectedVars.length === 0) {
        setSegments([]);
        setTablesData({});
        return;
      }

      setIsLoading(true);
      setError(null);

      try {
        const response = await getFrequencyStats(sessionId, selectedVars, segmentBy || undefined);
        setSegments(response.segments);
        setTablesData(response.tables);
        setActiveSegment('General');
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Error al cargar frecuencias');
        console.error('Error fetching frequency stats:', err);
      } finally {
        setIsLoading(false);
      }
    }

    fetchFrequencyData();
  }, [sessionId, selectedVars, segmentBy]);

  // Sincronizar editedData cuando tablesData cambia
  useEffect(() => {
    setEditedData(JSON.parse(JSON.stringify(tablesData)));
    setIsEditing(false);
  }, [tablesData]);

  // Reset AI state when selection changes
  useEffect(() => {
    setAnalysisResult(null);
    setActiveChatId(null);
  }, [selectedVars, segmentBy, activeSegment]);

  // Handler para editar celda
  const handleCellEdit = (
    segment: string,
    variableName: string,
    rowIndex: number,
    field: 'categoria' | 'frecuencia' | 'porcentaje',
    value: string
  ) => {
    setEditedData(prev => {
      const newData = { ...prev };
      const tables = [...(newData[segment] || [])];
      const tableIndex = tables.findIndex(t => t.variable === variableName);

      if (tableIndex === -1) return prev;

      const table = { ...tables[tableIndex] };
      const rows = [...table.rows];
      const row = { ...rows[rowIndex] };

      if (field === 'categoria') {
        row.categoria = value;
      } else if (field === 'frecuencia') {
        const numValue = parseInt(value) || 0;
        row.frecuencia = numValue;
        // Recalcular porcentaje si es posible
        const newTotal = rows.reduce((sum, r, i) =>
          sum + (i === rowIndex ? numValue : r.frecuencia), 0
        );
        if (newTotal > 0) {
          row.porcentaje = (numValue / newTotal) * 100;
          // Recalcular todos los porcentajes
          rows.forEach((r, i) => {
            if (i !== rowIndex) {
              r.porcentaje = (r.frecuencia / newTotal) * 100;
            }
          });
        }
        table.total = newTotal;
      } else if (field === 'porcentaje') {
        row.porcentaje = parseFloat(value) || 0;
      }

      rows[rowIndex] = row;
      // Recalcular porcentaje acumulado
      let cumulative = 0;
      rows.forEach(r => {
        cumulative += r.porcentaje;
        r.porcentaje_acumulado = cumulative;
      });

      table.rows = rows;
      tables[tableIndex] = table;
      newData[segment] = tables;

      return newData;
    });
  };

  // Guardar cambios (aplicar editedData como nuevo tablesData)
  const handleSaveEdits = () => {
    setTablesData(editedData);
    setIsEditing(false);
  };

  // Cancelar edición
  const handleCancelEdit = () => {
    setEditedData(JSON.parse(JSON.stringify(tablesData)));
    setIsEditing(false);
  };

  const toggleVariable = (variable: string) => {
    setSelectedVars(prev =>
      prev.includes(variable) ? prev.filter(v => v !== variable) : [...prev, variable]
    );
  };

  // Exportar a Excel con estilos WYSIWYG
  const handleExportExcel = () => {
    try {
      console.log('Iniciando exportación a Excel...');

      if (!XLSX || !XLSX.utils) {
        throw new Error('La biblioteca de Excel no se ha cargado correctamente.');
      }

      const dataSource = isEditing ? editedData : tablesData;
      const wb = XLSX.utils.book_new();
      let hasData = false;

      // Iterar sobre las variables seleccionadas
      for (const variable of selectedVars) {
        // Preparar datos planos (Array of Objects) para json_to_sheet
        const flatData: any[] = [];

        // Determinar segmentos a exportar
        const segmentsToExport = segmentBy ? segments : [segments[0] || 'General'];

        segmentsToExport.forEach(segment => {
          const tables = dataSource[segment] || [];
          const table = tables.find(t => t.variable === variable);

          if (table) {
            // Filas de datos
            table.rows.forEach(row => {
              const rowData: any = {
                'Categoría': row.categoria,
                'N': row.frecuencia,
                '%': parseFloat(row.porcentaje.toFixed(1)),
                '% Acum.': parseFloat(row.porcentaje_acumulado.toFixed(1))
              };

              if (segmentBy) {
                rowData['Segmento'] = segment;
              }

              flatData.push(rowData);
            });

            // Fila de Total
            const totalRow: any = {
              'Categoría': 'Total',
              'N': table.total,
              '%': 100.0,
              '% Acum.': 100.0
            };
            if (segmentBy) {
              totalRow['Segmento'] = segment;
            }
            flatData.push(totalRow);
            flatData.push({}); // Fila vacía para separar segmentos
          }
        });

        if (flatData.length > 0) {
          // Crear hoja desde JSON
          const ws = XLSX.utils.json_to_sheet(flatData);

          // Aplicar estilos básicos a la cabecera
          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

          for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_col(C) + "1"; // Primera fila
            if (!ws[address]) continue;
            ws[address].s = excelStyles.header;
          }

          // Ajustar anchos de columna
          ws['!cols'] = segmentBy
            ? [{ wch: 15 }, { wch: 30 }, { wch: 10 }, { wch: 10 }, { wch: 12 }]
            : [{ wch: 30 }, { wch: 10 }, { wch: 10 }, { wch: 12 }];

          XLSX.utils.book_append_sheet(wb, ws, variable.substring(0, 31));
          hasData = true;
        }
      }

      if (!hasData) {
        alert('No se encontraron datos para exportar.');
        return;
      }

      XLSX.writeFile(wb, `Tablas_Frecuencia${segmentBy ? `_por_${segmentBy}` : ''}.xlsx`);
      console.log('Exportación completada.');

    } catch (error) {
      console.error('Error exporting Excel:', error);
      alert(`Error al exportar: ${error instanceof Error ? error.message : String(error)}`);
    }
  };

  // Exportar a PDF
  const handleExportPDF = async () => {
    if (segments.length === 0) {
      alert('No hay datos para exportar');
      return;
    }

    try {
      // Importación dinámica de jspdf y autotable
      const jsPDFModule = await import('jspdf');
      const autoTableModule = await import('jspdf-autotable');

      const jsPDF = jsPDFModule.default || jsPDFModule.jsPDF;
      const autoTable = autoTableModule.default;

      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();
      let yPos = 20;

      // Título
      doc.setFontSize(16);
      doc.setFont('helvetica', 'bold');
      doc.text('Tablas de Frecuencia', pageWidth / 2, yPos, { align: 'center' });
      yPos += 8;

      // Fecha
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.text(`Generado: ${new Date().toLocaleString()}`, pageWidth / 2, yPos, { align: 'center' });
      yPos += 5;

      if (segmentBy) {
        doc.text(`Segmentado por: ${segmentBy}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 5;
      }
      yPos += 10;

      // Para cada variable
      for (const variable of selectedVars) {
        // Título de variable
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text(`Variable: ${variable}`, 14, yPos);
        yPos += 8;

        // Para cada segmento
        const dataSource = isEditing ? editedData : tablesData;
        for (const segment of segments) {
          const tables = dataSource[segment] || [];
          const table = tables.find(t => t.variable === variable);

          if (!table) continue;

          if (segmentBy && segments.length > 1) {
            doc.setFontSize(10);
            doc.setFont('helvetica', 'italic');
            doc.text(`Segmento: ${segment}`, 14, yPos);
            yPos += 6;
          }

          const tableData = table.rows.map(row => [
            row.categoria,
            row.frecuencia.toString(),
            `${row.porcentaje.toFixed(1)}`,
            `${row.porcentaje_acumulado.toFixed(1)}`
          ]);

          // Agregar fila de total
          tableData.push(['Total', table.total.toString(), '100.0', '100.0']);

          autoTable(doc, {
            startY: yPos,
            head: [['Categoría', 'N', '%', '% Acum.']],
            body: tableData,
            theme: 'grid',
            headStyles: {
              fillColor: [243, 244, 246],
              textColor: [31, 41, 55],
              lineColor: [209, 213, 219],
              lineWidth: 0.1,
              fontStyle: 'bold'
            },
            bodyStyles: {
              lineColor: [209, 213, 219],
              lineWidth: 0.1,
              textColor: [55, 65, 81]
            },
            footStyles: {
              fillColor: [229, 231, 235],
              textColor: [0, 0, 0],
              fontStyle: 'bold'
            },
            alternateRowStyles: { fillColor: [255, 255, 255] },
            margin: { left: 14, right: 14 },
            styles: { fontSize: 9, cellPadding: 3 },
            didParseCell: (data) => {
              // Estilo especial para fila Total
              if (data.row.index === tableData.length - 1) {
                data.cell.styles.fillColor = [229, 231, 235];
                data.cell.styles.fontStyle = 'bold';
              }
            }
          });

          yPos = (doc as any).lastAutoTable.finalY + 10;

          // Nueva página si es necesario
          if (yPos > 250) {
            doc.addPage();
            yPos = 20;
          }
        }

        yPos += 5;
      }

      doc.save(`Tablas_Frecuencia${segmentBy ? `_por_${segmentBy}` : ''}.pdf`);
    } catch (error) {
      console.error('Error al exportar PDF:', error);
      alert('Error al exportar a PDF. Por favor, intenta de nuevo.');
    }
  };

  // Verificar si hay datos para habilitar acciones
  const hasData = segments.length > 0 && selectedVars.length > 0;

  // Helper function to generate context for AI
  const generateFrequencyContext = (): string => {
    if (!selectedVars.length || !activeSegment) {
      return "No hay datos disponibles.";
    }

    const currentData = (isEditing ? editedData : tablesData)[activeSegment] || [];
    if (currentData.length === 0) {
      return "No hay tablas de frecuencia disponibles.";
    }

    let context = `Análisis de Frecuencias${segmentBy ? ` (Segmento: ${activeSegment})` : ''}:\n\n`;

    currentData.forEach(table => {
      context += `Variable: ${table.variable}\n`;
      context += `Total de casos: ${table.total}\n`;
      context += `Distribución:\n`;

      table.rows.forEach(row => {
        if (row.categoria !== '(Perdidos)') {
          context += `  - ${row.categoria}: ${row.frecuencia} casos (${row.porcentaje.toFixed(1)}%)\n`;
        }
      });

      // Identify mode (highest frequency)
      const validRows = table.rows.filter(r => r.categoria !== '(Perdidos)');
      if (validRows.length > 0) {
        const modeRow = validRows.reduce((max, row) => row.frecuencia > max.frecuencia ? row : max);
        context += `  Moda: ${modeRow.categoria} (${modeRow.frecuencia} casos)\n`;
      }
      context += '\n';
    });

    return context;
  };

  // Handler para interpretación con IA
  const handleAIInterpretation = async () => {
    if (analysisResult || isAnalyzing) return; // Ya existe o está cargando

    if (selectedVars.length === 0) {
      alert('Selecciona al menos una variable primero');
      return;
    }

    setIsAnalyzing(true);
    try {
      const frequencyContext = generateFrequencyContext();
      const prompt = `Actúa como un bioestadístico. Analiza la siguiente tabla de frecuencias para las variables seleccionadas. Identifica la moda, la heterogeneidad de los grupos y cualquier desbalance notable. Sé breve y clínico.\n\n${frequencyContext}`;

      const response = await sendChatMessage({
        session_id: sessionId || undefined,
        chat_id: activeChatId || undefined,
        message: prompt,
        history: []
      });

      if (response.success) {
        setAnalysisResult(response.response);
        if (response.chat_id) {
          setActiveChatId(response.chat_id);
        }
      } else {
        setError("No se pudo obtener la interpretación de IA.");
      }
    } catch (err) {
      console.error("AI Error:", err);
      setError("Error de conexión con el servicio de IA.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Handler para continuar al chat
  const handleContinueToChat = async () => {
    if (!onNavigateToChat || !sessionId) return;

    let finalChatId = activeChatId;

    try {
      setIsNavigating(true);

      // Si no hay análisis previo, enviar contexto silenciosamente
      if (!activeChatId) {
        const frequencyContext = generateFrequencyContext();
        const interpretationContext = analysisResult
          ? `\n\nInterpretación Previa Realizada:\n${analysisResult}`
          : '\n\n(El usuario aún no ha solicitado una interpretación automática)';

        const handoffMessage = `**SYSTEM CONTEXT HANDOFF**
        
Estás recibiendo el contexto de la sesión actual de "Tablas de Frecuencia" para asistir al usuario.

DATOS ESTRUCTURADOS:
${frequencyContext}
${interpretationContext}

INSTRUCCIÓN:
El usuario ha sido transferido al chat principal. Mantén este contexto en memoria para responder preguntas sobre estas variables. Saluda brevemente confirmando que tienes los datos.`;

        const response = await sendChatMessage({
          session_id: sessionId,
          chat_id: activeChatId || undefined,
          message: handoffMessage,
          history: []
        });

        if (response.chat_id) {
          finalChatId = response.chat_id;
          setActiveChatId(response.chat_id);
        }
      }

    } catch (error) {
      console.error("Error transfiriendo contexto:", error);
    } finally {
      setIsNavigating(false);
      onNavigateToChat(finalChatId || undefined);
    }
  };

  // Componente FrequencyTable editable
  const FrequencyTable = ({
    title,
    data,
    total,
    variableName,
    segment,
    editable,
    onEdit
  }: {
    title?: string;
    data: FrequencyRow[];
    total: number;
    variableName: string;
    segment: string;
    editable: boolean;
    onEdit: (rowIndex: number, field: 'categoria' | 'frecuencia' | 'porcentaje', value: string) => void;
  }) => (
    <div className={`bg-white mb-6 rounded-lg border overflow-hidden shadow-sm ${editable ? 'border-blue-300 ring-2 ring-blue-100' : 'border-slate-200'}`}>
      {title && (
        <div className={`px-4 py-3 border-b ${editable ? 'bg-blue-50 border-blue-200' : 'bg-slate-50 border-slate-200'}`}>
          <h3 className="text-sm font-semibold text-slate-800">
            Tabla: {title}
            {editable && <span className="ml-2 text-xs text-blue-600">(Editando)</span>}
          </h3>
        </div>
      )}
      <div className="overflow-x-auto">
        <table className="w-full" style={{ fontVariantNumeric: 'tabular-nums' }}>
          <thead>
            <tr className="bg-slate-100 border-b-2 border-slate-300">
              <th className="text-left px-4 py-3 text-sm font-bold text-slate-800 border border-slate-200">Categoría</th>
              <th className="text-center px-4 py-3 text-sm font-bold text-slate-800 border border-slate-200">N</th>
              <th className="text-center px-4 py-3 text-sm font-bold text-slate-800 border border-slate-200">%</th>
              <th className="text-center px-4 py-3 text-sm font-bold text-slate-800 border border-slate-200">% Acum.</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, index) => (
              <tr key={index} className={`transition-colors ${editable ? 'bg-white' : 'hover:bg-slate-50'}`}>
                <td className={`px-4 py-3 text-sm border border-slate-200 ${row.categoria === '(Perdidos)' ? 'text-slate-500 italic bg-amber-50/50' : 'text-slate-900'}`}>
                  {editable ? (
                    <input
                      type="text"
                      value={row.categoria}
                      onChange={(e) => onEdit(index, 'categoria', e.target.value)}
                      className="w-full bg-transparent border-none outline-none focus:ring-1 focus:ring-blue-400 rounded px-1 -mx-1"
                    />
                  ) : row.categoria}
                </td>
                <td className="px-4 py-3 text-sm text-slate-900 text-center font-mono border border-slate-200">
                  {editable ? (
                    <input
                      type="number"
                      value={row.frecuencia}
                      onChange={(e) => onEdit(index, 'frecuencia', e.target.value)}
                      className="w-full bg-transparent border-none outline-none text-center focus:ring-1 focus:ring-blue-400 rounded"
                      min="0"
                    />
                  ) : row.frecuencia}
                </td>
                <td className="px-4 py-3 text-sm text-slate-900 text-center font-mono border border-slate-200">
                  {row.porcentaje.toFixed(1)}
                </td>
                <td className="px-4 py-3 text-sm text-slate-900 text-center font-mono border border-slate-200">
                  {row.porcentaje_acumulado.toFixed(1)}
                </td>
              </tr>
            ))}
            <tr className="bg-slate-100 border-t-2 border-slate-300">
              <td className="px-4 py-3 text-sm font-bold text-slate-900 border border-slate-200">Total</td>
              <td className="px-4 py-3 text-sm font-bold text-slate-900 text-center font-mono border border-slate-200">{total}</td>
              <td className="px-4 py-3 text-sm font-bold text-slate-900 text-center font-mono border border-slate-200">100.0</td>
              <td className="px-4 py-3 text-sm font-bold text-slate-900 text-center font-mono border border-slate-200">100.0</td>
            </tr>
          </tbody>
        </table>
      </div>
      <p className="text-xs text-slate-500 px-4 py-2 bg-slate-50 border-t border-slate-200">
        {editable ? 'Edita las celdas directamente. Los porcentajes se recalculan automáticamente.' : 'Nota: Los porcentajes pueden no sumar exactamente 100% debido al redondeo.'}
      </p>
    </div>
  );

  // Obtener tablas del segmento activo (usar editedData si está editando)
  const currentTables = (isEditing ? editedData : tablesData)[activeSegment] || [];

  return (
    <div className="h-full flex flex-col bg-slate-50">
      {/* Header */}
      <div className="bg-white border-b border-slate-200 px-6 py-6 shadow-sm">
        <div className="flex items-center gap-4 mb-4">
          <button
            onClick={onBack}
            className="p-2 hover:bg-slate-100 rounded-lg transition-colors"
            title="Volver al menú"
          >
            <ArrowLeft className="w-5 h-5 text-slate-600" />
          </button>
          <div className="flex-1">
            <div className="flex items-center gap-3 mb-2">
              <div className="w-1 h-6 bg-gradient-to-b from-teal-500 to-teal-600 rounded-full"></div>
              <h2 className="text-2xl font-bold text-slate-900 tracking-tight">Tablas de Frecuencia</h2>
            </div>
            <p className="text-slate-600 leading-relaxed ml-4">Análisis de distribución de variables categóricas</p>
          </div>
        </div>

        {/* Control Bar */}
        <div className="flex items-center gap-3 ml-4 flex-wrap">
          {/* Multi-select Variables */}
          <div className="relative" ref={varsRef}>
            <button
              onClick={() => setShowVarsDropdown(!showVarsDropdown)}
              className="px-4 py-2.5 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors flex items-center gap-2 text-sm min-w-[280px] shadow-sm"
            >
              <span className="text-slate-700">
                {selectedVars.length === 0 ? 'Selecciona variables' : `${selectedVars.length} variable(s)`}
              </span>
              <ChevronDown className="w-4 h-4 text-slate-500 ml-auto" />
            </button>

            {showVarsDropdown && (
              <div className="absolute top-full mt-1 w-full bg-white border border-slate-300 rounded-lg shadow-lg z-10 max-h-64 overflow-auto">
                {categoricalVariables.length === 0 ? (
                  <div className="px-4 py-3 text-sm text-slate-500 italic">
                    No hay columnas disponibles
                  </div>
                ) : (
                  categoricalVariables.map((variable) => (
                    <label
                      key={variable}
                      className="flex items-center gap-3 px-4 py-2.5 hover:bg-slate-50 cursor-pointer"
                    >
                      <input
                        type="checkbox"
                        checked={selectedVars.includes(variable)}
                        onChange={() => toggleVariable(variable)}
                        className="rounded border-slate-300 text-teal-600 focus:ring-teal-500"
                      />
                      <span className="text-sm text-slate-900">{variable}</span>
                    </label>
                  ))
                )}
              </div>
            )}
          </div>

          {/* Segment By Dropdown */}
          <div className="relative" ref={segmentRef}>
            <button
              onClick={() => setShowSegmentDropdown(!showSegmentDropdown)}
              className="px-4 py-2.5 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors flex items-center gap-2 text-sm min-w-[220px] shadow-sm"
            >
              <span className="text-slate-700">
                {segmentBy ? `Segmentar por: ${segmentBy}` : 'Segmentar por: (opcional)'}
              </span>
              <ChevronDown className="w-4 h-4 text-slate-500 ml-auto" />
            </button>

            {showSegmentDropdown && (
              <div className="absolute top-full mt-1 w-full bg-white border border-slate-300 rounded-lg shadow-lg z-10 max-h-64 overflow-auto">
                <button
                  onClick={() => {
                    setSegmentBy('');
                    setShowSegmentDropdown(false);
                  }}
                  className="w-full text-left px-4 py-2.5 hover:bg-slate-50 text-sm text-slate-600 italic"
                >
                  Sin segmentación
                </button>
                {categoricalVariables
                  .filter(v => !selectedVars.includes(v))
                  .map((variable) => (
                    <button
                      key={variable}
                      onClick={() => {
                        setSegmentBy(variable);
                        setShowSegmentDropdown(false);
                      }}
                      className="w-full text-left px-4 py-2.5 hover:bg-slate-50 text-sm text-slate-900"
                    >
                      {variable}
                    </button>
                  ))}
              </div>
            )}
          </div>

          {(selectedVars.length > 0 || segmentBy) && (
            <button
              onClick={() => {
                setSelectedVars([]);
                setSegmentBy('');
              }}
              className="px-3 py-2.5 text-sm text-slate-600 hover:text-slate-900 flex items-center gap-2"
            >
              <X className="w-4 h-4" />
              Limpiar
            </button>
          )}

          {/* Botón Modo Edición */}
          {hasData && (
            <div className="ml-auto flex items-center gap-2">
              {isEditing ? (
                <>
                  <button
                    onClick={handleSaveEdits}
                    className="px-4 py-2.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm font-medium flex items-center gap-2 transition-colors shadow-sm"
                  >
                    <Save className="w-4 h-4" />
                    Guardar Cambios
                  </button>
                  <button
                    onClick={handleCancelEdit}
                    className="px-4 py-2.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-sm font-medium flex items-center gap-2 transition-colors"
                  >
                    Cancelar
                  </button>
                </>
              ) : (
                <button
                  onClick={() => setIsEditing(true)}
                  className="px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-medium flex items-center gap-2 transition-colors shadow-sm"
                >
                  <Pencil className="w-4 h-4" />
                  Modo Edición
                </button>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Content */}
      <div className="flex-1 overflow-auto p-8">
        <div className="max-w-7xl mx-auto space-y-6">
          {/* Loading State */}
          {isLoading && (
            <div className="flex items-center justify-center py-12">
              <Loader2 className="w-8 h-8 text-teal-600 animate-spin" />
              <span className="ml-3 text-slate-600">Cargando frecuencias...</span>
            </div>
          )}

          {/* Error State */}
          {error && (
            <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-red-700">
              {error}
            </div>
          )}

          {/* Empty State */}
          {!isLoading && !error && selectedVars.length === 0 && (
            <div className="bg-white rounded-xl shadow-md border border-slate-200 p-12 text-center">
              <div className="text-slate-400 mb-4">
                <ChevronDown className="w-12 h-12 mx-auto" />
              </div>
              <h3 className="text-lg font-medium text-slate-700 mb-2">Selecciona variables para analizar</h3>
              <p className="text-slate-500">Elige una o más columnas del dropdown para ver sus tablas de frecuencia</p>
            </div>
          )}

          {/* Panel de Análisis IA - Diseño Mejorado */}
          {(analysisResult || isAnalyzing) && (
            <div className="mb-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <div className="bg-white rounded-xl shadow-[0_4px_20px_-4px_rgba(0,0,0,0.1)] border border-indigo-100 overflow-hidden relative">
                {/* Barra decorativa superior */}
                <div className="h-1 w-full bg-gradient-to-r from-indigo-500 via-purple-500 to-pink-500"></div>

                {/* Fondo decorativo sutil */}
                <div className="absolute top-0 right-0 -mt-10 -mr-10 w-40 h-40 bg-indigo-50 rounded-full blur-3xl opacity-50 pointer-events-none"></div>

                <div className="p-6 relative z-10">
                  <div className="flex items-start gap-5">
                    {/* Icono con halo */}
                    <div className="flex-shrink-0">
                      <div className={`w-12 h-12 rounded-xl flex items-center justify-center shadow-sm border border-indigo-50 ${isAnalyzing ? 'bg-indigo-50' : 'bg-white'
                        }`}>
                        {isAnalyzing ? (
                          <Loader2 className="w-6 h-6 text-indigo-600 animate-spin" />
                        ) : (
                          <Sparkles className="w-6 h-6 text-indigo-600 fill-indigo-50" />
                        )}
                      </div>
                    </div>

                    {/* Contenido */}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center justify-between mb-3">
                        <div>
                          <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                            {isAnalyzing ? 'Analizando distribución...' : 'Insights Estadísticos'}
                            {!isAnalyzing && <span className="px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 text-[10px] font-bold uppercase tracking-wider border border-indigo-100">AI Analysis</span>}
                          </h3>
                          <p className="text-xs text-slate-500 mt-0.5">
                            {isAnalyzing ? 'Procesando frecuencias y patrones...' : `Análisis generado para: ${selectedVars.join(', ')}`}
                          </p>
                        </div>
                        {!isAnalyzing && (
                          <button
                            onClick={() => setAnalysisResult(null)}
                            className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-all"
                            title="Cerrar análisis"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}
                      </div>

                      <div className="text-slate-600 text-sm leading-relaxed bg-slate-50/50 rounded-lg p-4 border border-slate-100">
                        {isAnalyzing ? (
                          <div className="space-y-3 py-2">
                            <div className="h-2 bg-indigo-100 rounded w-3/4 animate-pulse"></div>
                            <div className="h-2 bg-indigo-100 rounded w-full animate-pulse"></div>
                            <div className="h-2 bg-indigo-100 rounded w-5/6 animate-pulse"></div>
                          </div>
                        ) : (
                          <div className="prose prose-sm prose-indigo max-w-none">
                            {analysisResult!.split('\n').map((line, i) => (
                              <p key={i} className={`min-h-[1em] ${line.trim() === '' ? 'h-2' : 'mb-2 last:mb-0'}`}>
                                {line.split('**').map((part, index) =>
                                  index % 2 === 1 ? <strong key={index} className="text-indigo-900 font-semibold">{part}</strong> : part
                                )}
                              </p>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Segment Tabs */}
          {!isLoading && !error && segments.length > 1 && (
            <div className="bg-white rounded-lg border border-slate-200 shadow-sm">
              <div className="flex items-center gap-1 px-4 border-b border-slate-200 overflow-x-auto">
                <span className="text-xs text-slate-500 mr-2 py-3 whitespace-nowrap">
                  Segmentado por: <span className="font-medium text-slate-700">{segmentBy}</span>
                </span>
                {segments.map((segment) => (
                  <button
                    key={segment}
                    onClick={() => setActiveSegment(segment)}
                    className={`px-4 py-3 text-sm font-medium transition-colors relative whitespace-nowrap ${activeSegment === segment
                      ? 'text-teal-700'
                      : 'text-slate-600 hover:text-slate-900'
                      }`}
                  >
                    {segment}
                    {activeSegment === segment && (
                      <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-teal-600"></div>
                    )}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Frequency Tables Grid */}
          {!isLoading && !error && currentTables.length > 0 && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {currentTables.map((table: FrequencyTableResult) => (
                <FrequencyTable
                  key={`${activeSegment}-${table.variable}`}
                  title={table.variable}
                  data={table.rows}
                  total={table.total}
                  variableName={table.variable}
                  segment={activeSegment}
                  editable={isEditing}
                  onEdit={(rowIndex, field, value) => handleCellEdit(activeSegment, table.variable, rowIndex, field, value)}
                />
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Action Toolbar - Fixed at bottom */}
      {hasData && (
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