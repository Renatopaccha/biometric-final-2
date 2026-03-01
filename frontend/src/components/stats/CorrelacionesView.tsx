import { ArrowLeft, ChevronDown, Plus, Trash2, Filter, ChevronRight, Play, Loader2, X, Sparkles } from 'lucide-react';
import { useState, useRef, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx-js-style';
import { ActionToolbar } from './ActionToolbar';
import { useCorrelations, CorrelationResponse, FilterRule } from '../../hooks/useCorrelations';
import { useDataContext } from '../../context/DataContext';
import { sendChatMessage } from '../../api/ai';

interface CorrelacionesViewProps {
  onBack: () => void;
  onNavigateToChat?: (chatId?: string) => void;
}

// Academic-style cell background - softer, educational look
const getCellBackground = (value: number | null, isDiagonal: boolean): string => {
  if (isDiagonal) {
    return 'bg-blue-50/80'; // Soft blue highlight for diagonal
  }
  if (value === null) {
    return 'bg-white';
  }
  const absVal = Math.abs(value);
  if (absVal >= 0.7) {
    return 'bg-blue-50/60'; // Slight highlight for strong correlations
  }
  return 'bg-white';
};

// Get coefficient text color based on strength
const getCoefColor = (value: number | null, isDiagonal: boolean): string => {
  if (isDiagonal) {
    return 'text-blue-600'; // Blue for diagonal
  }
  if (value === null) {
    return 'text-slate-400';
  }
  const absVal = Math.abs(value);
  if (absVal >= 0.7) {
    return 'text-blue-600 font-bold'; // Strong = bold blue
  }
  if (absVal >= 0.3) {
    return 'text-slate-800'; // Moderate = dark text
  }
  return 'text-slate-500'; // Weak = lighter text
};

type CorrelationMethod = 'comparar_todos' | 'pearson' | 'spearman' | 'kendall';

// Excel styles for correlation matrix - academic look
const excelStyles = {
  header: {
    fill: { fgColor: { rgb: "F2F2F2" } },
    font: { bold: true, color: { rgb: "444746" }, sz: 11 },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "medium", color: { rgb: "9CA3AF" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  diagonal: {
    fill: { fgColor: { rgb: "EBF8FF" } }, // Light blue for diagonal
    font: { bold: true, color: { rgb: "2563EB" }, sz: 10 },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  strong: {
    fill: { fgColor: { rgb: "DBEAFE" } }, // Soft blue for strong correlations
    font: { bold: true, color: { rgb: "1D4ED8" }, sz: 10 },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  cell: {
    font: { sz: 10, color: { rgb: "374151" } },
    border: {
      top: { style: "thin", color: { rgb: "E5E7EB" } },
      bottom: { style: "thin", color: { rgb: "E5E7EB" } },
      left: { style: "thin", color: { rgb: "E5E7EB" } },
      right: { style: "thin", color: { rgb: "E5E7EB" } }
    },
    alignment: { horizontal: "center", vertical: "center" }
  },
  variableLabel: {
    fill: { fgColor: { rgb: "F8FAFC" } },
    font: { bold: true, color: { rgb: "1E293B" }, sz: 10 },
    border: {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    },
    alignment: { horizontal: "left", vertical: "center" }
  }
};

export function CorrelacionesView({ onBack, onNavigateToChat }: CorrelacionesViewProps) {
  // Context and hooks
  const { sessionId, columns: availableColumns } = useDataContext();
  const { correlationData, calculateCorrelations, loading, error } = useCorrelations();

  // AI Interpretation State
  const [analysisResult, setAnalysisResult] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [activeChatId, setActiveChatId] = useState<string | null>(null);
  const [isNavigating, setIsNavigating] = useState(false);

  // UI state
  const [selectedVars, setSelectedVars] = useState<string[]>([]);
  const [method, setMethod] = useState<CorrelationMethod>('pearson');
  const [segmentBy, setSegmentBy] = useState<string>('');
  const [activeSegmentTab, setActiveSegmentTab] = useState<string>('General');

  // Dropdown states
  const [showVarsDropdown, setShowVarsDropdown] = useState(false);
  const [showSegmentDropdown, setShowSegmentDropdown] = useState(false);

  const varsRef = useRef<HTMLDivElement>(null);
  const segmentRef = useRef<HTMLDivElement>(null);

  // Create stable key for selectedVars to avoid unnecessary re-renders
  const selectedVarsKey = useMemo(() => selectedVars.join(','), [selectedVars]);

  // Effect to trigger calculation automatically with debounce
  useEffect(() => {
    // 1. Avoid firing if not enough variables
    if (selectedVars.length < 2 || !sessionId) {
      return;
    }

    // 2. Manual debounce (500ms)
    const timer = setTimeout(() => {
      const methods = method === 'comparar_todos'
        ? ['pearson', 'spearman', 'kendall']
        : [method];

      calculateCorrelations(
        sessionId,
        selectedVars,
        methods,
        segmentBy || null,
        [], // Current filters logic removed for brevity in this fix
        'AND'
      );
    }, 500);

    return () => clearTimeout(timer);
  }, [selectedVarsKey, method, segmentBy, sessionId, calculateCorrelations]);

  // Auto-update activeSegmentTab when correlationData changes
  useEffect(() => {
    if (correlationData?.segments && correlationData.segments.length > 0) {
      setActiveSegmentTab(correlationData.segments[0]);
    } else {
      setActiveSegmentTab('General');
    }
  }, [correlationData]);

  // Click outside handlers
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

  const toggleVariable = (variable: string) => {
    setSelectedVars(prev =>
      prev.includes(variable) ? prev.filter(v => v !== variable) : [...prev, variable]
    );
  };

  // Helper function to render table rows
  const renderTableRows = (methodOverride?: 'pearson' | 'spearman' | 'kendall') => {
    if (!correlationData) return null;

    const currentMethod = methodOverride
      ? methodOverride
      : (method === 'comparar_todos' ? 'pearson' : method as 'pearson' | 'spearman' | 'kendall');

    // Fallback safely if segment doesn't exist in data
    const availableSegments = correlationData?.segments || ['General'];
    const currentSegment = availableSegments.includes(activeSegmentTab) ? activeSegmentTab : availableSegments[0] || 'General';

    const matrixData = correlationData?.tables?.[currentSegment]?.[currentMethod]?.matrix;

    if (!matrixData) {
      return <tr><td colSpan={selectedVars.length + 2} className="text-center p-4">Sin datos</td></tr>;
    }

    return selectedVars.map((rowVar) => (
      <tr key={`${currentSegment}-${currentMethod}-${rowVar}`}>
        <td className="py-4 px-4 font-medium border-b">{rowVar}</td>
        <td className="py-4 px-4 border-b bg-gray-50 text-xs text-gray-500">
          <div>Coef</div><div>p-val</div><div>N</div>
        </td>
        {selectedVars.map((colVar) => {
          const pairData = matrixData[rowVar]?.[colVar];
          const isDiagonal = rowVar === colVar;

          if (!pairData) return <td key={colVar} className="border-b text-center">-</td>;

          const cellBg = getCellBackground(pairData.r, isDiagonal);
          const coefColor = getCoefColor(pairData.r, isDiagonal);

          return (
            <td key={colVar} className={`py-4 px-4 text-center border-b ${cellBg}`}>
              <div className={`font-mono ${coefColor}`}>{pairData.r?.toFixed(3) ?? '-'}</div>
              <div className="text-xs text-slate-500">{pairData.p_value?.toFixed(3) ?? '-'}</div>
              <div className="text-xs text-slate-400">N={pairData.n ?? '-'}</div>
            </td>
          );
        })}
      </tr>
    ));
  };


  // ========================================================================
  // EXPORT FUNCTIONS (CRÍTICO: AoA Implementation)
  // ========================================================================

  const handleExportExcel = () => {
    try {
      console.log('Iniciando exportación Correlaciones (AoA)...');

      if (!correlationData || selectedVars.length < 2) {
        alert('No hay datos suficientes para exportar (mínimo 2 variables).');
        return;
      }

      if (!XLSX || !XLSX.utils) {
        throw new Error('Librería XLSX no cargada');
      }

      const wb = XLSX.utils.book_new();
      let hasData = false;

      // Métodos a exportar
      const methodsToExport: Array<'pearson' | 'spearman' | 'kendall'> =
        method === 'comparar_todos'
          ? ['pearson', 'spearman', 'kendall']
          : [method as 'pearson' | 'spearman' | 'kendall'];

      // Segmentos a exportar: exportamos SOLO el activo para no saturar,
      // o todos si el usuario prefiere. Por simplicidad de la vista actual (tabs),
      // exportaremos TODOS los segmentos disponibles.
      const segmentsToExport = correlationData.segments && correlationData.segments.length > 0
        ? correlationData.segments
        : ['General'];

      segmentsToExport.forEach(segment => {
        methodsToExport.forEach(m => {
          const matrixData = correlationData.tables?.[segment]?.[m]?.matrix;
          if (!matrixData) return;

          // Crear AoA
          const aoa: any[][] = [];

          // Header Info
          aoa.push([`Matriz de Correlación (${m})`, `Segmento: ${segment}`]);
          aoa.push([]); // Espacio

          // Headers de columna
          const headers = ['Variable', ...selectedVars];
          aoa.push(headers);

          // Filas de datos
          selectedVars.forEach(rowVar => {
            const row = [rowVar];
            selectedVars.forEach(colVar => {
              const cell = matrixData[rowVar]?.[colVar];
              if (rowVar === colVar) {
                row.push(1.0); // Diagonal
              } else if (cell && cell.r !== null) {
                row.push(cell.r);
              } else {
                row.push('');
              }
            });
            aoa.push(row);
          });

          // Crear hoja
          const ws = XLSX.utils.aoa_to_sheet(aoa);

          // Estilos Condicionales
          // Datos empiezan en fila 3 (índice 3, pues row 0=Title, row 1=Empty, row 2=Headers)
          const startRow = 3;
          const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

          // Estilo Header
          for (let C = 0; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({ r: 2, c: C }); // Fila headers
            if (ws[addr]) ws[addr].s = excelStyles.header;
          }

          // Estilos Celdas (Diagonal y Negrita > 0.7)
          for (let R = startRow; R <= range.e.r; ++R) {
            // Col 0 es nombre variable, estilo Label
            const labelAddr = XLSX.utils.encode_cell({ r: R, c: 0 });
            if (ws[labelAddr]) ws[labelAddr].s = excelStyles.variableLabel;

            for (let C = 1; C <= range.e.c; ++C) {
              const addr = XLSX.utils.encode_cell({ r: R, c: C });
              if (!ws[addr]) continue;

              const val = ws[addr].v;

              // Diagonal? (Si R-startRow == C-1) -> R = C + 2
              // idxRow = R - 3. idxCol = C - 1. Diagonal si idxRow == idxCol
              // R - 3 = C - 1 => R = C + 2
              const isDiagonal = (R - startRow) === (C - 1);

              if (isDiagonal) {
                ws[addr].s = excelStyles.diagonal;
              } else if (typeof val === 'number') {
                ws[addr].s = Math.abs(val) >= 0.7 ? excelStyles.strong : excelStyles.cell;
                ws[addr].t = 'n';
                ws[addr].z = '0.000'; // Formato numérico
              }
            }
          }

          // Anchos
          ws['!cols'] = [{ wch: 20 }, ...Array(selectedVars.length).fill({ wch: 12 })];

          // Append sheet
          // Nombre hoja: Segmento - Método (truncado)
          const sheetName = `${segment.substring(0, 10)}_${m.substring(0, 10)}`;
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
          hasData = true;
        });
      });

      if (!hasData) {
        alert('No hay datos para exportar.');
        return;
      }

      const fileName = `Correlaciones_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, fileName);
      console.log('Exportación Correlaciones completada.');

    } catch (error) {
      console.error('Error exporting Correlations:', error);
      alert(`Error: ${error instanceof Error ? error.message : String(error)}`);
    }
  };

  // Exportar a PDF (Placeholder simple)
  const handleExportPDF = async () => {
    alert('Funcionalidad PDF simplificada para esta vista. Use Excel para mejor detalle.');
  };

  // AI Handlers
  const handleAIInterpretation = async () => {
    if (!selectedVars.length) return;
    setIsAnalyzing(true);
    try {
      const prompt = `Analiza correlaciones entre: ${selectedVars.join(', ')}.`;
      const resp = await sendChatMessage({ session_id: sessionId, message: prompt, history: [] });
      if (resp.success) setAnalysisResult(resp.response);
    } catch (e) { console.error(e); }
    finally { setIsAnalyzing(false); }
  };

  const handleContinueToChat = async () => {
    if (onNavigateToChat) onNavigateToChat(activeChatId || undefined);
  };

  return (
    <div className="h-full flex flex-col bg-slate-50">
      <div className="bg-white border-b p-4 shadow-sm flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button onClick={onBack}><ArrowLeft /></button>
          <h2 className="text-2xl font-bold">Correlaciones</h2>
        </div>
        {/* Controls Simplified */}
        <div className="flex gap-2">
          <div className="relative" ref={varsRef}>
            <button onClick={() => setShowVarsDropdown(!showVarsDropdown)} className="px-4 py-2 border rounded bg-white">
              Variables ({selectedVars.length}) <ChevronDown className="inline w-4" />
            </button>
            {showVarsDropdown && (
              <div className="absolute top-full right-0 bg-white border shadow-lg z-10 w-64 max-h-60 overflow-auto p-2">
                {availableColumns.map(c => (
                  <label key={c} className="flex gap-2 p-2 hover:bg-gray-50 cursor-pointer">
                    <input type="checkbox" checked={selectedVars.includes(c)} onChange={() => toggleVariable(c)} />
                    {c}
                  </label>
                ))}
              </div>
            )}
          </div>
          {/* Segment */}
          <div className="relative" ref={segmentRef}>
            <button onClick={() => setShowSegmentDropdown(!showSegmentDropdown)} className="px-4 py-2 border rounded bg-white">
              Segmentar <ChevronDown className="inline w-4" />
            </button>
            {showSegmentDropdown && (
              <div className="absolute top-full right-0 bg-white border shadow-lg z-10 w-48 p-2">
                <div onClick={() => { setSegmentBy(''); setShowSegmentDropdown(false) }} className="p-2 hover:bg-gray-50 cursor-pointer">Ninguno</div>
                {availableColumns.map(c => (
                  <div key={c} onClick={() => { setSegmentBy(c); setShowSegmentDropdown(false) }} className="p-2 hover:bg-gray-50 cursor-pointer">{c}</div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="flex-1 overflow-auto p-8">
        {loading && <div className="flex justify-center"><Loader2 className="animate-spin" /></div>}
        {!loading && !error && selectedVars.length >= 2 && (
          <div className="bg-white p-6 rounded shadow overflow-auto">
            <table className="w-full border-collapse">
              <tbody>{renderTableRows()}</tbody>
            </table>
          </div>
        )}
        {!loading && selectedVars.length < 2 && (
          <div className="text-center text-gray-500 mt-20">Selecciona al menos 2 variables</div>
        )}
      </div>

      <ActionToolbar
        onExportExcel={handleExportExcel}
        onExportPDF={handleExportPDF}
        onAIInterpretation={handleAIInterpretation}
        onContinueToChat={handleContinueToChat}
        isAnalyzing={isAnalyzing}
        isNavigating={isNavigating}
      />
    </div>
  );
}
