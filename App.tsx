
import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  FileSpreadsheet, 
  Save,
  Upload, 
  Trash2, 
  Printer,
  Search,
  X,
  Plus,
  ArrowDown,
  Eraser,
  Check,
  Download,
  FileText,
  ChevronDown,
  Calculator,
  CheckSquare,
  Undo,
  Redo,
  FolderOpen,
  AlertTriangle,
  Calendar,
  Hash,
  AlignLeft,
  Clock,
  Percent,
  Minus,
  Divide,
  X as XIcon,
  Equal,
  Info
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { BudgetItem, ProjectInfo, AppState } from './types';
import { formatCurrency, formatNumber, calculateItemTotals, roundToTwo } from './utils/formatters';

// --- INITIAL DATA (CLEAN) ---
const INITIAL_PROJECT_INFO: ProjectInfo = {
  name: "",
  projectNumber: "",
  orderNumber: "",
  location: "",
  client: "",
  certificationNumber: 1,
  date: new Date().toISOString().split('T')[0],
  isAveria: false,
  averiaNumber: "",
  averiaDate: new Date().toISOString().split('T')[0],
  averiaDescription: "",
  averiaTiming: "diurna"
};

const STORAGE_KEY = 'certipro_autosave_v1';

// Selection State Interface
interface SelectionState {
    start: { r: number; c: number } | null;
    end: { r: number; c: number } | null;
    isDragging: boolean;
}

// --- DRAGGABLE CALCULATOR COMPONENT ---
const DraggableCalculator: React.FC<{ onClose: () => void }> = ({ onClose }) => {
    const [pos, setPos] = useState({ x: window.innerWidth - 340, y: 120 });
    const [isDragging, setIsDragging] = useState(false);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [display, setDisplay] = useState('0');
    const [prevOperation, setPrevOperation] = useState('');
    const [isNewNumber, setIsNewNumber] = useState(true);

    const handleMouseDown = (e: React.MouseEvent) => {
        setIsDragging(true);
        setDragOffset({
            x: e.clientX - pos.x,
            y: e.clientY - pos.y
        });
    };

    useEffect(() => {
        const handleMouseMove = (e: MouseEvent) => {
            if (isDragging) {
                setPos({
                    x: e.clientX - dragOffset.x,
                    y: e.clientY - dragOffset.y
                });
            }
        };
        const handleMouseUp = () => {
            setIsDragging(false);
        };

        if (isDragging) {
            window.addEventListener('mousemove', handleMouseMove);
            window.addEventListener('mouseup', handleMouseUp);
        }
        return () => {
            window.removeEventListener('mousemove', handleMouseMove);
            window.removeEventListener('mouseup', handleMouseUp);
        };
    }, [isDragging, dragOffset]);

    const handleNumber = (num: string) => {
        if (isNewNumber) {
            setDisplay(num);
            setIsNewNumber(false);
        } else {
            setDisplay(display === '0' ? num : display + num);
        }
    };

    const handleOperator = (op: string) => {
        setPrevOperation(display + ' ' + op + ' ');
        setIsNewNumber(true);
    };

    const calculate = () => {
        try {
            // Safe eval logic
            const fullExpr = prevOperation + display;
            // Replace visual X with *
            const sanitized = fullExpr.replace(/x/g, '*').replace(/÷/g, '/');
            // eslint-disable-next-line no-new-func
            const result = new Function('return ' + sanitized)();
            const formatted = String(parseFloat(result.toFixed(4))); // Avoid long decimals
            setDisplay(formatted);
            setPrevOperation('');
            setIsNewNumber(true);
        } catch (e) {
            setDisplay('Error');
            setIsNewNumber(true);
        }
    };

    const clear = () => {
        setDisplay('0');
        setPrevOperation('');
        setIsNewNumber(true);
    };

    const btnClass = "h-12 flex items-center justify-center rounded bg-slate-700 hover:bg-slate-600 text-white font-bold text-lg transition-colors shadow-sm active:transform active:scale-95";
    const opClass = "h-12 flex items-center justify-center rounded bg-orange-500 hover:bg-orange-600 text-white font-bold text-lg transition-colors shadow-sm";

    return (
        <div 
            className="fixed z-[200] w-80 bg-slate-800 rounded-lg shadow-2xl border border-slate-600 overflow-hidden flex flex-col"
            style={{ left: pos.x, top: pos.y }}
        >
            {/* Header / Drag Handle */}
            <div 
                className="bg-slate-900 px-4 py-2 flex items-center justify-between cursor-move select-none border-b border-slate-700"
                onMouseDown={handleMouseDown}
            >
                <div className="flex items-center gap-2 text-slate-300 font-bold text-sm">
                    <Calculator className="w-4 h-4" /> Calculadora
                </div>
                <button 
                    onClick={onClose}
                    className="text-slate-400 hover:text-white p-1 hover:bg-slate-700 rounded"
                >
                    <X className="w-4 h-4" />
                </button>
            </div>

            {/* Display */}
            <div className="p-4 bg-slate-800">
                <div className="text-slate-400 text-right text-xs h-4 mb-1 font-mono">{prevOperation}</div>
                <div className="text-white text-right text-3xl font-mono font-bold truncate tracking-widest bg-slate-900/50 p-2 rounded border border-slate-700/50">
                    {display}
                </div>
            </div>

            {/* Keypad */}
            <div className="grid grid-cols-4 gap-2 p-4 pt-0 bg-slate-800">
                <button onClick={clear} className="col-span-3 h-12 flex items-center justify-center rounded bg-slate-600 hover:bg-slate-500 text-red-200 font-bold">AC</button>
                <button onClick={() => handleOperator('÷')} className={opClass}><Divide className="w-5 h-5"/></button>

                <button onClick={() => handleNumber('7')} className={btnClass}>7</button>
                <button onClick={() => handleNumber('8')} className={btnClass}>8</button>
                <button onClick={() => handleNumber('9')} className={btnClass}>9</button>
                <button onClick={() => handleOperator('x')} className={opClass}><XIcon className="w-5 h-5"/></button>

                <button onClick={() => handleNumber('4')} className={btnClass}>4</button>
                <button onClick={() => handleNumber('5')} className={btnClass}>5</button>
                <button onClick={() => handleNumber('6')} className={btnClass}>6</button>
                <button onClick={() => handleOperator('-')} className={opClass}><Minus className="w-5 h-5"/></button>

                <button onClick={() => handleNumber('1')} className={btnClass}>1</button>
                <button onClick={() => handleNumber('2')} className={btnClass}>2</button>
                <button onClick={() => handleNumber('3')} className={btnClass}>3</button>
                <button onClick={() => handleOperator('+')} className={opClass}><Plus className="w-5 h-5"/></button>

                <button onClick={() => handleNumber('0')} className={`${btnClass} col-span-2`}>0</button>
                <button onClick={() => handleNumber('.')} className={btnClass}>.</button>
                <button onClick={calculate} className="h-12 flex items-center justify-center rounded bg-emerald-500 hover:bg-emerald-600 text-white font-bold text-lg shadow-sm"><Equal className="w-6 h-6"/></button>
            </div>
        </div>
    );
};

const App: React.FC = () => {
  // --- STATE ---
  // Lazy initialization logic for Autosave
  const [state, setState] = useState<AppState>(() => {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            const parsed = JSON.parse(saved);
            // Rehydrate Set from Array
            return {
                ...parsed,
                // Ensure checkedRowIds is a Set
                checkedRowIds: new Set(Array.isArray(parsed.checkedRowIds) ? parsed.checkedRowIds : []),
                isLoading: false // Always reset loading
            };
        }
    } catch (e) {
        console.error("Failed to load autosave", e);
    }
    // Fallback if no save found
    return {
        masterItems: [],
        items: [],
        projectInfo: INITIAL_PROJECT_INFO,
        isLoading: false,
        checkedRowIds: new Set()
    };
  });
  
  // --- AUTOSAVE EFFECT ---
  useEffect(() => {
     if (!state) return;
     try {
         const stateToSave = {
             ...state,
             // Serialize Set to Array for JSON storage
             checkedRowIds: Array.from(state.checkedRowIds || new Set()),
             isLoading: false
         };
         localStorage.setItem(STORAGE_KEY, JSON.stringify(stateToSave));
     } catch (e) {
         console.error("Failed to autosave", e);
     }
  }, [state]);

  // History State
  const [history, setHistory] = useState<{ past: AppState[], future: AppState[] }>({
      past: [],
      future: []
  });

  // Ref to snapshot state before edits (for undoing text input changes)
  const historySnapshot = useRef<AppState | null>(null);
  
  // Refs
  const dropdownRef = useRef<HTMLDivElement>(null);
  const exportBtnRef = useRef<HTMLDivElement>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const marginInputRef = useRef<HTMLInputElement>(null);
  // Ref for auto-scroll speed during drag (negative = up, positive = down)
  const autoScrollSpeed = useRef<number>(0);
  
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showCalculator, setShowCalculator] = useState(false);
  
  // Dialog States
  const [showProformaDialog, setShowProformaDialog] = useState(false);
  const [proformaMargin, setProformaMargin] = useState("0");
  const [showClearDialog, setShowClearDialog] = useState(false);

  // Excel-like Drag Selection State
  const [selection, setSelection] = useState<SelectionState>({
      start: null,
      end: null,
      isDragging: false
  });

  // Row Reordering Drag State
  const [draggedRowIndex, setDraggedRowIndex] = useState<number | null>(null);

  // Inline Search State
  const [activeSearch, setActiveSearch] = useState<{ rowId: string, field: 'code' | 'description' } | null>(null);
  // Dropdown placement state: 'bottom' (default) or 'top' (if close to screen bottom)
  const [dropdownPlacement, setDropdownPlacement] = useState<'bottom' | 'top'>('bottom');
  
  // Editing Cell State for View/Edit mode switching (Price column)
  const [editingCell, setEditingCell] = useState<{ rowId: string, field: string } | null>(null);

  // Context Menu State
  const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; rowId: string | null }>({
    visible: false,
    x: 0,
    y: 0,
    rowId: null
  });

  // --- AUTO-SCROLL LOGIC FOR DRAGGING ROWS ---
  useEffect(() => {
    let animationFrameId: number;

    const scrollStep = () => {
        if (autoScrollSpeed.current !== 0 && tableContainerRef.current) {
            tableContainerRef.current.scrollTop += autoScrollSpeed.current;
        }
        animationFrameId = requestAnimationFrame(scrollStep);
    };

    animationFrameId = requestAnimationFrame(scrollStep);
    return () => cancelAnimationFrame(animationFrameId);
  }, []);

  // --- AUTO-SCROLL FOR DROPDOWN VISIBILITY ---
  useEffect(() => {
    // Only run if we have an active search AND the placement is 'bottom'.
    // If it's 'top', we don't need to scroll down.
    if (activeSearch && dropdownRef.current && tableContainerRef.current && dropdownPlacement === 'bottom') {
        const timer = setTimeout(() => {
            if (!dropdownRef.current || !tableContainerRef.current) return;

            const dropdownRect = dropdownRef.current.getBoundingClientRect();
            const containerRect = tableContainerRef.current.getBoundingClientRect();

            // Check distance from bottom of dropdown to bottom of container visible area
            const hiddenAmount = dropdownRect.bottom - containerRect.bottom;

            // If hidden (positive amount), scroll down
            if (hiddenAmount > 0) {
                tableContainerRef.current.scrollBy({
                    top: hiddenAmount + 24, // Scroll enough to show + 24px padding
                    behavior: 'smooth'
                });
            }
        }, 50); // Small delay to ensure DOM paint
        return () => clearTimeout(timer);
    }
  }, [activeSearch, state.items, dropdownPlacement]); 

  // Focus margin input when dialog opens
  useEffect(() => {
    if (showProformaDialog && marginInputRef.current) {
        setTimeout(() => marginInputRef.current?.select(), 50);
    }
  }, [showProformaDialog]);

  // --- HISTORY LOGIC ---

  // Improved saveHistory that accepts the state to save explicitly
  // and prevents adding duplicate states to history
  const saveHistory = (currentState: AppState) => {
      if (!currentState) return;
      
      setHistory(prev => {
          // Prevent saving if the state hasn't actually changed reference (basic duplication check)
          // or if it matches the very last entry in history
          const lastPastState = prev.past[prev.past.length - 1];
          if (lastPastState === currentState) {
              return prev;
          }

          return {
              past: [...prev.past, currentState],
              future: []
          };
      });
  };

  const undo = () => {
      // 1. Handle uncommitted edit (e.g. user clicks Undo while input is focused and typing)
      if (historySnapshot.current) {
          // If the snapshot is different from current state (meaning we typed something),
          // we treat this as the first "undo" step: reverting the text field.
          if (historySnapshot.current !== state) {
              const snapshot = historySnapshot.current;
              const dirtyState = state;
              
              setHistory(prev => ({
                  past: prev.past, 
                  future: [dirtyState, ...prev.future] 
              }));
              setState(snapshot);
              
              // Keep snapshot reference active because we are effectively still "in" the previous state context
              historySnapshot.current = snapshot;
              return;
          }
      }

      // 2. Standard Undo
      if (history.past.length === 0) return;
      const previous = history.past[history.past.length - 1];
      
      if (!previous) return;

      setHistory(prev => ({
          past: prev.past.slice(0, -1),
          future: [state, ...prev.future]
      }));
      setState(previous);

      // If we are focused on an input, update the snapshot to reflect the state we just landed on
      // This allows typing immediately after undoing without breaking logic
      if (historySnapshot.current) {
          historySnapshot.current = previous;
      }
  };

  const redo = () => {
      if (history.future.length === 0) return;
      const next = history.future[0];

      if (!next) return;
      
      setHistory(prev => ({
          past: [...prev.past, state],
          future: prev.future.slice(1)
      }));
      setState(next);

      if (historySnapshot.current) {
          historySnapshot.current = next;
      }
  };

  // Handlers for Continuous Inputs (Text/Number fields)
  // Saves history only when the user finishes editing (onBlur) and if value actually changed.
  const handleInputFocus = () => {
      if (state) {
        historySnapshot.current = state;
      }
  };

  const handleInputBlur = () => {
      // Only save if there was a change and we have a valid snapshot
      if (historySnapshot.current && state && historySnapshot.current !== state) {
           // We save the SNAPSHOT (clean state) to history, so undo brings us back to before typing
           saveHistory(historySnapshot.current);
      }
      historySnapshot.current = null;
  };

  // Keyboard Shortcuts for Undo/Redo
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
        if ((e.metaKey || e.ctrlKey) && e.key === 'z') {
            e.preventDefault();
            if (e.shiftKey) redo();
            else undo();
        }
        if ((e.metaKey || e.ctrlKey) && e.key === 'y') {
             e.preventDefault();
             redo();
        }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [history, state]);

  // --- PERSISTENCE (SAVE/LOAD) LOGIC ---
  const handleSaveProject = () => {
      if (!state) return;
      
      const backupData = {
          version: "1.0",
          timestamp: new Date().toISOString(),
          projectInfo: state.projectInfo,
          items: state.items,
          masterItems: state.masterItems,
          // Convert Set to Array for JSON serialization (safely)
          checkedRowIds: Array.from(state.checkedRowIds || []) 
      };

      const jsonString = JSON.stringify(backupData, null, 2);
      const blob = new Blob([jsonString], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `Copia_${state.projectInfo.projectNumber || 'Obra'}_${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
  };

  const handleLoadProject = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (event) => {
          try {
              const jsonString = event.target?.result as string;
              if (!jsonString) {
                  alert("No se pudo leer el archivo (archivo vacío).");
                  return;
              }

              let data;
              try {
                  data = JSON.parse(jsonString);
              } catch (parseError) {
                  alert("El archivo seleccionado no es un archivo JSON válido. \n\nSi intenta cargar una tabla de Excel, utilice el botón 'Importar Tabla Rec.'");
                  return;
              }

              // Basic validation
              if (!data || typeof data !== 'object') {
                  alert("Formato de archivo corrupto.");
                  return;
              }

              if (!data.items && !data.projectInfo) {
                   alert("El archivo no contiene datos reconocibles de CertiPro.");
                   return;
              }

              // Save current state to history before overwriting (if we have a valid previous state)
              if (state && state.items.length > 0) {
                  saveHistory(state);
              }

              setState({
                  projectInfo: data.projectInfo || INITIAL_PROJECT_INFO,
                  items: Array.isArray(data.items) ? data.items : [],
                  masterItems: Array.isArray(data.masterItems) ? data.masterItems : (state.masterItems || []), 
                  isLoading: false,
                  // Convert Array back to Set safely
                  checkedRowIds: new Set(Array.isArray(data.checkedRowIds) ? data.checkedRowIds : [])
              });

          } catch (error) {
              console.error(error);
              alert("Ocurrió un error inesperado al cargar el proyecto.");
          }
      };
      reader.readAsText(file);
  };

  // --- EXCEL IMPORT LOGIC ---
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Capture state for history before loading starts
    const currentState = state;
    if (!currentState) return;

    setState(prev => prev ? ({ ...prev, isLoading: true }) : prev);
    const reader = new FileReader();
    
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        if (rawData.length < 2) {
            alert("El archivo Excel parece vacío.");
            setState(prev => prev ? ({ ...prev, isLoading: false }) : prev);
            return;
        }

        let headerRowIndex = 0;
        let headers = rawData[0].map(h => String(h).toLowerCase());
        
        const findColIndex = (variants: string[]) => headers.findIndex(h => variants.some(v => h.includes(v)));

        let idxCode = findColIndex(['recurso', 'código', 'codigo', 'id', 'partida', 'code']);
        let idxDesc = findColIndex(['denominación', 'denominacion', 'descripción', 'descripcion', 'description', 'nombre']);
        let idxUnit = findColIndex(['unidad', 'unid', 'unit', 'medida']);
        let idxPrice = findColIndex(['precio', 'precio unitario', 'unit price', 'p.u.', 'price']);

        if (idxCode === -1 && idxDesc === -1) {
            idxCode = 0; idxDesc = 1; idxUnit = 2; idxPrice = 4;
        }

        const mappedItems: BudgetItem[] = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            const code = row[idxCode] ? String(row[idxCode]) : `R-${i}`;
            if ((!row[idxCode] && !row[idxDesc])) continue;

            const desc = row[idxDesc] ? String(row[idxDesc]) : 'Sin denominación';
            const unit = row[idxUnit] ? String(row[idxUnit]) : 'ud';
            const price = row[idxPrice] ? parseFloat(String(row[idxPrice]).replace(',', '.')) : 0;

            mappedItems.push({
                id: `master-${i}-${Date.now()}`,
                code,
                description: desc,
                unit,
                plannedQuantity: 0,
                unitPrice: isNaN(price) ? 0 : roundToTwo(price),
                previousQuantity: 0,
                currentQuantity: 0,
                totalQuantity: 0,
                totalAmount: 0,
                observations: '',
                kFactor: 1
            });
        }
        
        // Save history explicitly using our safe helper
        saveHistory(currentState);

        setState(prev => prev ? ({ 
          ...prev, 
          masterItems: mappedItems, // Update only the Master Items
          // We DO NOT reset items or checkedRowIds here anymore
          isLoading: false,
        }) : prev);
      } catch (err) {
        console.error(err);
        alert("Error al leer Excel.");
        setState(prev => prev ? ({ ...prev, isLoading: false }) : prev);
      }
    };
    reader.readAsBinaryString(file);
  };

  // --- CHECKBOX LOGIC (INTEGRATED INTO HISTORY) ---
  const toggleRowCheck = (id: string) => {
      saveHistory(state); // Save before change
      setState(prev => {
          if (!prev) return prev;
          const newSet = new Set(prev.checkedRowIds);
          if (newSet.has(id)) {
              newSet.delete(id);
          } else {
              newSet.add(id);
          }
          return { ...prev, checkedRowIds: newSet };
      });
  };

  const toggleAllChecks = () => {
      saveHistory(state); // Save before change
      setState(prev => {
          if (!prev) return prev;
          const newSet = new Set<string>();
          // If not all are checked, check all. If all checked, uncheck all (empty set).
          if (prev.checkedRowIds.size !== prev.items.length) {
              prev.items.forEach(i => newSet.add(i.id));
          }
          return { ...prev, checkedRowIds: newSet };
      });
  };

  // --- CLEAR ALL LOGIC ---
  const handleClearAll = () => {
    // Enable clear if items exist OR if general info exists OR if averia data is present
    const hasItems = state.items.length > 0;
    const hasAveria = state.projectInfo.isAveria;
    const hasInfo = !!(state.projectInfo.name || state.projectInfo.projectNumber || state.projectInfo.orderNumber || state.projectInfo.client);
    
    if (!hasItems && !hasAveria && !hasInfo) return;
    setShowClearDialog(true);
  };

  const confirmClearAll = () => {
    // Reset History completely to 0
    setHistory({ past: [], future: [] });
    // Also clear input snapshot to be safe
    historySnapshot.current = null;

    setState(prev => ({
        ...prev,
        items: [],
        checkedRowIds: new Set(),
        projectInfo: {
            ...INITIAL_PROJECT_INFO,
            // Ensure date is fresh
            date: new Date().toISOString().split('T')[0],
            averiaDate: new Date().toISOString().split('T')[0]
        }
    }));
    setSelectedRowId(null);
    setActiveSearch(null);
    setShowClearDialog(false);
  };

  // --- EXPORT LOGIC ---
  const exportToExcel = () => {
    if (!state) return;
    const { projectInfo, items } = state;
    const isAveria = projectInfo.isAveria;
    // Filter out rows that are effectively empty (no code/desc) to avoid exporting blank lines
    const validItems = items.filter(item => item.code.trim() !== '' || item.description.trim() !== '');

    // 1. Header Information (Project Details)
    const headerRows = [
        [isAveria ? "CERTIFICACIÓN DE OBRA (AVERÍA)" : "CERTIFICACIÓN DE OBRA"],
        [""],
        ["Denominación:", projectInfo.name],
        ["Nº Obra:", projectInfo.projectNumber],
        ["Nº Pedido:", projectInfo.orderNumber],
        ["Cliente:", projectInfo.client],
        ["Fecha:", projectInfo.date],
    ];

    // Inject Averia details if enabled
    if (isAveria) {
        headerRows.push(
            [""],
            ["DETALLES DE AVERÍA"],
            ["Nº Avería:", projectInfo.averiaNumber || ''],
            ["Fecha Avería:", projectInfo.averiaDate || ''],
            ["Horario:", projectInfo.averiaTiming === 'nocturna_finde' ? 'Avería nocturna/ Fin de Semana K=1,75' : 'Avería diurna K=1,25'],
            ["Descripción:", projectInfo.averiaDescription || '']
        );
    }

    headerRows.push(
        [""], // Spacer
        // Table Headers
        isAveria 
          ? ["Recurso", "Descripción", "Ud", "K", "Precio Unitario", "Importe Total", "Observaciones"]
          : ["Recurso", "Descripción", "Ud", "Precio Unitario", "Importe Total", "Observaciones"]
    );

    // 2. Data Rows
    const dataRows = validItems.map(item => {
        const k = item.kFactor || 1;
        const total = roundToTwo(item.currentQuantity * (isAveria ? k : 1) * item.unitPrice);
        
        if (isAveria) {
            return [
                item.code,
                item.description,
                item.currentQuantity,
                k, // New K Column
                roundToTwo(item.unitPrice),
                total,
                item.observations || ''
            ];
        } else {
            return [
                item.code,
                item.description,
                item.currentQuantity,
                roundToTwo(item.unitPrice),
                total,
                item.observations || ''
            ];
        }
    });

    // 3. Totals Row
    const totalAmountExport = validItems.reduce((acc, curr) => {
        const k = isAveria ? (curr.kFactor || 1) : 1;
        return acc + roundToTwo(curr.currentQuantity * k * curr.unitPrice);
    }, 0);
    
    const totalRow = isAveria 
        ? ["", "TOTAL CERTIFICACIÓN", "", "", "", roundToTwo(totalAmountExport), ""]
        : ["", "TOTAL CERTIFICACIÓN", "", "", roundToTwo(totalAmountExport), ""];

    // 4. Combine
    const finalData = [...headerRows, ...dataRows, totalRow];

    // 5. Create Sheet
    const ws = XLSX.utils.aoa_to_sheet(finalData);

    // 6. Set Column Widths (Adjusted for request)
    if (isAveria) {
        ws['!cols'] = [
            { wch: 30 }, // Recurso
            { wch: 50 }, // Descripción
            { wch: 10 }, // Ud
            { wch: 5 },  // K
            { wch: 15 }, // Precio
            { wch: 15 }, // Importe
            { wch: 30 }  // Observaciones
        ];
    } else {
        ws['!cols'] = [
            { wch: 30 }, // Recurso
            { wch: 50 }, // Descripción
            { wch: 10 }, // Ud
            { wch: 15 }, // Precio
            { wch: 15 }, // Importe
            { wch: 30 }  // Observaciones
        ];
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Certificación");
    XLSX.writeFile(wb, `Certificacion_${projectInfo.projectNumber || 'Obra'}.xlsx`);
    setShowExportMenu(false);
  };

  const handleProformaClick = () => {
    // Only allow proforma if checked items exist
    const selectedCount = state.checkedRowIds.size;
    if (selectedCount === 0) {
        alert("Por favor, marque al menos una casilla (tick) para generar la Proforma.");
        return;
    }
    
    // Close export menu and open proforma dialog
    setShowExportMenu(false);
    setProformaMargin("0"); // Reset to 0
    setShowProformaDialog(true);
  };

  const confirmProformaExport = () => {
      const margin = parseFloat(proformaMargin.replace(',', '.'));
      const finalMargin = isNaN(margin) ? 0 : margin;
      
      downloadPDF('landscape', 'proforma', finalMargin);
      setShowProformaDialog(false);
  };

  const downloadPDF = (orientation: 'portrait' | 'landscape', type: 'certification' | 'proforma', marginPercentage: number = 0) => {
    // 1. Filter items based on type
    const items = type === 'proforma' 
      ? state.items.filter(item => state.checkedRowIds.has(item.id)) 
      : state.items.filter(item => item.code.trim() !== '' || item.description.trim() !== '');

    if (type === 'proforma' && items.length === 0) {
        alert("Por favor, marque al menos una casilla (tick) para generar la Proforma.");
        return;
    }

    // --- MARGIN CALCULATION LOGIC ---
    let marginDivisor = 1;
    if (type === 'proforma' && marginPercentage !== 0) {
        marginDivisor = 1 + (marginPercentage / 100);
    }
    // --------------------------------

    const { projectInfo } = state;
    const isAveria = projectInfo.isAveria;

    // 2. Initialize jsPDF
    // @ts-ignore
    const doc = new jsPDF({
        orientation: orientation,
        unit: 'mm',
        format: 'a4'
    });
    
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 12;

    // 3. HEADER GENERATION (Replicating the "Perfect" HTML Layout)
    
    // -- Right Side: TENSA SA Info --
    const rightMargin = pageWidth - 14;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8);
    doc.setTextColor(50); // Dark Gray
    doc.text("TENSA SA", rightMargin, 12, { align: "right" });
    
    doc.setFont("helvetica", "normal");
    doc.setTextColor(100); // Lighter Gray
    doc.text("Pol. Villamuriel de Cerrato", rightMargin, 16, { align: "right" });
    doc.text("C/España, parc.79", rightMargin, 20, { align: "right" });
    doc.text("Palencia, CP: 33429", rightMargin, 24, { align: "right" });
    doc.text("CIF: A33020074", rightMargin, 28, { align: "right" });

    doc.setFont("helvetica", "bold");
    doc.setTextColor(150);
    doc.text("FECHA", rightMargin, 36, { align: "right" });
    doc.setFontSize(12);
    doc.setTextColor(0);
    doc.text(new Date(projectInfo.date).toLocaleDateString(), rightMargin, 41, { align: "right" });

    // -- Left Side: Title & Project Info --
    const title = type === 'proforma' 
        ? (isAveria ? "FACTURA PROFORMA (AVERÍA)" : "FACTURA PROFORMA") 
        : (isAveria ? "CERTIFICACIÓN DE OBRA (AVERÍA)" : "CERTIFICACIÓN DE OBRA");

    doc.setTextColor(0); // Black
    doc.setFontSize(18);
    doc.setFont("helvetica", "bold");
    doc.text(title.toUpperCase(), margin, 15);

    doc.setFontSize(14);
    doc.text(projectInfo.name.toUpperCase(), margin, 24);

    // Info Row
    let yPos = 35;
    doc.setFontSize(10);
    
    doc.setFont("helvetica", "bold");
    doc.text("Nº Obra:", margin, yPos);
    doc.setFont("helvetica", "normal");
    doc.text(projectInfo.projectNumber, margin + 18, yPos);

    doc.setFont("helvetica", "bold");
    doc.text("Nº Pedido:", margin + 60, yPos);
    doc.setFont("helvetica", "normal");
    doc.text(projectInfo.orderNumber, margin + 82, yPos);

    // Cliente Row
    yPos += 6;
    doc.setFont("helvetica", "bold");
    doc.text("Cliente:", margin, yPos);
    doc.setFont("helvetica", "normal");
    doc.text(projectInfo.client, margin + 18, yPos);

    yPos += 5; // Spacing

    // -- AVERIA BLOCK (Conditional) --
    if (isAveria) {
        yPos += 4;
        // Top Border Line
        doc.setDrawColor(220); // Light Gray
        doc.setLineWidth(0.1);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 6;

        const redColor: [number, number, number] = [153, 27, 27]; // Dark Red

        // Row 1: Nº, Fecha, Horario
        doc.setFontSize(10);
        
        // Nº Avería
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("Nº AVERÍA:", margin, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        doc.text(projectInfo.averiaNumber || '', margin + 22, yPos);

        // Fecha Avería
        const col2X = margin + 55;
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("FECHA AVERÍA:", col2X, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        doc.text(projectInfo.averiaDate ? new Date(projectInfo.averiaDate).toLocaleDateString() : '', col2X + 28, yPos);

        // Horario
        const col3X = margin + 110;
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("HORARIO:", col3X, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        const horario = projectInfo.averiaTiming === 'nocturna_finde' ? 'Avería nocturna/ Fin de Semana K=1,75' : 'Avería diurna K=1,25';
        doc.text(horario, col3X + 20, yPos);

        yPos += 6;

        // Description Row
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("DESCRIPCIÓN:", margin, yPos);
        
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        
        // Justified text block
        const descText = projectInfo.averiaDescription || '';
        const textX = margin + 28;
        const maxTextWidth = pageWidth - textX - margin;
        
        // Use text with maxWidth + align justify
        doc.text(descText, textX, yPos, { maxWidth: maxTextWidth, align: 'justify' });
        
        // Calculate height of text block to move yPos
        const dims = doc.getTextDimensions(descText, { maxWidth: maxTextWidth });
        yPos += Math.max(dims.h, 5) + 4;
        
        // Bottom Border Line (Stronger)
        doc.setDrawColor(0);
        doc.setLineWidth(0.5);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 8;

    } else {
        // Standard line if not averia
        yPos += 4;
        doc.setDrawColor(0);
        doc.setLineWidth(0.5);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 8;
    }

    // 4. TABLE GENERATION
    
    // Prepare Data
    const tableBody = items.map(item => {
        const k = isAveria ? (item.kFactor || 1) : 1;
        // Adjust price using marginDivisor
        const adjustedUnitPrice = item.unitPrice / marginDivisor;
        const total = roundToTwo(item.currentQuantity * k * adjustedUnitPrice);
        
        return [
            item.code,
            item.description,
            formatNumber(item.currentQuantity),
            ...(isAveria ? [formatNumber(k)] : []),
            formatCurrency(adjustedUnitPrice),
            formatCurrency(total),
            item.observations || ''
        ];
    });

    const tableHead = [
        "RECURSO", 
        "DESCRIPCIÓN", 
        "UD", 
        ...(isAveria ? ["K"] : []), 
        "PRECIO", 
        "IMPORTE", 
        "OBSERVACIONES"
    ];

    // Variable to track the end X position of the Importe column
    let importeColRightEdge = 0;

    autoTable(doc, {
        startY: yPos,
        head: [tableHead],
        body: tableBody,
        theme: 'plain',
        styles: {
            fontSize: 9,
            cellPadding: 3,
            valign: 'top', // ALIGN TOP so first lines match
            textColor: 20,
            overflow: 'linebreak'
        },
        headStyles: {
            fontStyle: 'bold',
            textColor: 0,
            lineWidth: { bottom: 0.1 },
            lineColor: 0,
            valign: 'middle', // Align headers vertically center
        },
        columnStyles: {
            0: { cellWidth: 45, fontStyle: 'bold' }, // Recurso INCREASED WIDTH
            1: { cellWidth: 'auto' }, // Descripción
            2: { cellWidth: 15, halign: 'center' }, // Ud CENTERED
            // Dynamic column index handling for Averia
            ...(isAveria ? {
                3: { cellWidth: 12, halign: 'center' }, // K CENTERED
                4: { cellWidth: 25, halign: 'right' }, // Precio RIGHT (WIDER)
                5: { cellWidth: 28, halign: 'right', fontStyle: 'bold' }, // Importe RIGHT (WIDER)
                6: { cellWidth: 35, halign: 'left' } // Obs LEFT
            } : {
                3: { cellWidth: 25, halign: 'right' }, // Precio RIGHT (WIDER)
                4: { cellWidth: 28, halign: 'right', fontStyle: 'bold' }, // Importe RIGHT (WIDER)
                5: { cellWidth: 35, halign: 'left' } // Obs LEFT
            })
        },
        // Capture the X position of the Importe column to align the total later
        didDrawCell: (data) => {
            const importeIndex = isAveria ? 5 : 4;
            if (data.column.index === importeIndex && data.section === 'body') {
                importeColRightEdge = data.cell.x + data.cell.width;
            }
        },
        // Custom Hook to justify Description text inside table
        didParseCell: (data) => {
             // Force Header Alignment to match data
             if (data.section === 'head') {
                const idx = data.column.index;
                const priceIdx = isAveria ? 4 : 3;
                const totalIdx = isAveria ? 5 : 4;
                const udIdx = 2;
                const kIdx = 3;

                if (idx === priceIdx || idx === totalIdx) {
                    data.cell.styles.halign = 'right';
                }
                if (idx === udIdx || (isAveria && idx === kIdx)) {
                    data.cell.styles.halign = 'center';
                }
            }
            if (data.section === 'body' && data.column.index === 1) {
                // Keep description left aligned
            }
        }
    });

    // 5. FOOTER (Total)
    const finalY = (doc as any).lastAutoTable.finalY;
    const totalAmount = items.reduce((acc, curr) => {
        const k = isAveria ? (curr.kFactor || 1) : 1;
        // Adjust price using marginDivisor
        const adjustedUnitPrice = curr.unitPrice / marginDivisor;
        return acc + roundToTwo(curr.currentQuantity * k * adjustedUnitPrice);
    }, 0);

    // Check if we need a new page for total
    if (finalY > pageHeight - 30) {
        doc.addPage();
        yPos = 20;
    } else {
        yPos = finalY + 5;
    }

    // Total Line
    // We start the line based on the importe column edge approx, or just a bit wide
    const lineStart = pageWidth - 80; 
    doc.setDrawColor(0);
    doc.setLineWidth(0.5);
    doc.line(lineStart, yPos, pageWidth - margin, yPos);
    
    yPos += 6;
    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    
    // ALIGN TOTAL WITH IMPORTE COLUMN
    // Use the captured importeColRightEdge for the numeric value alignment
    // Fallback to page margin if for some reason it wasn't captured (empty table)
    const totalValueX = importeColRightEdge > 0 ? importeColRightEdge : (pageWidth - margin - 35 - margin); // approx backup
    
    doc.text("TOTAL:", totalValueX - 30, yPos, { align: 'right' });
    doc.text(formatCurrency(totalAmount), totalValueX, yPos, { align: 'right' });


    // 6. PAGE NUMBERING (Hoja X de Y)
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(9);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(100);
        doc.text(`Hoja ${i} de ${totalPages}`, pageWidth - margin, pageHeight - 10, { align: 'right' });
    }

    // 7. Save
    const fileName = `${type === 'proforma' ? 'Proforma' : 'Certificacion'}_${projectInfo.projectNumber || 'Obra'}.pdf`;
    doc.save(fileName);
    setShowExportMenu(false);
  };

  // Helper to get PX from MM (approx for 96DPI / Browser Rendering)
  const mmToPx = (mm: number) => mm * 3.78;

  // --- ITEM MANAGEMENT ---

  const fillRowWithMaster = (rowId: string, masterItem: BudgetItem) => {
    saveHistory(state); // Save current state before changing
    setState(prev => {
        if (!prev) return prev;
        const isAveria = prev.projectInfo.isAveria;
        return {
            ...prev,
            items: prev.items.map(i => {
                if (i.id === rowId) {
                    const k = i.kFactor || 1;
                    const finalK = isAveria ? k : 1;
                    return {
                        ...i,
                        code: masterItem.code,
                        description: masterItem.description,
                        unit: masterItem.unit,
                        unitPrice: masterItem.unitPrice, // Master items already rounded on import
                        // Keep quantities if user already typed them, otherwise 0
                        totalAmount: roundToTwo(i.currentQuantity * finalK * masterItem.unitPrice),
                        observations: i.observations || ''
                    };
                }
                return i;
            })
        };
    });
    setActiveSearch(null);
  };

  const addEmptyItem = (targetIndex?: number) => {
    saveHistory(state); // Save current state before changing
    // 1. Add Item (No validation restrictions)
    const newItem: BudgetItem = {
        id: `manual-${Date.now()}-${Math.random()}`,
        code: '',
        description: '',
        unit: 'ud',
        plannedQuantity: 0,
        unitPrice: 0,
        previousQuantity: 0,
        currentQuantity: 0,
        totalQuantity: 0,
        totalAmount: 0,
        observations: '',
        kFactor: 1
    };
    
    setState(prev => {
        if (!prev) return prev;
        const newItems = [...prev.items];
        if (typeof targetIndex === 'number') {
            newItems.splice(targetIndex + 1, 0, newItem);
        } else if (selectedRowId) {
            const idx = newItems.findIndex(i => i.id === selectedRowId);
            if (idx !== -1) {
                newItems.splice(idx + 1, 0, newItem);
            } else {
                newItems.push(newItem);
            }
        } else {
            newItems.push(newItem);
        }
        return { ...prev, items: newItems };
    });

    // Auto-select the newly added empty row
    setTimeout(() => {
        setSelectedRowId(newItem.id);
    }, 50);
  };

  const deleteItem = (id: string | null) => {
    if (!id) return;
    saveHistory(state); // Save current state before changing
    setState(prev => {
      if (!prev) return prev;
      // Remove item and its checkbox state
      const newItems = prev.items.filter(i => i.id !== id);
      const newChecks = new Set(prev.checkedRowIds);
      if (newChecks.has(id)) newChecks.delete(id);
      
      return {
        ...prev,
        items: newItems,
        checkedRowIds: newChecks
      };
    });
    setContextMenu({ ...contextMenu, visible: false });
    if (selectedRowId === id) setSelectedRowId(null);
  };

  // Reorder Row Function
  const moveRow = (fromIndex: number, toIndex: number) => {
      if (fromIndex === toIndex) return;
      saveHistory(state);
      setState(prev => {
          if (!prev) return prev;
          const newItems = [...prev.items];
          const [movedItem] = newItems.splice(fromIndex, 1);
          newItems.splice(toIndex, 0, movedItem);
          return { ...prev, items: newItems };
      });
  };

  const updateQuantity = (id: string, qty: number) => {
    // Note: We do NOT saveHistory here on every keystroke/change. 
    // It's handled by handleInputBlur when the user leaves the field.
    setState(prev => {
      if (!prev) return prev;
      return {
        ...prev,
        items: prev.items.map(i => i.id === id ? calculateItemTotals({ ...i, currentQuantity: qty }) : i)
      };
    });
  };
  
  const updateField = (id: string, field: keyof BudgetItem, value: string | number) => {
     // Note: History saved on Blur
     setState(prev => {
        if (!prev) return prev;
        return {
            ...prev,
            items: prev.items.map(i => {
                if (i.id === id) {
                    const updatedItem = { ...i, [field]: value };
                    // If we change Price or K, we must recalculate total
                    if (field === 'unitPrice' || field === 'kFactor') {
                        return calculateItemTotals(updatedItem);
                    }
                    return updatedItem;
                }
                return i;
            })
        };
     });
     
     // Trigger search dropdown if typing in code or description
     if (field === 'code' || field === 'description') {
         setActiveSearch({ rowId: id, field });
     }
  };

  const updateProjectInfo = (field: keyof ProjectInfo, value: string | number | boolean) => {
    // Note: History saved on Blur
    setState(prev => {
      if (!prev) return prev;
      
      // If we are toggling AVERIA mode, trigger a recalc of all items
      if (field === 'isAveria') {
          const newIsAveria = value as boolean;
          // When switching back to normal, we might want to recalculate totals assuming K=1 implicitly
          // The calculateItemTotals function uses item.kFactor if present.
          // To strictly follow "Total = Ud * K * Price", we need to ensure that when isAveria is FALSE,
          // the effective K is 1.
          
          const recalculatedItems = prev.items.map(item => {
               const k = newIsAveria ? (item.kFactor || 1) : 1; 
               // We don't overwrite item.kFactor, just use 'k' for calculation
               const totalAmount = roundToTwo(item.currentQuantity * k * item.unitPrice);
               return { ...item, totalAmount };
          });

          return {
              ...prev,
              items: recalculatedItems,
              projectInfo: {
                  ...prev.projectInfo,
                  [field]: value
              }
          };
      }

      return {
          ...prev,
          projectInfo: {
            ...prev.projectInfo,
            [field]: value
          }
      };
    });
  };

  // --- INLINE SEARCH HELPER ---
  const getInlineSearchResults = (term: string) => {
      if (!term || term.length < 2 || !state) return [];
      const lowerTerm = term.toLowerCase();
      // REMOVED .slice(0, 10) to show all results as requested
      return state.masterItems.filter(i => 
        i.code.toLowerCase().includes(lowerTerm) || 
        i.description.toLowerCase().includes(lowerTerm)
      );
  };
  
  // Calculate Dropdown Placement based on screen position
  const calculateDropdownPlacement = (target: HTMLElement) => {
      const rect = target.getBoundingClientRect();
      const spaceBelow = window.innerHeight - rect.bottom;
      // If less than 300px space below, position upwards
      setDropdownPlacement(spaceBelow < 300 ? 'top' : 'bottom');
  };

  // SAFEGUARD: If state is null (e.g. during heavy operations or initialization glitches), do not render
  if (!state) return <div className="h-screen flex items-center justify-center bg-slate-50 text-slate-400">Cargando aplicación...</div>;

  // Calculate global total respecting Averia mode
  const totalAmount = state.items.reduce((acc, curr) => {
      const k = state.projectInfo.isAveria ? (curr.kFactor || 1) : 1;
      return acc + roundToTwo(curr.currentQuantity * k * curr.unitPrice);
  }, 0);

  // --- SELECTION & STATS LOGIC ---
  
  // Register mouse up globally to stop dragging
  useEffect(() => {
      const handleWindowMouseUp = () => {
          if (selection.isDragging) {
              setSelection(prev => ({ ...prev, isDragging: false }));
          }
      };
      window.addEventListener('mouseup', handleWindowMouseUp);
      return () => window.removeEventListener('mouseup', handleWindowMouseUp);
  }, [selection.isDragging]);

  const handleCellMouseDown = (r: number, c: number, e: React.MouseEvent) => {
      // Allow default input focus if it's a left click without movement yet
      // We set dragging to true, but we reset selection start
      setSelection({
          start: { r, c },
          end: { r, c },
          isDragging: true
      });
  };

  const handleCellMouseEnter = (r: number, c: number) => {
      if (selection.isDragging && selection.start) {
          // KEY FIX: Remove native text selection to prevent "blue mess" across cells
          if (window.getSelection) {
              window.getSelection()?.removeAllRanges();
          }
          setSelection(prev => ({
              ...prev,
              end: { r, c }
          }));
      }
  };

  const isCellSelected = (r: number, c: number) => {
      if (!selection.start || !selection.end) return false;
      
      const minR = Math.min(selection.start.r, selection.end.r);
      const maxR = Math.max(selection.start.r, selection.end.r);
      const minC = Math.min(selection.start.c, selection.end.c);
      const maxC = Math.max(selection.start.c, selection.end.c);

      return r >= minR && r <= maxR && c >= minC && c <= maxC;
  };

  // Calculate stats from selection
  const selectionStats = useMemo(() => {
      if (!selection.start || !selection.end || !state) return null;
      
      const minR = Math.min(selection.start.r, selection.end.r);
      const maxR = Math.max(selection.start.r, selection.end.r);
      const minC = Math.min(selection.start.c, selection.end.c);
      const maxC = Math.max(selection.start.c, selection.end.c);

      let count = 0;
      let sum = 0;
      let hasNumbers = false;
      const isAveria = state.projectInfo.isAveria;

      for (let r = minR; r <= maxR; r++) {
          const item = state.items[r];
          if (!item) continue;

          for (let c = minC; c <= maxC; c++) {
              count++;
              // Map visual columns to data values
              // Adjusted Col Mapping for Averia Mode
              // Normal: 1:Code, 2:Desc, 3:Qty, 4:Price, 5:Total, 6:Obs
              // Averia: 1:Code, 2:Desc, 3:Qty, 4:K, 5:Price, 6:Total, 7:Obs
              let value = 0;
              let isNum = false;

              if (isAveria) {
                 if (c === 3) { value = item.currentQuantity; isNum = true; }
                 else if (c === 4) { value = item.kFactor || 1; isNum = true; }
                 else if (c === 5) { value = item.unitPrice; isNum = true; }
                 else if (c === 6) { value = item.totalAmount; isNum = true; }
              } else {
                 if (c === 3) { value = item.currentQuantity; isNum = true; }
                 else if (c === 4) { value = item.unitPrice; isNum = true; }
                 else if (c === 5) { value = item.totalAmount; isNum = true; }
              }

              if (isNum) {
                  sum += value;
                  hasNumbers = true;
              }
          }
      }

      return { count, sum, hasNumbers };

  }, [selection, state]);

  // Event Handlers for UI
  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      // Close context menu
      if (contextMenu.visible) {
         setContextMenu(prev => ({ ...prev, visible: false }));
      }
      // Close export menu
      if (exportBtnRef.current && !exportBtnRef.current.contains(e.target as Node)) {
        setShowExportMenu(false);
      }
      // Close inline search dropdown if clicking outside
      if (activeSearch && dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
          const target = e.target as HTMLElement;
          if (target.tagName !== 'INPUT' && target.tagName !== 'TEXTAREA') {
              setActiveSearch(null);
          }
      }
    };
    document.addEventListener('click', handleClick);
    return () => document.removeEventListener('click', handleClick);
  }, [contextMenu.visible, showExportMenu, activeSearch]);

  const handleContextMenu = (e: React.MouseEvent, id: string) => {
    e.preventDefault();
    setSelectedRowId(id);
    setContextMenu({
        visible: true,
        x: e.clientX,
        y: e.clientY,
        rowId: id
    });
  };

  // Auto-resize textarea height
  const adjustTextareaHeight = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      e.target.style.height = 'auto';
      e.target.style.height = `${e.target.scrollHeight}px`;
  };

  return (
    <div className="h-screen flex flex-col bg-slate-50 font-sans text-slate-900 overflow-hidden relative">
      
      {/* PROFESSIONAL HEADER AREA - Visible on Screen */}
      <div className="bg-white border-b border-slate-300 flex flex-col shrink-0 z-30 shadow-sm relative">
         
         {/* TOP STRIP: Title + Actions */}
         <div className="flex items-center justify-between px-6 py-3 border-b border-slate-100">
             <div className="flex items-center gap-3">
                 <div className="bg-slate-900 text-white p-1.5 rounded shadow-sm">
                    <FileSpreadsheet className="w-5 h-5"/>
                 </div>
                 <div>
                    <h1 className="text-xl font-bold text-slate-900 leading-tight">TENSA SA</h1>
                 </div>
             </div>

             <div className="flex items-center gap-2">
                 {/* UNDO / REDO BUTTONS */}
                 <div className="flex items-center gap-2 mr-4 border-r border-slate-200 pr-4">
                     <button 
                        onClick={undo} 
                        onMouseDown={(e) => e.preventDefault()}
                        disabled={history.past.length === 0 && (!historySnapshot.current || historySnapshot.current === state)}
                        className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-blue-600 hover:border-blue-400 rounded-md shadow-sm disabled:opacity-40 disabled:cursor-not-allowed transition-all"
                        title="Deshacer (Ctrl+Z)"
                     >
                         <Undo className="w-4 h-4" />
                         <span className="text-xs font-semibold hidden lg:inline">Deshacer</span>
                     </button>
                     <button 
                        onClick={redo}
                        onMouseDown={(e) => e.preventDefault()}
                        disabled={history.future.length === 0} 
                        className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-blue-600 hover:border-blue-400 rounded-md shadow-sm disabled:opacity-40 disabled:cursor-not-allowed transition-all"
                        title="Rehacer (Ctrl+Y)"
                     >
                         <Redo className="w-4 h-4" />
                         <span className="text-xs font-semibold hidden lg:inline">Rehacer</span>
                     </button>
                 </div>

                 {/* LOAD PROJECT BUTTON - REFACTORED TO NATIVE LABEL */}
                 <label className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer text-sm text-slate-700 font-medium transition-colors select-none" title="Cargar proyecto guardado (.json)">
                     <FolderOpen className="w-4 h-4 text-orange-500" />
                     Cargar Trabajo
                     <input 
                        type="file" 
                        className="hidden" 
                        accept=".json" 
                        onChange={handleLoadProject} 
                        onClick={(e) => (e.currentTarget.value = '')}
                     />
                 </label>

                 {/* SAVE PROJECT BUTTON */}
                 <button 
                     onClick={handleSaveProject}
                     className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer text-sm text-slate-700 font-medium transition-colors"
                     title="Guardar proyecto actual (.json)"
                 >
                     <Save className="w-4 h-4 text-blue-500" />
                     Guardar Trabajo
                 </button>

                 <label 
                    className={`flex items-center gap-2 px-3 py-2 rounded border cursor-pointer text-sm font-medium transition-colors ml-4 ${
                        state.masterItems.length > 0 
                        ? "bg-emerald-100 border-emerald-300 text-emerald-800 hover:bg-emerald-200 shadow-sm" 
                        : "bg-white border-slate-300 text-slate-700 hover:bg-slate-50"
                    }`}
                    title={state.masterItems.length > 0 ? `Tabla de recursos cargada (${state.masterItems.length} elementos)` : "Importar tabla de recursos desde Excel"}
                 >
                     <Upload className={`w-4 h-4 ${state.masterItems.length > 0 ? "text-emerald-700" : "text-slate-500"}`} />
                     {state.masterItems.length > 0 ? "Tabla Rec. (OK)" : "Importar Tabla Rec."}
                     <input 
                        type="file" 
                        className="hidden" 
                        accept=".xlsx, .xls" 
                        onChange={handleFileUpload}
                        onClick={(e) => (e.currentTarget.value = '')}
                     />
                 </label>

                 <button 
                     onClick={handleClearAll}
                     disabled={state.items.length === 0 && !state.projectInfo.isAveria && !state.projectInfo.name && !state.projectInfo.projectNumber && !state.projectInfo.orderNumber && !state.projectInfo.client}
                     className="flex items-center gap-2 px-3 py-2 bg-white border border-red-200 text-red-600 rounded hover:bg-red-50 disabled:opacity-50 disabled:cursor-not-allowed text-sm font-medium transition-colors ml-2"
                     title="Borrar todos los datos de la certificación (filas y cabecera)"
                 >
                     <Eraser className="w-4 h-4" />
                     Limpiar
                 </button>
                 
                 {/* EXPORT DROPDOWN */}
                 <div className="relative" ref={exportBtnRef}>
                    <button 
                        className="flex items-center gap-2 px-3 py-2 bg-emerald-600 text-white rounded hover:bg-emerald-700 text-sm font-medium shadow-sm transition-colors"
                        onClick={() => setShowExportMenu(!showExportMenu)}
                    >
                        <Download className="w-4 h-4" />
                        Exportar
                        <ChevronDown className="w-3 h-3 opacity-70" />
                    </button>
                    
                    {showExportMenu && (
                        <div className="absolute right-0 mt-1 w-64 bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50 animate-in fade-in slide-in-from-top-2">
                             <button onClick={exportToExcel} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2">
                                <FileSpreadsheet className="w-4 h-4 text-green-600" /> Excel (.xlsx)
                             </button>
                             <div className="h-px bg-slate-100 my-1"></div>
                             <div className="px-4 py-1 text-xs font-bold text-slate-400 uppercase tracking-wider">PDF</div>
                             <button onClick={() => downloadPDF('landscape', 'certification')} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2">
                                <FileText className="w-4 h-4 text-red-600" /> Exportación PDF
                             </button>
                             <button onClick={handleProformaClick} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2">
                                <CheckSquare className="w-4 h-4 text-blue-600" /> Factura Proforma (Selección)
                             </button>
                        </div>
                    )}
                 </div>

                 {/* CALCULATOR TOGGLE BUTTON */}
                 <button 
                     onClick={() => setShowCalculator(!showCalculator)}
                     className={`flex items-center gap-2 px-3 py-2 rounded text-sm font-medium transition-colors ml-2 shadow-sm ${
                         showCalculator 
                            ? "bg-slate-700 text-white border border-slate-600" 
                            : "bg-white border border-slate-300 text-slate-700 hover:bg-slate-50"
                     }`}
                     title="Calculadora Flotante"
                 >
                     <Calculator className="w-4 h-4" />
                 </button>

             </div>
         </div>

         {/* MAIN INFO GRID (Editable) */}
         <div className="px-6 py-4 bg-slate-50/50">
             <div className="grid grid-cols-12 gap-x-6 gap-y-4">
                 
                 {/* Row 1: Denominación, Nº Obra, Nº Pedido, Checkbox */}
                 <div className="col-span-6">
                     <label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Denominación de la Obra</label>
                     <input 
                        type="text" 
                        className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-2xl font-bold text-slate-800 pb-1 focus:bg-white transition-colors placeholder-slate-300"
                        value={state.projectInfo.name}
                        onChange={(e) => updateProjectInfo('name', e.target.value)}
                        onFocus={handleInputFocus}
                        onBlur={handleInputBlur}
                     />
                 </div>
                 <div className="col-span-2">
                     <label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Nº Obra</label>
                     <input 
                        type="text" 
                        className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-sans text-slate-700 pb-1 focus:bg-white transition-colors"
                        value={state.projectInfo.projectNumber}
                        onChange={(e) => updateProjectInfo('projectNumber', e.target.value)}
                        onFocus={handleInputFocus}
                        onBlur={handleInputBlur}
                     />
                 </div>
                 <div className="col-span-2">
                     <label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Nº Pedido</label>
                     <input 
                        type="text" 
                        className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-sans text-slate-700 pb-1 focus:bg-white transition-colors"
                        value={state.projectInfo.orderNumber}
                        onChange={(e) => updateProjectInfo('orderNumber', e.target.value)}
                        onFocus={handleInputFocus}
                        onBlur={handleInputBlur}
                     />
                 </div>
                 <div className="col-span-2 flex items-end">
                     <label className="flex items-center gap-2 cursor-pointer select-none bg-red-50 px-3 py-2 rounded border border-red-100 hover:bg-red-100 transition-colors w-full">
                         <input 
                            type="checkbox" 
                            className="w-5 h-5 rounded border-red-400 text-red-600 focus:ring-red-500"
                            checked={state.projectInfo.isAveria || false}
                            onChange={(e) => updateProjectInfo('isAveria', e.target.checked)}
                         />
                         <span className="font-bold text-red-700 uppercase">AVERÍA</span>
                     </label>
                 </div>

                 {/* Row 2: Cliente, Fecha, Total */}
                 <div className="col-span-8">
                     <label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Cliente</label>
                     <input 
                        type="text" 
                        className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-medium text-slate-700 pb-1 focus:bg-white transition-colors"
                        value={state.projectInfo.client}
                        onChange={(e) => updateProjectInfo('client', e.target.value)}
                        onFocus={handleInputFocus}
                        onBlur={handleInputBlur}
                     />
                 </div>
                 <div className="col-span-2">
                     <label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Fecha</label>
                     <input 
                        type="date" 
                        className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-medium text-slate-700 pb-1 focus:bg-white transition-colors"
                        value={state.projectInfo.date}
                        onChange={(e) => updateProjectInfo('date', e.target.value)}
                        onFocus={handleInputFocus}
                        onBlur={handleInputBlur}
                     />
                 </div>
                  <div className="col-span-2 text-right">
                     <label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Total Acumulado</label>
                     <div className="text-3xl font-sans font-bold text-emerald-700 leading-none pb-1 tabular-nums">{formatCurrency(totalAmount)}</div>
                 </div>

                 {/* AVERIA EXTRA DETAILS FIELDS - Only visible if checked */}
                 {state.projectInfo.isAveria && (
                   <div className="col-span-12 bg-red-50 p-4 rounded border border-red-100 mt-2 animate-in fade-in slide-in-from-top-2 shadow-sm">
                      <div className="flex items-center gap-2 mb-3 text-red-800 font-bold uppercase text-sm border-b border-red-200 pb-1">
                          <AlertTriangle className="w-4 h-4" /> Detalles de la Avería
                      </div>
                      <div className="grid grid-cols-12 gap-6">
                          <div className="col-span-2">
                              <label className="block text-xs font-bold text-red-600 uppercase mb-1">Nº Avería</label>
                              <div className="relative">
                                <Hash className="w-4 h-4 absolute left-2 top-2.5 text-red-300" />
                                <input 
                                    type="text" 
                                    autoFocus
                                    className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium"
                                    placeholder="Ej: Axxxxx, Sxxxxx, Txxxxx"
                                    value={state.projectInfo.averiaNumber || ''}
                                    onChange={(e) => updateProjectInfo('averiaNumber', e.target.value)}
                                    onBlur={handleInputBlur}
                                    onFocus={handleInputBlur}
                                />
                              </div>
                          </div>
                          <div className="col-span-2">
                              <label className="block text-xs font-bold text-red-600 uppercase mb-1">Fecha de Avería</label>
                              <div className="relative">
                                <Calendar className="w-4 h-4 absolute left-2 top-2.5 text-red-300" />
                                <input 
                                    type="date" 
                                    className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium"
                                    value={state.projectInfo.averiaDate || ''}
                                    onChange={(e) => updateProjectInfo('averiaDate', e.target.value)}
                                    onBlur={handleInputBlur}
                                    onFocus={handleInputBlur}
                                />
                              </div>
                          </div>
                          <div className="col-span-2">
                               <div className="flex items-center gap-1 mb-1">
                                    <label className="block text-xs font-bold text-red-600 uppercase">Horario</label>
                                    <div className="group relative">
                                        <Info className="w-3.5 h-3.5 text-red-400 cursor-help" />
                                        <div className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-60 bg-slate-800 text-white text-xs p-2.5 rounded shadow-xl opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-50 leading-relaxed border border-slate-600">
                                            <div className="font-bold border-b border-slate-600 pb-1 mb-1 text-slate-200">Definición de Horarios</div>
                                            <p className="mb-1"><span className="text-orange-300 font-semibold">Diurno (K=1,25):</span> Lunes a Viernes laborables de 07:00 a 19:00h.</p>
                                            <p><span className="text-indigo-300 font-semibold">Nocturno/Finde (K=1,75):</span> Resto de horas, fines de semana y festivos.</p>
                                            <div className="absolute top-full left-1/2 -translate-x-1/2 -mt-1 border-4 border-transparent border-t-slate-800"></div>
                                        </div>
                                    </div>
                               </div>
                               <div className="relative">
                                   <Clock className="w-4 h-4 absolute left-2 top-2.5 text-red-300" />
                                   <select 
                                       className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium appearance-none"
                                       value={state.projectInfo.averiaTiming || 'diurna'}
                                       onChange={(e) => updateProjectInfo('averiaTiming', e.target.value)}
                                       onBlur={handleInputBlur}
                                       onFocus={handleInputBlur}
                                   >
                                       <option value="diurna">Avería diurna K=1,25</option>
                                       <option value="nocturna_finde">Avería nocturna/ Fin de Semana K=1,75</option>
                                   </select>
                                   <ChevronDown className="w-4 h-4 absolute right-2 top-2.5 text-red-300 pointer-events-none" />
                                </div>
                           </div>
                          <div className="col-span-6">
                              <label className="block text-xs font-bold text-red-600 uppercase mb-1">Descripción de la Avería</label>
                              <div className="relative">
                                <AlignLeft className="w-4 h-4 absolute left-2 top-2.5 text-red-300" />
                                <textarea 
                                    rows={2}
                                    className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium resize-none text-justify"
                                    placeholder="Describir brevemente la avería..."
                                    value={state.projectInfo.averiaDescription || ''}
                                    onChange={(e) => updateProjectInfo('averiaDescription', e.target.value)}
                                    onBlur={handleInputBlur}
                                    onFocus={handleInputBlur}
                                />
                              </div>
                          </div>
                      </div>
                   </div>
                 )}
             </div>
         </div>
      </div>

      {/* MAIN CONTENT AREA */}
      <div className="flex-1 overflow-hidden relative">
        <div className="flex flex-col h-full bg-white relative">
          
          {/* BARRA SUPERIOR (Limpia) */}
          <div className="flex items-center gap-2 px-4 py-2 border-b border-slate-300 bg-slate-50 sticky top-0 z-20 h-10 shrink-0">
            {!selectedRowId && (
                <div className="text-sm text-slate-400 italic ml-auto">
                    Seleccione una fila para editar o arrastre el ratón para calcular totales
                </div>
            )}
          </div>

          {/* GRID PRINCIPAL */}
          <div ref={tableContainerRef} className={`flex-1 overflow-auto relative bg-slate-100/50 ${selection.isDragging ? 'select-none' : ''}`}>
            <table className="w-full border-collapse text-base table-fixed min-w-[1050px] bg-white">
              <thead className="sticky top-0 z-10 bg-orange-100 text-orange-900 font-semibold border-b border-orange-200 shadow-sm">
                 <tr className="text-base uppercase tracking-wider">
                   {/* CHECKBOX HEADER - REPLACED WITH CCAA LABEL */}
                   <th className="w-12 text-center py-4 border-r border-orange-200 bg-orange-200 text-xs font-bold select-none text-orange-800">
                       CCAA
                   </th>
                   <th className="w-10 text-center py-4 border-r border-orange-200 bg-orange-200 text-sm select-none">#</th>
                   <th className="w-64 px-4 py-4 text-center border-r border-orange-200 font-bold">RECURSO</th>
                   <th className="w-[450px] px-4 py-4 text-center border-r border-orange-200 font-bold">DESCRIPCIÓN</th>
                   <th className="w-24 px-4 py-4 text-center border-r border-orange-200 bg-orange-50 text-orange-800 font-bold">UD</th>
                   {/* CONDITIONAL K COLUMN */}
                   {state.projectInfo.isAveria && (
                        <th className="w-20 px-4 py-4 text-center border-r border-orange-200 bg-red-100 text-red-800 font-bold">K</th>
                   )}
                   <th className="w-24 px-4 py-4 text-center border-r border-orange-200 font-bold">PRECIO</th>
                   <th className="w-28 px-4 py-4 text-center font-bold">TOTAL</th>
                   <th className="w-64 px-4 py-4 text-center border-r border-orange-200 border-l border-orange-200 bg-orange-50/50 font-bold">OBSERVACIONES</th>
                   <th className="w-24 px-3 py-4 text-center border-l border-orange-200 bg-orange-200 font-bold">ACCIONES</th>
                 </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                 {state.items.map((item, index) => {
                   const effectiveTotal = roundToTwo(item.currentQuantity * (state.projectInfo.isAveria ? (item.kFactor || 1) : 1) * item.unitPrice);
                   return (
                   <tr 
                      key={item.id} 
                      className={`group cursor-default transition-colors ${
                        selectedRowId === item.id 
                            ? 'bg-blue-50' 
                            : index % 2 === 0 ? 'bg-white hover:bg-slate-200' : 'bg-slate-100 hover:bg-slate-200'
                      }`}
                      onClick={() => setSelectedRowId(item.id)}
                      onContextMenu={(e) => handleContextMenu(e, item.id)}
                   >
                     {/* CHECKBOX COLUMN */}
                     <td className="border-r border-slate-300 text-center bg-slate-100 align-top pt-4">
                         <input 
                            type="checkbox" 
                            className="w-4 h-4 rounded border-slate-400 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                            checked={state.checkedRowIds.has(item.id)}
                            onChange={(e) => {
                                e.stopPropagation(); // Prevent row selection logic
                                toggleRowCheck(item.id);
                            }}
                            onClick={(e) => e.stopPropagation()}
                         />
                     </td>

                     {/* ROW NUMBER / HANDLE - NOW DRAGGABLE */}
                     <td 
                        className={`border-r border-slate-300 text-center text-base text-slate-400 select-none align-top pt-4 cursor-move hover:text-slate-600 active:text-slate-800 ${selectedRowId === item.id ? 'bg-blue-200 text-blue-700 font-bold' : 'bg-slate-200'}`}
                        draggable
                        onDragStart={(e) => {
                            setDraggedRowIndex(index);
                            e.dataTransfer.effectAllowed = "move";
                            // Optional: Transparent ghost image if needed, but browser default is usually fine for rows
                        }}
                        onDragOver={(e) => {
                            e.preventDefault(); // Necessary to allow dropping
                            e.dataTransfer.dropEffect = "move";
                            
                            // Auto-scroll logic
                            if (tableContainerRef.current) {
                                const { top, bottom } = tableContainerRef.current.getBoundingClientRect();
                                const y = e.clientY;
                                const threshold = 100; // Pixels from edge to trigger scroll

                                if (y < top + threshold) {
                                    // Scroll Up - faster as we get closer to edge
                                    autoScrollSpeed.current = -15;
                                } else if (y > bottom - threshold) {
                                    // Scroll Down
                                    autoScrollSpeed.current = 15;
                                } else {
                                    // Stop scrolling
                                    autoScrollSpeed.current = 0;
                                }
                            }
                        }}
                        onDragEnd={() => {
                            setDraggedRowIndex(null);
                            autoScrollSpeed.current = 0;
                        }}
                        onDrop={(e) => {
                            e.preventDefault();
                            autoScrollSpeed.current = 0;
                            if (draggedRowIndex !== null) {
                                moveRow(draggedRowIndex, index);
                                setDraggedRowIndex(null);
                            }
                        }}
                        title="Arrastrar para reordenar fila"
                     >
                        {index + 1}
                     </td>

                     {/* RECURSO (Con Búsqueda Inline) - TEXTAREA JUSTIFICADA */}
                     <td 
                        className={`border-r border-slate-200 p-0 relative align-top ${isCellSelected(index, 1) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(index, 1, e)}
                        onMouseEnter={() => handleCellMouseEnter(index, 1)}
                     >
                        <textarea 
                            rows={1}
                            className="w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none font-sans text-base text-slate-800 focus:bg-white focus:ring-2 focus:ring-inset focus:ring-emerald-500 text-justify resize-none overflow-hidden leading-relaxed relative z-0"
                            value={item.code}
                            onChange={(e) => {
                                updateField(item.id, 'code', e.target.value);
                                adjustTextareaHeight(e);
                            }}
                            onFocus={(e) => {
                                handleInputFocus();
                                setSelectedRowId(item.id);
                                setActiveSearch({ rowId: item.id, field: 'code' });
                                adjustTextareaHeight(e);
                                calculateDropdownPlacement(e.target); // Calculate position on focus
                            }}
                            onBlur={handleInputBlur}
                            placeholder="Buscar..."
                        />
                        {/* Dropdown de búsqueda */}
                        {activeSearch?.rowId === item.id && activeSearch.field === 'code' && item.code.length > 0 && (
                             <div ref={dropdownRef} className={`absolute left-0 w-[400px] bg-white border border-slate-300 shadow-xl z-50 max-h-60 overflow-y-auto rounded-sm ${dropdownPlacement === 'top' ? 'bottom-full mb-1 shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.1)]' : 'top-full mt-1'}`}>
                                 {getInlineSearchResults(item.code).map(res => (
                                     <div key={res.id} onClick={() => fillRowWithMaster(item.id, res)} className="px-3 py-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-100 text-base flex gap-2">
                                         <span className="font-bold font-sans text-slate-800">{res.code}</span>
                                         <span className="truncate flex-1">{res.description}</span>
                                     </div>
                                 ))}
                                 {getInlineSearchResults(item.code).length === 0 && <div className="p-3 text-sm text-slate-400 italic">Sin resultados</div>}
                             </div>
                        )}
                     </td>

                     {/* DESCRIPCIÓN (Con Búsqueda Inline) - TEXTAREA JUSTIFICADA */}
                     <td 
                        className={`border-r border-slate-200 p-0 relative align-top ${isCellSelected(index, 2) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(index, 2, e)}
                        onMouseEnter={() => handleCellMouseEnter(index, 2)}
                     >
                        <textarea 
                            rows={1}
                            className="w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none text-base text-slate-800 font-sans focus:bg-white focus:ring-2 focus:ring-inset focus:ring-emerald-500 text-justify resize-none overflow-hidden leading-relaxed relative z-0"
                            value={item.description}
                            onChange={(e) => {
                                updateField(item.id, 'description', e.target.value);
                                adjustTextareaHeight(e);
                            }}
                            onFocus={(e) => {
                                handleInputFocus();
                                setSelectedRowId(item.id);
                                setActiveSearch({ rowId: item.id, field: 'description' });
                                adjustTextareaHeight(e);
                                calculateDropdownPlacement(e.target); // Calculate position on focus
                            }}
                            onBlur={handleInputBlur}
                            placeholder="Descripción..."
                        />
                         {/* Dropdown de búsqueda */}
                         {activeSearch?.rowId === item.id && activeSearch.field === 'description' && item.description.length > 1 && (
                             <div ref={dropdownRef} className={`absolute left-0 w-full bg-white border border-slate-300 shadow-xl z-50 max-h-60 overflow-y-auto rounded-sm ${dropdownPlacement === 'top' ? 'bottom-full mb-1 shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.1)]' : 'top-full mt-1'}`}>
                                 {getInlineSearchResults(item.description).map(res => (
                                     <div key={res.id} onClick={() => fillRowWithMaster(item.id, res)} className="px-3 py-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-100 text-base flex gap-2">
                                         <span className="font-bold font-sans text-slate-500">{res.code}</span>
                                         <span className="truncate flex-1">{res.description}</span>
                                     </div>
                                 ))}
                             </div>
                        )}
                     </td>
                     
                     {/* MAIN INPUT: QUANTITY */}
                     <td 
                        className={`border-r border-slate-200 p-0 relative bg-yellow-50/30 align-top ${isCellSelected(index, 3) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(index, 3, e)}
                        onMouseEnter={() => handleCellMouseEnter(index, 3)}
                     >
                        <input 
                          type="number"
                          className="w-full px-3 py-3 text-center font-sans text-base text-slate-800 bg-transparent focus:bg-white focus:ring-2 focus:ring-inset focus:ring-emerald-500 outline-none tabular-nums relative z-0"
                          value={item.currentQuantity || ''}
                          placeholder="0"
                          onChange={(e) => updateQuantity(item.id, parseFloat(e.target.value) || 0)}
                          onFocus={() => {
                              handleInputFocus();
                              setSelectedRowId(item.id);
                          }}
                          onBlur={handleInputBlur}
                        />
                     </td>

                     {/* K FACTOR INPUT (CONDITIONAL) */}
                     {state.projectInfo.isAveria && (
                         <td 
                            className={`border-r border-slate-200 p-0 relative bg-red-50/30 align-top ${isCellSelected(index, 4) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                            onMouseDown={(e) => handleCellMouseDown(index, 4, e)}
                            onMouseEnter={() => handleCellMouseEnter(index, 4)}
                         >
                            <input 
                              type="number"
                              className="w-full px-3 py-3 text-center font-sans text-base text-red-700 font-bold bg-transparent focus:bg-white focus:ring-2 focus:ring-inset focus:ring-red-500 outline-none tabular-nums relative z-0"
                              value={item.kFactor || 1}
                              step="0.1"
                              onChange={(e) => updateField(item.id, 'kFactor', parseFloat(e.target.value) || 0)}
                              onFocus={() => {
                                  handleInputFocus();
                                  setSelectedRowId(item.id);
                              }}
                              onBlur={handleInputBlur}
                            />
                         </td>
                     )}

                     {/* PRICE INPUT (VIEW/EDIT MODE) */}
                     <td 
                        className={`border-r border-slate-200 p-0 align-top ${isCellSelected(index, state.projectInfo.isAveria ? 5 : 4) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(index, state.projectInfo.isAveria ? 5 : 4, e)}
                        onMouseEnter={() => handleCellMouseEnter(index, state.projectInfo.isAveria ? 5 : 4)}
                        onClick={() => setEditingCell({ rowId: item.id, field: 'unitPrice' })}
                     >
                        {editingCell?.rowId === item.id && editingCell?.field === 'unitPrice' ? (
                            <input 
                                autoFocus
                                type="number"
                                className="w-full px-3 py-3 text-right bg-white outline-none font-sans text-base text-slate-800 focus:ring-2 focus:ring-inset focus:ring-emerald-500 tabular-nums relative z-20 shadow-inner"
                                value={item.unitPrice}
                                onChange={(e) => updateField(item.id, 'unitPrice', parseFloat(e.target.value) || 0)}
                                onFocus={(e) => {
                                    handleInputFocus();
                                    e.target.select();
                                }}
                                onBlur={() => {
                                    handleInputBlur();
                                    setEditingCell(null);
                                }}
                                onKeyDown={(e) => {
                                    if (e.key === 'Enter') e.currentTarget.blur();
                                }}
                            />
                        ) : (
                            <div 
                                className="w-full h-full px-3 py-3 text-right font-sans text-base text-slate-800 tabular-nums cursor-text"
                                tabIndex={0}
                                onFocus={() => setEditingCell({ rowId: item.id, field: 'unitPrice' })}
                            >
                                {formatCurrency(item.unitPrice)}
                            </div>
                        )}
                     </td>
                     
                     {/* TOTAL COLUMN - Highlights RED if 0 */}
                     <td 
                        className={`px-3 py-3 text-right font-sans text-base border-r border-slate-200 align-top tabular-nums ${effectiveTotal === 0 ? 'bg-red-100 text-red-600 font-medium' : 'text-slate-800 bg-slate-50/50'} ${isCellSelected(index, state.projectInfo.isAveria ? 6 : 5) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(index, state.projectInfo.isAveria ? 6 : 5, e)}
                        onMouseEnter={() => handleCellMouseEnter(index, state.projectInfo.isAveria ? 6 : 5)}
                     >
                       {formatCurrency(effectiveTotal)}
                     </td>

                     {/* OBSERVATIONS COLUMN - TEXTAREA JUSTIFICADA */}
                     <td 
                        className={`border-r border-slate-200 p-0 bg-slate-50/30 align-top ${isCellSelected(index, state.projectInfo.isAveria ? 7 : 6) ? 'bg-blue-200/50 ring-2 ring-inset ring-blue-500 z-10' : ''}`}
                        onMouseDown={(e) => handleCellMouseDown(index, state.projectInfo.isAveria ? 7 : 6, e)}
                        onMouseEnter={() => handleCellMouseEnter(index, state.projectInfo.isAveria ? 7 : 6)}
                     >
                        <textarea 
                            rows={1}
                            className="w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none text-base text-slate-800 font-sans focus:bg-white focus:ring-2 focus:ring-inset focus:ring-emerald-500 placeholder-slate-300 text-justify resize-none overflow-hidden leading-relaxed relative z-0"
                            value={item.observations || ''}
                            onChange={(e) => {
                                updateField(item.id, 'observations', e.target.value);
                                adjustTextareaHeight(e);
                            }}
                            placeholder="..."
                            onFocus={(e) => {
                                handleInputFocus();
                                setSelectedRowId(item.id);
                                adjustTextareaHeight(e);
                            }}
                            onBlur={handleInputBlur}
                        />
                     </td>

                     {/* ACTIONS COLUMN */}
                     <td className="border-l border-slate-200 p-0 text-center bg-slate-50 align-top">
                        <div className="flex items-center justify-center pt-3 gap-1 opacity-20 group-hover:opacity-100 transition-opacity">
                            <button 
                                onClick={(e) => { e.stopPropagation(); addEmptyItem(index); }}
                                className="p-2 hover:bg-emerald-100 text-emerald-600 rounded transition-colors"
                                title="Insertar fila vacía debajo"
                            >
                                <Plus className="w-5 h-5" />
                            </button>
                            <button 
                                onClick={(e) => { e.stopPropagation(); deleteItem(item.id); }}
                                className="p-2 hover:bg-red-100 text-red-600 rounded transition-colors"
                                title="Eliminar fila"
                            >
                                <Trash2 className="w-5 h-5" />
                            </button>
                        </div>
                     </td>
                   </tr>
                 );})}
                 
                 {/* EMPTY STATE */}
                 {state.items.length === 0 && (
                   <tr>
                     <td colSpan={state.projectInfo.isAveria ? 10 : 9} className="h-64 text-center text-slate-400 bg-slate-50">
                        <div className="flex flex-col items-center justify-center h-full">
                           <FileSpreadsheet className="w-20 h-20 mb-6 opacity-20" />
                           <p className="text-xl font-medium">Hoja vacía</p>
                           <p className="text-base mt-2 opacity-70">Importe un Excel de Recursos o use el botón "+" para añadir líneas manualmente</p>
                           <button 
                             onClick={() => addEmptyItem()}
                             className="mt-6 px-6 py-3 bg-emerald-600 text-white rounded shadow hover:bg-emerald-700 transition flex items-center gap-2 text-lg"
                           >
                              <Plus className="w-6 h-6" /> Añadir primera partida
                           </button>
                        </div>
                     </td>
                   </tr>
                 )}

                 {/* TOTAL ROW */}
                 {state.items.length > 0 && (
                   <tr className="bg-slate-100 font-bold border-t-2 border-slate-300 sticky bottom-0 z-20 shadow-[0_-2px_10px_rgba(0,0,0,0.05)]">
                     <td colSpan={state.projectInfo.isAveria ? 7 : 6} className="px-4 py-4 text-right text-base uppercase text-slate-500 tracking-wider">Total Certificación</td>
                     <td className="px-3 py-4 text-right text-xl text-slate-900 border-l border-slate-300 font-sans tabular-nums">{formatCurrency(totalAmount)}</td>
                     <td colSpan={2} className="bg-slate-100 border-l border-slate-300"></td>
                   </tr>
                 )}
              </tbody>
            </table>
          </div>
          
          {/* CUSTOM CONTEXT MENU */}
          {contextMenu.visible && (
            <div 
                className="fixed bg-white border border-slate-300 shadow-xl rounded-sm z-[100] w-64 py-2 flex flex-col text-left"
                style={{ top: contextMenu.y, left: contextMenu.x }}
            >
                <div className="px-4 py-2 text-sm text-slate-400 uppercase font-bold border-b border-slate-100 mb-1">
                    Opciones de Fila
                </div>
                <button 
                    className="px-4 py-3 text-left hover:bg-blue-50 text-slate-700 text-base flex items-center gap-2"
                    onClick={() => {
                        const idx = state.items.findIndex(i => i.id === contextMenu.rowId);
                        addEmptyItem(idx !== -1 ? idx : undefined);
                        setContextMenu(prev => ({...prev, visible: false}));
                    }}
                >
                    <Plus className="w-5 h-5" /> Insertar Fila Vacía
                </button>
                <div className="h-px bg-slate-200 my-1"></div>
                <button 
                    className="px-4 py-3 text-left hover:bg-red-50 text-red-600 text-base flex items-center gap-2"
                    onClick={() => deleteItem(contextMenu.rowId)}
                >
                    <Trash2 className="w-5 h-5" /> Eliminar Fila
                </button>
            </div>
          )}

           {/* PROFORMA MARGIN DIALOG */}
           {showProformaDialog && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-sm animate-in fade-in">
                <div className="bg-white rounded-lg shadow-2xl w-full max-w-md overflow-hidden transform transition-all scale-100">
                    <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                        <h3 className="font-bold text-lg text-slate-800 flex items-center gap-2">
                           <CheckSquare className="w-5 h-5 text-blue-600"/> Generar Proforma
                        </h3>
                        <button 
                           onClick={() => setShowProformaDialog(false)}
                           className="text-slate-400 hover:text-slate-600"
                        >
                            <X className="w-5 h-5"/>
                        </button>
                    </div>
                    
                    <div className="p-6">
                        <p className="text-slate-600 mb-4 text-sm leading-relaxed">
                            Indique el porcentaje de margen de beneficio que desea descontar en la proforma.
                            <br/>
                            <span className="text-xs text-slate-400 mt-1 block">
                                (El precio mostrado será: <span className="font-mono bg-slate-100 px-1 rounded">Precio / (1 + Margen/100)</span>)
                            </span>
                        </p>
                        
                        <div className="relative mb-6">
                            <label className="block text-xs font-bold text-slate-500 uppercase mb-1.5 ml-1">Margen de Beneficio (%)</label>
                            <div className="relative">
                                <Percent className="absolute left-3 top-3 w-5 h-5 text-slate-400" />
                                <input 
                                    ref={marginInputRef}
                                    type="number" 
                                    className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none font-bold text-slate-800 text-lg shadow-sm"
                                    placeholder="0"
                                    value={proformaMargin}
                                    onChange={(e) => setProformaMargin(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && confirmProformaExport()}
                                />
                            </div>
                        </div>

                        <div className="flex gap-3 justify-end">
                            <button 
                                onClick={() => setShowProformaDialog(false)}
                                className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-100 rounded transition-colors"
                            >
                                Cancelar
                            </button>
                            <button 
                                onClick={confirmProformaExport}
                                className="px-5 py-2 bg-blue-600 text-white font-bold rounded shadow hover:bg-blue-700 transition-colors flex items-center gap-2"
                            >
                                <FileText className="w-4 h-4" />
                                Generar PDF
                            </button>
                        </div>
                    </div>
                </div>
            </div>
          )}

           {/* DELETE ALL CONFIRMATION DIALOG */}
           {showClearDialog && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm animate-in fade-in">
                <div className="bg-white rounded-lg shadow-2xl w-full max-w-md overflow-hidden border-t-4 border-red-500 transform transition-all scale-100">
                     <div className="p-6">
                        <div className="flex items-center gap-3 text-red-600 mb-4">
                            <div className="p-3 bg-red-100 rounded-full">
                                <AlertTriangle className="w-8 h-8" />
                            </div>
                            <h3 className="text-xl font-bold text-slate-900">¿Borrar todo?</h3>
                        </div>
                        
                        <p className="text-slate-600 leading-relaxed mb-6">
                            Se borrarán <strong>todos los datos generales</strong> de la obra (Cliente, Nº Obra, etc).<br/>
                            {state.items.length > 0 && (
                                <>También se eliminarán <strong>{state.items.length} partidas</strong> de la lista.<br/></>
                            )}
                            {state.projectInfo.isAveria && (
                                <>Se restablecerán los datos de <strong>Avería</strong>.<br/></>
                            )}
                            <br/>
                            Esta acción dejará la hoja completamente limpia.
                        </p>

                        <div className="flex gap-3 justify-end">
                            <button 
                                onClick={() => setShowClearDialog(false)}
                                className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-100 rounded transition-colors"
                            >
                                Cancelar
                            </button>
                            <button 
                                onClick={confirmClearAll}
                                className="px-5 py-2 bg-red-600 text-white font-bold rounded shadow hover:bg-red-700 transition-colors flex items-center gap-2"
                            >
                                <Trash2 className="w-4 h-4" />
                                Sí, borrar todo
                            </button>
                        </div>
                     </div>
                </div>
            </div>
        )}
        
        {/* DRAGGABLE CALCULATOR RENDER */}
        {showCalculator && <DraggableCalculator onClose={() => setShowCalculator(false)} />}

        </div>
      </div>

      {/* STATUS BAR */}
      <div className="bg-slate-50 border-t border-slate-300 px-4 py-2 text-sm text-slate-500 flex justify-between items-center shrink-0 select-none font-medium">
         <div className="flex gap-4 items-center">
            <span className="flex items-center gap-1"><Check className="w-4 h-4 text-emerald-500"/> Listo</span>
            <span>{state.masterItems.length} Refs</span>
            <span>{state.items.length} Filas</span>
            <span>{state.checkedRowIds.size} Marcadas</span>
            
            {/* SELECTION STATS DISPLAY */}
            {selectionStats && selectionStats.count > 0 && (
                <>
                    <div className="h-4 w-px bg-slate-300 mx-1"></div>
                    <div className="flex items-center gap-4 text-slate-700 font-semibold animate-in fade-in">
                        <span className="flex items-center gap-1">
                             Recuento: {selectionStats.count}
                        </span>
                        {selectionStats.hasNumbers && (
                             <span className="flex items-center gap-1">
                                <Calculator className="w-3 h-3 text-slate-400" />
                                Suma: {formatNumber(selectionStats.sum)}
                             </span>
                        )}
                    </div>
                </>
            )}
         </div>
         <div className="font-mono opacity-50">
            v4.4 Professional
         </div>
      </div>
      
      {/* Styles */}
      <style>{`
        input[type=number]::-webkit-inner-spin-button, 
        input[type=number]::-webkit-outer-spin-button { 
            -webkit-appearance: none; 
            margin: 0; 
        }
        tr { page-break-inside: avoid; }
        thead { display: table-header-group; }
        .avoid-break, .avoid-break > * {
            page-break-inside: avoid !important;
            break-inside: avoid !important;
            -webkit-column-break-inside: avoid;
        }
      `}</style>
    </div>
  );
};

export default App;
