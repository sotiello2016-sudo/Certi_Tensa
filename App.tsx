import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import { 
  FileSpreadsheet, 
  Save,
  Upload, 
  Trash2, 
  Search,
  X,
  Plus,
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
  Info,
  HelpCircle,
  MousePointer2,
  Keyboard,
  Briefcase,
  Layers,
  Printer,
  Sigma,
  Database,
  FileOutput,
  Move,
  HardDrive,
  Settings,
  Lock,
  Globe,
  ShieldAlert,
  Wifi,
  WifiOff
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
const SECURITY_STORAGE_KEY = 'certipro_allowed_ips_v1';
const ADMIN_PASSWORD_STORAGE_KEY = 'certipro_admin_password_v1'; // New storage key for admin password

// Helper to generate N empty rows
const generateEmptyRows = (count: number): BudgetItem[] => {
    return Array.from({ length: count }).map((_, i) => ({
        id: `empty-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 9)}`,
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
    }));
};

// --- OPTIMIZED ROW COMPONENT ---
interface RowProps {
    item: BudgetItem;
    index: number;
    isChecked: boolean;
    isSelected: boolean;
    isInSelection: boolean;
    isAveria: boolean;
    searchTerm: string;
    activeSearch: { rowId: string, field: 'code' | 'description' } | null;
    editingCell: { rowId: string, field: string } | null;
    masterItems: BudgetItem[];
    
    // Actions
    onToggleCheck: (id: string) => void;
    onSetSelectedRow: (id: string) => void;
    onDragStart: (e: React.DragEvent, index: number) => void;
    onDragOver: (e: React.DragEvent) => void;
    onDragEnd: () => void;
    onDrop: (e: React.DragEvent, index: number) => void;
    onUpdateField: (id: string, field: keyof BudgetItem, value: string | number) => void;
    onUpdateQuantity: (id: string, value: number) => void;
    onFillRow: (id: string, masterItem: BudgetItem) => void;
    onAddEmpty: (id: string) => void;
    onDelete: (id: string) => void;
    onSetActiveSearch: (data: { rowId: string, field: 'code' | 'description' } | null) => void;
    onSetEditingCell: (data: { rowId: string, field: string } | null) => void;
    onCellMouseDown: (index: number, e: React.MouseEvent) => void;
    onCellMouseEnter: (index: number) => void;
    onInputFocus: () => void;
    onInputBlur: () => void;
    dropdownRef: React.RefObject<HTMLDivElement | null>;
}

const BudgetItemRow = React.memo(({
    item, index, isChecked, isSelected, isInSelection, isAveria, searchTerm, 
    activeSearch, editingCell, masterItems,
    onToggleCheck, onSetSelectedRow, onDragStart, onDragOver, onDragEnd, onDrop,
    onUpdateField, onUpdateQuantity, onFillRow, onAddEmpty, onDelete,
    onSetActiveSearch, onSetEditingCell, onCellMouseDown, onCellMouseEnter,
    onInputFocus, onInputBlur, dropdownRef
}: RowProps) => {

    const effectiveTotal = roundToTwo(item.currentQuantity * (isAveria ? (item.kFactor || 1) : 1) * item.unitPrice);

    const handleFocusSelect = (e: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement>) => {
        onInputFocus();
        const target = e.currentTarget;
        target.select();
        setTimeout(() => target.select(), 50);
    };

    const adjustTextareaHeight = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
        const target = e.target;
        requestAnimationFrame(() => {
            target.style.height = 'auto';
            target.style.height = `${target.scrollHeight}px`;
        });
    };

    const preventDrag = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
    };

    const getInlineSearchResults = (term: string) => {
        if (!term || term.length < 2) return [];
        const lowerTerm = term.toLowerCase();
        return masterItems.filter(i => 
          i.code.toLowerCase().includes(lowerTerm) || 
          i.description.toLowerCase().includes(lowerTerm)
        ).slice(0, 50);
    };

    const searchResults = (activeSearch?.rowId === item.id && (item.code.length > 0 || item.description.length > 1))
        ? getInlineSearchResults(activeSearch.field === 'code' ? item.code : item.description)
        : [];

    return (
        <tr 
            className={`group cursor-default transition-colors ${
                isSelected ? 'bg-blue-100 border-blue-200' : index % 2 === 0 ? 'bg-white' : 'bg-slate-50'
            }`}
            onClick={() => onSetSelectedRow(item.id)}
        >
            <td className="border-r border-slate-300 text-center bg-transparent align-top pt-4">
                <input 
                    type="checkbox" 
                    className="w-4 h-4 rounded border-slate-400 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                    checked={isChecked}
                    onChange={(e) => {
                        e.stopPropagation();
                        onToggleCheck(item.id);
                    }}
                    onClick={(e) => e.stopPropagation()}
                />
            </td>

            <td 
                className={`border-r border-slate-300 text-center text-base text-slate-400 select-none align-top pt-4 ${!searchTerm ? 'cursor-move hover:text-slate-600 active:text-slate-800' : 'cursor-default opacity-50'} ${isSelected ? 'bg-blue-200 text-blue-700 font-bold' : 'bg-slate-100'}`}
                draggable={!searchTerm}
                onDragStart={(e) => {
                    if (searchTerm) { e.preventDefault(); return; }
                    onDragStart(e, index);
                }}
                onDragOver={onDragOver}
                onDragEnd={onDragEnd}
                onDrop={(e) => onDrop(e, index)}
                title={searchTerm ? "Reordenar desactivado durante la búsqueda" : "Arrastrar para reordenar fila"}
            >
                {index + 1}
            </td>

            <td className="border-r border-slate-200 p-0 relative align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20">
                <textarea 
                    rows={1}
                    className="w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none font-sans text-base text-slate-800 focus:bg-transparent focus:outline-none text-justify resize-none overflow-hidden leading-relaxed relative z-0 caret-black cursor-default focus:cursor-text selection:bg-blue-600 selection:text-white"
                    value={item.code}
                    draggable={false}
                    onDragStart={preventDrag}
                    onChange={(e) => {
                        onUpdateField(item.id, 'code', e.target.value);
                        adjustTextareaHeight(e);
                    }}
                    onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.shiftKey) {
                            e.preventDefault();
                            e.currentTarget.blur();
                        }
                    }}
                    onFocus={(e) => {
                        handleFocusSelect(e);
                        onSetSelectedRow(item.id);
                        onSetActiveSearch({ rowId: item.id, field: 'code' });
                        adjustTextareaHeight(e);
                    }}
                    onBlur={onInputBlur}
                    placeholder="Buscar..."
                />
                {activeSearch?.rowId === item.id && activeSearch.field === 'code' && searchResults.length > 0 && (
                    <div 
                        ref={dropdownRef}
                        className="absolute left-0 w-[400px] bg-white border border-slate-300 shadow-xl z-50 max-h-60 overflow-y-auto rounded-sm top-full mt-1"
                        onMouseDown={(e) => e.preventDefault()}
                    >
                        {searchResults.map(res => (
                            <div key={res.id} onClick={() => onFillRow(item.id, res)} className="px-3 py-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-100 text-base flex gap-2">
                                <span className="font-bold font-sans text-slate-800">{res.code}</span>
                                <span className="truncate flex-1">{res.description}</span>
                            </div>
                        ))}
                    </div>
                )}
            </td>

            <td className="border-r border-slate-200 p-0 relative align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20">
                <textarea 
                    rows={1}
                    className="w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none text-base text-slate-800 font-sans focus:bg-transparent focus:outline-none text-justify resize-none overflow-hidden leading-relaxed relative z-0 caret-black cursor-default focus:cursor-text selection:bg-blue-600 selection:text-white"
                    value={item.description}
                    draggable={false}
                    onDragStart={preventDrag}
                    onChange={(e) => {
                        onUpdateField(item.id, 'description', e.target.value);
                        adjustTextareaHeight(e);
                    }}
                    onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.shiftKey) {
                            e.preventDefault();
                            e.currentTarget.blur();
                        }
                    }}
                    onFocus={(e) => {
                        handleFocusSelect(e);
                        onSetSelectedRow(item.id);
                        onSetActiveSearch({ rowId: item.id, field: 'description' });
                        adjustTextareaHeight(e);
                    }}
                    onBlur={onInputBlur}
                    placeholder="Descripción..."
                />
                {activeSearch?.rowId === item.id && activeSearch.field === 'description' && searchResults.length > 0 && (
                    <div 
                        ref={dropdownRef}
                        className="absolute left-0 w-full bg-white border border-slate-300 shadow-xl z-50 max-h-60 overflow-y-auto rounded-sm top-full mt-1"
                        onMouseDown={(e) => e.preventDefault()}
                    >
                        {searchResults.map(res => (
                            <div key={res.id} onClick={() => onFillRow(item.id, res)} className="px-3 py-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-100 text-base flex gap-2">
                                <span className="font-bold font-sans text-slate-500">{res.code}</span>
                                <span className="truncate flex-1">{res.description}</span>
                            </div>
                        ))}
                    </div>
                )}
            </td>

            <td className="border-r border-slate-200 p-0 relative bg-yellow-50/30 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20">
                <input 
                    type="number"
                    className="w-full px-3 py-3 text-center font-sans text-base text-slate-800 bg-transparent focus:bg-transparent focus:outline-none outline-none tabular-nums relative z-0 caret-black cursor-default focus:cursor-text selection:bg-blue-600 selection:text-white"
                    value={item.currentQuantity || ''}
                    draggable={false}
                    onDragStart={preventDrag}
                    placeholder="0"
                    onChange={(e) => onUpdateQuantity(item.id, parseFloat(e.target.value) || 0)}
                    onFocus={(e) => {
                        handleFocusSelect(e);
                        onSetSelectedRow(item.id);
                    }}
                    onKeyDown={(e) => { if(e.key === 'Enter') e.currentTarget.blur(); }}
                    onBlur={onInputBlur}
                />
            </td>

            {isAveria && (
                <td className="border-r border-slate-200 p-0 relative bg-red-50/30 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20">
                    <input 
                        type="number"
                        className="w-full px-3 py-3 text-center font-sans text-base text-red-700 font-bold bg-transparent focus:bg-transparent focus:outline-none outline-none tabular-nums relative z-0 caret-black cursor-default focus:cursor-text selection:bg-blue-600 selection:text-white"
                        value={item.kFactor || 1}
                        draggable={false}
                        onDragStart={preventDrag}
                        step="0.1"
                        onChange={(e) => onUpdateField(item.id, 'kFactor', parseFloat(e.target.value) || 0)}
                        onFocus={(e) => {
                            handleFocusSelect(e);
                            onSetSelectedRow(item.id);
                        }}
                        onKeyDown={(e) => { if(e.key === 'Enter') e.currentTarget.blur(); }}
                        onBlur={onInputBlur}
                    />
                </td>
            )}

            <td 
                className="border-r border-slate-200 p-0 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20"
                onClick={() => {
                    if (editingCell?.rowId !== item.id || editingCell?.field !== 'unitPrice') {
                        onSetEditingCell({ rowId: item.id, field: 'unitPrice' });
                    }
                }}
            >
                {editingCell?.rowId === item.id && editingCell?.field === 'unitPrice' ? (
                    <input 
                        autoFocus
                        type="number"
                        className="w-full px-3 py-3 text-right bg-transparent outline-none font-sans text-base text-slate-800 focus:outline-none tabular-nums relative z-20 caret-black cursor-default focus:cursor-text selection:bg-blue-600 selection:text-white"
                        value={item.unitPrice}
                        draggable={false}
                        onDragStart={preventDrag}
                        onChange={(e) => onUpdateField(item.id, 'unitPrice', parseFloat(e.target.value) || 0)}
                        onFocus={handleFocusSelect}
                        onBlur={() => {
                            onInputBlur();
                            onSetEditingCell(null);
                        }}
                        onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
                    />
                ) : (
                    <div 
                        className="w-full h-full px-3 py-3 text-right font-sans text-base text-slate-800 tabular-nums cursor-default outline-none"
                        tabIndex={0}
                        draggable={false}
                        onDragStart={preventDrag}
                        onFocus={() => onSetEditingCell({ rowId: item.id, field: 'unitPrice' })}
                    >
                        {formatCurrency(item.unitPrice)}
                    </div>
                )}
            </td>

            <td 
                className={`px-3 py-3 text-right font-sans text-base border-r border-slate-200 align-top tabular-nums cursor-cell select-none ${
                    isInSelection 
                        ? 'bg-blue-600 text-white font-bold' 
                        : effectiveTotal === 0 ? 'bg-red-100 text-red-600 font-medium' : 'text-slate-800 bg-slate-50/50'
                }`}
                data-col="total"
                onMouseDown={(e) => onCellMouseDown(index, e)}
                onMouseEnter={() => onCellMouseEnter(index)}
            >
                {formatCurrency(effectiveTotal)}
            </td>

            <td className="border-r border-slate-200 p-0 bg-slate-50/30 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20">
                <textarea 
                    rows={1}
                    className="w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none text-base text-slate-800 font-sans focus:bg-transparent focus:outline-none placeholder-slate-300 text-justify resize-none overflow-hidden leading-relaxed relative z-0 caret-black cursor-default focus:cursor-text selection:bg-blue-600 selection:text-white"
                    value={item.observations || ''}
                    draggable={false}
                    onDragStart={preventDrag}
                    onChange={(e) => {
                        onUpdateField(item.id, 'observations', e.target.value);
                        adjustTextareaHeight(e);
                    }}
                    onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.shiftKey) {
                            e.preventDefault();
                            e.currentTarget.blur();
                        }
                    }}
                    placeholder="..."
                    onFocus={(e) => {
                        handleFocusSelect(e);
                        onSetSelectedRow(item.id);
                        adjustTextareaHeight(e);
                    }}
                    onBlur={onInputBlur}
                />
            </td>

            <td className="border-l border-slate-200 p-0 text-center bg-slate-50 align-top">
                <div className="flex items-center justify-center pt-3 gap-1 opacity-20 group-hover:opacity-100 transition-opacity">
                    <button 
                        onClick={(e) => { e.stopPropagation(); onAddEmpty(item.id); }}
                        className="p-2 hover:bg-emerald-100 text-emerald-600 rounded transition-colors"
                        title="Insertar fila vacía debajo"
                    >
                        <Plus className="w-5 h-5" />
                    </button>
                    <button 
                        onClick={(e) => { e.stopPropagation(); onDelete(item.id); }}
                        className="p-2 hover:bg-red-100 text-red-600 rounded transition-colors"
                        title="Eliminar fila"
                    >
                        <Trash2 className="w-5 h-5" />
                    </button>
                </div>
            </td>
        </tr>
    );
}, (prev, next) => {
    if (prev.index !== next.index) return false;
    if (prev.item !== next.item) return false;
    if (prev.isChecked !== next.isChecked) return false;
    if (prev.isSelected !== next.isSelected) return false;
    if (prev.isInSelection !== next.isInSelection) return false;
    if (prev.isAveria !== next.isAveria) return false;
    if (prev.searchTerm !== next.searchTerm) return false;
    
    const prevEdit = prev.editingCell;
    const nextEdit = next.editingCell;
    if (prevEdit !== nextEdit) {
        if (!prevEdit || !nextEdit) return false;
        if (prevEdit.rowId === prev.item.id || nextEdit.rowId === next.item.id) return false;
    }

    const prevSearch = prev.activeSearch;
    const nextSearch = next.activeSearch;
    if (prevSearch !== nextSearch) {
        if (!prevSearch || !nextSearch) return false;
        if (prevSearch.rowId === prev.item.id || nextSearch.rowId === next.item.id) return false;
    }

    return true;
});

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
            const fullExpr = prevOperation + display;
            const sanitized = fullExpr.replace(/x/g, '*').replace(/÷/g, '/');
            const result = new Function('return ' + sanitized)();
            const formatted = String(parseFloat(result.toFixed(4)));
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

            <div className="p-4 bg-slate-800">
                <div className="text-slate-400 text-right text-xs h-4 mb-1 font-mono">{prevOperation}</div>
                <div className="text-white text-right text-3xl font-mono font-bold truncate tracking-widest bg-slate-900/50 p-2 rounded border border-slate-700/50">
                    {display}
                </div>
            </div>

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
  const [state, setState] = useState<AppState>(() => {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            const parsed = JSON.parse(saved);
            return {
                ...parsed,
                checkedRowIds: new Set(Array.isArray(parsed.checkedRowIds) ? parsed.checkedRowIds : []),
                isLoading: false
            };
        }
    } catch (e) {
        console.error("Failed to load autosave", e);
    }
    return {
        masterItems: [],
        items: generateEmptyRows(200),
        projectInfo: INITIAL_PROJECT_INFO,
        isLoading: false,
        checkedRowIds: new Set(),
        loadedFileName: undefined
    };
  });

  // --- SECURITY STATE ---
  const [currentIP, setCurrentIP] = useState<string>('');
  const [allowedIPs, setAllowedIPs] = useState<string[]>(() => {
      try {
          const saved = localStorage.getItem(SECURITY_STORAGE_KEY);
          return saved ? JSON.parse(saved) : [];
      } catch (e) { return []; }
  });
  const [adminPassword, setAdminPassword] = useState<string>(() => {
      try {
          const saved = localStorage.getItem(ADMIN_PASSWORD_STORAGE_KEY);
          // Changed default password here
          return saved || '32293229'; 
      } catch (e) { return '32293229'; } // Changed default password here
  });
  const [showSettingsLogin, setShowSettingsLogin] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [loginPasswordAttempt, setLoginPasswordAttempt] = useState('');
  const [loginErrorMessage, setLoginErrorMessage] = useState<string | null>(null); // New state for login error message
  const [ipInput, setIpInput] = useState('');

  // Fetch IP
  useEffect(() => {
    fetch('https://api.ipify.org?format=json')
        .then(response => response.json())
        .then(data => setCurrentIP(data.ip))
        .catch(err => {
            console.error("Error fetching IP", err);
            // Optional: Set a dummy IP or handle error state if needed
        });
  }, []);

  // Persist allowed IPs
  useEffect(() => {
    localStorage.setItem(SECURITY_STORAGE_KEY, JSON.stringify(allowedIPs));
  }, [allowedIPs]);

  // Persist admin password
  useEffect(() => {
    localStorage.setItem(ADMIN_PASSWORD_STORAGE_KEY, adminPassword);
  }, [adminPassword]);

  const isAccessAllowed = useMemo(() => {
    if (!currentIP) return false; // Loading IP or error
    if (allowedIPs.length === 0) return false; // No allowed IPs configured, block all
    return allowedIPs.includes(currentIP);
  }, [currentIP, allowedIPs]);

  const handleSettingsLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginPasswordAttempt === adminPassword) {
        setShowSettingsLogin(false);
        setLoginPasswordAttempt('');
        setLoginErrorMessage(null); // Clear login error on success
        setShowSettingsModal(true);
    } else {
        setLoginErrorMessage("Contraseña incorrecta"); // Set login error message
    }
  };

  const handleAddIP = () => {
      if (ipInput && !allowedIPs.includes(ipInput)) {
          setAllowedIPs(prev => [...prev, ipInput.trim()]);
          setIpInput('');
      }
  };

  const handleRemoveIP = (ip: string) => {
      setAllowedIPs(prev => prev.filter(item => item !== ip));
  };
  
  const [searchTerm, setSearchTerm] = useState('');
  useEffect(() => {
     if (!state) return;
     try {
         const stateToSave = {
             ...state,
             checkedRowIds: Array.from(state.checkedRowIds || new Set()),
             isLoading: false
         };
         localStorage.setItem(STORAGE_KEY, JSON.stringify(stateToSave));
     } catch (e) {
         console.error("Failed to autosave", e);
     }
  }, [state]);

  const [history, setHistory] = useState<{ past: AppState[], future: AppState[] }>({
      past: [],
      future: []
  });

  const historySnapshot = useRef<AppState | null>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const exportBtnRef = useRef<HTMLDivElement>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const marginInputRef = useRef<HTMLInputElement>(null);
  const autoScrollSpeed = useRef<number>(0);
  
  const [selectedCellIndices, setSelectedCellIndices] = useState<Set<number>>(new Set());
  const [isSelectingCells, setIsSelectingCells] = useState(false);
  const selectionScrollSpeed = useRef<number>(0);
  
  const isSelectingCellsRef = useRef(false);
  const selectionAnchorRef = useRef<number | null>(null);
  const selectionSnapshotRef = useRef<Set<number>>(new Set());
  const draggedRowIndexRef = useRef<number | null>(null);
  const selectedCellIndicesRef = useRef<Set<number>>(new Set());

  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showCalculator, setShowCalculator] = useState(false);
  
  const [showProformaDialog, setShowProformaDialog] = useState(false);
  const [proformaMargin, setProformaMargin] = useState("0");
  const [showClearDialog, setShowClearDialog] = useState(false);
  const [showHelpDialog, setShowHelpDialog] = useState(false);
  const [draggedRowIndex, setDraggedRowIndex] = useState<number | null>(null);
  const [activeSearch, setActiveSearch] = useState<{ rowId: string, field: 'code' | 'description' } | null>(null);
  const [editingCell, setEditingCell] = useState<{ rowId: string, field: string } | null>(null);

  const filteredItems = useMemo(() => {
    if (!searchTerm.trim()) return state.items;
    const lower = searchTerm.toLowerCase();
    return state.items.filter(i => 
        i.code.toLowerCase().includes(lower) || 
        i.description.toLowerCase().includes(lower) || 
        (i.observations && i.observations.toLowerCase().includes(lower))
    );
  }, [state.items, searchTerm]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
        if (!isAccessAllowed) return; // Disable shortcuts if blocked

        if (e.key === 'Escape') {
            setSelectedCellIndices(new Set());
            selectedCellIndicesRef.current = new Set();
            setIsSelectingCells(false);
            isSelectingCellsRef.current = false;
        }
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
  }, [history, state, isAccessAllowed]);

  useEffect(() => {
    let animationFrameId: number;
    const scrollStep = () => {
        if (autoScrollSpeed.current !== 0 && tableContainerRef.current) {
            tableContainerRef.current.scrollTop += autoScrollSpeed.current;
        }
        if (selectionScrollSpeed.current !== 0 && tableContainerRef.current) {
             tableContainerRef.current.scrollTop += selectionScrollSpeed.current;
        }
        animationFrameId = requestAnimationFrame(scrollStep);
    };
    animationFrameId = requestAnimationFrame(scrollStep);
    return () => cancelAnimationFrame(animationFrameId);
  }, []);

  useEffect(() => {
      const handleWindowMouseMove = (e: MouseEvent) => {
          if (!isSelectingCells || !tableContainerRef.current) return;
          const { top, bottom } = tableContainerRef.current.getBoundingClientRect();
          const threshold = 60;
          if (e.clientY < top + threshold) {
              const intensity = (top + threshold - e.clientY) / threshold;
              selectionScrollSpeed.current = -10 * Math.max(0.1, intensity);
          } else if (e.clientY > bottom - threshold) {
              const intensity = (e.clientY - (bottom - threshold)) / threshold;
              selectionScrollSpeed.current = 10 * Math.max(0.1, intensity);
          } else {
              selectionScrollSpeed.current = 0;
          }
      };
      const handleWindowMouseUp = () => {
          if (isSelectingCellsRef.current) {
            setIsSelectingCells(false);
            isSelectingCellsRef.current = false;
            selectionScrollSpeed.current = 0;
          }
      };
      if (isSelectingCells) {
          window.addEventListener('mousemove', handleWindowMouseMove);
          window.addEventListener('mouseup', handleWindowMouseUp);
      } else {
          selectionScrollSpeed.current = 0;
      }
      return () => {
          window.removeEventListener('mousemove', handleWindowMouseMove);
          window.removeEventListener('mouseup', handleWindowMouseUp);
      };
  }, [isSelectingCells]);

  useEffect(() => {
    if (activeSearch && dropdownRef.current && tableContainerRef.current) {
        const timer = setTimeout(() => {
            if (!dropdownRef.current || !tableContainerRef.current) return;
            const dropdownRect = dropdownRef.current.getBoundingClientRect();
            const containerRect = tableContainerRef.current.getBoundingClientRect();
            const hiddenAmount = dropdownRect.bottom - containerRect.bottom;
            if (hiddenAmount > 0) {
                tableContainerRef.current.scrollBy({
                    top: hiddenAmount + 24,
                    behavior: 'smooth'
                });
            }
        }, 50);
        return () => clearTimeout(timer);
    }
  }, [activeSearch, state.items]); 

  useEffect(() => {
    if (showProformaDialog && marginInputRef.current) {
        setTimeout(() => marginInputRef.current?.select(), 50);
    }
  }, [showProformaDialog]);

  const saveHistory = (currentState: AppState) => {
      if (!currentState) return;
      setHistory(prev => {
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
      if (historySnapshot.current) {
          if (historySnapshot.current !== state) {
              const snapshot = historySnapshot.current;
              const dirtyState = state;
              setHistory(prev => ({
                  past: prev.past, 
                  future: [dirtyState, ...prev.future] 
              }));
              setState(snapshot);
              historySnapshot.current = snapshot;
              return;
          }
      }
      if (history.past.length === 0) return;
      const previous = history.past[history.past.length - 1];
      if (!previous) return;
      setHistory(prev => ({
          past: prev.past.slice(0, -1),
          future: [state, ...prev.future]
      }));
      setState(previous);
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

  const handleInputFocus = useCallback(() => {
      if (state) {
        historySnapshot.current = state;
      }
  }, [state]);

  const handleInputBlur = useCallback(() => {
      if (historySnapshot.current && state && historySnapshot.current !== state) {
           saveHistory(historySnapshot.current);
      }
      historySnapshot.current = null;
      setActiveSearch(null);
  }, [state]);

  const handleSaveProject = () => {
      if (!state) return;
      const backupData = {
          version: "1.0",
          timestamp: new Date().toISOString(),
          projectInfo: state.projectInfo,
          items: state.items,
          masterItems: state.masterItems,
          checkedRowIds: Array.from(state.checkedRowIds || []),
          loadedFileName: state.loadedFileName
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
                  alert("El archivo seleccionado no es un archivo JSON válido.");
                  return;
              }
              if (!data || typeof data !== 'object') {
                  alert("Formato de archivo corrupto.");
                  return;
              }
              if (!data.items && !data.projectInfo) {
                   alert("El archivo no contiene datos reconocibles.");
                   return;
              }
              if (state && state.items.length > 0) {
                  saveHistory(state);
              }
              const hasLocalMasterItems = state.masterItems && state.masterItems.length > 0;
              let loadedItems = Array.isArray(data.items) ? data.items : [];
              if (loadedItems.length < 200) {
                  const rowsNeeded = 200 - loadedItems.length;
                  loadedItems = [...loadedItems, ...generateEmptyRows(rowsNeeded)];
              }
              setState({
                  projectInfo: data.projectInfo || INITIAL_PROJECT_INFO,
                  items: loadedItems,
                  masterItems: hasLocalMasterItems 
                      ? state.masterItems 
                      : (Array.isArray(data.masterItems) ? data.masterItems : []), 
                  isLoading: false,
                  checkedRowIds: new Set(Array.isArray(data.checkedRowIds) ? data.checkedRowIds : []),
                  loadedFileName: hasLocalMasterItems ? state.loadedFileName : data.loadedFileName
              });
          } catch (error) {
              console.error(error);
              alert("Ocurrió un error inesperado al cargar el proyecto.");
          }
      };
      reader.readAsText(file);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
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
        saveHistory(currentState);
        setState(prev => prev ? ({ 
          ...prev, 
          masterItems: mappedItems, 
          isLoading: false,
          loadedFileName: file.name
        }) : prev);
      } catch (err) {
        console.error(err);
        alert("Error al leer Excel.");
        setState(prev => prev ? ({ ...prev, isLoading: false }) : prev);
      }
    };
    reader.readAsBinaryString(file);
  };

  const toggleRowCheck = useCallback((id: string) => {
      saveHistory(state);
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
  }, [state]);

  const handleClearAll = () => {
    const hasItems = state.items.length > 0;
    const hasAveria = state.projectInfo.isAveria;
    const hasInfo = !!(state.projectInfo.name || state.projectInfo.projectNumber || state.projectInfo.orderNumber || state.projectInfo.client);
    if (!hasItems && !hasAveria && !hasInfo) return;
    setShowClearDialog(true);
  };

  const confirmClearAll = () => {
    setHistory({ past: [], future: [] });
    historySnapshot.current = null;
    setState(prev => ({
        ...prev,
        items: generateEmptyRows(200),
        checkedRowIds: new Set(),
        projectInfo: {
            ...INITIAL_PROJECT_INFO,
            date: new Date().toISOString().split('T')[0],
            averiaDate: new Date().toISOString().split('T')[0]
        },
        loadedFileName: prev.loadedFileName
    }));
    setSelectedRowId(null);
    setActiveSearch(null);
    setShowClearDialog(false);
  };

  const exportToExcel = () => {
    if (!state) return;
    const { projectInfo, items } = state;
    const isAveria = projectInfo.isAveria;
    const validItems = items.filter(item => item.code.trim() !== '' || item.description.trim() !== '');
    const headerRows = [
        [isAveria ? "CERTIFICACIÓN DE OBRA (AVERÍA)" : "CERTIFICACIÓN DE OBRA"],
        [""],
        ["Denominación:", projectInfo.name],
        ["Nº Obra:", projectInfo.projectNumber],
        ["Nº Pedido:", projectInfo.orderNumber],
        ["Cliente:", projectInfo.client],
        ["Fecha:", projectInfo.date],
    ];
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
        [""],
        isAveria 
          ? ["Recurso", "Descripción", "Ud", "K", "Precio Unitario", "Importe Total", "Observaciones"]
          : ["Recurso", "Descripción", "Ud", "Precio Unitario", "Importe Total", "Observaciones"]
    );
    const dataRows = validItems.map(item => {
        const k = item.kFactor || 1;
        const total = roundToTwo(item.currentQuantity * (isAveria ? k : 1) * item.unitPrice);
        if (isAveria) {
            return [item.code, item.description, item.currentQuantity, k, roundToTwo(item.unitPrice), total, item.observations || ''];
        } else {
            return [item.code, item.description, item.currentQuantity, roundToTwo(item.unitPrice), total, item.observations || ''];
        }
    });
    const totalAmountExport = validItems.reduce((acc, curr) => {
        const k = isAveria ? (curr.kFactor || 1) : 1;
        return acc + roundToTwo(curr.currentQuantity * k * curr.unitPrice);
    }, 0);
    const totalRow = isAveria 
        ? ["", "TOTAL CERTIFICACIÓN", "", "", "", roundToTwo(totalAmountExport), ""]
        : ["", "TOTAL CERTIFICACIÓN", "", "", roundToTwo(totalAmountExport), ""];
    const finalData = [...headerRows, ...dataRows, totalRow];
    const ws = XLSX.utils.aoa_to_sheet(finalData);
    if (isAveria) {
        ws['!cols'] = [{ wch: 30 }, { wch: 50 }, { wch: 10 }, { wch: 5 },  { wch: 15 }, { wch: 15 }, { wch: 30 }];
    } else {
        ws['!cols'] = [{ wch: 30 }, { wch: 50 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 30 }];
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Certificación");
    XLSX.writeFile(wb, `Certificacion_${projectInfo.projectNumber || 'Obra'}.xlsx`);
    setShowExportMenu(false);
  };

  const handleProformaClick = () => {
    const selectedCount = state.checkedRowIds.size;
    if (selectedCount === 0) {
        alert("Por favor, marque al menos una casilla (tick) para generar la Proforma.");
        return;
    }
    setShowExportMenu(false);
    setProformaMargin("0");
    setShowProformaDialog(true);
  };

  const confirmProformaExport = () => {
      const margin = parseFloat(proformaMargin.replace(',', '.'));
      const finalMargin = isNaN(margin) ? 0 : margin;
      downloadPDF('landscape', 'proforma', finalMargin);
      setShowProformaDialog(false);
  };

  const downloadPDF = (orientation: 'portrait' | 'landscape', type: 'certification' | 'proforma', marginPercentage: number = 0) => {
    const items = type === 'proforma' 
      ? state.items.filter(item => state.checkedRowIds.has(item.id)) 
      : state.items.filter(item => item.code.trim() !== '' || item.description.trim() !== '');
    if (type === 'proforma' && items.length === 0) {
        alert("Por favor, marque al menos una casilla (tick) para generar la Proforma.");
        return;
    }
    let marginDivisor = 1;
    if (type === 'proforma' && marginPercentage !== 0) {
        marginDivisor = 1 + (marginPercentage / 100);
    }
    const { projectInfo } = state;
    const isAveria = projectInfo.isAveria;
    const doc = new jsPDF({ orientation: orientation, unit: 'mm', format: 'a4' });
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 12;
    const rightMargin = pageWidth - 14;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8);
    doc.setTextColor(50);
    doc.text("TENSA SA", rightMargin, 12, { align: "right" });
    doc.setFont("helvetica", "normal");
    doc.setTextColor(100);
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
    const title = type === 'proforma' 
        ? (isAveria ? "FACTURA PROFORMA (AVERÍA)" : "FACTURA PROFORMA") 
        : (isAveria ? "CERTIFICACIÓN DE OBRA (AVERÍA)" : "CERTIFICACIÓN DE OBRA");
    doc.setTextColor(0);
    doc.setFontSize(18);
    doc.setFont("helvetica", "bold");
    doc.text(title.toUpperCase(), margin, 15);
    doc.setFontSize(14);
    doc.text(projectInfo.name.toUpperCase(), margin, 24);
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
    yPos += 6;
    doc.setFont("helvetica", "bold");
    doc.text("Cliente:", margin, yPos);
    doc.setFont("helvetica", "normal");
    doc.text(projectInfo.client, margin + 18, yPos);
    yPos += 5;
    if (isAveria) {
        yPos += 4;
        doc.setDrawColor(220);
        doc.setLineWidth(0.1);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 6;
        const redColor: [number, number, number] = [153, 27, 27];
        doc.setFontSize(10);
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("Nº AVERÍA:", margin, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        doc.text(projectInfo.averiaNumber || '', margin + 22, yPos);
        const col2X = margin + 55;
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("FECHA AVERÍA:", col2X, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        doc.text(projectInfo.averiaDate ? new Date(projectInfo.averiaDate).toLocaleDateString() : '', col2X + 28, yPos);
        const col3X = margin + 110;
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("HORARIO:", col3X, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        const horario = projectInfo.averiaTiming === 'nocturna_finde' ? 'Avería nocturna/ Fin de Semana K=1,75' : 'Avería diurna K=1,25';
        doc.text(horario, col3X + 20, yPos);
        yPos += 6;
        doc.setTextColor(...redColor);
        doc.setFont("helvetica", "bold");
        doc.text("DESCRIPCIÓN:", margin, yPos);
        doc.setTextColor(0);
        doc.setFont("helvetica", "normal");
        const descText = projectInfo.averiaDescription || '';
        const textX = margin + 28;
        const maxTextWidth = pageWidth - textX - margin;
        doc.text(descText, textX, yPos, { maxWidth: maxTextWidth, align: 'justify' });
        const dims = doc.getTextDimensions(descText, { maxWidth: maxTextWidth });
        yPos += Math.max(dims.h, 5) + 4;
        doc.setDrawColor(0);
        doc.setLineWidth(0.5);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 8;
    } else {
        yPos += 4;
        doc.setDrawColor(0);
        doc.setLineWidth(0.5);
        doc.line(margin, yPos, pageWidth - margin, yPos);
        yPos += 8;
    }
    const tableBody = items.map(item => {
        const k = isAveria ? (item.kFactor || 1) : 1;
        const adjustedUnitPrice = item.unitPrice / marginDivisor;
        const total = roundToTwo(item.currentQuantity * k * adjustedUnitPrice);
        return [item.code, item.description, formatNumber(item.currentQuantity), ...(isAveria ? [formatNumber(k)] : []), formatCurrency(adjustedUnitPrice), formatCurrency(total), item.observations || ''];
    });
    const tableHead = ["RECURSO", "DESCRIPCIÓN", "UD", ...(isAveria ? ["K"] : []), "PRECIO", "IMPORTE", "OBSERVACIONES"];
    let importeColRightEdge = 0;
    autoTable(doc, {
        startY: yPos,
        head: [tableHead],
        body: tableBody,
        theme: 'plain',
        styles: { fontSize: 9, cellPadding: 3, valign: 'top', textColor: 20, overflow: 'linebreak' },
        headStyles: { fontStyle: 'bold', textColor: 0, lineWidth: { bottom: 0.1 }, lineColor: 0, valign: 'middle' },
        columnStyles: {
            0: { cellWidth: 45, fontStyle: 'normal' },
            1: { cellWidth: 'auto' },
            2: { cellWidth: 15, halign: 'center' },
            ...(isAveria ? {
                3: { cellWidth: 12, halign: 'center' },
                4: { cellWidth: 25, halign: 'right' },
                5: { cellWidth: 28, halign: 'right', fontStyle: 'normal' },
                6: { cellWidth: 35, halign: 'left' }
            } : {
                3: { cellWidth: 25, halign: 'right' },
                4: { cellWidth: 28, halign: 'right', fontStyle: 'normal' },
                5: { cellWidth: 35, halign: 'left' }
            })
        },
        didDrawCell: (data) => {
            const importeIndex = isAveria ? 5 : 4;
            if (data.column.index === importeIndex && data.section === 'body') {
                importeColRightEdge = data.cell.x + data.cell.width;
            }
        },
        didParseCell: (data) => {
             if (data.section === 'head') {
                const idx = data.column.index;
                const priceIdx = isAveria ? 4 : 3;
                const totalIdx = isAveria ? 5 : 4;
                const udIdx = 2;
                const kIdx = 3;
                if (idx === priceIdx || idx === totalIdx) data.cell.styles.halign = 'right';
                if (idx === udIdx || (isAveria && idx === kIdx)) data.cell.styles.halign = 'center';
            }
        }
    });
    const finalY = (doc as any).lastAutoTable.finalY;
    const totalAmount = items.reduce((acc, curr) => {
        const k = isAveria ? (curr.kFactor || 1) : 1;
        const adjustedUnitPrice = curr.unitPrice / marginDivisor;
        return acc + roundToTwo(curr.currentQuantity * k * adjustedUnitPrice);
    }, 0);
    if (finalY > pageHeight - 30) {
        doc.addPage();
        yPos = 20;
    } else {
        yPos = finalY + 5;
    }
    const lineStart = pageWidth - 80; 
    doc.setDrawColor(0);
    doc.setLineWidth(0.5);
    doc.line(lineStart, yPos, pageWidth - margin, yPos);
    yPos += 6;
    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    const totalValueX = importeColRightEdge > 0 ? importeColRightEdge : (pageWidth - margin - 35 - margin);
    doc.text("TOTAL:", totalValueX - 30, yPos, { align: 'right' });
    doc.text(formatCurrency(totalAmount), totalValueX, yPos, { align: 'right' });
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(9);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(100);
        doc.text(`Hoja ${i} de ${totalPages}`, pageWidth - margin, pageHeight - 10, { align: 'right' });
    }
    const fileName = `${type === 'proforma' ? 'Proforma' : 'Certificacion'}_${projectInfo.projectNumber || 'Obra'}.pdf`;
    doc.save(fileName);
    setShowExportMenu(false);
  };

  const fillRowWithMaster = (rowId: string, masterItem: BudgetItem) => {
    saveHistory(state);
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
                        unitPrice: masterItem.unitPrice,
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

  const addEmptyItem = (referenceId?: string) => {
    saveHistory(state);
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
        let insertIndex = newItems.length;
        if (referenceId) {
             const idx = newItems.findIndex(i => i.id === referenceId);
             if (idx !== -1) insertIndex = idx + 1;
        } else if (selectedRowId) {
             const idx = newItems.findIndex(i => i.id === selectedRowId);
             if (idx !== -1) insertIndex = idx + 1;
        }
        newItems.splice(insertIndex, 0, newItem);
        return { ...prev, items: newItems };
    });
    setTimeout(() => setSelectedRowId(newItem.id), 50);
  };

  const deleteItem = (id: string | null) => {
    if (!id) return;
    saveHistory(state);
    setState(prev => {
      if (!prev) return prev;
      const newItems = prev.items.filter(i => i.id !== id);
      const newChecks = new Set(prev.checkedRowIds);
      if (newChecks.has(id)) newChecks.delete(id);
      return { ...prev, items: newItems, checkedRowIds: newChecks };
    });
    if (selectedRowId === id) setSelectedRowId(null);
  };

  const moveRow = useCallback((fromIndex: number, toIndex: number) => {
      if (fromIndex === toIndex) return;
      setState(prev => {
          if (!prev) return prev;
          saveHistory(prev); 
          const newItems = [...prev.items];
          const [movedItem] = newItems.splice(fromIndex, 1);
          newItems.splice(toIndex, 0, movedItem);
          return { ...prev, items: newItems };
      });
  }, []);

  const updateQuantity = (id: string, qty: number) => {
    setState(prev => {
      if (!prev) return prev;
      return {
        ...prev,
        items: prev.items.map(i => i.id === id ? calculateItemTotals({ ...i, currentQuantity: qty }) : i)
      };
    });
  };
  
  const updateField = (id: string, field: keyof BudgetItem, value: string | number) => {
     setState(prev => {
        if (!prev) return prev;
        return {
            ...prev,
            items: prev.items.map(i => {
                if (i.id === id) {
                    const updatedItem = { ...i, [field]: value };
                    if (field === 'unitPrice' || field === 'kFactor') return calculateItemTotals(updatedItem);
                    return updatedItem;
                }
                return i;
            })
        };
     });
     if (field === 'code' || field === 'description') setActiveSearch({ rowId: id, field });
  };

  const updateProjectInfo = (field: keyof ProjectInfo, value: string | number | boolean) => {
    setState(prev => {
      if (!prev) return prev;
      if (field === 'isAveria') {
          const newIsAveria = value as boolean;
          const recalculatedItems = prev.items.map(item => {
               const k = newIsAveria ? (item.kFactor || 1) : 1; 
               const totalAmount = roundToTwo(item.currentQuantity * k * item.unitPrice);
               return { ...item, totalAmount };
          });
          return { ...prev, items: recalculatedItems, projectInfo: { ...prev.projectInfo, [field]: value } };
      }
      return { ...prev, projectInfo: { ...prev.projectInfo, [field]: value } };
    });
  };

  const handleCellMouseDown = useCallback((index: number, e: React.MouseEvent) => {
      if (e.button !== 0) return;
      setIsSelectingCells(true);
      isSelectingCellsRef.current = true;
      selectionAnchorRef.current = index;
      if (e.ctrlKey || e.metaKey) {
          const currentSelection = selectedCellIndicesRef.current;
          selectionSnapshotRef.current = new Set(currentSelection);
          const newSet = new Set(currentSelection);
          newSet.add(index);
          setSelectedCellIndices(newSet);
          selectedCellIndicesRef.current = newSet;
      } else {
          selectionSnapshotRef.current = new Set();
          const newSet = new Set([index]);
          setSelectedCellIndices(newSet);
          selectedCellIndicesRef.current = newSet;
      }
  }, []);

  const handleCellMouseEnter = useCallback((index: number) => {
      if (isSelectingCellsRef.current && selectionAnchorRef.current !== null) {
          const start = Math.min(selectionAnchorRef.current, index);
          const end = Math.max(selectionAnchorRef.current, index);
          const newSet = new Set(selectionSnapshotRef.current);
          for (let i = start; i <= end; i++) newSet.add(i);
          setSelectedCellIndices(newSet);
          selectedCellIndicesRef.current = newSet; 
      }
  }, []);

  const selectedSum = useMemo(() => {
      let sum = 0;
      selectedCellIndices.forEach(index => {
          const item = filteredItems[index];
          if (item) {
              const k = state.projectInfo.isAveria ? (item.kFactor || 1) : 1;
              sum += roundToTwo(item.currentQuantity * k * item.unitPrice);
          }
      });
      return sum;
  }, [selectedCellIndices, filteredItems, state.projectInfo.isAveria]);

  if (!state) return <div className="h-screen flex items-center justify-center bg-slate-50 text-slate-400">Cargando aplicación...</div>;

  const totalAmount = state.items.reduce((acc, curr) => {
      const k = state.projectInfo.isAveria ? (curr.kFactor || 1) : 1;
      return acc + roundToTwo(curr.currentQuantity * k * curr.unitPrice);
  }, 0);

  useEffect(() => {
    const handleGlobalMouseDown = (e: MouseEvent) => {
      if (exportBtnRef.current && !exportBtnRef.current.contains(e.target as Node)) setShowExportMenu(false);
      if (activeSearch && dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
          const target = e.target as HTMLElement;
          if (target.tagName !== 'INPUT' && target.tagName !== 'TEXTAREA') setActiveSearch(null);
      }
      const target = e.target as HTMLElement;
      if (!target.closest('td[data-col="total"]')) {
          setSelectedCellIndices(new Set());
          selectedCellIndicesRef.current = new Set();
      }
    };
    const handleWindowMouseUp = () => {
        if (isSelectingCellsRef.current) {
            setIsSelectingCells(false);
            isSelectingCellsRef.current = false;
            selectionScrollSpeed.current = 0;
        }
    };
    document.addEventListener('mousedown', handleGlobalMouseDown);
    window.addEventListener('mouseup', handleWindowMouseUp);
    return () => {
        document.removeEventListener('mousedown', handleGlobalMouseDown);
        window.removeEventListener('mouseup', handleWindowMouseUp);
    };
  }, [showExportMenu, activeSearch]);

  const handleDragStart = useCallback((e: React.DragEvent, index: number) => {
      if (searchTerm) { e.preventDefault(); return; }
      setDraggedRowIndex(index);
      draggedRowIndexRef.current = index;
      e.dataTransfer.effectAllowed = "move";
  }, [searchTerm]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
      if (searchTerm) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      if (tableContainerRef.current) {
          const { top, bottom } = tableContainerRef.current.getBoundingClientRect();
          const y = e.clientY;
          const threshold = 100;
          if (y < top + threshold) autoScrollSpeed.current = -15;
          else if (y > bottom - threshold) autoScrollSpeed.current = 15;
          else autoScrollSpeed.current = 0;
      }
  }, [searchTerm]);

  const handleDragEnd = useCallback(() => {
      setDraggedRowIndex(null);
      draggedRowIndexRef.current = null;
      autoScrollSpeed.current = 0;
  }, []);

  const handleDrop = useCallback((e: React.DragEvent, index: number) => {
      if (searchTerm) return;
      e.preventDefault();
      autoScrollSpeed.current = 0;
      const fromIndex = draggedRowIndexRef.current;
      if (fromIndex !== null) {
          moveRow(fromIndex, index);
          setDraggedRowIndex(null);
          draggedRowIndexRef.current = null;
      }
  }, [searchTerm, moveRow]);

  return (
    <div className="h-screen flex flex-col bg-slate-50 font-sans text-slate-900 overflow-hidden relative">
      <div className="bg-white border-b border-slate-300 flex flex-col shrink-0 z-30 shadow-sm relative">
         <div className="flex items-center justify-between px-6 py-3 border-b border-slate-100">
             <div className="flex items-center gap-3">
                 <div className="bg-slate-900 text-white p-1.5 rounded shadow-sm"><FileSpreadsheet className="w-5 h-5"/></div>
                 <div><h1 className="text-xl font-bold text-slate-900 leading-tight">TENSA SA</h1></div>
             </div>
             <div className="flex items-center gap-2">
                 <div className="flex items-center gap-2 mr-4 border-r border-slate-200 pr-4">
                     <button onClick={undo} onMouseDown={(e) => e.preventDefault()} disabled={!isAccessAllowed || (history.past.length === 0 && (!historySnapshot.current || historySnapshot.current === state))} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-blue-600 hover:border-blue-400 rounded-md shadow-sm disabled:opacity-40 disabled:cursor-not-allowed transition-all" title="Deshacer (Ctrl+Z)"><Undo className="w-4 h-4" /><span className="text-xs font-semibold hidden lg:inline">Deshacer</span></button>
                     <button onClick={redo} onMouseDown={(e) => e.preventDefault()} disabled={!isAccessAllowed || history.future.length === 0} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-blue-600 hover:border-blue-400 rounded-md shadow-sm disabled:opacity-40 disabled:cursor-not-allowed transition-all" title="Rehacer (Ctrl+Y)"><Redo className="w-4 h-4" /><span className="text-xs font-semibold hidden lg:inline">Rehacer</span></button>
                 </div>
                 <button onClick={() => setShowSettingsLogin(true)} className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded cursor-pointer text-sm font-medium transition-colors shadow-sm ml-2" title="Ajustes de Seguridad"><Settings className="w-4 h-4 text-slate-500" />Ajustes</button>
                 <label className={`flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer text-sm text-slate-700 font-medium transition-colors select-none ${!isAccessAllowed && "opacity-50 pointer-events-none"}`} title="Cargar proyecto guardado (.json)"><FolderOpen className="w-4 h-4 text-orange-500" />Cargar Trabajo<input type="file" className="hidden" accept=".json" onChange={handleLoadProject} onClick={(e) => (e.currentTarget.value = '')} disabled={!isAccessAllowed} /></label>
                 <button onClick={handleSaveProject} className={`flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer text-sm text-slate-700 font-medium transition-colors ${!isAccessAllowed && "opacity-50 pointer-events-none"}`} title="Guardar proyecto actual (.json)" disabled={!isAccessAllowed}><Save className="w-4 h-4 text-blue-500" />Guardar Trabajo</button>
                 <div className={`flex items-center rounded border ml-4 transition-colors ${!isAccessAllowed ? "opacity-50 pointer-events-none" : ""} ${state.masterItems.length > 0 ? "bg-emerald-100 border-emerald-300 text-emerald-800 shadow-sm" : "bg-white border-slate-300 text-slate-700 hover:bg-slate-50"}`}>
                     <label className="flex items-center gap-2 px-3 py-2 cursor-pointer text-sm font-medium grow select-none hover:bg-opacity-80 rounded-l" title={state.masterItems.length > 0 ? "Tabla de recursos cargada" : "Importar tabla de recursos desde Excel"}><Upload className={`w-4 h-4 ${state.masterItems.length > 0 ? "text-emerald-700" : "text-slate-500"}`} />{state.masterItems.length > 0 ? "Tabla Rec. (OK)" : "Importar Tabla Rec."}<input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload} onClick={(e) => (e.currentTarget.value = '')} disabled={!isAccessAllowed}/></label>
                     {state.masterItems.length > 0 && (
                        <button className="px-2 py-2 border-l border-emerald-200 hover:bg-emerald-200 text-emerald-700 rounded-r focus:outline-none relative group cursor-help" onClick={(e) => { e.preventDefault(); /* Removed alert */ }} aria-label="Información"><Info className="w-4 h-4" />
                            <div className="absolute right-0 top-full mt-3 w-72 bg-white p-0 rounded-lg shadow-[0_10px_40px_-10px_rgba(0,0,0,0.2)] opacity-0 group-hover:opacity-100 transition-all duration-300 pointer-events-none z-[100] border border-slate-100 ring-1 ring-slate-900/5 transform origin-top scale-95 group-hover:scale-100 text-left">
                                <div className="bg-slate-50 px-4 py-3 rounded-t-lg border-b border-slate-100 flex items-center gap-2"><FileSpreadsheet className="w-4 h-4 text-emerald-600" /><span className="font-bold text-slate-700 text-sm">Archivo de Recursos</span></div>
                                <div className="p-4"><div className="relative pl-3"><div className="absolute left-0 top-1.5 w-1 h-8 bg-emerald-500 rounded-full"></div><div className="flex flex-col gap-1"><span className="font-bold text-slate-800 text-xs uppercase tracking-wide">Excel Cargado</span><p className="text-sm text-slate-600 font-medium break-all leading-tight">{state.loadedFileName || 'Desconocido'}</p></div></div><div className="mt-4 pt-3 border-t border-slate-50 flex justify-between items-center"><span className="text-[10px] uppercase font-bold text-slate-400">Total Registros</span><span className="bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded text-xs font-bold border border-emerald-100">{state.masterItems.length} items</span></div></div>
                                <div className="absolute bottom-full right-2.5 -mb-2 w-4 h-4 bg-white border-t border-l border-slate-100 transform rotate-45 rounded-sm"></div>
                            </div>
                        </button>
                     )}
                 </div>
                 <button onClick={handleClearAll} disabled={(!isAccessAllowed) || (state.items.length === 0 && !state.projectInfo.isAveria && !state.projectInfo.name && !state.projectInfo.projectNumber && !state.projectInfo.orderNumber && !state.projectInfo.client)} className="flex items-center gap-2 px-3 py-2 bg-white border border-red-200 text-red-600 rounded hover:bg-red-50 disabled:opacity-50 disabled:cursor-not-allowed text-sm font-medium transition-colors ml-2" title="Limpiar todo"><Eraser className="w-4 h-4" />Limpiar</button>
                 <div className={`relative ${!isAccessAllowed ? "opacity-50 pointer-events-none" : ""}`} ref={exportBtnRef}>
                    <button className="flex items-center gap-2 px-3 py-2 bg-emerald-600 text-white rounded hover:bg-emerald-700 text-sm font-medium shadow-sm transition-colors" onClick={() => setShowExportMenu(!showExportMenu)} disabled={!isAccessAllowed}><Download className="w-4 h-4" />Exportar<ChevronDown className="w-3 h-3 opacity-70" /></button>
                    {showExportMenu && (
                        <div className="absolute right-0 mt-1 w-64 bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50 animate-in fade-in slide-in-from-top-2">
                             <button onClick={exportToExcel} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2"><FileSpreadsheet className="w-4 h-4 text-green-600" /> Excel (.xlsx)</button>
                             <div className="h-px bg-slate-100 my-1"></div>
                             <div className="px-4 py-1 text-xs font-bold text-slate-400 uppercase tracking-wider">PDF</div>
                             <button onClick={() => downloadPDF('landscape', 'certification')} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2"><FileText className="w-4 h-4 text-red-600" /> Exportación PDF</button>
                             <button onClick={handleProformaClick} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2"><CheckSquare className="w-4 h-4 text-blue-600" /> Factura Proforma (Selección)</button>
                        </div>
                    )}
                 </div>
                 <button onClick={() => setShowCalculator(!showCalculator)} className={`flex items-center gap-2 px-3 py-2 rounded text-sm font-medium transition-colors ml-2 shadow-sm ${!isAccessAllowed ? "opacity-50 pointer-events-none" : ""} ${showCalculator ? "bg-slate-700 text-white border border-slate-600" : "bg-white border border-slate-300 text-slate-700 hover:bg-slate-50"}`} title="Calculadora" disabled={!isAccessAllowed}><Calculator className="w-4 h-4" /></button>
                 <button onClick={() => setShowHelpDialog(true)} className={`flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded ml-2 transition-colors ${!isAccessAllowed ? "opacity-50 pointer-events-none" : ""}`} title="Ayuda" disabled={!isAccessAllowed}><HelpCircle className="w-4 h-4 text-purple-500" /></button>
             </div>
         </div>
         {isAccessAllowed && (
             <div className="px-6 py-4 bg-slate-50/50">
                 <div className="grid grid-cols-12 gap-x-6 gap-y-4">
                     <div className="col-span-6"><label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Denominación</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-2xl font-bold text-slate-800 pb-1 focus:bg-white transition-colors placeholder-slate-300" value={state.projectInfo.name} onChange={(e) => updateProjectInfo('name', e.target.value)} onFocus={handleInputFocus} onBlur={handleInputBlur} /></div>
                     <div className="col-span-2"><label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Nº Obra</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-sans text-slate-700 pb-1 focus:bg-white transition-colors" value={state.projectInfo.projectNumber} onChange={(e) => updateProjectInfo('projectNumber', e.target.value)} onFocus={handleInputFocus} onBlur={handleInputBlur} /></div>
                     <div className="col-span-2"><label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Nº Pedido</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-sans text-slate-700 pb-1 focus:bg-white transition-colors" value={state.projectInfo.orderNumber} onChange={(e) => updateProjectInfo('orderNumber', e.target.value)} onFocus={handleInputFocus} onBlur={handleInputBlur} /></div>
                     <div className="col-span-2 flex items-end"><label className="flex items-center gap-2 cursor-pointer select-none bg-red-50 px-3 py-2 rounded border border-red-100 hover:bg-red-100 transition-colors w-full"><input type="checkbox" className="w-5 h-5 rounded border-red-400 text-red-600 focus:ring-red-500" checked={state.projectInfo.isAveria || false} onChange={(e) => updateProjectInfo('isAveria', e.target.checked)} /><span className="font-bold text-red-700 uppercase">AVERÍA</span></label></div>
                     <div className="col-span-8"><label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Cliente</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-medium text-slate-700 pb-1 focus:bg-white transition-colors" value={state.projectInfo.client} onChange={(e) => updateProjectInfo('client', e.target.value)} onFocus={handleInputFocus} onBlur={handleInputBlur} /></div>
                     <div className="col-span-2"><label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Fecha</label><input type="date" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-medium text-slate-700 pb-1 focus:bg-white transition-colors" value={state.projectInfo.date} onChange={(e) => updateProjectInfo('date', e.target.value)} onFocus={handleInputFocus} onBlur={handleInputBlur} /></div>
                     <div className="col-span-2 text-right"><label className="block text-sm font-bold text-slate-400 uppercase tracking-wider mb-1">Total Acumulado</label><div className="text-3xl font-sans font-bold text-emerald-700 leading-none pb-1 tabular-nums">{formatCurrency(totalAmount)}</div></div>
                     {state.projectInfo.isAveria && (
                       <div className="col-span-12 bg-red-50 p-4 rounded border border-red-100 mt-2 animate-in fade-in slide-in-from-top-2 shadow-sm"><div className="flex items-center gap-2 mb-3 text-red-800 font-bold uppercase text-sm border-b border-red-200 pb-1"><AlertTriangle className="w-4 h-4" /> Detalles de la Avería</div><div className="grid grid-cols-12 gap-6"><div className="col-span-2"><label className="block text-xs font-bold text-red-600 uppercase mb-1">Nº Avería</label><div className="relative"><Hash className="w-4 h-4 absolute left-2 top-2.5 text-red-300" /><input type="text" autoFocus className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium" placeholder="Axxxxx..." value={state.projectInfo.averiaNumber || ''} onChange={(e) => updateProjectInfo('averiaNumber', e.target.value)} onBlur={handleInputBlur} onFocus={handleInputFocus} /></div></div><div className="col-span-2"><label className="block text-xs font-bold text-red-600 uppercase mb-1">Fecha Avería</label><div className="relative"><Calendar className="w-4 h-4 absolute left-2 top-2.5 text-red-300" /><input type="date" className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium" value={state.projectInfo.averiaDate || ''} onChange={(e) => updateProjectInfo('averiaDate', e.target.value)} onBlur={handleInputBlur} onFocus={handleInputFocus} /></div></div><div className="col-span-2"><div className="flex items-center gap-1 mb-1"><label className="block text-xs font-bold text-red-600 uppercase">Horario</label><div className="group relative z-10"><Info className="w-4 h-4 text-slate-400 hover:text-blue-500 transition-colors cursor-help" /><div className="absolute left-1/2 -translate-x-1/2 bottom-full mb-3 w-72 bg-white p-0 rounded-lg shadow-[0_10px_40px_-10px_rgba(0,0,0,0.2)] opacity-0 group-hover:opacity-100 transition-all duration-300 pointer-events-none z-[100] border border-slate-100 ring-1 ring-slate-900/5 transform origin-bottom scale-95 group-hover:scale-100"><div className="bg-slate-50 px-4 py-3 rounded-t-lg border-b border-slate-100 flex items-center gap-2"><Clock className="w-4 h-4 text-blue-500" /><span className="font-bold text-slate-700 text-sm">Horarios y Coeficientes</span></div><div className="p-4 space-y-4"><div className="relative pl-3"><div className="absolute left-0 top-1.5 w-1 h-8 bg-orange-400 rounded-full"></div><div className="flex justify-between items-baseline mb-1"><span className="font-bold text-slate-800 text-xs uppercase tracking-wide">Diurno</span><span className="bg-orange-100 text-orange-700 px-1.5 py-0.5 rounded text-[10px] font-bold border border-orange-200">K = 1,25</span></div><p className="text-xs text-slate-500 leading-relaxed">Lunes a Viernes laborables de <span className="font-semibold text-slate-700">07:00</span> a <span className="font-semibold text-slate-700">19:00h</span>.</p></div><div className="relative pl-3"><div className="absolute left-0 top-1.5 w-1 h-8 bg-indigo-500 rounded-full"></div><div className="flex justify-between items-baseline mb-1"><span className="font-bold text-slate-800 text-xs uppercase tracking-wide">Nocturno / Finde</span><span className="bg-indigo-100 text-indigo-700 px-1.5 py-0.5 rounded text-[10px] font-bold border border-indigo-200">K = 1,75</span></div><p className="text-xs text-slate-500 leading-relaxed">Resto de horas, fines de semana y festivos.</p></div></div><div className="absolute top-full left-1/2 -translate-x-1/2 -mt-2 w-4 h-4 bg-white border-r border-b border-slate-100 transform rotate-45 rounded-sm"></div></div></div></div><div className="relative"><Clock className="w-4 h-4 absolute left-2 top-2.5 text-red-300" /><select className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium appearance-none" value={state.projectInfo.averiaTiming || 'diurna'} onChange={(e) => updateProjectInfo('averiaTiming', e.target.value)} onBlur={handleInputBlur} onFocus={handleInputFocus}><option value="diurna">Diurna K=1,25</option><option value="nocturna_finde">Nocturna K=1,75</option></select><ChevronDown className="w-4 h-4 absolute right-2 top-2.5 text-red-300 pointer-events-none" /></div></div><div className="col-span-6"><label className="block text-xs font-bold text-red-600 uppercase mb-1">Descripción</label><div className="relative"><AlignLeft className="w-4 h-4 absolute left-2 top-2.5 text-red-300" /><textarea rows={2} className="w-full pl-8 pr-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-transparent font-medium resize-none" value={state.projectInfo.averiaDescription || ''} onChange={(e) => updateProjectInfo('averiaDescription', e.target.value)} onBlur={handleInputBlur} onFocus={handleInputFocus} /></div></div></div></div>
                     )}
                 </div>
             </div>
         )}
      </div>

      <div className="flex-1 overflow-hidden relative">
        {isAccessAllowed ? (
            <div className="flex flex-col h-full bg-white relative">
              <div className="flex items-center justify-between px-4 py-2 border-b border-slate-300 bg-slate-50 sticky top-0 z-20 h-12 shrink-0">
                 <div className="relative group"><Search className="w-4 h-4 absolute left-3 top-2.5 text-slate-400 group-focus-within:text-blue-500 transition-colors" /><input type="text" className="pl-9 pr-8 py-1.5 bg-white border border-slate-300 rounded-full text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-64 transition-all shadow-sm" placeholder="Buscar..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />{searchTerm && <button onClick={() => setSearchTerm('')} className="absolute right-2 top-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full p-0.5"><X className="w-3 h-3" /></button>}</div>
                 <div className="text-sm text-slate-500 font-medium">{searchTerm ? <span>Resultados: <span className="text-slate-900 font-bold">{filteredItems.length}</span></span> : !selectedRowId && <span className="text-slate-400 italic font-normal hidden md:inline">Seleccione una fila para editar</span>}</div>
              </div>
              <div ref={tableContainerRef} className="flex-1 overflow-auto relative bg-slate-100/50">
                <table className="w-full border-collapse text-base table-fixed min-w-[1050px] bg-white select-none">
                  <thead className="sticky top-0 z-10 bg-orange-100 text-orange-900 font-semibold border-b border-orange-200 shadow-sm"><tr className="text-base uppercase tracking-wider"><th className="w-12 text-center py-4 border-r border-orange-200 bg-orange-200 text-xs font-bold select-none text-orange-800">CCAA</th><th className="w-10 text-center py-4 border-r border-orange-200 bg-orange-200 text-sm select-none">#</th><th className="w-64 px-4 py-4 text-center border-r border-orange-200 font-bold">RECURSO</th><th className="w-[450px] px-4 py-4 text-center border-r border-orange-200 font-bold">DESCRIPCIÓN</th><th className="w-24 px-4 py-4 text-center border-r border-orange-200 bg-orange-50 text-orange-800 font-bold">UD</th>{state.projectInfo.isAveria && <th className="w-20 px-4 py-4 text-center border-r border-orange-200 bg-red-100 text-red-800 font-bold">K</th>}<th className="w-24 px-4 py-4 text-center border-r border-orange-200 font-bold">PRECIO</th><th className="w-28 px-4 py-4 text-center font-bold">TOTAL</th><th className="w-64 px-4 py-4 text-center border-r border-orange-200 border-l border-orange-200 bg-orange-50/50 font-bold">OBSERVACIONES</th><th className="w-24 px-3 py-4 text-center border-l border-orange-200 bg-orange-200 font-bold">ACCIONES</th></tr></thead>
                  <tbody className="divide-y divide-slate-200">{filteredItems.map((item, index) => (<BudgetItemRow key={item.id} item={item} index={index} isChecked={state.checkedRowIds.has(item.id)} isSelected={selectedRowId === item.id} isInSelection={selectedCellIndices.has(index)} isAveria={!!state.projectInfo.isAveria} searchTerm={searchTerm} activeSearch={activeSearch} editingCell={editingCell} masterItems={state.masterItems} onToggleCheck={toggleRowCheck} onSetSelectedRow={setSelectedRowId} onDragStart={handleDragStart} onDragOver={handleDragOver} onDragEnd={handleDragEnd} onDrop={handleDrop} onUpdateField={updateField} onUpdateQuantity={updateQuantity} onFillRow={fillRowWithMaster} onAddEmpty={addEmptyItem} onDelete={deleteItem} onSetActiveSearch={setActiveSearch} onSetEditingCell={setEditingCell} onCellMouseDown={handleCellMouseDown} onCellMouseEnter={handleCellMouseEnter} onInputFocus={handleInputFocus} onInputBlur={handleInputBlur} dropdownRef={dropdownRef} />))}</tbody>
                </table>
              </div>
              {showProformaDialog && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-sm animate-in fade-in duration-200"><div className="bg-white rounded-xl shadow-[0_20px_60px_-15px_rgba(0,0,0,0.3)] w-full max-w-sm overflow-hidden border border-slate-100"><div className="bg-slate-50/80 px-5 py-4 border-b border-slate-100 flex items-center justify-between"><div className="flex items-center gap-2.5"><div className="bg-blue-100 p-1.5 rounded-md"><CheckSquare className="w-4 h-4 text-blue-600"/></div><span className="font-bold text-slate-700 text-sm tracking-tight">Generar Factura Proforma</span></div><button onClick={() => setShowProformaDialog(false)} className="text-slate-400 hover:text-slate-600 hover:bg-slate-200/50 p-1 rounded-full"><X className="w-4 h-4"/></button></div><div className="p-6"><div className="relative pl-4 mb-6"><div className="absolute left-0 top-1 w-1 h-full max-h-[40px] bg-blue-500 rounded-full opacity-20"></div><p className="text-sm text-slate-600 font-medium">Margen de Beneficio</p><p className="text-xs text-slate-400 mt-1">Indique el porcentaje a descontar.</p></div><div className="mb-8"><div className="relative group"><div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none"><Percent className="h-5 w-5 text-slate-400 group-focus-within:text-blue-500" /></div><input ref={marginInputRef} type="number" className="block w-full pl-10 pr-12 py-3 bg-white border border-slate-200 rounded-lg text-slate-700 text-xl font-bold placeholder-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all shadow-sm" placeholder="0" value={proformaMargin} onChange={(e) => setProformaMargin(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && confirmProformaExport()} /><div className="absolute inset-y-0 right-0 pr-4 flex items-center pointer-events-none"><span className="text-slate-400 font-bold text-sm">%</span></div></div></div><div className="flex gap-3"><button onClick={() => setShowProformaDialog(false)} className="flex-1 px-4 py-2.5 text-slate-600 font-bold text-sm hover:bg-slate-50 rounded-lg border border-transparent hover:border-slate-200 transition-all">Cancelar</button><button onClick={confirmProformaExport} className="flex-1 px-4 py-2.5 bg-blue-600 text-white font-bold text-sm rounded-lg shadow-lg hover:bg-blue-700 transition-all flex items-center justify-center gap-2"><FileText className="w-4 h-4" />Generar PDF</button></div></div></div></div>
              )}
              {showClearDialog && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm animate-in fade-in"><div className="bg-white rounded-lg shadow-2xl w-full max-w-md overflow-hidden border-t-4 border-red-500"><div className="p-6"><div className="flex items-center gap-3 text-red-600 mb-4"><div className="p-3 bg-red-100 rounded-full"><AlertTriangle className="w-8 h-8" /></div><h3 className="text-xl font-bold text-slate-900">¿Borrar todo?</h3></div><p className="text-slate-600 mb-6">Se borrarán todos los datos. Esta acción dejará la hoja completamente limpia.</p><div className="flex gap-3 justify-end"><button onClick={() => setShowClearDialog(false)} className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-100 rounded">Cancelar</button><button onClick={confirmClearAll} className="px-5 py-2 bg-red-600 text-white font-bold rounded hover:bg-red-700 transition-colors flex items-center justify-center gap-2"><Trash2 className="w-4 h-4" />Sí, borrar todo</button></div></div></div></div>
              )}
            {showHelpDialog && (
                <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-md animate-in fade-in duration-300">
                  <div className="bg-white rounded-3xl shadow-[0_32px_128px_-16px_rgba(0,0,0,0.4)] w-full max-w-4xl overflow-hidden border border-slate-200 flex flex-col max-h-[92vh] transform transition-all scale-100">
                    
                    {/* Modern Light Header */}
                    <div className="bg-slate-50 px-8 py-7 flex items-center justify-between border-b border-slate-200 relative overflow-hidden">
                      <div className="absolute top-0 right-0 w-64 h-64 bg-indigo-500/5 rounded-full blur-3xl -mr-20 -mt-20"></div>
                      <div className="absolute bottom-0 left-0 w-48 h-48 bg-blue-500/5 rounded-full blur-3xl -ml-16 -mb-16"></div>
                      
                      <div className="flex items-center gap-4 relative z-10">
                        <div className="bg-gradient-to-br from-purple-500 to-indigo-600 p-3 rounded-2xl text-white shadow-lg shadow-purple-500/20">
                          <HelpCircle className="w-7 h-7" />
                        </div>
                        <div>
                          <h2 className="text-2xl font-black text-slate-900 tracking-tight">Guía de Uso CertiTensa</h2>
                        </div>
                      </div>
                      <button 
                        onClick={() => setShowHelpDialog(false)}
                        className="text-slate-400 hover:text-slate-900 p-2 rounded-xl hover:bg-slate-200 transition-all duration-200 relative z-10"
                      >
                        <X className="w-7 h-7" />
                      </button>
                    </div>
                    
                    <div className="flex-1 overflow-y-auto p-10 space-y-6 bg-slate-50/50">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm hover:shadow-md hover:border-indigo-300 transition-all group flex flex-col h-full">
                          <div className="flex items-center gap-3 mb-4">
                            <div className="bg-indigo-50 text-indigo-600 p-2.5 rounded-xl group-hover:bg-indigo-600 group-hover:text-white transition-all duration-300 shadow-sm">
                              <Database className="w-6 h-6" />
                            </div>
                            <h3 className="font-extrabold text-slate-800 text-lg">1. Carga de Datos</h3>
                          </div>
                          <p className="text-slate-600 leading-relaxed text-sm grow">
                            Antes de iniciar cualquier certificación, cargue el <strong>Excel de recursos de Iberdrola</strong>. Esto permite al sistema realizar búsquedas automáticas y autocompletar partidas de forma precisa.
                          </p>
                        </div>

                        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm hover:shadow-md hover:border-emerald-300 transition-all group flex flex-col h-full">
                          <div className="flex items-center gap-3 mb-4">
                            <div className="bg-emerald-50 text-emerald-600 p-2.5 rounded-xl group-hover:bg-emerald-600 group-hover:text-white transition-all duration-300 shadow-sm">
                              <FileOutput className="w-6 h-6" />
                            </div>
                            <h3 className="font-extrabold text-slate-800 text-lg">2. Exportación y Proformas</h3>
                          </div>
                          <p className="text-slate-600 leading-relaxed text-sm grow">
                            Exporte a <strong>PDF o EXCEL</strong> desde el menú superior. Para generar <strong>Facturas Proforma</strong>, marque las casillas en la columna <strong>CCAA</strong> e indique el margen de beneficio solicitado antes de emitir.
                          </p>
                        </div>

                        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm hover:shadow-md hover:border-orange-300 transition-all group flex flex-col h-full">
                          <div className="flex items-center gap-3 mb-4">
                            <div className="bg-orange-50 text-orange-600 p-2.5 rounded-xl group-hover:bg-orange-600 group-hover:text-white transition-all duration-300 shadow-sm">
                              <Move className="w-6 h-6" />
                            </div>
                            <h3 className="font-extrabold text-slate-800 text-lg">3. Organización de Filas</h3>
                          </div>
                          <p className="text-slate-600 leading-relaxed text-sm grow">
                            Personalice el orden de las partidas pulsando sobre el <strong>número de fila (#)</strong> y arrastrándola a la posición deseada. La numeración se actualizará de forma correlativa automáticamente.
                          </p>
                        </div>

                        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm hover:shadow-md hover:border-red-300 transition-all group flex flex-col h-full">
                          <div className="flex items-center gap-3 mb-4">
                            <div className="bg-red-50 text-red-600 p-2.5 rounded-xl group-hover:bg-red-600 group-hover:text-white transition-all duration-300 shadow-sm">
                              <AlertTriangle className="w-6 h-6" />
                            </div>
                            <h3 className="font-extrabold text-slate-800 text-lg">4. Módulo de Averías</h3>
                          </div>
                          <p className="text-slate-600 leading-relaxed text-sm grow">
                            Active el modo <strong>Avería</strong> para habilitar campos especiales (Nº avería, horarios y descripción). Se añadirá una columna para el coeficiente <strong>K</strong>, que ajustará el importe total de forma automática.
                          </p>
                        </div>

                        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm hover:shadow-md hover:border-purple-300 transition-all group flex flex-col h-full md:col-span-2">
                          <div className="flex items-center gap-3 mb-4">
                            <div className="bg-purple-50 text-purple-600 p-2.5 rounded-xl group-hover:bg-purple-600 group-hover:text-white transition-all duration-300 shadow-sm">
                              <HardDrive className="w-6 h-6" />
                            </div>
                            <h3 className="font-extrabold text-slate-800 text-lg">5. Persistencia de Datos (.JSON)</h3>
                          </div>
                          <p className="text-slate-600 leading-relaxed text-sm grow">
                            Utilice los botones <strong>Guardar Trabajo</strong> y <strong>Cargar Trabajo</strong> para descargar su progreso en un archivo .JSON. Esto le permite pausar su labor y retomarla cualquier día desde donde la dejó.
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="bg-white p-8 border-t border-slate-100 flex justify-end gap-4 items-center">
                      <button 
                        onClick={() => setShowHelpDialog(false)}
                        className="px-10 py-3.5 bg-slate-900 text-white font-black rounded-2xl hover:bg-slate-800 shadow-2xl shadow-slate-900/20 active:transform active:scale-95 transition-all flex items-center gap-3 text-sm"
                      >
                        COMPRENDIDO <Check className="w-5 h-5" />
                      </button>
                    </div>

                  </div>
                </div>
            )}
            {showCalculator && <DraggableCalculator onClose={() => setShowCalculator(false)} />}
            </div>
        ) : (
            <div className="flex flex-col items-center justify-center h-full bg-slate-100 text-center p-8">
                <div className="bg-white p-12 rounded-2xl shadow-xl border border-red-100 max-w-lg w-full">
                    <div className="mx-auto w-24 h-24 bg-red-100 text-red-600 rounded-full flex items-center justify-center mb-6">
                        <ShieldAlert className="w-12 h-12" />
                    </div>
                    <h2 className="text-3xl font-bold text-slate-900 mb-4">Acceso Denegado</h2>
                    <p className="text-slate-600 mb-8 leading-relaxed">
                        Esta aplicación está protegida. Su dirección IP pública no está autorizada para acceder al contenido.
                    </p>
                    <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 mb-8 flex items-center justify-center gap-3">
                        <Globe className="w-5 h-5 text-slate-400" />
                        <span className="font-mono text-lg font-bold text-slate-700">{currentIP || "Detectando IP..."}</span>
                    </div>
                    <p className="text-xs text-slate-400 uppercase tracking-wide font-bold">Contacte con el administrador o acceda a Ajustes</p>
                </div>
            </div>
        )}
      </div>

      <div className="bg-slate-50 border-t border-slate-300 px-4 py-2 text-sm text-slate-500 flex justify-between items-center shrink-0 select-none font-medium">
         <div className="flex gap-4 items-center"><span className="flex items-center gap-1"><Check className="w-4 h-4 text-emerald-500"/> Listo</span><span>{state.masterItems.length} Refs</span><span>{state.items.length} Filas</span><span>{state.checkedRowIds.size} Marcadas</span></div>
         {selectedSum > 0 && isAccessAllowed && (<div className="flex items-center gap-2 bg-blue-100 text-blue-800 px-3 py-1 rounded-full animate-in fade-in slide-in-from-bottom-2"><Sigma className="w-4 h-4" /><span className="uppercase text-xs font-bold tracking-wider">Suma Seleccionada:</span><span className="font-mono font-bold text-base">{formatCurrency(selectedSum)}</span></div>)}
         <div className="font-mono opacity-50 hidden md:block">CertiTensa v4.5</div>
      </div>

      {/* LOGIN MODAL */}
      {showSettingsLogin && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in">
              <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm overflow-hidden border border-slate-200">
                  <div className="bg-slate-900 p-6 flex flex-col items-center">
                      <div className="w-12 h-12 bg-white/10 rounded-full flex items-center justify-center text-white mb-3">
                          <Lock className="w-6 h-6" />
                      </div>
                      <h3 className="text-white font-bold text-lg">Seguridad</h3>
                  </div>
                  <form onSubmit={handleSettingsLogin} className="p-6">
                      <label className="block text-sm font-bold text-slate-600 mb-2">Contraseña de Administrador</label>
                      <input 
                        autoFocus
                        type="password" 
                        className="w-full px-4 py-3 bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all mb-3 text-center text-lg tracking-widest"
                        placeholder="••••"
                        value={loginPasswordAttempt}
                        onChange={(e) => setLoginPasswordAttempt(e.target.value)}
                        onFocus={() => setLoginErrorMessage(null)} // Clear error message when user starts typing
                      />
                      {loginErrorMessage && (
                          <div className="p-2 mb-4 bg-red-100 text-red-700 text-sm font-medium rounded-lg border border-red-200">
                              {loginErrorMessage}
                          </div>
                      )}
                      <div className="flex gap-3">
                          <button type="button" onClick={() => { setShowSettingsLogin(false); setLoginErrorMessage(null); }} className="flex-1 py-3 text-slate-600 font-bold hover:bg-slate-50 rounded-lg transition-colors">Cancelar</button>
                          <button type="submit" className="flex-1 py-3 bg-blue-600 text-white font-bold rounded-lg hover:bg-blue-700 transition-colors shadow-lg">Entrar</button>
                      </div>
                  </form>
              </div>
          </div>
      )}

      {/* SETTINGS MODAL */}
      {showSettingsModal && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in">
              <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[85vh]">
                  <div className="bg-white border-b border-slate-100 p-6 flex items-center justify-between sticky top-0 z-10">
                      <div className="flex items-center gap-3">
                          <div className="bg-slate-100 p-2 rounded-lg text-slate-700"><Settings className="w-6 h-6" /></div>
                          <div>
                              <h2 className="text-xl font-bold text-slate-900">Control de Acceso y Ajustes</h2>
                              <p className="text-sm text-slate-500">Gestione direcciones IP autorizadas.</p>
                          </div>
                      </div>
                      <button onClick={() => setShowSettingsModal(false)} className="text-slate-400 hover:text-slate-600 p-2 hover:bg-slate-50 rounded-full transition-colors"><X className="w-6 h-6" /></button>
                  </div>
                  
                  <div className="p-6 overflow-y-auto flex-1 bg-slate-50">
                      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-6 shadow-sm">
                          <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4">Añadir Nueva IP</h4>
                          <div className="flex gap-2 mb-3">
                              <input 
                                type="text" 
                                className="flex-1 px-4 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent font-mono text-sm"
                                placeholder="Ej: 192.168.1.1"
                                value={ipInput}
                                onChange={(e) => setIpInput(e.target.value)}
                                onKeyDown={(e) => e.key === 'Enter' && handleAddIP()}
                              />
                              <button onClick={handleAddIP} className="px-4 py-2 bg-emerald-600 text-white font-bold rounded-lg hover:bg-emerald-700 transition-colors flex items-center gap-2"><Plus className="w-4 h-4" /> Añadir</button>
                          </div>
                          <div className="flex items-center justify-between bg-blue-50 px-4 py-3 rounded-lg border border-blue-100">
                              <div className="flex items-center gap-2 text-blue-800">
                                  <Globe className="w-4 h-4" />
                                  <span className="text-sm font-medium">Su IP actual es: <strong>{currentIP}</strong></span>
                              </div>
                              <button 
                                onClick={() => { setIpInput(currentIP); handleAddIP(); }}
                                className="text-xs font-bold text-blue-600 hover:text-blue-800 hover:underline"
                              >
                                Autorizar mi IP
                              </button>
                          </div>
                      </div>

                      <div className="mb-6">
                          <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">IPs Autorizadas ({allowedIPs.length})</h4>
                          {allowedIPs.length === 0 ? (
                              <div className="text-center py-8 text-slate-400 bg-slate-100 rounded-xl border border-dashed border-slate-300">
                                  <WifiOff className="w-8 h-8 mx-auto mb-2 opacity-50" />
                                  <p>No hay IPs configuradas.</p>
                                  <p className="text-xs mt-1">Nadie puede acceder a la aplicación.</p>
                              </div>
                          ) : (
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  {allowedIPs.map(ip => (
                                      <div key={ip} className="flex items-center justify-between bg-white px-4 py-3 rounded-lg border border-slate-200 shadow-sm group">
                                          <div className="flex items-center gap-3">
                                              <div className={`w-2 h-2 rounded-full ${ip === currentIP ? 'bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.4)]' : 'bg-slate-300'}`}></div>
                                              <span className="font-mono text-slate-700 font-medium">{ip}</span>
                                              {ip === currentIP && <span className="text-[10px] font-bold bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded border border-emerald-200">TÚ</span>}
                                          </div>
                                          <button onClick={() => handleRemoveIP(ip)} className="text-slate-300 hover:text-red-500 transition-colors p-1"><Trash2 className="w-4 h-4" /></button>
                                      </div>
                                  ))}
                              </div>
                          )}
                      </div>
                  </div>
                  <div className="p-4 border-t border-slate-200 bg-white flex justify-end">
                      <button onClick={() => setShowSettingsModal(false)} className="px-6 py-2 bg-slate-900 text-white font-bold rounded-lg hover:bg-slate-800 transition-colors shadow-lg">Cerrar</button>
                  </div>
              </div>
          </div>
      )}

      <style>{`
        input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        tr { page-break-inside: avoid; }
        thead { display: table-header-group; }
        .avoid-break, .avoid-break > * { page-break-inside: avoid !important; break-inside: avoid !important; }
      `}</style>
    </div>
  );
};

export default App;