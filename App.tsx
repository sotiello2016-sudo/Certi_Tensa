import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import { 
  FileSpreadsheet, 
  Save,
  Upload, 
  Trash2, 
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
  WifiOff,
  Maximize,
  Minimize,
  Type,
  User,
  FlaskConical,
  Binary,
  Search,
  ChevronLeft,
  ChevronRight
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { BudgetItem, ProjectInfo, AppState, CertificationType } from './types';
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
  certificationType: 'iberdrola',
  isAveria: false,
  averiaNumber: "",
  averiaDate: new Date().toISOString().split('T')[0],
  averiaDescription: "",
  averiaTiming: "diurna"
};

const STORAGE_KEY = 'certipro_autosave_v1';
const SECURITY_STORAGE_KEY = 'certipro_allowed_ips_v1';
const ADMIN_PASSWORD_STORAGE_KEY = 'certipro_admin_password_v1';

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

interface RowProps {
    item: BudgetItem;
    index: number;
    isChecked: boolean;
    isSelected: boolean;
    isInSelection: boolean;
    isAveria: boolean;
    certificationType: CertificationType;
    activeSearch: { rowId: string, field: 'code' | 'description' } | null;
    editingCell: { rowId: string, field: string } | null;
    activeSearchCell?: {rowIndex: number, field: keyof BudgetItem, rowId: string};
    masterItems: BudgetItem[];
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
    item, index, isChecked, isSelected, isInSelection, isAveria, certificationType, 
    activeSearch, editingCell, activeSearchCell, masterItems,
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
        if (certificationType === 'others' || !term || term.length < 2) return [];
        const lowerTerm = term.toLowerCase();
        return masterItems.filter(i => 
          i.code.toLowerCase().includes(lowerTerm) || 
          i.description.toLowerCase().includes(lowerTerm)
        ).slice(0, 50);
    };

    const searchResults = (certificationType === 'iberdrola' && activeSearch?.rowId === item.id && (item.code.length > 0 || item.description.length > 1))
        ? getInlineSearchResults(activeSearch.field === 'code' ? item.code : item.description)
        : [];
    
    const getHighlightClass = (field: keyof BudgetItem) => {
        if (activeSearchCell?.rowId === item.id && activeSearchCell?.field === field) {
            return 'ring-4 ring-green-600 ring-inset relative z-10';
        }
        return '';
    };

    const getHighlightInputClass = (field: keyof BudgetItem) => {
         if (activeSearchCell?.rowId === item.id && activeSearchCell?.field === field) {
            return 'placeholder-slate-400 selection:bg-green-200 caret-black';
        }
        return 'selection:bg-blue-600 selection:text-white caret-black';
    }

    return (
        <tr 
            data-row-id={item.id}
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
                className={`border-r border-slate-300 text-center text-base text-slate-400 select-none align-top pt-4 grab-handle hover:text-slate-600 active:text-slate-800 ${isSelected ? 'bg-blue-200 text-blue-700 font-bold' : 'bg-slate-100'}`}
                draggable={true}
                onDragStart={(e) => onDragStart(e, index)}
                onDragOver={onDragOver}
                onDragEnd={onDragEnd}
                onDrop={(e) => onDrop(e, index)}
                title="Arrastrar para reordenar fila"
            >
                {index + 1}
            </td>

            {certificationType === 'iberdrola' && (
                <td className={`border-r border-slate-200 p-0 relative align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20 ${getHighlightClass('code')}`}>
                    <textarea 
                        rows={1}
                        className={`w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none font-sans text-base text-slate-800 focus:bg-transparent focus:outline-none text-justify resize-none overflow-hidden leading-relaxed relative z-0 ${getHighlightInputClass('code')}`}
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
            )}

            <td className={`border-r border-slate-200 p-0 relative align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20 ${getHighlightClass('description')}`}>
                <textarea 
                    rows={1}
                    className={`w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none text-base text-slate-800 font-sans focus:bg-transparent focus:outline-none text-justify resize-none overflow-hidden leading-relaxed relative z-0 ${getHighlightInputClass('description')}`}
                    value={item.description}
                    draggable={false}
                    onDragStart={preventDrag}
                    onChange={(e) => {
                        onUpdateField(item.id, 'description', e.target.value);
                        adjustTextareaHeight(e);
                    }}
                    onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.shiftKey) {
                            switch (e.key) { case 'Enter': e.preventDefault(); e.currentTarget.blur(); break; }
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
                {certificationType === 'iberdrola' && activeSearch?.rowId === item.id && activeSearch.field === 'description' && searchResults.length > 0 && (
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

            <td className={`border-r border-slate-200 p-0 relative bg-yellow-50/30 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20 ${getHighlightClass('currentQuantity')}`}>
                <input 
                    type="number"
                    className={`w-full px-3 py-3 text-center font-sans text-base text-slate-800 bg-transparent focus:bg-transparent focus:outline-none outline-none tabular-nums relative z-0 ${getHighlightInputClass('currentQuantity')}`}
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

            {isAveria && certificationType === 'iberdrola' && (
                <td className={`border-r border-slate-200 p-0 relative bg-red-50/30 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20 ${getHighlightClass('kFactor')}`}>
                    <input 
                        type="number"
                        className={`w-full px-3 py-3 text-center font-sans text-base text-red-700 font-bold bg-transparent focus:bg-transparent focus:outline-none outline-none tabular-nums relative z-0 ${getHighlightInputClass('kFactor')}`}
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
                className={`border-r border-slate-200 p-0 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20 ${getHighlightClass('unitPrice')}`}
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
                        className={`w-full px-3 py-3 text-right bg-transparent outline-none font-sans text-base text-slate-800 focus:outline-none tabular-nums relative z-20 ${getHighlightInputClass('unitPrice')}`}
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
                        className={`w-full h-full px-3 py-3 text-right font-sans text-base text-slate-800 tabular-nums cursor-default outline-none ${getHighlightInputClass('unitPrice')}`}
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
                className={`px-3 py-3 text-right font-sans text-base border-r border-slate-200 align-top tabular-nums selectable-cell select-none ${
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

            <td className={`border-r border-slate-200 p-0 bg-slate-50/30 align-top focus-within:ring-2 focus-within:ring-inset focus-within:ring-blue-500 focus-within:z-20 ${getHighlightClass('observations')}`}>
                <textarea 
                    rows={1}
                    className={`w-full h-full min-h-[56px] px-3 py-3 bg-transparent outline-none text-base text-slate-800 font-sans focus:bg-transparent focus:outline-none placeholder-slate-300 text-justify resize-none overflow-hidden leading-relaxed relative z-0 ${getHighlightInputClass('observations')}`}
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
    if (prev.certificationType !== next.certificationType) return false;
    if (prev.activeSearchCell !== next.activeSearchCell) {
      if (prev.activeSearchCell?.rowId === prev.item.id || next.activeSearchCell?.rowId === next.item.id) {
        return false;
      }
    }
    
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

const DraggableCalculator: React.FC<{ onClose: () => void }> = ({ onClose }) => {
    const containerRef = useRef<HTMLDivElement>(null);
    const draggingRef = useRef(false);
    const offsetRef = useRef({ x: 0, y: 0 });
    const posRef = useRef({ x: window.innerWidth - 440, y: 120 });

    const [display, setDisplay] = useState('0');
    const [prevOperation, setPrevOperation] = useState('');
    const [isNewNumber, setIsNewNumber] = useState(true);
    const [isScientific, setIsScientific] = useState(false);

    const handleMouseDown = (e: React.MouseEvent) => {
        draggingRef.current = true;
        offsetRef.current = {
            x: e.clientX - posRef.current.x,
            y: e.clientY - posRef.current.y
        };
        window.addEventListener('mousemove', handleMouseMove);
        window.addEventListener('mouseup', handleMouseUp);
    };

    const handleMouseMove = useCallback((e: MouseEvent) => {
        if (!draggingRef.current || !containerRef.current) return;
        const newX = e.clientX - offsetRef.current.x;
        const newY = e.clientY - offsetRef.current.y;
        posRef.current = { x: newX, y: newY };
        containerRef.current.style.transform = `translate3d(${newX}px, ${newY}px, 0)`;
    }, []);

    const handleMouseUp = useCallback(() => {
        draggingRef.current = false;
        window.removeEventListener('mousemove', handleMouseMove);
        window.removeEventListener('mouseup', handleMouseUp);
    }, [handleMouseMove]);

    useEffect(() => {
        if (containerRef.current) {
            containerRef.current.style.transform = `translate3d(${posRef.current.x}px, ${posRef.current.y}px, 0)`;
        }
    }, []);

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

    const handleFunction = (func: string) => {
        try {
            const currentVal = parseFloat(display);
            let result = 0;
            switch(func) {
                case 'sin': result = Math.sin(currentVal * Math.PI / 180); break;
                case 'cos': result = Math.cos(currentVal * Math.PI / 180); break;
                case 'tan': result = Math.tan(currentVal * Math.PI / 180); break;
                case 'log': result = Math.log10(currentVal); break;
                case 'ln': result = Math.log(currentVal); break;
                case 'sqrt': result = Math.sqrt(currentVal); break;
                case 'sqr': result = Math.pow(currentVal, 2); break;
                case 'exp': result = Math.exp(currentVal); break;
                case 'pi': setDisplay(Math.PI.toString()); return;
                case 'e': setDisplay(Math.E.toString()); return;
                default: return;
            }
            setDisplay(String(parseFloat(result.toFixed(8))));
            setIsNewNumber(true);
        } catch (e) {
            setDisplay('Error');
        }
    };

    const calculate = () => {
        try {
            const fullExpr = prevOperation + display;
            const sanitized = fullExpr.replace(/x/g, '*').replace(/÷/g, '/');
            const result = new Function('return ' + sanitized)();
            const formatted = String(parseFloat(result.toFixed(8)));
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

    const btnClass = "h-11 flex items-center justify-center rounded bg-slate-700 hover:bg-slate-600 text-white font-bold text-base transition-colors shadow-sm active:transform active:scale-95";
    const opClass = "h-11 flex items-center justify-center rounded bg-orange-500 hover:bg-orange-600 text-white font-bold text-lg transition-colors shadow-sm";
    const sciBtnClass = "h-11 flex items-center justify-center rounded bg-slate-600 hover:bg-slate-500 text-blue-200 font-bold text-sm transition-colors shadow-sm italic";

    return (
        <div 
            ref={containerRef}
            className={`fixed z-[200] ${isScientific ? 'w-[480px]' : 'w-80'} bg-slate-800 rounded-lg shadow-2xl border border-slate-600 overflow-hidden flex flex-col top-0 left-0 will-change-transform`}
            style={{ transition: 'width 0.3s ease-in-out' }}
        >
            <div 
                className="bg-slate-900 px-4 py-2 flex items-center justify-between cursor-move select-none border-b border-slate-700"
                onMouseDown={handleMouseDown}
            >
                <div className="flex items-center gap-3">
                    <div className="flex items-center gap-2 text-slate-300 font-bold text-sm">
                        <Calculator className="w-4 h-4" /> Calculadora {isScientific ? 'Científica' : ''}
                    </div>
                    <button 
                        onClick={() => setIsScientific(!isScientific)}
                        className={`text-[10px] uppercase font-black px-2 py-0.5 rounded border transition-colors ${isScientific ? 'bg-blue-600 border-blue-400 text-white' : 'bg-slate-800 border-slate-600 text-slate-400 hover:text-white'}`}
                    >
                        {isScientific ? 'Básica' : 'Científica'}
                    </button>
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
                <div className="text-white text-right text-2xl font-mono font-bold truncate tracking-widest bg-slate-900/50 p-2 rounded border border-slate-700/50 min-h-[48px] flex items-center justify-end">
                    {display}
                </div>
            </div>

            <div className="flex">
                {isScientific && (
                    <div className="grid grid-cols-3 gap-2 p-4 pt-0 bg-slate-800 border-r border-slate-700 w-48 animate-in slide-in-from-left-4 duration-300">
                        <button onClick={() => handleFunction('sin')} className={sciBtnClass}>sin</button>
                        <button onClick={() => handleFunction('cos')} className={sciBtnClass}>cos</button>
                        <button onClick={() => handleFunction('tan')} className={sciBtnClass}>tan</button>
                        <button onClick={() => handleFunction('log')} className={sciBtnClass}>log</button>
                        <button onClick={() => handleFunction('ln')} className={sciBtnClass}>ln</button>
                        <button onClick={() => handleFunction('sqrt')} className={sciBtnClass}>√</button>
                        <button onClick={() => handleFunction('sqr')} className={sciBtnClass}>x²</button>
                        <button onClick={() => handleFunction('exp')} className={sciBtnClass}>eˣ</button>
                        <button onClick={() => handleOperator('**')} className={sciBtnClass}>xʸ</button>
                        <button onClick={() => handleFunction('pi')} className={sciBtnClass}>π</button>
                        <button onClick={() => handleFunction('e')} className={sciBtnClass}>e</button>
                        <button onClick={() => handleNumber('(')} className={sciBtnClass}>(</button>
                        <button onClick={() => handleNumber(')')} className={`${sciBtnClass} col-span-3`}>)</button>
                    </div>
                )}
                <div className={`grid grid-cols-4 gap-2 p-4 pt-0 bg-slate-800 flex-1`}>
                    <button onClick={clear} className="col-span-3 h-11 flex items-center justify-center rounded bg-slate-600 hover:bg-slate-500 text-red-200 font-bold">AC</button>
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
                    <button onClick={calculate} className="h-11 flex items-center justify-center rounded bg-emerald-500 hover:bg-emerald-600 text-white font-bold text-lg shadow-sm"><Equal className="w-6 h-6"/></button>
                </div>
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
          return saved || '32293229'; 
      } catch (e) { return '32293229'; }
  });
  const [showSettingsLogin, setShowSettingsLogin] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [loginPasswordAttempt, setLoginPasswordAttempt] = useState('');
  const [loginErrorMessage, setLoginErrorMessage] = useState<string | null>(null);
  const [ipInput, setIpInput] = useState('');
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [showTypeMenu, setShowTypeMenu] = useState(false);

  const typeBtnRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onFsChange = () => setIsFullscreen(!!document.fullscreenElement);
    document.addEventListener('fullscreenchange', onFsChange);
    return () => {
      document.removeEventListener('fullscreenchange', onFsChange);
    };
  }, []);

  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => console.error(err));
    } else {
      if (document.exitFullscreen) document.exitFullscreen();
    }
  };

  useEffect(() => {
    fetch('https://api.ipify.org?format=json')
        .then(response => response.json())
        .then(data => setCurrentIP(data.ip))
        .catch(err => console.error(err));
  }, []);

  useEffect(() => {
    localStorage.setItem(SECURITY_STORAGE_KEY, JSON.stringify(allowedIPs));
  }, [allowedIPs]);

  useEffect(() => {
    localStorage.setItem(ADMIN_PASSWORD_STORAGE_KEY, adminPassword);
  }, [adminPassword]);

  const isAccessAllowed = useMemo(() => {
    if (!currentIP) return false; 
    if (allowedIPs.length === 0) return false; 
    return allowedIPs.includes(currentIP);
  }, [currentIP, allowedIPs]);

  const handleSettingsLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginPasswordAttempt === adminPassword) {
        setShowSettingsLogin(false);
        setLoginPasswordAttempt('');
        setLoginErrorMessage(null);
        setShowSettingsModal(true);
    } else {
        setLoginErrorMessage("Contraseña incorrecta");
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

  // --- General Search State ---
  const [searchQuery, setSearchQuery] = useState('');
  const [lastSearchedQuery, setLastSearchedQuery] = useState('');
  const [searchResults, setSearchResults] = useState<{rowIndex: number, field: keyof BudgetItem, rowId: string}[]>([]);
  const [currentMatchIndex, setCurrentMatchIndex] = useState(-1);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
        if (!isAccessAllowed) return; 
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

  // --- General Search Scroll Effect ---
  useEffect(() => {
    if (currentMatchIndex > -1 && searchResults[currentMatchIndex]) {
        const { rowId } = searchResults[currentMatchIndex];
        const rowElement = tableContainerRef.current?.querySelector(`tr[data-row-id="${rowId}"]`);
        if (rowElement) {
            rowElement.scrollIntoView({
                behavior: 'smooth',
                block: 'center',
            });
        }
    }
  }, [currentMatchIndex, searchResults]);

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
          version: "1.1",
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
              if (!jsonString) return;
              let data = JSON.parse(jsonString);
              if (state && state.items.length > 0) saveHistory(state);
              const currentCertType = state.projectInfo.certificationType;
              const hasLocalMasterItems = state.masterItems && state.masterItems.length > 0;
              let loadedItems = Array.isArray(data.items) ? data.items : [];
              if (loadedItems.length < 200) {
                  loadedItems = [...loadedItems, ...generateEmptyRows(200 - loadedItems.length)];
              }
              setState({
                  projectInfo: {
                      ...(data.projectInfo || INITIAL_PROJECT_INFO),
                      certificationType: currentCertType
                  },
                  items: loadedItems,
                  masterItems: hasLocalMasterItems ? state.masterItems : (Array.isArray(data.masterItems) ? data.masterItems : []), 
                  isLoading: false,
                  checkedRowIds: new Set(Array.isArray(data.checkedRowIds) ? data.checkedRowIds : []),
                  loadedFileName: hasLocalMasterItems ? state.loadedFileName : data.loadedFileName
              });
          } catch (error) {
              console.error(error);
          }
      };
      reader.readAsText(file);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setState(prev => ({ ...prev, isLoading: true }));
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        if (rawData.length < 2) {
            setState(prev => ({ ...prev, isLoading: false }));
            return;
        }
        let headers = rawData[0].map(h => String(h).toLowerCase());
        const findColIndex = (variants: string[]) => headers.findIndex(h => variants.some(v => h.includes(v)));
        let idxCode = findColIndex(['recurso', 'código', 'codigo', 'id', 'partida', 'code']);
        let idxDesc = findColIndex(['denominación', 'denominacion', 'descripción', 'descripcion', 'description', 'nombre']);
        let idxUnit = findColIndex(['unidad', 'unid', 'unit', 'medida']);
        let idxPrice = findColIndex(['precio', 'precio unitario', 'unit price', 'p.u.', 'price']);
        if (idxCode === -1 && idxDesc === -1) { idxCode = 0; idxDesc = 1; idxUnit = 2; idxPrice = 4; }
        const mappedItems: BudgetItem[] = [];
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || (!row[idxCode] && !row[idxDesc])) continue;
            const price = row[idxPrice] ? parseFloat(String(row[idxPrice]).replace(',', '.')) : 0;
            mappedItems.push({
                id: `master-${i}-${Date.now()}`,
                code: row[idxCode] ? String(row[idxCode]) : `R-${i}`,
                description: row[idxDesc] ? String(row[idxDesc]) : 'Sin denominación',
                unit: row[idxUnit] ? String(row[idxUnit]) : 'ud',
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
        saveHistory(state);
        setState(prev => ({ ...prev, masterItems: mappedItems, isLoading: false, loadedFileName: file.name }));
      } catch (err) {
        console.error(err);
        setState(prev => ({ ...prev, isLoading: false }));
      }
    };
    reader.readAsBinaryString(file);
  };

  const toggleRowCheck = useCallback((id: string) => {
      saveHistory(state);
      setState(prev => {
          const newSet = new Set(prev.checkedRowIds);
          newSet.has(id) ? newSet.delete(id) : newSet.add(id);
          return { ...prev, checkedRowIds: newSet };
      });
  }, [state]);

  const handleClearAll = () => setShowClearDialog(true);

  const confirmClearAll = () => {
    setHistory({ past: [], future: [] });
    historySnapshot.current = null;
    setState(prev => ({
        ...prev,
        items: generateEmptyRows(200),
        checkedRowIds: new Set(),
        projectInfo: { ...INITIAL_PROJECT_INFO, date: new Date().toISOString().split('T')[0], averiaDate: new Date().toISOString().split('T')[0] },
    }));
    setSelectedRowId(null);
    setActiveSearch(null);
    setShowClearDialog(false);
  };

  const exportToExcel = () => {
    if (!state) return;
    const { projectInfo, items } = state;
    const isAveria = projectInfo.isAveria && projectInfo.certificationType === 'iberdrola';
    const isIberdrola = projectInfo.certificationType === 'iberdrola';
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
          : isIberdrola 
            ? ["Recurso", "Descripción", "Ud", "Precio Unitario", "Importe Total", "Observaciones"]
            : ["Descripción", "Ud", "Precio Unitario", "Importe Total", "Observaciones"]
    );
    const dataRows = validItems.map(item => {
        const k = (isAveria ? (item.kFactor || 1) : 1);
        const total = roundToTwo(item.currentQuantity * k * item.unitPrice);
        if (isAveria) return [item.code, item.description, item.currentQuantity, k, roundToTwo(item.unitPrice), total, item.observations || ''];
        if (isIberdrola) return [item.code, item.description, item.currentQuantity, roundToTwo(item.unitPrice), total, item.observations || ''];
        return [item.description, item.currentQuantity, roundToTwo(item.unitPrice), total, item.observations || ''];
    });
    const totalAmountExport = validItems.reduce((acc, curr) => {
        const k = (isAveria ? (curr.kFactor || 1) : 1);
        return acc + roundToTwo(curr.currentQuantity * k * curr.unitPrice);
    }, 0);
    const totalLabelRow = isAveria ? ["", "", "", "", "TOTAL CERTIFICACIÓN", roundToTwo(totalAmountExport)] : ["", "", "TOTAL CERTIFICACIÓN", roundToTwo(totalAmountExport)];
    const finalData = [...headerRows, ...dataRows, [""], totalLabelRow];
    const ws = XLSX.utils.aoa_to_sheet(finalData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Certificación");
    XLSX.writeFile(wb, `Certificacion_${projectInfo.projectNumber || 'Obra'}.xlsx`);
    setShowExportMenu(false);
  };

  const downloadPDF = (orientation: 'portrait' | 'landscape', type: 'certification' | 'proforma', marginPercentage: number = 0) => {
    const items = type === 'proforma' 
      ? state.items.filter(item => state.checkedRowIds.has(item.id)) 
      : state.items.filter(item => item.code.trim() !== '' || item.description.trim() !== '');
    if (type === 'proforma' && items.length === 0) {
        alert("Por favor, marque al menos una casilla (tick) para generar la Proforma.");
        return;
    }
    const { projectInfo } = state;

    const formatDateDDMMYYYY = (dateString: string | undefined): string => {
        if (!dateString) return '';
        const parts = dateString.split('-');
        if (parts.length === 3) {
            return `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
        return dateString;
    };

    const isAveria = projectInfo.isAveria && projectInfo.certificationType === 'iberdrola';
    const isIberdrola = projectInfo.certificationType === 'iberdrola';
    const marginDivisor = type === 'proforma' && marginPercentage !== 0 ? (1 + marginPercentage / 100) : 1;
    const doc = new jsPDF({ orientation: orientation, unit: 'mm', format: 'a4' });
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 12;
    const rightMargin = pageWidth - 14;
    doc.setFont("helvetica", "bold").setFontSize(8).setTextColor(50);
    doc.text("TENSA SA", rightMargin, 12, { align: "right" });
    doc.setFont("helvetica", "normal").setTextColor(100);
    doc.text("Pol. Villamuriel de Cerrato", rightMargin, 16, { align: "right" });
    doc.text("C/España, parc.79", rightMargin, 20, { align: "right" });
    doc.text("Palencia, CP: 33429", rightMargin, 24, { align: "right" });
    doc.text("CIF: A33020074", rightMargin, 28, { align: "right" });
    doc.setFont("helvetica", "bold").setTextColor(150).text("FECHA", rightMargin, 36, { align: "right" });
    doc.setFontSize(12).setTextColor(0).text(formatDateDDMMYYYY(projectInfo.date), rightMargin, 41, { align: "right" });
    const title = type === 'proforma' 
        ? (isAveria ? "FACTURA PROFORMA (AVERÍA)" : "FACTURA PROFORMA") 
        : (isAveria ? "CERTIFICACIÓN DE OBRA (AVERÍA)" : "CERTIFICACIÓN DE OBRA");
    doc.setTextColor(0).setFontSize(18).setFont("helvetica", "bold").text(title.toUpperCase(), margin, 15);
    doc.setFontSize(14).text(projectInfo.name.toUpperCase(), margin, 24);
    let yPos = 35;
    doc.setFontSize(10).setFont("helvetica", "bold").text("Nº Obra:", margin, yPos);
    doc.setFont("helvetica", "normal").text(projectInfo.projectNumber, margin + 18, yPos);
    doc.setFont("helvetica", "bold").text("Nº Pedido:", margin + 60, yPos);
    doc.setFont("helvetica", "normal").text(projectInfo.orderNumber, margin + 82, yPos);
    yPos += 6;
    doc.setFont("helvetica", "bold").text("Cliente:", margin, yPos);
    doc.setFont("helvetica", "normal").text(projectInfo.client, margin + 18, yPos);
    yPos += 10;
    if (isAveria) {
        doc.setDrawColor(220).setLineWidth(0.1).line(margin, yPos, pageWidth - margin, yPos);
        yPos += 6;
        doc.setFontSize(10).setTextColor(153, 27, 27).setFont("helvetica", "bold").text("Nº AVERÍA:", margin, yPos);
        doc.setTextColor(0).setFont("helvetica", "normal").text(projectInfo.averiaNumber || '', margin + 22, yPos);
        doc.setTextColor(153, 27, 27).setFont("helvetica", "bold").text("FECHA AVERÍA:", margin + 55, yPos);
        doc.setTextColor(0).setFont("helvetica", "normal").text(formatDateDDMMYYYY(projectInfo.averiaDate), margin + 83, yPos);
        doc.setTextColor(153, 27, 27).setFont("helvetica", "bold").text("HORARIO:", margin + 110, yPos);
        const horario = projectInfo.averiaTiming === 'nocturna_finde' ? 'Nocturna K=1,75' : 'Diurna K=1,25';
        doc.setTextColor(0).setFont("helvetica", "normal").text(horario, margin + 130, yPos);
        yPos += 6;
        doc.setTextColor(153, 27, 27).setFont("helvetica", "bold").text("DESCRIPCIÓN:", margin, yPos);
        const descText = projectInfo.averiaDescription || '';
        const maxWidth = pageWidth - margin * 2 - 28;
        doc.setTextColor(0).setFont("helvetica", "normal");
        const wrappedDesc = doc.splitTextToSize(descText, maxWidth);
        doc.text(wrappedDesc, margin + 28, yPos);
        const textDim = doc.getTextDimensions(wrappedDesc);
        yPos += textDim.h + 8;
    }
    doc.setDrawColor(0).setLineWidth(0.5).line(margin, yPos, pageWidth - margin, yPos);
    yPos += 5;
    const tableHead = isAveria 
        ? ["RECURSO", "DESCRIPCIÓN", "UD", "K", "PRECIO", "IMPORTE", "OBSERVACIONES"]
        : isIberdrola 
            ? ["RECURSO", "DESCRIPCIÓN", "UD", "PRECIO", "IMPORTE", "OBSERVACIONES"]
            : ["DESCRIPCIÓN", "UD", "PRECIO", "IMPORTE", "OBSERVACIONES"];
    const tableBody = items.map(item => {
        const k = isAveria ? (item.kFactor || 1) : 1;
        const adjPrice = item.unitPrice / marginDivisor;
        const total = roundToTwo(item.currentQuantity * k * adjPrice);
        if (isAveria) return [item.code, item.description, formatNumber(item.currentQuantity), formatNumber(k), formatCurrency(adjPrice), formatCurrency(total), item.observations || ''];
        if (isIberdrola) return [item.code, item.description, formatNumber(item.currentQuantity), formatCurrency(adjPrice), formatCurrency(total), item.observations || ''];
        return [item.description, formatNumber(item.currentQuantity), formatCurrency(adjPrice), formatCurrency(total), item.observations || ''];
    });
    let importeColRightEdge = 0;
    autoTable(doc, {
        startY: yPos,
        head: [tableHead],
        body: tableBody,
        theme: 'plain',
        rowPageBreak: 'avoid',
        styles: { fontSize: 9, cellPadding: 3, valign: 'top', textColor: 20 },
        headStyles: { fontStyle: 'bold', textColor: 0, lineWidth: { bottom: 0.1 }, lineColor: 0 },
        columnStyles: isAveria ? {
            0: { cellWidth: 55 }, 1: { cellWidth: 'auto' }, 2: { cellWidth: 15, halign: 'center' }, 3: { cellWidth: 12, halign: 'center' }, 4: { cellWidth: 28, halign: 'right' }, 5: { cellWidth: 32, halign: 'right' }, 6: { cellWidth: 35 }
        } : isIberdrola ? {
            0: { cellWidth: 55 }, 1: { cellWidth: 'auto' }, 2: { cellWidth: 15, halign: 'center' }, 3: { cellWidth: 28, halign: 'right' }, 4: { cellWidth: 32, halign: 'right' }, 5: { cellWidth: 35 }
        } : {
            0: { cellWidth: 'auto' }, 1: { cellWidth: 18, halign: 'center' }, 2: { cellWidth: 30, halign: 'right' }, 3: { cellWidth: 35, halign: 'right' }, 4: { cellWidth: 45 }
        },
        didParseCell: (data) => {
          if (data.section === 'head') {
            const idx = data.column.index;
            if (isAveria) {
              if (idx === 2 || idx === 3) data.cell.styles.halign = 'center';
              if (idx === 4 || idx === 5) data.cell.styles.halign = 'right';
            } else if (isIberdrola) {
              if (idx === 2) data.cell.styles.halign = 'center';
              if (idx === 3 || idx === 4) data.cell.styles.halign = 'right';
            } else {
              if (idx === 1) data.cell.styles.halign = 'center';
              if (idx === 2 || idx === 3) data.cell.styles.halign = 'right';
            }
          }
        },
        didDrawCell: (data) => {
            const totalIdx = isAveria ? 5 : isIberdrola ? 4 : 3;
            if (data.column.index === totalIdx && data.section === 'body') importeColRightEdge = data.cell.x + data.cell.width;
        }
    });
    const lastY = (doc as any).lastAutoTable.finalY;
    const totalAmount = items.reduce((acc, curr) => {
        const k = isAveria ? (curr.kFactor || 1) : 1;
        return acc + roundToTwo(curr.currentQuantity * k * (curr.unitPrice / marginDivisor));
    }, 0);
    yPos = lastY + 10;
    if (yPos > pageHeight - 30) { doc.addPage(); yPos = 20; }
    doc.setFontSize(12).setFont("helvetica", "bold");
    doc.text("TOTAL:", importeColRightEdge - 40, yPos, { align: 'right' });
    doc.text(formatCurrency(totalAmount), importeColRightEdge, yPos, { align: 'right' });
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(10);
        doc.setTextColor(150);
        doc.setFont("helvetica", "normal");
        doc.text(`hoja ${i} de ${totalPages}`, pageWidth - margin, pageHeight - 10, { align: 'right' });
    }
    doc.save(`${type === 'proforma' ? 'Proforma' : 'Certificacion'}_${projectInfo.projectNumber || 'Obra'}.pdf`);
    setShowExportMenu(false);
  };

  const confirmProformaExport = () => {
    const margin = parseFloat(proformaMargin) || 0;
    downloadPDF('landscape', 'proforma', margin);
    setShowProformaDialog(false);
  };

  const fillRowWithMaster = (rowId: string, masterItem: BudgetItem) => {
    saveHistory(state);
    setState(prev => ({
        ...prev,
        items: prev.items.map(i => {
            if (i.id === rowId) {
                return { ...i, code: masterItem.code, description: masterItem.description, unit: masterItem.unit, unitPrice: masterItem.unitPrice, totalAmount: roundToTwo(i.currentQuantity * (prev.projectInfo.isAveria ? (i.kFactor || 1) : 1) * masterItem.unitPrice) };
            }
            return i;
        })
    }));
    setActiveSearch(null);
  };

  const addEmptyItem = (referenceId?: string) => {
    saveHistory(state);
    const newItem: BudgetItem = { ...generateEmptyRows(1)[0], id: `manual-${Date.now()}-${Math.random()}` };
    setState(prev => {
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
    setState(prev => ({ ...prev, items: prev.items.filter(i => i.id !== id), checkedRowIds: new Set(Array.from(prev.checkedRowIds).filter(cid => cid !== id)) }));
    if (selectedRowId === id) setSelectedRowId(null);
  };

  const moveRow = useCallback((fromIndex: number, toIndex: number) => {
      if (fromIndex === toIndex) return;
      setState(prev => {
          saveHistory(prev); 
          const newItems = [...prev.items];
          const [movedItem] = newItems.splice(fromIndex, 1);
          newItems.splice(toIndex, 0, movedItem);
          return { ...prev, items: newItems };
      });
  }, []);

  const updateQuantity = (id: string, qty: number) => {
    setState(prev => ({
        ...prev,
        items: prev.items.map(i => i.id === id ? calculateItemTotals({ ...i, currentQuantity: qty, kFactor: prev.projectInfo.certificationType === 'others' ? 1 : i.kFactor }) : i)
    }));
  };
  
  const updateField = (id: string, field: keyof BudgetItem, value: string | number) => {
     setState(prev => ({
        ...prev,
        items: prev.items.map(i => i.id === id ? calculateItemTotals({ ...i, [field]: value }) : i)
     }));
     if ((field === 'code' || field === 'description') && state.projectInfo.certificationType === 'iberdrola') setActiveSearch({ rowId: id, field });
  };

  const updateProjectInfo = (field: keyof ProjectInfo, value: any) => {
    setState(prev => {
      const newInfo = { ...prev.projectInfo, [field]: value };
      
      let newItems = prev.items;

      // Only recalculate items if certificationType or isAveria changes, as these affect item calculations
      if (field === 'certificationType') {
          // If switching to 'others', ensure isAveria is false
          if (value === 'others') {
              newInfo.isAveria = false;
          }
          // Recalculate kFactor-dependent totals for all items
          newItems = prev.items.map(item => calculateItemTotals({ ...item, kFactor: newInfo.isAveria ? item.kFactor : 1 }));
      } else if (field === 'isAveria') {
          // Recalculate kFactor-dependent totals for all items when isAveria toggles
          newItems = prev.items.map(item => calculateItemTotals({ ...item, kFactor: newInfo.isAveria ? (item.kFactor || 1) : 1 }));
      }
      // For other fields (like averiaDescription), newItems remains prev.items, avoiding unnecessary map operations

      return { ...prev, items: newItems, projectInfo: newInfo };
    });
  };

  const handleCellMouseDown = useCallback((index: number, e: React.MouseEvent) => {
      if (e.button !== 0) return;
      setIsSelectingCells(true);
      isSelectingCellsRef.current = true;
      selectionAnchorRef.current = index;
      const newSet = (e.ctrlKey || e.metaKey) ? new Set<number>(selectedCellIndicesRef.current) : new Set<number>();
      newSet.add(index);
      setSelectedCellIndices(newSet);
      selectedCellIndicesRef.current = newSet;
      selectionSnapshotRef.current = new Set<number>(newSet);
  }, []);

  const handleCellMouseEnter = useCallback((index: number) => {
      if (isSelectingCellsRef.current && selectionAnchorRef.current !== null) {
          const start = Math.min(selectionAnchorRef.current, index);
          const end = Math.max(selectionAnchorRef.current, index);
          const newSet = new Set<number>(selectionSnapshotRef.current);
          for (let i = start; i <= end; i++) newSet.add(i);
          setSelectedCellIndices(newSet);
          selectedCellIndicesRef.current = newSet; 
      }
  }, []);

  const handleDragEnd = useCallback(() => {
      setDraggedRowIndex(null);
      draggedRowIndexRef.current = null;
      autoScrollSpeed.current = 0;
  }, []);

  const goToPreviousMatch = () => {
    if (searchResults.length > 0) {
        setCurrentMatchIndex(prev => (prev - 1 + searchResults.length) % searchResults.length);
    }
  };

  const startOrContinueSearch = () => {
    const currentQuery = searchQuery.trim().toLowerCase();

    if (!currentQuery) {
      setSearchResults([]);
      setCurrentMatchIndex(-1);
      setLastSearchedQuery('');
      return;
    }

    const results: { rowIndex: number; field: keyof BudgetItem; rowId: string }[] = [];
    const isIberdrolaSearch = state.projectInfo.certificationType === 'iberdrola';
    
    const searchableFields: (keyof BudgetItem)[] = isIberdrolaSearch
        ? ['code', 'description', 'currentQuantity', 'kFactor', 'unitPrice', 'observations']
        : ['description', 'currentQuantity', 'unitPrice', 'observations'];

    state.items.forEach((item, index) => {
      for (const field of searchableFields) {
        const value = item[field];
        if (value !== null && value !== undefined && String(value).toLowerCase().includes(currentQuery)) {
          results.push({ rowIndex: index, field, rowId: item.id });
        }
      }
    });

    let nextIndex;
    if (results.length === 0) {
      nextIndex = -1;
    } else if (currentQuery !== lastSearchedQuery) {
      nextIndex = 0;
    } else {
      nextIndex = (currentMatchIndex + 1) % results.length;
    }
    
    setSearchResults(results);
    setCurrentMatchIndex(nextIndex);
    setLastSearchedQuery(currentQuery);
  };

  const handleSearchQueryChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newQuery = e.target.value;
    setSearchQuery(newQuery);
    if (newQuery.trim() === '') {
        setSearchResults([]);
        setCurrentMatchIndex(-1);
        setLastSearchedQuery('');
    }
  };

  const handleSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      startOrContinueSearch();
    }
  };

  const selectedSum = useMemo(() => {
      let sum = 0;
      selectedCellIndices.forEach(index => {
          const item = state.items[index];
          if (item) {
              const k = (state.projectInfo.certificationType === 'iberdrola' && state.projectInfo.isAveria) ? (item.kFactor || 1) : 1;
              sum += roundToTwo(item.currentQuantity * k * item.unitPrice);
          }
      });
      return sum;
  }, [selectedCellIndices, state.items, state.projectInfo.isAveria, state.projectInfo.certificationType]);

  const totalAmount = state.items.reduce((acc, curr) => {
      const k = (state.projectInfo.certificationType === 'iberdrola' && state.projectInfo.isAveria) ? (curr.kFactor || 1) : 1;
      return acc + roundToTwo(curr.currentQuantity * k * curr.unitPrice);
  }, 0);

  useEffect(() => {
    const handleGlobalMouseDown = (e: MouseEvent) => {
      if (exportBtnRef.current && !exportBtnRef.current.contains(e.target as Node)) setShowExportMenu(false);
      if (typeBtnRef.current && !typeBtnRef.current.contains(e.target as Node)) setShowTypeMenu(false);
      if (activeSearch && dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
          const target = e.target as HTMLElement;
          if (target.tagName !== 'INPUT' && target.tagName !== 'TEXTAREA') setActiveSearch(null);
      }
      if (!(e.target as HTMLElement).closest('td[data-col="total"]')) {
          setSelectedCellIndices(new Set<number>());
          selectedCellIndicesRef.current = new Set<number>();
      }
    };
    document.addEventListener('mousedown', handleGlobalMouseDown);
    return () => document.removeEventListener('mousedown', handleGlobalMouseDown);
  }, [activeSearch]);

  const handleDragStart = useCallback((e: React.DragEvent, index: number) => {
      setDraggedRowIndex(index);
      draggedRowIndexRef.current = index;
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent) => {
      e.preventDefault();
      if (tableContainerRef.current) {
          const { top, bottom } = tableContainerRef.current.getBoundingClientRect();
          const threshold = 100;
          if (e.clientY < top + threshold) autoScrollSpeed.current = -15;
          else if (e.clientY > bottom - threshold) autoScrollSpeed.current = 15;
          else autoScrollSpeed.current = 0;
      }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent, index: number) => {
      e.preventDefault();
      autoScrollSpeed.current = 0;
      if (draggedRowIndexRef.current !== null) {
          moveRow(draggedRowIndexRef.current, index);
          setDraggedRowIndex(null);
          draggedRowIndexRef.current = null;
      }
  }, [moveRow]);

  if (!state) return <div className="h-screen flex items-center justify-center bg-slate-50 text-slate-400">Cargando aplicación...</div>;

  const isIberdrola = state.projectInfo.certificationType === 'iberdrola';
  const activeSearchCell = searchResults[currentMatchIndex];

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
                     <button onClick={undo} disabled={!isAccessAllowed || (history.past.length === 0 && (!historySnapshot.current || historySnapshot.current === state))} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-blue-600 rounded-md shadow-sm disabled:opacity-40 transition-all"><Undo className="w-4 h-4" /><span className="text-xs font-semibold hidden lg:inline">Deshacer</span></button>
                     <button onClick={redo} disabled={!isAccessAllowed || history.future.length === 0} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 hover:text-blue-600 rounded-md shadow-sm disabled:opacity-40 transition-all"><Redo className="w-4 h-4" /><span className="text-xs font-semibold hidden lg:inline">Rehacer</span></button>
                 </div>
                 <div className={`relative ${!isAccessAllowed && "opacity-50 pointer-events-none"}`} ref={typeBtnRef}>
                    <button className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 text-slate-700 rounded hover:bg-slate-50 text-sm font-medium shadow-sm transition-colors" onClick={() => setShowTypeMenu(!showTypeMenu)}>
                        <Layers className="w-4 h-4 text-indigo-500" />
                        Tipo de certificación
                        <ChevronDown className="w-3 h-3 opacity-70" />
                    </button>
                    {showTypeMenu && (
                        <div className="absolute left-0 mt-1 w-64 bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50">
                             <button onClick={() => { updateProjectInfo('certificationType', 'iberdrola'); setShowTypeMenu(false); }} className={`w-full text-left px-4 py-2 hover:bg-slate-50 text-sm flex items-center gap-2 ${isIberdrola ? 'text-indigo-600 font-bold' : 'text-slate-700'}`}><FileText className="w-4 h-4 text-slate-400" /> Certificación Iberdrola</button>
                             <button onClick={() => { updateProjectInfo('certificationType', 'others'); setShowTypeMenu(false); }} className={`w-full text-left px-4 py-2 hover:bg-slate-50 text-sm flex items-center gap-2 ${!isIberdrola ? 'text-indigo-600 font-bold' : 'text-slate-700'}`}><User className="w-4 h-4 text-slate-400" /> Certificación otros clientes</button>
                        </div>
                    )}
                 </div>
                 <label className={`flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer text-sm text-slate-700 font-medium transition-colors select-none ${!isAccessAllowed && "opacity-50 pointer-events-none"}`} title="Cargar proyecto guardado (.json)"><FolderOpen className="w-4 h-4 text-orange-500" />Cargar Trabajo<input type="file" className="hidden" accept=".json" onChange={handleLoadProject} onClick={(e) => (e.currentTarget.value = '')} disabled={!isAccessAllowed} /></label>
                 <button onClick={handleSaveProject} className={`flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded hover:bg-slate-50 text-sm text-slate-700 font-medium transition-colors ${!isAccessAllowed && "opacity-50 pointer-events-none"}`} disabled={!isAccessAllowed}><Save className="w-4 h-4 text-blue-500" />Guardar Trabajo</button>
                 <div className={`flex items-center rounded border ml-4 transition-colors ${(!isAccessAllowed || !isIberdrola) ? "opacity-50 pointer-events-none" : ""} ${state.masterItems.length > 0 ? "bg-emerald-100 border-emerald-300 text-emerald-800 shadow-sm" : "bg-white border-slate-300 text-slate-700 hover:bg-slate-50"}`}>
                     <label className="flex items-center gap-2 px-3 py-2 cursor-pointer text-sm font-medium grow select-none hover:bg-opacity-80 rounded-l"><Upload className={`w-4 h-4 ${state.masterItems.length > 0 ? "text-emerald-700" : "text-slate-500"}`} />{state.masterItems.length > 0 ? "Tabla Rec. (OK)" : "Importar Tabla Rec."}<input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload} onClick={(e) => (e.currentTarget.value = '')} disabled={!isAccessAllowed || !isIberdrola}/></label>
                     {state.masterItems.length > 0 && isIberdrola && (
                        <button className="px-2 py-2 border-l border-emerald-200 hover:bg-emerald-200 text-emerald-700 rounded-r relative group cursor-help" onClick={(e) => e.preventDefault()}><Info className="w-4 h-4" />
                            <div className="absolute right-0 top-full mt-3 w-72 bg-white p-4 rounded-lg shadow-xl opacity-0 group-hover:opacity-100 transition-all pointer-events-none z-[100] border text-left">
                                <span className="font-bold text-slate-800 text-xs uppercase block mb-1">Excel Cargado</span>
                                <p className="text-sm text-slate-600 break-all">{state.loadedFileName || 'Desconocido'}</p>
                                <div className="mt-4 pt-3 border-t text-[10px] uppercase font-bold text-slate-400 flex justify-between">Total Registros <span className="text-emerald-700">{state.masterItems.length} items</span></div>
                            </div>
                        </button>
                     )}
                 </div>
                 <button onClick={handleClearAll} disabled={!isAccessAllowed} className="flex items-center gap-2 px-3 py-2 bg-white border border-red-200 text-red-600 rounded hover:bg-red-50 text-sm font-medium transition-colors ml-2"><Eraser className="w-4 h-4" />Limpiar</button>
                 <div className={`relative ${!isAccessAllowed ? "opacity-50 pointer-events-none" : ""}`} ref={exportBtnRef}>
                    <button className="flex items-center gap-2 px-3 py-2 bg-emerald-600 text-white rounded hover:bg-emerald-700 text-sm font-medium shadow-sm transition-colors" onClick={() => setShowExportMenu(!showExportMenu)}>
                        <Download className="w-4 h-4" />Exportar<ChevronDown className="w-3 h-3 opacity-70" />
                    </button>
                    {showExportMenu && (
                        <div className="absolute right-0 mt-1 w-64 bg-white border border-slate-200 shadow-xl rounded-md py-1 z-50">
                             <button onClick={exportToExcel} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2"><FileSpreadsheet className="w-4 h-4 text-green-600" /> Excel (.xlsx)</button>
                             <div className="h-px bg-slate-100 my-1"></div>
                             <button onClick={() => downloadPDF('landscape', 'certification')} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2"><FileText className="w-4 h-4 text-red-600" /> PDF Certificación</button>
                             <button onClick={() => { setShowExportMenu(false); setShowProformaDialog(true); }} className="w-full text-left px-4 py-2 hover:bg-slate-50 text-sm text-slate-700 flex items-center gap-2"><CheckSquare className="w-4 h-4 text-blue-600" /> PDF Proforma (Selección)</button>
                        </div>
                    )}
                 </div>
                 <button onClick={() => setShowCalculator(!showCalculator)} className={`flex items-center gap-2 px-3 py-2 rounded text-sm font-medium ml-2 shadow-sm ${!isAccessAllowed ? "opacity-50 pointer-events-none" : ""} ${showCalculator ? "bg-slate-700 text-white" : "bg-white text-slate-700"}`} disabled={!isAccessAllowed}><Calculator className="w-4 h-4" /></button>
                 <button onClick={toggleFullscreen} className={`flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 rounded ml-2 ${!isAccessAllowed && "opacity-50"}`}>{isFullscreen ? <Minimize className="w-4 h-4" /> : <Maximize className="w-4 h-4" />}</button>
                 <button onClick={() => setShowSettingsLogin(true)} className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 text-slate-700 rounded text-sm font-medium ml-2 transition-colors"><Settings className="w-4 h-4" />Ajustes</button>
                 <button onClick={() => setShowHelpDialog(true)} className="flex items-center gap-2 px-3 py-2 bg-white border border-slate-300 text-slate-700 rounded ml-2"><HelpCircle className="w-4 h-4 text-purple-500" /></button>
             </div>
         </div>
         {isAccessAllowed && (
             <div className="px-6 py-4 bg-slate-50/50">
                 <div className="grid grid-cols-12 gap-x-6 gap-y-4">
                     <div className="col-span-6"><label className="block text-sm font-bold text-slate-400 uppercase mb-1">Denominación</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-2xl font-bold text-slate-800 pb-1 hover:border-slate-400 transition-colors" value={state.projectInfo.name} onChange={(e) => updateProjectInfo('name', e.target.value)} /></div>
                     <div className="col-span-2"><label className="block text-sm font-bold text-slate-400 uppercase mb-1">Nº Obra</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-sans text-slate-700 pb-1 hover:border-slate-400 transition-colors" value={state.projectInfo.projectNumber} onChange={(e) => updateProjectInfo('projectNumber', e.target.value)} /></div>
                     <div className="col-span-2"><label className="block text-sm font-bold text-slate-400 uppercase mb-1">Nº Pedido</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-sans text-slate-700 pb-1 hover:border-slate-400 transition-colors" value={state.projectInfo.orderNumber} onChange={(e) => updateProjectInfo('orderNumber', e.target.value)} /></div>
                     <div className="col-span-2 flex items-end">
                        {isIberdrola && (
                            <label className="flex items-center gap-2 cursor-pointer select-none bg-red-50 px-3 py-2 rounded border border-red-100 hover:bg-red-100 transition-colors w-full">
                                <input type="checkbox" className="w-5 h-5 rounded border-red-400 text-red-600" checked={state.projectInfo.isAveria || false} onChange={(e) => updateProjectInfo('isAveria', e.target.checked)} />
                                <span className="font-bold text-red-700 uppercase">AVERÍA</span>
                            </label>
                        )}
                     </div>
                     <div className="col-span-8"><label className="block text-sm text-slate-400 font-bold uppercase mb-1">Cliente</label><input type="text" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-medium text-slate-700 pb-1 hover:border-slate-400 transition-colors" value={state.projectInfo.client} onChange={(e) => updateProjectInfo('client', e.target.value)} /></div>
                     <div className="col-span-2"><label className="block text-sm text-slate-400 font-bold uppercase mb-1">Fecha</label><input type="date" className="w-full bg-transparent border-b border-slate-300 focus:border-emerald-500 outline-none text-lg font-medium text-slate-700 pb-1 cursor-pointer hover:border-slate-400 transition-colors" value={state.projectInfo.date} onChange={(e) => updateProjectInfo('date', e.target.value)} /></div>
                     <div className="col-span-2 text-right"><label className="block text-sm text-slate-400 font-bold uppercase mb-1">Total Acumulado</label><div className="text-3xl font-sans font-bold text-emerald-700 leading-none pb-1 tabular-nums">{formatCurrency(totalAmount)}</div></div>
                     {state.projectInfo.isAveria && isIberdrola && (
                       <div className="col-span-12 bg-red-50 p-4 rounded border border-red-100 mt-2 animate-in fade-in slide-in-from-top-2 shadow-sm">
                           <div className="flex items-center gap-2 mb-3 text-red-800 font-bold uppercase text-sm border-b border-red-200 pb-1"><AlertTriangle className="w-4 h-4" /> Detalles de la Avería</div>
                           <div className="grid grid-cols-12 gap-6">
                               <div className="col-span-2"><label className="block text-xs font-bold text-red-600 uppercase mb-1">Nº Avería</label><input type="text" className="w-full px-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none focus:cursor-text" value={state.projectInfo.averiaNumber || ''} onChange={(e) => updateProjectInfo('averiaNumber', e.target.value)} /></div>
                               <div className="col-span-2"><label className="block text-xs font-bold text-red-600 uppercase mb-1">Fecha Avería</label><input type="date" className="w-full px-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none cursor-pointer" value={state.projectInfo.averiaDate || ''} onChange={(e) => updateProjectInfo('averiaDate', e.target.value)} /></div>
                               <div className="col-span-2">
                                   <label className="block text-xs font-bold text-red-600 uppercase mb-1 flex items-center gap-1">
                                       Horario 
                                       <span className="relative group flex items-center">
                                           <Info className="w-3.5 h-3.5 inline text-red-400 cursor-help" />
                                           <div className="absolute left-full bottom-full mb-2 ml-1 w-80 bg-white p-4 rounded-xl shadow-2xl opacity-0 group-hover:opacity-100 transition-all pointer-events-none z-[100] border border-red-100 animate-in fade-in slide-in-from-bottom-2">
                                               <div className="flex items-center gap-2 mb-2 pb-1 border-b border-slate-100">
                                                   <Clock className="w-4 h-4 text-red-600" />
                                                   <span className="font-bold text-slate-800 text-xs uppercase tracking-tight">Criterios de Horario</span>
                                               </div>
                                               <div className="space-y-3">
                                                   <div>
                                                       <div className="flex justify-between items-center mb-0.5">
                                                           <span className="text-[10px] font-black text-red-700 uppercase">Diurna (K=1,25)</span>
                                                           <span className="bg-red-50 text-red-700 text-[9px] px-1.5 py-0.5 rounded font-bold">Laborables</span>
                                                       </div>
                                                       <p className="text-xs text-slate-600 leading-relaxed font-medium">De 07:00h a 19:00h de Lunes a Viernes.</p>
                                                   </div>
                                                   <div>
                                                       <div className="flex justify-between items-center mb-0.5">
                                                           <span className="text-[10px] font-black text-indigo-700 uppercase">Nocturna (K=1,75)</span>
                                                           <span className="bg-indigo-50 text-indigo-700 text-[9px] px-1.5 py-0.5 rounded font-bold">Especial</span>
                                                       </div>
                                                       <p className="text-xs text-slate-600 leading-relaxed font-medium">Resto de horas (19:00h a 07:00h), Fines de Semana y Festivos completos.</p>
                                                   </div>
                                               </div>
                                           </div>
                                       </span>
                                   </label>
                                   <select className="w-full px-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none" value={state.projectInfo.averiaTiming || 'diurna'} onChange={(e) => updateProjectInfo('averiaTiming', e.target.value)}><option value="diurna">Diurna K=1,25</option><option value="nocturna_finde">Nocturna K=1,75</option></select>
                               </div>
                               <div className="col-span-6"><label className="block text-xs font-bold text-red-600 uppercase mb-1">Descripción</label><textarea rows={2} className="w-full px-3 py-2 bg-white border border-red-200 rounded text-red-900 focus:outline-none resize-none focus:cursor-text" value={state.projectInfo.averiaDescription || ''} onChange={(e) => updateProjectInfo('averiaDescription', e.target.value)} /></div>
                           </div>
                       </div>
                     )}
                 </div>
             </div>
         )}
      </div>
      <div className="flex-1 overflow-hidden relative">
        {isAccessAllowed ? (
            <div className="flex flex-col h-full bg-white relative">
              <div className="flex items-center justify-between px-4 py-2 border-b border-slate-300 bg-slate-50 sticky top-0 z-20 h-12 shrink-0">
                 <div className="flex items-center gap-2">
                     <div className="relative">
                         <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
                         <input
                             type="text"
                             placeholder="Buscar en tabla... (Intro para buscar)"
                             className={`border border-slate-300 rounded-md pl-9 pr-4 py-1.5 w-72 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none transition-all ${searchQuery ? 'bg-yellow-100' : 'bg-white'}`}
                             value={searchQuery}
                             onChange={handleSearchQueryChange}
                             onKeyDown={handleSearchKeyDown}
                         />
                     </div>
                     <button
                        onClick={goToPreviousMatch}
                        disabled={searchResults.length === 0}
                        className="p-1.5 bg-white border border-slate-300 rounded-md text-slate-600 hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed"
                        aria-label="Resultado anterior"
                     >
                        <ChevronLeft className="w-4 h-4" />
                     </button>
                    <button
                        onClick={startOrContinueSearch}
                        disabled={!searchQuery.trim()}
                        className="p-1.5 bg-white border border-slate-300 rounded-md text-slate-600 hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed"
                        aria-label="Iniciar búsqueda o ir al resultado siguiente"
                    >
                        <ChevronRight className="w-4 h-4" />
                    </button>
                     {lastSearchedQuery && searchResults.length > 0 && (
                         <div className="text-sm font-semibold text-slate-600 bg-slate-200 px-3 py-1.5 rounded-md tabular-nums">
                             <span className="font-bold text-slate-800">{currentMatchIndex + 1}</span>
                             <span className="mx-1 text-slate-400">/</span>
                             <span>{searchResults.length}</span>
                         </div>
                     )}
                     {lastSearchedQuery && searchResults.length === 0 && (
                        <div className="text-sm font-medium text-red-700 bg-red-100 px-3 py-1.5 rounded-md">
                            Sin resultados
                        </div>
                     )}
                 </div>
              </div>
              <div ref={tableContainerRef} className="flex-1 overflow-auto relative bg-slate-100/50">
                <table className="w-full border-collapse text-base table-fixed min-w-[1050px] bg-white select-none">
                  <thead className="sticky top-0 z-10 bg-orange-100 text-orange-900 font-semibold border-b border-orange-200 shadow-sm">
                    <tr className="text-base uppercase tracking-wider">
                        <th className="w-12 text-center py-4 border-r border-orange-200 bg-orange-200 text-xs font-bold select-none text-orange-800">CCAA</th>
                        <th className="w-10 text-center py-4 border-r border-orange-200 bg-orange-200 text-sm select-none">#</th>
                        {isIberdrola && <th className="w-64 px-4 py-4 text-center border-r border-orange-200 font-bold">RECURSO</th>}
                        <th className="px-4 py-4 text-center border-r border-orange-200 font-bold">DESCRIPCIÓN</th>
                        <th className="w-24 px-4 py-4 text-center border-r border-orange-200 bg-orange-50 text-orange-800 font-bold">UD</th>
                        {state.projectInfo.isAveria && isIberdrola && <th className="w-20 px-4 py-4 text-center border-r border-orange-200 bg-red-100 text-red-800 font-bold">K</th>}
                        <th className="w-32 px-4 py-4 text-center border-r border-orange-200 font-bold">PRECIO</th>
                        <th className="w-36 px-4 py-4 text-center font-bold">TOTAL</th>
                        <th className="w-64 px-4 py-4 text-center border-l border-orange-200 bg-orange-50/50 font-bold">OBSERVACIONES</th>
                        <th className="w-24 px-3 py-4 text-center border-l border-orange-200 bg-orange-200 font-bold">ACCIONES</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {state.items.map((item, index) => (
                        <BudgetItemRow key={item.id} item={item} index={index} isChecked={state.checkedRowIds.has(item.id)} isSelected={selectedRowId === item.id} isInSelection={selectedCellIndices.has(index)} isAveria={!!state.projectInfo.isAveria && isIberdrola} certificationType={state.projectInfo.certificationType} activeSearch={activeSearch} editingCell={editingCell} activeSearchCell={activeSearchCell} masterItems={state.masterItems} onToggleCheck={toggleRowCheck} onSetSelectedRow={setSelectedRowId} onDragStart={handleDragStart} onDragOver={handleDragOver} onDragEnd={handleDragEnd} onDrop={handleDrop} onUpdateField={updateField} onUpdateQuantity={updateQuantity} onFillRow={fillRowWithMaster} onAddEmpty={addEmptyItem} onDelete={deleteItem} onSetActiveSearch={setActiveSearch} onSetEditingCell={setEditingCell} onCellMouseDown={handleCellMouseDown} onCellMouseEnter={handleCellMouseEnter} onInputFocus={handleInputFocus} onInputBlur={handleInputBlur} dropdownRef={dropdownRef} />
                    ))}
                  </tbody>
                </table>
              </div>
              {showProformaDialog && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-sm animate-in fade-in duration-200">
                    <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm overflow-hidden">
                        <div className="bg-slate-50 px-5 py-4 border-b flex items-center justify-between"><span className="font-bold text-slate-700">Factura Proforma</span><button onClick={() => setShowProformaDialog(false)}><X className="w-4 h-4"/></button></div>
                        <div className="p-6">
                            <p className="text-sm text-slate-600 mb-4">Margen de Beneficio (%)</p>
                            <input ref={marginInputRef} type="number" className="w-full px-4 py-2 border rounded-lg mb-6 text-xl font-bold" placeholder="0" value={proformaMargin} onChange={(e) => setProformaMargin(e.target.value)} />
                            <div className="flex gap-3"><button onClick={() => setShowProformaDialog(false)} className="flex-1 px-4 py-2 text-slate-600 hover:bg-slate-50 rounded">Cancelar</button><button onClick={() => confirmProformaExport()} className="flex-1 px-4 py-2 bg-blue-600 text-white font-bold rounded">Generar PDF</button></div>
                        </div>
                    </div>
                </div>
              )}
              {showClearDialog && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm animate-in fade-in"><div className="bg-white rounded-lg shadow-2xl w-full max-w-md overflow-hidden"><div className="p-6"><h3 className="text-xl font-bold text-slate-900 mb-4">¿Borrar todo?</h3><p className="text-slate-600 mb-6">Esta acción dejará la hoja completamente limpia.</p><div className="flex gap-3 justify-end"><button onClick={() => setShowClearDialog(false)} className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-100 rounded">Cancelar</button><button onClick={confirmClearAll} className="px-5 py-2 bg-red-600 text-white font-bold rounded hover:bg-red-700 flex items-center gap-2"><Trash2 className="w-4 h-4" />Borrar todo</button></div></div></div></div>
              )}
            {showHelpDialog && (
                <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm animate-in fade-in duration-200">
                  <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[85vh]">
                    <div className="bg-slate-900 px-6 py-4 flex items-center justify-between text-white">
                      <div className="flex items-center gap-3">
                        <HelpCircle className="w-6 h-6 text-emerald-400" />
                        <h2 className="text-xl font-bold">Guía de Uso CertiTensa</h2>
                      </div>
                      <button onClick={() => setShowHelpDialog(false)} className="text-slate-400 hover:text-white transition-colors"><X className="w-6 h-6" /></button>
                    </div>
                    <div className="flex-1 overflow-y-auto p-8 space-y-8 bg-white">
                        <section className="flex gap-4">
                            <div className="shrink-0 w-10 h-10 bg-indigo-100 text-indigo-600 rounded-full flex items-center justify-center font-bold text-lg">1</div>
                            <div>
                                <h3 className="font-bold text-slate-900 text-lg mb-1">Seleccionar Formato</h3>
                                <p className="text-slate-600 leading-relaxed">Utilice el selector superior para cambiar entre <strong>Iberdrola</strong> (que incluye códigos de recurso y cálculos de avería) y <strong>Otros Clientes</strong> (un formato simplificado de descripción y precio).</p>
                            </div>
                        </section>
                        <section className="flex gap-4">
                            <div className="shrink-0 w-10 h-10 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center font-bold text-lg">2</div>
                            <div>
                                <h3 className="font-bold text-slate-900 text-lg mb-1">Cargar Presupuesto</h3>
                                <p className="text-slate-600 leading-relaxed">En el modo Iberdrola, importe su archivo Excel para habilitar el autocompletado. Al empezar a escribir un código o descripción en las filas, se mostrarán sugerencias automáticas.</p>
                            </div>
                        </section>
                        <section className="flex gap-4">
                            <div className="shrink-0 w-10 h-10 bg-orange-100 text-orange-600 rounded-full flex items-center justify-center font-bold text-lg">3</div>
                            <div>
                                <h3 className="font-bold text-slate-900 text-lg mb-1">Gestión de Partidas</h3>
                                <p className="text-slate-600 leading-relaxed">Puede reordenar las filas arrastrándolas desde el número de posición (#). Utilice el botón <strong>"+"</strong> para insertar filas nuevas o el icono de papelera para eliminar partidas innecesarias.</p>
                            </div>
                        </section>
                        <section className="flex gap-4">
                            <div className="shrink-0 w-10 h-10 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center font-bold text-lg">4</div>
                            <div>
                                <h3 className="font-bold text-slate-900 text-lg mb-1">Generar Informes</h3>
                                <p className="text-slate-600 leading-relaxed">Exportar a <strong>Excel</strong> para cálculos internos o a <strong>PDF</strong> para entrega oficial. La opción "Proforma" permite aplicar un margen de beneficio porcentual a los precios antes de generar el documento.</p>
                            </div>
                        </section>
                    </div>
                    <div className="bg-slate-50 p-6 border-t flex justify-end">
                      <button onClick={() => setShowHelpDialog(false)} className="px-6 py-2 bg-slate-900 text-white font-bold rounded-lg hover:bg-slate-800 transition-colors">Entendido</button>
                    </div>
                  </div>
                </div>
            )}
            {showCalculator && <DraggableCalculator onClose={() => setShowCalculator(false)} />}
            </div>
        ) : (
            <div className="flex flex-col items-center justify-center h-full bg-slate-100 text-center p-8">
                <div className="bg-white p-12 rounded-2xl shadow-xl border border-red-100 max-w-lg w-full">
                    <div className="mx-auto w-24 h-24 bg-red-100 text-red-600 rounded-full flex items-center justify-center mb-6"><ShieldAlert className="w-12 h-12" /></div>
                    <h2 className="text-3xl font-bold text-slate-900 mb-4">Acceso Denegado</h2>
                    <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 mb-8 flex items-center justify-center gap-3">
                        <Globe className="w-5 h-5 text-slate-400" />
                        <span className="font-mono text-lg font-bold text-slate-700">{currentIP || "Detectando IP..."}</span>
                    </div>
                </div>
            </div>
        )}
      </div>
      <div className="bg-slate-50 border-t border-slate-300 px-4 py-2 text-sm text-slate-500 flex justify-between items-center shrink-0 font-medium">
         <div className="flex gap-4 items-center"><span className="flex items-center gap-1"><Check className="w-4 h-4 text-emerald-500"/> Listo</span><span>{state.masterItems.length} Refs</span><span>{state.items.length} Filas</span><span>{state.checkedRowIds.size} Marcadas</span></div>
         {selectedSum > 0 && isAccessAllowed && (<div className="flex items-center gap-2 bg-blue-100 text-blue-800 px-3 py-1 rounded-full"><Sigma className="w-4 h-4" /><span className="uppercase text-xs font-bold tracking-wider">Suma Seleccionada:</span><span className="font-mono font-bold text-base">{formatCurrency(selectedSum)}</span></div>)}
         <div className="font-mono opacity-50 hidden md:block">CertiTensa v5.0</div>
      </div>
      {showSettingsLogin && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in">
              <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm overflow-hidden">
                  <div className="bg-slate-900 p-6 flex flex-col items-center"><Lock className="w-6 h-6 text-white mb-3" /><h3 className="text-white font-bold text-lg">Seguridad</h3></div>
                  <form onSubmit={handleSettingsLogin} className="p-6">
                      <input autoFocus type="password" className="w-full px-4 py-3 bg-slate-50 border rounded-lg text-lg text-center tracking-widest mb-3" placeholder="••••" value={loginPasswordAttempt} onChange={(e) => setLoginPasswordAttempt(e.target.value)} />
                      {loginErrorMessage && <div className="p-2 mb-4 bg-red-100 text-red-700 text-sm font-medium rounded-lg border border-red-200">{loginErrorMessage}</div>}
                      <div className="flex gap-3"><button type="button" onClick={() => setShowSettingsLogin(false)} className="flex-1 py-3 text-slate-600 font-bold hover:bg-slate-50 rounded-lg">Cancelar</button><button type="submit" className="flex-1 py-3 bg-blue-600 text-white font-bold rounded-lg shadow-lg">Entrar</button></div>
                  </form>
              </div>
          </div>
      )}
      {showSettingsModal && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in">
              <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[85vh]">
                  <div className="bg-white border-b p-6 flex items-center justify-between"><div className="flex items-center gap-3"><div className="bg-slate-100 p-2 rounded-lg text-slate-700"><Settings className="w-6 h-6" /></div><div><h2 className="text-xl font-bold text-slate-900">Ajustes</h2></div></div><button onClick={() => setShowSettingsModal(false)}><X className="w-6 h-6" /></button></div>
                  <div className="p-6 overflow-y-auto flex-1 bg-slate-50">
                      <div className="bg-white rounded-xl border p-4 mb-6 shadow-sm">
                          <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4">Añadir IP</h4>
                          <div className="flex gap-2 mb-3"><input type="text" className="flex-1 px-4 py-2 border rounded-lg font-mono text-sm" placeholder="192.168.1.1" value={ipInput} onChange={(e) => setIpInput(e.target.value)} /><button onClick={handleAddIP} className="px-4 py-2 bg-emerald-600 text-white font-bold rounded-lg hover:bg-emerald-700 transition-colors flex items-center gap-2"><Plus className="w-4 h-4" /> Añadir</button></div>
                          <div className="bg-blue-50 px-4 py-3 rounded-lg border flex justify-between items-center text-sm"><span>Su IP actual: <strong>{currentIP}</strong></span><button onClick={() => { setIpInput(currentIP); handleAddIP(); }} className="font-bold text-blue-600">Autorizar mi IP</button></div>
                      </div>
                      <div>
                          <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">IPs Autorizadas ({allowedIPs.length})</h4>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                              {allowedIPs.map(ip => (
                                  <div key={ip} className="flex items-center justify-between bg-white px-4 py-3 rounded-lg border shadow-sm"><span className="font-mono text-slate-700">{ip}</span><button onClick={() => handleRemoveIP(ip)} className="text-slate-300 hover:text-red-500"><Trash2 className="w-4 h-4" /></button></div>
                              ))}
                          </div>
                      </div>
                  </div>
                  <div className="p-4 border-t bg-white flex justify-end"><button onClick={() => setShowSettingsModal(false)} className="px-6 py-2 bg-slate-900 text-white font-bold rounded-lg shadow-lg">Cerrar</button></div>
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