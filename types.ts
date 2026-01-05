
export interface BudgetItem {
  id: string;
  code: string;
  description: string;
  unit: string;
  plannedQuantity: number;
  unitPrice: number;
  // Progress tracking
  previousQuantity: number;
  currentQuantity: number;
  // New K Factor for "Averia" mode (defaults to 1)
  kFactor?: number; 
  totalQuantity: number;
  totalAmount: number;
  observations?: string;
}

export interface ProjectInfo {
  name: string;
  projectNumber: string;
  orderNumber: string;
  location: string;
  client: string;
  certificationNumber: number;
  date: string;
  // New flag for Breakdown/Fault mode
  isAveria?: boolean;
  // Averia specific details
  averiaNumber?: string;
  averiaDate?: string;
  averiaDescription?: string;
  averiaTiming?: 'diurna' | 'nocturna_finde';
}

export interface AppState {
  masterItems: BudgetItem[]; // All items from Excel
  items: BudgetItem[]; // Items selected for the current certification
  projectInfo: ProjectInfo;
  isLoading: boolean;
  checkedRowIds: Set<string>; // Items currently checked/ticked
  loadedFileName?: string; // Name of the loaded resources Excel file
  authorizedIPs?: string[]; // List of public IPs allowed to use the app
}
