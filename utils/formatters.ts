
export const formatCurrency = (value: number): string => {
  return new Intl.NumberFormat('es-ES', {
    style: 'currency',
    currency: 'EUR',
  }).format(value);
};

export const formatNumber = (value: number): string => {
  return new Intl.NumberFormat('es-ES', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
};

export const roundToTwo = (num: number): number => {
  return Math.round((num + Number.EPSILON) * 100) / 100;
};

export const calculateItemTotals = (item: any) => {
  const totalQuantity = (item.previousQuantity || 0) + (item.currentQuantity || 0);
  // Default to 1 if K is not set or if we are not in a mode that uses it
  const k = (item.kFactor !== undefined && item.kFactor !== null) ? item.kFactor : 1;
  
  // Total Amount = Quantity * K * Price
  const totalAmount = roundToTwo(totalQuantity * k * item.unitPrice);
  const currentAmount = roundToTwo(item.currentQuantity * k * item.unitPrice);
  
  return { ...item, totalQuantity, totalAmount, currentAmount };
};