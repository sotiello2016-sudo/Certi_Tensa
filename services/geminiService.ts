
import { GoogleGenAI, Type } from "@google/genai";
import { BudgetItem, ProjectInfo } from "../types";

export const getProjectAnalysis = async (items: BudgetItem[], project: ProjectInfo) => {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY || '' });
  
  const summaryData = items.map(item => ({
    desc: item.description,
    planned: item.plannedQuantity,
    total: item.totalQuantity,
    percent: (item.totalQuantity / item.plannedQuantity) * 100
  }));

  const prompt = `Analiza el estado de esta certificación de obra para el proyecto "${project.name}". 
  Información del proyecto:
  - Cliente: ${project.client}
  - Certificación #: ${project.certificationNumber}
  
  Datos de ejecución:
  ${JSON.stringify(summaryData.slice(0, 20))}
  
  Por favor, proporciona un resumen ejecutivo profesional en español que incluya:
  1. Estado general del avance (en términos de presupuesto).
  2. Identificación de posibles retrasos en partidas críticas.
  3. Recomendaciones para el próximo periodo.
  
  El tono debe ser técnico y constructivo. Máximo 300 palabras.`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
    });
    return response.text;
  } catch (error) {
    console.error("Error calling Gemini API:", error);
    return "No se pudo generar el análisis automático en este momento.";
  }
};
