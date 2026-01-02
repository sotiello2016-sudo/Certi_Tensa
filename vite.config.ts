import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Configuración para convertir tu código en una web real
export default defineConfig({
  plugins: [react()],
})