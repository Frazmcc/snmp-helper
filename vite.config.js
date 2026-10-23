import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes('node_modules')) return
          if (id.includes('/exceljs/')) return 'excel'
          if (id.includes('/jszip/') || id.includes('/file-saver/')) return 'zip'
          if (id.includes('/js-yaml/')) return 'yaml'
          if (id.includes('/react/') || id.includes('/react-dom/')) return 'vendor'
        },
      },
    },
  },
})
