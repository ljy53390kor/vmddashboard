import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: process.env.NODE_ENV === 'production' ? '/vmddashboard/' : '/',
  server: {
    port: 5173,
    open: true,
    proxy: {
      '/holiday-api': {
        target: 'https://apis.data.go.kr',
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/holiday-api/, ''),
      },
    },
  },
})
