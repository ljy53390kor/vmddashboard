import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  const isVercel = !!process.env.VERCEL
  return {
    plugins: [react()],
    // Vercel 배포 시에는 루트 경로('/')로 서빙되므로 Playground용 base 경로 설정을 무시
    base: isVercel ? '/' : (env.VITE_BASE_URL || '/'),
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
  }
})
