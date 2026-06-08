import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  // Tải các biến môi trường từ file .env.development hoặc .env tương ứng
  const env = loadEnv(mode, process.cwd(), '')
  const port = parseInt(env.PORT || '5173', 10)

  console.log(`\x1b[36m[Vite Config] Chế độ hoạt động (mode): ${mode}\x1b[0m`)
  console.log(`\x1b[35m[Vite Config] Cổng dịch vụ (port) được tải: ${port}\x1b[0m`)

  return {
    plugins: [react()],
    server: {
      port: isNaN(port) ? 5173 : port,
    }
  }
})
