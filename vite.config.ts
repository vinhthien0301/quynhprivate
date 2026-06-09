import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  // Tải các biến môi trường từ file .env.development hoặc .env tương ứng
  const env = loadEnv(mode, process.cwd(), '')
  const port = parseInt(env.PORT || '5173', 10)
  const host = env.HOST || '127.0.0.1'
  
  // Xử lý biến ALLOWED_HOSTS (Vite yêu cầu kiểu: true | string[] | undefined)
  let allowedHosts: true | string[] | undefined = undefined
  if (env.ALLOWED_HOSTS === 'true') {
    allowedHosts = true
  } else if (env.ALLOWED_HOSTS) {
    allowedHosts = env.ALLOWED_HOSTS.split(',').map(h => h.trim())
  }

  console.log(`\x1b[36m[Vite Config] Chế độ hoạt động (mode): ${mode}\x1b[0m`)
  console.log(`\x1b[35m[Vite Config] Cổng dịch vụ (port): ${port}\x1b[0m`)
  console.log(`\x1b[32m[Vite Config] Địa chỉ host: ${host}\x1b[0m`)

  return {
    plugins: [react()],
    server: {
      port: isNaN(port) ? 5173 : port,
      host: host,
      allowedHosts: allowedHosts,
    }
  }
})
