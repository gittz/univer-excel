import { defineConfig } from 'vite'

export default defineConfig({
  // 用相对路径，部署到域名根目录或子路径都能用
  base: './',
  build: {
    outDir: 'dist',
    // Univer 本身体积较大，关掉无意义的体积警告
    chunkSizeWarningLimit: 20000,
  },
})
