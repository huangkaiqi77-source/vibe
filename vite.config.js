import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  base: './', // 核心，不加打包上传必空白
  plugins: [vue()]
})
