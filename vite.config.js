import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',
      includeAssets: ['qoyod_template.xlsx'],
      manifest: {
        name: 'مدقق استيراد القيود - قيود',
        short_name: 'مدقق القيود',
        description: 'أداة فحص وتجهيز ملفات القيود المحاسبية للاستيراد في قيود',
        theme_color: '#0E3B36',
        background_color: '#F7F4EC',
        display: 'standalone',
        start_url: '/',
        dir: 'rtl',
        lang: 'ar',
        icons: [
          { src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' },
          { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png' },
          { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
        ],
      },
      workbox: {
        globPatterns: ['**/*.{js,css,html,ico,png,svg,xlsx}'],
      },
    }),
  ],
})
