import { defineConfig } from 'vite'
import { fileURLToPath } from 'url'
// exceljs: the readable browser build, which patches/exceljs+*.patch modifies (outline "collapsed" flag, rows
// without cells but with hidden / outline data are written).
const exceljs = fileURLToPath(new URL('./node_modules/exceljs/dist/exceljs.js', import.meta.url))
export default defineConfig({ root: 'src', base: './', resolve: { alias: [{ find: /^exceljs$/, replacement: exceljs }] },
  build: { outDir: '../dist', emptyOutDir: true } })
