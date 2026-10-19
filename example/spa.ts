// The Excel-like app in the browser: a workbook calculated by xlcalc in a Web Worker, opened, edited, formatted and saved as xlsx.
import { createApp } from 'vue'
import SpaApp from './SpaApp.vue'

createApp(SpaApp).mount('#app')
