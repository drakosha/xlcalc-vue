// The local source's page: a synthetic workbook calculated by xlcalc in a Web Worker, displayed and edited by the sheet.
import { createApp } from 'vue'
import LocalApp from './LocalApp.vue'

createApp(LocalApp).mount('#app')
