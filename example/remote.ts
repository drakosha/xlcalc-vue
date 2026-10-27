// The page of a workbook on a server (`npm run check:remote` serves it): the sheet and the toolbar over `remoteSource`.
import { createApp } from 'vue'
import RemoteApp from './RemoteApp.vue'

createApp(RemoteApp).mount('#app')
