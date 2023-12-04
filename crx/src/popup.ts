import { createApp } from 'vue'
import App from './popup/App.vue'
import pinia from '@/store'

const app = createApp(App)
app.use(pinia).mount('#app')
