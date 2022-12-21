import { createApp } from 'vue'
import { createPinia } from 'pinia'
import Datepicker from '@vuepic/vue-datepicker';
import '@vuepic/vue-datepicker/dist/main.css'

import App from './App.vue'
import router from './router'

import './assets/main.css'

const app = createApp(App)

app.use(createPinia())
app.use(router)
app.component('Datepicker', Datepicker);

app.mount('#app')
