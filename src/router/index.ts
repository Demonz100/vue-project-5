import { createRouter, createWebHistory } from 'vue-router'
import Import from '../views/Import.vue'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      redirect: '/import'
    },
    {
      path: '/import',
      name: 'import',
      component: Import
    },
    {
      path: '/export',
      name: 'export',
      component: () => import('../views/Export.vue')
    },
    {
      path: '/generate-pdf',
      name: 'generatePdf',
      component: () => import('../views/GeneratePDF.vue')
    },
    {
      path: '/test-pdf',
      name: 'testPdf',
      component: () => import('../views/TestPDF.vue')
    }
  ]
})

export default router
