import { createRouter, createWebHistory } from 'vue-router'
import CardUploadFile from '../components/CardUploadFile.vue';

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'cardpploadfile',
      component: CardUploadFile
    },
  ]
})

export default router
