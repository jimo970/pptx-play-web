export default defineNuxtConfig({
  compatibilityDate: '2026-09-30',
  css: ['~/assets/css/main.css'],
  devtools: { enabled: false },
  app: {
    head: {
      title: 'Deckline — PPTX Player',
      meta: [
        { name: 'description', content: 'Play PowerPoint presentations locally in your browser.' },
        { name: 'viewport', content: 'width=device-width, initial-scale=1' },
      ],
    },
  },
})
