export default defineNuxtConfig({
    compatibilityDate: '2026-09-30',
    modules: ['@nuxtjs/i18n'],
    css: ['~/assets/css/main.css'],
    devtools: {enabled: false},
    i18n: {
        defaultLocale: 'en',
        detectBrowserLanguage: false,
        locales: [
            {code: 'en', language: 'en-US', name: 'English'},
            {code: 'zh', language: 'zh-CN', name: '中文'}
        ],
        strategy: 'no_prefix',
        vueI18n: './i18n.config.ts'
    },

    hooks: {
        'build:manifest'(manifest) {
            for (const [key, resource] of Object.entries(manifest)) {
                if (key === 'utils/pptx.ts' || resource.file.startsWith('pptx-parser.worker-')) {
                    resource.preload = false
                    resource.prefetch = false
                }
            }
        }
    },
    app: {
        head: {
            title: 'Deckline — PPTX Player',
            meta: [{name: 'viewport', content: 'width=device-width, initial-scale=1'}]
        }
    }
})
