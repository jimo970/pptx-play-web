export default defineNuxtConfig({
    compatibilityDate: '2026-09-30',
    css: ['~/assets/css/main.css'],
    devtools: {enabled: false},

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
