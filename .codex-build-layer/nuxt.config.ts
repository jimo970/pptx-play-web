export default defineNuxtConfig({
  buildDir: '.nuxt-codex-build-check',
  nitro: {
    output: {
      dir: '.output-codex-build-check',
      serverDir: '.output-codex-build-check/server',
      publicDir: '.output-codex-build-check/public',
    },
  },
})
