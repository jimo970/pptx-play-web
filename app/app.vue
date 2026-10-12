<template>
  <NuxtPage />
</template>
<script setup lang="ts">
const { locale, t } = useLocale()
let localeChannel: BroadcastChannel | undefined

useHead(() => ({
  title: t('brandAria'),
  meta: [{ name: 'description', content: t('metaDescription') }],
  htmlAttrs: { lang: locale.value === 'zh' ? 'zh-CN' : 'en' },
}))

watch(locale, value => {
  if (!import.meta.client) return
  localeChannel?.postMessage(value)
})

onMounted(() => {
  if (typeof BroadcastChannel !== 'undefined') {
    localeChannel = new BroadcastChannel('deckline-locale-sync')
    localeChannel.addEventListener('message', event => {
      if ((event.data === 'zh' || event.data === 'en') && event.data !== locale.value) locale.value = event.data
    })
  }
})

onBeforeUnmount(() => {
  localeChannel?.close()
  localeChannel = undefined
})
</script>
