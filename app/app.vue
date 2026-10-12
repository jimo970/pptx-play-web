<template>
  <NuxtPage />
</template>
<script setup lang="ts">
import type { UiLocale } from '~/composables/useLocale'

const localePreference = useCookie<UiLocale | ''>('deckline-locale', { default: () => '', maxAge: 31_536_000 })
const { locale, t } = useLocale()
let localeChannel: BroadcastChannel | undefined

if (localePreference.value === 'en' || localePreference.value === 'zh') locale.value = localePreference.value

useHead(() => ({ title: t('brandAria'), htmlAttrs: { lang: locale.value === 'zh' ? 'zh-CN' : 'en' } }))

watch(locale, value => {
  if (!import.meta.client) return
  document.documentElement.lang = value === 'zh' ? 'zh-CN' : 'en'
  localePreference.value = value
  localeChannel?.postMessage(value)
})

onMounted(() => {
  if (!localePreference.value) locale.value = navigator.language.toLowerCase().startsWith('zh') ? 'zh' : 'en'
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
