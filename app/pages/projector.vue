<script setup lang="ts">
import type { PptxSlide, PptxTransitionPlayback, PptxTriggerPlayback } from '~/utils/pptx'
import { presenterChannelName, type PresenterMessage } from '~/utils/presenter-channel'
import { useSlideTransition } from '~/composables/useSlideTransition'

interface ProjectorDocument { name: string; width: number; height: number; totalSlides: number }

const route = useRoute()
const documentInfo = shallowRef<ProjectorDocument | null>(null)
const slide = shallowRef<PptxSlide | null>(null)
const slideNumber = ref(0)
const animationStep = ref(0)
const blank = ref<'black' | 'white' | null>(null)
const triggerPlayback = shallowRef<PptxTriggerPlayback[]>([])
const transitionPlayback = shallowRef<PptxTransitionPlayback | null>(null)
const projectorStageShell = ref<HTMLElement | null>(null)
const projectorSlideHost = ref<HTMLElement | null>(null)
const projectorTransitionUnderlay = ref<HTMLElement | null>(null)
const projectorFlashOverlay = ref<HTMLElement | null>(null)
const channelError = ref('')
let channel: BroadcastChannel | undefined

const slideCountLabel = computed(() => `${slideNumber.value} / ${documentInfo.value?.totalSlides || 0}`)

function isSlide(value: unknown): value is PptxSlide {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<PptxSlide>
  return typeof candidate.id === 'string' && Array.isArray(candidate.elements)
}

function isTriggerPlayback(value: unknown, target: PptxSlide | null = slide.value): value is PptxTriggerPlayback[] {
  if (!Array.isArray(value) || !target) return false
  const sequences = target.triggeredAnimations || []
  const maximumEntries = sequences.reduce((count, sequence) => count + sequence.steps.length, 0)
  if (value.length > maximumEntries) return false
  const seen = new Set<string>()
  return value.every((item) => {
    if (!item || typeof item !== 'object') return false
    const playback = item as Partial<PptxTriggerPlayback>
    if (!Number.isSafeInteger(playback.sequence) || playback.sequence! < 0 || playback.sequence! >= sequences.length
      || !Number.isSafeInteger(playback.step) || playback.step! < 0 || playback.step! >= sequences[playback.sequence!]!.steps.length
      || !Number.isFinite(playback.startedAt)) return false
    const key = `${playback.sequence}:${playback.step}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function isTransitionPlayback(value: unknown, target: PptxSlide): value is PptxTransitionPlayback | null {
  if (value === null) return true
  if (!value || typeof value !== 'object') return false
  const playback = value as Partial<PptxTransitionPlayback>
  const transition = playback.transition
  return playback.slideId === target.id && Number.isFinite(playback.startedAtEpoch)
    && !!transition && typeof transition === 'object'
    && typeof transition.effect === 'string' && Number.isFinite(transition.durationMs) && transition.durationMs >= 0
}

function receiveMessage(message: PresenterMessage) {
  if (message.type === 'document') {
    if (!message.name || !Number.isFinite(message.width) || !Number.isFinite(message.height) || message.width <= 0 || message.height <= 0 || !Number.isInteger(message.totalSlides)) return
    documentInfo.value = { name: message.name, width: message.width, height: message.height, totalSlides: message.totalSlides }
  } else if (message.type === 'slide') {
    if (!isSlide(message.slide) || !Number.isInteger(message.number) || message.number < 1
      || !Number.isInteger(message.animationStep) || message.animationStep < 0
      || !isTriggerPlayback(message.triggerPlayback, message.slide) || !isTransitionPlayback(message.transitionPlayback, message.slide)) return
    transitionPlayback.value = message.transitionPlayback
    slide.value = message.slide
    slideNumber.value = message.number
    animationStep.value = message.animationStep
    blank.value = message.blank === 'black' || message.blank === 'white' ? message.blank : null
    triggerPlayback.value = message.triggerPlayback
  } else if (message.type === 'state') {
    if (!Number.isInteger(message.animationStep) || message.animationStep < 0 || !isTriggerPlayback(message.triggerPlayback)) return
    animationStep.value = message.animationStep
    blank.value = message.blank === 'black' || message.blank === 'white' ? message.blank : null
    triggerPlayback.value = message.triggerPlayback
  } else if (message.type === 'clear') {
    documentInfo.value = null
    slide.value = null
    slideNumber.value = 0
    animationStep.value = 0
    blank.value = null
    triggerPlayback.value = []
    transitionPlayback.value = null
  }
}

useSlideTransition({ slide, shell: projectorStageShell, incomingHost: projectorSlideHost, underlay: projectorTransitionUnderlay, flash: projectorFlashOverlay, playback: transitionPlayback })

function sendInput(action: Extract<PresenterMessage, { type: 'input' }>['action']) {
  channel?.postMessage({ type: 'input', action } satisfies PresenterMessage)
}

function toggleProjectorFullscreen() {
  if (document.fullscreenElement) void document.exitFullscreen()
  else void document.documentElement.requestFullscreen().catch(() => {})
}

function onKeydown(event: KeyboardEvent) {
  const target = event.target as HTMLElement | null
  if (target && ['INPUT', 'TEXTAREA', 'SELECT', 'BUTTON'].includes(target.tagName)) return
  if (['ArrowRight', 'ArrowDown', 'PageDown', 'Enter', ' '].includes(event.key)) {
    event.preventDefault()
    sendInput('next')
  } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(event.key)) {
    event.preventDefault()
    sendInput('previous')
  } else if (event.key.toLowerCase() === 'b') {
    sendInput(blank.value === 'black' ? 'blank-clear' : 'blank-black')
  } else if (event.key.toLowerCase() === 'w') {
    sendInput(blank.value === 'white' ? 'blank-clear' : 'blank-white')
  } else if (event.key.toLowerCase() === 'f') {
    event.preventDefault()
    if (document.fullscreenElement) void document.exitFullscreen()
    else void document.documentElement.requestFullscreen().catch(() => {})
  } else if (event.key === 'Escape') {
    if (document.fullscreenElement) return
    sendInput('close')
  }
}

watch([slide, slideNumber, animationStep, blank, triggerPlayback], () => {
  if (slide.value) channel?.postMessage({ type: 'rendered', slideId: slide.value.id, number: slideNumber.value, animationStep: animationStep.value, blank: blank.value, triggerCount: triggerPlayback.value.length } satisfies PresenterMessage)
}, { flush: 'post' })

onMounted(() => {
  const id = route.query.channel
  if (typeof id !== 'string' || !/^[a-f0-9]{32}$/.test(id)) {
    channelError.value = 'This projector link is invalid. Open the second display from Speaker View.'
    return
  }
  channel = new BroadcastChannel(presenterChannelName(id))
  channel.addEventListener('message', (event: MessageEvent<PresenterMessage>) => receiveMessage(event.data))
  channel.postMessage({ type: 'ready' } satisfies PresenterMessage)
  window.addEventListener('keydown', onKeydown)
})

onBeforeUnmount(() => {
  window.removeEventListener('keydown', onKeydown)
  channel?.postMessage({ type: 'closed' } satisfies PresenterMessage)
  channel?.close()
})
</script>

<template>
  <main class="projector-window" :class="{ 'is-projector-blank': blank }" data-projector-window>
    <template v-if="slide && documentInfo && !blank">
      <div ref="projectorStageShell" class="slide-transition-shell projector-transition-shell">
        <div ref="projectorTransitionUnderlay" class="slide-transition-underlay" :class="{ 'is-pulling': slide.transition?.effect === 'pull', 'is-splitting-in': slide.transition?.effect === 'split' && slide.transition.direction === 'in' }" aria-hidden="true" />
        <div ref="projectorSlideHost" class="slide-host" :data-slide-number="slideNumber" :style="{ '--slide-ratio': String(documentInfo.width / documentInfo.height) }">
          <SlideStage
            class="projector-stage"
            :style="{ '--slide-ratio': String(documentInfo.width / documentInfo.height) }"
            :slide="slide"
            :width="documentInfo.width"
            :height="documentInfo.height"
            :number="slideNumber"
            :animation-step="animationStep"
            :trigger-playback="triggerPlayback"
            thumbnail
            playback-preview
            :data-slide-id="slide.id"
            :data-slide-number="slideNumber"
            :data-animation-step="animationStep"
          />
        </div>
        <div ref="projectorFlashOverlay" class="flash-transition-overlay" aria-hidden="true" />
      </div>
      <div class="projector-toolbar" aria-label="Projector controls">
        <button class="projector-fullscreen-button" @click="toggleProjectorFullscreen">Fullscreen</button>
        <span>{{ slideCountLabel }} · ← → to navigate · B/W blank · Esc close</span>
      </div>
    </template>
    <div v-else-if="blank" class="projector-blank" :class="`is-${blank}`" :aria-label="`${blank} screen`" />
    <p v-else-if="channelError" class="projector-message" role="alert">{{ channelError }}</p>
    <p v-else class="projector-message" data-projector-waiting>{{ documentInfo?.name || 'Waiting for the presenter…' }}</p>
  </main>
</template>
