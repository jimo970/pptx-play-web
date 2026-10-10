<script setup lang="ts">
import type { PptxSlide, PptxTransitionPlayback, PptxTriggerPlayback } from '~/utils/pptx-model'
import { registerPptxEmbeddedFonts, releasePptxFontFaces } from '~/utils/pptx-fonts'
import { presenterChannelName, type PresenterInkStroke, type PresenterMessage, type PresenterPoint } from '~/utils/presenter-channel'
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
const ink = shallowRef<PresenterInkStroke[]>([])
const draftStroke = shallowRef<PresenterInkStroke | null>(null)
const laserPoint = shallowRef<PresenterPoint | null>(null)
const projectorStageShell = ref<HTMLElement | null>(null)
const projectorSlideHost = ref<HTMLElement | null>(null)
const projectorTransitionUnderlay = ref<HTMLElement | null>(null)
const projectorFlashOverlay = ref<HTMLElement | null>(null)
const channelError = ref('')
const fontWarnings = ref<string[]>([])
const fontsReady = ref(false)
let channel: BroadcastChannel | undefined
let fontGeneration = 0
let fontSignature = ''
let projectorFontFaces: FontFace[] = []
let slideNumberInput = ''
let slideNumberTimer: ReturnType<typeof setTimeout> | undefined

const slideCountLabel = computed(() => `${slideNumber.value} / ${documentInfo.value?.totalSlides || 0}`)

function isSlide(value: unknown): value is PptxSlide {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<PptxSlide>
  return typeof candidate.id === 'string' && Array.isArray(candidate.elements)
}

function isPresenterPoint(value: unknown): value is PresenterPoint {
  if (!value || typeof value !== 'object') return false
  const point = value as Partial<PresenterPoint>
  return Number.isFinite(point.x) && Number.isFinite(point.y)
}

function isPresenterInkStroke(value: unknown): value is PresenterInkStroke {
  if (!value || typeof value !== 'object') return false
  const stroke = value as Partial<PresenterInkStroke>
  return (stroke.tool === 'pen' || stroke.tool === 'highlighter') && typeof stroke.points === 'string'
}

function stopAllAudio() {
  for (const audio of document.querySelectorAll<HTMLAudioElement>('audio')) audio.pause()
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
    fontWarnings.value = [...message.fontWarnings]
    const signature = JSON.stringify(message.embeddedFonts)
    if (signature !== fontSignature) {
      fontSignature = signature
      const generation = ++fontGeneration
      fontsReady.value = false
      void registerPptxEmbeddedFonts(message.embeddedFonts).then(loaded => {
        if (generation !== fontGeneration) {
          releasePptxFontFaces(loaded.faces)
          return
        }
        releasePptxFontFaces(projectorFontFaces)
        projectorFontFaces = loaded.faces
        fontWarnings.value = [...new Set([...fontWarnings.value, ...loaded.warnings])]
        fontsReady.value = true
      })
    } else if (fontsReady.value) {
      return
    }
  } else if (message.type === 'slide') {
    if (!isSlide(message.slide) || !Number.isInteger(message.number) || message.number < 1
      || !Number.isInteger(message.animationStep) || message.animationStep < 0
      || !Array.isArray(message.ink) || !message.ink.every(isPresenterInkStroke)
      || !(message.draftStroke === null || isPresenterInkStroke(message.draftStroke))
      || !(message.laserPoint === null || isPresenterPoint(message.laserPoint))
      || !isTriggerPlayback(message.triggerPlayback, message.slide) || !isTransitionPlayback(message.transitionPlayback, message.slide)) return
    transitionPlayback.value = message.transitionPlayback
    slide.value = message.slide
    slideNumber.value = message.number
    animationStep.value = message.animationStep
    blank.value = message.blank === 'black' || message.blank === 'white' ? message.blank : null
    triggerPlayback.value = message.triggerPlayback
    ink.value = message.ink
    draftStroke.value = message.draftStroke
    laserPoint.value = message.laserPoint
  } else if (message.type === 'state') {
    if (!Number.isInteger(message.animationStep) || message.animationStep < 0 || !isTriggerPlayback(message.triggerPlayback)) return
    animationStep.value = message.animationStep
    blank.value = message.blank === 'black' || message.blank === 'white' ? message.blank : null
    triggerPlayback.value = message.triggerPlayback
  } else if (message.type === 'ink-start') {
    if (message.slideId === slide.value?.id && isPresenterInkStroke(message.stroke)) draftStroke.value = message.stroke
  } else if (message.type === 'ink-point') {
    if (message.slideId === slide.value?.id && isPresenterPoint(message.point) && draftStroke.value) {
      draftStroke.value = { ...draftStroke.value, points: `${draftStroke.value.points} ${message.point.x},${message.point.y}` }
    }
  } else if (message.type === 'ink-end') {
    if (message.slideId === slide.value?.id) {
      if (draftStroke.value?.points.includes(' ')) ink.value = [...ink.value, draftStroke.value]
      draftStroke.value = null
    }
  } else if (message.type === 'ink-clear') {
    if (message.slideId === slide.value?.id) {
      ink.value = []
      draftStroke.value = null
    }
  } else if (message.type === 'laser-point') {
    if (message.slideId === slide.value?.id && (message.point === null || isPresenterPoint(message.point))) laserPoint.value = message.point
  } else if (message.type === 'clear') {
    fontGeneration++
    fontSignature = ''
    fontsReady.value = true
    releasePptxFontFaces(projectorFontFaces)
    projectorFontFaces = []
    fontWarnings.value = []
    documentInfo.value = null
    slide.value = null
    slideNumber.value = 0
    animationStep.value = 0
    blank.value = null
    triggerPlayback.value = []
    transitionPlayback.value = null
    ink.value = []
    draftStroke.value = null
    laserPoint.value = null
  }
}

useSlideTransition({ slide, shell: projectorStageShell, incomingHost: projectorSlideHost, underlay: projectorTransitionUnderlay, flash: projectorFlashOverlay, playback: transitionPlayback })

function sendInput(action: Exclude<Extract<PresenterMessage, { type: 'input' }>['action'], 'go-to'>) {
  channel?.postMessage({ type: 'input', action } satisfies PresenterMessage)
}

function sendGoToSlide(slideNumber: number) {
  channel?.postMessage({ type: 'input', action: 'go-to', slideNumber } satisfies PresenterMessage)
}

function clearSlideNumberInput() {
  slideNumberInput = ''
  if (slideNumberTimer !== undefined) clearTimeout(slideNumberTimer)
  slideNumberTimer = undefined
}

function toggleProjectorFullscreen() {
  if (document.fullscreenElement) void document.exitFullscreen()
  else void document.documentElement.requestFullscreen().catch(() => {})
}

function onKeydown(event: KeyboardEvent) {
  const target = event.target instanceof HTMLElement ? event.target : null
  if (target && (target.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT', 'AUDIO', 'VIDEO'].includes(target.tagName) || ([' ', 'Enter'].includes(event.key) && !(event.key === 'Enter' && slideNumberInput) && Boolean(target.closest('button, a, [role="button"]'))))) return
  if (event.altKey || event.ctrlKey || event.metaKey) return
  if (/^\d$/.test(event.key)) {
    event.preventDefault()
    const maxDigits = Math.max(3, String(documentInfo.value?.totalSlides || 0).length)
    slideNumberInput = `${slideNumberInput}${event.key}`.slice(-maxDigits)
    if (slideNumberTimer !== undefined) clearTimeout(slideNumberTimer)
    slideNumberTimer = setTimeout(clearSlideNumberInput, 2_000)
  } else if (event.key === 'Enter' && slideNumberInput) {
    event.preventDefault()
    const number = Number(slideNumberInput)
    clearSlideNumberInput()
    if (number > 0 && number <= (documentInfo.value?.totalSlides || 0)) sendGoToSlide(number)
  } else if (event.key === 'F5' && documentInfo.value?.totalSlides) {
    event.preventDefault()
    clearSlideNumberInput()
    if (!event.shiftKey) sendGoToSlide(1)
    if (!document.fullscreenElement) void document.documentElement.requestFullscreen().catch(() => {})
  } else if (event.key === 'Home' && documentInfo.value?.totalSlides) {
    event.preventDefault()
    clearSlideNumberInput()
    sendGoToSlide(1)
  } else if (event.key === 'End' && documentInfo.value?.totalSlides) {
    event.preventDefault()
    clearSlideNumberInput()
    sendInput('end')
  } else if (['ArrowRight', 'ArrowDown', 'PageDown', 'Enter', ' '].includes(event.key)) {
    event.preventDefault()
    sendInput('next')
  } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(event.key)) {
    event.preventDefault()
    sendInput('previous')
  } else if (['b', '.'].includes(event.key.toLowerCase())) {
    sendInput(blank.value === 'black' ? 'blank-clear' : 'blank-black')
  } else if (['w', ','].includes(event.key.toLowerCase())) {
    sendInput(blank.value === 'white' ? 'blank-clear' : 'blank-white')
  } else if (event.key.toLowerCase() === 'f') {
    event.preventDefault()
    if (document.fullscreenElement) void document.exitFullscreen()
    else void document.documentElement.requestFullscreen().catch(() => {})
  } else if (event.key === 'Escape') {
    if (slideNumberInput) {
      clearSlideNumberInput()
      return
    }
    event.preventDefault()
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => {})
    sendInput('close')
  }
}

watch([slide, slideNumber, animationStep, blank, triggerPlayback], () => {
  if (slide.value && fontsReady.value) channel?.postMessage({ type: 'rendered', slideId: slide.value.id, number: slideNumber.value, animationStep: animationStep.value, blank: blank.value, triggerCount: triggerPlayback.value.length } satisfies PresenterMessage)
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
  fontGeneration++
  clearSlideNumberInput()
  releasePptxFontFaces(projectorFontFaces)
  projectorFontFaces = []
  window.removeEventListener('keydown', onKeydown)
  channel?.postMessage({ type: 'closed' } satisfies PresenterMessage)
  channel?.close()
})
</script>

<template>
  <main class="projector-window" :class="{ 'is-projector-blank': blank }" data-projector-window>
    <template v-if="slide && documentInfo && fontsReady && !blank">
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
            @stop-all-audio="stopAllAudio"
            thumbnail
            playback-preview
            :data-slide-id="slide.id"
            :data-slide-number="slideNumber"
            :data-animation-step="animationStep"
          />
          <svg class="ink-layer projector-ink-layer" :viewBox="`0 0 ${documentInfo.width} ${documentInfo.height}`" preserveAspectRatio="none" aria-hidden="true">
            <polyline v-for="(stroke, index) in ink" :key="index" :points="stroke.points" fill="none" :stroke="stroke.tool === 'pen' ? '#ef4444' : '#fde047'" :stroke-width="stroke.tool === 'pen' ? 3 : 16" :opacity="stroke.tool === 'pen' ? 1 : .4" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
            <polyline v-if="draftStroke" :points="draftStroke.points" fill="none" :stroke="draftStroke.tool === 'pen' ? '#ef4444' : '#fde047'" :stroke-width="draftStroke.tool === 'pen' ? 3 : 16" :opacity="draftStroke.tool === 'pen' ? 1 : .4" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
            <circle v-if="laserPoint" :cx="laserPoint.x" :cy="laserPoint.y" :r="documentInfo.width / 120" fill="#ef4444" opacity=".9" />
          </svg>
        </div>
        <div ref="projectorFlashOverlay" class="flash-transition-overlay" aria-hidden="true" />
      </div>
      <div class="projector-toolbar" aria-label="Projector controls">
        <button class="projector-fullscreen-button" @click="toggleProjectorFullscreen">Fullscreen</button>
        <span>{{ slideCountLabel }} · ← → to navigate · B/W blank · Esc close</span>
        <span v-for="warning in fontWarnings" :key="warning" class="warning-pill" role="status">{{ warning }}</span>
      </div>
    </template>
    <div v-else-if="blank" class="projector-blank" :class="`is-${blank}`" :aria-label="`${blank} screen`" />
    <p v-else-if="channelError" class="projector-message" role="alert">{{ channelError }}</p>
    <p v-else class="projector-message" data-projector-waiting>{{ documentInfo?.name || 'Waiting for the presenter…' }}</p>
  </main>
</template>
